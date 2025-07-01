import pandas as pd
from openpyxl import load_workbook
from colormath.color_objects import LabColor, sRGBColor
from colormath.color_diff import delta_e_cie2000
from colormath.color_conversions import convert_color
import numpy as np

# Monkey-patch для numpy.asscalar, который был удален в новых версиях.
# colormath 3.0.0 все еще его использует.
if not hasattr(np, 'asscalar'):
    np.asscalar = lambda x: x.item()

def get_rgb_from_lab(lab_color):
    """Преобразует объект LabColor в кортеж RGB."""
    srgb = convert_color(lab_color, sRGBColor)
    # Ограничиваем значения в диапазоне 0-1 перед конвертацией
    clamped_rgb = tuple(max(0, min(1, val)) for val in srgb.get_value_tuple())
    return tuple(int(c * 255) for c in clamped_rgb)

def extract_colors_from_excel(file_path):
    """
    Извлекает данные о цветах из указанного Excel-файла,
    обрабатывая ссылки на общие строки и пропуская заголовки.
    """
    all_colors = []
    try:
        # data_only=True необходимо, чтобы openpyxl разрешал ссылки (shared strings)
        workbook = load_workbook(filename=file_path, read_only=True, data_only=True)
        
        # Убеждаемся, что лист существует
        if 'Каталог Folio' not in workbook.sheetnames:
            print("Лист 'Каталог Folio' не найден в файле.")
            return []
            
        sheet = workbook['Каталог Folio']

        # Пропускаем заголовки, начинаем с 3-й строки
        for row in sheet.iter_rows(min_row=3):
            # Извлекаем значения по индексам колонок (A=0, B=1, и т.д.)
            # Name -> D (индекс 3)
            # L* -> AQ (индекс 42)
            # a* -> AS (индекс 44)
            # b* -> AU (индекс 46)
            name = row[3].value
            l_val = row[42].value
            a_val = row[44].value
            b_val = row[46].value

            if name and l_val is not None and a_val is not None and b_val is not None:
                try:
                    # Преобразуем значения в float, обрабатывая возможные ошибки
                    lab = LabColor(float(l_val), float(a_val), float(b_val))
                    all_colors.append({"name": str(name), "lab": lab})
                except (ValueError, TypeError):
                    # Пропускаем строки с некорректными числовыми значениями
                    continue
        
        print(f"Успешно извлечено {len(all_colors)} цветов из файла.")
        return all_colors

    except Exception as e:
        print(f"Произошла ошибка при чтении Excel файла: {e}")
        return []

def find_closest_color(target_lab, color_catalog):
    """Находит ближайший цвет в каталоге по формуле CIEDE2000."""
    return min(color_catalog, key=lambda color: delta_e_cie2000(target_lab, color['lab']))

def generate_color_scale_html(color_scale, filename="color_scale.html"):
    """Генерирует HTML-файл для визуализации цветовой шкалы."""
    html_content = """
<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <title>Color Scale</title>
    <style>
        body { font-family: sans-serif; display: flex; justify-content: center; align-items: center; min-height: 100vh; background-color: #f0f0f0; margin: 0; }
        .container { display: flex; flex-direction: column; border: 1px solid #ccc; box-shadow: 0 4px 8px rgba(0,0,0,0.1); }
        .color-band { padding: 20px; color: white; text-shadow: 1px 1px 2px rgba(0,0,0,0.7); min-height: 80px; display: flex; align-items: center; justify-content: center; }
        .color-info { text-align: center; }
        .color-info p { margin: 2px 0; }
    </style>
</head>
<body>
    <div class="container">
"""
    for item in color_scale:
        i = item['step'] - 1
        rgb = item['rgb']
        lab_l, lab_a, lab_b = item['lab'].lab_l, item['lab'].lab_a, item['lab'].lab_b
        source = item['source']
        match_name = item['match_name']
        delta_e = item['delta_e']
        
        # Выбираем цвет текста в зависимости от яркости фона
        text_color = "black" if lab_l > 50 else "white"

        html_content += f"""
        <div class="color-band" style="background-color: rgb({rgb[0]}, {rgb[1]}, {rgb[2]}); color: {text_color};">
            <div class="color-info">
                <p>Step {i+1}</p>
                <p>LAB: {lab_l:.2f}, {lab_a:.2f}, {lab_b:.2f}</p>
                <p>RGB: {rgb[0]}, {rgb[1]}, {rgb[2]}</p>
                <p>Closest Match: {match_name} (dE2000: {delta_e:.2f})</p>
            </div>
        </div>"""
    html_content += """
    </div>
</body>
</html>"""
    with open(filename, "w", encoding='utf-8') as f:
        f.write(html_content)

def main():
    """
    Главная функция для генерации градиентной шкалы.
    """
    excel_file = '27.06.2025г. Каталог Folio (составы) .xlsx'
    
    # 1. Извлекаем всю палитру цветов из Excel
    color_catalog = extract_colors_from_excel(excel_file)
    
    if not color_catalog:
        print("Не удалось извлечь цвета из каталога. Завершение работы.")
        return

    # 2. Определяем начальный и конечный цвета для градиента
    start_lab = LabColor(33.0, 0.0, 0.0) # Темно-серый
    end_lab = LabColor(93.0, 0.0, 0.0)   # Светло-серый
    
    steps = 10
    color_scale = []

    # 3. Генерируем градиент и подбираем цвета
    for i in range(steps):
        t = i / (steps - 1)
        l = start_lab.lab_l + (end_lab.lab_l - start_lab.lab_l) * t
        a = start_lab.lab_a + (end_lab.lab_a - start_lab.lab_a) * t
        b = start_lab.lab_b + (end_lab.lab_b - start_lab.lab_b) * t
        interpolated_lab = LabColor(l, a, b)
        
        # Находим ближайший реальный цвет из каталога
        closest_match = find_closest_color(interpolated_lab, color_catalog)
        
        # Получаем RGB для отображения реального цвета
        closest_rgb = get_rgb_from_lab(closest_match['lab'])
        
        delta_e = delta_e_cie2000(interpolated_lab, closest_match['lab'])

        color_scale.append({
            "step": i + 1,
            "source": "Interpolated",
            "lab": interpolated_lab,
            "rgb": closest_rgb, # Используем RGB ближайшего реального цвета
            "match_name": closest_match['name'],
            "delta_e": delta_e
        })

    # 4. Создаем HTML-файл для визуализации
    generate_color_scale_html(color_scale, "color_scale.html")
    print("Генерация color_scale.html завершена.")

if __name__ == "__main__":
    main() 