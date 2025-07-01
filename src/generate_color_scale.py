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
        
        if 'Каталог Folio' not in workbook.sheetnames:
            print("Лист 'Каталог Folio' не найден в файле.")
            return []
            
        sheet = workbook['Каталог Folio']

        for row in sheet.iter_rows(min_row=3):
            name = row[3].value
            l_val = row[42].value
            a_val = row[44].value
            b_val = row[46].value

            if name and l_val is not None and a_val is not None and b_val is not None:
                try:
                    lab = LabColor(float(l_val), float(a_val), float(b_val))
                    all_colors.append({"name": str(name), "lab": lab})
                except (ValueError, TypeError):
                    continue
        
        print(f"Успешно извлечено {len(all_colors)} цветов из файла.")
        return all_colors

    except Exception as e:
        print(f"Произошла ошибка при чтении Excel файла: {e}")
        return []

def find_closest_color(target_lab, color_catalog):
    """Находит ближайший цвет в каталоге по формуле CIEDE2000."""
    return min(color_catalog, key=lambda color: delta_e_cie2000(target_lab, color['lab']))

def generate_color_scale_html(ideal_scale, real_scale, filename="color_scale.html"):
    """Генерирует HTML-файл для визуализации двух цветовых шкал."""
    
    def generate_strip(title, scale_data):
        """Вспомогательная функция для генерации одной полосы градиента."""
        strip_html = f"<h2>{title}</h2><div class='container'>"
        for item in scale_data:
            rgb = item['rgb']
            lab_l, lab_a, lab_b = item['lab'].lab_l, item['lab'].lab_a, item['lab'].lab_b
            text_color = "black" if lab_l > 50 else "white"
            
            tooltip_text = f"LAB: {lab_l:.2f}, {lab_a:.2f}, {lab_b:.2f}<br>RGB: {rgb[0]}, {rgb[1]}, {rgb[2]}"
            if 'match_name' in item:
                tooltip_text += f"<br>Match: {item['match_name']}<br>dE2000: {item['delta_e']:.2f}"

            strip_html += f"""
            <div class="color-band" style="background-color: rgb({rgb[0]}, {rgb[1]}, {rgb[2]}); color: {text_color};" title="{tooltip_text.replace('<br>', ' ')}">
                <div class="color-info">
                    <p>Step {item['step']}</p>
                </div>
            </div>"""
        strip_html += "</div>"
        return strip_html

    html_content = """
<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <title>Color Scale Comparison</title>
    <style>
        body { font-family: sans-serif; display: flex; flex-direction: column; justify-content: center; align-items: center; min-height: 100vh; background-color: #f0f0f0; margin: 2em; }
        .container { display: flex; flex-direction: row; border: 1px solid #ccc; box-shadow: 0 4px 8px rgba(0,0,0,0.1); margin-bottom: 2em; }
        .color-band { padding: 10px; min-height: 100px; min-width: 80px; display: flex; align-items: center; justify-content: center; text-align: center; }
        .color-info p { margin: 2px 0; font-size: 0.8em; }
        h2 { text-align: center; }
    </style>
</head>
<body>
"""
    html_content += generate_strip("Ideal Gradient (Interpolated)", ideal_scale)
    html_content += generate_strip("Real Folio Colors (Closest Match)", real_scale)

    html_content += """
</body>
</html>"""
    with open(filename, "w", encoding='utf-8') as f:
        f.write(html_content)


def main():
    """
    Главная функция для генерации градиентной шкалы.
    """
    excel_file = '27.06.2025г. Каталог Folio (составы) .xlsx'
    
    color_catalog = extract_colors_from_excel(excel_file)
    
    if not color_catalog:
        print("Не удалось извлечь цвета из каталога. Завершение работы.")
        return

    start_lab = LabColor(33.0, 0.0, 0.0) # Темно-серый
    end_lab = LabColor(93.0, 0.0, 0.0)   # Светло-серый
    
    steps = 10
    ideal_color_scale = []
    real_color_scale = []

    for i in range(steps):
        t = i / (steps - 1)
        l = start_lab.lab_l + (end_lab.lab_l - start_lab.lab_l) * t
        a = start_lab.lab_a + (end_lab.lab_a - start_lab.lab_a) * t
        b = start_lab.lab_b + (end_lab.lab_b - start_lab.lab_b) * t
        interpolated_lab = LabColor(l, a, b)
        
        closest_match = find_closest_color(interpolated_lab, color_catalog)
        
        ideal_rgb = get_rgb_from_lab(interpolated_lab)
        ideal_color_scale.append({
            "step": i + 1,
            "lab": interpolated_lab,
            "rgb": ideal_rgb
        })

        closest_rgb = get_rgb_from_lab(closest_match['lab'])
        delta_e = delta_e_cie2000(interpolated_lab, closest_match['lab'])
        real_color_scale.append({
            "step": i + 1,
            "lab": closest_match['lab'],
            "rgb": closest_rgb,
            "match_name": closest_match['name'],
            "delta_e": delta_e
        })

    generate_color_scale_html(ideal_color_scale, real_color_scale, "color_scale.html")
    print("Генерация color_scale.html завершена.")

if __name__ == "__main__":
    main() 