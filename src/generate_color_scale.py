import pandas as pd
import numpy as np
from skimage.color import lab2rgb, deltaE_cie76
import sys
import os

def load_folio_catalog(input_path):
    """Loads the full Folio color catalog for matching."""
    print("Loading full Folio catalog from Excel...")
    try:
        df = pd.read_excel(input_path, sheet_name='Каталог Folio', header=0)
        df = df.iloc[1:].reset_index(drop=True) # Skip pseudo-header

        lab_cols = ['Target_Coordinate1', 'Target_Coordinate2', 'Target_Coordinate3']
        for col in lab_cols:
            df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', '.'), errors='coerce')

        df.dropna(subset=lab_cols + ['TargetName'], inplace=True)
        print(f"Loaded {len(df)} valid colors from catalog.")
        return df
    except Exception as e:
        print(f"Failed to load catalog: {e}", file=sys.stderr)
        sys.exit(1)

def interpolate_lab_colors(lab_start, lab_end, steps=10):
    """Generates a color scale by interpolating between two LAB colors."""
    print(f"Interpolating between LAB {np.round(lab_start, 2)} and {np.round(lab_end, 2)}...")
    return [lab_start + (lab_end - lab_start) * i / (steps - 1) for i in range(steps)]

def find_closest_folio_color(lab_color, catalog_df):
    """Finds the closest color in the Folio catalog using Delta E."""
    target_lab = np.array(lab_color)
    catalog_labs = catalog_df[['Target_Coordinate1', 'Target_Coordinate2', 'Target_Coordinate3']].values
    deltas = deltaE_cie76(target_lab, catalog_labs)
    closest_idx = np.argmin(deltas)
    return catalog_df.iloc[closest_idx]

def lab_to_hex(lab):
    """Converts a LAB color to an RGB hex string."""
    rgb_0_1 = lab2rgb(np.array([[lab]]))
    rgb_0_1 = np.clip(rgb_0_1, 0, 1)
    rgb_0_255 = (rgb_0_1[0][0] * 255).astype(int)
    return f"#{rgb_0_255[0]:02x}{rgb_0_255[1]:02x}{rgb_0_255[2]:02x}"

def generate_html(colors_data, output_path):
    """Generates an HTML file to display the color scale."""
    print("Generating HTML file...")
    html = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>LAB Color Gradient (Matched to Folio Catalog)</title>
    <style>
        body { 
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
            background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
            min-height: 100vh;
            margin: 0;
            padding: 20px;
            display: flex;
            flex-direction: column;
            align-items: center;
        }
        .header {
            text-align: center;
            color: #333;
            margin-bottom: 30px;
        }
        .header h1 {
            font-size: 2.5em;
            margin: 0;
            text-shadow: 1px 1px 2px rgba(0,0,0,0.1);
        }
        .header p {
            font-size: 1.2em;
            opacity: 0.9;
            margin: 10px 0;
        }
        .container {
            display: flex;
            flex-wrap: wrap;
            gap: 20px;
            justify-content: center;
            max-width: 1200px;
        }
        .card { 
            background: white;
            border-radius: 15px;
            box-shadow: 0 8px 32px rgba(0,0,0,0.1);
            width: 180px;
            transition: transform 0.3s ease, box-shadow 0.3s ease;
            overflow: hidden;
            border: 1px solid #e0e0e0;
        }
        .card:hover {
            transform: translateY(-5px);
            box-shadow: 0 12px 40px rgba(0,0,0,0.15);
        }
        .swatch { 
            width: 100%; 
            height: 120px;
        }
        .info { 
            padding: 15px; 
            text-align: center; 
            border-top: 1px solid #eee;
        }
        .info .folio { 
            font-weight: bold; 
            font-size: 1.2em; 
            color: #333;
            margin-bottom: 8px;
            font-family: 'Courier New', monospace;
        }
        .info .lab { 
            font-size: 0.85em; 
            color: #666;
            line-height: 1.4;
        }
        .info .hex {
            font-family: 'Courier New', monospace;
            font-size: 0.9em;
            color: #888;
            margin-top: 5px;
        }
        .source-info {
            background: rgba(255,255,255,0.9);
            border-radius: 10px;
            padding: 20px;
            margin: 30px 0;
            text-align: center;
            max-width: 600px;
            border: 1px solid #ddd;
        }
        .source-info h3 {
            margin: 0 0 10px 0;
            color: #333;
        }
        .source-info p {
            margin: 5px 0;
            color: #666;
        }
    </style>
</head>
<body>
    <div class="header">
        <h1>LAB Color Gradient</h1>
        <p>Interpolated colors matched to the nearest real Folio catalog color</p>
    </div>
    
    <div class="source-info">
        <h3>Source Colors from Screenshot</h3>
        <p><strong>Dark/Saturated:</strong> L*: 22.26, a*: 3.29, b*: -5.94</p>
        <p><strong>Light/Unsaturated:</strong> L*: 92.52, a*: 1.05, b*: -1.82</p>
    </div>
    
    <div class="container">
"""

    for i, item in enumerate(colors_data):
        hex_color = lab_to_hex(item['lab'])
        html += f"""
        <div class="card">
            <div class="swatch" style="background-color: {hex_color};"></div>
            <div class="info">
                <div class="folio">{item['folio_code']}</div>
                <div class="lab">L*: {item['lab'][0]:.2f}<br>a*: {item['lab'][1]:.2f}<br>b*: {item['lab'][2]:.2f}</div>
                <div class="hex">{hex_color.upper()}</div>
            </div>
        </div>
        """
    
    html += """
    </div>
</body>
</html>"""

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"Successfully created HTML file at '{output_path}'")

def main(catalog_file, output_html):
    catalog = load_folio_catalog(catalog_file)

    # LAB values from screenshot analysis (corrected)
    lab_dark_saturated = np.array([22.260, 3.294, -5.936])
    lab_light_unsaturated = np.array([92.5239, 1.0497, -1.8174])
    
    interpolated_steps = interpolate_lab_colors(lab_dark_saturated, lab_light_unsaturated)
    
    final_colors = []
    print("\nMatching interpolated steps to Folio catalog...")
    for i, step_lab in enumerate(interpolated_steps):
        closest_match = find_closest_folio_color(step_lab, catalog)
        match_data = {
            "folio_code": closest_match['TargetName'],
            "lab": (closest_match['Target_Coordinate1'], closest_match['Target_Coordinate2'], closest_match['Target_Coordinate3'])
        }
        final_colors.append(match_data)
        print(f"Step {i+1}: Theoretical LAB {np.round(step_lab, 1)} -> Closest Folio: {match_data['folio_code']} LAB {np.round(match_data['lab'], 1)}")

    generate_html(final_colors, output_html)

if __name__ == "__main__":
    catalog_path = "27.06.2025г. Каталог Folio (составы) .xlsx"
    if not os.path.exists(catalog_path):
        print(f"Error: Catalog file '{catalog_path}' not found.", file=sys.stderr)
        sys.exit(1)
    main(catalog_path, "index.html")
