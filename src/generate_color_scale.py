import pandas as pd
import numpy as np
from skimage.color import lab2rgb, deltaE_cie76, lab2lch, lch2lab
import sys
import os
import math

def load_folio_catalog(input_path):
    """Loads the full Folio color catalog from the Excel file."""
    print("Loading full Folio catalog from Excel...")
    try:
        df = pd.read_excel(input_path, sheet_name='Каталог Folio', header=0)
        df = df.iloc[1:].reset_index(drop=True)
        lab_cols = ['Target_Coordinate1', 'Target_Coordinate2', 'Target_Coordinate3']
        for col in lab_cols:
            df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', '.'), errors='coerce')
        df.dropna(subset=lab_cols + ['TargetName'], inplace=True)
        print(f"Loaded {len(df)} valid colors from catalog.")
        return df
    except Exception as e:
        print(f"Failed to load catalog: {e}", file=sys.stderr)
        sys.exit(1)

def generate_alternating_gradient(lab_start, lab_end, steps=10, start_with_lightness=True):
    """
    Generates a gradient by alternating changes in Lightness and Chroma
    at a constant average Hue.
    """
    print("Generating alternating gradient...")
    lch_start = lab2lch(np.array([[lab_start]]))[0][0]
    lch_end = lab2lch(np.array([[lab_end]]))[0][0]

    l_start, c_start, h_start_deg = lch_start
    l_end, c_end, h_end_deg = lch_end
    
    # Correctly average hue (angles in degrees)
    h_start_rad = math.radians(h_start_deg)
    h_end_rad = math.radians(h_end_deg)
    avg_hue_rad = math.atan2(
        (math.sin(h_start_rad) + math.sin(h_end_rad)) / 2,
        (math.cos(h_start_rad) + math.cos(h_end_rad)) / 2
    )
    avg_hue_deg = math.degrees(avg_hue_rad) % 360

    print(f"Start LCH: {np.round(lch_start,1)}, End LCH: {np.round(lch_end,1)}")
    print(f"Using constant average Hue: {avg_hue_deg:.1f}°")

    l_current, c_current = l_start, c_start
    
    num_l_steps = math.ceil((steps -1) / 2) if start_with_lightness else math.floor((steps -1) / 2)
    num_c_steps = math.floor((steps -1) / 2) if start_with_lightness else math.ceil((steps-1) / 2)

    delta_l_per_step = (l_end - l_start) / num_l_steps if num_l_steps > 0 else 0
    delta_c_per_step = (c_end - c_start) / num_c_steps if num_c_steps > 0 else 0
    
    gradient_lch = [lch_start]
    for i in range(1, steps):
        is_lightness_step = (i % 2 != 0) if start_with_lightness else (i % 2 == 0)
        if is_lightness_step:
            l_current += delta_l_per_step
        else:
            c_current += delta_c_per_step
        
        # Ensure chroma is non-negative
        c_actual = max(0, c_current)
        gradient_lch.append([l_current, c_actual, avg_hue_deg])

    # Convert back to LAB
    return [lch2lab(np.array([[lch]]))[0][0] for lch in gradient_lch]


def find_closest_folio_color(lab_color, catalog_df):
    """Finds the closest color in the Folio catalog using Delta E."""
    target_lab = np.array(lab_color)
    catalog_labs = catalog_df[['Target_Coordinate1', 'Target_Coordinate2', 'Target_Coordinate3']].values
    deltas = deltaE_cie76(target_lab, catalog_labs)
    closest_idx = np.argmin(deltas)
    return catalog_df.iloc[closest_idx]

def lab_to_hex(lab):
    """Converts a LAB color to an RGB hex string."""
    with np.errstate(invalid='ignore'): # Suppress minor warnings from skimage
        rgb_0_1 = lab2rgb(np.array([[lab]]))
    rgb_0_1 = np.clip(rgb_0_1, 0, 1)
    rgb_0_255 = (rgb_0_1[0][0] * 255).astype(int)
    return f"#{rgb_0_255[0]:02x}{rgb_0_255[1]:02x}{rgb_0_255[2]:02x}"

def generate_html(theoretical_colors, matched_colors, output_path):
    """Generates an HTML file to display both color scales."""
    print("Generating HTML file with two scales...")

    def create_scale_html(title, colors, show_folio_code=False):
        scale_html = f"<h2>{title}</h2><div class='container'>"
        for i, item in enumerate(colors):
            hex_color = lab_to_hex(item['lab'])
            folio_code_html = f"<div class='folio'>{item['folio_code']}</div>" if show_folio_code else f"<div class='step'>Step {i+1}</div>"
            scale_html += f"""
            <div class="card">
                <div class="swatch" style="background-color: {hex_color};"></div>
                <div class="info">
                    {folio_code_html}
                    <div class="lab">L*: {item['lab'][0]:.2f}<br>a*: {item['lab'][1]:.2f}<br>b*: {item['lab'][2]:.2f}</div>
                    <div class="hex">{hex_color.upper()}</div>
                </div>
            </div>
            """
        scale_html += "</div>"
        return scale_html

    html = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Color Scale Comparison</title>
    <style>
        body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; background-color: #f0f2f5; margin: 0; padding: 20px; }
        h1, h2 { text-align: center; color: #333; }
        .container { display: flex; flex-wrap: wrap; gap: 20px; justify-content: center; margin-bottom: 40px; }
        .card { background: white; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); width: 160px; overflow: hidden; }
        .swatch { width: 100%; height: 120px; }
        .info { padding: 12px; text-align: center; font-size: 0.9em; }
        .info .folio, .info .step { font-weight: bold; font-size: 1.1em; color: #333; margin-bottom: 6px; }
        .info .lab { font-size: 0.85em; color: #666; line-height: 1.4; }
        .info .hex { font-family: 'SF Mono', 'Courier New', monospace; font-size: 0.9em; color: #888; margin-top: 5px; }
    </style>
</head>
<body>
    <h1>Color Scale Comparison</h1>
"""
    html += create_scale_html("Theoretical Gradient (Alternating L*/C*)", theoretical_colors)
    html += create_scale_html("Matched to Folio Catalog", matched_colors, show_folio_code=True)
    html += "</body></html>"

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"Successfully created HTML file at '{output_path}'")

def main(catalog_file, output_html):
    catalog = load_folio_catalog(catalog_file)
    lab_start = np.array([22.260, 3.294, -5.936])
    lab_end = np.array([92.5239, 1.0497, -1.8174])
    
    # 1. Generate theoretical "clean" scale
    theoretical_lab_steps = generate_alternating_gradient(lab_start, lab_end, steps=10, start_with_lightness=True)
    theoretical_colors_data = [{"lab": lab} for lab in theoretical_lab_steps]
    
    # 2. Generate matched scale
    matched_colors_data = []
    print("\nMatching theoretical steps to Folio catalog...")
    for i, step_lab in enumerate(theoretical_lab_steps):
        closest_match = find_closest_folio_color(step_lab, catalog)
        match_data = {
            "folio_code": closest_match['TargetName'],
            "lab": (closest_match['Target_Coordinate1'], closest_match['Target_Coordinate2'], closest_match['Target_Coordinate3'])
        }
        matched_colors_data.append(match_data)
        print(f"Step {i+1}: Theoretical LAB {np.round(step_lab, 1)} -> Closest Folio: {match_data['folio_code']} LAB {np.round(match_data['lab'], 1)}")

    generate_html(theoretical_colors_data, matched_colors_data, output_html)

if __name__ == "__main__":
    catalog_path = "27.06.2025г. Каталог Folio (составы) .xlsx"
    if not os.path.exists(catalog_path):
        print(f"Error: Catalog file '{catalog_path}' not found.", file=sys.stderr)
        sys.exit(1)
    main(catalog_path, "index.html")
