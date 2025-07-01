import pandas as pd
import numpy as np
from skimage.color import lab2rgb, deltaE_cie76, deltaE_ciede2000
import sys
import os

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

def linear_interpolate_gradient(lab_start, lab_end, steps=10):
    """Generates a gradient using simple linear interpolation."""
    print("Generating linear interpolation gradient...")
    gradient_lab = []
    for i in range(steps):
        fraction = i / (steps - 1)
        lab_step = lab_start + (lab_end - lab_start) * fraction
        gradient_lab.append(lab_step)
    return gradient_lab

def find_closest_folio_color_hybrid(lab_color, catalog_df, lightness_tolerance=5.0):
    """
    Finds the closest color using a hybrid two-stage approach:
    1. Filter by a lightness (L*) tolerance.
    2. Find the best match within the filtered candidates using CIEDE2000.
    """
    target_lab = np.array(lab_color)
    
    # Stage 1: Filter by Lightness
    l_target = target_lab[0]
    l_min, l_max = l_target - lightness_tolerance, l_target + lightness_tolerance
    
    candidates = catalog_df[catalog_df['Target_Coordinate1'].between(l_min, l_max)]
    
    # If no candidates are found, widen the search to the full catalog
    if candidates.empty:
        candidates = catalog_df

    # Stage 2: Find the best match in the candidate pool using CIEDE2000
    candidate_labs = candidates[['Target_Coordinate1', 'Target_Coordinate2', 'Target_Coordinate3']].values
    deltas = deltaE_ciede2000(target_lab, candidate_labs)
    
    closest_idx_in_candidates = np.argmin(deltas)
    return candidates.iloc[closest_idx_in_candidates]

def lab_to_hex(lab):
    """Converts a LAB color to an RGB hex string."""
    with np.errstate(invalid='ignore'):
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
    <title>Color Scale Comparison (Linear + CIE76)</title>
    <style>
        body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; background-color: #f0f2f5; margin: 0; padding: 20px; }
        h1, h2 { text-align: center; color: #333; }
        h2 { border-top: 1px solid #ddd; padding-top: 20px; margin-top: 40px;}
        .container { display: flex; flex-wrap: wrap; gap: 20px; justify-content: center; margin-bottom: 20px; }
        .card { background: white; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); width: 160px; overflow: hidden; }
        .swatch { width: 100%; height: 120px; }
        .info { padding: 12px; text-align: center; font-size: 0.9em; }
        .info .folio, .info .step { font-weight: bold; font-size: 1.1em; color: #333; margin-bottom: 6px; }
        .info .lab { font-size: 0.85em; color: #666; line-height: 1.4; }
        .info .hex { font-family: 'SF Mono', 'Courier New', monospace; font-size: 0.9em; color: #888; margin-top: 5px; }
    </style>
</head>
<body>
    <h1>Color Scale Comparison (Linear Gradient)</h1>
"""
    html += create_scale_html("Theoretical Gradient (Linear Interpolation)", theoretical_colors)
    html += create_scale_html("Matched to Folio Catalog (Hybrid Method)", matched_colors, show_folio_code=True)
    html += "</body></html>"

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"Successfully created HTML file at '{output_path}'")

def main(catalog_file, output_html):
    catalog = load_folio_catalog(catalog_file)
    lab_start = np.array([22.260, 3.294, -5.936])
    lab_end = np.array([92.5239, 1.0497, -1.8174])
    
    # 1. Generate theoretical "clean" scale using linear interpolation
    theoretical_lab_steps = linear_interpolate_gradient(lab_start, lab_end, steps=10)
    theoretical_colors_data = [{"lab": lab} for lab in theoretical_lab_steps]
    
    # 2. Generate matched scale using the hybrid method
    matched_colors_data = []
    print("\nMatching theoretical steps to Folio catalog using Hybrid Method (L* filter + CIEDE2000)...")
    for i, step_lab in enumerate(theoretical_lab_steps):
        closest_match = find_closest_folio_color_hybrid(step_lab, catalog, lightness_tolerance=7.5) # Increased tolerance slightly for better matching
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
