import pandas as pd
import numpy as np
from skimage.color import lab2rgb, deltaE_ciede2000
import sys
import os

def load_folio_catalog(file_path):
    """Loads the color catalog from the cleaned CSV file."""
    print("Loading full Folio catalog from CSV...")
    try:
        df = pd.read_csv(file_path)
        print(f"Loaded {len(df)} valid colors from catalog.")
        return df
    except FileNotFoundError:
        print(f"Error: Catalog file not found at {file_path}")
        sys.exit(1)

def linear_interpolate_gradient(lab_start, lab_end, steps=10):
    """Generates a list of LAB colors forming a linear gradient."""
    print("Generating linear interpolation gradient...")
    gradient_lab = []
    for i in range(steps):
        alpha = i / (steps - 1)
        interpolated_lab = [(1 - alpha) * start + alpha * end for start, end in zip(lab_start, lab_end)]
        gradient_lab.append(interpolated_lab)
    return gradient_lab

def find_closest_folio_color_robust(lab_color, catalog_df, chroma_tolerance=12.0, lightness_tolerance=10.0):
    """
    Finds the closest color using a robust three-stage approach:
    1. Filter by Chroma (C*) to match saturation level.
    2. Filter the result by Lightness (L*).
    3. Find the best match within the candidates using CIEDE2000.
    """
    target_lab = np.array(lab_color)
    target_l = target_lab[0]
    target_c = np.sqrt(target_lab[1]**2 + target_lab[2]**2)
    
    if 'Chroma' not in catalog_df.columns:
        catalog_df['Chroma'] = np.sqrt(catalog_df['Target_Coordinate2']**2 + catalog_df['Target_Coordinate3']**2)

    chroma_candidates = catalog_df[catalog_df['Chroma'] < (target_c + chroma_tolerance)]
    if chroma_candidates.empty:
        chroma_candidates = catalog_df

    l_min, l_max = target_l - lightness_tolerance, target_l + lightness_tolerance
    final_candidates = chroma_candidates[chroma_candidates['Target_Coordinate1'].between(l_min, l_max)]
    
    if final_candidates.empty:
        final_candidates = chroma_candidates
        if final_candidates.empty:
             final_candidates = catalog_df
    
    candidate_labs = final_candidates[['Target_Coordinate1', 'Target_Coordinate2', 'Target_Coordinate3']].values
    deltas = deltaE_ciede2000(target_lab, candidate_labs)
    
    closest_idx = np.argmin(deltas)
    return final_candidates.iloc[closest_idx]

def lab_to_hex(lab):
    """Converts a LAB color list to a HEX string."""
    lab_array = np.array(lab, dtype=np.float64)
    rgb_0_1 = lab2rgb(lab_array.reshape(1, 1, 3))
    rgb_0_255 = (rgb_0_1.reshape(3) * 255).astype(int)
    return f"#{rgb_0_255[0]:02x}{rgb_0_255[1]:02x}{rgb_0_255[2]:02x}".upper()

def create_scale_html(title, colors_data, show_folio_code=False):
    """Creates an HTML section for a single color scale."""
    html = f'<h2>{title}</h2><div class="color-scale">'
    for i, item in enumerate(colors_data):
        hex_color = lab_to_hex(item['lab'])
        lab_l, lab_a, lab_b = item['lab']
        
        html += '<div class="color-card">'
        html += f'<div class="color-box" style="background-color: {hex_color};"></div>'
        html += '<div class="color-info">'
        
        if show_folio_code:
            html += f"<b>{item['folio_code']}</b>"
        else:
            html += f"<b>Step {i + 1}</b>"
            
        html += f"<br>L*: {lab_l:.2f}<br>a*: {lab_a:.2f}<br>b*: {lab_b:.2f}<br>{hex_color}"
        html += '</div></div>'
    html += '</div>'
    return html

def generate_html(theoretical_colors_data, matched_colors_data, output_path):
    """Generates the complete HTML file for displaying color scales."""
    print("Generating HTML file with two scales...")
    html_template = f"""
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Color Scale Comparison</title>
        <style>
            body {{ font-family: sans-serif; margin: 2em; background-color: #f4f4f9; }}
            h1, h2 {{ text-align: center; color: #333; }}
            .color-scale {{ display: flex; flex-wrap: wrap; justify-content: center; gap: 1em; margin-bottom: 2em; }}
            .color-card {{ background-color: white; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); overflow: hidden; width: 150px; text-align: center; }}
            .color-box {{ width: 100%; height: 100px; }}
            .color-info {{ padding: 1em; font-size: 0.9em; color: #555; }}
            b {{ color: #000; }}
        </style>
    </head>
    <body>
        <h1>Color Scale Comparison (Linear Gradient)</h1>
        {create_scale_html("Theoretical Gradient (Linear Interpolation)", theoretical_colors_data)}
        {create_scale_html("Matched to Folio Catalog (Robust Method)", matched_colors_data, show_folio_code=True)}
    </body>
    </html>
    """
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html_template)
    print(f"Successfully created HTML file at '{output_path}'")

def main(catalog_path, output_html):
    """Main function to run the color scale generation."""
    catalog = load_folio_catalog(catalog_path)
    
    # LAB values for the gradient provided by user
    lab_start = [22.26, 3.29, -5.94]
    lab_end = [92.52, 1.05, -1.82]

    # 1. Generate theoretical scale
    theoretical_lab_steps = linear_interpolate_gradient(lab_start, lab_end, steps=10)
    theoretical_colors_data = [{"lab": lab} for lab in theoretical_lab_steps]
    
    # 2. Generate matched scale using the robust 3-stage method
    matched_colors_data = []
    print("\nMatching theoretical steps to Folio catalog using Robust 3-Stage Method (Chroma -> Lightness -> CIEDE2000)...")
    for i, step_lab in enumerate(theoretical_lab_steps):
        closest_match = find_closest_folio_color_robust(step_lab, catalog)
        
        matched_lab = [
            float(closest_match['Target_Coordinate1']),
            float(closest_match['Target_Coordinate2']),
            float(closest_match['Target_Coordinate3'])
        ]
        
        match_data = {
            "folio_code": closest_match['TargetName'],
            "lab": matched_lab,
        }
        matched_colors_data.append(match_data)
        print(f"Step {i+1}: Theoretical LAB {np.round(step_lab, 1)} -> Closest Folio: {match_data['folio_code']} LAB {np.round(matched_lab, 1)}")

    # 3. Generate HTML
    generate_html(theoretical_colors_data, matched_colors_data, output_html)

if __name__ == "__main__":
    script_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.dirname(script_dir)
    # Use the new clean CSV file as the data source
    catalog_path = os.path.join(project_root, 'folio_catalog_clean.csv')
    output_html_path = os.path.join(project_root, 'index.html')
    main(catalog_path, output_html_path)
