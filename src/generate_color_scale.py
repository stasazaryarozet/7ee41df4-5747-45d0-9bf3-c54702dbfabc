import pandas as pd
import numpy as np
from skimage.color import lab2rgb, deltaE_cie76
import sys
import os

def interpolate_lab_colors(lab_start, lab_end, steps=10):
    """Generates a color scale by interpolating between two LAB colors."""
    print(f"Interpolating between LAB {np.round(lab_start, 2)} and {np.round(lab_end, 2)}...")
    return [lab_start + (lab_end - lab_start) * i / (steps - 1) for i in range(steps)]

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
    <title>LAB Color Gradient from Screenshot</title>
    <style>
        body { 
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            margin: 0;
            padding: 20px;
            display: flex;
            flex-direction: column;
            align-items: center;
        }
        .header {
            text-align: center;
            color: white;
            margin-bottom: 30px;
        }
        .header h1 {
            font-size: 2.5em;
            margin: 0;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
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
        }
        .card:hover {
            transform: translateY(-5px);
            box-shadow: 0 12px 40px rgba(0,0,0,0.2);
        }
        .swatch { 
            width: 100%; 
            height: 120px;
            position: relative;
        }
        .swatch::after {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: linear-gradient(45deg, transparent 40%, rgba(255,255,255,0.1) 50%, transparent 60%);
        }
        .info { 
            padding: 15px; 
            text-align: center; 
        }
        .info .step { 
            font-weight: bold; 
            font-size: 1.1em; 
            color: #333;
            margin-bottom: 8px;
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
        <p>Generated from Paletton Screenshot Analysis</p>
    </div>
    
         <div class="source-info">
         <h3>Source Colors from Screenshot</h3>
         <p><strong>Dark/Saturated:</strong> L*: 22.26, a*: 3.29, b*: -5.94</p>
         <p><strong>Light/Unsaturated:</strong> L*: 92.52, a*: 1.05, b*: -1.82</p>
         <p>10-step linear interpolation in LAB color space</p>
     </div>
    
    <div class="container">
"""

    for i, item in enumerate(colors_data):
        hex_color = lab_to_hex(item['lab'])
        html += f"""
        <div class="card">
            <div class="swatch" style="background-color: {hex_color};"></div>
            <div class="info">
                <div class="step">Step {i+1}</div>
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

def main():
    # LAB values from screenshot analysis (corrected)
    lab_dark_saturated = np.array([22.260, 3.294, -5.936])  # Темный/насыщенный
    lab_light_unsaturated = np.array([92.5239, 1.0497, -1.8174])  # Светлый/ненасыщенный
    
    print("=== LAB Parameters from Screenshot (Corrected) ===")
    print(f"Dark/Saturated: L*: {lab_dark_saturated[0]}, a*: {lab_dark_saturated[1]}, b*: {lab_dark_saturated[2]}")
    print(f"Light/Unsaturated: L*: {lab_light_unsaturated[0]}, a*: {lab_light_unsaturated[1]}, b*: {lab_light_unsaturated[2]}")
    print("=================================================")
    
    interpolated_steps = interpolate_lab_colors(lab_dark_saturated, lab_light_unsaturated)
    
    final_colors = []
    print("\nGenerating 10-step gradient...")
    for i, step_lab in enumerate(interpolated_steps):
        color_data = {
            "lab": step_lab
        }
        final_colors.append(color_data)
        print(f"Step {i+1}: LAB ({step_lab[0]:.2f}, {step_lab[1]:.2f}, {step_lab[2]:.2f}) -> HEX {lab_to_hex(step_lab)}")

    generate_html(final_colors, "index.html")

if __name__ == "__main__":
    main()
