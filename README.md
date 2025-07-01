# LAB Color Gradient from Paletton Screenshot

This project generates a 10-step color gradient based on LAB color space interpolation between two colors extracted from a Paletton screenshot.

## Source Colors

- **Dark/Saturated**: L*: 22.36, a*: -5.84, b*: -0.81
- **Light/Unsaturated**: L*: 92.52, a*: 1.54, b*: -0.82

## Live Demo

Visit the live demo: [LAB Color Gradient](https://azaryarozet.github.io/lab-color-gradient/)

## Technical Details

- Linear interpolation in LAB color space
- 10-step gradient generation
- LAB to RGB conversion using scikit-image
- Responsive HTML/CSS interface

## Files

- `index.html` - Main color gradient visualization
- `src/generate_color_scale.py` - Python script for gradient generation
- Color values extracted from Paletton interface screenshot
