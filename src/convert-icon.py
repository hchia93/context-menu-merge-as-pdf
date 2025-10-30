"""
Convert icon.png to icon.ico for Windows context menu.
Windows context menus require .ico format for proper icon display.
"""

from PIL import Image
import sys

def convert_png_to_ico(png_path="icon.png", ico_path="icon.ico"):
    try:
        # Open the PNG image
        img = Image.open(png_path)
        
        # Convert to RGB if necessary (remove alpha channel issues)
        if img.mode != 'RGBA':
            img = img.convert('RGBA')
        
        # Create multiple sizes for better display at different resolutions
        sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (256, 256)]
        
        # Save as ICO with multiple sizes
        img.save(ico_path, format='ICO', sizes=sizes)
        
        print(f"Successfully converted {png_path} to {ico_path}")
        print(f"Generated sizes: {', '.join([f'{w}x{h}' for w, h in sizes])}")
        
    except FileNotFoundError:
        print(f"ERROR: {png_path} not found in current directory")
        sys.exit(1)
    except Exception as e:
        print(f"ERROR: Error converting icon: {e}")
        sys.exit(1)

if __name__ == "__main__":
    convert_png_to_ico()