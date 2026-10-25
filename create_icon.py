from PIL import Image, ImageDraw, ImageFont

def make_excel_icon():
    sizes = [(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)]
    images = []
    
    for size in sizes:
        w, h = size
        img = Image.new("RGBA", (w, h), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        
        # Rounded green background
        r = int(w * 0.18)
        draw.rounded_rectangle([0, 0, w - 1, h - 1], radius=r, fill="#107C41")
        
        # Dark green folded document shape / sheet accent
        margin = int(w * 0.12)
        draw.rounded_rectangle([margin, margin, w - margin, h - margin], radius=int(r * 0.7), fill="#0E6435")
        
        # Inner Excel grid lines or letter X
        x_pad = int(w * 0.28)
        y_pad = int(h * 0.28)
        stroke_w = max(2, int(w * 0.14))
        
        # Draw "X"
        draw.line([(x_pad, y_pad), (w - x_pad, h - y_pad)], fill="white", width=stroke_w)
        draw.line([(x_pad, h - y_pad), (w - x_pad, y_pad)], fill="white", width=stroke_w)
        
        images.append(img)
        
    images[0].save(
        "app_icon.ico",
        format="ICO",
        sizes=[(s[0], s[1]) for s in sizes],
        append_images=images[1:]
    )
    images[0].save("static/favicon.ico", format="ICO")
    print("app_icon.ico created successfully!")

if __name__ == '__main__':
    make_excel_icon()
