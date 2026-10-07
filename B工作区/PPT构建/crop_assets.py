from pathlib import Path
from PIL import Image

root = Path(__file__).resolve().parent
src = root / "paper-renders"
out = root / "assets"
out.mkdir(parents=True, exist_ok=True)

jobs = [
    ("scratch-2.png", (380, 75, 730, 510), "微划痕_暗场系统与图像.png"),
    ("meng-2.png", (55, 80, 705, 430), "Meng_双微镜TIR暗场_Fig1.png"),
    ("meng-3.png", (55, 75, 705, 430), "Meng_定位与粒径性能_Fig2.png"),
]

for source_name, box, output_name in jobs:
    with Image.open(src / source_name) as image:
        cropped = image.crop(box)
        cropped.save(out / output_name, quality=95)
        print(output_name, cropped.size)
