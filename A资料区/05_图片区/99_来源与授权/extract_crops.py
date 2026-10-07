from pathlib import Path
from PIL import Image


ROOT = Path(__file__).resolve().parent


JOBS = [
    # source, crop box (left, top, right, bottom), output
    ("Davidson_p-25.png", (680, 845, 1315, 1645), ROOT.parent / "01_原理图" / "01_透射暗场_遮住直射光的空心光锥.png"),
    ("Davidson_p-26.png", (420, 220, 960, 850), ROOT.parent / "04_典型成像" / "01_硅藻暗场图像.png"),
    ("Zsigmondy_p-05.png", (190, 930, 970, 1515), ROOT.parent / "03_典型结构" / "01_1902超显微镜侧向照明结构.png"),
    ("Zsigmondy_p-06.png", (145, 780, 980, 1460), ROOT.parent / "02_实物与核心部件" / "01_历史浸没式超显微镜实物.png"),
    ("Zamora_p-02.png", (240, 1040, 1090, 1530), ROOT.parent / "05_现代扩展" / "01_高光谱增强暗场系统.png"),
    ("Enoki_p-2.png", (95, 90, 1040, 690), ROOT.parent / "03_典型结构" / "02_物镜型TIR暗场光路.png"),
    ("Enoki_p-3.png", (95, 90, 1010, 760), ROOT.parent / "04_典型成像" / "02_流感病毒浓度与暗场亮点.png"),
]


for source_name, box, output_path in JOBS:
    with Image.open(ROOT / source_name) as image:
        cropped = image.crop(box)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        cropped.save(output_path, optimize=True)
        print(f"{output_path.name}: {cropped.size}")
