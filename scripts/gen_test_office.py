# -*- coding: utf-8 -*-
"""
在 G: 盘（Test5GB）生成 .xls / .ppt 测试文件，用于数据恢复扫描与预览验证。

用法:
  pip install xlwt olefile
  python gen_test_office.py

可选: 把图片放到 G:\\recovery-test\\assets\\ 下，脚本会复制到测试目录（仅作附带文件，不嵌入 PPT）。
"""

from __future__ import annotations

import os
import random
import shutil
import ssl
import urllib.request
from datetime import datetime
from pathlib import Path

try:
    import xlwt
except ImportError:
    raise SystemExit("请先安装: pip install xlwt")

# ========== 可改配置 ==========
TARGET_DIR = Path(r"G:\recovery-test")
XLS_COUNT = 8
PPT_DOWNLOAD_COUNT = 6
# 若下载失败，可把本地 .ppt 模板放到此目录，脚本会复制并改名
LOCAL_PPT_DIR = Path(__file__).resolve().parent / "ppt_templates"
ASSETS_DIR = TARGET_DIR / "assets"
# ==============================

PPT_SAMPLE_URLS = [
    "https://filesamples.com/samples/document/ppt/sample1.ppt",
    "https://filesamples.com/samples/document/ppt/sample2.ppt",
    "https://filesamples.com/samples/document/ppt/sample3.ppt",
]


def ensure_dir(p: Path) -> None:
    p.mkdir(parents=True, exist_ok=True)


def write_xls(path: Path, title: str, rows: int) -> None:
    wb = xlwt.Workbook(encoding="utf-8")
    ws = wb.add_sheet("数据表")
    header_style = xlwt.easyxf("font: bold on; pattern: pattern solid, fore_colour light_blue;")
    ws.write(0, 0, "编号", header_style)
    ws.write(0, 1, "名称", header_style)
    ws.write(0, 2, "金额", header_style)
    ws.write(0, 3, "备注", header_style)
    ws.write(0, 4, "生成时间", header_style)
    for i in range(1, rows + 1):
        ws.write(i, 0, i)
        ws.write(i, 1, f"测试项目-{title}-{i}")
        ws.write(i, 2, round(random.uniform(10, 9999), 2))
        ws.write(i, 3, "数据恢复测试用" if i % 2 == 0 else "稳定加载预览测试")
        ws.write(i, 4, datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    wb.save(str(path))


def download_ppt(url: str, dest: Path) -> bool:
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        ctx = ssl.create_default_context()
        with urllib.request.urlopen(req, context=ctx, timeout=30) as resp:
            data = resp.read()
        if len(data) < 512 or data[:4] != b"\xd0\xcf\x11\xe0":
            return False
        dest.write_bytes(data)
        return True
    except Exception as e:
        print(f"  下载失败 {url}: {e}")
        return False


def copy_local_ppts(out_dir: Path) -> int:
    if not LOCAL_PPT_DIR.is_dir():
        return 0
    n = 0
    for src in LOCAL_PPT_DIR.glob("*.ppt"):
        dst = out_dir / f"local_{src.stem}_{n + 1}.ppt"
        shutil.copy2(src, dst)
        print(f"  复制本地模板: {dst.name}")
        n += 1
    return n


def copy_assets(out_dir: Path) -> int:
    if not ASSETS_DIR.is_dir():
        return 0
    n = 0
    exts = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".webp"}
    for src in ASSETS_DIR.iterdir():
        if src.suffix.lower() in exts and src.is_file():
            shutil.copy2(src, out_dir / f"asset_{src.name}")
            n += 1
    return n


def main() -> None:
    if not Path("G:/").exists():
        raise SystemExit("未找到 G: 盘，请确认 Test5GB 分区已挂载。")

    xls_dir = TARGET_DIR / "xls"
    ppt_dir = TARGET_DIR / "ppt"
    ensure_dir(xls_dir)
    ensure_dir(ppt_dir)
    ensure_dir(ASSETS_DIR)

    print(f"输出目录: {TARGET_DIR}\n")

    print("=== 生成 .xls ===")
    sizes = [20, 50, 100, 200, 30, 80, 150, 500]
    for i in range(XLS_COUNT):
        rows = sizes[i % len(sizes)]
        name = f"测试报表_{i + 1:02d}_{rows}行.xls"
        path = xls_dir / name
        write_xls(path, f"表{i + 1}", rows)
        print(f"  {name} ({path.stat().st_size // 1024} KB)")

    print("\n=== 获取 .ppt ===")
    ppt_ok = 0
    for i, url in enumerate(PPT_SAMPLE_URLS[:PPT_DOWNLOAD_COUNT]):
        name = f"下载样本_{i + 1:02d}.ppt"
        path = ppt_dir / name
        if download_ppt(url, path):
            print(f"  {name} ({path.stat().st_size // 1024} KB)")
            ppt_ok += 1

    local_n = copy_local_ppts(ppt_dir)
    ppt_ok += local_n

    if ppt_ok == 0:
        print(
            "  未能获取任何 .ppt。可选方案:\n"
            "  1) 检查网络后重试\n"
            "  2) 把任意 .ppt 放到 scripts/ppt_templates/ 再运行\n"
            "  3) 用 PowerPoint/WPS 另存为 .ppt 放到 G:\\recovery-test\\ppt\\"
        )

    asset_n = copy_assets(TARGET_DIR)
    if asset_n:
        print(f"\n=== 已复制 {asset_n} 个图片到测试目录 ===")
    else:
        print(
            "\n提示: 图片非必须。若需要，可放到 G:\\recovery-test\\assets\\ 后重新运行。"
        )

    print("\n完成。建议测试流程:")
    print("  1) 记录当前扫描基线")
    print("  2) 删除 G:\\recovery-test 下部分文件")
    print("  3) 再扫描，对比能否找到并稳定预览 xls/ppt")


if __name__ == "__main__":
    main()
