# -*- coding: utf-8 -*-
"""
【精准模式】PDF周期性页提取工具
需求：单个PDF中，每4页为1组（1-4, 5-8, 9-12...），每组提取前2页 → 页码序列：1,2,5,6,9,10,13,14...
输入：用户指定的PDF路径（支持中文、空格）
输出：当前目录下生成 extracted_4pages_pattern.pdf
执行后暂停窗口，便于查看日志
"""
import os
import sys
from pathlib import Path

# --- 🔧 用户可配置区（仅修改此处）---
INPUT_PDF_PATH = r"C:\Users\Administrator\Desktop\215\combined.pdf"  # ← 替换为您要处理的PDF完整路径！
OUTPUT_FILENAME = "extracted_4pages_pattern.pdf"                     # ← 输出文件名（保存在本脚本所在目录）

# --- 📦 依赖自动检测与导入 ---
try:
    from pypdf import PdfReader, PdfWriter
except ImportError:
    try:
        from PyPDF2 import PdfReader, PdfWriter
    except ImportError:
        print("❌ 错误：未安装 pypdf 或 PyPDF2。")
        print("   请运行：pip install pypdf （推荐） 或 pip install PyPDF2")
        input("按任意键退出...")
        sys.exit(1)

# --- 📁 输入文件验证 ---
if not os.path.isfile(INPUT_PDF_PATH):
    print(f"❌ 错误：输入PDF文件不存在 → {INPUT_PDF_PATH}")
    input("按任意键退出...")
    sys.exit(1)

# --- 📄 加载PDF并分析 ---
try:
    reader = PdfReader(INPUT_PDF_PATH)
    total_pages = len(reader.pages)
    if total_pages == 0:
        raise ValueError("PDF文件为空（0页）")
except Exception as e:
    if "encrypted" in str(e).lower():
        print(f"⚠️  提示：PDF可能被密码保护。pypdf默认不支持解密。")
        print("   如需处理加密PDF，请改用 PyMuPDF（fitz）或 Spire.PDF 库。")
        input("按任意键退出...")
        sys.exit(1)
    else:
        print(f"❌ 读取PDF失败：{e}")
        input("按任意键退出...")
        sys.exit(1)

print(f"📄 检测到输入PDF共 {total_pages} 页")
print("🔍 提取规则解析：每4页为1组，取每组前2页 → 对应1-based页码：")
print("   组1 (1-4): 取 1,2  |  组2 (5-8): 取 5,6  |  组3 (9-12): 取 9,10  | ...")

# --- ⚙️ 核心逻辑：生成目标页码列表（1-based）---
target_pages_1based = []
for n in range(1, total_pages + 1):
    # 计算所属组号（从0开始）和组内偏移（0-3）
    group_idx = (n - 1) // 4
    offset_in_group = (n - 1) % 4  # 0,1,2,3
    if offset_in_group in (0, 1):  # 即第1页和第2页（1-based）
        target_pages_1based.append(n)

print(f"✅ 规则匹配页码（共{len(target_pages_1based)}页）：")
# 分行打印，每行最多10个数字，防刷屏
for i in range(0, len(target_pages_1based), 10):
    line = ", ".join(map(str, target_pages_1based[i:i+10]))
    print(f"   {line}")

# --- 📥 提取并合并 ---
writer = PdfWriter()
success_count = 0
for page_num_1based in target_pages_1based:
    try:
        # pypdf使用0-based索引
        page = reader.pages[page_num_1based - 1]
        writer.add_page(page)
        success_count += 1
    except Exception as e:
        print(f"⚠️  跳过页码 {page_num_1based}（异常：{type(e).__name__}）")

# --- 💾 保存结果 ---
output_path = Path(__file__).parent / OUTPUT_FILENAME
try:
    with open(output_path, "wb") as f:
        writer.write(f)
    print(f"\n🎉 提取合并成功！")
    print(f"   • 原文件：{os.path.basename(INPUT_PDF_PATH)}（{total_pages}页）")
    print(f"   • 提取页数：{len(target_pages_1based)}页（规则：每4页取前2页）")
    print(f"   • 输出文件：{output_path.resolve()}")
except Exception as e:
    print(f"\n❌ 保存失败：{e}")
    input("按任意键退出...")
    sys.exit(1)

# --- 🛑 调试暂停（关键！确保窗口不关闭）---
print("\n" + "="*70)
print("🔧 【调试模式】任务已完成。请检查：")
print("   • 上方页码列表是否符合您的预期（如：1,2,5,6,9,10,...）")
print("   • 输出PDF打开后页面顺序与内容是否正确")
print("   • 若需处理其他PDF，请修改 INPUT_PDF_PATH 后重新运行")
print("="*70)
input("\n👉 按任意键退出窗口...")