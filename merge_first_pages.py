# -*- coding: utf-8 -*-
"""
批量提取PDF首页并合并为一个PDF
来源目录：C:\\Users\\Administrator\\Desktop\\215
输出文件：当前目录下的 all_first_pages.pdf
执行完毕后暂停窗口（按任意键退出）
"""
import os
import sys
from pathlib import Path

# --- 🔧 配置区（请勿修改下方代码，仅调整此处）---
SOURCE_DIR = r"C:\Users\Administrator\Desktop\215"  # ← 源PDF所在文件夹（绝对路径，Windows反斜杠需加r前缀）
OUTPUT_FILENAME = "all_first_pages.pdf"             # ← 合并后输出的PDF文件名（保存在当前脚本所在目录）

# --- 📦 依赖检查与导入 ---
try:
    from pypdf import PdfReader, PdfWriter  # 优先尝试 pypdf（现代、维护活跃、API更清晰）
except ImportError:
    try:
        from PyPDF2 import PdfReader, PdfWriter  # 兜底尝试旧版 PyPDF2
    except ImportError:
        print("❌ 错误：未安装 pypdf 或 PyPDF2。")
        print("   请在命令行中运行：pip install pypdf")
        input("按任意键退出...")
        sys.exit(1)

# --- 📁 验证源目录 ---
if not os.path.isdir(SOURCE_DIR):
    print(f"❌ 错误：源目录不存在 → {SOURCE_DIR}")
    input("按任意键退出...")
    sys.exit(1)

pdf_files = [f for f in os.listdir(SOURCE_DIR) if f.lower().endswith(".pdf")]
if not pdf_files:
    print(f"⚠️  提示：目录 {SOURCE_DIR} 中未找到任何 PDF 文件。")
    input("按任意键退出...")
    sys.exit(0)

print(f"🔍 扫描到 {len(pdf_files)} 个 PDF 文件：")
for i, f in enumerate(pdf_files[:5], 1):  # 仅显示前5个，防刷屏
    print(f"  {i}. {f}")
if len(pdf_files) > 5:
    print(f"  ... 还有 {len(pdf_files)-5} 个文件")

# --- 📄 初始化合并器 ---
merger = PdfWriter()

# --- 📥 批量提取首页 ---
success_count = 0
error_list = []

for idx, filename in enumerate(pdf_files, 1):
    full_path = os.path.join(SOURCE_DIR, filename)
    try:
        reader = PdfReader(full_path)
        if len(reader.pages) == 0:
            raise ValueError("PDF 文件为空（0页）")
        
        first_page = reader.pages[0]
        merger.add_page(first_page)
        success_count += 1
        print(f"✅ [{idx}/{len(pdf_files)}] 已添加首页 → {filename}")

    except Exception as e:
        error_msg = f"[{idx}] {filename} → {str(e)[:80]}"
        error_list.append(error_msg)
        print(f"❌ [{idx}/{len(pdf_files)}] 跳过 → {filename}（{type(e).__name__}）")

# --- 💾 保存合并结果（输出到当前脚本所在目录）---
output_path = Path(__file__).parent / OUTPUT_FILENAME
try:
    with open(output_path, "wb") as f:
        merger.write(f)
    print(f"\n🎉 合并完成！共成功提取 {success_count}/{len(pdf_files)} 个首页。")
    print(f"📁 输出文件已保存至：{output_path.resolve()}")
except Exception as e:
    print(f"\n❌ 保存失败：{e}")
    input("按任意键退出...")
    sys.exit(1)

# --- 🛑 调试暂停（关键！确保窗口不关闭）---
print("\n" + "="*60)
print("🔧 调试模式：任务已完成。")
print("   • 检查上方是否有报错信息")
print("   • 查看输出PDF是否符合预期")
print("   • 如需重新运行，请修改脚本后双击或命令行执行")
print("="*60)
if error_list:
    print(f"\n⚠️  以下文件处理失败（共{len(error_list)}个）：")
    for err in error_list:
        print(f"   {err}")

input("\n👉 按任意键退出窗口...")  # 核心：阻塞等待用户按键，防止窗口关闭
