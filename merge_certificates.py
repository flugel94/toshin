import fitz  # PyMuPDF
from pathlib import Path

# ===== 結合対象のPDFがあるディレクトリ =====
pdf_dir = Path.home() / "Desktop" / "example" / "example" # 各々のパスに合わせてください。
output_pdf = pdf_dir / "修了書_まとめ.pdf"

# ===== PDFの読み込み・結合 =====
merged_doc = fitz.open()

# 「_修了書.pdf」で終わるファイルをソートして処理
pdf_files = sorted(pdf_dir.glob("*_修了書.pdf"))

for pdf_path in pdf_files:
    single_doc = fitz.open(pdf_path)
    merged_doc.insert_pdf(single_doc)
    single_doc.close()
    print(f"{pdf_path.name} を結合")

# ===== 結合後のPDF保存 =====
merged_doc.save(output_pdf)
merged_doc.close()

print(f"修了書を結合したPDFを出力しました: {output_pdf}")
