import fitz  # PyMuPDF
import pandas as pd
from pathlib import Path

# ===== パスの設定 =====
base_dir = Path.home() / "Desktop" / "example" / "example" # 各々のパスに合わせてください。
template_path = base_dir / "修了証.pdf"
csv_path = base_dir / "名簿.csv"
output_dir = base_dir / "修了証"
output_dir.mkdir(exist_ok=True)

# ===== CSVの読み込み =====
df = pd.read_csv(csv_path)

# ===== 各行に対してループ処理 =====
for idx, row in df.iterrows():
    name = str(row.iloc[1])           # B列の値（名前として使う）
    file_prefix = name                # 出力ファイル名の先頭にも使う

    # PDF読み込み
    doc = fitz.open(template_path)
    if len(doc) == 0:
        raise ValueError("修了証.pdf にページが含まれていません。中身を確認してください。")

    page = doc[0]

    # 名前を挿入（24pt、Times-Roman）
    page.insert_text(
        (105, 270),
        name,
        fontsize=28,
        fontname="Times-Roman",
        color=(0, 0, 0),
    )

    # 保存
    output_path = output_dir / f"{file_prefix}_修了書.pdf"
    doc.save(output_path)
    doc.close()

    print(f"{file_prefix}_修了書.pdf を出力しました。")

print("全員分の修了書が完成しました！")
