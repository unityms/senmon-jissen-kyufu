"""ジョブ・カード等の Excel 様式にセル値を書き込み、LibreOffice で PDF 化する。

使い方:
  python3 xlsx_fill.py 様式.xlsx 値.json 出力.xlsx [--pdf]

値.json の形式（シート名省略時は先頭シート）:
  {"C2": "YAMADA", "F2": "TARO", "B14": "複数行は\\n で改行"}
  {"様式2（Ver.1.23）": {"C3": "YAMADA"}}
値に null を指定するとセルを空にする（旧申請者の値を消すとき）。
"""
import json
import subprocess
import sys
from pathlib import Path

import openpyxl


def fill(src, values, dst):
    wb = openpyxl.load_workbook(src)
    if values and all(isinstance(v, dict) for v in values.values()):
        per_sheet = values
    else:
        per_sheet = {wb.worksheets[0].title: values}
    for sheet, cells in per_sheet.items():
        ws = wb[sheet]
        for coord, v in cells.items():
            ws[coord].value = v
    wb.save(dst)


def to_pdf(xlsx):
    out = Path(xlsx).resolve().parent
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(out), str(xlsx)],
                   check=True, capture_output=True, timeout=300)
    return out / (Path(xlsx).stem + ".pdf")


if __name__ == "__main__":
    src, vals, dst = sys.argv[1:4]
    fill(src, json.loads(Path(vals).read_text(encoding="utf-8")), dst)
    if "--pdf" in sys.argv[4:]:
        print(to_pdf(dst))
