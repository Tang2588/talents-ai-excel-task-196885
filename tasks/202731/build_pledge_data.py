"""生成题目 202731 的附件：股权质押数据表.xlsx"""

from pathlib import Path

import xlsxwriter

OUT = Path(__file__).resolve().parent / "股权质押数据表.xlsx"

# 股东, 上市公司, 质押股份数(万股), 融资额(万元), 质押日股价, 最新股价,
# 预警线, 平仓线, 股东持股数(万股), 公司总股本(万股)
ROWS = [
    ("华曜集团", "华曜科技", 8000, 80000, 13.80, 12.50, 1.50, 1.30, 15000, 40000),
    ("云帆投资", "云帆软件", 6000, 54000, 12.60, 11.20, 1.50, 1.30, 6600, 30000),
    ("江海控股", "江海银行", 12000, 96000, 11.60, 11.00, 1.50, 1.30, 20000, 60000),
    ("安信实业", "安信证券", 5000, 40000, 11.40, 10.50, 1.50, 1.30, 6000, 25000),
    ("远岭医药集团", "远岭医药", 9000, 72000, 13.60, 13.00, 1.50, 1.30, 16000, 40000),
    ("康泽生物控股", "康泽生物", 4000, 36000, 12.80, 12.00, 1.50, 1.30, 5000, 20000),
    ("东辰消费集团", "东辰消费", 7000, 84000, 16.20, 15.00, 1.50, 1.30, 12000, 30000),
    ("新禾食品控股", "新禾食品", 5500, 44000, 10.80, 10.00, 1.50, 1.30, 10000, 25000),
    ("瀚源能源集团", "瀚源能源", 10000, 80000, 14.80, 14.00, 1.50, 1.30, 20000, 50000),
    ("启明电力控股", "启明电力", 6500, 52000, 10.20, 9.60, 1.50, 1.30, 7000, 30000),
    ("联港工业集团", "联港工业", 7500, 69000, 13.10, 12.40, 1.50, 1.30, 9000, 14000),
    ("锐虎机械控股", "锐虎机械", 3000, 24000, 16.80, 16.00, 1.50, 1.30, 6000, 20000),
]


def main() -> None:
    book = xlsxwriter.Workbook(str(OUT))
    header = book.add_format({"bold": True, "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True})
    text = book.add_format({"border": 1})
    number = book.add_format({"border": 1, "num_format": "0.00"})
    integer = book.add_format({"border": 1, "num_format": "#,##0"})
    percent = book.add_format({"border": 1, "num_format": "0.0%"})

    sheet = book.add_worksheet("股权质押明细")
    sheet.hide_gridlines(2)
    columns = [
        "股东名称", "上市公司", "质押股份数(万股)", "融资额(万元)", "质押日股价(元)",
        "最新股价(元)", "预警线(%)", "平仓线(%)", "股东持股数(万股)", "公司总股本(万股)",
    ]
    sheet.write_row(0, 0, columns, header)
    for idx, row in enumerate(ROWS, start=1):
        sheet.write(idx, 0, row[0], text)
        sheet.write(idx, 1, row[1], text)
        sheet.write_number(idx, 2, row[2], integer)
        sheet.write_number(idx, 3, row[3], integer)
        sheet.write_number(idx, 4, row[4], number)
        sheet.write_number(idx, 5, row[5], number)
        sheet.write_number(idx, 6, row[6], percent)
        sheet.write_number(idx, 7, row[7], percent)
        sheet.write_number(idx, 8, row[8], integer)
        sheet.write_number(idx, 9, row[9], integer)

    widths = [16, 12, 16, 13, 14, 12, 11, 11, 16, 16]
    for col, width in enumerate(widths):
        sheet.set_column(col, col, width)
    sheet.set_row(0, 30)

    notice = book.add_worksheet("说明")
    notice.write(0, 0, "本表为模拟数据，仅用于题目作答。", text)
    notice.set_column(0, 0, 40)
    book.close()
    print(OUT)


if __name__ == "__main__":
    main()
