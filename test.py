import os
import pdfplumber
import re

filepath = "/Users/65148/Desktop/汇总/"
filenames = os.listdir(filepath)
result = []

for file in filenames:
    pdf = pdfplumber.open(filepath + file)
    first_page = pdf.pages[0]
    items = first_page.extract_table()

    if items is not None:  # 检查是否提取到了表格
        for item in items:
            if item[0] is None:
                continue
            if '合计' in item[0]:
                for i in item:
                    if i is None:
                        continue
                    if '小写' in i:
                        result.append(i)

# 最终金额产出
final_value = 0.0
for string in result:
    final_value += float(re.findall(r"\d+\.?\d*", string)[-1])

print("最终金额产出:", final_value)
