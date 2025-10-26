import pandas as pd

# Create a test CSV file that mimics the problematic format
test_content = '''订单号,外部订单号,订单金额,其他列
12345678901234567890,ProductP123,100,A
12345678901234567891,ProductP456,-50,B
"12345,678901234567892",ProductP789,200,C
11111111111111111111,ProductP000,-30,D
'''

with open('ExcelForHandel/problematic.csv', 'w', encoding='utf-8') as f:
    f.write(test_content)

print("Created problematic CSV file")