import time
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
import random

# 创建一个Workbook对象
wb = Workbook()
# 激活默认的工作表
ws = wb.active

# 写入表头
tableheaders = [
    "货币名称",
    "现汇买入价",
    "现钞买入价",
    "现汇卖出价",
    "现钞卖出价",
    "中行折算价",
    "发布时间",
]
ws.append(tableheaders)

url = "https://srh.bankofchina.com/search/whpj/search_cn.jsp"

params = {"erectDate": "2024-05-01", "nothing": "2024-05-31", "pjname": "美元"}

for page in range(1, 360):
    params["page"] = str(page)
    print(params)

    # 创建一个会话对象
    session = requests.Session()

    # 随机生成一个Cookie
    cookie_value = f"JSESSIONID={random.randint(100000000, 999999999)}:-1"
    mheaders = {
        "Content-Type": "application/x-www-form-urlencoded",
        "Cookie": cookie_value,
    }

    response = session.post(url, headers=mheaders, data=params)

    # 使用Beautiful Soup解析HTML
    soup = BeautifulSoup(response.text, "html.parser")
    # 找到表格
    table = soup.find("div", class_="BOC_main publish")
    # 如果找到表格，提取表格数据
    if table:
        # 找到表头
        header_row = table.find("tr")
        headers = [header.text.strip() for header in header_row.find_all("th")]
        data_rows = []
        for row in table.find_all("tr")[1:]:
            cells = [cell.text.strip() for cell in row.find_all("td")]
            data_rows.append(cells)
        # 重新请求处理
        while not any(data_rows):  # 重试
            print(f"Retrying page {page}...")
            time.sleep(2)
            response = session.post(url, headers=mheaders, data=params)
            soup = BeautifulSoup(response.text, "html.parser")
            table = soup.find("div", class_="BOC_main publish")
            data_rows = []
            for row in table.find_all("tr")[1:]:
                cells = [cell.text.strip() for cell in row.find_all("td")]
                data_rows.append(cells)

        # 将数据行添加到工作表中
        for row_data in data_rows:
            ws.append(row_data)

# 保存Excel文件
wb.save("exchange_rates.xlsx")
