from selenium import webdriver
import win32com.client

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
xlsx_path = r"C:\Win32\Download Table.xlsx"
workbook = excel.Workbooks.Open(xlsx_path)


driver = webdriver.Chrome()
driver.get('https://www.fxempire.com/macro/vietnam')


titles = driver.find_elements('xpath', "//*[@class='FakeButton-sc-fclsxo-0 gUfUNa']")
title_num = len(titles)
print(title_num)

temp_sheet = workbook.Worksheets.Add()
while len(workbook.Worksheets) > 1:
    workbook.Worksheets(1).Delete()

for i in range(1,title_num+1):
    if i == 1:
        title = driver.find_element('xpath', f"//*[(@class='FakeButton-sc-fclsxo-0 gUfUNa')][1]")
        workbook.Worksheets(1).Name = title.text

    else:
        title = driver.find_element('xpath', f"//*[(@class='FakeButton-sc-fclsxo-0 gUfUNa')][{i}]")
        title.click()

    worksheet = workbook.Worksheets[title.text]

    column_num = len(driver.find_elements('xpath', "//*[@class='Card-sc-1ib64vn-0 fNBbuL']/div/table/thead/tr/th"))
    row_num = len(driver.find_elements('xpath', "//*[@class='Card-sc-1ib64vn-0 fNBbuL']/div/table/tbody/tr"))

    for header in range(1,column_num+1):
        worksheet.cells(1,header).Value = driver.find_element('xpath', f"//*[@class='Card-sc-1ib64vn-0 fNBbuL']/div/table/thead/tr/th[{header}]").text

    for row in range(1,row_num+1):
        for column in range(1, column_num+1):
            worksheet.cells(row+1, column).Value = driver.find_element('xpath',f"//*[@class='Card-sc-1ib64vn-0 fNBbuL']/div/table/tbody/tr[{row}]/td[{column}]").text

    if i != title_num:
        last_sheet = worksheet
        new_sheet = workbook.Worksheets.Add(Before=None,After=last_sheet)
        new_sheet.Name = driver.find_element('xpath', f"//*[(@class='FakeButton-sc-fclsxo-0 gUfUNa')][{i+1}]").text


workbook.Save()

#//*[@id='content']/div[2]/div[6]/div[1]/div[7]/div/div[2]/div/table/thead/tr/th