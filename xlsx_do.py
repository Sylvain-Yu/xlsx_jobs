from openpyxl import load_workbook
wb = load_workbook("./MAGELEC_TestData/1000rpm_290V_Mot.xlsx")

current_sheet = wb["Instantly Data"]

for i in range(2,30):
    data = current_sheet["E"+str(i)].value
    print(data)