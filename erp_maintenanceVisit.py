from selenium import webdriver
from selenium.webdriver.common.by import By
import openpyxl as ox
from getdriver import *
import time

excel_file = ox.load_workbook('erp_all.xlsx')
excel_sheet = excel_file['ERP ALL']

station = 'L06'
start_date = '23-12-2023'
end_date = '31-12-2037'


# Capture the equipment name ID for the choosen station or distributor
station_cell_index = 'C'
equipment_list=[]
for station_row in range(3,1569):
    station_cell = station_cell_index  + str(station_row)
    station_cell_value = excel_sheet[station_cell].value
    if station_cell_value == station:
        first_row = station_row
        current_station_value = station_cell_value
        while current_station_value == station_cell_value:
            equipment_name_cell_index = 'D' + str(station_row)
            equipment_name_value = excel_sheet[equipment_name_cell_index].value
            equipment_list.append(equipment_name_value)
            station_row+=1
            current_station_index = station_cell_index  + str(station_row)
            current_station_value = excel_sheet[current_station_index].value

        break

getDriverfunc()

email = 'gadelhak.badr@madkour.com.eg'
passw = 'g0e9b7ssD@'
driver = webdriver.Chrome()

#Login
driver.get('https://erp.madkour.com.eg/app')

        
l = driver.find_element(By.XPATH,'//*[@id="page-app"]/div/main/div[2]/div/div/div/a')
l.click()

l = driver.find_element(By.XPATH,'//*[@id="login_email"]')
l.send_keys(email)


l = driver.find_element(By.XPATH,'//*[@id="login_password"]')
l.send_keys(passw)

l = driver.find_element(By.XPATH,'//*[@id="page-login"]/div/main/div[2]/div/section[1]/div[1]/form/div[2]/button')
l.click()
print('Logged In Successfully')


time.sleep(1)
print('Logging IN')

for equipment in equipment_list:
    driver.get(f'https://erp.madkour.com.eg/app/equipment-name/{equipment}')
    
    break

        

