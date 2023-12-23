from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import WebDriverException,NoSuchElementException,NoSuchWindowException,ElementClickInterceptedException
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import openpyxl as ox
from getdriver import *
import time
from selenium.webdriver.common.keys import Keys
import os

excel_file = ox.load_workbook('erp_all.xlsx')
excel_sheet = excel_file['ERP ALL']
ex = 0

station = 'I01'
start_date = '25-12-2023'
end_date = '31-12-2037'
periodicity = 'Quarterly'
speed_time = 0.3


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

# User configuration
email = 'gadelhak.badr@madkour.com.eg'
passw = 'g0e9b7ssD@'

user_profile_path = os.path.abspath('./user_profile/')
options = Options()
options.add_argument(f"user-data-dir={user_profile_path}")
driver = webdriver.Chrome(options=options)


#Login
while True:
    try:
        l = driver.find_element(By.XPATH,'/html/body/div[1]/header/div/div/ul/li[5]/a/span/span')
        break
    except:
        pass
    
    try:    
        driver.get('https://erp.madkour.com.eg/app')

        time.sleep(2)
        
        l = driver.find_element(By.XPATH,'//*[@id="page-app"]/div/main/div[2]/div/div/div/a')
        l.click()

        l = driver.find_element(By.XPATH,'//*[@id="login_email"]')
        l.send_keys(email)


        l = driver.find_element(By.XPATH,'//*[@id="login_password"]')
        l.send_keys(passw)

        l = driver.find_element(By.XPATH,'//*[@id="page-login"]/div/main/div[2]/div/section[1]/div[1]/form/div[2]/button')
        l.click()
        
        break
        
    except:

        time.sleep(1)
        print('Logging IN')

msg = 'Logged In Successfully'
print(msg)

while True:
    try:
        
        driver.find_element(By.XPATH,'//*[@id="navbar-search"]')
        msg = 'ERP is Ready'
        print(msg)
        break
    except:
        time.sleep(1)
        
num_of_items = len(equipment_list)

with open('Logging.txt','r') as lr:
    cont = lr.read()
    item_num = cont[-3:-1]
    
if item_num=='ed':
    item_num=0
else:
    item_num = int(item_num)+1

for equipment in equipment_list:
    
        
        equipment = equipment_list[item_num]
       
        try: 
            #Open page for equipment
            driver.get(f'https://erp.madkour.com.eg/app/equipment-name/{equipment}')
            time.sleep(2)
            while True:
                try:
                    driver.find_element(By.XPATH,'//*[@id="page-Equipment Name"]/div[2]/div[2]/div/div[3]/div[2]/div[1]/div[2]/div/div/div[2]/div[6]/div[2]/div/form/div/div/div[3]/div/div[1]/button[4]')
                    break
                except NoSuchElementException:
                    time.sleep(2)
                    msg = 'Waiting the Add row Button'
                    print(msg)

            time.sleep(2)
            #Click Add Row 
            elea = driver.find_element(By.XPATH,'//*[@id="page-Equipment Name"]/div[2]/div[2]/div/div[3]/div[2]/div[1]/div[2]/div/div/div[2]/div[6]/div[2]/div/form/div/div/div[3]/div/div[1]/button[4]')
            elea.click()
            
            time.sleep(1)
            
            #Scroll down to see the adding row event
            body = driver.find_element(By.TAG_NAME,'body')
            for i in range(1,7):
                body.send_keys(Keys.DOWN)
            
            # Waiting Start Date Field
            while True:
                try:
                    driver.find_element(By.XPATH,'//*[@id="page-Equipment Name"]/div[2]/div[2]/div/div[3]/div[2]/div[1]/div[2]/div/div/div[2]/div[6]/div[2]/div/form/div/div/div[2]/div/div[2]/div[1]/div[2]/div/div[3]')
                    break
                except NoSuchElementException:
                    time.sleep(1)
                    msg = 'Waiting the start date field'
                    print(msg)
                    elea.click()
            
            #Start Date
            ele = driver.find_element(By.XPATH,' //*[@id="page-Equipment Name"]/div[2]/div[2]/div/div[3]/div[2]/div[1]/div[2]/div/div/div[2]/div[6]/div[2]/div/form/div/div/div[2]/div/div[2]/div[1]/div[2]/div/div[3]')
            ele.click()
            time.sleep(speed_time)
            ele_input = driver.find_element(By.XPATH,'//*[@id="page-Equipment Name"]/div[2]/div[2]/div/div[3]/div[2]/div[1]/div[2]/div/div/div[2]/div[6]/div[2]/div/form/div/div/div[2]/div/div[2]/div[1]/div[2]/div/div[3]/div[1]/div/input')
            ele_input.send_keys(start_date)
            
            time.sleep(speed_time)
            
            # End Date
            ele = driver.find_element(By.XPATH,'//*[@id="page-Equipment Name"]/div[2]/div[2]/div/div[3]/div[2]/div[1]/div[2]/div/div/div[2]/div[6]/div[2]/div/form/div/div/div[2]/div/div[2]/div[1]/div[2]/div/div[4]/div[1]/div/input')
            ele.click()
            time.sleep(speed_time)
            ele.send_keys(end_date)
            
            time.sleep(speed_time)
            
            # Select periodicity
            ele = WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH,'//*[@id="page-Equipment Name"]/div[2]/div[2]/div/div[3]/div[2]/div[1]/div[2]/div/div/div[2]/div[6]/div[2]/div/form/div/div/div[2]/div/div[2]/div[1]/div[2]/div/div[5]/div[1]/div/select')))
            select = Select(ele)
            select.select_by_visible_text(periodicity)
            
            #Get Job Type 
            ele = driver.find_element(By.XPATH,'//*[@id="page-Equipment Name"]/div[2]/div[2]/div/div[3]/div[2]/div[1]/div[2]/div/div/div[2]/div[6]/div[2]/div/form/div/div/div[2]/div/div[2]/div[1]/div[1]/div/div[6]/div[2]/a')
            jop_type_text = ele.text

            # Enter Jop Type
            ele = driver.find_element(By.XPATH,'//*[@id="page-Equipment Name"]/div[2]/div[2]/div/div[3]/div[2]/div[1]/div[2]/div/div/div[2]/div[6]/div[2]/div/form/div/div/div[2]/div/div[2]/div[1]/div[2]/div/div[6]/div[1]/div/div/div/input')
            ele.click()
            time.sleep(speed_time)
            ele.send_keys(jop_type_text)
            
            time.sleep(speed_time)
            
            #Scroll up to see the update button
            body = driver.find_element(By.TAG_NAME,'body')
            body.send_keys(Keys.UP)
            
            time.sleep(1)
            
            #Click Update Button 
            ele = driver.find_element(By.XPATH,'//*[@id="page-Equipment Name"]/div[1]/div/div/div[2]/div[3]/button[2]')
            ele.click()
            
            time.sleep(2)
            
            for i in range(1,6):
                body.send_keys(Keys.DOWN)
            
            # Wait Edit Button
            while True:
                try:
                    driver.find_element(By.XPATH,'//*[@id="page-Equipment Name"]/div[2]/div[2]/div/div[3]/div[2]/div[1]/div[2]/div/div/div[2]/div[6]/div[2]/div/form/div/div/div[2]/div/div[2]/div[1]/div[2]/div[1]/div[8]/div')
                    break
                except NoSuchElementException:
                    time.sleep(1)
                    msg = 'Waiting the Edit Button'
                    print(msg)
                    

            time.sleep(1)
            
            # Click Edit Button
            ele = driver.find_element(By.XPATH,'//*[@id="page-Equipment Name"]/div[2]/div[2]/div/div[3]/div[2]/div[1]/div[2]/div/div/div[2]/div[6]/div[2]/div/form/div/div/div[2]/div/div[2]/div[1]/div[2]/div[1]/div[8]/div')
            while True:
                try:
                    ele.click()
                    break
                except ElementClickInterceptedException:
                    time.sleep(2)
                    msg = 'Trying again to press Edit'
                    print(msg)
                    
            
            time.sleep(1)
            
            for i in range(1,6):
                body.send_keys(Keys.DOWN)
            
            
            # Click Generate Schedule
            ele = driver.find_element(By.XPATH,'//*[@id="page-Equipment Name"]/div[2]/div[2]/div/div[3]/div[2]/div[1]/div[2]/div/div/div[2]/div[6]/div[2]/div/form/div/div/div[2]/div/div[2]/div[1]/div[2]/div[2]/div[2]/div[1]/div/div[2]/div[2]/div[2]/div/form/div/div/div[2]/div[1]/button')
            ele.click()
            
            time.sleep(2)
            
            c=True
            # Ensure it's generated successfully
            while True:
                try: 
                    driver.find_element(By.XPATH,'/html/body/div[7]/div/div')
                    time.sleep(2)
                    break
                except:
                    time.sleep(2)
                    msg = 'Waiting Page Final Message'
                    driver.current_url
                    print(msg)
                    if c:
                        try:
                            ele.click()
                            c=False
                        except ElementClickInterceptedException:
                            pass
                        
            try:
                ele = driver.find_element(By.XPATH,'/html/body/div[7]/div/div/div[2]/div[1]/div')
                if 'Schedule And Events Created successfully' in ele.text:
                    msg = 'Schedule And Events Created successfully'  
            except:
                msg = 'Something went wrong'
                
            fmsg = f'{item_num+1}/{num_of_items} , {msg} equipment {equipment} item {item_num}'
            print(fmsg)
            
            with open('Logging.txt','a') as l:
                l.write(fmsg+'\n')
            
            time.sleep(0.5)
            item_num+=1
            
        except NoSuchWindowException or WebDriverException:
            msg = 'Browser was Closed Unexpectedly'
            print(msg)
            ex = 1
            break
        except NoSuchElementException:
            msg = "I can't Locate My Next Step Please Make the browser Window bigger"
            print(msg)
            time.sleep(2)
            break
        
        if item_num == num_of_items:
            msg = 'Station Finished'
            print(msg)
            with open('Logging.txt','a') as l:
                l.write(f'Station {station} Finished'+'\n') 
            break