from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
import datetime
import pandas as pd
from datetime import datetime, timedelta
from openpyxl.workbook import Workbook
import re
import time

def scraping (url, driver_path, dpts, arrs, cartype_element, scrape_zoning):
    # date variable
    today = datetime.now()
    flight_date = datetime.today() + timedelta( days = 2)
    flight_month = flight_date.month
    flight_day = flight_date.day
    
    # get chrome driver ready
    chromedriver_path = driver_path
    service = Service(chromedriver_path)

    # open browser
    driver = webdriver.Chrome(service=service)
    driver.get(url)
    wait = WebDriverWait(driver,20)

    ## Setup search as history and back to re-search page for further loop (loop for different location search and getting data)
    # click pickup service
    click_service = wait.until(EC.visibility_of_element_located((By.XPATH, '//li[@class="is-tranform"]/div[@class="customtab-item fr-cc"]/span[contains(text(),"接机")]'))).click()
    # click arr textbox
    arr_txtbox_click = wait.until(EC.element_to_be_clickable((By.XPATH, '//div[@class="searchbox-item-placeholder fr-lc" and contains(text(), "降落的机场")]'))).click()
    click_dpt_txtbox = wait.until(EC.visibility_of_element_located((By.XPATH, '//div[@class="thanos-rpx is-bottom thanos-span__input"]/div[@class="input__content"]'))).click()
    dpt_txtbox = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="dpt_city"]/div[1]/div'))).click()

    flight_search_txtbox = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="city_input_pick"]')))
    try:
        flight_search_txtbox.clear()
    except:
        pass
    flight_search_txtbox.send_keys("上海") # key-in new search name
    select_dpt_city = wait.until(EC.element_to_be_clickable((By.XPATH, '//span[contains(text(),"PVG")]'))).click()
    # choose arr city & flight
    click_arr_flight = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="arr_city"]/div[2]/span[2]/span'))).click()
    text_box = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="city_input_drop"]')))
    text_box.clear()  # clear history search
    text_box.send_keys('曼谷素万那普机场')  # key-in new search name
    select_arv_city = wait.until(EC.element_to_be_clickable((By.XPATH, '//span[contains(text(), "BKK 曼谷")]'))).click()
    # choose flight schedule
    click_date_bottom = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="dpt_time"]/div/div/div'))).click()
    flight_month = wait.until(EC.element_to_be_clickable((By.XPATH,f'//div[@class="calendar__month--title fx-c" and contains(text(), "2023年{flight_month}月")]/../ul/li[@class="calendar__month--day fx-c" and contains (span[@class="day__date"], {flight_day})]'))).click()
    search_flight = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="flight_search"]'))).click()
    select_flight = wait.until(EC.visibility_of_element_located((By.XPATH, "(//div[@class='thanos-rpx is-bottom'])[1]"))).click()
    # setup domestic destination for PICKUP CAR
    select_dos_dest = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="pick_address"]/div'))).click()
    text_box = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="arrive_right_address"]')))
    text_box.clear()  # clear history search
    text_box.send_keys('W Hotel')  # key-in domestic destination
    select_destination = wait.until(EC.element_to_be_clickable((By.XPATH, '//li[@class="touch-block addr__item"][@data-idx="0"]'))).click()
    # Click search
    click_serch_bottom = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="__next"]/div/div/div[2]/div/div/div[1]/div[2]/div/div[2]/div/div[2]/div/div[6]/div'))).click()
    back = wait.until(EC.element_to_be_clickable((By.XPATH, '//div[@id="__next"]/div/div/div[1]/div[1]/a/i'))).click()

    for dpt,arr in zip(dpts,arrs):
        ## change arrival airport
        change_airport = wait.until(EC.element_to_be_clickable((By.XPATH, '//div[@id="pick_flight"]//div[@class="searchbox-item-value fr-lc"]'))).click()
        click_airport = wait.until(EC.element_to_be_clickable((By.XPATH, '//div[@id="arr_city"]'))).click()
        text_box = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="city_input_drop"]')))
        text_box.clear()  # clear history search
        text_box.send_keys(dpt)  # key-in new search name
        choose_airport = wait.until(EC.element_to_be_clickable((By.XPATH, f'//ul[@class="result__list"]//div[@class="touch-light item-bar-wrapper thanos-rpx is-bottom fr-lt"]/span'))).click()
        search_flight = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="flight_search"]'))).click()
        select_flight = wait.until(EC.visibility_of_element_located((By.XPATH, "(//div[@class='thanos-rpx is-bottom'])[1]"))).click()
        
        ## change destination
        click_dpt_box = wait.until(EC.element_to_be_clickable((By.XPATH, '//div[@id="pick_address"]/div'))).click()
        input_box = wait.until(EC.element_to_be_clickable((By.XPATH, '//div["thanos-input"]/input[@id="arrive_right_address"]')))
        input_box.clear()
        input_box.send_keys(arr)
        time.sleep(1)
        select_destination = wait.until(EC.visibility_of_element_located((By.XPATH, '//li[@class="touch-block addr__item"][@data-idx="0"][1]'))).click()
        time.sleep(1)
        click_serch_bottom = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="__next"]/div/div/div[2]/div/div/div[1]/div[2]/div/div[2]/div/div[2]/div/div[6]/div'))).click()
        
        for cartype, zoning in zip(cartype_element, scrape_zoning):
            car_typeID = wait.until(EC.visibility_of_element_located((By.XPATH, f'//div[@id= {cartype}]/div[1]'))).click()

            ## Get all elements contain in place for further extraction data by for loop
            cars = wait.until(EC.visibility_of_element_located((By.XPATH, f'//div[@grpidx="{zoning}"]//div[@class="sticky-header"]//span[@class="listv2-car-name"]')))
            names = wait.until(EC.visibility_of_all_elements_located((By.XPATH, f'//div[@grpidx="{zoning}"]//div[@class="listv2-prds-list"]//span[@class="prdv2-vnd-name "]')))
            prices = wait.until(EC.visibility_of_all_elements_located((By.XPATH, f'//div[@grpidx="{zoning}"]//div[@class="listv2-prds-list"]//p[@class="listv2-opriceShow"]')))
            dpts = wait.until(EC.visibility_of_element_located((By.XPATH,'//div[@class="line1txt _listJourney-infotxt"][1]')))
            arrs = wait.until(EC.visibility_of_element_located((By.XPATH,'//div[@class="line1txt _listJourney-infotxt"][2]')))
            
            ## loop to get each text from elements
            for i,e in zip(names,prices) :
                competitor_names.append(i.text)
                competitor_prices.append(int(''.join(re.findall(r'\d',e.text))))
                car_types.append(cars.text)
                departure.append(dpts.text)
                arrival.append(arrs.text)
                
        clk_12s_van = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="__next"]/div/div/div[2]/div/div[3]/div/div/div/div[6]/div/div[1]/div[3]/div[2]/div'))).click()
        cars = wait.until(EC.visibility_of_element_located((By.XPATH, '//div[@class="sticky-header"][@grpid="127"]//span[@class="listv2-car-name"]')))
        names = wait.until(EC.visibility_of_all_elements_located((By.XPATH, '//div[@grpidx="5"]//div[@class="listv2-prds-list"]//span[@class="prdv2-vnd-name "]')))
        prices = wait.until(EC.visibility_of_all_elements_located((By.XPATH, '//div[@grpidx="5"]//div[@class="listv2-prds-list"]//p[@class="listv2-opriceShow"]')))        
        dpts = wait.until(EC.visibility_of_element_located((By.XPATH,'//div[@class="line1txt _listJourney-infotxt"][1]')))
        arrs = wait.until(EC.visibility_of_element_located((By.XPATH,'//div[@class="line1txt _listJourney-infotxt"][2]')))

        for a,e in zip(names,prices) :
            competitor_names.append(a.text)
            competitor_prices.append(int(''.join(re.findall(r'\d',e.text))))
            car_types.append(cars.text)
            departure.append(dpts.text)
            arrival.append(arrs.text)

            ## get back to search page for new location search
        back = wait.until(EC.element_to_be_clickable((By.XPATH, '//div[@id="__next"]/div/div/div[1]/div[1]/a/i'))).click()
    return competitor_names, competitor_prices, car_types, departure, arrival

# ----------------------
#    Variable Setup
# ----------------------
url = ('https://m.ctrip.com/webapp/zhuanche/airport-transfers/index?s=car&ptgroup=17&biztype=32&channelid=90189')
driver_path = r'T:\Python_project\main\chrome-win32\chrome-win32\chorme.exe'
search_df = pd.read_excel('search_lists.xlsx')
dpts = search_df['departure']
arrs = search_df['destination']
cartype_element = ["1170", "1181", "1192", "200203", "1204", "05"]
scrape_zoning = ["0","1","2","3","4","5"]

today = datetime.now()
flight_date = datetime.today() + timedelta( days = 2)
flight_month = flight_date.month
flight_day = flight_date.day
competitor_names = []
competitor_prices = []
car_types = []
departure = []
arrival = []

# Call for function [scraping]
scraping (url, driver_path, dpts, arrs, cartype_element, scrape_zoning)

# Combine all data in dataFrame
df = pd.DataFrame({
    '类别':'airport_pickup',
    '区域':departure,
    '目的地':arrival,
    '车型':car_types,
    '竟对':competitor_names, 
    '价格':competitor_prices,
    '记录时间':today
})

# file name 
file_name = today.strftime("%Y-%m-%d_%H-%M") + '.xlsx'

# export file
df.to_excel(f'T:\OneDrive\Desktop\Price Adjustment\competitor_log\{file_name}',index=False)

