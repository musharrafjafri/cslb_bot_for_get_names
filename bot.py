import xlrd  # xlrd use to read excel files.
from xlwt import Workbook  # xlwt use to write into excel's workbook and sheets.
from time import sleep  # sleep/wait is use to apply wait to load all the attributes in browser.
from selenium import webdriver  # It is out main library to make automation through python language.
from selenium.webdriver.support.select import Select

print('                       BOT STARTED...')
city_low = []  # We are getting city name from our input_data.xlsx in lower or whatever case.
lic_type = []  # We are getting license types from our input_data.xlsx file.
city = []  # now here we will convert are city list to upper case.
d_path = ''  # Here we're getting res_files folder's path.
with open('res_files_path.txt') as file:
    d_path = file.read()  # here its reading path and parsing to d_path variable.

workbook = xlrd.open_workbook('input_data/input_data.xlsx')  # Opoen excel workbook to read in put_data.xlsx
sheet = workbook.sheet_by_index(0)  # Getting sheet's number.

for row in range(sheet.nrows):  # Now we are getting rows from sheet.
    city_low.append(sheet.cell_value(row, 0))  # parsing city data from input_data.xlsx file.
    lic_type.append(sheet.cell_value(row, 1))  # parsing license type data from input_data.xlsx file.

for citi in city_low:
    city.append(citi.upper())  # Here we converting city list to upper case.

wrong_city = []
for num in range(len(city)):
    options = webdriver.ChromeOptions()
    # options.add_argument("--headless")
    options.add_argument('start-maximized')
    preferances = {"download.default_directory": d_path}  # Parsing res_files folder's path to download all files here.
    options.add_experimental_option("prefs", preferances)
    options.add_experimental_option("useAutomationExtension", False)
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    driver = webdriver.Chrome(options=options)
    driver.maximize_window()  # to open chrome in maximaze mode.
    driver.get("https://cslb.ca.gov/OnlineServices/CheckLicenseII/ZipCodeSearch.aspx")  # Giving url to open in chrome browser.
    sleep(1)  # wait 1 second to load page.


    try:
        driver.find_element_by_name('ctl00$MainContent$txtCity').send_keys(city[num])  # Entering city name on site's page.
        selection = Select(driver.find_element_by_name('ctl00$MainContent$ddlLicenseType'))  # Getting all options of license type from list.
        selection.select_by_visible_text(lic_type[num])  # Selecting our desired license type from option's list
        driver.find_element_by_name('ctl00$MainContent$btnZipCodeSearch').click()  # Click on selected license type.
        sleep(1)  # wait 1 second to load page.
        driver.find_element_by_name('ctl00$MainContent$ibExportToExcell').click()  # Click on button to go to result.
        sleep(5)  # wait 5 second to wait page until file download.

        des_city = city[num]  # Getting our desired city name one by one with every loop.
        cur_lic_type = lic_type[num]  # Getting our desired license type one by one with every loop.
        location = 'res_files/Contractor List.xlsx'  # Getiing files from res_files folder to get license numbers.
        workbook = xlrd.open_workbook(location)  # Opening workbook by location.
        sheet = workbook.sheet_by_index(0)  # getting sheet from excel workbook.
        file_city = sheet.cell_value(8, 2)  # getting sheet's city name to verify that we are getting right file to extract or not.
        p_num = 0  # It will use if our desired matched file is not on current location so we will push location one ahead.

        try:
            while des_city != file_city:  # comparing that our desired city is matched with the current get file. If not than run this loop again and again.
                location = f"res_files/Contractor List ({p_num + 1}).xlsx"
                workbook = xlrd.open_workbook(location)
                sheet = workbook.sheet_by_index(0)
                file_city = sheet.cell_value(8, 2)
                p_num += 1
        except:
            pass

        lic = []  # It will use to store all license number from next for loop.
        for row in range(sheet.nrows):  # Getting all the license numbers from donwloaded file.
            if row > 7:
                lic.append(sheet.cell_value(row, 5))

        print('length of license list: ', len(lic))
        wb = Workbook()  # Open new workbook to save Names against license number.
        sheet1 = wb.add_sheet('Sheet 1')
        sheet1.write(0, 0, 'LICENSE NUMBER')
        sheet1.write(0, 1, 'NAME')

        for i in range(len(lic)):  # run loop to get license number's Personal name one by one.
            print(f'{i+1} out of {len(lic)} - {des_city}')
            try:
                url = f'https://cslb.ca.gov/OnlineServices/CheckLicenseII/LicenseDetail.aspx?LicNum={int(lic[i])}'
                driver.get(url)
                # sleep(1)
                driver.find_element_by_name('ctl00$MainContent$PersonnelLink').click()
                # sleep(1)
                sheet1.write(i+1, 0, lic[i])
                sheet1.write(i+1, 1, driver.find_element_by_id('MainContent_dlAssociated_hlName_0').text)
            except:  # If there is some issue on license number's page so NOT FOUND will written agianst that license number.
                print('Except block...')
                sheet1.write(i+1, 0, lic[i])
                sheet1.write(i+1, 1, 'NOT FOUND')

        wb.save(f'output_data/{des_city}-{cur_lic_type}-Names.xls')  # Finally we're saving out file.
    except:
        driver.close()
        print(city[num], ' with ', lic_type[num], 'is wrong.')
        wrong_city.append(city[num])
        # wrong_lic_type.append(lic_type[num])

    with open('wrong_input/wrong_inputs.txt', 'w') as w_in:
            for i in range(len(wrong_city)):
                w_in.write(f'{wrong_city[i]}, ')

    driver.close()  # Closing browser to go to next city and license type.
print('                       Files generated successfully...')