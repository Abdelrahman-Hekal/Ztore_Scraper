from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait as wait
from selenium.webdriver.chrome.service import Service as ChromeService 
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.common.action_chains import ActionChains
import undetected_chromedriver as uc
import time
import os
from datetime import datetime
import pandas as pd
import warnings
import re
import sys
import shutil
import xlsxwriter
warnings.filterwarnings('ignore')

def initialize_bot():
    
    print('Initializing the web driver ...')
    # Setting up chrome driver for the bot
    chrome_options = uc.ChromeOptions()
    chrome_options.add_argument('--log-level=3')
    chrome_options.add_argument('--headless')
    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
    # installing the chrome driver
    driver_path = ChromeDriverManager().install()
    chrome_service = ChromeService(driver_path)
    # configuring the driver
    driver = webdriver.Chrome(options=chrome_options, service=chrome_service)
    ver = int(driver.capabilities['chrome']['chromedriverVersion'].split('.')[0])
    driver.quit()
    chrome_options = uc.ChromeOptions()
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36")
    chrome_options.add_argument('--log-level=3')
    chrome_options.add_argument("--enable-javascript")
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--incognito")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--lang=en")
    chrome_options.add_argument('--headless=new')
    chrome_options.page_load_strategy = 'normal'
    prefs = {"profile.default_content_setting_values.geolocation": 2, "profile.managed_default_content_settings.images": 2, "profile.managed_default_content_settings.notifications": 2}  
    chrome_options.add_experimental_option("prefs", prefs)
    driver = uc.Chrome(version_main = ver, options=chrome_options) 
    driver.set_window_size(1920, 1080)
    driver.maximize_window()
    driver.set_page_load_timeout(10000)

    return driver

def process_links(driver, links, settings):

    print('-'*100)
    print('Processing links before scraping')
    print('-'*100)
    df = pd.DataFrame()
    prods_limit = settings['Product Limit']
    n = len(links)
    for i, link in enumerate(links):

        print(f'Processing input link {i+1}/{n}...')
        # single product link
        if '/product/' in link:
            df = df.append([{'Link': link}])
            continue

        driver.get(link)
        time.sleep(3)
        nprods = 0   
        try:
            # clicking on view all products if available
            try:
                button = wait(driver, 4).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[class='viewAllButton']")))
                driver.execute_script("arguments[0].click();", button)
                time.sleep(2)
            except:
                pass

            # handling lazy loading
            while True:  
                try:
                    height1 = driver.execute_script("return document.body.scrollHeight")
                    driver.execute_script(f"window.scrollTo(0, {height1})")
                    time.sleep(3)
                    height2 = driver.execute_script("return document.body.scrollHeight")
                    if height1 == height2: 
                        break
                    if prods_limit > 0:
                        prods = wait(driver, 5).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.ProductItem")))
                        if len(prods) > prods_limit:
                            break
                except Exception as err:
                    break
            
            prods = wait(driver, 5).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.ProductItem")))
            for prod in prods:
                try:
                    url = wait(prod, 5).until(EC.presence_of_element_located((By.TAG_NAME, "a"))).get_attribute('href')
                    if '/product/' in url:
                        df = df.append([{'Link': url}])
                        nprods += 1
                    if nprods == prods_limit:
                        break
                except:
                    pass

        except Exception as err:
            print(f"The below error occurred while scraping the products urls under link {i+1} \n")
            print(err)
            print('-'*100)
            driver.quit()
            time.sleep(5)
            driver = initialize_bot()
            driver.get(link)
            time.sleep(3)
            
    # return products links
    df.drop_duplicates(inplace=True)
    prod_links = df['Link'].values.tolist()
    return prod_links


def scrape_prods(driver, prod_links, output1, output2, settings):

    keys = ["Product ID","Product URL",	"Product Title","Product Price","Product Origin","Product Category","Product Description","Product Delivery","Product Rating","Product Image","Return Info","Store Name","Store Rating","Sold"]

    print('-'*100)
    print('Scraping links')
    print('-'*100)

    stamp = datetime.now().strftime("%d_%m_%Y")
    # reading scraped links for skipping
    scraped = []
    try:
        df = pd.read_excel(output1)
        scraped = df['Product URL'].values.tolist()
    except:
        pass

    prods = pd.DataFrame()
    comments = pd.DataFrame()
    nlinks = len(prod_links)  
    for i, link in enumerate(prod_links):

        if link in scraped: continue
        prod = {}
        for key in keys:
            prod[key] = ''

        print(f'Scraping the details from link {i+1}/{nlinks} ...')
        driver.get(link)
        time.sleep(1)

        # handling 404 error
        try:
            wait(driver, 3).until(EC.presence_of_element_located((By.XPATH, "div.NotFound")))  
            print(f'Error 404 in link: {link}')
            continue
        except:
            pass
           
        # scrolling across the page 
        try:
            htmlelement= wait(driver, 3).until(EC.presence_of_element_located((By.TAG_NAME, "html")))
            total_height = driver.execute_script("return document.body.scrollHeight")
            height = total_height/30
            new_height = 0
            for _ in range(30):
                prev_hight = new_height
                new_height += height             
                driver.execute_script(f"window.scrollTo({prev_hight}, {new_height})")
                time.sleep(0.1)
        except:
            pass

        try:
            # scraping product URL
            prod['Product URL'] = driver.current_url

            # scraping product ID
            try:            
                try:
                    img_div = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.product-image-wrapper")))
                    url = wait(img_div, 2).until(EC.presence_of_element_located((By.TAG_NAME, "img"))).get_attribute("src")
                except:
                    url = ''
                nums = re.findall(r'[0-9]+', link)
                if len(nums) == 1:
                    ID = nums[0]
                else:
                    found = False
                    for num in nums:
                        if len(num) < 6: continue
                        if num in url:
                            found = True
                            ID = num
                            break

                    if not found:
                        for num in nums:
                            if len(num) > 5:
                                ID = num
                                break

                prod['Product ID'] = ID    
            except:
                prod['Product ID'] = ''

            # scraping product title
            try:
                title_div = wait(driver, 3).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.name-rating-container")))
                title = wait(title_div, 3).until(EC.presence_of_element_located((By.TAG_NAME, "h2"))).get_attribute('textContent')
                prod['Product Title'] = title
                brand = wait(driver, 3).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.brand"))).get_attribute('textContent').strip()
                prod['Product Title'] = brand + ' - ' + title
            except:
                prod['Product Title'] = ''
                
            # scraping product price
            try:
                price_div = wait(driver, 3).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.price")))
                try:
                    price = wait(price_div, 3).until(EC.presence_of_element_located((By.TAG_NAME, "span"))).get_attribute('textContent').replace('$', '').replace(',', '').strip()
                except:
                    price = wait(price_div, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "span.original"))).get_attribute('textContent').replace('$', '').replace(',', '').strip()

                prod['Product Price'] = price
            except:
                prod['Product Price'] = ''

            # scraping product origion
            try:
                origin = wait(driver, 3).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.info-row-country"))).get_attribute('textContent').strip()
                prod['Product Origin'] = origin       
            except:
                prod['Product Origin'] = ''
            
            # scraping product delivery
            try:
                ships_div = wait(driver, 3).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.shippings")))[-1]
                methods = wait(ships_div, 3).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.shipping")))
                delivery = ''
                for method in methods:
                    name = wait(method, 3).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.name"))).get_attribute('textContent')
                    desc = wait(method, 3).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.desc"))).get_attribute('textContent')
                    delivery += name +' : '+ desc +'\n'
                    
                prod['Product Delivery'] = delivery.strip("\n")
            except:
                prod['Product Delivery'] = ''
                
            # scraping product description
            try:
                des = wait(driver, 1).until(EC.presence_of_element_located((By.CSS_SELECTOR, "section.ProductDetailSection"))).get_attribute('textContent').replace('\n\n', '\n').replace('Product Details', '') 
                prod['Product Description'] = des
            except:
                prod['Product Description'] = ''

            # scraping product category
            try:
                cat_div = wait(driver, 3).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.DropDownList")))
                cat = wait(cat_div, 3).until(EC.presence_of_element_located((By.TAG_NAME, "input"))).get_attribute("value")
                prod['Product Category'] = cat
            except:
                prod['Product Category'] = ''

            # scraping product rating
            try:
                rating = wait(driver, 3).until(EC.presence_of_element_located((By.CSS_SELECTOR, "span.rating"))).get_attribute('textContent')
                rating = float(rating)
                if rating > 0:
                    prod['Product Rating'] = rating
                else:
                    prod['Product Rating'] = ''
            except:
                prod['Product Rating'] = ''

            # scraping product image link
            try:
                img_div = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.product-image-wrapper")))
                url = wait(img_div, 2).until(EC.presence_of_element_located((By.TAG_NAME, "img"))).get_attribute("src")
                if url[:6].lower() == 'https:':
                    prod['Product Image'] = url
                else:
                    prod['Product Image'] = 'https:' + url
            except:
                prod['Product Image'] = ''

            prod['Store Name'], prod['Store Rating'], prod['Return Info'] , prod['Sold']= '', '', '', ''
            prod['Extraction Date'] = stamp
               
            # scraping product comments
            if settings['Scrape Comments'] != 0 and prod['Product ID'] != '' and prod['Product Title'] != '' and prod['Product Price'] != '':
                revs_limit = settings['Comment Limit']
                try:
                    rev_div = wait(driver, 3).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.ProductReview")))
                    revs = wait(rev_div, 3).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.Review")))

                    # applying the comments limit
                    nrevs = len(revs)
                    if nrevs > revs_limit:
                        nrevs = revs_limit

                    for k in range(nrevs):
                        try:
                            comm = {}
                            try:
                                rev = wait(revs[k], 3).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.review"))).get_attribute('textContent')
                            except:
                                rev = ''
                            try:
                                date = wait(revs[k], 3).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.date"))).get_attribute('textContent')
                            except:
                                date = ''

                            try:
                                stars_div = wait(revs[k], 3).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.full")))
                                stars = wait(stars_div, 3).until(EC.presence_of_all_elements_located((By.TAG_NAME, "span")))
                                rating = len(stars)
                            except:
                                rating = ''

                            comm['Product ID'] = ID
                            comm['Comment Content'] = rev
                            comm['Comment Rating'] = rating
                            comm['Comment Date'] = date
                            comm['Extraction Date'] = stamp
                            comments = comments.append([comm.copy()]) 
                        except:
                            pass
                except:
                    # No product reviews are available
                    pass

            # checking if the produc data has been scraped successfully
            if prod['Product ID'] != '' and prod['Product Title'] != '' and prod['Product Price'] != '':
                # output scraped data
                prods = prods.append([prod.copy()])
                
        except Exception as err:
            print(f'The error below ocurred during scraping link {i+1}/{nlinks}, skipping ...\n') 
            print(err)
            print('-'*100)
            continue 
        
    # output data
    if prods.shape[0] > 0:
        prods['Extraction Date'] = pd.to_datetime(prods['Extraction Date'], errors='coerce', format="%d_%m_%Y")
        prods['Extraction Date'] = prods['Extraction Date'].dt.date   
        writer = pd.ExcelWriter(output1, date_format='d/m/yyyy')
        prods.to_excel(writer, index=False)
        writer.close()
    if comments.shape[0] > 0:
        comments['Extraction Date'] = pd.to_datetime(comments['Extraction Date'], errors='coerce', format="%d_%m_%Y")
        comments['Extraction Date'] = comments['Extraction Date'].dt.date
        comments['Comment Date'] = pd.to_datetime(comments['Comment Date'], errors='coerce', format="%d-%m-%Y")
        comments['Comment Date'] = comments['Comment Date'].dt.date
        writer = pd.ExcelWriter(output2, date_format='d/m/yyyy')
        comments.to_excel(writer, index=False)
        writer.close()

def initialize_output():

    # removing the previous output file
    stamp = datetime.now().strftime("%d_%m_%Y_%H_%M")
    path = os.getcwd() + '\\scraped_data\\' + stamp
    if os.path.exists(path):
        shutil.rmtree(path)

    os.makedirs(path)

    file1 = f'Ztore_{stamp}.xlsx'
    file2 = f'Ztore_Comments_{stamp}.xlsx'

    # Windws and Linux slashes
    if os.getcwd().find('/') != -1:
        output1 = path.replace('\\', '/') + "/" + file1
        output2 = path.replace('\\', '/') + "/" + file2

    else:
        output1 = path + "\\" + file1
        output2 = path + "\\" + file2


    workbook1 = xlsxwriter.Workbook(output1)
    workbook1.add_worksheet()
    workbook1.close()    
    workbook2 = xlsxwriter.Workbook(output2)
    workbook2.add_worksheet()
    workbook2.close()    

    return output1, output2

def get_inputs():
 
    print('Processing The Settings Sheet ...')
    print('-'*100)
    # assuming the inputs to be in the same script directory
    path = os.getcwd()
    if '\\' in path:
        path += '\\Ztore_settings.xlsx'
    else:
        path += '/Ztore_settings.xlsx'

    if not os.path.isfile(path):
        print('Error: Missing the settings file "Ztore_settings.xlsx"')
        input('Press any key to exit')
        sys.exit(1)
    try:
        settings = {}
        links = []
        df = pd.read_excel(path)
        cols  = df.columns
        for col in cols:
            df[col] = df[col].astype(str)

        inds = df.index
        for ind in inds:
            row = df.iloc[ind]
            for col in cols:
                if row[col] == 'nan': continue
                elif col == 'Product Link':
                    links.append(row[col])                
                elif col == 'Search Link':
                    links.append(row[col])
                else:
                    settings[col] = row[col]

    except:
        print('Error: Failed to process the settings sheet')
        input('Press any key to exit')
        sys.exit(1)

    # checking the settings dictionary
    keys = ["Scrape Comments", "Comment Limit", "Product Limit"]
    for key in keys:
        if key not in settings.keys():
            print(f"Warning: the setting '{key}' is not present in the settings file")
            settings[key] = 0
        try:
            settings[key] = int(float(settings[key]))
        except:
            input(f"Error: Incorrect value for '{key}', values must be numeric only, press an key to exit.")
            sys.exit(1)

    return settings, links

if __name__ == '__main__':

    start = time.time()  
    settings, links = get_inputs()
    output1, output2 = initialize_output()  
    while True:
        try:
            driver = initialize_bot()
            prod_links = process_links(driver, links, settings)
            scrape_prods(driver, prod_links, output1, output2, settings)
            driver.quit()
            break
        except Exception as err:
            print('The below error occurred:\n')
            print(err)
            driver.quit()
            time.sleep(5)

    print('-'*100)
    elapsed = round(((time.time() - start)/60), 3)
    hrs = round(elapsed/60, 3)
    input(f'Process is completed successfully in {elapsed} mins ({hrs} hours). Press any key to exit.')
    sys.exit()


