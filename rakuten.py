import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
from selenium.common.exceptions import TimeoutException
import validators
import time
import tkinter as tk
from tkinter import ttk
from tkinter.ttk import Progressbar

def main():
    
    site_url = []

    product_dataset = {
        '総合ジャンル':[],
        'ランク': [],
        '商品名':[],
        '価格':[],
    }

    my_dataset = pd.DataFrame(product_dataset)
        
    current_date = datetime.now().strftime('%Y-%m-%d')
    
    item_no = 0
    page_no = 0
    rank_no = 0
    options = webdriver.ChromeOptions()
    options.add_experimental_option("excludeSwitches",["ignore-certificate-errors"]) 
    # options.add_argument('--headless')
    browser = webdriver.Chrome(options=options)
    browser.minimize_window()
    # browser.implicitly_wait(200)
    
    
    for page_no in range(13):
        # if page_no > 0: break
        try:            
            starturl = f"https://ranking.rakuten.co.jp/weekly/p={page_no + 1}"
            browser.get(starturl)
            wait = WebDriverWait(browser, 20)
        except Exception as e:
            print({e})
            
        browser.execute_script("window.scrollTo(0, document.body.scrollHeight)")
        
        for item_no in range(80):
            rank_no = rank_no + 1
            
            if rank_no > 1000: 
                break
            
            if page_no == 0 and item_no >= 20:
                No = item_no + 2
            else:
                No = item_no + 1
            
            try:
                
                rank_name = wait.until(EC.visibility_of_element_located((By.XPATH, f'//*[@id="rnkRankingMain"]/div[{No}]/div[3]/div[1]/div/div/div/div[1]/a')))
                rank_price = wait.until(EC.visibility_of_element_located((By.XPATH, f'//*[@id="rnkRankingMain"]/div[{No}]/div[3]/div[2]/div[1]/div[1]')))
                browser.execute_script("arguments[0].scrollIntoView(true);", rank_price)
                sub_url = rank_name.get_attribute("href")
                
                site_url.append(sub_url)
                # time.sleep(0.1)
                text = f"{int(rank_no)}------{rank_name.text}-----{rank_price.text}" 
                print(text)
                
                my_dataset.at[rank_no, 'ランク'] = int(rank_no)
                my_dataset.at[rank_no, '商品名'] = rank_name.text
                my_dataset.at[rank_no, '価格'] = rank_price.text
                
            except Exception as e:
                print({e})
    browser.quit()
    
    # Generating overall genre    
    cou = 0
    
    for url in site_url:
        try:
            cou = cou + 1
        
            is_valid = validators.url(url)
            
            if not is_valid: 
                continue
            
            print(url)
            try:
                sub_browser = webdriver.Chrome(options=options)
                sub_browser.minimize_window()                           
                sub_browser.get(url)
                sub_wait = WebDriverWait(sub_browser, 3)
                # sub_browser.implicitly_wait(20)
            except Exception as e:
                pass      
            finally:
                pass
            
            genre = ""
            
            try:
                genre = sub_wait.until(EC.visibility_of_element_located((By.XPATH, f'//dd[@itemprop="breadcrumb"]'))).text
            except TimeoutException as e:
                pass
            
            if genre == "":        
                try:
                    genre = sub_wait.until(EC.visibility_of_element_located((By.XPATH,'//td[@class="sdtext"]/parent::tr/parent::tbody'))).text
                except TimeoutException as e:
                    pass
            
            my_dataset.at[cou, '総合ジャンル'] = genre   
            
            sub_browser.quit()
                
            with pd.ExcelWriter(f'{current_date}.xlsx', engine='openpyxl') as writer:
                my_dataset.to_excel(writer, sheet_name='Sheet1', encoding='utf-8', index=False)    
                
        except Exception as e:
            pass
        finally:
            pass
        
    
            
def implement_wait(time, day_of_week):
    current_date = datetime.now().strftime('%H-%M')
    
    # print(day_of_week.strip() + " ----- " + date[datetime.now().weekday()] )
    # print( time.strip() + " ----- " + current_date )
    
    if date[datetime.now().weekday()] == day_of_week.strip() and  current_date == time.strip():
        
        print(f"Start ...")
        main()
        
        print(f"Start ---> {current_date}")
        print(f"End ---> {datetime.now().strftime('%Y-%m-%d %H-%M-%S')}")
    else:
        print("Wait...")
            
   
# Start scraping
if __name__ == "__main__":
    date = ["月", "火", "水", "木", "金", "土", "日"]
    
    global window
    
    window = tk.Tk()
    window.title("Rakuten-Rank")

    window.geometry("630x50")
    window.resizable(False, False)
    
    label = tk.Label(window, text="時間:", font=("Arial", 15))
    label.grid(row=0, column=0, padx=5, pady=10)
    
    entry = tk.Entry(window, font=("Arial", 15))
    entry.insert(0, "01-00")
    entry.grid(row=0, column=1, padx=10, pady=10)
    
    label = tk.Label(window, text="曜日:", font=("Arial", 15))
    label.grid(row=0, column=2, padx=5, pady=10)

    combo = ttk.Combobox(window, font=("Arial", 15), values=["月", "火", "水", "木", "金", "土", "日"])
    combo.set("月")
    combo.grid(row=0, column=3, padx=5, pady=10)
    
  
    def update():
        
        implement_wait(str(entry.get()), str(combo.get()))
        # window.update_idletasks()
        window.after(3000, update)

    update()

    # Run the main loop
    window.mainloop()
    