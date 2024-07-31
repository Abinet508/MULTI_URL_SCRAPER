import datetime
import os
import time
import urllib
import urllib.parse
import urllib.request
from playwright.sync_api import sync_playwright
import requests
from thefuzz import fuzz, process
from requests_html import HTML
import pandas as pd
from multiprocessing.pool import ThreadPool
import pandas as pd
import os
import time
import urllib.parse
from playwright.sync_api import sync_playwright
from requests_html import HTML
import pandas as pd
import pygsheets


class GoogleSearch:
    def __init__(self) -> None:
        """
        Initializes a GoogleSearch object.

        Attributes:
            urls (list): A list to store the URLs.
            results (list): A list to store the scraped results.
            df (DataFrame): A DataFrame to store the final results.
            current_dir (str): The current directory path.
            output_file (str): The path of the output Excel file.
            input_file (str): The path of the input text file.
        """
        self.urls = []
        self.results = []
        self.df = pd.DataFrame(columns=['PAGE', 'WEBSITE', 'LINK', 'HAS_PRODUCT', 'SCORE', 'TEXT'])
        self.current_dir = os.path.dirname(os.path.abspath(__file__))
        self.output_file = os.path.join(self.current_dir,'results', 'sample_v1.xlsx')
        self.input_file = os.path.join(self.current_dir,'Source', 'input.txt')
        os.makedirs(os.path.join(self.current_dir, 'CREDENTIALS'), exist_ok=True)
        os.makedirs(os.path.join(self.current_dir, 'Source'), exist_ok=True)
        os.makedirs(os.path.join(self.current_dir, 'results'), exist_ok=True)
        self.bing_q = True
        self.sheet_name = "Zoom Add on List"
        self.current_path = os.path.dirname(os.path.abspath(__file__))
        self.sheet_key = ""
        self.gc = pygsheets.authorize(service_account_file=os.path.join(self.current_path,"..","pygsheets","pygsheets-service-399508-aafa08c0b1b4.json"))
        
        self.file_path = os.path.join(self.current_dir, 'results')
    
    def timer(self):
        """
        This function is used to calculate the time taken by the function to execute.

        Returns:
            None
        """
        start_time = time.time()
        yield
        end_time = time.time()
        execution_time = end_time - start_time
        print(f"Execution time: {execution_time} seconds")

    def progress_bar(self, iteration, total, prefix='', suffix='', decimals=1, length=100, fill='█', printEnd="\r"):
        """
        This function is used to display the progress bar.

        Args:
            iteration (int): The current iteration number.
            total (int): The total number of iterations.
            prefix (str): The prefix string for the progress bar.
            suffix (str): The suffix string for the progress bar.
            decimals (int): The number of decimals to display.
            length (int): The length of the progress bar.
            fill (str): The character used to fill the progress bar.
            printEnd (str): The character used to print the end of the progress bar.

        Returns:
            None
        """
        percent = ("{0:." + str(decimals) + "f}").format(100 * (iteration / float(total)))
        filledLength = int(length * iteration // total)
        bar = fill * filledLength + '-' * (length - filledLength)
        print(f'\r{prefix} |{bar}| {percent}% {suffix}', end=printEnd)
        if iteration == total:
            print(flush=True)

    def add_to_df(self):
        """
        Adds the scraped results to the DataFrame.

        Returns:
            None
        """
        self.df = pd.DataFrame(self.results)
        # if take is True, then add the link to Has Product as hyperlink of YES else add the PAGE as hyperlink of NO
        self.df['HAS_PRODUCT'] = self.df.apply(lambda x: f'=HYPERLINK("{x["LINK"]}", "YES")' if x['TAKE'] else "NO", axis=1)
        for website in self.df['WEBSITE'].unique():
            website_rows = self.df[self.df['WEBSITE'] == website]
            website_rows_length = len(website_rows)
            row_index = website_rows.index[0]
            if website_rows_length > 1:
                for i in range(1, website_rows_length):
                    self.df.at[website_rows.index[0], f'PRODUCT{i}'] = f'=HYPERLINK("{website_rows.iloc[i]["LINK"]}", "{website_rows.iloc[i]["LINK_TEXT"]}")'
            else:
                self.df.at[row_index, 'PRODUCT1'] = f'=HYPERLINK("{website_rows.iloc[0]["LINK"]}", "{website_rows.iloc[0]["LINK_TEXT"]}")'
        self.df.drop(columns=['TAKE'], inplace=True)
        self.df.drop(columns=['LINK'], inplace=True)
        self.df.drop(columns=['LINK_TEXT'], inplace=True)
        #drop duplicates leaving the first row
        self.df.drop_duplicates(subset=['WEBSITE'], keep='first', inplace=True)

    def setup_GoogleSheet(self):
        """ setup_google_sheet is a function that setup google sheet
        """
        print("Google sheet setup started")
        
        os.makedirs(exist_ok=True,name=self.file_path)
        while True:
            try:
                self.sh = self.gc.open_by_key(self.sheet_key)
                #create self.sheet_name if it does not exist
                try:
                    self.wks = self.sh.worksheet_by_title(self.sheet_name)
                except:
                    self.sh.add_worksheet(self.sheet_name,src_worksheet=self.sh.sheet1)
                    self.wks = self.sh.worksheet_by_title(self.sheet_name)
                self.all_rows=self.wks.rows
                self.columunName=self.wks.get_row(1)
                sheet_title = self.sh.title
                try:
                    sheet_title = sheet_title.split("-")[0]
                except Exception as e:
                    pass
                finally:
                    get_date = datetime.datetime.now().strftime("%Y/%m/%d %I:%M %p")
                    sheet_title = sheet_title + "-" + get_date
                    self.sh.title = sheet_title
                    print("Sheet title is: ", sheet_title)
                    print("Google sheet setup successfully")
                    return True
            except Exception as e:
                print("Error in setup_GoogleSheet",e.__str__())
                time.sleep(5)
                continue
                
    def write_To_GoogleSheet(self):
        """ write_to_google_sheet is a function that write data to google sheet from excel file

        Args:
            file_name (string): file name
        """
        try:
            
            file_name=os.path.join(self.file_path,self.file_name)
            if os.path.exists(self.file_path)==False:
                os.makedirs(self.file_path)
                df = self.wks.get_as_df()
                df.to_excel(file_name,index=False)
            if os.path.exists(file_name)==False:
                df=self.wks.get_as_df()
                df.to_excel(file_name,index=False)
            
            df=pd.read_excel(file_name)
            self.wks.clear()
            self.wks.set_dataframe(df,(1,1))
            print("Data written to google sheet successfully")
            return True
        except Exception as e: 
            print(e.__str__())
            return False   
        
    def read_GoogleSheet_write_excel(self,reset=False,reset_sheet=False):
        """ read_GoogleSheet_write_excel is a function that read data from google sheet and save local changes to an excel file
        """
        file_name=os.path.join(self.file_path,self.file_name)
        
        print("Reading google sheet")
        try:
            if reset_sheet==True:
                df = self.wks.get_as_df()
                df = df.iloc[0:0]
                self.wks.clear()
                self.wks.set_dataframe(df,(1,1))
            if not os.path.exists(self.file_path):
                os.makedirs(self.file_path)
                df=self.wks.get_as_df()
                df.to_excel(file_name)
                return len(df)+1
            elif not os.path.exists(file_name):
                df=self.wks.get_as_df()
                df.to_excel(file_name,index=False)
                return len(df)+1
            else:
                if reset==True:
                    df=self.wks.get_as_df()
                    try:
                        os.remove(file_name)
                        print("file removed")
                    except Exception as e:
                        pass
                    df.to_excel(file_name,index=False)
                    return len(df)+1
                else:
                    df=pd.read_excel(file_name)
                    df=df.reset_index(drop=True)
                    df=df.drop_duplicates()
                    df=df.reset_index(drop=True)
                    df=df.fillna("")
                    df=df.replace("nan","")
                    df=df.replace("NaN","")
                    df=df.replace("None","")
                    df=df.replace("NONE","")
                    df=df.replace("none","")
                    df=df.replace("NAN","")
                    df=df.reset_index(drop=True)
                    self.wks.clear()
                    self.wks.set_dataframe(df,(1,1))
                    print("Google sheet read successfully")
                    return len(df)+1
        except Exception as e:
            print("Error in read_GoogleSheet_write_excel",e.__str__())
            return False
        
    def read_excel_file(self, file_path):
        """
        Reads the URLs from an Excel file.

        Args:
            file_path (str): The path of the Excel file.

        Returns:
            None
        """
        df = pd.read_excel(file_path)
        self.urls = df['WEBSITE'].tolist()

    def write_to_excel(self):
        """
        Writes the final results to an Excel file.

        Returns:
            None
        """
        
        self.df.to_excel(self.output_file, index=False)

    def read_txt_file(self):
        """
        Reads the URLs from a text file.

        Args:
            file_path (str): The path of the text file.

        Returns:
            list: A list of URLs read from the text file.
        """
        lines = []
        with open(self.input_file, 'r') as file:
            for line in file:
                lines.append(line.strip())  # Remove newline characters
        return lines

    def create_google_search_url(self, query):
        """
        Creates a Google search URL for the given query.

        Args:
            query (str): The search query.

        Returns:
            str: The Google search URL.
        """
        if self.bing_q:
            base_url = "https://www.bing.com/search?q="
        else:
            base_url = "https://www.google.com/search?q="
    
        encoded_query = urllib.parse.quote(query)
        full_url = base_url + encoded_query
        return full_url
    
    def setup_playwright(self):
        """
        Sets up the Playwright browser and context.
        """
        try:
            os.remove(os.path.join(self.current_dir, 'CREDENTIALS', 'storage_state.json'))
        except FileNotFoundError:
            pass
        try:
            context.close()
            browser.close()
        except:
            pass
        p = sync_playwright().start()
        browser = p.chromium.launch(headless=False,channel='msedge')
        if os.path.exists(os.path.join(self.current_dir, 'CREDENTIALS', 'storage_state.json')):
            context = browser.new_context(storage_state=os.path.join(self.current_dir, 'CREDENTIALS', 'storage_state.json'))
        else:
            context = browser.new_context()
        page = context.new_page()
        return page, context
    
    def scrape_data(self, urls):
        """
        Scrapes data from the given URLs.

        Args:
            urls (list): A list of URLs to scrape.

        Returns:
            None
        """
        
        response_list = []
        #urls = urls[:25]
        page, context = self.setup_playwright()
        try:
            url_counter = 0
            total_urls = len(urls)
            for url in urls:
                try:
                    url_counter += 1
                    url = url.replace('www.', '')
                    key_Words = "Change Connectors, Quick Disconnect Connectors, Fluidic Connectors, Hydraulic Connectors, Couplings"
                    key_Words_main = key_Words.split(', ')
                    key_Words = ' OR '.join(key_Words_main)
                    key_Words = f'{key_Words}'
                    #print(f"Scraping data from {key_Words}")
                    google_q = f'''inurl:{url} ("{key_Words}")'''
                    bing_q = f'''site:{url} ({key_Words})'''
                    #q = f'''inurl:{url} ("photovoltaic cable" OR "photovoltaic" OR "photovoltaic connector" OR "photovoltaic cable assembly" OR PV OR "PV cable assembly") (product OR buy OR shop OR store OR price OR catalog OR specifications)'''

                    response = page.goto(self.create_google_search_url(bing_q))
                    robot = page.locator('[id="captcha-form"]').all()
                    if robot:
                        print("Robot detected")
                        new_url = page.url
                        while robot:
                            time.sleep(1)
                            robot = page.locator('[id="captcha-form"]').all()
                        try:
                            os.remove(os.path.join(self.current_dir, 'CREDENTIALS', 'storage_state.json'))
                        except FileNotFoundError:
                            pass
                        os.makedirs(os.path.join(self.current_dir, 'CREDENTIALS'), exist_ok=True)
                        
                        context.storage_state(path=os.path.join(self.current_dir, 'CREDENTIALS', 'storage_state.json'))
                        page.goto(new_url)
                        response = page.goto(self.create_google_search_url(bing_q))
                    html_body = response.body()
                    html = HTML(html=html_body)
                    response_list.append(html_body)
                    if not self.bing_q:
                        texts = html.xpath('//div[@id="search"]//div[contains(@jscontroller,"") and contains(@lang,"en")]')
                    else:
                        texts = html.xpath('//ol[@id="b_results"]/li[contains(@class,"b_algo")]')
                    found = False
                    index = 1
                    for text in texts:
                        with ThreadPool() as pool:
                            link = text.xpath('//a')[0].attrs['href']
                            page1 = context.new_page()
                            page1.goto(link, timeout=60000, wait_until='domcontentloaded')
                            
                            print(f"Link: {page1.url}")
                            page1.close()
                            link_text = text.xpath('//a')[0].text
                            #check if any of in the key words is in the text at least partially meaning that the text contains the key words in the past or present form
                            score = 0
                            with ThreadPool() as pool:
                                results = pool.map(lambda x: fuzz.partial_ratio(x, text.text), key_Words_main)
                                score = max(results)
                            if score > 50:
                                data = {'PAGE': page.url, 'WEBSITE': url,'TAKE': True, 'LINK': link, 'SCORE': score, 'TEXT': text.text, 'LINK_TEXT': link_text}
                                self.results.append(data)
                                found = True
                                index += 1
                    if not found:
                        data = {'PAGE': page.url, 'WEBSITE': url, 'TAKE': False, 'LINK': 'N/A', 'SCORE': 0, 'TEXT': 'N/A', 'LINK_TEXT': 'N/A'}
                        self.results.append(data)    
                    self.progress_bar(url_counter, total_urls, prefix='CROWLING PROGRESS:', suffix=f'{url_counter}/{total_urls} URLS COMPLETED ', length=50)
                except Exception as e:
                    pass
            self.add_to_df()
        except KeyboardInterrupt:
            print("Script terminated by user.")

    def main(self, urls):
        """
        The main function to execute the scraping process.

        Args:
            urls (list): A list of URLs to scrape.

        Returns:
            None
        """
        self.scrape_data(urls)
        self.wks.clear()
        self.wks.set_dataframe(self.df,(1,1))
        #self.write_to_excel()
        
if __name__ == '__main__':
    gs = GoogleSearch()
    gs.setup_GoogleSheet()
    input_urls = gs.read_txt_file()
    gs.main(input_urls)
