import undetected_chromedriver as uc
import os
import time
from requests_html import HTML
import os
import time
import urllib
import urllib.parse
from thefuzz import fuzz
from requests_html import HTML
import pandas as pd
from multiprocessing.pool import ThreadPool
import pandas as pd
import os
import time
import urllib.parse
from requests_html import HTML
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
        #group by the WEBSITE if there are multiple row for same website add new column for the rest of the link and link text Product1	Product2	Product3	Product4	Product5
        # use the link text as the hyperlink text and link as the hyperlink for rows with take as True for same website except the first row name the columns as Product1, Product2, Product3, Product4, Product5
        #get number of products with same website
        for website in self.df['WEBSITE'].unique():
            website_rows = self.df[self.df['WEBSITE'] == website]
            website_rows_length = len(website_rows)
            row_index = website_rows.index[0]
            if website_rows_length > 1:
                for i in range(1, website_rows_length):
                    self.df.at[website_rows.index[0], f'PRODUCT{i}'] = f'=HYPERLINK("{website_rows.iloc[i]["LINK"]}", "{website_rows.iloc[i]["LINK_TEXT"]}")'
        
        self.df.drop(columns=['TAKE'], inplace=True)
        self.df.drop(columns=['LINK'], inplace=True)
        self.df.drop(columns=['LINK_TEXT'], inplace=True)
        #drop duplicates leaving the first row
        self.df.drop_duplicates(subset=['WEBSITE'], keep='first', inplace=True)

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
        base_url = "https://www.google.com/search?q="
        encoded_query = urllib.parse.quote(query)
        full_url = base_url + encoded_query
        return full_url
    
    def scrape_data(self, urls):
        driver = uc.Chrome()
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
                q = f'''inurl:{url} ("{key_Words}")'''
                #q = f'''inurl:{url} ("photovoltaic cable" OR "photovoltaic" OR "photovoltaic connector" OR "photovoltaic cable assembly" OR PV OR "PV cable assembly") (product OR buy OR shop OR store OR price OR catalog OR specifications)'''

                #response = page.goto(self.create_google_search_url(q))
                driver.get(self.create_google_search_url(q))

                # Handle CAPTCHA
                robot = driver.find_element(by='id', value='captcha-form').is_displayed()
                while robot:
                    time.sleep(1)
                    robot = driver.find_element(by='id', value='captcha-form').is_displayed()
                    # if robot:
                    #     input('Please solve the CAPTCHA and press Enter to continue...')

                # Remove and create directories
                try:
                    os.remove(os.path.join(self.current_dir, 'CREDENTIALS', 'storage_state.json'))
                except FileNotFoundError:
                    pass
                os.makedirs(os.path.join(self.current_dir, 'CREDENTIALS'), exist_ok=True)

                # Save storage state
                storage_state_path = os.path.join(self.current_dir, 'CREDENTIALS', 'storage_state.json')
                with open(storage_state_path, 'w') as f:
                    f.write(driver.get_cookies())

                # Perform Google search
                search_url = self.create_google_search_url(q)
                driver.get(search_url)
                response_body = driver.page_source
                html = HTML(html=response_body)
                texts = html.xpath('//div[@id="search"]//div[contains(@jscontroller,"") and contains(@lang,"en")]')
                found = False
                index = 1
                for text in texts:
                    link = text.xpath('//a')[0].links.pop()
                    link_text = text.xpath('//a')[0].text
                    #check if any of in the key words is in the text at least partially meaning that the text contains the key words in the past or present form
                    score = 0
                    with ThreadPool() as pool:
                        results = pool.map(lambda x: fuzz.partial_ratio(x, text.text), key_Words_main)
                        score = max(results)
                    if score > 50:
                        data = {'PAGE': driver.current_url, 'WEBSITE': url,'TAKE': True, 'LINK': link, 'SCORE': score, 'TEXT': text.text, 'LINK_TEXT': link_text}
                    #pprint.pprint(link)
                        self.results.append(data)
                        found = True
                        index += 1
                    if not found:
                        data = {'PAGE': driver.current_url, 'WEBSITE': url, 'TAKE': False, 'LINK': 'N/A', 'SCORE': 0, 'TEXT': 'N/A', 'LINK_TEXT': 'N/A'}
                        self.results.append(data)
                    self.progress_bar(url_counter, total_urls, prefix='CROWLING PROGRESS:', suffix=f'{url_counter}/{total_urls} URLS COMPLETED ', length=50)
            except Exception as e:
                pass
        self.add_to_df()
        self.write_to_excel()
        driver.quit()
    def main(self):
        """
        The main function.
        """
        #with self.timer():
        input_urls = self.read_txt_file()
        self.scrape_data(input_urls)
if __name__ == '__main__':
    google_search = GoogleSearch()
    google_search.main()