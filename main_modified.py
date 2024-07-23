import os
import time
import urllib
import urllib.parse
from playwright.sync_api import sync_playwright
from thefuzz import fuzz, process
from requests_html import HTML
import pandas as pd

import pandas as pd
import os
import time
import urllib.parse
from playwright.sync_api import sync_playwright
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
        self.df.drop(columns=['TAKE'], inplace=True)
        self.df.drop(columns=['LINK'], inplace=True)

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
        """
        Scrapes data from the given URLs.

        Args:
            urls (list): A list of URLs to scrape.

        Returns:
            None
        """
        with sync_playwright() as p:
            browser = p.firefox.launch(headless=False)
            if os.path.exists(os.path.join(self.current_dir, 'CREDENTIALS', 'storage_state.json')):
                context = browser.new_context(storage_state=os.path.join(self.current_dir, 'CREDENTIALS', 'storage_state.json'))
            else:
                context = browser.new_context()
            page = context.new_page()
            response_list = []
            urls = urls[:25]
            try:
                url_counter = 0
                total_urls = len(urls)
                for url in urls:
                    url_counter += 1
                    url = url.replace('www.', '')
                    key_Words = ["Quick Change Connectors", "Quick Disconnect Connectors", "Fluidic Connectors", "Hydraulic Connectors", "Couplings"]
                    q = f'''inurl:{url} ("Quick Change Connectors" OR "Quick Disconnect Connectors" OR "Fluidic Connectors" OR "Hydraulic Connectors" OR "Couplings") (product OR buy OR shop OR store OR price OR catalog OR specifications)'''
                    #q = f'''inurl:{url} ("photovoltaic cable" OR "photovoltaic" OR "photovoltaic connector" OR "photovoltaic cable assembly" OR PV OR "PV cable assembly") (product OR buy OR shop OR store OR price OR catalog OR specifications)'''

                    response = page.goto(self.create_google_search_url(q))
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
                    html_body = response.body()
                    html = HTML(html=html_body)
                    response_list.append(html_body)
                    texts = html.xpath('//div[@id="search"]//div[contains(@jscontroller,"") and contains(@lang,"en")]')
                    found = False
                    for text in texts:
                        link = text.xpath('//a')[0].links.pop()
                        score = process.extractBests(text.text, key_Words, scorer=fuzz.token_sort_ratio)
                        if score[0][1] > 10:
                            data = {'PAGE': page.url, 'WEBSITE': url, 'TAKE': True, 'LINK': link, 'SCORE': score, 'TEXT': text.text}
                            #pprint.pprint(link)
                            self.results.append(data)
                            found = True
                    if not found:
                        data = {'PAGE': page.url, 'WEBSITE': url, 'TAKE': False, 'LINK': None, 'SCORE': None, 'TEXT': None}
                        self.results.append(data)
                    self.progress_bar(url_counter, total_urls, prefix='CROWLING PROGRESS:', suffix=f'{url_counter}/{total_urls} URLS COMPLETED ', length=50)
                self.add_to_df()
            except KeyboardInterrupt:
                print("Script terminated by user.")

            finally:
                context.close()

    def main(self, urls):
        """
        The main function to execute the scraping process.

        Args:
            urls (list): A list of URLs to scrape.

        Returns:
            None
        """
        self.scrape_data(urls)
        self.write_to_excel()
        
if __name__ == '__main__':
    gs = GoogleSearch()
    input_urls = gs.read_txt_file('Input.txt')
    gs.main(input_urls)
