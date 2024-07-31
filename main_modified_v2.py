import time
import pprint
import os
import pandas as pd
from googleapiclient.discovery import build
from thefuzz import fuzz, process

class GoogleSearch:
    def __init__(self) -> None:
        self.urls = []
        self.results = []
        self.df = pd.DataFrame(columns=['WEBSITE', 'LINK', 'HAS_PRODUCT', 'SEARCH_TERM', 'SNIPPET', 'TITLE', 'FUZZY_MATCH'])
        self.current_dir = os.path.dirname(os.path.abspath(__file__))
        self.output_file = os.path.join(self.current_dir,"results", 'sample_v2.xlsx')
        self.input_file = os.path.join(self.current_dir, "source", 'input.txt')
        os.makedirs(os.path.join(self.current_dir, 'results'), exist_ok=True)
        os.makedirs(os.path.join(self.current_dir, 'source'), exist_ok=True)
    def read_txt_file(self):
        lines = []
        with open(self.input_file, 'r') as file:
            for line in file:
                lines.append(line.strip())  # Remove newline characters
        return lines
    
    def main(self):
        urls = self.read_txt_file()
        urls = urls[:4]  # For testing purposes, only use the first 4 URLs
        # Build a service object for interacting with the API. Visit
        # the Google APIs Console <http://code.google.com/apis/console>
        # to get an API key for your own application.
        service = build(
            "customsearch", "v1", developerKey="<API_KEY>" # Replace <API_KEY> with your API key
        )
        cx="c4c9c752c7ba54abc"
        key_Words = ['Abinet Tesfu']
        query = " ".join(key_Words)
        for url in urls:
            res = (
                service.cse()
                .list(q=query,
                    cx=cx,
                    siteSearch=url)
                .execute()
            )
            time.sleep(1)
            self.results.append(res)

        # Process the results and store in DataFrame
        for result in self.results:
            # Process the result and extract required information
            if 'items' in result:
                for item in result['items']:
                    website = item['displayLink']
                    link = item['link']
                    snippet = item['snippet'] if 'snippet' in item else ''
                    search_term = result['queries']['request'][0]['searchTerms']
                    text = item['title'] + ' ' + snippet
                    fuzzy_match = process.extractBests(text.capitalize(),key_Words, score_cutoff=15)
                    has_product = any(fuzzy_match)
                    self.df = pd.concat([self.df, pd.DataFrame([[website, link, "YES" if has_product else "NO", search_term, snippet, item['title'], fuzzy_match]], columns=['WEBSITE', 'LINK', 'HAS_PRODUCT', 'SEARCH_TERM', 'SNIPPET', 'TITLE', 'FUZZY_MATCH'])])
                    if has_product:
                        self.df.loc[self.df['LINK'] == link, 'HAS_PRODUCT'] = f'=HYPERLINK("{link}","YES")'
                    else:
                        self.df.loc[self.df['LINK'] == link, 'HAS_PRODUCT'] = f'=HYPERLINK("{link}","NO")'
            else:
            
                website = result['queries']['request'][0]['siteSearch']
                search_term = result['queries']['request'][0]['searchTerms']
                self.df = pd.concat([self.df, pd.DataFrame([[website, '', 'NO', search_term, '', '', '']], columns=['WEBSITE', 'LINK', 'HAS_PRODUCT', 'SEARCH_TERM', 'SNIPPET', 'TITLE', 'FUZZY_MATCH'])])
        # Write the DataFrame to the output Excel file
        self.df.to_excel(self.output_file, index=False)

if __name__ == "__main__":
    gs = GoogleSearch()
    gs.main()