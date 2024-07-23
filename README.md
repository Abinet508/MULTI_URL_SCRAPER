# Multi URL Scraper

This project is a multi URL scraper designed to extract specific information from multiple web pages. It uses XPath for parsing HTML and fuzzy matching to find relevant text. The results are stored in a structured format (Excel) for easy analysis. The scraper is secure and handles storage of credentials securely. The project is easy to configure and use, making it a versatile tool for web scraping tasks.

## Features

## Features

- **Multi-URL Scraping**: Scrape multiple URLs in a single run.
- **XPath Parsing**: Use XPath to accurately parse HTML content.
- **Fuzzy Matching**: Implement fuzzy matching to find relevant text even if it is not an exact match.
- **Structured Results**: Store results in a structured format (Excel) for easy analysis.
- **Secure Storage**: Handle storage of credentials securely.


## Installation

1. **Clone the repository**:
    ```sh
    git clone https://github.com/yourusername/multi-url-scraper.git
    cd multi-url-scraper
    ```

2. **Create a virtual environment**:
    ```sh
    python -m venv venv
    ```

3. **Activate the virtual environment**:
    - On Windows:
        ```sh
        venv\Scripts\activate
        ```
    - On macOS/Linux:
        ```sh
        source venv/bin/activate
        ```

4. **Install the dependencies**:
    ```sh
    pip install -r requirements.txt
    ```

## Usage

1. **Configure the scraper**:
    - Update the `config.json` file with the URLs and keywords you want to scrape.

2. **Run the scraper**:
    ```sh
    python main_modified.py
    ```

3. **View the results**:
    - The results will be stored in the `results` directory in XLSX format.

## Configuration

- **URLs**: List of URLs to scrape.
- **Keywords**: List of keywords for fuzzy matching.
- **Storage Path**: Directory where credentials and results will be stored.

## Dependencies

- Python 3.7+
- `requests`
- `lxml`
- `fuzzywuzzy`
- `pandas`
- `playwright`


