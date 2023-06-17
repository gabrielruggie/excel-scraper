import os 

class ScraperSettings:
    """
    Configures excel scraper to scrape specific excel workbook 
    and corresponding worksheets

    NOTE: Excel workbook must be in the same directory as both scraper scripts.
          If not then get_path_to_workbook must be redefined by consumer
    """
    XLSX_FILE_NAME: str = 'example2.xlsx'

    DATA_SHEET: str = 'Sheet1'

    GRAPH_SHEET: str = 'Sheet2'

    CWD: str = os.getcwd()

    def get_path_to_workbook (cls) -> str:
        """
        Constructs full path to excel workbook
        """
        return cls.CWD + '/' + cls.XLSX_FILE_NAME