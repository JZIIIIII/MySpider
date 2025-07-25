# Spider/__init__.py

from .Tao_collection_tool import TaobaoScraper
from .PDD_collection_tool import PDDSpider
from .JD_collection_tool import JDScraper
from .Ali1688_collection_tool import Ali_1688Scraper
from .BaseScraper import BaseScraper
from .AntiScrapingException import CaptchaHandler
from .ExceLChange import ExcelImageProcessor

__version__ = "1.0.0"

__all__ = [
    "TaobaoScraper",
    "PDDSpider",
    "JDScraper",
    "Ali_1688Scraper",
    "BaseScraper",
    "CaptchaHandler",
    "ExcelImageProcessor"
]




