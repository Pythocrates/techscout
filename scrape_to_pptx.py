#! /usr/bin/env python

from scrapy.crawler import CrawlerProcess
from scrapy.utils.project import get_project_settings

from techscout.spiders.engadget_spider import EngadgetSpider


if __name__ == '__main__':
    crawler = CrawlerProcess(get_project_settings())

    crawler.crawl(EngadgetSpider, follow_pages=0)
    crawler.start()
