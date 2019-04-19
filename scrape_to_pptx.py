#! /usr/bin/env python

from scrapy.crawler import CrawlerProcess
from scrapy.utils.project import get_project_settings

from techscout.spiders.engadget_spider import EngadgetSpider


if __name__ == '__main__':
    crawler_process = CrawlerProcess(get_project_settings())

    crawler_process.crawl(
        EngadgetSpider,
        follow_pages=-1,
        categories=['gear', 'gaming', 'entertainment', 'tomorrow'])
    crawler_process.start()
