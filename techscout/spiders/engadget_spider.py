# -*- coding: utf-8 -*-

import scrapy


class EngadgetSpider(scrapy.Spider):
    name = "engadget"
    next_page_selector =\
        'body > div.o-h > div > div div.table-cell > a.o-btn::attr(href)'

    def __init__(self, follow_pages=None, categories=None, *args, **kwargs):
        super(EngadgetSpider, self).__init__(*args, **kwargs)
        self.__pages_left = follow_pages if follow_pages is not None else -1
        self.__categories = categories or list()

    def start_requests(self):
        for category in self.__categories:
            url = 'https://www.engadget.com/topics/{c}/'.format(c=category)
            yield scrapy.Request(url=url, callback=self.parse)

    def parse(self, response):
        # Link to the next page.
        for href in response.css(self.next_page_selector):
            if self.__pages_left != 0:
                self.__pages_left -= 1
                yield response.follow(href, self.parse)

        # Links to the articles.
        # Top article.
        for href in response.css(
            '#engadget-the-latest div article h2 a::attr(href)'
        ):
            yield response.follow(href, self.parse_article)

        # Other articles.
        for href in response.css(
            '#engadget-the-latest div article.o-hit a.o-hit__link::attr(href)'
        ):
            yield response.follow(href, self.parse_article)

    def parse_article(self, response):
        page = response.url.split("/")[-2]
        print('Parsing article {a}...'.format(a=page))
        post_id = response.css('meta[name="post_id"]::attr(content)').get()
        tags = response.css('meta[property="article:tag"]::attr(content)').getall()
        title = response.css('meta[property="og:title"]::attr(content)').getall()
        text = response.css('.article-text > p *::text').getall()
        img_src = response.css('#page_body > div div > figure > img::attr(src)').get()
        yield dict(
            post_id=post_id,
            tags=tags,
            url=response.url,
            title=title,
            text=text,
            image_urls=[img_src],
        )
