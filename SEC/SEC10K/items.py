# -*- coding: utf-8 -*-

# Define here the models for your scraped items
#
# See documentation in:
# https://doc.scrapy.org/en/latest/topics/items.html

import scrapy


class SecScrapingItem(scrapy.Item):
#    Company name
    FirmName = scrapy.Field()
#    CIK
    FirmCIK = scrapy.Field()
#    Form type
    FormType = scrapy.Field()
#    Date
    FormDate = scrapy.Field()
#    Firm TIC
    FirmTIC = scrapy.Field()
#    Form Doc
    FormDoc = scrapy.Field()
#   Form Period
    FormPeriod = scrapy.Field()
