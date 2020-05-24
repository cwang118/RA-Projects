# -*- coding: utf-8 -*-
from scrapy.spider import BaseSpider
from scrapy.http import Request
from SEC10K.items import SecScrapingItem
import os
import pandas as pd
import re

os.chdir("D:/SEC10K/data/")
SP500 = pd.read_excel('SP500.xlsx',converters={'CIK':str})\
#.iloc[[5,27,151,182,340,453,459]]
SP500 = pd.DataFrame(SP500.values, columns = list(SP500.columns))
CIK = SP500['CIK']
TIC = SP500['Ticker symbol']
#CIK = SP500['CIK'].iloc[5];TIC = SP500['Ticker symbol'].iloc[5]
class Sec10kSpider(BaseSpider):
    name = 'SEC10K'
    start_urls = ['https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK=' + CIK[i] + '&type=10-k&dateb=&owner=include&count=100' for i in range(len(CIK))]
    MapCIK = dict(zip(CIK,TIC))
#    start_urls = ['https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK=0000718877&type=10-k&dateb=&owner=include&count=100']
    def parse(self, response):
        node_list = response.xpath("//tr/td[@nowrap='nowrap'][1]")
        for node in node_list:
            yield Request(str("https://www.sec.gov" + node.xpath("./following::td//a[1]/@href").extract_first()), callback=self.parseSecondLevel)
    def parseSecondLevel(self, response):
        item = SecScrapingItem()
        if ".htm" in response.xpath("//tr/td[@scope='row'][3]/a/text()").extract_first():
            #    Company name
            item['FirmName'] = re.search(r"(.*?)\ \(",response.xpath("//span[@class='companyName'][1]/text()").extract()[0]).group(1)
            #    CIK
            CIK_list = response.xpath("//span[@class='companyName']/a/text()").extract()
            for cik in CIK_list:
                if cik[:10] in CIK.values:
                    item['FirmCIK'] = cik[:10]
            #    Form type
            item['FormType'] = response.xpath("//tr/td[@scope='row'][4]/text()").extract_first()
            #    Date
            item['FormDate'] = response.xpath("//div[@class='formGrouping'][1]/div[@class='info'][1]/text()").extract_first()
            #    Form Doc
            item['FormDoc'] = str("https://www.sec.gov" + response.xpath("//tr/td[@scope='row'][3]/a/@href").extract_first())
            #   Form Period
            item['FormPeriod'] = response.xpath("//div[@class='formGrouping'][2]/div[@class='info'][1]/text()").extract_first()
            #    Firm TIC
            item['FirmTIC'] = self.MapCIK[str(item['FirmCIK'])]
            yield item



