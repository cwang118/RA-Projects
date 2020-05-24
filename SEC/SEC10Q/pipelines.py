# -*- coding: utf-8 -*-

# Define your item pipelines here
#
# Don't forget to add your pipeline to the ITEM_PIPELINES setting
# See: https://doc.scrapy.org/en/latest/topics/item-pipeline.html
import json
import scrapy
from scrapy.exporters import CsvItemExporter
from scrapy.pipelines.files import FilesPipeline
import os
import re

#from scrapy.utils.project import get_project_settings
#settings = get_project_settings()
FILES_STORE = 'D:/SEC10Q/data/Download/'

class SecScrapingPipeline(object):
    def __init__(self):
        self.f = open("SEC10Q.json","w")
    
    def process_item(self, item, spider):
        content = json.dumps(dict(item)) + ", \n"
        self.f.write(content)
        return item
    
    def close_spider(self, spider):
        self.f.close()


class CsvPipeline(object):
    def __init__(self):
        self.file = open('SEC10Q.csv', 'wb')
        self.exporter = CsvItemExporter(self.file)
        self.exporter.start_exporting()

    def close_spider(self, spider):
        self.exporter.finish_exporting()
        self.file.close()

    def process_item(self, item, spider):
        self.exporter.export_item(item)
        return item


class DownloadFiles(FilesPipeline):
    def get_media_requests(self, item, info):
        image_link = item['FormDoc']
        yield scrapy.Request(image_link)
    
    def item_completed(self, results, item, info):
        file_path = [x["path"] for ok, x in results if ok]
        item['FormType'] = re.sub("[\/\\\>\<\:\"\*\|\?]", '-', item['FormType'])
        Newfolder = item['FirmTIC'] + " " + item['FirmCIK'] + "/"
        if not os.path.exists(FILES_STORE + Newfolder):
            os.makedirs(FILES_STORE + Newfolder)
        os.rename(FILES_STORE + file_path[0], FILES_STORE + Newfolder + item['FirmTIC']\
                  + " " + item['FirmCIK'] + " " + item['FormType'] + " " + item['FormPeriod']\
                  + " " + item['FormDate'] + re.search(r"\.\w+$",item['FormDoc']).group(0))
        return item
        

