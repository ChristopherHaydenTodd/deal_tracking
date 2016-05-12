#!/usr/bin/python
"""
    Pull RSS Data, Parse, Create Excel Report, and Email
"""

from datetime import datetime
import logging
import os
import sys
import xml.etree.ElementTree as ET

import requests
import xlwt
import pytz

CURRENT_PATH = os.path.dirname(os.path.realpath(__file__))
FILENAME = os.path.splitext(os.path.basename(__file__))[0]
sys.path.insert(0, CURRENT_PATH + '/../config')

from config.config import Config
CONFIGS = Config()

logging.basicConfig(
    level=CONFIGS.LOG_LEVEL, format=CONFIGS.LOG_FORMAT,
    datefmt=CONFIGS.LOG_DATEFORMAT, filemode=CONFIGS.LOG_FILEMODE,
    filename='../logs/{0}.log'.format(FILENAME))


def main():
    """
        Pull RSS Data from Hip2Save and Parse
    """

    logging.info('Starting Script to Pull Hip2Save RSS Feed')

    rss_xml = pull_hip_2_save_rss_xml()
    hip_2_save_deals = parse_hip_2_save_rss_xml(rss_xml)
    hip_2_save_deal_report =\
        build_hip_2_save_deal_report(hip_2_save_deals)

    logging.info('Script to Pull Hip2Save RSS Feed Complete')

    return


def pull_hip_2_save_rss_xml():
    """
        Pull RSS Data from Hip2Save's RSS Feed XML
    """
    logging.info('Pulling Hip2Save RSS Feed XML')

    request = requests.get(CONFIGS.HIP_2_SAVE_RSS_URL)
    rss_feed_xml = request.text

    if not rss_feed_xml:
        raise Exception('Cannot Pull RSS Data from Hip2Save')

    return rss_feed_xml


def parse_hip_2_save_rss_xml(rss_xml):
    """
        Parse XML RSS Feed from Hip2Save
    """
    logging.info('Parsing Hip2Saves RSS Feed XML')

    hip_2_save_deals = []

    xml_root = ET.fromstring(rss_xml.encode('utf-8'))
    channel_root = xml_root.find('channel')

    for item in channel_root.findall('item'):

        hip_2_save_deals.append({
            'title': item.find('title').text,
            'link': item.find('link').text,
            'published_date': item.find('pubDate').text,
            'category': item.find('category').text,
            'description': item.find('description').text
        })

    return hip_2_save_deals


def build_hip_2_save_deal_report(hip_2_save_deals):
    """
        Build Hip2Save Report
    """
    logging.info('Creating Hip2Save Deal Report')

    workbook = xlwt.Workbook()

    headers_style = xlwt.easyxf(
        'align:wrap on; pattern: pattern solid, fore_colour gray25')
    headers = hip_2_save_deals[0].keys()

    sheet = workbook.add_sheet('Hip2Save Deals')
    for index, value in enumerate(headers):
        sheet.write(0, index, value, headers_style)
        sheet.col(index).width = ((1 + len(value)) * 300)

    for row, deal in enumerate(hip_2_save_deals, 1):
        for column, key in enumerate(headers):
            sheet.write(row, column, deal[key])

    report_date =\
        datetime.now(pytz.timezone('US/Eastern')).strftime('%Y-%m-%d')
    hip_2_save_deal_report =\
        '../data/deals_{0}.xls'.format(report_date)
    workbook.save(hip_2_save_deal_report)

    return hip_2_save_deal_report


if __name__ == '__main__':

    try:
        main()
    except Exception, err:
        logging.error('Failed to Run Script %s', err)
        raise
