#!/usr/bin/python
"""
    Pull RSS Data, Parse, Create Excel Report, and Email
"""

import argparse
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email import encoders
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
import logging
import mimetypes
import os
import smtplib
import sys
import xml.etree.ElementTree as ET

import requests
import xlwt
import pytz

CURRENT_PATH = os.path.dirname(os.path.realpath(__file__))
FILENAME = os.path.splitext(os.path.basename(__file__))[0]
sys.path.insert(0, CURRENT_PATH + '/../../')

from config.config import Config
CONFIGS = Config()

logging.basicConfig(
    level=CONFIGS.LOG_LEVEL, format=CONFIGS.LOG_FORMAT,
    datefmt=CONFIGS.LOG_DATEFORMAT, filemode=CONFIGS.LOG_FILEMODE,
    filename='../../logs/{0}.log'.format(FILENAME))


def main(email, password):
    """
        Pull RSS Data from Hip2Save and Parse
    """

    logging.info('Starting Script to Pull Hip2Save RSS Feed')

    rss_xml = pull_hip_2_save_rss_xml()
    hip_2_save_deals = parse_hip_2_save_rss_xml(rss_xml)
    hip_2_save_deal_report =\
        build_hip_2_save_deal_report(hip_2_save_deals)
    email_hip_2_save_deal_report(email, password, hip_2_save_deal_report)

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
        '../../data/hip_2_save/deals_{0}.xls'.format(report_date)
    workbook.save(hip_2_save_deal_report)

    return hip_2_save_deal_report


def email_hip_2_save_deal_report(email, password, hip_2_save_deal_report):
    """
        Email Hip2Save Deal Report
    """
    logging.info('Emailing Hip2Save Deal Report')

    # Create Message
    msg = MIMEMultipart(_subtype='related')
    msg.set_charset('UTF-8')

    # Set Message Address and Subject
    report_date =\
        datetime.now(pytz.timezone('US/Eastern')).strftime('%Y-%m-%d')
    msg['Subject'] = 'Hip2Save Deal Report {0}'.format(report_date)
    msg['To'] = email
    msg['From'] = email 

    # Set Message Body
    html = MIMEMultipart()
    html.attach(MIMEText('Attached is the Daily Report', _subtype='html'))
    msg.attach(html)

    # Attach Report
    ctype, encoding = mimetypes.guess_type(hip_2_save_deal_report)
    maintype, subtype = ctype.split("/", 1)

    attachement_file = open(hip_2_save_deal_report, "rb")
    attachment = MIMEBase(maintype, subtype)
    attachment.set_payload(attachement_file.read())
    encoders.encode_base64(attachment)
    attachement_file.close()

    attachment.add_header(
        'Content-Disposition', 'attachment',
        filename=hip_2_save_deal_report.split('/')[-1])
    msg.attach(attachment)

    mailserver = smtplib.SMTP('smtp.gmail.com',587)
    mailserver.ehlo()
    mailserver.starttls()
    mailserver.ehlo()
    mailserver.login(email, password)
    mailserver.sendmail(email, email, msg.as_string())

    return


if __name__ == '__main__':

    parser = argparse.ArgumentParser(description='Process some integers.')
    parser.add_argument(
        '-e', '--email', help='Gmail Address to Send Report', required=True)
    parser.add_argument(
        '-p', '--password', help='Gmail Password', required=True)
    args = parser.parse_args()

    try:
        main(args.email, args.password)
    except Exception, err:
        logging.error('Failed to Run Script %s', err)
        raise
