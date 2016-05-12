#!/usr/bin/python
"""
    Main Config File for Hip2Save Project
"""

import logging


class Config(object):
    """
        Project Configuration Class
    """

    ###
    # Log Configs
    ###

    LOG_LEVEL = logging.INFO
    LOG_FORMAT = '%(asctime)s %(levelname)-8s %(message)s'
    LOG_DATEFORMAT = '%a, %d %b %Y %H:%M:%S'
    LOG_FILEMODE = 'a'

    ###
    # Connection Configs
    ###

    HIP_2_SAVE_RSS_URL = 'https://hip2save.com/feed/'
