# -*- coding: utf-8 -*-
"""
Created on Fri Mar 11 15:53:38 2022

@author: tbednall
"""

data = []

def max_peers():
    '''A function that shows the largest number of peers'''
    maxp = 0
    for ID in globals.canvas_data:
        len_p = len(rater_peerID(int(ID)))
        if len_p > maxp: maxp = len_p
    return(maxp)