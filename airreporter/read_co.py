# -*- coding: utf-8 -*-
"""
Created on Mon Jan 04 15:00:28 2016

@author: bbatt
"""

"""assumes comma delimited input from CAL3QHC"""

bld_input = r'C:\Users\bbatt\Dropbox\!Python\air-reporter\auxfiles\CO\Build.in'
bld_output = r'C:\Users\bbatt\Dropbox\!Python\air-reporter\auxfiles\CO\Build.out'


bld_in = open(bld_input, 'rb')
bld_out = open(bld_output, 'rb')

hdr = bld_in.readline().split(',')
print hdr

class CO(object):
    def __init__(self, co_input, co_output):
        pass
