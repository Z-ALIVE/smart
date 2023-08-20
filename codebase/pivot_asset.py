from apinfvi import eSight
from DateTime import DateTime
from shutil import copyfile
import datetime
#from assetpivot import *
import vars
import pandas as pd
import json
import os, sys
import glob
import numpy as np
import openpyxl
import win32com.client as win32
import win32api
import win32con
import time
import shutil
import timeit

def open_id(site):
    openid = eSight(site).get_openid()
    return openid

def all_serverlist(sites, sites_n):
    with open('asset_pivot_serverlist.json', 'w+') as file:
        site_list = []
        for site in sites:
            openid = open_id(site)
            json_data = eSight(site).get_serverlist(openid, 'blade')
            data = json.loads(json_data)
            site_list.append(data['data'])
        site_list_json = json.dumps(site_list, indent=4, sort_keys=True)
        file.write(site_list_json)

def dn_blades(site,i=0,n=0):
    dn_blades = []
    openid = open_id(site)
    json_data = eSight(site).get_serverlist(openid, 'blade')
    data = json.loads(json_data)
    count_e9000 = len(data['data'])
    for i in range(count_e9000):
        childBlades = data['data'][i]['childBlades']
        count_blades = len(childBlades)
        for n in range(count_blades):
            dn = childBlades[n]['dn']
            dn_blades.append(dn)
    return dn_blades

def all_serverdetails(sites,sites_n):
    with open('asset_pivot_serverdetails.json', 'w+') as file:
        site_details = []
        for site,site_n in zip(sites,sites_n):
            DN_blades = dn_blades(site)
            openid = open_id(site)
            final_data = []
            for dn in DN_blades:
                json_data = eSight(site).get_serverdetails(openid, dn)
                data = json.loads(json_data)
                final_data.append(data['data'])
            server_details_json = json.dumps(final_data)
            server_details = json.loads(server_details_json)
            site_details.append(server_details)
        site_details_json = json.dumps(site_details, indent=4, sort_keys=True)
        file.write(site_details_json)
        return site_details_json

def all_networklist(sites, sites_n):
    with open('asset_pivot_networklist.json', 'w+') as file:
        site_list = []
        for site in sites:
            openid = open_id(site)
            json_data = eSight(site).get_networklist(openid)
            data = json.loads(json_data)
            site_list.append(data['data'])
        site_list_json = json.dumps(site_list, indent=4, sort_keys=True)
        file.write(site_list_json)

def all_storagelist(sites, sites_n):
    with open('asset_pivot_storagelist.json', 'w+') as file:
        site_list = []
        for site in sites:
            openid = open_id(site)
            json_data = eSight(site).get_storagelist(openid)
            data = json.loads(json_data)
            site_list.append(data['data'])
        site_list_json = json.dumps(site_list, indent=4, sort_keys=True)
        file.write(site_list_json)

def dn_storage(site,i=0,n=0):
    dn_storage = []
    openid = open_id(site)
    json_data = eSight(site).get_storagelist(openid)
    data = json.loads(json_data)
    count_stor = len(data['data'])
    for i in range(count_stor):
        dn = data['data'][i]['dn']
        dn_storage.append(dn)
    return dn_storage

def all_storagedisklist(sites, sites_n):
    with open('asset_pivot_storagedisklist.json', 'w+') as file:
        site_details = []
        for site in sites:
            DN_storage = dn_storage(site)
            openid = open_id(site)
            final_data = []
            for dn in DN_storage:
                json_data = eSight(site).get_storagedisklist(openid, dn)
                data = json.loads(json_data)
                for k in range(len(data['data'])):
                    data_n = data['data'][k]
                    data_n['dn'] = dn
                    final_data.append(data_n)
            storage_details_json = json.dumps(final_data)
            storage_details = json.loads(storage_details_json)
            site_details.append(storage_details)
        site_details_json = json.dumps(site_details, indent=4, sort_keys=True)
        file.write(site_details_json)


print("Start Running Script for PIVOT ASSET INVENTORY\n-----------------------------------------------------\n")
start = timeit.default_timer()
current_time = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
print('Start Time: {}'.format (current_time))

all_serverdetails(vars.sites, vars.sites_n)
all_serverlist(vars.sites, vars.sites_n)
all_networklist(vars.sites, vars.sites_n)
all_storagelist(vars.sites, vars.sites_n)
all_storagedisklist(vars.sites, vars.sites_n)

current_time2 = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
print('End Time: {}'.format(current_time2))
stop = timeit.default_timer()
print('Total time to run the program: ', stop - start)
print("_________________________END_________________________")


#if __name__ == '__main__':
    #open_id(site=vars.sites)
    #all_serverlist()
    #dn_blades()
    #all_serverdetails()

