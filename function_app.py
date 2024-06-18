import json
from math import sqrt, cos
import re
import azure.functions as func
import os, logging, uuid
from azure.storage.blob import BlobServiceClient
from azure.identity import DefaultAzureCredential
import urllib.request
import pandas as pd
import xml.etree.ElementTree as ET
import defusedxml.ElementTree as dET

# from loguru import logger as LOGGER
import traceback
import openpyxl
# from openpyxl import load_workbook

import numpy as np

import socket
print(socket.gethostname())
import pprint

# import office365
# from office365.runtime.auth.authentication_context import AuthenticationContext
# from office365.sharepoint.client_context import ClientContext
# from office365.sharepoint.files.file import File 



MAX_REAL_FLOORS = 10

# https://ksnmagicplanfunc3e54b9.file.core.windows.net/attachment/Survey Portal Excel Sheet_Export_Template.xlsx


def cart_distance(p1 : tuple[float, float], p2 : tuple[float, float]) -> float:
    (x1, y1) = p1
    (x2, y2) = p2
    return sqrt((x2 - x1) ** 2 + (y2 - y1) ** 2)

def create_table(dict : dict[str, list[float]], headers : list,
                  do_not_sum : list[str] = [], 
                  styling: str = "", colour_table : bool = False) -> str:
    
    output = f'<table {styling}><tr>'
    
    for header in headers:
        output += f'<th>{header}</th>'
    output += '</tr>'
    
    for i, key in enumerate(dict):
        if colour_table:
            output += f'<tr><td><font color="{key[:len(key)-2]}"><b>Colour {i}</b></font></td>'
        else:
            output += f'<tr><td>{key}</td>'
        for elem in dict[key]:
            output += f'<td>{round(elem, 2)}</td>'
        if do_not_sum != ['All']:
            if key in do_not_sum:
                output += '<td>N/A</td></tr>'
            else:
                output += f'<td>{round(sum(dict[key]), 2)}</td></tr>'
    
    output += '</table>'
    return output


def create_table_text(dict, headers : list,
                  do_not_sum : list[str] = [], 
                  styling: str = "", order_list = []) -> str:
    try:
        
        # print(dict)
        
        
        
        
        output = f'<table {styling}><tr>'
        
        for header in headers:
            output += f'<th>{header}</th>'
        output += '</tr>'
        
        if len(order_list) != 0:
            for item in order_list:
                if item.isupper():
                    output += f'<tr><td><strong>{item}</strong></td>'
                else:
                    output += f'<tr><td>{item}</td>'
                value = dict[item] if item in dict.keys() else ''
                if (type(value) == bool and value == True):
                    value = "Yes"
                if (type(value) == bool and value == False):
                    value = "No"
                output += f'<td>{value}</td>'
                # print(item, value)
        else:
            for i, key in enumerate(dict):
                # print(key, dict[key])
                if key.isupper():
                    output += f'<tr><td><strong>{key}</strong></td>'
                else:
                    output += f'<tr><td>{key}</td>'
                output += f'<td>{dict[key]}</td>'

        output += '</table>'
        
        
    except Exception as ex:
        # exc_type, exc_obj, exc_tb = sys.exc_info()
        # output = "Line " + str(exc_tb.tb_lineno) + ": " + exc_type 
        
        output = str(ex)
        output = traceback.format_exc()
        # LOGGER.info('Exception : ' + str(traceback.format_exc()))
        print(output)
        
        # fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        # print(exc_type, fname, exc_tb.tb_lineno)
    finally:
        return output
    return output

def create_table_new(dict, headers : list,
                  do_not_sum : list[str] = [], 
                  styling: str = "", order_list = []) -> str:
    try:
        
        # print(dict)
        
        
        
        
        output = f'<table {styling}><tr>'
        
        for header in headers:
            output += f'<th>{header}</th>'
        output += '</tr>'
        
        if len(order_list) != 0:
            for item in order_list:
                if item.isupper():
                    output += f'<tr><td><strong>{item}</strong></td>'
                else:
                    output += f'<tr><td>{item}</td>'
                value = dict[item]["value"] if (item in dict.keys() and "value" in dict[item].keys()) else ''
                if (type(value) == bool and value == True):
                    value = "Yes"
                if (type(value) == bool and value == False):
                    value = "No"
                output += f'<td>{value}</td>'
                # print(item, value)
        else:
            for i, key in enumerate(dict):
                # print(key, dict[key])
                if key.isupper():
                    output += f'<tr><td><strong>{key}</strong></td>'
                else:
                    output += f'<tr><td>{key}</td>'
                output += f'<td>{dict[key]}</td>'

        output += '</table>'
        
        
    except Exception as ex:
        # exc_type, exc_obj, exc_tb = sys.exc_info()
        # output = "Line " + str(exc_tb.tb_lineno) + ": " + exc_type 
        
        output = str(ex)
        output = traceback.format_exc()
        # LOGGER.info('Exception : ' + str(traceback.format_exc()))
        print(output)
        
        # fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        # print(exc_type, fname, exc_tb.tb_lineno)
    finally:
        return output
    return output


def roof_general(json_val_dict):
    json_val_dict["Room in Roof"] = False
    json_val_dict["Suitable for Insulation *"] = False
    json_val_dict["Roof not suitable details*"] = ''
    json_val_dict["Notes (Roof)"] = ''
    
    if "sfi" not in json_val_dict.keys():
        json_val_dict["sfi"] = []
    if "sfi_dict" not in json_val_dict.keys():
        json_val_dict["sfi_dict"] = {}
    
    
    
    for n in range(1, 5):
        if f"Roof Type {n} Suitable for Insulation" in json_val_dict.keys():
            # print(f"Roof Type {n} Suitable for Insulation", json_val_dict[f"Roof Type {n} Suitable for Insulation"])
            if json_val_dict[f"Roof Type {n} Suitable for Insulation"] == True:
                json_val_dict["Suitable for Insulation *"] = True
                json_val_dict["sfi"].append(n)
                # json_val_dict["sfi_dict"][n] = 300

        if f"Roof type {n} Suitable for Insulation*" in json_val_dict.keys():
            # print(f"Roof type {n} Suitable for Insulation*", json_val_dict[f"Roof type {n} Suitable for Insulation*"])
            if json_val_dict[f"Roof type {n} Suitable for Insulation*"] == True:
                json_val_dict["Suitable for Insulation *"] = True
                json_val_dict["sfi"].append(n)
                # json_val_dict["sfi_dict"][n] = 300

        if f"Roof Type {n} Suitable for Insulation*" in json_val_dict.keys():
            # print(f"Roof Type {n} Suitable for Insulation*", json_val_dict[f"Roof Type {n} Suitable for Insulation*"])
            if json_val_dict[f"Roof Type {n} Suitable for Insulation*"] == True:
                json_val_dict["Suitable for Insulation *"] = True
                json_val_dict["sfi"].append(n)
                # json_val_dict["sfi_dict"][n] = 300

        if f"Roof Type {n} Sloping Ceiling Suitable for Insulation*" in json_val_dict.keys():
            # print(f"Roof Type {n} Sloping Ceiling Suitable for Insulation*", json_val_dict[f"Roof Type {n} Sloping Ceiling Suitable for Insulation*"])
            if json_val_dict[f"Roof Type {n} Sloping Ceiling Suitable for Insulation*"] == True:
                json_val_dict["Suitable for Insulation *"] = True
                json_val_dict["sfi"].append(str(n) + 's')
                # json_val_dict["sfi_dict"][str(n) + 's'] = 300

    # print('sfi: ', json_val_dict["sfi"])
    # print('sfi: ', json_val_dict["sfi_dict"])

    for n in range(1, 5):
        # Suitable?
        if n not in json_val_dict["sfi"]:
            continue
        
        # Existing? (Thickness)
        e = 0
        if f"Roof {n} Thickness (mm)*" in json_val_dict.keys():
            e = json_val_dict[f"Roof {n} Thickness (mm)*"]
        elif f"Roof {n} Thickness (mm)" in json_val_dict.keys():
            e = json_val_dict[f"Roof {n} Thickness (mm)"]
        print(f"Roof {n} Thickness (mm)", ": ",  e)
        
        # value: Area - add to appropriate dict entry
        # print(f"roof_{n}_area", ": ", json_val_dict[f"roof_{n}_area"])
        for t in [100, 150, 200, 250, 300]:
            if int(e) + int(t) >= 300:
                # print(f"need to add roof_{n}_area to dict entry {t}")
                key = str(t)
                if key not in json_val_dict["sfi_dict"].keys():
                    json_val_dict["sfi_dict"][key] = json_val_dict[f"roof_{n}_area"]
                else:
                    json_val_dict["sfi_dict"][key] += json_val_dict[f"roof_{n}_area"]
                break


    for n in range(1, 5):
        if json_val_dict["Suitable for Insulation *"] == False:
            if f"Roof Type {n} Not Suitable Details" in json_val_dict.keys():
                # print('n', ':', n, json_val_dict[f"Roof Type {n} Not Suitable Details"])
                json_val_dict["Roof not suitable details*"] += f"Roof Type {n} Not Suitable Details: "
                json_val_dict["Roof not suitable details*"] += json_val_dict[f"Roof Type {n} Not Suitable Details"]
                json_val_dict["Roof not suitable details*"] += "<BR>"
            if f"Roof Type {n} Sloping Ceiling Not Suitable Details*" in json_val_dict.keys():
                # print('n', ':', n, json_val_dict[f"Roof Type {n} Sloping Ceiling Not Suitable Details*"])
                json_val_dict["Roof not suitable details*"] += f"Roof Type {n} Sloping Ceiling Not Suitable Details: "
                json_val_dict["Roof not suitable details*"] += json_val_dict[f"Roof Type {n} Sloping Ceiling Not Suitable Details*"]
                json_val_dict["Roof not suitable details*"] += "<BR>"
        else:
            json_val_dict["Roof not suitable details*"] = 'N/A'
    # print('json_val_dict["Roof not suitable details*"]: ', json_val_dict["Roof not suitable details*"])
    
    for n in range(1, 5):
        if f"Notes (Roof Type {n})" in json_val_dict.keys():
            if json_val_dict[f"Notes (Roof Type {n})"] not in json_val_dict["Notes (Roof)"]:
                json_val_dict["Notes (Roof)"] += f"Notes (Roof Type {n}): "
                json_val_dict["Notes (Roof)"] += json_val_dict[f"Notes (Roof Type {n})"]
                json_val_dict["Notes (Roof)"] += "<BR>"
        if f"Notes (Roof Type {n})*" in json_val_dict.keys():
            if json_val_dict[f"Notes (Roof Type {n})*"] not in json_val_dict["Notes (Roof)"]:
                json_val_dict["Notes (Roof)"] += f"Notes (Roof Type {n})*: "
                json_val_dict["Notes (Roof)"] += json_val_dict[f"Notes (Roof Type {n})*"]
                json_val_dict["Notes (Roof)"] += "<BR>"
    # print('json_val_dict["Notes (Roof)"]: ', json_val_dict["Notes (Roof)"])
    
    for n in range(1, 5):
        if f"Roof {n} Type*" in json_val_dict.keys():
            # print(f"Roof {n} Type*", json_val_dict[f"Roof {n} Type*"])
            if json_val_dict[f"Roof {n} Type*"] == "Dormer / room in roof":
                json_val_dict["Room in Roof"] = True
        if f"Roof {n} Type" in json_val_dict.keys():
            # print(f"Roof {n} Type", json_val_dict[f"Roof {n} Type"])
            if json_val_dict[f"Roof {n} Type"] == "Dormer / room in roof":
                json_val_dict["Room in Roof"] = True
    



def walls_general(json_val_dict):

    json_val_dict["Is the property suitable for wall insulation? *"] = False
    json_val_dict["No wall insulation details *"] = ''
    json_val_dict["Notes (Walls)"] = ''
    
    if "sfwi" not in json_val_dict.keys():
        json_val_dict["sfwi"] = []
    if "sfwi_dict" not in json_val_dict.keys():
        json_val_dict["sfwi_dict"] = {}
    
    
    
    
    for n in range(1, 5):
        if f"Is wall type {n} suitable for wall insulation?" in json_val_dict.keys():
            # print(f"Is wall type {n} suitable for wall insulation?", json_val_dict[f"Is wall type {n} suitable for wall insulation?"])
            if json_val_dict[f"Is wall type {n} suitable for wall insulation?"] == True:
                json_val_dict["Is the property suitable for wall insulation? *"] = True
                json_val_dict["sfwi"].append(n)
                # json_val_dict["sfwi_dict"][n] = 300

        if f"Is wall type {n} suitable for wall insulation?*" in json_val_dict.keys():
            # print(f"Is wall type {n} suitable for wall insulation?*", json_val_dict[f"Is wall type {n} suitable for wall insulation?*"])
            if json_val_dict[f"Is wall type {n} suitable for wall insulation?*"] == True:
                json_val_dict["Is the property suitable for wall insulation? *"] = True
                json_val_dict["sfwi"].append(n)
                # json_val_dict["sfwi_dict"][n] = 300


    # print('sfwi: ', json_val_dict["sfwi"])
    # print('sfwi: ', json_val_dict["sfwi_dict"])



    for n in range(1, 5):
        if json_val_dict["Is the property suitable for wall insulation? *"] == False:
            if f"No wall type {n} insulation details" in json_val_dict.keys():
                # print('n', ':', n, json_val_dict[f"No wall type {n} insulation details"])
                json_val_dict["No wall insulation details *"] += f"No wall type {n} insulation details: "
                json_val_dict["No wall insulation details *"] += json_val_dict[f"No wall type {n} insulation details"]
                json_val_dict["No wall insulation details *"] += "<BR>"
        else:
            json_val_dict["No wall insulation details *"] = 'N/A'
    # print('json_val_dict["No wall insulation details *"]: ', json_val_dict["No wall insulation details *"])
    
    for n in range(1, 5):
        if f"Notes (Wall type {n} Walls)" in json_val_dict.keys():
            if json_val_dict[f"Notes (Wall type {n} Walls)"] not in json_val_dict["Notes (Walls)"]:
                json_val_dict["Notes (Walls)"] += f"Notes (Wall type {n} Walls): "
                json_val_dict["Notes (Walls)"] += json_val_dict[f"Notes (Wall type {n} Walls)"]
                json_val_dict["Notes (Walls)"] += "<BR>"
        if f"Notes (Wall type {n} Walls)*" in json_val_dict.keys():
            if json_val_dict[f"Notes (Wall type {n} Walls)*"] not in json_val_dict["Notes (Walls)"]:
                json_val_dict["Notes (Walls)"] += f"Notes (Wall type {n} Walls)*: "
                json_val_dict["Notes (Walls)"] += json_val_dict[f"Notes (Wall type {n} Walls)*"]
                json_val_dict["Notes (Walls)"] += "<BR>"
    # print('json_val_dict["Notes (Walls)"]: ', json_val_dict["Notes (Walls)"])
    
    if json_val_dict["Is the property suitable for wall insulation? *"] == False:
        json_val_dict["No wall insulation details *"] += json_val_dict["Notes (Walls)"]

def is_point_in_line_segment(x1, y1, a1, b1, a2, b2):
    # print('checking if (' + str(x1) + ',' + str(y1) + ') is contained in (' + str(a1) + ',' + str(b1) + ') -> (' + str(a2) + ',' + str(b2) + ')')
    epsilon = 0.001
    
    cp = (y1 - b1) * (a2 - a1) - (x1 - a1) * (b2 - b1)
    if abs(cp) > epsilon:
        return False
    
    dp = (x1 - a1) * (a2 - a1) + (y1 - b1) * (b2 - b1)
    if dp < 0:
        return False
    
    slba = (a2 - a1) * (a2 - a1) + (b2 - b1) * (b2 - b1)
    if dp > slba:
        return False
    
    return True



def linear_subset(x1, y1, x2, y2, a1, b1, a2, b2):
    if not is_point_in_line_segment(x1, y1, a1, b1, a2, b2):
        return False
    
    if not is_point_in_line_segment(x2, y2, a1, b1, a2, b2):
        return False
    
    return True



def XML_2_dict(root, t = "floor"):
    try:
        # d = {}
        xml_ref_dict = {}
        nwa_dict = {}
        obj_dict = {}
        xml_val_dict = {}
        
        id = root.get('id')
        plan_name = root.get('name')
        if plan_name[-1] == ' ':
            plan_name = plan_name[:-1]
        print('id', ':', id)
        print('plan_name', ':', plan_name)
        xml_val_dict['Application ID'] = plan_name
        xml_val_dict['id'] = id
        xml_val_dict['plan_name'] = plan_name
        
        xml_val_dict['Client Address'] = ''
        address_fields = ['street', 'city', 'province', 'country', 'postalCode']
        for af in address_fields:
            f = root.get(af)
            if f is not None:
                xml_val_dict['Client Address'] = (xml_val_dict['Client Address'] + ', ' + str(f)) if xml_val_dict['Client Address'] != '' else str(f)
        xml_val_dict['Eircode'] = root.get('postalCode')
        
        
        # date = root.find('values/value[@key="date"]').text
        # xml_val_dict['Survey Date *'] = date
        
        
        MagicPlan_2_SEAI_dict = {"date": "Survey Date *", "qf.34d66ce4q3": "rating_type", "qf.34d66ce4q4": "rating_purpose", "author": "Surveyor"}
        
        # note MagicPlan also has a separate "Surveyor" field ("qf.34d66ce4q1") but "author" is the one used for SEAI survey purposes
        
        
        values = root.findall('values/value')
        for value in values:
            k = value.attrib["key"]
            for mpk in MagicPlan_2_SEAI_dict:
                if k == mpk:
                    xml_val_dict[MagicPlan_2_SEAI_dict[mpk]] = value.text
                    print(MagicPlan_2_SEAI_dict[mpk], ':', xml_val_dict[MagicPlan_2_SEAI_dict[mpk]])
        
        
        
        
        # w = {}
        wd_list = ['634004d284d12@edit:0063fa41-fa2d-4493-9f86-dcd0263e8108', '634004d284d12@edit:0ecdca7d-a4c3-4692-893a-89e6eaa76e74', '634004d284d12@edit:28960da1-84f6-4f3b-a446-7c72b9febe9f', '634004d284d12@edit:28b0fb8c-47a4-4d9e-8ce5-2b35a1a0404e', '634004d284d12@edit:2b72a58f-7380-4b6c-9d74-667f937a9b57', '634004d284d12@edit:32b043c7-432a-409f-972d-a75b386b1789', '634004d284d12@edit:60194a47-84ce-414b-8368-69ec53167111', '634004d284d12@edit:6976cc78-3a2e-4935-99c6-6aff8011be8a', '634004d284d12@edit:735122f1-ab8b-47e8-b5ca-d4ec4d492f1c', '634004d284d12@edit:7d851726-6ff6-48f7-8371-9ea09bd5179f', '634004d284d12@edit:7f6101da-4b6d-4c31-9293-d59552aeff3a', '634004d284d12@edit:a9a0a953-0fd3-4733-b161-de4f08fe5d49', '634004d284d12@edit:e6026a1e-3089-4fe7-9ec4-8504b001eb2e', '634004d284d12@edit:fc02c0c5-d9d8-4679-8a77-dc75edf7f592', 'arcdoor', 'doorbypass', 'doorbypassglass', 'doordoublefolding', 'doordoublehinged', 'doordoublesliding', 'doorfolding', 'doorfrench', 'doorgarage', 'doorglass', 'doorhinged', 'doorpocket', 'doorsliding', 'doorslidingglass', 'doorswing', 'doorwithwindow', 'windowarched', 'windowawning', 'windowbay', 'windowbow', 'windowcasement', 'windowfixed', 'windowfrench', 'windowhopper', 'windowhung', 'windowsliding', 'windowtrapezoid', 'windowtriangle', 'windowtskylight1', 'windowtskylight2', 'windowtskylight3']
        xml_ref_dict['habitable_rooms'] = []
        xml_ref_dict['wet_rooms'] = []
        xml_ref_dict['exclude_rooms'] = []
        # xml_ref_dict['include_rooms'] = []
        xml_ref_dict['exclude_room_types'] = ['Attic', 'Balcony', 'Storage', 'Patio', 'Deck', 'Porch', 'Cellar', 'Garage', 'Furnace Room', 'Outbuilding', 'Unfinished Basement', 'Workshop']
        
        xml_ref_dict['habitable_room_types'] = ['Kitchen', 'Dining Room', 'Living Room', 'Bedroom', 'Primary Bedroom', "Children's Bedroom", 'Study', 'Music Room']
        xml_ref_dict['wet_room_types'] = ['Kitchen', 'Bathroom', 'Half Bathroom', 'Laundry Room', 'Toilet', 'Primary Bathroom']

        
        
        
        
        
        
        floors = root.findall('interiorRoomPoints/floor')
        for floor in floors:
            ft = floor.get('floorType')
            xml_ref_dict[floor.get('floorType')] = floor.get('uid')
            xml_ref_dict[floor.get('uid')] = floor.get('floorType')
            nwa_dict[ft] = {}
            
            for room in floor.findall('floorRoom'):
                if room.get('type') not in xml_ref_dict.keys():
                    xml_ref_dict[room.get('type')] = []
                xml_ref_dict[room.get('type')].append(room.get('uid'))
                xml_ref_dict[room.get('uid')] = room.get('type')
                # print(room.get('type'))
                if room.get('type') in xml_ref_dict['habitable_room_types']:
                    xml_ref_dict['habitable_rooms'].append(room.get('uid'))
                    xml_ref_dict['habitable_rooms'].append('floor ' + ft + " - " + room.get('type') + " - " + room.get('uid'))
                
                if room.get('type') in xml_ref_dict['wet_room_types']:
                    xml_ref_dict['wet_rooms'].append(room.get('uid'))
                    xml_ref_dict['wet_rooms'].append('floor ' + ft + " - " + room.get('type') + " - " + room.get('uid'))
                
                if room.get('type') in xml_ref_dict['exclude_room_types']:
                    xml_ref_dict['exclude_rooms'].append(room.get('uid'))
                    xml_ref_dict['exclude_rooms'].append('floor ' + ft + " - " + room.get('type') + " - " + room.get('uid') + " (" + room.get('area') + ")")
                # else:
                    # xml_ref_dict['include_rooms'].append(room.get('uid'))
                    # xml_ref_dict['include_rooms'].append('floor ' + ft + " - " + room.get('type') + " - " + room.get('uid') + " (" + room.get('area') + ")")
                
                # print('exclude_rooms', ':', xml_ref_dict['exclude_rooms'])
                
                for value in room.findall('values/value'):
                    key = value.get('key')
                    # print(key)
                    if key == "qcustomfield.2979903aq1": # Include?
                        # print(room.get('type'))
                        floor_area_include = value.text
                        # print('floor_area_include', ':', floor_area_include)
                        # if floor_area_include == '0':
                            # xml_ref_dict['exclude_rooms'].append(room.get('uid'))
                        if floor_area_include == '1':
                            if room.get('uid') in xml_ref_dict['exclude_rooms']:
                                xml_ref_dict['exclude_rooms'].remove(room.get('uid'))
                                xml_ref_dict['exclude_rooms'].remove('floor ' + ft + " - " + room.get('type') + " - " + room.get('uid') + " (" + room.get('area') + ")")
                                # print(xml_ref_dict['exclude_rooms'])
                                # print(room.get('type'))
                
                # print('exclude_rooms', ':', xml_ref_dict['exclude_rooms'])
                
                rt = room.get('type') + ' (' + room.get('uid') + ')'
                x = {}
                room_x = room.get('x')
                room_y = room.get('y')
                w_index = 0
                for point in room.findall('point'):
                    w_index += 1
                    # uid = point.get('uid')
                    x[w_index] = {}
                    for value in point.findall('values/value'):
                        if value.get('key') in ['qf.c52807ebq1', 'qf.bdbaf056q1', 'qf.c52807ebq1']:
                            x[w_index]['type'] = value.text
                    # if 'type' not in list(x[w_index].keys()):
                        # x.pop(w_index)
                        # continue
                    x[w_index]['uid'] = point.get('uid')
                    x[w_index]['x1'] = float(point.get('snappedX')) + float(room_x)
                    x[w_index]['y1'] = -float(point.get('snappedY')) - float(room_y)
                    x[w_index]['h'] = point.get('height')
                    for value in point.findall('values/value'):
                        if value.get('key') == "loadBearingWall":
                            # print("loadBearingWall", ':', value.text)
                            x[w_index]['loadBearingWall'] = value.text
                # print('ft', ':', ft)
                # print('rt', ':', rt)
                # print('x', ':', x)
                # print('len(x)', ':', len(x))
                
                        
                w_index = 0
                for wall in x:
                    w_index += 1
                    # print(list(x[1].keys()))
                    if w_index + 1 in list(x.keys()):
                        x[w_index]['x2'] = x[w_index + 1]['x1']
                        x[w_index]['y2'] = x[w_index + 1]['y1']
                    else:
                        x[w_index]['x2'] = x[1]['x1']
                        x[w_index]['y2'] = x[1]['y1']
                    x[w_index]['l'] = cart_distance((x[w_index]['x1'], x[w_index]['y1']), (x[w_index]['x2'], x[w_index]['y2']))
                    x[w_index]['a'] = float(x[w_index]['l']) * float(x[w_index]['h'])
                
                y = {}
                for wall in x:
                    uid = x[wall]['uid']
                    y[uid] = {}
                    if 'type' in list(x[wall].keys()):
                        y[uid]['type'] = x[wall]['type']
                    y[uid]['x1'] = x[wall]['x1']
                    y[uid]['y1'] = x[wall]['y1']
                    y[uid]['x2'] = x[wall]['x2']
                    y[uid]['y2'] = x[wall]['y2']
                    y[uid]['h'] = x[wall]['h']
                    y[uid]['l'] = x[wall]['l']
                    y[uid]['a'] = x[wall]['a']
                    # print(list(x[wall].keys()))
                    if 'loadBearingWall' in list(x[wall].keys()):
                        y[uid]['loadBearingWall'] = x[wall]['loadBearingWall']
                        
                
                # print('len(y)', ':', len(y))
                # print('y', ':', y)
                # print('adding wall dict y for room ' + rt + ' to nwa_dict')
                nwa_dict[ft][rt] = y
                
        # print('nwa_dict', ':', nwa_dict)
        
        # print("xml_ref_dict['exclude_rooms']", ':', str(xml_ref_dict['exclude_rooms']))
        # print("xml_ref_dict['include_rooms']", ':', str(xml_ref_dict['include_rooms']))
        
        
        
        
        # Create Object Dictionary 
            # - first get list of all objects on each floor
            # - then add any additional details available from "exploded" section (linked via "id" e.g. "W-1-5")
        floors = root.findall('floor')
        for floor in floors:
            ft = floor.get('floorType')
            
            o = {}
            
            for p in floor.findall('symbolInstance'):
                # print("p.get('uid')", ':', p.get('uid'))
                if p.get('symbol') in wd_list:
                    id = p.get('id')
                    o[id] = {}
                    o[id]['uid'] = p.get('uid')
                    o[id]['symbol'] = p.get('symbol')
            # print('o', ':', o)
                
            
            
            for p in floor.findall('exploded/door'):
                si = p.get('symbolInstance')
                # print('si', ':', si)
                if si in list(o.keys()):
                # o[si] = {}
                # o[si]['symbolInstance'] = window.get('symbolInstance')
                    o[si]['x1'] = p.get('x1')
                    o[si]['y1'] = -float(p.get('y1'))
                    o[si]['x2'] = p.get('x2')
                    o[si]['y2'] = -float(p.get('y2'))
                    o[si]['w'] = p.get('width')
                    o[si]['d'] = p.get('depth')
                    o[si]['h'] = p.get('height')
                    o[si]['a'] = float(o[si]['w']) * float(o[si]['h'])
            
            for p in floor.findall('exploded/window'):
                # o_index += 1
                si = p.get('symbolInstance')
                # print('si', ':', si)
                if si in list(o.keys()):
                # o[si] = {}
                # o[si]['symbolInstance'] = window.get('symbolInstance')
                    o[si]['x1'] = p.get('x1')
                    o[si]['y1'] = -float(p.get('y1'))
                    o[si]['x2'] = p.get('x2')
                    o[si]['y2'] = -float(p.get('y2'))
                    o[si]['w'] = p.get('width')
                    o[si]['d'] = p.get('depth')
                    o[si]['h'] = p.get('height')
                    o[si]['a'] = float(o[si]['w']) * float(o[si]['h'])
            
            
            
            
            # print('o', ':', o)
            obj_dict[ft] = o
            
            for room in floor.findall('floorRoom'):
                rt = room.get('type') + ' (' + room.get('uid') + ')'
                
                w = {}
                room_x = room.get('x')
                room_y = room.get('y')
                w_index = 0
                for point in room.findall('point'): # get (x3, y3)
                    w_index += 1
                    w[w_index] = {}
                    w[w_index]['uid'] = point.get('uid')
                    w[w_index]['x3'] = float(point.get('snappedX')) + float(room_x)
                    w[w_index]['y3'] = -float(point.get('snappedY')) - float(room_y)

                w_index = 0
                for wall in w: # get (x4, y4), the second point in each line segment - WARNING: relies on the assumption that the points are in order
                    w_index += 1
                    if w_index + 1 in list(w.keys()):
                        w[w_index]['x4'] = w[w_index + 1]['x3']
                        w[w_index]['y4'] = w[w_index + 1]['y3']
                    else:
                        w[w_index]['x4'] = w[1]['x3']
                        w[w_index]['y4'] = w[1]['y3']
                
                # print('ft', ':', ft)
                # print('rt', ':', '"' + rt + '"')
                # print('w', ':', w)
                
                for wall in w: # transfer values to nwa_dict (where wall key is "uid" instead of numbered index)
                    uid = w[wall]['uid']
                    nwa_dict[ft][rt][uid]['x3'] = w[wall]['x3']
                    nwa_dict[ft][rt][uid]['y3'] = w[wall]['y3']
                    nwa_dict[ft][rt][uid]['x4'] = w[wall]['x4']
                    nwa_dict[ft][rt][uid]['y4'] = w[wall]['y4']
                
                y = nwa_dict[ft][rt] # for brevity
                
                w_index = 0
                for wall in y:
                    w_index += 1
                    # print('wall', ':', wall)
                    
                    
                    
                    y[wall]['windows'] = []
                    y[wall]['net_a'] = y[wall]['a']
                    y[wall]['total_window_a'] = 0
                    for window in o:
                        if 'x1' not in list(o[window].keys()):
                            continue
                        if 'x3' not in list(o[window].keys()):
                            continue
                        # print(window)
                        if linear_subset(float(o[window]['x1']), float(o[window]['y1']), float(o[window]['x2']), float(o[window]['y2']), float(y[wall]['x3']), float(y[wall]['y3']), float(y[wall]['x4']), float(y[wall]['y4'])) == True:
                            y[wall]['windows'].append(window + ' (' + str(o[window]['a']) + ')')
                            y[wall]['net_a'] -= o[window]['a']
                            y[wall]['total_window_a'] += o[window]['a']
                            
                            # print('object ' + str(window) + ' (' + str(o[window]['x1']) + '\t' + str(o[window]['y1']) + ') -> (' + str(o[window]['x2']) + '\t' + str(o[window]['y2']) + ') is colinear with wall ' + str(wall) + ' (' + str(y[wall]['x3']) + '\t' + str(y[wall]['y3']) + ') -> (' + str(y[wall]['x4']) + '\t' + str(y[wall]['y4']) + ')')
                            # print('yes')
                    # print("w[wall]['windows']", ':', w[wall]['windows'])
                
                # print('y', ':', y)
                nwa_dict[ft][rt] = y
                
                
    except Exception as ex:
        output = str(ex) + "\n\n" + traceback.format_exc()
        # LOGGER.info('Exception : ' + str(traceback.format_exc()))
        print(output)
    
    finally:
        return xml_ref_dict, nwa_dict, xml_val_dict



def XML_2_dict_new(root, t = "floor"):
    '''
    Returns [xml_ref_dict, # links floors and rooms with their uids
            , nwa_dict # net wall area
            , obj_dict
            , xml_val_dict # includes any variables required as output
            , 
            ]
        
    '''
    try:
        # d = {}
        xml_ref_dict = {}
        nwa_dict = {}
        obj_dict = {}
        xml_val_dict = {}
        
        xml_val_dict['project_id'] = root.get('id')
        
        project_name = root.get('name')
        if project_name[-1] == ' ':
            project_name = project_name[:-1]
        xml_val_dict['project_name'] = project_name
        
        project_address = ''
        address_fields = ['street', 'city', 'province', 'country', 'postalCode']
        for af in address_fields:
            f = root.get(af)
            if f is not None:
                project_address = (project_address + ', ' + str(f)) if project_address != '' else str(f)
        xml_val_dict['project_address'] = project_address
        
        
        xml_val_dict['eircode'] = root.get('postalCode')
        
        
        
        
        MagicPlan_2_SEAI_dict = {
            "date": "survey_date"
            , "qf.34d66ce4q3": "rating_type"
            , "qf.34d66ce4q4": "rating_purpose"
            , "author": "Surveyor"
            }
        # note MagicPlan also has a separate "Surveyor" field ("qf.34d66ce4q1") but "author" is the one used for SEAI survey purposes
        # ToDo: confirm if also the case for BER
        values = root.findall('values/value')
        for mpk in MagicPlan_2_SEAI_dict:
            xml_val_dict[MagicPlan_2_SEAI_dict[mpk]] = ''
            for value in values:
                k = value.attrib["key"]
                if k == mpk:
                    xml_val_dict[MagicPlan_2_SEAI_dict[mpk]] = value.text
            print(MagicPlan_2_SEAI_dict[mpk], ':', xml_val_dict[MagicPlan_2_SEAI_dict[mpk]])
        
        # w = {}
        wd_list = ['634004d284d12@edit:0063fa41-fa2d-4493-9f86-dcd0263e8108', '634004d284d12@edit:0ecdca7d-a4c3-4692-893a-89e6eaa76e74', '634004d284d12@edit:28960da1-84f6-4f3b-a446-7c72b9febe9f', '634004d284d12@edit:28b0fb8c-47a4-4d9e-8ce5-2b35a1a0404e', '634004d284d12@edit:2b72a58f-7380-4b6c-9d74-667f937a9b57', '634004d284d12@edit:32b043c7-432a-409f-972d-a75b386b1789', '634004d284d12@edit:60194a47-84ce-414b-8368-69ec53167111', '634004d284d12@edit:6976cc78-3a2e-4935-99c6-6aff8011be8a', '634004d284d12@edit:735122f1-ab8b-47e8-b5ca-d4ec4d492f1c', '634004d284d12@edit:7d851726-6ff6-48f7-8371-9ea09bd5179f', '634004d284d12@edit:7f6101da-4b6d-4c31-9293-d59552aeff3a', '634004d284d12@edit:a9a0a953-0fd3-4733-b161-de4f08fe5d49', '634004d284d12@edit:e6026a1e-3089-4fe7-9ec4-8504b001eb2e', '634004d284d12@edit:fc02c0c5-d9d8-4679-8a77-dc75edf7f592', 'arcdoor', 'doorbypass', 'doorbypassglass', 'doordoublefolding', 'doordoublehinged', 'doordoublesliding', 'doorfolding', 'doorfrench', 'doorgarage', 'doorglass', 'doorhinged', 'doorpocket', 'doorsliding', 'doorslidingglass', 'doorswing', 'doorwithwindow', 'windowarched', 'windowawning', 'windowbay', 'windowbow', 'windowcasement', 'windowfixed', 'windowfrench', 'windowhopper', 'windowhung', 'windowsliding', 'windowtrapezoid', 'windowtriangle', 'windowtskylight1', 'windowtskylight2', 'windowtskylight3']
        xml_ref_dict['habitable_rooms'] = []
        xml_ref_dict['wet_rooms'] = []
        xml_ref_dict['exclude_rooms'] = []
        # xml_ref_dict['include_rooms'] = []
        xml_ref_dict['exclude_room_types'] = ['Attic', 'Balcony', 'Storage', 'Patio', 'Deck', 'Porch', 'Cellar', 'Garage', 'Furnace Room', 'Outbuilding', 'Unfinished Basement', 'Workshop']
        
        xml_ref_dict['habitable_room_types'] = ['Kitchen', 'Dining Room', 'Living Room', 'Bedroom', 'Primary Bedroom', "Children's Bedroom", 'Study', 'Music Room']
        xml_ref_dict['wet_room_types'] = ['Kitchen', 'Bathroom', 'Half Bathroom', 'Laundry Room', 'Toilet', 'Primary Bathroom']

        
        
        
        
        
        
        floors = root.findall('interiorRoomPoints/floor')
        for floor in floors:
            ft = floor.get('floorType')
            xml_ref_dict[floor.get('floorType')] = floor.get('uid')
            xml_ref_dict[floor.get('uid')] = floor.get('floorType')
            nwa_dict[ft] = {}
            
            for room in floor.findall('floorRoom'):
                if room.get('type') not in xml_ref_dict.keys():
                    xml_ref_dict[room.get('type')] = []
                xml_ref_dict[room.get('type')].append(room.get('uid'))
                xml_ref_dict[room.get('uid')] = room.get('type')
                # print(room.get('type'))
                if room.get('type') in xml_ref_dict['habitable_room_types']:
                    xml_ref_dict['habitable_rooms'].append(room.get('uid'))
                    xml_ref_dict['habitable_rooms'].append('floor ' + ft + " - " + room.get('type') + " - " + room.get('uid'))
                
                if room.get('type') in xml_ref_dict['wet_room_types']:
                    xml_ref_dict['wet_rooms'].append(room.get('uid'))
                    xml_ref_dict['wet_rooms'].append('floor ' + ft + " - " + room.get('type') + " - " + room.get('uid'))
                
                if room.get('type') in xml_ref_dict['exclude_room_types']:
                    xml_ref_dict['exclude_rooms'].append(room.get('uid'))
                    xml_ref_dict['exclude_rooms'].append('floor ' + ft + " - " + room.get('type') + " - " + room.get('uid') + " (" + room.get('area') + ")")
                # else:
                    # xml_ref_dict['include_rooms'].append(room.get('uid'))
                    # xml_ref_dict['include_rooms'].append('floor ' + ft + " - " + room.get('type') + " - " + room.get('uid') + " (" + room.get('area') + ")")
                
                # print('exclude_rooms', ':', xml_ref_dict['exclude_rooms'])
                
                for value in room.findall('values/value'):
                    key = value.get('key')
                    # print(key)
                    if key == "qcustomfield.2979903aq1": # Include?
                        # print(room.get('type'))
                        floor_area_include = value.text
                        # print('floor_area_include', ':', floor_area_include)
                        # if floor_area_include == '0':
                            # xml_ref_dict['exclude_rooms'].append(room.get('uid'))
                        if floor_area_include == '1':
                            if room.get('uid') in xml_ref_dict['exclude_rooms']:
                                xml_ref_dict['exclude_rooms'].remove(room.get('uid'))
                                xml_ref_dict['exclude_rooms'].remove('floor ' + ft + " - " + room.get('type') + " - " + room.get('uid') + " (" + room.get('area') + ")")
                                # print(xml_ref_dict['exclude_rooms'])
                                # print(room.get('type'))
                
                # print('exclude_rooms', ':', xml_ref_dict['exclude_rooms'])
                
                rt = room.get('type') + ' (' + room.get('uid') + ')'
                x = {}
                room_x = room.get('x')
                room_y = room.get('y')
                w_index = 0
                for point in room.findall('point'):
                    w_index += 1
                    # uid = point.get('uid')
                    x[w_index] = {}
                    for value in point.findall('values/value'):
                        if value.get('key') in ['qf.c52807ebq1', 'qf.bdbaf056q1', 'qf.c52807ebq1']:
                            x[w_index]['type'] = value.text
                    # if 'type' not in list(x[w_index].keys()):
                        # x.pop(w_index)
                        # continue
                    x[w_index]['uid'] = point.get('uid')
                    x[w_index]['x1'] = float(point.get('snappedX')) + float(room_x)
                    x[w_index]['y1'] = -float(point.get('snappedY')) - float(room_y)
                    x[w_index]['h'] = point.get('height')
                    for value in point.findall('values/value'):
                        if value.get('key') == "loadBearingWall":
                            # print("loadBearingWall", ':', value.text)
                            x[w_index]['loadBearingWall'] = value.text
                # print('ft', ':', ft)
                # print('rt', ':', rt)
                # print('x', ':', x)
                # print('len(x)', ':', len(x))
                
                        
                w_index = 0
                for wall in x:
                    w_index += 1
                    # print(list(x[1].keys()))
                    if w_index + 1 in list(x.keys()):
                        x[w_index]['x2'] = x[w_index + 1]['x1']
                        x[w_index]['y2'] = x[w_index + 1]['y1']
                    else:
                        x[w_index]['x2'] = x[1]['x1']
                        x[w_index]['y2'] = x[1]['y1']
                    x[w_index]['l'] = cart_distance((x[w_index]['x1'], x[w_index]['y1']), (x[w_index]['x2'], x[w_index]['y2']))
                    x[w_index]['a'] = float(x[w_index]['l']) * float(x[w_index]['h'])
                
                y = {}
                for wall in x:
                    uid = x[wall]['uid']
                    y[uid] = {}
                    if 'type' in list(x[wall].keys()):
                        y[uid]['type'] = x[wall]['type']
                    y[uid]['x1'] = x[wall]['x1']
                    y[uid]['y1'] = x[wall]['y1']
                    y[uid]['x2'] = x[wall]['x2']
                    y[uid]['y2'] = x[wall]['y2']
                    y[uid]['h'] = x[wall]['h']
                    y[uid]['l'] = x[wall]['l']
                    y[uid]['a'] = x[wall]['a']
                    # print(list(x[wall].keys()))
                    if 'loadBearingWall' in list(x[wall].keys()):
                        y[uid]['loadBearingWall'] = x[wall]['loadBearingWall']
                        
                
                # print('len(y)', ':', len(y))
                # print('y', ':', y)
                # print('adding wall dict y for room ' + rt + ' to nwa_dict')
                nwa_dict[ft][rt] = y
                
        # print('nwa_dict', ':', nwa_dict)
        
        # print("xml_ref_dict['exclude_rooms']", ':', str(xml_ref_dict['exclude_rooms']))
        # print("xml_ref_dict['include_rooms']", ':', str(xml_ref_dict['include_rooms']))
        
        
        
        
        # Create Object Dictionary 
            # - first get list of all objects on each floor
            # - then add any additional details available from "exploded" section (linked via "id" e.g. "W-1-5")
        floors = root.findall('floor')
        for floor in floors:
            ft = floor.get('floorType')
            
            o = {}
            
            for p in floor.findall('symbolInstance'):
                # print("p.get('uid')", ':', p.get('uid'))
                if p.get('symbol') in wd_list:
                    id = p.get('id')
                    o[id] = {}
                    o[id]['uid'] = p.get('uid')
                    o[id]['symbol'] = p.get('symbol')
            # print('o', ':', o)
                
            
            
            for p in floor.findall('exploded/door'):
                si = p.get('symbolInstance')
                # print('si', ':', si)
                if si in list(o.keys()):
                # o[si] = {}
                # o[si]['symbolInstance'] = window.get('symbolInstance')
                    o[si]['x1'] = p.get('x1')
                    o[si]['y1'] = -float(p.get('y1'))
                    o[si]['x2'] = p.get('x2')
                    o[si]['y2'] = -float(p.get('y2'))
                    o[si]['w'] = p.get('width')
                    o[si]['d'] = p.get('depth')
                    o[si]['h'] = p.get('height')
                    o[si]['a'] = float(o[si]['w']) * float(o[si]['h'])
            
            for p in floor.findall('exploded/window'):
                # o_index += 1
                si = p.get('symbolInstance')
                # print('si', ':', si)
                if si in list(o.keys()):
                # o[si] = {}
                # o[si]['symbolInstance'] = window.get('symbolInstance')
                    o[si]['x1'] = p.get('x1')
                    o[si]['y1'] = -float(p.get('y1'))
                    o[si]['x2'] = p.get('x2')
                    o[si]['y2'] = -float(p.get('y2'))
                    o[si]['w'] = p.get('width')
                    o[si]['d'] = p.get('depth')
                    o[si]['h'] = p.get('height')
                    o[si]['a'] = float(o[si]['w']) * float(o[si]['h'])
            
            
            
            
            # print('o', ':', o)
            obj_dict[ft] = o
            
            for room in floor.findall('floorRoom'):
                rt = room.get('type') + ' (' + room.get('uid') + ')'
                
                w = {}
                room_x = room.get('x')
                room_y = room.get('y')
                w_index = 0
                for point in room.findall('point'): # get (x3, y3)
                    w_index += 1
                    w[w_index] = {}
                    w[w_index]['uid'] = point.get('uid')
                    w[w_index]['x3'] = float(point.get('snappedX')) + float(room_x)
                    w[w_index]['y3'] = -float(point.get('snappedY')) - float(room_y)

                w_index = 0
                for wall in w: # get (x4, y4), the second point in each line segment - WARNING: relies on the assumption that the points are in order
                    w_index += 1
                    if w_index + 1 in list(w.keys()):
                        w[w_index]['x4'] = w[w_index + 1]['x3']
                        w[w_index]['y4'] = w[w_index + 1]['y3']
                    else:
                        w[w_index]['x4'] = w[1]['x3']
                        w[w_index]['y4'] = w[1]['y3']
                
                # print('ft', ':', ft)
                # print('rt', ':', '"' + rt + '"')
                # print('w', ':', w)
                
                for wall in w: # transfer values to nwa_dict (where wall key is "uid" instead of numbered index)
                    uid = w[wall]['uid']
                    nwa_dict[ft][rt][uid]['x3'] = w[wall]['x3']
                    nwa_dict[ft][rt][uid]['y3'] = w[wall]['y3']
                    nwa_dict[ft][rt][uid]['x4'] = w[wall]['x4']
                    nwa_dict[ft][rt][uid]['y4'] = w[wall]['y4']
                
                y = nwa_dict[ft][rt] # for brevity
                
                w_index = 0
                for wall in y:
                    w_index += 1
                    # print('wall', ':', wall)
                    
                    
                    
                    y[wall]['windows'] = []
                    y[wall]['net_a'] = y[wall]['a']
                    y[wall]['total_window_a'] = 0
                    for window in o:
                        if 'x1' not in list(o[window].keys()):
                            continue
                        if 'x3' not in list(o[window].keys()):
                            continue
                        # print(window)
                        if linear_subset(float(o[window]['x1']), float(o[window]['y1']), float(o[window]['x2']), float(o[window]['y2']), float(y[wall]['x3']), float(y[wall]['y3']), float(y[wall]['x4']), float(y[wall]['y4'])) == True:
                            y[wall]['windows'].append(window + ' (' + str(o[window]['a']) + ')')
                            y[wall]['net_a'] -= o[window]['a']
                            y[wall]['total_window_a'] += o[window]['a']
                            
                            # print('object ' + str(window) + ' (' + str(o[window]['x1']) + '\t' + str(o[window]['y1']) + ') -> (' + str(o[window]['x2']) + '\t' + str(o[window]['y2']) + ') is colinear with wall ' + str(wall) + ' (' + str(y[wall]['x3']) + '\t' + str(y[wall]['y3']) + ') -> (' + str(y[wall]['x4']) + '\t' + str(y[wall]['y4']) + ')')
                            # print('yes')
                    # print("w[wall]['windows']", ':', w[wall]['windows'])
                
                # print('y', ':', y)
                nwa_dict[ft][rt] = y
                
                
    except Exception as ex:
        output = str(ex) + "\n\n" + traceback.format_exc()
        # LOGGER.info('Exception : ' + str(traceback.format_exc()))
        print(output)
    
    finally:
        return xml_ref_dict, nwa_dict, xml_val_dict


def no_2_alph(no):
    alph = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
    return alph[no - 1]


def XL_2_dict_new(url = ''
                , headers = {"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/39.0.2171.95 Safari/537.36"
                , "accept": "application/json"
                }
                , username = "mpsharepoint@ksnenergy.ie"
                , password = "S$658651652748ag"
                ):
    '''
    Read the contents of an Excel Workbook from Azure Blob Storage and return as a dictionary
    
    Three types of expected worksheet format: [trivial, table, lookup_table]
    
    trivial:
        2 columns, B & C, field_name and field_req
        field_req is one of the following four : [exact text, variable name, count, logic_dict]
        
    table: 
        need to update populate_template_new() to handle multicol outputs
        need to establish how many columns?
    
    lookup_table:
        2 cols
        create lookup_table_dict
        lookup is: field_value = lookup_table_dict[col_1]?
    
    '''
    # url = "https://ksnenergy.sharepoint.com/:x:/g/BERarea/QA/ETPUQWYZG1NOpv6HEPTLG6QBZhxJVA5Tbjda4m5if2Dngw?e=Tisozl"
    # url = "https://ksnenergy.sharepoint.com/BERarea/QA/Shared Documents/MagicPlan/Grant Thornton/BER Export Button/BER Export Button Tables.xlsx"
    
    try:
        # request = urllib.request.Request(url, headers=headers)
        # file_content = urllib.request.urlopen(request).read()
        
        
        print(username)
        print(password)
        
        
        ctx_auth = AuthenticationContext(url)
        if ctx_auth.acquire_token_for_user(username, password):
            ctx = ClientContext(url, ctx_auth)
            web = ctx.web
            ctx.load(web)
            ctx.execute_query()
            print("Authentication successful")
        
        
        
        
        wb = openpyxl.load_workbook(file_content, data_only = True)
        
        output = {}
        for sheet in wb.worksheets:
            # print(sheet.title)
            output[sheet.title] = {}
        
            for i, row in enumerate(list(sheet.values)):
                field_name = row[1]
                if field_name == None:
                    continue
                # print('field_name', ':', field_name)
                
                field_req = row[2]
                field_loc = no_2_alph(2) + str(i)
                # print('field_loc', ':', field_loc)
                output[sheet.title][field_name] = {"field_req": field_req, "field_loc": field_loc}
        
    except:
        output = traceback.format_exc()
        print(output)
        
    
    return output




def XL_2_dict(
                file_name
                , account_url = "https://ksnmagicplanfunc3e54b9.blob.core.windows.net"
                , default_credential = DefaultAzureCredential()
                , container_from = 'attachment'
                , local_path_from = "/tmp"
                # , container_to = 'project-files'
                # , local_path_to = plan_name
                ):
    '''
    Read the contents of an Excel Workbook from Azure Blob Storage and return as a dictionary
    
    Three types of worksheet expected format: [trivial, table, lookup_table]
    
    trivial:
        2 columns, B & C, field_name and field_req
        field_req is one of the following four : [exact text, variable name, count, logic_dict]
        
    table: 
        need to update populate_template_new() to handle multicol outputs
        need to establish how many columns?
    
    lookup_table:
        2 cols
        create lookup_table_dict
        lookup is: field_value = lookup_table_dict[col_1]?
    
    '''
    
    try:
        azure_source_fp = file_name
        instance_fp = os.path.join(os.getcwd(), local_path_from, file_name)
        print(instance_fp)
        
        # file_name = plan_name + ' Major Renovation calculation' + '.xlsx'


        # Create the BlobServiceClient object
        blob_service_client = BlobServiceClient(account_url, credential=default_credential)
        container_client = blob_service_client.get_container_client(container= container_from) 
        
        with open(file=instance_fp, mode="wb") as download_file:
            download_file.write(container_client.download_blob(azure_source_fp).readall())
        
        # file_content = container_client.download_blob(azure_source_fp).readall()
        
        wb = openpyxl.load_workbook(instance_fp, data_only = True)
        
        output_tables = ["1. Survey Details P1"
                        , "7. Thermal Mass P1"
                        , "8. Ventilation P1"
                        , "11. Lighting P1"
                        ]
        
        
        lookup_tables = ['1.1 Assessor Details Table'
                            , '2.2 Referance Table S8'
                            , '3.1 Referance Table S5'
                            , '3.2 Referance Table S4'
                            , '3.3 Referance Table Non Default'
                            , '7.1 Referance Table'
                            ]
       
        multicol_tables = ['5.1 Windows Summary Table'
                            , '5.2 Window Schedule Table'
                            , '5.3 Building | Doors P1'
                            , '6. Colour Area Table P1'
                            , '8.1 Ventilation Items'
                            , '11.1 Lighting Schedule'

                            ]
        output = {}
        lookup = {}
        for sheet in wb.worksheets:
            # print(sheet.title)

            # if sheet.title == '2.2 Referance Table S8':
                # for i, row in enumerate(list(sheet.values)):
                    # for j, x in enumerate(['Α', 'B', 'C', 'D', 'E', 'F', 'G', 'Η', 'I', 'J', 'K']):
                        # line = str(row[1]) + "_" + str(row[2]) + "_" + str(x) + "\t" + str(row[j+3])
                        # print(line)
                        
            
            if sheet.title in multicol_tables:
                print('multicol_table', ':', sheet.title)
            
            elif sheet.title in output_tables:
                    output[sheet.title] = {}
                    for i, row in enumerate(list(sheet.values)):
                        field_name = row[1]
                        if field_name == None:
                            continue
                        print('field_name', ':', field_name)
                        
                        field_req = row[2]
                        field_loc = no_2_alph(2) + str(i)
                        # print('field_loc', ':', field_loc)
                        output[sheet.title][field_name] = {"field_req": field_req, "field_loc": field_loc}
                
            elif sheet.title in lookup_tables:
                    lookup[sheet.title] = {}
                    for i, row in enumerate(list(sheet.values)):
                        field_name = row[1]
                        if field_name == None:
                            continue
                        print('field_name', ':', field_name)
                        
                        field_req = row[2]
                        field_loc = no_2_alph(2) + str(i)
                        # print('field_loc', ':', field_loc)
                        lookup[sheet.title][field_name] = {"field_req": field_req, "field_loc": field_loc}
                

            
    except:
        output = traceback.format_exc()
        print(output)
        
        
    return output, lookup


def JSON_2_dict(project_id, headers = {
            "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/39.0.2171.95 Safari/537.36"
            ,"key": "45170e50321733db78952dfa5901b0dfeeb8"
            , "customer": "63b5a4ae69c91"
            , "accept": "application/json"
            }):
    '''
    This function retrieves Forms & Statistics JSON files from MagicPlan
    Combines and filters the data to return json_dict:
        floor; room; object; details
        label_val_dict
    '''
    try:
        json_dict = {}
        
        # need json_Forms and json_Statistics
        
        # first go through Statistics
        # populate  json_dict = {floor; room; object; detail}
        # Floors -1 to 9
        # Then go through Forms and add any interesting information to the appropriate entry in json_dict
        # 
        # Count - should these be contained in json_dict? leaning towards no
        # 
        json_dict, form_dict = get_forms_data(project_id)
        
        
        
        pprint.pprint(form_dict)
        
        
        
        
        
        
        output = json_dict
    
    except:
        output = traceback.format_exc()
        print(output)
   
    return output

def BER(root, output = '', email = ''):
    # read template xlsx (xlst?)
    # create dicts of required table/tab contents
    # populate the dicts
    # use the dicts to populate the template and save it to be sent as an email attachment
    # use the dicts to generate this function's "output" HTML to serve as the body of the return email
    
    try:
        project_id = root.get("id")
        project_name = root.get("name") # ToDo: are we going to pass in xml_dict or do we need to produce a project-specific one?
        xml_dict = XML_2_dict_new(root)
        xml_val_dict = xml_dict[2]
        
        print('xml_val_dict', ':')
        pprint.pprint(xml_val_dict)
        
        
        
        json_dict = JSON_2_dict(project_id) # does this need to be project-specific?
        for f in ["Building Regulations Era", "Date of Plans", "Planning Reference"]:
            if f not in json_dict.keys():
                json_dict[f] = 'NOT FOUND'
        print("Orientation of front of building", ':', json_dict["Orientation of front of building"])
        
        ofl_filelist = []
        ofl_filelist = get_project_files(project_id, plan_name = project_name) # ofl_filelist is part of this function's output
        # print('finished getting project files')
        print('warning: did not get project files')
        
        # read template "BER template.xlsx"
        # create dictionary of required table/tab contents
        
        output_dict, lookup_dict = XL_2_dict(
                                    file_name = "template_ber.xlsx"
                                    , account_url = "https://ksnmagicplanfunc3e54b9.blob.core.windows.net"
                                    , default_credential = DefaultAzureCredential()
                                    , container_from = "attachment"
                                    , local_path_from = "/tmp"
                                    # , container_to = 'project-files'
                                    # , local_path_to = project_name
                                    )
        
        # output_dict = XL_2_dict_new(url)

        
        sheet_names = []
        for section in output_dict:
            sheet_names.append(section)
            # print('section', ':', section)
        print('sheet_names', ':', sheet_names)
        
        
        
        # populate the dicts 
        for sheet_name in sheet_names:
            print(sheet_name, ':')
            pprint.pprint(output_dict[sheet_name])
            for field in output_dict[sheet_name]:
                print()
                field_req = output_dict[sheet_name][field]['field_req']
                if field_req == None:
                    continue
                print('field', ':', field)
                print('field_req', ':', field_req)
                # first check if it's Exact Text (Forms question):
                if field_req in json_dict.keys():
                    output_dict[sheet_name][field]['value'] = json_dict[field_req]

                # then check if it's a variable name from xml_val_dict (i.e. xml_val_dict)
                elif field_req in xml_val_dict.keys():
                    output_dict[sheet_name][field]['value'] = xml_val_dict[field_req]
                
                # lookup_table
                elif field_req[0:6] == "lookup":
                    lu = field_req.split("|")
                    # print(eval(lu[2]))
                    output_dict[sheet_name][field]['value'] = lookup_dict[lu[1]][eval(lu[2])]['field_req']
                
                # count, need json_dict...?
                elif field_req[0:5] == "count":
                    output_dict[sheet_name][field]['value'] = "COUNT"
                
                # either/or logic
                elif field_req[0:5] == "logic":
                    # print(field_req)
                    lu = field_req.split("|")
                    # print(lu)
                    # print(eval(lu[1]))
                    # print(eval(lu[2]))
                    # print(lu[2])
                    # print(lu[4])
                    # print(eval(lu[4]))
                    output_dict[sheet_name][field]['value'] = eval(lu[1]) if json_dict[eval(lu[2])] == eval(lu[3]) else eval(lu[4])
                
                # logic is uncharted
                else:
                    output_dict[sheet_name][field]['value'] = "NOT FOUND"
                    
        
                # print('value', ':', output_dict[sheet_name][field]['value'])
        
        # use output_dict to populate (a copy of) the Excel template and save it to be sent as an email attachment
        
        
        # use output_dict to generate this function's "output" HTML to serve as the body of the return email
        
        if output == '':
            styling = "border=\"1\""
            output = f"""\
                <h1>File List</h1> \
                {create_table_text(output_dict, headers = ['name', 'value'], styling=styling, do_not_sum=['All'], order_list = ofl_filelist)} \

                </div>"""

            for section in output_dict:
                order_list = [field for field in output_dict[section]]
                # print(order_list)
                section_output = f"""\
                                <h1>{section}</h1> \
                                {create_table_new(output_dict[section], headers = ['name', 'value'], styling=styling, do_not_sum=['All'], order_list = order_list)} \
                                </div>"""
                                
                output = output + section_output
                
            output = output + "</div>"
        
    except:
        output = traceback.format_exc()
        print(output)
    
    return output



def preBER(root):
    return
def inspection(root):
    return
def QA(root):
    return



def get_project_files(id, plan_name, headers = {
            "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/39.0.2171.95 Safari/537.36"
            ,"key": "45170e50321733db78952dfa5901b0dfeeb8"
            , "customer": "63b5a4ae69c91"
            , "accept": "application/json"
            }):
            
    try:
        output = []
        generate_locally = True
        if (socket.gethostname()) == "PC1VXW6X":
            generate_locally = False
        # azure_upload(json_data)
        # account_url = os.environ['AZ_STR_URL']
        account_url = "https://ksnmagicplanfunc3e54b9.blob.core.windows.net"
        default_credential = DefaultAzureCredential()
        blob_service_client = BlobServiceClient(account_url, credential=default_credential)
        
        
        container_name = "project-files"
        container_client = blob_service_client.get_container_client(container_name)
        if not container_client.exists():
            container_client = blob_service_client.create_container(container_name)

        
        
        json_url = "https://cloud.magicplan.app/api/v2/plans/" + str(id) + "/files?include_photos=true"
        request = urllib.request.Request(json_url, headers=headers)
        JSON = urllib.request.urlopen(request).read()
        JSON = json.loads(JSON)
        
        # local_path = str(json_val_dict['Application ID'])
        # if not os.path.exists(local_path):
            # os.mkdir(local_path)
        
        
        
        
        for file in JSON["data"]["files"]:
            if file["file_type"] == "pdf":
                output.append(file["name"])
                # print('getting file: ' + file["name"])
                if generate_locally == True:
                    request = urllib.request.Request(file["url"], headers=headers)
                    file_content = urllib.request.urlopen(request).read()
                    local_file_name = file["name"]
                    blob_client = blob_service_client.get_blob_client(container=container_name, blob=os.path.join(plan_name, file["name"]))
                    blob_client.upload_blob(file_content, overwrite=True)

        
        for file in JSON["data"]["photos"]:
            local_file_name = file["name"]
            local_file_name = local_file_name.replace("10th Floor", "Ground Floor")
            local_file_name = local_file_name.replace("11th Floor", "1st Floor")
            local_file_name = local_file_name.replace("12th Floor", "2nd Floor")
            local_file_name = local_file_name.replace("13th Floor", "3rd Floor")
            local_file_name = local_file_name.replace("14th Floor", "4th Floor")
            output.append(local_file_name)
            # print('getting file: ' + file["name"], 'local: ', local_file_name)
            if generate_locally == True:
                request = urllib.request.Request(file["url"], headers=headers)
                file_content = urllib.request.urlopen(request).read()
                blob_client = blob_service_client.get_blob_client(container=container_name, blob=os.path.join(plan_name, local_file_name))
                blob_client.upload_blob(file_content, overwrite=True)
    
    except Exception as ex:
        # exc_type, exc_obj, exc_tb = sys.exc_info()
        # output = "Line " + str(exc_tb.tb_lineno) + ": " + exc_type 
        
        # output = str(ex)
        output = traceback.format_exc()
        print(output)
        
    finally:
        return output


def survey(root):
    try:
        output = ''
        xml_ref_dict, nwa_dict, xml_val_dict = XML_2_dict(root)
        # id = xml_val_dict['Application ID'] # take this out once all erroneous references have been updated
        id = xml_val_dict['id'] # take this out once all erroneous references have been updated
        plan_name = xml_val_dict['plan_name'] # take this out once all erroneous references have been updated
        
        
        # print('xml_val_dict', ':', xml_val_dict)
        json_val_dict = xml_val_dict # take this out once all erroneous references have been updated
        
        
        sfi = [] # a list to hold the numbers of roof (also wall?) types that are suitable for insulation
        
        headers = {
            "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/39.0.2171.95 Safari/537.36"
            ,"key": "45170e50321733db78952dfa5901b0dfeeb8"
            , "customer": "63b5a4ae69c91"
            , "accept": "application/json"
            }


        ofl_filelist = []
        
        # if (socket.gethostname()) != "PC1VXW6X":
        # print('about to get project files for ' + plan_name + " (id: " + str(id) + ")")
        ofl_filelist = get_project_files(id, plan_name)
        # print('finished getting project files')
        # for filename in ofl_filelist:
            # print(filename)
        
   
        # print('about to create (almost) empty attachment files for ' + plan_name + " (id: " + str(id) + ")")
        # populate_template(xml_val_dict) # adds an (almost) empty copy of the template to avoid potential Logic App error if file not found
        populate_template_new(xml_val_dict, 'template') # adds an (almost) empty copy of the template to avoid potential Logic App error if file not found
        # populate_template_new(xml_val_dict, 'template_mrc')
        print('finished creating empty attachment file')
            

        
        
        ofl_wos = ['ATTIC INSULATION'
                , 'Internal Wall Insulation: Sloped or flat (horizontal) surface'
                , 'Attic (Loft) Insulation 100 mm top-up'
                , 'Attic (Loft) Insulation 150 mm top-up'
                , 'Attic (Loft) Insulation 200 mm top-up'
                , 'Attic (Loft) Insulation 250 mm top up'
                , 'Attic (Loft) Insulation 300 mm'
                , 'Attic Storage (5m2)'
                , 'Installation of new attic hatch'
                , 'Additional Roof Ventilation (High Level)'
                , 'Additional Roof Ventilation (Low Level)'
                , 'WALLS'
                , 'Draught Proofing (<= 20m installed)'
                , 'Draught Proofing (> 20m installed)'
                , 'MEV 15l/s Bathroom'
                , 'MEV 30l/s Utility'
                , 'MEV 60l/s Kitchen'
                , 'Permanent ventilation wall vent (Certified Proprietary Integrated System)'
                , 'Background ventilation wall vent (Certified Proprietary Integrated System)'
                , 'Ducting existing cooker hood to exterior'
                , 'Cavity Wall Insulation Bonded Bead'
                , 'Loose Fibre Extraction'
                , 'External Wall Insulation: Less than 60m2'
                , 'External Wall Insulation: 60m2 to 85m2'
                , 'External Wall Insulation: Greater than 85m2'
                , 'ESB alteration'
                , 'GNI meter alteration'
                , 'GNI new connection'
                , 'RGI Meter_No Heating'
                , 'Internal Wall Insulation: Vertical Surface'
                , 'External wall insulation and CWI: less than 60m2'
                , 'External wall insulation and CWI: 60m2 to 85m2'
                , 'External wall insulation and CWI: greater than 85m2'
                , 'Window (same m2 rate will apply to windows with certified trickle vents)'
                , 'HEATING'
                , 'Basic gas heating system'
                , 'Basic oil heating system'
                , 'Full gas heating system installation'
                , 'Full oil heating system installation'
                , 'Gas boiler and controls (Basic & controls pack)'
                , 'Oil boiler and controls (Basic & controls pack)'
                , 'Hot Water Cylinder Jacket'
                , 'MECHANICAL VENTILATION SYSTEMS AND AIR TIGHTNESS TESTING & ENERGY'
                , 'Air-tightness test recommended?'
                , 'LED Bulbs: supply only (4 no.)'

                ]
        
        
        ofl_wos_2 = ['ATTIC INSULATION'
                , 'Internal Wall Insulation: Sloped or flat (horizontal) surface'
                , 'Attic (Loft) Insulation 100 mm top-up'
                , 'Attic (Loft) Insulation 150 mm top-up'
                , 'Attic (Loft) Insulation 200 mm top-up'
                , 'Attic (Loft) Insulation 250 mm top up'
                , 'Attic (Loft) Insulation 300 mm'
                , 'Attic Storage (5m2)'
                , 'Installation of new attic hatch'
                , 'Additional Roof Ventilation (High Level)'
                , 'Additional Roof Ventilation (Low Level)'
                , 'WALLS'
                , 'Draught Proofing (<= 20m installed)'
                , 'Draught Proofing (> 20m installed)'
                , 'MEV 15l/s Bathroom'
                , 'MEV 30l/s Utility'
                , 'MEV 60l/s Kitchen'
                , 'Permanent ventilation wall vent (Certified Proprietary Integrated System)'
                , 'Background ventilation wall vent (Certified Proprietary Integrated System)'
                , 'Ducting existing cooker hood to exterior'
                , 'Cavity Wall Insulation Bonded Bead'
                , 'Loose Fibre Extraction'
                , 'External Wall Insulation: Less than 60m2'
                , 'External Wall Insulation: 60m2 to 85m2'
                , 'External Wall Insulation: Greater than 85m2'
                , 'ESB alteration'
                , 'GNI meter alteration'
                , 'GNI new connection'
                , 'RGI Meter_No Heating'
                , 'Internal Wall Insulation: Vertical Surface'
                , 'External wall insulation and CWI: less than 60m2'
                , 'External wall insulation and CWI: 60m2 to 85m2'
                , 'External wall insulation and CWI: greater than 85m2'
                , 'Window (same m2 rate will apply to windows with certified trickle vents)'
                , 'HEATING'
                , 'Basic gas heating system'
                , 'Basic oil heating system'
                , 'Full gas heating system installation'
                , 'Full oil heating system installation'
                , 'Gas boiler and controls (Basic & controls pack)'
                , 'Oil boiler and controls (Basic & controls pack)'
                , 'Hot Water Cylinder Jacket'
                , 'MECHANICAL VENTILATION SYSTEMS AND AIR TIGHTNESS TESTING & ENERGY'
                , 'Air-tightness test recommended?'
                , 'LED Bulbs: supply only (4 no.)'

                ]
        
        ofl_pm = ['Internal Wall Insulation: Sloped or flat (horizontal) surface'
                , 'Attic (Loft) Insulation 100 mm top-up'
                , 'Attic (Loft) Insulation 150 mm top-up'
                , 'Attic (Loft) Insulation 200 mm top-up'
                , 'Attic (Loft) Insulation 250 mm top up'
                , 'Attic (Loft) Insulation 300 mm'
                , 'Cavity Wall Insulation Bonded Bead'
                , 'Loose Fibre Extraction'
                , 'External Wall Insulation: Less than 60m2'
                , 'External Wall Insulation: 60m2 to 85m2'
                , 'External Wall Insulation: Greater than 85m2'
                , 'Internal Wall Insulation: Vertical Surface'
                , 'External wall insulation and CWI: less than 60m2'
                , 'External wall insulation and CWI: 60m2 to 85m2'
                , 'External wall insulation and CWI: greater than 85m2'
                ]
        
        ofl_hpm = ['Basic gas heating system'
                , 'Basic oil heating system'
                , 'Full gas heating system installation'
                , 'Full oil heating system installation'
                , 'Gas boiler and controls (Basic & controls pack)'
                , 'Oil boiler and controls (Basic & controls pack)'
                , 'Hot Water Cylinder Jacket'
                ]
        

        
        ofl_s = ["Adequate Access*"
            , "No Access Details*"
            , "Cherry Picker Required*"
            , "Cherry Picker Required Details*"
            , "Mould/Mildew identified by surveyor; or reported by the applicant*"
            , "Mould/Mildew Details*"
            , "As confirmed by homeowner; property is a protected structure*"
            , "Protected Structure Details*"
            ]
        
        ofl_mr = ['Thermal Envelope - Heat loss walls, windows and doors' # all walls from floors 10 - 13 except Loadbearing/Party walls and internal walls)
                    , 'Thermal Envelope - Heat loss floor area' # surface area of floor 10 (excluding rooms as per attribute)
                    , 'Thermal Envelope - Heat loss roof area' # repeat value from above
                    , 'Heat loss Wall Area recommended for EWI and IWI' # 
                    , 'New Windows being recommended for replacement' # use value "replace_window_area" 
                    , 'Total Surface Area (m2)' # sum of first three above fields
                    , 'Total Surface Area receiving EWWR (m2)' # sum of 4th and 5th above fields
                    , 'Result %' # 7th divided by 6th * 100
                    , 'Is Major Renovation?' # yes if greater than or equal to 23%
                    ]
        
        ofl_mae = ["Number of habitable rooms in the property"
                    , "Number of wet rooms in the property"
                    , "No. of habitable/wet rooms w/ open flued appliance"
                    , "LED Bulbs: supply only (4 no.)"
                    , "Air-tightness test recommended?"
                    ]
        
        ofl_general = ['Dwelling Type*'
                        , 'Dwelling Age*'
                        , 'Age Extension 1'
                        , 'Age Extension 2'
                        , 'Asbestos Suspected'
                        , 'Asbestos Details' # only if suspected after 2000 
                        , 'Lot *' # blank for now
                        , 'Survey Date *' # project creation date in MP
                        , 'Gross floor area (m2) *'
                        , 'Number of Storeys *' # 0-9
                        , 'Room in Roof' # yes if any "dormer / room in roof" else no
                        , 'No. Single Glazed Windows *'
                        , 'No. Double Glazed Windows *'
                        , 'Property Height (m)*'
                        , 'Internet Available'
                        ]
        
        ofl_roof = ['sloped_surface_area'
                , 'ins_100_area'
                , 'ins_150_area'
                , 'ins_200_area'
                , 'ins_250_area'
                , 'ins_300_area'
                , 'storage'
                , 'new_hatch_count'
                , 'high_roof_vent_area'
                , 'Roof 1 Type*'
                , 'Other Details Roof 1*'
                , 'Sloped Ceiling Roof 1*'
                , 'Roof 1 greater than 2/3 floor area*'
                , 'Roof 1 Pitch (degrees)*'
                , 'Roof Type 1 Insulation Exists*'
                , 'Can Roof Type 1 Insulation Thickness be Measured?*'
                , 'Roof 1 Thickness (mm)*'
                , 'Roof 1 Insulation Type*'
                , 'Required per standards (mm2) *'
                , 'Existing (mm2)*'
                , 'Area of Roof Type 1 with fixed flooring (m2)*'
                , 'Folding/stair ladder in Roof Type 1*'
                , 'Fixed light in Roof Type 1*'
                , 'Downlighters in Roof Type 1*'
                , 'High power cable in Roof Type 1 (6sq/10sq or higher)*'
                , 'Roof 2 Type'
                , 'Other Details Roof 2*'
                , 'Sloped Ceiling Roof 2*'
                , 'Roof 2 greater than 2/3 floor area*'
                , 'Roof 2 Pitch (degrees)*'
                , 'Roof 2 Insulation Exists*'
                , 'Can Roof Type 2 Insulation Thickness be Measured?'
                , 'Roof 2 Thickness (mm)'
                , 'Roof 2 Insulation Type'
                , 'Roof 2 Required per standards (mm2) *'
                , 'Roof 2 Existing (mm2) *'
                , 'Area of Roof Type 2 with fixed flooring (m2)'
                , 'Folding/stair ladder in Roof Type 2'
                , 'Fixed light in Roof Type 2'
                , 'Downlighters in Roof Type 2'
                , 'High power cable in Roof Type 2 (6sq/10sq or higher)'
                , 'Roof 3 Type'
                , 'Other Details Roof 3*'
                , 'Sloped Ceiling Roof 3*'
                , 'Roof 3 greater than 2/3 floor area*'
                , 'Roof 3 Pitch (degrees)*'
                , 'Roof 3 Insulation Exists*'
                , 'Can Roof Type 3 Insulation Thickness be Measured?*'
                , 'Roof 3 Thickness (mm)*'
                , 'Roof 3 Insulation Type*'
                , 'Roof 3 Required per standards (mm2) *'
                , 'Roof 3 Existing (mm2) *'
                , 'Area of Roof Type 3 with fixed flooring (m2)*'
                , 'Folding/stair ladder in Roof Type 3*'
                , 'Fixed light in Roof Type 3*'
                , 'Downlighters in Roof Type 3*'
                , 'High power cable in Roof Type 3 (6sq/10sq or higher)*'
                , 'Roof 4 Type'
                , 'Other Details Roof 4*'
                , 'Sloped Ceiling Roof 4*'
                , 'Roof 4 greater than 2/3 floor area'
                , 'Roof 4 Pitch (degrees)*'
                , 'Roof Type 4 Insulation Exists*'
                , 'Can Roof Type 4 Insulation Thickness be Measured?*'
                , 'Roof 4 Thickness (mm)'
                , 'Roof 4 Insulation Type*'
                , 'Roof 4 Required per standards (mm2) *'
                , 'Roof 4 Existing (mm2) *'
                , 'Area of Roof Type 4 with fixed flooring (m2)*'
                , 'Folding/stair ladder in Roof Type 4*'
                , 'Fixed light in Roof Type 4*'
                , 'Downlighters in Roof Type 4*'
                , 'High power cable in Roof Type 4 (6sq/10sq or higher)*'
                , 'Suitable for Insulation *'
                , 'Roof not suitable details*'
                , 'Notes (Roof)']
        
        ofl_walls = ['Wall Type 1*'
                    , 'Other Wall type 1 Details*'
                    , 'Wall 1 wall thickness (mm)*'
                    , 'Wall 1 Insulation Present?*'
                    , 'Wall 1 Insulation Type*'
                    , "Wall 1 Fill Type*"
                    , 'Wall 1 Residual Cavity Width (mm)*'
                    , 'Can Wall type 1 Insulation Thickness be Measured?*'
                    , "If 'Yes' enter Wall type 1 insulation thickness (mm)*"
                    , 'Wall Type 2'
                    , 'Other Wall type 2 Details*'
                    , 'Wall 2 wall thickness (mm)*'
                    , 'Wall 2 Insulation Present?*'
                    , 'Wall 2 Insulation Type*'
                    , "Wall 2 Fill Type*"
                    , 'Wall 2 Residual Cavity Width (mm)*'
                    , 'Can Wall type 2 Insulation Thickness be Measured?*'
                    , "If 'Yes' enter Wall type 2 insulation thickness (mm)*"
                    , 'Wall Type 3'
                    , 'Other Wall type 3 Details*'
                    , 'Wall 3 wall thickness (mm)*'
                    , 'Wall 3 Insulation Present?*'
                    , 'Wall 3 Insulation Type*'
                    , "Wall 3 Fill Type*"
                    , 'Wall 3 Residual Cavity Width (mm)*'
                    , 'Can Wall type 3 Insulation Thickness be Measured?*'
                    , "If 'Yes' enter Wall type 3 insulation thickness (mm)*"
                    , 'Wall Type 4'
                    , 'Other Wall type 4 Details*'
                    , 'Wall 4 wall thickness (mm)*'
                    , 'Wall 4 Insulation Present?*'
                    , 'Wall 4 Insulation Type*'
                    , "Wall 4 Fill Type*"
                    , 'Wall 4 Residual Cavity Width (mm)*'
                    , 'Can Wall type 4 Insulation Thickness be Measured?*'
                    , "If 'Yes' enter Wall type 4 insulation thickness (mm)*"
                    , "Is the property suitable for wall insulation? *"
                    , "No wall insulation details *"
                    , "EWI/IWI > 25% *"
                    , 'Suitable for Draught Proofing'
                    , 'Not suitable details Draughtproofing*'
                    , "Notes (Walls)"
                    
                    , "Draught Proofing (<= 20m installed)"
                    , "Draught Proofing (> 20m installed)"
                    , "MEV 15l/s Bathroom"
                    , "MEV 30l/s Utility"
                    , "MEV 60l/s Kitchen"
                    , "New Permanent Vent"
                    , "New Background Vent"
                    , "Duct Cooker Hood"
                    , "Cavity Wall Insulation Bonded Bead"
                    , "Loose Fibre Extraction"
                    , "External Wall Insulation: Less than 60m2"
                    , "External Wall Insulation: 60m2 to 85m2"
                    , "External Wall Insulation: Greater than 85m2"
                    , "ESB alteration"
                    , "GNI meter alteration"
                    # , "GNI new connection"
                    , "New Gas Connection"
                    , "RGI Meter_No Heating"
                    , 'Internal Wall Insulation: Vertical Surface'
                    , "External wall insulation and CWI: less than 60m2"
                    , "External wall insulation and CWI: 60m2 to 85m2"
                    , "External wall insulation and CWI: greater than 85m2"
                    , 'replace_window_area'
                    , 'Notes (Windows and Doors)'
                    ]
        
        ofl_heating = ['Heating System *'
                    , 'Qualifying Boiler'
                    , 'System Age *'
                    , 'Fully Working *'
                    , 'Requires Service *'
                    , "Other Primary Heating Details *"
                    , 'Not Working Details Primary Heating *'
                    , 'Requires Service (App?)*'
                    , 'Requires Service Details Primary Heating *'
                    
                    , 'Hot Water System Exists *'
                    , 'From Primary heating system'
                    , 'From Secondary heating system'
                    , 'Electric Immersion'
                    , 'Electric Instantaneous'
                    , 'Instantaneous Combi Boiler'
                    , 'Other'
                    , 'Other HW Details *'
                    , 'HWS'
                    
                    , 'Hot Water Cylinder*'
                    , 'Insulation *'
                    , 'Condition of Lagging Jacket *'
                    , 'HWC Controls *'
                    
                    , 'Heating Systems Controls *'
                    , 'Partial Details *'
                    , 'Programmer / Timeclock *'
                    , 'Room Thermostat Number *'
                    , 'Rads Number *'
                    , 'TRVs Number *'
                    
                    , 'Suitable for Heating Measures *'
                    , 'Not suitable details*'
                    , 'Notes (Heating)'
                    , 'Secondary Heating System'
                    , 'Secondary System Age *'
                    , 'Secondary System Fully Working *'
                    , 'Secondary System Requires Service *'
                    , 'Not Working Details Secondary Heating *'
                    , 'Secondary System Requires Service (App?)*'
                    , 'Requires Service Details Secondary Heating *'





                    ]

        
        
        
        
        
        
        
        
        
        habitable_room_types = ['Kitchen', 'Dining Room', 'Living Room', 'Bedroom', 'Primary Bedroom', "Children's Bedroom", 'Study', 'Music Room']
        wet_room_types = ['Kitchen', 'Bathroom', 'Half Bathroom', 'Laundry Room', 'Toilet', 'Primary Bathroom']
        
        
        
        json_val_dict['Hot Water Cylinder Jacket'] = ''
        
        json_val_dict["Electric Storage Heater age (years)*"] = ''
        json_val_dict["Warm Air System age (years)*"] = ''
        json_val_dict["Is there Mains Gas in the Area?"] = ''
        json_val_dict["Number of habitable rooms in the property"] = 0
        json_val_dict["Number of wet rooms in the property"] = 0
        json_val_dict["No. of habitable/wet rooms w/ open flued appliance"] = 0
        
        
        
        wt_dict = {}
        wt_dict['ext_wall_area_gross'], exploded_wall_dict = exterior_walls(root)
        # print("wt_dict['ext_wall_area_gross']", ':', wt_dict['ext_wall_area_gross'])
        wt_dict['gross'] = 0
        wt_dict['total'] = 0
        wt_dict['total_window_a'] = 0
        wt_dict['total_party_a'] = 0
        wt_dict['External.Wall.not.recieving.EWI.or.IWI'] = 0
        nwa_temp_dict = {}
        for floor in nwa_dict.keys():
            for room in nwa_dict[floor]:
                for wall in nwa_dict[floor][room]:
                    if 10 <= int(floor) <= 13:
                        # print(nwa_dict[floor][room][wall]['total_window_a'])
                        wt_dict['total_window_a'] += nwa_dict[floor][room][wall]['total_window_a']
                        if 'loadBearingWall' in list(nwa_dict[floor][room][wall].keys()):
                            if nwa_dict[floor][room][wall]['loadBearingWall'] == '1':
                                wt_dict['total_party_a'] += nwa_dict[floor][room][wall]['a']
                        # wt_dict['total_party_a'] += nwa_dict[floor][room][wall]['total_window_a']
                    if 10 <= int(floor) <= 13:
                        if 'type' in list(nwa_dict[floor][room][wall].keys()): # indicates wall recommended for insulation
                            # print("nwa_dict[floor][room][wall]['a']", ':', nwa_dict[floor][room][wall]['a'])
                            wt_dict['gross'] += float(nwa_dict[floor][room][wall]['a'])
                    for key in nwa_dict[floor][room][wall]:
                        name = 'floor ' + floor + '_' + room + '_wall ' + str(wall) + '_' + key
                        nwa_temp_dict[name] = nwa_dict[floor][room][wall][key]
                        if key == 'type':
                            if nwa_dict[floor][room][wall][key] == 'External.Wall.not.recieving.EWI.or.IWI':
                                req_area = nwa_dict[floor][room][wall]['a']
                            else:
                                req_area = nwa_dict[floor][room][wall]['net_a']
                            # print(name, 'req_area', ':', str(req_area))
                            if nwa_dict[floor][room][wall][key] in wt_dict.keys():
                                wt_dict[nwa_dict[floor][room][wall][key]] += req_area
                            else:
                                wt_dict[nwa_dict[floor][room][wall][key]] = req_area
                            wt_dict['total'] += req_area
        
        # print(nwa_temp_dict)
        
        
        
        
        wt_dict['ext_wall_area_net'] = wt_dict['ext_wall_area_gross'] - wt_dict['total_party_a']
        # print(wt_dict)
        
        # (if any value blank then 0)

        
        
        # print('nwa_dict[10]:')
        # print(nwa_dict['10'])
        # print('exploded_wall_dict:')
        # print(exploded_wall_dict)
        
        
        json_data = json.dumps(
        nwa_dict
        )
        j = r"d:\USERS\gshortall\Documents\Shortcut\investigate_A.json"
        # with open(j, "w") as investigate_file:
            # investigate_file.write(json_data)
        
        json_data = json.dumps(
        exploded_wall_dict
        )
        j = r"d:\USERS\gshortall\Documents\Shortcut\investigate_B.json"
        # with open(j, "w") as investigate_file:
            # investigate_file.write(json_data)
        
        print(nwa_dict.keys())
        for floor in list(nwa_dict.keys()):
            if floor != "10":
                continue
            print(floor)
            for room in nwa_dict[floor].keys():
                for wall in nwa_dict[floor][room].keys():
                    x1 = nwa_dict[floor][room][wall]["x1"]
                    y1 = nwa_dict[floor][room][wall]["y1"]
                    x2 = nwa_dict[floor][room][wall]["x2"]
                    y2 = nwa_dict[floor][room][wall]["y2"]
                    x3 = nwa_dict[floor][room][wall]["x3"]
                    y3 = nwa_dict[floor][room][wall]["y3"]
                    x4 = nwa_dict[floor][room][wall]["x4"]
                    y4 = nwa_dict[floor][room][wall]["y4"]
                    
                    # print(str(x1), str(y1))
                    # print(room, str(x1), str(y1), str(x3), str(y3))
                    # print(room, str(x2), str(y2), str(x4), str(y4))
                    # print(room, str(x3), str(y3))
                    # print(room, str(x4), str(y4))

        # else:
            # print('WARNING: No floor 10')
        
                    # for floor in exploded_wall_dict:
                    # print(floor)
                    d_min = 1
                    w_candidates = []
                    for wall in exploded_wall_dict[floor].keys():
                        # print(exploded_wall_dict[floor][wall])
                        x5 = exploded_wall_dict[floor][wall]["x1"]
                        y5 = exploded_wall_dict[floor][wall]["y1"]
                        x6 = exploded_wall_dict[floor][wall]["x2"]
                        y6 = exploded_wall_dict[floor][wall]["y2"]
                        w_type = exploded_wall_dict[floor][wall]["type"]
                        # print(wall, str(x5), str(y5))
                        # print(wall, str(x6), str(y6))
                        
                        d = cart_distance((x3, y3), (x5, y5))
                        if d == d_min:
                            w_candidates.append(wall)
                        if d < d_min:
                            d_min = d
                            w_candidates.append(wall)
                        d = cart_distance((x3, y3), (x6, y6))
                        if d == d_min:
                            w_candidates.append(wall)
                        if d < d_min:
                            d_min = d
                            w_candidates.append(wall)
                            
                    # print('d_min', ':', d_min)
                    print('w_candidates', ':', w_candidates)

        
        
        
        

        
        
        
        json_url = "https://cloud.magicplan.app/api/v2/plans/forms/" + str(xml_val_dict['id'])
        try:
            request = urllib.request.Request(json_url, headers=headers)
        except:
            output = traceback.format_exc()
            print(output)
        # finally:
            
        JSON = urllib.request.urlopen(request).read()
        JSON = json.loads(JSON)
        print('len(JSON["data"]): ', len(JSON["data"]))
        df = pd.DataFrame(JSON["data"])
        
        
        

        
        json_val_dict["Existing (mm2)*"] = int(0)
        json_val_dict['No. Single Glazed Windows *'] = 0
        # json_ref_dict = {}
        alt_double_glazed_count = 0
        



        replace_windows = []


        cylinder_stat = False
        json_val_dict['Heating System *'] = 'N/A'
        json_val_dict['Secondary Heating System'] = 'N/A'
        json_val_dict['HWS'] = ''
        json_val_dict['HWC Controls *'] = 'None'
        single_glazed_windows = []
        programmers = []
        room_thermostats = []
        
        
        json_val_dict["ESB alteration"] = ''
        json_val_dict["GNI meter alteration"] = ''
        json_val_dict["RGI Meter_No Heating"] = ''
        json_val_dict["New Gas Connection"] = ''
        
        # esb_alterations = []
        # gni_alterations = []
        # rgi_meter_no_heating = []
        # new_gas_connection = []
        
        
        
        condensing = ''
        linked_stove_bb = ''
        other_heating_notes = ''
        
        json_val_dict["Duct Cooker Hood"] = 0
        
        balanced_flues = []
        req_lagging_jackets = []
        slope_dict = {}
        roof_type_dict = {}
        h = {}
        h_index = 0
        for datum in JSON["data"]:
            # print(datum["symbol_name"])
            if 'Combi Boiler' in datum["symbol_name"]:
                json_val_dict['Hot Water System Exists *'] = True
                
            if 'Hot Water Cylinder' in datum["symbol_name"]:
                json_val_dict['Hot Water Cylinder*'] = True
                json_val_dict['Hot Water System Exists *'] = True
                if 'Bad' in datum["symbol_name"] or 'No Insulation' in datum["symbol_name"]:
                    json_val_dict['Condition of Lagging Jacket *'] = 'Bad'
                    req_lagging_jackets.append(datum["symbol_instance_id"])
                else:
                    json_val_dict['Condition of Lagging Jacket *'] = 'Good'
                
                if 'Lagging Jacket' in datum["symbol_name"]:
                    json_val_dict['Insulation *'] = 'Lagging Jacket'
                if 'Factory Fitted' in datum["symbol_name"]:
                    json_val_dict['Insulation *'] = 'Factory Fitted'
                if 'No Insulation' in datum["symbol_name"]:
                    json_val_dict['Insulation *'] = 'No Insulation'
                    
                    
                    
                
                
            if datum["symbol_name"] == "Programmer":
                programmers += datum["symbol_instance_id"]
            if datum["symbol_name"] == "Room Thermostat":
                room_thermostats += datum["symbol_instance_id"]


            if datum["symbol_name"] == "ESB alteration":
                json_val_dict["ESB alteration"] = 1
                # esb_alterations.append(datum["symbol_instance_id"])
            if datum["symbol_name"] == "GNI meter alteration":
                json_val_dict["GNI meter alteration"] = 1
                # gni_alterations.append(datum["symbol_instance_id"])
            
            if datum["symbol_name"] == "New Gas Connection":
                json_val_dict["New Gas Connection"] = 1
                # new_gas_connection.append(datum["symbol_instance_id"])
            if datum["symbol_name"] == "RGI Meter_No Heating":
                json_val_dict["RGI Meter_No Heating"] = 1
                # rgi_meter_no_heating.append(datum["symbol_instance_id"])
            
            for form in datum["forms"]:
                for section in form["sections"]:
                    for field in section["fields"]:
                        v = ''
                        if field["value"]["value"] == None:
                            vals = [val["value"] for val in field["value"]["values"]]
                            for val in vals:
                                v += val
                                v += '<BR>'
                        else:
                            v = field["value"]["value"]
                        # print(field["label"], ':', v)
                        json_val_dict[field["label"]] = v
                        
                        im = field["label"].replace(' *', '')
                        im = im.replace('*', '')
                        json_val_dict[im] = v # if the field is marked as mandatory this creates a non-marked copy with the same answer, note this doesn't solve all issues though
                        

                            
                        if field["label"] == "Is it a Balanced Flue?" and field["value"]["value"] == False:
                            balanced_flues.append(datum["symbol_instance_id"])
                        
                        if field["label"] == "Heating designation on Portal*" and field["value"]["value"] == "Primary":
                            json_val_dict['Heating System *'] = datum["symbol_name"]
                        if field["label"] == "Heating designation on Portal*" and field["value"]["value"] == "Secondary":
                            json_val_dict['Secondary Heating System'] = datum["symbol_name"]
                        if field["label"] == "Is there a timer?" and field["value"]["value"] == True:
                            json_val_dict['HWC Controls *'] = 'Independent Timer'
                        if field["label"] == "Is there a cylinder stat?" and field["value"]["value"] == True:
                            json_val_dict['HWC Controls *'] = 'Cylinder Thermostat'
                            cylinder_stat = True
                        if field["label"] == "Is the cylinder heated from the primary heating system?":
                            if field["value"]["value"] == True:
                                json_val_dict['From Primary heating system'] = True
                                if 'From Primary heating system' not in json_val_dict['HWS']:
                                    json_val_dict['HWS'] += ('From Primary heating system' + '<BR>')
                        if field["label"] == "Is the cylinder heated from the secondary heating system?":
                            if field["value"]["value"] == True:
                                json_val_dict['From Secondary heating system'] = True
                                if 'From Secondary heating system' not in json_val_dict['HWS']:
                                    json_val_dict['HWS'] += ('From Secondary heating system' + '<BR>')
                        if field["label"] == "Is there an electric immersion?" and field["value"]["value"] == True:
                            json_val_dict['Electric Immersion'] = True
                            if 'Electric Immersion' not in json_val_dict['HWS']:
                                json_val_dict['HWS'] += ('Electric Immersion' + '<BR>')
                        if field["label"] == "How is the cylinder heated? (Do not include immersion)" and field["value"]["has_value"] == True:
                            json_val_dict["Other HW Details *"] = field["value"]["value"]
                            if 'Other' not in json_val_dict['HWS']:
                                json_val_dict['HWS'] += ('Other' + '<BR>')
                            
                        if field['label'] == "Heating notes*":
                            if datum["symbol_name"] not in [json_val_dict['Heating System *'], json_val_dict['Secondary Heating System']]:
                                n = 'Other System: ' + field["value"]["value"] + '<BR>'
                                if n not in other_heating_notes:
                                    other_heating_notes += n



                        if field["label"] == "Existing Roof Ventilation (mm2)*":
                            if not field["value"]["value"].isdigit():
                                continue
                            json_val_dict["Existing (mm2)*"] += int(field["value"]["value"])
                        
                        if field["label"] == "Is the window Single glazed?" and field["value"]["value"] == True:
                            # json_val_dict['No. Single Glazed Windows *'] += 1
                            # single_glazed_windows.append(datum["symbol_instance_id"])
                            single_glazed_windows.append(datum["symbol_instance_id"])
                        if field["label"] == "Is it being recommended for replacement?" and field["value"]["value"] == True:
                            replace_windows.append(datum["symbol_instance_id"]) 
                        if datum["symbol_instance_id"] in xml_ref_dict.keys():
                            if field["label"] == "Roof Type*":
                                # print('before: ', xml_ref_dict[datum["symbol_instance_id"]])
                                xml_ref_dict[datum["symbol_instance_id"]] = field["value"]["value"]
                                # print('after: ', xml_ref_dict[datum["symbol_instance_id"]])
                                for n in range(1, 5):
                                    if field["value"]["value"] == f"Roof Type {n}":
                                        roof_type_dict[datum["symbol_instance_id"]] = n
                    g = {} # temporary dictionary containing the answers to all questions in a section (roof-related)
                    for field in section["fields"]:
                        g[field["label"]] = field["value"]["value"]
                    for n in range(1, 5):
                        if "Roof Type*" in g.keys() and f"Roof Type {n} Sloping Ceiling Suitable for Insulation*" in g.keys():
                            # print('g: ', g)
                            if g["Roof Type*"] == "Sloped Ceiling" and g[f"Roof Type {n} Sloping Ceiling Suitable for Insulation*"] == True:
                                if f"Roof {n} Pitch (degrees)*" in json_val_dict.keys():
                                    pitch = json_val_dict[f"Roof {n} Pitch (degrees)*"]
                                else:
                                    pitch = 30
                                slope_dict[datum["symbol_instance_id"]] = pitch
        # print('balanced_flues', ':', str(balanced_flues))

        print(1)


        json_val_dict['Notes (Heating)'] = ''
        # Go through Forms again to get values for Primary & Secondary Heating Systems
        for datum in JSON["data"]:
            if datum["symbol_name"] == json_val_dict['Heating System *']:
                print(datum["symbol_name"])
                for form in datum["forms"]:
                    for section in form["sections"]:
                        for field in section["fields"]:
                            print(field["label"])
                            if 'age (years)' in field['label']:
                                json_val_dict['System Age *'] = field["value"]["value"]
                            if field['label'] == 'Fully Working?':
                                json_val_dict['Fully Working *'] = field["value"]["value"]
                            if 'require service?' in field['label']:
                                json_val_dict['Requires Service *'] = field["value"]["value"]
                            if field['label'] == '':
                                json_val_dict["Other Primary Heating Details *"] = field["value"]["value"]
                            if field['label'] == "Not working details*":
                                # print(field["value"]["value"])
                                json_val_dict['Not Working Details Primary Heating *'] = field["value"]["value"]
                            # if field['label'] == 'Does the appliance require service?':
                                # json_val_dict['Requires Service (App?)*'] = field["value"]["value"]
                            if field['label'] == 'Service details':
                                json_val_dict['Requires Service Details Primary Heating *'] = field["value"]["value"]
                            if field['label'] == "Is the boiler Condensing?*" and field["value"]["value"] == True:
                                condensing = True
                            if field['label'] == "Interlinked with?" and field["value"]["value"] == "Stove + Back Boiler":
                                linked_stove_bb = True
                            if field['label'] == "Heating notes*" and field["value"]["value"] != None:
                                n = 'Primary System: ' + field["value"]["value"] + '<BR>'
                                # print(n)
                                if n not in json_val_dict['Notes (Heating)']:
                                    json_val_dict['Notes (Heating)'] += n
                                    # print('Notes (Heating)', ':', json_val_dict['Notes (Heating)'])
            # print(2)
            
            if datum["symbol_name"] == json_val_dict['Secondary Heating System']:
                for form in datum["forms"]:
                    for section in form["sections"]:
                        for field in section["fields"]:
                            if 'age (years)' in field['label']:
                                json_val_dict['Secondary System Age *'] = field["value"]["value"]
                            if field['label'] == 'Fully Working?':
                                json_val_dict['Secondary System Fully Working *'] = field["value"]["value"]
                            if 'require service?' in field['label']:
                                json_val_dict['Secondary System Requires Service *'] = field["value"]["value"]
                            if field['label'] == '':
                                json_val_dict["Other Primary Heating Details *"] = field["value"]["value"]
                            if field['label'] == 'Not working details*':
                                json_val_dict['Not Working Details Secondary Heating *'] = field["value"]["value"]
                            # if field['label'] == 'Does the appliance require service?':
                                # json_val_dict['Secondary System Requires Service (App?)*'] = field["value"]["value"]
                            if field['label'] == 'Service details':
                                json_val_dict['Requires Service Details Secondary Heating *'] = field["value"]["value"]
                            if field['label'] == "Heating notes*":
                                n = 'Secondary System: ' + field["value"]["value"] + '<BR>'
                                if n not in json_val_dict['Notes (Heating)']:
                                    json_val_dict['Notes (Heating)'] += n
                    
        # print('Notes (Heating)', ':', json_val_dict['Notes (Heating)'])
        
        json_val_dict['Programmer / Timeclock *'] = 0
        json_val_dict['Room Thermostat Number *'] = 0
        json_val_dict['Rads Number *'] = 0
        json_val_dict['TRVs Number *'] = 0
        
        if other_heating_notes != '':
            json_val_dict['Notes (Heating)'] = json_val_dict['Notes (Heating)'] + other_heating_notes
        



        
        json_url = "https://cloud.magicplan.app/api/v2/plans/statistics/" + str(xml_val_dict['id'])
        request = urllib.request.Request(json_url, headers=headers)
        JSON = urllib.request.urlopen(request).read()
        JSON = json.loads(JSON)
        
        # df = pd.DataFrame(JSON["data"])
        
        df = pd.DataFrame(JSON["data"]["project_statistics"])
        # print()
        

        
        
        json_val_dict['No. Double Glazed Windows *'] = 0
        json_val_dict['Number of Storeys *'] = 0
        json_val_dict['Gross floor area (m2) *'] = 0.00
        new_draughtproofing = 0
        sum_low = 0
        sum_high = 0
        roof_area_sum = 0
        slope_roof_area_sum = 0
        new_hatch_count = 0
        
        json_val_dict["MEV 15l/s Bathroom"] = 0
        json_val_dict["MEV 30l/s Utility"] = 0
        json_val_dict["MEV 60l/s Kitchen"] = 0
        json_val_dict["New Permanent Vent"] = 0
        json_val_dict["New Background Vent"] = 0
        
        json_val_dict["Cavity Wall Insulation Bonded Bead"] = 0
        json_val_dict["Loose Fibre Extraction"] = 0
        external_wall_insulation = 0
        json_val_dict["External Wall Insulation: Less than 60m2"] = 0
        json_val_dict["External Wall Insulation: 60m2 to 85m2"] = 0
        json_val_dict["External Wall Insulation: Greater than 85m2"] = 0
        json_val_dict["Internal Wall Insulation: Vertical Surface"] = 0
        external_wall_insulation_and_cwi = 0
        json_val_dict["External wall insulation and CWI: less than 60m2"] = 0
        json_val_dict["External wall insulation and CWI: 60m2 to 85m2"] = 0
        json_val_dict["External wall insulation and CWI: greater than 85m2"] = 0

        # , 'Heating Systems Controls *'
        # , 'Partial Details *'
        
        
        rooms_with_balanced_flues = []
        req_lagging_jacket_count = 0
        
        
        json_val_dict['Thermal Envelope - Heat loss floor area'] = 0
        json_val_dict['replace_window_area'] = 0
        # print(xml_ref_dict)
        # print('replace_windows', ':', replace_windows)
        for floor in df.floors:
            if int(xml_ref_dict[floor["uid"]]) == 20:
                external_wall_insulation += floor["area_with_interior_walls_only"]
            if int(xml_ref_dict[floor["uid"]]) == 21:
                json_val_dict["Cavity Wall Insulation Bonded Bead"] += floor["area_with_interior_walls_only"]
            if int(xml_ref_dict[floor["uid"]]) == 22:
                json_val_dict["Internal Wall Insulation: Vertical Surface"] += floor["area_with_interior_walls_only"]
            if int(xml_ref_dict[floor["uid"]]) == 23:
                json_val_dict["Loose Fibre Extraction"] += floor["area_with_interior_walls_only"]
            if int(xml_ref_dict[floor["uid"]]) == 24:
                external_wall_insulation_and_cwi += floor["area_with_interior_walls_only"]
            if -1 <= int(xml_ref_dict[floor["uid"]]) <= 9:
                for room in floor["rooms"]:
                    
                    
                    for furniture in room["furnitures"]:
                        if furniture["uid"] in req_lagging_jackets:
                            # json_val_dict['Hot Water Cylinder Jacket'] += 1
                            req_lagging_jacket_count += 1
                        if furniture["name"] in ["Radiator", "Radiator with TRV", "Water Radiator"]:
                            json_val_dict['Rads Number *'] += 1
                        if furniture["name"] == "Radiator with TRV":
                            json_val_dict['TRVs Number *'] += 1
                        if furniture["name"] == "Electric Instantaneous":
                            json_val_dict['Electric Instantaneous'] = True
                        if furniture["name"] in ["Gas Combi Boiler", "Oil Combi Boiler"]:
                            json_val_dict['Instantaneous Combi Boiler'] = True
                    for wall_item in room["wall_items"]:
                        if wall_item["name"] == "Room Thermostat":
                            json_val_dict['Room Thermostat Number *'] += 1
                        if wall_item["name"] == "Programmer":
                            json_val_dict['Programmer / Timeclock *'] += 1
                            
                            
            
            
            
            
            if int(xml_ref_dict[floor["uid"]]) == 10:
                json_val_dict['Thermal Envelope - Heat loss floor area'] = floor["area_with_interior_walls_only"]
                for room in floor["rooms"]:
                    if room["uid"] in xml_ref_dict['exclude_rooms']:
                        json_val_dict['Thermal Envelope - Heat loss floor area'] -= room["area_with_interior_walls_only"]

            if -1 <= int(xml_ref_dict[floor["uid"]]) <= 9:
                json_val_dict['Number of Storeys *'] += 1
                json_val_dict['Gross floor area (m2) *'] += floor["area_with_interior_walls_only"]
                json_val_dict['No. Double Glazed Windows *'] += floor["window_count"]

                for room in floor["rooms"]:
                    # print(xml_ref_dict[room["uid"]])
                    if room["uid"] in xml_ref_dict['habitable_rooms']:
                        json_val_dict["Number of habitable rooms in the property"] += 1
                    if room["uid"] in xml_ref_dict['wet_rooms']:
                        json_val_dict["Number of wet rooms in the property"] += 1
                    if room["uid"] in xml_ref_dict['exclude_rooms']:
                        json_val_dict['Gross floor area (m2) *'] -= room["area_with_interior_walls_only"]
                        
                    for wall_item in room["wall_items"]:
                        if wall_item["uid"] in replace_windows:
                            json_val_dict['replace_window_area'] += (wall_item["width"] * wall_item["height"])
                        if wall_item["uid"] in single_glazed_windows:
                            # print(wall_item["uid"] + " found in " + str(single_glazed_windows))
                            json_val_dict['No. Single Glazed Windows *'] += 1
                            
                    for furniture in room["furnitures"]:
                        # print(furniture["name"])
                        if furniture["uid"] in balanced_flues:
                            rooms_with_balanced_flues.append(room["uid"])
                            
                        if furniture["name"] == "New Draughtproofing":
                            new_draughtproofing += 1
                        if furniture["name"] == "New Mechanical Vent":
                            # print(xml_ref_dict[room["uid"]])
                            if xml_ref_dict[room["uid"]] in ['Bathroom', 'Half Bathroom', 'Toilet']:
                                json_val_dict["MEV 15l/s Bathroom"] += 1
                            if xml_ref_dict[room["uid"]] in ['Laundry Room']:
                                json_val_dict["MEV 30l/s Utility"] += 1
                            if xml_ref_dict[room["uid"]] in ['Kitchen']:
                                json_val_dict["MEV 60l/s Kitchen"] += 1
                        if furniture["name"] == "New Permanent Vent":
                            json_val_dict["New Permanent Vent"] += 1
                        if furniture["name"] == "New Background Vent":
                            json_val_dict["New Background Vent"] += 1
                        if furniture["name"] == "Duct Cooker Hood":
                            json_val_dict["Duct Cooker Hood"] += 1
                        if furniture["name"] == "Duct Mechanical Extract Vent":
                            json_val_dict["Duct Cooker Hood"] += 1
                        if furniture["name"] == "New Hatch": # only found in -1 to 9 and Roof?
                            new_hatch_count += 1
                
                
            for room in floor["rooms"]:
                for furniture in room["furnitures"]:
                    if furniture["name"] == "New Low Level Roof Ventilation": # only found in Roof?
                        sum_low += float(furniture["width"])
                    if furniture["name"] == "New High Level Roof Ventilation": # only found in Roof?
                        sum_high += float(furniture["width"])
        
        
        json_val_dict['No. Double Glazed Windows *'] -= json_val_dict['No. Single Glazed Windows *']

        for room in rooms_with_balanced_flues:
            if room in (xml_ref_dict['habitable_rooms'] + xml_ref_dict['wet_rooms']):
                json_val_dict["No. of habitable/wet rooms w/ open flued appliance"] += 1
        


        
        if new_draughtproofing == 0:
            json_val_dict["Draught Proofing (<= 20m installed)"] = 'N/A'
            json_val_dict["Draught Proofing (> 20m installed)"] = 'N/A'
        if 1 <= new_draughtproofing <= 3:
            json_val_dict["Draught Proofing (<= 20m installed)"] = 1
            json_val_dict["Draught Proofing (> 20m installed)"] = 'N/A'
        if new_draughtproofing >= 4:
            json_val_dict["Draught Proofing (<= 20m installed)"] = 'N/A'
            json_val_dict["Draught Proofing (> 20m installed)"] = 1
        
        
        json_val_dict["External Wall Insulation: Less than 60m2"] = round(external_wall_insulation) if external_wall_insulation <= 60 else 'N/A'
        
        json_val_dict["External Wall Insulation: 60m2 to 85m2"] = round(external_wall_insulation) if 60 < external_wall_insulation <= 85  else 'N/A'
        
        json_val_dict["External Wall Insulation: Greater than 85m2"] = round(external_wall_insulation) if external_wall_insulation > 85  else 'N/A'


        json_val_dict["External wall insulation and CWI: less than 60m2"] = round(external_wall_insulation_and_cwi) if external_wall_insulation_and_cwi <= 60 else 'N/A'
        
        json_val_dict["External wall insulation and CWI: 60m2 to 85m2"] = round(external_wall_insulation_and_cwi) if 60 < external_wall_insulation_and_cwi <= 85 else 'N/A'
        
        json_val_dict["External wall insulation and CWI: greater than 85m2"] = round(external_wall_insulation_and_cwi) if external_wall_insulation_and_cwi > 85 else 'N/A'
            
        
        
        
        for floor in df.floors:
            if int(xml_ref_dict[floor["uid"]]) == 1000: # i.e. type "Roof"
                # roof_area_total = floor["area_with_interior_walls_only"]
                # print('roof_area_total (unused variable): ', roof_area_total)
                
                slope_roof_area_sum = 0
                for n in range(1, 5):
                    json_val_dict[f"roof_{n}_area"] = 0

                for room in floor["rooms"]:
                    # print('room["uid"]', ':', room["uid"])
                    # print(xml_ref_dict[room["uid"]])
                    # print('area:', room["area_with_interior_walls_only"])
                    
                    if room["uid"] in roof_type_dict.keys():
                        n = roof_type_dict[room["uid"]]
                        json_val_dict[f"roof_{n}_area"] += room["area_with_interior_walls_only"]
                    if room["uid"] in slope_dict.keys():
                        this_slope_area = room["area_with_interior_walls_only"] / cos(slope_dict[room["uid"]]/57.2958)
                        slope_roof_area_sum += this_slope_area
                    for furniture in room["furnitures"]:
                        if furniture["name"] == "New Hatch": # only found in -1 to 9 and Roof?
                            new_hatch_count += 1
        # print(json_val_dict)
        
        
        roof_general(json_val_dict) # adds a number of fields contingent on the above
        
        walls_general(json_val_dict)
        
        
        json_val_dict['Gross floor area (m2) *'] = round(json_val_dict['Gross floor area (m2) *'], 2)
        json_val_dict['Required per standards (mm2) *'] = round(sum_low * 10000)
        

        
        
        HSC_count = 0
        # Yes to Cylinder stat in form for hot water cylinder
        if cylinder_stat == True:
            HSC_count += 1
        # Object count of "Programmer" >0
        if json_val_dict['Programmer / Timeclock *'] > 0:
            HSC_count += 1
        # Object count of "Room Thermostat" >0
        if json_val_dict['Room Thermostat Number *'] > 0:
            HSC_count += 1
        # % of Radiators/Rads with TRVs >=50%
        if json_val_dict['TRVs Number *'] > 0:
            if json_val_dict['Rads Number *'] > 0: # should be redundant due to preceeding condition... (apparently not!)
                r = json_val_dict['TRVs Number *'] / json_val_dict['Rads Number *']
                if r >= 0.5:
                    HSC_count += 1
        else:
            r = 0
        # print('json_val_dict["TRVs Number *"]', ':', json_val_dict['TRVs Number *'])
        # print('json_val_dict["Rads Number *"]', ':', json_val_dict['Rads Number *'])
        # print('r', ':', str(r))
        # print('HSC_count', ':', str(HSC_count))
        
        if HSC_count == 0:
            json_val_dict['Heating Systems Controls *'] = 'No Controls'
        if 1 <= HSC_count <= 3:
            cylinder_stat_yn = "Yes" if cylinder_stat == True else "No"
            percentage = str(round(r * 100)) + '%'
            json_val_dict['Heating Systems Controls *'] = 'Partial Controls'
            json_val_dict["Partial Details *"] = 'No of Programmers: ' + str(json_val_dict['Programmer / Timeclock *']) + "<BR>" + 'No of Room Stats: ' + str(json_val_dict['Room Thermostat Number *']) + "<BR>" + '% of Radiators  with TRVs: ' + percentage + "<BR>" + 'Cylinder Stat?: ' + cylinder_stat_yn
        if HSC_count == 4:
            json_val_dict['Heating Systems Controls *'] = 'Full zone control to spec'
            
        
        
        print(3)
        
        
        
        # Work Order Recommendation (Roof):
        json_val_dict['sloped_surface_area'] = round(slope_roof_area_sum) if round(slope_roof_area_sum) != 0 else 'N/A'
        
        print('sfi_dict', ':', json_val_dict["sfi_dict"])
        json_val_dict['Attic Storage (5m2)'] = 0
        for t in [100, 150, 200, 250, 300]:
            if str(t) in json_val_dict["sfi_dict"].keys():
                json_val_dict[f'ins_{t}_area'] = round(json_val_dict["sfi_dict"][str(t)])
                json_val_dict['Attic Storage (5m2)'] = 1
        
        
        json_val_dict['new_hatch_count'] = new_hatch_count
        json_val_dict['high_roof_vent_area'] = round(sum_high * 5000)
        # json_val_dict['low_roof_vent_area'] = json_val_dict['Required per standards (mm2) *']
        
        for n in range(1, 5):
            if f"Wall Type {n}" in json_val_dict.keys():
                json_val_dict[f"Wall Type {n}"] = json_val_dict[f"Wall Type {n}"] if json_val_dict[f"Wall Type {n}"] != '' else 'N/A'
            if f"Wall Type {n} Residual Cavity Width (mm)" in json_val_dict.keys():
                json_val_dict[f"Wall Type {n} Residual Cavity Width (mm)"] = json_val_dict[f"Wall Type {n} Residual Cavity Width (mm)"] if json_val_dict[f"Wall Type {n} Residual Cavity Width (mm)"] != 0 else 'N/A'
            if f"Wall Type {n} Fill Type" in json_val_dict.keys():
                json_val_dict[f"Wall Type {n} Fill Type"] = json_val_dict[f"Wall Type {n} Fill Type"] if json_val_dict[f"Wall Type {n} Fill Type"] != 0 else 'N/A'
        
        
        
        
        # Fixed Values: (should these only be added to output_dict as they are not JSON values?)
        json_val_dict['Roof 2 Required per standards (mm2) *'] = 0
        json_val_dict['Roof 2 Existing (mm2) *'] = 0
        json_val_dict['Roof 3 Required per standards (mm2) *'] = 0
        json_val_dict['Roof 3 Existing (mm2) *'] = 0
        json_val_dict['Roof 4 Required per standards (mm2) *'] = 0
        json_val_dict['Roof 4 Existing (mm2) *'] = 0
        
        
        wt_dict['EWI/IWI'] = round(wt_dict['total'] - wt_dict['External.Wall.not.recieving.EWI.or.IWI'])
        
        # (if any value blank then 0)
        json_val_dict['Thermal Envelope - Heat loss walls, windows and doors'] = round(wt_dict['gross'], 2) # round(wt_dict['ext_wall_area_net'], 2)
        json_val_dict['Thermal Envelope - Heat loss floor area'] = round(json_val_dict['Thermal Envelope - Heat loss floor area'], 2)
        json_val_dict['Thermal Envelope - Heat loss roof area'] = json_val_dict['Thermal Envelope - Heat loss floor area']
        json_val_dict['Heat loss Wall Area recommended for EWI and IWI'] = round(wt_dict['EWI/IWI'], 2) # round(wt_dict['total'], 2)
        json_val_dict['New Windows being recommended for replacement'] = round(json_val_dict['replace_window_area'], 2) if json_val_dict['replace_window_area'] > 0.5 else 0
        json_val_dict['Total Surface Area (m2)'] = round(json_val_dict['Thermal Envelope - Heat loss walls, windows and doors'] + (2 * json_val_dict['Thermal Envelope - Heat loss floor area']), 2)
        json_val_dict['Total Surface Area receiving EWWR (m2)'] = round(float(wt_dict['EWI/IWI']) + float(json_val_dict['replace_window_area']), 2)
        json_val_dict['Result %'] = round(100 * (json_val_dict['Total Surface Area receiving EWWR (m2)'] / json_val_dict['Total Surface Area (m2)']), 2) if json_val_dict['Total Surface Area (m2)'] > 0 else 0
        json_val_dict['Is Major Renovation?'] = 'Yes' if json_val_dict['Result %'] >= 23 else 'No'
        
        json_val_dict['THERMAL ENVELOPE OF BUILDING AREA'] = json_val_dict['Total Surface Area (m2)']
        json_val_dict['TOTAL SURFACE AREA FOR MAJOR RENOVATION WORKS'] = json_val_dict['Total Surface Area receiving EWWR (m2)']
        json_val_dict['Total surface area for MR works / Thermal Envelope'] = json_val_dict['Result %']
        json_val_dict['WARMER HOMES MAJOR RENOVATION RESULT'] = json_val_dict['Is Major Renovation?']
        
        
        
        json_val_dict['EWI/IWI > 25% *'] = json_val_dict['Is Major Renovation?']
        
        

        
        json_val_dict["ESB alteration"] = json_val_dict["ESB alteration"] if json_val_dict["ESB alteration"] != 0 else ''
        json_val_dict["GNI meter alteration"] = json_val_dict["GNI meter alteration"] if json_val_dict["GNI meter alteration"] != 0 else ''
        
        
        

        # print(1)
        
        
        
        
        for pm in ofl_pm:
            # print(pm)
            if pm not in json_val_dict.keys():
                json_val_dict[pm] = ''
            # print('json_val_dict[pm]', ':', json_val_dict[pm])
        
        
        # print('ins_200_area', ':', ins_200_area)
        # print("json_val_dict['ins_200_area']", ':', json_val_dict['ins_200_area'])
        
        json_val_dict['Internal Wall Insulation: Sloped or flat (horizontal) surface'] = json_val_dict['sloped_surface_area']
        if 'ins_100_area' in json_val_dict.keys():
            json_val_dict['Attic (Loft) Insulation 100 mm top-up'] = json_val_dict['ins_100_area']
        if 'ins_150_area' in json_val_dict.keys():
            json_val_dict['Attic (Loft) Insulation 150 mm top-up'] = json_val_dict['ins_150_area']
        if 'ins_200_area' in json_val_dict.keys():
            json_val_dict['Attic (Loft) Insulation 200 mm top-up'] = json_val_dict['ins_200_area']
        if 'ins_250_area' in json_val_dict.keys():
            json_val_dict['Attic (Loft) Insulation 250 mm top up'] = json_val_dict['ins_250_area']
        if 'ins_300_area' in json_val_dict.keys():
            json_val_dict['Attic (Loft) Insulation 300 mm'] = json_val_dict['ins_300_area']
        
        
        # json_val_dict['Cavity Wall Insulation Bonded Bead']
        # json_val_dict['Loose Fibre Extraction']
        # json_val_dict['External Wall Insulation: Less than 60m2']
        # json_val_dict['External Wall Insulation: 60m2 to 85m2']
        # json_val_dict['External Wall Insulation: Greater than 85m2']
        # json_val_dict['Internal Wall Insulation: Vertical Surface']
        # json_val_dict['External wall insulation and CWI: less than 60m2']
        # json_val_dict['External wall insulation and CWI: 60m2 to 85m2']
        # json_val_dict['External wall insulation and CWI: greater than 85m2']
        json_val_dict['Basic gas heating system'] = ''
        json_val_dict['Basic oil heating system'] = ''
        json_val_dict['Full gas heating system installation'] = ''
        json_val_dict['Full oil heating system installation'] = ''
        json_val_dict['Gas boiler and controls (Basic & controls pack)'] = ''
        json_val_dict['Oil boiler and controls (Basic & controls pack)'] = ''
        
        
        
        # print('req_lagging_jackets', ':', req_lagging_jackets)
        # print('req_lagging_jacket_count', ':', req_lagging_jacket_count)
        
        for pm in ofl_pm:
            # print(json_val_dict[pm])
            if str(json_val_dict[pm]) not in ['', '0', 'N/A']: # if any primary measure has any valid value
                json_val_dict["LED Bulbs: supply only (4 no.)"] = 1
                json_val_dict["Hot Water Cylinder Jacket"] = req_lagging_jacket_count
        
        # print("Hot Water Cylinder Jacket", ':', json_val_dict["Hot Water Cylinder Jacket"])
        # print('req_lagging_jacket_count', ':', req_lagging_jacket_count)
        
        
        # print(external_wall_insulation)

        # print(json_val_dict["Internal Wall Insulation: Vertical Surface"])
        
        # print('sum of Ex/In: ', float(external_wall_insulation) + float(json_val_dict["Internal Wall Insulation: Vertical Surface"]))
        json_val_dict["Air-tightness test recommended?"] = 1 if float(external_wall_insulation) + float(json_val_dict["Internal Wall Insulation: Vertical Surface"]) > 0 else ''
        

        json_val_dict["Cavity Wall Insulation Bonded Bead"] = round(json_val_dict["Cavity Wall Insulation Bonded Bead"]) if json_val_dict["Cavity Wall Insulation Bonded Bead"] != 0 else 'N/A'
        json_val_dict["Loose Fibre Extraction"] = round(json_val_dict["Loose Fibre Extraction"]) if json_val_dict["Loose Fibre Extraction"] != 0 else 'N/A'
        json_val_dict["Internal Wall Insulation: Vertical Surface"] = round(json_val_dict["Internal Wall Insulation: Vertical Surface"]) if json_val_dict["Internal Wall Insulation: Vertical Surface"] != 0 else 'N/A'
        json_val_dict['replace_window_area'] = round(json_val_dict['replace_window_area']) if json_val_dict['replace_window_area'] != 0 else 'N/A'
        json_val_dict['replace_window_area'] = 1 if json_val_dict['replace_window_area'] == 0 else json_val_dict['replace_window_area']
        if 'Notes (Windows and Doors)' in json_val_dict.keys():
            json_val_dict['Notes (Windows and Doors)'] = json_val_dict['Notes (Windows and Doors)'] if json_val_dict['Notes (Windows and Doors)'] != '' else 'N/A'
        # json_val_dict['No. Double Glazed Windows *'] = json_val_dict['No. Double Glazed Windows *'] - json_val_dict['No. Single Glazed Windows *']
        
        
        
        


            
        # json_val_dict["Is a Major Renovation calculation necessary?*"] = True
        # json_val_dict['Thermal Envelope - Heat loss walls, windows and doors'] = 0
        # json_val_dict['Thermal Envelope - Heat loss floor area'] = 0
        # json_val_dict["Reason Major Renovation calculation is not necessary?*"] = "The proportions of EWI/IWI or significantly greater than 25%"
        # json_val_dict['Qualifying Boiler'] = 'N/A'
        
        warnings = 'Major Renovation Error:'
        print("Is a Major Renovation calculation necessary?*", ':', json_val_dict["Is a Major Renovation calculation necessary?*"])
        if json_val_dict["Is a Major Renovation calculation necessary?*"] in [True, "Yes"]:
            print(json_val_dict['Thermal Envelope - Heat loss walls, windows and doors'])
            print(json_val_dict['Thermal Envelope - Heat loss floor area'])
            print(json_val_dict['Heat loss Wall Area recommended for EWI and IWI'])
            if 0 in [json_val_dict['Thermal Envelope - Heat loss walls, windows and doors'], json_val_dict['Thermal Envelope - Heat loss floor area'], json_val_dict['Heat loss Wall Area recommended for EWI and IWI']]:
                warnings = warnings + "<BR>" + 'Major Renovation calculation has been confirmed as necessary but not provided. The Values for "EWI/IWI >25%", "Qualifying Boiler", Heating Recommendations and the Lot Type may be incorrect as the required information has not been provided. Please resubmit to include Major Renovation calculations on Heat Loss 10th floor, 11th floor etc. and include values for Walls receiving EWI or IWI and External Walls not receiving EWI or IWI and confirm the response to the form question "Is there Mains Gas in the Area?"'
                json_val_dict["EWI/IWI > 25% *"] = "No"

            # for q in ['Thermal Envelope - Heat loss walls, windows and doors', 'Thermal Envelope - Heat loss floor area', 'Heat loss Wall Area recommended for EWI and IWI']:
                # if json_val_dict[q] == 0:
                    # warnings = warnings + "<BR>" + "Setting 'EWI/IWI > 25% *' to 'No', since Major Renovation calculation has been confirmed as necessary but zero value provided for '" + q + "'"
                    # json_val_dict["EWI/IWI > 25% *"] = "No"
            
            
            
        print(warnings)
            
            
            
        if json_val_dict["Is a Major Renovation calculation necessary?*"] == False:
            # print(json_val_dict["Reason Major Renovation calculation is not necessary?*"])
            if json_val_dict["Reason Major Renovation calculation is not necessary?*"] == "The proportions of EWI/IWI are significantly greater than 25%":
                json_val_dict["EWI/IWI > 25% *"] = "Yes"
                # if json_val_dict['Qualifying Boiler'] == 'N/A':
                    # warnings = warnings + "<BR>" + "Qualifying Boiler question must be answered Yes/No"
                    
        
        if warnings == 'Major Renovation Error:':
            warnings = ''
        
        
        json_val_dict['Qualifying Boiler'] = False
        
        if json_val_dict['EWI/IWI > 25% *'] == 'No':
            json_val_dict['Qualifying Boiler'] = 'N/A'
        else:
            print('condensing', ':', str(condensing))
            if condensing == False:
                json_val_dict['Qualifying Boiler'] = True
                if linked_stove_bb == True:
                    json_val_dict['Qualifying Boiler'] = False
        
        
        
        print(4)
        
        json_val_dict['Suitable for Heating Measures *'] = False
        
        if json_val_dict['Qualifying Boiler'] == True:
            json_val_dict['Suitable for Heating Measures *'] = True
        
        # print(json_val_dict['Suitable for Heating Measures *'])
        # print("Is there Mains Gas in the Area?", ':', json_val_dict["Is there Mains Gas in the Area?"])
        
        
        if json_val_dict['Suitable for Heating Measures *'] == False:
            json_val_dict['Not suitable details*'] = json_val_dict['Notes (Heating)']
        
        print("Is there Mains Gas in the Area?", ':', json_val_dict["Is there Mains Gas in the Area?"])
        
        if json_val_dict["Is there Mains Gas in the Area?"] != "N/A - No heating measured being recommended":
            if json_val_dict['Qualifying Boiler'] == True:
                if json_val_dict['Heating Systems Controls *'] == 'Full zone control to spec':
                    if json_val_dict["Is there Mains Gas in the Area?"] == "Yes":
                        json_val_dict['Basic gas heating system'] = 1
                    else:
                        json_val_dict['Basic oil heating system'] = 1
            
            # print("Electric Storage Heater age (years)*", ':', json_val_dict["Electric Storage Heater age (years)*"])
            # print("Warm Air System age (years)*", ':', json_val_dict["Warm Air System age (years)*"])
            # print('Heating System *', ':', json_val_dict['Heating System *'])
            
            if (json_val_dict['Heating System *'] in ["Electric Storage Heater", "Warm Air System"] and json_val_dict['System Age *'] == "25+") or json_val_dict['Heating System *'] in ["Open Fire with Back Boiler", "Open Fire with Back Boiler With Enclosure Door", "Solid Fuel Range", "Solid Fuel Range with Back Boiler", "Fixed Electric Inset Fire", "Gas Fire Room Heater", "Oil Stove Room Heater", "Open Fire", "Open Fire With Enclosure Door"]:
                if json_val_dict["Is there Mains Gas in the Area?"] == "Yes":
                    json_val_dict['Full gas heating system installation'] = 1
                else:
                    json_val_dict['Full oil heating system installation'] = 1
            
            if json_val_dict['Qualifying Boiler'] == True:
                print('Heating Systems Controls *', ':', json_val_dict['Heating Systems Controls *'])
                if json_val_dict['Heating Systems Controls *'] != 'Full zone control to spec':
                    if json_val_dict["Is there Mains Gas in the Area?"] == "Yes":
                        json_val_dict['Gas boiler and controls (Basic & controls pack)'] = 1
                    else:
                        json_val_dict['Oil boiler and controls (Basic & controls pack)'] = 1
            
        # print(2)
        # print("Hot Water Cylinder Jacket", ':', json_val_dict["Hot Water Cylinder Jacket"])
        for field in ofl_hpm:
            if json_val_dict[field] != '' and field != 'Hot Water Cylinder Jacket':
                print(field, ':', json_val_dict[field])
                json_val_dict['Hot Water Cylinder Jacket'] = ''
        # print("Hot Water Cylinder Jacket", ':', json_val_dict["Hot Water Cylinder Jacket"])
        # print(3)
        json_val_dict['Permanent ventilation wall vent (Certified Proprietary Integrated System)'] = json_val_dict["New Permanent Vent"]
        json_val_dict['Background ventilation wall vent (Certified Proprietary Integrated System)'] = json_val_dict["New Background Vent"]
        json_val_dict['Ducting existing cooker hood to exterior'] = json_val_dict['Duct Cooker Hood']
        json_val_dict['Window (same m2 rate will apply to windows with certified trickle vents)'] = int(round(json_val_dict['New Windows being recommended for replacement'], 0))
        json_val_dict['GNI new connection'] = json_val_dict['New Gas Connection']
        json_val_dict['Additional Roof Ventilation (Low Level)'] = json_val_dict['Required per standards (mm2) *'] - json_val_dict["Existing (mm2)*"]
        json_val_dict['Additional Roof Ventilation (High Level)'] = json_val_dict['high_roof_vent_area']
        
        # print('Window (same m2 rate will apply to windows with certified trickle vents)', ':', json_val_dict['Window (same m2 rate will apply to windows with certified trickle vents)'])

        
        # xl_2_pdf(xl_path)
        # print(json_val_dict)
        output_dict = json_val_dict
        
        for field in output_dict:
            value = output_dict[field]
            if (type(value) == bool and value == True):
                output_dict[field] = "Yes"
            if (type(value) == bool and value == False):
                output_dict[field] = "No"
        
        # xl_path = 
        # populate_template(output_dict)
        
        
        # Below all N/A unless type = other        
        
        missing = ['Thermal Envelope - Heat loss walls, windows and doors'
                , 'Other Details Roof 1 *'
                , 'Other Details Roof 2 *'
                , 'Other Details Roof 3 *'
                , 'Other Details Roof 4 *'
                , 'Other W1 Details *'
                , 'Other W2 Details *'
                , 'Other W3 Details *'
                , 'Other W4 Details *'

                , 'Internal Wall Insulation: Sloped or flat (horizontal) surface'
                , 'Attic (Loft) Insulation 100 mm top-up'
                , 'Attic (Loft) Insulation 150 mm top-up'
                , 'Attic (Loft) Insulation 200 mm top-up'
                , 'Attic (Loft) Insulation 250 mm top up'
                , 'Attic (Loft) Insulation 300 mm'
                
                , 'Attic Storage (5m2)' # "1" if any of the 5 above (not sloped) is non-zero
                
                , 'Additional Roof Ventilation (High Level)' # high_roof_vent_area
                , 'Additional Roof Ventilation (Low Level)' # 'Required per standards (mm2) *' - 'Existing (mm2)*'
                
                # , 'Roof 2 Insulation Type'
                , 'Wall 1 Residual Cavity Width (mm)*'
                , 'Wall 2 Residual Cavity Width (mm)*'
                , 'Wall 3 Residual Cavity Width (mm)*'
                , 'Wall 4 Residual Cavity Width (mm)*'

                # , 'Insulation *'
                # , 'Condition of Lagging Jacket *'
                , 'Partial Details *'
                
                ]
        
        print('missing:')
        for m in missing:
            if m not in json_val_dict.keys():
                print(m, 'not in json_val_dict.keys')
            else:
                print(m, ':', json_val_dict[m])
        
        
        
        
        # print(output_dict)
        output_dict['Lot *'] = 'Lot ' + lot(output_dict)
        print(output_dict['plan_name'], 'Lot *', ':', output_dict['Lot *'])
        
        # print('Not Working Details Primary Heating *', ':', json_val_dict['Not Working Details Primary Heating *'])
        
        populate_template_new(output_dict, 'template')
        
                
        # print(ofl_wos)
        
        # ofl_wos_2 = ofl_wos
        
        for f in ofl_wos_2:
            # print('f', ':', f)
            if f.isupper():
                continue
            else:
                if f in json_val_dict.keys():
                    # print(f, 'found in json_val_dict:', "'" + str(json_val_dict[f]) + "'")
                    if json_val_dict[f] in ['', 'N/A', None, 0]:
                        ofl_wos.remove(f)
                        # print('removed ' + f)
                else:
                    # print('removing ', f)
                    ofl_wos.remove(f)
        output = ''
        # json_val_dict["Is a Major Renovation calculation necessary?*"] = True
        print('json_val_dict["Is a Major Renovation calculation necessary?*"]', ':', json_val_dict["Is a Major Renovation calculation necessary?*"])
        if json_val_dict["Is a Major Renovation calculation necessary?*"] == "Yes":
            print('generating template_mrc')
            output, filename = populate_template_new(output_dict, 'template_mrc')
            
            ofl_filelist.append(filename)
            # ofl_filelist.append(filename)
            # print(ofl_filelist)
        

        
        
        # print(output)
        # print('Attic (Loft) Insulation 200 mm top-up', ':', json_val_dict['Attic (Loft) Insulation 200 mm top-up'])
        
        # print("Hot Water Cylinder Jacket", ':', json_val_dict["Hot Water Cylinder Jacket"])
        if output == '':
            styling = "border=\"1\""
            output = f"""\
                <h1>Work Order Summary</h1> \
                {create_table_text(output_dict, headers = ['Measure Item Name', 'Total Quantity'], styling=styling, do_not_sum=['All'], order_list = ofl_wos)} \
                <h1>General</h1> \
                {create_table_text(output_dict, headers = ['name', 'value'], styling=styling, do_not_sum=['All'], order_list = ofl_general)} \
                <h1>Major Renovation</h1> \
                {create_table_text(output_dict, headers = ['Building Thermal Envelope', 'm2'], styling=styling, do_not_sum=['All'], order_list = ofl_mr)} \
                <h1>Primary Measures</h1> \
                {create_table_text(output_dict, headers = ['name', 'value'], styling=styling, do_not_sum=['All'], order_list = ofl_pm)} \
                <h1>Roof</h1> \
                {create_table_text(output_dict, headers = ['name', 'value'], styling=styling, do_not_sum=['All'], order_list = ofl_roof)} \
                <h1>Walls</h1> \
                {create_table_text(output_dict, headers = ['name', 'value'], styling=styling, do_not_sum=['All'], order_list = ofl_walls)} \
                <h1>Heating</h1> \
                {create_table_text(output_dict, headers = ['name', 'value'], styling=styling, do_not_sum=['All'], order_list = ofl_heating)} \
                <h1>Mechanical Ventilation Systems, Air Tightness Testing & Energy</h1> \
                {create_table_text(output_dict, headers = ['name', 'value'], styling=styling, do_not_sum=['All'], order_list = ofl_mae)} \
                <h1>Supplementary</h1> \
                {create_table_text(output_dict, headers = ['name', 'value'], styling=styling, do_not_sum=['All'], order_list = ofl_s)} \
                <h1>Heating Primary Measures</h1> \
                {create_table_text(output_dict, headers = ['name', 'value'], styling=styling, do_not_sum=['All'], order_list = ofl_hpm)} \
                <h1>File List</h1> \
                {create_table_text(output_dict, headers = ['name', 'value'], styling=styling, do_not_sum=['All'], order_list = ofl_filelist)} \

                </div>"""


            if warnings != '':
                output = f"""\
                    <h1>Warnings</h1> \
                    {warnings} \
                    </div>""" + output

        # print(output)

    except Exception as ex:
        # exc_type, exc_obj, exc_tb = sys.exc_info()
        # output = "Line " + str(exc_tb.tb_lineno) + ": " + exc_type 
        
        # populate_template(xml_val_dict) # adds an (almost) empty copy of the template to avoid potential Logic App error if file not found
        
        populate_template_new(xml_val_dict, 'template')
        # populate_template_new(output_dict, 'template_mrc')
        
        # output = str(ex)
        output = traceback.format_exc()
        # LOGGER.info('Exception : ' + str(traceback.format_exc()))
        
        # fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        # print(exc_type, fname, exc_tb.tb_lineno)
    finally:
        return output
    return output


def get_forms_data(id, headers = {
            "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/39.0.2171.95 Safari/537.36"
            ,"key": "45170e50321733db78952dfa5901b0dfeeb8"
            , "customer": "63b5a4ae69c91"
            , "accept": "application/json"
            }):
    
    try:
        json_val_dict = {}
        form_dict = {}
        json_url = "https://cloud.magicplan.app/api/v2/plans/forms/" + id
        request = urllib.request.Request(json_url, headers=headers)
        
        JSON = urllib.request.urlopen(request).read()
        JSON = json.loads(JSON)
        for datum in JSON["data"]:
            if datum["symbol_type"] not in form_dict.keys():
                form_dict[datum["symbol_type"]] = {}
            if datum["symbol_name"] not in form_dict[datum["symbol_type"]].keys():
                form_dict[datum["symbol_type"]][datum["symbol_name"]] = {}
            form_dict[datum["symbol_type"]][datum["symbol_name"]][datum["symbol_instance_id"]] = {}
            for form in datum["forms"]:
                for section in form["sections"]:
                    for field in section["fields"]:
                        v = ''
                        if field["value"]["value"] == None:
                            vals = []
                            if field["type_as_string"] == "list":
                                # print("type_as_string", ':', field["type_as_string"])
                                vals = [val["value"] for val in field["value"]["values"]]
                                for val in vals:
                                    v += val
                                    v += '<BR>'
                        else:
                            v = field["value"]["value"]
                        # print(field["label"], ':', v)
                        # form_dict[datum["symbol_type"]][datum["symbol_name"]][datum["symbol_instance_id"]][field["label"]] = v
                        # json_val_dict[field["label"]] = v
                        
                        im = field["label"].replace(' *', '')
                        im = im.replace('*', '')
                        form_dict[datum["symbol_type"]][datum["symbol_name"]][datum["symbol_instance_id"]][im] = v
                        json_val_dict[im] = v
        
        output = json_val_dict, form_dict
    
    except:
        output = traceback.format_exc()
        print(output)
    
    return output


def distributor_function(form, root = ''):
    
    # Processing steps common to all projects are performed first
        # Establish fundamental parameters
            # plan_name (name given to the project by the user)
            # email (of the user who exported the project)
            # xml (a url where the project XML data can be retrieved)
        # Retrieve XML data
            # Load as dict
            # Get the MagicPlan Project ID (ToDo: check is this available directly from the form?)
        
        # Retrieve API data (forms, statistics, files) # Need forms here to get "This project is a" but others could probably wait... don't forget "get_project_files" returns a list that is currently included in the email body HTML tables so probably leave that where it is in the project-specific function
            # Load as dict x 3
            # ToDo: avoid any potential Key Errors by giving these all a default value of ''
        # Identify what type of project it is
    
    # Project-specific processing is then carried out by dedicated functions:
        # survey()
        # BER()
    # These return a HTML formatted table as output which will appear as the email body
    
    # Return "json_data" to be uploaded to Azure blob storage where it will be processed by the Logic App
    
    
    try:

        if root == '':
            email = form['email']
            xml = form['xml']
            root : ET.Element # why is this line necessary? Does this initialise the variable? Does it enforce the type?
            with urllib.request.urlopen(xml) as f:
                s = f.read().decode('utf-8')
            root = dET.fromstring(s)
        else:
            email = 'gtsupport@ie.gt.com'
            xml = ''

        project_name = root.get('name')
        project_id = root.get('id')
        
        forms_val_dict, form_dict = get_forms_data(project_id)
        
        
        output = ''
        
        if "This project is a" in forms_val_dict.keys():
            print("This project is a", ':', forms_val_dict["This project is a"])
            if forms_val_dict["This project is a"] == "BER":
                output = BER(root, email = 'gtsupport@ie.gt.com')
            if forms_val_dict["This project is a"] == "Survey":
                output = survey(root)
        
        if output == '':
            output = survey(root)
        
        output = output + '<h2>' + xml + '</h2></div>'
        
        
        if project_name[-1] == ' ':
            project_name = project_name[:-1]

        json_data = json.dumps({
            'email' : email,
            'name'  : project_name, 
            'table' : output
        })
        output = json_data
    except:
        output = traceback.format_exc()
        print(output)

    return output




app = func.FunctionApp()
@app.function_name(name="MagicplanTrigger")
@app.route(route="magicplan", auth_level=func.AuthLevel.ANONYMOUS)


def test_function(req: func.HttpRequest) -> func.HttpResponse:
    try:

        form = dict(req.form)
        json_data = distributor_function(form)




        sc = 200    # OK

    except Exception as ex:
        output = str(ex)
        output = traceback.format_exc()
        sc = 503    # Service Unavailable
        

    finally:        
        try:
            # azure_upload(json_data)
            account_url = os.environ['AZ_STR_URL']
            # account_url = "https://ksnmagicplanfunc3e54b9.blob.core.windows.net"
            default_credential = DefaultAzureCredential()
            blob_service_client = BlobServiceClient(account_url, credential=default_credential)
            
            # container_name = os.environ['AZ_CNTR_ST']
            # container_name = "magicplan-container"
            container_name = "function-app-output"
            container_client = blob_service_client.get_container_client(container_name)
            if not container_client.exists():
                container_client = blob_service_client.create_container(container_name)
            
            local_file_name = str(uuid.uuid4()) + '.json'
            blob_client = blob_service_client.get_blob_client(container=container_name, blob=local_file_name)
            blob_client.upload_blob(json_data)
            
            container_name = "project-files"
            container_client = blob_service_client.get_container_client(container_name)
            if not container_client.exists():
                container_client = blob_service_client.create_container(container_name)


            # local_file_name = str(uuid.uuid4()) + '_post' + ".txt"
            # data = "Hello, World!"
            # blob_client = blob_service_client.get_blob_client(container=container_name, blob=local_file_name)
            # blob_client.upload_blob(data)



            # json_url = "https://cloud.magicplan.app/api/v2/plans/" + str(xml_val_dict['Application ID']) + "/files?include_photos=true"
            # request = urllib.request.Request(json_url, headers=headers)
            # JSON = urllib.request.urlopen(request).read()
            # JSON = json.loads(JSON)

            # for file in JSON["data"]["files"]:
                # print(file["file_type"])
                # if file["file_type"] == "pdf":
                    # request = urllib.request.Request(file["url"], headers=headers)
                    # file_content = urllib.request.urlopen(request).read()
                    # local_file_name = file["name"]
                    # blob_client = blob_service_client.get_blob_client(container=container_name, blob=local_file_name)
                    # blob_client.upload_blob(file_content)
            
            # for file in JSON["data"]["photos"]:
                # request = urllib.request.Request(file["url"], headers=headers)
                # file_content = urllib.request.urlopen(request).read()
                # local_file_name = file["name"]
                # blob_client = blob_service_client.get_blob_client(container=container_name, blob=local_file_name)
                # blob_client.upload_blob(file_content)
            
            
            
            
            return_body = '0'
            
        except Exception as ex:
            output = str(ex)
            output = traceback.format_exc()
            sc = 500     # Internal Server Error
            # return_body = output
            json_data = json.dumps({
                'email' : email,
                'name'  : plan_name, 
                'table' : output
            })
            azure_upload(json_data)
        return func.HttpResponse(status_code=sc, body=return_body)



def azure_upload(file_data, container_name = 'attachment'):
    # account_url = os.environ['AZ_STR_URL']
    account_url = "https://ksnmagicplanfunc3e54b9.blob.core.windows.net"
    
    default_credential = DefaultAzureCredential()
    blob_service_client = BlobServiceClient(account_url, credential=default_credential)
    
    # container_name = os.environ['AZ_CNTR_ST']
    container_client = blob_service_client.get_container_client(container_name)
    if not container_client.exists():
        container_client = blob_service_client.create_container(container_name)
    
    local_file_name = str(uuid.uuid4()) + '.json'
    blob_client = blob_service_client.get_blob_client(container=container_name, blob=local_file_name)
    blob_client.upload_blob(file_data)
    
    return



def populate_template_new(json_val_dict, template):
    try:

        account_url = "https://ksnmagicplanfunc3e54b9.blob.core.windows.net"
        default_credential = DefaultAzureCredential()
        output = ''
        return_filename = ''
        
        if template == 'template':
            filename = json_val_dict['plan_name'] + '.xlsx'
            container_name = 'attachment'
            local_path = "/tmp"
            print('local_path', ':', local_path)
            instance_file_path = os.path.join(local_path, filename)
            print('instance_file_path', ':', instance_file_path)
            v = {
                'Applicant Name': { 'Value': '' , 'Tab': 'General' , 'Cell': 'C4'}
                , 'Application ID': { 'Value': '' , 'Tab': 'General' , 'Cell': 'E4'}
                , 'Client Address': { 'Value': '' , 'Tab': 'General' , 'Cell': 'C6'}
                , 'MPRN': { 'Value': '' , 'Tab': 'General' , 'Cell': 'E6'}
                , 'Surveyor': { 'Value': '' , 'Tab': 'General' , 'Cell': 'C8'}
                , 'Survey Date *': { 'Value': '' , 'Tab': 'General' , 'Cell': 'E8'}
                , 'Dwelling Type*': { 'Value': '' , 'Tab': 'General' , 'Cell': 'C10'}
                , 'Gross floor area (m2) *': { 'Value': '' , 'Tab': 'General' , 'Cell': 'E10'}
                , 'Dwelling Age*': { 'Value': '' , 'Tab': 'General' , 'Cell': 'C12'}
                , 'Number of Storeys *': { 'Value': '' , 'Tab': 'General' , 'Cell': 'E12'}
                , 'Age Extension 1': { 'Value': '' , 'Tab': 'General' , 'Cell': 'C14'}
                , 'Room in Roof': { 'Value': '' , 'Tab': 'General' , 'Cell': 'E14'}
                , 'Age Extension 2': { 'Value': '' , 'Tab': 'General' , 'Cell': 'C16'}
                , 'No. Single Glazed Windows *': { 'Value': '' , 'Tab': 'General' , 'Cell': 'E16'}
                , 'Asbestos Suspected': { 'Value': '' , 'Tab': 'General' , 'Cell': 'C18'}
                , 'No. Double Glazed Windows *': { 'Value': '' , 'Tab': 'General' , 'Cell': 'E18'}
                , 'Asbestos Details': { 'Value': '' , 'Tab': 'General' , 'Cell': 'C20'}
                , 'Lot *': { 'Value': '' , 'Tab': 'General' , 'Cell': 'C22'}
                , 'Property Height (m)*': { 'Value': '' , 'Tab': 'General' , 'Cell': 'E22'}
                , 'Eircode': { 'Value': '' , 'Tab': 'General' , 'Cell': 'C24'}
                , 'Internet Available': { 'Value': '' , 'Tab': 'General' , 'Cell': 'E24'}


                , 'Roof 1 Type*': { 'Value': '' , 'Tab': 'Roof' , 'Cell': 'E25'}
                , 'Sloped Ceiling Roof 1*': { 'Value': '' , 'Tab': 'Roof' , 'Cell': 'H25'}
                , 'Other Details Roof 1*': { 'Value': '' , 'Tab': 'Roof' , 'Cell': 'E27'}
                , 'Roof 1 greater than 2/3 floor area*': { 'Value': '' , 'Tab': 'Roof' , 'Cell': 'H27'}
                , 'Roof 1 Pitch (degrees)*': { 'Value': '' , 'Tab': 'Roof' , 'Cell': 'H29'}
                , 'Roof Type 1 Insulation Exists*': { 'Value': '' , 'Tab': 'Roof' , 'Cell': 'E33'}
                , 'Can Roof Type 1 Insulation Thickness be Measured?*': { 'Value': '' , 'Tab': 'Roof' , 'Cell': 'E35'}
                , 'Roof 1 Thickness (mm)*': { 'Value': '' , 'Tab': 'Roof' , 'Cell': 'E37'}
                , 'Roof 1 Insulation Type*': { 'Value': '' , 'Tab': 'Roof' , 'Cell': 'H33'}
                #N/A
                , 'Required per standards (mm2) *': { 'Value': '' , 'Tab': 'Roof' , 'Cell': 'E42'}
                , 'Existing (mm2)*': { 'Value': '' , 'Tab': 'Roof' , 'Cell': 'H42'}
                , 'Area of Roof Type 1 with fixed flooring (m2)*': { 'Value': '' , 'Tab': 'Roof' , 'Cell': 'E46'}
                , 'Folding/stair ladder in Roof Type 1*': { 'Value': '' , 'Tab': 'Roof' , 'Cell': 'E48'}
                , 'Fixed light in Roof Type 1*': { 'Value': '' , 'Tab': 'Roof' , 'Cell': 'E50'}
                , 'Downlighters in Roof Type 1*': { 'Value': '' , 'Tab': 'Roof' , 'Cell': 'H46'}
                , 'High power cable in Roof Type 1 (6sq/10sq or higher)*': { 'Value': '' , 'Tab': 'Roof' , 'Cell': 'H48'}
                , 'Roof 2 Type': { 'Value': '' , 'Tab': 'Roof Type 2' , 'Cell': 'E25'}
                , 'Other Details Roof 2*': { 'Value': '' , 'Tab': 'Roof Type 2' , 'Cell': 'E27'}
                , 'Sloped Ceiling Roof 2*': { 'Value': '' , 'Tab': 'Roof Type 2' , 'Cell': 'H25'}
                , 'Roof 2 greater than 2/3 floor area*': { 'Value': '' , 'Tab': 'Roof Type 2' , 'Cell': 'H27'}
                , 'Roof 2 Pitch (degrees)*': { 'Value': '' , 'Tab': 'Roof Type 2' , 'Cell': 'H29'}
                , 'Roof 2 Insulation Exists*': { 'Value': '' , 'Tab': 'Roof Type 2' , 'Cell': 'E33'}
                , 'Can Roof Type 2 Insulation Thickness be Measured?': { 'Value': '' , 'Tab': 'Roof Type 2' , 'Cell': 'E35'}
                , 'Roof 2 Thickness (mm)': { 'Value': '' , 'Tab': 'Roof Type 2' , 'Cell': 'E37'}
                , 'Roof 2 Insulation Type': { 'Value': '' , 'Tab': 'Roof Type 2' , 'Cell': 'H33'}
                #N/A
                , 'Roof 2 Required per standards (mm2) *': { 'Value': '' , 'Tab': 'Roof Type 2' , 'Cell': 'E42'}
                , 'Roof 2 Existing (mm2) *': { 'Value': '' , 'Tab': 'Roof Type 2' , 'Cell': 'H42'}
                , 'Area of Roof Type 2 with fixed flooring (m2)': { 'Value': '' , 'Tab': 'Roof Type 2' , 'Cell': 'E46'}
                , 'Folding/stair ladder in Roof Type 2': { 'Value': '' , 'Tab': 'Roof Type 2' , 'Cell': 'E48'}
                , 'Fixed light in Roof Type 2': { 'Value': '' , 'Tab': 'Roof Type 2' , 'Cell': 'E50'}
                , 'Downlighters in Roof Type 2': { 'Value': '' , 'Tab': 'Roof Type 2' , 'Cell': 'H46'}
                , 'High power cable in Roof Type 2 (6sq/10sq or higher)': { 'Value': '' , 'Tab': 'Roof Type 2' , 'Cell': 'H48'}
                , 'Roof 3 Type': { 'Value': '' , 'Tab': 'Roof Type 3' , 'Cell': 'E25'}
                , 'Other Details Roof 3*': { 'Value': '' , 'Tab': 'Roof Type 3' , 'Cell': 'E27'}
                , 'Sloped Ceiling Roof 3*': { 'Value': '' , 'Tab': 'Roof Type 3' , 'Cell': 'H25'}
                , 'Roof 3 greater than 2/3 floor area*': { 'Value': '' , 'Tab': 'Roof Type 3' , 'Cell': 'H27'}
                , 'Roof 3 Pitch (degrees)*': { 'Value': '' , 'Tab': 'Roof Type 3' , 'Cell': 'H29'}
                , 'Roof 3 Insulation Exists*': { 'Value': '' , 'Tab': 'Roof Type 3' , 'Cell': 'E33'}
                , 'Can Roof Type 3 Insulation Thickness be Measured?*': { 'Value': '' , 'Tab': 'Roof Type 3' , 'Cell': 'E35'}
                , 'Roof 3 Thickness (mm)*': { 'Value': '' , 'Tab': 'Roof Type 3' , 'Cell': 'E37'}
                , 'Roof 3 Insulation Type*': { 'Value': '' , 'Tab': 'Roof Type 3' , 'Cell': 'H33'}
                #N/A
                , 'Roof 3 Required per standards (mm2) *': { 'Value': '' , 'Tab': 'Roof Type 3' , 'Cell': 'E42'}
                , 'Roof 3 Existing (mm2) *': { 'Value': '' , 'Tab': 'Roof Type 3' , 'Cell': 'H42'}
                , 'Area of Roof Type 3 with fixed flooring (m2)*': { 'Value': '' , 'Tab': 'Roof Type 3' , 'Cell': 'E46'}
                , 'Folding/stair ladder in Roof Type 3*': { 'Value': '' , 'Tab': 'Roof Type 3' , 'Cell': 'E48'}
                , 'Fixed light in Roof Type 3*': { 'Value': '' , 'Tab': 'Roof Type 3' , 'Cell': 'E50'}
                , 'Downlighters in Roof Type 3*': { 'Value': '' , 'Tab': 'Roof Type 3' , 'Cell': 'H46'}
                , 'High power cable in Roof Type 3 (6sq/10sq or higher)*': { 'Value': '' , 'Tab': 'Roof Type 3' , 'Cell': 'H48'}
                , 'Roof 4 Type': { 'Value': '' , 'Tab': 'Roof Type 4' , 'Cell': 'E25'}
                , 'Other Details Roof 4*': { 'Value': '' , 'Tab': 'Roof Type 4' , 'Cell': 'E27'}
                , 'Sloped Ceiling Roof 4*': { 'Value': '' , 'Tab': 'Roof Type 4' , 'Cell': 'H25'}
                , 'Roof 4 greater than 2/3 floor area': { 'Value': '' , 'Tab': 'Roof Type 4' , 'Cell': 'H27'}
                , 'Roof 4 Pitch (degrees)*': { 'Value': '' , 'Tab': 'Roof Type 4' , 'Cell': 'H29'}
                , 'Roof Type 4 Insulation Exists*': { 'Value': '' , 'Tab': 'Roof Type 4' , 'Cell': 'E33'}
                , 'Can Roof Type 4 Insulation Thickness be Measured?*': { 'Value': '' , 'Tab': 'Roof Type 4' , 'Cell': 'E35'}
                , 'Roof 4 Thickness (mm)': { 'Value': '' , 'Tab': 'Roof Type 4' , 'Cell': 'E37'}
                , 'Roof 4 Insulation Type*': { 'Value': '' , 'Tab': 'Roof Type 4' , 'Cell': 'H33'}
                #N/A
                , 'Roof 4 Required per standards (mm2) *': { 'Value': '' , 'Tab': 'Roof Type 4' , 'Cell': 'E42'}
                , 'Roof 4 Existing (mm2) *': { 'Value': '' , 'Tab': 'Roof Type 4' , 'Cell': 'H42'}
                , 'Area of Roof Type 4 with fixed flooring (m2)*': { 'Value': '' , 'Tab': 'Roof Type 4' , 'Cell': 'E46'}
                , 'Folding/stair ladder in Roof Type 4*': { 'Value': '' , 'Tab': 'Roof Type 4' , 'Cell': 'E48'}
                , 'Fixed light in Roof Type 4*': { 'Value': '' , 'Tab': 'Roof Type 4' , 'Cell': 'E50'}
                , 'Downlighters in Roof Type 4*': { 'Value': '' , 'Tab': 'Roof Type 4' , 'Cell': 'H46'}
                , 'High power cable in Roof Type 4 (6sq/10sq or higher)*': { 'Value': '' , 'Tab': 'Roof Type 4' , 'Cell': 'H48'}
                , 'Suitable for Insulation *': { 'Value': '' , 'Tab': 'Roof' , 'Cell': 'E54'}
                , 'Roof not suitable details*': { 'Value': '' , 'Tab': 'Roof' , 'Cell': 'E56'}
                , 'Notes (Roof)': { 'Value': '' , 'Tab': 'Roof' , 'Cell': 'E59'}
                , 'Internal Wall Insulation: Sloped or flat (horizontal) surface': { 'Value': '' , 'Tab': 'Roof' , 'Cell': 'I67', 'Default': ''}
                , 'Attic (Loft) Insulation 100 mm top-up': { 'Value': '' , 'Tab': 'Roof' , 'Cell': 'I68', 'Default': ''}
                , 'Attic (Loft) Insulation 150 mm top-up': { 'Value': '' , 'Tab': 'Roof' , 'Cell': 'I69', 'Default': ''}
                , 'Attic (Loft) Insulation 200 mm top-up': { 'Value': '' , 'Tab': 'Roof' , 'Cell': 'I70', 'Default': ''}
                , 'Attic (Loft) Insulation 250 mm top up': { 'Value': '' , 'Tab': 'Roof' , 'Cell': 'I71', 'Default': ''}
                , 'Attic (Loft) Insulation 300 mm': { 'Value': '' , 'Tab': 'Roof' , 'Cell': 'I72', 'Default': ''}
                
                , 'Attic Storage (5m2)': { 'Value': '' , 'Tab': 'Roof' , 'Cell': 'I76', 'Default': ''}
                , 'Replacement of CWST/F&E tank': { 'Value': '' , 'Tab': 'Roof' , 'Cell': 'I78', 'Default': ''}
                , 'Additional Roof Ventilation (High Level)': { 'Value': '' , 'Tab': 'Roof' , 'Cell': 'I84', 'Default': ''}
                , 'Additional Roof Ventilation (Low Level)': { 'Value': '' , 'Tab': 'Roof' , 'Cell': 'I85', 'Default': ''}

                
                
                
                

                , 'Wall Type 1*': { 'Value': '' , 'Tab': 'Wall' , 'Cell': 'E27'}
                , 'Wall 1 wall thickness (mm)*': { 'Value': '' , 'Tab': 'Wall' , 'Cell': 'E29'}
                , 'Wall 1 Insulation Present?*': { 'Value': '' , 'Tab': 'Wall' , 'Cell': 'E33'}
                , 'Wall 1 Insulation Type*': { 'Value': '' , 'Tab': 'Wall' , 'Cell': 'E35'}
                , 'Wall 1 Fill Type*': { 'Value': '' , 'Tab': 'Wall' , 'Cell': 'E37'}
                , 'Wall 1 Residual Cavity Width (mm)*': { 'Value': '' , 'Tab': 'Wall' , 'Cell': 'K37'}
                , 'Can Wall type 1 Insulation Thickness be Measured?*': { 'Value': '' , 'Tab': 'Wall' , 'Cell': 'K33'}
                , "If 'Yes' enter Wall type 1 insulation thickness (mm)*": { 'Value': '' , 'Tab': 'Wall' , 'Cell': 'K35'}
                , 'Wall Type 2': { 'Value': '' , 'Tab': 'Wall Type 2' , 'Cell': 'E27'}
                , 'Wall 2 wall thickness (mm)*': { 'Value': '' , 'Tab': 'Wall Type 2' , 'Cell': 'E29'}
                , 'Wall 2 Insulation Present?*': { 'Value': '' , 'Tab': 'Wall Type 2' , 'Cell': 'E33'}
                , 'Wall 2 Insulation Type*': { 'Value': '' , 'Tab': 'Wall Type 2' , 'Cell': 'E35'}
                , 'Wall 2 Fill Type*': { 'Value': '' , 'Tab': 'Wall Type 2' , 'Cell': 'E37'}
                , 'Wall 2 Residual Cavity Width (mm)*': { 'Value': '' , 'Tab': 'Wall Type 2' , 'Cell': 'K37'}
                , 'Can Wall type 2 Insulation Thickness be Measured?*': { 'Value': '' , 'Tab': 'Wall Type 2' , 'Cell': 'K33'}
                , "If 'Yes' enter Wall type 2 insulation thickness (mm)*": { 'Value': '' , 'Tab': 'Wall Type 2' , 'Cell': 'K35'}
                , 'Wall Type 3': { 'Value': '' , 'Tab': 'Wall Type 3' , 'Cell': 'E27'}
                , 'Wall 3 wall thickness (mm)*': { 'Value': '' , 'Tab': 'Wall Type 3' , 'Cell': 'E29'}
                , 'Wall 3 Insulation Present?*': { 'Value': '' , 'Tab': 'Wall Type 3' , 'Cell': 'E33'}
                , 'Wall 3 Insulation Type*': { 'Value': '' , 'Tab': 'Wall Type 3' , 'Cell': 'E35'}
                , 'Wall 3 Fill Type*': { 'Value': '' , 'Tab': 'Wall Type 3' , 'Cell': 'E37'}
                , 'Wall 3 Residual Cavity Width (mm)*': { 'Value': '' , 'Tab': 'Wall Type 3' , 'Cell': 'K37'}
                , 'Can Wall type 3 Insulation Thickness be Measured?*': { 'Value': '' , 'Tab': 'Wall Type 3' , 'Cell': 'K33'}
                , "If 'Yes' enter Wall type 3 insulation thickness (mm)*": { 'Value': '' , 'Tab': 'Wall Type 3' , 'Cell': 'K35'}
                , 'Wall Type 4': { 'Value': '' , 'Tab': 'Wall Type 4' , 'Cell': 'E27'}
                , 'Wall 4 wall thickness (mm)*': { 'Value': '' , 'Tab': 'Wall Type 4' , 'Cell': 'E29'}
                , 'Wall 4 Insulation Present?*': { 'Value': '' , 'Tab': 'Wall Type 4' , 'Cell': 'E33'}
                , 'Wall 4 Insulation Type*': { 'Value': '' , 'Tab': 'Wall Type 4' , 'Cell': 'E35'}
                , 'Wall 4 Fill Type*': { 'Value': '' , 'Tab': 'Wall Type 4' , 'Cell': 'E37'}
                , 'Wall 4 Residual Cavity Width (mm)*': { 'Value': '' , 'Tab': 'Wall Type 4' , 'Cell': 'K37'}
                , 'Can Wall type 4 Insulation Thickness be Measured?*': { 'Value': '' , 'Tab': 'Wall Type 4' , 'Cell': 'K33'}
                , "If 'Yes' enter Wall type 4 insulation thickness (mm)*": { 'Value': '' , 'Tab': 'Wall Type 4' , 'Cell': 'K35'}
                , 'Is the property suitable for wall insulation? *': { 'Value': '' , 'Tab': 'Wall' , 'Cell': 'E41'}
                , 'No wall insulation details *': { 'Value': '' , 'Tab': 'Wall' , 'Cell': 'E43'}
                , 'Notes (Walls)': { 'Value': '' , 'Tab': 'Wall' , 'Cell': 'E54'}
                , 'EWI/IWI > 25% *': { 'Value': '' , 'Tab': 'Wall' , 'Cell': 'E47'}
                , 'Suitable for Draught Proofing': { 'Value': '' , 'Tab': 'Wall' , 'Cell': 'E49'}
                , 'Not suitable details Draughtproofing*': { 'Value': '' , 'Tab': 'Wall' , 'Cell': 'E51'}
                , 'Draught Proofing (<= 20m installed)': { 'Value': '' , 'Tab': 'Wall' , 'Cell': 'L68', 'Default': ''}
                , 'Draught Proofing (> 20m installed)': { 'Value': '' , 'Tab': 'Wall' , 'Cell': 'L69', 'Default': ''}
                , 'MEV 15l/s Bathroom': { 'Value': '' , 'Tab': 'Wall' , 'Cell': 'L74', 'Default': ''}
                , 'MEV 30l/s Utility': { 'Value': '' , 'Tab': 'Wall' , 'Cell': 'L75', 'Default': ''}
                , 'MEV 60l/s Kitchen': { 'Value': '' , 'Tab': 'Wall' , 'Cell': 'L76', 'Default': ''}
                , 'New Permanent Vent': { 'Value': '' , 'Tab': 'Wall' , 'Cell': 'L80', 'Default': ''}
                , 'New Background Vent': { 'Value': '' , 'Tab': 'Wall' , 'Cell': 'L81', 'Default': ''}
                , 'Duct Cooker Hood': { 'Value': '' , 'Tab': 'Wall' , 'Cell': 'L82', 'Default': ''}
                , 'Cavity Wall Insulation Bonded Bead': { 'Value': '' , 'Tab': 'Wall' , 'Cell': 'F86', 'Default': ''}
                , 'Loose Fibre Extraction': { 'Value': '' , 'Tab': 'Wall' , 'Cell': 'F87', 'Default': ''}
                , 'External Wall Insulation: Less than 60m2': { 'Value': '' , 'Tab': 'Wall' , 'Cell': 'F88', 'Default': ''}
                , 'External Wall Insulation: 60m2 to 85m2': { 'Value': '' , 'Tab': 'Wall' , 'Cell': 'F89', 'Default': ''}
                , 'External Wall Insulation: Greater than 85m2': { 'Value': '' , 'Tab': 'Wall' , 'Cell': 'F90', 'Default': ''}
                , 'ESB alteration': { 'Value': '' , 'Tab': 'Wall' , 'Cell': 'F91', 'Default': ''}
                , 'GNI meter alteration': { 'Value': '' , 'Tab': 'Wall' , 'Cell': 'F92', 'Default': ''}
                , 'New Gas Connection': { 'Value': '' , 'Tab': 'Wall' , 'Cell': 'F93', 'Default': ''}
                , 'RGI Meter_No Heating': { 'Value': '' , 'Tab': 'Wall' , 'Cell': 'F94', 'Default': ''}
                , 'Internal Wall Insulation: Vertical Surface': { 'Value': '' , 'Tab': 'Wall' , 'Cell': 'F99', 'Default': ''}
                , 'External wall insulation and CWI: less than 60m2': { 'Value': '' , 'Tab': 'Wall' , 'Cell': 'F104', 'Default': ''}
                , 'External wall insulation and CWI: 60m2 to 85m2': { 'Value': '' , 'Tab': 'Wall' , 'Cell': 'F105', 'Default': ''}
                , 'External wall insulation and CWI: greater than 85m2': { 'Value': '' , 'Tab': 'Wall' , 'Cell': 'F106', 'Default': ''}
                , 'replace_window_area': { 'Value': '' , 'Tab': 'Wall' , 'Cell': 'L113', 'Default': ''}
                , 'Notes (Windows and Doors)': { 'Value': '' , 'Tab': 'Wall' , 'Cell': 'E117', 'Default': ''}
                
                
                
                , 'Heating System *': { 'Value': '' , 'Tab': 'Heating' , 'Cell': 'E25'}
                , 'Qualifying Boiler': { 'Value': '' , 'Tab': 'Heating' , 'Cell': 'E27'}
                , 'Major Renovation': { 'Value': '' , 'Tab': 'Heating' , 'Cell': 'E29'}
                , 'System Age *': { 'Value': '' , 'Tab': 'Heating' , 'Cell': 'E31'}
                , 'Fully Working *': { 'Value': '' , 'Tab': 'Heating' , 'Cell': 'E33'}
                , 'Requires Service *': { 'Value': '' , 'Tab': 'Heating' , 'Cell': 'E37'}
                , 'Other Primary Heating Details *': { 'Value': '' , 'Tab': 'Heating' , 'Cell': 'I25'}
                , 'Not Working Details Primary Heating *': { 'Value': '' , 'Tab': 'Heating' , 'Cell': 'I33'}
                , 'Requires Service Details Primary Heating *': { 'Value': '' , 'Tab': 'Heating' , 'Cell': 'I37'}
                , 'Hot Water System Exists *': { 'Value': '' , 'Tab': 'Heating' , 'Cell': 'E44'}
                , 'HWS': { 'Value': '' , 'Tab': 'Heating' , 'Cell': 'G44'}
                , 'Other HW Details *': { 'Value': '' , 'Tab': 'Heating' , 'Cell': 'I49'}
                , 'Hot Water Cylinder*': { 'Value': '' , 'Tab': 'Heating' , 'Cell': 'E55'}
                , 'Insulation *': { 'Value': '' , 'Tab': 'Heating' , 'Cell': 'E57'}
                , 'Condition of Lagging Jacket *': { 'Value': '' , 'Tab': 'Heating' , 'Cell': 'H55'}
                , 'HWC Controls *': { 'Value': 'Cylinder Thermostat Controls *' , 'Tab': 'Heating' , 'Cell': 'H57'}
                , 'Heating Systems Controls *': { 'Value': '' , 'Tab': 'Heating' , 'Cell': 'E62'}
                , 'Partial Details *': { 'Value': '' , 'Tab': 'Heating' , 'Cell': 'E64'}
                , 'Programmer / Timeclock *': { 'Value': '' , 'Tab': 'Heating' , 'Cell': 'H62'}
                , 'Room Thermostat Number *': { 'Value': '' , 'Tab': 'Heating' , 'Cell': 'H64'}
                , 'Rads Number *': { 'Value': '' , 'Tab': 'Heating' , 'Cell': 'K62'}
                , 'TRVs Number *': { 'Value': '' , 'Tab': 'Heating' , 'Cell': 'K64'}
                , 'Suitable for Heating Measures *': { 'Value': '' , 'Tab': 'Heating' , 'Cell': 'E68'}
                , 'Not suitable details*': { 'Value': '' , 'Tab': 'Heating' , 'Cell': 'E70'}
                , 'Notes (Heating)': { 'Value': '' , 'Tab': 'Heating' , 'Cell': 'E74'}
                
                
                , 'Basic gas heating system': { 'Value': '' , 'Tab': 'Heating' , 'Cell': 'K81', 'Default': ''}
                , 'Basic oil heating system': { 'Value': '' , 'Tab': 'Heating' , 'Cell': 'K85', 'Default': ''}
                , 'Full gas heating system installation': { 'Value': '' , 'Tab': 'Heating' , 'Cell': 'K89', 'Default': ''}
                , 'Full oil heating system installation': { 'Value': '' , 'Tab': 'Heating' , 'Cell': 'K114', 'Default': ''}
                , 'Gas boiler and controls (Basic & controls pack)': { 'Value': '' , 'Tab': 'Heating' , 'Cell': 'K118', 'Default': ''}
                , 'Hot Water Cylinder Jacket': { 'Value': '' , 'Tab': 'Heating' , 'Cell': 'K127', 'Default': ''}
                , 'Oil boiler and controls (Basic & controls pack)': { 'Value': '' , 'Tab': 'Heating' , 'Cell': 'K167', 'Default': ''}
                
                
                
                , 'Secondary Heating System': { 'Value': '' , 'Tab': 'Heating - Secondary' , 'Cell': 'E25'}
                , 'Secondary System Age *': { 'Value': '' , 'Tab': 'Heating - Secondary' , 'Cell': 'E27'}
                , 'Secondary System Fully Working *': { 'Value': '' , 'Tab': 'Heating - Secondary' , 'Cell': 'E29'}
                , 'Secondary System Requires Service *': { 'Value': '' , 'Tab': 'Heating - Secondary' , 'Cell': 'E33'}
                , 'Not Working Details Secondary Heating *': { 'Value': '' , 'Tab': 'Heating - Secondary' , 'Cell': 'I29'}
                , 'Requires Service Details Secondary Heating *': { 'Value': '' , 'Tab': 'Heating - Secondary' , 'Cell': 'I33'}

                , 'Number of habitable rooms in the property': { 'Value': '' , 'Tab': 'Mechanical Ventilation Systems' , 'Cell': 'D55'}
                , 'Number of wet rooms in the property': { 'Value': '' , 'Tab': 'Mechanical Ventilation Systems' , 'Cell': 'D57'}
                , 'No. of habitable/wet rooms w/ open flued appliance': { 'Value': '' , 'Tab': 'Mechanical Ventilation Systems' , 'Cell': 'D59'}
                , 'LED Bulbs: supply only (4 no.)': { 'Value': '' , 'Tab': 'Mechanical Ventilation Systems' , 'Cell': 'E31', 'Default': ''}
                , 'Air-tightness test recommended?': { 'Value': '' , 'Tab': 'Mechanical Ventilation Systems' , 'Cell': 'E27', 'Default': ''}

                , 'Adequate Access*': { 'Value': '' , 'Tab': 'Supplementary' , 'Cell': 'D25'}
                , 'No Access Details*': { 'Value': '' , 'Tab': 'Supplementary' , 'Cell': 'I25'}
                , 'Cherry Picker Required*': { 'Value': '' , 'Tab': 'Supplementary' , 'Cell': 'D27'}
                , 'Cherry Picker Required Details*': { 'Value': '' , 'Tab': 'Supplementary' , 'Cell': 'I27'}
                , 'Mould/Mildew identified by surveyor; or reported by the applicant*': { 'Value': '' , 'Tab': 'Supplementary' , 'Cell': 'D29'}
                , 'Mould/Mildew Details*': { 'Value': '' , 'Tab': 'Supplementary' , 'Cell': 'I29'}
                , 'As confirmed by homeowner; property is a protected structure*': { 'Value': '' , 'Tab': 'Supplementary' , 'Cell': 'D31'}
                , 'Protected Structure Details*': { 'Value': '' , 'Tab': 'Supplementary' , 'Cell': 'I31'}

                # , 'Surveyor Signature': { 'Value': '' , 'Tab': 'Declarations' , 'Cell': 'D27'}
                # , 'Surveyor Signature Date': { 'Value': '' , 'Tab': 'Declarations' , 'Cell': 'L27'}
                # , 'Customer Signature': { 'Value': '' , 'Tab': 'Declarations' , 'Cell': 'D29'}
                # , 'Customer Signature Date': { 'Value': '' , 'Tab': 'Declarations' , 'Cell': 'L29'}

                }
            
        if template == 'template_mrc':
            filename = json_val_dict['plan_name'] + ' Major Renovation calculation.xlsx'
            return_filename = filename
            # container_name = "project-files"
            container_name = 'attachment'
            # local_path = json_val_dict['plan_name']
            local_path = "/tmp"
            print('local_path', ':', local_path)
            # instance_file_path = filename
            instance_file_path = os.path.join(local_path, filename)
            
            # instance_file_path = instance_file_path.replace('\\\\', '\\')
            # instance_file_path = instance_file_path.replace('\\', '/')
            # print('instance_file_path', ':', instance_file_path)
            
            v = {
                'plan_name': { 'Value': '' , 'Tab': 'Results' , 'Cell': 'D2'}
                , 'Thermal Envelope - Heat loss walls, windows and doors': { 'Value': '' , 'Tab': 'Results' , 'Cell': 'E6'}
                , 'Thermal Envelope - Heat loss floor area': { 'Value': '' , 'Tab': 'Results' , 'Cell': 'E7'}
                , 'Thermal Envelope - Heat loss roof area': { 'Value': '' , 'Tab': 'Results' , 'Cell': 'E8'}
                , 'THERMAL ENVELOPE OF BUILDING AREA': { 'Value': '' , 'Tab': 'Results' , 'Cell': 'E9'}
                , 'Heat loss Wall Area recommended for EWI and IWI': { 'Value': '' , 'Tab': 'Results' , 'Cell': 'E12'}
                , 'New Windows being recommended for replacement': { 'Value': '' , 'Tab': 'Results' , 'Cell': 'E13'}
                , 'TOTAL SURFACE AREA FOR MAJOR RENOVATION WORKS': { 'Value': '' , 'Tab': 'Results' , 'Cell': 'E14'}
                , 'Total surface area for MR works / Thermal Envelope': { 'Value': '' , 'Tab': 'Results' , 'Cell': 'E17'}
                , 'WARMER HOMES MAJOR RENOVATION RESULT': { 'Value': '' , 'Tab': 'Results' , 'Cell': 'E18'}
                
                
                }
        # print("json_val_dict['Thermal Envelope - Heat loss walls, windows and doors']", ':', json_val_dict['Thermal Envelope - Heat loss walls, windows and doors'])
        for field in v:
            # print('field', ':', field)
            if field in json_val_dict.keys():
                v[field]['Value'] = json_val_dict[field]
                # print("v[field]['Value']", ':', v[field]['Value'])
                
                if v[field]['Value'] in ['', 'N/A', 0]:
                    if "Default" in v[field].keys():
                        v[field]['Value'] = v[field]['Default']
                
                if type(v[field]['Value']) == str and '<BR>' in v[field]['Value']:
                    v[field]['Value'] = v[field]['Value'].replace('<BR>', '\n')
                # print(field, ':', v[field]['Value'])

        # print(v)


        # Create the BlobServiceClient object
        blob_service_client = BlobServiceClient(account_url, credential=default_credential)
        container_client = blob_service_client.get_container_client(container= container_name) 

        # if not os.path.exists(local_path):
            # print('about to create local_path', ':', local_path)
        try:
            error = ''
            os.mkdir(local_path)
        except Exception as ex:
            error = traceback.format_exc()
            # print('error', ':', error)
            
        
        
        
        print("\nDownloading blob to \n\t" + instance_file_path)
        
        
        template_file_name = template + '.xlsx'
        with open(file=instance_file_path, mode="wb") as download_file:
            download_file.write(container_client.download_blob(template_file_name).readall())

        xfile = openpyxl.load_workbook(instance_file_path)
        
        for field in v:
            # print(field)
            sheet = xfile[v[field]['Tab']]
            sheet[v[field]['Cell']] = v[field]['Value']



        xfile.save(instance_file_path)


        instance_file_path = os.path.join(local_path, filename)
        with open(file=instance_file_path, mode="rb") as upload_file:
            blob_client = blob_service_client.get_blob_client(container=container_name, blob=instance_file_path)
            blob_client.upload_blob(upload_file, overwrite=True)
    

        
        print('created')
        
        
        if template == 'template_mrc':
            output = copy_from_container(json_val_dict['plan_name'])
        
        
    except Exception as ex:
        # exc_type, exc_obj, exc_tb = sys.exc_info()
        # output = "Line " + str(exc_tb.tb_lineno) + ": " + exc_type 
        
        # output = str(ex)
        output = traceback.format_exc()
        # LOGGER.info('Exception : ' + str(traceback.format_exc()))
        
        # fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        # print(exc_type, fname, exc_tb.tb_lineno)
        print(output)

    finally:
        return output, return_filename



def copy_from_container(plan_name):

    try:
        output = ''
        account_url = "https://ksnmagicplanfunc3e54b9.blob.core.windows.net"
        default_credential = DefaultAzureCredential()
        container_from = 'attachment'
        local_path_from = "/tmp"
        container_to = 'project-files'
        local_path_to = plan_name
        
        filenames = [plan_name + '.xlsx', plan_name + ' Major Renovation calculation' + '.xlsx']
        # filename = plan_name + ' Major Renovation calculation' + '.xlsx'
        
        for filename in filenames:
            instance_file_path_from = os.path.join(local_path_from, filename)
            print('instance_file_path_from', ':', instance_file_path_from)
            instance_file_path_to = os.path.join(local_path_to, filename)
            print('instance_file_path_to', ':', instance_file_path_to)



            # Create the BlobServiceClient object
            blob_service_client = BlobServiceClient(account_url, credential=default_credential)
            container_client = blob_service_client.get_container_client(container= container_from) 
            
            
            file_content = container_client.download_blob(instance_file_path_from).readall()
            print('got file_content')

            blob_client = blob_service_client.get_blob_client(container=container_to, blob=instance_file_path_to)
            blob_client.upload_blob(file_content, overwrite=True)

        # with open(file=instance_file_path_to, mode="rb") as upload_file:
            # blob_client = blob_service_client.get_blob_client(container=container_to, blob=instance_file_path_from)
            # blob_client.upload_blob(upload_file, overwrite=True)

    except:
        output = traceback.format_exc()

    finally:
        # print('output', ':', output)
        return output
        

def lot(output_dict):
    try:
        Lot = 'S' # No works recommended
        
        recommended_works_dict = {
            "Internal Wall Insulation: Sloped or flat (horizontal) surface": "a"
            , "Attic (Loft) Insulation 100 mm top-up": "a"
            , "Attic (Loft) Insulation 150 mm top-up": "a"
            , "Attic (Loft) Insulation 200 mm top-up": "a"
            , "Attic (Loft) Insulation 250 mm top up": "a"
            , "Attic (Loft) Insulation 300 mm": "a"
            , "Cavity Wall Insulation Bonded Bead": "c"
            , "Internal Wall Insulation: Vertical Surface": "I"
            , "Basic gas heating system": "H"
            , "Basic oil heating system": "H"
            , "Full gas heating system installation": "H"
            , "Full oil heating system installation": "H"
            , "Gas boiler and controls (Basic & controls pack)": "H"
            , "Oil boiler and controls (Basic & controls pack)": "H"
            , "External Wall Insulation: Less than 60m2": "E"
            , "External Wall Insulation: 60m2 to 85m2": "E"
            , "External Wall Insulation: Greater than 85m2": "E"
            , "External wall insulation and CWI: less than 60m2": "E"
            , "External wall insulation and CWI: 60m2 to 85m2": "E"
            , "External wall insulation and CWI: greater than 85m2": "E"
            , "Window (same m2 rate will apply to windows with certified trickle vents)": "w"
        }
        Lot_upper = ''
        Lot_lower = ''
        for k in recommended_works_dict:
            v = recommended_works_dict[k]
            if k in output_dict.keys() and output_dict[k] not in ['N/A', 0, ""] and v not in (Lot_upper + Lot_lower):
                # print('adding: ', output_dict[k])
                if v.isupper():
                    Lot_upper += v
                else:
                    Lot_lower += v
        
        Lot_upper = Lot_upper.replace("IE", "E")
        Lot_upper = Lot_upper.replace("IH", "H")
        if Lot_upper == '':
            Lot_upper = 'S'
        else:
            Lot_lower = Lot_lower.replace('a', '')
            Lot_lower = Lot_lower.replace('c', '')
        Lot = Lot_upper + Lot_lower
        
        
        # print(Lot)
        
        # IEacw
        
        valid_lots = ['Sa', 'Sc', 'Sac', 'Scw', 'Sacw', 'I', 'Iw', 'E', 'Ew', 'H', 'Hw', 'HE', 'HEw', 'S']
        if Lot not in valid_lots:
            # print('invalid')
            Lot = Lot + ' (invalid)'
            print(list(valid_lots))

        

        # No longer in use: Sw, a, b, c
    except:
        Lot = traceback.format_exc()

    
    return Lot


def exterior_walls(root):
    ext_wall_area_gross = 0
    plan_name = root.get('name')
    interior_wall_width = root.get('interiorWallWidth') # always available?
    exteriorWallWidth = float(root.get('exteriorWallWidth')) # always available?
    extern_width_offset = interior_wall_width * 4
    extern_perim = 0
    exploded_wall_dict = {}
    
    floors = root.findall('interiorRoomPoints/floor')
    # floors = root.findall('floor')
    # print('len(floors)', ':', len(floors))
    for floor in floors:
        floor_type = floor.get('floorType')
        ft = floor_type
        if floor_type not in ['10', '11', '12', '13']:
            continue
        exterior_walls = [] # {} 
        # print('floor_type', ':', floor_type)
        walls = floor.findall('exploded/wall')
        exploded_wall_dict[ft] = {}
        for i, wall in enumerate(walls):
            exploded_wall_dict[ft][i] = {}
            w_type = wall.find('type').text
            points = wall.findall('point') 
            p1, p2, *rest = points
            x1 = float(p1.get('x'))
            y1 = -float(p1.get('y'))
            x2 = float(p2.get('x'))
            y2 = -float(p2.get('y'))
            length = cart_distance((x1, y1), (x2, y2)) - (0.25 * exteriorWallWidth)
            # print(length)
            # wall_height = (float(p1.get('height')) + float(p2.get('height'))) / 2
            if floor_type == '10':
                wall_height = 2.4
            if floor_type == '11':
                wall_height = 2.7
            if floor_type in ['12', '13']:
                wall_height = 2
                
            area = wall_height * length
            # print('length ' + str(i) , ':', length)
            # print('wall_height ' + str(i) , ':', wall_height)
            # print('area ' + str(i) , ':', area)
            
            (x1, y1), (x2, y2)
            
            
            exploded_wall_dict[ft][i]['x1'] = x1
            exploded_wall_dict[ft][i]['y1'] = y1
            exploded_wall_dict[ft][i]['x2'] = x2
            exploded_wall_dict[ft][i]['y2'] = y2
            exploded_wall_dict[ft][i]['type'] = w_type
            
            if w_type == 'exterior':
                ext_wall_area_gross += wall_height * length
                extern_perim += length
                # print(w_type, x1, y1, x2, y2, wall_height)
                # print(w)
                # exterior_walls.append(w)
        # print(exterior_walls)
        
    
    
            # extern_perim -= extern_width_offset
            # floors_perims.append(extern_perim)
            # wall_height = wall_height/nwalls if nwalls != 0 else wall_height
            # floors_heights.append(wall_height)
            

            # ext_wall_area_gross -= wall_types['Party Wall Area'][floor_index_adj] if 'Party Wall Area' in wall_types else 0
            # ext_wall_area_gross -= wall_types['Internal Wall Area'][floor_index_adj] if 'Internal Wall Area' in wall_types else 0

            # walls_area_gross.append(wall_area_gross)
        # print('ext_wall_area_gross', ':', str(ext_wall_area_gross))
        # print('extern_perim', ':', str(extern_perim))
    return ext_wall_area_gross, exploded_wall_dict





