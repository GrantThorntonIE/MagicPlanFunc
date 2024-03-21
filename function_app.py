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

MAX_REAL_FLOORS = 10




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
                output += f'<tr><td>{item}</td>'
                value = dict[item] if item in dict.keys() else 'N/A'
                if (type(value) == bool and value == True):
                    value = "Yes"
                if (type(value) == bool and value == False):
                    value = "No"
                output += f'<td>{value}</td>'
                # print(item, value)
        else:
            for i, key in enumerate(dict):
                # print(key, dict[key])
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
    json_val_dict["Not suitable details*"] = ''
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
        # print(f"Roof {n} Thickness (mm)", ": ",  e)
        
        # value: Area - add to appropriate dict entry
        # print(f"roof_{n}_area", ": ", json_val_dict[f"roof_{n}_area"])
        for t in [100, 150, 200, 250, 300]:
            if e + t >= 300:
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
                # print('n', ': ', n, json_val_dict[f"Roof Type {n} Not Suitable Details"])
                json_val_dict["Not suitable details*"] += f"Roof Type {n} Not Suitable Details: "
                json_val_dict["Not suitable details*"] += json_val_dict[f"Roof Type {n} Not Suitable Details"]
                json_val_dict["Not suitable details*"] += "<BR>"
            if f"Roof Type {n} Sloping Ceiling Not Suitable Details*" in json_val_dict.keys():
                # print('n', ': ', n, json_val_dict[f"Roof Type {n} Sloping Ceiling Not Suitable Details*"])
                json_val_dict["Not suitable details*"] += f"Roof Type {n} Sloping Ceiling Not Suitable Details: "
                json_val_dict["Not suitable details*"] += json_val_dict[f"Roof Type {n} Sloping Ceiling Not Suitable Details*"]
                json_val_dict["Not suitable details*"] += "<BR>"
        else:
            json_val_dict["Not suitable details*"] = 'N/A'
    # print('json_val_dict["Not suitable details*"]: ', json_val_dict["Not suitable details*"])
    
    for n in range(1, 5):
        if f"Notes (Roof Type {n})" in json_val_dict.keys():
            json_val_dict["Notes (Roof)"] += f"Notes (Roof Type {n}): "
            json_val_dict["Notes (Roof)"] += json_val_dict[f"Notes (Roof Type {n})"]
            json_val_dict["Notes (Roof)"] += "<BR>"
        if f"Notes (Roof Type {n})*" in json_val_dict.keys():
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
                # print('n', ': ', n, json_val_dict[f"No wall type {n} insulation details"])
                json_val_dict["No wall insulation details *"] += f"No wall type {n} insulation details: "
                json_val_dict["No wall insulation details *"] += json_val_dict[f"No wall type {n} insulation details"]
                json_val_dict["No wall insulation details *"] += "<BR>"
        else:
            json_val_dict["No wall insulation details *"] = 'N/A'
    # print('json_val_dict["No wall insulation details *"]: ', json_val_dict["No wall insulation details *"])
    
    for n in range(1, 5):
        if f"Notes (Wall type {n} Walls)" in json_val_dict.keys():
            json_val_dict["Notes (Walls)"] += f"Notes (Wall type {n} Walls): "
            json_val_dict["Notes (Walls)"] += json_val_dict[f"Notes (Wall type {n} Walls)"]
            json_val_dict["Notes (Walls)"] += "<BR>"
        if f"Notes (Wall type {n} Walls)*" in json_val_dict.keys():
            json_val_dict["Notes (Walls)"] += f"Notes (Wall type {n} Walls)*: "
            json_val_dict["Notes (Walls)"] += json_val_dict[f"Notes (Wall type {n} Walls)*"]
            json_val_dict["Notes (Walls)"] += "<BR>"
    # print('json_val_dict["Notes (Walls)"]: ', json_val_dict["Notes (Walls)"])
    
    if json_val_dict["Is the property suitable for wall insulation? *"] == False:
        json_val_dict["No wall insulation details *"] += json_val_dict["Notes (Walls)"]




def XML_2_dict(root):
    try:
        d = {}
        floors = root.findall('interiorRoomPoints/floor')
        for floor in floors:
            # print('floorType: ' + floor.get('floorType'))
            # print('uid: ' + floor.get('uid'))
            d[floor.get('floorType')] = floor.get('uid')
            d[floor.get('uid')] = floor.get('floorType')
            for room in floor.findall('floorRoom'):
                # print('type: ' + room.get('type'))
                # print('uid: ' + room.get('uid'))
                if room.get('type') not in d.keys():
                    d[room.get('type')] = []
                d[room.get('type')].append(room.get('uid'))
                d[room.get('uid')] = room.get('type')
    
    except Exception as ex:
        output = str(ex) + "\n\n" + traceback.format_exc()
        # LOGGER.info('Exception : ' + str(traceback.format_exc()))
        print(output)
    
    finally:
        return d


def ber_old(root):
    

    lookup = {
        'LED/CFL'           : 'co-3a9c9ff6-2bad-4d62-9526-1df98538cbad',
        'Halogen Lamp'      : 'co-94486aec-b47a-4d75-aaf3-0645576bae56',
        'Halogen LV'        : 'co-b21a94da-ad62-40e5-bfe0-c1aa0b8461d5',
        'Linear Fluorescent': 'co-497bde35-eb4a-41ec-ba91-b24e35099799',
        'Incandescent'      : 'co-44a1cdea-ff05-40a8-afb5-fe5b9c7f086a',
        'EMV'               : 'co-0c7d0ada-8a17-41f9-8746-e7007a1c40b1',
        'NMV'               : 'co-accd48a4-43b8-4381-b569-c8404f52dec5',
        'NPV'               : 'co-483ab20e-2762-4733-9db5-19d21e1d090d',
        'NBV'               : 'co-88da8a48-da48-4c3d-a459-ee8eef96bcdd',
        'EPV'               : 'co-4d2e52df-c793-4c02-953a-f4ed0b7eaae0',
        'DCH'               : 'co-03b80e12-32b7-45be-8b44-9b4b03a09b4c',
        'EMVB'              : 'co-33fd6b69-25ae-4e55-bf7b-f91af6112ac4',
        'ECHB'              : 'co-d0f4acd2-8598-49ec-ac3c-8b39edc724e9',
        'Chimney'           : 'co-5651e005-f73c-4d26-8b75-6d07291d7839',
        'Flue'              : 'co-9276562f-535a-4c1f-a708-d869f581194d',
        'Flueless'          : 'co-86083af5-5a40-4831-9990-4f60aa537d99'
    }

    internal_width = float(root.get('interiorWallWidth'))
    extern_width_offset = internal_width * 4

    plan_name = root.get('name')
    walls_area_net = []
    walls_area_gross = []
    floors_heights = []
    floors_perims = []
    floors_area = []
    doors_area = []
    windows_area = []

    led_count = []
    lf_count = []
    inc_count = []
    hlv_count = []
    hl_count = []

    in_f_count = []
    pnc_count = []
    disc_vent_count = []
    total_vent_count = []

    flue_count = []
    chimney_count = []
    flueless_count = []

    rad_count = []
    rad_trv_count = []
    rs_count = []
    programmer_count = []
    er_count = []
    esh_count = []

    bath_count = []
    ies_count = []
    msv_count = []
    msvp_count = []
    msu_count = []

    floor_enum = ['Floor']
    real_floor_enum = ['Floor']
    imaginary_floor_enum = ['Floor']

    floor_index = 0
    object_floor_enum = ['Name']

    living_room_area = 0

    wall_types = {}
    floor_table : dict[str, list[float]] = {}
    colours : dict[str, list[float]] = {}
    roof_table : dict[str, list[float]] = {}
    window_door_table : dict[str, list[float]] = {}

    floors = root.findall('interiorRoomPoints/floor')
    empty_array = [0] * len([floor for floor in floors if not re.search('[0-9][012346789]', floor.find('name').text)])
    imaginary_array = [0] * len([floor for floor in floors if re.search('[0-9][012346789]', floor.find('name').text)])
    full_floor_array = [0] * len(floors)


    wd = pd.DataFrame(None, columns=['Window Type', 'Number of Openings', 'Number of Openings Draught Stripped', 'In Roof', 'Shading', 'Orientation', 'Area'])

    dt = pd.DataFrame(None, columns=['Type', 'Number of Openings', 'Number of Openings Draught Stripped', 'Glazed Area', 'Glazed Area (%)', 'Glazing Type', 'U-Value', 'Area'])

    for floor in floors:

        floor_name = floor.find('name').text
        floor_index_adj : int
        floor_num : int
        if floor_name == 'Ground Floor':
            floor_index_adj = 0
            floor_num = 0
        elif floor_name[0:2].isnumeric():
            floor_index_adj = int(floor_name[0:2]) - 10
            floor_num = int(floor_name[0:2])
        elif floor_name[0].isnumeric():
            floor_num = int(floor_name[0])
            floor_index_adj = int(floor_name[0])
        else:
            floor_index_adj = None
            floor_num = None
        

        
        extern_perim = 0
        wall_height = 0
        window_area = 0
        nwalls = 0
        door_area = 0
        wall_area_gross = 0

        room_points = floor.findall('floorRoom/point/values/value[@key="qcustomfield.e8660a0cq0.lo6b23iucno"]../../..')
        for room in floor.findall('floorRoom/values/value[@key="ground.color"]../..'):
            colour = room.find('values/value[@key="ground.color"]').text
            area = float(room.get('area'))
            if colour not in colours:
                colours[colour] = full_floor_array.copy()
            colours[colour][floor_index] += area

        if floor_num != None and floor_num < MAX_REAL_FLOORS:
            for room in floor.findall('floorRoom[@type="Living Room"]'):
                living_room_area += float(room.get('area')) if room.get('area') != None else 0

            real_floor_enum.append(floor_name)

            floor_area = float(floor.get('areaWithInteriorWallsOnly'))
            floors_area.append(floor_area)
            windows_doors = floor.findall('symbolInstance')
            windows_doors = [window for window in windows_doors if ('W' in window.get('id') or 'F' in window.get('id'))]

            for window in windows_doors:
                id = window.get('id')
                # LOGGER.info('id: ' + id)
                symbol = window.get('symbol')
                if symbol == None or 'window' not in symbol:
                    if window.find('values/value[@key="clonedFrom"]') == None:
                        continue
                    if 'window' not in window.find('values/value[@key="clonedFrom"]').text:
                        continue
                    
                wall_elem : ET.Element

                wall_elem = window.find('values/value[@key="qcustomfield.bebb2096q3"]')

                wall_type = ''
                if wall_elem != None:
                    wall_type = wall_elem.text[-1] if not wall_elem.text[-2].isnumeric() else wall_elem.text[-2:]
                
                window_elem = floor.find(f'exploded/window[@symbolInstance="{id}"]')
                if window_elem == None:
                    window_elem = floor.find(f'exploded/furniture[@symbolInstance="{id}"]')
                if window_elem == None:
                    window_elem = floor.find(f'floorRoom/window[@symbolInstance="{id}"]') # another place you might find it
                if window_elem == None:
                    window_elem = floor.find(f'floorRoom/door[@symbolInstance="{id}"]') # another place you might find it

                
                area = float(window_elem.get('height')) * float(window_elem.get('width'))
                if wall_type not in window_door_table and wall_type != '':
                    window_door_table[wall_type] = empty_array.copy()
                if wall_type != '':
                    window_door_table[wall_type][floor_index_adj] += area

                shading_type : ET.Element
                window_type : ET.Element
                direction : ET.Element
                openings_elem : ET.Element
                ds_openings_elem : ET.Element
                in_roof : bool

                cloned_from = window.find('values/value[@key="clonedFrom"]')
                if cloned_from == None or not 'skylight' in window.find('values/value[@key="clonedFrom"]').text:
                    shading_type = window.find('values/value[@key="qcustomfield.bebb2096q0.vvvvtj3gbp8"]')
                    window_type = window.find('values/value[@key="qcustomfield.bebb2096q2"]')
                    direction = window.find('values/value[@key="qcustomfield.bebb2096q0.b8o7vbr534"]')
                    openings_elem = window.find('values/value[@key="qcustomfield.bebb2096q0.47fm2211clg"]')
                    ds_openings_elem = window.find('values/value[@key="qcustomfield.bebb2096q0.shu7ct5p1l8"]')
                    in_roof = False
                else:
                    shading_type = window.find('values/value[@key="qcustomfield.91cb4548q0.d5skr1o2ol"]')
                    window_type = window.find('values/value[@key="qcustomfield.91cb4548q0.knium9uou08"]')
                    direction = window.find('values/value[@key="qcustomfield.91cb4548q0.p2meoelvuao"]')
                    openings_elem = window.find('values/value[@key="qcustomfield.91cb4548q0.073aprtkrs8"]')
                    ds_openings_elem = window.find('values/value[@key="qcustomfield.91cb4548q0.v88utngglp"]')
                    in_roof = True

                openings = 0 if openings_elem == None else int(openings_elem.text)
                ds_openings = 0 if ds_openings_elem == None else int(ds_openings_elem.text)

                if window_type != None and direction != None:
                    shading_type_text = ''
                    if shading_type == None:
                        shading_type_text = 'Average or Unknown 20 60'
                    else:
                        shading_type_text = shading_type.text.replace('.', ' ')
                    direction_text = direction.text.replace('.', ' ')
                    window_type_int = int(window_type.text.split('Type.')[1])
                    if ((wd['Window Type'] == window_type_int) & (wd['In Roof'] == in_roof) & \
                        (wd['Shading'] == shading_type_text) & (wd['Orientation'] == direction_text)).any():
                        index = wd.index[(wd['Window Type'] == window_type_int) & (wd['In Roof'] == in_roof) & \
                                 (wd['Shading'] == shading_type_text) & (wd['Orientation'] == direction_text)].to_list()[0]
                        wd.loc[index, 'Number of Openings'] += openings
                        wd.loc[index, 'Number of Openings Draught Stripped'] += ds_openings
                        wd.loc[index, 'Area'] += area
                    else:
                        wd.loc[len(wd.index)] = [
                            window_type_int,
                            openings,
                            ds_openings,
                            in_roof,
                            shading_type_text,
                            direction_text,
                            area
                        ]

            door_question_key = {
                'g_t' : {
                    'Solid.Exposed.Door.30.60.Glazed' : 'qcustomfield.ddc14d2eq0.vmacape1ks',
                    'Solid.Semi.Exposed.Glazed.Door.30.60.Glazed' : 'qcustomfield.ddc14d2eq0.ij3dcce5clo'
                },
                'u_v' : {
                    'Solid.Exposed.Door' : 'qcustomfield.ddc14d2eq0.0v6l9n35trg',
                    'Solid.Semi.Exposed.Door' : 'qcustomfield.ddc14d2eq0.pl6roqhqj3o',
                    'Solid.Exposed.Door.30.60.Glazed' : '',
                    'Solid.Semi.Exposed.Glazed.Door.30.60.Glazed' : '',
                    'Metal.Uninsulated.Garage.Door' : 'qcustomfield.ddc14d2eq0.o51v05s6veg',
                    'Certified.Door.Data' : '',
                },
                'g_a' : {
                    'Solid.Exposed.Door.30.60.Glazed' : 'qcustomfield.ddc14d2eq0.7r2dd1lsr7o',
                    'Solid.Semi.Exposed.Glazed.Door.30.60.Glazed' : 'qcustomfield.ddc14d2eq0.e6oefhpmmjo'
                }
            }

            for door in windows_doors:
                id = door.get('id')
                # LOGGER.info('door id: ' + id)

                symbol = door.get('symbol')
                # LOGGER.info('door symbol: ' + symbol)
                if symbol == None or 'door' not in symbol:
                    if door.find('values/value[@key="clonedFrom"]') == None:
                        continue
                    if 'door' not in door.find('values/value[@key="clonedFrom"]').text:
                        continue
                    
                door_type = door.find('values/value[@key="qcustomfield.ddc14d2eq0.31bdk91s35o"]')
                if door_type == None:
                    continue
                door_type_text = door_type.text

                u_value = door.find(f'values/value[@key="{door_question_key["u_v"][door_type_text]}"]')
                n_openings = door.find('values/value[@key="qcustomfield.ddc14d2eq0.lko7143kejg"]')
                n_openings_ds = door.find('values/value[@key="qcustomfield.ddc14d2eq0.84vs7q5icu"]')

                glazed_area : ET.Element = None
                glazing_type : ET.Element = None
                if 'Glazed' in door_type_text:
                    glazed_area = door.find(f'values/value[@key="{door_question_key["g_a"][door_type_text]}"]')
                    glazing_type = door.find(f'values/value[@key="{door_question_key["g_t"][door_type_text]}"]')

                u_value_text = u_value.text if u_value != None else 'N/A'
                glazed_area_val = float(glazed_area.text) if glazed_area != None else 0
                glazing_type_text = glazing_type.text if glazing_type != None else 'N/A'
                n_openings_int = int(n_openings.text) if n_openings != None else 0
                n_openings_ds_int = int(n_openings_ds.text) if n_openings_ds != None else 0

                door_elem = floor.find(f'exploded/door[@symbolInstance="{id}"]')
                if door_elem == None:
                    door_elem = floor.find(f'floorRoom/door[@symbolInstance="{id}"]') # another place you might find it
                
                area = float(door_elem.get('height')) * float(door_elem.get('width'))
                door_area += area

                dt.loc[len(dt.index)] = [
                    door_type_text.replace('.', ' '), 
                    n_openings_int,
                    n_openings_ds_int,
                    glazed_area_val,
                    glazed_area_val/area * 100,
                    glazing_type_text,
                    u_value_text,
                    area
                ]

                wall_elem : ET.Element
                wall_elem = door.find(f'values/value[@key="qcustomfield.ddc14d2eq1"]')
                if wall_elem == None:
                    continue
                
                wall_type = wall_elem.text[-1]
                if wall_type not in window_door_table:
                    window_door_table[wall_type] = empty_array.copy()
                window_door_table[wall_type][floor_index_adj] += area

        if floor_num != None and floor_num >= MAX_REAL_FLOORS and floor_name != '15th Floor':
            for room in room_points:
                points = room.findall('point/values/value[@key="qcustomfield.e8660a0cq0.lo6b23iucno"]../..')
                all_points = room.findall('point')
                floor_type_elem = room.find('values/value[@key="qcustomfield.86272860q0.rc9aflbaq2"]')
                floor_type = ''

                if floor_type_elem != None:
                    floor_type = floor_type_elem.text.replace('.', ' ')
                if floor_type != '':
                    floor_key = f'{floor_type} Area'
                    if floor_key not in floor_table:
                        floor_table[floor_key] = imaginary_array.copy()
                    floor_table[floor_key][floor_index_adj] += float(room.get('area')) if room.get('area') != None else 0

                for point in points:
                    wall_type = point.find('values/value[@key="qcustomfield.e8660a0cq0.lo6b23iucno"]')
                    if len(wall_type.text) == 3:
                        continue
                    w_type = wall_type.text[-1] if wall_type.text[-1].isnumeric() else wall_type.text.replace('.', ' ')
                    if wall_type.text[-2].isnumeric():
                        w_type = wall_type.text[-2:]
                    type_area = f'Wall Type {w_type} Area Gross'
                    perim = f'Wall Type {w_type} Perimeter'
                    x1 = float(point.get('snappedX'))
                    y1 = float(point.get('snappedY'))
                    next_index = all_points.index(point) + 2 #Element Tree Indexes from 1, this index returns the index from 0, to get the next element we add 2.
                    next = room.find(f'point[{next_index}]')
                    if next == None:
                        next = room.find('point[1]')
                    x2 = float(next.get('snappedX'))
                    y2 = float(next.get('snappedY'))
                    height = (float(point.get('height')) + float(next.get('height')))/2
                    wall_length = cart_distance((x1, y1), (x2, y2))
                    wall_area = wall_length * height

                    if floor_type != '' and w_type not in ['', 'Party Wall', 'Internal Wall'] and floor_name == '10th Floor':
                        floor_perim_key = f'{floor_type} Perim'
                        if floor_perim_key not in floor_table:
                            floor_table[floor_perim_key] = imaginary_array.copy()
                        floor_table[floor_perim_key][floor_index_adj] += wall_length

                    if type_area not in wall_types:
                        wall_types[type_area] = imaginary_array.copy()
                    wall_types[type_area][floor_index_adj] += wall_area
                    if perim not in wall_types:
                        wall_types[perim] = imaginary_array.copy()
                    wall_types[perim][floor_index_adj] += wall_length


            for wall in floor.findall('exploded/wall'):
                if wall.find('type').text == 'exterior':
                    points = wall.findall('point')
                    p1, p2, *rest = points
                    x1 = float(p1.get('x'))
                    x2 = float(p2.get('x'))
                    y1 = float(p1.get('y'))
                    y2 = float(p2.get('y'))
                    length = cart_distance((x1, y1), (x2, y2))
                    wall_height = (float(p1.get('height')) + float(p2.get('height')))/2
                    wall_area_gross += wall_height * length
                    extern_perim += length

            extern_perim -= extern_width_offset
            floors_perims.append(extern_perim)
            wall_height = wall_height/nwalls if nwalls != 0 else wall_height
            floors_heights.append(wall_height)

            wall_area_gross -= wall_types['Party Wall Area'][floor_index_adj] if 'Party Wall Area' in wall_types else 0
            wall_area_gross -= wall_types['Internal Wall Area'][floor_index_adj] if 'Internal Wall Area' in wall_types else 0

            walls_area_gross.append(wall_area_gross)


            # LOGGER.info('window_door_table len: ' + str(len(window_door_table)))
            # LOGGER.info('wall_types len: ' + str(len(wall_types)))
            
            for type in window_door_table:

                window_door_area = window_door_table[type][floor_index_adj]
                
                gross_area_key = f'Wall Type {type} Area Gross'
                net_area_key = f'Wall Type {type} Area Net'
                try:
                    area : float
                    if net_area_key not in wall_types:
                        # LOGGER.info('net_area_key not in wall_types: ' + str(wall_types))
                        print(wall_types[gross_area_key])
                        wall_types[net_area_key] = wall_types[gross_area_key].copy()
                    

                    area = wall_types[net_area_key][floor_index_adj]

                    # net_area = area - window_door_area
                    # LOGGER.info('net_area: ' + str(net_area))
                    # wall_types[net_area_key][floor_index_adj] = net_area
                except Exception as ex:
                    # exc_type, exc_obj, exc_tb = sys.exc_info()
                    # fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
                    # print(exc_type, fname, exc_tb.tb_lineno)
                    print('Could not find wall type in wall_type dict')


            imaginary_floor_enum.append(floor_name)
            wall_area_net = wall_area_gross - sum([window_door_table[key][floor_index_adj] for key in window_door_table])
            walls_area_net.append(wall_area_net)

        if floor_num != None and floor_num < MAX_REAL_FLOORS:
            for window in floor.findall('exploded/window'):
                window_area += float(window.get('height')) * float(window.get('width'))                  

        doors_area.append(door_area)
        windows_area.append(window_area)
        floor_enum.append(floor_name)

        if floor_name in ['Ground Floor', '1st Floor', '2nd Floor', '3rd Floor', 'Roof']:

            led_count.append(len(floor.findall(f'symbolInstance[@symbol=\'{lookup["LED/CFL"]}\']')))
            lf_count.append(len(floor.findall(f'symbolInstance[@symbol=\'{lookup["Linear Fluorescent"]}\']')))
            hl_count.append(len(floor.findall(f'symbolInstance[@symbol=\'{lookup["Halogen Lamp"]}\']')))
            inc_count.append(len(floor.findall(f'symbolInstance[@symbol=\'{lookup["Incandescent"]}\']')))
            hlv_count.append(len(floor.findall(f'symbolInstance[@symbol=\'{lookup["Halogen LV"]}\']')))

            flueless_count.append(len(floor.findall('symbolInstance/values/value[@key="qcustomfield.122c26d158"]')))
            default_flues = len(floor.findall('symbolInstance/values/value[@key="qcustomfield.f8a9c5deq0.5i3vasj3i78"]'))
            non_default = floor.findall('symbolInstance/values/value[@key="qcustomfield.3f240a7858"]')
            non_default_flues = len(
                list(
                    filter(
                        lambda x: True if x.text == 'Flue' else False,
                        non_default
                    )
                )
            )

            boiler_flues = floor.findall('symbolInstance/values/value[@key="qcustomfield.733f024958"]..')
            boiler_flue_count = len(
                list(
                    filter(
                        lambda x: True if x.find('value[@key="qcustomfield.733f024958"]') != None and x.find('value[@key="qcustomfield.733f024958"]').text == 'Open.Flue' \
                        and x.find('value[@key="qcustomfield.733f0249q0.6ouelp9umr8"]') != None and x.find('value[@key="qcustomfield.733f0249q0.6ouelp9umr8"]').text == '0' else False,
                        boiler_flues
                    )
                )
            )
            
            flue_count.append(non_default_flues + default_flues + boiler_flue_count)
            chimney_count.append(len(non_default) - non_default_flues)

            in_f_count.append(
                len(floor.findall(f'symbolInstance[@symbol=\'{lookup["NMV"]}\']')) + 
                len(floor.findall(f'symbolInstance[@symbol=\'{lookup["EMV"]}\']')) + 
                len(floor.findall(f'symbolInstance[@symbol=\'{lookup["DCH"]}\']')) + 
                len(floor.findall(f'symbolInstance[@symbol=\'{lookup["EMVB"]}\']')) + 
                len(floor.findall(f'symbolInstance[@symbol=\'{lookup["ECHB"]}\']'))
            )

            vents = floor.findall(f'symbolInstance[@symbol=\'{lookup["EPV"]}\']') + floor.findall(f'symbolInstance[@symbol=\'{lookup["NPV"]}\']')
            pnc_count.append(
                len(
                    list(
                        filter(lambda x: False if x.find('values/value[@key="qcustomfield.8d83fdcaq0.46r9ir0vvd"]') != None and \
                            x.find('values/value[@key="qcustomfield.8d83fdcaq0.46r9ir0vvd"]').text == '1' else True, 
                            vents
                        )
                    )
                )
            )

            discounted_vents = len(vents) - pnc_count[-1]
            disc_vent_count.append(discounted_vents)
            total_vent_count.append(len(vents))

            rad_count.append(len(floor.findall('symbolInstance[@symbol="co-afc6eed1-0e5c-4189-b955-4d98f616baa3"]') + 
                                 floor.findall('symbolInstance[@symbol="co-a2b10df6-429a-49b7-bfbf-8824a91c6e39"]')))
            rad_trv_count.append(len(floor.findall('symbolInstance[@symbol="co-a2b10df6-429a-49b7-bfbf-8824a91c6e39"]')))
            rs_count.append(len(floor.findall('symbolInstance[@symbol="co-8e288bb1-7947-41a0-9224-5d1d32bbacd4"]')))
            programmer_count.append(len(floor.findall('symbolInstance[@symbol="co-88d188fc-8cd9-413f-8dce-6a5d4a987047"]')))
            er_count.append(len(floor.findall('symbolInstance[@symbol="co-e49d64d3-e0f2-47c9-bfc3-dfd8ece4e61c"]')))
            esh_count.append(len(floor.findall('symbolInstance[@symbol="co-30b97448-fe04-4202-b701-2f54cd5ad4b0"]')))

            bath_count.append(len(floor.findall('symbolInstance[@symbol="co-064a7f28-56e6-4d08-bfa5-d9f0aae885a1"]') + 
                                  floor.findall('symbolInstance[@symbol="co-9fe51e91-80c4-4114-8ce8-3cdb3eaadb86"]') +
                                  floor.findall('symbolInstance[@symbol="co-bdc6fc6b-7ab1-4b00-b6f3-2aa346c91d14"]') + 
                                  floor.findall('symbolInstance[@symbol="co-7d191d92-4a25-4c60-b2f0-65c9921b386d"]')
                                ))

            ies_count.append(len(floor.findall('symbolInstance[@symbol="co-9fe51e91-80c4-4114-8ce8-3cdb3eaadb86"]') + 
                                 floor.findall('symbolInstance[@symbol="co-f6f1173a-8abe-4a31-9f1f-0eb2ff93e00f"]')
                                ))

            mixer_showers = floor.findall('symbolInstance[@symbol="co-bdc6fc6b-7ab1-4b00-b6f3-2aa346c91d14"]') + \
                            floor.findall('symbolInstance[@symbol="co-8b8a81b5-b070-4d65-ae52-3cd5262c0215"]')

            msv_count.append(len([shower for shower in mixer_showers if shower.find('values/value[@key="qcustomfield.22ba7c63q0.bja6s075v1o"]') != None and shower.find('values/value[@key="qcustomfield.22ba7c63q0.bja6s075v1o"]').text == 'Vented']))

            msvp_count.append(len(floor.findall('symbolInstance[@symbol="co-7d191d92-4a25-4c60-b2f0-65c9921b386d"]') + 
                                  floor.findall('symbolInstance[@symbol="co-acd8e516-6f7a-4397-a890-fde87994fb80"]')  
                                ))
            msu_count.append(len([shower for shower in mixer_showers if shower.find('values/value[@key="qcustomfield.22ba7c63q0.bja6s075v1o"]') != None and shower.find('values/value[@key="qcustomfield.22ba7c63q0.bja6s075v1o"]').text == 'Unvented']))

            object_floor_enum += [floor_name]


        if floor_name == 'Roof':
            for room in floor.findall('floorRoom'):
                roof_type = room.find('values/value[@key="qcustomfield.8fd606fcq2"]')
                if roof_type == None:
                    continue
                area_str = room.get('area')
                if area_str == None:
                    continue
                area = float(area_str)
                roof_key = roof_type.text.replace('.', ' ')
                if roof_key not in roof_table:
                    roof_table[roof_key] = [0]
                roof_table[roof_key][0] += area

        floor_index += 1

    floor_enum.append('Total')

    summary_values = {
        'Perimeter'   : [sum([wall_types[perim][0] for perim in wall_types if 'Perimeter' in perim and re.search('[0-9]', perim)])],
        'Living Area' : [living_room_area]
    }

    wall_types_less_perim = wall_types.copy()
    for key in wall_types:
        if 'Perimeter' in key:
            del wall_types_less_perim[key]

    shower_bath_table = {
        'Count of Baths'                            : bath_count,
        'Count of Electric Showers'                 : ies_count,
        'Count of Mixer Showers - Vented'           : msv_count,
        'Count of Mixer Showers - Vented + Pump'    : msvp_count,
        'Count of Mixer Showers - Unvented'         : msu_count
    }

    lighting_table = {
        'LED/CFL'                         : led_count,
        'Halogen Lamp'                    : hl_count,
        'Halogen LV'                      : hlv_count,
        'Linear Fluorescent'              : lf_count,
        'Incandescent'                    : inc_count,
    }

    ventilation_table = {
        'Intermittent Fan Count'          : in_f_count,
        'Passive Non-Closable'            : pnc_count,
        'Discounted Vents'                : disc_vent_count,
        'Total Vent Count'                : total_vent_count,
        'Flueless Combustion Room Heater' : flueless_count,
        'Flue'                            : flue_count,
        'Chimney'                         : chimney_count
    }

    space_heating_table = {
        'Count of Radiators'                : rad_count,
        'Count of Radiators With TRVs'      : rad_trv_count,
        'Percentage of Radiators With TRVs' : map(lambda a, b : (a / b) * 100 if b != 0 else 0, rad_trv_count, rad_count),
        'Count of Programmers'              : programmer_count,
        'Count of Room Stats'               : rs_count,
        'Count of Electric Radiators'       : er_count,
        'Count of Electric Storage Heaters' : esh_count
    }

    wd.loc[len(wd.index)] = [
        'Totals', 
        wd['Number of Openings'].sum(), 
        wd['Number of Openings Draught Stripped'].sum(), 
        'N/A', 
        'N/A', 
        'N/A', 
        wd['Area'].sum()
    ]

    dt.loc[len(dt.index)] = [
        'Totals',
        dt['Number of Openings'].sum(),
        dt['Number of Openings Draught Stripped'].sum(),
        dt['Glazed Area'].sum(),
        'N/A',
        'N/A',
        'N/A',
        dt['Area'].sum()
    ]

    object_floor_enum += ['Total']
    real_floor_enum += ['Total']
    imaginary_floor_enum += ['Total']

    styling = "border=\"1\""
    output = f"""\
        <h1>Summary Table</h1> \
        {create_table(summary_values, ['Name', 'Sum'], styling=styling, do_not_sum=['All'])} \
        <h1>Lighting Table</h1> \
        {create_table(lighting_table, object_floor_enum, styling=styling)} \
        <h1>Ventilation Table</h1> \
        {create_table(ventilation_table, object_floor_enum, styling=styling)} \
        <h1>Space Heating Table</h1> \
        {create_table(space_heating_table, object_floor_enum, styling=styling, do_not_sum=['Percentage of Radiators With TRVs'])} \
        <h1>Shower and Bath Table</h1>
        {create_table(shower_bath_table, object_floor_enum, styling=styling)}
        {"<h1>Colour Area Table</h1>" + create_table(colours, floor_enum, styling=styling, colour_table=True) if len(colours) > 0 else ""} \
        <h1>Wall Types</h1> \
        {create_table(wall_types_less_perim, imaginary_floor_enum, styling=styling)} \
        <h1>Window Table</h1> \
        {wd.to_html()} \
        <h1>Door Table</h1> \
        {dt.to_html()} \
        <h1>Floor Area Table</h1> \
        {create_table(floor_table, imaginary_floor_enum, styling=styling)} \
        <h1>Roof Table</h1> \
        {create_table(roof_table, ['Name', 'Sum'], styling=styling, do_not_sum=['All'])} \
        <h2>""" + xml + """</h2>
        </div>"""
    
    return output


def preBER(root):
    return
def inspection(root):
    return
def qa(root):
    return






def survey(root):
    try:
        sfi = [] # a list to hold the numbers of roof (also wall?) types that are suitable for insulation
        
        plan_name = root.get('name')
        json_val_dict = {'plan_name': plan_name} # specifically NOT JSON
        
        # Values from root also need to be accessed outside this function, but it is here that they can be inserted into the HTML Output - in future consider separating out the two activities by returning the output_dict?
        
        # date = root.find('values/value[@key="date"]').text
        # assessor = root.find('values/value[@key="qf.34d66ce4q1"]').text
        # rating_type = root.find('values/value[@key="qf.34d66ce4q3"]').text
        # rating_purpose = root.find('values/value[@key="qf.34d66ce4q4"]').text
        
        date = root.find('values/value[@key="date"]').text
        json_val_dict['Survey Date *'] = date
        
        

        xml_ref_dict = XML_2_dict(root)
        
        
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
                , 'Can Roof Type 2 Insulation Thickness be Measured?*'
                , 'Roof 2 Thickness (mm)*'
                , 'Roof 2 Insulation Type*'
                , 'Roof 2 Required per standards (mm2) *'
                , 'Roof 2 Existing (mm2) *'
                , 'Area of Roof Type 2 with fixed flooring (m2)*'
                , 'Folding/stair ladder in Roof Type 2*'
                , 'Fixed light in Roof Type 2*'
                , 'Downlighters in Roof Type 2*'
                , 'High power cable in Roof Type 2 (6sq/10sq or higher)*'
                , 'Roof 3 Type'
                , 'Other Details Roof 3*'
                , 'Sloped Ceiling Roof 3*'
                , 'Roof 3 greater than 2/3 floor area*'
                , 'Roof 3 Pitch (degrees)*'
                , 'Roof 3 Insulation Exists*'
                , 'Can Roof Type 3 Insulation Thickness be Measured?'
                , 'Roof 3 Thickness (mm)'
                , 'Roof 3 Insulation Type'
                , 'Roof 3 Required per standards (mm2) *'
                , 'Roof 3 Existing (mm2) *'
                , 'Area of Roof Type 3 with fixed flooring (m2)'
                , 'Folding/stair ladder in Roof Type 3'
                , 'Fixed light in Roof Type 3'
                , 'Downlighters in Roof Type 3'
                , 'High power cable in Roof Type 3 (6sq/10sq or higher)'
                , 'Roof 4 Type'
                , 'Other Details Roof 4*'
                , 'Sloped Ceiling Roof 4*'
                , 'Roof 4 greater than 2/3 floor area*'
                , 'Roof 4 Pitch (degrees)*'
                , 'Roof 4 Insulation Exists*'
                , 'Can Roof Type 4 Insulation Thickness be Measured?*'
                , 'Roof 4 Thickness (mm)*'
                , 'Roof 4 Insulation Type*'
                , 'Roof 4 Required per standards (mm2) *'
                , 'Roof 4 Existing (mm2) *'
                , 'Area of Roof Type 4 with fixed flooring (m2)*'
                , 'Folding/stair ladder in Roof Type 4*'
                , 'Fixed light in Roof Type 4*'
                , 'Downlighters in Roof Type 4*'
                , 'High power cable in Roof Type 4 (6sq/10sq or higher)*'
                , 'Suitable for Insulation *'
                , 'Not suitable details*'
                , 'Notes (Roof)']
        ofl_walls = ['Wall Type 1*'
                    , 'Wall 1 wall thickness (mm)*'
                    , 'Wall 1 Insulation Present?*'
                    , 'Wall 1 Insulation Type*'
                    , "Wall 1 Fill Type*"
                    , 'Wall 1 Residual Cavity Width (mm)*'
                    , 'Can Wall type 1 Insulation Thickness be Measured?*'
                    , "If 'Yes' enter insulation thickness (mm)*"
                    , 'Wall Type 2'
                    , 'Wall 2 wall thickness (mm)*'
                    , 'Wall 2 Insulation Present?*'
                    , 'Wall 2 Insulation Type*'
                    , "Wall 2 Fill Type*"
                    , 'Wall 2 Residual Cavity Width (mm)*'
                    , 'Can Wall type 2 Insulation Thickness be Measured?*'
                    , "If 'Yes' enter Wall type 2 insulation thickness (mm)*"
                    , 'Wall Type 3'
                    , 'Wall 3 wall thickness (mm)*'
                    , 'Wall 3 Insulation Present?*'
                    , 'Wall 3 Insulation Type*'
                    , "Wall 3 Fill Type*"
                    , 'Wall 3 Residual Cavity Width (mm)*'
                    , 'Can Wall type 3 Insulation Thickness be Measured?*'
                    , "If 'Yes' enter Wall type 3 insulation thickness (mm)*"
                    , 'Wall Type 4'
                    , 'Wall 4 wall thickness (mm)*'
                    , 'Wall 4 Insulation Present?*'
                    , 'Wall 4 Insulation Type*'
                    , "Wall 4 Fill Type*"
                    , 'Wall 4 Residual Cavity Width (mm)*'
                    , 'Can Wall type 4 Insulation Thickness be Measured?*'
                    , "If 'Yes' enter Wall type 4 insulation thickness (mm)*"
                    , "Is the property suitable for wall insulation? *"
                    , "No wall insulation details *"
                    , "Notes (Walls)"
                    , 'Suitable for Draught Proofing'
                    , 'Not suitable details Draughtproofing*'
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
                    , "GNI new connection"
                    , "RGI Meter_No Heating"
                    , 'Internal Wall Insulation: Vertical Surface'
                    , "External wall insulation and CWI: less than 60m2"
                    , "External wall insulation and CWI: 60m2 to 85m2"
                    , "External wall insulation and CWI: greater than 85m2"
                    , 'replace_window_area'
                    , 'Notes (Windows and Doors)'
                    ]
        

        
        
        id = root.get('id')
        print(id)
        
        headers = {
            "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/39.0.2171.95 Safari/537.36"
            ,"key": "45170e50321733db78952dfa5901b0dfeeb8"
            , "customer": "63b5a4ae69c91"
            , "accept": "application/json"
            }
        
        
        json_url = "https://cloud.magicplan.app/api/v2/plans/forms/" + str(id)
        request = urllib.request.Request(json_url, headers=headers)
        JSON = urllib.request.urlopen(request).read()
        JSON = json.loads(JSON)
        print('len(JSON["data"]): ', len(JSON["data"]))
        df = pd.DataFrame(JSON["data"])
        
        
            
        # json_val_dict['new_hatch_count'] = 0
        json_val_dict["Existing (mm2)*"] = int(0)
        json_val_dict['No. Single Glazed Windows *'] = 0
        # json_ref_dict = {}
        alt_double_glazed_count = 0
        

        
        
        
        replace_windows = []
        
        # print(df)
        print('len(df.forms): ', len(df.forms))
        for form in df.forms:
            # print(form["symbol_instance_id"])
            df2 = pd.DataFrame(form)
            for section in df2.sections:
                df3 = pd.DataFrame(section)
                for field in df3.fields:
                    df4 = pd.DataFrame(field)
                    for index, row in df4.iterrows():
                        
                        if row["label"] == "Is the window Single glazed?":
                            if row["value"]["value"] == True:
                                json_val_dict['No. Single Glazed Windows *'] += 1
                            if row["value"]["value"] == False:
                                alt_double_glazed_count += 1
                        
                        if row["label"] == "Existing Roof Ventilation (mm2)*":
                            if not row["value"]["value"].isdigit():
                                continue
                            json_val_dict["Existing (mm2)*"] += int(row["value"]["value"])
                        
                        v = ''
                        if row["value"]["value"] == None:
                            vals = [val["value"] for val in row["value"]["values"]]
                            for val in vals:
                                v += val
                                v += '<BR>'
                        else:
                            v = row["value"]["value"]
                        json_val_dict[row["label"]] = v


        
        
        json_val_dict["ESB alteration"] = 0
        esb_alterations = []
        json_val_dict["GNI meter alteration"] = 0
        
        json_val_dict["Duct Cooker Hood"] = 0
        
        slope_dict = {}
        roof_type_dict = {}
        for datum in JSON["data"]:
            # print(datum["symbol_name"])
            if datum["symbol_name"] == "ESB alteration":
                json_val_dict["ESB alteration"] += 1
                esb_alterations += datum["symbol_instance_id"]
            if datum["symbol_name"] == "GNI meter alteration":
                json_val_dict["GNI meter alteration"] += 1
            # print(datum["symbol_instance_id"])
            if datum["symbol_name"] == "GNI new connection":
                json_val_dict["GNI new connection"] += 1
            if datum["symbol_name"] == "RGI Meter_No Heating":
                json_val_dict["RGI Meter_No Heating"] += 1
            for form in datum["forms"]:
                for section in form["sections"]:
                    for field in section["fields"]:
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
                        g = {} # temporary dictionary containing the answers to questions
                        for field in section["fields"]:
                            g[field["label"]] = field["value"]["value"]
                    for n in range(1, 5):
                        if "Roof Type*" in g.keys() and f"Roof Type {n} Sloping Ceiling Suitable for Insulation*" in g.keys():
                            # print('g: ', g)
                            if g["Roof Type*"] == "Sloped Ceiling" and g[f"Roof Type {n} Sloping Ceiling Suitable for Insulation*"] == True:
                                # print("Need slope for: ", g["Part of roof type?*"])
                                # slopes.append(g["Part of roof type?*"])
                                if f"Roof {n} Pitch (degrees)*" in json_val_dict.keys():
                                    pitch = json_val_dict[f"Roof {n} Pitch (degrees)*"]
                                else:
                                    pitch = 30
                                slope_dict[datum["symbol_instance_id"]] = pitch
        # print('slopes', ': ', slopes)
        # print('roof_type_dict', ': ', roof_type_dict)
        # print('slope_dict', ': ', slope_dict)

        # for x in json_val_dict:
            # print(x, json_val_dict[x])
        json_val_dict["ESB alteration"] = json_val_dict["ESB alteration"] if json_val_dict["ESB alteration"] != 0 else 'N/A'
        json_val_dict["GNI meter alteration"] = json_val_dict["GNI meter alteration"] if json_val_dict["GNI meter alteration"] != 0 else 'N/A'

        
        
        json_url = "https://cloud.magicplan.app/api/v2/plans/statistics/" + str(id)
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





        json_val_dict['replace_window_area'] = 0
        # print(xml_ref_dict)
        print(replace_windows)
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
                json_val_dict['Number of Storeys *'] += 1
                json_val_dict['Gross floor area (m2) *'] += floor["area_with_interior_walls_only"]
                json_val_dict['No. Double Glazed Windows *'] += floor["window_count"]

                for room in floor["rooms"]:
                    for wall_item in room["wall_items"]:
                        if wall_item["uid"] in replace_windows:
                            json_val_dict['replace_window_area'] += (wall_item["width"] * wall_item["height"])
                    for furniture in room["furnitures"]:
                        print(furniture["name"])
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
                        if furniture["name"] == "New Hatch": # only found in -1 to 9 and Roof?
                            new_hatch_count += 1
                            
            for room in floor["rooms"]:
                for furniture in room["furnitures"]:
                    if furniture["name"] == "New Low Level Roof Ventilation": # only found in Roof?
                        sum_low += float(furniture["width"])
                    if furniture["name"] == "New High Level Roof Ventilation": # only found in Roof?
                        sum_high += float(furniture["width"])
        
        json_val_dict["Cavity Wall Insulation Bonded Bead"] = round(json_val_dict["Cavity Wall Insulation Bonded Bead"]) if json_val_dict["Cavity Wall Insulation Bonded Bead"] != 0 else 'N/A'
        json_val_dict["Loose Fibre Extraction"] = round(json_val_dict["Loose Fibre Extraction"]) if json_val_dict["Loose Fibre Extraction"] != 0 else 'N/A'
        json_val_dict["Internal Wall Insulation: Vertical Surface"] = round(json_val_dict["Internal Wall Insulation: Vertical Surface"]) if json_val_dict["Internal Wall Insulation: Vertical Surface"] != 0 else 'N/A'
        json_val_dict['replace_window_area'] = round(json_val_dict['replace_window_area']) if json_val_dict['replace_window_area'] != 0 else 'N/A'
        json_val_dict['replace_window_area'] = 1 if json_val_dict['replace_window_area'] == 0 else json_val_dict['replace_window_area']
        json_val_dict['Notes (Windows and Doors)'] = json_val_dict['Notes (Windows and Doors)'] if json_val_dict['Notes (Windows and Doors)'] != '' else 'N/A'
        json_val_dict['No. Double Glazed Windows *'] = json_val_dict['No. Double Glazed Windows *'] - json_val_dict['No. Single Glazed Windows *']
        
        
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
                    # print('room["uid"]', ': ', room["uid"])
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
        
        
        # Work Order Recommendation (Roof):
        json_val_dict['sloped_surface_area'] = round(slope_roof_area_sum) if round(slope_roof_area_sum) != 0 else 'N/A'
        
        print('sfi_dict', ': ', json_val_dict["sfi_dict"])
        json_val_dict['storage'] = 0
        for t in [100, 150, 200, 250, 300]:
            if str(t) in json_val_dict["sfi_dict"].keys():
                json_val_dict[f'ins_{t}_area'] = round(json_val_dict["sfi_dict"][str(t)])
                json_val_dict['storage'] = 1
        
        
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
        
        

        # print(json_val_dict)
        output_dict = json_val_dict


        styling = "border=\"1\""
        output = f"""\
            <h1>General</h1> \
            {create_table_text(output_dict, headers = ['name', 'value'], styling=styling, do_not_sum=['All'], order_list = ofl_general)} \
            <h1>Roof</h1> \
            {create_table_text(output_dict, headers = ['name', 'value'], styling=styling, do_not_sum=['All'], order_list = ofl_roof)} \
            <h1>Walls</h1> \
            {create_table_text(output_dict, headers = ['name', 'value'], styling=styling, do_not_sum=['All'], order_list = ofl_walls)} \
            </div>"""

    except Exception as ex:
        # exc_type, exc_obj, exc_tb = sys.exc_info()
        # output = "Line " + str(exc_tb.tb_lineno) + ": " + exc_type 
        
        output = str(ex)
        output = traceback.format_exc()
        # LOGGER.info('Exception : ' + str(traceback.format_exc()))
        
        # fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        # print(exc_type, fname, exc_tb.tb_lineno)
    finally:
        return output
    return output


def XML_old():
    # Go through the XML, referring to the JSON data whenever we need to - this is now disused but might need it again if there are any required values not included in the JSON (e.g. counts of objects)
    
    # values = root.findall('values/value')
    # for value in values:
        # k = value.attrib["key"]
        # if value.attrib["key"] in json_ref_dict.keys():
            # k = json_ref_dict[value.attrib["key"]]
        # if 'statistics.' in k:
            # continue
        # output_dict[k] = value.text
    
    
    # values = root.findall('floor/floorRoom/values/value')
    # for value in values:
        # k = value.attrib["key"]
        # if value.attrib["key"] in json_ref_dict.keys():
            # k = json_ref_dict[value.attrib["key"]]
        # output_dict[k] = value.text
    
    
            
    floors = root.findall('floor')
    # LOGGER.info('no of floors:' + str(len(floors)))
    
    # Calculated Field. Equals SUM of "Ground surface without walls: m²" for floors Basement level 1, Ground Floor, higher ground floor, 1st floor, 2nd floor, 3rd floor……...up to 9th floor
    floor_area = 0
    floor_area_without_walls = 0
    floor_area_with_walls = 0
    # for floor in root.findall('floor[@floorType="10"]'):
    for floor in floors:
        if int(floor.get('floorType')) > 9:
            continue
        print('floorType: ' + floor.get('floorType'))
        floor_area_without_walls += float(floor.get('areaWithoutWalls')) if floor.get('areaWithoutWalls') != None else 0
        floor_area_with_walls += float(floor.get('areaWithInteriorWallsOnly')) if floor.get('areaWithInteriorWallsOnly') != None else 0
    output_dict['floor_area_without_walls'] = floor_area_without_walls
    output_dict['floor_area_with_walls'] = floor_area_with_walls
    
    # Count of floors Basement level 1, Ground Floor, higher ground floor, 1st floor, 2nd floor, 3rd floor……...up to 9th floor
    output_dict['no_of_floors'] = len(floors)




def distributor_function(form):
    # At this point we can hopefully use info from the form to identify what type it is
    # The extracted XML "root" is then sent to the appropriate function which returns a HTML formatted table as output
    # which is then included in the JSON 
    
    plan_name = form['title']
    email = form['email']
    xml = form['xml']
    root : ET.Element
    with urllib.request.urlopen(xml) as f:
        s = f.read().decode('utf-8')
    root = dET.fromstring(s)

    # if "Survey" in plan_name:
        # output = survey(root)
    # elif "BER" in plan_name:
        # output = ber_old(root)
    # elif "Pre BER" in plan_name:
        # output = preBER(root)
    # elif "Inspection" in plan_name:
        # output = inspection(root)
    # elif "QA" in plan_name:
        # output = qa(root)
    
    output = survey(root)
    # output = ber_old(root)

    output = output + '<h2>' + xml + '</h2></div>'

    json_data = json.dumps({
        'email' : email,
        'name'  : plan_name, 
        'table' : output
    })



    return json_data




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
        sc = 503    # Service Unavailable
        

    finally:        
        try:
            account_url = os.environ['AZ_STR_URL']
            default_credential = DefaultAzureCredential()
            blob_service_client = BlobServiceClient(account_url, credential=default_credential)
            container_name = os.environ['AZ_CNTR_ST']
            container_client = blob_service_client.get_container_client(container_name)
            if not container_client.exists():
                container_client = blob_service_client.create_container(container_name)
            local_file_name = str(uuid.uuid4()) + '.json'
            blob_client = blob_service_client.get_blob_client(container=container_name, blob=local_file_name)
            blob_client.upload_blob(json_data)
        except Exception as ex:
            output = str(ex)
            sc = 500     # Internal Server Error
        return func.HttpResponse(status_code=sc)

    
