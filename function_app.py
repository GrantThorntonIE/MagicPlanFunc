import ToAzure
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
from loguru import logger as LOGGER

# import win32com.client
import matplotlib.pyplot as plt
import numpy as np
from shapely.geometry import Point, MultiPoint




def survey_new(root, json_data):
    try:
        plan_name = root.get('name')
        output = {'plan_name': plan_name}
        floors = root.findall('floor')


        floor_area = 0 # Calculated Field. Equals SUM of "Ground surface without walls: m²" for floors Basement level 1, Ground Floor, higher ground floor, 1st floor, 2nd floor, 3rd floor……...up to 9th floor
        floor_area_without_walls = 0
        floor_area_with_walls = 0
        # for floor in root.findall('floor[@floorType="10"]'):
        for floor in floors:
            if int(floor.get('floorType')) > 9:
                continue
            print('floorType: ' + floor.get('floorType'))
            floor_area_without_walls += float(floor.get('areaWithoutWalls')) if floor.get('areaWithoutWalls') != None else 0
            floor_area_with_walls += float(floor.get('areaWithInteriorWallsOnly')) if floor.get('areaWithInteriorWallsOnly') != None else 0
        output['floor_area_without_walls'] = floor_area_without_walls
        output['floor_area_with_walls'] = floor_area_with_walls
        
        
        
        # Count of floors Basement level 1, Ground Floor, higher ground floor, 1st floor, 2nd floor, 3rd floor……...up to 9th floor
        output['no_of_floors'] = len(floors)
        
        
        
        values = root.findall('values/value')
        print(len(values))
        for value in values:
            print(value.attrib["key"])
            
        
        
        # items = json_data['data']['forms']
        # for item in items:
            # if 'label' in item.keys():
                # print 'Label Found'

        # photos?
        
        
        styling = "border=\"1\""
        output = create_table_text(output, headers = ['name', 'value'], styling=styling, do_not_sum=['All'])
        
        # print(output)


        
    except Exception as ex:
        
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        print(exc_type, fname, exc_tb.tb_lineno)
        # print('Could not find wall type in wall_type dict')
        LOGGER.info('Exception : ' + str(ex))
    finally:
        return output
    return output


def cart_distance(p1 : tuple[float, float], p2 : tuple[float, float]) -> float:
    (x1, y1) = p1
    (x2, y2) = p2
    return sqrt((x2 - x1) ** 2 + (y2 - y1) ** 2)


    
    
def review():

    
    print('exterior_walls', ':')
    for ew in exterior_walls:
        print(ew)
    
    
    

    ew_ordered = []
    ew_ordered.append(exterior_walls[0])
    print('ew_ordered', ':', ew_ordered)
    
    
    ew = exterior_walls[0]
    no_of_walls = len(exterior_walls)
    for x in range(1, no_of_walls):
        print('ew', ':', ew)
        p1 = ew[0]
        print('p1', ':', p1)
        p2 = ew[1]
        print('p2', ':', p2)
        for w in exterior_walls:
            if p1 in w and w != ew and w not in ew_ordered:
                print('w', ':', w)
                print('found p1 in w, adding w to ew_ordered')
                ew_ordered.append(w)
                print('ew_ordered', ':', ew_ordered)
                ew = w
            if p2 in w and w != ew and w not in ew_ordered:
                print('w', ':', w)
                print('found p2 in w, adding w to ew_ordered')
                ew_ordered.append(w)
                print('ew_ordered', ':', ew_ordered)
                ew = w
        
    print('ew_ordered', ':')
    for e in ew_ordered:
        print(e)
    # return exterior_walls

    rooms = floor.findall('floorRoom')
    print('len(rooms)', ':', len(rooms))
    for room in rooms:
        xpoints = np.array([])
        ypoints = np.array([])
        rx = round(float(room.get('x')), 2)
        print('rx', ':', rx)
        ry = round(float(room.get('y')), 2)
        points = room.findall('point')
        print('len(points)', ':', len(points))
        for point in points:
            x = round(float(point.get('snappedX')), 2)
            print('x', ':', x)
            x = rx + x
            print('x', ':', x)
            y = -ry - round(float(point.get('snappedY')), 2)
            h = point.get('height')
            uid = point.get('uid')
            
            xpoints = np.append(xpoints, [x])
            ypoints = np.append(ypoints, [y])
        xpoints = np.append(xpoints, xpoints[0])
        ypoints = np.append(ypoints, ypoints[0])
        
        # print(xpoints)
        # print(ypoints)
        # plt.plot(xpoints, ypoints)
        
        # for i, x in enumerate(xpoints):
            # all_coordinates.append([x, ypoints[i-1]])
        # print('all_coordinates', ':', all_coordinates)
        
    # print('len(floors)', ':', len(floors))
    # for floor in floors:
        # all_coordinates = []
        # rooms = floor.findall('floorRoom')
        # print('len(rooms)', ':', len(rooms))
        # for room in rooms:
            # xpoints = np.array([])
            # ypoints = np.array([])
            # rx = float(room.get('x'))
            # print('rx', ':', rx)
            # ry = float(room.get('y'))
            # points = room.findall('point')
            # print('len(points)', ':', len(points))
            # for point in points:
                # x = float(point.get('snappedX'))
                # print('x', ':', x)
                # x = rx + x
                # print('x', ':', x)
                # y = -ry - float(point.get('snappedY'))
                # h = point.get('height')
                # uid = point.get('uid')
                
                # xpoints = np.append(xpoints, [x])
                # ypoints = np.append(ypoints, [y])
            # xpoints = np.append(xpoints, xpoints[0])
            # ypoints = np.append(ypoints, ypoints[0])
            
            # print(xpoints)
            # print(ypoints)
            # plt.plot(xpoints, ypoints)
            
            # for i, x in enumerate(xpoints):
                # all_coordinates.append([x, ypoints[i-1]])
            # print('all_coordinates', ':', all_coordinates)
        
        
        
        # point_coordinates = [Point(z[0], z[1]) for z in all_coordinates]
        # print(point_coordinates)
        # multi_point = MultiPoint(point_coordinates)
        # convex_hull = multi_point.convex_hull
        # print('convex_hull.exterior.coords.xy', ':', convex_hull.exterior.coords.xy)
        # mark = [vals.index(i) for i in convex_hull[0]]
        # print(mark)
        
        # bounding_box = multi_point.bounds
        # print('bounding_box', ':', bounding_box) 
        
        
        
        
        
        
        # plt.axis('equal')
        # plt.show()

    # print(xpoints)
    # print(ypoints)
    # plt.plot(xpoints, ypoints)
    # plt.show()
    return





def read_XML_from_filepath(xmlfilepath):
    try:
        output = ''
        email = 'RPASupport@ie.gt.com' # req._HttpRequest__params['email']


        root : ET.Element
        with open(xmlfilepath) as f:
            xml_data_as_string = f.read()
        root = dET.fromstring(xml_data_as_string)
    
    except:
        root = 'problem getting root'
    
    finally:
        return root


def get_JSON(root, req_JSON):
    try:
        plan_name = root.get('name')
        id = root.get('id')
        print(id)
        
        headers = {
            "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/39.0.2171.95 Safari/537.36"
            ,"key": "45170e50321733db78952dfa5901b0dfeeb8"
            , "customer": "63b5a4ae69c91"
            , "accept": "application/json"
            }
        
        
        
        if req_JSON == 'files':
            json_url = "https://cloud.magicplan.app/api/v2/plans/" + str(id) + "/files?include_photos=true"
        else:
            json_url = "https://cloud.magicplan.app/api/v2/plans/" + req_JSON + "/" + str(id)
        
        # json_url = "https://cloud.magicplan.app/api/v2/plans/" + req_JSON + "/" + str(id)
        request = urllib.request.Request(json_url, headers=headers)
        JSON = urllib.request.urlopen(request).read()
        JSON = json.loads(JSON)
        
        # df = pd.DataFrame(JSON["data"])
        
        file = plan_name + '_' + req_JSON + '.json'
        with open(file, 'w') as outfile:
            json.dump(JSON, outfile, indent=4)
    
        print(file)
        
    except Exception as ex:
        output = str(ex)
        output = traceback.format_exc()
        # LOGGER.info('Exception : ' + str(traceback.format_exc()))
        print(output)
    
    finally:
        return JSON












# projects = ['Full Survey Test 1', 'Full Survey Test 2', 'Full Survey Test 3', 'Full Survey Test 4', 'Full Survey Test 5', 'Full Survey Test 6', 'Full Survey Test 7', 'Full Survey Test 8', 'Full Survey Test 9', 'Full Survey Test 10']
# projects = ['Major Renovation Survey Portal 4']
# projects = ['Doors and Windows Symbol Instance']
# projects = ['Full Survey Test 3 1']
# projects = ['WH5735507 Major Renovation Survey Portal 6']
# projects = ['General And Roof Test 1']
# projects = ['WH5735507 Major Renovation Survey Portal 6 1']
# projects = ['Full Survey Test 6 1']
# projects = ['WH568343 QA']
# projects = ['MM Home 2']
# projects = ['Full Survey Test 10']
# projects = ['WH573384 QA']
# projects = ['WH568631 Mr Issue']
# projects = ['Wh568184 QA']
# projects = ['WH567983 QA']
# projects = ['WH566181 Survey']
# projects = ['WH563791 QA']
# projects = ['WH567672 - Survey']
# projects = ['WH565958 Survey']
# projects = ['WH563782 QA']
projects = ['Wh573577 QA'] # Heating Notes
projects = ['Heating Appliances'] # Heating Notes
# projects = ['WH563782 QA 2']
# projects = ['WH563944 QA']
# projects = ['WH573032 Survey']
projects = ['WH564273 QA']
projects = ['WH565598 QA']


for project in projects:
    xmlfilepath = 'd:\\Users\\gshortall\\Documents\\KSN\\MagicPlan\\' + project + '.magicplan'
    # LOGGER.info('Processing file: ' + xmlfilepath)
    
    root = read_XML_from_filepath(xmlfilepath)
    
    # d = ToAzure.XML_2_dict(root)
    # print(d)
    
    # JSON_forms = get_JSON(root, 'forms')
    # JSON_statistics = get_JSON(root, 'statistics')
    # JSON_files = get_JSON(root, 'files')
    
    html_output = ''
    # html_output = exterior_walls(root)
    html_output = ToAzure.survey(root)
    
    if html_output != '' and html_output != None and html_output != 0:
        # print('html_output', ':', html_output)
        outputfilepath = 'd:\\Users\\gshortall\\Documents\\KSN\\' + project + '.htm'
        # outputfilepath = 'd:\\Users\\gshortall\\Documents\\KSN\\' + project + '.json'
        f = open(outputfilepath, "w")
        f.write(html_output)
        f.close()


quit()



def ber_old(root):
    

    lookup = {
        'LED/CFL'           : 'co-3a9c9ff6-2bad-4d62-9526-1df98538cbad',
        'Halogen Lamp'      : 'co-94486aec-b47a-4d75-aaf3-0645576bae56',
        'Halogen LV'        : 'co-b21a94da-ad62-40e5-bfe0-c1aa0b8461d5',
        'Linear Fluorescent':'co-497bde35-eb4a-41ec-ba91-b24e35099799',
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
    if plan_name[-1] == ' ':
        plan_name = plan_name[:-1]
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
                        # print(wall_types[gross_area_key])
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

