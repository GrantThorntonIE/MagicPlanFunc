import json
from math import sqrt
import re
import azure.functions as func
import os, logging, uuid
from azure.storage.blob import BlobServiceClient
from azure.identity import DefaultAzureCredential
import urllib.request
import pandas as pd
import xml.etree.ElementTree as ET
import defusedxml.ElementTree as dET


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

app = func.FunctionApp()
@app.function_name(name="MagicplanTrigger")
@app.route(route="magicplan", auth_level=func.AuthLevel.ANONYMOUS)


def test_function(req: func.HttpRequest) -> func.HttpResponse:
    try:

        # email = req._HttpRequest__params['email']
        # xml = req._HttpRequest__params['xml']

        plan_name = 'test'
        email = 'RPASupport@ie.gt.com'
        # xml = req._HttpRequest__body_str
        xml = req.get_json()
        


        account_url = os.environ['AZ_STR_URL']
        default_credential = DefaultAzureCredential()
        blob_service_client = BlobServiceClient(account_url, credential=default_credential)
        container_name = os.environ['AZ_CNTR_ST']
        container_client = blob_service_client.get_container_client(container_name)
        if not container_client.exists():
            container_client = blob_service_client.create_container(container_name)

        output = xml
        
        json_data = json.dumps({
            'email' : email,
            'name'  : plan_name, 
            'table' : output
        })

        local_file_name = str(uuid.uuid4()) + '.json'

        blob_client = blob_service_client.get_blob_client(container=container_name, blob=local_file_name)

        blob_client.upload_blob(json_data)
        return func.HttpResponse(status_code=200)

    except Exception as ex:
        logging.error(ex)
        print(f"Exception: {ex}")
    return func.HttpResponse(status_code=200)
