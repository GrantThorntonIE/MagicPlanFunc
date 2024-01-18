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
        

        output = {}
        output['_HttpRequest__body_bytes'] = hasattr(req._HttpRequest__body_bytes, '__getitem__')
        output['_HttpRequest__body_str'] = hasattr(req._HttpRequest__body_str, '__getitem__')
        output['_HttpRequest__body_type'] = hasattr(req._HttpRequest__body_type, '__getitem__')
        output['_HttpRequest__files'] = hasattr(req._HttpRequest__files, '__getitem__')
        output['_HttpRequest__form'] = hasattr(req._HttpRequest__form, '__getitem__')
        output['_HttpRequest__form_parsed'] = hasattr(req._HttpRequest__form_parsed, '__getitem__')
        output['_HttpRequest__headers'] = hasattr(req._HttpRequest__headers, '__getitem__')
        output['_HttpRequest__method'] = hasattr(req._HttpRequest__method, '__getitem__')
        output['_HttpRequest__params'] = hasattr(req._HttpRequest__params, '__getitem__')
        output['_HttpRequest__route_params'] = hasattr(req._HttpRequest__route_params, '__getitem__')
        output['_HttpRequest__url'] = hasattr(req._HttpRequest__url, '__getitem__')
        output['__abstractmethods__'] = hasattr(req.__abstractmethods__, '__getitem__')
        output['__annotations__'] = hasattr(req.__annotations__, '__getitem__')
        output['__class__'] = hasattr(req.__class__, '__getitem__')
        output['__delattr__'] = hasattr(req.__delattr__, '__getitem__')
        output['__dict__'] = hasattr(req.__dict__, '__getitem__')
        output['__dir__'] = hasattr(req.__dir__, '__getitem__')
        output['__doc__'] = hasattr(req.__doc__, '__getitem__')
        output['__eq__'] = hasattr(req.__eq__, '__getitem__')
        output['__format__'] = hasattr(req.__format__, '__getitem__')
        output['__ge__'] = hasattr(req.__ge__, '__getitem__')
        output['__getattribute__'] = hasattr(req.__getattribute__, '__getitem__')
        output['__gt__'] = hasattr(req.__gt__, '__getitem__')
        output['__hash__'] = hasattr(req.__hash__, '__getitem__')
        output['__init__'] = hasattr(req.__init__, '__getitem__')
        output['__init_subclass__'] = hasattr(req.__init_subclass__, '__getitem__')
        output['__le__'] = hasattr(req.__le__, '__getitem__')
        output['__lt__'] = hasattr(req.__lt__, '__getitem__')
        output['__module__'] = hasattr(req.__module__, '__getitem__')
        output['__ne__'] = hasattr(req.__ne__, '__getitem__')
        output['__new__'] = hasattr(req.__new__, '__getitem__')
        output['__reduce__'] = hasattr(req.__reduce__, '__getitem__')
        output['__reduce_ex__'] = hasattr(req.__reduce_ex__, '__getitem__')
        output['__repr__'] = hasattr(req.__repr__, '__getitem__')
        output['__setattr__'] = hasattr(req.__setattr__, '__getitem__')
        output['__sizeof__'] = hasattr(req.__sizeof__, '__getitem__')
        output['__slots__'] = hasattr(req.__slots__, '__getitem__')
        output['__str__'] = hasattr(req.__str__, '__getitem__')
        output['__subclasshook__'] = hasattr(req.__subclasshook__, '__getitem__')
        output['__weakref__'] = hasattr(req.__weakref__, '__getitem__')
        output['_abc_impl'] = hasattr(req._abc_impl, '__getitem__')
        output['_parse_form_data'] = hasattr(req._parse_form_data, '__getitem__')
        output['files'] = hasattr(req.files, '__getitem__')
        output['form'] = hasattr(req.form, '__getitem__')
        output['get_body'] = hasattr(req.get_body, '__getitem__')
        output['get_json'] = hasattr(req.get_json, '__getitem__')
        output['headers'] = hasattr(req.headers, '__getitem__')
        output['method'] = hasattr(req.method, '__getitem__')
        output['params'] = hasattr(req.params, '__getitem__')
        output['route_params'] = hasattr(req.route_params, '__getitem__')
        output['url'] = hasattr(req.url, '__getitem__')



        
        
        
        # output = req.get_body().decode('utf-8')
        
        

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
    
            json_data = json.dumps({
                'email' : email,
                'name'  : plan_name, 
                'table' : str(output)
            })
    
            local_file_name = str(uuid.uuid4()) + '.json'
    
            blob_client = blob_service_client.get_blob_client(container=container_name, blob=local_file_name)
    
            blob_client.upload_blob(json_data)
        except:
            sc = 500     # Internal Server Error
        
    
        return func.HttpResponse(status_code=sc)

    
