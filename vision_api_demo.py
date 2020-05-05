import os,io
from google.cloud import vision
from google.cloud.vision import types
import pandas as pd
import json
os.environ['GOOGLE_APPLICATION_CREDENTIALS']= r'vision_key.json'
from google.protobuf.json_format import MessageToJson,MessageToDict
from protobuf_to_dict import protobuf_to_dict, dict_to_protobuf
import testing

def detect_text(path,filename):
    
    filename = filename[:-4]
    
    """Detects text in the file."""
    client = vision.ImageAnnotatorClient()

    with io.open(path, 'rb') as image_file:
        content = image_file.read()

    image = vision.types.Image(content=content)

    response = client.text_detection(image=image)
    # json_texts =protobuf_to_dict(response.text_annotations)
    json_texts = MessageToDict(response, preserving_proto_field_name = True)    
    with open('./jsons/{}.json'.format(filename), 'w', encoding='utf-8') as f:
        json.dump(json_texts, f, indent=4)
    
    testing.save_to_sheet()
    # lines = testing.testing_driver('./jsons/10.json')   
    if response.error.message:
        raise Exception(
            '{}\nFor more info on error messages, check: '
            'https://cloud.google.com/apis/design/errors'.format(
                response.error.message))

def start_detection():
    for img in os.listdir('./upload_img/'): 
        detect_text('./upload_img/{}'.format(img),img)
        print(img,": done..") 