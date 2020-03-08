# -*- coding: utf-8 -*-
import requests
import json
import time
import xlsxwriter
import os
import json
import cv2
import numpy as np
import xml.etree.ElementTree as ET
from pdf2image import convert_from_path

subscription_key = "5a0c67d7759345d69939db5a53fc4336"
endpoint = "https://southeastasia.api.cognitive.microsoft.com/vision/v2.0/ocr"
headers = {'Ocp-Apim-Subscription-Key': subscription_key, 'Content-Type':'application/octet-stream'}
params = {'language': 'ja', 'detectOrientation ': 'true'}

def call_api(image_dir, filename):
    image_path = image_dir + filename
    extension = filename.split('.')[-1]        
    if extension == 'pdf':
        convert_from_path(image_path, dpi=300, output_folder=image_dir)
        print("convert successful!")
        jpeged_image = filename.replace(".pdf", ".jpeg")
        jpeged_image[0].save('{}.jpg'.format(filename.strip('pdf')), 'jpg')
        data = open('{}{}.jpg'.format(image_dir, filename.strip('.pdf')), 'wb')
        response = requests.post(endpoint, headers=headers, params=params, data=data)
        filedata = response.json()
    else:    
        data = open(image_path, 'rb')
        response = requests.post(endpoint, headers=headers, params=params, data=data)
        filedata = response.json()
        print("call api for {} ok".format(filename))
    return filedata

def read_json(json): 
    pred_boxes = []
    if 'regions' in json.keys():
        for i in range(len(json['regions'])):
            for obj in json['regions'][i]['lines']:
                str_coord = obj['boundingBox']
                splits = str_coord.split(',')
                word = ''
                for w in obj["words"]:
                    word += w["text"]
                box = [int(splits[0]),int(splits[1]),int(splits[0])+int(splits[2]),int(splits[1])+int(splits[3]),word] 
                pred_boxes.append(box)
    else:
        print('No response was returned from Microsoft API')
        print(json['error']['message'])
    print(pred_boxes)

    return pred_boxes
    
def write_xml_word(xml_path, img_name, word_boxes, img_height, img_width, img_depth):
    root = ET.Element('annotation')
    file_node = ET.SubElement(root, 'filename')
    file_node.text = img_name
    
    node_size = ET.SubElement(root,'size')
    node_width = ET.SubElement(node_size, 'width')
    node_width.text = str(img_width)
    node_height = ET.SubElement(node_size,'height')
    node_height.text = str(img_height)
    node_depth = ET.SubElement(node_size,'depth')
    node_depth.text = str(img_depth)        

    for wb in word_boxes:
        xmin, ymin, xmax, ymax, word = wb
        node_object = ET.SubElement(root,'object')
        node_name = ET.SubElement(node_object,'name')
        node_name.text = word
        node_difficult = ET.SubElement(node_object,'difficult')
        node_difficult.text = '0'    
        node_bndbox = ET.SubElement(node_object,'bndbox')

        node_xmin = ET.SubElement(node_bndbox,'xmin')
        node_xmin.text = str(xmin)
        node_ymin = ET.SubElement(node_bndbox,'ymin')
        node_ymin.text = str(ymin)
        node_xmax = ET.SubElement(node_bndbox,'xmax')
        node_xmax.text = str(xmax)
        node_ymax = ET.SubElement(node_bndbox,'ymax')
        node_ymax.text = str(ymax)       
            
    tree = ET.ElementTree(root)
    tree.write(xml_path, encoding='utf-8', xml_declaration=True) 

def convert(filename, json_data, des_path):
    extension = filename.split('.')[-1]        
    xml_path = '{}{}'.format(des_path, filename.replace('.{}'.format(extension),".xml"))
    img = cv2.imread('{}/{}'.format(image_dir, filename))  
    h, w, depth = img.shape
    word_boxes = read_json(json_data)
    write_xml_word(xml_path, filename, word_boxes, h, w, depth)
    
    open('./jijilla/json_results/{}json'.format(filename.strip(extension)), 'w').write(
        json.dumps(json_data, sort_keys=True, indent=4, ensure_ascii=False)
    )
    
    print("annotation file {} created\n".format(xml_path))     
        
if __name__ == '__main__':
    image_dir = '/home/asilla/ichijo/microsoft_vision/jijilla/images/'
    des_path = '/home/asilla/ichijo/microsoft_vision/jijilla/annotations/'
    files = os.listdir(image_dir)
    for i, filename in enumerate(os.listdir(image_dir)):
        res = call_api(image_dir, filename)
        print(res)
        print("{} finished recognizing".format(filename))
        convert(filename, res, des_path)
    if not os.path.isdir(des_path):
        os.makedirs(des_path)


def write_to_excel(json_by_line):
    workbook = xlsxwriter.Workbook('microsoft_vision_eval.xlsx') 
    cell_format = workbook.add_format()
    cell_format.set_border()

    worksheet = workbook.add_worksheet() 
    worksheet.write('A1', 'T',cell_format) 
    worksheet.write('B1', 'File name',cell_format) 
    worksheet.write('C1', 'Predict',cell_format) 
    worksheet.write('D1', 'Real',cell_format) 
    worksheet.write('E1', 'type',cell_format) 
    worksheet.write('J1', 'Error',cell_format)
    worksheet.write('K1', 'Total',cell_format)     

    line_count = 1

    for i, textline in enumerate(json_by_line, 1):
        for text in textline:
            line_count += 1
            if len(textline) == 0:
                continue
            worksheet.write('A' + str(line_count), str(line_count-1), cell_format)
            worksheet.write('B'+ str(line_count), 'msvision_eval00{}.jpg'.format(i), cell_format)
            worksheet.write('C' + str(line_count), text, cell_format)
            
    workbook.close() 