# Importing required libraries
import os
os.environ['TF_ENABLE_ONEDNN_OPTS'] = '0'
import re
import numpy as np
import pandas as pd
import cv2
from openpyxl import load_workbook, Workbook
from openpyxl.drawing.image import Image as XLImage
from paddleocr import PPStructure,draw_structure_result,save_structure_res
from paddleocr import PaddleOCR, draw_ocr
import layoutparser as lp
import tensorflow as tf
import subprocess
from functions import pdf_to_images,save_images,read_image,get_ifsc

# Function to find the intersection of two boxes
def intersection(box_1, box_2):
  return [box_2[0], box_1[1],box_2[2], box_1[3]]

def iou(box_1, box_2):

  x_1 = max(box_1[0], box_2[0])
  y_1 = max(box_1[1], box_2[1])
  x_2 = min(box_1[2], box_2[2])
  y_2 = min(box_1[3], box_2[3])

  inter = abs(max((x_2 - x_1, 0)) * max((y_2 - y_1), 0))
  if inter == 0:
      return 0

  box_1_area = abs((box_1[2] - box_1[0]) * (box_1[3] - box_1[1]))
  box_2_area = abs((box_2[2] - box_2[0]) * (box_2[3] - box_2[1]))

  return inter / float(box_1_area + box_2_area - inter)

# Function to convert the pdf to images and save them
# def PdfToImage(pdf_path):
#     root_directory = r"C:\Users\Lenovo\OneDrive\Desktop\Folders\NaharOm\BSA\Main_Project\PdfData"
    
#     regex = r"BankStatement\\(.+)(?=\.pdf)"
    
#     folder_path = re.search(regex, pdf_path).group(1)

#     excel_folder_path = os.path.join(root_directory, folder_path, "ExcelData")
#     if not os.path.exists(excel_folder_path):
#         os.makedirs(excel_folder_path)
    
#     images_folder_path = os.path.join(root_directory, folder_path, "Images")
#     if not os.path.exists(images_folder_path):
#         os.makedirs(images_folder_path)
    
#     processed_images_folder_path = os.path.join(root_directory, folder_path, "ProcessedImages")
#     if not os.path.exists(processed_images_folder_path):
#         os.makedirs(processed_images_folder_path)
    
#     images = pdf_to_images(pdf_path)

#     save_images(images, images_folder_path)

# Convering the saved images to csv
def ImgToCsv(image_path,processed_images_path,excel_output_path):
    regex = r"Images\\(.+)(?=\.png)"
    ocr = PaddleOCR(lang='en')
    image_name = re.search(regex, image_path).group(1)
    output_list = []
    image = cv2.imread(image_path)

    image = image[..., ::-1]
    img_height = image.shape[0]
    img_width = image.shape[0]
    # load model
    model = lp.PaddleDetectionLayoutModel(config_path="lp://PubLayNet/ppyolov2_r50vd_dcn_365e_publaynet/config",
                                    threshold=0.5,
                                    label_map={0: "Text", 1: "Title", 2: "List", 3:"Table", 4:"Figure"},
                                    enforce_cpu=False,
                                    enable_mkldnn=True)#math kernel library
    # detect
    layout = model.detect(image)

    x_1=0
    y_1=0
    x_2=0
    y_2=0

    tables = []
    for l in layout:
        #print(l)
        if l.type == 'Table':
            x_1 = int(l.block.x_1)
            y_1 = int(l.block.y_1)
            x_2 = int(l.block.x_2)
            y_2 = int(l.block.y_2)
            tables.append(image[max(y_1-30,0):min(y_2+30,img_height),max(x_1-30,0):min(x_2+30,img_width)])
    
    if not len(tables):
        return []
    
    k = 0
    for i, table in enumerate(tables):
        cropped_img_path = processed_images_path+f"\{image_name}_cropped_{i}.png"
        cv2.imwrite(cropped_img_path,table) # type: ignore
        # if table.shape[0] > 4000 :
        #     # Calculate the middle row index
        #     middle_row = table.shape[0] // 2

        #     # Initialize the sum of intensities
        #     intensity_sum = 0

        #     # Find the row where the sum of intensities exceeds the threshold
        #     split_row = None
        #     for row in range(middle_row, table.shape[0]):
        #         count = np.count_nonzero(255 - table[row, :])
        #         if count < 10 :
        #             split_row = row
        #             break

        #     # Split the image into two images
        #     if split_row is not None:
        #         image1 = table[:split_row, :]
        #         cropped_img_path1 = processed_images_path+f"\{image_name}_cropped_{k}.png"
        #         k += 1
        #         cv2.imwrite(cropped_img_path1,image1)
        #         image2 = table[split_row:, :]
        #         cropped_img_path2 = processed_images_path+f"\{image_name}_cropped_{k}.png"
        #         k += 1
        #         cv2.imwrite(cropped_img_path2,image2)
        #     else:
        #         cv2.imwrite(cropped_img_path,table)
        #         k += 1
        # else :
        #     cv2.imwrite(cropped_img_path,table)
        #     k += 1

        output = ocr.ocr(cropped_img_path)[0]

        out_array = CroppedToCSV(output,tables,i,processed_images_path,image_name)
        array_output_path = excel_output_path+f"\{image_name}_output_{i}.csv" # type: ignore
        output_list.append(array_output_path)
        
        pd.DataFrame(out_array).to_csv(array_output_path,index=False,
                                       header=False)
    
    return output_list

def CroppedToCSV(output,tables,i,processed_images_path,image_name):
    table = tables[i]
    img_height = table.shape[0]
    img_width = table.shape[1]

    boxes = [line[0] for line in output]
    texts = [line[1][0] for line in output]
    probabilities = [line[1][1] for line in output]

    image_boxes = table.copy()

    for box,text in zip(boxes,texts):
        cv2.rectangle(image_boxes, (int(box[0][0]),int(box[0][1])), (int(box[2][0]),int(box[2][1])),(0,0,255),1)
        cv2.putText(image_boxes, text,(int(box[0][0]),int(box[0][1])),cv2.FONT_HERSHEY_SIMPLEX,1,(222,0,0),1)

    boxxed_img_path = processed_images_path+f"\{image_name}_detections_{i}.jpg"
    cv2.imwrite(boxxed_img_path, image_boxes)

    horiz_boxes = []
    vert_boxes = []

    for box in boxes:
        x_h, x_v = 0,int(box[0][0])
        y_h, y_v = int(box[0][1]),0
        width_h,width_v = img_width, int(box[2][0]-box[0][0])
        height_h,height_v = int(box[2][1]-box[0][1]),img_height

        horiz_boxes.append([x_h,y_h,x_h+width_h,y_h+height_h])
        vert_boxes.append([x_v,y_v,x_v+width_v,y_v+height_v])
    
    horiz_out = tf.image.non_max_suppression(
        horiz_boxes,
        probabilities,
        max_output_size = 1000,
        iou_threshold=0.1,
        score_threshold=float('-inf'),
        name=None)

    horiz_lines = np.sort(np.array(horiz_out))

    vert_out = tf.image.non_max_suppression(
        vert_boxes,
        probabilities,
        max_output_size = 1000,
        iou_threshold=0.1,
        score_threshold=float('-inf'),
        name=None
    )

    vert_lines = np.sort(np.array(vert_out))


    out_array = [["" for i in range(len(vert_lines))] for j in range(len(horiz_lines))]

    unordered_boxes = []

    for i in vert_lines:
        unordered_boxes.append(vert_boxes[i][0])
    
    ordered_boxes = np.argsort(unordered_boxes)

    for i in range(len(horiz_lines)):
        for j in range(len(vert_lines)):
            resultant = intersection(horiz_boxes[horiz_lines[i]], vert_boxes[vert_lines[ordered_boxes[j]]] )

            for b in range(len(boxes)):
                the_box = [boxes[b][0][0],boxes[b][0][1],boxes[b][2][0],boxes[b][2][1]]
                if(iou(resultant,the_box)>0.1):
                    out_array[i][j] = texts[b]

    out_array=np.array(out_array)

    return out_array
