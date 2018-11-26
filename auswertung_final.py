# -*- coding: utf-8 -*-
"""
Created on Tue Nov 13 08:34:03 2018

@author: q463532
"""

import cv2
import os
import numpy as np
from matplotlib import pyplot as plt
import xlsxwriter
from io import BytesIO
from matplotlib.ticker import FormatStrFormatter
from tkinter import filedialog
import tkinter as tk
import time

#---------------import image-----------------#
#import window
print('Select an image')
root = tk.Tk()
root.withdraw()
file_name = filedialog.askopenfilename(filetypes = (("jpeg files","*.jpg"),("png files","*.png*"),("all files","*.*")))
image = cv2.imread(file_name)
#image = cv2.imread('PROT_181017_1.jpg')
#image = cv2.imread(image_file)

if image is not None:
    img_gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY,0)
    print('Image imported')
else: print('Image not found')


'''
find ROI using mouse click
'''
#----------ROI selection using mouse click------------------------------------# 
#drawing = False
#refPt = []
##function to select ROI
#def draw_roi(event, x, y, flags, param):
#    global drawing, refPt, num, img_gray
#    if event == cv2.EVENT_LBUTTONDOWN:
#        refPt = [(x, y)]
#    elif event == cv2.EVENT_LBUTTONUP:
#        refPt.append((x, y))
#        cv2.rectangle(img_gray, refPt[0], refPt[1], (0, 255, 0), 2)
#
#clone = img_gray.copy()
#windowName = "Press 'Esc' to reset the image, 'spacebar' to save and exit the image selection"
#cv2.namedWindow(windowName,cv2.WINDOW_NORMAL)
#cv2.setMouseCallback(windowName, draw_roi)
#while (True):
#    #img_gray = cv2.resize(img_gray, (960, 960))
#    cv2.imshow(windowName,img_gray)
#    key = cv2.waitKey(1) & 0xFF
#    #press esc to reset
#    if key == 27:
#        img_gray = clone.copy()
#    #press spacebar to save and exit
#    elif key == 32:
#        break
#cv2.destroyAllWindows()
#new_image = image[refPt[0][1]:refPt[1][1],refPt[0][0]:refPt[1][0]]


while(True):
    windowName = "Press SPACE to save the image selection or press C to cancel it"
    #cv2.namedWindow(windowName,cv2.WINDOW_NORMAL)
    cv2.namedWindow(windowName,cv2.WINDOW_NORMAL)
    refPt = cv2.selectROI(windowName,img_gray)  
    imCrop = img_gray[int(refPt[1]):int(refPt[1]+refPt[3]), int(refPt[0]):int(refPt[0]+refPt[2])]
    cv2.imshow("Cropped Image: press 'SPACE' to save and exit or 'ESC' to restart image selection",imCrop)
    print("press 'SPACE' to save and exit or 'ESC' to restart image selection")
    key = cv2.waitKey(0) & 0xFF
    if key == 27:
        cv2.destroyWindow("Cropped Image: press 'SPACE' to save and exit or 'ESC' to restart image selection")
        continue
    elif key == 32:
        cv2.destroyAllWindows()
        break



#----------crop thr image with respect to bounding box------------------------#
new_image = image[int(refPt[1]):int(refPt[1]+refPt[3]), int(refPt[0]):int(refPt[0]+refPt[2])]
#new_image_seg = new_image
new_image = new_image[:,:,0]

#-------------drawthe ROI rectangle on the actual image------------------------

#image_rect = cv2.rectangle(image.copy(),refPt[0],refPt[1],(0,255,0),2)
image_rect = cv2.rectangle(image.copy(),(refPt[0],refPt[1]),(refPt[2],refPt[3]),(0,255,0),2)
#cv2.imshow('Image with ROI',image_rect)

#Set histogram
fig = plt.figure()
fig.canvas.set_window_title('Histogram to select apropriate threshold')
plt.hist(imCrop.ravel(),256,[0,256])   #Histogram for threshold
plt.show(True)


'''
set the threshold for binary image
'''

#------------set threshold using trackbar-----------------------#

def nothing(x):
  pass
cv2.namedWindow("Set threshold: press 'SPACE' to save and exit the threshold window. Default is 200",cv2.WINDOW_NORMAL)
print("Set the threshold and then press SPACE to save it. Default is 200")
cv2.createTrackbar("Set threshold", "Set threshold: press 'SPACE' to save and exit the threshold window. Default is 200",200,255,nothing)

while(1):
    value_threshold=cv2.getTrackbarPos("Set threshold", "Set threshold: press 'SPACE' to save and exit the threshold window. Default is 200")
    _,threshold_binary = cv2.threshold(new_image,value_threshold,255,cv2.THRESH_BINARY)
    #cv2.imshow("Set threshold: press 'c' to cancel and save the threshold. Default is 200",threshold_binary)
    cv2.imshow("Set threshold: press 'SPACE' to save and exit the threshold window. Default is 200",threshold_binary)
    key = cv2.waitKey(100) & 0xFF
    #if key == ord('c'):
    if key == 32:
        break
cv2.destroyAllWindows()
plt.close('all')


#-------------segment the images------------------------------------
segment_step = int(len(new_image[:,0])/10)
segment_roi = []


for i in range(0,int(len(new_image[:,0])),segment_step-1):
    segment = new_image[i:i+segment_step,:]
    ret,threshold = cv2.threshold(segment,value_threshold,255,cv2.THRESH_BINARY_INV)
    roi = threshold[:,:]
    segment_roi.append(roi)
#plt.imshow(segment_roi[9],'gray')         

#-------------find ROI of whole part (part_roi) and binarize it -------------------------------
part = new_image[:,:]
ret,part_roi = cv2.threshold(part,value_threshold,255,cv2.THRESH_BINARY_INV)
#plt.imshow(part_roi,'gray')
#---------------
'''
calculation of sq.pixel to sq.mm
'''

pixelsPerMetric = 107      #1mm = 107 pixels
pixelAreaPerMetric = pixelsPerMetric*pixelsPerMetric     #1sq.mm = x sq.pixel
pixelArea = 1/pixelAreaPerMetric     #Size of 1 sq.pixel
# Actual area = area in pixels x pixelArea(size of 1 sq.pixel)  
 

'''
Porosity
'''
#porosity of all the segments
segment_porosity = []

for i in range(1,11,1):   
    porosity1 = cv2.countNonZero(segment_roi[i])
    total_density1 = segment_roi[i].size
    porosity = ((porosity1)/total_density1)*100
    #porosity="{0:.2f}".format(porosity)
    segment_porosity.append(porosity)

#porosity of the entire part
porosity = cv2.countNonZero(part_roi)
total_density = part_roi.size
part_porosity = (porosity/total_density)*100
#part_porosity="{0:.2f}".format(porosity)

'''
functions
'''
#find the pores
def find_pores(roi):
    (_,contour1,_)  = cv2.findContours(roi, cv2.RETR_TREE,cv2.CHAIN_APPROX_NONE)
    contours_pores = []
    for con in contour1:
        area = cv2.contourArea(con)
        if 3 < area:
            contours_pores.append(con)
            
    return contours_pores

#find the area of the pores
def find_area(contour_array):
    area_contours = []
    for i in range(0,len(contour_array)):
        area = cv2.contourArea(contour_array[i])
        area_actual = pixelArea*area
        area_contours.append(area_actual)
    return area_contours

##Find solidity  #also convexity  
def find_solidity(contour_array):
    solidity_contours = []
    for i in range(0,len(contour_array)):
        area = cv2.contourArea(contour_array[i])
        hull = cv2.convexHull(contour_array[i])
        hull_area = cv2.contourArea(hull)
        solidity = float(area)/hull_area
        solidity_contours.append(solidity)
    return solidity_contours

#find circularity                        (1 for circle)
def find_circularity(contour_array):
    circularity_contours = []
    for i in range(0,len(contour_array)):
        area = cv2.contourArea(contour_array[i])
        perimeter = cv2.arcLength(contour_array[i],True)             #closed contour True, else false
        circularity = 4*np.pi*float(area)/(perimeter)**2
        circularity_contours.append(circularity)   
    return circularity_contours

#Elongation #intertia ratio #ratio of major axis to minor axis
def find_elongation(contour_array):
    elongation_contours = []
    for i in range(0,len(contour_array)):
        (x,y),(MA,ma),angle = cv2.fitEllipse(contour_array[i])
        inertia_ratio = MA/ma
        elongation_contours.append(inertia_ratio)
    return elongation_contours
'''
find pores
'''
#find pores in every segment
segment_pores = []
for i in range(0,10):
    contour = find_pores(segment_roi[i])  
    segment_pores.append(contour)


#find pores in the part
part_pores = find_pores(part_roi)
'''
Area
'''
#find pore area in every segment
segment_area = []
for j in range(len(segment_pores)):
    area1 = find_area(segment_pores[j])
    segment_area.append(area1)

#find pore area in the part
part_area = find_area(part_pores)

'''
solidity
'''
#solidity of pores in evry segment
segment_solidity = []
for j in range(len(segment_pores)):
    solidity1 = find_solidity(segment_pores[j])
    segment_solidity.append(solidity1)
    
#solidity of pores in the entire part
part_solidity = find_solidity(part_pores)    

'''
circularity
'''
#circularity of pores in every segment
segment_circularity = []
for j in range(len(segment_pores)):
    circularity1 = find_circularity(segment_pores[j])
    segment_circularity.append(circularity1)
    
#circularity of pores in the entire part
part_circularity = find_circularity(part_pores)

'''
Elongation/Inertia ratio/ratio of major axis to minor axis
'''
#elongation of pores in every segment
segment_elongation = []
for j in range(len(segment_pores)):
    elongation1 = find_elongation(segment_pores[j])
    segment_elongation.append(elongation1)

#elongation of pores in entire part
part_elongation = find_elongation(part_pores)



print('Exporting the results to excel')
'''
export to excel
'''
#-----------Create an Excel file------------------------------
#extract file name
path, file = os.path.split(file_name)
new_file_name = file.replace('.jpg','')

timestr = time.strftime("%Y%m%d_%H%M")
if os.path.exists(new_file_name+'_'+timestr+'.xlsx'):
    workbook = xlsxwriter.Workbook(new_file_name+'_'+timestr+"_1"+'.xlsx')
else:
    workbook = xlsxwriter.Workbook(new_file_name+'_'+timestr+'.xlsx')

#-------formatting for title---------------------------------
merge_title = workbook.add_format({
    'bold': 1,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': 'yellow'})
merge_comment = workbook.add_format({
    'align': 'left',
    'valign': 'vcenter',})
bold = workbook.add_format({'bold': 1})
bold_border = bold = workbook.add_format({'bold': 1,'border': True,'align': 'center'})


#---------Create a tab(worksheet)----------------------
worksheet1 = workbook.add_worksheet('porosität')
 
'''
porosity
'''
#title
worksheet1.set_row(0,20)
worksheet1.set_column('A:A',17)
worksheet1.merge_range('A1:D1', 'Porosität', merge_title)
worksheet1.write(2,0, 'Gesamt porosität:',bold_border)
worksheet1.write(2,1,"{:.2f} %".format(part_porosity),bold_border)
worksheet1.write(5,0, 'Schichten',bold_border)
worksheet1.write(5,1,"Porosität",bold_border)

#layer number
layers = ['Schichten 181-200','Schichten 161-180','Schichten 141-160','Schichten 121-140','Schichten 101-120',
          'Schichten 81-100','Schichten 61-80','Schichten 41-60','Schichten 21-40','Schichten 1-20']
row = 6
col = 0
for i in range(0,len(layers)):
    worksheet1.write(row,col, layers[i])
    row += 1
#porosity of each layer
row1 = 6
col1 = 1
for i in range(0,len(segment_porosity)):
    worksheet1.write(row1,col1,"{:.2f} %".format(segment_porosity[i]))
    row1 += 1

worksheet1.write_column('D7',segment_porosity)    

#add porosity chart
imgdata = BytesIO()
fig,ax = plt.subplots(tight_layout=True)
ax.plot(segment_porosity,marker='o')
ax.set_xticks(np.arange(0,len(segment_porosity),1))
ax.set_xticklabels(layers,rotation='vertical')
plt.title('Porosität')
#plt.xlabel('Schichte Nr.')
plt.ylabel('porosity(%)')
plt.grid(True)
fig.savefig(imgdata, format="png")
worksheet1.insert_image('D6', "",{'image_data': imgdata})
plt.close()

#--------Display the image with ROI--------
#worksheet1.merge_range('A32:D32', 'Image with ROI', merge_title)
#roi_image = BytesIO()
#plt.imshow(image_rect)
#plt.savefig(roi_image, format="png")
#plt.tight_layout()
#worksheet1.insert_image('D34', "",{'image_data': roi_image})
#plt.close()


'''
Area
'''
worksheet2 = workbook.add_worksheet('fläche')
worksheet2.merge_range('A1:D1', 'Pore fläche', merge_title)
worksheet2.set_row(0,20)


#worksheet2.write(3,0, 'Maximum fläche:',bold_border)
worksheet2.merge_range('A4:B4', 'Maximum fläche:',bold_border)
#worksheet2.write(3,1,"{:.4f} mm²".format(max(part_area)),bold_border)
worksheet2.merge_range('C4:D4', "{:.4f} mm²".format(max(part_area)),bold_border)
worksheet2.merge_range('A5:B5', 'Minimum fläche:',bold_border)
worksheet2.merge_range('C5:D5', "{:.4f} mm²".format(min(part_area)),bold_border)


step = [7,7,30,30,55,55,80,80,105,105]
step1 =[0,9,0,9,0,9,0,9,0,9]
binwidth= (max(part_area)-min(part_area))/10

for i in range(0,10):
    imgdata = BytesIO()
    fig,ax = plt.subplots(tight_layout=True)
    counts,bins,patches = ax.hist(segment_area[i],bins=np.arange(min(part_area),max(part_area)+binwidth,binwidth))
    ax.xaxis.set_ticks(np.arange(min(part_area),max(part_area)+0.002,0.002))
    ax.yaxis.set_ticks(range(0,160,10))
    ax.set_xticklabels(bins,rotation='vertical')
    ax.xaxis.set_major_formatter(FormatStrFormatter('%0.4f'))
    plt.grid(True)
    plt.title(layers[i])
    plt.xlabel('Area (mm²)')
    plt.ylabel('Frequency of occurence')
    fig.savefig(imgdata, format="png")
    #imgdata.seek()
    worksheet2.insert_image(step[i],step1[i], "",{'image_data': imgdata})
    plt.close()



'''
Solidity
'''

worksheet3 = workbook.add_worksheet('solidity')

#title
worksheet3.merge_range('A1:D1', 'Solidity/Convexity', merge_title)
worksheet3.merge_range('E1:L1', 'Its is measure of irregulatity, It is 1 for circle, 0 for highly irregular shape', merge_comment)
worksheet3.set_row(0,20)
worksheet3.set_column('A:A',17)



step = [5,5,30,30,55,55,80,80,105,105]
step1 =[0,9,0,9,0,9,0,9,0,9]



for i in range(0,10):
    imgdata = BytesIO()
    fig,ax = plt.subplots(tight_layout=True)
    counts,bins,patches = ax.hist(segment_solidity[i], bins=10,range = (0,1))
    ax.set_xticks(bins)
    ax.xaxis.set_major_formatter(FormatStrFormatter('%0.1f'))
    ax.yaxis.set_ticks(range(0,110,10))
    plt.grid(True)
    plt.title(layers[i])
    plt.xlabel('Solidity')
    plt.ylabel('Frequency of occurence')
    fig.savefig(imgdata, format="png")
    #imgdata.seek()
    worksheet3.insert_image(step[i],step1[i], "",{'image_data': imgdata})
    plt.close()
worksheet3.insert_image('B133','convexity.png')    
'''
Circularity
'''
worksheet4 = workbook.add_worksheet('circularity')

#title
worksheet4.merge_range('A1:D1', 'Circularity', merge_title)
worksheet4.merge_range('E1:I1', 'Circularity is 1 for circle, 0 for line', merge_comment)
worksheet4.set_row(0,20)
worksheet4.set_column('A:A',17)

step = [5,5,30,30,55,55,80,80,105,105]
step1 =[0,9,0,9,0,9,0,9,0,9]

for i in range(0,10):
    imgdata = BytesIO()
    fig,ax = plt.subplots(tight_layout=True)
    counts,bins,patches = ax.hist(segment_circularity[i], bins=10,range = (0,1))
    ax.set_xticks(bins)
    ax.xaxis.set_major_formatter(FormatStrFormatter('%0.1f'))
    ax.yaxis.set_ticks(range(0,60,10))
    plt.grid(True)
    plt.title(layers[i])
    plt.xlabel('Circularity')
    plt.ylabel('Frequency of occurence')
    fig.savefig(imgdata, format="png")
    #imgdata.seek()
    worksheet4.insert_image(step[i],step1[i], "",{'image_data': imgdata})
    plt.close()

'''
Elongation/Inertia ratio/ratio of major axis to minor axis
'''
worksheet5 = workbook.add_worksheet('elongation')

#title
worksheet5.merge_range('A1:D1', 'Elongation', merge_title)
worksheet5.merge_range('E1:I1', 'also known as Inertia ratio/ratio of major axis to minor axis', merge_comment)
worksheet5.set_row(0,20)
worksheet5.set_column('A:A',17)

step = [5,5,30,30,55,55,80,80,105,105]
step1 =[0,9,0,9,0,9,0,9,0,9]

for i in range(0,10):
    imgdata = BytesIO()
    fig,ax = plt.subplots(tight_layout=True)
    counts,bins,patches = ax.hist(segment_elongation[i], bins=10,range = (0,1))
    ax.set_xticks(bins)
    ax.xaxis.set_major_formatter(FormatStrFormatter('%0.1f'))
    ax.yaxis.set_ticks(range(0,60,10))
    plt.grid(True)
    plt.title(layers[i])
    plt.xlabel('Elongation')
    plt.ylabel('Frequency of occurence')
    fig.savefig(imgdata, format="png")
    #imgdata.seek()
    worksheet5.insert_image(step[i],step1[i], "",{'image_data': imgdata})
    plt.close()
    


workbook.close()

print('Result file created')
print('Program ended')
print('')
print('')
print('')
print('------------------------------------------------------')
print('              Gesamt porosität:',"{0:.2f}".format(part_porosity),'%')
print('------------------------------------------------------')
print('')
print('')
print('')