import glob
import numpy as np
import cv2
import os
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import ttk
import xlsxwriter
from io import BytesIO
from matplotlib.ticker import FormatStrFormatter
import time

'''
Import the image
'''
image_list = []
image_name = []
for filename in glob.glob('O:/TI-6/TI-67/09_Intern/09-06_Temporaere_MA/Kumar/05_Probekoerper/01_MetalFab/*.jpg'):
    path, name = os.path.split(filename)
    name = name.replace('.jpg', '')
    image_name.append(name)
    im = cv2.imread(filename)
    img_gray = cv2.cvtColor(im, cv2.COLOR_BGR2GRAY,0)
    image_list.append(img_gray)

'''
function of 'input name dialogue box'
'''
def input_name():
    #callback()
    def callback():
        #print(e.get())
        image_roi_name.append(e.get())
        master.destroy()
    master = tk.Tk()
    e = ttk.Entry(master)
    NORM_FONT = ("Helvetica", 10)
    label = ttk.Label(master,text='Enter the name of the ROI', font=NORM_FONT)
    label.pack(side="top", fill="x", pady=10)
    e.pack(side = 'top',padx = 10, pady = 10)
    e.focus_set()
    b = ttk.Button(master, text = "OK", width = 10, command = callback)
    b.pack()
    master.mainloop()


'''
Select ROI
'''

image_roi_list = []
image_roi_name = []   
for i in range(len(image_list)):
    while(True):
        windowName = "Press SPACE to save the image selection or press C to cancel it"
        cv2.namedWindow(windowName,cv2.WINDOW_NORMAL)
        refPt = cv2.selectROI(windowName,image_list[i]) 
        imCrop = image_list[i][int(refPt[1]):int(refPt[1]+refPt[3]), int(refPt[0]):int(refPt[0]+refPt[2])]
        image_roi_list.append(imCrop)
        cv2.imshow("Cropped Image: press 'SPACE' to save and exit or 'ESC' to select another region in same image",imCrop)
        print("press 'SPACE' to save and exit or 'ESC' to select another region in same image")
        key = cv2.waitKey(0) & 0xFF
        if key == 27:
            cv2.destroyWindow("Cropped Image: press 'SPACE' to save and exit or 'ESC' to select another region in same image")
            input_name() #popup to  input name
            continue
        elif key == 32:
            cv2.destroyAllWindows()
            input_name() #popup to input name
            break

'''
Set the threshold
'''
threshold_list = []    
for i in range(len(image_roi_list)):    
    def nothing(x):
      pass
  
    cv2.namedWindow("Set threshold: press 'SPACE' to save and exit the threshold window. Default is 200",cv2.WINDOW_NORMAL)
    print("Set the threshold and then press SPACE to save it. Default is 200")
    cv2.createTrackbar("Set threshold", "Set threshold: press 'SPACE' to save and exit the threshold window. Default is 200",200,255,nothing)
    
    while(1):
        value_threshold=cv2.getTrackbarPos("Set threshold", "Set threshold: press 'SPACE' to save and exit the threshold window. Default is 200")
        _,threshold_binary = cv2.threshold(image_roi_list[i],value_threshold,255,cv2.THRESH_BINARY)        
        cv2.imshow("Set threshold: press 'SPACE' to save and exit the threshold window. Default is 200",threshold_binary)
        plt.hist(image_roi_list[i].ravel(),256,[0,256])
        key = cv2.waitKey(100) & 0xFF
        #if key == ord('c'):
        if key == 32:
            threshold_list.append(value_threshold)
            plt.close()
            break
    cv2.destroyAllWindows()
    plt.close('all')

'''
segment the image, use the threshold and convert to binary
'''
segment_roi_all_images = []

for i in range(len(image_roi_list)):
    current_image = image_roi_list[i]
    segment_step = int(len(current_image[:,0])/10)
    segment_roi = []
    
    for j in range(0,int(len(current_image[:,0])),segment_step-1):
        segment = current_image[j:j+segment_step,:]
        ret,threshold = cv2.threshold(segment,threshold_list[i],255,cv2.THRESH_BINARY_INV)
        roi = threshold[:,:]
        segment_roi.append(roi)
        segment_roi1 = segment_roi[::-1]
    segment_roi_all_images.append(segment_roi1)
    
    
#Convert the entire image to binary(without segmentation)
part_roi_all_images = []
for i in range(len(image_roi_list)):
    _,part_roi = cv2.threshold(image_roi_list[i],threshold_list[i],255,cv2.THRESH_BINARY_INV)
    part_roi_all_images.append(part_roi)
    
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
#porosity of all the segments of all the images
segment_porosity_all_images = []
for i in range(len(segment_roi_all_images)):
    segment_porosity = []
    for j in range(1,len(segment_roi_all_images[i])):   
        porosity1 = cv2.countNonZero(segment_roi_all_images[i][j])
        total_density1 = segment_roi_all_images[i][j].size
        porosity = ((porosity1)/total_density1)*100
        #porosity="{0:.2f}".format(porosity)
        segment_porosity.append(porosity)
    segment_porosity_all_images.append(segment_porosity)
 
#porosity of all the images
part_porosity_all_images = []
for i in range(len(part_roi_all_images)):
    porosity = cv2.countNonZero(part_roi_all_images[i])
    total_density = part_roi_all_images[i].size
    part_porosity = (porosity/total_density)*100
    #part_porosity="{0:.2f}".format(porosity)
    part_porosity_all_images.append(part_porosity)


'''
functions for the pore property
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
#find pores in every segment of all the images
segment_pores_all_images = []
for i in range(len(segment_roi_all_images)):
    segment_pores = []
    for j in range(0,len(segment_roi_all_images[i])-1):
        contour = find_pores(segment_roi_all_images[i][j])  
        segment_pores.append(contour)
    segment_pores_all_images.append(segment_pores)

#find pores in the entire part
part_pores_all_images = []
for i in range(len(part_roi_all_images)):
    part_pores = find_pores(part_roi_all_images[i])
    part_pores_all_images.append(part_pores)

'''
Area
'''
#find pore area in every segment of all the images
segment_area_all_images = []
for i in range(len(segment_pores_all_images)):
    segment_area = []
    for j in range(len(segment_pores_all_images[i])):
        area1 = find_area(segment_pores_all_images[i][j])
        segment_area.append(area1)
        #segment_area = segment_area[::-1]
    segment_area_all_images.append(segment_area)

#find pore area in every part
part_area_all_images = []
for i in range(len(part_pores_all_images)):
    part_area = find_area(part_pores_all_images[i])
    part_area_all_images.append(part_area)


'''
Solidity
'''
#find pore solidity in every segment of all the images
segment_solidity_all_images = []
for i in range(len(segment_pores_all_images)):
    segment_solidity = []
    for j in range(len(segment_pores_all_images[i])):
        solidity1 = find_solidity(segment_pores_all_images[i][j])
        segment_solidity.append(solidity1)
    segment_solidity_all_images.append(segment_solidity)

#find pore solidity in every part
part_solidity_all_images = []
for i in range(len(part_pores_all_images)):
    part_solidity = find_solidity(part_pores_all_images[i])
    part_solidity_all_images.append(part_solidity)


'''
Circularity
'''
#find pore circularity in every segment of all the images
segment_circularity_all_images = []
for i in range(len(segment_pores_all_images)):
    segment_circularity = []
    for j in range(len(segment_pores_all_images[i])):
        circularity1 = find_circularity(segment_pores_all_images[i][j])
        segment_circularity.append(circularity1)
    segment_circularity_all_images.append(segment_circularity)

#find pore circularity in every part
part_circularity_all_images = []
for i in range(len(part_pores_all_images)):
    part_circularity = find_circularity(part_pores_all_images[i])
    part_circularity_all_images.append(part_circularity)

'''
Elongation/Inertia ratio/ratio of major axis to minor axis
'''
#find pore elongation in every segment of all the images
segment_elongation_all_images = []
for i in range(len(segment_pores_all_images)):
    segment_elongation = []
    for j in range(len(segment_pores_all_images[i])):
        elongation1 = find_elongation(segment_pores_all_images[i][j])
        segment_elongation.append(elongation1)
    segment_elongation_all_images.append(segment_elongation)

#find pore elongation in every part
part_elongation_all_images = []
for i in range(len(part_pores_all_images)):
    part_elongation = find_elongation(part_pores_all_images[i])
    part_elongation_all_images.append(part_elongation)

print('Exporting the results to excel')


'''
Export to excel
'''
#--------Create a new file----------
timestr = time.strftime("%Y%m%d_%H%M")
if os.path.exists('Schilffbuild_auswertung_'+timestr+'.xlsx'):
    workbook = xlsxwriter.Workbook('Schilffbuild_auswertung_'+timestr+'_1'+'.xlsx')
else:
    workbook = xlsxwriter.Workbook('Schilffbuild_auswertung_'+timestr+'.xlsx')
    
#-----formatting styles for the title----------
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

#-------create a tab/workseet for 'Zusammenfassung'---------------
worksheet1 = workbook.add_worksheet('zusammenfassung')

#-----Create tab/worksheet for all the part name
part_worksheet = []
for i in range(len(image_roi_name)):
    worksheet = workbook.add_worksheet(image_roi_name[i])
    part_worksheet.append(worksheet)
    
layers = ['Schichten 1-20','Schichten 21-40','Schichten 41-60','Schichten 61-80','Schichten 81-100',
          'Schichten 101-120','Schichten 121-140','Schichten 141-160','Schichten 161-180','Schichten 180-200']

'''
Zusammenfassung
'''
#title
worksheet1.set_row(0,20)
worksheet1.set_column('A:A',17)
worksheet1.set_column('E:E',17)
worksheet1.merge_range('A1:D1', 'Zusammenfassung', merge_title)


worksheet1.write('A3','Würfel Porosität',merge_title)
worksheet1.write_column('A4',image_roi_name,bold_border)

#Write part porosity
row = 3
col = 1
for i in range(len(part_porosity_all_images)):
    worksheet1.write(row,col,"{:.2f} %".format(part_porosity_all_images[i]))
    row += 1
#write title for segment porosiy tabel    
worksheet1.write_column('E3',layers,bold_border)
worksheet1.write_row('F2',image_roi_name,bold_border)

#convert segment porosity to %
segment_porosity_all_images_1 = []
for i in range(len(segment_porosity_all_images)):
    segment_porosity_1 = ["{:.1f} %".format(j) for j in segment_porosity_all_images[i]]
    segment_porosity_all_images_1.append(segment_porosity_1)

#write segment porosity in table format
row = 2
col = 5
for i in range(len(segment_porosity_all_images_1)):
    worksheet1.write_column(row,col,segment_porosity_all_images_1[i])
    col += 1

'''
#Add pore properties for every image
'''

#porosity
for i in range(len(part_worksheet)):
    #part_worksheet[i].write_url('O2','internal:part_worksheet[i]!A159',string='Solidity')
    part_worksheet[i].set_row(0,20)
    part_worksheet[i].set_column('A:A',17)
    part_worksheet[i].set_column('B:B',12)
    part_worksheet[i].merge_range('A1:D1', 'Porosität', merge_title)
    part_worksheet[i].write(2,0, 'Gesamt porosität:',bold_border)
    part_worksheet[i].write(2,1,"{:.2f} %".format(part_porosity_all_images[i]),bold_border)
    part_worksheet[i].write(5,0, 'Schichten',bold_border)
    part_worksheet[i].write(5,1,"Porosität",bold_border)
    part_worksheet[i].write_column('A6',layers,bold_border)
    part_worksheet[i].write_column('B6',segment_porosity_all_images_1[i])
    #Porosity Chart
    imgdata = BytesIO()
    fig,ax = plt.subplots(tight_layout=True)
    ax.plot(segment_porosity_all_images[i],marker='o')
    ax.set_xticks(np.arange(0,len(segment_porosity_all_images[i]),1))
    ax.set_xticklabels(layers,rotation='vertical')
    ax.set_yticks(np.arange(0,7.5,0.5))
    #ax.set_yticks(range(15))
    #ax.set_yticklabels(np.arange(0,7.5,0.5))
    plt.title('Porosität')
    #plt.xlabel('Schichte Nr.')
    plt.ylabel('porosity(%)')
    plt.grid(True)
    fig.savefig(imgdata, format="png")
    part_worksheet[i].insert_image('D3', "",{'image_data': imgdata})
    plt.close()


#Area
    part_worksheet[i].write('A24', 'Pore fläche', merge_title)
    part_worksheet[i].write('A26', 'Maximum fläche:',bold_border)
    part_worksheet[i].write('B26', "{:.4f} mm²".format(max(part_area_all_images[i])),bold_border)
    part_worksheet[i].write('A27', 'Minimum fläche:',bold_border)
    part_worksheet[i].write('B27', "{:.4f} mm²".format(min(part_area_all_images[i])),bold_border)

#Area chart
    step = [30,30,55,55,80,80,105,105,130,130]
    step1 =[0,9,0,9,0,9,0,9,0,9]
    #binwidth= (max(part_area_all_images[i])-min(part_area_all_images[i]))/10 
    binwidth = 0.002
    for j in range(0,10,1):
        imgdata = BytesIO()
        fig,ax = plt.subplots(tight_layout=True)
        counts,bins,patches = ax.hist(segment_area_all_images[i][j],bins=np.arange(min(part_area_all_images[i]),max(part_area_all_images[i])+binwidth,binwidth))
        ax.xaxis.set_ticks(np.arange(0,0.032,0.002))
        ax.yaxis.set_ticks(range(0,160,10))
        ax.set_xticklabels(np.arange(0,0.032,0.002),rotation='vertical')
        ax.xaxis.set_major_formatter(FormatStrFormatter('%0.4f'))
        plt.grid(True)
        plt.title(layers[j])
        plt.xlabel('Area (mm²)')
        plt.ylabel('Frequency of occurence')
        fig.savefig(imgdata, format="png")
        #imgdata.seek()
        part_worksheet[i].insert_image(step[j],step1[j], "",{'image_data': imgdata})
        plt.close()
        
#Solidity
    part_worksheet[i].write('A159', 'Solidity/Convexity', merge_title)
    part_worksheet[i].merge_range('B159:I159', 'Its is measure of irregulatity, It is 1 for circle, 0 for highly irregular shape', merge_comment)
#Solidity chart    
    step2 = [162,162,187,187,212,212,237,237,262,262]
    step3 =[0,9,0,9,0,9,0,9,0,9]
    for k in range(0,10,1):
        imgdata = BytesIO()
        fig,ax = plt.subplots(tight_layout=True)
        counts,bins,patches = ax.hist(segment_solidity_all_images[i][k], bins=10,range = (0,1))
        ax.set_xticks(bins)
        ax.xaxis.set_major_formatter(FormatStrFormatter('%0.1f'))
        ax.yaxis.set_ticks(range(0,60,10))
        plt.grid(True)
        plt.title(layers[k])
        plt.xlabel('Solidity')
        plt.ylabel('Frequency of occurence')
        fig.savefig(imgdata, format="png")
        #imgdata.seek()
        part_worksheet[i].insert_image(step2[k],step3[k], "",{'image_data': imgdata})
        plt.close()
        
#circularity
    part_worksheet[i].write('A290', 'Circularity', merge_title)
    part_worksheet[i].merge_range('B290:I290', 'Circularity is 1 for circle, 0 for line', merge_comment)
#circularity chart
    step4 = [293,293,318,318,343,343,368,368,393,393]
    step5 =[0,9,0,9,0,9,0,9,0,9]
    for l in range(0,10):
        imgdata = BytesIO()
        fig,ax = plt.subplots(tight_layout=True)
        counts,bins,patches = ax.hist(segment_circularity_all_images[i][l], bins=10,range = (0,1))
        ax.set_xticks(bins)
        ax.xaxis.set_major_formatter(FormatStrFormatter('%0.1f'))
        ax.yaxis.set_ticks(range(0,60,10))
        plt.grid(True)
        plt.title(layers[l])
        plt.xlabel('Circularity')
        plt.ylabel('Frequency of occurence')
        fig.savefig(imgdata, format="png")
        #imgdata.seek()
        part_worksheet[i].insert_image(step4[l],step5[l], "",{'image_data': imgdata})
        plt.close()

#elongation
    part_worksheet[i].write('A420', 'Elongation', merge_title)
    part_worksheet[i].merge_range('B420:I420', 'also known as Inertia ratio/ratio of major axis to minor axis', merge_comment)
#elongation chart
    step6 = [423,423,448,448,473,473,498,498,523,523]
    step7 =[0,9,0,9,0,9,0,9,0,9]
    for m in range(0,10):
        imgdata = BytesIO()
        fig,ax = plt.subplots(tight_layout=True)
        counts,bins,patches = ax.hist(segment_elongation_all_images[i][m], bins=10,range = (0,1))
        ax.set_xticks(bins)
        ax.xaxis.set_major_formatter(FormatStrFormatter('%0.1f'))
        ax.yaxis.set_ticks(range(0,60,10))
        plt.grid(True)
        plt.title(layers[i])
        plt.xlabel('Elongation')
        plt.ylabel('Frequency of occurence')
        fig.savefig(imgdata, format="png")
        #imgdata.seek()
        part_worksheet[i].insert_image(step6[m],step7[m], "",{'image_data': imgdata})
        plt.close()

print('Result exported')
workbook.close()