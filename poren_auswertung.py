import numpy as np
import cv2
import os
import matplotlib.pyplot as plt
import tkinter as tk
import xlsxwriter
from io import BytesIO
from matplotlib.ticker import FormatStrFormatter
import time
from tkinter import filedialog
import tkinter.simpledialog as tksd
import imutils

'''
Import single or multiple images images
'''
image_list1 = []
root = tk.Tk()
root.withdraw()
root.focus_set()
image_files = filedialog.askopenfilenames(parent=root,
                                     title="ein oder mehrere Bilder auswählen:",
                                     filetypes=[('All Files', '.*'),('JPEG', '.jpg'), ('PNG', '.png')])
for filename in image_files: 
    im = cv2.imread(filename)
    img_gray = cv2.cvtColor(im, cv2.COLOR_BGR2GRAY,0)
    image_list1.append(img_gray)


'''
Crop and rotate
'''
#------rotate image function--------------
def nothing(x):
    pass
def rotate_image(image):
    window_name_rot = "Rotieren: Press SPACE to rotate, C to save and close" 
    cv2.namedWindow(window_name_rot,cv2.WINDOW_NORMAL)
    cv2.createTrackbar("Angle",window_name_rot,0,20,nothing)
    switch = '0:CW\nCCW:-'
    cv2.createTrackbar(switch,window_name_rot,0,1,nothing)
    while(1):
        s = cv2.cv2.getTrackbarPos(switch,window_name_rot)
        rot = cv2.getTrackbarPos("Angle",window_name_rot)
        if s == 0:
            rotated = imutils.rotate_bound(image,rot)
        else:
            rotated = imutils.rotate_bound(image,-rot)
        cv2.imshow(window_name_rot, rotated)
        key = cv2.waitKey(0) & 0xFF
        if key == ord('c'):
        #if key == 27:
            break
    cv2.destroyAllWindows()
    return rotated

image_list = []
image_roi_name = []
for i in range(len(image_list1)):
    while(True):
        windowName = "Crop image (used for rotation in the next step): press SPACE to confirm selection"
        cv2.namedWindow(windowName,cv2.WINDOW_NORMAL)
        cv2.resizeWindow(windowName, 1100, 900)
        try:
            refPt1 = cv2.selectROI(windowName,image_list1[i])
            imCrop1 = image_list1[i][int(refPt1[1]):int(refPt1[1]+refPt1[3]), int(refPt1[0]):int(refPt1[0]+refPt1[2])]
            windowName_crop_output = "Press SPACE for next image, ESC for same image"
            cv2.namedWindow(windowName_crop_output,cv2.WINDOW_NORMAL)
            cv2.resizeWindow(windowName_crop_output, 1100,900)
            cv2.imshow(windowName_crop_output,imCrop1)
            key = cv2.waitKey(0) & 0xFF
            if key == 27:
                rot = rotate_image(imCrop1)
                image_list.append(rot)   
                name = tksd.askstring("","Enter the name of the Part:")
                image_roi_name.append(name)
                continue
            elif key == 32:
                cv2.destroyAllWindows()
                rot = rotate_image(imCrop1)
                image_list.append(rot)
                name = tksd.askstring("","Enter the name of the Part:")
                image_roi_name.append(name)
                #cv2.destroyAllWindows()
                break
        except Exception as e:
            cv2.destroyAllWindows()
            tk.messagebox.showerror("Error","Bounding box not found, Please crop the image")

'''
Select ROI
'''
image_roi_list = []
   
for i in range(len(image_list)):
    while(True):
        clone = image_list[i].copy
        windowName = "Select ROI: Press SPACE to confirm the image selection"
        cv2.namedWindow(windowName,cv2.WINDOW_NORMAL)
        cv2.resizeWindow(windowName, 1100,900)
        try:
            refPt = cv2.selectROI(windowName,image_list[i]) 
            imCrop = image_list[i][int(refPt[1]):int(refPt[1]+refPt[3]), int(refPt[0]):int(refPt[0]+refPt[2])]
            #image_roi_list.append(imCrop)
            #
            windowName_roi_output = "Press 'SPACE' to save and exit, ESC to reselect"
            cv2.namedWindow(windowName_roi_output,cv2.WINDOW_NORMAL)
            cv2.resizeWindow(windowName_roi_output, 1100,900)
            cv2.imshow(windowName_roi_output,imCrop)
            #
            #cv2.imshow("Press 'SPACE' to save and exit, ESC to reselect",imCrop)
            #print("press 'SPACE' to save and exit")
            key = cv2.waitKey(0) & 0xFF
            if key == 27:
                cv2.destroyWindow("Press 'SPACE' to save and exit, ESC to reselect")            
                continue
            elif key == 32:
                image_roi_list.append(imCrop)
                cv2.destroyAllWindows()
                break
        except Exception as e:
            cv2.destroyAllWindows()
            tk.messagebox.showerror("Error","ROI not found, Please select the ROI")

'''
Set the threshold and blend the scratches
'''
threshold_list1 = []
threshold_list = []
for i in range(len(image_roi_list)):
    plt.hist(image_roi_list[i].ravel(),256,[0,256])
    plt.show()
    def nothing1(x):
        pass 
    window_name_thresh = "Set threshold: press 'SPACE' to save and exit the threshold window. Default is 200" 
    cv2.namedWindow(window_name_thresh,cv2.WINDOW_NORMAL)
    cv2.resizeWindow(window_name_thresh, 1100,900)
    cv2.createTrackbar("Set threshold", window_name_thresh,200,255,nothing1)   
    while(1):
        value_threshold=cv2.getTrackbarPos("Set threshold", window_name_thresh)
        _,threshold_binary = cv2.threshold(image_roi_list[i],value_threshold,255,cv2.THRESH_BINARY)        
        cv2.imshow(window_name_thresh,threshold_binary)        
        key = cv2.waitKey(100) & 0xFF
        #if key == ord('c'):
        if key == 32:
            threshold_list1.append(value_threshold)
            plt.close()
            break        
    cv2.destroyAllWindows()
    plt.close('all')
    
    #Mouse paint brush (Blend the scratches)
    drawing = False
    def draw_circle(event,x,y,flags,param):
        global drawing,img,threshold_binary, threshold_binary
        if event == cv2.EVENT_LBUTTONDOWN:
            drawing = True
            cv2.circle(threshold_binary,(x,y),size,(255),-1)
        elif event == cv2.EVENT_MOUSEMOVE:
            if drawing == True:
                cv2.circle(threshold_binary,(x,y),size,(255),-1)
        elif event == cv2.EVENT_LBUTTONUP:
            drawing = False    
    threshold_binary2 = threshold_binary.copy()
    window_name_paint = "Kratzer ausblenden, Press SPACE to save, ESC to reset" 
    cv2.namedWindow(window_name_paint,cv2.WINDOW_NORMAL)
    cv2.resizeWindow(window_name_paint, 1100,900)
    cv2.createTrackbar("Grosse", window_name_paint,3,50,nothing1)
    cv2.setMouseCallback(window_name_paint, draw_circle)   
    while(1):
        size=cv2.getTrackbarPos("Grosse", window_name_paint)     
        cv2.imshow(window_name_paint,threshold_binary)        
        key = cv2.waitKey(100) & 0xFF
        if key == 27:
            threshold_binary = threshold_binary2.copy()
        elif key == 32:
            threshold_list.append(threshold_binary)
            break
    cv2.destroyAllWindows()
    
    
'''
segment the image, use the threshold and convert to binary
'''
segment_roi_all_images = []

#number of segments
number_of_segments = 10

for i in range(len(image_roi_list)):
    current_image = threshold_list[i]
    segment_step = int(len(current_image[:,0])/number_of_segments)
    segment_roi = []
    
    for j in range(0,int(len(current_image[:,0])),segment_step-1):
        segment = current_image[j:j+segment_step,:]
        roi = cv2.bitwise_not(segment)
        segment_roi.append(roi)
        segment_roi1 = segment_roi[::-1]
    segment_roi_all_images.append(segment_roi1)
    
    
#Convert the entire image to binary(without segmentation)
part_roi_all_images = []
for i in range(len(image_roi_list)):
    part_roi = cv2.bitwise_not(threshold_list[i])
    part_roi_all_images.append(part_roi)



'''
calculation of sq.pixel to sq.mm
'''

pixelsPerMetric = 107      #1mm = 107 pixels
pixelAreaPerMetric = pixelsPerMetric*pixelsPerMetric     #1sq.mm = x sq.pixel
pixelArea = 1/pixelAreaPerMetric     #Size of 1 sq.pixel
# Actual area = area in pixels x pixelArea(size of 1 sq.pixel) 


#minimum area of pore to be considered (in pixels)
min_area_pore = 10

#'''
#Porosity 
#''' 
##porosity of all the segments of all the images
#segment_porosity_all_images = []
#for i in range(len(segment_roi_all_images)):
#    segment_porosity = []
#    for j in range(1,len(segment_roi_all_images[i])):   
#        porosity1 = cv2.countNonZero(segment_roi_all_images[i][j])
#        total_density1 = segment_roi_all_images[i][j].size
#        porosity = ((porosity1)/total_density1)*100
#        #porosity="{0:.2f}".format(porosity)
#        segment_porosity.append(porosity)
#    segment_porosity_all_images.append(segment_porosity)
# 
##porosity of all the images
#part_porosity_all_images = []
#for i in range(len(part_roi_all_images)):
#    porosity = cv2.countNonZero(part_roi_all_images[i])
#    total_density = part_roi_all_images[i].size
#    part_porosity = (porosity/total_density)*100
#    #part_porosity="{0:.2f}".format(porosity)
#    part_porosity_all_images.append(part_porosity)
#    

'''
functions for the pore property
'''
#find the pores
def find_pores(roi):
    (_,contour1,_)  = cv2.findContours(roi, cv2.RETR_TREE,cv2.CHAIN_APPROX_NONE)
    contours_pores = []
    for con in contour1:
        area = cv2.contourArea(con)
        if min_area_pore < area:
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
    for j in range(1,len(segment_roi_all_images[i])):
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


'''
Porosity 
''' 
#porosity of all the segments of all the images
segment_porosity_all_images = []
for i in range(len(segment_roi_all_images)):
    segment_porosity = []
    #segment_roi_all_images_i = segment_roi_all_images[i]
    for j in range(1,len(segment_roi_all_images[i])):
        (_,contour_porosity_seg,_)  = cv2.findContours(segment_roi_all_images[i][j], cv2.RETR_TREE,cv2.CHAIN_APPROX_NONE)
        segPorosity = []
        for con_seg in contour_porosity_seg:
            area_seg = cv2.contourArea(con_seg)
            if min_area_pore < area_seg:
                segPorosity.append(area_seg)
        porosity_segment = sum(segPorosity)
        total_density1 = segment_roi_all_images[i][j].size
        porosity = (porosity_segment/total_density1)*100
        segment_porosity.append(porosity)
    segment_porosity_all_images.append(segment_porosity)

 
#porosity of all the images
part_porosity_all_images = []
for i in range(len(part_roi_all_images)):
    (_,contour_porosity,_)  = cv2.findContours(part_roi_all_images[i], cv2.RETR_TREE,cv2.CHAIN_APPROX_NONE)
    contoursPorosity = []
    for con_1 in contour_porosity:
        area = cv2.contourArea(con_1)
        if min_area_pore < area:
            contoursPorosity.append(area)
    porosity = sum(contoursPorosity)
    total_density = part_roi_all_images[i].size
    part_porosity = (porosity/total_density)*100
    part_porosity_all_images.append(part_porosity)
    


tk.messagebox.showinfo("Information","Exportieren der Ergebnisse zu Excel")

'''
Export to excel
'''
#--------Create a new file----------
timestr = time.strftime("%Y%m%d_%H%M")
directory_name = os.path.dirname(filename)
if os.path.exists(directory_name+'/Schilffbuilder_auswertung_'+timestr+'.xlsx'):
    workbook = xlsxwriter.Workbook(directory_name+'/Schilffbuilder_auswertung_'+timestr+'_1'+'.xlsx')
else:
    workbook = xlsxwriter.Workbook(directory_name+'/Schilffbuilder_auswertung_'+timestr+'.xlsx')

#timestr = time.strftime("%Y%m%d_%H%M")
#directory_name = os.path.dirname(filename)
#if os.path.exists(directory_name+'/Schilffbuild_auswertung_'+timestr+'.xlsx'):
#    output_file_name = directory_name+'/Schilffbuild_auswertung_'+timestr+'_1'+'.xlsx'
#else:
#    output_file_name = directory_name+'/Schilffbuild_auswertung_'+timestr+'.xlsx'
#    
#workbook = xlsxwriter.Workbook(output_file_name)
    
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
    
#name of the lasers/segments
layers = []
for num in range(number_of_segments):
    layers.append("Schichten "+str((num*20)+1)+"-"+str((num*20)+20))

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
    plt.ylabel('porosität(%)')
    plt.grid(True)
    fig.savefig(imgdata, format="png")
    part_worksheet[i].insert_image('D3', "",{'image_data': imgdata})
    plt.close()


#Position of the graphs
    step_pos = [0,9]      #column value
    step = (step_pos*int(number_of_segments/2)) #displaying the graphs side-by-side (column value)  
    
    step1_1 = []
    for j in range(int(number_of_segments/2)):
        step1_1.append([35+(j*25)]*2)
    step1 = sum(step1_1,[])    #position of area graph (row value)
    
    step2_1 = []
    for j in range(int(number_of_segments/2)):
        step2_1.append([(step1[-1]+35) + (j*25)]*2)
    step2 = sum(step2_1,[])         #position for solidity (row value)

    step3_1 = []
    for j in range(int(number_of_segments/2)):
        step3_1.append([(step2[-1]+30) + (j*25)]*2)
    step3 = sum(step3_1,[])         #position for circularity (row value)

    step4_1 = []
    for j in range(int(number_of_segments/2)):
        step4_1.append([(step3[-1]+30) + (j*25)]*2)
    step4 = sum(step4_1,[])         #position for elongation (row value)

#Area
    part_worksheet[i].write(29,0, 'Pore fläche', merge_title)
    part_worksheet[i].write(31,0, 'Maximum fläche:',bold_border)
    part_worksheet[i].write(31,1, "{:.4f} mm²".format(max(part_area_all_images[i])),bold_border)
    part_worksheet[i].write(32,0, 'Minimum fläche:',bold_border)
    part_worksheet[i].write(32,1, "{:.4f} mm²".format(min(part_area_all_images[i])),bold_border)
#Area chart
    #binwidth= (max(part_area_all_images[i])-min(part_area_all_images[i]))/10 
    binwidth = 0.002
    for j in range(len(step1)):
        imgdata = BytesIO()
        fig,ax = plt.subplots(tight_layout=True)
        counts,bins,patches = ax.hist(segment_area_all_images[i][j],bins=np.arange(min(part_area_all_images[i]),max(part_area_all_images[i])+binwidth,binwidth))
        ax.xaxis.set_ticks(np.arange(0,0.032,0.002))
        ax.yaxis.set_ticks(range(0,60,10))
        ax.set_xticklabels(np.arange(0,0.032,0.002),rotation='vertical')
        ax.xaxis.set_major_formatter(FormatStrFormatter('%0.4f'))
        plt.grid(True)
        plt.title(layers[j])
        plt.xlabel('Fläche (mm²)')
        plt.ylabel('Anzahl')
        fig.savefig(imgdata, format="png")
        #imgdata.seek()
        part_worksheet[i].insert_image(step1[j],step[j], "",{'image_data': imgdata})
        plt.close()
        
#Solidity
    solidity_title_row = step1[-1]+30
    part_worksheet[i].write(solidity_title_row,0, 'Solidität/Konvexität', merge_title)
    part_worksheet[i].merge_range(solidity_title_row,1,solidity_title_row,10, 'Es ist das Maß für die Unregelmäßigkeit, es ist 1 für Kreis, 0 für unregelmäßige Formen', merge_comment)
#Solidity chart
    for k in range(len(step2)):
        imgdata = BytesIO()
        fig,ax = plt.subplots(tight_layout=True)
        counts,bins,patches = ax.hist(segment_solidity_all_images[i][k], bins=10,range = (0,1))
        ax.set_xticks(bins)
        ax.xaxis.set_major_formatter(FormatStrFormatter('%0.1f'))
        ax.yaxis.set_ticks(range(0,60,10))
        plt.grid(True)
        plt.title(layers[k])
        plt.xlabel('Solidität')
        plt.ylabel('Anzahl')
        fig.savefig(imgdata, format="png")
        #imgdata.seek()
        part_worksheet[i].insert_image(step2[k],step[k], "",{'image_data': imgdata})
        plt.close()
        
#circularity
    circularity_title_row = step2[-1]+27 
    part_worksheet[i].write(circularity_title_row,0,'Rundheit', merge_title)
    part_worksheet[i].merge_range(circularity_title_row,1,circularity_title_row,10, 'es ist 1 für Kreis, 0 für linie', merge_comment)
#circularity chart
    for l in range(len(step3)):
        imgdata = BytesIO()
        fig,ax = plt.subplots(tight_layout=True)
        counts,bins,patches = ax.hist(segment_circularity_all_images[i][l], bins=10,range = (0,1))
        ax.set_xticks(bins)
        ax.xaxis.set_major_formatter(FormatStrFormatter('%0.1f'))
        ax.yaxis.set_ticks(range(0,60,10))
        plt.grid(True)
        plt.title(layers[l])
        plt.xlabel('Rundheit')
        plt.ylabel('Anzahl')
        fig.savefig(imgdata, format="png")
        #imgdata.seek()
        part_worksheet[i].insert_image(step3[l],step[l], "",{'image_data': imgdata})
        plt.close()

#elongation
    elongation_title_row = (step3[-1]+27)
    part_worksheet[i].write(elongation_title_row,0, 'Elongation', merge_title)
    part_worksheet[i].merge_range(elongation_title_row,1,elongation_title_row,10, 'auch bekannt als Inertia ratio oder ratio of major axis to minor axis', merge_comment)
#elongation chart    
    for m in range(len(step4)):
        imgdata = BytesIO()
        fig,ax = plt.subplots(tight_layout=True)
        counts,bins,patches = ax.hist(segment_elongation_all_images[i][m], bins=10,range = (0,1))
        ax.set_xticks(bins)
        ax.xaxis.set_major_formatter(FormatStrFormatter('%0.1f'))
        ax.yaxis.set_ticks(range(0,60,10))
        plt.grid(True)
        plt.title(layers[m])
        plt.xlabel('Elongation')
        plt.ylabel('Anzahl')
        fig.savefig(imgdata, format="png")
        #imgdata.seek()
        part_worksheet[i].insert_image(step4[m],step[m], "",{'image_data': imgdata})
        plt.close()

tk.messagebox.showinfo("Information","Excel Bericht generiert")
workbook.close()