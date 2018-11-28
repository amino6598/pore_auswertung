# -*- coding: utf-8 -*-
"""
Created on Tue Nov 27 14:44:12 2018

@author: q463532
"""

import cv2
import numpy as np

img = cv2.imread('PROT_181017_1.jpg')

def nothing(x):
    pass
num_rows, num_cols = img.shape[:2]


cv2.namedWindow("Rotate, Press 'C' to cancel",cv2.WINDOW_NORMAL)
cv2.createTrackbar("Rotate","Rotate, Press 'C' to cancel",0,20,nothing)
switch = '0 : CW \n1: CCW'
cv2.createTrackbar(switch,"Rotate, Press 'C' to cancel",0,1,nothing)
while(1):
    s = cv2.cv2.getTrackbarPos(switch,"Rotate, Press 'C' to cancel")
    rot = cv2.getTrackbarPos("Rotate","Rotate, Press 'C' to cancel")
    if s == 0:
        rotation_matrix = cv2.getRotationMatrix2D((num_cols/2, num_rows/2), -rot, 1)
    else:
        rotation_matrix = cv2.getRotationMatrix2D((num_cols/2, num_rows/2), rot, 1)
    img = cv2.warpAffine(img, rotation_matrix, (num_cols, num_rows))
    cv2.imshow("Rotate, Press 'C' to cancel", img)
    key = cv2.waitKey(0) & 0xFF
    if key == ord('c'):
    #if key == 27:
        break
cv2.destroyAllWindows()