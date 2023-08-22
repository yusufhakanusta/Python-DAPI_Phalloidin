#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Aug 21 13:04:24 2023

@author: s92622yu
"""

import cv2
import numpy as np
import matplotlib.pyplot as plt
import os
import openpyxl
from skimage import measure

# Conversion factor: pixels to micrometers (µm)
conversion_factor = 0.1  # Example value: 1 pixel = 0.1 µm 


# Path to the directory containing the images
images_directory = '/Users/user/Documents/MATLAB/'

# List all image files in the directory
image_files = [f for f in os.listdir(images_directory) if f.endswith('.tif')]

# Create a new Excel workbook and add a worksheet
workbook = openpyxl.Workbook()
worksheet = workbook.active
worksheet.append(["Image File", "Cell Area (px^2)", "Perimeter (px)", "Circularity", "Solidity", "Eccentricity", "elongation", "distribution", "volume", "intensity"])

# Loop through each image file
for image_file in image_files:
    image_path = os.path.join(images_directory, image_file)
    phalloidin_image = cv2.imread(image_path, cv2.IMREAD_GRAYSCALE)
    
    # Preprocess the phalloidin image (thresholding, denoising, etc.)
    # You may need to adapt this preprocessing based on your images
    _, thresholded = cv2.threshold(phalloidin_image, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    denoised = cv2.medianBlur(thresholded, 5)
    
    # Find contours of cells
    contours, _ = cv2.findContours(denoised, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    # Loop through each cell contour
    for idx, contour in enumerate(contours):
        # Calculate cell properties
        area = cv2.contourArea(contour)
        if area < 10000:
            continue
        
        # Convert area and perimeter from pixels to µm
        area_um2 = area * (conversion_factor ** 2)
        perimeter_um = cv2.arcLength(contour, True) * conversion_factor
        perimeter = cv2.arcLength(contour, True)
        if perimeter != 0:
            circularity = (4 * np.pi * area_um2) / (perimeter_um ** 2)
        else:
            circularity = 0.0  # Default value for zero perimeter
        
        # Calculate hull area for solidity (if possible)
        if len(contour) > 2:  # Ensure the contour has enough points for hull calculation
            hull = cv2.convexHull(contour)
            hull_area = cv2.contourArea(hull) * (conversion_factor ** 2)
            
            # Calculate solidity only if hull_area is non-zero
            if hull_area != 0:
                solidity = area_um2 / hull_area
            else:
                solidity = 0.0  # Assign a default value (you can adjust this as needed)
        else:
            solidity = 0.0  # Default value for small or invalid contours
        
        # Calculate eccentricity and other properties
        moments = cv2.moments(contour)
        eccentricity = np.sqrt((moments["mu20"] + moments["mu02"]) ** 2 - 4 * (moments["mu11"] ** 2)) / (moments["mu20"] + moments["mu02"])
        
        # Approximate the contour to handle merged cells
        epsilon = 0.02 * cv2.arcLength(contour, True)
        approx_contour = cv2.approxPolyDP(contour, epsilon, True)
        
        # Calculate elongation from center for the approximate contour
        moments = cv2.moments(approx_contour)
        center = moments["m10"] / moments["m00"], moments["m01"] / moments["m00"]
        elongation = np.sqrt((moments["mu20"] - moments["mu02"])**2 + 4 * moments["mu11"]**2) / (moments["mu20"] + moments["mu02"])
        
        distribution = area_um2 / perimeter_um
        volume = area_um2
        intensity = np.sum(phalloidin_image * (denoised > 0))  # Intensity within the cell
        
        # Append cell properties to the Excel worksheet
        worksheet.append([image_file, area_um2, perimeter_um, circularity, solidity, eccentricity, elongation, distribution, volume, intensity])
    
    
    
    
# Save the Excel workbook
results_excel_path = '/Users/user/Documents/MATLAB/results.xlsx'
workbook.save(results_excel_path)

# Display the cell images (comment this out if not needed)
for image_file in image_files:
    image_path = os.path.join(images_directory, image_file)
    phalloidin_image = cv2.imread(image_path, cv2.IMREAD_GRAYSCALE)
    plt.imshow(phalloidin_image, cmap='gray')
    plt.title(image_file)
    plt.axis('off')
    plt.show()
