# NetCDF Extractor with Map Preview for Climate and Hydrological Datas

Download:
https://github.com/Alireza-Dindar/NetCDF-Climate-and-Hydrological-Data-Extractor/releases/download/v.1.3.0/NetCDFExtractorSetup.exe
<img width="500" height="200" alt="historical avg" src="https://github.com/user-attachments/assets/9a588bd8-623c-428b-9eb3-66d95ff59cc5" />


**New updates and fixes:**

-User can visualzie .nc files using Bicubic interpolation tool 
  
-Now users are able to see and select the files they want to extract from the parent folder.

-Users can select the necessary columns they want to see in excel outputs.

-Issues with not displaying coordinates in preview map are now fixed

-New version is compatible with ERA5 and RCM climate .NC products

-Issues with desktop shortcut are now fixed

-In case of massive data sizes, app automatically will convert the outputs to .CSV format to reduce the memory and space usage.

----------------------------------------------------------------------------------------------------------------------
# A desktop application built with Python (PyQt5) to visualize and extract data from NetCDF (.nc) files.

## Features
- 🌍 **Interactive Map:** Preview grid nodes on a world map using Folium before the extraction.
- 📅 **Time Slicing:** Automatically detects available date ranges (in case of multiple files, shortest period will be displayed) and allows subsetting.
- 📊 **Excel Export:** Extracts grid point coordinates within given region and climate data within a bounding box to respectively csv. and .xlsx.
- 🚀 **Batch Processing:** Handle multiple folders and files at once.

## How to Run 

1. Simply install and open the NetCDF Extractor.exe (Run as Adminstrator) 
2. Browse and find your NetCDF data parent folder 
3. Select subfolders in which .NC files reside
4. Specify coordinates of your desired region
5. Click on "Detect available  date range" button if there are no previous details available
6. Set the start and end date 
7. [optional] Should you have checking the available grid points of the dataset in mind, press "Preview grids for selected folder" and then "Show nodes on world map". After a map displayed, tick the "All nodes in box" option in the layout menu on the top-right corner. 
8. Filter the bands that you want to extract.
9. Select the target folder that you want the extracted excel files to be saved. 
10. Press "Start Extraction"
11. In case if you need the coordinates of the nodes available within the given boundaries, press "Export grid list to CSV" and choose the target path. 
12. In case of any errors, they will be displayed in Log box at the bottom 

 
---------------------------------------------------------------
!!! Feel free to give feedbacks on GitHub or via email: alireza.dindar.1998@gmail.com !!!
