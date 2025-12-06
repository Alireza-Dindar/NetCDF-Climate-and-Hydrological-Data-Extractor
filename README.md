# NetCDF Extractor with Map Preview for Climate and Hydrological Datas

A desktop application built with Python (PyQt5) to visualize and extract data from NetCDF (.nc) files.

## Features
- 🌍 **Interactive Map:** Preview grid nodes on a world map using Folium before the extraction.
- 📅 **Time Slicing:** Automatically detects available date ranges (in case of multiple files, shortest period will be displayed) and allows subsetting.
- 📊 **Excel Export:** Extracts grid point coordinates within given region and climate data within a bounding box to respectively csv. and .xlsx.
- 🚀 **Batch Processing:** Handle multiple folders and files at once.

## How to Run (Source)
1. Clone the repository.
2. Simply install and open the NetCDF Extractor.exe
3. Browse and find your NetCDF data parent folder 
4. Select subfolders in which .NC files reside
5. Specify coordinates of your desired region
6. Click on "Detect available  date range" button if there are no previous details available
7. Set the start and end date 
8. [optional] Should you have checking the available grid points of the dataset in mind, press "Preview grids for selected folder" and then "Show nodes on world map". After a map displayed, tick the "All nodes in box" option in the layout menu on the top-right corner. 
9. Filter the bands that you want to extract.
10. Select the target folder that you want the extracted excel files to be saved. 
11. Press "Start Extraction"
12. In case if you need the coordinates of the nodes available within the given boundaries, press "Export grid list to CSV" and choose the target path. 
13. In case of any errors, they will be displayed in Log box at the bottom 

 
---------------------------------------------------------------
!!! Feel free to give feedbacks on GitHub or via email: alireza.dindar.1998@gmail.com !!!
