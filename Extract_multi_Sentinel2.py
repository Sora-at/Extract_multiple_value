import os 
import arcpy 
import numpy
import openpyxl
import pandas as pd
from arcpy import env
from arcpy.sa import*


#Set location file
env.workspace = "D:/Drive1/GIS/Rubber/KU_rubber plot_reshape/2022"

#Import point & copy point
inPointFeatures = arcpy.CopyFeatures_management("NM_Point5.shp","NM_20221229_QVE.shp")

#Import raster
path_Base = "D:/Drive1/Work_Sentinal/QVE/2022_12/MSIL2A_20221229T033139_N0509_R018_T48QVE_20221229T065723/S2B_MSIL2A_20221229T033139_N0509_R018_T48QVE_20221229T065723_"
path_Band1 =  path_Base +  's2resampled.data/' 
path_Bio_f1 = path_Base +  'biophysical10m.data/' 
path_Bio_l1 = path_Base +  'biophysical.data/' 
path_Band2 = path_Base + 'mosaic.data/'
path_Bio_f2 = path_Base + 'mosaic_biophysical10m.data/'
path_Bio_l2 = path_Base + 'mosaic_biophysical.data/'

raster1 =[[path_Band1 + "B1.img"],
          [path_Band1 + "B2.img"], 
          [path_Band1 + "B3.img"],
          [path_Band1 + "B4.img"],
          [path_Band1 + "B5.img"],
          [path_Band1 + "B6.img"],
          [path_Band1 + "B7.img"],
          [path_Band1 + "B8.img"],
          [path_Band1 + "B8A.img"],
          [path_Band1 + "B9.img"],
          [path_Band1 + "B11.img"],
          [path_Band1 + "B12.img"],
          [path_Band1 + "quality_cloud_confidence.img"],
          [path_Bio_f1 + "fapar.img"],
          [path_Bio_f1 + "fcover.img"],
          [path_Bio_f1 + "lai.img"],
          [path_Bio_l1 + "lai_cab.img"],
          [path_Bio_l1 + "lai_cw.img"]
         ]
raster2 =[[path_Band2 + "B1.img"],
          [path_Band2 + "B2.img"], 
          [path_Band2 + "B3.img"],
          [path_Band2 + "B4.img"],
          [path_Band2 + "B5.img"],
          [path_Band2 + "B6.img"],
          [path_Band2 + "B7.img"],
          [path_Band2 + "B8.img"],
          [path_Band2 + "B8A.img"],
          [path_Band2 + "B9.img"],
          [path_Band2 + "B11.img"],
          [path_Band2 + "B12.img"],
          [path_Band2 + "quality_cloud_confidence.img"],
          [path_Bio_f2 + "fapar.img"],
          [path_Bio_f2 + "fcover.img"],
          [path_Bio_f2 + "lai.img"],
          [path_Bio_l2 + "lai_cab.img"],
          [path_Bio_l2 + "lai_cw.img"]
         ]

#Extract multi value to point
arcpy.CheckOutExtension("Spatial")
#raster1 for normal  file & raster2 for mosaic file
#raster1[0:] this file has B1 & raster1[1:] this file has not B1 and start index [1:] is B2
ExtractMultiValuesToPoints(inPointFeatures, raster1[0:], "NONE") 

#Export to excel
out_path = 'D:/Drive2/Work_Sentinal/Excel_result/Rubber/V2/Point_Extract_Pixel/Nakonpanom2022_QVE/Extract2/'
out_xls = out_path + 't1.xls' 
arcpy.TableToExcel_conversion(inPointFeatures, out_xls)

#Convert xls to xlsx
pd.read_excel(out_xls).to_excel(out_path + "NM_20221229_QVE.xlsx")

#remove .xls
os.remove(out_xls)