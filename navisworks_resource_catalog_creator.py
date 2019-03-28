#C.C.J. Claus 2019, Amersfoort, the Netherlands.
#A python script to create STABU XML files for Navisworks Manage 2019

from lxml import etree
import xlrd
import uuid 

###########################################################################################
####################### GET STABU DATA FROM EXTERNAL EXCEL FILE ###########################
###########################################################################################
excel_file_name = "stabu_hoofdstukken_2017-1.xlsx"
    
workbook = xlrd.open_workbook(excel_file_name)
worksheet = workbook.sheet_by_name("Table 1") 

current_row = 0
row_list = []
stabu_list = []


while current_row < (worksheet.nrows - 1):
    current_row += 1
    row = worksheet.row(current_row)
    row_list.append(row)
    stabu_list = [[ele.value for ele in each] for each in row_list ]


###########################################################################################
####################### INITIALISE PARAMETERS FOR XML FILE  ###############################
########################################################################################### 

take_off = "Takeoff"
catalog = "Catalog"
resource_group = "ResourceGroup"

root = etree.Element(take_off)
doc = etree.SubElement(root, catalog)

#<?xml version="1.0" encoding="utf-8"?>
#<Takeoff xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns="http://download.autodesk.com/us/navisworks/schemas/nw-TakeoffCatalog-10.0.xsd">

for j in stabu_list:
    if len(j[0]) == 2:
        
     
        resourcegroup = etree.SubElement(doc, resource_group, Name=str(j[1]), RBS=str(j[0]), CatalogId=str(uuid.uuid4())) 
    
    if len(j[0]) != 2:
        resource = etree.SubElement(resourcegroup, "Resource", Name=str(j[1]), RBS=str(j[0]), CatalogId=str(uuid.uuid4()))

    
tree = etree.ElementTree(root)
tree.write('stabu_hoofdstukken_2017.xml', encoding="utf-8", xml_declaration=True, pretty_print=True) 


