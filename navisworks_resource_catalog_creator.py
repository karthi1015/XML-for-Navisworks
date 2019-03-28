#C.C.J. Claus 2019, Amersfoort, the Netherlands.
#A python script to create STABU XML files for Navisworks Manage 2019

from lxml import etree
import lxml.etree
import lxml.builder 
import xlrd
import uuid 
import os
import csv
import re 


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
#resource = "Resource"
#variable_collection = "VariableCollection"
#variable = "Variable"


root = etree.Element(take_off)
doc = etree.SubElement(root, catalog)

#<?xml version="1.0" encoding="utf-8"?>
#<Takeoff xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns="http://download.autodesk.com/us/navisworks/schemas/nw-TakeoffCatalog-10.0.xsd">

for j in stabu_list:
    if len(j[0]) == 2:
        
     
        resourcegroup = etree.SubElement(doc, resource_group, Name=str(j[1]), RBS=str(j[0]), CatalogId=str(uuid.uuid4())) 
    
    if len(j[0]) != 2:
        resource = etree.SubElement(resourcegroup, "Resource", Name=str(j[1]), RBS=str(j[0]), CatalogId=str(uuid.uuid4()))
        
        #for i in j:
        #    variable_collection = etree.SubElement(resource, "VariableCollection")
            
            
            #etree.SubElement(variable_collection, "Variable", Name="Length", Formula="ModelLength", Units="Meter")
            #etree.SubElement(variable_collection, "Variable", Name="Width", Formula="ModelWidth", Units="Meter")
            #etree.SubElement(variable_collection, "Variable", Name="Thickness", Formula="ModelThickness", Units="Meter")
            #etree.SubElement(variable_collection, "Variable", Name="Height", Formula="ModelHeight", Units="Meter")
            #etree.SubElement(variable_collection, "Variable", Name="Perimeter", Formula="ModelPerimeter", Units="Meter")
            #etree.SubElement(variable_collection, "Variable", Name="Area", Formula="ModelArea", Units="Square Meter")
            #etree.SubElement(variable_collection, "Variable", Name="Volume", Formula="ModelVolume", Units="Cubic Meter")
            #etree.SubElement(variable_collection, "Variable", Name="Weight", Formula="ModelWeight", Units="Kilogram")
            #etree.SubElement(variable_collection, "Variable", Name="Count", Formula="1")

            
            
            
        
        


    
tree = etree.ElementTree(root)
tree.write('stabu_hoofdstukken_2017.xml', encoding="utf-8", xml_declaration=True, pretty_print=True) 
























#stabu_bestandslocatie = "stabu_hoofdstukken_2017-1.pdf"
#pdf_file = open(stabu_bestandslocatie, 'rb')
#fileReader = PyPDF2.PdfFileReader(pdf_file)
#print fileReader 













funderingspalen_en_damwanden = {    "20.00":"ALGEMEEN",
                                    "20.13":"BEPROEVEN, CONTROLEREN",
                                    "20.21":"DEMONTEREN/VERWIJDEREN BESTAAND WERK",
                                    "20.22":"REINIGEN BESTAAND WERK",
                                    "20.23":"VOORBEHANDELEN ONDERGROND BESTAAND WERK",
                                    "20.24":"INJECTEREN BESTAAND WERK",
                                    "20.25":"HERSTELLEN BESTAAND WERK",
                                    "20.26":"AANPASSEN BESTAAND WERK",
                                    "20.27":"HERPLAATSEN BESTAAND WERK",
                                    "20.28":"ONDERHOUDEN BESTAAND WERK",
                                    "20.31":"PAALFUNDERINGEN VAN VOORAF VERVAARDIGDE PALEN",
                                    "20.32":"IN DE GROND GEVORMDE PALEN",
                                    "20.37":"BEWERKINGEN FUNDERINGSPALEN",
                                    "20.41":"DAMWANDEN VAN VOORAF VERVAARDIGDE DAMWANDPROFIELEN",
                                    "20.42":"IN DE GROND GEVORMDE DAMWANDEN",
                                    "20.43":"DAMWANDVERANKERINGEN",
                                    "20.47":"BEWERKINGEN DAMWANDEN"
                                    }





#for k, v in funderingspalen_en_damwanden.iteritems():
#    print k,v 
