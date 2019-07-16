from lxml import etree
import lxml.etree
import lxml.builder 
from xlrd import open_workbook 
import fileinput

import sys
import uuid 
import os
import csv
import re 

path_to_file = "C:\\Users\\CClaus\\eclipse-workspace\\BasicBIMChecker_for_Navisworks\\xml_files\\ILS_nlsfb_navisworks.xml"  
exchange = '<exchange xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="http://download.autodesk.com/us/navisworks/schemas/nw-exchange-12.0.xsd">' 

#INITALISATION OF READING EXCEL FILE
book = open_workbook('NL-SfB_Tabel_0-4.xlsx') 
sheet = book.sheet_by_name("NL-SfB_Tabel 1")

#INITALISATION OF WRITING XML FILE
exchange="exchange"              
root = etree.Element(exchange)  
doc = etree.SubElement(root, "selectionsets")

nl_sfb_folder_list = ["0_PROJECT",
                      "1_FUNDERINGEN",
                      "2_RUWBOUW",
                      "3_AFBOUW",
                      "4_AFWERKINGEN",
                      "5_MECHANISCHE_INSTALLATIES",
                      "6_ELECTRISCHE_INSTALLATIES",
                      "7_VASTE_INRICHTINGEN",
                      "8_LOSSE_INVENTARIS",
                      "9_TERREIN"
                      ]

nlsfb_conditions_list = [["Classification: NL/SfB (4 cijfers)", "Reference"],
                         ["Classification: NL/SfB (4 cijfers) 2005", "Reference"],
                         ["Classification:NL/SfB (4 cijfers) 2005", "Reference"],
                         ["Classification: Uniformat","Reference"],
                         ["Classification: NL/SfB (4 cijfers)", "ITEMREFERENCE"],
                         ["Classification: NL/SfB (4 cijfers) 2005", "ITEMREFERENCE"],
                         ["Classification:NL/SfB (4 cijfers) 2005", "ITEMREFERENCE"],
                         ["Classification:NL/SfB (4 cijfers)", "ITEMREFERENCE"],
                         ["Classification: Uniformat", "ITEMREFERENCE"],
                         ["Revit Type", "Assembly Code"]
                         ]



def get_first_two_number_classsification():
    
    first_two_number_list=[]
    
    for i in sheet.col(4):
        if len(i.value) > 0:
            
            if str(i.value).startswith('0'):
                pass  
            else:
                if i.value != "Class-codenotatie":
                    
                    first_two_number_list.append(i.value[0:2])
    
    return first_two_number_list
        

def get_classificiation_number_description():
    
    nl_sfb_list=[]


    for i, j in enumerate(sheet.col(5)):
        if len((j.value.split(';'))) >1:
            nl_sfb_list.append([sheet.col(4)[i].value, j.value.split(';')[1]])
           
    return nl_sfb_list


def get_subfolder():
    
    nl_sfb_subfolder_list = []
    
    for i, v in enumerate(sheet.col(4)):
        if len(v.value) == 4:
            if v.value.endswith('0'):
                nl_sfb_subfolder_list.append([v.value, sheet.col(5)[i].value])
 
    return nl_sfb_subfolder_list


def write_first_two(items, description):
    print items, description 
        
    
def write_xml_navisworks(items, description):
    

    selectionset = etree.SubElement(viewfolder, "selectionset", name=(str(items) + str(description)), guid=str(uuid.uuid4()))
    findspec = etree.SubElement(selectionset, "findspec", mode="all", disjoint="0")
    conditions = etree.SubElement(findspec, "conditions")
    
    
    for condition in  nlsfb_conditions_list:

        #############################################################
        ########## IFCB NLSFB CLASSIFICATION CONDITION ##############
        #############################################################
        condition_ifc = etree.SubElement(conditions, "condition", test="equals", flags="74")
        
        category_ifc = etree.SubElement(condition_ifc, "category")
        name_ifc = etree.SubElement(category_ifc, "name", internal="lcldrevit_parameter_2476_classification_tab").text  = str(condition[0])
        
        property_ifc = etree.SubElement(condition_ifc, "property")
        name_ifc = etree.SubElement(property_ifc, "name", internal="lcldrevit_parameter_3118_reference").text = str(condition[1])
        
        value_ifc = etree.SubElement(condition_ifc, "value")
        data_ifc = etree.SubElement(value_ifc, "data", type="wstring").text = str(items) 
        
    
     
    locator = etree.SubElement(findspec, "locator").text = "/"

    

#CREATE MAIN FOLDERS
for nlsfb_folder in nl_sfb_folder_list:
    mainfolder = etree.SubElement(doc, "viewfolder", name=str(nlsfb_folder), guid=str(uuid.uuid4()))

    #CREATE SUB FOLDER
    for sub_folder in get_subfolder():
        if sub_folder[0].startswith(str(nlsfb_folder[0])):

            viewfolder = etree.SubElement(mainfolder, "viewfolder", name=str(sub_folder[0] + " " + sub_folder[1]), guid=str(uuid.uuid4()))
            first_two_view_folder =  sub_folder[0][0] + sub_folder[0][1]
            
            
            for i in get_classificiation_number_description():
                
                first_two = str(i[0][0]) + str(i[0][1])
                
                if first_two_view_folder.startswith(first_two):

                    write_xml_navisworks(items=str(i[0]), description=str(i[1]))
                    
  
  
tree = etree.ElementTree(root)
tree.write('xml_files\\ILS_nlsfb_navisworks.xml', encoding="utf-8", xml_declaration=True, pretty_print=True)  


def replace_all(xml_file, search_expression, replace_expression):
    for line in fileinput.input(xml_file, inplace=1):
        if search_expression in line:
            line = line.replace(search_expression, replace_expression)
            
        sys.stdout.write(line)
        
replace_all(path_to_file, '<exchange>','<exchange xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="http://download.autodesk.com/us/navisworks/schemas/nw-exchange-12.0.xsd">' )





    
         
                





