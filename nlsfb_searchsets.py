from lxml import etree
import lxml.etree
import lxml.builder 
from xlrd import open_workbook 

import sys
import uuid 
import os
import csv
import re 

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

def get_first_two_number_classsification():
    
    first_two_number_list=[]
    
    #tokkie code
    
    for i in sheet.col(4):
        if len(i.value) > 0:
            
            if str(i.value).startswith('0'):
                pass  
            else:
                if i.value != "Class-codenotatie":
                    
                    first_two_number_list.append(i.value[0:2])
    
    return first_two_number_list
        
        
        
#for x in get_first_two_number_classsification():
#    print x 
    

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
    
    ######################################################
    ########## REVIT ASSEMBLY CODE CONDITION #############
    ######################################################
    condition = etree.SubElement(conditions, "condition", test="equals", flags="10")
    
    category = etree.SubElement(condition, "category")
    
    name = etree.SubElement(category, "name", internal="LcRevitData_Type").text = "Revit Type"

    property = etree.SubElement(condition, "property")
    name = etree.SubElement(property, "name", internal="lcldrevit_parameter_-1002500").text = "Assembly Code"
    
    value = etree.SubElement(condition, "value")
    data = etree.SubElement(value, "data", type="wstring").text = str(items) 
    

    
    #############################################################
    ########## IFCB NLSFB CLASSIFICATION CONDITION ##############
    #############################################################
    
    condition_ifc = etree.SubElement(conditions, "condition", test="equals", flags="74")
    
    category_ifc = etree.SubElement(condition_ifc, "category")
    name_ifc = etree.SubElement(category_ifc, "name", internal="lcldrevit_parameter_2476_classification_tab").text  = "Classification: NL/SfB (4 cijfers)" 
    
    property_ifc = etree.SubElement(condition_ifc, "property")
    name_ifc = etree.SubElement(property_ifc, "name", internal="lcldrevit_parameter_3118_reference").text = "Reference"
    
    value_ifc = etree.SubElement(condition_ifc, "value")
    data_ifc = etree.SubElement(value_ifc, "data", type="wstring").text = str(items) 
    
    
    
    
    
    ###############################################################
    ########## IFC NLSFB CLASSIFICATION CONDITION 2005 ############
    ###############################################################
    
    condition_ifc_2005 = etree.SubElement(conditions, "condition", test="equals", flags="74")
    
    category_ifc_2005 = etree.SubElement(condition_ifc_2005, "category")
    name_ifc_2005 = etree.SubElement(category_ifc_2005, "name", internal="lcldrevit_parameter_2476_classification_tab").text  = "Classification: NL/SfB (4 cijfers) 2005" 
    
    property_ifc_2005 = etree.SubElement(condition_ifc_2005, "property")
    name_ifc_2005 = etree.SubElement(property_ifc_2005, "name", internal="lcldrevit_parameter_3118_reference").text = "Reference"
    
    value_ifc_2005 = etree.SubElement(condition_ifc_2005, "value")
    data_ifc_2005 = etree.SubElement(value_ifc_2005, "data", type="wstring").text = str(items) 
    
    
    
    
    
    ###############################################################
    ############### ARCHICAD ITEM REFERENCE #######################
    ###############################################################
    
    condition_archicad = etree.SubElement(conditions, "condition", test="equals", flags="74")
    
    category_archicad = etree.SubElement(condition_archicad, "category")
    name_archicad = etree.SubElement(category_archicad, "name", internal="lcldrevit_parameter_2476_classification_tab").text  = "Classification:NL/SfB (4 cijfers) 2005" 
    
    property_archicad = etree.SubElement(condition_archicad, "property")
    name_archicad = etree.SubElement(property_archicad, "name", internal="lcldrevit_parameter_3118_reference").text = "Reference"
    
    value_archicad = etree.SubElement(condition_archicad, "value")
    data_archicad = etree.SubElement(value_archicad, "data", type="wstring").text =   str(items) 
    
    
    
    
    
    
    
    #Deze nog afmaken aan de hand van voorbeeld IFC bestand de Loef 
    ###############################################################
    ############### UNIFORMAT CLASSIFICATION ######################
    ###############################################################
    
    condition_uniformat_classfication = etree.SubElement(conditions, "condition", test="equals", flags="74")
    
    category_uniformat_classification = etree.SubElement(condition_uniformat_classfication, "category")
    name_uniformat_classification = etree.SubElement(category_uniformat_classification, "name", internal="lcldrevit_parameter_2476_classification_tab").text  = "Classification: Uniformat" 
    
    property_uniformat_classification = etree.SubElement(condition_uniformat_classfication, "property")
    name_uniformat_classification = etree.SubElement(property_uniformat_classification, "name", internal="lcldrevit_parameter_3118_reference").text = "Reference"
    
    value_uniformat_classification = etree.SubElement(condition_uniformat_classfication, "value")
    data_uniformat_classification = etree.SubElement(value_uniformat_classification , "data", type="wstring").text =   str(items) 
    

    
    #CLASSIFICATION: UNIFORMAT CLASSIFCATION
    
    ###########################################################################
    ############# GEEN IDEE WAT DIT DOET, MAAR IS BLIJKBAAR NODIG #############
    ###########################################################################
    
    locator = etree.SubElement(findspec, "locator").text = "/"

    


#STAP 1: NEGEN HOOFDMAPEN MAKEN
#STAP 2: SUBMAPPEN MAKEN
#STAP 3: SEARCHSETS MAKEN
#STAP 4: IFC SEARCHSETS TOEVOEGEN

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
tree.write('nlsfb_check_neverworks.xml', encoding="utf-8", xml_declaration=True, pretty_print=True)  





"""
# Declare filenames
navis_works_xml_file = 'nlsfb_check_neverworks.xml'
tree = etree.ElementTree(root)
tree.write(navis_works_xml_file, encoding="utf-8", xml_declaration=True, pretty_print=True)           
 
 

# Open original file
tree = etree.parse(navis_works_xml_file)
root = tree.getroot()

# Edit file
for tag_exchange in (root.iter('exchange')):
    tag_exchange.set('xmlns:xsi', 'http://www.w3.org/2001/XMLSchema-instance') 
    tag_exchange.set('xsi:noNamespaceSchemaLocation', 'http://download.autodesk.com/us/navisworks/schemas/nw-exchange-12.0.xsd')
 

 
tree.write(navis_works_xml_file, encoding="utf-8", xml_declaration=True)

    
"""    


    
         
                






