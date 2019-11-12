import os
import ifcopenshell
import lxml 
from lxml import etree 
import uuid
from collections import OrderedDict
import itertools 


def check_file(file_name):
    
    print (file_name)
    print ('check filename')
    
    
def check_position_and_orientation():
    print ('check position and orientation')
    
def check_building_storey():
    print ('check building storey')
    
def check_use_of_entities(file_name):
    print ('check use of entities')
    ifc_file = ifcopenshell.open(file_name)
    products = ifc_file.by_type('IfcProduct')
    
    product_list = []
    product_names_list = []

    for product in products:
        product_list.append(product.is_a())
        product_names_list.append([(product.is_a(),product.Name)])
      
        
    ifcproduct_list = (list(OrderedDict.fromkeys(product_list)))
    product_names_list.sort() 
    ifcproduct_name_list = list(product_names_list for product_names_list, _ in itertools.groupby(product_names_list))
                

    root = etree.Element('exchange')
    doc = etree.SubElement(root, "selectionsets")


    for folder_name in ifcproduct_list:
        viewfolder = etree.SubElement(doc, "viewfolder", name=str(folder_name), guid=str(uuid.uuid4()))

        for product in ifcproduct_name_list:
            if product[0][0] == folder_name:
                
                selectionset = etree.SubElement(viewfolder, "selectionset", name=str(product[0][1]), guid=str(uuid.uuid4()))
                
                findspec = etree.SubElement(selectionset, "findspec", mode="all", disjoint="0")
                
                conditions = etree.SubElement(findspec, "conditions")
                
                condition = etree.SubElement(conditions, "condition", test="equals", flags="10")
                
                category = etree.SubElement(condition, "category")
                
                name = etree.SubElement(category, "name", internal="LcIFCProperty").text = "Item"
                    
                property = etree.SubElement(condition, "property")
                name = etree.SubElement(property, "name", internal="IFCString").text = "Name"
                
                value = etree.SubElement(condition, "value")
                data = etree.SubElement(value, "data", type="wstring").text = str(product[0][1])
                
                locator = etree.SubElement(findspec, "locator").text = "/"
                
           
    schema_location_string = '<exchange xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="http://download.autodesk.com/us/navisworks/schemas/nw-exchange-12.0.xsd">'
    tree = etree.ElementTree(root)
    tree.write('xml_files/ifc_searchsets_navisworks.xml', encoding="utf-8", xml_declaration=True, pretty_print=True)
    
    
    with open('xml_files/ifc_searchsets_navisworks.xml', 'r') as xml_file:
        xml_data = xml_file.readlines()
        
    xml_data[1] = schema_location_string + '\n'
    
    with open('xml_files/ifc_searchsets_navisworks.xml', 'w') as xml_file:
        xml_file.writelines(xml_data)
    
def check_structure_and_naming():
    print ('check structure and naming')
    
def check_classification(file_name):
    print ('check classification')
    
    ifc_file = ifcopenshell.open(file_name)
    products = ifc_file.by_type('IfcProduct')
    

    for product in products:
        print (product)
        
        
    
    
def check_correct_material():
    print ('check correct material')
    
def check_duplicates_and_clashes():
    print ('check duplicates and clashes') 
    
    
def check_pset_loadbearing():
    print ('check pset loadbearing')
    
def check_pset_isexternal():
    print ('pset isexternal')
    
def check_firerating():
    print ('fire rating')
    
def check_project_specific():
    print ('check project specific')


###########################################################
### declare a global variable to be used in the methods ###
###########################################################
file_name='K:\\BIM\\03_BIM_standaarden\\3.3_Navisworks\\Navisworks handleiding\\01_testproject\\IFC modellen Schependomlaan\\Root\\IFC-gebouw.ifc'

check_use_of_entities(file_name)
#check_classification(file_name)   





