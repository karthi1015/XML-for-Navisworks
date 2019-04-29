"""
MIT License

Copyright (c) 2018 C. Claus 

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
"""

    
    
import ifcopenshell
from lxml import etree
import uuid 
import os 



def get_software(filename):
    print 'get software'

    ifcfile = ifcopenshell.open(filename)
    application = ifcfile.by_type('IfcApplication')
    
    software_list = []
    
   
    for software in application:
        print software 
    
     
        if 'Autodesk Revit' in software.ApplicationFullName:
            print 'Revit gevonden'
            software_list.append(software.ApplicationFullName)
            
        if 'ArchiCAD' in software.ApplicationFullName:
            print 'ARCHICAD gevonden'
            software_list.append(software.ApplicationFullName)
            
        if 'Tekla' in software.ApplicationFullName:
            software_list.append(software.ApplicationFullName)    
            
            
            
    print 'hier', software_list     
    return software_list  
        
        
def get_revit_description(filename):
    print 'get revit description'
    
    product_types_list = []
    
    ifcfile = ifcopenshell.open(filename) 
    products = ifcfile.by_type('IfcProduct')
    
    for product in products:
        if (product.is_a() != 'IfcBuilding' and 
            product.is_a() !=  'IfcBuildingStorey' and 
            product.is_a() != 'IfcOpeningElement' and 
            product.is_a() != 'IfcAnnotation'):
            
            if product.ObjectType is not None:
                    #print product.is_a(), product.Name, product.ObjectType
                    product_types_list.append([product.is_a(), product.ObjectType])
            
    return product_types_list


def get_archicad_description(filename):
    print 'get archicad description'
    
    product_types_list = []
    
    ifcfile = ifcopenshell.open(filename)
    products = ifcfile.by_type('IfcProduct')
    
    for product in products:
        if (product.is_a() != 'IfcBuilding' and 
            product.is_a() !=  'IfcBuildingStorey' and 
            product.is_a() != 'IfcOpeningElement' and 
            product.is_a() != 'IfcAnnotation' and
            product.is_a() != 'IfcVirtualElement'):
            if product.ObjectType is None:
                    #print product.is_a(), product.Name
                    product_types_list.append([product.is_a(), product.Name])
                    
    return product_types_list  

def write_to_xml_revit(file_name):
     
    #ROOT EXCHANGE
    root = etree.Element("exchange")
    
    #SELECTION SETS
    doc = etree.SubElement(root, "selectionsets")
    
    ######################################################################
    ##################### REMOVE DUPLICATES FROM LIST ####################
    #####################################################################
    some_type_list = get_revit_description(filename=file_name)
    
    y_list=[]          
    for i in some_type_list:
        y_list.append(i[0])

    unique_type_list = list(dict.fromkeys(y_list)) 
 
    ##############################################################################
    ############### CREATE XML FOLDER STRUCTURE AND SEARCHSETS ###################
    ##############################################################################

    a = []
    for k in some_type_list:
        a.append(k)
        
    b_set = set(map(tuple, a))    
    
    b = map(list, b_set)
    

    schema_location_string = '<exchange xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="http://download.autodesk.com/us/navisworks/schemas/nw-exchange-12.0.xsd">' 
    
    for i in unique_type_list:
        viewfolder = etree.SubElement(doc, "viewfolder", name=str(i), guid=str(uuid.uuid4()))
        for j in b:
            if i == j[0]:
                if j[1] is not None:
                    
                    selectionset = etree.SubElement(viewfolder, "selectionset", name=j[1].encode('ascii', 'ignore').decode('ascii'), guid=str(uuid.uuid4()))
                    findspec = etree.SubElement(selectionset, "findspec", mode="all", disjoint="0")
                    
                    conditions = etree.SubElement(findspec, "conditions")
                    
                    
                    ##################################################################
                    ################# REVIT OBJECTTYPE CONDITION #####################
                    ##################################################################
                    condition = etree.SubElement(conditions, "condition", test="equals", flags="10")
                    
                    category = etree.SubElement(condition, "category")
                    name = etree.SubElement(category, "name", internal="LcIFCProperty").text = 'IFC'#j[1].encode('ascii', 'ignore').decode('ascii')
                        
                    
                    property = etree.SubElement(condition, "property")
                    name = etree.SubElement(property, "name", internal="IFCString").text = "OBJECTTYPE" #'NAME'
                    
                    
                    value = etree.SubElement(condition, "value")
                    data = etree.SubElement(value, "data", type="wstring").text = j[1].encode('ascii', 'ignore').decode('ascii')
                    
                    locator = etree.SubElement(findspec, "locator").text = "/"
                    
                    
    tree = etree.ElementTree(root)
    head, tail = os.path.split(file_name)
    xml_file_ifc = tail.replace('.ifc','')
    tree.write('ifc_searchsets_navisworks_' + str(xml_file_ifc) + '.xml', encoding="utf-8", xml_declaration=True, pretty_print=True)

    folder_path = head.replace('/','//')
    ##########################################################################################################
    ####################### READ EXISTING XML FILE TO CREATE EXCHANGE TAG ON LINE 2 ##########################  
    ##########################################################################################################

    with open('ifc_searchsets_navisworks_' + str(xml_file_ifc) + '.xml', 'r') as file:
        data = file.readlines()
    
    
    data[1] = schema_location_string + '\n'
    
    
    with open(str(head) + '//XML_ifc_searchsets_navisworks_' + str(xml_file_ifc) + '.xml', 'w') as file:
        file.writelines(data)
        
        print 'script completed'
        
        path = os.path.realpath(folder_path)   
        os.startfile(path) 
                                 
    return
 
def write_to_xml_archicad(file_name):
     
    #ROOT EXCHANGE
    root = etree.Element("exchange")
    
    #SELECTION SETS
    doc = etree.SubElement(root, "selectionsets")
    
    
    ######################################################################
    ##################### REMOVE DUPLICATES FROM LIST ####################
    #####################################################################
    some_type_list = get_archicad_description(filename=file_name)
    
    y_list=[]          
    for i in some_type_list:
        y_list.append(i[0])

    unique_type_list = list(dict.fromkeys(y_list)) 
 


    
    ##############################################################################
    ############### CREATE XML FOLDER STRUCTURE AND SEARCHSETS ###################
    ##############################################################################

    a = []
    for k in some_type_list:
        a.append(k)
        
    b_set = set(map(tuple, a))    
    
    b = map(list, b_set)
    
    
    schema_location_string = '<exchange xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="http://download.autodesk.com/us/navisworks/schemas/nw-exchange-12.0.xsd">' 
    
    for i in unique_type_list:
        viewfolder = etree.SubElement(doc, "viewfolder", name=str(i), guid=str(uuid.uuid4()))

        for j in b:
   
            if i == j[0]:
                
                if j[1] is not None:
                    

                    selectionset = etree.SubElement(viewfolder, "selectionset", name=j[1].encode('ascii', 'ignore').decode('ascii'), guid=str(uuid.uuid4()))
                    
                    
                    findspec = etree.SubElement(selectionset, "findspec", mode="all", disjoint="0")
                    
                    conditions = etree.SubElement(findspec, "conditions")
                    
                    
                    condition_archicad = etree.SubElement(conditions, "condition", test="equals", flags="10")
                    
                    category_archicad = etree.SubElement(condition_archicad, "category")
                    name_archicad = etree.SubElement(category_archicad, "name", internal="LcIFCProperty").text = 'IFC'
                    
                    property_archicad = etree.SubElement(condition_archicad, "property")
                    name = etree.SubElement(property_archicad, "name", internal="IFCString").text = 'NAME'
                    
                    value_archicad = etree.SubElement(condition_archicad, "value")
                    data_archicad = etree.SubElement(value_archicad, "data", type="wstring").text = j[1].encode('ascii', 'ignore').decode('ascii')
                    
                    
                    
                    
                    locator = etree.SubElement(findspec, "locator").text = "/"
                    
                 
            
    
            
    tree = etree.ElementTree(root)
    
    head, tail = os.path.split(file_name)
    
    xml_file_ifc = tail.replace('.ifc','')
    
    tree.write('ifc_searchsets_navisworks_' + str(xml_file_ifc) + '.xml', encoding="utf-8", xml_declaration=True, pretty_print=True)

    #print 'HEAD', head
    
    folder_path = head.replace('/','//')
    
 
    ##########################################################################################################
    ####################### READ EXISTING XML FILE TO CREATE EXCHANGE TAG ON LINE 2 ##########################  
    ##########################################################################################################

    with open('ifc_searchsets_navisworks_' + str(xml_file_ifc) + '.xml', 'r') as file:
        data = file.readlines()
    
    
    data[1] = schema_location_string + '\n'
    
    
    with open(str(head) + '//XML_ifc_searchsets_navisworks_' + str(xml_file_ifc) + '.xml', 'w') as file:
        file.writelines(data)
        
        print 'script completed'
        #os.system(head.replace('/','//')) 
     
        path = os.path.realpath(folder_path)   
        os.startfile(path)   
           
    return



if __name__ == '__main__':
    if 'ArchiCAD' in get_software(filename=file_name)[0]:
        print 'archicad method'
        
        get_archicad_description(filename=file_name)
        write_to_xml_archicad(file_name=file_name)
        
    
        
    if 'Revit' in get_software(filename=file_name)[0]:
        print 'revit method'
        
        get_revit_description(filename=file_name)
        write_to_xml_revit(file_name=file_name)
    
    
    if 'Tekla' in get_software(filename=file_name)[0]:
        print 'tekla method'
        write_to_xml_revit(file_name)
        
        

