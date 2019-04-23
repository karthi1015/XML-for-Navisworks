import ifcopenshell
from lxml import etree
import uuid 
import os 

######################################################################
#################### GET TYPE DATA FROM IFC  #########################
######################################################################

def get_type(filename):
    

    ifc_file = ifcopenshell.open(filename)
    
    product_type = ifc_file.by_type('IfcTypeProduct')
    
    product_types_list = []

    for product_types in product_type:
        product_types_list.append([product_types.is_a(), product_types.Name])

   
    return product_types_list 




def write_to_xml(file_name):
    ######################################################################
    ########################### WRITE XML  ###############################
    ###################################################################### 
    
    #ROOT EXCHANGE
    root = etree.Element("exchange")
    
    #SELECTION SETS
    doc = etree.SubElement(root, "selectionsets")
    
    
    ######################################################################
    ##################### REMOVE DUPLICATES FROM LIST ####################
    #####################################################################
    some_type_list = get_type(filename=file_name)
    
    y_list=[]          
    for i in some_type_list:
        y_list.append(i[0])

    unique_type_list = list(dict.fromkeys(y_list))    
        
    
    ##############################################################################
    ############### CREATE XML FOLDER STRUCTURE AND SEARCHSETS ###################
    ##############################################################################
    schema_location_string = '<exchange xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="http://download.autodesk.com/us/navisworks/schemas/nw-exchange-12.0.xsd">' 
    
    for i in unique_type_list:
        viewfolder = etree.SubElement(doc, "viewfolder", name=str(i), guid=str(uuid.uuid4()))
    
        for j in some_type_list:
            if i == j[0]:
                
                selectionset = etree.SubElement(viewfolder, "selectionset", name=j[1], guid=str(uuid.uuid4()))
                
                findspec = etree.SubElement(selectionset, "findspec", mode="all", disjoint="0")
                
                conditions = etree.SubElement(findspec, "conditions")
                
                condition = etree.SubElement(conditions, "condition", test="equals", flags="10")
                
                category = etree.SubElement(condition, "category")
                
                name = etree.SubElement(category, "name", internal="LcIFCProperty").text = str(j[0]).upper()
                    
                property = etree.SubElement(condition, "property")
                name = etree.SubElement(property, "name", internal="IFCString").text = "NAME"
                
                value = etree.SubElement(condition, "value")
                data = etree.SubElement(value, "data", type="wstring").text = j[1]
                
                locator = etree.SubElement(findspec, "locator").text = "/"
                

            
    tree = etree.ElementTree(root)
    
    head, tail = os.path.split(file_name)
    
    xml_file_ifc = tail.replace('.ifc','')
    
    tree.write('ifc_searchsets_navisworks_' + str(xml_file_ifc) + '.xml', encoding="utf-8", xml_declaration=True, pretty_print=True)



 
    ##########################################################################################################
    ####################### READ EXISTING XML FILE TO CREATE EXCHANGE TAG ON LINE 2 ##########################  
    ##########################################################################################################

    with open('ifc_searchsets_navisworks_' + str(xml_file_ifc) + '.xml', 'r') as file:
        data = file.readlines()
    
    
    data[1] = schema_location_string + '\n'
    
    
    with open('xml_navisworks/XML_ifc_searchsets_navisworks_' + str(xml_file_ifc) + '.xml', 'w') as file:
        file.writelines(data)
        
        print 'script completed'



if __name__ == '__main__':
 
    write_to_xml(file_name)
        
     
        
    
    