import os
import sys
# from pandas import pd
import uuid
import xml.dom.minidom

import xml.etree.ElementTree as ET

import pandas as pd


Operating_unit = sys.argv[1]
coveragefile = os.getcwd()+"/output.xlsx"

# Specify the path to your XML file
combined_xml = 'C:/Users/Gokul/Desktop/python/x.xml'
coverage_name = pd.read_excel(coveragefile, sheet_name='Sheet2')

# Parse the XML file
tree = ET.parse(combined_xml)
root = tree.getroot()
# Find all elements named 'FormPattern' using XPath
form_patterns = root.findall('.//FormPattern')
# Loop through each 'FormPattern' element
for s,cov in coverage_name.iterrows():      
  for form_pattern in form_patterns:
      if form_pattern.find('Code') != None and form_pattern.find('Code').text == cov['Code']:
        list = []
        publicId = form_pattern.attrib.get('public-id', 'N/A')
        # Access 'FormPatternOU_Ext' elements
        form_pattern_ou_ext_elements = form_pattern.find('.//FormPatternOU_Ext')
        for form_pattern_ou_ext in form_pattern_ou_ext_elements:
            publicId2 = str(uuid.uuid4())
            availability_ext = form_pattern_ou_ext.find('Availability_Ext').text
            operating_unit_ext = form_pattern_ou_ext.find('OperatingUnit_Ext').text
            if(availability_ext == 'Available'):
              list.append(operating_unit_ext)
            


        
            print(list, end=' ')
            print(cov['Code'], end = ' ') 
            print(cov['Jira#']) # For separating each form pattern's details
            # if Operating_unit in list :
            #   print(form_pattern.find('Code').text )
  #             new_form_pattern_ou_ext = ET.SubElement(form_pattern_ou_ext, "FormPatternOU_Ext")
                        
  #             availability_ext_element = ET.SubElement(new_form_pattern_ou_ext, "Availability_Ext")
  #             availability_ext_element.text = "Available"
                        
  #             operating_unit_ext_element = ET.SubElement(new_form_pattern_ou_ext, "OperatingUnit_Ext")
  #             operating_unit_ext_element.text = Operating_unit
  #             print(new_form_pattern_ou_ext.text)
  #             form_pattern_ou_ext.append(new_form_pattern_ou_ext)
  # tree.write('operating_unit_.xml')