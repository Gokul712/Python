import xml.etree.ElementTree as ET

xml_string = '''
<fp>
  <formpro>
    <formpro pid="123">
       <name>tt</name>
       <age>20</age>
    </formpro>
    <formpro pid="1234">
       <name>tt</name>
       <age>21</age>
    </formpro>
  </formpro>
</fp>
'''

# Parse the XML string
root = ET.fromstring(xml_string)

# Find the parent <formpro> element
parent_formpro = root.find('.//formpro')

# Create a new <formpro> element
new_formpro = ET.Element('formpro', {'pid': '12354'})
name = ET.SubElement(new_formpro, 'name')
name.text = 't2t'
age = ET.SubElement(new_formpro, 'age')
age.text = '25'

# Append the new <formpro> element to the parent <formpro> element
parent_formpro.append(new_formpro)

# Convert the modified XML tree back to a string
modified_xml_string = ET.tostring(root, encoding='unicode')

print(modified_xml_string)