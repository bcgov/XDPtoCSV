import os
import csv
import xml.etree.ElementTree as ET

folder_path = 'xdp'
rows = [] # Stores the data for each form field.

# Iterate over all XDP files in the folder and store in the rows array
for file_name in os.listdir(folder_path):
    
    if file_name.endswith('.xdp'):
        
        file_path = os.path.join(folder_path, file_name)
        tree = ET.parse(file_path)
        root = tree.getroot()

        # Locate template and its namespace dynamically        
        template = None
        namespace = ''
        for elem in root.iter():
            if elem.tag.endswith('template'):  # Ignores namespace in search
                template = elem
                if '}' in elem.tag: # get namespace
                    namespace_url = elem.tag.split('}')[0].strip('{')
                    namespace = {'xfa': namespace_url}
                break

        # find all subform elements throughout
        for subform in template.findall('.//xfa:subform', namespace):
            
            # Find all field elements immediately within this subform (depth 1)
            for field in subform.findall('./xfa:field', namespace):
            
                # field name
                field_name = field.get('name', '')

                # get any data binding from the subform
                sub_bind = subform.find('./xfa:bind', namespace)
                sub_bind_text = sub_bind.get('ref', '') if sub_bind is not None else ''

                # get any data binding from the field
                bind = field.find('./xfa:bind', namespace)
                bind_text = bind.get('ref', '') if bind is not None else ''

                # get field label
                caption = field.find('.//xfa:caption//xfa:value//xfa:text', namespace)
                caption_text = caption.text if caption is not None else ''
                
                # get field type
                type_text = ''
                for tag in ['textEdit', 'checkButton', 'choiceList', 'numericEdit', 'button', 'barcode', 'dateTimeEdit']:
                    if field.find(f'.//xfa:{tag}', namespace) is not None:
                        type_text = tag
                        break

                # get Speak
                speak = field.find('.//xfa:assist//xfa:speak', namespace)
                speak_text = speak.text if speak is not None else ''

                # get Items
                items = field.find('.//xfa:items', namespace)
                item_texts = []
                if items is not None:
                    for item in items.findall('.//xfa:text', namespace):
                        if item.text is not None: 
                            item_texts.append(item.text) 

                #item_texts = [item.text for item in items.findall('.//xfa:text', namespace)] if items is not None else []
                
                # group name
                group_name = subform.get('name','')

                # check if subform is a repeater group
                repeater = 'no' if subform.find('xfa:./occur', namespace) is None else 'yes'

                # Add row
                rows.append({
                    'Field': field_name,
                    'Label': caption_text,
                    'Type': type_text,
                    'Speak': speak_text,
                    'Items': ' | '.join(item_texts),
                    'Group Binding': sub_bind_text,
                    'Field Binding': bind_text,
                    'Group Name': group_name,
                    'Is Repeater': repeater,
                    'Forms': file_name
                })

# Sort rows by Field (case insensitive)
rows.sort(key=lambda x: x['Field'].lower())

print('Generating CSV...')
# Write to CSV
with open('form_fields.csv', mode='w', newline='', encoding='utf-8') as csvfile:
    writer = csv.DictWriter(csvfile, fieldnames=['Field', 'Label', 'Type', 'Speak', 'Items', 'Group Binding', 'Field Binding', 'Group Name','Is Repeater','Forms'])
    writer.writeheader()
    writer.writerows(rows)

print('Done!')