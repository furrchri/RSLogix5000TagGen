import os
import re
import pandas as pd
import openpyxl
import xml.etree.ElementTree as ET
from xml.dom import minidom
from datetime import datetime
import csv

system_name_PH = "PH1"
system_name_PR = "PR1"
wm_area = "WM"
inf_area = "INF"

#Use arguments when calling this function or hard code paramaters to meet your tag needs
#This particular function was sourced and modified from Reddy Controls solution for automated tag creation in Logix Designer
#https://github.com/ReddyControls/PLC/blob/main/Tags_RPA
def add_custom_tags_to_csv_interactive(csv_file_path, tag_type_input= None,datatype_input = None, base_name_input = None, start_index_input = None):
    #tag_type = input("Enter the tag type (e.g., TAG, ALIAS): ")
    tag_type = "TAG"
    #scope = input("Enter the scope (e.g., MainProgram): ")
    scope = ""
    #base_name = input("Enter the base name for the tags (e.g., Motor): ")
    base_name = base_name_input
    #datatype = input("Enter the datatype (e.g., BOOL): ")
    datatype = "GCS_Motor_V2_6"
    #attributes = input("Enter the attributes (e.g., '(RADIX := Decimal, Constant := false, ExternalAccess := Read/Write)'): ")
    attributes = '(Constant := false, ExternalAccess := Read/Write)'
    try:
        #num_tags = int(input("Enter the number of tags to add: "))
        num_tags = 1
        if num_tags < 1:
            raise ValueError("Number of tags must be at least 1.")
        #start_index = int(input("Enter the starting index for the tags: ")
        start_index = start_index_input
    except ValueError as e:
        print(f"Invalid input: {e}")
        return

    new_rows = [
        [tag_type, scope, f'{system_name_PH}_{inf_area}_{wm_area}_{base_name}', '', datatype, '', attributes]
        for i in range(start_index, start_index + num_tags)
    ]

    try:
        with open(csv_file_path, 'a', newline='') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerows(new_rows)
        print(f"New rows have been added to {csv_file_path}.")
    except Exception as e:
        print(f"Failed to write to the file: {e}")

#The next two functions are to convert non-decimal representations in a given cell 
# such as "Horsepower" and other relevant columns when reading.
def parse_fraction(frac_str):
    try:
        numerator, denominator = map(float, frac_str.split('/'))
        return numerator / denominator
    except (ValueError, ZeroDivisionError):
        return None

def parse_number(number_str):
    # Handle None/empty
    if not number_str or not str(number_str).strip():
        return None
    
    s = str(number_str).strip()
    
    # Case 1: Simple float ("1.25")
    try:
        return float(s)
    except ValueError:
        pass
    
    # Case 2: Simple fraction ("3/4")
    if '/' in s:
        result = parse_fraction(s)
        if result is not None:
            return result
    
    # Case 3: Mixed number ("1 3/4")
    if ' ' in s and '/' in s:
        whole, fraction = s.split(' ', 1)
        try:
            whole_num = float(whole)
            frac_part = parse_fraction(fraction)
            if frac_part is not None:
                return whole_num + frac_part
        except ValueError:
            pass
    
    # Case 4: Scientific notation ("1.23e-4")
    if 'e' in s.lower():
        try:
            return float(s)
        except ValueError:
            pass
    
    # Return original if no conversion worked
    return None  # or return s to keep as string

# TODO: Combining module files for a future single import of multiple modules. 
# Currently unused
def combine_l5x_files(input_dir, output_file="combined_modules.L5X"):
    """
    Combines multiple L5X module files into one consolidated file
    
    Args:
        input_dir (str): Directory containing individual .L5X files
        output_file (str): Path for the combined output file
    """
    # Create root structure for combined file
    combined_root = ET.Element("RSLogix5000Content",
                             SchemaRevision="1.0",
                             SoftwareRevision="32.00",
                             TargetName="Combined_Modules",
                             TargetType="Controller",
                             ContainsContext="true",
                             Owner="Rockwell Automation, Inc.",
                             ExportOptions="DecoratedData ForceProtectedEncoding AllProjDocTrans")
    

    controller = ET.SubElement(combined_root, "Controller", 
                                Use="Target",
                                ProcessorType="Logix5380",
                                MajorRev="32",  # For Studio 5000 v32
                                MinorRev="11")
    
    modules = ET.SubElement(controller, "Modules")

    # Process each L5X file in directory
    for filename in os.listdir(input_dir):
        if filename.endswith(".L5X"):
            filepath = os.path.join(input_dir, filename)
            
            try:
                tree = ET.parse(filepath)
                module_root = tree.getroot()
                
                # Find the Module element in each file
                for module in module_root.findall(".//Module"):
                    # Import the module node into our combined structure
                    modules.append(module)
                    
                print(f"Added {filename} to combined file")
                
            except Exception as e:
                print(f"Error processing {filename}: {str(e)}")
                continue

    # Generate pretty-printed XML
    rough_string = ET.tostring(combined_root, 'utf-8')
    parsed = minidom.parseString(rough_string)
    pretty_xml = parsed.toprettyxml(indent="    ")
    
    # Write combined file
    with open(output_file, 'w') as f:
        f.write(pretty_xml)
    
    print(f"\nSuccessfully created combined file: {output_file}")
    print(f"Total modules combined: {len(modules.findall('Module'))}")

#This function places all generated modules in a single folder specified by user
def create_output_folder(base_directory):
    folder_name = f"RSLogix_Modules_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    full_path = os.path.join(base_directory, folder_name)
    os.makedirs(full_path, exist_ok=True)
    return full_path


#Creating the sub array (python list) for each row in the sheet
def array_from_sheet(file_path, sheet_name=None):
    # Reads an Excel sheet and returns a list of lists with numeric indexing
    
    # Args:
    #     file_path (str): Path to Excel file
    #     sheet_name (str): Name of sheet to read (uses first sheet if None)
    
    # Returns:
    #     list: A list where each element is a sublist representing a row,
    #           with numeric indexes for columns (0-based)
    
    # Load the workbook and select sheet
    wb = openpyxl.load_workbook(file_path, read_only=True)
    #sheet = wb[sheet_name] if sheet_name else wb.worksheets[1]
    sheet = wb[sheet_name] if sheet_name else wb.worksheets[0]

    print(f"Reading sheet: {sheet.title}")
    module_names = []
    # Determine starting row read from sheet. Modify to better fit your documents structure
    start_row = 0

    
    for row_idx, row in enumerate(sheet.iter_rows(values_only=True), start=start_row):
        # Create subarray for current row with numeric column indexes
        row_array = []
        for col_idx, cell_value in enumerate(row):
            row_array.append(cell_value)
        module_names.append(row_array)
    wb.close()
    return module_names
        

#Creates generic module with manually entered specifications. Currently unused
def create_module_xml(module_name):
    """Creates XML for a single module (using only the module name)"""
    root = ET.Element("RSLogix5000Content", 
                     SchemaRevision="1.0",
                     SoftwareRevision="32.00",
                     TargetName=module_name,
                     TargetType="Module",
                     ContainsContext="true",
                     Owner="Rockwell Automation, Inc.",
                     ExportDate=datetime.now().strftime("%a %b %d %H:%M:%S %Y"),
                     ExportOptions="DecoratedData ForceProtectedEncoding AllProjDocTrans")
    
    controller = ET.SubElement(root, "Controller", Use="Source")
    modules = ET.SubElement(controller, "Modules")
    
    # Default values (modify if needed)
    module = ET.SubElement(modules, "Module",
                         Name=module_name,
                         CatalogNumber="1756-IF8",  # Default example
                         Vendor="1",
                         ProductType="14",
                         ProductCode="58",
                         Major="1",
                         Minor="0",
                         ParentModule="Local",
                         ParentModPortId="1",
                         Inhibited="false",
                         MajorFault="true")
    
    ports = ET.SubElement(module, "Ports")
    port = ET.SubElement(ports, "Port", Id="1", Address="0", Type="ICP", Upstream="false")
    ET.SubElement(port, "Bus", Size="10")
    
    ET.SubElement(module, "Communications",
                CommMethod="7059",
                ElemSize="4",
                ConfigTag=f"{module_name}_Config")
    
    ext_props = ET.SubElement(module, "ExtendedProperties")
    ET.SubElement(ext_props, "ExtendedProperty", Name="Slot", Value="1")  # Default slot
    
    rough_string = ET.tostring(root, 'utf-8')
    parsed = minidom.parseString(rough_string)
    return parsed.toprettyxml(indent="    ")

#This function takes the previously exported template module and assigns a new
#name and IP address and slot number based on previously determined
# hard coded values or passed arguments (ideally) read from the document
def modify_module_template(
    template_path, 
    new_name,
    new_ip=None, 
    new_slot=None,
):
    """
    Imports a module template, modifies key values, and exports new L5X
    
    Args:
        template_path: Path to existing L5X template
        new_name: New module name
        new_ip: New IP address (for network devices)
        new_slot: New slot number
    """
    # Parse template
    tree = ET.parse(template_path)
    root = tree.getroot()
    
    # Find the module element
    module = root.find(".//Module")
    if module is None:
        raise ValueError("No Module found in template")
    
    # Update module attributes
    module.set("Name", f"{system_name_PH}_{inf_area}_{wm_area}_{new_name}_VFD")
    
    # Update IP if specified (for PowerFlex/Ethernet modules)
    if new_ip:
        port = module.find(".//Port[@Type='Ethernet']")
        if port is not None:
            port.set("Address", new_ip)
        
        # Update IP in ExtendedProperties
        for prop in module.findall(".//ExtendedProperty[@Name='IPAddress']"):
            prop.set("Value", new_ip)
    
    # Update slot if specified
    if new_slot is not None:
        for prop in module.findall(".//ExtendedProperty[@Name='Slot']"):
            prop.set("Value", str(new_slot))
    
    # Update ConfigTag to match new name
    comms = module.find("Communications")
    if comms is not None:
        comms.set("ConfigTag", f"{new_name}_Config")

    # # TODO: Update Drive Rating extended properties
    
    # Return formatted XML
    rough_string = ET.tostring(root, encoding='utf-8')
    parsed = minidom.parseString(rough_string)
    return parsed.toprettyxml(indent="    ")



#Reads spreadsheet and places rows into sub arrays.
# the function will row-wise read the sheet with a name specified in a workbook or default to sheet 0,
# then it will read from specified folder and file path arguments)
# an xml file named "template.L5X" already exported as a module from Logix Designer unless changed
# depending on how your data is organized, it will check the columns specified 
# for keywords or a sequence of characters in the string to determine whether modules and tags are generated
def read_excel_and_generate_xml(file_path,base_dir, sheet_name=None):
    """Reads Excel and generates XML using only data from column 2 (index 1)"""
    try:
     
        # Create output folder
        outputs_folder = create_output_folder(base_dir)
        print(f"Created output folder: {outputs_folder}")
        template_path_input = os.path.join(base_dir, "template.L5X")

        module_data = array_from_sheet(file_path)
        #print("first entry is: ", module_data[6])
        #combined_modules = []
        success_count = 0
        #output_path = os.path.join(output_folder, file_name) 
        for row in module_data:
                cell_data = row
                equipment_title = str(cell_data[1])
                drive_type = str(cell_data[5])
                # if cell_data[4] != None:
                #      drive_hp  = parse_number(cell_data[4])

                motor_string = "M1"
                powerflex_string  = "AB PF525"
                print(equipment_title)

                match_1 = re.search(motor_string, equipment_title)
                match_1a = re.search(powerflex_string, drive_type)
                print(match_1)
                csv_path = r'Test_Gen.csv'
                if  match_1a:
                    xml_content = modify_module_template(template_path = template_path_input,new_name= equipment_title)
                    print("made vfd")
                else: 
                    #xml_content = create_module_xml(equipment_title)
                    print("no match")
                    xml_content = None
                if xml_content:
                    file_name = f"{equipment_title}_VFD.L5X"
                    output_file = os.path.join(outputs_folder, file_name)
                    add_custom_tags_to_csv_interactive(csv_path,base_name_input=equipment_title, start_index_input=success_count)
                    with open(output_file, 'w') as f:
                        f.write(xml_content)
                    print(f"Generated {output_file}")
                    success_count += 1
        #combine_l5x_files(input_dir=outputs,output_file="combined_project.L5X")      
        #print(f"Generated {output_path}")
        print(f"\nSuccessfully generated {success_count} module configuration files")
 
    except Exception as e:
        print(f"ERROR: {str(e)}")

if __name__ == "__main__":
    #excel_file = "modules_configuration.xlsx"
    excel_file = "EXAMPLE.xlsx"
    #When not specifying output directory, use empty string to output
    #to the folder in which the program executes
    output_folder_directory = ""
    read_excel_and_generate_xml(excel_file,output_folder_directory)
    
    
    
