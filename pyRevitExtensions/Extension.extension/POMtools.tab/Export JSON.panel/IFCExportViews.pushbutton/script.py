# -*- coding: utf-8 -*-
"""Export IFC to JSON"""
from pyrevit import forms, script
import json
from collections import defaultdict
import os
import sys
import codecs

__title__ = 'Export\nIFC to JSON'
__author__ = 'Yazid'
__doc__ = 'Exports IFC files data to JSON with element type selection'

logger = script.get_logger()
output = script.get_output()

def read_ifc_file(file_path):
    """Read IFC file and extract elements by type."""
    try:
        output.print_md("Analyzing file: **{}**".format(os.path.basename(file_path)))
        
        # Dictionary to store element counts by type
        element_types = defaultdict(int)
        elements_by_type = defaultdict(list)
        
        # Basic file info
        file_info = {
            "file_name": os.path.basename(file_path),
            "file_path": file_path,
            "file_size": os.path.getsize(file_path)
        }
        
        # Parse IFC file line by line - IronPython compatible version
        with open(file_path, 'r') as f:
            line_num = 0
            for line in f:
                try:
                    # Try to decode line if needed
                    decoded_line = line
                    if isinstance(line, bytes):
                        decoded_line = line.decode('utf-8', errors='ignore')
                    
                    line_num += 1
                    # Only process lines that define IFC elements
                    if '=IFC' in decoded_line and decoded_line.startswith('#'):
                        # Extract entity ID and type
                        parts = decoded_line.split('=')
                        entity_id = parts[0].strip('#').strip()
                        entity_type = parts[1].split('(')[0].strip()
                        
                        # Store basic element info
                        element_info = {
                            "id": entity_id,
                            "type": entity_type,
                            "line": line_num,
                            "raw_data": decoded_line.strip()
                        }
                        
                        # Add to collections
                        element_types[entity_type] += 1
                        elements_by_type[entity_type].append(element_info)
                except:
                    continue
        
        return {
            "file_info": file_info,
            "element_types": dict(element_types),
            "elements_by_type": elements_by_type
        }
    except Exception as ex:
        logger.error("Error reading IFC file {}: {}".format(file_path, str(ex)))
        return None

def extract_elements_from_ifc(ifc_data, selected_types):
    """Extract selected element types from IFC data."""
    extracted_data = {
        "file_info": ifc_data["file_info"],
        "selected_types": selected_types,
        "elements": {}
    }
    
    # Add elements of selected types
    for element_type in selected_types:
        if element_type in ifc_data["elements_by_type"]:
            extracted_data["elements"][element_type] = ifc_data["elements_by_type"][element_type]
    
    return extracted_data

def export_ifc_to_json(ifc_data, export_folder, file_name):
    """Export IFC data to JSON file."""
    try:
        if not os.path.exists(export_folder):
            os.makedirs(export_folder)
        
        # Create filename for JSON
        safe_filename = file_name + ".json"
        filepath = os.path.join(export_folder, safe_filename)
        
        # Export to JSON
        with codecs.open(filepath, 'w', encoding='utf-8') as f:
            json_str = json.dumps(ifc_data, indent=4, ensure_ascii=False)
            f.write(json_str)
        
        return filepath
    except Exception as ex:
        logger.error("Error exporting IFC data to JSON: {}".format(str(ex)))
        return None

def process_ifc_file(file_path, export_folder):
    """Process a single IFC file and export selected element types."""
    # Read IFC file
    ifc_data = read_ifc_file(file_path)
    
    if not ifc_data or not ifc_data["element_types"]:
        output.print_md("No valid elements found in IFC file: {}".format(os.path.basename(file_path)))
        return []
    
    # Create options for element type selection
    element_types = ifc_data["element_types"]
    type_options = sorted(["{}  ({} elements)".format(t, element_types[t]) for t in element_types])
    
    # Show element type selection dialog
    selected_options = forms.SelectFromList.show(
        type_options,
        title='Select Element Types to Export from: {}'.format(os.path.basename(file_path)),
        button_name='Export to JSON',
        multiselect=True,
        width=500,
        height=600
    )
    
    if not selected_options:
        return []
    
    # Extract selected types from options
    selected_types = [opt.split('  (')[0] for opt in selected_options]
    
    # Create file name for export
    file_name = os.path.splitext(os.path.basename(file_path))[0]
    
    # Extract and export data
    exported_data = extract_elements_from_ifc(ifc_data, selected_types)
    json_file = export_ifc_to_json(exported_data, export_folder, file_name)
    
    return [json_file] if json_file else []

def main():
    try:
        # Let user select IFC files
        ifc_file_paths = forms.pick_file(
            file_ext='ifc',
            multi_file=True,
            title='Sélectionner les fichiers IFC à exporter'
        )
        
        if not ifc_file_paths:
            return
        
        # Let user select export folder
        export_folder = forms.pick_folder(
            title='Sélectionner un dossier de destination pour l\'export'
        )
        
        if not export_folder:
            return
        
        exported_files = []
        total_files = len(ifc_file_paths)
        
        with forms.ProgressBar() as pb:
            for idx, file_path in enumerate(ifc_file_paths):
                pb.update_progress(idx, max_value=total_files)
                output.print_md("Processing file: **{}** ({}/{})".format(
                    os.path.basename(file_path), idx + 1, total_files))
                
                # Process IFC file
                exported = process_ifc_file(file_path, export_folder)
                exported_files.extend(exported)
        
        if exported_files:
            message = 'Export completed successfully.'
            details = 'Files saved to:\n' + export_folder + '\n\nExported files:\n'
            details += '\n'.join(['- ' + os.path.basename(f) for f in exported_files])
            forms.alert(message, sub_msg=details)
        else:
            forms.alert('No files were exported.')
        
    except:
        error_msg = str(sys.exc_info()[1])
        logger.error(error_msg)
        forms.alert('An error occurred during export.', sub_msg=error_msg)

if __name__ == '__main__':
    main()