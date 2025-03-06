# -*- coding: utf-8 -*-
"""Export Schedules to JSON"""
from pyrevit import revit, DB, forms, script
import json
from collections import defaultdict
import os
import sys
import codecs

__title__ = 'Export\nSchedules to JSON'
__author__ = 'Yazid Ben Said'
__doc__ = 'Exports selected schedules to JSON from active or selected Revit files'

logger = script.get_logger()
output = script.get_output()

def get_schedule_data(schedule, doc):
    """Extract all data from a schedule."""
    try:
        schedule_data = {
            "id": schedule.Id.IntegerValue,
            "name": schedule.Name,
            "headers": [],
            "rows": []
        }
        
        try:
            if hasattr(schedule, 'CategoryId') and schedule.CategoryId:
                schedule_data["category"] = schedule.CategoryId.IntegerValue
            else:
                schedule_data["category"] = None
        except:
            schedule_data["category"] = None
        
        # Get schedule definition to access fields
        schedule_definition = schedule.Definition
        
        for field_id in schedule_definition.GetFieldOrder():
            field = schedule_definition.GetField(field_id)
            field_info = {
                "id": field_id.IntegerValue,
                "name": field.GetName(),
                "type": str(field.FieldType)
            }
            schedule_data["headers"].append(field_info)
        
        # Create schedule export options
        schedule_export_options = DB.ViewScheduleExportOptions()
        
        # Export schedule to temporary file
        temp_file = os.path.join(os.environ['TEMP'], 'temp_schedule.txt')
        schedule.Export(os.path.dirname(temp_file), os.path.basename(temp_file), schedule_export_options)
        
        # Read and parse the exported file
        with open(temp_file, 'r') as f:
            lines = f.readlines()
            
            
            header_rows = 0
            for i, line in enumerate(lines):
                if line.strip() and i < len(lines) - 1:
                    if all(header["name"] in line for header in schedule_data["headers"] if header["name"]):
                        header_rows = i + 1
                        break
            
            
            if header_rows == 0 and lines:
                header_rows = 1
            
            # Process data rows
            for line in lines[header_rows:]:
                if line.strip(): 
                    row_data = line.strip().split('\t')
                    if len(row_data) >= 1:
                        row = {}
                        for i, cell in enumerate(row_data):
                            if i < len(schedule_data["headers"]):
                                header_name = schedule_data["headers"][i]["name"]
                                row[header_name] = cell.strip()
                        schedule_data["rows"].append(row)
        
        try:
            os.remove(temp_file)
        except:
            pass
            
        return schedule_data
    except Exception as ex:
        logger.error("Error processing schedule {}: {}".format(schedule.Name, str(ex)))
        return None

def export_schedules_to_json(schedules_data, folder_path, file_name=""):
    """Export multiple schedules data to JSON files."""
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
    
    exported_files = []
    for schedule_name, schedule_data in schedules_data.items():
        try:
            safe_filename = "".join([c for c in schedule_name if c.isalnum() or c in (' ','-','_')]).rstrip()
            if file_name:
                safe_filename = file_name + "_" + safe_filename
            filepath = os.path.join(folder_path, safe_filename + ".json")
            
            with codecs.open(filepath, 'w', encoding='utf-8') as f:
                json_str = json.dumps(schedule_data, indent=4, ensure_ascii=False)
                f.write(json_str)
            exported_files.append(filepath)
        except Exception as ex:
            output.print_md("Error exporting schedule {}: {}".format(schedule_name, str(ex)))
            continue
    return exported_files

def process_document(doc, export_folder, file_name=""):
    """Process a single document and export selected schedules."""
    # Get all schedules
    schedules = DB.FilteredElementCollector(doc)\
                 .OfClass(DB.ViewSchedule)\
                 .WhereElementIsNotElementType()\
                 .ToElements()
    
    # Filter out invalid schedules
    valid_schedules = [s for s in schedules if not s.IsTemplate and s.Name != 'Schedule']
    
    if not valid_schedules:
        output.print_md("No valid schedules found in document: {}".format(doc.Title))
        return []
    
    # Show schedule selection dialog
    schedule_options = sorted([s.Name for s in valid_schedules])
    selected_schedules = forms.SelectFromList.show(
        schedule_options,
        title='Sélectionner les nomenclatures à exporter: {}'.format(doc.Title),
        button_name='Exporter en JSON',
        multiselect=True,
        width=500,
        height=600
    )
    
    if not selected_schedules:
        return []
    
    schedules_to_export = [s for s in valid_schedules if s.Name in selected_schedules]
    schedules_data = {}
    
    # Process schedules with progress bar
    with forms.ProgressBar() as pb:
        total_schedules = len(schedules_to_export)
        for idx, schedule in enumerate(schedules_to_export):
            pb.update_progress(idx, max_value=total_schedules)
            output.print_md("Processing schedule: **{}**".format(schedule.Name))
            schedule_data = get_schedule_data(schedule, doc)
            if schedule_data:
                schedules_data[schedule.Name] = schedule_data
    
    return export_schedules_to_json(schedules_data, export_folder, file_name)

def select_export_mode():
    """Let user select export mode."""
    options = {
        'Exporter un document actif': 'active',
        'Exporter d\'autres fichiers Revit': 'files'
    }
    selected_option = forms.CommandSwitchWindow.show(
        options.keys(),
        message='Sélectionner le mode d\'export:'
    )
    return options.get(selected_option)

def main():
    try:
        # Let user choose mode
        mode = select_export_mode()
        if not mode:
            return

        # Let user select export folder
        export_folder = forms.pick_folder(
            title='Sélectionner un dossier de destination pour l\'export'
        )
        if not export_folder:
            return

        exported_files = []
        
        if mode == 'active':
            # Process active document
            doc = revit.doc
            file_name = doc.Title.replace('.rvt', '')
            exported_files = process_document(doc, export_folder, file_name)
        else:
            # Let user select Revit files
            file_paths = forms.pick_file(
                file_ext='rvt',
                multi_file=True,
                title='Sélectionner les fichiers Revit à exporter'
            )
            
            if not file_paths:
                return
            
            # Get application handle
            app = __revit__.Application
            
            total_files = len(file_paths)
            with forms.ProgressBar() as pb:
                for idx, file_path in enumerate(file_paths):
                    pb.update_progress(idx, max_value=total_files)
                    file_name = os.path.splitext(os.path.basename(file_path))[0]
                    output.print_md("Processing file: **{}** ({}/{})".format(
                        file_name, idx + 1, total_files))
                    
                    try:
                        doc = app.OpenDocumentFile(file_path)
                        exported = process_document(doc, export_folder, file_name)
                        exported_files.extend(exported)
                        doc.Close(False)
                    except Exception as ex:
                        logger.error("Error processing file {}: {}".format(file_path, str(ex)))
                        continue

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