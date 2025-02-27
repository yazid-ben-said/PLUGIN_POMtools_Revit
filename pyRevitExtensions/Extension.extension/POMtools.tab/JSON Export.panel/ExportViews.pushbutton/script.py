# -*- coding: utf-8 -*-
"""Export Views to JSON"""
from pyrevit import revit, DB, forms, script
import json
from collections import defaultdict
import os
import sys
import codecs

__title__ = 'Export\nViews to JSON'
__author__ = 'Yazid'
__doc__ = 'Exports selected views to JSON from active or selected Revit files'

logger = script.get_logger()
output = script.get_output()

def get_parameter_value(param):
   """Get parameter value based on its storage type."""
   try:
       if param.StorageType == DB.StorageType.String:
           val = param.AsString()
           return val.encode('utf-8').decode('utf-8') if val else None
       elif param.StorageType == DB.StorageType.Double:
           return param.AsDouble()
       elif param.StorageType == DB.StorageType.Integer:
           return param.AsInteger()
       elif param.StorageType == DB.StorageType.ElementId:
           return param.AsElementId().IntegerValue
       else:
           return None
   except:
       return None

def get_element_data(element):
   """Extract relevant data from an element."""
   try:
       # Base element data
       element_data = {
           "id": element.Id.IntegerValue,
           "category": element.Category.Name if element.Category else "Uncategorized",
           "type_name": element.Name if hasattr(element, 'Name') else None,
           "parameters": {}
       }

       # Handle family name differently for different element types
       try:
           if isinstance(element, DB.TextNote):
               element_data["family_name"] = "Text Note"
               element_data["text_content"] = element.Text
           elif isinstance(element, DB.Dimension):
               element_data["family_name"] = "Dimension"
           elif hasattr(element, 'Symbol') and element.Symbol:
               element_data["family_name"] = element.Symbol.Family.Name
           else:
               element_data["family_name"] = None
       except:
           element_data["family_name"] = None
       
       # Get all parameters
       for param in element.Parameters:
           if param.Definition:
               param_value = get_parameter_value(param)
               if param_value is not None:
                   element_data["parameters"][param.Definition.Name] = param_value
       
       # Get location data if available
       try:
           location = element.Location
           if location:
               if isinstance(location, DB.LocationPoint):
                   point = location.Point
                   element_data["location"] = {
                       "x": point.X,
                       "y": point.Y,
                       "z": point.Z
                   }
               elif isinstance(location, DB.LocationCurve):
                   curve = location.Curve
                   start = curve.GetEndPoint(0)
                   end = curve.GetEndPoint(1)
                   element_data["location"] = {
                       "start_point": {"x": start.X, "y": start.Y, "z": start.Z},
                       "end_point": {"x": end.X, "y": end.Y, "z": end.Z}
                   }
       except:
           pass
       
       return element_data
   except Exception as ex:
       logger.error("Error processing element: {}".format(str(ex)))
       return None

def get_view_data(view, doc):
   """Extract all data from the given view."""
   try:
       discipline = str(view.Discipline) if hasattr(view, 'Discipline') else "Unknown"
   except:
       discipline = "Unknown"
       
   try:
       detail_level = str(view.DetailLevel) if hasattr(view, 'DetailLevel') else "Unknown"
   except:
       detail_level = "Unknown"

   view_data = {
       "id": view.Id.IntegerValue,
       "name": view.Name,
       "view_type": str(view.ViewType),
       "scale": view.Scale,
       "level": view.GenLevel.Name if hasattr(view, 'GenLevel') and view.GenLevel else None,
       "template": view.ViewTemplateId.IntegerValue if hasattr(view, 'ViewTemplateId') else None,
       "detail_level": detail_level,
       "discipline": discipline,
       "elements": []
   }
   
   try:
       collector = DB.FilteredElementCollector(doc, view.Id)\
                    .WhereElementIsNotElementType()\
                    .ToElements()
       
       categorized_elements = defaultdict(list)
       for element in collector:
           if element and element.Category:
               element_data = get_element_data(element)
               if element_data:
                   categorized_elements[element.Category.Name].append(element_data)
       
       for category, elements in categorized_elements.items():
           category_data = {
               "category": category,
               "elements": elements
           }
           view_data["elements"].append(category_data)
   except Exception as ex:
       logger.error("Error collecting elements from view: {}".format(str(ex)))
       view_data["elements"] = []
       
   return view_data

def export_views_to_json(views_data, folder_path, file_name=""):
   """Export multiple views data to JSON files."""
   if not os.path.exists(folder_path):
       os.makedirs(folder_path)
   
   exported_files = []
   for view_name, view_data in views_data.items():
       try:
           safe_filename = "".join([c for c in view_name if c.isalnum() or c in (' ','-','_')]).rstrip()
           if file_name:
               safe_filename = file_name + "_" + safe_filename
           filepath = os.path.join(folder_path, safe_filename + ".json")
           
           with codecs.open(filepath, 'w', encoding='utf-8') as f:
               json_str = json.dumps(view_data, indent=4, ensure_ascii=False)
               f.write(json_str)
           exported_files.append(filepath)
       except Exception as ex:
           output.print_md("Error exporting view {}: {}".format(view_name, str(ex)))
           continue
   return exported_files

def process_document(doc, export_folder, file_name=""):
   """Process a single document and export selected views."""
   # Get all valid views
   views = DB.FilteredElementCollector(doc).OfClass(DB.View).ToElements()
   valid_views = [v for v in views if 
                 not v.IsTemplate and
                 v.CanBePrinted and
                 v.ViewType not in [DB.ViewType.Schedule, DB.ViewType.Undefined]]
   
   if not valid_views:
       output.print_md("No valid views found in document: {}".format(doc.Title))
       return []
   
   # Show view selection dialog
   view_options = sorted([v.Name for v in valid_views])
   selected_views = forms.SelectFromList.show(
       view_options,
       title='Sélectionner les vues à exporter: {}'.format(doc.Title),
       button_name='Exporter en JSON',
       multiselect=True,
       width=500,
       height=600
   )
   
   if not selected_views:
       return []
   
   views_to_export = [v for v in valid_views if v.Name in selected_views]
   views_data = {}
   
   # Process views with progress bar
   with forms.ProgressBar() as pb:
       total_views = len(views_to_export)
       for idx, view in enumerate(views_to_export):
           pb.update_progress(idx, max_value=total_views)
           output.print_md("Processing view: **{}**".format(view.Name))
           views_data[view.Name] = get_view_data(view, doc)
   
   return export_views_to_json(views_data, export_folder, file_name)

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