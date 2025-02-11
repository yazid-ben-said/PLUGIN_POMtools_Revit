# -*- coding: utf-8 -*-
"""Export Views to JSON"""
from pyrevit import revit, DB, forms, script
import json
from collections import defaultdict
import os
import sys
import codecs

__title__ = 'Export\nViews to JSON'
__author__ = 'Yazid Ben Said'
__doc__ = 'Exports selected views data to JSON'

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
               # Add text specific data
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
   
   # Get all visible elements in the view
   try:
       collector = DB.FilteredElementCollector(doc, view.Id)\
                    .WhereElementIsNotElementType()\
                    .ToElements()
       
       # Group elements by category
       categorized_elements = defaultdict(list)
       for element in collector:
           if element and element.Category:
               element_data = get_element_data(element)
               if element_data:
                   categorized_elements[element.Category.Name].append(element_data)
       
       # Process elements by category
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

def export_views_to_json(views_data, folder_path):
   """Export multiple views data to JSON files."""
   # Create the folder if it doesn't exist
   if not os.path.exists(folder_path):
       os.makedirs(folder_path)
   
   # Export each view to a separate JSON file
   for view_name, view_data in views_data.items():
       try:
           # Create safe filename
           safe_filename = "".join([c for c in view_name if c.isalnum() or c in (' ','-','_')]).rstrip()
           filepath = os.path.join(folder_path, safe_filename + ".json")
           
           # Export to JSON with proper formatting and encoding
           with codecs.open(filepath, 'w', encoding='utf-8') as f:
               json_str = json.dumps(view_data, indent=4, ensure_ascii=False)
               f.write(json_str)
       except Exception as ex:
           output.print_md("Error exporting view {}: {}".format(view_name, str(ex)))
           continue

def main():
   try:
       # Get current document
       doc = revit.doc
       
       # Get all valid views
       views = DB.FilteredElementCollector(doc).OfClass(DB.View).ToElements()
       
       # Filter views: exclude templates, schedules, and undefined views
       valid_views = [v for v in views if 
                     not v.IsTemplate and
                     v.CanBePrinted and
                     v.ViewType not in [DB.ViewType.Schedule, DB.ViewType.Undefined]]
       
       if not valid_views:
           forms.alert('No valid views found in the document.')
           return
       
       # Show view selection dialog with checkboxes
       view_options = sorted([v.Name for v in valid_views])  # Sort views alphabetically
       selected_views = forms.SelectFromList.show(
           view_options,
           title='Selectionner les vues à exporter',
           button_name='Exporter en JSON',
           multiselect=True,
           width=500,
           height=600
       )
       
       if not selected_views:
           return
       
       # Get the actual view objects from the selected names
       views_to_export = [v for v in valid_views if v.Name in selected_views]
       
       # Let user select export folder
       export_folder = forms.pick_folder()
       if not export_folder:
           return
       
       views_data = {}
       
       # Setup progress bar
       total_views = len(views_to_export)
       with forms.ProgressBar() as pb:
           # Process each view
           for idx, view in enumerate(views_to_export):
               # Update progress
               pb.update_progress(idx, max_value=total_views)
               output.print_md("Processing view: **{}**".format(view.Name))
               
               views_data[view.Name] = get_view_data(view, doc)
           
           # Export all views
           export_views_to_json(views_data, export_folder)
       
       # Show success message with export details
       message = 'Export completed successfully!'
       details = 'Files saved to:\n' + export_folder + '\n\nExported views:\n'
       details += '\n'.join(['- ' + view for view in selected_views])
       
       forms.alert(
           message,
           sub_msg=details
       )
       
   except:
       error_msg = str(sys.exc_info()[1])
       logger.error(error_msg)
       forms.alert(
           'An error occurred during export.',
           sub_msg=error_msg
       )

if __name__ == '__main__':
   main()