# -*- coding: utf-8 -*-
"""Export Views to IFC (Active Document)"""
from pyrevit import revit, DB, forms, script
import json
import os
import sys
import codecs
from System.Collections.Generic import List

__title__ = 'Export\nViews to IFC\n(Document Actif)'
__author__ = 'Yazid Ben Said'
__doc__ = 'Exports selected views to IFC from active Revit document'

logger = script.get_logger()
output = script.get_output()

# IFC Export Constants
IFC_VERSION_2x3 = "IFC2x3"
IFC_EXPORT_CONFIG_NAME = "VUE de coordination 2.0"

def load_ifc_config_from_json(json_path):
    """Load IFC export configuration from a JSON file."""
    try:
        with codecs.open(json_path, 'r', encoding='utf-8') as f:
            config_data = json.load(f)
        return config_data
    except Exception as ex:
        logger.error("Error loading IFC configuration from JSON: {}".format(str(ex)))
        forms.alert('Erreur de chargement de la configuration IFC.', sub_msg=str(ex))
        return None

def apply_ifc_config_to_options(config_data, ifc_options):
    """Apply configuration from JSON to IFC export options."""
    try:
        if "IFCVersion" in config_data:
            try:
                if config_data["IFCVersion"] == 21:
                    ifc_options.FileVersion = DB.IFCVersion.IFC2x3CV2
                elif config_data["IFCVersion"] == 23 or config_data["IFCVersion"] == 25:  
                    ifc_options.FileVersion = DB.IFCVersion.IFC4
            except Exception as ex:
                logger.warning("Could not set IFC version: {}".format(str(ex)))
        

        if "SpaceBoundaries" in config_data:
            try:
                ifc_options.SpaceBoundaryLevel = config_data["SpaceBoundaries"]
            except Exception as ex:
                logger.warning("Could not set space boundaries: {}".format(str(ex)))
        

        if "SplitWallsAndColumns" in config_data:
            try:
                ifc_options.WallAndColumnSplitting = config_data["SplitWallsAndColumns"]
            except Exception as ex:
                logger.warning("Could not set wall and column splitting: {}".format(str(ex)))
        

        if "ExportBaseQuantities" in config_data:
            try:
                ifc_options.ExportBaseQuantities = config_data["ExportBaseQuantities"]
            except Exception as ex:
                logger.warning("Could not set base quantities: {}".format(str(ex)))
        
        
        if "ExportRoomsInView" in config_data:
            try:
                ifc_options.AddOption("ExportRoomsInView", str(config_data["ExportRoomsInView"]).lower())
            except Exception as ex:
                logger.warning("Could not set ExportRoomsInView: {}".format(str(ex)))

        
        if "ExportSolidModelRep" in config_data:
            try:
                ifc_options.AddOption("ExportSolidModelRep", str(config_data["ExportSolidModelRep"]).lower())
            except Exception as ex:
                logger.warning("Could not set ExportSolidModelRep: {}".format(str(ex)))
                
        
        if "ExcludeFilter" in config_data:
            try:
                ifc_options.AddOption("ExcludeFilter", config_data["ExcludeFilter"])
            except Exception as ex:
                logger.warning("Could not set ExcludeFilter: {}".format(str(ex)))

        simple_options = {
            "VisibleElementsOfCurrentView": lambda x: str(x).lower(),
            "Export2DElements": lambda x: str(x).lower(),
            "ExportLinkedFiles": lambda x: str(x).lower(),
            "UseActiveViewGeometry": lambda x: str(x).lower(),
            "ExportIFCCommonPropertySets": lambda x: str(x).lower(),
            "Use2DRoomBoundaryForVolume": lambda x: str(x).lower(),
            "UseOnlyTriangulation": lambda x: str(x).lower(),
            "UseTypeNameOnlyForIfcType": lambda x: str(x).lower()
        }
        

        for option_name, converter in simple_options.items():
            if option_name in config_data:
                try:
                    option_value = converter(config_data[option_name])
                    ifc_options.AddOption(option_name, option_value)
                except Exception as ex:
                    logger.warning("Could not set option {}: {}".format(option_name, str(ex)))
                
        return True
    except Exception as ex:
        logger.error("Error applying IFC configuration: {}".format(str(ex)))
        return False

def export_view_to_ifc(doc, view, doc_folder, config_file=None, use_active_view_only=False, ifc_version='default'):
    """Export a view to IFC format."""
    try:
        # Create IFC export options with settings based on selected version
        ifc_options = DB.IFCExportOptions()
        
        if ifc_version == 'ifc4':
            # IFC4 - Reference View settings
            ifc_options.FileVersion = DB.IFCVersion.IFC4
            ifc_options.SpaceBoundaryLevel = 0
            ifc_options.ExportBaseQuantities = False
            ifc_options.WallAndColumnSplitting = False
            
            # Add IFC4 Reference View specific options
            ifc_options.AddOption("ExchangeRequirement", "ReferenceView") 
            ifc_options.AddOption("IFCVersion", "IFC4") 
            ifc_options.AddOption("ExportBoundingBox", "false")
            ifc_options.AddOption("UseTypeNameOnlyForIfcType", "true")
            ifc_options.AddOption("UseOnlyTriangulation", "true")
        else:
            # Default IFC 2x3 settings - Coordination View 2.0
            ifc_options.FileVersion = DB.IFCVersion.IFC2x3CV2
            ifc_options.SpaceBoundaryLevel = 0
            ifc_options.ExportBaseQuantities = False
            ifc_options.WallAndColumnSplitting = False
        
        # Default to visible elements of current view for all versions
        ifc_options.AddOption("VisibleElementsOfCurrentView", "true")
        
        # If active view only is selected, enforce that setting for all versions
        if use_active_view_only:
            ifc_options.AddOption("VisibleElementsOfCurrentView", "true")
            ifc_options.AddOption("ExportLinkedFiles", "false")
            ifc_options.AddOption("Export2DElements", "false")
        
        # If we have a custom config file, apply it after defaults
        if config_file and os.path.exists(config_file):
            config_data = load_ifc_config_from_json(config_file)
            if config_data:
                apply_ifc_config_to_options(config_data, ifc_options)
                if use_active_view_only:
                    ifc_options.AddOption("VisibleElementsOfCurrentView", "true")
        
        # Set the view to export
        ifc_options.FilterViewId = view.Id
        
        # Create export filename - just the view name
        safe_viewname = "".join([c for c in view.Name if c.isalnum() or c in (' ','-','_')]).rstrip()
        
        # Add a prefix based on IFC version
        if ifc_version == 'ifc4' or (config_file and "IFCVersion" in load_ifc_config_from_json(config_file) and load_ifc_config_from_json(config_file)["IFCVersion"] in [23, 25]):
            prefix = "IFC4_"
        else:
            prefix = "IFC2x3_"
            
        safe_filename = prefix + safe_viewname
            
        filepath = os.path.join(doc_folder, safe_filename + ".ifc")
        
        # Start a transaction to allow document modifications
        with DB.Transaction(doc, "IFC Export") as t:
            t.Start()
            
            try:
                # Execute the export
                result = doc.Export(doc_folder, safe_filename, ifc_options)
                t.Commit()
                
                if result:
                    version_display = "IFC 2x3 - Vue de coordination 2.0"
                    if ifc_version == 'ifc4' or (config_file and "IFCVersion" in load_ifc_config_from_json(config_file) and load_ifc_config_from_json(config_file)["IFCVersion"] in [23, 25]):
                        version_display = "IFC4 - Vue de référence"
                    elif config_file:
                        version_display = "configuration personnalisée"
                    
                    output.print_md("Vue **{}** exportée avec succès en {}!".format(
                        view.Name, version_display
                    ))
                    return filepath
                else:
                    logger.error("IFC export failed for view: {}".format(view.Name))
                    output.print_md("**ÉCHEC** de l'export pour la vue: {}".format(view.Name))
                    return None
            except Exception as export_ex:
                # If export fails, roll back the transaction
                t.RollBack()
                logger.error("Error during export process: {}".format(str(export_ex)))
                raise
            
    except Exception as ex:
        logger.error("Error exporting view to IFC: {}".format(str(ex)))
        return None

def process_document(doc, export_folder, config_file=None, use_active_view_only=False, ifc_version='default'):
    """Process a single document and export selected views to IFC."""
    # Create specific folder for this document
    file_name = doc.Title.replace('.rvt', '')
    # Remove username if present (after underscore)
    if '_' in file_name:
        file_name = file_name.split('_')[0]
        
    doc_folder = os.path.join(export_folder, file_name)
    if not os.path.exists(doc_folder):
        os.makedirs(doc_folder)
        
    # Get all valid views
    views = DB.FilteredElementCollector(doc).OfClass(DB.View).ToElements()
    valid_views = [v for v in views if 
                  not v.IsTemplate and
                  v.CanBePrinted and
                  v.ViewType not in [DB.ViewType.Schedule, DB.ViewType.Undefined]]
    
    if not valid_views:
        output.print_md("Aucune vue valide trouvée dans le document: {}".format(doc.Title))
        return []
    
    # Show view selection dialog
    view_options = sorted([v.Name for v in valid_views])
    selected_views = forms.SelectFromList.show(
        view_options,
        title='Sélectionner les vues à exporter: {}'.format(doc.Title),
        button_name='Exporter en IFC',
        multiselect=True,
        width=500,
        height=600
    )
    
    if not selected_views:
        return []
    
    views_to_export = [v for v in valid_views if v.Name in selected_views]
    exported_files = []
    
    # Process views with progress bar
    with forms.ProgressBar() as pb:
        total_views = len(views_to_export)
        for idx, view in enumerate(views_to_export):
            pb.update_progress(idx, max_value=total_views)
            output.print_md("Export de la vue: **{}**".format(view.Name))
            
            try:
                exported_file = export_view_to_ifc(doc, view, doc_folder, config_file, use_active_view_only, ifc_version)
                if exported_file:
                    exported_files.append(exported_file)
            except Exception as ex:
                output.print_md("Erreur d'export de la vue: **{}** - {}".format(view.Name, str(ex)))
                continue
    
    return exported_files

def config_options():
    """Let user select configuration options."""
    options = {
        'Utiliser la configuration IFC 2x3 - Vue de coordination 2.0': 'default',
        'Utiliser la configuration IFC 4 - Vue de référence': 'ifc4',
        'Utiliser une configuration IFC personnalisée (JSON)': 'custom'
    }
    selected_option = forms.CommandSwitchWindow.show(
        options.keys(),
        message='Options de configuration:'
    )
    return options.get(selected_option)

def main():
    try:
        # Let user choose configuration option
        config_option = config_options()
        if not config_option:
            return
            
        config_file = None
        if config_option == 'custom':
            config_file = forms.pick_file(
                file_ext='json',
                title='Sélectionner un fichier de configuration IFC (JSON)'
            )
            if not config_file:
                # Fallback to default if no file selected
                config_option = 'default'
                
        # Ask about use active view only option
        use_active_view_only = forms.alert(
            'Voulez-vous utiliser l\'option "Vue active uniquement" pour réduire la taille d\'export?',
            yes=True, no=True, ok=False
        )
        
        # Let user select export folder
        export_folder = forms.pick_folder(
            title='Sélectionner un dossier de destination pour l\'export IFC'
        )
        if not export_folder:
            return

        # Process active document
        doc = revit.doc
        exported_files = process_document(doc, export_folder, config_file, use_active_view_only, config_option)

        if exported_files:
            file_name = doc.Title.replace('.rvt', '')
            # Remove username if present (after underscore)
            if '_' in file_name:
                file_name = file_name.split('_')[0]
                
            doc_folder = os.path.join(export_folder, file_name)
            
            # Determine IFC version used for message
            version_text = "IFC 2x3 - Vue de coordination 2.0"
            if config_option == 'ifc4':
                version_text = "IFC 4 - Vue de référence"
            elif config_option == 'custom':
                if config_file and os.path.exists(config_file):
                    config_data = load_ifc_config_from_json(config_file)
                    if config_data and "IFCVersion" in config_data and config_data["IFCVersion"] in [23, 25]:
                        version_text = "IFC 4 - Configuration personnalisée"
                    else:
                        version_text = "IFC 2x3 - Configuration personnalisée"
                else:
                    version_text = "Configuration personnalisée"
            
            message = 'Export IFC terminé avec succès.'
            details = 'Format utilisé: {}\n\nFichiers enregistrés dans:\n{}\n\nNombre de vues exportées: {}'.format(
                version_text, doc_folder, len(exported_files))
            forms.alert(message, sub_msg=details)
        else:
            forms.alert('Aucun fichier n\'a été exporté.')
        
    except Exception as ex:
        error_msg = str(ex)
        logger.error(error_msg)
        forms.alert('Une erreur s\'est produite pendant l\'export IFC.', sub_msg=error_msg)

if __name__ == '__main__':
    main()

