# -*- coding: utf-8 -*-
"""Export Views to IFC (External Files)"""
from pyrevit import revit, DB, forms, script
import json
import os
import sys
import codecs
from System.Collections.Generic import List

__title__ = 'Export\nViews to IFC\n(Autres Fichiers)'
__author__ = 'Yazid Ben Said'
__doc__ = 'Exports selected views to IFC from other Revit files'

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

def generate_default_ifc_config(filepath):
    """Generate a default IFC configuration file."""
    try:
        default_config = {
            "IFCVersion": 21,
            "ExchangeRequirement": 3,
            "IFCFileType": 0,
            "SpaceBoundaries": 0,
            "SplitWallsAndColumns": False,
            "IncludeSteelElements": True,
            "ExportBaseQuantities": False,
            "Export2DElements": False,
            "ExportLinkedFiles": False,
            "VisibleElementsOfCurrentView": True,
            "ExportRoomsInView": False,
            "ExportInternalRevitPropertySets": True,
            "ExportIFCCommonPropertySets": False,
            "TessellationLevelOfDetail": 0.5,
            "UseActiveViewGeometry": True,
            "UseFamilyAndTypeNameForReference": False,
            "Use2DRoomBoundaryForVolume": False,
            "IncludeSiteElevation": True,
            "StoreIFCGUID": True
        }
        
        with codecs.open(filepath, 'w', encoding='utf-8') as f:
            json.dump(default_config, f, indent=2)
            
        return True
    except Exception as ex:
        logger.error("Error generating default IFC configuration: {}".format(str(ex)))
        return False

def apply_ifc_config_to_options(config_data, ifc_options):
    """Apply configuration from JSON to IFC export options."""
    try:
        
        # 1. IFC version - essential
        if "IFCVersion" in config_data:
            try:
                if config_data["IFCVersion"] == 21:
                    ifc_options.FileVersion = DB.IFCVersion.IFC2x3CV2
                elif config_data["IFCVersion"] == 23:
                    ifc_options.FileVersion = DB.IFCVersion.IFC4
            except Exception as ex:
                logger.warning("Could not set IFC version: {}".format(str(ex)))
        
        # 2. Space boundaries - essential
        if "SpaceBoundaries" in config_data:
            try:
                ifc_options.SpaceBoundaryLevel = config_data["SpaceBoundaries"]
            except Exception as ex:
                logger.warning("Could not set space boundaries: {}".format(str(ex)))
        
        # 3. Wall and column splitting - essential
        if "SplitWallsAndColumns" in config_data:
            try:
                ifc_options.WallAndColumnSplitting = config_data["SplitWallsAndColumns"]
            except Exception as ex:
                logger.warning("Could not set wall and column splitting: {}".format(str(ex)))
        
        # 4. Base quantities - essential
        if "ExportBaseQuantities" in config_data:
            try:
                ifc_options.ExportBaseQuantities = config_data["ExportBaseQuantities"]
            except Exception as ex:
                logger.warning("Could not set base quantities: {}".format(str(ex)))
        
        # Use simplified option mapping to reduce potential errors
        simple_options = {
            "VisibleElementsOfCurrentView": lambda x: str(x).lower(),
            "Export2DElements": lambda x: str(x).lower(),
            "ExportLinkedFiles": lambda x: str(x).lower(),
            "UseActiveViewGeometry": lambda x: str(x).lower()
        }
        
        # Apply only the most essential options using AddOption
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

def open_and_process_revit_files(file_paths, export_folder, config_file=None, use_active_view_only=False, ifc_version='default'):
    """Opens and processes multiple Revit files."""
    if not file_paths:
        return []
        
    # Get application handle
    app = __revit__.Application
    exported_files = []
    
    version_text = "IFC 2x3 - Vue de coordination 2.0"
    if ifc_version == 'ifc4':
        version_text = "IFC 4 - Vue de référence"
    elif ifc_version == 'custom':
        version_text = "configuration personnalisée"
        
    forms.alert(
        'Information importante pour le traitement par lots',
        sub_msg='Ce script va traiter séquentiellement les fichiers Revit sélectionnés.'
        '\n\nFormat IFC: {}'
        '\n\nPour chaque fichier, une boîte de dialogue apparaîtra pour sélectionner les vues à exporter.'
        '\n\nLes fichiers seront ouverts un par un, puis fermés une fois l\'export terminé.'
        '\n\nCliquez sur OK pour commencer.'.format(version_text)
    )
 
    total_files = len(file_paths)
    successful_files = 0
    failed_files = []
    
    with forms.ProgressBar() as pb:
        for idx, file_path in enumerate(file_paths):
            pb.update_progress(idx, max_value=total_files)
            file_name = os.path.splitext(os.path.basename(file_path))[0]
            output.print_md("Traitement du fichier: **{}** ({}/{})".format(
                file_name, idx + 1, total_files))
            
            try:
                # Open document in detached mode to prevent locking
                open_options = DB.OpenOptions()
                open_options.DetachFromCentralOption = DB.DetachFromCentralOption.DetachAndPreserveWorksets
                
                output.print_md("Ouverture du fichier...")
                # Convert string path to ModelPath object required by Revit API
                model_path = DB.ModelPathUtils.ConvertUserVisiblePathToModelPath(file_path)
                doc = app.OpenDocumentFile(model_path, open_options)
                
                # Process this document
                output.print_md("Fichier ouvert. Sélection des vues...")
                try:
                    this_file_exports = process_document(doc, export_folder, config_file, file_name, use_active_view_only, ifc_version)
                    
                    if this_file_exports:
                        exported_files.extend(this_file_exports)
                        successful_files += 1
                        output.print_md("**Export terminé** pour {}. {} vues exportées.".format(
                            file_name, len(this_file_exports)))
                    else:
                        output.print_md("Aucune vue n'a été exportée pour {}".format(file_name))
                        failed_files.append(file_name + " (aucune vue exportée)")
                    
                    # Check if the document is a linked file
                    is_linked = False
                    try:
                        # Try creating and immediately rolling back a transaction to test
                        test_transaction = DB.Transaction(doc, "Test Transaction")
                        test_transaction.Start()
                        test_transaction.RollBack()
                    except Exception:
                        # If transaction fails, it's a linked file
                        is_linked = True
                    
                    # Close document only if it's not a linked file
                    if not is_linked:
                        output.print_md("Fermeture du fichier...")
                        doc.Close(False)
                        output.print_md("Fichier fermé.")
                    else:
                        output.print_md("Le fichier est détecté comme lié, pas besoin de fermeture explicite.")
                
                except Exception as process_ex:
                    output.print_md("**ERREUR** lors du traitement du document: {}".format(str(process_ex)))
                    failed_files.append(file_name + " (" + str(process_ex) + ")")
                
            except Exception as ex:
                output.print_md("**ERREUR** lors du traitement de {}: {}".format(file_name, str(ex)))
                logger.error("Error processing file {}: {}".format(file_path, str(ex)))
                failed_files.append(file_name + " (" + str(ex) + ")")
                continue
    
    # Return summary info
    return {
        "exported_files": exported_files,
        "successful_files": successful_files,
        "failed_files": failed_files,
        "total_files": total_files,
        "ifc_version": version_text
    }

def process_document(doc, export_folder, config_file=None, file_name="", use_active_view_only=False, ifc_version='default'):
    """Process a single document and export selected views to IFC."""
    # Create specific folder for this document
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
    
    # Check if the document is a linked file or not (for transaction handling)
    try:
        # Try creating and immediately rolling back a transaction to test
        test_transaction = DB.Transaction(doc, "Test Transaction")
        test_transaction.Start()
        test_transaction.RollBack()
        is_linked_file = False
    except Exception:
        # If transaction fails, it's a linked file
        is_linked_file = True
        forms.alert(
            "Fichier détecté comme fichier 'spécial'",
            sub_msg="Le fichier {} semble avoir un statut particulier dans Revit (fichier central, fichier lié, etc.).\n\n"
                   "En raison des limitations de l'API Revit, ces fichiers ne peuvent pas être exportés avec cette méthode.\n\n"
                   "Pour exporter ce fichier :\n"
                   "1. Fermez ce script\n"
                   "2. Ouvrez le fichier directement dans Revit (via la commande standard 'Ouvrir')\n"
                   "3. Utilisez l'outil 'Export Views to IFC (Document Actif)'"
        )
        return []
    
    # Process views with progress bar (only for regular, non-linked documents)
    with forms.ProgressBar() as pb:
        total_views = len(views_to_export)
        for idx, view in enumerate(views_to_export):
            pb.update_progress(idx, max_value=total_views)
            output.print_md("Export de la vue: **{}**".format(view.Name))
            
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
                    ifc_options.AddOption("ExchangeRequirement", "ReferenceView")  # Reference View MVD
                    ifc_options.AddOption("IFCVersion", "IFC4")  # Explicitly set IFC4
                    ifc_options.AddOption("ExportBoundingBox", "false")
                    ifc_options.AddOption("UseTypeNameOnlyForIfcType", "true")
                    ifc_options.AddOption("UseOnlyTriangulation", "true")  # For reference view
                else:
                    # Default IFC 2x3 settings - Coordination View 2.0
                    ifc_options.FileVersion = DB.IFCVersion.IFC2x3CV2
                    ifc_options.SpaceBoundaryLevel = 0
                    ifc_options.ExportBaseQuantities = False
                    ifc_options.WallAndColumnSplitting = False
                
                # Default to visible elements of current view
                ifc_options.AddOption("VisibleElementsOfCurrentView", "true")
                
                # If active view only is selected, enforce that setting
                if use_active_view_only:
                    ifc_options.AddOption("VisibleElementsOfCurrentView", "true")
                    # Try to reduce exported data by setting more restrictive options
                    ifc_options.AddOption("ExportLinkedFiles", "false")
                    ifc_options.AddOption("Export2DElements", "false")
                
                # If we have a custom config file, apply it after defaults
                if config_file and os.path.exists(config_file):
                    config_data = load_ifc_config_from_json(config_file)
                    if config_data:
                        apply_ifc_config_to_options(config_data, ifc_options)
                        # If using active view only, override this setting
                        if use_active_view_only:
                            ifc_options.AddOption("VisibleElementsOfCurrentView", "true")
                
                # Set the view to export
                ifc_options.FilterViewId = view.Id
                
                # Create export filename - add prefix based on IFC version
                safe_viewname = "".join([c for c in view.Name if c.isalnum() or c in (' ','-','_')]).rstrip()
                
                # Add a prefix based on IFC version
                if ifc_version == 'ifc4':
                    prefix = "IFC4_"
                else:
                    prefix = "IFC2x3_"
                    
                safe_filename = prefix + safe_viewname
                    
                filepath = os.path.join(doc_folder, safe_filename + ".ifc")
                
                # For regular documents, use transaction
                with DB.Transaction(doc, "IFC Export") as t:
                    t.Start()
                    
                    try:
                        result = doc.Export(doc_folder, safe_filename, ifc_options)
                        t.Commit()
                        
                        if result:
                            exported_files.append(filepath)
                            output.print_md("Vue exportée avec succès en {}!".format(
                                "IFC4 - Vue de référence" if ifc_version == 'ifc4' else "IFC 2x3 - Vue de coordination 2.0"
                            ))
                        else:
                            output.print_md("**Export échoué** pour cette vue.")
                    except Exception as export_ex:
                        # If export fails, roll back the transaction
                        t.RollBack()
                        logger.error("Error during export process: {}".format(str(export_ex)))
                        raise
                
            except Exception as ex:
                output.print_md("Erreur d'export de la vue: **{}** - {}".format(view.Name, str(ex)))
                # Continue with next view even if one fails
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
            
        # Let user select Revit files
        file_paths = forms.pick_file(
            file_ext='rvt',
            multi_file=True,
            title='Sélectionner les fichiers Revit à exporter'
        )
        
        if not file_paths:
            forms.alert('Aucun fichier sélectionné. Opération annulée.')
            return
            
        # Process the files
        results = open_and_process_revit_files(file_paths, export_folder, config_file, use_active_view_only, config_option)
        
        if results["exported_files"]:
            message = 'Export IFC terminé.'
            details = 'Résumé:\n'
            details += '- Format utilisé: {}\n'.format(results["ifc_version"])
            details += '- Fichiers traités: {}/{}\n'.format(results["successful_files"], results["total_files"])
            details += '- Vues exportées: {}\n\n'.format(len(results["exported_files"]))
            
            if results["failed_files"]:
                details += 'Échecs:\n'
                for failed in results["failed_files"]:
                    details += '- {}\n'.format(failed)
                    
            details += '\nFichiers exportés dans: {}'.format(export_folder)
            forms.alert(message, sub_msg=details)
        else:
            forms.alert('Aucun fichier n\'a été exporté.')
    
    except Exception as ex:
        error_msg = str(ex)
        logger.error(error_msg)
        forms.alert('Une erreur s\'est produite pendant l\'export IFC.', sub_msg=error_msg)

if __name__ == '__main__':
    main()