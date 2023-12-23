#-------------------------------------------------------------------------------------------
# Copyright (c) 2020-2022 5@xes
# 
# ImportExportProfiles is released under the terms of the AGPLv3 or higher.
#
# Version 0.0.3 : First functionnal release
# Version 1.0.5 : top_bottom for new release (Ready Arachne or futur 4.9?)
# Version 1.0.6 : bug correction
# Version 1.0.7 : Add sniff function for the import csv function
#
# Version 1.1.0 : Update Cura 5.0
#
# Version 1.2.0 : Integrate French Translation
# Version 1.2.1 : New texte translated
# Version 1.2.2 : Add import CuraProfile
# Version 1.2.3 : Analyse Quality to Substitute non-existing quality
# Version 1.2.4 : Change i18n location
#
# Version 1.3.0 : Export also Setting Label & add import by step
#-------------------------------------------------------------------------------------------


VERSION_QT5 = False
try:
    from PyQt6.QtCore import QObject
    from PyQt6.QtCore import QTimer
    from PyQt6.QtCore import pyqtSlot
    from PyQt6.QtWidgets import QFileDialog, QMessageBox
except ImportError:
    from PyQt5.QtCore import QObject
    from PyQt5.QtCore import QTimer
    from PyQt5.QtCore import pyqtSlot
    from PyQt5.QtWidgets import QFileDialog, QMessageBox
    VERSION_QT5 = True
    
    
import os
import platform
import os.path
import sys
import re
import time

from datetime import datetime
from typing import cast, Dict, List, Optional, Tuple, Any, Set
from cura.CuraApplication import CuraApplication
from cura.Settings.cura_empty_instance_containers import empty_quality_container
from cura.Machines.ContainerTree import ContainerTree
from cura.ReaderWriters.ProfileReader import NoProfileException, ProfileReader


from cura.CuraVersion import CuraVersion  # type: ignore



# Python csv  : https://docs.python.org/fr/2/library/csv.html
#               https://docs.python.org/3/library/csv.html
# Code from Aldo Hoeben / fieldOfView for this tips
try:
    import csv
except ImportError:
    # older versions of Cura somehow ship with a python version that does not include
    # this file, so a local copy is supplied as a fallback
    # thanks to Aldo Hoeben / fieldOfView for this tips
    from . import csv

from UM.Extension import Extension
from UM.Application import Application
from UM.Logger import Logger
from UM.Message import Message
from UM.Version import Version
from UM.i18n import i18nCatalog
from UM.Resources import Resources
from UM.PluginRegistry import PluginRegistry  # For getting the possible profile writers to write with.
from UM.Settings.ContainerRegistry import ContainerRegistry
from UM.Settings.Interfaces import ContainerInterface, ContainerRegistryInterface
from UM.Settings.InstanceContainer import InstanceContainer
from UM.Util import parseBool

i18n_cura_catalog = i18nCatalog("cura")

Resources.addSearchPath(
	os.path.join(os.path.abspath(os.path.dirname(__file__)),'resources')
)  # Plugin translation file import

catalog = i18nCatalog("profiles")

if catalog.hasTranslationLoaded():
	Logger.log("i", "Import Export Profiles Plugin translation loaded!")

class ImportExportProfiles(Extension, QObject,):
    def __init__(self, parent = None) -> None:
        QObject.__init__(self, parent)
        Extension.__init__(self)
        
        self._Section =""

        self._application = Application.getInstance()
        self._preferences = self._application.getPreferences()
        self._preferences.addPreference("import_export_tools/dialog_path", "")
        self._change_dialog = None
        self._update_timer = QTimer()
        self._update_timer.setInterval(0)
        self._update_timer.setSingleShot(True)
        
        self.Major=1
        self.Minor=0

        # Test version for futur release 4.9
        # Logger.log('d', "Info Version CuraVersion --> " + str(Version(CuraVersion)))
        Logger.log('d', "Info CuraVersion --> " + str(CuraVersion))        
        
        self._qml_folder = "qml_qt6" if not VERSION_QT5 else "qml_qt5"
        
        if "master" in CuraVersion :
            # Master is always a developement version.
            self.Major=4
            self.Minor=20
            
        else:
            try:
                self.Major = int(CuraVersion.split(".")[0])
                self.Minor = int(CuraVersion.split(".")[1])

            except:
                pass
                
        # Thanks to Aldo Hoeben / fieldOfView for this code
        # QFileDialog.Options
        if VERSION_QT5:
            self._dialog_options = QFileDialog.Options()
            if sys.platform == "linux" and "KDE_FULL_SESSION" in os.environ:
                self._dialog_options |= QFileDialog.DontUseNativeDialog
        else:
            self._dialog_options = None

        self.setMenuName(catalog.i18nc("@item:inmenu", "Import/Export Settings"))
        self.addMenuItem(catalog.i18nc("@item:inmenu", "Export Current Settings"), self.exportData)
        self.addMenuItem(catalog.i18nc("@item:inmenu", "Export Current Profile"), self.exportProfile)
        self.addMenuItem("", lambda: None)
        self.addMenuItem(catalog.i18nc("@item:inmenu", "Merge a CSV File"), self.importDataDirect)
        self.addMenuItem(catalog.i18nc("@item:inmenu", "Import Cura Profile"), self.importProfile)
        self.addMenuItem(" ", lambda: None)
        self.addMenuItem(catalog.i18nc("@item:inmenu", "Merge by Step a CSV File"), self.importDataByStep)

    # Return Actual ProfileName
    def profileName(self)->str:
        # Check for Profile Name
        value = ''
        for extruder_stack in CuraApplication.getInstance().getExtruderManager().getActiveExtruderStacks():
            for container in extruder_stack.getContainers():
                # Logger.log("d", "Extruder_stack Type : %s", container.getMetaDataEntry("type") )
                if str(container.getMetaDataEntry("type")) == "quality_changes" :
                    value = container.getName()
        return value
    
    # Export CuraProfile
    def exportProfile(self) -> None:

        _containerRegistry = CuraApplication.getInstance().getContainerRegistry()
        value = self.profileName()
        Logger.log("d", "Attempting to Export ProfileName {}".format(value))
        
        #container_list = [cast(InstanceContainer, _containerRegistry.findContainers(id = quality_changes_group.metadata_for_global["id"])[0])]  # type: List[InstanceContainer]
        #for metadata in quality_changes_group.metadata_per_extruder.values():
        container_list = [] 
        for extruder_stack in CuraApplication.getInstance().getExtruderManager().getActiveExtruderStacks():
            for container in extruder_stack.getContainers():
                if str(container.getMetaDataEntry("type")) == "quality_changes" :
                    if container.getName() != "empty" :
                        container_list.append(cast(InstanceContainer, container))
                    else :
                        Logger.log("d", "Container empty : {}".format(container) )
                    
        Cstack = CuraApplication.getInstance().getGlobalContainerStack()
        for container in Cstack.getContainers():
            if str(container.getMetaDataEntry("type")) == "quality_changes" :
                if container.getName() != "empty" :
                    container_list.append(cast(InstanceContainer, container))
        
        if len(container_list) :
            file_name = ""
            tempo_file_name = self.profileName() + ".curaprofile"
            if VERSION_QT5:
                path = os.path.join(self._preferences.getValue("import_export_tools/dialog_path"), tempo_file_name)
                file_name = QFileDialog.getSaveFileName(
                    parent = None,
                    caption = catalog.i18nc("@title:window", "Save as"),
                    directory = path,
                    filter = catalog.i18nc("@filter", "Cura Profile (*.curaprofile)"),
                    options = self._dialog_options
                )[0]
            else:
                dialog = QFileDialog()
                dialog.setWindowTitle(catalog.i18nc("@title:window", "Save as"))
                dialog.setDirectory(self._preferences.getValue("import_export_tools/dialog_path"))
                dialog.setNameFilters([catalog.i18nc("@filter", "Cura Profile (*.curaprofile)")])
                dialog.setAcceptMode(QFileDialog.AcceptMode.AcceptSave)
                dialog.setFileMode(QFileDialog.FileMode.AnyFile)
                dialog.selectFile(tempo_file_name)
                if dialog.exec():
                    file_name = dialog.selectedFiles()[0]               
                    
            if not file_name:
                Logger.log("d", "No file to export selected")
                return
                
            _containerRegistry.exportQualityProfile(container_list, file_name, catalog.i18nc("@filter", "Cura Profile (*.curaprofile)"))
            self._preferences.setValue("import_export_tools/dialog_path", os.path.dirname(file_name))
            
        else:
            Message().hide()
            Message(catalog.i18nc("@text", "Nothing to export !"), title = catalog.i18nc("@title", "Export Profiles Tools")).show()            
    
    # Export CSV File    
    def exportData(self) -> None:
        # Thanks to Aldo Hoeben / fieldOfView for this part of the code
        file_name = ""
        tempo_file_name = self.profileName() + ".csv"
        if VERSION_QT5:
            path = os.path.join(self._preferences.getValue("import_export_tools/dialog_path"), tempo_file_name)
            file_name = QFileDialog.getSaveFileName(
                parent = None,
                caption = catalog.i18nc("@title:window", "Save as"),
                directory = path,
                filter = catalog.i18nc("@filter", "CSV files (*.csv)"),
                options = self._dialog_options
            )[0]
        else:    
            dialog = QFileDialog()
            dialog.setWindowTitle(catalog.i18nc("@title:window", "Save as"))
            dialog.setDirectory(self._preferences.getValue("import_export_tools/dialog_path"))
            dialog.setNameFilters([catalog.i18nc("@filter", "CSV files (*.csv)")])
            dialog.setAcceptMode(QFileDialog.AcceptMode.AcceptSave)
            dialog.setFileMode(QFileDialog.FileMode.AnyFile)
            dialog.selectFile(tempo_file_name)
            if dialog.exec():
                file_name = dialog.selectedFiles()[0]
                
                
        if not file_name:
            Logger.log("d", "No file to export selected")
            return

        self._preferences.setValue("import_export_tools/dialog_path", os.path.dirname(file_name))
        # -----
        
        machine_manager = CuraApplication.getInstance().getMachineManager()        
        stack = CuraApplication.getInstance().getGlobalContainerStack()

        global_stack = machine_manager.activeMachine

        # Get extruder count
        extruder_count=stack.getProperty("machine_extruder_count", "value")
        
        # for name in sorted(csv.list_dialects()):
        #             Logger.log("d", "Dialect = %s" % name)
        #             dialect = csv.get_dialect(name)
        #             Logger.log("d", "Delimiter = %s" % dialect.delimiter)
        
        exported_count = 0
        try:
            with open(file_name, 'w', newline='') as csv_file:
                # csv.QUOTE_MINIMAL  or csv.QUOTE_NONNUMERIC ?
                csv_writer = csv.writer(csv_file, delimiter=';', quotechar='"', quoting=csv.QUOTE_MINIMAL)
                # E_dialect = csv.get_dialect("excel")
                # csv_writer = csv.writer(csv_file, dialect=E_dialect)
                
                csv_writer.writerow([
                    "Section",
                    "Extruder",
                    "Key",
                    "Type",
                    "Label",
                    "Value"
                ])
                 
                # Date
                self._WriteRow(csv_writer,"general",0,"Date","str","Date",datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
                # Platform
                self._WriteRow(csv_writer,"general",0,"Os","str","Os",str(platform.system()) + " " + str(platform.version())) 
                # Version  
                self._WriteRow(csv_writer,"general",0,"Cura_Version","str","Cura Version",CuraVersion)
                # Profile
                P_Name = global_stack.qualityChanges.getMetaData().get("name", "")
                self._WriteRow(csv_writer,"general",0,"Profile","str","Profile",P_Name)
                # Quality
                Q_Name = global_stack.quality.getMetaData().get("name", "")
                self._WriteRow(csv_writer,"general",0,"Quality","str","Quality",Q_Name)
                # Extruder_Count
                self._WriteRow(csv_writer,"general",0,"Extruder_Count","int","Extruder_Count",str(extruder_count))
                
                # Material
                # extruders = list(global_stack.extruders.values())  
                extruder_stack = CuraApplication.getInstance().getExtruderManager().getActiveExtruderStacks()
 
                # Define every section to get the same order as in the Cura Interface
                # Modification from global_stack to extruders[0]
                i=0
                for Extrud in extruder_stack:    
                    i += 1                        
                    self._doTree(Extrud,"resolution",csv_writer,0,i)
                    # Shell before 4.9 and now Walls
                    self._doTree(Extrud,"shell",csv_writer,0,i)
                    # New section Arachne and 4.9 ?
                    if self.Major > 4 or ( self.Major == 4 and self.Minor >= 9 ) :
                        self._doTree(Extrud,"top_bottom",csv_writer,0,i)
                    self._doTree(Extrud,"infill",csv_writer,0,i)
                    self._doTree(Extrud,"material",csv_writer,0,i)
                    self._doTree(Extrud,"speed",csv_writer,0,i)
                    self._doTree(Extrud,"travel",csv_writer,0,i)
                    self._doTree(Extrud,"cooling",csv_writer,0,i)
                    # If single extruder doesn't export the data
                    if extruder_count>1 :
                        self._doTree(Extrud,"dual",csv_writer,0,i)
                        
                    self._doTree(Extrud,"support",csv_writer,0,i)
                    self._doTree(Extrud,"platform_adhesion",csv_writer,0,i)                   
                    self._doTree(Extrud,"meshfix",csv_writer,0,i)             
                    self._doTree(Extrud,"blackmagic",csv_writer,0,i)
                    self._doTree(Extrud,"experimental",csv_writer,0,i)
                    
                    # Machine_settings
                    # Not Updated by This Plugin
                    # self._doTree(Extrud,"machine_settings",csv_writer,0,i)
                    
        except:
            Logger.logException("e", "Could not export profile to the selected file")
            return

        Message().hide()
        Message(catalog.i18nc("@text", "Exported data for profil %s") % P_Name, title = catalog.i18nc("@title", "Import Export CSV Profiles Tools")).show()

    def _WriteRow(self,csvwriter,Section,Extrud,Key,KType,KeyLbl,ValStr):
        
        csvwriter.writerow([
                     Section,
                     "%d" % Extrud,
                     Key,
                     KType,
                     KeyLbl,
                     str(ValStr)
                ])
               
    def _doTree(self,stack,key,csvwriter,depth,extrud):   
        #output node     
        Pos=0
        if stack.getProperty(key,"type") == "category":
            self._Section=key
        else:
            if stack.getProperty(key,"enabled") == True:
                GetType=stack.getProperty(key,"type")
                GetVal=stack.getProperty(key,"value")
                GetKeyLabl=stack.getProperty(key,"label")
                
                if str(GetType)=='float':
                    # GelValStr="{:.2f}".format(GetVal).replace(".00", "")  # Formatage
                    GelValStr="{:.4f}".format(GetVal).rstrip("0").rstrip(".") # Formatage
                else:
                    # enum = Option list
                    if str(GetType)=='enum':
                        definition_option=key + " option " + str(GetVal)
                        get_option=str(GetVal)
                        GetOption=stack.getProperty(key,"options")
                        GetOptionDetail=GetOption[get_option]
                        GelValStr=str(GetVal)
                    else:
                        GelValStr=str(GetVal)
                
                self._WriteRow(csvwriter,self._Section,extrud,key,str(GetType),str(GetKeyLabl),GelValStr)
                depth += 1

        #look for children
        if len(CuraApplication.getInstance().getGlobalContainerStack().getSettingDefinition(key).children) > 0:
            for i in CuraApplication.getInstance().getGlobalContainerStack().getSettingDefinition(key).children:       
                self._doTree(stack,i.key,csvwriter,depth,extrud)       
 
    def importProfile(self) -> None:
        # 
        file_name = ""
        if VERSION_QT5:
            file_name = QFileDialog.getOpenFileName(
                parent = None,
                caption = catalog.i18nc("@title:window", "Open File"),
                directory = self._preferences.getValue("import_export_tools/dialog_path"),
                filter = catalog.i18nc("@filter", "Cura Profile (*.curaprofile)"),
                options = self._dialog_options
            )[0]
        else:
            dialog = QFileDialog()
            dialog.setWindowTitle(catalog.i18nc("@title:window", "Open File"))
            dialog.setDirectory(self._preferences.getValue("import_export_tools/dialog_path"))
            dialog.setNameFilters([catalog.i18nc("@filter", "Cura Profile (*.curaprofile)"),catalog.i18nc("@filter", "G-Code File (*.gcode)")])
            dialog.setAcceptMode(QFileDialog.AcceptMode.AcceptOpen)
            dialog.setFileMode(QFileDialog.FileMode.ExistingFile)
            if dialog.exec():
                file_name = dialog.selectedFiles()[0]
                
                
        if not file_name:
            Logger.log("d", "No file to import from selected")
            return

        #result = CuraApplication.getInstance().getContainerRegistry().importProfile(file_name)
        result = self.importMyProfile(file_name)
        
        Message(result["message"] , title = catalog.i18nc("@title", "Import Profiles Tools")).show()

    def _getIOPlugins(self, io_type):
        """Gets a list of profile writer plugins

        :return: List of tuples of (plugin_id, meta_data).
        """
        plugin_registry = PluginRegistry.getInstance()
        active_plugin_ids = plugin_registry.getActivePlugins()

        result = []
        for plugin_id in active_plugin_ids:
            meta_data = plugin_registry.getMetaData(plugin_id)
            if io_type in meta_data:
                result.append( (plugin_id, meta_data) )
        return result

    # Original source Code from Ultimaker
    # ContainerManager.py https://github.com/Ultimaker/Cura/blob/main/cura/Settings/ContainerManager.py
    def importMyProfile(self, file_name: str) -> Dict[str, str]:
        """Imports a profile from a file

        :param file_name: The full path and filename of the profile to import.
        :return: Dict with a 'status' key containing the string 'ok', 'warning' or 'error',
            and a 'message' key containing a message for the user.
        """

        Logger.log("d", "Attempting to import profile %s", file_name)
        if not file_name:
            return { "status": "error", "message": i18n_cura_catalog.i18nc("@info:status Don't translate the XML tags <filename>!", "Failed to import profile from <filename>{0}</filename>: {1}", file_name, "Invalid path")}

        global_stack = CuraApplication.getInstance().getGlobalContainerStack()
        if not global_stack:
            return {"status": "error", "message": i18n_cura_catalog.i18nc("@info:status Don't translate the XML tags <filename>!", "Can't import profile from <filename>{0}</filename> before a printer is added.", file_name)}
        container_tree = ContainerTree.getInstance()

        machine_extruders = global_stack.extruderList

        plugin_registry = PluginRegistry.getInstance()
        extension = file_name.split(".")[-1]

        for plugin_id, meta_data in self._getIOPlugins("profile_reader"):
            if meta_data["profile_reader"][0]["extension"] != extension:
                continue
            profile_reader = cast(ProfileReader, plugin_registry.getPluginObject(plugin_id))
            try:
                profile_or_list = profile_reader.read(file_name)  # Try to open the file with the profile reader.
            except NoProfileException:
                return { "status": "ok", "message": i18n_cura_catalog.i18nc("@info:status Don't translate the XML tags <filename>!", "No custom profile to import in file <filename>{0}</filename>", file_name)}
            except Exception as e:
                # Note that this will fail quickly. That is, if any profile reader throws an exception, it will stop reading. It will only continue reading if the reader returned None.
                Logger.log("e", "Failed to import profile from %s: %s while using profile reader. Got exception %s", file_name, profile_reader.getPluginId(), str(e))
                return { "status": "error", "message": i18n_cura_catalog.i18nc("@info:status Don't translate the XML tags <filename>!", "Failed to import profile from <filename>{0}</filename>:", file_name) + "\n<message>" + str(e) + "</message>"}

            if profile_or_list:
                # Ensure it is always a list of profiles
                if not isinstance(profile_or_list, list):
                    profile_or_list = [profile_or_list]

                # First check if this profile is suitable for this machine
                global_profile = None
                extruder_profiles = []
                if len(profile_or_list) == 1:
                    global_profile = profile_or_list[0]
                else:
                    for profile in profile_or_list:
                        if not profile.getMetaDataEntry("position"):
                            global_profile = profile
                        else:
                            extruder_profiles.append(profile)
                extruder_profiles = sorted(extruder_profiles, key = lambda x: int(x.getMetaDataEntry("position", default = "0")))
                profile_or_list = [global_profile] + extruder_profiles

                if not global_profile:
                    Logger.log("e", "Incorrect profile [%s]. Could not find global profile", file_name)
                    return { "status": "error",
                             "message": i18n_cura_catalog.i18nc("@info:status Don't translate the XML tags <filename>!", "This profile <filename>{0}</filename> contains incorrect data, could not import it.", file_name)}
                profile_definition = global_profile.getMetaDataEntry("definition")

                # Make sure we have a profile_definition in the file:
                if profile_definition is None:
                    break
                
                # Logger.log("d", "Profile_definition {}".format(profile_definition))
                _containerRegistry = CuraApplication.getInstance().getContainerRegistry() #ContainerRegistry()  # type: ContainerRegistryInterface
                machine_definitions = _containerRegistry.findContainers(id = profile_definition)
                if not machine_definitions:
                    Logger.log("e", "Incorrect profile [%s]. Unknown machine type [%s]", file_name, profile_definition)
                    return {"status": "error",
                            "message": i18n_cura_catalog.i18nc("@info:status Don't translate the XML tags <filename>!", "This profile <filename>{0}</filename> contains incorrect data, could not import it.", file_name)
                            }
                machine_definition = machine_definitions[0]

                # Get the expected machine definition.
                # i.e.: We expect gcode for a UM2 Extended to be defined as normal UM2 gcode...
                has_machine_quality = parseBool(machine_definition.getMetaDataEntry("has_machine_quality", "false"))
                profile_definition = machine_definition.getMetaDataEntry("quality_definition", machine_definition.getId()) if has_machine_quality else "fdmprinter"
                expected_machine_definition = container_tree.machines[global_stack.definition.getId()].quality_definition

                # And check if the profile_definition matches either one (showing error if not):
                if profile_definition != expected_machine_definition:
                    Logger.log("d", "Profile {file_name} is for machine {profile_definition}, but the current active machine is {expected_machine_definition}. Changing profile's definition.".format(file_name = file_name, profile_definition = profile_definition, expected_machine_definition = expected_machine_definition))
                    global_profile.setMetaDataEntry("definition", expected_machine_definition)
                    for extruder_profile in extruder_profiles:
                        extruder_profile.setMetaDataEntry("definition", expected_machine_definition)

                quality_name = global_profile.getName()
                quality_type = global_profile.getMetaDataEntry("quality_type")

                name_seed = os.path.splitext(os.path.basename(file_name))[0]
                new_name = _containerRegistry.uniqueName(name_seed)

                # Ensure it is always a list of profiles
                if type(profile_or_list) is not list:
                    profile_or_list = [profile_or_list]

                # Make sure that there are also extruder stacks' quality_changes, not just one for the global stack
                if len(profile_or_list) == 1:
                    global_profile = profile_or_list[0]
                    extruder_profiles = []
                    for idx, extruder in enumerate(global_stack.extruderList):
                        profile_id = ContainerRegistry.getInstance().uniqueName(global_stack.getId() + "_extruder_" + str(idx + 1))
                        profile = InstanceContainer(profile_id)
                        profile.setName(quality_name)
                        profile.setMetaDataEntry("setting_version", CuraApplication.SettingVersion)
                        profile.setMetaDataEntry("type", "quality_changes")
                        profile.setMetaDataEntry("definition", expected_machine_definition)
                        profile.setMetaDataEntry("quality_type", quality_type)
                        profile.setDirty(True)
                        if idx == 0:
                            # Move all per-extruder settings to the first extruder's quality_changes
                            for qc_setting_key in global_profile.getAllKeys():
                                settable_per_extruder = global_stack.getProperty(qc_setting_key, "settable_per_extruder")
                                if settable_per_extruder:
                                    setting_value = global_profile.getProperty(qc_setting_key, "value")

                                    setting_definition = global_stack.getSettingDefinition(qc_setting_key)
                                    if setting_definition is not None:
                                        new_instance = SettingInstance(setting_definition, profile)
                                        new_instance.setProperty("value", setting_value)
                                        new_instance.resetState()  # Ensure that the state is not seen as a user state.
                                        profile.addInstance(new_instance)
                                        profile.setDirty(True)

                                    global_profile.removeInstance(qc_setting_key, postpone_emit = True)
                        extruder_profiles.append(profile)

                    for profile in extruder_profiles:
                        profile_or_list.append(profile)

                # Import all profiles
                profile_ids_added = []  # type: List[str]
                additional_message = None
                for profile_index, profile in enumerate(profile_or_list):
                    if profile_index == 0:
                        # This is assumed to be the global profile
                        profile_id = (cast(ContainerInterface, global_stack.getBottom()).getId() + "_" + name_seed).lower().replace(" ", "_")

                    elif profile_index < len(machine_extruders) + 1:
                        # This is assumed to be an extruder profile
                        extruder_id = machine_extruders[profile_index - 1].definition.getId()
                        extruder_position = str(profile_index - 1)
                        if not profile.getMetaDataEntry("position"):
                            profile.setMetaDataEntry("position", extruder_position)
                        else:
                            profile.setMetaDataEntry("position", extruder_position)
                        profile_id = (extruder_id + "_" + name_seed).lower().replace(" ", "_")

                    else:  # More extruders in the imported file than in the machine.
                        continue  # Delete the additional profiles.
                    
                    available_quality_groups_dict = {name: quality_group for name, quality_group in ContainerTree.getInstance().getCurrentQualityGroups().items() if quality_group.is_available}
                    all_quality_groups_dict = ContainerTree.getInstance().getCurrentQualityGroups()
                    
                    quality_type = profile.getMetaDataEntry("quality_type")
                    quality_message = ''
                    if quality_type not in available_quality_groups_dict:
                        
                        # Logger.log("d", "quality_type {}".format(quality_type))
                        # Logger.log("d", "available_quality_groups_dict {} / {}".format(available_quality_groups_dict, all_quality_groups_dict))
                        mode ="standard"
                        Cstack = CuraApplication.getInstance().getGlobalContainerStack()
                        for container in Cstack.getContainers():                          
                            if str(container.getMetaDataEntry("type")) == "quality" :
                                # Logger.log("d", "Container : {}".format(container.getMetaDataEntry("quality_type")) )
                                if container.getMetaDataEntry("quality_type") != "empty" :
                                    mode = container.getMetaDataEntry("quality_type")  
                                else:
                                    mode ="standard"
                        
                        Logger.log("d", "Profile {file_name} is for quality {quality_type}, changed to {mode}. Changing profile's definition.".format(file_name = file_name, quality_type = quality_type, mode = mode))
                        profile.setMetaDataEntry("quality_type", mode)
                    
                        quality_message = catalog.i18nc("@info:status", "\nWarning: The profile have been switch from the quality '{}' to the Quality '{}'".format(quality_type, mode))

                    # This function return the message 
                    # catalog.i18nc("@info:status", "Warning: The profile is not visible because its quality type '{0}' is not available for the current configuration. Switch to a material/nozzle combination that can use this quality type.", quality_type)
                    configuration_successful, message = _containerRegistry._configureProfile(profile, profile_id, new_name, expected_machine_definition)
                    
                    if quality_message :
                        if message == None :
                            message = quality_message
                        else :
                            message += quality_message 
                    
                    if configuration_successful:
                        additional_message = message
                    else:
                        # Remove any profiles that were added.
                        for profile_id in profile_ids_added + [profile.getId()]:
                            _containerRegistry.removeContainer(profile_id)
                        if not message:
                            message = ""
                        return {"status": "error", "message": i18n_cura_catalog.i18nc(
                                "@info:status Don't translate the XML tag <filename>!",
                                "Failed to import profile from <filename>{0}</filename>:",
                                file_name) + " " + message}
                    profile_ids_added.append(profile.getId())
                result_status = "ok"
                success_message = i18n_cura_catalog.i18nc("@info:status", "Successfully imported profile {0}.", profile_or_list[0].getName())
                if additional_message:
                    result_status = "warning"
                    success_message += additional_message
                return {"status": result_status, "message": success_message}

            # This message is throw when the profile reader doesn't find any profile in the file
            return {"status": "error", "message": i18n_cura_catalog.i18nc("@info:status", "File {0} does not contain any valid profile.", file_name)}

        # If it hasn't returned by now, none of the plugins loaded the profile successfully.
        return {"status": "error", "message": i18n_cura_catalog.i18nc("@info:status", "Profile {0} has an unknown file type or is corrupted.", file_name)}
    
    def importDataDirect(self) -> None:
        self.importData(False)
        
    def importDataByStep(self) -> None:
        self.importData(True)
        
    # Import CSV file
    def importData(self, byStep: bool) -> None:
        # thanks to Aldo Hoeben / fieldOfView for this part of the code
        file_name = ""
        if VERSION_QT5:
            file_name = QFileDialog.getOpenFileName(
                parent = None,
                caption = catalog.i18nc("@title:window", "Open File"),
                directory = self._preferences.getValue("import_export_tools/dialog_path"),
                filter = catalog.i18nc("@filter", "CSV files (*.csv)"),
                options = self._dialog_options
            )[0]
        else:
            dialog = QFileDialog()
            dialog.setWindowTitle(catalog.i18nc("@title:window", "Open File"))
            dialog.setDirectory(self._preferences.getValue("import_export_tools/dialog_path"))
            dialog.setNameFilters(["CSV files (*.csv)"])
            dialog.setAcceptMode(QFileDialog.AcceptMode.AcceptOpen)
            dialog.setFileMode(QFileDialog.FileMode.ExistingFile)
            if dialog.exec():
                file_name = dialog.selectedFiles()[0]
                
                
        if not file_name:
            Logger.log("d", "No file to import from selected")
            return

        self._preferences.setValue("import_export_tools/dialog_path", os.path.dirname(file_name))
        # -----
        
        machine_manager = CuraApplication.getInstance().getMachineManager()        
        stack = CuraApplication.getInstance().getGlobalContainerStack()
        global_stack = machine_manager.activeMachine

        #Get extruder count
        extruder_count=stack.getProperty("machine_extruder_count", "value")
        
        #extruders = list(global_stack.extruders.values())   
        extruder_stack = CuraApplication.getInstance().getExtruderManager().getActiveExtruderStacks()
        
        imported_count = 0
        CPro = ""
        try:
            with open(file_name, 'r', newline='') as csv_file:
                C_dialect = csv.Sniffer().sniff(csv_file.read(1024))
                # Reset to begining file position
                csv_file.seek(0, 0)
                Logger.log("d", "Csv Import %s : Delimiter = %s Quotechar = %s ByStep = %s", file_name, C_dialect.delimiter, C_dialect.quotechar , byStep)
                # csv.QUOTE_MINIMAL  or csv.QUOTE_NONNUMERIC ?
                # csv_reader = csv.reader(csv_file, delimiter=';', quotechar='"', quoting=csv.QUOTE_MINIMAL)
                csv_reader = csv.reader(csv_file, dialect=C_dialect)
                line_number = -1
                for row in csv_reader:
                    line_number += 1
                    if line_number == 0:
                        if len(row) < 5:
                            continue         
                    else:
                        # Logger.log("d", "Import Data = %s | %s | %s | %s | %s | %s",row[0], row[1], row[2], row[3], row[4], row[5])
                        try:
                            #(section, extrud, kkey, ktype, kvalue) = row[0:5]
                            section=row[0]
                            extrud=int(row[1])
                            extrud -= 1
                            kkey=row[2]
                            ktype=row[3]
                            klbl=row[4]
                            kvalue=row[5]
                            
                            # Logger.log("d", "Current Data = %s | %d | %s | %s | %s | %s", section,extrud, kkey, ktype, klbl, kvalue) 
                            update_setting = True                            
                            if extrud<extruder_count:
                                try:
                                    container=extruder_stack[extrud]
                                    try:
                                        prop_value = container.getProperty(kkey, "value")
                                        if prop_value != None :
                                            
                                            settable_per_extruder= container.getProperty(kkey, "settable_per_extruder")
                                            # Logger.log("d", "%s settable_per_extruder : %s", kkey, str(settable_per_extruder))
                                            
                                            if ktype == "str" or ktype == "enum":
                                                if prop_value != kvalue :
                                                    if extrud == 0 :
                                                        if byStep :
                                                            update_setting = self.changeValue(klbl)
                                                        if update_setting :
                                                            stack.setProperty(kkey,"value",kvalue)
                                                            Logger.log("d", "prop_value changed: %s = %s / %s", kkey ,kvalue, prop_value)
                                                            
                                                    if settable_per_extruder == True : 
                                                        if byStep or update_setting == False :
                                                            update_setting = self.changeValue(klbl)
                                                        if update_setting :
                                                            container.setProperty(kkey,"value",kvalue)
                                                            Logger.log("d", "prop_value per extruder changed: %s = %s / %s", kkey ,kvalue, prop_value)
                                                    else:
                                                        Logger.log("d", "%s not settable_per_extruder", kkey)
                                                    imported_count += 1
                                                    
                                            elif ktype == "bool" :
                                                if kvalue == "True" or kvalue == "true" :
                                                    C_bool=True
                                                else:
                                                    C_bool=False
                                                
                                                if prop_value != C_bool :
                                                    if extrud == 0 :
                                                        if byStep :
                                                            update_setting = self.changeValue(klbl)
                                                        if update_setting :
                                                            stack.setProperty(kkey,"value",C_bool)
                                                            Logger.log("d", "prop_value changed: %s = %s / %s", kkey ,C_bool, prop_value)
                                                        
                                                    if settable_per_extruder == True : 
                                                        if byStep or update_setting == False :
                                                            update_setting = self.changeValue(klbl)
                                                        if update_setting :
                                                            container.setProperty(kkey,"value",C_bool)
                                                            Logger.log("d", "prop_value per extruder changed: %s = %s / %s", kkey ,C_bool, prop_value)                                                       
                                                        
                                                    else:
                                                        Logger.log("d", "%s not settable_per_extruder", kkey)
                                                    imported_count += 1
                                                    
                                            elif ktype == "int" :
                                                if prop_value != int(kvalue) :
                                                    if extrud == 0 :
                                                        if byStep :
                                                            update_setting = self.changeValue(klbl)
                                                        if update_setting :
                                                            stack.setProperty(kkey,"value",int(kvalue))
                                                            Logger.log("d", "prop_value changed: %s = %s / %s", kkey ,kvalue, prop_value)
                                                        
                                                    if settable_per_extruder == True :
                                                        if byStep or update_setting == False :
                                                            update_setting = self.changeValue(klbl)
                                                        if update_setting :
                                                            container.setProperty(kkey,"value",int(kvalue))
                                                            Logger.log("d", "prop_value per extruder changed: %s = %s / %s", kkey ,int(kvalue), prop_value)
                                                    else:
                                                        Logger.log("d", "%s not settable_per_extruder", kkey)
                                                    imported_count += 1
                                            
                                            elif ktype == "float" :
                                                TransVal=round(float(kvalue),4)
                                                if round(prop_value,4) != TransVal :
                                                    if extrud == 0 :
                                                        if byStep :
                                                            update_setting = self.changeValue(klbl)
                                                        if update_setting :
                                                            stack.setProperty(kkey,"value",TransVal)
                                                            Logger.log("d", "prop_value changed: %s = %s / %s", kkey ,TransVal, round(prop_value,4))
                                                            
                                                    if settable_per_extruder == True : 
                                                        if byStep or update_setting == False :
                                                            update_setting = self.changeValue(klbl)
                                                        if update_setting :
                                                            container.setProperty(kkey,"value",TransVal)
                                                            Logger.log("d", "prop_value per extruder changed: %s = %s / %s", kkey ,TransVal, prop_value)
                                                    else:
                                                        Logger.log("d", "%s not settable_per_extruder", kkey)
                                                        
                                                    imported_count += 1
                                            else :
                                                # Case of the tables                                              
                                                try:
                                                    container.setProperty(kkey,"value",kvalue)
                                                    if byStep :
                                                        update_setting = self.changeValue(klbl)
                                                  
                                                    Logger.log("d", "prop_value changed: %s = %s / %s", kkey ,kvalue, prop_value)
                                                except:
                                                    Logger.log("d", "Value type Else = %d | %s | %s | %s",extrud, kkey, ktype, kvalue)
                                                    continue
                                                 
                                        else:
                                            # Logger.log("d", "Current Data = %s | %d | %s | %s | %s | %s", section,extrud, kkey, ktype, klbl, kvalue) 
                                            if kkey=="Profile" :
                                                CPro=kvalue                                               
                                    except:
                                        Logger.log("e", "Error kkey: %s" % kkey)
                                        continue                                       
                                except:
                                    Logger.log("e", "Error Extruder: %s" % row)
                                    continue                             
                        except:
                            Logger.log("e", "Row does not have enough data: %s" % row)
                            continue
                            
        except:
            Logger.logException("e", "Could not import settings from the selected file")
            return

        Message().hide()
        Message(catalog.i18nc("@text", "Imported profil : %d changed keys from %s") % (imported_count, CPro) , title = catalog.i18nc("@title", "Import Export CSV Profiles Tools")).show()

    def changeValue(self, lblkey) -> bool:
        # Logger.logException("d", "In ChangeValue")
        validValue = False 
        dialog = self._createConfirmationDialog(lblkey)

        returnValue = dialog.exec()
        
        if VERSION_QT5:
            if returnValue == QMessageBox.Ok:
                validValue = True
            else:
                validValue = False
        else:
            if returnValue == QMessageBox.StandardButton.Ok:
                validValue = True
            else:
                validValue = False               
            
        Logger.log("d", "validValue : %s : %s", lblkey ,validValue)
            
        if validValue:
            self._application.backend.forceSlice()
            self._application.backend.slice()
        
        return validValue

    def _createConfirmationDialog(self, lblkey):
        '''Create a message box prompting the user if they want to update parameter.'''
        msgBox = QMessageBox()
        msgBox.setIcon(QMessageBox.Information if VERSION_QT5 else QMessageBox.Icon.Information)
        msgBox.setText(catalog.i18nc("@text", "Would you like to update the slice for : %s ?") % (lblkey) )
        msgBox.setWindowTitle(catalog.i18nc("@title", "Update slice"))
        msgBox.setStandardButtons((QMessageBox.Ok if VERSION_QT5 else QMessageBox.StandardButton.Ok) | (QMessageBox.Cancel if VERSION_QT5 else QMessageBox.StandardButton.Cancel))
        msgBox.setDefaultButton(QMessageBox.Ok if VERSION_QT5 else QMessageBox.StandardButton.Ok)

        return msgBox
        
