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
# Version 1.2.0 : Integrate French Translation
#
#-------------------------------------------------------------------------------------------


VERSION_QT5 = False
try:
    from PyQt6.QtCore import QObject
    from PyQt6.QtWidgets import QFileDialog, QMessageBox
except ImportError:
    from PyQt5.QtCore import QObject
    from PyQt5.QtWidgets import QFileDialog, QMessageBox
    VERSION_QT5 = True
    
    
import os
import platform
import os.path
import sys
import re

from datetime import datetime
# from typing import cast, Dict, List, Optional, Tuple, Any, Set
from cura.CuraApplication import CuraApplication
from cura.CuraVersion import CuraVersion  # type: ignore
from UM.Version import Version

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

from UM.i18n import i18nCatalog
from UM.Resources import Resources

Resources.addSearchPath(
	os.path.join(os.path.abspath(os.path.dirname(__file__)))
)  # Plugin translation file import

catalog = i18nCatalog("ImportExportProfiles")

if catalog.hasTranslationLoaded():
	Logger.log("i", "ImportExportProfiles Plugin translation loaded!")


class ImportExportProfiles(Extension, QObject,):
    def __init__(self, parent = None) -> None:
        QObject.__init__(self, parent)
        Extension.__init__(self)
        
        self._Section =""

        self._application = Application.getInstance()
        self._preferences = self._application.getPreferences()
        self._preferences.addPreference("import_export_tools/dialog_path", "")
        
        self.Major=1
        self.Minor=0

        # Test version for futur release 4.9
        # Logger.log('d', "Info Version CuraVersion --> " + str(Version(CuraVersion)))
        Logger.log('d', "Info CuraVersion --> " + str(CuraVersion))        
        
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

        self.setMenuName(catalog.i18nc("@item:inmenu", "Import Export Profiles"))
        self.addMenuItem(catalog.i18nc("@item:inmenu", "Export current profile"), self.exportData)
        self.addMenuItem("", lambda: None)
        self.addMenuItem(catalog.i18nc("@item:inmenu", "Merge a profile"), self.importData)


    def exportData(self) -> None:
        # thanks to Aldo Hoeben / fieldOfView for this part of the code
        file_name = ""
        if VERSION_QT5:
            file_name = QFileDialog.getSaveFileName(
                parent = None,
                caption = catalog.i18nc("@title:window", "Save as"),
                directory = self._preferences.getValue("import_export_tools/dialog_path"),
                filter = "CSV files (*.csv)",
                options = self._dialog_options
            )[0]
        else:
            dialog = QFileDialog()
            dialog.setWindowTitle(catalog.i18nc("@title:window", "Save as"))
            dialog.setDirectory(self._preferences.getValue("import_export_tools/dialog_path"))
            dialog.setNameFilters(["CSV files (*.csv)"])
            dialog.setAcceptMode(QFileDialog.AcceptMode.AcceptSave)
            dialog.setFileMode(QFileDialog.FileMode.AnyFile)
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
                    "Value"
                ])
                 
                # Date
                self._WriteRow(csv_writer,"general",0,"Date","str",datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
                # Platform
                self._WriteRow(csv_writer,"general",0,"Os","str",str(platform.system()) + " " + str(platform.version())) 
                # Version  
                self._WriteRow(csv_writer,"general",0,"Cura_Version","str",CuraVersion)
                # Profile
                P_Name = global_stack.qualityChanges.getMetaData().get("name", "")
                self._WriteRow(csv_writer,"general",0,"Profile","str",P_Name)
                # Quality
                Q_Name = global_stack.quality.getMetaData().get("name", "")
                self._WriteRow(csv_writer,"general",0,"Quality","str",Q_Name)
                # Extruder_Count
                self._WriteRow(csv_writer,"general",0,"Extruder_Count","int",str(extruder_count))
                
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
                    
                    # machine_settings
                    # self._doTree(Extrud,"machine_settings",csv_writer,0,i)
                    
        except:
            Logger.logException("e", "Could not export profile to the selected file")
            return

        Message().hide()
        Message("Exported data for profil %s" % P_Name, title = "Import Export CSV Profiles Tools").show()

    def _WriteRow(self,csvwriter,Section,Extrud,Key,KType,ValStr):
        
        csvwriter.writerow([
                     Section,
                     "%d" % Extrud,
                     Key,
                     KType,
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
                        # Logger.log("d", "GetType_doTree = %s ; %s ; %s ; %s",definition_option, GelValStr, GetOption, GetOptionDetail)
                    else:
                        GelValStr=str(GetVal)
                
                self._WriteRow(csvwriter,self._Section,extrud,key,str(GetType),GelValStr)
                depth += 1

        #look for children
        if len(CuraApplication.getInstance().getGlobalContainerStack().getSettingDefinition(key).children) > 0:
            for i in CuraApplication.getInstance().getGlobalContainerStack().getSettingDefinition(key).children:       
                self._doTree(stack,i.key,csvwriter,depth,extrud)       
                
    def importData(self) -> None:
        # thanks to Aldo Hoeben / fieldOfView for this part of the code
        file_name = ""
        if VERSION_QT5:
            file_name = QFileDialog.getOpenFileName(
                parent = None,
                caption = catalog.i18nc("@title:window", "Open File"),
                directory = self._preferences.getValue("import_export_tools/dialog_path"),
                filter = "CSV files (*.csv)",
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
                Logger.log("d", "Csv Import %s : Delimiter = %s Quotechar = %s", file_name, C_dialect.delimiter, C_dialect.quotechar)
                # csv.QUOTE_MINIMAL  or csv.QUOTE_NONNUMERIC ?
                # csv_reader = csv.reader(csv_file, delimiter=';', quotechar='"', quoting=csv.QUOTE_MINIMAL)
                csv_reader = csv.reader(csv_file, dialect=C_dialect)
                line_number = -1
                for row in csv_reader:
                    line_number += 1
                    if line_number == 0:
                        if len(row) < 4:
                            continue         
                    else:
                        # Logger.log("d", "Import Data = %s | %s | %s | %s | %s",row[0], row[1], row[2], row[3], row[4])
                        try:
                            #(section, extrud, kkey, ktype, kvalue) = row[0:4]
                            section=row[0]
                            extrud=int(row[1])
                            extrud -= 1
                            kkey=row[2]
                            ktype=row[3]
                            kvalue=row[4]
                            
                            #Logger.log("d", "Current Data = %s | %d | %s | %s | %s", section,extrud, kkey, ktype, kvalue)  
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
                                                    if extrud == 0 : stack.setProperty(kkey,"value",kvalue)
                                                    if settable_per_extruder == True : 
                                                        container.setProperty(kkey,"value",kvalue)
                                                        Logger.log("d", "prop_value changed: %s = %s / %s", kkey ,kvalue, prop_value)
                                                    else:
                                                        Logger.log("d", "%s not settable_per_extruder", kkey)
                                                    imported_count += 1
                                                    
                                            elif ktype == "bool" :
                                                if kvalue == "True" or kvalue == "true" :
                                                    C_bool=True
                                                else:
                                                    C_bool=False
                                                
                                                if prop_value != C_bool :
                                                    if extrud == 0 : stack.setProperty(kkey,"value",C_bool)
                                                    if settable_per_extruder == True : 
                                                        container.setProperty(kkey,"value",C_bool)
                                                        Logger.log("d", "prop_value changed: %s = %s / %s", kkey ,C_bool, prop_value)
                                                    else:
                                                        Logger.log("d", "%s not settable_per_extruder", kkey)
                                                    imported_count += 1
                                                    
                                            elif ktype == "int" :
                                                if prop_value != int(kvalue) :
                                                    if extrud == 0 : stack.setProperty(kkey,"value",int(kvalue))
                                                    if settable_per_extruder == True :
                                                        container.setProperty(kkey,"value",int(kvalue))
                                                        Logger.log("d", "prop_value changed: %s = %s / %s", kkey ,kvalue, prop_value)
                                                    else:
                                                        Logger.log("d", "%s not settable_per_extruder", kkey)
                                                    imported_count += 1
                                            
                                            elif ktype == "float" :
                                                TransVal=round(float(kvalue),4)
                                                if round(prop_value,4) != TransVal :
                                                    if extrud == 0 : stack.setProperty(kkey,"value",TransVal)
                                                    if settable_per_extruder == True : 
                                                        container.setProperty(kkey,"value",TransVal)
                                                        Logger.log("d", "prop_value changed: %s = %s / %s", kkey ,TransVal, prop_value)
                                                    else:
                                                        Logger.log("d", "%s not settable_per_extruder", kkey)
                                                    imported_count += 1
                                            else :
                                                Logger.log("d", "Value type Else = %d | %s | %s | %s",extrud, kkey, ktype, kvalue)
                                        else:
                                            # Logger.log("d", "Value None = %d | %s | %s | %s",extrud, kkey, ktype, kvalue)
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
        Message("Imported profil %d changed keys from %s" % (imported_count, CPro) , title = "Import Export CSV Profiles Tools").show()


