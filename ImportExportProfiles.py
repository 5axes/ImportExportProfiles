# Copyright (c) 2020 5@xes
# ImportExportProfiles is released under the terms of the AGPLv3 or higher.

from PyQt5.QtCore import QObject
from PyQt5.QtWidgets import QFileDialog, QMessageBox

import os.path
import sys
import json
import re
from uuid import UUID
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
from UM.Settings.ContainerRegistry import ContainerRegistry

use_container_tree = True
try:
    from cura.Machines.ContainerTree import ContainerTree
except ImportError:
    use_container_tree = False

from UM.i18n import i18nCatalog
catalog = i18nCatalog("cura")

from typing import List, Dict, Any

class ImportExportProfiles(Extension, QObject,):
    def __init__(self, parent = None) -> None:
        QObject.__init__(self, parent)
        Extension.__init__(self)

        self._application = Application.getInstance()
        self._preferences = self._application.getPreferences()
        self._preferences.addPreference("import_export_tools/dialog_path", "")

        self._message = Message()

        self._dialog_options = QFileDialog.Options()
        if sys.platform == "linux" and "KDE_FULL_SESSION" in os.environ:
            self._dialog_options |= QFileDialog.DontUseNativeDialog

        self.setMenuName(catalog.i18nc("@item:inmenu", "Import Export Profiles"))
        self.addMenuItem(catalog.i18nc("@item:inmenu", "Export current profile..."), self.exportProfilData)
        self.addMenuItem("", lambda: None)
        self.addMenuItem(catalog.i18nc("@item:inmenu", "Import a profile..."), self.importData)


    def exportProfilData(self):
        global_stack = self._application.getGlobalContainerStack()
        if not global_stack or not global_stack.getMetaDataEntry("has_materials", False):
            return
        extruder_stack = global_stack.extruders.get("0")
        if not extruder_stack:
            return

        approximate_material_diameter = extruder_stack.getApproximateMaterialDiameter()

        if use_container_tree:
            nozzle_name = extruder_stack.variant.getName()
            machine_node = ContainerTree.getInstance().machines[global_stack.definition.getId()]
            if nozzle_name not in machine_node.variants:
                Logger.log("w", "Unable to find variant %s in container tree", nozzle_name)
                return

            material_nodes = machine_node.variants[nozzle_name].materials
            materials_metadata = [
                m.getMetadata() for m in material_nodes.values()
                if float(m.getMetaDataEntry("approximate_diameter", -1)) == approximate_material_diameter
            ]
        else:
            materials_metadata = [
                m for m in ContainerRegistry.getInstance().findInstanceContainersMetadata(type = "material")
                if "base_file" in m and m.get("approximate_diameter", -1) == approximate_material_diameter and m["id"] == m["base_file"]
            ]

        self._exportData(materials_metadata)


    def _exportData(self, materials_metadata: List[Dict[str, Any]]) -> None:
        try:
            material_settings = json.loads(self._preferences.getValue("cura/material_settings"))
        except:
            Logger.logException("e", "Could not load material settings from preferences")
            return

        file_name = QFileDialog.getSaveFileName(
            parent = None,
            caption = catalog.i18nc("@title:window", "Save as"),
            directory = self._preferences.getValue("import_export_tools/dialog_path"),
            filter = "CSV files (*.csv)",
            options = self._dialog_options
        )[0]

        if not file_name:
            Logger.log("d", "No file to export to selected")
            return

        self._preferences.setValue("import_export_tools/dialog_path", os.path.dirname(file_name))

        materials_metadata = [
            {
                "guid": m["GUID"],
                "material": m["material"],
                "brand": m.get("brand",""),
                "name": m["name"],
                "spool_weight": material_settings.get(m["GUID"], {}).get("spool_weight", ""),
                "spool_cost": material_settings.get(m["GUID"], {}).get("spool_cost", "")
            }
            for m in materials_metadata
            if "brand" in m
        ]
        materials_metadata.sort(key = lambda k: (k["brand"], k["material"], k["name"]))

        exported_count = 0
        try:
            with open(file_name, 'w', newline='') as csv_file:
                csv_writer = csv.writer(csv_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
                csv_writer.writerow([
                    "guid",
                    "name",
                    "weight (g)",
                    "cost (%s)" % self._preferences.getValue("cura/currency")
                ])

                for material in materials_metadata:
                    try:
                        csv_writer.writerow([
                            material["guid"],
                            "%s %s" % (material["brand"], material["name"]),
                            material["spool_weight"],
                            material["spool_cost"]
                        ])
                        exported_count += 1
                    except:
                        continue
        except:
            Logger.logException("e", "Could not export settings to the selected file")
            return

        self._message.hide()
        self._message = Message(
            catalog.i18ncp(
                "@info:status {0} is count", "Exported data for {0} material.", "Exported data for {0} materials.", exported_count
            ).format(exported_count),
            title=catalog.i18nc("@info:title", "Material Cost Tools")
        )
        self._message.show()


    def importData(self) -> None:
        file_name = QFileDialog.getOpenFileName(
            parent = None,
            caption = catalog.i18nc("@title:window", "Open File"),
            directory = self._preferences.getValue("import_export_tools/dialog_path"),
            filter = "CSV files (*.csv)",
            options = self._dialog_options
        )[0]

        if not file_name:
            Logger.log("d", "No file to import from selected")
            return

        self._preferences.setValue("import_export_tools/dialog_path", os.path.dirname(file_name))

        try:
            material_settings = json.loads(self._preferences.getValue("cura/material_settings"))
        except:
            Logger.logException("e", "Could not load material settings from preferences")
            return

        imported_count = 0
        try:
            with open(file_name, 'r', newline='') as csv_file:
                csv_reader = csv.reader(csv_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
                line_number = -1
                for row in csv_reader:
                    line_number += 1
                    if line_number == 0:
                        if len(row) < 4:
                            continue
                        match = re.search("cost\s\((.*)\)", row[3])
                        if not match:
                            continue

                        currency = match.group(1)

                        if currency != self._preferences.getValue("cura/currency"):

                            result = QMessageBox.question(
                                None,
                                catalog.i18nc("@title:window", "Import weights and prices"),
                                catalog.i18nc("@label",
                                    "The file contains prices specified in %s, but your Cura is configured to use %s.\nAre you sure you want to import these prices as is?" % (
                                        currency, self._preferences.getValue("cura/currency")
                                    )
                                )
                            )

                            if result == QMessageBox.No:
                                return
                    else:
                        try:
                            (guid, name, weight, cost) = row[0:4]
                        except:
                            Logger.log("e", "Row does not have enough data: %s" % row)
                            continue

                        try:
                            uuid = UUID(guid)
                        except:
                            Logger.log("e", "UUID is malformed: %s" % row)
                            continue

                        data = {}
                        try:
                             data["spool_cost"] = float(cost)
                        except:
                            pass
                        try:
                            data["spool_weight"] = int(weight)
                        except:
                            pass
                        if data:
                            material_settings[guid] = data
                            imported_count += 1
        except:
            Logger.logException("e", "Could not import settings from the selected file")
            return

        self._preferences.setValue("cura/material_settings", json.dumps(material_settings))

        self._message.hide()
        self._message = Message(
            catalog.i18ncp(
                "@info:status {0} is count", "Imported weight & price for {0} material.", "Imported weights & prices for {0} materials.", imported_count
            ).format(imported_count),
            title=catalog.i18nc("@info:title", "Material Cost Tools")
        )
        self._message.show()


