"""Portable Inventor control panel — pywebview shell over core.py.

Build a standalone exe with:
    pyinstaller --onefile --windowed --name InventorPult --add-data "ui;ui" webview_app.py

Run directly for development:
    python webview_app.py
"""
import os
import sys

import webview

import core


def _resource_path(*parts):
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, *parts)


class Api:
    def get_status(self):
        return core.get_status()

    def list_documents(self):
        return core.list_documents()

    def list_parameters(self):
        return core.list_parameters()

    def set_parameter(self, name, expression):
        return core.set_parameter(name, expression)

    def list_ilogic_rules(self):
        return core.list_ilogic_rules()

    def run_ilogic_rule(self, rule_name):
        return core.run_ilogic_rule(rule_name)

    def save_document(self):
        return core.save_document()

    def open_document_dialog(self):
        result = window.create_file_dialog(
            webview.OPEN_DIALOG,
            file_types=("Inventor files (*.ipt;*.iam;*.idw;*.dwg)", "All files (*.*)"),
        )
        if not result:
            return None
        return core.open_document(result[0])

    def list_materials(self):
        return core.list_materials()

    def suggest_designation(self, level, product, module):
        return core.suggest_designation(
            level,
            int(product) if product else None,
            int(module) if module else None,
        )

    def designation_exists(self, designation):
        return core.designation_exists(designation)

    def create_assembly(self, designation, description):
        return core.create_assembly(designation, description)

    def list_known_products(self):
        return core.list_known_products()

    def list_known_modules(self, product):
        return core.list_known_modules(product) if product else []

    def list_tube_presets(self):
        return core.list_tube_presets()

    def list_sheet_thicknesses(self):
        return core.list_sheet_thicknesses()

    def list_part_name_templates(self):
        return core.list_part_name_templates()

    def open_catalog(self):
        return core.open_catalog()

    def create_sheet_part(self, designation, length_mm, width_mm, thickness_mm, description, material_name):
        return core.create_sheet_part(
            designation, float(length_mm), float(width_mm), float(thickness_mm),
            description, material_name or None)

    def create_tube_part(self, designation, outer_width_mm, outer_height_mm, wall_mm, length_mm,
                          description, material_name):
        return core.create_tube_part(
            designation, float(outer_width_mm), float(outer_height_mm), float(wall_mm), float(length_mm),
            description, material_name or None)


api = Api()
window = webview.create_window(
    "Inventor Pult", _resource_path("ui", "index.html"), js_api=api,
    width=880, height=780, min_size=(720, 520), resizable=True, background_color="#05060f")

if __name__ == "__main__":
    webview.start()
