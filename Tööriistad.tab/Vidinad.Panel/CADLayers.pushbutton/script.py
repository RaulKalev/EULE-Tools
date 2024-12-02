# -*- coding: utf-8 -*-
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')

import os
import json
import codecs
import threading
from Autodesk.Revit.DB import FilteredElementCollector, ImportInstance, ElementId, Transaction
from System.Windows.Forms import Form, ListView, View, Label, TextBox, Button, SaveFileDialog, OpenFileDialog, DialogResult, MessageBox, MessageBoxButtons, ListViewItem

# Access the Revit application and active document
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Variable to store the path to the JSON file
SETTINGS_FILE = None

def load_data_in_background(self):
    def load_data():
        data = get_dwg_links_with_layers(self.doc, self.view)
        self.Invoke(lambda: self.populate_list(data))  # Update UI from the main thread

    threading.Thread(target=load_data, daemon=True).start()

def get_dwg_links_with_layers(doc, view):
    """
    Optimized retrieval of DWG categories and their layers visible in the given Revit view.
    """
    try:
        dwg_data = {}
        dwg_links = list(FilteredElementCollector(doc, view.Id).OfClass(ImportInstance).ToElements())
        cached_categories = {dwg.Id: dwg.Category for dwg in dwg_links if dwg.Category}


        for dwg in dwg_links:
            category = dwg.Category
            if category and hasattr(category, 'Name'):
                category_name = category.Name
                if category.SubCategories:
                    subcategories = category.SubCategories
                    layers = [
                        (subcategory.Name, subcategory.get_Visible(view))
                        for subcategory in subcategories
                    ]
                    dwg_data[category_name] = {
                        "layers": sorted(layers, key=lambda x: x[0]),
                        "visibility": any(layer[1] for layer in layers)
                    }

        return dwg_data

    except Exception as e:
        MessageBox.Show("Error retrieving DWG links and layers: " + str(e), "Error", MessageBoxButtons.OK)
        return {}

class DWGCategoryAndLayerForm(Form):
    """
    Windows Form to display DWG categories and their layers with checkboxes, real-time search,
    and options to save and load visibility settings.
    """
    def __init__(self, dwg_data, doc, view):
        Form.__init__(self)

        self.Text = "DWG Categories and Layers with Save and Load"
        self.Width = 700
        self.Height = 750
        self.TopMost = True  # Ensures the window stays on top

        self.doc = doc
        self.view = view
        self.original_data = dwg_data  # Store original data for resetting after search
        self.filtered_data = dwg_data  # Start with all data visible

        self.init_ui()

    def init_ui(self):
        """
        Initialize UI components including real-time search functionality.
        """
        label = Label()
        label.Text = "DWG Categories and Layers:"
        label.Top = 10
        label.Left = 10
        self.Controls.Add(label)

        self.search_box = TextBox()
        self.search_box.Top = 40
        self.search_box.Left = 10
        self.search_box.Width = 600
        self.search_box.TextChanged += self.perform_search
        self.Controls.Add(self.search_box)

        self.listview = ListView()
        self.listview.Top = 80
        self.listview.Left = 10
        self.listview.Width = 650
        self.listview.Height = 500
        self.listview.View = View.Details
        self.listview.CheckBoxes = True
        self.listview.FullRowSelect = True
        self.listview.Columns.Add("Name", 400)
        self.listview.Columns.Add("Type", 150)
        self.listview.ItemChecked += self.on_item_checked

        self.populate_list(self.filtered_data)
        self.Controls.Add(self.listview)

        save_button = Button()
        save_button.Text = "Save Selections"
        save_button.Top = 600
        save_button.Left = 150
        save_button.Click += self.save_selections
        self.Controls.Add(save_button)

        load_button = Button()
        load_button.Text = "Load Selections"
        load_button.Top = 600
        load_button.Left = 350
        load_button.Click += self.load_selections
        self.Controls.Add(load_button)

        close_button = Button()
        close_button.Text = "Close"
        close_button.Top = 650
        close_button.Left = 300
        close_button.Click += self.close_form
        self.Controls.Add(close_button)

    def populate_list(self, data):
        """
        Populate the ListView with DWG categories and layers in batches for better performance.
        """
        self.listview.BeginUpdate()  # Suspend drawing to improve performance
        try:
            self.listview.Items.Clear()
            for category, info in sorted(data.items()):
                # Add DWG category
                category_item = ListViewItem(category)
                category_item.SubItems.Add("DWG")
                category_item.Checked = info["visibility"]
                self.listview.Items.Add(category_item)

                # Add layers
                for layer, is_visible in info["layers"]:
                    layer_item = ListViewItem("    " + layer)
                    layer_item.SubItems.Add("Layer")
                    layer_item.Checked = is_visible
                    self.listview.Items.Add(layer_item)
        finally:
            self.listview.EndUpdate()  # Resume drawing


    def perform_search(self, sender, event):
        search_term = self.search_box.Text.strip().lower()
        if not search_term:
            self.filtered_data = self.original_data
        else:
            self.filtered_data = {
                category: {
                    "layers": [
                        (layer, visible) for layer, visible in info["layers"] if search_term in layer.lower() or search_term in category.lower()
                    ],
                    "visibility": info["visibility"]
                }
                for category, info in self.original_data.items()
            }
            self.filtered_data = {k: v for k, v in self.filtered_data.items() if v["layers"]}
        self.populate_list(self.filtered_data)

    def save_selections(self, sender, event):
        """
        Save the current visibility selections from the ListView to a file, using the DWG's Category.Name as the key.
        """
        global SETTINGS_FILE

        # Update the internal state from the current ListView state
        for item in self.listview.Items:
            name = item.Text.strip()
            item_type = item.SubItems[1].Text
            is_visible = item.Checked

            if item_type == "DWG":
                if name in self.original_data:
                    self.original_data[name]["visibility"] = is_visible
            elif item_type == "Layer":
                for category, info in self.original_data.items():
                    for layer, visible in info["layers"]:
                        if layer == name:
                            layer_index = info["layers"].index((layer, visible))
                            info["layers"][layer_index] = (layer, is_visible)

        # Prompt the user to select a save location if no file is set
        if SETTINGS_FILE is None:
            save_dialog = SaveFileDialog()
            save_dialog.Title = "Save Layer Visibility Selections"
            save_dialog.Filter = "JSON Files (*.json)|*.json"
            save_dialog.DefaultExt = "json"
            save_dialog.AddExtension = True
            save_dialog.OverwritePrompt = True

            if save_dialog.ShowDialog() == DialogResult.OK:
                SETTINGS_FILE = save_dialog.FileName
            else:
                MessageBox.Show("Save canceled.", "Info", MessageBoxButtons.OK)
                return

        # Save selections using the Category.Name
        selections = {
            category: {
                "layers": {layer: visible for layer, visible in info["layers"]},
                "visibility": info["visibility"]
            }
            for category, info in self.original_data.items()
        }

        try:
            # Save the updated selections to the chosen file
            with codecs.open(SETTINGS_FILE, "w", "utf-8") as f:
                json.dump(selections, f, ensure_ascii=False, indent=4)

            MessageBox.Show("Selections saved successfully to " + SETTINGS_FILE, "Info", MessageBoxButtons.OK)

        except Exception as e:
            MessageBox.Show("Error saving selections: " + str(e), "Error", MessageBoxButtons.OK)

    def load_selections(self, sender, event):
        """
        Load visibility selections from a JSON file, apply them to the current view, and refresh the list while ensuring DWG names match.
        """
        load_dialog = OpenFileDialog()
        load_dialog.Title = "Load Layer Visibility Selections"
        load_dialog.Filter = "JSON Files (*.json)|*.json"
        load_dialog.DefaultExt = "json"

        if load_dialog.ShowDialog() == DialogResult.OK:
            file_path = load_dialog.FileName
        else:
            MessageBox.Show("Load canceled.", "Info", MessageBoxButtons.OK)
            return

        try:
            with codecs.open(file_path, "r", "utf-8") as f:
                saved_selections = json.load(f)

            # Get the DWG names present in the current view
            current_dwg_names = [
                dwg.Category.Name for dwg in FilteredElementCollector(self.doc, self.view.Id).OfClass(ImportInstance).ToElements()
                if dwg.Category
            ]

            # Check if all DWGs in the settings exist in the current view
            for category in saved_selections.keys():
                if category not in current_dwg_names:
                    MessageBox.Show("Wrong DWG: '{}' is not present in the current view.".format(category), "Error", MessageBoxButtons.OK)
                    return  # Exit without applying settings

            # Apply the loaded selections using the DWG Category.Name
            for category, info in saved_selections.items():
                self.toggle_dwg_visibility(category, info["visibility"])
                for layer, visible in info["layers"].items():
                    self.toggle_layer_visibility(layer, visible)

            # Update the internal state with loaded selections
            self.original_data = {
                category: {
                    "layers": sorted([(layer, visible) for layer, visible in info["layers"].items()], key=lambda x: x[0]),
                    "visibility": info["visibility"]
                }
                for category, info in sorted(saved_selections.items())
            }
            self.filtered_data = self.original_data

            # Refresh the list
            self.populate_list(self.filtered_data)

            MessageBox.Show("Selections loaded and applied successfully from " + file_path, "Info", MessageBoxButtons.OK)

        except Exception as e:
            MessageBox.Show("Error loading selections: " + str(e), "Error", MessageBoxButtons.OK)



    def on_item_checked(self, sender, event):
        item = event.Item
        name = item.Text.strip()
        item_type = item.SubItems[1].Text
        is_visible = item.Checked

        if item_type == "DWG":
            self.toggle_dwg_visibility(name, is_visible)
        elif item_type == "Layer":
            self.toggle_layer_visibility(name, is_visible)

    def toggle_dwg_visibility(self, dwg_name, is_visible):
        try:
            view_template_id = self.view.ViewTemplateId
            if view_template_id == ElementId.InvalidElementId:
                raise ValueError("No View Template is applied to the active view.")
            template = self.doc.GetElement(view_template_id)
            with Transaction(self.doc, "Toggle DWG Visibility") as t:
                t.Start()
                graphics_styles = FilteredElementCollector(self.doc).OfClass(ImportInstance).ToElements()
                for gs in graphics_styles:
                    if gs.Category and gs.Category.Name == dwg_name:
                        template.SetCategoryHidden(gs.Category.Id, not is_visible)
                t.Commit()
        except Exception as e:
            MessageBox.Show("Error toggling DWG visibility: " + str(e), "Error", MessageBoxButtons.OK)

    def toggle_layer_visibility(self, layer_name, is_visible):
        try:
            view_template_id = self.view.ViewTemplateId
            if view_template_id == ElementId.InvalidElementId:
                raise ValueError("No View Template is applied to the active view.")
            template = self.doc.GetElement(view_template_id)
            with Transaction(self.doc, "Toggle Layer Visibility") as t:
                t.Start()
                graphics_styles = FilteredElementCollector(self.doc).OfClass(ImportInstance).ToElements()
                for gs in graphics_styles:
                    if gs.Category and gs.Category.SubCategories:
                        for subcategory in gs.Category.SubCategories:
                            if subcategory.Name == layer_name:
                                template.SetCategoryHidden(subcategory.Id, not is_visible)
                t.Commit()
        except Exception as e:
            MessageBox.Show("Error toggling layer visibility: " + str(e), "Error", MessageBoxButtons.OK)

    def close_form(self, sender, event):
        self.Close()

def main():
    try:
        view = doc.ActiveView
        if view is None:
            raise ValueError("No active view. Please open a valid view and try again.")

        dwg_data = get_dwg_links_with_layers(doc, view)
        form = DWGCategoryAndLayerForm(dwg_data, doc, view)
        form.ShowDialog()

    except Exception as e:
        MessageBox.Show("An error occurred: " + str(e), "Error", MessageBoxButtons.OK)

main()
