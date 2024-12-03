# -*- coding: utf-8 -*-
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

import os
import json
import codecs
import threading
from Autodesk.Revit.DB import FilteredElementCollector, ImportInstance, ElementId, Transaction
from System.Windows.Forms import Form, ListView, View, Label, TextBox, Button, SaveFileDialog, OpenFileDialog, DialogResult, MessageBox, MessageBoxButtons, ListViewItem, FormBorderStyle, Panel
from System.Drawing import Color, Size, Point, Bitmap, Font, FontStyle, ContentAlignment, StringFormat, Graphics, Rectangle, SolidBrush, Brushes, Pen, Drawing2D

from Snippets._titlebar import TitleBar
from Snippets._imagePath import ImagePathHelper
from Snippets._windowResize import WindowResizer

# Access the Revit application and active document
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

#Setting up colors
backgroundColor = Color.FromArgb(31,31,31)
textColor = Color.FromArgb(240,240,240)
lineColor = Color.FromArgb(69,69,69)

# Instantiate ImagePathHelper
image_helper = ImagePathHelper()
windowWidth = 500
windowHeight = 750
titleBar = 40

# Load images using the get_image_path function
minimize_image_path = image_helper.get_image_path('Minimize.png')
close_image_path = image_helper.get_image_path('Close.png')
clear_image_path = image_helper.get_image_path('Clear.png')
logo_image_path = image_helper.get_image_path('Logo.png')

# Create Bitmap objects from the paths
minimize_image = Bitmap(minimize_image_path)
close_image = Bitmap(close_image_path)
clear_image = Bitmap(clear_image_path)
logo_image = Bitmap(logo_image_path)

# Variable to store the path to the JSON file
SETTINGS_FILE = None

def get_project_save_folder():
    """
    Get or create the LayerToggles folder in the location of the opened Revit model.
    """
    try:
        # Get the directory of the currently opened Revit project
        project_path = doc.PathName  # Full path to the Revit project
        if not project_path:
            raise ValueError("The project must be saved before using the save/load feature.")

        project_dir = os.path.dirname(project_path)
        save_folder = os.path.join(project_dir, "LayerToggles")

        # Create the LayerToggles folder if it doesn't exist
        if not os.path.exists(save_folder):
            os.makedirs(save_folder)

        return save_folder

    except Exception as e:
        MessageBox.Show("Error creating project-specific folder: " + str(e), "Error", MessageBoxButtons.OK)
        return None

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
        appName = "DWG Layer manipulation"
        self.Width = windowWidth
        self.Height = windowHeight
        self.TopMost = True  # Ensures the window stays on top
        self.titleBar = TitleBar(self, appName, logo_image, minimize_image, close_image)
        self.FormBorderStyle = FormBorderStyle.None
        self.Text = appName
        self.Size = Size(windowWidth, windowHeight)
        self.BackColor = backgroundColor
        self.ForeColor = textColor
        self.is_expanding = False

        self.doc = doc
        self.view = view
        self.original_data = dwg_data  # Store original data for resetting after search
        self.filtered_data = dwg_data  # Start with all data visible

        self.Controls.Add(self.titleBar)

        self.dragging = False
        self.offset = None
        
        self.init_ui()

    def init_ui(self):
        """
        Initialize UI components including real-time search functionality.
        """
        listStartHeight = titleBar+5
        listHeight = 600
        buttonRow = 700
        searchTop = titleBar+5
        searchLeft = 275

        testColor = Color.FromArgb(100,240,100)

        searchBox_CoverTop = Panel()
        searchBox_CoverTop.Size = Size(200,2)
        searchBox_CoverTop.Location = Point(searchLeft,searchTop)
        searchBox_CoverTop.BackColor = backgroundColor

        searchBox_BorderTop = Panel()
        searchBox_BorderTop.Size = Size(199,1)
        searchBox_BorderTop.Location = Point(searchLeft,searchTop)
        searchBox_BorderTop.BackColor = lineColor

        searchBox_CoverBottom = Panel()
        searchBox_CoverBottom.Size = Size(200,2)
        searchBox_CoverBottom.Location = Point(searchLeft,searchTop+18)
        searchBox_CoverBottom.BackColor = backgroundColor

        searchBox_BorderBottom = Panel()
        searchBox_BorderBottom.Size = Size(199,1)
        searchBox_BorderBottom.Location = Point(searchLeft,searchTop+18)
        searchBox_BorderBottom.BackColor = lineColor

        searchBox_CoverLeft = Panel()
        searchBox_CoverLeft.Size = Size(2,20)
        searchBox_CoverLeft.Location = Point(searchLeft,searchTop)
        searchBox_CoverLeft.BackColor = backgroundColor

        searchBox_BorderLeft = Panel()
        searchBox_BorderLeft.Size = Size(1,19)
        searchBox_BorderLeft.Location = Point(searchLeft,searchTop)
        searchBox_BorderLeft.BackColor = lineColor

        searchBox_CoverRight = Panel()
        searchBox_CoverRight.Size = Size(2,20)
        searchBox_CoverRight.Location = Point(searchLeft+198,searchTop)
        searchBox_CoverRight.BackColor = backgroundColor

        searchBox_BorderRight = Panel()
        searchBox_BorderRight.Size = Size(1,19)
        searchBox_BorderRight.Location = Point(searchLeft+198,searchTop)
        searchBox_BorderRight.BackColor = lineColor

        self.search_box = TextBox()
        self.search_box.Top = searchTop
        self.search_box.Left = searchLeft
        self.search_box.Width = 200
        self.search_box.TextChanged += self.perform_search
        self.search_box.BackColor = backgroundColor
        self.search_box.ForeColor = textColor

        self.Controls.Add(searchBox_BorderRight)
        self.Controls.Add(searchBox_BorderLeft)
        self.Controls.Add(searchBox_BorderTop)
        self.Controls.Add(searchBox_BorderBottom)
        self.Controls.Add(searchBox_CoverRight)
        self.Controls.Add(searchBox_CoverLeft)
        self.Controls.Add(searchBox_CoverTop)
        self.Controls.Add(searchBox_CoverBottom)
        self.Controls.Add(self.search_box)

        listLabel = Label()
        listLabel.AutoSize = False
        listLabel.TextAlign = ContentAlignment.MiddleLeft
        listLabel.Text = "CAD kihid:"
        listLabel.Font = Font("Helvetica", 11, FontStyle.Regular)
        listLabel.Location = Point(5, listStartHeight+3)
        listLabel.Size = Size(100,15)
        listLabel.BackColor = backgroundColor
        listLabel.ForeColor = textColor

        listBorderTop = Panel()
        listBorderTop.Location = Point(0, listStartHeight)
        listBorderTop.Size = Size(520, 26)
        listBorderTop.BackColor = backgroundColor

        listBorderTop_Border = Panel()
        listBorderTop_Border.Location = Point(0, listStartHeight+23)
        listBorderTop_Border.Size = Size(520, 1)
        listBorderTop_Border.BackColor = lineColor

        listBorderBottom = Panel()
        listBorderBottom.Location = Point(0, listStartHeight+listHeight-20)
        listBorderBottom.Size = Size(520, 20)
        listBorderBottom.BackColor = backgroundColor

        self.listview = ListView()
        self.listview.Top = listStartHeight
        self.listview.Left = -1
        self.listview.Width = windowWidth+20
        self.listview.Height = listHeight
        self.listview.View = View.Details
        self.listview.CheckBoxes = True
        self.listview.FullRowSelect = True
        self.listview.Columns.Add("Name", 550)
        #self.listview.Columns.Add("Type", 150)
        self.listview.Font = Font("Helvetica", 10, FontStyle.Regular)
        self.listview.BackColor = backgroundColor
        self.listview.ForeColor = textColor
        self.listview.ItemChecked += self.on_item_checked

        self.listview.ColumnWidthChanging += self.prevent_column_resize

        self.Controls.Add(listLabel)
        self.Controls.Add(listBorderTop_Border)
        self.Controls.Add(listBorderTop)
        self.Controls.Add(listBorderBottom)
        self.populate_list(self.filtered_data)
        self.Controls.Add(self.listview)

        save_button = Button()
        save_button.Text = "Save Selections"
        save_button.Top = buttonRow
        save_button.Left = 150
        save_button.Click += self.save_selections
        self.Controls.Add(save_button)

        load_button = Button()
        load_button.Text = "Load Selections"
        load_button.Top = buttonRow
        load_button.Left = 350
        load_button.Click += self.load_selections
        self.Controls.Add(load_button)

        #close_button = Button()
        #close_button.Text = "Close"
        #close_button.Top = buttonRow
        #close_button.Left = 300
        #close_button.Click += self.close_form
        #self.Controls.Add(close_button)

    def load_data_in_background(self):
        def load_data():
            data = get_dwg_links_with_layers(self.doc, self.view)
            self.Invoke(lambda: self.populate_list(data))  # Update UI from the main thread

        threading.Thread(target=load_data, daemon=True).start()

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

    def prevent_column_resize(self, sender, event):
        """
        Prevent columns from being resized.
        """
        event.Cancel = True
        event.NewWidth = sender.Columns[event.ColumnIndex].Width  # Keep the current width


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
        Save the current visibility selections from the ListView to a file,
        using the DWG's Category.Name as the key and the current view's name as the filename.
        """
        try:
            save_folder = get_project_save_folder()
            if not save_folder:
                return

            view_name = self.view.Name.replace(" ", "_")  # Replace spaces to make it filename-friendly
            save_path = os.path.join(save_folder, view_name + "_LayerToggles.json")

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

            # Prepare selections to save
            selections = {
                category: {
                    "layers": {layer: visible for layer, visible in info["layers"]},
                    "visibility": info["visibility"]
                }
                for category, info in self.original_data.items()
            }

            # Save to the JSON file
            with codecs.open(save_path, "w", "utf-8") as f:
                json.dump(selections, f, ensure_ascii=False, indent=4)

            MessageBox.Show("Selections saved successfully to " + save_path, "Info", MessageBoxButtons.OK)

        except Exception as e:
            MessageBox.Show("Error saving selections: " + str(e), "Error", MessageBoxButtons.OK)

    def load_selections(self, sender, event):
        """
        Load visibility selections from a JSON file by matching DWG names,
        apply them, and refresh the list.
        """
        try:
            save_folder = get_project_save_folder()
            if not save_folder:
                return

            # List all JSON files in the save folder
            json_files = [f for f in os.listdir(save_folder) if f.endswith("_LayerToggles.json")]

            if not json_files:
                MessageBox.Show("No saved selections found for this project.", "Info", MessageBoxButtons.OK)
                return

            # Get the DWG names present in the current view
            current_dwg_names = [
                dwg.Category.Name for dwg in FilteredElementCollector(self.doc, self.view.Id).OfClass(ImportInstance).ToElements()
                if dwg.Category
            ]

            # Iterate through the JSON files and look for matches
            loaded = False
            for file_name in json_files:
                file_path = os.path.join(save_folder, file_name)

                with codecs.open(file_path, "r", "utf-8") as f:
                    saved_selections = json.load(f)

                # Check if any DWG name in the saved file matches the current view's DWG names
                matching_dwg_names = [name for name in saved_selections.keys() if name in current_dwg_names]

                if matching_dwg_names:
                    # Apply the loaded selections for matched DWG names
                    for category in matching_dwg_names:
                        info = saved_selections[category]
                        self.toggle_dwg_visibility(category, info["visibility"])
                        for layer, visible in info["layers"].items():
                            self.toggle_layer_visibility(layer, visible)

                    # Update the internal state with loaded selections
                    self.original_data = {
                        category: {
                            "layers": sorted([(layer, visible) for layer, visible in info["layers"].items()], key=lambda x: x[0]),
                            "visibility": info["visibility"]
                        }
                        for category, info in saved_selections.items() if category in matching_dwg_names
                    }
                    self.filtered_data = self.original_data

                    # Refresh the list
                    self.populate_list(self.filtered_data)

                    loaded = True
                    MessageBox.Show("Selections loaded and applied successfully from " + file_path, "Info", MessageBoxButtons.OK)
                    break

            if not loaded:
                MessageBox.Show("No matching DWG names found in saved files.", "Info", MessageBoxButtons.OK)

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
