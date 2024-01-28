# -*- coding: utf-8 -*-
import clr
import os
import csv
import sys
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from Autodesk.Revit.DB import FilteredElementCollector, ViewSchedule, SectionType
from System.Windows.Forms import (Form, CheckBox, PictureBox, PictureBoxSizeMode, ListView, View,
                                  DialogResult, AnchorStyles, MessageBox, MessageBoxButtons, Panel, HorizontalAlignment,
                                  MessageBoxIcon, FolderBrowserDialog, FormBorderStyle)
from System.Drawing import Color, Size, Point, Bitmap, Font, FontStyle
from Snippets._imagePath import ImagePathHelper
from Snippets._interactivePictureBox import InteractivePictureBox
from Snippets._titlebar import TitleBar
from Snippets._searchBox import SearchBox
from Snippets._windowResize import WindowResizer

# Instantiate ImagePathHelper
image_helper = ImagePathHelper()
windowWidth = 500 
windowHeight = 390

# Load images using the get_image_path function
export_image_path = image_helper.get_image_path('Export.png')
minimize_image_path = image_helper.get_image_path('Minimize.png')
close_image_path = image_helper.get_image_path('Close.png')
search_image_path = image_helper.get_image_path('Search.png')
clear_image_path = image_helper.get_image_path('Clear.png')
logo_image_path = image_helper.get_image_path('Logo.png')

# Create Bitmap objects from the paths
export_image = Bitmap(export_image_path)
minimize_image = Bitmap(minimize_image_path)
close_image = Bitmap(close_image_path)
search_image = Bitmap(search_image_path)
clear_image = Bitmap(clear_image_path)
logo_image = Bitmap(logo_image_path)

class ScheduleListForm(Form):
    def __init__(self, schedules):
        self.schedules = schedules
        # Store schedules for filtering
        self.allSchedules = schedules
        # Constructor for the form
        appName = "Schedule Exporter"

        self.titleBar = TitleBar(self, appName, logo_image, minimize_image, close_image)
        self.searchBox = SearchBox(self, search_image, clear_image)
        self.resizer = WindowResizer(self)
        
        # Set the title and size of the window
        self.FormBorderStyle = FormBorderStyle.None
        self.Text = appName
        self.Size = Size(windowWidth, windowHeight)
        color3 = Color.FromArgb(49, 49, 49)
        color1 = Color.FromArgb(24, 24, 24)
        colorText = Color.FromArgb(240,240,240)
        panelSize = 30
        self.minWidth = 500
        self.minHeight = 400
        self.BackColor = color1
        self.ForeColor = colorText        

        # Handling form movement
        self.dragging = False
        self.offset = None

        # Panel to contain the CheckedListBox
        self.panel = Panel()
        self.panel.Location = Point(11, panelSize + 41)
        self.panel.Size = Size(windowWidth-22, self.ClientSize.Height - (panelSize + 52))
        self.panel.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        self.panel.BackColor = color3  # Match the form's background color
        self.panelBorder = Panel()
        self.panelBorder.Location = Point(15, panelSize + 45)
        self.panelBorder.Size = Size(windowWidth-35, self.ClientSize.Height - (panelSize + 63))
        self.panelBorder.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        self.panelBorder.BackColor = colorText
        self.panelBorderLine = Panel()
        self.panelBorderLine.Location = Point(10, panelSize + 40)
        self.panelBorderLine.Size = Size(windowWidth-20, self.ClientSize.Height - (panelSize + 50))
        self.panelBorderLine.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        self.panelBorderLine.BackColor = Color.FromArgb(69,69,69)

        # Create a checklist box to list schedules
        self.checklist = ListView()
        self.checklist.View = View.Details
        self.checklist.CheckBoxes = True
        self.checklist.FullRowSelect = True
        self.checklist.GridLines = False
        self.checklist.MultiSelect = True
        self.checklist.Location = Point(-2,-26)  # Slightly offset inside the panel
        self.checklist.Size = Size(windowWidth, self.panel.Height+34)  # Slightly larger to hide borders
        self.checklist.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        self.checklist.BackColor = color3
        self.checklist.ForeColor = colorText
        self.checklist.Font = Font("Helvetica", 9, FontStyle.Regular)
        self.checklist.Columns.Add("Schedules", -2, HorizontalAlignment.Left)
        for schedule in schedules:
            self.checklist.Items.Add(schedule.Name)  # Add schedule names to the checklist
        self.checklist.ItemCheck += self.on_listview_item_check
    
        # Create a checkbox for selecting all schedules
        self.selectAll = CheckBox()
        self.selectAll.Text = "Select All"
        self.selectAll.Location = Point(19, panelSize+13)
        self.selectAll.Anchor = AnchorStyles.Top | AnchorStyles.Left
        self.selectAll.Font = Font("Helvetica", 8, FontStyle.Regular)
        self.selectAll.CheckedChanged += self.on_select_all_changed  # Event handler for checkbox change
        
        # Create an export button
        self.exportButton = PictureBox()
        self.exportButton.Location = Point(windowWidth-39, (panelSize+7))
        self.exportButton.Size = Size(30,30)
        self.exportButton.Image = export_image
        self.exportButton.SizeMode = PictureBoxSizeMode.StretchImage
        self.exportButton.BackColor = color1
        self.exportButton.Anchor = AnchorStyles.Top | AnchorStyles.Right
        self.exportButtonInteractive = InteractivePictureBox(
        self.exportButton, 'Export.png', 'ExportHover.png', 'ExportClick.png')
        self.exportButton.Click += self.on_export_clicked  # Event handler for button click
        
        self.Controls.Add(self.searchBox)
        self.Controls.Add(self.titleBar)
        
        self.Controls.Add(self.panelBorder)
        self.Controls.Add(self.panel)
        self.Controls.Add(self.panelBorderLine)
        
        self.panelBorder.Controls.Add(self.checklist)
        
        self.Controls.Add(self.selectAll)
        self.Controls.Add(self.exportButton)

    def on_listview_item_check(self, sender, item_check_event_args):
        # Logic to handle item check events
        item = self.checklist.Items[item_check_event_args.Index]
        # If you need to do something when an item is checked/unchecked, you can add it here

    def on_select_all_changed(self, sender, e):
        for item in self.checklist.Items:
            item.Checked = self.selectAll.Checked

    # Event handler for the export button
    def on_export_clicked(self, sender, e):
        selectedSchedules = [item.Text for item in self.checklist.Items if item.Checked]
        if not selectedSchedules:
            MessageBox.Show("No schedules selected for export.", "Export", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return

        folderBrowserDialog = FolderBrowserDialog()

        if folderBrowserDialog.ShowDialog() == DialogResult.OK:
            folderPath = folderBrowserDialog.SelectedPath
            existing_files = []

            # Check for existing files
            for scheduleName in selectedSchedules:
                filePath = os.path.join(folderPath, "{}.csv".format(scheduleName))
                if os.path.exists(filePath):
                    existing_files.append(os.path.basename(filePath))

            if existing_files:
                if not self.confirm_overwrite(existing_files):
                    return  # User chose not to overwrite

            # Export schedules
            for scheduleName in selectedSchedules:
                scheduleElement = next((s for s in self.schedules if s.Name == scheduleName), None)
                if scheduleElement:
                    self.export_schedule(scheduleElement, folderPath)
            MessageBox.Show("Schedules exported successfully to {}".format(folderPath), "Export Successful", MessageBoxButtons.OK, MessageBoxIcon.Information)

    def export_schedule(self, schedule, folderPath):
        filePath = os.path.join(folderPath, "{}.csv".format(schedule.Name))
        with open(filePath, mode='w') as file:
            writer = csv.writer(file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)

            # Write the column headers
            scheduleDefinition = schedule.Definition
            fieldIds = scheduleDefinition.GetFieldOrder()
            header = [scheduleDefinition.GetField(fieldId).GetName() for fieldId in fieldIds]
            writer.writerow(header)  # This should write the header row once

            # Fetching schedule data
            tableData = schedule.GetTableData()
            sectionData = tableData.GetSectionData(SectionType.Body)

            # Iterate through all rows and columns
            for rowIndex in range(sectionData.NumberOfRows):
                if rowIndex == 0:  # Skip the first row if it's a duplicate of the header
                    continue
                row = []
                for columnIndex in range(sectionData.NumberOfColumns):
                    cell = schedule.GetCellText(SectionType.Body, rowIndex, columnIndex)
                    row.append(cell)
                # Check if the row contains any non-whitespace data
                if not all(cell is None or cell.strip() == '' for cell in row):
                    writer.writerow(row)  # Write only rows with data

    def confirm_overwrite(self, existing_files):
        message = "The following files already exist in the selected folder:\n\n"
        message += "\n".join(existing_files)
        message += "\n\nDo you want to overwrite these files?"
        result = MessageBox.Show(message, "Confirm Overwrite", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        return result == DialogResult.Yes
    
# Function to retrieve all ViewSchedules from the Revit document
def get_schedules(doc):
    return FilteredElementCollector(doc).OfClass(ViewSchedule)

# Main function to execute the script
def main():
    doc = __revit__.ActiveUIDocument.Document
    schedules = get_schedules(doc)  # Retrieve schedules
    try:
        form = ScheduleListForm(schedules)
        form.ShowDialog()
    except Exception as e:
        print("Error occurred:", e)

# Entry point for the script
if __name__ == "__main__":
    main()
