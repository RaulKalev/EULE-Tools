# -*- coding: utf-8 -*-
# Import necessary CLR (Common Language Runtime) references to access .NET libraries
import clr
import os
import csv
clr.AddReference('RevitAPI')  # Access to Revit API
clr.AddReference('RevitAPIUI')  # Access to Revit UI API
clr.AddReference('System.Windows.Forms')  # Access to Windows Forms for UI creation
clr.AddReference('System.Drawing')  # Access to basic graphic functionality

# Import specific classes from the Revit API
from Autodesk.Revit.DB import FilteredElementCollector, ViewSchedule, SectionType

# Import specific classes from the .NET libraries for UI
from System.Windows.Forms import (Application, Form, CheckedListBox, Button, CheckBox, Label, PictureBox, PictureBoxSizeMode,
                                  DialogResult, AnchorStyles, MessageBox, MessageBoxButtons, Panel, Cursors,
                                  MessageBoxIcon, SaveFileDialog, FolderBrowserDialog, FlatStyle, FormBorderStyle, Control, MouseButtons,FormWindowState)
from System.Drawing import Color, Size, Point, Bitmap,ContentAlignment
# Define a class for the main form

def get_image_path(image_filename):
    # Define the path to your image folder here
    image_folder = "Icons"  # This goes up one level and then into the ImagesFolder
    current_script_dir = os.path.dirname(__file__)
    relative_image_path = os.path.join(current_script_dir, image_folder, image_filename)
    return relative_image_path

class ScheduleListForm(Form):
    def __init__(self, schedules):
        self.schedules = schedules
        # Constructor for the form

        appName = "Schedule Exporter"
        windowWidth = 285
        windowHeight = 400

        # Set the title and size of the window
        self.FormBorderStyle = FormBorderStyle.None
        self.Text = appName
        self.Size = Size(windowWidth, windowHeight)
        color2 = Color.FromArgb(31, 31, 31)
        color1 = Color.FromArgb(24, 24, 24)
        colorText = Color.FromArgb(240,240,240)
        panelSize = 30
        self.minWidth = 285
        self.minHeight = 400
        self.BackColor = color1
        self.ForeColor = colorText        

        # Handling form resizing
        self.isResizing = False
        self.resizeHandleSize = 10  # Size of the area from the corner for resize handle

        # Attach mouse events for resizing
        self.MouseDown += self.on_resize_mouse_down
        self.MouseMove += self.on_resize_mouse_move
        self.MouseUp += self.on_resize_mouse_up

        # Handling form movement
        self.dragging = False
        self.offset = None

        # Load images using the get_image_path function
        export_image_path = get_image_path('Export.png')
        minimize_image_path = get_image_path('Minimize.png')
        close_image_path = get_image_path('Close.png')

        # Create Bitmap objects from the paths
        export_image = Bitmap(export_image_path)
        minimize_image = Bitmap(minimize_image_path)
        close_image = Bitmap(close_image_path)

        #Title bar
        self.titleBar = Panel()
        self.titleBar.MouseDown += self.form_mouse_down
        self.titleBar.MouseMove += self.form_mouse_move
        self.titleBar.MouseUp += self.form_mouse_up
        self.titleBar.Location = Point(0,0)
        self.titleBar.Size = Size(self.ClientSize.Width, panelSize)
        self.titleBar.Anchor = AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Left
        self.titleBar.BackColor = color2

        # Title label
        self.titleLabel = Label()
        self.titleLabel.MouseDown += self.form_mouse_down
        self.titleLabel.MouseMove += self.form_mouse_move
        self.titleLabel.MouseUp += self.form_mouse_up
        self.titleLabel.Text = appName
        self.titleLabel.TextAlign = ContentAlignment.MiddleCenter
        self.titleLabel.Location = Point(0,0)
        self.titleLabel.Size = Size(self.titleBar.Width, self.titleBar.Height)
        self.titleLabel.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        self.titleLabel.BackColor = Color.Transparent
        self.titleLabel.ForeColor = colorText

        #Title bar buttons
        #close
        self.titleClose = PictureBox()
        self.titleClose.Location = Point(255,5)
        self.titleClose.Anchor = AnchorStyles.Top | AnchorStyles.Right
        self.titleClose.Size = Size(20,20)
        self.titleClose.Image = close_image
        self.titleClose.SizeMode = PictureBoxSizeMode.StretchImage
        self.titleClose.BackColor = color2
        self.titleClose.Click += self.close_button_clicked
        #minimize
        self.titleMinimize = Button()
        self.titleMinimize.Location = Point(225,5)
        self.titleMinimize.Text = "—"
        self.titleMinimize.Anchor = AnchorStyles.Top | AnchorStyles.Right
        self.titleMinimize.Size = Size(20,20)
        self.titleMinimize.FlatStyle = FlatStyle.Flat
        self.titleMinimize.FlatAppearance.BorderSize = 0
        self.titleMinimize.BackColor = color2
        self.titleMinimize.ForeColor = colorText
        self.titleMinimize.Click += self.minimize_button_clicked
        
        # Panel to contain the CheckedListBox
        self.panel = Panel()
        self.panel.Location = Point(10, panelSize + 50)
        self.panel.Size = Size(265, self.ClientSize.Height - (panelSize + 60))
        self.panel.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        self.panel.BackColor = color2  # Match the form's background color

        # Create a checklist box to list schedules
        self.checklist = CheckedListBox()
        self.checklist.Location = Point(-2, -2)  # Slightly offset inside the panel
        self.checklist.Size = Size(269, self.panel.Height + 18)  # Slightly larger to hide borders
        self.checklist.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        self.checklist.CheckOnClick = True
        self.checklist.BackColor = color2
        self.checklist.ForeColor = colorText
        for schedule in schedules:
            self.checklist.Items.Add(schedule.Name)  # Add schedule names to the checklist
        
        # Create a checkbox for selecting all schedules
        self.selectAll = CheckBox()
        self.selectAll.Text = "Select All"
        self.selectAll.Location = Point(10, panelSize+10)
        self.selectAll.Anchor = AnchorStyles.Top | AnchorStyles.Left
        self.selectAll.CheckedChanged += self.on_select_all_changed  # Event handler for checkbox change
        
        # Create an export button
        self.exportButton = PictureBox()
        self.exportButton.Location = Point(230, (panelSize+2.9))
        self.exportButton.Size = Size(45,45)
        self.exportButton.Image = export_image
        self.exportButton.SizeMode = PictureBoxSizeMode.StretchImage
        self.exportButton.BackColor = color1
        self.exportButton.Anchor = AnchorStyles.Top | AnchorStyles.Right
        self.exportButton.Click += self.on_export_clicked  # Event handler for button click
        
        self.titleBar.Controls.Add(self.titleLabel)
        self.Controls.Add(self.titleClose)
        self.Controls.Add(self.titleMinimize)
        self.Controls.Add(self.titleBar)
        self.panel.Controls.Add(self.checklist)
        self.Controls.Add(self.panel)
        self.Controls.Add(self.selectAll)  # Add the checkbox to the form
        self.Controls.Add(self.exportButton)  # Add the button to the form

    def on_button_focus(self, sender, e):
        sender.FlatAppearance.BorderSize = 0

    def on_button_lose_focus(self, sender, e):
        sender.FlatAppearance.BorderSize = 0

    # Moving the window
    def form_mouse_down(self, sender, e):
        if e.Button == MouseButtons.Left:
            self.dragging = True
            self.offset = Point(e.X - self.titleLabel.Left, e.Y - self.titleLabel.Top)

    def form_mouse_move(self, sender, e):
        if self.dragging:
            screenPosition = Point(e.X + self.Location.X, e.Y + self.Location.Y)
            self.Location = Point(screenPosition.X - self.offset.X, screenPosition.Y - self.offset.Y)

    def form_mouse_up(self, sender, e):
        self.dragging = False    

    # Event handlers for resizing
    def on_resize_mouse_down(self, sender, e):
        if e.X >= self.Width - self.resizeHandleSize and e.Y >= self.Height - self.resizeHandleSize:
            self.isResizing = True

    def on_resize_mouse_move(self, sender, e):
        if self.isResizing:
            newWidth = max(e.X, self.minWidth)
            newHeight = max(e.Y, self.minHeight)
            self.Width = newWidth
            self.Height = newHeight
        elif e.X >= self.Width - self.resizeHandleSize and e.Y >= self.Height - self.resizeHandleSize:
            self.Cursor = Cursors.SizeNWSE
        else:
            self.Cursor = Cursors.Default

    def on_resize_mouse_up(self, sender, e):
        self.isResizing = False

    def close_button_clicked(self, sender, e):
        self.Close()

    def minimize_button_clicked(self, sender, e):
        self.WindowState = FormWindowState.Minimized


    # Event handler for the 'Select All' checkbox
    def on_select_all_changed(self, sender, e):
        # Check or uncheck all items in the checklist based on the checkbox state
        for i in range(self.checklist.Items.Count):
            self.checklist.SetItemChecked(i, self.selectAll.Checked)

    # Event handler for the export button
    def on_export_clicked(self, sender, e):
        selectedSchedules = [self.checklist.Items[i] for i in range(self.checklist.Items.Count) if self.checklist.GetItemChecked(i)]
        if not selectedSchedules:
            MessageBox.Show("No schedules selected for export.", "Export", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return

        folderBrowserDialog = FolderBrowserDialog()

        if folderBrowserDialog.ShowDialog() == DialogResult.OK:
            folderPath = folderBrowserDialog.SelectedPath
            for schedule in selectedSchedules:
                scheduleElement = [s for s in self.schedules if s.Name == schedule][0]
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

                    
# Function to retrieve all ViewSchedules from the Revit document
def get_schedules(doc):
    return FilteredElementCollector(doc).OfClass(ViewSchedule)



# Main function to execute the script
def main():
    doc = __revit__.ActiveUIDocument.Document
    schedules = get_schedules(doc)  # Retrieve schedules
    form = ScheduleListForm(schedules)  # Pass the schedules to the form
    form.ShowDialog()

# Entry point for the script
if __name__ == "__main__":
    main()
