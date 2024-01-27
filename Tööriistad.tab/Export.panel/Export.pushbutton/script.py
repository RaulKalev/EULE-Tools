# -*- coding: utf-8 -*-
import clr
import os
import csv
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from Autodesk.Revit.DB import FilteredElementCollector, ViewSchedule, SectionType
from System.Drawing import Font, FontStyle
from System.Windows.Forms import (Form, CheckedListBox, CheckBox, Label, PictureBox, PictureBoxSizeMode, ControlStyles,
                                  DialogResult, AnchorStyles, MessageBox, MessageBoxButtons, Panel, Cursors, TextBox,BorderStyle, HorizontalAlignment,
                                  MessageBoxIcon, FolderBrowserDialog, FormBorderStyle, Control, MouseButtons,FormWindowState)
from System.Drawing import Color, Size, Point, Bitmap, ContentAlignment

def get_image_path(image_filename):
    image_folder = "Icons"
    current_script_dir = os.path.dirname(__file__)
    relative_image_path = os.path.join(current_script_dir, image_folder, image_filename)
    return relative_image_path
try:
    test_image_path = get_image_path('Close.png')  # Replace with your image file name
    test_bitmap = Bitmap(test_image_path)
except Exception as e:
    print("Error loading image:", e)


class InteractivePictureBox:
    def __init__(self, pictureBox, normalImage, hoverImage, clickImage):
        self.pictureBox = pictureBox
        # Load images from file paths
        self.normalImage = Bitmap(get_image_path(normalImage))
        self.hoverImage = Bitmap(get_image_path(hoverImage))
        self.clickImage = Bitmap(get_image_path(clickImage))

        # Set initial image
        self.pictureBox.Image = self.normalImage
        self.pictureBox.SizeMode = PictureBoxSizeMode.StretchImage

        # Subscribe to mouse events
        self.pictureBox.MouseEnter += self.on_mouse_enter
        self.pictureBox.MouseLeave += self.on_mouse_leave
        self.pictureBox.MouseDown += self.on_mouse_down
        self.pictureBox.MouseUp += self.on_mouse_up

    def on_mouse_enter(self, sender, e):
        self.pictureBox.Image = self.hoverImage

    def on_mouse_leave(self, sender, e):
        self.pictureBox.Image = self.normalImage

    def on_mouse_down(self, sender, e):
        self.pictureBox.Image = self.clickImage

    def on_mouse_up(self, sender, e):
        if self.pictureBox.ClientRectangle.Contains(self.pictureBox.PointToClient(Control.MousePosition)):
            self.pictureBox.Image = self.hoverImage
        else:
            self.pictureBox.Image = self.normalImage

    def update_image(self, newImage):
        if self.pictureBox.Image is not None:
            self.pictureBox.Image.Dispose()  # Dispose the old image
        self.pictureBox.Image = newImage

    def Dispose(self):
        # Dispose all Bitmap resources
        if self.normalImage is not None:
            self.normalImage.Dispose()
        if self.hoverImage is not None:
            self.hoverImage.Dispose()
        if self.clickImage is not None:
            self.clickImage.Dispose()

        if self.pictureBox.Image is not None:
            self.pictureBox.Image.Dispose()

class BorderlessTextBox(TextBox):
    def __init__(self):
        self.SetStyle(ControlStyles.UserPaint, True)
        self.SetStyle(ControlStyles.AllPaintingInWmPaint, True)
        self.SetStyle(ControlStyles.OptimizedDoubleBuffer, True)

    def OnPaint(self, e):
        # Call the base class method
        TextBox.OnPaint(self, e)

        # Draw a black border (comment this out for no border)
        e.Graphics.DrawRectangle(Pens.Black, 0, 0, self.Width - 1, self.Height - 1)

class ScheduleListForm(Form):
    def __init__(self, schedules):
        self.schedules = schedules
        # Store schedules for filtering
        self.allSchedules = schedules
        # Constructor for the form
        appName = "Schedule Exporter"
        windowWidth = 500 #285before
        windowHeight = 400

        # Set the title and size of the window
        self.FormBorderStyle = FormBorderStyle.None
        self.Text = appName
        self.Size = Size(windowWidth, windowHeight)
        color3 = Color.FromArgb(49, 49, 49)
        color2 = Color.FromArgb(31, 31, 31)
        color1 = Color.FromArgb(24, 24, 24)
        colorText = Color.FromArgb(240,240,240)
        panelSize = 30
        self.minWidth = 500
        self.minHeight = 400
        self.BackColor = color1
        self.ForeColor = colorText        

        # Handling form resizing
        self.isResizing = False
        self.resizeHandleSize = 10  # Size of the area from the corner for resize handle
        self.Resize += self.on_form_resize


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
        self.titleBar.BackColor = color3

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
        # Create a PictureBox for the close button
        self.titleClose = PictureBox()
        self.titleClose.Location = Point(windowWidth-25, 5)
        self.titleClose.Size = Size(20, 20)
        self.titleClose.Image = Bitmap(get_image_path('Close.png'))
        self.titleClose.SizeMode = PictureBoxSizeMode.StretchImage
        self.titleClose.BackColor = color3
        self.titleClose.Anchor = AnchorStyles.Top | AnchorStyles.Right
        self.titleClose.Click += self.close_button_clicked
        self.titleCloseInteractive = InteractivePictureBox(
            self.titleClose, 'Close.png', 'CloseHover.png', 'CloseClick.png')

        #minimize
        self.titleMinimize = PictureBox()
        self.titleMinimize.Location = Point(windowWidth-45,5)
        self.titleMinimize.Size = Size(20,20)
        self.titleMinimize.Image = Bitmap(get_image_path('Minimize.png')) 
        self.titleMinimize.SizeMode = PictureBoxSizeMode.StretchImage
        self.titleMinimize.BackColor = color3
        self.titleMinimize.Anchor = AnchorStyles.Top | AnchorStyles.Right
        self.titleMinimizeInteractive = InteractivePictureBox(
            self.titleMinimize, 'Minimize.png', 'MinimizeHover.png', 'MinimizeClick.png')
        self.titleMinimize.Click += self.minimize_button_clicked
       
        # Create a search box
        self.searchBox = TextBox()
        self.searchBox.Text = "Search"
        self.searchBox.Font = Font(self.searchBox.Font.FontFamily, 9, FontStyle.Regular)  # Adjust the font size as needed
        self.searchBox.TextAlign = HorizontalAlignment.Right
        self.searchBox.Location = Point(windowWidth-250, panelSize + 15)
        self.searchBox.Size = Size(200, 20)
        self.searchBox.BorderStyle = BorderStyle.None  # Remove border
        self.searchBox.BackColor = color3  # Match background color
        self.searchBox.ForeColor = colorText
        self.searchBox.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        self.searchBox.TextChanged += self.on_search_changed  # Event handler for text change
        self.searchBox.Enter += self.search_box_enter
        self.searchBox.Leave += self.search_box_leave

        self.searchPanel = Panel()
        self.searchPanel.Location = Point(windowWidth-260, panelSize + 12)
        self.searchPanel.Size = Size(220, 22)  # Adjust size to accommodate padding
        self.searchPanel.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        self.searchPanel.BackColor = self.searchBox.BackColor  # Match the background color
        self.searchPanel.Controls.Add(self.searchBox)

        # Panel to contain the CheckedListBox
        self.panel = Panel()
        self.panel.Location = Point(10, panelSize + 50)
        self.panel.Size = Size(windowWidth-20, self.ClientSize.Height - (panelSize + 60))
        self.panel.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        self.panel.BackColor = color3  # Match the form's background color

        # Create a checklist box to list schedules
        self.checklist = CheckedListBox()
        self.checklist.Location = Point(-2, -2)  # Slightly offset inside the panel
        self.checklist.Size = Size(windowWidth-16, self.panel.Height + 18)  # Slightly larger to hide borders
        self.checklist.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        self.checklist.CheckOnClick = True
        self.checklist.BackColor = color3
        self.checklist.ForeColor = colorText
        for schedule in schedules:
            self.checklist.Items.Add(schedule.Name)  # Add schedule names to the checklist
        
        # Create a checkbox for selecting all schedules
        self.selectAll = CheckBox()
        self.selectAll.Text = "Select All"
        self.selectAll.Location = Point(10, panelSize+13)
        self.selectAll.Anchor = AnchorStyles.Top | AnchorStyles.Left
        self.selectAll.CheckedChanged += self.on_select_all_changed  # Event handler for checkbox change
        
        # Create an export button
        self.exportButton = PictureBox()
        self.exportButton.Location = Point(windowWidth-35, (panelSize+10))
        self.exportButton.Size = Size(30,30)
        self.exportButton.Image = export_image
        self.exportButton.SizeMode = PictureBoxSizeMode.StretchImage
        self.exportButton.BackColor = color1
        self.exportButton.Anchor = AnchorStyles.Top | AnchorStyles.Right
        self.exportButtonInteractive = InteractivePictureBox(
        self.exportButton, 'Export.png', 'ExportHover.png', 'ExportClick.png')
        self.exportButton.Click += self.on_export_clicked  # Event handler for button click
        
        self.titleBar.Controls.Add(self.titleLabel)
        self.Controls.Add(self.titleClose)
        self.Controls.Add(self.titleMinimize)
        self.Controls.Add(self.titleBar)
        self.Controls.Add(self.searchBox)
        self.Controls.Add(self.searchPanel)
        self.panel.Controls.Add(self.checklist)
        self.Controls.Add(self.panel)
        self.Controls.Add(self.selectAll)  # Add the checkbox to the form
        self.Controls.Add(self.exportButton)  # Add the button to the form
        self.Shown += self.on_form_shown

    def on_form_resize(self, sender, e):
        # Redraw the form and its contents
        self.Invalidate()

    def search_box_enter(self, sender, e):
        if self.searchBox.Text == "Search":
            self.searchBox.Text = ""

    def search_box_leave(self, sender, e):
        if self.searchBox.Text.strip() == "":
            self.searchBox.Text = "Search"

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

    def on_form_shown(self, sender, e):
        # Set focus to the checklist or another control when the form is first shown
        self.checklist.Focus()

    def on_resize_mouse_up(self, sender, e):
        self.isResizing = False

    def close_button_clicked(self, sender, e):
        self.Close()

    def minimize_button_clicked(self, sender, e):
        self.WindowState = FormWindowState.Minimized

    def on_search_changed(self, sender, e):
        searchText = self.searchBox.Text.lower()
        if searchText == "search":  # Ignore if the text is the default "Search"
            return

        self.checklist.Items.Clear()
        for schedule in self.allSchedules:
            if searchText in schedule.Name.lower():
                self.checklist.Items.Add(schedule.Name)
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
    try:
        form = ScheduleListForm(schedules)
        form.ShowDialog()
    except Exception as e:
        print("Error occurred:", e)

# Entry point for the script
if __name__ == "__main__":
    main()
