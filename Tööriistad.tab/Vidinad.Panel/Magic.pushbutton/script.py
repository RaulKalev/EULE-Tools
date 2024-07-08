# -*- coding: utf-8 -*-
import clr
import System
import math
import System.Threading.Thread
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
import os
import ConfigParser as configparser  # Use ConfigParser for IronPython compatibility
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import Color as RevitColor
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from System.Windows.Forms import Application, TrackBar, TickStyle, Form, Button, Label, ColorDialog, CheckBox, DialogResult, TextBox, Timer, Cursors, DrawMode, HorizontalAlignment, DialogResult, RichTextBox, FormWindowState, MessageBox, FormBorderStyle, BorderStyle, RadioButton, ListBox, GroupBox, ComboBox, ComboBoxStyle, AnchorStyles, Panel, PictureBoxSizeMode, PictureBox, ToolTip, FormStartPosition
from System.Drawing import Color, Size, Point, Bitmap, Font, FontStyle, ContentAlignment, StringFormat, Graphics, Rectangle, SolidBrush, Brushes
from System.Drawing.Drawing2D import SmoothingMode
from System import Array, Object
from math import radians, sin, cos
from decimal import Decimal
from Snippets._titlebar import TitleBar
from Snippets._imagePath import ImagePathHelper
from Snippets._windowResize import WindowResizer
from Snippets._interactivePictureBox import InteractivePictureBox
from Snippets._intersections import line_intersection,line_segment_intersection,find_closest_intersection,get_intersection_point
from Snippets._fovCalculations import rotate_vector,calculate_fov_endpoints
from Scripts._advancedCamera import calculator_1
from Snippets._revitUtilities import list_filled_region_type_names_and_ids,get_custom_detail_lines,draw_line,simulate_camera_fov,select_cameras

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
app   = __revit__.Application
rvt_year = int(app.VersionNumber)

# Instantiate ImagePathHelper
image_helper = ImagePathHelper()
windowWidth = 280
windowHeight = 570
titleBar = 40

# Load images using the get_image_path function
minimize_image_path = image_helper.get_image_path('Minimize.png')
close_image_path = image_helper.get_image_path('Close.png')
clear_image_path = image_helper.get_image_path('Clear.png')
logo_image_path = image_helper.get_image_path('Logo.png')
run_image_path = image_helper.get_image_path('Run.png')
select_image_path = image_helper.get_image_path('Select.png')
expand_image_path = image_helper.get_image_path('Expand.png')
contract_image_path = image_helper.get_image_path('Contract.png')
selectColor_image_path = image_helper.get_image_path('SelectColor.png')
createRegion_image_path = image_helper.get_image_path('CreateRegion.png')
# Create Bitmap objects from the paths
minimize_image = Bitmap(minimize_image_path)
close_image = Bitmap(close_image_path)
clear_image = Bitmap(clear_image_path)
logo_image = Bitmap(logo_image_path)
run_image = Bitmap(run_image_path)
select_image = Bitmap(select_image_path)
expand_image = Bitmap(expand_image_path)
contract_image = Bitmap(contract_image_path)
selectColor_image = Bitmap(selectColor_image_path)
createRegion_image = Bitmap(createRegion_image_path)
# Get Solid Pattern
all_patterns       = FilteredElementCollector(doc).OfClass(FillPatternElement).ToElements()
all_solid_patterns = [pat for pat in all_patterns if pat.GetFillPattern().IsSolidFill]
solid_pattern      = all_solid_patterns[0]

def get_filled_region(filled_region_name):
    """Function to get FireWall Types based on FamilyName"""
    # Create Filter
    pvp         = ParameterValueProvider(ElementId(BuiltInParameter.ALL_MODEL_TYPE_NAME))
    condition   = FilterStringEquals()
    fRule =     FilterStringRule(pvp, condition, filled_region_name, True) if rvt_year < 2022 \
             else FilterStringRule(pvp, condition, filled_region_name)
    my_filter   = ElementParameterFilter(fRule)

    # Get Types
    return FilteredElementCollector(doc).OfClass(FilledRegionType).WherePasses(my_filter).FirstElement()

def create_RegionType(name, color, masking = False, lineweight = 1):
    #type: (str, Color, bool, int) -> FilledRegionType
    """Create FilledRegionType with Solid Pattern."""
    random_filled_region = FilteredElementCollector(doc).OfClass(FilledRegionType).FirstElement()
    new_region           = random_filled_region.Duplicate(name)

    # Set Solid Pattern
    new_region.BackgroundPatternId = ElementId(-1)
    new_region.ForegroundPatternId = solid_pattern.Id

    # Set Colour
    new_region.BackgroundPatternColor = color
    new_region.ForegroundPatternColor = color

    # Masking
    new_region.IsMasking = masking

    # LineWeight
    new_region.LineWeight = lineweight

    return new_region

class DetailLineFilter(ISelectionFilter):
    def AllowElement(self, elem):
        return elem.Category.Id.IntegerValue == int(BuiltInCategory.OST_Lines)

def main_script(camera_info, fov_angle, max_distance_mm, detail_lines, filled_region_type_id):
    # Unpack camera position, from_linked_file flag, and rotation angle from camera_info
    camera_position, from_linked_file, rotation_angle = camera_info

    activeView = doc.ActiveView
    # Assuming camera_position needs to be an XYZ object
    if not from_linked_file:
        # If the camera is from the current project and camera_info[0] is a FamilyInstance
        camera_element = camera_info[0]
        # Extract the XYZ position from the FamilyInstance
        if hasattr(camera_element, "Location") and hasattr(camera_element.Location, "Point"):
            camera_position = camera_element.Location.Point
        else:
            MessageBox.Show("Camera position could not be determined.")
            return

    if not isinstance(activeView, ViewPlan):
        MessageBox.Show("The active view is not a plan view. Please switch to a plan view to draw detail lines.")
        return

    # Convert fov_angle and max_distance_mm to float if they are not already
    fov_angle = float(fov_angle)
    max_distance_mm = float(max_distance_mm)
    rotation_angle = float(rotation_angle)  # Ensure rotation_angle is a float

    # Call the simulate_camera_fov function using the unpacked and converted values
    simulate_camera_fov(doc, camera_position, fov_angle, max_distance_mm, detail_lines, activeView, rotation_angle, filled_region_type_id)

class RevitLinkSelectionFilter(ISelectionFilter):
    """Selection filter to allow only Revit link instances."""
    def AllowElement(self, elem):
        return isinstance(elem, RevitLinkInstance)
    
class LinkedFileCameraSelectionFilter(ISelectionFilter):
    def __init__(self, doc):
        self.doc = doc

    def AllowElement(self, elem):
        # Allow any element from the linked file
        if isinstance(self.doc.GetElement(elem), RevitLinkInstance):
            return True
        else:
            return False

class CameraFOVApp(Form):
    def __init__(self):
        System.Threading.Thread.Sleep(100)
        self.setupCameraSelectionGroup()
        self.factorIncrement = 0.2
        self.checkboxes = []
        self.colorIndicatorLabels = []
        self.filledRegionCreation()
        self.createFilledRegionComboBox()
        self.StartPosition = FormStartPosition.Manual  # Set start position to Manual
        self.Location = Point(400, 200)  # Set the start location of the form
        self.selected_link = None
        self.Text = "Camera FOV Configuration"
        self.Width = windowWidth
        self.Height = windowHeight
        self.toolTip = ToolTip()
        appName = "Draw FOV"
        self.TopMost = True
        self.titleBar = TitleBar(self, appName, logo_image, minimize_image, close_image)
        self.selected_cameras = []
        self.FormBorderStyle = FormBorderStyle.None
        self.Text = appName
        self.Size = Size(windowWidth, windowHeight)
        color3 = Color.FromArgb(49, 49, 49)
        color1 = Color.FromArgb(24, 24, 24)
        colorText = Color.FromArgb(240,240,240)
        panelSize = 30
        self.BackColor = color1
        self.ForeColor = colorText  
        self.expanded_width = 560  # Target expanded width
        self.original_width = 280  # Original width
        self.is_expanding = False  # Track direction of animation

        self.Controls.Add(self.titleBar)
        self.suffixLabel("°", 78, "fov_angle", "93")
        self.suffixLabel("px", 108, "horizontal_resolution", "1920")
        self.suffixLabel("m", 138, "max_distance", "0")
        # Labels and TextBoxes for user input with default values
        self.create_label_and_textbox("FOV Angle:", 78, "fov_angle", "93", "Enter cameras horizontal fov angle")
        # Replace the following line:
        # self.create_label_and_textbox("Horizontal Resolution:", 108, "horizontal_resolution", "1920", "Enter your cameras horizontal resolution")
        # With the ComboBox for horizontal resolution:
        self.create_label_and_combobox("Horizontal Resolution:", 108, "horizontal_resolution", ["3840","2592","2688","1920","1280","800","640"], "1920", "Select your cameras horizontal resolution")
        self.create_label_and_textbox("Max Distance:", 138, "max_distance", "0","Will be calculated once one DORI option is selected")

        self.run_button = PictureBox()
        self.run_button.Image = run_image
        self.run_button.Location = System.Drawing.Point(170, titleBar+420)
        self.run_button.Size = System.Drawing.Size(80,40)
        self.run_button.SizeMode = PictureBoxSizeMode.StretchImage
        self.run_button.Click += self.run_script
        self.run_buttonInteractive = InteractivePictureBox(
        self.run_button, 'Run.png', 'RunHover.png', 'RunClick.png')
        self.Controls.Add(self.run_button)
        self.AddDORIRadioButtons()
        self.setupRotationAngleControls()

        border1Text = Label()
        border1Text.AutoSize = False
        border1Text.TextAlign = ContentAlignment.MiddleLeft
        border1Text.Text = "Camera Parameters:"
        border1Text.Font = Font("Helvetica", 8, FontStyle.Regular)
        border1Text.Location = System.Drawing.Point(20,titleBar+10)
        border1Text.Size = System.Drawing.Size(106,10)
        border1Text.BackColor = Color.Transparent
        border1 = Panel()
        border1.Location = System.Drawing.Point(20,titleBar+20)
        border1.Size = System.Drawing.Size(windowWidth-40,110)#bottom=130
        border1.BackColor = Color.FromArgb(69,69,69)
        panel1 = Panel()
        panel1.Location = System.Drawing.Point(21,titleBar+21)
        panel1.Size = System.Drawing.Size(windowWidth-42,108)
        panel1.BackColor = Color.FromArgb(24,24,24)
        self.Controls.Add(border1Text)
        self.Controls.Add(panel1)
        self.Controls.Add(border1)

        # Create the PictureBox for displaying the pie chart
        self.pieChartPictureBox = PictureBox()
        self.pieChartPictureBox.SizeMode = PictureBoxSizeMode.Normal
        self.pieChartPictureBox.Size = Size(220, 220)  # Adjust size as needed
        self.pieChartPictureBox.Location = Point(310, 300)  # Adjust location as needed

        # Add PictureBox to the form's controls
        self.Controls.Add(self.pieChartPictureBox)

        self.expand_button = PictureBox()
        self.expand_button.Image = expand_image
        self.expand_button.Location = Point(self.Width - 35, self.Height - 60)  # Adjust as necessary
        self.expand_button.Size = Size(15, 15)  # Adjust as necessary
        self.expand_button.SizeMode = PictureBoxSizeMode.StretchImage
        #self.expand_button.Cursor = Cursors.Hand  # Change the cursor to indicate clickable
        self.expand_button.Click += self.toggle_expand
        self.Controls.Add(self.expand_button)
        # Ensure the expand button moves with window resizing

        self.expand_button.Anchor = AnchorStyles.Bottom | AnchorStyles.Right        
        self.animation_timer = Timer()
        self.animation_timer.Interval = 10  # Milliseconds between ticks, adjust for speed
        self.animation_timer.Tick += self.animate_resize

        # Handling form movement
        self.dragging = False
        self.offset = None

        self.toolTip = ToolTip()
        self.toolTip.SetToolTip(self.select_camera_button, "Select cameras from the project or a linked file")
        self.toolTip.SetToolTip(self.run_button, "Draw the FOV with the specified parameters")
        self.toolTip.SetToolTip(self.filledRegionTypeComboBox, "Select filled region type to be drawn")
        self.toolTip.SetToolTip(self.radio_current_project, "Cameras are located in the current project")
        self.toolTip.SetToolTip(self.radio_linked_file, "Cameras are located in a linked file")
        self.toolTip.SetToolTip(self.expand_button, "Create filled regions")

        self.FormClosing += self.on_form_closing  # Register the FormClosing event handler

        self.load_settings()  # Load settings during initialization

    def on_form_closing(self, sender, e):
        self.save_settings(sender, e)  # Save settings when the form is closed

    def draw_combo_item(self, sender, e):
        e.DrawBackground()
        e.Graphics.DrawString(self.filledRegionTypeComboBox.Items[e.Index].ToString(),
                            e.Font, System.Drawing.Brushes.White, e.Bounds, StringFormat.GenericDefault)
        e.DrawFocusRectangle()
        
    def suffixLabel(self, label_text, y, name, default_value="", label_size=(120, 20), textbox_size=(200, 20), label_color=System.Drawing.Color.LightGray, textbox_color=System.Drawing.Color.White, text_color=System.Drawing.Color.Black):
        suffixLabel = Label()
        suffixLabel.Text = label_text
        suffixLabel.Location = System.Drawing.Point(230, (y))
        suffixLabel.Font = Font("Helvetica", 9, FontStyle.Regular)
        suffixLabel.Size = System.Drawing.Size(18,15)
        suffixLabel.ForeColor = Color.FromArgb(140,140,140)
        suffixLabel.BackColor = Color.FromArgb(49, 49, 49)
        self.Controls.Add(suffixLabel)

    def create_label_and_textbox(self, label_text, y, name, default_value="", label_size=(120, 20), textbox_size=(200, 20), label_color=System.Drawing.Color.LightGray, textbox_color=System.Drawing.Color.White, text_color=System.Drawing.Color.Black, tooltip_text=""):
        label = Label()
        label.Text = label_text
        label.Location = System.Drawing.Point(30, (y-3))
        label.Font = Font("Helvetica", 10, FontStyle.Regular)
        label.Size = System.Drawing.Size(150,20)
        label.ForeColor = Color.FromArgb(240,240,240)
        
        textbox = RichTextBox()
        textbox.Location = System.Drawing.Point(172, y)
        textbox.Name = name
        textbox.Text = default_value
        textbox.Font = Font(default_value, 9, FontStyle.Regular)  # Adjust the font size as needed
        textbox.Size = System.Drawing.Size(57,15)
        textbox.BackColor = Color.FromArgb(49, 49, 49)
        textbox.ForeColor = Color.FromArgb(240,240,240)
        textbox.BorderStyle = BorderStyle.None
        textbox.Multiline = False

        textbox.SelectAll()
        textbox.SelectionAlignment = HorizontalAlignment.Center
        textbox.DeselectAll()

        boxBorder = Panel()
        boxBorder.Location = System.Drawing.Point(169, (y-3))
        boxBorder.Size = System.Drawing.Size(81,20)
        boxBorder.BackColor = Color.FromArgb(69,69,69)
        boxBorder.Anchor = AnchorStyles.Top | AnchorStyles.Left

        boxCover = Panel()
        boxCover.Location = System.Drawing.Point(170, (y-2))
        boxCover.Size = System.Drawing.Size(79,18)
        boxCover.BackColor = Color.FromArgb(49, 49, 49)
        boxCover.Anchor = AnchorStyles.Top | AnchorStyles.Left
    
        self.Controls.Add(textbox)
        self.Controls.Add(boxCover)
        self.Controls.Add(boxBorder)
        self.Controls.Add(label)
        self.toolTip.SetToolTip(textbox, tooltip_text)

    def create_label_and_combobox(self, label_text, y, name, options, default_value="", label_size=(120, 20), combobox_size=(200, 20), label_color=System.Drawing.Color.LightGray, combobox_color=System.Drawing.Color.White, text_color=System.Drawing.Color.Black, tooltip_text=""):
        label = Label()
        label.Text = label_text
        label.Location = System.Drawing.Point(30, (y-3))
        label.Font = Font("Helvetica", 10, FontStyle.Regular)
        label.Size = System.Drawing.Size(150,20)
        label.ForeColor = Color.FromArgb(240,240,240)
        
        combobox = ComboBox()
        combobox.Location = System.Drawing.Point(172, y)
        combobox.Name = name
        combobox.Items.AddRange(Array[Object](options))
        combobox.SelectedItem = default_value
        combobox.Font = Font(default_value, 9, FontStyle.Regular)  # Adjust the font size as needed
        combobox.Size = System.Drawing.Size(57,15)
        combobox.BackColor = Color.FromArgb(49, 49, 49)
        combobox.ForeColor = Color.FromArgb(240,240,240)
        combobox.DropDownStyle = ComboBoxStyle.DropDownList

        # Add event handler for SelectedIndexChanged event
        combobox.SelectedIndexChanged += self.on_combobox_selection_changed

        boxBorder = Panel()
        boxBorder.Location = System.Drawing.Point(169, (y-3))
        boxBorder.Size = System.Drawing.Size(81,20)
        boxBorder.BackColor = Color.FromArgb(69,69,69)
        boxBorder.Anchor = AnchorStyles.Top | AnchorStyles.Left

        boxCover = Panel()
        boxCover.Location = System.Drawing.Point(170, (y-2))
        boxCover.Size = System.Drawing.Size(79,18)
        boxCover.BackColor = Color.FromArgb(49, 49, 49)
        boxCover.Anchor = AnchorStyles.Top | AnchorStyles.Left
    
        self.Controls.Add(combobox)
        self.Controls.Add(boxCover)
        self.Controls.Add(boxBorder)
        self.Controls.Add(label)
        self.toolTip.SetToolTip(combobox, tooltip_text)

    def on_combobox_selection_changed(self, sender, event):
        self.update_distance()

    def createFilledRegionComboBox(self):
        self.border1 = Panel()
        self.border1.Location = Point(30,365)
        self.border1.Size = Size(220,2)
        self.border1.BackColor = Color.FromArgb(49, 49, 49)
        self.border2 = Panel()
        self.border2.Location = Point(30,384)
        self.border2.Size = Size(220,2)
        self.border2.BackColor = Color.FromArgb(49, 49, 49)
        self.border3 = Panel()
        self.border3.Location = Point(30,365)
        self.border3.Size = Size(2,20)
        self.border3.BackColor = Color.FromArgb(49, 49, 49)
        self.border4 = Panel()
        self.border4.Location = Point(231,365)
        self.border4.Size = Size(19,20)
        self.border4.BackColor = Color.FromArgb(49, 49, 49)
        self.border5 = Panel()
        self.border5.Location = Point((30-1),(365-1))
        self.border5.Size = Size(222,23)
        self.border5.BackColor = Color.FromArgb(69, 69, 69)
        
        self.Controls.Add(self.border2)
        self.Controls.Add(self.border1)
        self.Controls.Add(self.border3)
        self.Controls.Add(self.border4)

        # ComboBox for filled region type selection
        self.filledRegionTypeComboBox = ComboBox()
        self.filledRegionTypeComboBox.DropDownStyle = ComboBoxStyle.DropDownList
        self.filledRegionTypeComboBox.Location = Point(30, 365)
        self.filledRegionTypeComboBox.Width = 220
        self.filledRegionTypeComboBox.BackColor = Color.FromArgb(49, 49, 49)  # Example color
        self.filledRegionTypeComboBox.ForeColor = Color.FromArgb(240, 240, 240)  # Example text color
        self.filledRegionTypeComboBox.DrawMode = DrawMode.OwnerDrawFixed
        self.filledRegionTypeComboBox.DrawItem += self.draw_combo_item

        filled_region_types = list_filled_region_type_names_and_ids(doc)
        for name, _ in filled_region_types:
            self.filledRegionTypeComboBox.Items.Add(name)
        if filled_region_types:
            self.filledRegionTypeComboBox.SelectedIndex = 0
        self.Controls.Add(self.filledRegionTypeComboBox)
        self.Controls.Add(self.border5)
    def update_filled_region_combobox(self, sender, event):
        # Create a new list for the updated filled region types
        new_items = []
        filled_region_types = list_filled_region_type_names_and_ids(doc)
        for name, _ in filled_region_types:
            new_items.append(name)
        
        # Remove existing items by iterating backwards to avoid modifying the collection while iterating
        for i in range(self.filledRegionTypeComboBox.Items.Count - 1, -1, -1):
            self.filledRegionTypeComboBox.Items.RemoveAt(i)
        
        # Add new items from the list
        for item in new_items:
            self.filledRegionTypeComboBox.Items.Add(item)
        
        # Safely update the selected index
        self.filledRegionTypeComboBox.SelectedIndex = 0 if filled_region_types else -1
    def filledRegionCreation(self,startX=300):
        FRC_borderText = Label()
        FRC_borderText.AutoSize = False
        FRC_borderText.TextAlign = ContentAlignment.MiddleLeft
        FRC_borderText.Text = "Create Filled Regions:"
        FRC_borderText.Font = Font("Helvetica", 8, FontStyle.Regular)
        FRC_borderText.Location = System.Drawing.Point((startX), (titleBar+10))
        FRC_borderText.Size = System.Drawing.Size(120, 10)
        FRC_borderText.BackColor = Color.Transparent

        FRC_border = Panel()
        FRC_border.Location = System.Drawing.Point((startX), (titleBar+20))
        FRC_border.Size = System.Drawing.Size(windowWidth-40, 202)
        FRC_border.BackColor = Color.FromArgb(69,69,69)

        FRC_panel = Panel()
        FRC_panel.Location = System.Drawing.Point((startX+1), (titleBar+21))
        FRC_panel.Size = System.Drawing.Size(windowWidth-42, 200)
        FRC_panel.BackColor = Color.FromArgb(24,24,24)

        prefix_label = Label()
        prefix_label.Text = "Prefix Name:"
        prefix_label.Location = System.Drawing.Point(10, (17+-3))
        prefix_label.Font = Font("Helvetica", 10, FontStyle.Regular)
        prefix_label.Size = System.Drawing.Size(150,20)
        prefix_label.ForeColor = Color.FromArgb(240,240,240)   
        
        self.prefix_textbox = RichTextBox()
        self.prefix_textbox.Location = System.Drawing.Point(152, 17)
        self.prefix_textbox.Font = Font("Helvetica", 9, FontStyle.Regular)  # Adjust the font size as needed
        self.prefix_textbox.Size = System.Drawing.Size(57,15)
        self.prefix_textbox.BackColor = Color.FromArgb(49, 49, 49)
        self.prefix_textbox.ForeColor = Color.FromArgb(240,240,240)
        self.prefix_textbox.BorderStyle = BorderStyle.None
        self.prefix_textbox.Multiline = False
        self.prefix_textbox.SelectionAlignment = HorizontalAlignment.Center
        FRC_panel.Controls.Add(self.prefix_textbox)

        prefix_boxCover = Panel()
        prefix_boxCover.Location = System.Drawing.Point(150, (17-2))
        prefix_boxCover.Size = System.Drawing.Size(79,18)
        prefix_boxCover.BackColor = Color.FromArgb(49, 49, 49)
        prefix_boxCover.Anchor = AnchorStyles.Top | AnchorStyles.Left
        FRC_panel.Controls.Add(prefix_boxCover)

        prefix_boxBorder = Panel()
        prefix_boxBorder.Location = System.Drawing.Point(149, (17-3))
        prefix_boxBorder.Size = System.Drawing.Size(81,20)
        prefix_boxBorder.BackColor = Color.FromArgb(69,69,69)
        prefix_boxBorder.Anchor = AnchorStyles.Top | AnchorStyles.Left
        FRC_panel.Controls.Add(prefix_boxBorder)

        FRC_selectColor = PictureBox()
        FRC_selectColor.Image = run_image
        FRC_selectColor.Location = System.Drawing.Point(10, 45)
        FRC_selectColor.Size = System.Drawing.Size(80,40)
        FRC_selectColor.SizeMode = PictureBoxSizeMode.StretchImage
        FRC_selectColor.Click += self.SelectColor
        self.color_buttonInteractive = InteractivePictureBox(
        FRC_selectColor, 'SelectColor.png', 'SelectColorHover.png', 'SelectColorClick.png')
        FRC_panel.Controls.Add(FRC_selectColor)

        FRC_create = PictureBox()
        FRC_create.Image = run_image
        FRC_create.Location = System.Drawing.Point(150, 45)
        FRC_create.Size = System.Drawing.Size(80,40)
        FRC_create.SizeMode = PictureBoxSizeMode.StretchImage
        FRC_create.Click += self.create_filled_region
        self.createRegion_buttonInteractive = InteractivePictureBox(
        FRC_create, 'CreateRegion.png', 'CreateRegionHover.png', 'CreateRegionClick.png')
        FRC_panel.Controls.Add(FRC_create)
    
        self.checkboxes = []
        checkBoxNames = ["0px", "25px", "63px", "125px", "250px"]
        checkBoxX = 10
        checkBoxY = 105
        colorIndicatorYOffset = -10
        for i,name in enumerate(checkBoxNames):
            checkBox = CheckBox()
            checkBox.Text = name
            checkBox.ForeColor = Color.FromArgb(49, 49, 49)
            checkBox.Location = Point(checkBoxX, checkBoxY)
            checkBox.Size = Size(20, 20)
            checkBox.Enabled = False
            FRC_panel.Controls.Add(checkBox)
            checkBox.CheckedChanged += self.updateColorDisplay
            self.checkboxes.append(checkBox)

            label = Label()
            label.Text = name
            label.Font = Font("Helvetica", 10, FontStyle.Regular)
            label.Location = Point(checkBoxX - 20, (checkBoxY+20))
            label.Size = Size(50, 15)
            label.TextAlign = ContentAlignment.TopCenter
            FRC_panel.Controls.Add(label)

            # Create a list to hold the color indicator labels for this checkbox
            indicatorLabelsForThisCheckbox = []
            for i in range(5):
                colorIndicatorLabel = Label()
                colorIndicatorLabel.Location = Point(checkBoxX-15, checkBoxY + colorIndicatorYOffset)  # Adjust position for each label
                colorIndicatorLabel.Size = Size(44, 10)  # Adjusted size to 10x10 as requested
                colorIndicatorLabel.BackColor = Color.FromArgb(24, 24, 24)
                FRC_panel.Controls.Add(colorIndicatorLabel)
                indicatorLabelsForThisCheckbox.append(colorIndicatorLabel)
            self.colorIndicatorLabels.append(indicatorLabelsForThisCheckbox)  # Add the list of labels to the main list

            checkBoxX += 50

        FRC_panel.Controls.Add(prefix_label)
        self.Controls.Add(FRC_borderText)
        self.Controls.Add(FRC_panel)
        self.Controls.Add(FRC_border)

        self.toolTip = ToolTip()
        self.toolTip.SetToolTip(FRC_selectColor, "Select primary color")
        self.toolTip.SetToolTip(FRC_create, "Create selected filled regions")
        self.toolTip.SetToolTip(self.prefix_textbox, "Add a prefix to be used in created filled region names")

    def lighten_color(self, color, factor):
        """Lightens the given color by mixing it with white."""
        white = RevitColor(255, 255, 255)
        new_red = int((color.Red * (1 - factor)) + (white.Red * factor))
        new_green = int((color.Green * (1 - factor)) + (white.Green * factor))
        new_blue = int((color.Blue * (1 - factor)) + (white.Blue * factor))
        
        # Clamping the RGB values to ensure they are within 0-255
        new_red = max(0, min(255, new_red))
        new_green = max(0, min(255, new_green))
        new_blue = max(0, min(255, new_blue))
        
        return RevitColor(new_red, new_green, new_blue)  
    def SelectColor(self, sender, args):
        colorDialog = ColorDialog()
        if colorDialog.ShowDialog() == DialogResult.OK:
            self.selectedColor = (colorDialog.Color.R, colorDialog.Color.G, colorDialog.Color.B)
            # Enable checkboxes after color is selected
            for checkBox in self.checkboxes:
                checkBox.Enabled = True
            # Update the display immediately upon selection
            self.updateColorDisplay(None, None)
    def updateColorDisplay(self, sender, args):
        if hasattr(self, 'selectedColor'):
            r, g, b = self.selectedColor
            selectedColor = RevitColor(r, g, b)
            whiteColor = RevitColor(255, 255, 255)

            checkedIndices = [i for i, cb in enumerate(self.checkboxes) if cb.Checked]
            checkedIndices.reverse()  # Reverse the order for color application

            for labels in self.colorIndicatorLabels:
                for label in labels:
                    label.BackColor = System.Drawing.Color.FromArgb(24, 24, 24)  # Reset color

            if self.checkboxes[0].Checked:
                for label in self.colorIndicatorLabels[0]:
                    label.BackColor = System.Drawing.Color.FromArgb(whiteColor.Red, whiteColor.Green, whiteColor.Blue)

            for i, index in enumerate(checkedIndices):
                if index == 0:
                    continue  # Skip "0px" since it's always white
                factor = (i - 1 if checkedIndices[0] == 0 else i) * 0.2
                lightenedColor = self.lighten_color(selectedColor, factor)
                sysColor = System.Drawing.Color.FromArgb(lightenedColor.Red, lightenedColor.Green, lightenedColor.Blue)

                for label in self.colorIndicatorLabels[index]:
                    label.BackColor = sysColor

                self.draw_pie_chart(self.pieChartPictureBox)
        else:
            print("No color selected yet.")
    def draw_pie_chart(self, pie_chart_area):
        # Prepare the drawing surface
        if pie_chart_area.Image is None:
            pie_chart_area.Image = Bitmap(pie_chart_area.Width, pie_chart_area.Height)
        graphics = Graphics.FromImage(pie_chart_area.Image)
        graphics.SmoothingMode = SmoothingMode.AntiAlias

        # Define the pie chart rectangle area
        chart_area = Rectangle(10, 10, 200, 200)

        # Initialize colors list and names list
        active_colors = []
        checkbox_names = []

        # Determine active colors and names based on selected checkboxes
        if hasattr(self, 'selectedColor'):
            r, g, b = self.selectedColor
            selectedColor = RevitColor(r, g, b)
            whiteColor = RevitColor(255, 255, 255)  # White color for "0px"
            factorIncrement = 0.2
            checkedIndices = [i for i, cb in enumerate(self.checkboxes) if cb.Checked]

            # Include "0px" color if selected
            if self.checkboxes[0].Checked:
                active_colors.append(Color.FromArgb(whiteColor.Red, whiteColor.Green, whiteColor.Blue))
                checkbox_names.append(self.checkboxes[0].Text)

            # Apply colors to checked checkboxes beyond "0px"
            for i, index in enumerate(checkedIndices):
                if index == 0: continue  # Already handled "0px"
                factor = (i - 1 if checkedIndices[0] == 0 else i) * factorIncrement
                lightenedColor = self.lighten_color(selectedColor, factor)
                sysColor = Color.FromArgb(lightenedColor.Red, lightenedColor.Green, lightenedColor.Blue)
                active_colors.append(sysColor)
                checkbox_names.append(self.checkboxes[index].Text)

        # Reverse the order of colors and names
        active_colors.reverse()
        checkbox_names.reverse()

        # Draw pie chart segments and names
        segment_angle = 360 / len(active_colors) if active_colors else 360
        start_angle = 270  # Starting at the top

        for i, color in enumerate(active_colors):
            brush = SolidBrush(color)
            graphics.FillPie(brush, chart_area, start_angle, segment_angle)

            # Calculate the midpoint angle of the segment for text placement
            midpoint_angle = start_angle + segment_angle / 2
            radians = math.radians(midpoint_angle)
            radius = chart_area.Width / 4  # Adjust radius for text position
            text_x = chart_area.X + (chart_area.Width / 2) + (radius * cos(radians)) - 10
            text_y = chart_area.Y + (chart_area.Height / 2) + (radius * sin(radians)) - 10

            # Draw the checkbox name at the segment's center
            graphics.DrawString(checkbox_names[i], Font("Helvetica", 10), Brushes.Black, float(text_x), float(text_y))

            start_angle += segment_angle  # Move to the next segment

        pie_chart_area.Refresh()  # Refresh to display the updated pie chart

    def setupRotationAngleControls(self):
        border3Text = Label()
        border3Text.AutoSize = False
        border3Text.TextAlign = ContentAlignment.MiddleLeft
        border3Text.Text = "Camera Rotation:"
        border3Text.Font = Font("Helvetica", 8, FontStyle.Regular)
        border3Text.Location = System.Drawing.Point(20, titleBar+232)
        border3Text.Size = System.Drawing.Size(106, 10)
        border3Text.BackColor = Color.Transparent

        border3 = Panel()
        border3.Location = System.Drawing.Point(20, titleBar+242)
        border3.Size = System.Drawing.Size(windowWidth-40, 122)
        border3.BackColor = Color.FromArgb(69,69,69)

        panel3 = Panel()
        panel3.Location = System.Drawing.Point(21, titleBar+243)
        panel3.Size = System.Drawing.Size(windowWidth-42, 120)
        panel3.BackColor = Color.FromArgb(24,24,24)

        rotoSuffixLabel = Label()
        rotoSuffixLabel.Text = "°"
        rotoSuffixLabel.Location = System.Drawing.Point(230, 300)
        rotoSuffixLabel.Font = Font("Helvetica", 9, FontStyle.Regular)
        rotoSuffixLabel.Size = System.Drawing.Size(18,15)
        rotoSuffixLabel.ForeColor = Color.FromArgb(140,140,140)
        rotoSuffixLabel.BackColor = Color.FromArgb(49, 49, 49)
        self.Controls.Add(rotoSuffixLabel)

        rotationAngleLabel = Label()
        rotationAngleLabel.Text = "Rotation Angle:"
        rotationAngleLabel.Font = Font("Helvetica", 10, FontStyle.Regular)
        rotationAngleLabel.Location = Point(30, 297)  # Adjust location as needed
        rotationAngleLabel.Size = Size(120, 20)  # Adjust size as needed
        rotationAngleLabel.ForeColor = Color.FromArgb(240, 240, 240)
        self.Controls.Add(rotationAngleLabel)
        
        # Textbox for Rotation Angle
        rotationAngleTextBox = RichTextBox()
        rotationAngleTextBox.Name = "rotation_angle"
        rotationAngleTextBox.Text = "0"  # Default value
        rotationAngleTextBox.Location = Point(172, 300)  # Adjust location to align with the label
        rotationAngleTextBox.Size = Size(57, 15)  # Adjust size as needed
        rotationAngleTextBox.BackColor = Color.FromArgb(49, 49, 49)
        rotationAngleTextBox.ForeColor = Color.FromArgb(240, 240, 240)
        rotationAngleTextBox.BorderStyle = BorderStyle.None
        rotationAngleTextBox.Multiline = False
        rotationAngleTextBox.SelectAll()
        rotationAngleTextBox.SelectionAlignment = HorizontalAlignment.Center
        rotationAngleTextBox.DeselectAll()
        self.Controls.Add(rotationAngleTextBox)

        rotoBoxBorder = Panel()
        rotoBoxBorder.Location = System.Drawing.Point(169, (300-3))
        rotoBoxBorder.Size = System.Drawing.Size(81,20)
        rotoBoxBorder.BackColor = Color.FromArgb(69,69,69)
        rotoBoxBorder.Anchor = AnchorStyles.Top | AnchorStyles.Left

        rotoBoxCover = Panel()
        rotoBoxCover.Location = System.Drawing.Point(170, (300-2))
        rotoBoxCover.Size = System.Drawing.Size(79,18)
        rotoBoxCover.BackColor = Color.FromArgb(49, 49, 49)
        rotoBoxCover.Anchor = AnchorStyles.Top | AnchorStyles.Left
        self.Controls.Add(rotoBoxCover)
        self.Controls.Add(rotoBoxBorder)

        # Create a panel to hold the radio buttons
        angleSelectionPanel = Panel()
        angleSelectionPanel.Location = Point(30, 330)  # Adjust location as needed
        angleSelectionPanel.Size = Size(220, 20)  # Adjust size to fit the radio buttons

        angles = [0, 90, 180, 270]
        self.radio_angles = {}

        for i, angle in enumerate(angles):
            radio_button = RadioButton()
            radio_button.Text = "%d°" % angle
            radio_button.Font = Font("Helvetica", 8, FontStyle.Regular)
            radio_button.Location = Point(10 + i * (40 + 14), 0)  # Adjust spacing as needed
            radio_button.Size = Size(50, 20)
            radio_button.ForeColor = Color.FromArgb(240, 240, 240)  # Adjust text color as needed
            radio_button.Tag = angle
            radio_button.CheckedChanged += self.on_angle_changed  # Attach the event handler
            if angle == 0:
                radio_button.Checked = True  # Default selection

            # Add each radio button to the panel instead of the form
            angleSelectionPanel.Controls.Add(radio_button)
            self.radio_angles[angle] = radio_button

        # Add the panel to the form
        self.Controls.Add(angleSelectionPanel)
        self.Controls.Add(border3Text)
        self.Controls.Add(panel3)
        self.Controls.Add(border3)

        self.toolTip = ToolTip()
        self.toolTip.SetToolTip(rotationAngleTextBox, "Rotate the camera")

    def setupCameraSelectionGroup(self):
        border2Text = Label()
        border2Text.AutoSize = False
        border2Text.TextAlign = ContentAlignment.MiddleLeft
        border2Text.Text = "Camera Selection:"
        border2Text.Font = Font("Helvetica", 8, FontStyle.Regular)
        border2Text.Location = System.Drawing.Point(20, titleBar+140)
        border2Text.Size = System.Drawing.Size(106, 10)
        border2Text.BackColor = Color.Transparent

        border2 = Panel()
        border2.Location = System.Drawing.Point(20, titleBar+150)#bottom=222
        border2.Size = System.Drawing.Size(windowWidth-40, 72)
        border2.BackColor = Color.FromArgb(69,69,69)

        panel2 = Panel()
        panel2.Location = System.Drawing.Point(21, titleBar+151)
        panel2.Size = System.Drawing.Size(windowWidth-42, 70)
        panel2.BackColor = Color.FromArgb(24,24,24)

        # RadioButton for selecting cameras in the current project
        self.radio_current_project = RadioButton()
        self.radio_current_project.Text = "Current project"
        self.radio_current_project.Font = Font("Helvetica", 8, FontStyle.Regular)
        self.radio_current_project.Location = System.Drawing.Point(10, 10)  # Adjusted location
        self.radio_current_project.Size = System.Drawing.Size(140, 20)
        self.radio_current_project.Checked = True

        # RadioButton for selecting cameras from a linked file
        self.radio_linked_file = RadioButton()
        self.radio_linked_file.Text = "Linked file"
        self.radio_linked_file.Font = Font("Helvetica", 8, FontStyle.Regular)
        self.radio_linked_file.Location = System.Drawing.Point(10, 35)  # Adjusted location
        self.radio_linked_file.Size = System.Drawing.Size(140, 20)

        # Button for selecting camera
        self.select_camera_button = PictureBox()
        self.select_camera_button.Location = System.Drawing.Point(150, 15)  # Adjusted location
        self.select_camera_button.Size = System.Drawing.Size(80, 40)
        self.select_camera_button.SizeMode = PictureBoxSizeMode.StretchImage
        self.select_camera_button.Click += self.select_camera
        self.select_camera_buttonInteractive = InteractivePictureBox(
            self.select_camera_button, 'Select.png', 'SelectHover.png', 'SelectClick.png')

        # Add the controls to panel2 instead of the form
        panel2.Controls.Add(self.radio_current_project)
        panel2.Controls.Add(self.radio_linked_file)
        panel2.Controls.Add(self.select_camera_button)

        # Finally, add panel2 (and its border) to the form
        self.Controls.Add(border2Text)
        self.Controls.Add(panel2)
        self.Controls.Add(border2)
    def animate_resize(self, sender, event):
        step = 10  # Size change per tick, adjust for smoothness vs speed
        if self.is_expanding:
            if self.Width < self.expanded_width:
                self.Width += step
            else:
                self.animation_timer.Stop()
                self.Width = self.expanded_width  # Ensure it ends exactly at target size
                self.expand_button.Image = contract_image  # Change to contract image
                self.toolTip.SetToolTip(self.expand_button, "Close")
        else:
            if self.Width > self.original_width:
                self.Width -= step
            else:
                self.animation_timer.Stop()
                self.Width = self.original_width  # Ensure it ends exactly at target size
                self.expand_button.Image = expand_image  # Change to expand image
                self.toolTip.SetToolTip(self.expand_button, "Create filled regions")
    def toggle_expand(self, sender, event):
        # Toggle the direction of expansion based on the current width
        self.is_expanding = self.Width == self.original_width
        self.animation_timer.Start()  # Start the animation
    def AddDORIRadioButtons(self):
        border4Text = Label()
        border4Text.AutoSize = False
        border4Text.TextAlign = ContentAlignment.MiddleLeft
        border4Text.Text = "Draw FOV:"
        border4Text.Font = Font("Helvetica", 8, FontStyle.Regular)
        border4Text.Location = System.Drawing.Point(20, titleBar+374)
        border4Text.Size = System.Drawing.Size(106, 10)
        border4Text.BackColor = Color.Transparent

        border4 = Panel()
        border4.Location = System.Drawing.Point(20, titleBar+384)#bottom=222
        border4.Size = System.Drawing.Size(windowWidth-40, 122)
        border4.BackColor = Color.FromArgb(69,69,69)

        panel4 = Panel()
        panel4.Location = System.Drawing.Point(21, titleBar+385)
        panel4.Size = System.Drawing.Size(windowWidth-42, 120)
        panel4.BackColor = Color.FromArgb(24,24,24)

        # Panel for DORI Options
        doriOptionsPanel = Panel()
        doriOptionsPanel.Location = Point(30, 434)  # Adjust as needed, positioning below the label
        doriOptionsPanel.Size = Size(140, 110)  # Adjust as needed
        doriOptionsPanel.BackColor = Color.FromArgb(24, 24, 24)  # Optional: set background color
        self.Controls.Add(doriOptionsPanel)

        self.doriOptions = ["Detection (25px/m)", "Observation (63px/m)", "Recognition (125px/m)", "Identification (250px/m)"]
        startY = 0  # Starting Y position for the first radio button within the panel

        for i, option in enumerate(self.doriOptions):
            radioButtonDORI = RadioButton()
            radioButtonDORI.Text = option
            radioButtonDORI.Font = Font("Helvetica", 8, FontStyle.Regular)
            radioButtonDORI.Location = Point(0, startY + (i * 25))
            radioButtonDORI.Size = Size(150, 20)
            radioButtonDORI.Tag = option
            radioButtonDORI.CheckedChanged += self.onDORIOptionChanged
            doriOptionsPanel.Controls.Add(radioButtonDORI)  # Add to the panel instead of the form directly

        # Finally, add panel2 (and its border) to the form
        self.Controls.Add(border4Text)
        self.Controls.Add(panel4)
        self.Controls.Add(border4)
    def onDORIOptionChanged(self, sender, event):
        if sender.Checked:
            self.selected_dori_option_index = self.doriOptions.index(sender.Tag)
            self.update_distance()
            
    def update_distance(self):
        # Safely access the text boxes, checking if they actually exist
        hrTextBox = self.Controls.Find("horizontal_resolution", True)
        fovTextBox = self.Controls.Find("fov_angle", True)
        maxDistanceTextBox = self.Controls.Find("max_distance", True)[0]  # Directly access the RichTextBox

        if hrTextBox and fovTextBox and maxDistanceTextBox:
            try:
                hr = int(hrTextBox[0].Text)
                fov = int(fovTextBox[0].Text)
                distances = calculator_1(hr, fov)
                doriIndex = 0  # Default to the first DORI option (or change as needed)
                if hasattr(self, 'selected_dori_option_index'):
                    doriIndex = self.selected_dori_option_index
                
                calculatedValue = distances[doriIndex]
                
                # Update the max_distance textbox with the calculated value
                maxDistanceTextBox.Text = str(calculatedValue)
                
                # Reset the alignment to center after changing the text
                maxDistanceTextBox.SelectAll()
                maxDistanceTextBox.SelectionAlignment = HorizontalAlignment.Center
                maxDistanceTextBox.DeselectAll()
            except ValueError:
                # Handle cases where conversion to int fails
                MessageBox.Show("Please enter valid numeric values for horizontal resolution and FOV angle.")
        else:
            MessageBox.Show("One or more required fields are missing.")

    def on_angle_changed(self, sender, event):
        if sender.Checked:
            self.additional_rotation_angle = float(sender.Tag)  # Assuming sender.Tag holds the angle in degrees
    def select_camera(self, sender, event):
        self.WindowState = FormWindowState.Minimized
        if self.radio_current_project.Checked:
            self.select_cameras_current_project()
        elif self.radio_linked_file.Checked:
            self.select_cameras_linked_file()
        self.WindowState = FormWindowState.Normal
        self.Activate()
    def select_cameras_current_project(self):
        self.selected_cameras = []  # Reset the selected cameras list
        try:
            refs = uidoc.Selection.PickObjects(ObjectType.Element, "Please select cameras.")
            for ref in refs:
                camera = doc.GetElement(ref.ElementId)
                rotation_param = camera.LookupParameter("Camera Rotation")
                rotation_angle = math.degrees(rotation_param.AsDouble()) if rotation_param else 0
                camera_info = (camera, False, rotation_angle)  # Include rotation angle in camera_info
                self.selected_cameras.append(camera_info)
            self.Activate()  # Bring the window back into focus
        except Exception as e:
            MessageBox.Show("An error occurred during camera selection: " + str(e))
            self.Activate()
    def select_cameras_linked_file(self):
        self.selected_cameras = []  # Reset the list
        try:
            selectedObjs = uidoc.Selection.PickObjects(ObjectType.LinkedElement, "Select Linked Elements")
            for selectedObj in selectedObjs:
                linkInstance = doc.GetElement(selectedObj.ElementId)
                linkedDoc = linkInstance.GetLinkDocument()
                linkedCameraElement = linkedDoc.GetElement(selectedObj.LinkedElementId)
                rotation_param = linkedCameraElement.LookupParameter("Camera Rotation")
                rotation_angle = math.degrees(rotation_param.AsDouble()) if rotation_param else 0
                transformedPosition = linkInstance.GetTransform().OfPoint(linkedCameraElement.Location.Point)
                camera_info = (transformedPosition, True, rotation_angle)  # Include rotation angle in camera_info
                self.selected_cameras.append(camera_info)
            self.Activate()
        except Exception as e:
            MessageBox.Show("An error occurred during camera selection: " + str(e))
            self.Activate()
    def create_filled_region(self, sender, args):
        if not any(cb.Checked for cb in self.checkboxes):
            MessageBox.Show("Please select at least one checkbox.")
            return
        transaction_title = "Create Filled Regions Based on Selection"
        prefix = self.prefix_textbox.Text.strip()  # Use strip() to remove leading/trailing whitespace
        
        # Fetch selected line weights
        selectedWeights = [cb.Text for cb in self.checkboxes if cb.Checked]
        
        # Original color specified by the user
        originalColor = RevitColor(*self.selectedColor)  # Unpacking the RGB tuple
        whiteColor = RevitColor(255, 255, 255)  # White color for the "0px" checkbox

        message = ""  # Initialize an empty message string

        with Transaction(doc, transaction_title) as t:
            t.Start()
            
            for index, weight in enumerate(selectedWeights):
                base_name = prefix + ("_" if prefix else "") + weight  # Base name without the copy suffix
                
                # Initialize the suffix and name
                suffix = 0
                unique_region_name = base_name
                
                # Loop to find a unique name
                while get_filled_region(unique_region_name):
                    suffix += 1
                    unique_region_name = "{} Copy {}".format(base_name, suffix)
                
                # Determine the color based on the checkbox index
                if weight == "0px":
                    currentColor = whiteColor
                else:
                    factor = (index - 1) * 0.2
                    currentColor = self.lighten_color(originalColor, min(factor, 0.8))
                
                # Create the filled region type with the unique name and adjusted color
                create_RegionType(name=unique_region_name, color=currentColor)
                message += "Created Filled Region Type: {}\n".format(unique_region_name)  # Append to message string
            self.update_filled_region_combobox(self, sender)
            t.Commit()

        if message:
            MessageBox.Show(message, "Creation Summary")  # Show message box with all messages
        else:
            MessageBox.Show("No filled regions were created.", "Creation Summary")

    def setComboBoxSelection(self):
        # Ensure there are items before setting the selected index
        if self.filledRegionTypeComboBox.Items.Count > 0:
            self.filledRegionTypeComboBox.SelectedIndex = 0  # or another valid index
    def run_script(self, sender, event):
        try:
            fov_angle = float(Decimal(self.Controls["fov_angle"].Text))
            max_distance_m = float(Decimal(self.Controls["max_distance"].Text))
            max_distance_mm = max_distance_m * 1000  # Convert from meters to millimeters
            rotation_angle_input = float(Decimal(self.Controls["rotation_angle"].Text))  # Get the manual input rotation
            
            # Check if there's a valid selection before proceeding
            if self.filledRegionTypeComboBox.SelectedIndex >= 0:
                selected_filled_region_name = self.filledRegionTypeComboBox.SelectedItem.ToString()
                selected_filled_region_id = None
                for name, id in list_filled_region_type_names_and_ids(doc):
                    if name == selected_filled_region_name:
                        selected_filled_region_id = id
                        break

                if selected_filled_region_id is not None:
                    detail_lines = get_custom_detail_lines(doc, "Boundary")
                    for camera_info in self.selected_cameras:
                        camera_position, from_linked_file, camera_rotation_angle = camera_info
                        final_rotation_angle = camera_rotation_angle + rotation_angle_input + self.additional_rotation_angle
                        main_script((camera_position, from_linked_file, final_rotation_angle), fov_angle, max_distance_mm, detail_lines, selected_filled_region_id)
                else:
                    MessageBox.Show("Selected filled region type not found.")
            else:
                MessageBox.Show("Please select a filled region type.")
        except Exception as e:
            MessageBox.Show("Invalid input: " + str(e))
    def load_settings(self):
        config = configparser.ConfigParser()
        if os.path.exists('settings.ini'):
            config.read('settings.ini')
            if config.has_section('Settings'):
                if config.has_option('Settings', 'fov_angle'):
                    self.Controls["fov_angle"].Text = config.get('Settings', 'fov_angle')
                if config.has_option('Settings', 'horizontal_resolution'):
                    self.Controls["horizontal_resolution"].Text = config.get('Settings', 'horizontal_resolution')

    def save_settings(self, sender, e):
        config = configparser.ConfigParser()
        config.add_section('Settings')
        config.set('Settings', 'fov_angle', self.Controls["fov_angle"].Text)
        config.set('Settings', 'horizontal_resolution', self.Controls["horizontal_resolution"].Text)
        with open('settings.ini', 'w') as configfile:
            config.write(configfile)

if __name__ == "__main__":
    Application.EnableVisualStyles()
    form = CameraFOVApp()
    Application.Run(form)
