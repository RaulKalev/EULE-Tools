# -*- coding: utf-8 -*-
import clr
import System
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from Autodesk.Revit.DB import (XYZ, Line, Transaction, ViewDetailLevel, RevitLinkInstance,
                               ReferenceIntersector, FindReferenceTarget, FilledRegion, FilledRegionType, CurveLoop,
                               ElementCategoryFilter, BuiltInCategory, ViewFamilyType, ElementId, Element,
                               FilteredElementCollector, View3D, ElementId, Curve, CurveElement, ViewPlan)
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from System.Windows.Forms import Application, Form, Button, Label, TextBox, Timer, Cursors, DrawMode, DialogResult,FormWindowState, MessageBox, FormBorderStyle, BorderStyle, RadioButton, ListBox, GroupBox, ComboBox, ComboBoxStyle, AnchorStyles, Panel, PictureBoxSizeMode, PictureBox, ToolTip,FormStartPosition
from System.Drawing import Color, Size, Point, Bitmap, Font, FontStyle,ContentAlignment,StringFormat
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

# Create Bitmap objects from the paths
minimize_image = Bitmap(minimize_image_path)
close_image = Bitmap(close_image_path)
clear_image = Bitmap(clear_image_path)
logo_image = Bitmap(logo_image_path)
run_image = Bitmap(run_image_path)
select_image = Bitmap(select_image_path)
expand_image = Bitmap(expand_image_path)
contract_image = Bitmap(contract_image_path)

class DetailLineFilter(ISelectionFilter):
    def AllowElement(self, elem):
        return elem.Category.Id.IntegerValue == int(BuiltInCategory.OST_Lines)

# Main script execution
def main_script(camera_info, fov_angle, max_distance_mm, rotation_angle, detail_lines, filled_region_type_id):
    camera_position, from_linked_file = camera_info
    if from_linked_file:
        # If camera is from a linked file, the position is already transformed
        camera_position = camera_info[0]
    else:
        # If camera is from the current project
        camera_element = camera_info[0]
        camera_position = camera_element.GetTransform().Origin
    
    activeView = doc.ActiveView  # Get the active view from the document

    if not isinstance(activeView, ViewPlan):
        print("The active view is not a plan view. Please switch to a plan view to draw detail lines.")
        return

    # Convert Decimal to float
    fov_angle = float(fov_angle)
    max_distance_mm = float(max_distance_mm)
    rotation_angle = float(rotation_angle)

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
        self.setupCameraSelectionGroup()
        self.StartPosition = FormStartPosition.Manual  # Set start position to Manual
        self.Location = Point(400, 200)  # Set the start location of the form
        self.selected_link = None
        self.Text = "Camera FOV Configuration"
        self.Width = windowWidth
        self.Height = windowHeight
        appName = "FOV Magic"
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
        self.expanded_width = 580  # Target expanded width
        self.original_width = 280  # Original width
        self.is_expanding = False  # Track direction of animation

        self.Controls.Add(self.titleBar)

        # Labels and TextBoxes for user input with default values
        self.create_label_and_textbox("FOV Angle:", 78, "fov_angle", "55")
        self.create_label_and_textbox("Horizontal Resolution:", 108, "horizontal_resolution", "2560")
        self.create_label_and_textbox("Max Distance (m):", 138, "max_distance", "0")

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

        # Run script button
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

    #Groups
        #Group1
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
        ##############################

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
        self.toolTip.SetToolTip(self.expand_button, "Advanced features")

    def draw_combo_item(self, sender, e):
        e.DrawBackground()
        e.Graphics.DrawString(self.filledRegionTypeComboBox.Items[e.Index].ToString(),
                            e.Font, System.Drawing.Brushes.White, e.Bounds, StringFormat.GenericDefault)
        e.DrawFocusRectangle()
    def create_label_and_textbox(self, label_text, y, name, default_value="", label_size=(120, 20), textbox_size=(200, 20), label_color=System.Drawing.Color.LightGray, textbox_color=System.Drawing.Color.White, text_color=System.Drawing.Color.Black):
        label = Label()
        label.Text = label_text
        label.Location = System.Drawing.Point(30, (y-3))
        label.Font = Font("Helvetica", 10, FontStyle.Regular)
        label.Size = System.Drawing.Size(150,20)
        label.ForeColor = Color.FromArgb(240,240,240)
        
        textbox = TextBox()
        textbox.Location = System.Drawing.Point(172, y)
        textbox.Name = name
        textbox.Text = default_value
        textbox.Font = Font(default_value, 9, FontStyle.Regular)  # Adjust the font size as needed
        textbox.Size = System.Drawing.Size(75,20)
        textbox.BackColor = Color.FromArgb(49, 49, 49)
        textbox.ForeColor = Color.FromArgb(240,240,240)
        textbox.BorderStyle = BorderStyle.None

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

        rotationAngleLabel = Label()
        rotationAngleLabel.Text = "Rotation Angle:"
        rotationAngleLabel.Font = Font("Helvetica", 10, FontStyle.Regular)
        rotationAngleLabel.Location = Point(30, 297)  # Adjust location as needed
        rotationAngleLabel.Size = Size(120, 20)  # Adjust size as needed
        rotationAngleLabel.ForeColor = Color.FromArgb(240, 240, 240)
        self.Controls.Add(rotationAngleLabel)
        
        # Textbox for Rotation Angle
        rotationAngleTextBox = TextBox()
        rotationAngleTextBox.Name = "rotation_angle"
        rotationAngleTextBox.Text = "90"  # Default value
        rotationAngleTextBox.Location = Point(172, 300)  # Adjust location to align with the label
        rotationAngleTextBox.Size = Size(75, 20)  # Adjust size as needed
        rotationAngleTextBox.BackColor = Color.FromArgb(49, 49, 49)
        rotationAngleTextBox.ForeColor = Color.FromArgb(240, 240, 240)
        rotationAngleTextBox.BorderStyle = BorderStyle.None
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
            radio_button.Text = "{}°".format(angle)
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
                self.toolTip.SetToolTip(self.expand_button, "You wuss!")
        else:
            if self.Width > self.original_width:
                self.Width -= step
            else:
                self.animation_timer.Stop()
                self.Width = self.original_width  # Ensure it ends exactly at target size
                self.expand_button.Image = expand_image  # Change to expand image
                self.toolTip.SetToolTip(self.expand_button, "Advanced features")
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
            # Safely access the text boxes, checking if they actually exist
            hrTextBox = self.Controls.Find("horizontal_resolution", True)
            fovTextBox = self.Controls.Find("fov_angle", True)
            maxDistanceTextBox = self.Controls.Find("max_distance", True)

            if hrTextBox and fovTextBox and maxDistanceTextBox:
                try:
                    hr = int(hrTextBox[0].Text)
                    fov = int(fovTextBox[0].Text)
                    distances = calculator_1(hr, fov)
                    doriIndex = self.doriOptions.index(sender.Tag)
                    calculatedValue = distances[doriIndex]
                    
                    # Update the max_distance textbox with the calculated value
                    maxDistanceTextBox[0].Text = str(calculatedValue)
                except ValueError:
                    # Handle cases where conversion to int fails
                    MessageBox.Show("Please enter valid numeric values for horizontal resolution and FOV angle.")
            else:
                MessageBox.Show("One or more required fields are missing.")

    def on_angle_changed(self, sender, event):
        if sender.Checked:
            self.additional_rotation_angle = sender.Tag  # Update the additional rotation angle based on the selected radio button

    def select_camera(self, sender, event):
        self.WindowState = FormWindowState.Minimized
        if self.radio_current_project.Checked:
            self.select_cameras_current_project()
        elif self.radio_linked_file.Checked:
            self.select_cameras_linked_file()
        self.WindowState = FormWindowState.Normal
        self.Activate()
    def select_cameras_current_project(self):
        # Reset the selected cameras list
        self.selected_cameras = []
        try:
            refs = uidoc.Selection.PickObjects(ObjectType.Element, "Please select cameras.")
            for ref in refs:
                self.selected_cameras.append(doc.GetElement(ref.ElementId))
            self.Activate()  # Bring the window back into focus
        except:
            pass

    def select_cameras_linked_file(self):
        try:
            selectedObjs = uidoc.Selection.PickObjects(ObjectType.LinkedElement, "Select Linked Elements")
            self.selected_cameras = []  # Reset the list

            for selectedObj in selectedObjs:
                linkInstance = doc.GetElement(selectedObj.ElementId)  # Getting the RevitLinkInstance
                linkedDoc = linkInstance.GetLinkDocument()  # Accessing the linked document
                linkedCameraElement = linkedDoc.GetElement(selectedObj.LinkedElementId)  # Getting the element

                transform = linkInstance.GetTransform()
                transformedPosition = transform.OfPoint(linkedCameraElement.Location.Point)
                self.selected_cameras.append((transformedPosition, True))  # Store the transformed position

            MessageBox.Show("Cameras from linked file selected.")
            self.Activate()
        except Exception as e:
            MessageBox.Show("An error occurred during camera selection: " + str(e))
            self.Activate()

    def run_script(self, sender, event):

        try:
            fov_angle = float(Decimal(self.Controls["fov_angle"].Text))
            rotation_angle = float(Decimal(self.Controls["rotation_angle"].Text if self.Controls["rotation_angle"].Text != "" else "0"))
            rotation_angle += self.additional_rotation_angle
            max_distance_m = float(Decimal(self.Controls["max_distance"].Text))
            max_distance_mm = max_distance_m * 1000  # Convert from meters to millimeters

            selected_filled_region_name = self.filledRegionTypeComboBox.SelectedItem.ToString()
            selected_filled_region_id = None
            for name, id in list_filled_region_type_names_and_ids(doc):
                if name == selected_filled_region_name:
                    selected_filled_region_id = id
                    break

            if selected_filled_region_id is not None:
                detail_lines = get_custom_detail_lines(doc, "Boundary")
                if self.radio_current_project.Checked and self.selected_cameras:
                    for camera in self.selected_cameras:
                        # Ensure this call includes all six arguments
                        main_script((camera, False), fov_angle, max_distance_mm, rotation_angle, detail_lines, selected_filled_region_id)
                elif self.radio_linked_file.Checked and self.selected_cameras:
                    for camera_info in self.selected_cameras:
                        # Ensure this call includes all six arguments
                        main_script(camera_info, fov_angle, max_distance_mm, rotation_angle, detail_lines, selected_filled_region_id)
            else:
                MessageBox.Show("Selected filled region type not found.")
        except Exception as e:
            MessageBox.Show("Invalid input: " + str(e))
        
if __name__ == "__main__":
    Application.EnableVisualStyles()
    form = CameraFOVApp()
    Application.Run(form)