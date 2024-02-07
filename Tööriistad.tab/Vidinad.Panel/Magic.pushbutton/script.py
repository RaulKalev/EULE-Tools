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
from System.Windows.Forms import Application, Form, Button, Label, TextBox, Timer, Cursors, DialogResult, MessageBox, FormBorderStyle, BorderStyle, RadioButton, ListBox, GroupBox, ComboBox, ComboBoxStyle, AnchorStyles, Panel, PictureBoxSizeMode, PictureBox, ToolTip
from System.Drawing import Color, Size, Point, Bitmap, Font, FontStyle
from math import radians, sin, cos
from decimal import Decimal
from Snippets._titlebar import TitleBar
from Snippets._imagePath import ImagePathHelper
from Snippets._windowResize import WindowResizer
from Snippets._interactivePictureBox import InteractivePictureBox
from Snippets._intersections import line_intersection,line_segment_intersection,find_closest_intersection,get_intersection_point
from Snippets._fovCalculations import rotate_vector,calculate_fov_endpoints

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
# Instantiate ImagePathHelper
image_helper = ImagePathHelper()
windowWidth = 280
windowHeight = 345
titleBar = 30

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

def list_filled_region_type_names_and_ids(doc):
    f_region_types = FilteredElementCollector(doc).OfClass(FilledRegionType)
    names_ids = [(Element.Name.GetValue(fregion_type), fregion_type.Id) for fregion_type in f_region_types]
    return names_ids

class DetailLineFilter(ISelectionFilter):
    def AllowElement(self, elem):
        return elem.Category.Id.IntegerValue == int(BuiltInCategory.OST_Lines)

def select_cameras(uidoc):
    global selected_cameras
    selected_cameras = []
    try:
        refs = uidoc.Selection.PickObjects(ObjectType.Element, "Please select cameras.")
        for ref in refs:
            selected_cameras.append(doc.GetElement(ref.ElementId))
    except:
        pass
          
def get_custom_detail_lines(doc, line_type_name):
    custom_lines = []
    collector = FilteredElementCollector(doc).OfClass(CurveElement).WhereElementIsNotElementType()
    for line in collector:
        if line.LineStyle.Name == line_type_name:
            if hasattr(line, 'GeometryCurve'):
                custom_lines.append(line.GeometryCurve)
    return custom_lines

def draw_line(doc, start_point, end_point):
    transaction = Transaction(doc, "Draw Line")
    transaction.Start()
    try:
        line = Line.CreateBound(start_point, end_point)
        doc.Create.NewDetailCurve(doc.ActiveView, line)
        transaction.Commit()
    except Exception as e:
        print("Error: " + str(e))
        transaction.RollBack()

def simulate_camera_fov(doc, camera_position, fov_angle, max_distance_mm, detail_lines, activeView, rotation_angle=0, filled_region_type_id=None):
    # Calculate the rotated FOV endpoints
    left_end, right_end = calculate_fov_endpoints(camera_position, fov_angle, max_distance_mm, rotation_angle)

    # Initialize list to store the boundary points for the filled region
    filled_region_points = [camera_position]

    with Transaction(doc, "Create FOV Filled Region") as trans:
        trans.Start()
        try:
            # Iterate through the FOV angles
            for angle in range(0, int(fov_angle), 1):
                # Calculate the direction for the current angle with rotation
                current_angle = radians(-fov_angle/2 + angle + rotation_angle)
                direction = XYZ(sin(current_angle), -cos(current_angle), 0).Normalize()

                # Create a line in the FOV direction with a defined length
                end_point = camera_position + direction.Multiply(max_distance_mm / 304.8)
                fov_line = Line.CreateBound(camera_position, end_point)
                
                # Find the intersection with detail lines or use the endpoint if no intersection
                intersection_point = find_closest_intersection(fov_line, detail_lines)
                final_point = intersection_point if intersection_point is not None else end_point

                # Add the final point to the list for the filled region
                filled_region_points.append(final_point)

            # Create a CurveLoop for the filled region
            curve_loop = CurveLoop()
            for i in range(len(filled_region_points)):
                start_point = filled_region_points[i]
                end_point = filled_region_points[(i + 1) % len(filled_region_points)]
                curve_loop.Append(Line.CreateBound(start_point, end_point))

            if filled_region_type_id is None:
                # Optionally, find a default filled region type ID as a fallback
                filled_region_types = FilteredElementCollector(doc).OfClass(FilledRegionType)
                filled_region_type_id = filled_region_types.FirstElementId()
            
            # Create the filled region using the specified filled region type ID
            FilledRegion.Create(doc, filled_region_type_id, activeView.Id, [curve_loop])
        except Exception as e:
            print("Error: " + str(e))
            trans.RollBack()
        else:
            trans.Commit()

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
        self.expanded_width = 480  # Target expanded width
        self.original_width = 280  # Original width
        self.is_expanding = False  # Track direction of animation

        self.Controls.Add(self.titleBar)
        #RadioButton Label
        self.radioLabel = Label()
        self.radioLabel.Text = "Camera location:"
        self.radioLabel.Font = Font("Helvetica", 10, FontStyle.Regular)
        self.radioLabel.Location = System.Drawing.Point(30,160)
        self.radioLabel.Size = System.Drawing.Size(140,30)

        # RadioButton for selecting cameras in the current project
        self.radio_current_project = RadioButton()
        self.radio_current_project.Text = "Current project"
        self.radio_current_project.Location = System.Drawing.Point(30, 180)
        self.radio_current_project.Size = System.Drawing.Size(140, 30)
        self.radio_current_project.Checked = True  # Default to current project
        self.Controls.Add(self.radio_current_project)

        # RadioButton for selecting cameras from a linked file
        self.radio_linked_file = RadioButton()
        self.radio_linked_file.Text = "Linked file"
        self.radio_linked_file.Location = System.Drawing.Point(30, 205)
        self.Controls.Add(self.radio_linked_file)

        # Button for selecting camera
        self.select_camera_button = PictureBox()
        self.select_camera_button.Location = System.Drawing.Point(170, titleBar+118)
        self.select_camera_button.Size = System.Drawing.Size(80,40)
        self.select_camera_button.SizeMode = PictureBoxSizeMode.StretchImage
        self.select_camera_button.Click += self.select_camera
        self.select_camera_buttonInteractive = InteractivePictureBox(
        self.select_camera_button, 'Select.png', 'SelectHover.png', 'SelectClick.png')
        self.Controls.Add(self.select_camera_button)

        # Labels and TextBoxes for user input with default values
        self.create_label_and_textbox("FOV Angle:", 60, "fov_angle", "55")
        self.create_label_and_textbox("Rotation Angle:", 90, "rotation_angle", "90")
        self.create_label_and_textbox("Max Distance (m):", 120, "max_distance", "25")

        # ComboBox for filled region type selection
        self.filledRegionTypeComboBox = ComboBox()
        self.filledRegionTypeComboBox.DropDownStyle = ComboBoxStyle.DropDownList
        self.filledRegionTypeComboBox.Location = Point(30, 300)
        self.filledRegionTypeComboBox.Width = 220
        filled_region_types = list_filled_region_type_names_and_ids(doc)
        for name, _ in filled_region_types:
            self.filledRegionTypeComboBox.Items.Add(name)
        if filled_region_types:
            self.filledRegionTypeComboBox.SelectedIndex = 0
        self.Controls.Add(self.filledRegionTypeComboBox)

        # GroupBox for Preset Rotation Angles
        self.rotation_angle_group = GroupBox()
        self.rotation_angle_group.Text = "Preset Rotation Angles"
        self.rotation_angle_group.ForeColor = colorText
        self.rotation_angle_group.Location = Point(30, 240)  # Adjust location as needed
        self.rotation_angle_group.Size = Size(220, 50)  # Adjust size as needed

        # Radio Buttons for preset angles
        self.radio_angles = {}
        angles = [0, 90, 180, 270]
        button_width = 50  # Width of each radio button, adjust as needed
        for i, angle in enumerate(angles):
            radio_button = RadioButton()
            radio_button.Text = "{}°".format(angle)
            radio_button.Location = Point(10 + i * (40 + 10), 20)  # Space them horizontally
            radio_button.Size = Size(50, 20)  # Optional: Adjust size as needed
            radio_button.Tag = angle  # Store the angle value for easy access
            radio_button.CheckedChanged += self.on_angle_changed  # Event handler for change
            if angle == 0:
                radio_button.Checked = True  # Set 0° as the default selection
            self.rotation_angle_group.Controls.Add(radio_button)
            self.radio_angles[angle] = radio_button

        self.Controls.Add(self.rotation_angle_group)

        # Run script button
        self.run_button = PictureBox()
        self.run_button.Image = run_image
        self.run_button.Location = System.Drawing.Point(170, titleBar+165)
        self.run_button.Size = System.Drawing.Size(80,40)
        self.run_button.SizeMode = PictureBoxSizeMode.StretchImage
        self.run_button.Click += self.run_script
        self.run_buttonInteractive = InteractivePictureBox(
        self.run_button, 'Run.png', 'RunHover.png', 'RunClick.png')
        self.Controls.Add(self.run_button)
        self.Controls.Add(self.radioLabel)

        # Add to CameraFOVApp.__init__
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

        self.secretText = Label()
        self.secretText.Text = "Congrats! You found the secret menu"
        self.secretText.ForeColor = colorText
        self.secretText.BackColor = color1
        self.secretText.Font = Font("Helvetica", 18, FontStyle.Regular)
        self.secretText.Location = Point(290, 100)
        self.secretText.Size = Size (200,200)
        self.Controls.Add(self.secretText)

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

    def create_label_and_textbox(self, label_text, y, name, default_value="", label_size=(120, 20), textbox_size=(200, 20), label_color=System.Drawing.Color.LightGray, textbox_color=System.Drawing.Color.White, text_color=System.Drawing.Color.Black):
        label = Label()
        label.Text = label_text
        label.Location = System.Drawing.Point(30, y)
        label.Font = Font("Helvetica", 10, FontStyle.Regular)
        label.Size = System.Drawing.Size(120,20)
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

        border = Panel()
        border.Location = System.Drawing.Point(169, (y-3))
        border.Size = System.Drawing.Size(81,20)
        border.BackColor = Color.FromArgb(69,69,69)
        border.Anchor = AnchorStyles.Top | AnchorStyles.Left

        cover = Panel()
        cover.Location = System.Drawing.Point(170, (y-2))
        cover.Size = System.Drawing.Size(79,18)
        cover.BackColor = Color.FromArgb(49, 49, 49)
        cover.Anchor = AnchorStyles.Top | AnchorStyles.Left
      
        self.Controls.Add(textbox)
        self.Controls.Add(label) 
        self.Controls.Add(cover)
        self.Controls.Add(border)

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

    def on_angle_changed(self, sender, event):
        if sender.Checked:
            self.additional_rotation_angle = sender.Tag  # Update the additional rotation angle based on the selected radio button

    def select_camera(self, sender, event):
        if self.radio_current_project.Checked:
            self.select_cameras_current_project()
        elif self.radio_linked_file.Checked:
            self.select_cameras_linked_file()

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