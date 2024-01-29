import clr
import System
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from Autodesk.Revit.DB import (XYZ, Line, Transaction, ViewDetailLevel, 
                               ReferenceIntersector, FindReferenceTarget, FilledRegion, FilledRegionType, CurveLoop,
                               ElementCategoryFilter, BuiltInCategory, ViewFamilyType,
                               FilteredElementCollector, View3D, ElementId, Curve, CurveElement, ViewPlan)
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from System.Windows.Forms import Application, Form, Button, Label, TextBox, DialogResult, MessageBox, FormBorderStyle
from System.Drawing import Color, Size, Point, Bitmap, Font, FontStyle
from math import radians, sin, cos
from decimal import Decimal
from Snippets._titlebar import TitleBar
from Snippets._imagePath import ImagePathHelper
from Snippets._windowResize import WindowResizer

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

# Instantiate ImagePathHelper
image_helper = ImagePathHelper()
windowWidth = 300 
windowHeight = 205
titleBar = 30

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
        
def line_intersection(line1, line2):
    """
    Calculate the intersection point of two infinite lines.
    Returns the intersection point or None if there is no intersection.
    """
    p1, p2 = line1.GetEndPoint(0), line1.GetEndPoint(1)
    p3, p4 = line2.GetEndPoint(0), line2.GetEndPoint(1)

    def det(a, b):
        return a.X * b.Y - a.Y * b.X

    div = det(XYZ(p1.X - p2.X, p1.Y - p2.Y, 0), XYZ(p3.X - p4.X, p3.Y - p4.Y, 0))
    if abs(div) < 1e-9:  # Lines are parallel or coincident
        return None

    d = XYZ(det(p1, p2), det(p3, p4), 0)
    x = det(d, XYZ(p1.X - p2.X, p3.X - p4.X, 0)) / div
    y = det(d, XYZ(p1.Y - p2.Y, p3.Y - p4.Y, 0)) / div

    return XYZ(x, y, p1.Z) 
  
def rotate_vector(vector, angle_degrees, axis=XYZ(0, 0, 1)):
    """
    Rotates a vector around a given axis by a certain angle in degrees.
    """
    angle_radians = radians(angle_degrees)
    cos_angle = cos(angle_radians)
    sin_angle = sin(angle_radians)
    x, y, z = vector.X, vector.Y, vector.Z
    ux, uy, uz = axis.X, axis.Y, axis.Z
    
    # Rotation matrix components
    rot_matrix00 = cos_angle + ux * ux * (1 - cos_angle)
    rot_matrix01 = ux * uy * (1 - cos_angle) - uz * sin_angle
    rot_matrix02 = ux * uz * (1 - cos_angle) + uy * sin_angle
    
    rot_matrix10 = uy * ux * (1 - cos_angle) + uz * sin_angle
    rot_matrix11 = cos_angle + uy * uy * (1 - cos_angle)
    rot_matrix12 = uy * uz * (1 - cos_angle) - ux * sin_angle
    
    rot_matrix20 = uz * ux * (1 - cos_angle) - uy * sin_angle
    rot_matrix21 = uz * uy * (1 - cos_angle) + ux * sin_angle
    rot_matrix22 = cos_angle + uz * uz * (1 - cos_angle)
    
    # Apply rotation
    rotated_vector = XYZ(
        rot_matrix00 * x + rot_matrix01 * y + rot_matrix02 * z,
        rot_matrix10 * x + rot_matrix11 * y + rot_matrix12 * z,
        rot_matrix20 * x + rot_matrix21 * y + rot_matrix22 * z
    )
    
    return rotated_vector

def calculate_fov_endpoints(camera_position, fov_angle, max_distance_mm, rotation_angle):
    """
    Calculate the endpoints of the FOV based on the camera position, FOV angle, distance, and rotation angle.
    """
    distance_feet = max_distance_mm / 304.8
    half_fov = radians(fov_angle / 2)
    base_left_direction = XYZ(-sin(half_fov), -cos(half_fov), 0)
    base_right_direction = XYZ(sin(half_fov), -cos(half_fov), 0)

    # Rotate the direction vectors
    left_direction = rotate_vector(base_left_direction, rotation_angle)
    right_direction = rotate_vector(base_right_direction, rotation_angle)

    left_end = camera_position + left_direction.Multiply(distance_feet)
    right_end = camera_position + right_direction.Multiply(distance_feet)

    return (left_end, right_end)

def get_custom_detail_lines(doc, line_type_name):
    """
    Get all detail lines of a specific type in the project.
    :param doc: Revit document.
    :param line_type_name: Name of the custom detail line type.
    :return: List of detail lines of the specified type.
    """
    custom_lines = []
    collector = FilteredElementCollector(doc).OfClass(CurveElement).WhereElementIsNotElementType()
    for line in collector:
        if line.LineStyle.Name == line_type_name:
            if hasattr(line, 'GeometryCurve'):
                custom_lines.append(line.GeometryCurve)
    return custom_lines

def get_intersection_point(doc, origin, direction, view3D):
    filter = ElementCategoryFilter(BuiltInCategory.OST_Walls)
    intersector = ReferenceIntersector(filter, FindReferenceTarget.Element, view3D)
    intersector.FindReferencesInRevitLinks = True

    hits = intersector.Find(origin, direction)
    if hits:
        closest_hit = hits[0]
        ref = closest_hit.GetReference()

        # Check if the reference is in a linked model
        if ref.GlobalPoint is not None:
            # Transform the intersection point from the linked model's coordinate system to the host model's coordinate system
            linked_instance = doc.GetElement(ref.ElementId)
            if linked_instance is not None:
                transform = linked_instance.GetTransform()
                return transform.OfPoint(ref.GlobalPoint)

        return None
    else:
        return None

def line_segment_intersection(line1, line2, tolerance=0.0001):
    """
    Calculate the intersection point of two line segments.
    Returns the intersection point or None if there is no intersection.
    """
    p1, p2 = line1.GetEndPoint(0), line1.GetEndPoint(1)
    p3, p4 = line2.GetEndPoint(0), line2.GetEndPoint(1)

    def det(a, b):
        return a.X * b.Y - a.Y * b.X

    def on_segment(p, q, r):
        # Check if q lies within 'tolerance' of the line segment pr
        if (min(p.X, r.X) - tolerance <= q.X <= max(p.X, r.X) + tolerance and 
            min(p.Y, r.Y) - tolerance <= q.Y <= max(p.Y, r.Y) + tolerance):
            return True
        return False

    d1 = det(p3 - p1, p2 - p1)
    d2 = det(p4 - p1, p2 - p1)
    d3 = det(p1 - p3, p4 - p3)
    d4 = det(p2 - p3, p4 - p3)

    if d1 * d2 < 0 and d3 * d4 < 0:
        # Lines intersect, but need to check if the intersection point is within the line segments
        intersection = line_intersection(Line.CreateBound(p1, p2), Line.CreateBound(p3, p4))
        if intersection and on_segment(p1, intersection, p2) and on_segment(p3, intersection, p4):
            return intersection

    return None

def find_closest_intersection(fov_line, detail_lines):
    """
    Find the closest intersection point of the fov_line with any of the detail_lines.
    """
    closest_point = None
    min_distance = float('inf')

    for detail_line in detail_lines:
        intersection = line_segment_intersection(fov_line, detail_line)
        if intersection:
            distance = intersection.DistanceTo(fov_line.GetEndPoint(0))
            if distance < min_distance:
                min_distance = distance
                closest_point = intersection

    return closest_point

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

def simulate_camera_fov(doc, camera_position, fov_angle, max_distance_mm, detail_lines, activeView, rotation_angle=0):
    """
    Simulate the field of view for a camera and create a filled region in the active view.
    """
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

            # Find a filled region type (use the first one available)
            filled_region_types = FilteredElementCollector(doc).OfClass(FilledRegionType)
            filled_region_type_id = filled_region_types.FirstElementId()

            # Create the filled region
            FilledRegion.Create(doc, filled_region_type_id, activeView.Id, [curve_loop])
        except Exception as e:
            print("Error: " + str(e))
            trans.RollBack()
        else:
            trans.Commit()

# Main script execution
def main_script(camera, fov_angle, max_distance_mm, rotation_angle, detail_lines):
    if camera is not None:
        camera_position = camera.GetTransform().Origin
        activeView = doc.ActiveView  # Get the active view from the document

        if not isinstance(activeView, ViewPlan):
            print("The active view is not a plan view. Please switch to a plan view to draw detail lines.")
        else:
            # Convert Decimal to float
            fov_angle = float(fov_angle+1)
            max_distance_mm = float(max_distance_mm)
            rotation_angle = float(rotation_angle)

            simulate_camera_fov(doc, camera_position, fov_angle, max_distance_mm, detail_lines, activeView, rotation_angle)
    else:
        print("Camera selection was cancelled or no camera was selected.")

# Initialize and run the application
class CameraFOVApp(Form):
    def __init__(self):
        self.Text = "Camera FOV Configuration"
        self.Width = windowWidth
        self.Height = windowHeight
        appName = "FOV Magic"

        # Set your custom minWidth and minHeight
        #customMinWidth = 300
        #customMinHeight = 205

        self.titleBar = TitleBar(self, appName, logo_image, minimize_image, close_image)

        #self.resizer = WindowResizer(self, customMinWidth, customMinHeight)

        # Set the title and size of the window
        self.FormBorderStyle = FormBorderStyle.None
        self.Text = appName
        self.Size = Size(windowWidth, windowHeight)
        color3 = Color.FromArgb(49, 49, 49)
        color1 = Color.FromArgb(24, 24, 24)
        colorText = Color.FromArgb(240,240,240)
        panelSize = 30
        self.BackColor = color1
        self.ForeColor = colorText  

        self.Controls.Add(self.titleBar)

        # Button for selecting camera
        self.select_camera_button = Button()
        self.select_camera_button.Text = "Select Camera"
        self.select_camera_button.Location = System.Drawing.Point(100, titleBar+125)
        self.select_camera_button.Click += self.select_camera
        self.Controls.Add(self.select_camera_button)

        # Labels and TextBoxes for user input with default values
        self.create_label_and_textbox("FOV Angle:", 60, "fov_angle", "55")
        self.create_label_and_textbox("Rotation Angle:", 90, "rotation_angle", "90")
        self.create_label_and_textbox("Max Distance (m):", 120, "max_distance", "25")

        # Run script button
        self.run_button = Button()
        self.run_button.Text = "Run Script"
        self.run_button.Location = System.Drawing.Point(180, titleBar+125)
        self.run_button.Click += self.run_script
        self.Controls.Add(self.run_button)

        # Handling form movement
        self.dragging = False
        self.offset = None

    def create_label_and_textbox(self, label_text, y, name, default_value=""):
        label = Label()
        label.Text = label_text
        label.Location = System.Drawing.Point(30, y)
        self.Controls.Add(label)

        textbox = TextBox()
        textbox.Location = System.Drawing.Point(150, y)
        textbox.Name = name
        textbox.Text = default_value
        self.Controls.Add(textbox)

    def select_camera(self, sender, event):
        select_cameras(uidoc)
        if not selected_cameras:
            MessageBox.Show("No cameras selected.")
        self.Activate()  # Bring the window back into focus

    def run_script(self, sender, event):
        # Validate input and run script
        try:
            fov_angle = Decimal(self.Controls["fov_angle"].Text)
            rotation_angle_text = self.Controls["rotation_angle"].Text
            rotation_angle = Decimal(rotation_angle_text) if rotation_angle_text != "" else 0
            max_distance_m = Decimal(self.Controls["max_distance"].Text)

            if selected_cameras and fov_angle and max_distance_m and (rotation_angle is not None):
                max_distance_mm = max_distance_m * 1000  # Convert from meters to millimeters
                custom_line_type_name = "Boundary"
                detail_lines = get_custom_detail_lines(doc, custom_line_type_name)
                for camera in selected_cameras:
                    main_script(camera, fov_angle, max_distance_mm, rotation_angle, detail_lines)
            else:
                MessageBox.Show("Please fill all fields and select cameras.")
        except Exception as e:
            MessageBox.Show("Invalid input: " + str(e))
if __name__ == "__main__":
    Application.EnableVisualStyles()
    form = CameraFOVApp()
    Application.Run(form)