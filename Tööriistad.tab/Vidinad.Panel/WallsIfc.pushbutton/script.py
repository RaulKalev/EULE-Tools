import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from Autodesk.Revit.DB import (FilteredElementCollector, BuiltInCategory, Outline, BoundingBoxIntersectsFilter, Transaction, 
                               Level, BuiltInParameter, RevitLinkInstance, Transform, ElementIntersectsElementFilter, FamilySymbol, 
                               FamilyInstance, XYZ, Family, ElementId, Structure, Line, ElementTransformUtils, RevitLinkType)
from Autodesk.Revit.UI import TaskDialog
from System.Windows.Forms import Application, Form, Button, ComboBox, Label, FormStartPosition, TextBox, RadioButton, GroupBox, FormBorderStyle, FlatStyle
from System.Drawing import Size, Point, Color, Bitmap, Font, FontStyle
from Snippets._titlebar import TitleBar
from Snippets._imagePath import ImagePathHelper

image_helper = ImagePathHelper()
windowWidth = 400
windowHeight = 400
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

def get_linked_document(doc):
    link_instances = FilteredElementCollector(doc).OfClass(RevitLinkInstance).ToElements()
    for link_instance in link_instances:
        link_doc = link_instance.GetLinkDocument()
        if link_doc:
            return link_doc
    return None

def get_linked_elements(link_doc, category):
    collector = FilteredElementCollector(link_doc).OfCategory(category).WhereElementIsNotElementType().ToElements()
    return collector

def find_closest_wall_and_determine_orientation(doc, point):
    # Find all walls in the document
    walls = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElements()
    
    closest_wall = None
    min_distance = float('inf')
    
    # Iterate over all walls to find the closest one
    for wall in walls:
        # Get the wall's location curve
        location_curve = wall.Location.Curve
        closest_point_on_curve = location_curve.GetEndPoint(0)  # Just a placeholder for actual logic
        
        # Calculate distance from the point to this closest point on the wall's curve
        distance = point.DistanceTo(closest_point_on_curve)
        
        if distance < min_distance:
            min_distance = distance
            closest_wall = wall
    
    # Once you have the closest wall, determine its orientation
    if closest_wall:
        wall_bb = closest_wall.get_BoundingBox(None)
        wall_length_x = abs(wall_bb.Max.X - wall_bb.Min.X)
        wall_length_y = abs(wall_bb.Max.Y - wall_bb.Min.Y)
        is_wall_horizontal = wall_length_x > wall_length_y
        return is_wall_horizontal  # True if wall is horizontal, False if vertical
    
    return False  # Default to False if no wall is found

def determine_wall_orientation(doc, wall, wall_doc=None):
    target_doc = wall_doc if wall_doc else doc
    wall_bb = wall.get_BoundingBox(None)
    wall_length_y = abs(wall_bb.Max.X - wall_bb.Min.X)
    wall_length_x = abs(wall_bb.Max.Y - wall_bb.Min.Y)
    return wall_length_x > wall_length_y

def calculate_rotation_angle(doc, wall):
    wall_orientation_horizontal = determine_wall_orientation(doc, wall)
    return 1.57079632679 if wall_orientation_horizontal else 0

def get_transformed_wall_orientation(doc, wall):
    if isinstance(wall, RevitLinkInstance):
        # If the wall is from a linked document, apply the linked document's transformation.
        transform = wall.GetTotalTransform()
        transformed_orientation = transform.OfVector(wall.Orientation)
        return transformed_orientation
    else:
        # For walls within the host document, proceed as usual.
        return determine_wall_orientation(doc, wall)

def is_rotation_necessary(ladder_position, closest_wall, doc):
    # Determine the orientation of the closest wall
    wall_curve = closest_wall.Location.Curve
    start_point = wall_curve.GetEndPoint(0)
    end_point = wall_curve.GetEndPoint(1)
    wall_direction = end_point - start_point
    wall_orientation_horizontal = abs(wall_direction.X) > abs(wall_direction.Y)
    
    # Determine the position of the ladder relative to the closest wall
    # For simplicity, assume ladder_position is a point (XYZ object)
    wall_to_ladder_vector = ladder_position - start_point
    on_wall_side = (wall_to_ladder_vector.DotProduct(wall_direction)) > 0
    
    # Check for opposing walls
    # This is a simplified approach; a more complex method may be needed for accurate detection
    opposing_walls_exist = False  # Placeholder for actual check logic
    
    # Rotation decision logic
    # This is a simplified logic; you might want to refine it based on your project's requirements
    if wall_orientation_horizontal:
        # If the wall is horizontal, decide to rotate based on the ladder's position on the X-axis
        should_rotate = abs(wall_to_ladder_vector.X) > abs(wall_to_ladder_vector.Y)
    else:
        # If the wall is vertical, decide to rotate based on the ladder's position on the Y-axis
        should_rotate = abs(wall_to_ladder_vector.Y) > abs(wall_to_ladder_vector.X)
    
    # Consider opposing walls in the decision (if necessary)
    if opposing_walls_exist:
        should_rotate = not should_rotate
    
    return should_rotate

def find_clashes_and_widths(doc, include_project_walls, include_linked_walls):
    clash_points = []
    ladder_widths = []
    wall_thicknesses = []
    ladder_heights = []
    
    if include_project_walls:
        walls = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElements()
        wall_doc = doc
    elif include_linked_walls:
        link_doc = get_linked_document(doc)
        if not link_doc:
            TaskDialog.Show('Error', 'No linked documents found.')
            return [], [], [], []
        walls = get_linked_elements(link_doc, BuiltInCategory.OST_Walls)
        wall_doc = link_doc
    else:
        return [], [], [], []
    
    for wall in walls:
        wall_bb = wall.get_BoundingBox(None)
        if wall_bb:
            bb_filter = BoundingBoxIntersectsFilter(Outline(wall_bb.Min, wall_bb.Max))
            ladders = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_CableTray).WhereElementIsNotElementType().WherePasses(bb_filter).ToElements()
            for ladder in ladders:
                ladder_bb = ladder.get_BoundingBox(None)
                if ladder_bb:
                    width_param = ladder.LookupParameter("Width")
                    width = width_param.AsDouble() if width_param else 0
                    height_param = ladder.LookupParameter("Height")
                    height = height_param.AsDouble() if height_param else 0
                    
                    wall_thickness = wall_bb.Max.Z - wall_bb.Min.Z

                    if determine_wall_orientation(doc, wall, wall_doc):
                        y_coordinate = (ladder_bb.Min.Y + ladder_bb.Max.Y) / 2
                        x_coordinate = (wall_bb.Min.X + wall_bb.Max.X) / 2
                    else:
                        x_coordinate = (ladder_bb.Min.X + ladder_bb.Max.X) / 2
                        y_coordinate = (wall_bb.Min.Y + wall_bb.Max.Y) / 2

                    clash_point = XYZ(x_coordinate, y_coordinate, (ladder_bb.Min.Z + ladder_bb.Max.Z) / 2)
                    clash_points.append(clash_point)
                    ladder_widths.append(width)
                    wall_thicknesses.append(wall_thickness)
                    ladder_heights.append(height)

    return clash_points, ladder_widths, wall_thicknesses, ladder_heights

def get_generic_model_families(doc):
    families = FilteredElementCollector(doc).OfClass(Family).ToElements()
    return [f.Name for f in families if f.FamilyCategory.Id == ElementId(BuiltInCategory.OST_GenericModel)]

def get_family_symbol_by_name(doc, family_name, type_name):
    collector = FilteredElementCollector(doc).OfClass(FamilySymbol)
    for elem in collector:
        if elem.Family.Name == family_name and elem.Name == type_name:
            return elem
    return None
def split_and_move_cable_trays(doc, clash_points, depth):
    t = Transaction(doc, "Split and Move Cable Trays")
    t.Start()
    try:
        for clash_point in clash_points:
            # Assuming clash_point includes references to the cable tray and wall it clashes with
            cable_tray = clash_point['cable_tray']
            wall = clash_point['wall']
            intersection_point = clash_point['point']  # This needs to be calculated

            # Split cable tray at intersection point
            # Note: Revit API does not directly support splitting elements at an arbitrary point;
            # this is a conceptual placeholder for whatever method you devise to split the tray
            split_result = ElementTransformUtils.SplitElement(doc, cable_tray.Id, intersection_point)
            
            # Move each resulting segment of the cable tray
            for element_id in split_result:
                # Calculate move vector based on wall orientation and desired depth
                move_vector = XYZ(0, 0, -depth)  # Example move; adjust based on actual need
                ElementTransformUtils.MoveElement(doc, element_id, move_vector)
    except Exception as e:
        t.RollBack()
        TaskDialog.Show("Error", str(e))
    else:
        t.Commit()
def place_markers_at_clashes(doc, clash_points, family_name, ladder_widths, extension_width, extension_depth, ladder_heights, extension_height, wall_thicknesses):
    family = None
    for f in FilteredElementCollector(doc).OfClass(Family):
        if f.Name == family_name and f.FamilyCategory.Id == ElementId(BuiltInCategory.OST_GenericModel):
            family = f
            break

    if not family:
        TaskDialog.Show('Error', 'Family not found.')
        return

    family_symbols_collector = FilteredElementCollector(doc).OfClass(FamilySymbol).WhereElementIsElementType()
    family_symbol = next((s for s in family_symbols_collector if s.Family.Id == family.Id), None)

    if not family_symbol:
        TaskDialog.Show('Error', 'No types found for the family.')
        return

    if not family_symbol.IsActive:
        family_symbol.Activate()
        doc.Regenerate()

    t = Transaction(doc, 'Place Family Instances')
    t.Start()
    try:
        for i, clash_point in enumerate(clash_points):
            instance = doc.Create.NewFamilyInstance(clash_point, family_symbol, Structure.StructuralType.NonStructural)
            closest_wall = closest_wall_to_point(doc, clash_point)
            if closest_wall:
                if is_rotation_necessary(clash_point, closest_wall, doc):
                    # Calculate the required rotation angle based on the wall orientation
                    angle = calculate_rotation_angle(doc, closest_wall)  # Your existing function to calculate angle
                
                # Apply rotation to each instance individually
                if angle != 0:
                    axis = Line.CreateBound(clash_point, XYZ(clash_point.X, clash_point.Y, clash_point.Z + 1))
                    ElementTransformUtils.RotateElement(doc, instance.Id, axis, angle)

            # Set parameters as before
            param_width = instance.LookupParameter("Width")
            if param_width and i < len(ladder_widths):
                param_width.Set(ladder_widths[i] + 2 * extension_width)

            param_depth = instance.LookupParameter("Depth")
            if param_depth and i < len(wall_thicknesses):
                param_depth.Set(2 * extension_depth)

            param_height = instance.LookupParameter("Height")
            if param_height and i < len(ladder_heights):
                param_height.Set(ladder_heights[i] + extension_height)

    except Exception as e:
        t.RollBack()
        TaskDialog.Show("Error", str(e))
    else:
        t.Commit()

def closest_wall_to_point(doc, point):
    # Simplified example to find the closest wall to a given point
    walls = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElements()
    closest_wall = None
    min_distance = float('inf')
    for wall in walls:
        location_curve = wall.Location.Curve
        wall_start_point = location_curve.GetEndPoint(0)
        wall_end_point = location_curve.GetEndPoint(1)
        wall_mid_point = wall_start_point + (wall_end_point - wall_start_point) / 2
        distance = wall_mid_point.DistanceTo(point)
        if distance < min_distance:
            min_distance = distance
            closest_wall = wall
    return closest_wall

def transform_bounding_box(bb, transform):
    transformed_min = transform.OfPoint(bb.Min)
    transformed_max = transform.OfPoint(bb.Max)
    return Outline(transformed_min, transformed_max)

# UI Class for Clash Detection and Marker Placement
class ClashDetectionUI(Form):
    def __init__(self, doc):
        self.doc = doc
        self.InitializeComponent()
    
    def InitializeComponent(self):
        self.Text = "Clash Detection and Marker Placement"
        self.Size = Size(400, 400)
        self.StartPosition = FormStartPosition.CenterScreen
        appName = "Ladder intersections"
        color1 = Color.FromArgb(24, 24, 24)
        color2 = Color.FromArgb(31, 31, 31)
        colorText = Color.FromArgb(240,240,240)
        bodyText = Font("Helvetica", 12, FontStyle.Regular)
        boxText = Font("Helvetica", 10, FontStyle.Regular)

        self.FormBorderStyle = FormBorderStyle.None
        self.dragging = False
        self.offset = None
        self.titleBar = TitleBar(self, appName, logo_image, minimize_image, close_image)

        self.Text = appName
        self.Size = Size(windowWidth, windowHeight)
        self.ForeColor = colorText
        self.Controls.Add(self.titleBar)

        self.BackColor = color2
        self.label = Label()
        self.label.Text = "Family Type:"
        self.label.Font = bodyText
        self.label.Location = Point(10, titleBar+35)
        self.label.Size = Size(130, 30)
        self.Controls.Add(self.label)
        
        self.comboBox = ComboBox()
        self.comboBox.Location = Point(180, titleBar+35)
        self.comboBox.Size = Size(190, 30)
        self.comboBox.Font = boxText
        generic_model_families = get_generic_model_families(self.doc)
        for family_name in generic_model_families:
            self.comboBox.Items.Add(family_name)
        if self.comboBox.Items.Count > 0:
            self.comboBox.SelectedIndex = 0
        self.Controls.Add(self.comboBox)
        
        self.extensionGroup = GroupBox()
        self.extensionGroup.ForeColor = colorText
        self.extensionGroup.Location = Point(10, titleBar+75)
        self.extensionGroup.Size = Size(360, 140)

        self.widthExtensionLabel = Label()
        self.widthExtensionLabel.Text = "Width Extension:"
        self.widthExtensionLabel.Font = bodyText
        self.widthExtensionLabel.Location = Point(10, 25)
        self.widthExtensionLabel.Size = Size(200, 30)
        self.widthExtensionTextBox = TextBox()
        self.widthExtensionTextBox.Text = "0"
        self.widthExtensionTextBox.Font = boxText
        self.widthExtensionTextBox.Location = Point(180, 25)
        self.widthExtensionTextBox.Size = Size(50, 30)

        self.depthExtensionLabel = Label()
        self.depthExtensionLabel.Text = "Depth Extension:"
        self.depthExtensionLabel.Font = bodyText
        self.depthExtensionLabel.Location = Point(10, 60)
        self.depthExtensionLabel.Size = Size(200, 30)
        self.depthExtensionTextBox = TextBox()
        self.depthExtensionTextBox.Text = "100"
        self.depthExtensionTextBox.Font = boxText
        self.depthExtensionTextBox.Location = Point(180, 60)
        self.depthExtensionTextBox.Size = Size(50, 30)

        self.heightExtensionLabel = Label()
        self.heightExtensionLabel.Text = "Height Extension:"
        self.heightExtensionLabel.Font = bodyText
        self.heightExtensionLabel.Location = Point(10, 95)
        self.heightExtensionLabel.Size = Size(200, 30)
        self.heightExtensionTextBox = TextBox()
        self.heightExtensionTextBox.Text = "0"
        self.heightExtensionTextBox.Font = boxText
        self.heightExtensionTextBox.Location = Point(180, 95)
        self.heightExtensionTextBox.Size = Size(50, 30)

        self.runButton = Button()
        self.runButton.Text = "Run"
        self.runButton.Font = Font("Helvetica", 15, FontStyle.Regular)
        self.runButton.Location = Point(240, 25)
        self.runButton.Size = Size(105,95)
        self.runButton.BackColor = Color.FromArgb(69,69,69)
        self.runButton.FlatStyle = FlatStyle.Flat
        self.runButton.FlatAppearance.BorderSize = 0
        self.runButton.Click += self.RunButtonClick
                
        # Inside ClashDetectionUI's InitializeComponent method
        self.radioGroup = GroupBox()
        self.radioGroup.Location = Point(10, titleBar+207)
        self.radioGroup.Size = Size(360, 60)

        self.radioWallsInProject = RadioButton()
        self.radioWallsInProject.Text = "Walls in project"
        self.radioWallsInProject.Font = boxText
        self.radioWallsInProject.Location = Point(10, 20)
        self.radioWallsInProject.Size = Size(170, 20)
        self.radioWallsInProject.Checked = True  # Default selection

        self.radioWallsInLink = RadioButton()
        self.radioWallsInLink.Text = "Walls in link"
        self.radioWallsInLink.Font = boxText
        self.radioWallsInLink.Location = Point(200, 20)
        self.radioWallsInLink.Size = Size(150, 20)

        self.extensionGroup.Controls.Add(self.runButton)
        self.extensionGroup.Controls.Add(self.widthExtensionTextBox)
        self.extensionGroup.Controls.Add(self.widthExtensionLabel)
        self.extensionGroup.Controls.Add(self.depthExtensionTextBox)
        self.extensionGroup.Controls.Add(self.depthExtensionLabel)        
        self.extensionGroup.Controls.Add(self.heightExtensionTextBox)
        self.extensionGroup.Controls.Add(self.heightExtensionLabel)

        self.radioGroup.Controls.Add(self.radioWallsInProject)
        self.radioGroup.Controls.Add(self.radioWallsInLink)

        self.Controls.Add(self.extensionGroup)
        self.Controls.Add(self.radioGroup)
    
    def RunButtonClick(self, sender, args):
        include_project_walls = self.radioWallsInProject.Checked
        include_linked_walls = self.radioWallsInLink.Checked

        try:
            extension_width = float(self.widthExtensionTextBox.Text) / 304.8
            extension_depth = float(self.depthExtensionTextBox.Text) / 304.8
            extension_height = float(self.heightExtensionTextBox.Text) / 304.8
        except ValueError:
            TaskDialog.Show("Error", "Invalid extension value.")
            return

        clash_points, ladder_widths, wall_thicknesses, ladder_heights = find_clashes_and_widths(self.doc, include_project_walls, include_linked_walls)
        if not clash_points:
            TaskDialog.Show("Error", "No clashes found.")
            return

        selected_family_name = self.comboBox.SelectedItem.ToString()
        place_markers_at_clashes(self.doc, clash_points, selected_family_name, ladder_widths, extension_width, extension_depth, ladder_heights, extension_height, wall_thicknesses)
        TaskDialog.Show("Success", "Family instances placed at clash locations.")

# Function to show the UI
def ShowUI(doc):
    # Now show the main UI form
    form = ClashDetectionUI(doc)
    form.ShowDialog()

# Main script execution
if __name__ == "__main__":
    doc = __revit__.ActiveUIDocument.Document
    ShowUI(doc)
