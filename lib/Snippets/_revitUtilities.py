import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import (XYZ, Line, Transaction, FilledRegion, FilledRegionType, CurveLoop, BuiltInParameter,
                               ElementId, ViewPlan, FilteredElementCollector, CurveElement)
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from System import Exception
from math import radians, sin, cos
from decimal import Decimal
from Snippets._fovCalculations import calculate_fov_endpoints
from Snippets._intersections import find_closest_intersection

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
def list_filled_region_type_names_and_ids(doc):
    f_region_types = FilteredElementCollector(doc).OfClass(FilledRegionType)
    # Adjusted to ensure proper access to the Name property
    names_ids = [(fregion_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString(), fregion_type.Id) for fregion_type in f_region_types]
    return names_ids


def get_custom_detail_lines(doc, line_type_name):
    custom_lines = []
    collector = FilteredElementCollector(doc).OfClass(CurveElement).WhereElementIsNotElementType()
    for line in collector:
        if line.LineStyle.Name == line_type_name and hasattr(line, 'GeometryCurve'):
            custom_lines.append(line.GeometryCurve)
    return custom_lines

def draw_line(doc, start_point, end_point):
    with Transaction(doc, "Draw Line") as transaction:
        try:
            transaction.Start()
            line = Line.CreateBound(start_point, end_point)
            doc.Create.NewDetailCurve(doc.ActiveView, line)
            transaction.Commit()
        except Exception as e:
            transaction.RollBack()
            raise e

def select_cameras(uidoc):
    global selected_cameras
    selected_cameras = []
    try:
        refs = uidoc.Selection.PickObjects(ObjectType.Element, "Please select cameras.")
        for ref in refs:
            selected_cameras.append(doc.GetElement(ref.ElementId))
    except:
        pass

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
