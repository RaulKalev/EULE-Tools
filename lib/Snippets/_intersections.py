# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import Line, XYZ, ReferenceIntersector, FindReferenceTarget, ElementCategoryFilter, BuiltInCategory

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
