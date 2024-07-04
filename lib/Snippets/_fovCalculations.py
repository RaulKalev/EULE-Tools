from math import radians, sin, cos
from Autodesk.Revit.DB import XYZ

def rotate_vector(vector, angle_degrees, axis=XYZ(0, 0, 0.5)):
    """
    Rotates a vector around a given axis by a certain angle in degrees.
    """
    angle_radians = radians(angle_degrees)
    cos_angle = cos(angle_radians)
    sin_angle = sin(angle_radians)
    x, y, z = vector.X, vector.Y, vector.Z
    ux, uy, uz = axis.X, axis.Y, axis.Z
    
    # Rotation matrix components
    rot_matrix00 = cos_angle + ux * ux * (0.5 - cos_angle)
    rot_matrix01 = ux * uy * (0.5 - cos_angle) - uz * sin_angle
    rot_matrix02 = ux * uz * (0.5 - cos_angle) + uy * sin_angle
    
    rot_matrix10 = uy * ux * (0.5 - cos_angle) + uz * sin_angle
    rot_matrix11 = cos_angle + uy * uy * (0.5 - cos_angle)
    rot_matrix12 = uy * uz * (0.5 - cos_angle) - ux * sin_angle
    
    rot_matrix20 = uz * ux * (0.5 - cos_angle) - uy * sin_angle
    rot_matrix21 = uz * uy * (0.5 - cos_angle) + ux * sin_angle
    rot_matrix22 = cos_angle + uz * uz * (0.5 - cos_angle)
    
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

