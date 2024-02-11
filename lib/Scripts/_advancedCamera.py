# -*- coding: utf-8 -*-
from System.Drawing import Color, Font, FontStyle, Point, Size
from System.Windows.Forms import Form, Label, Panel, TextBox, BorderStyle, Button
import math

def calculator_1(hr, fov):
    # Conversion values
    d = 7.62
    o = 19.2024
    r = 38.1
    i = 76.2
    # Calculate distances for DORI
    distances = [0] * 4
    for index, ppm in enumerate([d, o, r, i], start=0):
        A = hr / ppm
        B = 360 / float(fov)
        C = A * B
        D = 2 * math.pi
        E = C / D
        distances[index] = round(E * 0.3048, 1)  # Convert to meters and round
    return distances
