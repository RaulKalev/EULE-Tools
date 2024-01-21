# -*- coding: utf-8 -*-
import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from Autodesk.Revit.UI import TaskDialog
from System.Drawing import Size, Point, Font, FontStyle
from System.Windows.Forms import Application, Button, Form, FormWindowState, ComboBox
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.DB import Line, Transaction, XYZ, CategoryType, BuiltInCategory, Category, CategoryType
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, CategoryType, Transaction, ElementId, GraphicsStyle


def get_line_styles(doc):
    line_styles = {}

    # Fetching all elements of type GraphicsStyle
    collector = FilteredElementCollector(doc).OfClass(GraphicsStyle)

    for elem in collector:
        # Filtering for elements that are line styles
        if elem.GraphicsStyleCategory is not None and elem.GraphicsStyleCategory.CategoryType == CategoryType.Annotation:
            line_styles[elem.Name] = elem.Id

    return line_styles

class SimpleForm(Form):
    def __init__(self, uidoc):
        self.uidoc = uidoc
        self.Text = 'Draw Lines'
        self.Size = Size(350, 90)
        self.midpoint_location = None

        # Custom font and size for buttons
        #buttonFont = Font("Arial", 12, FontStyle.Bold)
        buttonSize = Size(30, 30)
        yCord = 10

        # Variables to store element locations
        self.element1_location = None
        self.element2_location = None

        # Button for selecting Element 1
        self.button1 = Button()
        self.button1.Text = '1.'
        #self.button1.Font = buttonFont
        self.button1.Size = buttonSize
        self.button1.Location = Point(10, yCord)
        self.button1.Click += self.on_button1_click
        self.Controls.Add(self.button1)

        # Button for selecting Element 2
        self.button2 = Button()
        self.button2.Text = '2.'
        #self.button2.Font = buttonFont
        self.button2.Size = buttonSize
        self.button2.Location = Point(45, yCord)
        self.button2.Click += self.on_button2_click
        self.Controls.Add(self.button2)

        '''# Button to display stored values
        self.button3 = Button()
        self.button3.Text = 'Show Stored Values'
        #self.button3.Font = buttonFont
        self.button3.Size = buttonSize
        self.button3.Location = Point(10, 180)
        self.button3.Click += self.on_button3_click
        self.Controls.Add(self.button3)
        '''
        
       # Add this new button in your __init__ method
        self.button4 = Button()
        self.button4.Text = 'Draw'
        #self.button4.Font = buttonFont
        self.button4.Size = Size(50, 30)
        self.button4.Location = Point(80, yCord)
        self.button4.Click += self.on_button4_click
        self.Controls.Add(self.button4)

        self.line_styles = get_line_styles(uidoc.Document)
        self.dropdown = ComboBox()
        self.dropdown.Location = Point(150, 15)
        self.dropdown.Size = Size(150, 20)
        self.dropdown.Text = "Select line type"  # Set default text
            # Sort the line style names alphabetically and add them to the dropdown
        for name in sorted(self.line_styles.keys()):
            self.dropdown.Items.Add(name)
        self.Controls.Add(self.dropdown)

    def on_button1_click(self, sender, args):
        self.WindowState = FormWindowState.Minimized
        self.element1_location = self.select_element_and_get_location()
        self.WindowState = FormWindowState.Normal
        self.update_midpoint()

    def on_button2_click(self, sender, args):
        self.WindowState = FormWindowState.Minimized
        self.element2_location = self.select_element_and_get_location()
        self.WindowState = FormWindowState.Normal
        self.update_midpoint()

    def on_button3_click(self, sender, args):
        self.WindowState = FormWindowState.Minimized
        element1_msg = self.format_location(self.element1_location)
        element2_msg = self.format_location(self.element2_location)
        midpoint_msg = self.format_location(self.midpoint_location) if self.midpoint_location else 'Not Set'

        message = 'Element 1 Location: {}\nElement 2 Location: {}\nMidpoint Location: {}'.format(
            element1_msg, element2_msg, midpoint_msg)
        TaskDialog.Show('Stored Values', message)
        self.WindowState = FormWindowState.Normal

    def select_element_and_get_location(self):
        try:
            selected_ref = self.uidoc.Selection.PickObject(ObjectType.Element)
            if selected_ref is not None:
                element = self.uidoc.Document.GetElement(selected_ref)
                if hasattr(element, 'Location') and element.Location.Point:
                    location = element.Location.Point
                    location_point = XYZ(location.X, location.Y, location.Z)
                    return location_point
                else:
                    return None
        except Exception as e:
            return None

    def format_location(self, location):
        if location:
            return 'X: {:.2f}, Y: {:.2f}, Z: {:.2f}'.format(location.X, location.Y, location.Z)
        else:
            return 'Not Set'
        
    def update_midpoint(self):
        if self.element1_location and self.element2_location:
            mid_x = (self.element1_location.X + self.element2_location.X) / 2.0
            mid_y = (self.element1_location.Y + self.element2_location.Y) / 2.0
            self.midpoint_location = XYZ(mid_x, mid_y, 0)  # Assuming Z = 0 for simplicity
        
    def on_button4_click(self, sender, args):
        self.draw_detail_line()
    
    def on_button_show_midpoint_click(self, sender, args):
        if self.midpoint_location:
            message = 'Midpoint Location: X: {:.2f}, Y: {:.2f}, Z: {:.2f}'.format(
                self.midpoint_location.X, self.midpoint_location.Y, self.midpoint_location.Z)
            TaskDialog.Show('Midpoint', message)
        else:
            TaskDialog.Show('Midpoint', 'Midpoint is not set yet.')
    
    def draw_detail_line(self):
        selected_style_name = self.dropdown.SelectedItem
        selected_style_id = self.line_styles.get(selected_style_name)

        if selected_style_id and self.element1_location and self.element2_location:
            self.update_midpoint()
            midpoint = self.midpoint_location

            if midpoint:
                with Transaction(self.uidoc.Document, 'Draw Detail Line') as trans:
                    trans.Start()
                    try:
                        # Define the points for line segments
                        point1 = XYZ(self.element1_location.X, self.element1_location.Y, 0)
                        interm_point1 = XYZ(midpoint.X, self.element1_location.Y, 0)
                        interm_point2 = XYZ(midpoint.X, self.element2_location.Y, 0)
                        point2 = XYZ(self.element2_location.X, self.element2_location.Y, 0)

                        # Create line geometries
                        line1_geometry = Line.CreateBound(point1, interm_point1)
                        line2_geometry = Line.CreateBound(interm_point1, interm_point2)
                        line3_geometry = Line.CreateBound(interm_point2, point2)

                        # Set line style for each detail curve and create them
                        for line_geom in [line1_geometry, line2_geometry, line3_geometry]:
                            detail_line = self.uidoc.Document.Create.NewDetailCurve(self.uidoc.ActiveView, line_geom)
                            if hasattr(detail_line, 'LineStyle'):
                                detail_line.LineStyle = self.uidoc.Document.GetElement(selected_style_id)

                        trans.Commit()
                    except Exception as e:
                        TaskDialog.Show('Error', 'Failed to draw line: ' + str(e))
                        trans.RollBack()
            else:
                TaskDialog.Show('Error', 'Midpoint not set.')
        else:
            TaskDialog.Show('Error', 'Please select both elements and a line style.')

# To run this form, you need to pass the UIDocument from a RevitPythonShell script
uidoc = __revit__.ActiveUIDocument
form = SimpleForm(uidoc)
Application.Run(form)
