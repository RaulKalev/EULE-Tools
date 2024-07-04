from Autodesk.Revit.DB import RevitLinkInstance, BuiltInParameter
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.UI import TaskDialog
from System.Collections.Generic import List
from System.Windows.Forms import Application, Form, Button, Label, TextBox, DialogResult, MessageBox, FormBorderStyle, RadioButton, ListBox

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

# Function to select cameras within the current project and display their rotation values
def select_cameras_current_project():
    selected_cameras = []  # Reset the selected cameras list
    rotation_values = []  # List to store rotation values
    try:
        refs = uidoc.Selection.PickObjects(ObjectType.Element, "Please select cameras.")
        for ref in refs:
            camera = doc.GetElement(ref.ElementId)
            selected_cameras.append(camera)
            # Retrieve the "Camera Rotation" parameter value
            rotation_param = camera.LookupParameter("Camera Rotation")  # Ensure the parameter name matches
            if rotation_param:
                rotation_values.append(rotation_param.AsDouble())  # Assuming rotation is stored as a double
            else:
                rotation_values.append("N/A")  # Parameter not found

        # Display the rotation values in a message box
        rotation_message = "\n".join("Camera ID {}: Rotation = {}".format(cam.Id, rot) for cam, rot in zip(selected_cameras, rotation_values))
        MessageBox.Show(rotation_message, "Camera Rotations")

    except Exception as e:
        MessageBox.Show("An error occurred during camera selection: " + str(e))

# Function to select cameras from a linked file and display their rotation values
def select_cameras_linked_file(selected_link):
    selected_cameras = []  # Reset the list
    rotation_values = []  # List to store rotation values

    try:
        selectedObjs = uidoc.Selection.PickObjects(ObjectType.LinkedElement, "Select Linked Elements")
        for selectedObj in selectedObjs:
            linkInstance = doc.GetElement(selectedObj.ElementId)  # Getting the RevitLinkInstance
            if linkInstance and isinstance(linkInstance, RevitLinkInstance) and linkInstance.GetLinkDocument():
                linkedDoc = linkInstance.GetLinkDocument()  # Accessing the linked document
                linkedCameraElement = linkedDoc.GetElement(selectedObj.LinkedElementId)  # Getting the camera element

                # Retrieve the "Camera Rotation" parameter value
                rotation_param = linkedCameraElement.LookupParameter("Camera Rotation")  # Ensure the parameter name matches
                if rotation_param:
                    rotation_values.append(rotation_param.AsDouble())  # Assuming rotation is stored as a double
                else:
                    rotation_values.append("N/A")  # Parameter not found

                # Store the camera info with placeholder for transformed position
                selected_cameras.append((linkedCameraElement, "Transformed position placeholder"))

        # Display the rotation values in a message box for linked cameras
        rotation_message = "\n".join("Camera ID {}: Rotation = {}".format(cam.Id, rot) for cam, rot in zip(selected_cameras, rotation_values))
        MessageBox.Show(rotation_message, "Linked Camera Rotations")

    except Exception as e:
        MessageBox.Show("An error occurred during camera selection: " + str(e))

        self.Activate()
