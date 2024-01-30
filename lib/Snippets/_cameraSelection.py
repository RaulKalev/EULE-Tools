from Autodesk.Revit.DB import RevitLinkInstance
from Autodesk.Revit.UI.Selection import ObjectType
from System.Windows.Forms import Application, Form, Button, Label, TextBox, DialogResult, MessageBox, FormBorderStyle, RadioButton, ListBox
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
def select_cameras(uidoc):
    global selected_cameras
    selected_cameras = []
    try:
        refs = uidoc.Selection.PickObjects(ObjectType.Element, "Please select cameras.")
        for ref in refs:
            selected_cameras.append(doc.GetElement(ref.ElementId))
    except:
        pass

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
    if self.selected_link is None:
        MessageBox.Show("No link selected. Please select a link first.")
        return

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
