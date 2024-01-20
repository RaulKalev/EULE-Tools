import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from System.Windows.Forms import Form, Application, Button, ComboBox
from System.Drawing import Size, Point
from System import Array
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.DB import StorageType, ParameterSet, BuiltInParameterGroup

class SimpleForm(Form):
    def __init__(self, uidoc):
        self.uidoc = uidoc
        self.InitializeComponent()
        self.TopMost = True  # This will keep the form always on top
        self.parameters_in_memory = {}

    def InitializeComponent(self):
        self.Text = 'Select Element and Get Parameters'
        self.Size = Size(155, 225)  # Adjust the size as needed
        startCoordinate = 10
        midCoordinate = 20
        boxHeight = 45
        boxLength = 100

        # Create and configure the dropdown list (ComboBox)
        self.comboBox = ComboBox()
        paramTypes = Array[object](["Type Parameters", "Instance Parameters", "Shared Parameters"])
        self.comboBox.Items.AddRange(paramTypes)
        self.comboBox.SelectedIndex = 0
        self.comboBox.Location = Point(startCoordinate, 10)

        # Create and configure the button
        self.button = Button()
        self.button.Text = 'Select Element'
        self.button.Location = Point(midCoordinate, 40)
        self.button.Size = Size(boxLength, boxHeight)
        self.button.Click += self.OnButtonClick

        # Create and configure the button for displaying stored parameters
        self.showParamsButton = Button()
        self.showParamsButton.Text = 'Show Params'
        self.showParamsButton.Location = Point(midCoordinate, 85)  # Adjust location as needed
        self.showParamsButton.Size = Size(boxLength, boxHeight)
        self.showParamsButton.Click += self.OnShowParamsButtonClick

        # Create and configure the button for clearing stored parameters
        self.clearMemoryButton = Button()
        self.clearMemoryButton.Text = 'Clear'
        self.clearMemoryButton.Location = Point(midCoordinate, 130)  # Adjust location as needed
        self.clearMemoryButton.Size = Size(boxLength, boxHeight)
        self.clearMemoryButton.Click += self.OnClearMemoryButtonClick

        # Add the dropdown list and the button to the form
        self.Controls.Add(self.comboBox)
        self.Controls.Add(self.button)
        self.Controls.Add(self.showParamsButton)
        self.Controls.Add(self.clearMemoryButton)


    def OnButtonClick(self, sender, args):
        try:
            ref = self.uidoc.Selection.PickObject(ObjectType.Element, "Please select an element")
            element = self.uidoc.Document.GetElement(ref.ElementId)  # Define 'element' here
            elementType = self.uidoc.Document.GetElement(element.GetTypeId())  # Define 'elementType' here


            selectedParamType = self.comboBox.SelectedItem
            paramInfo = []

            if selectedParamType == "Instance Parameters":
                paramInfo = self.collectParameters(element.Parameters, "Instance")
            elif selectedParamType == "Type Parameters":
                if elementType:  # Ensure elementType is not None
                    # Exclude shared parameters from type parameter list
                    paramInfo = self.collectParameters(elementType.Parameters, "Type", excludeShared=True)
                else:
                    paramInfo = ["No type parameters found."]
            elif selectedParamType == "Shared Parameters":
                # Include shared parameters from both element and elementType
                paramInfo = self.collectSharedParameters(element.Parameters, element)
                if elementType:
                    paramInfo += self.collectSharedParameters(elementType.Parameters, elementType)

            TaskDialog.Show("Element Parameters", "\n".join(paramInfo))
        except Exception as e:
            TaskDialog.Show("Error", str(e))

    def collectParameters(self, parameters, paramType, excludeShared=False):
        paramInfo = []
        for p in parameters:
            if excludeShared and p.IsShared:
                continue  # Skip shared parameters if excludeShared is True
            
            paramName = p.Definition.Name
            paramValue = self.getParameterValue(p)
            paramInfo.append("{}: {}".format(paramName, paramValue))
            
            # Store the parameter with type information
            self.parameters_in_memory[(paramName, paramType)] = paramValue

        return paramInfo

    def collectSharedParameters(self, parameters, paramType):
        sharedParamInfo = []
        for p in parameters:
            if p.IsShared:
                paramName = p.Definition.Name
                paramValue = self.getParameterValue(p)
                sharedParamInfo.append("{}: {}".format(paramName, paramValue))
                
                # Store the parameter with type information
                self.parameters_in_memory[(paramName, paramType)] = paramValue

        return sharedParamInfo

    def getParameterValue(self, param):
        if param.StorageType == StorageType.Integer:
            return param.AsInteger()
        elif param.StorageType == StorageType.Double:
            return "{:.3f}".format(param.AsDouble())
        elif param.StorageType == StorageType.String:
            return param.AsString()
        elif param.StorageType == StorageType.ElementId:
            id = param.AsElementId()
            if id.IntegerValue >= 0:
                relatedElement = self.uidoc.Document.GetElement(id)
                return relatedElement.Name if relatedElement and hasattr(relatedElement, 'Name') else "Related Element"
            else:
                return "None"
        else:
            return "Not Supported"
        
    def OnShowParamsButtonClick(self, sender, e):
        if self.parameters_in_memory:
            paramInfo = ["{} ({}): {}".format(name, typ, value) 
                         for (name, typ), value in self.parameters_in_memory.items()]
            message = "\n".join(paramInfo)
            TaskDialog.Show("Stored Parameters", message)
        else:
            TaskDialog.Show("Info", "No parameters stored in memory.")

    def OnClearMemoryButtonClick(self, sender, e):
        # Function to execute when the clear memory button is clicked
        self.parameters_in_memory.clear()
        TaskDialog.Show("Info", "Memory cleared.")

# To run this form, you need to pass the UIDocument from a RevitPythonShell script
uidoc = __revit__.ActiveUIDocument
form = SimpleForm(uidoc)
Application.Run(form)
