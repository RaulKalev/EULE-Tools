# -*- coding: utf-8 -*-
import clr
import System
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from System.Windows.Forms import Form, Application, Button, Label, CheckBox, ListView, View, ListViewItem, HorizontalAlignment, TextBox, FlatStyle
from System.Drawing import Size, Point, Font, FontStyle, Color, Icon
from System.Collections import IComparer
from System import Array
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.DB import StorageType, UnitUtils, ParameterType, UnitTypeId, LabelUtils, BuiltInParameterGroup, Transaction

class ListViewColumnSorter(IComparer):
    def __init__(self):
        self.columnToSort = 0

    def Compare(self, x, y):
        textX = x.SubItems[self.columnToSort].Text
        textY = y.SubItems[self.columnToSort].Text
        return cmp(textX.lower(), textY.lower())  # Compare as lowercase for case-insensitive sorting

class SimpleForm(Form):
    def __init__(self, uidoc):
        self.uidoc = uidoc        
        self.parameters_in_memory = {}
        self.Icon = Icon("Tööriistad.tab/Parameetrid.Panel/Parameter Copy.pushbutton/icon.ico")
        self.listViewColumnSorter = ListViewColumnSorter()
        self.fullListViewItems = []  # List to store all items
        self.InitializeComponent()
        #self.BackColor = Color.FromArgb(24, 24, 24)
        self.TopMost = True  # This will keep the form always on top
        
    def InitializeComponent(self):
        self.Text = 'Copy parameter values from element to element'
        self.Size = Size(635, 515)  # Adjust the size as needed
        startCoordinate = 10
        midCoordinate = 412.5
        cord1 = 10
        cord2 = 70
        boxHeight = 35
        boxLength = 64
        color1 = Color.FromArgb(66, 66, 66)
        color2 = Color.FromArgb(245, 245, 245)
        #color3 = Color.FromArgb(24, 24, 24)
                       
        # Create and configure the label for displaying the selected element's info
        self.selectedElementNameLabel = Label()
        self.selectedElementNameLabel.Location = Point(10, 15)
        self.selectedElementNameLabel.Size = Size(400, 20)
        self.selectedElementNameLabel.Text = "Name: None"
        self.selectedElementNameLabel.Font = Font("Arial", 12, FontStyle.Bold)
        #self.selectedElementNameLabel.ForeColor = color2

        self.selectedElementTypeLabel = Label()
        self.selectedElementTypeLabel.Location = Point(10, 35)
        self.selectedElementTypeLabel.Size = Size(400, 20)
        self.selectedElementTypeLabel.Text = "Type: None"
        self.selectedElementTypeLabel.Font = Font("Arial", 12, FontStyle.Bold)
        #self.selectedElementTypeLabel.ForeColor = color2
    
        # Create and configure the button
        self.button = Button()
        self.button.Text = 'Select Element'
        self.button.Location = Point(midCoordinate, startCoordinate)
        self.button.Size = Size(boxLength, boxHeight)
        self.button.FlatStyle = FlatStyle.Flat
        self.button.FlatAppearance.BorderSize = 0
        self.button.BackColor = color1
        self.button.ForeColor = color2
        self.button.Click += self.OnButtonClick

        # Add the new "Add to Element" button
        self.addParametersButton = Button()
        self.addParametersButton.Text = 'Add to Element'
        self.addParametersButton.Location = Point(midCoordinate + 65.5, startCoordinate)  # Adjust location as needed
        self.addParametersButton.Size = Size(boxLength, boxHeight)
        self.addParametersButton.FlatStyle = FlatStyle.Flat
        self.addParametersButton.FlatAppearance.BorderSize = 0
        self.addParametersButton.BackColor = color1
        self.addParametersButton.ForeColor = color2
        self.addParametersButton.Click += self.OnAddParametersButtonClick

        # Create and configure the button for clearing stored parameters
        self.clearMemoryButton = Button()
        self.clearMemoryButton.Text = 'Clear'
        self.clearMemoryButton.Location = Point(midCoordinate+131.5, startCoordinate)  # Adjust location as needed
        self.clearMemoryButton.Size = Size(boxLength, boxHeight)
        self.clearMemoryButton.FlatStyle = FlatStyle.Flat
        self.clearMemoryButton.FlatAppearance.BorderSize = 0
        self.clearMemoryButton.BackColor = color1
        self.clearMemoryButton.ForeColor = color2
        self.clearMemoryButton.Click += self.OnClearMemoryButtonClick

        # Create and configure the "Select All/None" checkbox
        self.selectAllCheckBox = CheckBox()        
        self.selectAllCheckBox.Location = Point(cord1+6, cord2+7)  # Adjust location as needed
        self.selectAllCheckBox.Size = Size(13, 13)
        self.selectAllCheckBox.CheckedChanged += self.OnSelectAllCheckBoxChanged
        
        # Configure the ListView
        self.listView = ListView()
        self.listView.View = View.Details
        self.listView.CheckBoxes = True
        self.listView.FullRowSelect = True
        self.listView.GridLines = True
        self.listView.Location = Point(cord1, cord2)
        self.listView.Size = Size(600, 400)
        #self.listView.BackColor = color1
        #self.listView.ForeColor = color2

        # Adding columns to the ListView
        self.listView.Columns.Add("     Name", 210, HorizontalAlignment.Left)
        self.listView.Columns.Add("Value", 155, HorizontalAlignment.Left)
        self.listView.Columns.Add("Group", 100, HorizontalAlignment.Left)
        self.listView.Columns.Add("Type", 60, HorizontalAlignment.Left)
        self.listView.Columns.Add("Shared", 50, HorizontalAlignment.Center)

        # Create and configure the search box (TextBox)
        self.searchBox = TextBox()
        self.searchBox.Location = Point(cord1+400, cord2-24)
        self.searchBox.Size = Size(200, 20)
        self.searchBox.Text = "Search"
        #self.searchBox.BackColor = color1
        #self.searchBox.ForeColor = color2
        self.searchBox.TextChanged += self.OnSearchBoxTextChanged
        self.searchBox.Enter += self.OnSearchBoxEnter
        
        # Add the dropdown list and the button to the form
        self.Controls.Add(self.selectAllCheckBox)
        self.Controls.Add(self.selectedElementNameLabel)
        self.Controls.Add(self.selectedElementTypeLabel)
        self.Controls.Add(self.button)
        self.Controls.Add(self.clearMemoryButton)      
        self.Controls.Add(self.listView)        
        self.Controls.Add(self.searchBox)
        self.Controls.Add(self.addParametersButton)

        self.listView.ListViewItemSorter = self.listViewColumnSorter
        self.listView.ColumnClick += self.OnColumnHeaderClick
    
    def getParameterCategory(self, paramName, paramType, element, elementType):
        try:
            if paramType == "Instance":
                param = element.LookupParameter(paramName)
            elif paramType == "Type":
                param = elementType.LookupParameter(paramName)
            else:
                return "Unknown Type"

            if param.Definition is None:
                return "No Definition"

            paramGroup = param.Definition.ParameterGroup
            if paramGroup == BuiltInParameterGroup.INVALID:
                return "Invalid Group"

            return LabelUtils.GetLabelFor(paramGroup) or "No Label"
        except Exception as e:
            return "Error: " + str(e)

    def OnButtonClick(self, sender, args):
        try:
            # Disable the button to prevent further clicks until selection is complete
            self.button.Enabled = False

            ref = self.uidoc.Selection.PickObject(ObjectType.Element, "Please select an element")
            element = self.uidoc.Document.GetElement(ref.ElementId)
            self.listView.Items.Clear()

            # Initialize variables for family name and type
            familyName = "Unknown"
            typeName = "Unknown"

            # Fetching the ElementId of the ElementType
            typeId = element.GetTypeId()
            if typeId.IntegerValue >= 0:  # Checking if the typeId is valid
                elementType = self.uidoc.Document.GetElement(typeId)

                # Retrieve the family name and type name
                if elementType:
                    typeName = element.Name if hasattr(element, 'Name') else "Unknown"
                    if hasattr(elementType, 'Family'):
                        family = elementType.Family
                        familyName = family.Name if hasattr(family, 'Name') else "Unknown"

            # Update label with family name and type
            self.selectedElementNameLabel.Text = "Name: {0}".format(familyName)
            self.selectedElementTypeLabel.Text = "Type: {0}".format(typeName)

            # Collect and display all types of parameters
            self.displayParameters(element)

        except Exception as e:
            TaskDialog.Show("Error", str(e))
                    
            TaskDialog.Show("Element Parameters", "\n".join(paramInfo))

        finally:
            # Re-enable the button after selection is made or cancelled
            self.button.Enabled = True

    def OnAddParametersButtonClick(self, sender, e):
        transaction = None

        try:
            selectedParameters = [item for item in self.listView.Items if item.Checked]
            total_parameters = len(selectedParameters)
            parameters_changed = 0
            parameters_not_changed = []

            if not selectedParameters:
                TaskDialog.Show("Info", "Please select some parameters first.")
                return

            ref = self.uidoc.Selection.PickObject(ObjectType.Element, "Select another element to apply parameters")
            targetElement = self.uidoc.Document.GetElement(ref.ElementId)

            transaction = Transaction(self.uidoc.Document, "Add Parameters to Element")
            transaction.Start()

            for item in selectedParameters:
                paramName = item.SubItems[0].Text
                paramIsTypeParam = self.isTypeParameter(paramName)

                if paramName in self.parameters_in_memory:
                    originalValue, _, _ = self.parameters_in_memory[paramName]

                    if paramIsTypeParam:
                        typeId = targetElement.GetTypeId()
                        targetElementType = self.uidoc.Document.GetElement(typeId)
                        targetParam = targetElementType.LookupParameter(paramName)
                    else:
                        targetParam = targetElement.LookupParameter(paramName)

                    if not targetParam or targetParam.IsReadOnly:
                        parameters_not_changed.append(paramName)
                        continue

                    if targetParam.StorageType == StorageType.Double:
                        targetParam.Set(float(originalValue))
                    elif targetParam.StorageType == StorageType.Integer:
                        targetParam.Set(int(originalValue))
                    elif targetParam.StorageType == StorageType.String:
                        targetParam.Set(str(originalValue))
                    elif targetParam.StorageType == StorageType.ElementId:
                        # Handle ElementId if necessary
                        pass
                    else:
                        parameters_not_changed.append(paramName)
                        continue

                    parameters_changed += 1

            transaction.Commit()

            summary_message = "{}/{} parameters changed.".format(parameters_changed, total_parameters)
            if parameters_not_changed:
                summary_message += "\nParameters not changed: " + ", ".join(parameters_not_changed)

            TaskDialog.Show("Summary", summary_message)

        except Exception as e:
            if transaction and transaction.HasStarted():
                transaction.RollBack()
            TaskDialog.Show("Error", str(e))

    def isTypeParameter(self, paramName):
        # Implement logic to determine if the parameter is a type parameter
        # This could be based on how you're storing parameter information
        # Example logic (modify as per your data structure):
        _, _, paramType = self.parameters_in_memory.get(paramName, (None, None, None))
        return paramType == "Type"

    def displayParameters(self, element):
        self.parameters_in_memory = {}
        self.fullListViewItems = []  # Reset the full list of items

        # Collect instance parameters
        self.collectParameters(element.Parameters, "Instance", element)

        # Collect type parameters
        elementType = self.uidoc.Document.GetElement(element.GetTypeId())
        if elementType:
            self.collectParameters(elementType.Parameters, "Type", elementType)

        # Populate ListView with collected parameters
        for paramName, (originalValue, convertedValue, paramType) in self.parameters_in_memory.items():
            categoryStr = self.getParameterCategory(paramName, paramType, element, elementType)
            if categoryStr != "Invalid Group":  # Skip parameters with 'Invalid Group'
                item = ListViewItem(Array[str]([paramName, convertedValue, categoryStr, paramType]))
                self.listView.Items.Add(item)

        # After populating listView.Items, also populate fullListViewItems
        for item in self.listView.Items:
            self.fullListViewItems.append(item.Clone())

    def collectParameters(self, parameters, paramType, element):
        for param in parameters:
            paramName = param.Definition.Name
            originalValue = self.getOriginalParameterValue(param)
            convertedValue = self.getConvertedParameterValue(param)
            # Store the parameter along with its type and shared status
            self.parameters_in_memory[paramName] = (originalValue, convertedValue, paramType)
    
    def getOriginalParameterValue(self, param):
        if param.StorageType == StorageType.Double:
            return param.AsDouble()
        elif param.StorageType == StorageType.Integer:
            return param.AsInteger()
        elif param.StorageType == StorageType.String:
            return param.AsString()
        elif param.StorageType == StorageType.ElementId:
            return param.AsElementId()
        else:
            return None  # Handle unsupported types

    def getConvertedParameterValue(self, param):
        if param.StorageType == StorageType.Double:
            if param.Definition.ParameterType == ParameterType.Length:
                value_in_mm = UnitUtils.ConvertFromInternalUnits(param.AsDouble(), UnitTypeId.Millimeters)
                return "{:.1f}".format(value_in_mm)
            else:
                return "{:.1f}".format(param.AsDouble())
        elif param.StorageType == StorageType.Integer:
            return str(param.AsInteger())
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
        
    def getStorageTypeFromParamType(self, paramType):
        if paramType == "Double":
            return StorageType.Double
        elif paramType == "Integer":
            return StorageType.Integer
        elif paramType == "String":
            return StorageType.String
        # Add other cases as needed
        else:
            return None  # Or some default value
        
    def OnClearMemoryButtonClick(self, sender, e):
       # Clear both the memory and the ListView
        self.parameters_in_memory = {}
        self.fullListViewItems = []
        self.listView.Items.Clear()
        self.searchBox.Text = "Search"  # Clear the search box text

        # Reset the label to its original state
        self.selectedElementNameLabel.Text = "Name: None"
        self.selectedElementTypeLabel.Text = "Type: None"
            
    def OnColumnHeaderClick(self, sender, e):
        # Set the column number that is to be sorted; default to ascending
        self.listViewColumnSorter.columnToSort = e.Column
        self.listView.Sort()

    def OnSearchBoxEnter(self, sender, e):
        # Clear the text when the search box is entered
        if self.searchBox.Text == "Search":
            self.searchBox.Text = ""

    def OnSearchBoxTextChanged(self, sender, e):
        search_text = self.searchBox.Text.lower()
        self.listView.Items.Clear()  # Clear current items in ListView

        # Filter and repopulate the ListView based on the search text
        for item in self.fullListViewItems:
            if search_text in item.SubItems[0].Text.lower():
                self.listView.Items.Add(item.Clone())  # Clone the item to add it to ListView

    def OnSelectAllCheckBoxChanged(self, sender, e):
        for item in self.listView.Items:
            item.Checked = self.selectAllCheckBox.Checked

# To run this form, you need to pass the UIDocument from a RevitPythonShell script
uidoc = __revit__.ActiveUIDocument
form = SimpleForm(uidoc)
Application.Run(form)
