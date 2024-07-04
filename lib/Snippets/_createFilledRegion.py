# -*- coding: utf-8 -*-
import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import Color as RevitColor
from System.Windows.Forms import Application, Form, Button, TextBox, ColorDialog, DialogResult, CheckBox
from System.Drawing import Point, Size, Color

doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app   = __revit__.Application
rvt_year = int(app.VersionNumber)

region_name_A = 'New Region A'
region_name_B = 'New Region B'

# Get Solid Pattern
all_patterns       = FilteredElementCollector(doc).OfClass(FillPatternElement).ToElements()
all_solid_patterns = [pat for pat in all_patterns if pat.GetFillPattern().IsSolidFill]
solid_pattern      = all_solid_patterns[0]

def lighten_color(color, factor):
    """Lightens the given color by mixing it with white."""
    white = RevitColor(255, 255, 255)
    new_red = int((color.Red * (1 - factor)) + (white.Red * factor))
    new_green = int((color.Green * (1 - factor)) + (white.Green * factor))
    new_blue = int((color.Blue * (1 - factor)) + (white.Blue * factor))
    return RevitColor(new_red, new_green, new_blue)

def get_filled_region(filled_region_name):
    """Function to get FireWall Types based on FamilyName"""
    # Create Filter
    pvp         = ParameterValueProvider(ElementId(BuiltInParameter.ALL_MODEL_TYPE_NAME))
    condition   = FilterStringEquals()
    fRule =     FilterStringRule(pvp, condition, filled_region_name, True) if rvt_year < 2022 \
             else FilterStringRule(pvp, condition, filled_region_name)
    my_filter   = ElementParameterFilter(fRule)

    # Get Types
    return FilteredElementCollector(doc).OfClass(FilledRegionType).WherePasses(my_filter).FirstElement()

def create_RegionType(name, color, masking = False, lineweight = 1):
    #type: (str, Color, bool, int) -> FilledRegionType
    """Create FilledRegionType with Solid Pattern."""
    random_filled_region = FilteredElementCollector(doc).OfClass(FilledRegionType).FirstElement()
    new_region           = random_filled_region.Duplicate(name)

    # Set Solid Pattern
    new_region.BackgroundPatternId = ElementId(-1)
    new_region.ForegroundPatternId = solid_pattern.Id

    # Set Colour
    new_region.BackgroundPatternColor = color
    new_region.ForegroundPatternColor = color

    # Masking
    new_region.IsMasking = masking

    # LineWeight
    new_region.LineWeight = lineweight

    return new_region

class MainForm(Form):
    def __init__(self):
        self.selectedColorRGB = (0, 0, 0)  # Default color RGB as a tuple (R, G, B)
        self.InitializeComponent()
    
    def InitializeComponent(self):
        self.Text = "Create Filled Region"
        self.Width = 300
        self.Height = 300  # Adjusted for additional controls

        self.prefixTextBox = TextBox()
        self.prefixTextBox.Location = Point(50, 20)
        self.prefixTextBox.Size = Size(200, 20)
        self.Controls.Add(self.prefixTextBox)

        self.colorButton = Button()
        self.colorButton.Text = 'Select Color'
        self.colorButton.Location = Point(50, 50)
        self.colorButton.Click += self.SelectColor
        self.Controls.Add(self.colorButton)

        # CheckBoxes for line weights
        self.lineWeightCheckboxes = []
        lineWeights = ["0px", "25px", "63px", "125px", "250px"]
        for i, weight in enumerate(lineWeights):
            checkBox = CheckBox()
            checkBox.Text = weight
            checkBox.Location = Point(50, 80 + (i * 30))
            checkBox.Size = Size(100, 20)
            self.Controls.Add(checkBox)
            self.lineWeightCheckboxes.append(checkBox)


        self.runButton = Button()
        self.runButton.Text = 'Run'
        self.runButton.Location = Point(170, 90)  # Adjusted location
        self.runButton.Click += self.RunScript
        self.Controls.Add(self.runButton)

    def SelectColor(self, sender, args):
        colorDialog = ColorDialog()
        if colorDialog.ShowDialog() == DialogResult.OK:
            self.selectedColorRGB = (colorDialog.Color.R, colorDialog.Color.G, colorDialog.Color.B)

    def RunScript(self, sender, args):
        transaction_title = "Create Filled Regions Based on Selection"
        prefix = self.prefixTextBox.Text.strip()  # Use strip() to remove leading/trailing whitespace
        
        # Fetch selected line weights
        selectedWeights = [cb.Text for cb in self.lineWeightCheckboxes if cb.Checked]
        
        # Original color specified by the user
        originalColor = RevitColor(*self.selectedColorRGB)  # Unpacking the RGB tuple
        whiteColor = RevitColor(255, 255, 255)  # White color for the "0px" checkbox
        
        # Ensure there's at least one selection
        if not selectedWeights:
            print("No line weights selected.")
            return

        with Transaction(doc, transaction_title) as t:
            t.Start()
            
            for index, weight in enumerate(selectedWeights):
                # Conditionally format the unique_region_name to include an underscore if the prefix is not empty
                unique_region_name = prefix + ("_" if prefix else "") + weight  # Avoid extra underscore if prefix is empty
                
                # Determine the color based on the checkbox index
                if weight == "0px":
                    currentColor = whiteColor  # "0px" is always white
                else:
                    factor = (index - 1) * 0.2  # Adjust factor for subsequent checkboxes, starting from the selected color
                    currentColor = lighten_color(originalColor, min(factor, 0.8))  # Cap factor to avoid too light colors

                # Check if this filled region type already exists
                existing_region = get_filled_region(unique_region_name)
                
                if not existing_region:
                    # If it does not exist, create a new filled region type with the adjusted color
                    create_RegionType(name=unique_region_name, color=currentColor)
                    print("Created Filled Region Type: " + unique_region_name)
                else:
                    print("Filled Region Type already exists: " + unique_region_name)
            
            t.Commit()

if __name__ == "__main__":
    form = MainForm()
    Application.Run(form)
