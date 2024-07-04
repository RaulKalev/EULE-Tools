# -*- coding: utf-8 -*-
import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

class RevitRegionUtilities:
    def __init__(self, doc):
        self.doc = doc
        self.solid_pattern = self.get_solid_pattern()

    def get_solid_pattern(self):
        """Retrieve the first solid fill pattern available in the document."""
        all_patterns = FilteredElementCollector(self.doc).OfClass(FillPatternElement).ToElements()
        solid_patterns = [pat for pat in all_patterns if pat.GetFillPattern().IsSolidFill]
        return solid_patterns[0] if solid_patterns else None

    def lighten_color(self, color, factor):
        """Lightens the given color by mixing it with white."""
        white = Color(255, 255, 255)
        new_red = int((color.Red * (1 - factor)) + (white.Red * factor))
        new_green = int((color.Green * (1 - factor)) + (white.Green * factor))
        new_blue = int((color.Blue * (1 - factor)) + (white.Blue * factor))
        return Color(new_red, new_green, new_blue)

    def get_filled_region(self, filled_region_name):
        """Retrieve a filled region type by name."""
        pvp = ParameterValueProvider(ElementId(BuiltInParameter.ALL_MODEL_TYPE_NAME))
        condition = FilterStringEquals()
        fRule = FilterStringRule(pvp, condition, filled_region_name, True)
        my_filter = ElementParameterFilter(fRule)
        return FilteredElementCollector(self.doc).OfClass(FilledRegionType).WherePasses(my_filter).FirstElement()

    def create_region_type(self, name, color, masking=False, lineweight=1):
        """Create a new FilledRegionType with specified properties."""
        random_filled_region = FilteredElementCollector(self.doc).OfClass(FilledRegionType).FirstElement()
        if not random_filled_region:
            return None  # No filled region type available to duplicate
        new_region = random_filled_region.Duplicate(name)
        new_region.BackgroundPatternId = ElementId(-1)
        new_region.ForegroundPatternId = self.solid_pattern.Id
        new_region.BackgroundPatternColor = color
        new_region.ForegroundPatternColor = color
        new_region.IsMasking = masking
        new_region.LineWeight = lineweight
        return new_region
