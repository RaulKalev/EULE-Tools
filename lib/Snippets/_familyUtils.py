# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import FilteredElementCollector, Family
from System import Array

def get_families_by_category(doc, category_name):
    """Returns a sorted list of family names within a given category."""
    families = []
    collector = FilteredElementCollector(doc).OfClass(Family)
    for family in collector:
        if family.FamilyCategory.Name == category_name:
            families.append(family.Name)
    return sorted(set(families))

def on_category_changed(doc, comboBoxFamily, category_name):
    """Handles category selection changes and updates families ComboBox."""
    families = get_families_by_category(doc, category_name)
    comboBoxFamily.Items.Clear()
    comboBoxFamily.Items.AddRange(Array[object](families))
