# -*- coding: utf-8 -*-
import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, Transaction, Family

def get_family_categories(doc):
    categories = set()
    collector = FilteredElementCollector(doc).OfClass(Family)
    for elem in collector:
        category = elem.FamilyCategory
        if category:  # Ensure the category is not None
            categories.add(category.Name)
    return sorted(list(categories))
