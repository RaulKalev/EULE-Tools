# -*- coding: utf-8 -*-
import clr
import os
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Drawing import Bitmap

class ImagePathHelper:
    def __init__(self, image_folder="Icons"):
        self.image_folder = image_folder
        self.current_script_dir = os.path.dirname(__file__)

    def get_image_path(self, image_filename):
        relative_image_path = os.path.join(self.current_script_dir, self.image_folder, image_filename)
        return relative_image_path

# Example usage
try:
    image_helper = ImagePathHelper()
    test_image_path = image_helper.get_image_path('Close.png')  # Replace with your image file name
    test_bitmap = Bitmap(test_image_path)
except Exception as e:
    print("Error loading image:", e)
