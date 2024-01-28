# -*- coding: utf-8 -*-
import clr
import os
import csv
import sys
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from Autodesk.Revit.DB import *
from System.Drawing import Font, FontStyle
from System.Windows.Forms import (Form, CheckedListBox, CheckBox, Label, PictureBox, PictureBoxSizeMode, ControlStyles,DockStyle,
                                  DialogResult, AnchorStyles, MessageBox, MessageBoxButtons, Panel, Cursors, TextBox,BorderStyle, HorizontalAlignment,
                                  MessageBoxIcon, FolderBrowserDialog, FormBorderStyle, Control, MouseButtons,FormWindowState)
from System.Drawing import Color, Size, Point, Bitmap, ContentAlignment
from Snippets._imagePath import ImagePathHelper
image_helper = ImagePathHelper()
class InteractivePictureBox:
    def __init__(self, pictureBox, normalImage, hoverImage, clickImage):
        self.pictureBox = pictureBox
        # Load images from file paths
        self.normalImage = Bitmap(image_helper.get_image_path(normalImage))
        self.hoverImage = Bitmap(image_helper.get_image_path(hoverImage))
        self.clickImage = Bitmap(image_helper.get_image_path(clickImage))

        # Set initial image
        self.pictureBox.Image = self.normalImage
        self.pictureBox.SizeMode = PictureBoxSizeMode.StretchImage

        # Subscribe to mouse events
        self.pictureBox.MouseEnter += self.on_mouse_enter
        self.pictureBox.MouseLeave += self.on_mouse_leave
        self.pictureBox.MouseDown += self.on_mouse_down
        self.pictureBox.MouseUp += self.on_mouse_up

    def on_mouse_enter(self, sender, e):
        self.pictureBox.Image = self.hoverImage

    def on_mouse_leave(self, sender, e):
        self.pictureBox.Image = self.normalImage

    def on_mouse_down(self, sender, e):
        self.pictureBox.Image = self.clickImage

    def on_mouse_up(self, sender, e):
        if self.pictureBox.ClientRectangle.Contains(self.pictureBox.PointToClient(Control.MousePosition)):
            self.pictureBox.Image = self.hoverImage
        else:
            self.pictureBox.Image = self.normalImage

    def update_image(self, newImage):
        if self.pictureBox.Image is not None:
            self.pictureBox.Image.Dispose()  # Dispose the old image
        self.pictureBox.Image = newImage

    def Dispose(self):
        # Dispose all Bitmap resources
        if self.normalImage is not None:
            self.normalImage.Dispose()
        if self.hoverImage is not None:
            self.hoverImage.Dispose()
        if self.clickImage is not None:
            self.clickImage.Dispose()

        if self.pictureBox.Image is not None:
            self.pictureBox.Image.Dispose()