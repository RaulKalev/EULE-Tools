# -*- coding: utf-8 -*-
import clr
import os
import csv
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from Autodesk.Revit.DB import FilteredElementCollector, ViewSchedule, SectionType
from System.Drawing import Font, FontStyle
from System.Windows.Forms import (Form, CheckedListBox, CheckBox, Label, PictureBox, PictureBoxSizeMode, ControlStyles,DockStyle,
                                  DialogResult, AnchorStyles, MessageBox, MessageBoxButtons, Panel, Cursors, TextBox,BorderStyle, HorizontalAlignment,
                                  MessageBoxIcon, FolderBrowserDialog, FormBorderStyle, Control, MouseButtons,FormWindowState)
from System.Drawing import Color, Size, Point, Bitmap, ContentAlignment
from Snippets._interactivePictureBox import InteractivePictureBox
from Snippets._imagePath import ImagePathHelper

# Instantiate ImagePathHelper
image_helper = ImagePathHelper()

windowWidth = 500 

# Load images using the get_image_path function
minimize_image_path = image_helper.get_image_path('Minimize.png')
close_image_path = image_helper.get_image_path('Close.png')
logo_image_path = image_helper.get_image_path('Logo.png')

# Create Bitmap objects from the paths
minimize_image = Bitmap(minimize_image_path)
close_image = Bitmap(close_image_path)
logo_image = Bitmap(logo_image_path)

class TitleBar(Panel):
    def __init__(self, parent_form, title, logo_image, minimize_image, close_image):
        super(TitleBar, self).__init__()
        self.windowWidth = windowWidth
        self.parent_form = parent_form
        self.init_ui(title, logo_image, minimize_image, close_image)
        self.Resize += self.on_titlebar_resize


    def init_ui(self, title, logo_image, minimize_image, close_image):
        # Set title bar properties
        self.BackColor = Color.FromArgb(49, 49, 49)
        self.Dock = DockStyle.Top
        self.Height = 30
        self.MouseDown += self.form_mouse_down
        self.MouseMove += self.form_mouse_move
        self.MouseUp += self.form_mouse_up

        # Create and add title label
        self.titleLabel = Label()
        self.titleLabel.Text = title
        self.titleLabel.Font = Font("Helvetica", 10, FontStyle.Regular)  # Adjust the font size as needed
        self.titleLabel.TextAlign = ContentAlignment.MiddleCenter
        self.titleLabel.Dock = DockStyle.Fill
        self.titleLabel.MouseDown += self.form_mouse_down
        self.titleLabel.MouseMove += self.form_mouse_move
        self.titleLabel.MouseUp += self.form_mouse_up

        #Crate a line under title bar
        self.titleLine = Panel()
        self.titleLine.BackColor = Color.FromArgb(69,69,69)
        self.titleLine.Location = Point(0,29)

        # Add logo
        self.logo = PictureBox()
        self.logo.Image = Bitmap(logo_image)
        self.logo.SizeMode = PictureBoxSizeMode.StretchImage
        self.logo.Location = Point(10, 5)
        self.logo.Size = Size(28, 20)
    
        # Add close button
        self.closeButton = self.create_button(close_image, self.on_close_clicked)
        self.closeButton.Location = Point(self.windowWidth - 25, 5)
        self.exportButtonInteractive = InteractivePictureBox(
        self.closeButton, 'Close.png', 'CloseHover.png', 'CloseClick.png')

        # Add minimize button
        self.minimizeButton = self.create_button(minimize_image, self.on_minimize_clicked)
        self.minimizeButton.Location = Point(self.windowWidth - 45, 5)
        self.exportButtonInteractive = InteractivePictureBox(
        self.minimizeButton, 'Minimize.png', 'MinimizeHover.png', 'MinimizeClick.png')

        #Add and arrange in titlebar
        self.Controls.Add(self.titleLine)
        self.Controls.Add(self.logo)
        self.Controls.Add(self.closeButton)
        self.Controls.Add(self.minimizeButton)
        self.Controls.Add(self.titleLabel)
        

    def create_button(self, image_path, click_event):
        button = PictureBox()
        button.Image = Bitmap(image_path)
        button.SizeMode = PictureBoxSizeMode.StretchImage
        button.Size = Size(20, 20)
        button.Click += click_event
        return button

    def on_close_clicked(self, sender, e):
        self.parent_form.Close()

    def on_minimize_clicked(self, sender, e):
        self.parent_form.WindowState = FormWindowState.Minimized

    # Moving the window
    def form_mouse_down(self, sender, e):
        if e.Button == MouseButtons.Left:
            self.parent_form.dragging = True
            self.parent_form.offset = Point(e.X, e.Y)

    def form_mouse_move(self, sender, e):
        if self.parent_form.dragging:
            screenPosition = Point(self.parent_form.PointToScreen(Point(e.X, e.Y)).X - self.parent_form.offset.X,
                                   self.parent_form.PointToScreen(Point(e.X, e.Y)).Y - self.parent_form.offset.Y)
            self.parent_form.Location = screenPosition

    def form_mouse_up(self, sender, e):
        self.parent_form.dragging = False

    def on_titlebar_resize(self, sender, e):
        # Update the position of the minimize and close buttons when the title bar is resized
        self.closeButton.Location = Point(self.Width - 25, 5)
        self.minimizeButton.Location = Point(self.Width - 45, 5)
        self.titleLine.Size = Size(self.Width, 1)

