# -*- coding: utf-8 -*-
import clr
import os
import csv
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Drawing import Font, FontStyle
from System.Windows.Forms import (PictureBox, PictureBoxSizeMode, AnchorStyles, Panel, Cursors, TextBox, BorderStyle, HorizontalAlignment)
from System.Drawing import Color, Size, Point, Bitmap
from Snippets._imagePath import ImagePathHelper

# Instantiate ImagePathHelper
image_helper = ImagePathHelper()

# Load images using the get_image_path function
search_image_path = image_helper.get_image_path('Search.png')
clear_image_path = image_helper.get_image_path('Clear.png')

# Create Bitmap objects from the paths
search_image = Bitmap(search_image_path)
clear_image = Bitmap(clear_image_path)

class SearchBox(Panel):
    def __init__(self, parent_form, search_image, clear_image):
        super(SearchBox, self).__init__()
        self.parent_form = parent_form
        self.init_ui(search_image, clear_image)

    def init_ui(self, search_image, clear_image):
        # Define colors
        colorText = Color.FromArgb(240, 240, 240)
        colorBackground = Color.FromArgb(49, 49, 49)

        # Set the size and location of the panel
        self.Size = Size(180, 22)
        self.Location = Point(275, 42)
        self.BackColor = colorBackground
        self.Anchor = AnchorStyles.Top | AnchorStyles.Right

        self.border = Panel()
        self.border.Size = Size(180,22)
        self.border.BackColor = Color.FromArgb(69,69,69)
        self.border.Anchor = AnchorStyles.Top | AnchorStyles.Right

        self.cover = Panel()
        self.cover.Size = Size(178,20)
        self.cover.Location = Point(1,1)
        self.cover.BackColor = colorBackground
        self.cover.Anchor = AnchorStyles.Top | AnchorStyles.Right

        # Create the search icon PictureBox
        self.searchIcon = PictureBox()
        self.searchIcon.Image = search_image
        self.searchIcon.SizeMode = PictureBoxSizeMode.StretchImage
        self.searchIcon.Size = Size(15, 15)
        self.searchIcon.Location = Point(3, 4)
        self.searchIcon.BackColor = colorBackground

        # Create the search TextBox
        self.searchBox = TextBox()
        self.searchBox.Text = "Search"
        self.searchBox.Font = Font("Helvetica", 10, FontStyle.Regular)
        self.searchBox.Size = Size(135, 20)
        self.searchBox.Location = Point(self.searchIcon.Right + 5, 3)
        self.searchBox.BorderStyle = BorderStyle.None
        self.searchBox.BackColor = colorBackground
        self.searchBox.ForeColor = colorText
        self.searchBox.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        self.searchBox.TextChanged += self.on_search_changed

        # Create the clear button PictureBox
        self.clearButton = PictureBox()
        self.clearButton.Image = clear_image
        self.clearButton.SizeMode = PictureBoxSizeMode.StretchImage
        self.clearButton.Size = Size(15, 15)
        self.clearButton.Location = Point(self.searchBox.Right + 5, 4)
        self.clearButton.BackColor = colorBackground
        self.clearButton.Anchor = AnchorStyles.Top | AnchorStyles.Right
        self.clearButton.Cursor = Cursors.Hand
        self.clearButton.Click += self.clear_search_text
        self.clearButton.Visible = False  # Initially hidden

        # Add the controls to the panel
        self.Controls.Add(self.searchIcon)
        self.Controls.Add(self.searchBox)
        self.Controls.Add(self.clearButton)
        self.Controls.Add(self.cover)
        self.Controls.Add(self.border)

        # Event handlers
        self.searchBox.TextChanged += self.on_search_changed
        self.searchBox.Enter += self.search_box_enter
        self.searchBox.Leave += self.search_box_leave

    def search_box_enter(self, sender, e):
        if self.searchBox.Text == "Search":
            self.searchBox.Text = ""
            self.clearButton.Visible = True

    def search_box_leave(self, sender, e):
        if not self.searchBox.Text:
            self.searchBox.Text = "Search"
            self.clearButton.Visible = False

    def on_search_changed(self, sender, e):
        searchText = self.searchBox.Text.lower()
        # Check if the text is not the default placeholder, then conduct the search
        if searchText != "search" and searchText.strip() != "":
            # Make the clear button visible whenever there's text to clear
            self.clearButton.Visible = True

            # Clear current items in the checklist
            self.parent_form.checklist.Items.Clear()

            # Filter and add back only the items that match the search text
            for schedule in self.parent_form.allSchedules:
                if searchText in schedule.Name.lower():
                    self.parent_form.checklist.Items.Add(schedule.Name)
        elif searchText.strip() == "":
            # If the search box is cleared, show all schedules again
            self.clearButton.Visible = False
            self.parent_form.checklist.Items.Clear()
            for schedule in self.parent_form.allSchedules:
                self.parent_form.checklist.Items.Add(schedule.Name)

    def clear_search_text(self, sender, e):
        self.searchBox.Text = ""
        self.clearButton.Visible = False
        self.parent_form.checklist.Items.Clear()
        for schedule in self.parent_form.allSchedules:
            self.parent_form.checklist.Items.Add(schedule.Name)
        # Optionally, if you want to reset the selection state, you can add:
        #self.parent_form.checklist.ClearSelected()

