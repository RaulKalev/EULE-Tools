import clr
import System
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from System.Windows.Forms import Form, ListView, View, Button, ColumnHeader, HorizontalAlignment, ListViewItem, ComboBox, Label, Application, ButtonBase, ControlPaint, Padding, FlatStyle
from System.Drawing import Size, Point, Color, Pen, Drawing2D, Graphics, Rectangle, Font, FontStyle
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, ElementId, Transaction
from Autodesk.Revit.UI import TaskDialog


class SheetSelectorForm(Form):
    def __init__(self, sheets, doc):
        self.Text = "Select Sheets"
        self.Size = Size(415, 650)
        self.doc = doc

         # Variables for color controls
        backgroundColor1 = Color.FromArgb(24, 24, 24)
        backgroundColor2 = Color.FromArgb(31, 31, 31)
        textColor1 = Color.FromArgb(255, 255, 255)
        buttonColor1 = Color.FromArgb(66, 66, 66)
        font1 = Font("Arial", 12,)

        #self.BackColor = backgroundColor1

        self.listView = ListView()
        self.listView.View = View.Details
        self.listView.CheckBoxes = True
        self.listView.FullRowSelect = True
        self.listView.Columns.Add("Sheet Name", -2, HorizontalAlignment.Left)
        self.listView.Columns.Add("Sheet Number", -2, HorizontalAlignment.Left)
        self.listView.Columns.Add("Current Revision", -2, HorizontalAlignment.Left)
        self.listView.Size = Size(380, 500)
        self.listView.Location = Point(10, 10)
        #self.listView.BackColor = backgroundColor2
        #self.listView.ForeColor = textColor1

        for sheet in sheets:
            item = ListViewItem(sheet.Name)
            item.SubItems.Add(sheet.SheetNumber)
            item.Tag = sheet.Id  # Store the sheet ID in the tag property
            current_revision = self.get_latest_revision(sheet)
            item.SubItems.Add(current_revision)
            self.listView.Items.Add(item)

        self.revisionLabel = Label()
        self.revisionLabel.Text = "Select Revision:"
        self.revisionLabel.Font = font1
        self.revisionLabel.Location = Point(10, 520)
        self.revisionLabel.Size = Size(130, 20)
        #self.revisionLabel.ForeColor = textColor1
        self.Controls.Add(self.revisionLabel)

        self.revisionComboBox = ComboBox()
        self.revisionComboBox.Location = Point(140, 520)
        self.revisionComboBox.Size = Size(200, 20)
        #self.revisionComboBox.BackColor = backgroundColor1
        #self.revisionComboBox.ForeColor = textColor1
        self.load_revisions(doc)
        self.Controls.Add(self.revisionComboBox)

        self.applyButton = Button()
        self.applyButton.Text = "Apply"
        self.applyButton.Location = Point(300, 560)
        self.applyButton.Click += self.OnApply
        #self.applyButton.FlatStyle = FlatStyle.Flat
        #self.applyButton.FlatAppearance.BorderSize = 0
        #self.applyButton.BackColor = buttonColor1
        #self.applyButton.ForeColor = textColor1
        self.Controls.Add(self.applyButton)
        
        self.removeRevisionButton = Button()
        self.removeRevisionButton.Text = "Remove"
        self.removeRevisionButton.Location = Point(220, 560)
        self.removeRevisionButton.Click += self.OnRemoveRevision
        #self.removeRevisionButton.FlatStyle = FlatStyle.Flat
        #self.removeRevisionButton.FlatAppearance.BorderSize = 0
        #self.removeRevisionButton.BackColor = buttonColor1
        #self.removeRevisionButton.ForeColor = textColor1
        self.Controls.Add(self.removeRevisionButton)

        self.selectAllButton = Button()
        self.selectAllButton.Text = "Select All"
        self.selectAllButton.Location = Point(10, 560)
        self.selectAllButton.Click += self.OnSelectAll
        #self.selectAllButton.FlatStyle = FlatStyle.Flat
        #self.selectAllButton.FlatAppearance.BorderSize = 0
        #self.selectAllButton.BackColor = buttonColor1
        #self.selectAllButton.ForeColor = textColor1
        self.Controls.Add(self.selectAllButton)

        self.deselectAllButton = Button()
        self.deselectAllButton.Text = "Deselect All"
        self.deselectAllButton.Location = Point(90, 560)
        self.deselectAllButton.Click += self.OnDeselectAll
        #self.deselectAllButton.FlatStyle = FlatStyle.Flat
        #self.deselectAllButton.FlatAppearance.BorderSize = 0
        #self.deselectAllButton.BackColor = buttonColor1
        #self.deselectAllButton.ForeColor = textColor1
        self.Controls.Add(self.deselectAllButton)

        self.Controls.Add(self.listView)

    def load_revisions(self, doc):
        revisions = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Revisions).WhereElementIsNotElementType().ToElements()
        for rev in revisions:
            self.revisionComboBox.Items.Add(rev.Name)

    def get_latest_revision(self, sheet):
        revs = sheet.GetAllRevisionIds()
        if revs.Count > 0:
            latest_rev = max([self.doc.GetElement(rev_id).RevisionNumber for rev_id in revs])
            return latest_rev
        else:
            return "No Revision"

    def OnApply(self, sender, args):
        selected_revision_name = self.revisionComboBox.SelectedItem
        selected_revision_id = None

        revisions = FilteredElementCollector(self.doc).OfCategory(BuiltInCategory.OST_Revisions).WhereElementIsNotElementType().ToElements()
        for rev in revisions:
            if rev.Name == selected_revision_name:
                selected_revision_id = rev.Id
                break

        if selected_revision_id is None:
            TaskDialog.Show("Error", "Selected revision not found in document.")
            return

        selected_sheet_ids = [item.Tag for item in self.listView.Items if item.Checked]

        transaction = Transaction(self.doc)
        transaction.Start("Apply Revision")
        for sheet_id in selected_sheet_ids:
            sheet = self.doc.GetElement(sheet_id)
            current_revs = sheet.GetAllRevisionIds()
            if selected_revision_id not in current_revs:
                current_revs.Add(selected_revision_id)
                sheet.SetAdditionalRevisionIds(current_revs)
        transaction.Commit()

        TaskDialog.Show("Success", "Applied revision '{}' to {} sheets".format(selected_revision_name, len(selected_sheet_ids)))
        self.Close()

    def OnRemoveRevision(self, sender, args):
        selected_revision_name = self.revisionComboBox.SelectedItem
        selected_revision_id = None

        revisions = FilteredElementCollector(self.doc).OfCategory(BuiltInCategory.OST_Revisions).WhereElementIsNotElementType().ToElements()
        for rev in revisions:
            if rev.Name == selected_revision_name:
                selected_revision_id = rev.Id
                break

        if selected_revision_id is None:
            TaskDialog.Show("Error", "Selected revision not found in document.")
            return

        selected_sheet_ids = [item.Tag for item in self.listView.Items if item.Checked]

        transaction = Transaction(self.doc)
        transaction.Start("Remove Revision")
        for sheet_id in selected_sheet_ids:
            sheet = self.doc.GetElement(sheet_id)
            current_revs = sheet.GetAllRevisionIds()
            if selected_revision_id in current_revs:
                current_revs.Remove(selected_revision_id)
                sheet.SetAdditionalRevisionIds(current_revs)
        transaction.Commit()

        TaskDialog.Show("Success", "Removed revision '{}' from {} sheets".format(selected_revision_name, len(selected_sheet_ids)))

        self.Close()  # Add this line to close the form

    def OnSelectAll(self, sender, args):
        for item in self.listView.Items:
            item.Checked = True

    def OnDeselectAll(self, sender, args):
        for item in self.listView.Items:
            item.Checked = False

def show_sheet_selector():
    doc = __revit__.ActiveUIDocument.Document
    sheets = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Sheets).WhereElementIsNotElementType().ToElements()
    form = SheetSelectorForm(sheets, doc)
    form.ShowDialog()  # This line is changed

show_sheet_selector()
