# -*- coding: utf-8 -*-
from System.Windows.Forms import FormWindowState, Cursors
class WindowResizer:
    def __init__(self, form, minWidth=500, minHeight=400):
        self.form = form
        self.resizeHandleSize = 10
        self.minWidth = minWidth
        self.minHeight = minHeight

        # Attach events
        form.Resize += self.on_form_resize
        form.MouseDown += self.on_resize_mouse_down
        form.MouseMove += self.on_resize_mouse_move
        form.MouseUp += self.on_resize_mouse_up

    def on_form_resize(self, sender, e):
        # Redraw the form and its contents
        sender.Invalidate()

    def on_resize_mouse_down(self, sender, e):
        if e.X >= sender.Width - self.resizeHandleSize and e.Y >= sender.Height - self.resizeHandleSize:
            sender.Tag = True  # Tag used to indicate resizing

    def on_resize_mouse_move(self, sender, e):
        if sender.Tag:
            newWidth = max(e.X, self.minWidth)
            newHeight = max(e.Y, self.minHeight)
            sender.Width = newWidth
            sender.Height = newHeight
        elif e.X >= sender.Width - self.resizeHandleSize and e.Y >= sender.Height - self.resizeHandleSize:
            sender.Cursor = Cursors.SizeNWSE
        else:
            sender.Cursor = Cursors.Default

    def on_resize_mouse_up(self, sender, e):
        sender.Tag = False
