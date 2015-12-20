#!/usr/bin/env python
#-*- coding: utf-8 -*-

### 2008-2015 Charlie Barnes.

### This program is free software; you can redistribute it and/or modify
### it under the terms of the GNU General Public License as published by
### the Free Software Foundation; either version 2 of the License, or
### (at your option) any later version.

### This program is distributed in the hope that it will be useful,
### but WITHOUT ANY WARRANTY; without even the implied warranty of
### MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
### GNU General Public License for more details.

### You should have received a copy of the GNU General Public License
### along with this program; if not, write to the Free Software
### Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA

import gobject
import gtk
import cfg

def setup_vice_county_treeview(treeview):
    '''Create a model for a vice-county treeview and populate.'''
    
    store = gtk.ListStore(str, str)
    
    for vc in cfg.vc_list:
        store.append([str(vc[0]), vc[1]])

    treeview.set_rules_hint(True)
    treeview.get_selection().set_mode(gtk.SELECTION_MULTIPLE)
    treeview.set_model(store)

    rendererText = gtk.CellRendererText()
    column = gtk.TreeViewColumn("#", rendererText, text=0)
    column.set_sort_column_id(0)
    treeview.append_column(column)

    rendererText = gtk.CellRendererText()
    column = gtk.TreeViewColumn("Vice-county", rendererText, text=1)
    column.set_sort_column_id(1)
    treeview.append_column(column)

def setup_family_treeview(treeview):
    '''Create a model for a family treeview.'''

    treeview.set_rules_hint(True)
    treeview.get_selection().set_mode(gtk.SELECTION_MULTIPLE)

    rendererText = gtk.CellRendererText()
    column = gtk.TreeViewColumn("Family", rendererText, text=0)
    treeview.append_column(column)

def setup_combo_box(combobox, data):
    '''Create a model for a combobox using data.'''
    
    model = gtk.ListStore(gobject.TYPE_STRING)

    for i in range(len(data)):
        model.append([data[i]])

    combobox.set_model(model)
    cell = gtk.CellRendererText()
    combobox.pack_start(cell, True)
    combobox.add_attribute(cell, 'text',0)

def setup_image_file_chooser(filechooserbutton):
    '''Create filter and preview widget for image file chooser buttons.'''
    
    filter = gtk.FileFilter()
    filter.set_name("Supported image files")
    filter.add_pattern("*.png")
    filter.add_pattern("*.jpg")
    filter.add_pattern("*.jpeg")
    filter.add_pattern("*.gif")
    filter.add_mime_type("image/png")
    filter.add_mime_type("image/jpg")
    filter.add_mime_type("image/gif")
    filechooserbutton.add_filter(filter)
    filechooserbutton.set_filter(filter)

    preview = gtk.Image()
    filechooserbutton.set_preview_widget(preview)
    filechooserbutton.connect("update-preview", update_preview, preview)
    filechooserbutton.set_use_preview_label(False)
    
            
def update_preview(file_chooser, preview):
    filename = file_chooser.get_preview_filename()

    try:
        pixbuf = gtk.gdk.pixbuf_new_from_file_at_size(filename, 256, 256)
        preview.set_from_pixbuf(pixbuf)
        have_preview = True
    except:
        have_preview = False
    file_chooser.set_preview_widget_active(have_preview)
    return    
