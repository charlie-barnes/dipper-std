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

class CellRendererClickableText(gtk.CellRendererText):

    __gsignals__ = {'clicked': (gobject.SIGNAL_RUN_LAST, gobject.TYPE_NONE,
                                (gobject.TYPE_STRING,))
                   }

    def __init__(self):
        gtk.CellRendererText.__init__(self)
        self.set_property('mode', gtk.CELL_RENDERER_MODE_ACTIVATABLE)

    def do_activate(self, event, widget, path, background_area, cell_area,
                    flags):
        self.emit('clicked', path)

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

    model = gtk.TreeStore(str, str)
    
    treeview.set_model(model)       
    treeview.set_rules_hint(True)
    treeview.get_selection().set_mode(gtk.SELECTION_MULTIPLE)

    rendererText = gtk.CellRendererText()
    column = gtk.TreeViewColumn("name", rendererText, text=0)
    treeview.append_column(column)

def setup_combo_box(combobox, data):
    '''Create a model for a combobox using data.'''
    combobox.clear()
    
    model = gtk.ListStore(gobject.TYPE_STRING)

    for i in range(len(data)):
        model.append([data[i]])

    combobox.set_model(model)
    cell = gtk.CellRendererText()
    combobox.pack_start(cell, True)
    combobox.add_attribute(cell, 'text', 0)

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
    
def setup_mapping_layers_treeview(container, config=None):
    '''Create the mapping layers treeview, selecting config defaults if present'''

    try:    
        container.get_child().destroy()
    except AttributeError:
        pass

    notebook = gtk.Notebook()
        
    if config is not None:
        sel_config = []
        
        for sel_config_opt in config:
            sel_config.append([sel_config_opt[0],sel_config_opt[1]])
    
    #loop through each available category of mapping layers, adding the list for each one
    for category in cfg.gis.keys():      
        store = gtk.ListStore(str, str, str)
          
        treeview = gtk.TreeView()
        treeview.set_rules_hint(False)
        selection= treeview.get_selection()
        selection.set_mode(gtk.SELECTION_MULTIPLE)
        treeview.set_model(store)
        treeview.set_headers_visible(True)
        
        selection = treeview.get_selection()
                    
        #add each avaialble layer BUT not if present in our config file
        for layer in sorted(cfg.gis[category]):
            
            try:
                if [category, layer] not in sel_config:
                    iter = store.append([layer,'   <span background="#f2f2f2">      </span>   ', '   <span background="#797979">      </span>   '])
 
            except UnboundLocalError:
                iter = store.append([layer,'   <span background="#f2f2f2">      </span>   ', '   <span background="#797979">      </span>   '])
        
        #loop through the config options if set and add our selected mapping layers
        try:
            for opts in config:
                if opts[0] == category:
                    iter = store.append([opts[1],''.join(['   <span background="', opts[2],'">      </span>   ']), ''.join(['   <span background="', opts[3],'">      </span>   '])   ])
                    selection.select_iter(iter)
        except TypeError:
            pass

        store.set_sort_column_id(0,gtk.SORT_ASCENDING)  
        
        rendererText = gtk.CellRendererText()
        column = gtk.TreeViewColumn("Area", rendererText, text=0)
        column.set_property("expand", True)
        column.set_sort_column_id(0)
        treeview.append_column(column)
        
        renderer = CellRendererClickableText()
        renderer.connect("clicked", color_cell_edited, [store, 1, selection])
        column = gtk.TreeViewColumn("Fill", renderer, markup=1)   
        column.set_fixed_width(75) 
        treeview.append_column(column)
       
        renderer = CellRendererClickableText()
        renderer.connect("clicked", color_cell_edited, [store, 2, selection])
        column = gtk.TreeViewColumn("Border", renderer, markup=2)  
        column.set_fixed_width(75)
        treeview.append_column(column)

        model, selected = selection.get_selected_rows()
        try:
            treeview.scroll_to_cell(selected[0])
        except IndexError:
            pass
        
        label = gtk.Label(category)
        scrolled_window = gtk.ScrolledWindow()
        scrolled_window.set_shadow_type(gtk.SHADOW_NONE)
        scrolled_window.add(treeview) 
        notebook.append_page(scrolled_window, tab_label=label)

    container.add(notebook)
    notebook.show_all()



def color_cell_edited(widget, path, userdata):
    #hacky but it works        
    selection = userdata[2]
    
    pre_colour = userdata[0][path][userdata[1]].split('"')
    
    if selection.path_is_selected(path):
        dialog = gtk.ColorSelectionDialog('Pick a Colour')
        dialog.colorsel.set_current_color(gtk.gdk.color_parse(pre_colour[1]))

        response = dialog.run()
        dialog.destroy()

        if response == gtk.RESPONSE_OK:
            color = dialog.colorsel.get_current_color()   
            userdata[0][path][userdata[1]] = ''.join(['   <span background="', str(color), '">      </span>   '])
         
