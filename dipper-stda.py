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

__version__ = "1.0a4"

import gobject
import gtk
import tempfile
import goocanvas
import csv
import cairo
from datetime import datetime
import sqlite3
import mimetypes
import sys
import os
import shutil
import time
import glob
import pango
import xlrd
import copy
from PIL import ImageChops
import math
from subprocess import call
import ConfigParser
from pygtk_chart import bar_chart

from vaguedateparse import VagueDate
from geographiccoordinatesystem import Coordinate

try:
    from fpdf import FPDF
except ImportError:
    from pyfpdf import FPDF

import shapefile
from PIL import Image
from PIL import ImageDraw

markers = []

#walk the markers directory searching for GIS markers
for style in os.listdir('markers/'):
    markers.append(style)

#grid resolution list
grid_resolution = ['100km', '10km', '5km', '2km', '1km',]

#paper sizes list
paper_size = ['A4',]

#paper orientation list
paper_orientation = ['Portrait', 'Landscape',]
                  

#vice county list - number & filename
vc_list = [[1, 'West Cornwall'],
         [2, 'East Cornwall'],
         [3, 'South Devon'],
         [4, 'North Devon'],
         [5, 'South Somerset'],
         [6, 'North Somerset'],
         [7, 'North Wiltshire'],
         [8, 'South Wiltshire'],
         [9, 'Dorset'],
         [10, 'Isle of Wight'],
         [11, 'South Hampshire'],
         [12, 'North Hampshire'],
         [13, 'West Sussex'],
         [14, 'East Sussex'],
         [15, 'East Kent'],
         [16, 'West Kent'],
         [17, 'Surrey'],
         [18, 'South Essex'],
         [19, 'North Essex'],
         [20, 'Hertfordshire'],
         [21, 'Middlesex'],
         [22, 'Berkshire'],
         [23, 'Oxfordshire'],
         [24, 'Buckinghamshire'],
         [25, 'East Suffolk'],
         [26, 'West Suffolk'],
         [27, 'East Norfolk'],
         [28, 'West Norfolk'],
         [29, 'Cambridgeshire'],
         [30, 'Bedfordshire'],
         [31, 'Huntingdonshire'],
         [32, 'Northamptonshire'],
         [33, 'East Gloucestershire'],
         [34, 'West Gloucestershire'],
         [35, 'Monmouthshire'],
         [36, 'Herefordshire'],
         [37, 'Worcestershire'],
         [38, 'Warwickshire'],
         [39, 'Stafford'],
         [40, 'Shropshire'],
         [41, 'Glamorgan'],
         [42, 'Brecon'],
         [43, 'Radnorshire'],
         [44, 'Carmarthenshire'],
         [45, 'Pembrokeshire'],
         [46, 'Cardiganshire'],
         [47, 'Montgomeryshire'],
         [48, 'Merionethshire'],
         [49, 'Caernarvonshire'],
         [50, 'Denbighshire'],
         [51, 'Flintshire'],
         [52, 'Anglesey'],
         [53, 'South Lincolnshire'],
         [54, 'North Lincolnshire'],
         [55, 'Leicestershire'],
         [56, 'Nottinghamshire'],
         [57, 'Derbyshire'],
         [58, 'Cheshire'],
         [59, 'South Lancashire'],
         [60, 'West Lancashire'],
         [61, 'South East Yorkshire'],
         [62, 'North East Yorkshire'],
         [63, 'South West Yorkshire'],
         [64, 'Mid West Yorkshire'],
         [65, 'North West Yorkshire'],
         [66, 'Durham'],
         [67, 'South Northumberland'],
         [68, 'North Northumberland'],
         [69, 'Westmorland'],
         [70, 'Cumberland'],
         [71, 'Isle of Man'],
         [72, 'Dumfrieshire'],
         [73, 'Kirkcudbrightshire'],
         [74, 'Wigtownshire'],
         [75, 'Ayrshire'],
         [76, 'Renfrewshire'],
         [77, 'Lanarkshire'],
         [78, 'Peebles'],
         [79, 'Selkirk'],
         [80, 'Roxburgh'],
         [81, 'Berwickshire'],
         [82, 'East Lothian'],
         [83, 'Midlothian'],
         [84, 'West Lothian'],
         [85, 'Fifeshire'],
         [86, 'Stirling'],
         [87, 'West Perth'],
         [88, 'Mid Perth'],
         [89, 'East Perth'],
         [90, 'Angus'],
         [91, 'Kincardine'],
         [92, 'South Aberdeen'],
         [93, 'North Aberdeen'],
         [94, 'Banff'],
         [95, 'Moray'],
         [96, 'East Inverness'],
         [97, 'West Inverness'],
         [98, 'Argyllshire'],
         [99, 'Dunbarton'],
         [100, 'Clyde Islands'],
         [101, 'Kintyre'],
         [102, 'South Ebudes'],
         [103, 'Mid Ebudes'],
         [104, 'North Ebudes'],
         [105, 'West Ross'],
         [106, 'East Ross'],
         [107, 'East Sutherland'],
         [108, 'West Sutherland'],
         [109, 'Caithness'],
         [110, 'Outer Hebrides'],
         [111, 'Orkney'],
         [112, 'Shetland']]

def repeat_to_length(string_to_expand, length):
   return (string_to_expand * ((length/len(string_to_expand))+1))[:length]


class Run():


    def __init__(self, filename=None):

        self.builder = gtk.Builder()
        self.builder.add_from_file('./gui/gui.glade')

        signals = {'quit':self.quit,
                   'generate':self.generate,
                   'open_dataset':self.open_dataset,
                   'select_all_families':self.select_all_families,
                   'unselect_image':self.unselect_image,
                   'update_title':self.update_title,
                   'show_about':self.show_about,
                  }
        self.builder.connect_signals(signals)
        self.dataset = None#


        filter = gtk.FileFilter()
        filter.set_name("Supported data files")
        filter.add_pattern("*.xls")
        filter.add_mime_type("application/vnd.ms-excel")
        self.builder.get_object('filechooserbutton3').add_filter(filter)
        self.builder.get_object('filechooserbutton3').set_filter(filter)
        filter = gtk.FileFilter()
        filter.set_name("ssconvert-able data files")
        filter.add_pattern("*.csv")
        filter.add_pattern("*.txt")
        filter.add_pattern("*.xlsx")
        filter.add_pattern("*.gnumeric")
        filter.add_pattern("*.ods")
        filter.add_mime_type("text/csv")
        filter.add_mime_type("text/plain")
        filter.add_mime_type("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        filter.add_mime_type("application/x-gnumeric")
        filter.add_mime_type("application/vnd.oasis.opendocument.spreadsheet")
        self.builder.get_object('filechooserbutton3').add_filter(filter)



        filter = gtk.FileFilter()
        filter.set_name("Supported image files")
        filter.add_pattern("*.png")
        filter.add_pattern("*.jpg")
        filter.add_pattern("*.jpeg")
        filter.add_pattern("*.gif")
        filter.add_mime_type("image/png")
        filter.add_mime_type("image/jpg")
        filter.add_mime_type("image/gif")
        self.builder.get_object('filechooserbutton5').add_filter(filter)
        self.builder.get_object('filechooserbutton5').set_filter(filter)

        preview = gtk.Image()
        self.builder.get_object('filechooserbutton5').set_preview_widget(preview)
        self.builder.get_object('filechooserbutton5').connect("update-preview", self.update_preview_cb, preview)
        self.builder.get_object('filechooserbutton5').set_use_preview_label(False)

        filter = gtk.FileFilter()
        filter.set_name("Supported image files")
        filter.add_pattern("*.png")
        filter.add_pattern("*.jpg")
        filter.add_pattern("*.jpeg")
        filter.add_pattern("*.gif")
        filter.add_mime_type("image/png")
        filter.add_mime_type("image/jpg")
        filter.add_mime_type("image/gif")
        self.builder.get_object('filechooserbutton1').add_filter(filter)
        self.builder.get_object('filechooserbutton1').set_filter(filter)

        preview = gtk.Image()
        self.builder.get_object('filechooserbutton1').set_preview_widget(preview)
        self.builder.get_object('filechooserbutton1').connect("update-preview", self.update_preview_cb, preview)
        self.builder.get_object('filechooserbutton1').set_use_preview_label(False)


        atlas_orientation_liststore = gtk.ListStore(gobject.TYPE_STRING)
        
        for i in range(len(paper_orientation)):
            atlas_orientation_liststore.append([paper_orientation[i]])
        
        atlas_orientation_combo = self.builder.get_object('combobox8')
        atlas_orientation_combo.set_model(atlas_orientation_liststore)
        cell = gtk.CellRendererText()
        atlas_orientation_combo.pack_start(cell, True)
        atlas_orientation_combo.add_attribute(cell, 'text',0)


        list_orientation_liststore = gtk.ListStore(gobject.TYPE_STRING)
        
        for i in range(len(paper_orientation)):
            list_orientation_liststore.append([paper_orientation[i]])
            
        list_orientation_combo = self.builder.get_object('combobox9')
        list_orientation_combo.set_model(list_orientation_liststore)
        cell = gtk.CellRendererText()
        list_orientation_combo.pack_start(cell, True)
        list_orientation_combo.add_attribute(cell, 'text',0)


        atlas_page_size_liststore = gtk.ListStore(gobject.TYPE_STRING)
        
        for i in range(len(paper_size)):
            atlas_page_size_liststore.append([paper_size[i]])
            
        atlas_page_size_combo = self.builder.get_object('combobox10')
        atlas_page_size_combo.set_model(atlas_page_size_liststore)
        cell = gtk.CellRendererText()
        atlas_page_size_combo.pack_start(cell, True)
        atlas_page_size_combo.add_attribute(cell, 'text',0)


        list_page_size_liststore = gtk.ListStore(gobject.TYPE_STRING)
        
        for i in range(len(paper_size)):
            list_page_size_liststore.append([paper_size[i]])
            
        list_page_size_combo = self.builder.get_object('combobox11')
        list_page_size_combo.set_model(list_page_size_liststore)
        cell = gtk.CellRendererText()
        list_page_size_combo.pack_start(cell, True)
        list_page_size_combo.add_attribute(cell, 'text',0)


        atlas_mapping_level_liststore = gtk.ListStore(gobject.TYPE_STRING)
        
        for i in range(len(grid_resolution)):
            atlas_mapping_level_liststore.append([grid_resolution[i]])
            
        atlas_mapping_level_combo = self.builder.get_object('combobox3')
        atlas_mapping_level_combo.set_model(atlas_mapping_level_liststore)
        cell = gtk.CellRendererText()
        atlas_mapping_level_combo.pack_start(cell, True)
        atlas_mapping_level_combo.add_attribute(cell, 'text',0)

        list_mapping_level_liststore = gtk.ListStore(gobject.TYPE_STRING)
        
        for i in range(len(grid_resolution)):
            list_mapping_level_liststore.append([grid_resolution[i]])
            
        combo = self.builder.get_object('combobox5')
        combo.set_model(list_mapping_level_liststore)
        cell = gtk.CellRendererText()
        combo.pack_start(cell, True)
        combo.add_attribute(cell, 'text',0)

        coverage_liststore = gtk.ListStore(gobject.TYPE_STRING)

        for i in range(len(markers)):
            coverage_liststore.append([markers[i]])

        combo = self.builder.get_object('combobox6')
        combo.set_model(coverage_liststore)
        cell = gtk.CellRendererText()
        combo.pack_start(cell, True)
        combo.add_attribute(cell, 'text',0)

        date_band_1_liststore = gtk.ListStore(gobject.TYPE_STRING)

        for i in range(len(markers)):
            date_band_1_liststore.append([markers[i]])

        combo = self.builder.get_object('combobox2')
        combo.set_model(date_band_1_liststore)
        cell = gtk.CellRendererText()
        combo.pack_start(cell, True)
        combo.add_attribute(cell, 'text',0)

        date_band_2_liststore = gtk.ListStore(gobject.TYPE_STRING)

        for i in range(len(markers)):
            date_band_2_liststore.append([markers[i]])

        combo = self.builder.get_object('combobox4')
        combo.set_model(date_band_2_liststore)
        cell = gtk.CellRendererText()
        combo.pack_start(cell, True)
        combo.add_attribute(cell, 'text',0)

        date_band_3_liststore = gtk.ListStore(gobject.TYPE_STRING)

        for i in range(len(markers)):
            date_band_3_liststore.append([markers[i]])

        combo = self.builder.get_object('combobox7')
        combo.set_model(date_band_3_liststore)
        cell = gtk.CellRendererText()
        combo.pack_start(cell, True)
        combo.add_attribute(cell, 'text',0)

        grid_lines_liststore = gtk.ListStore(gobject.TYPE_STRING)

        for i in range(len(grid_resolution)):
            grid_lines_liststore.append([grid_resolution[i]])
            
        combo = self.builder.get_object('combobox1')
        combo.set_model(grid_lines_liststore)
        cell = gtk.CellRendererText()
        combo.pack_start(cell, True)
        combo.add_attribute(cell, 'text',0)

        #family treeviews
        #atlas
        treeView = self.builder.get_object('treeview2')
        treeView.set_rules_hint(True)
        treeView.get_selection().set_mode(gtk.SELECTION_MULTIPLE)

        rendererText = gtk.CellRendererText()
        column = gtk.TreeViewColumn("Family", rendererText, text=0)
        treeView.append_column(column)

        #list
        treeView = self.builder.get_object('treeview3')
        treeView.set_rules_hint(True)
        treeView.get_selection().set_mode(gtk.SELECTION_MULTIPLE)

        rendererText = gtk.CellRendererText()
        column = gtk.TreeViewColumn("Family", rendererText, text=0)
        treeView.append_column(column)

        #vc treeview
        #atlas
        store = gtk.ListStore(str, str)
        for vc in vc_list:
            store.append([str(vc[0]), vc[1]])

        treeView = self.builder.get_object('treeview1')
        treeView.set_rules_hint(True)
        treeView.get_selection().set_mode(gtk.SELECTION_MULTIPLE)
        treeView.set_model(store)

        rendererText = gtk.CellRendererText()
        column = gtk.TreeViewColumn("#", rendererText, text=0)
        column.set_sort_column_id(0)
        treeView.append_column(column)

        rendererText = gtk.CellRendererText()
        column = gtk.TreeViewColumn("Vice-county", rendererText, text=1)
        column.set_sort_column_id(1)
        treeView.append_column(column)

        #list
        store = gtk.ListStore(str, str)
        for vc in vc_list:
            store.append([str(vc[0]), vc[1]])

        treeView = self.builder.get_object('treeview4')
        treeView.set_rules_hint(True)
        treeView.get_selection().set_mode(gtk.SELECTION_MULTIPLE)
        treeView.set_model(store)

        rendererText = gtk.CellRendererText()
        column = gtk.TreeViewColumn("#", rendererText, text=0)
        column.set_sort_column_id(0)
        treeView.append_column(column)

        rendererText = gtk.CellRendererText()
        column = gtk.TreeViewColumn("Vice-county", rendererText, text=1)
        column.set_sort_column_id(1)
        treeView.append_column(column)

        dialog = self.builder.get_object('dialog1')
        dialog.show()


    def update_preview_cb(self, file_chooser, preview):
        filename = file_chooser.get_preview_filename()

        try:
            pixbuf = gtk.gdk.pixbuf_new_from_file_at_size(filename, 256, 256)
            preview.set_from_pixbuf(pixbuf)
            have_preview = True
        except:
            have_preview = False
        file_chooser.set_preview_widget_active(have_preview)
        return

    def show_about(self, widget):
        """Show the about dialog."""
        dialog = gtk.AboutDialog()
        dialog.set_name('VA&CG')
        dialog.set_version(__version__)
        dialog.set_authors(['Charlie Barnes'])
        dialog.set_license("This is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation; either version 2 of the Licence, or (at your option) any later version.\n\nThis program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.\n\nYou should have received a copy of the GNU General Public License along with this program; if not, write to the Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA")
        dialog.set_wrap_license(True)
        dialog.set_property('skip-taskbar-hint', True)
        dialog.run()
        dialog.destroy()

    def update_title(self, selection):
        """Update the title of the generated document based on the selection
        in the families list, appending the families to the current input.

        - if more than 2 are selected use a range approach.
        - if 2 are selected, append both.
        - if just 1, append it.

        """
        model, selected = selection.get_selected_rows()
        iters = [model.get_iter(path) for path in selected]

        orig_title = self.builder.get_object('entry3').get_text().split("\n")[0]

        if len(iters) > 2:
            self.builder.get_object('entry3').set_text(''.join([orig_title,
                                                                "\n",
                                                                model.get_value(iters[0], 0),
                                                                ' to ',
                                                                model.get_value(iters[len(iters)-1], 0)
                                                              ]))

        elif len(iters) == 2:
            self.builder.get_object('entry3').set_text(''.join([orig_title,
                                                                "\n",
                                                                model.get_value(iters[0], 0),
                                                                ' and ',
                                                                model.get_value(iters[1], 0),
                                                              ]))

        elif len(iters) == 1:
            self.builder.get_object('entry3').set_text(''.join([orig_title,
                                                                "\n",
                                                                model.get_value(iters[0], 0),
                                                              ]))


    def select_all_families(widget, treeview):
        """Select all families in the selection."""
        treeview.get_selection().select_all()

    def unselect_image(widget, filechoooserbutton):
        """Clear the cover image file selection."""
        filechoooserbutton.unselect_all()

    def open_dataset(self, widget):
        """Open a data file."""
        self.builder.get_object('notebook1').set_sensitive(False)
        self.dataset = Dataset(self, widget.get_filename())

        try:

            if self.dataset.data_source.read() == True:

                self.builder.get_object('notebook1').set_sensitive(True)
                
                config_files = []
                
                #add the default config file to the list
                config_files.append(''.join([os.path.splitext(self.dataset.filename)[0], '.cfg']))
                
                #search the path for additional config files
                for filename in glob.glob(''.join([os.path.splitext(self.dataset.filename)[0], '-*.cfg'])):
                    config_files.append(filename)
                
                #if we have more than one config file
                if len(config_files) > 1:
                    builder = gtk.Builder()
                    builder.add_from_file('./gui/select_dialog.glade')
                    builder.get_object('label68').set_text('Configuration file:')
                    builder.get_object('dialog').set_title('Select configuration file')
                    builder.get_object('button1').hide()
                    dialog = builder.get_object('dialog')

                    combobox = gtk.combo_box_new_text()

                    for filename in config_files:
                        combobox.append_text(os.path.basename(filename))

                    combobox.set_active(0)
                    combobox.show()
                    builder.get_object('hbox5').add(combobox)

                    response = dialog.run()

                    if response == 1:
                        config_file = '/'.join([os.path.dirname(self.dataset.filename), combobox.get_active_text()])
                        print config_file
                    else:
                        dialog.destroy()
                        return -1

                    dialog.destroy()
                else:
                    config_file = ''.join([os.path.splitext(self.dataset.filename)[0], '.cfg'])          

                self.dataset.config.read([config_file])
                self.dataset.config.filename = config_file
                
                ##set up the atlas gui based on config settings
                
                #title
                self.builder.get_object('entry3').set_text(self.dataset.config.get('Atlas', 'title'))
                
                #author
                self.builder.get_object('entry2').set_text(self.dataset.config.get('Atlas', 'author'))

                #cover image
                if self.dataset.config.get('Atlas', 'cover_image') == '':
                    self.unselect_image(self.builder.get_object('filechooserbutton1'))
                else:
                    self.builder.get_object('filechooserbutton1').set_filename(self.dataset.config.get('Atlas', 'cover_image'))
            
                #inside cover
                self.builder.get_object('textview1').get_buffer().set_text(self.dataset.config.get('Atlas', 'inside_cover'))
                
                #introduction
                self.builder.get_object('textview3').get_buffer().set_text(self.dataset.config.get('Atlas', 'introduction'))

                #distribution unit                                               
                self.builder.get_object('combobox3').set_active(grid_resolution.index(self.dataset.config.get('Atlas', 'distribution_unit')))
                
                #grid line style
                self.builder.get_object('combobox1').set_active(grid_resolution.index(self.dataset.config.get('Atlas', 'grid_lines_style')))
                
                #grid line colour
                self.builder.get_object('colorbutton1').set_color(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'grid_lines_colour')))
                
                #grid line visible
                self.builder.get_object('checkbutton2').set_active(self.dataset.config.getboolean('Atlas', 'grid_lines_visible'))
    
                #families
                store = gtk.ListStore(str)
                self.builder.get_object('treeview2').set_model(store)
                selection = self.builder.get_object('treeview2').get_selection()

                selection.unselect_all()
                self.builder.get_object('treeview2').scroll_to_point(0,0)
                for family in self.dataset.families:
                    iter = store.append([family])

                    if family.strip().lower() in ''.join(self.dataset.config.get('Atlas', 'families').split()).lower().split(','):
                        selection.select_path(store.get_path((iter)))

                model, selected = selection.get_selected_rows()
                try:
                    self.builder.get_object('treeview2').scroll_to_cell(selected[0])
                except IndexError:
                    pass
                    
                #vcs
                selection = self.builder.get_object('treeview1').get_selection()

                selection.unselect_all()
                self.builder.get_object('treeview1').scroll_to_point(0,0)
                try:
                    for vc in self.dataset.config.get('Atlas', 'vice-counties').split(','):
                        selection.select_path(int(float(vc))-1)
                    self.builder.get_object('treeview1').scroll_to_cell(int(float(self.dataset.config.get('Atlas', 'vice-counties').split(',')[0]))-1)
                except ValueError:
                    pass

                #paper size
                self.builder.get_object('combobox10').set_active(paper_size.index(self.dataset.config.get('Atlas', 'paper_size')))

                #paper orientation
                self.builder.get_object('combobox8').set_active(paper_orientation.index(self.dataset.config.get('Atlas', 'orientation')))
                
                #vice county colour
                self.builder.get_object('colorbutton5').set_color(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'vice-counties_colour')))
                
                #coverage style
                self.builder.get_object('combobox6').set_active(markers.index(self.dataset.config.get('Atlas', 'coverage_style')))
                
                #coverage colour
                self.builder.get_object('colorbutton4').set_color(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'coverage_colour')))

                #coverage visible
                self.builder.get_object('checkbutton1').set_active(self.dataset.config.getboolean('Atlas', 'coverage_visible'))                
                                
                #date band 1 style
                self.builder.get_object('combobox2').set_active(markers.index(self.dataset.config.get('Atlas', 'date_band_1_style')))
                
                #date band 1 from
                self.builder.get_object('spinbutton3').set_value(self.dataset.config.getfloat('Atlas', 'date_band_1_from'))
                
                #date band 1 to
                self.builder.get_object('spinbutton4').set_value(self.dataset.config.getfloat('Atlas', 'date_band_1_to'))       
                
                #date band 1 fill
                self.builder.get_object('colorbutton2').set_color(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_1_fill')))
                
                #date band 1 outline
                self.builder.get_object('colorbutton7').set_color(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_1_outline')))    

                #date band 2 visible
                self.builder.get_object('checkbutton4').set_active(self.dataset.config.getboolean('Atlas', 'date_band_2_visible'))
                                
                #date band 2 style
                self.builder.get_object('combobox4').set_active(markers.index(self.dataset.config.get('Atlas', 'date_band_2_style')))
                
                #date band 2 from
                self.builder.get_object('spinbutton1').set_value(self.dataset.config.getfloat('Atlas', 'date_band_2_from'))
                
                #date band 2 to
                self.builder.get_object('spinbutton2').set_value(self.dataset.config.getfloat('Atlas', 'date_band_2_to'))
                
                #date band 2 fill
                self.builder.get_object('colorbutton3').set_color(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_2_fill')))
                
                #date band 2 outline
                self.builder.get_object('colorbutton8').set_color(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_2_outline')))

                #date band 3 visible
                self.builder.get_object('checkbutton5').set_active(self.dataset.config.getboolean('Atlas', 'date_band_3_visible'))
                                
                #date band 3 style
                self.builder.get_object('combobox7').set_active(markers.index(self.dataset.config.get('Atlas', 'date_band_3_style')))
                
                #date band 3 from
                self.builder.get_object('spinbutton5').set_value(self.dataset.config.getfloat('Atlas', 'date_band_3_from'))
                
                #date band 3 to
                self.builder.get_object('spinbutton6').set_value(self.dataset.config.getfloat('Atlas', 'date_band_3_to'))
                
                #date band 3 fill
                self.builder.get_object('colorbutton6').set_color(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_3_fill')))
                
                #date band 3 outline
                self.builder.get_object('colorbutton9').set_color(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_3_outline')))
                                
                
                ##set up the list gui based on config settings
                
                #title
                self.builder.get_object('entry4').set_text(self.dataset.config.get('List', 'title'))
                
                #author
                self.builder.get_object('entry1').set_text(self.dataset.config.get('List', 'author'))

                #cover image
                if self.dataset.config.get('Atlas', 'cover_image') == '':
                    self.unselect_image(self.builder.get_object('filechooserbutton5'))
                else:
                    self.builder.get_object('filechooserbutton5').set_filename(self.dataset.config.get('List', 'cover_image'))
            
                #inside cover
                self.builder.get_object('textview4').get_buffer().set_text(self.dataset.config.get('List', 'inside_cover'))
                
                #introduction
                self.builder.get_object('textview6').get_buffer().set_text(self.dataset.config.get('List', 'introduction'))

                #distribution unit                                               
                self.builder.get_object('combobox5').set_active(grid_resolution.index(self.dataset.config.get('List', 'distribution_unit')))

                #families
                store = gtk.ListStore(str)
                self.builder.get_object('treeview3').set_model(store)
                selection = self.builder.get_object('treeview3').get_selection()

                selection.unselect_all()
                self.builder.get_object('treeview3').scroll_to_point(0,0)
                for family in self.dataset.families:
                    iter = store.append([family])

                    if family.strip().lower() in ''.join(self.dataset.config.get('List', 'families').split()).lower().split(','):
                        selection.select_path(store.get_path((iter)))

                model, selected = selection.get_selected_rows()
                try:
                    self.builder.get_object('treeview3').scroll_to_cell(selected[0])
                except IndexError:
                    pass
                    
                #vcs
                selection = self.builder.get_object('treeview4').get_selection()

                selection.unselect_all()
                self.builder.get_object('treeview4').scroll_to_point(0,0)
                try:
                    for vc in self.dataset.config.get('List', 'vice-counties').split(','):
                        selection.select_path(int(float(vc))-1)
                    self.builder.get_object('treeview4').scroll_to_cell(int(float(self.dataset.config.get('List', 'vice-counties').split(',')[0]))-1)
                except ValueError:
                    pass

                #paper size
                self.builder.get_object('combobox11').set_active(paper_size.index(self.dataset.config.get('List', 'paper_size')))

                #paper orientation
                self.builder.get_object('combobox9').set_active(paper_orientation.index(self.dataset.config.get('List', 'orientation')))

                self.builder.get_object('button3').set_sensitive(True)
            else:
                raise AttributeError()
        except AttributeError as e:
            self.builder.get_object('button3').set_sensitive(False)
            md = gtk.MessageDialog(None,
                gtk.DIALOG_DESTROY_WITH_PARENT, gtk.MESSAGE_ERROR,
                gtk.BUTTONS_CLOSE, ''.join(['Unable to open data file: ', str(e)]))
            md.run()
            md.destroy()


    def quit(self, widget, third=None):
        """Quit."""
        gtk.main_quit()
        sys.exit()

    def generate(self, widget):
        """Process the data file and configuration."""
        if self.dataset is not None:
            vbox = widget.get_parent().get_parent()
            notebook = vbox.get_children()[1]
            vbox.set_sensitive(False)

            dialog = gtk.FileChooserDialog('Save As...',
                                           None,
                                           gtk.FILE_CHOOSER_ACTION_SAVE,
                                           (gtk.STOCK_CANCEL, gtk.RESPONSE_CANCEL,
                                            gtk.STOCK_OPEN, gtk.RESPONSE_OK))
            dialog.set_default_response(gtk.RESPONSE_OK)

            if notebook.get_current_page() == 0:
                dialog.set_current_folder(os.path.dirname(os.path.abspath(self.dataset.filename)))
                dialog.set_current_name(''.join([os.path.splitext(self.dataset.config.filename)[0], '_atlas.pdf']))
            elif notebook.get_current_page() == 1:
                dialog.set_current_folder(os.path.dirname(os.path.abspath(self.dataset.filename)))
                dialog.set_current_name(''.join([os.path.splitext(self.dataset.config.filename)[0], '_checklist.pdf']))

            response = dialog.run()

            output = dialog.get_filename()

            dialog.destroy()

            watch = gtk.gdk.Cursor(gtk.gdk.WATCH)
            self.builder.get_object('dialog1').window.set_cursor(watch)

            if response == gtk.RESPONSE_OK:
                config = self.update_config()

                #add the extension if it's missing
                if output[-4:] != '.pdf':
                    output = ''.join([output, '.pdf'])

                #do the atlas
                if notebook.get_current_page() == 0:

                    atlas = Atlas(self.dataset)
                    atlas.save_in = output

                    ### convert these to config data

                    atlas.set_vcs(self.builder.get_object('treeview1'))
                    atlas.set_families(self.builder.get_object('treeview2'))
                    
                    temp_dir = tempfile.mkdtemp()
        
                    atlas.generate_base_map(temp_dir)
                    atlas.generate_density_map(temp_dir)
                    atlas.generate(temp_dir)

                elif notebook.get_current_page() == 1:

                    listing = List(self.dataset)

                    ### convert these to config data
                    listing.save_in = output
                    listing.set_vcs(self.builder.get_object('treeview4'))
                    listing.set_families(self.builder.get_object('treeview3'))
                    listing.generate()

            vbox.set_sensitive(True)
            self.builder.get_object('dialog1').window.set_cursor(None)

    def update_config(self):
        
        #atlas
        self.dataset.config.set('Atlas', 'title', self.builder.get_object('entry3').get_text())
        self.dataset.config.set('Atlas', 'author', self.builder.get_object('entry2').get_text())

        try:
            self.dataset.config.set('Atlas', 'cover_image', self.builder.get_object('filechooserbutton1').get_filename())
        except TypeError:
            self.dataset.config.set('Atlas', 'cover_image', '')

        buffer = self.builder.get_object('textview1').get_buffer()
        startiter, enditer = buffer.get_bounds()
        inside_cover = buffer.get_text(startiter, enditer, True)
        self.dataset.config.set('Atlas', 'inside_cover', inside_cover)

        buffer = self.builder.get_object('textview3').get_buffer()
        startiter, enditer = buffer.get_bounds()
        introduction = buffer.get_text(startiter, enditer, True)

        self.dataset.config.set('Atlas', 'introduction', introduction)
        self.dataset.config.set('Atlas', 'distribution_unit', self.builder.get_object('combobox3').get_active_text())

        #grab a comma delimited list of families
        selection = self.builder.get_object('treeview2').get_selection()
        model, selected = selection.get_selected_rows()
        iters = [model.get_iter(path) for path in selected]

        families = ''

        for iter in iters:
            families = ','.join([families, model.get_value(iter, 0)])

        self.dataset.config.set('Atlas', 'families', families[1:])

        #grab a comma delimited list of vcs
        selection = self.builder.get_object('treeview1').get_selection()
        model, selected = selection.get_selected_rows()
        iters = [model.get_iter(path) for path in selected]

        vcs = ''

        for iter in iters:
            vcs = ','.join([vcs, model.get_value(iter, 0)])

        self.dataset.config.set('Atlas', 'vice-counties', vcs[1:])
        self.dataset.config.set('Atlas', 'vice-counties_colour', str(self.builder.get_object('colorbutton5').get_color()))

        #date band 1
        self.dataset.config.set('Atlas', 'date_band_1_style', self.builder.get_object('combobox2').get_active_text())
        self.dataset.config.set('Atlas', 'date_band_1_fill', str(self.builder.get_object('colorbutton2').get_color()))
        self.dataset.config.set('Atlas', 'date_band_1_outline', str(self.builder.get_object('colorbutton7').get_color()))
        self.dataset.config.set('Atlas', 'date_band_1_from', str(self.builder.get_object('spinbutton3').get_value()))
        self.dataset.config.set('Atlas', 'date_band_1_to', str(self.builder.get_object('spinbutton4').get_value()))

        #date band 2
        self.dataset.config.set('Atlas', 'date_band_2_visible', str(self.builder.get_object('checkbutton4').get_active()))
        self.dataset.config.set('Atlas', 'date_band_2_style', self.builder.get_object('combobox4').get_active_text())
        self.dataset.config.set('Atlas', 'date_band_2_overlay', str(self.builder.get_object('checkbutton7').get_active()))
        self.dataset.config.set('Atlas', 'date_band_2_fill', str(self.builder.get_object('colorbutton3').get_color()))
        self.dataset.config.set('Atlas', 'date_band_2_outline', str(self.builder.get_object('colorbutton8').get_color()))
        self.dataset.config.set('Atlas', 'date_band_2_from', str(self.builder.get_object('spinbutton1').get_value()))
        self.dataset.config.set('Atlas', 'date_band_2_to', str(self.builder.get_object('spinbutton2').get_value()))

        #date band 3
        self.dataset.config.set('Atlas', 'date_band_3_visible', str(self.builder.get_object('checkbutton5').get_active()))
        self.dataset.config.set('Atlas', 'date_band_3_style', self.builder.get_object('combobox7').get_active_text())
        self.dataset.config.set('Atlas', 'date_band_3_overlay', str(self.builder.get_object('checkbutton11').get_active()))
        self.dataset.config.set('Atlas', 'date_band_3_fill', str(self.builder.get_object('colorbutton6').get_color()))
        self.dataset.config.set('Atlas', 'date_band_3_outline', str(self.builder.get_object('colorbutton9').get_color()))
        self.dataset.config.set('Atlas', 'date_band_3_from', str(self.builder.get_object('spinbutton5').get_value()))
        self.dataset.config.set('Atlas', 'date_band_3_to', str(self.builder.get_object('spinbutton6').get_value()))

        #coverage
        self.dataset.config.set('Atlas', 'coverage_visible', str(self.builder.get_object('checkbutton1').get_active()))
        self.dataset.config.set('Atlas', 'coverage_style', self.builder.get_object('combobox6').get_active_text())
        self.dataset.config.set('Atlas', 'coverage_colour', str(self.builder.get_object('colorbutton4').get_color()))

        #grid lines
        self.dataset.config.set('Atlas', 'grid_lines_visible', str(self.builder.get_object('checkbutton2').get_active()))
        self.dataset.config.set('Atlas', 'grid_lines_style', self.builder.get_object('combobox1').get_active_text())
        self.dataset.config.set('Atlas', 'grid_lines_colour', str(self.builder.get_object('colorbutton1').get_color()))

        #page setup
        self.dataset.config.set('Atlas', 'paper_size', self.builder.get_object('combobox10').get_active_text())
        self.dataset.config.set('Atlas', 'orientation', self.builder.get_object('combobox8').get_active_text())

        #table of contents
        self.dataset.config.set('Atlas', 'toc_show_families', str(self.builder.get_object('checkbutton6').get_active()))
        self.dataset.config.set('Atlas', 'toc_show_species_names', str(self.builder.get_object('checkbutton9').get_active()))
        self.dataset.config.set('Atlas', 'toc_show_common_names', str(self.builder.get_object('checkbutton10').get_active()))

        #species accounts
        self.dataset.config.set('Atlas', 'species_accounts_show_descriptions', str(self.builder.get_object('checkbutton12').get_active()))
        self.dataset.config.set('Atlas', 'species_accounts_show_latest', str(self.builder.get_object('checkbutton13').get_active()))
        self.dataset.config.set('Atlas', 'species_accounts_show_statistics', str(self.builder.get_object('checkbutton14').get_active()))
        self.dataset.config.set('Atlas', 'species_accounts_show_status', str(self.builder.get_object('checkbutton16').get_active()))
        self.dataset.config.set('Atlas', 'species_accounts_show_phenology', str(self.builder.get_object('checkbutton15').get_active()))


        #list
        self.dataset.config.set('List', 'title', self.builder.get_object('entry4').get_text())
        self.dataset.config.set('List', 'author', self.builder.get_object('entry1').get_text())

        try:
            self.dataset.config.set('List', 'cover_image', self.builder.get_object('filechooserbutton5').get_filename())
        except TypeError:
            self.dataset.config.set('List', 'cover_image', '')

        buffer = self.builder.get_object('textview4').get_buffer()
        startiter, enditer = buffer.get_bounds()
        inside_cover = buffer.get_text(startiter, enditer, True)
        self.dataset.config.set('List', 'inside_cover', inside_cover)

        buffer = self.builder.get_object('textview6').get_buffer()
        startiter, enditer = buffer.get_bounds()
        introduction = buffer.get_text(startiter, enditer, True)

        self.dataset.config.set('List', 'introduction', introduction)
        self.dataset.config.set('List', 'distribution_unit', self.builder.get_object('combobox3').get_active_text())

        #grab a comma delimited list of families
        selection = self.builder.get_object('treeview3').get_selection()
        model, selected = selection.get_selected_rows()
        iters = [model.get_iter(path) for path in selected]

        families = ''

        for iter in iters:
            families = ','.join([families, model.get_value(iter, 0)])

        self.dataset.config.set('List', 'families', families[1:])

        #grab a comma delimited list of vcs
        selection = self.builder.get_object('treeview4').get_selection()
        model, selected = selection.get_selected_rows()
        iters = [model.get_iter(path) for path in selected]

        vcs = ''

        for iter in iters:
            vcs = ','.join([vcs, model.get_value(iter, 0)])

        self.dataset.config.set('List', 'vice-counties', vcs[1:])

        #page setup
        self.dataset.config.set('List', 'paper_size', self.builder.get_object('combobox11').get_active_text())
        self.dataset.config.set('List', 'orientation', self.builder.get_object('combobox9').get_active_text())

        #write the config file
        with open(self.dataset.config.filename, 'wb') as configfile:
            self.dataset.config.write(configfile)


class Dataset(gobject.GObject):

    def __init__(self, instance, filename):
        gobject.GObject.__init__(self)

        self.instance = instance
        self.filename = filename
        self.mime = None

        self.records = None
        self.taxa = None
        self.families = []

        self.listing = { }

        self.connection = None
        self.cursor = None

        self.sql_filters = []

        self.chart = None

        self.atlas_config = {}
        self.list_config = {}

        self.cancel_reading = False

        if self.connection is None:
            self.connection = sqlite3.connect(':memory:')
            self.cursor = self.connection.cursor()

            #create the table to store the records
            self.cursor.execute('CREATE TABLE data \
                                 (taxon TEXT, \
                                  location TEXT, \
                                  grid_native TEXT, \
                                  grid_100km TEXT, \
                                  grid_10km TEXT, \
                                  grid_5km TEXT, \
                                  grid_2km TEXT, \
                                  grid_1km TEXT, \
                                  grid_100m TEXT, \
                                  grid_10m TEXT, \
                                  grid_1m TEXT, \
                                  easting NUMERIC, \
                                  northing NUMERIC, \
                                  accuracy NUMERIC, \
                                  date TEXT, \
                                  decade NUMERIC, \
                                  year NUMERIC, \
                                  month NUMERIC, \
                                  day NUMERIC, \
                                  decade_from NUMERIC, \
                                  year_from NUMERIC, \
                                  month_from NUMERIC, \
                                  day_from NUMERIC, \
                                  decade_to NUMERIC, \
                                  year_to NUMERIC, \
                                  month_to NUMERIC, \
                                  day_to NUMERIC, \
                                  recorder TEXT, \
                                  determiner TEXT, \
                                  vc NUMERIC)')

            self.cursor.execute('CREATE INDEX index_taxon \
                                 ON data (taxon)')

            #create the table to store the species data
            self.cursor.execute('CREATE TABLE species_data \
                                 (taxon TEXT, \
                                  family TEXT, \
                                  sort_order NUMERIC, \
                                  nbn_key TEXT, \
                                  national_status TEXT, \
                                  local_status TEXT, \
                                  description TEXT, \
                                  common_name TEXT)')

            self.cursor.execute('CREATE INDEX index_sp_taxon \
                                 ON species_data (taxon)')


            #initiate config with defaults
            self.config = ConfigParser.ConfigParser({'title': '',
                                                             'author': '',
                                                             'cover_image': '',
                                                             'inside_cover': '',
                                                             'introduction': '',
                                                             'distribution_unit': '2km',
                                                             'families': '',
                                                             'vice-counties': '',
                                                             'vice-counties_colour': '#000',
                                                             'date_band_1_style': 'squares',
                                                             'date_band_1_fill': '#000',
                                                             'date_band_1_outline': '#000',
                                                             'date_band_1_from': '1600.0',
                                                             'date_band_1_to': '1980.0',
                                                             'date_band_2_visible': 'True',
                                                             'date_band_2_style': 'squares',
                                                             'date_band_2_overlay': 'False',
                                                             'date_band_2_fill': '#000',
                                                             'date_band_2_outline': '#000',
                                                             'date_band_2_from': '1980.0',
                                                             'date_band_2_to': '2050.0',
                                                             'date_band_3_visible': 'False',
                                                             'date_band_3_style': 'squares',
                                                             'date_band_3_overlay': 'False',
                                                             'date_band_3_fill': '#000',
                                                             'date_band_3_outline': '#000',
                                                             'date_band_3_from': '0.0',
                                                             'date_band_3_to': '0.0',
                                                             'coverage_visible': 'True',
                                                             'coverage_style': 'squares',
                                                             'coverage_colour': '#d2d2d2',
                                                             'grid_lines_visible': 'True',
                                                             'grid_lines_style': '2km',
                                                             'grid_lines_colour': '#d2d2d2',
                                                             'paper_size': 'A4',
                                                             'orientation': 'Portrait',
                                                             'toc_show_families': 'True',
                                                             'toc_show_species_names': 'False',
                                                             'toc_show_common_names': 'False',
                                                             'species_accounts_show_descriptions': 'True',
                                                             'species_accounts_show_latest': 'True',
                                                             'species_accounts_show_statistics': 'True',
                                                             'species_accounts_show_status': 'True',
                                                             'species_accounts_show_phenology': 'True',
                                                            })
                                                
            self.config.add_section('Atlas')
            self.config.add_section('List')

        #guess the mimetype of the file
        self.mime = mimetypes.guess_type(self.filename)[0]

        if self.mime == 'application/vnd.ms-excel':
            self.data_source = Read(self.filename, self)
        else:
            tempfile = 'dipper-stand-convert.xls'

            try:
                returncode = call(["ssconvert", self.filename, tempfile])

                if returncode == 0:
                    self.data_source = Read(tempfile, self)
            except OSError:
                pass


    def close(self):
        self.connection = None
        self.cursor = None

class Read(gobject.GObject):
    __gsignals__ = {
      'progress-pre-begin': (gobject.SIGNAL_RUN_FIRST, gobject.TYPE_NONE, ()),
      'progress-pre-end': (gobject.SIGNAL_RUN_FIRST, gobject.TYPE_NONE, ()),
      'progress-begin': (gobject.SIGNAL_RUN_FIRST, gobject.TYPE_NONE, (str, str,)),
      'progress-end': (gobject.SIGNAL_RUN_FIRST, gobject.TYPE_NONE, ()),
      'progress-cancelled': (gobject.SIGNAL_RUN_FIRST, gobject.TYPE_NONE, ()),
      'progress-update': (gobject.SIGNAL_RUN_FIRST, gobject.TYPE_NONE, (float,)),
    }

    def __init__(self, filename, dataset):
        gobject.GObject.__init__(self)

        self.filename = filename
        self.dataset = dataset
        self.cancel_reading = False


    def read(self):
        '''Read the file and insert the data into the sqlite database.'''

        self.emit('progress-pre-begin')
        book = xlrd.open_workbook(self.filename)
        self.emit('progress-pre-end')

        #if more than one sheet is in the workbook, display sheet selection
        #dialog
        has_data = False
        has_config = False
        ignore_sheets = 0
        for name in book.sheet_names():
            if name[:2] == '--' and name [-2:] == '--':
                ignore_sheets = ignore_sheets + 1

            # do we have a data sheet?
            if name == '--data--':
                has_data = True

            # do we have a config sheet?
            if name == '--config--':
                has_config = True

        if (book.nsheets)-ignore_sheets > 1:
            builder = gtk.Builder()
            builder.add_from_file('./gui/select_dialog.glade')
            builder.get_object('label68').set_text('Sheet:')
            builder.get_object('dialog').set_title('Select sheet')
            dialog = builder.get_object('dialog')

            combobox = gtk.combo_box_new_text()

            for name in book.sheet_names():
                if not name[:2] == '--' and not name [-2:] == '--':
                    combobox.append_text(name)

            combobox.set_active(0)
            combobox.show()
            builder.get_object('hbox5').add(combobox)

            response = dialog.run()

            if response == 1:
                sheets = (book.sheet_by_name(combobox.get_active_text()),)
            elif response == 2:
                sheets = []
                for name in book.sheet_names():
                    if not name[:2] == '--' and not name [-2:] == '--':
                        sheets.append(book.sheet_by_name(name))
            else:
                dialog.destroy()
                return -1

            dialog.destroy()

        else:
            sheets = (book.sheet_by_index(0),)

        text = ''.join(['Opening <b>', os.path.basename(self.filename) ,'</b>', ' from ', '<b>', os.path.dirname(os.path.abspath(self.filename)), '</b>'])
        self.emit('progress-begin', text, 'open')

        vc_position = False

        temp_taxa_list = []


        try:
            #loop through the selected sheets of the workbook
            for sheet in sheets:
                # try and match up the column headings
                for col_index in range(sheet.ncols):
                    ######## what to do if these headings _dont_ exist?
                    if sheet.cell(0, col_index).value.lower() in ('taxon name', 'taxon', 'recommended taxon name'):
                        taxon_position = col_index
                    elif sheet.cell(0, col_index).value.lower() in ('grid reference', 'grid ref', 'grid ref.', 'gridref', 'sample spatial reference'):
                        grid_reference_position = col_index
                    elif sheet.cell(0, col_index).value.lower() in ('date', 'sample date'):
                        date_position = col_index
                    elif sheet.cell(0, col_index).value.lower() in ('location', 'site'):
                        location_position = col_index
                    elif sheet.cell(0, col_index).value.lower() in ('recorder', 'recorders'):
                        recorder_position = col_index
                    elif sheet.cell(0, col_index).value.lower() in ('determiner', 'determiners'):
                        determiner_position = col_index
                    elif sheet.cell(0, col_index).value.lower() in ('vc', 'vice-county', 'vice county'):
                        vc_position = col_index

                rownum = 0

                #loop through each row, skipping the header (first) row
                for row_index in range(1, sheet.nrows):
                    taxa = sheet.cell(row_index, taxon_position).value
                    location = sheet.cell(row_index, location_position).value
                    grid_reference = sheet.cell(row_index, grid_reference_position).value
                    date = sheet.cell(row_index, date_position).value
                    recorder = sheet.cell(row_index, recorder_position).value
                    determiner = sheet.cell(row_index, determiner_position).value

                    vaguedate = VagueDate(date)
                    decade, year, month, day, decade_from, year_from, month_from, day_from, decade_to, year_to, month_to, day_to = vaguedate.decade, vaguedate.year, vaguedate.month,  vaguedate.day, vaguedate.decade_from, vaguedate.year_from, vaguedate.month_from,  vaguedate.day_from, vaguedate.decade_to, vaguedate.year_to, vaguedate.month_to,  vaguedate.day_to

                    reference = Coordinate(str(grid_reference))

                    easting, northing = reference.e, reference.n
                    grid_100km = reference.os_100km
                    grid_10km = reference.os_10km
                    grid_5km = reference.os_5km
                    grid_2km = reference.os_2km
                    grid_1km = reference.os_1km
                    grid_100m = reference.os_100m
                    grid_10m = reference.os_10m
                    grid_1m = reference.os_1m


                    if vc_position:
                      vc = sheet.cell(row_index, vc_position).value
                    else:
                      vc = None

                    accuracy = reference.accuracy

                    if len(taxa) > 0:
                        temp_taxa_list.append(taxa)
                        self.dataset.cursor.execute('INSERT INTO data \
                                                     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)',
                                                     [taxa,
                                                     location,
                                                     grid_reference,
                                                     grid_100km, grid_10km, grid_5km, grid_2km,
                                                     grid_1km, grid_100m, grid_10m, grid_1m,
                                                     easting, northing, accuracy, date,
                                                     decade, year, month, day,
                                                     decade_from, year_from, month_from, day_from,
                                                     decade_to, year_to, month_to, day_to,
                                                     recorder,
                                                     determiner,
                                                     vc])

                    #emit a progress update every 100 rows
                    if rownum%100 == 0:
                        self.emit('progress-update', float(rownum)/sheet.nrows)

                    rownum = rownum + 1

                    #escape hatch
                    if self.cancel_reading is True:
                        self.cancel_reading = False
                        self.emit('progress-cancelled')
                        return False

            ########### we need to run the 'distinct species' SQL first, then loop through species data -
            ########### if species data sheet is only partially filled, any taxa not covered are ignored
            #######################################################################################

            #load the data sheet
            if has_data:
                sheet = book.sheet_by_name('--data--')

                # try and match up the column headings
                for col_index in range(sheet.ncols):

                    if sheet.cell(0, col_index).value.lower() in ('taxon', 'taxon name'):
                        taxon_position = col_index
                    elif sheet.cell(0, col_index).value.lower() in ('family'):
                        family_position = col_index
                    elif sheet.cell(0, col_index).value.lower() in ('sort order'):
                        sort_order_position = col_index
                    elif sheet.cell(0, col_index).value.lower() in ('nbn key'):
                        nbn_key_position = col_index
                    elif sheet.cell(0, col_index).value.lower() in ('national status (short)', 'status'):
                        national_status_position = col_index
                    elif sheet.cell(0, col_index).value.lower() in ('description'):
                        description_position = col_index
                    elif sheet.cell(0, col_index).value.lower() in ('common name'):
                        common_name_position = col_index

                #loop through each row, skipping the header (first) row
                for row_index in range(1, sheet.nrows):

                    taxa = sheet.cell(row_index, taxon_position).value
                    if taxa in temp_taxa_list:
                        try:
                            family = sheet.cell(row_index, family_position).value
                        except UnboundLocalError:
                            family = None

                        try:
                            sort_order = sheet.cell(row_index, sort_order_position).value
                        except UnboundLocalError:
                            sort_order = None

                        try:
                            nbn_key = sheet.cell(row_index, nbn_key_position).value
                        except UnboundLocalError:
                            nbn_key = None

                        try:
                            national_status = sheet.cell(row_index, national_status_position).value
                        except UnboundLocalError:
                            national_status = None

                        try:
                            description = sheet.cell(row_index, description_position).value
                        except UnboundLocalError:
                            description = None

                        try:
                            common_name = sheet.cell(row_index, common_name_position).value
                        except UnboundLocalError:
                            common_name = None

                        if family not in self.dataset.families:
                            self.dataset.families.append(family)

                        self.dataset.cursor.execute('INSERT INTO species_data \
                                                     VALUES (?, ?, ?, ?, ?, ?, ?, ?)',
                                                     [taxa,
                                                     family,
                                                     sort_order,
                                                     nbn_key,
                                                     national_status,
                                                     None,
                                                     description,
                                                     common_name])

                statuss = {}

                ######### hack

                self.dataset.cursor.execute('SELECT COUNT(DISTINCT(grid_2km)) AS count \
                                       FROM data \
                                       GROUP BY data.taxon \
                                       ORDER BY count DESC')

                data = self.dataset.cursor.fetchall()

                total_coverage = float(data[0][0])

                #get any taxa that haven't been recorded in n years
                self.dataset.cursor.execute('SELECT data.taxon, MAX(data.year), species_data.sort_order \
                                       FROM data \
                                       JOIN species_data ON data.taxon = species_data.taxon \
                                       GROUP BY data.taxon, species_data.sort_order')

                data = self.dataset.cursor.fetchall()

                for taxon in data:
                    if taxon[1] <= datetime.now().year-31:## move this to the SQL above
                        statuss[taxon[0]] = '*****'

                #get the taxa that have been recorded in the last 30 years

                #####hack

                self.dataset.cursor.execute('SELECT data.taxon, COUNT(DISTINCT(grid_2km)), species_data.sort_order \
                                       FROM data \
                                       JOIN species_data ON data.taxon = species_data.taxon \
                                       WHERE data.year >= ' + str(datetime.now().year-30) + ' \
                                       GROUP BY data.taxon, species_data.sort_order')

                data = self.dataset.cursor.fetchall()

                for taxon in data:
                    percent = (float(taxon[1])/total_coverage)*100

                    if percent <= 1:
                        status = '*****'
                    elif percent > 1 and percent <= 5:
                        status = '****'
                    elif percent > 5 and percent <= 15:
                        status = '***'
                    elif percent > 15 and percent <= 30:
                        status = '**'
                    elif percent > 30:
                        status = '*'

                    statuss[taxon[0]] = status

                for species in statuss:

                    self.dataset.cursor.execute('UPDATE species_data \
                                               SET local_status = "' + statuss[species] + '" \
                                               WHERE taxon = "' + species + '"')






            else:
                #if we dont have species data sheet, just create the data using distinct taxa
                self.dataset.cursor.execute('SELECT DISTINCT(data.taxon) \
                                            FROM data \
                                            ORDER BY data.taxon')

                data = self.dataset.cursor.fetchall()

                for row in data:
                    self.dataset.cursor.execute('INSERT INTO species_data \
                                                VALUES (?, ?, ?, ?, ?, ?, ?, ?)',
                                                [row[0],
                                                None,
                                                None,
                                                None,
                                                None,
                                                None,
                                                None,
                                                None])

            self.emit('progress-end')

            return True
        except UnboundLocalError:
            return False

    def cancel_read(self):
        self.cancel_reading = True

class PDF(FPDF):
    def __init__(self, orientation,unit,format):
        FPDF.__init__(self, orientation=orientation,unit=unit,format=format)
        self.toc = []
        self.numbering = False
        self.num_page_num = 0
        self.toc_page_break_count = 1
        self.set_left_margin(10)
        self.set_right_margin(10)
        self.do_header = False
        self.type = None
        self.toc_length = 0
        self.doing_the_list = False
        self.vcs = []
        self.toc_page_num = 2

    def p_add_page(self):
        #if(self.numbering):
        self.add_page()
        self.num_page_num = self.num_page_num + 1


    def num_page_no(self):
        return self.num_page_num


    def startPageNums(self):
        self.numbering = True

    def stopPageNums(self):
        self.numbering = False

    def TOC_Entry(self, txt, level=0):
        self.toc.append({'t':txt, 'l':level, 'p':str(self.num_page_no()+self.toc_length)})

    def insertTOC(self, location=1, labelSize=20, entrySize=10, tocfont='Helvetica', label='Table of Contents'):
        #make toc at end
        self.stopPageNums()
        self.section = 'Contents'
        self.p_add_page()
        tocstart = self.page

        self.set_font('Helvetica', '', 20)
        self.multi_cell(0, 20, 'Contents', 0, 'J', False)
        #self.set_font(tocfont, 'B', labelSize)
        #self.cell(0, 5, label, 0, 1, 'C')
        #self.ln(10)

        for t in self.toc:

            #Offset
            level = t['l']

            if level > 0:
                self.cell(level*8)

            weight = ''

            if level == 0:
                weight = 'B'

            txxt = t['t']
            self.set_font(tocfont, weight, entrySize)
            strsize = self.get_string_width(txxt)
            self.cell(strsize+2, self.font_size+2, txxt)

            #Filling dots
            self.set_font(tocfont, '', entrySize)
            PageCellSize = self.get_string_width(t['p'])+2
            w = self.w-self.l_margin-self.r_margin-PageCellSize-(level*8)-(strsize+2)
            nb = w/self.get_string_width('.')
            dots = repeat_to_length('.', int(nb))
            self.cell(w, self.font_size+2, dots, 0, 0, 'R')

            #Page number of the toc entry
            self.cell(PageCellSize, self.font_size+2, str(int(t['p'])), 0, 1, 'R')

        if self.toc_page_break_count%2 != 0:
            self.section = 'Contents'
            self.toc_page_break_count = self.toc_page_break_count + 1
            self.p_add_page()


        #Grab it and move to selected location
        n = self.page
        ntoc = n - tocstart + 1
        last = []


        #store toc pages
        i = tocstart
        while i <= n:
            last.append(self.pages[i])
            i = i + 1

        #move pages
        i = tocstart
        while i >= (location-1):
            self.pages[i+ntoc] = self.pages[i]
            i = i - 1

        #Put toc pages at insert point
        i = 0
        while i < ntoc:
            self.pages[location + i] = last[i]
            i = i + 1

    def header(self):
        if self.do_header:
            self.set_font('Helvetica', '', 8)
            self.set_text_color(0, 0, 0)
            self.set_line_width(0.1)

            if self.page_no()%2 == 0:
                self.cell(0, 5, self.section, 'B', 0, 'L', 0) # even page header
                self.cell(0, 5, self.title, 'B', 1, 'R', 0) # even page header

            else:
                self.cell(0, 5, self.section, 'B', 1, 'R', 0) #odd page header

            if self.type == 'list' and self.doing_the_list == True:

                col_width = 12.7#((self.w - self.l_margin - self.r_margin)/2)/7.5

                #vc headings
                self.set_font('Helvetica', '', 10)
                self.set_line_width(0.0)
                self.set_y(20)

                self.set_x(self.w-(7+col_width+(((col_width*3)+(col_width/4))*len(self.vcs))))

                self.cell(col_width, 5, '', '0', 0, 'C', 0)
                for vc in sorted(self.vcs):
                    self.cell((col_width*3), 5, ''.join(['VC',vc]), '0', 0, 'C', 0)
                    self.cell(col_width/4, 5, '', '0', 0, 'C', 0)

                self.ln()

                self.set_x(self.w-(7+col_width+(((col_width*3)+(col_width/4))*len(self.vcs))))
                self.set_font('Helvetica', '', 8)
                self.cell(col_width, 5, '', '0', 0, 'C', 0)

                for vc in sorted(self.vcs):
                    #colum headings
                    self.cell(col_width, 5, 'Tetrads', '0', 0, 'C', 0)
                    self.cell(col_width, 5, 'Records', '0', 0, 'C', 0)
                    self.cell(col_width, 5, 'Last in', '0', 0, 'C', 0)
                    self.cell(col_width/4, 5, '', '0', 0, 'C', 0)

                self.y0 = self.get_y()

            if self.section == 'Contributors':
                self.set_y(self.y0 + 20)

    def footer(self):
        self.set_y(-20)
        self.set_font('Helvetica','',8)

        #only show page numbers in the main body
        if self.num_page_no() >= 4 and self.section != 'Contents' and self.section != 'Index' and self.section != 'Contributors' and self.section != '':
            self.cell(0, 10, str(self.num_page_no()+self.toc_length), '', 0, 'C')


    def setcol(self, col):
        self.col = col
        x = 10 + (col*100)
        self.set_left_margin(x)
        self.set_x(x)

    def accept_page_break(self):

        if self.section == 'Contents':
            self.toc_page_break_count = self.toc_page_break_count + 1

        if self.section == 'Contributors':
            self.set_y(self.y0+20)

        if self.section == 'Index':
            if self.col < 1:
                self.setcol(self.col + 1)
                self.set_y(self.y0+20)
                return False
            else:
                self.setcol(0)
                self.p_add_page()
                self.set_y(self.y0+20)
                return False
        else:
            return True


class List(gobject.GObject):
    __gsignals__ = {
      'progress-pre-begin': (gobject.SIGNAL_RUN_FIRST, gobject.TYPE_NONE, ()),
      'progress-pre-end': (gobject.SIGNAL_RUN_FIRST, gobject.TYPE_NONE, ()),
      'progress-begin': (gobject.SIGNAL_RUN_FIRST, gobject.TYPE_NONE, (str, str,)),
      'progress-end': (gobject.SIGNAL_RUN_FIRST, gobject.TYPE_NONE, ()),
      'progress-cancelled': (gobject.SIGNAL_RUN_FIRST, gobject.TYPE_NONE, ()),
      'progress-update': (gobject.SIGNAL_RUN_FIRST, gobject.TYPE_NONE, (float,)),
    }

    def __init__(self, dataset):
        gobject.GObject.__init__(self)
        self.dataset = dataset
        self.cancel = False
        self.start_time = time.time()
        self.page_unit = 'mm'
        self.save_in = None
        self.vcs = []
        self.vcs_sql = ''
        self.families = ''

    def set_vcs(self, widget):

        selection = widget.get_selection()
        model, selected = selection.get_selected_rows()
        iters = [model.get_iter(path) for path in selected]

        sql = ''

        for iter in iters:
            sql = ','.join([sql, ''.join(["'", model.get_value(iter, 0), "'"])     ])
            self.vcs.append(model.get_value(iter, 0))

        if len(iters) > 0:
            self.vcs_sql = ''.join([' WHERE data.vc IN (', sql[1:], ') '])

    def set_families(self, widget):

        selection = widget.get_selection()
        model, selected = selection.get_selected_rows()
        iters = [model.get_iter(path) for path in selected]

        sql = ''

        for iter in iters:
            sql = ','.join([sql, ''.join(["'", model.get_value(iter, 0), "'"])     ])

        if self.vcs is not None:
            joiner = 'AND'
        else:
            joiner = 'WHERE'

        if len(iters) > 0:
            self.families = ''.join([joiner, ' species_data.family IN (', sql[1:], ') '])


    def generate(self):
        self.emit('progress-pre-begin')

        taxa_statistics = {}
        taxon_list = []

        self.dataset.cursor.execute('SELECT data.taxon, species_data.family, species_data.national_status, species_data.local_status, \
                                   COUNT(DISTINCT(grid_' + self.dataset.config.get('List', 'distribution_unit') + ')) AS tetrads, \
                                   COUNT(data.taxon) AS records, \
                                   MAX(data.year) AS year, \
                                   data.vc AS VC \
                                   FROM data \
                                   JOIN species_data ON data.taxon = species_data.taxon \
                                   ' + self.families + ' \
                                   ' + self.vcs_sql + ' \
                                   GROUP BY data.taxon, species_data.family, species_data.national_status, species_data.local_status, data.vc \
                                   ORDER BY species_data.sort_order')

        data = self.dataset.cursor.fetchall()
        numrecords = 0

        for row in data:
            numrecords = numrecords + int(row[5])

        for row in data:
            if row[0] not in taxon_list:
                taxon_list.append(row[0])

        for row in data:

            if row[0] in taxa_statistics:
                taxa_statistics[row[0]]['vc'][str(row[7])] = {}

                taxa_statistics[row[0]]['vc'][str(row[7])]['tetrads'] = str(row[4])
                taxa_statistics[row[0]]['vc'][str(row[7])]['records'] = str(row[5])
                taxa_statistics[row[0]]['vc'][str(row[7])]['year'] = str(row[6])

            else:
                taxa_statistics[row[0]] = {}
                taxa_statistics[row[0]]['vc'] = {}
                taxa_statistics[row[0]]['family'] = str(row[1])
                taxa_statistics[row[0]]['national_designation'] = str(row[2])
                taxa_statistics[row[0]]['local_designation'] = str(row[3])

                taxa_statistics[row[0]]['vc'][str(row[7])] = {}

                taxa_statistics[row[0]]['vc'][str(row[7])]['tetrads'] = str(row[4])
                taxa_statistics[row[0]]['vc'][str(row[7])]['records'] = str(row[5])
                taxa_statistics[row[0]]['vc'][str(row[7])]['year'] = str(row[6])

        #the pdf
        pdf = PDF(orientation=self.dataset.config.get('List', 'orientation'),unit=self.page_unit,format=self.dataset.config.get('List', 'paper_size'))
        pdf.type = 'list'
        pdf.do_header = False
        pdf.vcs = self.vcs

        pdf.col = 0
        pdf.y0 = 0
        pdf.set_title(self.dataset.config.get('List', 'title'))
        pdf.set_author(self.dataset.config.get('List', 'author'))
        pdf.section = ''

        #title page
        pdf.p_add_page()

        if self.dataset.config.get('List', 'cover_image') is not None and os.path.isfile(self.dataset.config.get('List', 'cover_image')):
            pdf.image(self.dataset.config.get('List', 'cover_image'), 0, 0, pdf.w, pdf.h)

        pdf.set_text_color(0)
        pdf.set_fill_color(255, 255, 255)
        pdf.set_font('Helvetica', '', 28)
        pdf.ln(15)
        pdf.multi_cell(0, 10, pdf.title, 0, 'L', False)

        pdf.ln(20)
        pdf.set_font('Helvetica', '', 18)
        pdf.multi_cell(0, 10, ''.join([pdf.author, '\n',datetime.now().strftime('%B %Y')]), 0, 'L', False)

        #inside cover
        pdf.p_add_page()
        pdf.set_font('Helvetica', '', 12)
        pdf.multi_cell(0, 6, self.dataset.config.get('List', 'inside_cover'), 0, 'J', False)

        #introduction
        pdf.do_header = True
        pdf.section = ('Introduction')
        pdf.p_add_page()
        pdf.startPageNums()
        pdf.set_font('Helvetica', '', 20)
        pdf.cell(0, 20, 'Introduction', 0, 0, 'L', 0)
        pdf.ln()
        pdf.set_font('Helvetica', '', 12)
        pdf.multi_cell(0, 6, self.dataset.config.get('List', 'introduction'), 0, 0, 'L')
        pdf.ln()
        pdf.set_font('Helvetica', '', 20)
        pdf.cell(0, 15, 'Checklist', 0, 1, 'L', 0)




        col_width = 12.7#((self.w - self.l_margin - self.r_margin)/2)/7.5

        #vc headings
        pdf.set_font('Helvetica', '', 10)
        pdf.set_line_width(0.0)

        pdf.set_x(pdf.w-(7+col_width+(((col_width*3)+(col_width/4))*len(self.vcs))))

        pdf.cell(col_width, 5, '', '0', 0, 'C', 0)
        for vc in sorted(self.vcs):
            pdf.cell((col_width*3), 5, ''.join(['VC',vc]), '0', 0, 'C', 0)
            pdf.cell(col_width/4, 5, '', '0', 0, 'C', 0)


        pdf.ln()

        pdf.set_x(pdf.w-(7+col_width+(((col_width*3)+(col_width/4))*len(self.vcs))))
        pdf.set_font('Helvetica', '', 8)
        pdf.cell(col_width, 5, '', '0', 0, 'C', 0)

        for vc in sorted(pdf.vcs):
            #colum headings
            pdf.cell(col_width, 5, 'Tetrads', '0', 0, 'C', 0)
            pdf.cell(col_width, 5, 'Records', '0', 0, 'C', 0)
            pdf.cell(col_width, 5, 'Last in', '0', 0, 'C', 0)
            pdf.cell(col_width/4, 5, '', '0', 0, 'C', 0)








        pdf.doing_the_list = True
        pdf.set_font('Helvetica', '', 8)


        for vckey in sorted(self.vcs):
            #print vckey

            col = self.vcs.index(vckey)+1

            pdf.cell(col_width/col, 5, '', '0', 0, 'C', 0)
            pdf.cell(col_width, 5, 'Tetrads', '0', 0, 'C', 0)
            pdf.cell(col_width, 5, 'Records', '0', 0, 'C', 0)
            pdf.cell(col_width, 5, 'Last in', '0', 0, 'C', 0)

        taxon_count = 0

        record_count = 0
        families = []
        for key in taxon_list:
            if taxa_statistics[key]['family'] not in families:
                pdf.section = (''.join(['Family ', taxa_statistics[key]['family'].upper()]))
                if pdf.get_y() > (pdf.h - 42):
                    pdf.p_add_page()
                pdf.set_font('Helvetica', 'B', 10)

                pdf.set_fill_color(50, 50, 50)
                pdf.set_text_color(255, 255, 255)

                pdf.ln()
                pdf.cell(0, 5, ''.join(['Family ', taxa_statistics[key]['family'].upper()]), 0, 1, 'L', 1)
                families.append(taxa_statistics[key]['family'])

                pdf.set_fill_color(255, 255, 255)
                pdf.set_text_color(0, 0, 0)

            elif pdf.get_y() > (pdf.h - 27):
                pdf.section = (''.join(['Family ', taxa_statistics[key]['family'].upper()]))
                pdf.p_add_page()
                pdf.set_y(30)

            #species name
            pdf.set_font('Helvetica', '', 10)
            strsize = pdf.get_string_width(key)+2
            pdf.cell(strsize, pdf.font_size+2, key, '', 0, 'L', 0)

            #dots
            ####to get dots to the tight length
            w = pdf.w-(4+col_width+col_width+(((col_width*3)+(col_width/4))*len(pdf.vcs))) - strsize
            nb = w/pdf.get_string_width('.')
            ####

            dots = repeat_to_length('.', int(nb))
            pdf.cell(w, pdf.font_size+2, dots, 0, 0, 'R', 0)

            pdf.set_font('Helvetica', '', 6)
            pdf.cell(col_width, pdf.font_size+3, taxa_statistics[key]['national_designation'], '', 0, 'L', 0)
            pdf.set_font('Helvetica', '', 10)

            for vckey in sorted(self.vcs):
                #print vckey

                pdf.set_fill_color(230, 230, 230)
                try:

                    if taxa_statistics[key]['vc'][vckey]['tetrads'] == '0':
                        tetrads = '-'
                    else:
                        tetrads = taxa_statistics[key]['vc'][vckey]['tetrads']

                    pdf.cell(col_width, pdf.font_size+2, tetrads, '', 0, 'L', 1)
                    pdf.cell(col_width, pdf.font_size+2, taxa_statistics[key]['vc'][vckey]['records'], '', 0, 'L', 1)
                    record_count = record_count + int(taxa_statistics[key]['vc'][vckey]['records'])

                    if taxa_statistics[key]['vc'][vckey]['year'] == 'None':
                        pdf.cell(col_width, pdf.font_size+2, '?', '', 0, 'L', 1)
                    else:
                        pdf.cell(col_width, pdf.font_size+2, taxa_statistics[key]['vc'][vckey]['year'], '', 0, 'C', 1)

                except KeyError:
                    pdf.cell(col_width, pdf.font_size+2, '', '', 0, 'L', 1)
                    pdf.cell(col_width, pdf.font_size+2, '', '', 0, 'L', 1)
                    pdf.cell(col_width, pdf.font_size+2, '', '', 0, 'C', 1)


                pdf.set_fill_color(255, 255, 255)

                pdf.cell((col_width/4), pdf.font_size+2, '', 0, 0, 'C', 0)


            pdf.ln()

            while gtk.events_pending():
                gtk.main_iteration()

            taxon_count = taxon_count + 1

            if self.cancel is True:
                self.cancel = False
                self.emit('progress-cancelled')
                return False

            self.emit('progress-update', float(taxon_count)/len(taxa_statistics))


        pdf.section = ''
        pdf.doing_the_list = False

        if len(families) > 1:
            family_text = ' '.join([str(len(families)), 'families and'])
        else:
            family_text = ''

        if len(data) > 1:
            taxa_text = 'taxa'
        else:
            taxa_text = 'taxon'

        if record_count > 1:
            record_text = 'records.'
        else:
            record_text = 'record.'

        pdf.set_y(pdf.get_y()+10)
        pdf.set_font('Helvetica', 'I', 10)
        pdf.multi_cell(0, 5, ' '.join([family_text,
                                       ' '.join([str(len(data)), taxa_text]),
                                       ' '.join(['listed from', str(record_count), record_text]),]),
                       0, 'J', False)


        #output
        pdf.output(self.save_in,'F')

        self.emit('progress-end')



class Atlas(gobject.GObject):
    __gsignals__ = {
      'progress-pre-begin': (gobject.SIGNAL_RUN_FIRST, gobject.TYPE_NONE, ()),
      'progress-pre-end': (gobject.SIGNAL_RUN_FIRST, gobject.TYPE_NONE, ()),
      'progress-begin': (gobject.SIGNAL_RUN_FIRST, gobject.TYPE_NONE, (str, str,)),
      'progress-end': (gobject.SIGNAL_RUN_FIRST, gobject.TYPE_NONE, ()),
      'progress-cancelled': (gobject.SIGNAL_RUN_FIRST, gobject.TYPE_NONE, ()),
      'progress-update': (gobject.SIGNAL_RUN_FIRST, gobject.TYPE_NONE, (float,)),
    }

    def __init__(self, dataset):
        gobject.GObject.__init__(self)
        self.dataset = dataset
        self.cancel = False
        self.start_time = time.time()
        self.save_in = None
        self.page_unit = 'mm'
        self.base_map = None
        self.date_band_1_fill_colour = None
        self.date_band_2_fill_colour = None
        self.date_band_3_fill_colour = None
        self.date_band_1_border_colour = None
        self.date_band_2_border_colour = None
        self.date_band_3_border_colour = None
        self.vcs_widget = None
        self.families = None
        self.date_band_1_style_coverage = []
        self.date_band_2_style_coverage = []
        self.date_band_3_style_coverage = []
        self.density_map_filename = None

    def set_vcs(self, widget):
        self.vcs_widget = widget

    def set_families(self, widget):

        selection = widget.get_selection()
        model, selected = selection.get_selected_rows()
        iters = [model.get_iter(path) for path in selected]

        sql = ''

        for iter in iters:
            sql = ','.join([sql, ''.join(["'", model.get_value(iter, 0), "'"])     ])

        if len(iters) > 0:
            self.families = ''.join([' WHERE species_data.family IN (', sql[1:], ') '])



    def generate_density_map(self, temp_dir):

        ##generate the base map
        scalefactor = 0.01

        layers = []
        selection = self.vcs_widget.get_selection()
        model, selected = selection.get_selected_rows()
        iters = [model.get_iter(path) for path in selected]
        for iter in iters:
            layers.append('./vice-counties/'+vc_list[int(model.get_value(iter, 0))-1][1]+'.shp')

        bounds_bottom_x = 700000
        bounds_bottom_y = 1300000
        bounds_top_x = 0
        bounds_top_y = 0

        # Read in the shapefiles to get the bounding box
        #need to round to the nearest dis unit so we don't cut off edge square###################
        #########################################################################################
        for shpfile in layers:
            r = shapefile.Reader(shpfile)

            if r.bbox[0] < bounds_bottom_x:
                bounds_bottom_x = r.bbox[0]

            if r.bbox[1] < bounds_bottom_y:
                bounds_bottom_y = r.bbox[1]

            if r.bbox[2] > bounds_top_x:
                bounds_top_x = r.bbox[2]

            if r.bbox[3] > bounds_top_y:
                bounds_top_y = r.bbox[3]

            # Geographic x & y distance
            xdist = bounds_top_x - bounds_bottom_x
            ydist = bounds_top_y - bounds_bottom_y

        base_map = Image.new('RGB', (int(xdist*scalefactor)+1, int(ydist*scalefactor)+1), 'white')
        base_map_draw = ImageDraw.Draw(base_map)

        miniscale = Image.open('./backgrounds/miniscale.png', 'r')
        region = miniscale.crop((int(bounds_bottom_x/100), (1300000/100)-int(bounds_top_y/100), int(bounds_top_x/100)+1, (1300000/100)-int(bounds_bottom_y/100)))
        region.save('crop', format='PNG')

        base_map.paste(region, (0, 0, (int(xdist*scalefactor)+1), (int(ydist*scalefactor)+1)) )


        #add the total coverage & calc first and date band 2 grid arrays
        self.dataset.cursor.execute('SELECT grid_' + self.dataset.config.get('Atlas', 'distribution_unit') + ' AS grids, COUNT(DISTINCT taxon) as species \
                                     FROM data \
                                     GROUP BY grid_' + self.dataset.config.get('Atlas', 'distribution_unit'))

        data = self.dataset.cursor.fetchall()
        #print data
        grids = []

        gridsdict = {}

        for tup in data:
            grids.append(tup[0])
            gridsdict[tup[0]] = tup[1]
            #print tup[0]

        r = shapefile.Reader('./markers/' + self.dataset.config.get('Atlas', 'coverage_style') + '/' + self.dataset.config.get('Atlas', 'distribution_unit'))
        #loop through each object in the shapefile
        for obj in r.shapeRecords():
            #if the grid is in our coverage, add it to the map
            if obj.record[0] in grids:
                #add the grid to to our holding layer so we can access it later without having to loop through all of them each time
                pixels = []
                #loop through each point in the object
                for x,y in obj.shape.points:
                    px = (xdist * scalefactor)- (bounds_top_x - x) * scalefactor
                    py = (bounds_top_y - y) * scalefactor
                    pixels.append((px,py))

                gradfill = 'rgb(255, 255, 255)'

                if gridsdict[obj.record[0]] >= 1 and gridsdict[obj.record[0]] <= 1:
                    gradfill = 'rgb(255, 255, 128)'
                elif gridsdict[obj.record[0]] >= 2 and gridsdict[obj.record[0]] <= 2:
                    gradfill = 'rgb(244, 236, 118)'
                elif gridsdict[obj.record[0]] >= 3 and gridsdict[obj.record[0]] <= 3:
                    gradfill = 'rgb(233, 218, 109)'
                elif gridsdict[obj.record[0]] >= 4 and gridsdict[obj.record[0]] <= 5:
                    gradfill = 'rgb(223, 200, 100)'
                elif gridsdict[obj.record[0]] >= 6 and gridsdict[obj.record[0]] <= 10:
                    gradfill = 'rgb(212, 182, 91)'
                elif gridsdict[obj.record[0]] >= 11 and gridsdict[obj.record[0]] <= 15:
                    gradfill = 'rgb(202, 164, 82)'
                elif gridsdict[obj.record[0]] >= 16 and gridsdict[obj.record[0]] <= 20:
                    gradfill = 'rgb(191, 146, 73)'
                elif gridsdict[obj.record[0]] >= 21 and gridsdict[obj.record[0]] <= 35:
                    gradfill = 'rgb(181, 127, 64)'
                elif gridsdict[obj.record[0]] >= 36 and gridsdict[obj.record[0]] <= 50:
                    gradfill = 'rgb(170, 109, 55)'
                elif gridsdict[obj.record[0]] >= 51 and gridsdict[obj.record[0]] <= 75:
                    gradfill = 'rgb(160, 91, 46)'
                elif gridsdict[obj.record[0]] >= 76 and gridsdict[obj.record[0]] <= 100:
                    gradfill = 'rgb(149, 73, 37)'
                elif gridsdict[obj.record[0]] >= 101 and gridsdict[obj.record[0]] <= 250:
                    gradfill = 'rgb(139, 55, 28)'
                elif gridsdict[obj.record[0]] >= 251 and gridsdict[obj.record[0]] <= 500:
                    gradfill = 'rgb(128, 37, 19)'
                elif gridsdict[obj.record[0]] >= 501:
                    gradfill = 'rgb(118, 19, 10)'

                base_map_draw.polygon(pixels, fill=gradfill, outline='rgb(0,0,0)')

        #add the grid lines
        if self.dataset.config.getboolean('Atlas', 'grid_lines_visible'):
            r = shapefile.Reader('./markers/squares/' + self.dataset.config.get('Atlas', 'grid_lines_style'))
            #loop through each object in the shapefile
            for obj in r.shapes():
                pixels = []
                #loop through each point in the object
                for x,y in obj.points:
                    px = (xdist * scalefactor)- (bounds_top_x - x) * scalefactor
                    py = (bounds_top_y - y) * scalefactor
                    pixels.append((px,py))
                base_map_draw.polygon(pixels, outline='rgb(' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'grid_lines_colour')).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'grid_lines_colour')).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'grid_lines_colour')).blue_float*255)) + ')')



        #mask off grid lines outside the boundary area
        mask = Image.new('RGBA', base_map.size)
        mask_draw = ImageDraw.Draw(mask)

        #mask each boundary shapefile
        for shpfile in layers:
            r = shapefile.Reader(shpfile)
            #loop through each object in the shapefile
            for obj in r.shapes():
                pixels = []
                counter = 1
                #loop through each point in the object
                for x,y in obj.points:
                    px = (xdist * scalefactor)- (bounds_top_x - x) * scalefactor
                    py = (bounds_top_y - y) * scalefactor
                    pixels.append((px,py))
                    #if we reach the start of new part, draw our polygon and clear pixels for the next
                    if counter in obj.parts:
                        mask_draw.polygon(pixels, fill='rgb(0,0,0)')
                        pixels = []
                    counter = counter + 1
                #draw the final polygon (or the only, if we have just the one)
                mask_draw.polygon(pixels, fill='rgb(0,0,0)')

        mask = ImageChops.invert(mask)
        base_map.paste(mask, (0,0), mask)

        #add each boundary shapefile
        for shpfile in layers:
            r = shapefile.Reader(shpfile)
            #loop through each object in the shapefile
            for obj in r.shapes():
                pixels = []
                counter = 1
                #loop through each point in the object
                for x,y in obj.points:
                    px = (xdist * scalefactor)- (bounds_top_x - x) * scalefactor
                    py = (bounds_top_y - y) * scalefactor
                    pixels.append((px,py))
                    #if we reach the start of new part, draw our polygon and clear pixels for the next
                    if counter in obj.parts:
                        base_map_draw.polygon(pixels, outline='rgb(' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'vice-counties_colour')).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'vice-counties_colour')).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'vice-counties_colour')).blue_float*255)) + ')')
                        pixels = []
                    counter = counter + 1
                #draw the final polygon (or the only, if we have just the one)
                base_map_draw.polygon(pixels, outline='rgb(' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'vice-counties_colour')).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'vice-counties_colour')).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'vice-counties_colour')).blue_float*255)) + ')')

        self.density_map_filename = tempfile.NamedTemporaryFile(dir=temp_dir).name
        base_map.save(self.density_map_filename, format='JPEG')

    def generate_base_map(self, temp_dir):

        ##generate the base map
        self.scalefactor = 0.0035

        layers = []
        selection = self.vcs_widget.get_selection()
        model, selected = selection.get_selected_rows()
        iters = [model.get_iter(path) for path in selected]
        for iter in iters:
            layers.append('./vice-counties/'+vc_list[int(model.get_value(iter, 0))-1][1]+'.shp')

        self.bounds_bottom_x = 700000
        self.bounds_bottom_y = 1300000
        self.bounds_top_x = 0
        self.bounds_top_y = 0

        # Read in the shapefiles to get the bounding box
        #need to round to the nearest dis unit so we don't cut off edge square###################
        #########################################################################################
        for shpfile in layers:
            r = shapefile.Reader(shpfile)

            if r.bbox[0] < self.bounds_bottom_x:
                self.bounds_bottom_x = r.bbox[0]

            if r.bbox[1] < self.bounds_bottom_y:
                self.bounds_bottom_y = r.bbox[1]

            if r.bbox[2] > self.bounds_top_x:
                self.bounds_top_x = r.bbox[2]

            if r.bbox[3] > self.bounds_top_y:
                self.bounds_top_y = r.bbox[3]

            # Geographic x & y distance
            self.xdist = self.bounds_top_x - self.bounds_bottom_x
            self.ydist = self.bounds_top_y - self.bounds_bottom_y

        self.base_map = Image.new('RGB', (int(self.xdist*self.scalefactor)+1, int(self.ydist*self.scalefactor)+1), 'white')
        self.base_map_draw = ImageDraw.Draw(self.base_map)

        #add the total coverage & calc first and date band 2 grid arrays
        self.dataset.cursor.execute('SELECT DISTINCT(grid_' + self.dataset.config.get('Atlas', 'distribution_unit') + ') AS grids \
                                     FROM data')

        data = self.dataset.cursor.fetchall()

        grids = []

        for tup in data:
            grids.append(tup[0])
            #print tup[0]

        r = shapefile.Reader('./markers/' + self.dataset.config.get('Atlas', 'coverage_style') + '/' + self.dataset.config.get('Atlas', 'distribution_unit'))
        #loop through each object in the shapefile
        for obj in r.shapeRecords():
            #if the grid is in our coverage, add it to the map
            if obj.record[0] in grids:
                #add the grid to to our holding layer so we can access it later without having to loop through all of them each time
                pixels = []
                #loop through each point in the object
                for x,y in obj.shape.points:
                    px = (self.xdist * self.scalefactor)- (self.bounds_top_x - x) * self.scalefactor
                    py = (self.bounds_top_y - y) * self.scalefactor
                    pixels.append((px,py))
                if self.dataset.config.getboolean('Atlas', 'coverage_visible'):
                    self.base_map_draw.polygon(pixels, fill='rgb(' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'coverage_colour')).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'coverage_colour')).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'coverage_colour')).blue_float*255)) + ')')

        #create date band 1 grid array
        #we always show date band 1
        r = shapefile.Reader('./markers/' + self.dataset.config.get('Atlas', 'date_band_1_style') + '/' + self.dataset.config.get('Atlas', 'distribution_unit'))
        #loop through each object in the shapefile
        for obj in r.shapeRecords():
            if obj.record[0] in grids:
                self.date_band_1_style_coverage.append(obj)

        #create date band 2 grid array
        if self.dataset.config.get('Atlas', 'date_band_2_visible'):
            r = shapefile.Reader('./markers/' + self.dataset.config.get('Atlas', 'date_band_2_style') + '/' + self.dataset.config.get('Atlas', 'distribution_unit'))
            #loop through each object in the shapefile
            for obj in r.shapeRecords():
                if obj.record[0] in grids:
                    self.date_band_2_style_coverage.append(obj)

        #create date band 3 grid array
        if self.dataset.config.get('Atlas', 'date_band_3_visible'):
            r = shapefile.Reader('./markers/' + self.dataset.config.get('Atlas', 'date_band_2_style') + '/' + self.dataset.config.get('Atlas', 'distribution_unit'))
            #loop through each object in the shapefile
            for obj in r.shapeRecords():
                if obj.record[0] in grids:
                    self.date_band_3_style_coverage.append(obj)

        #add the grid lines
        if self.dataset.config.getboolean('Atlas', 'grid_lines_visible'):
            r = shapefile.Reader('./markers/squares/' + self.dataset.config.get('Atlas', 'grid_lines_style'))
            #loop through each object in the shapefile
            for obj in r.shapes():
                pixels = []
                #loop through each point in the object
                for x,y in obj.points:
                    px = (self.xdist * self.scalefactor)- (self.bounds_top_x - x) * self.scalefactor
                    py = (self.bounds_top_y - y) * self.scalefactor
                    pixels.append((px,py))
                self.base_map_draw.polygon(pixels, outline='rgb(' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'grid_lines_colour')).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'grid_lines_colour')).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'grid_lines_colour')).blue_float*255)) + ')')



        #mask off grid lines outside the boundary area
        mask = Image.new('RGBA', self.base_map.size)
        mask_draw = ImageDraw.Draw(mask)

        #mask each boundary shapefile
        for shpfile in layers:
            r = shapefile.Reader(shpfile)
            #loop through each object in the shapefile
            for obj in r.shapes():
                pixels = []
                counter = 1
                #loop through each point in the object
                for x,y in obj.points:
                    px = (self.xdist * self.scalefactor)- (self.bounds_top_x - x) * self.scalefactor
                    py = (self.bounds_top_y - y) * self.scalefactor
                    pixels.append((px,py))
                    #if we reach the start of new part, draw our polygon and clear pixels for the next
                    if counter in obj.parts:
                        mask_draw.polygon(pixels, fill='rgb(0,0,0)')
                        pixels = []
                    counter = counter + 1
                #draw the final polygon (or the only, if we have just the one)
                mask_draw.polygon(pixels, fill='rgb(0,0,0)')

        mask = ImageChops.invert(mask)
        self.base_map.paste(mask, (0,0), mask)

        #add each boundary shapefile
        for shpfile in layers:
            r = shapefile.Reader(shpfile)
            #loop through each object in the shapefile
            for obj in r.shapes():
                pixels = []
                counter = 1
                #loop through each point in the object
                for x,y in obj.points:
                    px = (self.xdist * self.scalefactor)- (self.bounds_top_x - x) * self.scalefactor
                    py = (self.bounds_top_y - y) * self.scalefactor
                    pixels.append((px,py))
                    #if we reach the start of new part, draw our polygon and clear pixels for the next
                    if counter in obj.parts:
                        self.base_map_draw.polygon(pixels, outline='rgb(' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'vice-counties_colour')).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'vice-counties_colour')).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'vice-counties_colour')).blue_float*255)) + ')')
                        pixels = []
                    counter = counter + 1
                #draw the final polygon (or the only, if we have just the one)
                self.base_map_draw.polygon(pixels, outline='rgb(' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'vice-counties_colour')).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'vice-counties_colour')).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'vice-counties_colour')).blue_float*255)) + ')')

    def generate(self, temp_dir):
        
        self.emit('progress-pre-begin')

        self.dataset.cursor.execute('SELECT data.taxon, species_data.family, species_data.national_status, species_data.local_status, COUNT(data.taxon), MIN(data.year), MAX(data.year), COUNT(DISTINCT(grid_' + self.dataset.config.get('Atlas', 'distribution_unit') + ')), \
                                   COUNT(DISTINCT(grid_' + self.dataset.config.get('Atlas', 'distribution_unit') + ')) AS tetrads, \
                                   COUNT(data.taxon) AS records, \
                                   MAX(data.year) AS year, \
                                   species_data.description, \
                                   species_data.common_name \
                                   FROM data \
                                   JOIN species_data ON data.taxon = species_data.taxon \
                                   ' + self.families + ' \
                                   GROUP BY data.taxon, species_data.family, species_data.national_status, species_data.local_status, species_data.description, species_data.common_name \
                                   ORDER BY species_data.sort_order')

        data = self.dataset.cursor.fetchall()

        record_count = 0
        taxa_statistics = {}

        for stats in data:
            record_count = record_count + stats[4]

            taxa_statistics[stats[0]] = {}
            taxa_statistics[stats[0]]['family'] = str(stats[1])
            taxa_statistics[stats[0]]['national_designation'] = str(stats[2])
            taxa_statistics[stats[0]]['local_designation'] = str(stats[3])
            taxa_statistics[stats[0]]['count'] = str(stats[4])
            taxa_statistics[stats[0]]['description'] = str(stats[11])
            taxa_statistics[stats[0]]['common_name'] = str(stats[12])

            if stats[5] == None:
                taxa_statistics[stats[0]]['earliest'] = 'Unknown'
            else:
                taxa_statistics[stats[0]]['earliest'] = stats[5]

            if stats[6] == None:
                taxa_statistics[stats[0]]['latest'] = 'Unknown'
            else:
                taxa_statistics[stats[0]]['latest'] = stats[6]

            taxa_statistics[stats[0]]['dist_count'] = stats[7]

        families = []
        species = []
        for item in data:
            if item[0] not in species:
                species.append(item[0])
            if taxa_statistics[item[0]]['family'] not in families:
                families.append(taxa_statistics[item[0]]['family'])

        toc_length = int(math.ceil(float(len(families))/43.0))

        if toc_length%2 != 0:
            toc_length = toc_length + 1

        #contributors
        contrib_data = {}

        self.dataset.cursor.execute('SELECT DISTINCT(data.recorder) \
                                   FROM data')

        recorder_data = self.dataset.cursor.fetchall()

        for recorders in recorder_data:
            names = recorders[0].split(',')
            for name in names:
                if name.strip() not in contrib_data.keys():
                    parts = name.strip().split()

                    if len(name) > 0:
                        initials = []

                        if parts[0].strip() == 'Mr':
                            initials.append('Mr')
                        elif parts[0].strip() == 'Mrs':
                            initials.append('Mrs')
                        elif parts[0].strip() == 'Dr':
                            initials.append('Dr')

                        for qwert in parts[len(initials):]:
                            initials.append(qwert[0:1])

                        working_part = len(parts)-1
                        check_val = 1

                        while ''.join(initials) in contrib_data.values():
                            if check_val <= len(parts[working_part]):
                                initials[working_part] = parts[working_part][0:check_val]
                                check_val = check_val + 1
                            elif check_val > len(parts[working_part]):
                                working_part = working_part - 1
                                check_val = 1

                        contrib_data[name.strip()] = ''.join(initials)

        self.dataset.cursor.execute('SELECT DISTINCT(data.determiner) \
                                    FROM data')

        determiner_data = self.dataset.cursor.fetchall()

        for determiner in determiner_data:
            names = determiner[0].split(',')
            for name in names:
                if name.strip() not in contrib_data.keys():
                    parts = name.strip().split()

                    initials = []

                    if len(name) > 0:
                        if parts[0].strip() == 'Mr':
                            initials.append('Mr')
                        elif parts[0].strip() == 'Mrs':
                            initials.append('Mrs')
                        elif parts[0].strip() == 'Dr':
                            initials.append('Dr')

                        for qwert in parts[len(initials):]:
                            initials.append(qwert[0:1])

                        working_part = len(parts)-1
                        check_val = 1

                        while ''.join(initials) in contrib_data.values():
                            if check_val <= len(parts[working_part]):
                                initials[working_part] = parts[working_part][0:check_val]
                                check_val = check_val + 1
                            elif check_val > len(parts[working_part]):
                                working_part = working_part - 1
                                check_val = 1

                        contrib_data[name.strip()] = ''.join(initials)

        #the pdf
        pdf = PDF(orientation=self.dataset.config.get('Atlas', 'orientation'),unit=self.page_unit,format=self.dataset.config.get('Atlas', 'paper_size'))
        pdf.type = 'atlas'
        pdf.toc_length = toc_length

        pdf.col = 0
        pdf.y0 = 0
        pdf.set_title(self.dataset.config.get('Atlas', 'title'))
        pdf.set_author(self.dataset.config.get('Atlas', 'author'))
        pdf.section = ''

        #title page
        pdf.p_add_page()

        if self.dataset.config.get('Atlas', 'cover_image') is not None and os.path.isfile(self.dataset.config.get('Atlas', 'cover_image')):
            pdf.image(self.dataset.config.get('Atlas', 'cover_image'), 0, 0, pdf.w, pdf.h)

        pdf.set_text_color(0)
        pdf.set_fill_color(255, 255, 255)
        pdf.set_font('Helvetica', '', 28)
        pdf.ln(15)
        pdf.multi_cell(0, 10, pdf.title, 0, 'L', False)

        pdf.ln(20)
        pdf.set_font('Helvetica', '', 18)
        pdf.multi_cell(0, 10, ''.join([pdf.author, '\n',datetime.now().strftime('%B %Y')]), 0, 'L', False)

        #inside cover
        pdf.p_add_page()
        pdf.set_font('Helvetica', '', 12)
        pdf.multi_cell(0, 6, self.dataset.config.get('Atlas', 'inside_cover'), 0, 'J', False)

        pdf.do_header = True

        #introduction
        if len(self.dataset.config.get('Atlas', 'introduction')) > 0:
            pdf.section = ('Introduction')
            pdf.p_add_page()
            pdf.set_font('Helvetica', '', 20)
            pdf.multi_cell(0, 20, 'Introduction', 0, 'J', False)
            pdf.set_font('Helvetica', '', 12)
            pdf.multi_cell(0, 6, self.dataset.config.get('Atlas', 'introduction'), 0, 'J', False)
            pdf.p_add_page()
            pdf.set_font('Helvetica', '', 20)
            pdf.multi_cell(0, 20, 'Species density', 0, 'J', False)

            im = Image.open(self.density_map_filename)

            width, height = im.size

            if self.dataset.config.get('Atlas', 'orientation')[0:1] == 'P':
                scalefactor = (pdf.w-pdf.l_margin-pdf.r_margin)/width
                target_width = width*scalefactor
                target_height = height*scalefactor

                while target_height >= (pdf.h-40-30):
                    target_height = target_height - 1

                scalefactor = target_height/height
                target_width = width*scalefactor

            elif self.dataset.config.get('Atlas', 'orientation')[0:1] == 'L':
                scalefactor = (pdf.h-40-30)/height
                target_width = width*scalefactor
                target_height = height*scalefactor

                while target_width >= (pdf.w-pdf.l_margin-pdf.r_margin):
                    target_width = target_width - 1

                scalefactor = target_width/width
                target_height = height*scalefactor

            centerer = ((pdf.w-pdf.l_margin-pdf.r_margin)-target_width)/2

            pdf.image(self.density_map_filename, pdf.l_margin+centerer, 40, w=target_width, h=target_height, type='JPG')

        taxon_count = 0
        genus_index = {}
        species_index = {}
        common_name_index = {}

        families = []
        rownum = 0

        if self.dataset.config.get('Atlas', 'paper_size') == 'A4':
            max_region_count = 2

        region_count = 3

        #we should really use the selection & get unique taxa?
        for item in data:
            #print taxa_statistics[item[0]]
            taxon_blurb = ''

            pdf.section = ''.join(['Family ', taxa_statistics[item[0]]['family']])

            if region_count > max_region_count:
                region_count = 1
                pdf.startPageNums()
                pdf.p_add_page()

            if taxa_statistics[item[0]]['family'] not in families:
                families.append(taxa_statistics[item[0]]['family'])
                pdf.TOC_Entry(''.join(['Family ', taxa_statistics[item[0]]['family']]), 0)

            # add the toc entry
            #pdf.TOC_Entry(item[0], 1)

            designation = taxa_statistics[item[0]]['national_designation']
            if (designation == '') or (designation == 'None'):
                designation = ' '

            if (taxa_statistics[item[0]]['common_name'] == '') or (taxa_statistics[item[0]]['common_name'] == 'None'):
                common_name = ''
            else:
                common_name = taxa_statistics[item[0]]['common_name']

            if item[0] not in taxa_statistics:
                taxa_statistics[item[0]] = {}
                taxa_statistics[item[0]]['count'] = 0
                taxa_statistics[item[0]]['earliest'] = 'N/A'
                taxa_statistics[item[0]]['latest'] = 'N/A'
                taxa_statistics[item[0]]['dist_count'] = 0
                taxa_statistics[item[0]]['description'] = ''
                taxa_statistics[item[0]]['common_name'] = ''

            while gtk.events_pending():
                gtk.main_iteration()

            # calc the positining for the various elements
            if region_count == 1: # top left taxon map
                y_padding = 19
                x_padding = pdf.l_margin
            elif region_count == 2: # bottom left taxon map
                y_padding = 5 + (((pdf.w / 2)-pdf.l_margin-pdf.r_margin)+3+10+y_padding) + ((((pdf.w / 2)-pdf.l_margin-pdf.r_margin)+3)/3.75)
                x_padding = pdf.l_margin

            #taxon heading
            pdf.set_y(y_padding)
            pdf.set_x(x_padding)
            pdf.set_text_color(255)
            pdf.set_fill_color(59, 59, 59)
            pdf.set_line_width(0.1)
            pdf.set_font('Helvetica', 'BI', 12)
            pdf.cell(((pdf.w)-pdf.l_margin-pdf.r_margin)/2, 5, ''.join([item[0]]), 'TLB', 0, 'L', True)
            pdf.set_font('Helvetica', 'B', 12)
            pdf.cell(((pdf.w)-pdf.l_margin-pdf.r_margin)/2, 5, common_name, 'TRB', 1, 'R', True)
            pdf.set_x(x_padding)
            pdf.multi_cell(((pdf.w)-pdf.l_margin-pdf.r_margin), 5, ''.join([designation]), 1, 'L', True)

            ####compile list of last e.g. 10 records for use below

            self.dataset.cursor.execute('SELECT data.taxon, data.location, data.grid_native, data.grid_' + self.dataset.config.get('Atlas', 'distribution_unit') + ', data.date, data.decade_to, data.year_to, data.month_to, data.recorder, data.determiner, data.vc, data.grid_100m \
                                        FROM data \
                                        WHERE data.taxon = "' + item[0] + '" \
                                        ORDER BY data.year_to || data.month_to || data.day_to desc')

            indiv_taxon_data = self.dataset.cursor.fetchall()
            max_blurb_length = 900 - len(taxa_statistics[item[0]]['description'])
            left_count = 0

            #there has to be a better way?
            for indiv_record in indiv_taxon_data:
                if len(taxon_blurb) < max_blurb_length:
                    if indiv_record[8] != indiv_record[9]:

                        detees = indiv_record[9].split(',')
                        deter = ''

                        for deter_name in sorted(detees):
                            if deter_name != '':
                                deter = ','.join([deter, contrib_data[deter_name.strip()]])

                        if deter != '':
                            det = ''.join([' det. ', deter[1:]])
                    else:
                        det = ''

                    if indiv_record[4] == 'Unknown':
                        date = '[date unknown]'
                    else:
                        date = indiv_record[4]

                    if indiv_record[8] == 'Unknown':
                        rec = ' anon.'
                    else:
                        recs = indiv_record[8].split(',')
                        rec = ''

                        for recorder_name in sorted(recs):
                            rec = ','.join([rec, contrib_data[recorder_name.strip()]])

                    if len(indiv_record[2]) > 8:
                        grid = indiv_record[11]
                    else:
                        grid = indiv_record[2]

                    taxon_blurb = ''.join([taxon_blurb, indiv_record[1], ' (VC', str(indiv_record[10]), ') ', grid, ' ', date.replace('/', '.'), ' (', rec[1:], det, '); '])
                else:
                    left_count = left_count + 1

            if left_count > 0:
                left_blurb = ''.join([' [+ ', str(left_count), ' more]'])
            else:
                left_blurb = ''

            #taxon blurb
            pdf.set_y(y_padding+12)
            pdf.set_x(x_padding+(((pdf.w / 2)-pdf.l_margin-pdf.r_margin)+3+5))
            pdf.set_font('Helvetica', '', 10)
            pdf.set_text_color(0)
            pdf.set_fill_color(255, 255, 255)
            pdf.set_line_width(0.1)

            if len(taxa_statistics[item[0]]['description']) > 0:
                pdf.set_font('Helvetica', 'B', 10)
                pdf.multi_cell((((pdf.w / 2)-pdf.l_margin-pdf.r_margin)+12), 5, ''.join([taxa_statistics[item[0]]['description'], '\n\n']), 0, 'L', False)
                pdf.set_x(x_padding+(((pdf.w / 2)-pdf.l_margin-pdf.r_margin)+3+5))

            pdf.set_font('Helvetica', '', 10)
            pdf.multi_cell((((pdf.w / 2)-pdf.l_margin-pdf.r_margin)+12), 5, ''.join(['Records (most recent first): ', taxon_blurb[:-2], '.', left_blurb]), 0, 'L', False)

            #chart
            chart = Chart(self.dataset, item[0], temp_dir=temp_dir)
            if chart.temp_filename != None:
                pdf.image(chart.temp_filename, x_padding, ((pdf.w / 2)-pdf.l_margin-pdf.r_margin)+3+10+y_padding, ((pdf.w / 2)-pdf.l_margin-pdf.r_margin)+3, (((pdf.w / 2)-pdf.l_margin-pdf.r_margin)+3)/3.75, 'PNG')
            pdf.rect(x_padding, ((pdf.w / 2)-pdf.l_margin-pdf.r_margin)+3+10+y_padding, ((pdf.w / 2)-pdf.l_margin-pdf.r_margin)+3, (((pdf.w / 2)-pdf.l_margin-pdf.r_margin)+3)/3.75)

            #do the map

            current_map =  self.base_map.copy()
            current_map_draw = ImageDraw.Draw(current_map)

            if self.dataset.config.get('Atlas', 'date_band_3_visible'):
                #date band 3
                self.dataset.cursor.execute('SELECT DISTINCT(grid_' + self.dataset.config.get('Atlas', 'distribution_unit') + ') AS grids \
                                            FROM data \
                                            WHERE data.taxon = "' + item[0] + '" \
                                            AND data.year_to >= ' + str(int(float(self.dataset.config.get('Atlas', 'date_band_3_from')))) + '\
                                            AND data.year_to < ' + str(int(float(self.dataset.config.get('Atlas', 'date_band_3_to')))) + ' \
                                            AND data.year_from >= ' + str(int(float(self.dataset.config.get('Atlas', 'date_band_3_from')))) + ' \
                                            AND data.year_from < ' + str(int(float(self.dataset.config.get('Atlas', 'date_band_3_to')))))

                date_band_3 = self.dataset.cursor.fetchall()

                date_band_3_grids = []

                for tup in date_band_3:
                    date_band_3_grids.append(tup[0])


            if self.dataset.config.get('Atlas', 'date_band_2_visible'):
                #date band 2
                self.dataset.cursor.execute('SELECT DISTINCT(grid_' + self.dataset.config.get('Atlas', 'distribution_unit') + ') AS grids \
                                            FROM data \
                                            WHERE data.taxon = "' + item[0] + '" \
                                            AND data.year_to >= ' + str(int(float(self.dataset.config.get('Atlas', 'date_band_2_from')))) + '\
                                            AND data.year_to < ' + str(int(float(self.dataset.config.get('Atlas', 'date_band_2_to')))) + ' \
                                            AND data.year_from >= ' + str(int(float(self.dataset.config.get('Atlas', 'date_band_2_from')))) + ' \
                                            AND data.year_from < ' + str(int(float(self.dataset.config.get('Atlas', 'date_band_2_to')))))

                date_band_2 = self.dataset.cursor.fetchall()

                date_band_2_grids = []

                #show 3 and overlay 3
                if self.dataset.config.get('Atlas', 'date_band_3_visible') and not self.dataset.config.get('Atlas', 'date_band_3_overlay'):
                    for tup in date_band_2:
                        if tup[0] not in date_band_3_grids:
                            date_band_2_grids.append(tup[0])
                else:
                    for tup in date_band_2:
                        date_band_2_grids.append(tup[0])

            #we always show date band 1
            #date band 1
            self.dataset.cursor.execute('SELECT DISTINCT(grid_' + self.dataset.config.get('Atlas', 'distribution_unit') + ') AS grids \
                                        FROM data \
                                        WHERE data.taxon = "' + item[0] + '" \
                                        AND data.year_to >= ' + str(int(float(self.dataset.config.get('Atlas', 'date_band_1_from')))) + ' \
                                        AND data.year_to < ' + str(int(float(self.dataset.config.get('Atlas', 'date_band_1_to')))) + ' \
                                        AND data.year_from >= ' + str(int(float(self.dataset.config.get('Atlas', 'date_band_1_from')))) + ' \
                                        AND data.year_from < ' + str(int(float(self.dataset.config.get('Atlas', 'date_band_1_to')))))

            date_band_1 = self.dataset.cursor.fetchall()
            date_band_1_grids = []

            #show 2 and 3, don't overlay 2 and 3
            if self.dataset.config.get('Atlas', 'date_band_2_visible') and self.dataset.config.get('Atlas', 'date_band_3_visible') and not self.dataset.config.get('Atlas', 'date_band_2_overlay') and not self.dataset.config.get('Atlas', 'date_band_3_overlay'):
                for tup in date_band_1:
                    if tup[0] not in date_band_3_grids and tup[0] not in date_band_2_grids:
                        date_band_1_grids.append(tup[0])

            #show 2 and 3, overlay 2 not 3
            elif self.dataset.config.get('Atlas', 'date_band_2_visible') and self.dataset.config.get('Atlas', 'date_band_3_visible') and self.dataset.config.get('Atlas', 'date_band_2_overlay') and not self.dataset.config.get('Atlas', 'date_band_3_overlay'):
                for tup in date_band_1:
                    if tup[0] not in date_band_3_grids:
                        date_band_1_grids.append(tup[0])

            #show 2 and 3, overlay 3 not 2
            elif self.dataset.config.get('Atlas', 'date_band_2_visible') and self.dataset.config.get('Atlas', 'date_band_3_visible') and not self.dataset.config.get('Atlas', 'date_band_2_overlay') and self.dataset.config.get('Atlas', 'date_band_3_overlay'):
                for tup in date_band_1:
                    if tup[0] not in date_band_2_grids:
                        date_band_1_grids.append(tup[0])

            #show 2, don't overlay 2
            elif self.dataset.config.get('Atlas', 'date_band_2_visible') and not self.dataset.config.get('Atlas', 'date_band_3_visible') and not self.dataset.config.get('Atlas', 'date_band_2_overlay'):
                for tup in date_band_1:
                    if tup[0] and tup[0] not in date_band_2_grids:
                        date_band_1_grids.append(tup[0])

            #show 3, don't overlay 3
            elif not self.dataset.config.get('Atlas', 'date_band_2_visible') and self.dataset.config.get('Atlas', 'date_band_3_visible') and not self.dataset.config.get('Atlas', 'date_band_3_overlay'):
                for tup in date_band_1:
                    if tup[0] and tup[0] not in date_band_3_grids:
                        date_band_1_grids.append(tup[0])

            #else use all
            else:
                for tup in date_band_1:
                        date_band_1_grids.append(tup[0])



            #we always show date band 1
            #loop through each object in the date band 1 grids
            for obj in self.date_band_1_style_coverage:
                if obj.record[0] in date_band_1_grids:
                    pixels = []
                    #loop through each point in the object
                    for x,y in obj.shape.points:
                        px = (self.xdist * self.scalefactor)- (self.bounds_top_x - x) * self.scalefactor
                        py = (self.bounds_top_y - y) * self.scalefactor
                        pixels.append((px,py))
                    current_map_draw.polygon(pixels, fill='rgb(' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_1_fill')).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_1_fill')).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_1_fill')).blue_float*255)) + ')', outline='rgb(' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_1_outline')).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_1_outline')).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_1_outline')).blue_float*255)) + ')')


            if self.dataset.config.get('Atlas', 'date_band_2_visible'):
                #loop through each object in the date band 2 grids
                for obj in self.date_band_2_style_coverage:
                    #print item[0], obj.record[0]
                    if obj.record[0] in date_band_2_grids:
                        #print "yes"
                        pixels = []
                        #loop through each point in the object
                        for x,y in obj.shape.points:
                            px = (self.xdist * self.scalefactor)- (self.bounds_top_x - x) * self.scalefactor
                            py = (self.bounds_top_y - y) * self.scalefactor
                            pixels.append((px,py))
                        current_map_draw.polygon(pixels, fill='rgb(' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_2_fill')).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_2_fill')).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_2_fill')).blue_float*255)) + ')', outline='rgb(' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_2_outline')).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_2_outline')).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_2_outline')).blue_float*255)) + ')')


            if self.dataset.config.get('Atlas', 'date_band_3_visible'):
                #loop through each object in the date band 3 grids
                for obj in self.date_band_3_style_coverage:
                    #print item[0], obj.record[0]
                    if obj.record[0] in date_band_3_grids:
                        #print "yes"
                        pixels = []
                        #loop through each point in the object
                        for x,y in obj.shape.points:
                            px = (self.xdist * self.scalefactor)- (self.bounds_top_x - x) * self.scalefactor
                            py = (self.bounds_top_y - y) * self.scalefactor
                            pixels.append((px,py))
                        current_map_draw.polygon(pixels, fill='rgb(' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_3_fill')).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_3_fill')).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_3_fill')).blue_float*255)) + ')', outline='rgb(' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_3_outline')).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_3_outline')).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_3_outline')).blue_float*255)) + ')')




            temp_map_file = tempfile.NamedTemporaryFile(dir=temp_dir).name
            current_map.save(temp_map_file, format='PNG')

            (width, height) =  current_map.size

            max_width = (pdf.w / 2)-pdf.l_margin-pdf.r_margin
            max_height = (pdf.w / 2)-pdf.l_margin-pdf.r_margin

            if height > width:
                pdfim_height = max_width
                pdfim_width = (float(width)/float(height))*max_width
            else:
                pdfim_height = (float(height)/float(width))*max_width
                pdfim_width = max_width

            img_x_cent = ((max_width-pdfim_width)/2)+2
            img_y_cent = ((max_height-pdfim_height)/2)+12

            #print height,width
            #print pdfim_height,pdfim_width

            pdf.image(temp_map_file, x_padding+img_x_cent, y_padding+img_y_cent, int(pdfim_width), int(pdfim_height), 'PNG')


            #map container
            pdf.set_text_color(0)
            pdf.set_fill_color(255, 255, 255)
            pdf.set_line_width(0.1)
            pdf.rect(x_padding, 10+y_padding, ((pdf.w / 2)-pdf.l_margin-pdf.r_margin)+3, ((pdf.w / 2)-pdf.l_margin-pdf.r_margin)+3)

            #from
            pdf.set_y(11+y_padding)
            pdf.set_x(1+x_padding)
            pdf.set_font('Helvetica', '', 12)
            pdf.set_text_color(0)
            pdf.set_fill_color(255, 255, 255)
            pdf.set_line_width(0.1)

            if str(taxa_statistics[item[0]]['earliest']) == 'Unknown':
                val = '?'
            else:
                val = str(taxa_statistics[item[0]]['earliest'])
            pdf.multi_cell(18, 5, ''.join(['E ', val]), 0, 'L', False)

            #to
            pdf.set_y(11+y_padding)
            pdf.set_x((((pdf.w / 2)-pdf.l_margin-pdf.r_margin)-15)+x_padding)
            pdf.set_font('Helvetica', '', 12)
            pdf.set_text_color(0)
            pdf.set_fill_color(255, 255, 255)
            pdf.set_line_width(0.1)

            if str(taxa_statistics[item[0]]['latest']) == 'Unknown':
                val = '?'
            else:
                val = str(taxa_statistics[item[0]]['latest'])
            pdf.multi_cell(18, 5, ''.join(['L ', val]), 0, 'R', False)

            #records
            pdf.set_y((((pdf.w / 2)-pdf.l_margin-pdf.r_margin)+7)+y_padding)
            pdf.set_x(1+x_padding)
            pdf.set_font('Helvetica', '', 12)
            pdf.set_text_color(0)
            pdf.set_fill_color(255, 255, 255)
            pdf.set_line_width(0.1)
            pdf.multi_cell(18, 5, ''.join(['R ', str(taxa_statistics[item[0]]['count'])]), 0, 'L', False)

            #tetrads
            pdf.set_y((((pdf.w / 2)-pdf.l_margin-pdf.r_margin)+7)+y_padding)
            pdf.set_x((((pdf.w / 2)-pdf.l_margin-pdf.r_margin)-15)+x_padding)
            pdf.set_font('Helvetica', '', 12)
            pdf.set_text_color(0)
            pdf.set_fill_color(255, 255, 255)
            pdf.set_line_width(0.1)
            pdf.multi_cell(18, 5, ''.join(['T ', str(taxa_statistics[item[0]]['dist_count'])]), 0, 'R', False)

            taxon_parts = item[0].split(' ')

            if taxon_parts[0].lower() not in genus_index:
                genus_index[taxon_parts[0].lower()] = ['genus', pdf.num_page_no()+pdf.toc_length]

            if taxa_statistics[item[0]]['common_name'] not in common_name_index and taxa_statistics[item[0]]['common_name'] != 'None':
                common_name_index[taxa_statistics[item[0]]['common_name']] = ['common', pdf.num_page_no()+pdf.toc_length]

            if taxon_parts[len(taxon_parts)-1] == 'agg.':
                part = ' '.join([taxon_parts[len(taxon_parts)-2], taxon_parts[len(taxon_parts)-1]])
            else:
                part = taxon_parts[len(taxon_parts)-1]

            if ''.join([part, ' ', taxon_parts[0].lower()]) not in species_index:
                species_index[''.join([part, ', ', taxon_parts[0]]).lower()] = ['species', pdf.num_page_no()+pdf.toc_length]

            taxon_count = taxon_count + 1

            if self.cancel is True:
                self.cancel = False
                self.emit('progress-cancelled')
                return False

            rownum = rownum + 1
            #self.emit('progress-update', float(rownum)/len(paths))
            #except KeyError:
            #    print item[0]
            #    pass

            region_count = region_count + 1

        pdf.section = ''
        #pdf.do_header = False

        index = genus_index.copy()
        index.update(species_index)
        index.update(common_name_index)

        #index = sorted(index.iteritems())

        pdf.p_add_page()

        if len(families) > 1:
            family_text = ' '.join([str(len(families)), 'families and'])
        else:
            family_text = ''

        if len(data) > 1:
            taxa_text = 'taxa'
        else:
            taxa_text = 'taxon'

        if record_count > 1:
            record_text = 'records.'
        else:
            record_text = 'record'

        pdf.set_y(19)
        pdf.set_font('Helvetica', 'I', 10)
        pdf.multi_cell(0, 5, ' '.join([family_text,
                                       ' '.join([str(len(data)), taxa_text]),
                                       ' '.join(['mapped from', str(record_count), record_text]),]),
                       0, 'J', False)

        #pdf.section = ''
        #pdf.stopPageNums()
        pdf.section = 'Index'
        pdf.p_add_page()

        initial = ''

        pdf.set_y(pdf.y0+20)
        for taxon in sorted(index, key=lambda taxon: taxon.lower()):
            try:
                if taxon[0].upper() != initial:
                    if taxon[0].upper() != 'A':
                        pdf.ln(3)
                    pdf.set_font('Helvetica', 'B', 12)
                    pdf.cell(0, 5, taxon[0].upper(), 0, 1, 'L', 0)
                    initial = taxon[0].upper()

                if index[taxon][0] == 'species':
                    pos = taxon.find(', ')
                    #capitalize the first letter of the genus
                    display_taxon = list(taxon)
                    display_taxon[pos+2] = display_taxon[pos+2].upper()
                    display_taxon = ''.join(display_taxon)
                    pdf.set_font('Helvetica', '', 12)
                    pdf.cell(0, 5, '  '.join([display_taxon, str(index[taxon][1])]), 0, 1, 'L', 0)
                elif index[taxon][0] == 'genus':
                    pdf.set_font('Helvetica', '', 12)
                    pdf.cell(0, 5, '  '.join([taxon.upper(), str(index[taxon][1])]), 0, 1, 'L', 0)
                elif index[taxon][0] == 'common':
                    pdf.set_font('Helvetica', '', 12)
                    pdf.cell(0, 5, '  '.join([taxon, str(index[taxon][1])]), 0, 1, 'L', 0)
            except IndexError:
                pass

        pdf.setcol(0)

        pdf.section = 'Contributors'
        pdf.p_add_page()
        pdf.set_font('Helvetica', '', 20)
        pdf.multi_cell(0, 20, 'Contributors', 0, 'J', False)
        pdf.set_font('Helvetica', '', 12)

        contrib_blurb = []

        for name in sorted(contrib_data.keys()):
            if name != 'Unknown':
                contrib_blurb.append(''.join([name, ' (', contrib_data[name], ')']))

        pdf.multi_cell(0, 5, ''.join([', '.join(contrib_blurb), '.']), 0, 'J', False)

        pdf.section = ''

        pdf.set_y(-30)
        pdf.set_font('Helvetica','',8)

        if pdf.num_page_no() >= 4 and pdf.section != 'Contents':
            pdf.cell(0, 10, 'Generated in seconds using dipper-stda. For more information, see https://github.com/charlie-barnes/dipper-stda.', 0, 1, 'L')
            pdf.cell(0, 10, ''.join(['Vice-county boundaries provided by the National Biodiversity Network. Contains Ordnance Survey data (C) Crown copyright and database right ', str(datetime.now().year), '.']), 0, 1, 'L')

        #pdf.p_add_page()
        pdf.section = ''

        #toc
        pdf.insertTOC(3)

        #output
        pdf.output(self.save_in,'F')
        shutil.rmtree(temp_dir)
        self.emit('progress-end')


class Chart(gtk.Window):

    def __init__(self, dataset, item, temp_dir):
        gtk.Window.__init__(self)

        vbox = gtk.VBox()
        self.add(vbox)
        combo_box = gtk.combo_box_new_text()
        vbox.pack_end(combo_box, False, False, 0)
        combo_box.append_text('months')
        combo_box.append_text('decades')
        combo_box.set_active(0)

        self.set_decorated(False)
        self.item = item

        combo_box.connect('changed', self.set_mode)

        self.data = None
        self.division = None
        self.minimum_decade = None
        self.maximum_decade = None
        self.months = None
        self.decades = None
        self.mode = 'months'
        self.chart = None

        self.temp_filename = None
        self.temp_dir = temp_dir

        self.visible = False

        self.dataset = dataset

        self.transient_window = None
        self.keep_on_top = False
        self.set_transient_for(self.transient_window)

        self.__set_data__()

        if self.chart:
            self.chart.destroy()

        self.__redraw__()


    def toggle_visibility(self, widget=None, event=None):
        if self.visible:
            self.hide()
            self.visible = False
        else:
            self.show_all()
            self.visible = True

        return True

    def set_mode(self, widget):
        self.mode = widget.get_active_text()
        #self.filter_()

    def __redraw__(self):
        # this should be done more efficently. init the bar chart with null values, then just update?

        if self.mode == 'months':
            data = [('J', self.data['Jan'], 'J'),
                    ('F', self.data['Feb'], 'F'),
                    ('M', self.data['Mar'], 'M'),
                    ('A', self.data['Apr'], 'A'),
                    ('M', self.data['May'], 'M'),
                    ('J', self.data['Jun'], 'J'),
                    ('J', self.data['Jul'], 'J'),
                    ('A', self.data['Aug'], 'A'),
                    ('S', self.data['Sep'], 'S'),
                    ('O', self.data['Oct'], 'O'),
                    ('N', self.data['Nov'], 'N'),
                    ('D', self.data['Dec'], 'D')]
        elif self.mode == 'decades':
            print self.data.keys()
        try:
            barchart = bar_chart.BarChart()
            barchart.grid.set_visible(False)
            barchart.set_mode(bar_chart.MODE_VERTICAL)

            for bar_info in data:
                bar = bar_chart.Bar(*bar_info)
                bar.set_color(gtk.gdk.color_parse('darkgrey'))
                bar._label_object.set_property('size', 16)
                bar._value_label_object.set_property('size', 16)
                bar._label_object.set_property('weight', pango.WEIGHT_BOLD)
                bar._value_label_object.set_property('weight', pango.WEIGHT_BOLD)
                bar._label_object.set_property('color', gtk.gdk.color_parse('black'))
                bar._value_label_object.set_property('color', gtk.gdk.color_parse('black'))
                barchart.add_bar(bar)

            self.chart = barchart
            self.chart.set_enable_mouseover(False)
            self.get_children()[0].pack_start(self.chart, True, True, 0)
            self.chart.show()
            self.temp_filename = tempfile.NamedTemporaryFile(dir=self.temp_dir).name
            self.chart.export_png(self.temp_filename, size=(760,240))
        except ZeroDivisionError:
            self.chart = None
            self.temp_filename = None


    def __set_data__(self):
        #create the sql for the sql_filters
        self.dataset.cursor.execute('SELECT COUNT(DISTINCT(data.month)) \
                                   FROM data \
                                   WHERE data.taxon = "' + self.item + '" \
                                   AND data.month IS NOT NULL')

        data = self.dataset.cursor.fetchall()
        self.months = data[0][0]

        self.dataset.cursor.execute('SELECT COUNT(DISTINCT(data.decade)) \
                                   FROM data \
                                   WHERE data.taxon = "' + self.item + '" \
                                   AND data.decade IS NOT NULL')

        data = self.dataset.cursor.fetchall()
        self.decades = data[0][0]

        if self.mode == 'months':
            self.data = {'Jan': 0, 'Feb': 0, 'Mar': 0, 'Apr': 0, 'May': 0,
                         'Jun': 0, 'Jul': 0, 'Aug': 0, 'Sep': 0, 'Oct': 0,
                         'Nov': 0, 'Dec': 0,}

            self.dataset.cursor.execute('SELECT data.month, COUNT(data.taxon) \
                                   FROM data \
                                   WHERE data.taxon = "' + self.item + '" \
                                   AND data.month IS NOT NULL \
                                   GROUP BY data.month')

            data = self.dataset.cursor.fetchall()

            for month in data:
                if month[0] == 1:
                    self.data['Jan'] = month[1]
                elif month[0] == 2:
                    self.data['Feb'] = month[1]
                elif month[0] == 3:
                    self.data['Mar'] = month[1]
                elif month[0] == 4:
                    self.data['Apr'] = month[1]
                elif month[0] == 5:
                    self.data['May'] = month[1]
                elif month[0] == 6:
                    self.data['Jun'] = month[1]
                elif month[0] == 7:
                    self.data['Jul'] = month[1]
                elif month[0] == 8:
                    self.data['Aug'] = month[1]
                elif month[0] == 9:
                    self.data['Sep'] = month[1]
                elif month[0] == 10:
                    self.data['Oct'] = month[1]
                elif month[0] == 11:
                    self.data['Nov'] = month[1]
                elif month[0] == 12:
                    self.data['Dec'] = month[1]

        elif self.mode == 'decades':

            self.dataset.cursor.execute('SELECT data.decade, COUNT(data.taxon) \
                                         FROM data \
                                         WHERE data.taxon = "' + self.item + '" \
                                         AND data.decade IS NOT NULL \
                                         GROUP BY data.decade \
                                         ORDER BY data.decade')

            decades = self.dataset.cursor.fetchall()

            try:
                self.minimum_decade = decades[0][0]
            except IndexError:
                pass
            else:
                self.maximum_decade = decades[len(decades)-1][0]

                self.data = {}

                for decade in decades:
                    self.data[str(decade[0])] = decade[1]


if __name__ == '__main__':

    #Load the specified file if we have one
    if len(sys.argv) > 1:
        if os.path.exists(sys.argv[1]):
            Run(sys.argv[1])
        else:
            Run()
    else:
        Run()
    gtk.main()
