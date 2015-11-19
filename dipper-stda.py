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

__version__ = "1.1.0"

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
import glob
import random
import pango
import xlrd
import copy
from PIL import ImageChops
import math
from subprocess import call
import ConfigParser
from pygtk_chart import bar_chart
from colour import Color

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
backgrounds = []

#walk the markers directory searching for GIS markers
for style in os.listdir('markers/'):
    markers.append(style)

#walk the backgrounds directory searching for GIS markers
for background in os.listdir('backgrounds/'):
    backgrounds.append(background)

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

gradation_ranges = [[1,1],
                    [2,2],
                    [3,3],
                    [4,5],
                    [6,10],
                    [11,15],
                    [16,20],
                    [21,35],
                    [36,50],
                    [51,75],
                    [76,100],
                    [101,250],
                    [251,500],
                    [501,1000]]


def repeat_to_length(string_to_expand, length):
   return (string_to_expand * ((length/len(string_to_expand))+1))[:length]

class Run():

    def __init__(self, filename=None):

       
        self.builder = gtk.Builder()
        self.builder.add_from_file('./gui/gui.glade')

        win = self.builder.get_object('window1')
        win.maximize()   

        signals = {'quit':self.quit,
                   'generate_atlas':self.generate_atlas,
                   'generate_list':self.generate_list,
                   'open_dataset':self.open_dataset,
                   'select_all_families':self.select_all_families,
                   'unselect_image':self.unselect_image,
                   'update_title':self.update_title,
                   'show_about':self.show_about,
                   'save_config_as':self.save_config_as,
                   'save_configuration':self.save_configuration,
                   'load_config':self.load_config,
                   'switch_update_title':self.switch_update_title,
                   'open_file':self.open_file,
                   'navigation_change':self.navigation_change,
                  }
        self.builder.connect_signals(signals)
        self.dataset = None

        self.builder.get_object('notebook1').set_show_tabs(False)
        self.builder.get_object('notebook2').set_show_tabs(False)
        self.builder.get_object('notebook3').set_show_tabs(False)
        


        #filter for the cover image filechooser
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


        #navigation treeview        
        treeview = self.builder.get_object('treeview5')
        treeview.set_rules_hint(False)
        treeview.get_selection().set_mode(gtk.SELECTION_SINGLE)

        rendererText = gtk.CellRendererText()
        column = gtk.TreeViewColumn("Atlas", rendererText, text=0)
        treeview.append_column(column)
        
        treeselection = treeview.get_selection()
        
        store = gtk.TreeStore(str, int, int)
        self.builder.get_object('treeview5').set_model(store)
        iter = store.append(None, ['Atlas', 0, 0])
        treeselection.select_iter(iter)
        store.append(iter, ['Families', 0, 1])
        store.append(iter, ['Vice-counties', 0, 2])
        store.append(iter, ['Page Setup', 0, 3])
        store.append(iter, ['Table of Contents', 0, 4])
        store.append(iter, ['Species Density Map', 0, 5])
        store.append(iter, ['Species Accounts', 0, 6])
        iter = store.append(None, ['Checklist', 1, 0])
        store.append(iter, ['Families', 1, 1])
        store.append(iter, ['Vice-counties', 1, 2])
        store.append(iter, ['Page Setup', 1, 3])
        
        treeview.expand_all()

        #atlas paper orientation
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

        species_density_background_liststore = gtk.ListStore(gobject.TYPE_STRING)

        for i in range(len(backgrounds)):
            species_density_background_liststore.append([backgrounds[i]])

        combo = self.builder.get_object('combobox16')
        combo.set_model(species_density_background_liststore)
        cell = gtk.CellRendererText()
        combo.pack_start(cell, True)
        combo.add_attribute(cell, 'text',0)

        species_density_style_liststore = gtk.ListStore(gobject.TYPE_STRING)

        for i in range(len(markers)):
            species_density_style_liststore.append([markers[i]])

        combo = self.builder.get_object('combobox13')
        combo.set_model(species_density_style_liststore)
        cell = gtk.CellRendererText()
        combo.pack_start(cell, True)
        combo.add_attribute(cell, 'text',0)

        species_density_unit_liststore = gtk.ListStore(gobject.TYPE_STRING)

        for i in range(len(grid_resolution)):
            species_density_unit_liststore.append([grid_resolution[i]])

        combo = self.builder.get_object('combobox14')
        combo.set_model(species_density_unit_liststore)
        cell = gtk.CellRendererText()
        combo.pack_start(cell, True)
        combo.add_attribute(cell, 'text',0)

        species_density_grid_lines_liststore = gtk.ListStore(gobject.TYPE_STRING)

        for i in range(len(grid_resolution)):
            species_density_grid_lines_liststore.append([grid_resolution[i]])

        combo = self.builder.get_object('combobox15')
        combo.set_model(species_density_grid_lines_liststore)
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

        dialog = self.builder.get_object('window1')
        dialog.show()

        if filename != None:
            self.open_dataset(None, filename)
            
    def navigation_change(self, widget):
        store = widget.get_selected()[0]
        iter = widget.get_selected()[1]
        
        try:
            main_notebook_page = store.get_value(iter, 1)
            sub_notebook_page = store.get_value(iter, 2)
            
            self.builder.get_object('notebook1').set_current_page(main_notebook_page)
            
            if main_notebook_page == 0:
                self.builder.get_object('notebook2').set_current_page(sub_notebook_page)
            elif main_notebook_page == 1:
                self.builder.get_object('notebook3').set_current_page(sub_notebook_page)
        except TypeError:
            pass
            

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
        dialog.set_name('Atlas & Checklist Generator\n')
        dialog.set_version(''.join(['dipper-stda ', __version__]))
        dialog.set_authors(['Charlie Barnes'])
        dialog.set_website('https://github.com/charlie-barnes/dipper-stda')
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

        if self.builder.get_object('notebook1').get_current_page() == 0 and self.dataset.config.getboolean('Atlas', 'families_update_title'):
            entry = self.builder.get_object('entry3')
        elif self.builder.get_object('notebook1').get_current_page() == 1 and self.dataset.config.getboolean('List', 'families_update_title'):
            entry = self.builder.get_object('entry4')

        try:
            model, selected = selection.get_selected_rows()
            iters = [model.get_iter(path) for path in selected]

            orig_title = entry.get_text().split("\n")[0]

            if len(iters) > 2:
                entry.set_text(''.join([orig_title,
                                        "\n",
                                        model.get_value(iters[0], 0),
                                        ' to ',
                                        model.get_value(iters[len(iters)-1], 0)
                                        ]))

            elif len(iters) == 2:
                entry.set_text(''.join([orig_title,
                                        "\n",
                                        model.get_value(iters[0], 0),
                                        ' and ',
                                        model.get_value(iters[1], 0),
                                        ]))

            elif len(iters) == 1:
                entry.set_text(''.join([orig_title,
                                        "\n",
                                        model.get_value(iters[0], 0),
                                        ]))
        except UnboundLocalError:
            pass



    def select_all_families(widget, treeview):
        """Select all families in the selection."""
        treeview.get_selection().select_all()

    def unselect_image(widget, filechoooserbutton):
        """Clear the cover image file selection."""
        filechoooserbutton.unselect_all()

    def open_dataset(self, widget, filename=None):
        """Open a data file."""
        self.builder.get_object('notebook1').set_sensitive(False)

        #if this isn't the first dataset we've opened this session,
        #delete the preceeding temp directory
        if not self.dataset == None:
            try:
                shutil.rmtree(self.dataset.temp_dir)
            except OSError:
                pass

        try:
            self.dataset = Dataset(widget.get_filename())
        except AttributeError:
            self.dataset = Dataset(filename)

        try:

            if self.dataset.data_source.read() == True:

                self.builder.get_object('treeview4').set_sensitive(self.dataset.use_vcs)
                
                if self.dataset.use_vcs:
                    self.builder.get_object('label61').set_markup('<i>Data will be grouped as one if no vice-counties are selected</i>')
                    self.builder.get_object('label37').set_markup('<i>The selection is used to both filter the records and draw the map</i>')
                else:
                    self.builder.get_object('label61').set_markup('<i>Vice-county information is not present in the source file - data will be grouped as one</i>')
                    self.builder.get_object('label37').set_markup('<i>Vice-county information is not present in the source file - selection will just be used draw the maps</i>')

                while gtk.events_pending():
                    gtk.main_iteration_do(True)
                
                self.builder.get_object('window1').set_title(''.join([os.path.basename(self.dataset.filename), ' (',  os.path.dirname(self.dataset.filename), ') - Atlas & Checklist Generator',]) )
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
                    builder.get_object('label68').set_text('Settings file:')
                    builder.get_object('dialog').set_title('Select settings file')
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
                    else:
                        dialog.destroy()
                        return -1

                    dialog.destroy()
                else:
                    config_file = ''.join([os.path.splitext(self.dataset.filename)[0], '.cfg'])

                self.dataset.config.read([config_file])
                self.dataset.config.filename = config_file
                
                if os.path.isfile(self.dataset.config.filename):
                    config_file_txt = os.path.basename(self.dataset.config.filename)
                else:
                    config_file_txt = '(default)'
                                      
                self.builder.get_object('label24').set_markup(''.join(['<b>Sheets:</b> ', self.dataset.sheet, '      <b>Settings:</b> ', config_file_txt]))
                
                self.update_widgets()

                self.builder.get_object('hbox1').show()
                self.builder.get_object('menuitem5').set_sensitive(True)
                self.builder.get_object('menuitem6').set_sensitive(True)
                self.builder.get_object('menuitem7').set_sensitive(True)
                self.builder.get_object('imagemenuitem3').set_sensitive(True)
                self.builder.get_object('imagemenuitem4').set_sensitive(True)
                self.builder.get_object('toolbutton5').set_sensitive(True)
                self.builder.get_object('toolbutton2').set_sensitive(True)
                self.builder.get_object('toolbutton3').set_sensitive(True)
        except AttributeError as e:
            self.builder.get_object('menuitem5').set_sensitive(False)
            self.builder.get_object('menuitem6').set_sensitive(False)
            self.builder.get_object('menuitem7').set_sensitive(False)
            self.builder.get_object('imagemenuitem3').set_sensitive(False)
            self.builder.get_object('imagemenuitem4').set_sensitive(False)
            self.builder.get_object('toolbutton5').set_sensitive(False)
            self.builder.get_object('toolbutton2').set_sensitive(False)
            self.builder.get_object('toolbutton3').set_sensitive(False)
            md = gtk.MessageDialog(None,
                gtk.DIALOG_DESTROY_WITH_PARENT, gtk.MESSAGE_ERROR,
                gtk.BUTTONS_CLOSE, ''.join(['Unable to open data file: ', str(e)]))
            md.run()
            md.destroy()


    def quit(self, widget, third=None):
        """Quit."""
        if not self.dataset == None:
            try:
                shutil.rmtree(self.dataset.temp_dir)
            except:
                pass
        gtk.main_quit()
        sys.exit()

    def generate_atlas(self, widget):
        """Process the data file and configuration."""
        if self.dataset is not None:
            self.update_config()

            vbox = self.builder.get_object('vbox2') 
            notebook = self.builder.get_object('notebook1')           

            #we can't produce an atlas without any geographic entity!
            if not len(self.dataset.config.get('Atlas', 'vice-counties'))>0:
                md = gtk.MessageDialog(None,
                    gtk.DIALOG_DESTROY_WITH_PARENT, gtk.MESSAGE_ERROR,
                    gtk.BUTTONS_CLOSE, 'Select at least one vice-county.')
                md.run()
                md.destroy()
            #we can't produce an atlas or list without any families selected!
            elif not len(self.dataset.config.get('Atlas', 'families'))>0:
                md = gtk.MessageDialog(None,
                    gtk.DIALOG_DESTROY_WITH_PARENT, gtk.MESSAGE_ERROR,
                    gtk.BUTTONS_CLOSE, 'Select at least one family.')
                md.run()
                md.destroy()                   
            else:              

                vbox.set_sensitive(False)

                dialog = gtk.FileChooserDialog('Save As...',
                                               None,
                                               gtk.FILE_CHOOSER_ACTION_SAVE,
                                               (gtk.STOCK_CANCEL, gtk.RESPONSE_CANCEL,
                                                gtk.STOCK_SAVE, gtk.RESPONSE_OK))
                dialog.set_default_response(gtk.RESPONSE_OK)
                dialog.set_do_overwrite_confirmation(True)

                dialog.set_current_folder(os.path.dirname(os.path.abspath(self.dataset.filename)))
                dialog.set_current_name(''.join([os.path.splitext(os.path.basename(self.dataset.config.filename))[0], '_atlas.pdf']))

                response = dialog.run()

                output = dialog.get_filename()

                dialog.destroy()

                while gtk.events_pending():
                    gtk.main_iteration()

                watch = gtk.gdk.Cursor(gtk.gdk.WATCH)
                self.builder.get_object('window1').window.set_cursor(watch)

                while gtk.events_pending():
                    gtk.main_iteration()

                if response == gtk.RESPONSE_OK:

                    #add the extension if it's missing
                    if output[-4:] != '.pdf':
                        output = ''.join([output, '.pdf'])

                    #do the atlas
                    atlas = Atlas(self.dataset)
                    atlas.save_in = output

                    atlas.generate_base_map()

                    if self.dataset.config.getboolean('Atlas', 'species_density_map_visible'):
                        atlas.generate_density_map()

                    atlas.generate()
                    
                    if self.builder.get_object('menuitem9').get_active():
                        if sys.platform == 'linux2':
                            call(["xdg-open", output])
                        else:
                            os.startfile(output)
                        
            vbox.set_sensitive(True)
            self.builder.get_object('window1').window.set_cursor(None)


    def generate_list(self, widget):
        """Process the data file and configuration."""
        if self.dataset is not None:
            self.update_config()

            vbox = self.builder.get_object('vbox2') 
            notebook = self.builder.get_object('notebook1')           

            #we can't produce an atlas without any geographic entity!
            if not len(self.dataset.config.get('List', 'families'))>0:
                md = gtk.MessageDialog(None,
                    gtk.DIALOG_DESTROY_WITH_PARENT, gtk.MESSAGE_ERROR,
                    gtk.BUTTONS_CLOSE, 'Select at least one family.')
                md.run()
                md.destroy()                   
            else:              

                vbox.set_sensitive(False)

                dialog = gtk.FileChooserDialog('Save As...',
                                               None,
                                               gtk.FILE_CHOOSER_ACTION_SAVE,
                                               (gtk.STOCK_CANCEL, gtk.RESPONSE_CANCEL,
                                                gtk.STOCK_SAVE, gtk.RESPONSE_OK))
                dialog.set_default_response(gtk.RESPONSE_OK)
                dialog.set_do_overwrite_confirmation(True)

                dialog.set_current_folder(os.path.dirname(os.path.abspath(self.dataset.filename)))
                dialog.set_current_name(''.join([os.path.splitext(os.path.basename(self.dataset.config.filename))[0], '_checklist.pdf']))

                response = dialog.run()

                output = dialog.get_filename()

                dialog.destroy()

                while gtk.events_pending():
                    gtk.main_iteration()

                watch = gtk.gdk.Cursor(gtk.gdk.WATCH)
                self.builder.get_object('window1').window.set_cursor(watch)

                while gtk.events_pending():
                    gtk.main_iteration()

                if response == gtk.RESPONSE_OK:

                    #add the extension if it's missing
                    if output[-4:] != '.pdf':
                        output = ''.join([output, '.pdf'])

                    #do the list
                    listing = List(self.dataset)
                    listing.save_in = output
                    listing.generate()
                    
                    if self.builder.get_object('menuitem9').get_active():
                        if sys.platform == 'linux2':
                            call(["xdg-open", output])
                        else:
                            os.startfile(output)

            vbox.set_sensitive(True)
            self.builder.get_object('window1').window.set_cursor(None)

    def update_widgets(self):

        #set up the atlas gui based on config settings

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
        
        if self.builder.get_object('treeview2').get_realized():
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

        self.builder.get_object('checkbutton18').set_active(self.dataset.config.getboolean('Atlas', 'families_update_title'))

        #vcs
        selection = self.builder.get_object('treeview1').get_selection()

        selection.unselect_all()
        
        if self.builder.get_object('treeview1').get_realized():
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

        #vice county outline
        self.builder.get_object('colorbutton5').set_color(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'vice-counties_outline')))

        #vice county fill
        self.builder.get_object('colorbutton10').set_color(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'vice-counties_fill')))

        #coverage style
        self.builder.get_object('combobox6').set_active(markers.index(self.dataset.config.get('Atlas', 'coverage_style')))

        #coverage colour
        self.builder.get_object('colorbutton4').set_color(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'coverage_colour')))

        #coverage visible
        self.builder.get_object('checkbutton1').set_active(self.dataset.config.getboolean('Atlas', 'coverage_visible'))

        #species density visible
        self.builder.get_object('checkbutton19').set_active(self.dataset.config.getboolean('Atlas', 'species_density_map_visible'))

        #species density background visible
        self.builder.get_object('checkbutton20').set_active(self.dataset.config.getboolean('Atlas', 'species_density_map_background_visible'))

        #species density background
        self.builder.get_object('combobox16').set_active(backgrounds.index(self.dataset.config.get('Atlas', 'species_density_map_background')))

        #species density style
        self.builder.get_object('combobox13').set_active(markers.index(self.dataset.config.get('Atlas', 'species_density_map_style')))

        #species density unit
        self.builder.get_object('combobox14').set_active(grid_resolution.index(self.dataset.config.get('Atlas', 'species_density_map_unit')))

        #species density map low colour
        self.builder.get_object('colorbutton12').set_color(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'species_density_map_low_colour')))

        #species density map high colour
        self.builder.get_object('colorbutton13').set_color(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'species_density_map_high_colour')))

        #species density map grid line visible
        self.builder.get_object('checkbutton21').set_active(self.dataset.config.getboolean('Atlas', 'species_density_grid_lines_visible'))

        #species density map grid line style
        self.builder.get_object('combobox15').set_active(grid_resolution.index(self.dataset.config.get('Atlas', 'species_density_map_grid_lines_style')))

        #species density map grid line colour
        self.builder.get_object('colorbutton14').set_color(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'species_density_map_grid_lines_colour')))

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

        #table of contents
        self.builder.get_object('checkbutton6').set_active(self.dataset.config.getboolean('Atlas', 'toc_show_families'))
        self.builder.get_object('checkbutton9').set_active(self.dataset.config.getboolean('Atlas', 'toc_show_species_names'))
        self.builder.get_object('checkbutton10').set_active(self.dataset.config.getboolean('Atlas', 'toc_show_common_names'))

        #species accounts
        self.builder.get_object('checkbutton12').set_active(self.dataset.config.getboolean('Atlas', 'species_accounts_show_descriptions'))
        self.builder.get_object('checkbutton13').set_active(self.dataset.config.getboolean('Atlas', 'species_accounts_show_latest'))
        self.builder.get_object('checkbutton14').set_active(self.dataset.config.getboolean('Atlas', 'species_accounts_show_statistics'))
        self.builder.get_object('checkbutton16').set_active(self.dataset.config.getboolean('Atlas', 'species_accounts_show_status'))
        self.builder.get_object('checkbutton15').set_active(self.dataset.config.getboolean('Atlas', 'species_accounts_show_phenology'))
        self.builder.get_object('colorbutton11').set_color(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'species_accounts_phenology_colour')))
        self.builder.get_object('entry5').set_text(self.dataset.config.get('Atlas', 'species_accounts_latest_format'))

        #set up the list gui based on config settings
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
        
        if self.builder.get_object('treeview3').get_realized():
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

        self.dataset.config.set('List', 'families_update_title', str(self.builder.get_object('checkbutton17').get_active()))

        #vcs
        selection = self.builder.get_object('treeview4').get_selection()

        selection.unselect_all()
        
        if self.builder.get_object('treeview4').get_realized():
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
        self.dataset.config.set('Atlas', 'families_update_title', str(self.builder.get_object('checkbutton18').get_active()))

        #grab a comma delimited list of vcs
        selection = self.builder.get_object('treeview1').get_selection()
        model, selected = selection.get_selected_rows()
        iters = [model.get_iter(path) for path in selected]

        vcs = ''

        for iter in iters:
            vcs = ','.join([vcs, model.get_value(iter, 0)])

        self.dataset.config.set('Atlas', 'vice-counties', vcs[1:])
        self.dataset.config.set('Atlas', 'vice-counties_fill', str(self.builder.get_object('colorbutton10').get_color()))
        self.dataset.config.set('Atlas', 'vice-counties_outline', str(self.builder.get_object('colorbutton5').get_color()))

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

        #species density map visible
        self.dataset.config.set('Atlas', 'species_density_map_visible', str(self.builder.get_object('checkbutton19').get_active()))

        #species density background visible
        self.dataset.config.set('Atlas', 'species_density_map_background_visible', str(self.builder.get_object('checkbutton20').get_active()))

        #species density background
        self.dataset.config.set('Atlas', 'species_density_map_background', self.builder.get_object('combobox16').get_active_text())

        #species density style
        self.dataset.config.set('Atlas', 'species_density_map_style', self.builder.get_object('combobox13').get_active_text())

        #species density unit
        self.dataset.config.set('Atlas', 'species_density_map_unit', self.builder.get_object('combobox14').get_active_text())

        #species density map low colour
        self.dataset.config.set('Atlas', 'species_density_map_low_colour', str(self.builder.get_object('colorbutton12').get_color()))
        
        #species density map high colour
        self.dataset.config.set('Atlas', 'species_density_map_high_colour', str(self.builder.get_object('colorbutton13').get_color()))

        #species density map grid lines visible
        self.dataset.config.set('Atlas', 'species_density_grid_lines_visible', str(self.builder.get_object('checkbutton21').get_active()))
        
        #species density map grid line style
        self.dataset.config.set('Atlas', 'species_density_map_grid_lines_style', self.builder.get_object('combobox15').get_active_text())
        
        #species density map grid line colour
        self.dataset.config.set('Atlas', 'species_density_map_grid_lines_colour', str(self.builder.get_object('colorbutton14').get_color()))
        
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
        self.dataset.config.set('Atlas', 'species_accounts_phenology_colour', str(self.builder.get_object('colorbutton11').get_color()))
        self.dataset.config.set('Atlas', 'species_accounts_latest_format', self.builder.get_object('entry5').get_text())


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
        self.dataset.config.set('List', 'distribution_unit', self.builder.get_object('combobox5').get_active_text())

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

    def switch_update_title(self, widget):
        self.dataset.config.set('Atlas', 'families_update_title', str(self.builder.get_object('checkbutton18').get_active()))
        self.dataset.config.set('List', 'families_update_title', str(self.builder.get_object('checkbutton17').get_active()))

    def open_file(self, widget):
        '''Open a dataset file'''
        dialog = gtk.FileChooserDialog('Open...',
                                       None,
                                       gtk.FILE_CHOOSER_ACTION_OPEN,
                                       (gtk.STOCK_CANCEL, gtk.RESPONSE_CANCEL,
                                        gtk.STOCK_OPEN, gtk.RESPONSE_OK))
        dialog.set_default_response(gtk.RESPONSE_OK)

        #filter for the data file filechooser
        filter = gtk.FileFilter()
        filter.set_name("Supported data files")
        filter.add_pattern("*.xls")
        filter.add_mime_type("application/vnd.ms-excel")
        dialog.add_filter(filter)
        dialog.set_filter(filter)
        
        
        #if we can run ssconvert, add ssconvert-able filter to the data file filechooser
        try:
            returncode = call(["ssconvert"])

            if returncode == 1:
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
                dialog.add_filter(filter)
        except OSError:
            print "ssconvert isn't available - you're limited to reading XLS files. Install Gnumeric to make use of ssconvert."
            pass

        response = dialog.run()
        data_file = dialog.get_filename()
        dialog.destroy()

        if response == -5:
            self.open_dataset(widget, data_file)
          


        #self.update_widgets()

    def load_config(self, widget):
        '''Load a configuration file'''
        dialog = gtk.FileChooserDialog('Load...',
                                       None,
                                       gtk.FILE_CHOOSER_ACTION_OPEN,
                                       (gtk.STOCK_CANCEL, gtk.RESPONSE_CANCEL,
                                        gtk.STOCK_SAVE, gtk.RESPONSE_OK))
        dialog.set_default_response(gtk.RESPONSE_OK)
        dialog.set_current_folder(os.path.dirname(os.path.abspath(self.dataset.config.filename)))
        dialog.set_current_name(os.path.basename(self.dataset.config.filename))

        filter = gtk.FileFilter()
        filter.set_name("Config files")
        filter.add_pattern("*.cfg")
        dialog.add_filter(filter)

        response = dialog.run()

        config_file = dialog.get_filename()

        dialog.destroy()

        self.dataset.config.read([config_file])
        self.dataset.config.filename = config_file
        self.builder.get_object('label24').set_markup(''.join(['<b>Sheets:</b> ', self.dataset.sheet, '      <b>Settings:</b> ', os.path.basename(self.dataset.config.filename)]))
                    
        self.update_widgets()
        

    def save_configuration(self, widget):
        '''Save a configuration file based on the current settings'''

        self.update_config()

        #write the config file
        with open(self.dataset.config.filename, 'wb') as configfile:
            self.dataset.config.write(configfile)
            
        if os.path.isfile(self.dataset.config.filename):
            config_file_txt = os.path.basename(self.dataset.config.filename)
        else:
            config_file_txt = '(default)'
                              
        self.builder.get_object('label24').set_markup(''.join(['<b>Sheets:</b> ', self.dataset.sheet, '      <b>Settings:</b> ', config_file_txt]))
 

    def save_config_as(self, widget):
        '''Save a configuration file based on the current settings'''

        dialog = gtk.FileChooserDialog('Save As...',
                                       None,
                                       gtk.FILE_CHOOSER_ACTION_SAVE,
                                       (gtk.STOCK_CANCEL, gtk.RESPONSE_CANCEL,
                                        gtk.STOCK_SAVE, gtk.RESPONSE_OK))
        dialog.set_default_response(gtk.RESPONSE_OK)
        dialog.set_current_folder(os.path.dirname(os.path.abspath(self.dataset.config.filename)))
        dialog.set_current_name(os.path.basename(self.dataset.config.filename))
        dialog.set_do_overwrite_confirmation(True)

        filter = gtk.FileFilter()
        filter.set_name("Config files")
        filter.add_pattern("*.cfg")
        dialog.add_filter(filter)

        response = dialog.run()

        output = dialog.get_filename()

        dialog.destroy()

        self.update_config()

        #write the config file
        with open(output, 'wb') as configfile:
            self.dataset.config.write(configfile)
                              
        self.builder.get_object('label24').set_markup(''.join(['<b>Sheets:</b> ', self.dataset.sheet, '      <b>Settings:</b> ', config_file_txt]))


class Dataset(gobject.GObject):

    def __init__(self, filename):
        gobject.GObject.__init__(self)

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
        self.sheet = ''

        self.atlas_config = {}
        self.list_config = {}
        
        self.use_vcs = True

        self.temp_dir = tempfile.mkdtemp()

        temp_file = tempfile.NamedTemporaryFile(dir=self.temp_dir).name

        if self.connection is None:
            self.connection = sqlite3.connect(temp_file)
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
                                                             'families_update_title': 'True',
                                                             'vice-counties': '',
                                                             'vice-counties_fill': '#fff',
                                                             'vice-counties_outline': '#000',
                                                             'date_band_1_style': 'circles',
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
                                                             'species_density_map_visible': 'True',
                                                             'species_density_map_background_visible': 'True',
                                                             'species_density_map_background': 'miniscale.png',
                                                             'species_density_map_style': 'squares',
                                                             'species_density_map_unit': '10km',
                                                             'species_density_map_low_colour': '#FFFF80',
                                                             'species_density_map_high_colour': '#76130A',
                                                             'species_density_grid_lines_visible': 'True',
                                                             'species_density_map_grid_lines_style': '2km',
                                                             'species_density_map_grid_lines_colour': '#d2d2d2',
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
                                                             'species_accounts_latest_format': '%l (VC%v) %g %d (%r %i)',
                                                             'species_accounts_show_statistics': 'True',
                                                             'species_accounts_show_status': 'True',
                                                             'species_accounts_show_phenology': 'True',
                                                             'species_accounts_phenology_colour': '#000',
                                                            })

            self.config.add_section('Atlas')
            self.config.add_section('List')

        #guess the mimetype of the file
        self.mime = mimetypes.guess_type(self.filename)[0]

        if self.mime == 'application/vnd.ms-excel':
            self.data_source = Read(self.filename, self)
        else:
            temp_file = tempfile.NamedTemporaryFile(dir=self.temp_dir).name

            try:
                returncode = call(["ssconvert", self.filename, ''.join([temp_file, '.xls'])])

                if returncode == 0:
                    self.data_source = Read(''.join([temp_file, '.xls']), self)
            except OSError:
                pass


    def close(self):
        self.connection = None
        self.cursor = None

class Read(gobject.GObject):

    def __init__(self, filename, dataset):
        gobject.GObject.__init__(self)

        self.filename = filename
        self.dataset = dataset


    def read(self):
        '''Read the file and insert the data into the sqlite database.'''

        book = xlrd.open_workbook(self.filename)

        #if more than one sheet is in the workbook, display sheet selection
        #dialog
        has_data = False
        ignore_sheets = 0
        for name in book.sheet_names():
            if name[:2] == '--' and name [-2:] == '--':
                ignore_sheets = ignore_sheets + 1

            # do we have a data sheet?
            if name == '--data--':
                has_data = True

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

        temp_taxa_list = []


        try:
            #loop through the selected sheets of the workbook
            for sheet in sheets:
                self.dataset.sheet = ' + '.join([self.dataset.sheet, sheet.name])
                
                # try and match up the column headings
                for col_index in range(sheet.ncols):
                    if sheet.cell(0, col_index).value.lower() in ['taxon name', 'taxon', 'recommended taxon name']:
                        taxon_position = col_index
                    elif sheet.cell(0, col_index).value.lower() in ['grid reference', 'grid ref', 'grid ref.', 'gridref', 'sample spatial reference']:
                        grid_reference_position = col_index
                    elif sheet.cell(0, col_index).value.lower() in ['date', 'sample date']:
                        date_position = col_index
                    elif sheet.cell(0, col_index).value.lower() in ['location', 'site']:
                        location_position = col_index
                    elif sheet.cell(0, col_index).value.lower() in ['recorder', 'recorders']:
                        recorder_position = col_index
                    elif sheet.cell(0, col_index).value.lower() in ['determiner', 'determiners']:
                        determiner_position = col_index
                    elif sheet.cell(0, col_index).value.lower() in ['vc', 'vice-county', 'vice county']:
                        vc_position = col_index

                rownum = 0

                #loop through each row, skipping the header (first) row
                for row_index in range(1, sheet.nrows):
                    taxa = sheet.cell(row_index, taxon_position).value
                    location = sheet.cell(row_index, location_position).value
                    grid_reference = sheet.cell(row_index, grid_reference_position).value
                    date = sheet.cell(row_index, date_position).value
                    recorder = sheet.cell(row_index, recorder_position).value
                    
                    #we can allow null determiners
                    try:
                        determiner = sheet.cell(row_index, determiner_position).value
                    except UnboundLocalError:
                        determiner = None
                    
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

                    #we can allow null vcs 
                    try:
                        vc = sheet.cell(row_index, vc_position).value
                        self.dataset.use_vcs = True
                    except UnboundLocalError:
                        vc = None
                        self.dataset.use_vcs = False

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

                    rownum = rownum + 1

            self.dataset.sheet = self.dataset.sheet[3:]
            
            #load the data sheet
            if has_data:
                sheet = book.sheet_by_name('--data--')

                # try and match up the column headings
                for col_index in range(sheet.ncols):

                    if sheet.cell(0, col_index).value.lower() in ['taxon', 'taxon name']:
                        taxon_position = col_index
                    elif sheet.cell(0, col_index).value.lower() in ['family']:
                        family_position = col_index
                    elif sheet.cell(0, col_index).value.lower() in ['sort order']:
                        sort_order_position = col_index
                    elif sheet.cell(0, col_index).value.lower() in ['nbn key']:
                        nbn_key_position = col_index
                    elif sheet.cell(0, col_index).value.lower() in ['national status (short)', 'status']:
                        national_status_position = col_index
                    elif sheet.cell(0, col_index).value.lower() in ['description']:
                        description_position = col_index
                    elif sheet.cell(0, col_index).value.lower() in ['common name']:
                        common_name_position = col_index

                #loop through each row, skipping the header (first) row
                for row_index in range(1, sheet.nrows):

                    taxa = sheet.cell(row_index, taxon_position).value
                    if taxa in temp_taxa_list:
                        try:
                            family = sheet.cell(row_index, family_position).value
                        except UnboundLocalError:
                            family = ''

                        try:
                            sort_order = sheet.cell(row_index, sort_order_position).value
                        except UnboundLocalError:
                            sort_order = ''

                        try:
                            nbn_key = sheet.cell(row_index, nbn_key_position).value
                        except UnboundLocalError:
                            nbn_key = ''

                        try:
                            national_status = sheet.cell(row_index, national_status_position).value
                        except UnboundLocalError:
                            national_status = ''

                        try:
                            description = sheet.cell(row_index, description_position).value
                        except UnboundLocalError:
                            description = ''

                        try:
                            common_name = sheet.cell(row_index, common_name_position).value
                        except UnboundLocalError:
                            common_name = ''

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


            return True
        except UnboundLocalError as e:
            md = gtk.MessageDialog(None,
                gtk.DIALOG_DESTROY_WITH_PARENT, gtk.MESSAGE_ERROR,
                gtk.BUTTONS_CLOSE, ''.join(['Unable to open data file: ', str(e)]))
            md.run()
            md.destroy()
            return False


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
        self.dataset = None

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
                self.cell(0, 5, self.title.replace('\n', ' - '), 'B', 1, 'R', 0) # even page header
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
                    if vc == None:
                        vc_head_text = ''
                    else:
                        vc_head_text = ''.join(['VC',vc])
                    self.cell((col_width*3), 5, vc_head_text, '0', 0, 'C', 0)
                    self.cell(col_width/4, 5, '', '0', 0, 'C', 0)

                self.ln()

                self.set_x(self.w-(7+col_width+(((col_width*3)+(col_width/4))*len(self.vcs))))
                self.set_font('Helvetica', '', 8)
                self.cell(col_width, 5, '', '0', 0, 'C', 0)

                for vc in sorted(self.vcs):                    
                    #colum headings
                    self.cell(col_width, 5, ' '.join([self.dataset.config.get('List', 'distribution_unit'), 'sqs']), '0', 0, 'C', 0)
                    self.cell(col_width, 5, 'Records', '0', 0, 'C', 0)
                    self.cell(col_width, 5, 'Last in', '0', 0, 'C', 0)
                    self.cell(col_width/4, 5, '', '0', 0, 'C', 0)

                self.y0 = self.get_y()

            if self.section == 'Contributors' or self.section == 'Contents':
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

    def __init__(self, dataset):
        gobject.GObject.__init__(self)
        self.dataset = dataset
        self.page_unit = 'mm'
        self.save_in = None

    def generate(self):

        taxa_statistics = {}
        taxon_list = []

        if self.dataset.use_vcs:
            vcs_sql = ''.join(['data.vc IN (', self.dataset.config.get('List', 'vice-counties'), ') AND'])
            vcs_sql_sel = 'data.vc'
        else:
            vcs_sql = ''
            vcs_sql_sel = '"00"'
            
        families_sql = ''.join(['species_data.family IN ("', '","'.join(self.dataset.config.get('List', 'families').split(',')), '")'])

        self.dataset.cursor.execute('select * from species_data')

        self.dataset.cursor.execute('SELECT data.taxon, species_data.family, species_data.national_status, species_data.local_status, \
                                   COUNT(DISTINCT(grid_' + self.dataset.config.get('List', 'distribution_unit') + ')) AS squares, \
                                   COUNT(data.taxon) AS records, \
                                   MAX(data.year) AS year, \
                                   ' + vcs_sql_sel + ' AS VC \
                                   FROM data \
                                   JOIN species_data ON data.taxon = species_data.taxon \
                                   WHERE ' + vcs_sql + ' ' + families_sql + ' \
                                   GROUP BY data.taxon, species_data.family, species_data.national_status, species_data.local_status, data.vc \
                                   ORDER BY species_data.sort_order, species_data.family, data.taxon')

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

                taxa_statistics[row[0]]['vc'][str(row[7])]['squares'] = str(row[4])
                taxa_statistics[row[0]]['vc'][str(row[7])]['records'] = str(row[5])
                taxa_statistics[row[0]]['vc'][str(row[7])]['year'] = str(row[6])

            else:
                taxa_statistics[row[0]] = {}
                taxa_statistics[row[0]]['vc'] = {}
                taxa_statistics[row[0]]['family'] = str(row[1])
                taxa_statistics[row[0]]['national_designation'] = str(row[2])
                taxa_statistics[row[0]]['local_designation'] = str(row[3])

                taxa_statistics[row[0]]['vc'][str(row[7])] = {}

                taxa_statistics[row[0]]['vc'][str(row[7])]['squares'] = str(row[4])
                taxa_statistics[row[0]]['vc'][str(row[7])]['records'] = str(row[5])
                taxa_statistics[row[0]]['vc'][str(row[7])]['year'] = str(row[6])

        #the pdf
        pdf = PDF(orientation=self.dataset.config.get('List', 'orientation'),unit=self.page_unit,format=self.dataset.config.get('List', 'paper_size'))
        pdf.type = 'list'
        pdf.do_header = False
        pdf.dataset = self.dataset

        pdf.col = 0
        pdf.y0 = 0
        pdf.set_title(self.dataset.config.get('List', 'title'))
        pdf.set_author(self.dataset.config.get('List', 'author'))
        pdf.set_creator(' '.join(['dipper-stda', __version__])) 
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
        pdf.do_header = True
        pdf.set_font('Helvetica', '', 12)
        pdf.multi_cell(0, 6, self.dataset.config.get('List', 'inside_cover'), 0, 'J', False)

        #introduction
        if len(self.dataset.config.get('List', 'introduction')) > 0:
            pdf.section = ('Introduction')       
            pdf.p_add_page() 
            pdf.do_header = True
            pdf.startPageNums()
            pdf.set_font('Helvetica', '', 20)
            pdf.cell(0, 20, 'Introduction', 0, 0, 'L', 0)
            pdf.ln()
            pdf.set_font('Helvetica', '', 12)
            pdf.multi_cell(0, 6, self.dataset.config.get('List', 'introduction'), 0, 0, 'L')
            pdf.ln()
        else:
            pdf.section = (' '.join(['Family', data[0][1].upper()]))     
            pdf.p_add_page()
            pdf.ln()
            

        #main heading            
        pdf.set_font('Helvetica', '', 20)
        pdf.cell(0, 15, 'Checklist', 0, 1, 'L', 0)

        col_width = 12.7#((self.w - self.l_margin - self.r_margin)/2)/7.5

        #vc headings
        pdf.set_font('Helvetica', '', 10)
        pdf.set_line_width(0.0)

        pdf.set_x(pdf.w-(7+col_width+(((col_width*3)+(col_width/4))*len(self.dataset.config.get('List', 'vice-counties').split(',')))))

        pdf.cell(col_width, 5, '', '0', 0, 'C', 0)
        
        if self.dataset.use_vcs:
            for vc in sorted(self.dataset.config.get('List', 'vice-counties').split(',')):
                pdf.vcs = self.dataset.config.get('List', 'vice-counties').split(',')
                pdf.cell((col_width*3), 5, ''.join(['VC',vc]), '0', 0, 'C', 0)
                pdf.cell(col_width/4, 5, '', '0', 0, 'C', 0)
        else:
            pdf.vcs = [None]
            pdf.cell((col_width*3), 5, '', '0', 0, 'C', 0)
            pdf.cell(col_width/4, 5, '', '0', 0, 'C', 0)                


        pdf.ln()

        pdf.set_x(pdf.w-(7+col_width+(((col_width*3)+(col_width/4))*len(self.dataset.config.get('List', 'vice-counties').split(',')))))
        pdf.set_font('Helvetica', '', 8)
        pdf.cell(col_width, 5, '', '0', 0, 'C', 0)

        for vc in sorted(pdf.vcs):
            #colum headings
            pdf.cell(col_width, 5, ' '.join([self.dataset.config.get('List', 'distribution_unit'), 'sqs']), '0', 0, 'C', 0)
            pdf.cell(col_width, 5, 'Records', '0', 0, 'C', 0)
            pdf.cell(col_width, 5, 'Last in', '0', 0, 'C', 0)
            pdf.cell(col_width/4, 5, '', '0', 0, 'C', 0)

        pdf.doing_the_list = True
        pdf.set_font('Helvetica', '', 8)
        
        if self.dataset.use_vcs:
            for vckey in sorted(self.dataset.config.get('List', 'vice-counties').split(',')):
                #print vckey
    
                col = self.dataset.config.get('List', 'vice-counties').split(',').index(vckey)+1
    
                pdf.cell(col_width/col, 5, '', '0', 0, 'C', 0)
                pdf.cell(col_width, 5, ' '.join([self.dataset.config.get('List', 'distribution_unit'), 'sqs']), '0', 0, 'C', 0)
                pdf.cell(col_width, 5, 'Records', '0', 0, 'C', 0)
                pdf.cell(col_width, 5, 'Last in', '0', 0, 'C', 0)
        else:
            pdf.cell(col_width/1, 5, '', '0', 0, 'C', 0)
            pdf.cell(col_width, 5, ' '.join([self.dataset.config.get('List', 'distribution_unit'), 'sqs']), '0', 0, 'C', 0)
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
            w = pdf.w-(4+col_width+col_width+(((col_width*3)+(col_width/4))*len(pdf.vcs))) - strsize
            nb = w/pdf.get_string_width('.')

            dots = repeat_to_length('.', int(nb))
            pdf.cell(w, pdf.font_size+2, dots, 0, 0, 'R', 0)

            pdf.set_font('Helvetica', '', 6)
            pdf.cell(col_width, pdf.font_size+3, taxa_statistics[key]['national_designation'], '', 0, 'L', 0)
            pdf.set_font('Helvetica', '', 10)

            if self.dataset.use_vcs:
                for vckey in sorted(self.dataset.config.get('List', 'vice-counties').split(',')):
                    #print vckey
    
                    pdf.set_fill_color(230, 230, 230)
                    try:
    
                        if taxa_statistics[key]['vc'][vckey]['squares'] == '0':
                            squares = '-'
                        else:
                            squares = taxa_statistics[key]['vc'][vckey]['squares']
    
                        pdf.cell(col_width, pdf.font_size+2, squares, '', 0, 'L', 1)
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
            else:
 
                  pdf.set_fill_color(230, 230, 230)
                  try:
  
                      if taxa_statistics[key]['vc']['00']['squares'] == '0':
                          squares = '-'
                      else:
                          squares = taxa_statistics[key]['vc']['00']['squares']
  
                      pdf.cell(col_width, pdf.font_size+2, squares, '', 0, 'L', 1)
                      pdf.cell(col_width, pdf.font_size+2, taxa_statistics[key]['vc']['00']['records'], '', 0, 'L', 1)
                      record_count = record_count + int(taxa_statistics[key]['vc']['00']['records'])
  
                      if taxa_statistics[key]['vc']['00']['year'] == 'None':
                          pdf.cell(col_width, pdf.font_size+2, '?', '', 0, 'L', 1)
                      else:
                          pdf.cell(col_width, pdf.font_size+2, taxa_statistics[key]['vc']['00']['year'], '', 0, 'C', 1)
  
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
        try:
            pdf.output(self.save_in,'F')
        except IOError:
            md = gtk.MessageDialog(None,
                gtk.DIALOG_DESTROY_WITH_PARENT, gtk.MESSAGE_ERROR,
                gtk.BUTTONS_CLOSE, 'Unable to write to file. This usually means it''s open - close it and try again.')
            md.run()
            md.destroy()        


class Atlas(gobject.GObject):

    def __init__(self, dataset):
        gobject.GObject.__init__(self)
        self.dataset = dataset
        self.save_in = None
        self.page_unit = 'mm'
        self.base_map = None
        self.density_map_filename = None
        self.date_band_1_style_coverage = []
        self.date_band_2_style_coverage = []
        self.date_band_3_style_coverage = []
        self.increments = 14               
        
        
    def generate_density_map(self):

        #generate the base map
        scalefactor = 0.01

        layers = []
        for vc in self.dataset.config.get('Atlas', 'vice-counties').split(','):
            layers.append(''.join(['./vice-counties/',vc_list[int(vc)-1][1],'.shp']))

        bounds_bottom_x = 700000
        bounds_bottom_y = 1300000
        bounds_top_x = 0
        bounds_top_y = 0

        # Read in the shapefiles to get the bounding box
        # BUG https://github.com/charlie-barnes/dipper-stda/issues/1
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

        temp_file = tempfile.NamedTemporaryFile(dir=self.dataset.temp_dir).name
        
        if self.dataset.config.getboolean('Atlas', 'species_density_map_background_visible'):
            background_map = Image.open(''.join(['./backgrounds/', self.dataset.config.get('Atlas', 'species_density_map_background')]), 'r')

            region = background_map.crop((int(bounds_bottom_x/100), (1300000/100)-int(bounds_top_y/100), int(bounds_top_x/100)+1, ((1300000/100)-int(bounds_bottom_y/100))))
            region.save(temp_file, format='PNG')

            ###HACK: for some reason the crop of the background map isn't always the right
            #size. height seems to be off by 1 pixel in some cases.
            (region_width, region_height) = region.size

            if region_height == (int(ydist*scalefactor)+1)-1:
                hack_diff = 1
            else:
                hack_diff = 0

            base_map.paste(region, (0, 0, (int(xdist*scalefactor)+1), (int(ydist*scalefactor)+1)-hack_diff  ))

        if self.dataset.use_vcs:
            vcs_sql = ''.join(['WHERE data.vc IN (', self.dataset.config.get('Atlas', 'vice-counties'), ')'])
        else:
            vcs_sql = ''

        #add the total coverage & calc first and date band 2 grid arrays
        self.dataset.cursor.execute('SELECT grid_' + self.dataset.config.get('Atlas', 'species_density_map_unit') + ' AS grids, COUNT(DISTINCT taxon) as species \
                                     FROM data \
                                     ' + vcs_sql + ' \
                                     GROUP BY grid_' + self.dataset.config.get('Atlas', 'species_density_map_unit'))

        data = self.dataset.cursor.fetchall()
        #print data
        grids = []

        gridsdict = {}
        max_count = 0

        for tup in data:
            grids.append(tup[0])
            gridsdict[tup[0]] = tup[1]
            if tup[1] > max_count:
                max_count = tup[1]

        #work out how many increments we need
        for ranges in gradation_ranges:
            if max_count >= ranges[0] and max_count <= ranges[1]:
                self.increments = gradation_ranges.index(ranges)+1

        self.grad_ranges = gradation_ranges[:self.increments]
        
        #calculate the colour gradient
        low = Color(self.dataset.config.get('Atlas', 'species_density_map_low_colour'))
        high = Color(self.dataset.config.get('Atlas', 'species_density_map_high_colour'))
        self.grad_fills = list(low.range_to(high, self.increments))
        
        r = shapefile.Reader('./markers/' + self.dataset.config.get('Atlas', 'species_density_map_style') + '/' + self.dataset.config.get('Atlas', 'species_density_map_unit'))
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

                for swatch_ranges in self.grad_ranges:
                    if gridsdict[obj.record[0]] >= swatch_ranges[0] and gridsdict[obj.record[0]] <= swatch_ranges[1]:
                        base_map_draw.polygon(pixels, fill='rgb(' + str(int(self.grad_fills[self.grad_ranges.index(swatch_ranges)].red*255)) + ',' + str(int(self.grad_fills[self.grad_ranges.index(swatch_ranges)].green*255)) + ',' + str(int(self.grad_fills[self.grad_ranges.index(swatch_ranges)].blue*255)) + ')', outline='rgb(0,0,0)')

        #add the grid lines
        if self.dataset.config.getboolean('Atlas', 'species_density_grid_lines_visible'):
            r = shapefile.Reader('./markers/squares/' + self.dataset.config.get('Atlas', 'species_density_map_grid_lines_style'))
            #loop through each object in the shapefile
            for obj in r.shapes():
                pixels = []
                #loop through each point in the object
                for x,y in obj.points:
                    px = (xdist * scalefactor)- (bounds_top_x - x) * scalefactor
                    py = (bounds_top_y - y) * scalefactor
                    pixels.append((px,py))
                base_map_draw.polygon(pixels, outline='rgb(' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'species_density_map_grid_lines_colour')).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'species_density_map_grid_lines_colour')).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'species_density_map_grid_lines_colour')).blue_float*255)) + ')')

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
                        base_map_draw.polygon(pixels, outline='rgb(' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'vice-counties_outline')).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'vice-counties_outline')).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'vice-counties_outline')).blue_float*255)) + ')')
                        pixels = []
                    counter = counter + 1
                #draw the final polygon (or the only, if we have just the one)
                base_map_draw.polygon(pixels, outline='rgb(' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'vice-counties_outline')).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'vice-counties_outline')).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'vice-counties_outline')).blue_float*255)) + ')')

        self.density_map_filename = tempfile.NamedTemporaryFile(dir=self.dataset.temp_dir).name
        base_map.save(self.density_map_filename, format='PNG')

    def generate_base_map(self):

        #generate the base map
        self.scalefactor = 0.0035

        if self.dataset.use_vcs:
            vcs_sql = ''.join(['WHERE data.vc IN (', self.dataset.config.get('Atlas', 'vice-counties'), ')'])
        else:
            vcs_sql = ''

        layers = []
        for vc in self.dataset.config.get('Atlas', 'vice-counties').split(','):
            layers.append(''.join(['./vice-counties/',vc_list[int(vc)-1][1],'.shp']))

        #add the total coverage & calc first and date band 2 grid arrays
        self.dataset.cursor.execute('SELECT DISTINCT(grid_' + self.dataset.config.get('Atlas', 'distribution_unit') + ') AS grids \
                                     FROM data \
                                    ' + vcs_sql)

        data = self.dataset.cursor.fetchall()

        grids = []

        for tup in data:
            grids.append(tup[0])
            #print tup[0]

        self.bounds_bottom_x = 700000
        self.bounds_bottom_y = 1300000
        self.bounds_top_x = 0
        self.bounds_top_y = 0

        # Read in the coverage grid ref shapefiles and extend the bounding box
        r = shapefile.Reader('./markers/' + self.dataset.config.get('Atlas', 'coverage_style') + '/' + self.dataset.config.get('Atlas', 'distribution_unit'))
        #loop through each object in the shapefile
        for obj in r.shapeRecords():
            #if the grid is in our coverage, extend the bounds to match
            if obj.record[0] in grids:
                if obj.shape.bbox[0] < self.bounds_bottom_x:
                    self.bounds_bottom_x = obj.shape.bbox[0]

                if obj.shape.bbox[1] < self.bounds_bottom_y:
                    self.bounds_bottom_y = obj.shape.bbox[1]

                if obj.shape.bbox[2] > self.bounds_top_x:
                    self.bounds_top_x = obj.shape.bbox[2]

                if obj.shape.bbox[3] > self.bounds_top_y:
                    self.bounds_top_y = obj.shape.bbox[3]

        # Read in the date band 1 grid ref shapefiles and extend the bounding box
        r = shapefile.Reader('./markers/' + self.dataset.config.get('Atlas', 'date_band_1_style') + '/' + self.dataset.config.get('Atlas', 'distribution_unit'))
        #loop through each object in the shapefile
        for obj in r.shapeRecords():
            #if the grid is in our coverage, extend the bounds to match
            if obj.record[0] in grids:
                if obj.shape.bbox[0] < self.bounds_bottom_x:
                    self.bounds_bottom_x = obj.shape.bbox[0]

                if obj.shape.bbox[1] < self.bounds_bottom_y:
                    self.bounds_bottom_y = obj.shape.bbox[1]

                if obj.shape.bbox[2] > self.bounds_top_x:
                    self.bounds_top_x = obj.shape.bbox[2]

                if obj.shape.bbox[3] > self.bounds_top_y:
                    self.bounds_top_y = obj.shape.bbox[3]

        # Read in the date band 2 grid ref shapefiles and extend the bounding box
        r = shapefile.Reader('./markers/' + self.dataset.config.get('Atlas', 'date_band_2_style') + '/' + self.dataset.config.get('Atlas', 'distribution_unit'))
        #loop through each object in the shapefile
        for obj in r.shapeRecords():
            #if the grid is in our coverage, extend the bounds to match
            if obj.record[0] in grids:
                if obj.shape.bbox[0] < self.bounds_bottom_x:
                    self.bounds_bottom_x = obj.shape.bbox[0]

                if obj.shape.bbox[1] < self.bounds_bottom_y:
                    self.bounds_bottom_y = obj.shape.bbox[1]

                if obj.shape.bbox[2] > self.bounds_top_x:
                    self.bounds_top_x = obj.shape.bbox[2]

                if obj.shape.bbox[3] > self.bounds_top_y:
                    self.bounds_top_y = obj.shape.bbox[3]

        # Read in the date band 3 grid ref shapefiles and extend the bounding box
        r = shapefile.Reader('./markers/' + self.dataset.config.get('Atlas', 'date_band_3_style') + '/' + self.dataset.config.get('Atlas', 'distribution_unit'))
        #loop through each object in the shapefile
        for obj in r.shapeRecords():
            #if the grid is in our coverage, extend the bounds to match
            if obj.record[0] in grids:
                if obj.shape.bbox[0] < self.bounds_bottom_x:
                    self.bounds_bottom_x = obj.shape.bbox[0]

                if obj.shape.bbox[1] < self.bounds_bottom_y:
                    self.bounds_bottom_y = obj.shape.bbox[1]

                if obj.shape.bbox[2] > self.bounds_top_x:
                    self.bounds_top_x = obj.shape.bbox[2]

                if obj.shape.bbox[3] > self.bounds_top_y:
                    self.bounds_top_y = obj.shape.bbox[3]

        # Read in the vc shapefiles and extend the bounding box
        # BUG https://github.com/charlie-barnes/dipper-stda/issues/1
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
        #this was indented under the for above. this didn't seem right?
        self.xdist = self.bounds_top_x - self.bounds_bottom_x
        self.ydist = self.bounds_top_y - self.bounds_bottom_y

        self.base_map = Image.new('RGB', (int(self.xdist*self.scalefactor)+1, int(self.ydist*self.scalefactor)+1), 'white')
        self.base_map_draw = ImageDraw.Draw(self.base_map)

        #grab the grids we're dealing with and draw them
        r = shapefile.Reader('./markers/' + self.dataset.config.get('Atlas', 'coverage_style') + '/' + self.dataset.config.get('Atlas', 'distribution_unit'))
        #loop through each object in the shapefile
        for obj in r.shapeRecords():
            #if the grid is in our coverage, add it to the map
            if obj.record[0] in grids:
                #add the grid to our holding layer so we can access it later without having to loop through all of them each time
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
                        self.base_map_draw.polygon(pixels, fill='rgb(' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'vice-counties_fill')).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'vice-counties_fill')).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'vice-counties_fill')).blue_float*255)) + ')', outline='rgb(' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'vice-counties_outline')).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'vice-counties_outline')).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'vice-counties_outline')).blue_float*255)) + ')')
                        pixels = []
                    counter = counter + 1
                #draw the final polygon (or the only, if we have just the one)
                self.base_map_draw.polygon(pixels, fill='rgb(' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'vice-counties_fill')).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'vice-counties_fill')).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'vice-counties_fill')).blue_float*255)) + ')', outline='rgb(' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'vice-counties_outline')).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'vice-counties_outline')).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'vice-counties_outline')).blue_float*255)) + ')')

    def generate(self):

        if self.dataset.use_vcs == True:
            vcs_sql = ''.join(['data.vc IN (', self.dataset.config.get('Atlas', 'vice-counties'), ') AND'])
        else:
            vcs_sql = ''
            
        families_sql = ''.join(['species_data.family IN ("', '","'.join(self.dataset.config.get('Atlas', 'families').split(',')), '")'])

        self.dataset.cursor.execute('SELECT data.taxon, species_data.family, species_data.national_status, species_data.local_status, COUNT(data.taxon), MIN(data.year), MAX(data.year), COUNT(DISTINCT(grid_' + self.dataset.config.get('Atlas', 'distribution_unit') + ')), \
                                   COUNT(DISTINCT(grid_' + self.dataset.config.get('Atlas', 'distribution_unit') + ')) AS squares, \
                                   COUNT(data.taxon) AS records, \
                                   MAX(data.year) AS year, \
                                   species_data.description, \
                                   species_data.common_name \
                                   FROM data \
                                   JOIN species_data ON data.taxon = species_data.taxon \
                                   WHERE ' + vcs_sql + ' ' + families_sql + ' \
                                   GROUP BY data.taxon, species_data.family, species_data.national_status, species_data.local_status, species_data.description, species_data.common_name \
                                   ORDER BY species_data.sort_order, species_data.family, data.taxon')
        
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
            try:
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
            except AttributeError:
                pass

        #the pdf
        pdf = PDF(orientation=self.dataset.config.get('Atlas', 'orientation'),unit=self.page_unit,format=self.dataset.config.get('Atlas', 'paper_size'))
        pdf.type = 'atlas'
        pdf.toc_length = toc_length

        pdf.col = 0
        pdf.y0 = 0
        pdf.set_title(self.dataset.config.get('Atlas', 'title'))
        pdf.set_author(self.dataset.config.get('Atlas', 'author'))
        pdf.set_creator(' '.join(['dipper-stda', __version__])) 
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
            pdf.p_add_page()
            pdf.section = ('Introduction')
            pdf.set_font('Helvetica', '', 20)
            pdf.multi_cell(0, 20, 'Introduction', 0, 'J', False)
            pdf.set_font('Helvetica', '', 12)
            pdf.multi_cell(0, 6, self.dataset.config.get('Atlas', 'introduction'), 0, 'J', False)

        #species density map
        if self.dataset.config.getboolean('Atlas', 'species_density_map_visible'):
            pdf.section = ('Introduction')
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

            pdf.image(self.density_map_filename, pdf.l_margin+centerer, 40, w=target_width, h=target_height, type='PNG')

            #add the colour swatches
            
            #scale down the marker to a sensible size
            if self.dataset.config.get('Atlas', 'distribution_unit') == '100km':
                scalefactor = 0.00025
            elif self.dataset.config.get('Atlas', 'distribution_unit') == '10km':
                scalefactor = 0.0025
            elif self.dataset.config.get('Atlas', 'distribution_unit') == '5km':
                scalefactor = 0.005
            elif self.dataset.config.get('Atlas', 'distribution_unit') == '2km':
                scalefactor = 0.0125
            elif self.dataset.config.get('Atlas', 'distribution_unit') == '1km':
                scalefactor = 0.025
                
            #draw each band marker in turn and save out
            #we always show date band 1
            r = shapefile.Reader('./markers/' + self.dataset.config.get('Atlas', 'coverage_style') + '/' + self.dataset.config.get('Atlas', 'distribution_unit'))
            shapes = r.shapes()
            pixels = []

            #grab the first marker we come to - no need to be fussy
            for x, y in shapes[0].points:
                px = 2+int(float(x-shapes[0].bbox[0]) * scalefactor)
                py = 2+int(float(shapes[0].bbox[3]-y) * scalefactor)
                pixels.append((px,py))

            count = 0

            pdf.set_font('Helvetica', '', 9)
            #loop through the grad fills drawing a swatch for each
            for swatch_ranges in reversed(self.grad_ranges):                
                swatch = Image.new('RGB',
                                     (  4+int(float((shapes[0].bbox[2]-shapes[0].bbox[0])) * scalefactor),
                                        4+int(float((shapes[0].bbox[3]-shapes[0].bbox[1])) * scalefactor)     ),
                                     'white')
                swatch_draw = ImageDraw.Draw(swatch)
                swatch_draw.polygon(pixels, fill='rgb(' + str(int(self.grad_fills[self.grad_ranges.index(swatch_ranges)].red*255)) + ',' + str(int(self.grad_fills[self.grad_ranges.index(swatch_ranges)].green*255)) + ',' + str(int(self.grad_fills[self.grad_ranges.index(swatch_ranges)].blue*255)) + ')', outline='rgb(0, 0, 0)')
                swatch_temp_file = tempfile.NamedTemporaryFile(dir=self.dataset.temp_dir).name
                swatch.save(swatch_temp_file, format='PNG')
    
                pdf.image(swatch_temp_file, 10, (200 + (count*5) + ((14-self.increments)*5)), h=4, type='PNG')

                pdf.set_y(200 + (count*5) + ((14-self.increments)*5))
                pdf.cell(4)
                
                if swatch_ranges[0] == swatch_ranges[1]:
                    swatch_text = str(swatch_ranges[0])
                else:
                    swatch_text = ' '.join([str(swatch_ranges[0]), '-', str(swatch_ranges[1])])
                    
                pdf.cell(10, 5, swatch_text, 0, 1, 'L', True)
                
                count = count + 1         
                

        #explanation map
        
        #HACK - this really needs converting a function to create the 'species
        #package' and then choose one at randomn for the explanation, then
        #loop through for the rest. The only difference is the extra Y padding?
        #the explanation map#######################
        random_species = random.choice(list(taxa_statistics.keys()))

        designation = taxa_statistics[random_species]['national_designation']
        if (designation == '') or (designation == 'None'):
            designation = ' '

        if (taxa_statistics[random_species]['common_name'] == '') or (taxa_statistics[random_species]['common_name'] == 'None'):
            common_name = ''
        else:
            common_name = taxa_statistics[random_species]['common_name']

        pdf.section = ('Introduction')
        pdf.p_add_page()
        pdf.set_font('Helvetica', '', 20)
        pdf.multi_cell(0, 20, 'Species account explanation', 0, 'J', False)

        y_padding = 39#######extra Y padding to centralize
        y_padding = (5 + (((pdf.w / 2)-pdf.l_margin-pdf.r_margin)+3+10+y_padding) + ((((pdf.w / 2)-pdf.l_margin-pdf.r_margin)+3)/3.75))/2
        x_padding = pdf.l_margin

        #taxon heading
        pdf.set_y(y_padding)
        pdf.set_x(x_padding)
        pdf.set_text_color(255)
        pdf.set_fill_color(59, 59, 59)
        pdf.set_line_width(0.1)
        pdf.set_font('Helvetica', 'BI', 12)
        pdf.cell(((pdf.w)-pdf.l_margin-pdf.r_margin)/2, 5, random_species, 'TLB', 0, 'L', True)
        pdf.set_font('Helvetica', 'B', 12)
        pdf.cell(((pdf.w)-pdf.l_margin-pdf.r_margin)/2, 5, common_name, 'TRB', 1, 'R', True)
        pdf.set_x(x_padding)

        status_text = ''
        if self.dataset.config.getboolean('Atlas', 'species_accounts_show_status'):
            status_text = ''.join([designation])

        pdf.multi_cell(((pdf.w)-pdf.l_margin-pdf.r_margin), 5, status_text, 1, 'L', True)

        #compile list of last e.g. 10 records for use below
        self.dataset.cursor.execute('SELECT data.taxon, data.location, data.grid_native, data.grid_' + self.dataset.config.get('Atlas', 'distribution_unit') + ', data.date, data.decade_to, data.year_to, data.month_to, data.recorder, data.determiner, data.vc, data.grid_100m \
                                    FROM data \
                                    WHERE data.taxon = "' + random_species + '" \
                                    ORDER BY data.year_to || data.month_to || data.day_to desc')

        indiv_taxon_data = self.dataset.cursor.fetchall()

        max_species_records_length = 900

        if self.dataset.config.getboolean('Atlas', 'species_accounts_show_descriptions'):
            max_species_records_length = max_species_records_length - len(taxa_statistics[random_species]['description'])

        remaining_records = 0

        #there has to be a better way?
        taxon_recent_records = ''
        for indiv_record in indiv_taxon_data:
            if len(taxon_recent_records) < max_species_records_length:

                #if the determiner is different to the recorder, set the
                #determinater (!)
                try:
                    if indiv_record[8] != indiv_record[9]:

                        detees = indiv_record[9].split(',')
                        deter = ''

                        for deter_name in sorted(detees):
                            if deter_name != '':
                                deter = ','.join([deter, contrib_data[deter_name.strip()]])

                        if deter == 'Unknown' or deter == 'Unknown Unknown' or deter == '':
                            det = ' anon.'
                        else:
                            det = ''.join([' det. ', deter[1:]])
                    else:
                        det = ''
                except AttributeError:
                    det = ''

                if indiv_record[4] == 'Unknown':
                    date = '[date unknown]'
                else:
                    date = indiv_record[4]


                if indiv_record[8] == 'Unknown' or indiv_record[8] == 'Unknown Unknown' or indiv_record[8] == '':
                    rec = ' anon.'
                else:
                    recs = indiv_record[8].split(',')
                    rec = ''

                    for recorder_name in sorted(recs):
                        rec = ','.join([rec, contrib_data[recorder_name.strip()]])

                #limit grid reference to 100m
                if len(indiv_record[2]) > 8:
                    grid = indiv_record[11]
                else:
                    grid = indiv_record[2]

                #taxon_recent_records = ''.join([taxon_recent_records, indiv_record[1], ' (VC', str(indiv_record[10]), ') ', grid, ' ', date.replace('/', '.'), ' (', rec[1:], det, '); '])

                #substitute parameters for record values
                #det and loc values can be empty - to remove empty spaces
                #we check for preceeeding and trailing spaces first with
                #these if they are empty strings
                current_rec = self.dataset.config.get('Atlas', 'species_accounts_latest_format')

                if indiv_record[1] == '':
                    current_rec = current_rec.replace(' %l', indiv_record[1])
                    current_rec = current_rec.replace('%l ', indiv_record[1])
                else:
                    current_rec = current_rec.replace('%l', indiv_record[1])

                if det == '':
                    current_rec = current_rec.replace(' %i', det)
                    current_rec = current_rec.replace('%i ', det)
                else:
                    current_rec = current_rec.replace('%i', det)

                current_rec = current_rec.replace('%v', str(indiv_record[10]))
                current_rec = current_rec.replace('%g', grid)
                current_rec = current_rec.replace('%d', date.replace('/', '.'))
                current_rec = current_rec.replace('%r', rec[1:])

                #append current record to the output
                taxon_recent_records = ''.join([taxon_recent_records, current_rec, '; '])
            else:
                remaining_records = remaining_records + 1

        #if any records remain, add a note to the output
        if remaining_records > 0:
            remaining_records_text = ''.join([' [+ ', str(remaining_records), ' more]'])
        else:
            remaining_records_text = ''

        #taxon blurb
        pdf.set_y(y_padding+12)
        pdf.set_x(x_padding+(((pdf.w / 2)-pdf.l_margin-pdf.r_margin)+3+5))
        pdf.set_font('Helvetica', '', 10)
        pdf.set_text_color(0)
        pdf.set_fill_color(255, 255, 255)
        pdf.set_line_width(0.1)

        if len(taxa_statistics[random_species]['description']) > 0 and self.dataset.config.getboolean('Atlas', 'species_accounts_show_descriptions'):
            pdf.set_font('Helvetica', 'B', 10)
            pdf.multi_cell((((pdf.w / 2)-pdf.l_margin-pdf.r_margin)+12), 5, ''.join([taxa_statistics[random_species]['description'], '\n\n']), 0, 'L', False)
            pdf.set_x(x_padding+(((pdf.w / 2)-pdf.l_margin-pdf.r_margin)+3+5))

        if self.dataset.config.getboolean('Atlas', 'species_accounts_show_latest'):
            pdf.set_font('Helvetica', '', 10)
            pdf.multi_cell((((pdf.w / 2)-pdf.l_margin-pdf.r_margin)+12), 5, ''.join(['Records (most recent first): ', taxon_recent_records[:-2], '.', remaining_records_text]), 0, 'L', False)

        y_for_explanation = pdf.get_y()

        #chart
        if self.dataset.config.getboolean('Atlas', 'species_accounts_show_phenology'):
            chart = Chart(self.dataset, random_species)
            if chart.temp_filename != None:
                pdf.image(chart.temp_filename, x_padding, ((pdf.w / 2)-pdf.l_margin-pdf.r_margin)+3+10+y_padding, ((pdf.w / 2)-pdf.l_margin-pdf.r_margin)+3, (((pdf.w / 2)-pdf.l_margin-pdf.r_margin)+3)/3.75, 'PNG')
            pdf.rect(x_padding, ((pdf.w / 2)-pdf.l_margin-pdf.r_margin)+3+10+y_padding, ((pdf.w / 2)-pdf.l_margin-pdf.r_margin)+3, (((pdf.w / 2)-pdf.l_margin-pdf.r_margin)+3)/3.75)


        current_map =  self.base_map.copy()
        current_map_draw = ImageDraw.Draw(current_map)

        if self.dataset.config.get('Atlas', 'date_band_3_visible'):
            #date band 3
            self.dataset.cursor.execute('SELECT DISTINCT(grid_' + self.dataset.config.get('Atlas', 'distribution_unit') + ') AS grids \
                                        FROM data \
                                        WHERE data.taxon = "' + random_species + '" \
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
                                        WHERE data.taxon = "' + random_species + '" \
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
                                    WHERE data.taxon = "' + random_species + '" \
                                    AND data.year_to >= ' + str(int(float(self.dataset.config.get('Atlas', 'date_band_1_from')))) + ' \
                                    AND data.year_to < ' + str(int(float(self.dataset.config.get('Atlas', 'date_band_1_to')))) + ' \
                                    AND data.year_from >= ' + str(int(float(self.dataset.config.get('Atlas', 'date_band_1_from')))) + ' \
                                    AND data.year_from < ' + str(int(float(self.dataset.config.get('Atlas', 'date_band_1_to')))))

        date_band_1 = self.dataset.cursor.fetchall()
        date_band_1_grids = []

        ###
        ###Overlay:
        ### rather than work out which we should/shouln't display, 
        ### would it be easier to draw a white square if we're
        ### not overlaying? would make it easier should we
        ### transition to user specified number of date bands
        ###

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
                #print random_species, obj.record[0]
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
                #print random_species, obj.record[0]
                if obj.record[0] in date_band_3_grids:
                    #print "yes"
                    pixels = []
                    #loop through each point in the object
                    for x,y in obj.shape.points:
                        px = (self.xdist * self.scalefactor)- (self.bounds_top_x - x) * self.scalefactor
                        py = (self.bounds_top_y - y) * self.scalefactor
                        pixels.append((px,py))
                    current_map_draw.polygon(pixels, fill='rgb(' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_3_fill')).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_3_fill')).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_3_fill')).blue_float*255)) + ')', outline='rgb(' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_3_outline')).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_3_outline')).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_3_outline')).blue_float*255)) + ')')

        temp_map_file = tempfile.NamedTemporaryFile(dir=self.dataset.temp_dir).name
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

        pdf.image(temp_map_file, x_padding+img_x_cent, y_padding+img_y_cent, int(pdfim_width), int(pdfim_height), 'PNG')


        #map container
        pdf.set_text_color(0)
        pdf.set_fill_color(255, 255, 255)
        pdf.set_line_width(0.1)
        pdf.rect(x_padding, 10+y_padding, ((pdf.w / 2)-pdf.l_margin-pdf.r_margin)+3, ((pdf.w / 2)-pdf.l_margin-pdf.r_margin)+3)

        if self.dataset.config.getboolean('Atlas', 'species_accounts_show_statistics'):

            #from
            pdf.set_y(11+y_padding)
            pdf.set_x(1+x_padding)
            pdf.set_font('Helvetica', '', 12)
            pdf.set_text_color(0)
            pdf.set_fill_color(255, 255, 255)
            pdf.set_line_width(0.1)

            if str(taxa_statistics[random_species]['earliest']) == 'Unknown':
                val = '?'
            else:
                val = str(taxa_statistics[random_species]['earliest'])
            pdf.multi_cell(18, 5, ''.join(['E ', val]), 0, 'L', False)

            #to
            pdf.set_y(11+y_padding)
            pdf.set_x((((pdf.w / 2)-pdf.l_margin-pdf.r_margin)-15)+x_padding)
            pdf.set_font('Helvetica', '', 12)
            pdf.set_text_color(0)
            pdf.set_fill_color(255, 255, 255)
            pdf.set_line_width(0.1)

            if str(taxa_statistics[random_species]['latest']) == 'Unknown':
                val = '?'
            else:
                val = str(taxa_statistics[random_species]['latest'])
            pdf.multi_cell(18, 5, ''.join(['L ', val]), 0, 'R', False)

            #records
            pdf.set_y((((pdf.w / 2)-pdf.l_margin-pdf.r_margin)+7)+y_padding)
            pdf.set_x(1+x_padding)
            pdf.set_font('Helvetica', '', 12)
            pdf.set_text_color(0)
            pdf.set_fill_color(255, 255, 255)
            pdf.set_line_width(0.1)
            pdf.multi_cell(18, 5, ''.join(['R ', str(taxa_statistics[random_species]['count'])]), 0, 'L', False)

            #squares
            pdf.set_y((((pdf.w / 2)-pdf.l_margin-pdf.r_margin)+7)+y_padding)
            pdf.set_x((((pdf.w / 2)-pdf.l_margin-pdf.r_margin)-15)+x_padding)
            pdf.set_font('Helvetica', '', 12)
            pdf.set_text_color(0)
            pdf.set_fill_color(255, 255, 255)
            pdf.set_line_width(0.1)
            pdf.multi_cell(18, 5, ''.join(['S ', str(taxa_statistics[random_species]['dist_count'])]), 0, 'R', False)



        #### the explanations
        pdf.set_font('Helvetica', '', 9)
        pdf.set_draw_color(0,0,0)

        #species name
        pdf.line(20,
                 50,
                 x_padding+7,
                 y_padding)

        pdf.set_x(1+x_padding +15)
        pdf.set_y(11+y_padding -50)
        pdf.cell(1)
        pdf.cell(10, 5, 'Species name', 0, 0, 'L', True)

        #common name
        pdf.line(pdf.w-pdf.l_margin-30,
                 50,
                 pdf.w-pdf.l_margin-x_padding-10,
                 y_padding)

        pdf.set_x(1+x_padding +15)
        pdf.set_y(11+y_padding -48)
        pdf.cell(150)
        pdf.cell(10, 5, 'Common name', 0, 0, 'L', True)


        pdf.set_y(y_padding+12)
        pdf.set_x(x_padding+(((pdf.w / 2)-pdf.l_margin-pdf.r_margin)+3+5))

        #taxon blurb
        if self.dataset.config.getboolean('Atlas', 'species_accounts_show_latest') or self.dataset.config.getboolean('Atlas', 'species_accounts_show_descriptions'):
            pdf.line(x_padding+(((pdf.w / 2)-pdf.l_margin-pdf.r_margin)+3+5)   +50, #x1
                     y_for_explanation,                                             #y1
                     170,                                                           #x2
                     240)                                                           #y2

            pdf.set_x(10)
            pdf.set_y(240)
            pdf.cell(130)

            if self.dataset.config.getboolean('Atlas', 'species_accounts_show_latest') and self.dataset.config.getboolean('Atlas', 'species_accounts_show_descriptions'):
                pdf.cell(10, 5, 'Species description and most recent records', 0, 0, 'L', True)
            elif self.dataset.config.getboolean('Atlas', 'species_accounts_show_latest') and not self.dataset.config.getboolean('Atlas', 'species_accounts_show_descriptions'):
                pdf.cell(10, 5, 'Most recent records', 0, 0, 'L', True)
            elif not self.dataset.config.getboolean('Atlas', 'species_accounts_show_latest') and self.dataset.config.getboolean('Atlas', 'species_accounts_show_descriptions'):
                pdf.cell(10, 5, 'Species description', 0, 0, 'L', True)

        #phenology chart
        if self.dataset.config.getboolean('Atlas', 'species_accounts_show_phenology'):
            pdf.line((   ((pdf.w / 2)-pdf.l_margin-pdf.r_margin)+3     )       / 2, #x1
                     ((pdf.w / 2)-pdf.l_margin-pdf.r_margin)+3+10+y_padding    +22, #y1
                     40,                                                            #x2
                     260)                                                           #y2

            pdf.set_x(10)
            pdf.set_y(260)
            pdf.cell(15)
            pdf.cell(10, 5, 'Monthly phenology chart', 0, 0, 'L', True)

        #status
        if self.dataset.config.getboolean('Atlas', 'species_accounts_show_status'):
            pdf.line(1+x_padding                                               +10, #x1
                     11+y_padding                                              -5,  #y1
                     1+x_padding                                               +40, #x2
                     11+y_padding                                              -37) #y2

            pdf.set_x(10)
            pdf.set_y(11+y_padding -43)
            pdf.cell(30)
            pdf.cell(10, 5, 'Species status', 0, 0, 'L', True)

        #statistics
        if self.dataset.config.getboolean('Atlas', 'species_accounts_show_statistics'):
            #earliest
            pdf.line(1+x_padding                                               +15, #x1
                     11+y_padding                                              +2,  #y1
                     1+x_padding                                               +50, #x2
                     11+y_padding                                              -20) #y2

            pdf.set_x(10)
            pdf.set_y(11+y_padding -25)
            pdf.cell(45)
            pdf.cell(10, 5, 'Earliest record', 0, 0, 'L', True)

            #latest
            pdf.line((((pdf.w / 2)-pdf.l_margin-pdf.r_margin)-15)+x_padding    +10, #x1
                     11+y_padding,                                                  #y1
                     (((pdf.w / 2)-pdf.l_margin-pdf.r_margin)-15)+x_padding    +40, #x2
                     11+y_padding -40)                                              #y2

            pdf.set_x(10)
            pdf.set_y(11+y_padding -45)
            pdf.cell(100)
            pdf.cell(10, 5, 'Latest record', 0, 0, 'L', True)

            #number of records
            pdf.line(1+x_padding                                               +5,  #x1
                     (((pdf.w / 2)-pdf.l_margin-pdf.r_margin)+7)+y_padding     +5,  #y1
                     20,                                                            #x2
                     220)                                                           #y2

            pdf.set_x(10)
            pdf.set_y(220)
            pdf.cell(1)
            pdf.cell(10, 5, 'Number of records', 0, 0, 'L', True)

            #number of squares
            pdf.line((((pdf.w / 2)-pdf.l_margin-pdf.r_margin)-15)+x_padding    +15, #x1
                     (((pdf.w / 2)-pdf.l_margin-pdf.r_margin)+7)+y_padding     +5,  #y1
                     (((pdf.w / 2)-pdf.l_margin-pdf.r_margin)-15)+x_padding    +30, #x2
                     240                                                       -20) #y2

            pdf.set_x(10)
            pdf.set_y(220)
            pdf.cell(70)
            pdf.cell(10, 5, ' '.join(['Number of', self.dataset.config.get('Atlas', 'distribution_unit'), 'squares the species occurs in']), 0, 0, 'L', True)


        # the date classes

        pdf.line((((pdf.w / 2)-pdf.l_margin-pdf.r_margin)-15)+x_padding    -30, #x1
                 (((pdf.w / 2)-pdf.l_margin-pdf.r_margin)+7)+y_padding     -30, #y1
                 80, #x2
                 240                                                          ) #y2            
        
        pdf.set_font('Helvetica', '', 9)
        pdf.set_x(70)
        pdf.set_y(240)
        pdf.cell(60)
        pdf.cell(10, 5, 'Date classes:', 0, 1, 'L', True)

        #scale down the marker to a sensible size
        if self.dataset.config.get('Atlas', 'distribution_unit') == '100km':
            scalefactor = 0.00025
        elif self.dataset.config.get('Atlas', 'distribution_unit') == '10km':
            scalefactor = 0.0025
        elif self.dataset.config.get('Atlas', 'distribution_unit') == '5km':
            scalefactor = 0.005
        elif self.dataset.config.get('Atlas', 'distribution_unit') == '2km':
            scalefactor = 0.0125
        elif self.dataset.config.get('Atlas', 'distribution_unit') == '1km':
            scalefactor = 0.025

        #draw each band marker in turn and save out
        #we always show date band 1
        r = shapefile.Reader('./markers/' + self.dataset.config.get('Atlas', 'date_band_1_style') + '/' + self.dataset.config.get('Atlas', 'distribution_unit'))
        shapes = r.shapes()
        pixels = []

        #grab the first marker we come to - no need to be fussy
        for x, y in shapes[0].points:
            px = 2+int(float(x-shapes[0].bbox[0]) * scalefactor)
            py = 2+int(float(shapes[0].bbox[3]-y) * scalefactor)
            pixels.append((px,py))
        
        date_band_1 = Image.new('RGB',
                             (  4+int(float((shapes[0].bbox[2]-shapes[0].bbox[0])) * scalefactor),
                                4+int(float((shapes[0].bbox[3]-shapes[0].bbox[1])) * scalefactor)     ),
                             'white')
        date_band_1_draw = ImageDraw.Draw(date_band_1)
        date_band_1_draw.polygon(pixels, fill='rgb(' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_1_fill')).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_1_fill')).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_1_fill')).blue_float*255)) + ')', outline='rgb(' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_1_outline')).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_1_outline')).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_1_outline')).blue_float*255)) + ')')
        #and as a line so we can increase the thickness
        date_band_1_draw.line(pixels, width=3, fill='rgb(' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_1_outline')).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_1_outline')).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_1_outline')).blue_float*255)) + ')')

        date_band_1_temp_file = tempfile.NamedTemporaryFile(dir=self.dataset.temp_dir).name
        date_band_1.save(date_band_1_temp_file, format='PNG')


        pdf.image(date_band_1_temp_file, 80, 245, h=4, type='PNG')
        pdf.cell(75)
        
        if self.dataset.config.get('Atlas', 'date_band_1_from') == '1600.0':
            date_band_1_from_text = ' '.join(['before', str(int(float(self.dataset.config.get('Atlas', 'date_band_1_to'))))])
        else:
            date_band_1_from_text = ' '.join([str(int(float(self.dataset.config.get('Atlas', 'date_band_1_from')))), 'to', str(int(float(self.dataset.config.get('Atlas', 'date_band_1_to'))))])
        
        pdf.cell(10, 5, date_band_1_from_text, 0, 1, 'L', True)
            

        #date band 2
        if self.dataset.config.getboolean('Atlas', 'date_band_2_visible'):
            r = shapefile.Reader('./markers/' + self.dataset.config.get('Atlas', 'date_band_2_style') + '/' + self.dataset.config.get('Atlas', 'distribution_unit'))
            shapes = r.shapes()
            pixels = []

            #grab the first marker we come to - no need to be fussy
            for x, y in shapes[0].points:
                px = 2+int(float(x-shapes[0].bbox[0]) * scalefactor)
                py = 2+int(float(shapes[0].bbox[3]-y) * scalefactor)
                pixels.append((px,py))

            date_band_2 = Image.new('RGB',
                                 (  4+int(float((shapes[0].bbox[2]-shapes[0].bbox[0])) * scalefactor),
                                    4+int(float((shapes[0].bbox[3]-shapes[0].bbox[1])) * scalefactor)     ),
                                 'white')
            date_band_2_draw = ImageDraw.Draw(date_band_2)
            date_band_2_draw.polygon(pixels, fill='rgb(' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_2_fill')).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_2_fill')).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_2_fill')).blue_float*255)) + ')', outline='rgb(' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_2_outline')).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_2_outline')).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_2_outline')).blue_float*255)) + ')')
            #and as a line so we can increase the thickness
            date_band_2_draw.line(pixels, width=3, fill='rgb(' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_2_outline')).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_2_outline')).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_2_outline')).blue_float*255)) + ')')

            date_band_2_temp_file = tempfile.NamedTemporaryFile(dir=self.dataset.temp_dir).name
            date_band_2.save(date_band_2_temp_file, format='PNG')

            pdf.image(date_band_2_temp_file, 80, 250, h=4, type='PNG')
            pdf.cell(75)
            
            if self.dataset.config.get('Atlas', 'date_band_2_to') == '2050.0':
                date_band_2_from_text = ' '.join([str(int(float(self.dataset.config.get('Atlas', 'date_band_2_from')))), 'onwards'])
            else:
                date_band_2_from_text = ' '.join([str(int(float(self.dataset.config.get('Atlas', 'date_band_2_from')))), 'to', str(int(float(self.dataset.config.get('Atlas', 'date_band_2_to'))))])
            
            pdf.cell(10, 5, date_band_2_from_text, 0, 1, 'L', True)

        #date band 3
        if self.dataset.config.getboolean('Atlas', 'date_band_3_visible'):
            r = shapefile.Reader('./markers/' + self.dataset.config.get('Atlas', 'date_band_3_style') + '/' + self.dataset.config.get('Atlas', 'distribution_unit'))
            shapes = r.shapes()
            pixels = []

            #grab the first marker we come to - no need to be fussy
            for x, y in shapes[0].points:
                px = 2+int(float(x-shapes[0].bbox[0]) * scalefactor)
                py = 2+int(float(shapes[0].bbox[3]-y) * scalefactor)
                pixels.append((px,py))

            date_band_3 = Image.new('RGB',
                                 (  4+int(float((shapes[0].bbox[2]-shapes[0].bbox[0])) * scalefactor),
                                    4+int(float((shapes[0].bbox[3]-shapes[0].bbox[1])) * scalefactor)     ),
                                 'white')
            date_band_3_draw = ImageDraw.Draw(date_band_3)
            date_band_3_draw.polygon(pixels, fill='rgb(' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_3_fill')).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_3_fill')).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_3_fill')).blue_float*255)) + ')', outline='rgb(' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_3_outline')).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_3_outline')).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_3_outline')).blue_float*255)) + ')')
            #and as a line so we can increase the thickness
            date_band_3_draw.line(pixels, width=3, fill='rgb(' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_3_outline')).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_3_outline')).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'date_band_3_outline')).blue_float*255)) + ')')

            date_band_3_temp_file = tempfile.NamedTemporaryFile(dir=self.dataset.temp_dir).name
            date_band_3.save(date_band_3_temp_file, format='PNG')

            pdf.image(date_band_3_temp_file, 80, 255, h=4, type='PNG')
            pdf.cell(75)
            
            if self.dataset.config.get('Atlas', 'date_band_3_to') == '2050.0':
                date_band_3_from_text = ' '.join([str(int(float(self.dataset.config.get('Atlas', 'date_band_3_from')))), 'onwards'])
            else:
                date_band_3_from_text = ' '.join([str(int(float(self.dataset.config.get('Atlas', 'date_band_3_from')))), 'to', str(int(float(self.dataset.config.get('Atlas', 'date_band_3_to'))))])
                        
            pdf.cell(10, 5, date_band_3_from_text, 0, 1, 'L', True)

        #### end the explanations

        #end of explanation map code###############################

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
            taxon_recent_records = ''

            pdf.section = ''.join(['Family ', taxa_statistics[item[0]]['family']])

            if region_count > max_region_count:
                region_count = 1
                pdf.startPageNums()
                pdf.p_add_page()

            if taxa_statistics[item[0]]['family'] not in families:

                families.append(taxa_statistics[item[0]]['family'])

                if self.dataset.config.getboolean('Atlas', 'toc_show_families'):
                    pdf.TOC_Entry(''.join(['Family ', taxa_statistics[item[0]]['family']]), level=0)

            if self.dataset.config.getboolean('Atlas', 'toc_show_species_names') and self.dataset.config.getboolean('Atlas', 'toc_show_common_names'):
                pdf.TOC_Entry(''.join([item[0], ' - ', taxa_statistics[item[0]]['common_name']]), level=1)
            elif not self.dataset.config.getboolean('Atlas', 'toc_show_species_names') and self.dataset.config.getboolean('Atlas', 'toc_show_common_names'):
                pdf.TOC_Entry(taxa_statistics[item[0]]['common_name'], level=1)
            elif self.dataset.config.getboolean('Atlas', 'toc_show_species_names') and not self.dataset.config.getboolean('Atlas', 'toc_show_common_names'):
                pdf.TOC_Entry(item[0], level=1)

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
            if region_count == 1: # top taxon map
                y_padding = 19
                x_padding = pdf.l_margin
            elif region_count == 2: # bottom taxon map
                y_padding = 5 + (((pdf.w / 2)-pdf.l_margin-pdf.r_margin)+3+10+y_padding) + ((((pdf.w / 2)-pdf.l_margin-pdf.r_margin)+3)/3.75)
                x_padding = pdf.l_margin

            #taxon heading
            pdf.set_y(y_padding)
            pdf.set_x(x_padding)
            pdf.set_text_color(255)
            pdf.set_fill_color(59, 59, 59)
            pdf.set_line_width(0.1)
            pdf.set_font('Helvetica', 'BI', 12)
            pdf.cell(((pdf.w)-pdf.l_margin-pdf.r_margin)/2, 5, item[0], 'TLB', 0, 'L', True)
            pdf.set_font('Helvetica', 'B', 12)
            pdf.cell(((pdf.w)-pdf.l_margin-pdf.r_margin)/2, 5, common_name, 'TRB', 1, 'R', True)
            pdf.set_x(x_padding)

            status_text = ''
            if self.dataset.config.getboolean('Atlas', 'species_accounts_show_status'):
                status_text = ''.join([designation])

            pdf.multi_cell(((pdf.w)-pdf.l_margin-pdf.r_margin), 5, status_text, 1, 'L', True)

            #compile list of last e.g. 10 records for use below
            self.dataset.cursor.execute('SELECT data.taxon, data.location, data.grid_native, data.grid_' + self.dataset.config.get('Atlas', 'distribution_unit') + ', data.date, data.decade_to, data.year_to, data.month_to, data.recorder, data.determiner, data.vc, data.grid_100m \
                                        FROM data \
                                        WHERE data.taxon = "' + item[0] + '" \
                                        ORDER BY data.year_to || data.month_to || data.day_to desc')

            indiv_taxon_data = self.dataset.cursor.fetchall()

            max_species_records_length = 900

            if self.dataset.config.getboolean('Atlas', 'species_accounts_show_descriptions'):
                max_species_records_length = max_species_records_length - len(taxa_statistics[item[0]]['description'])

            remaining_records = 0

            #there has to be a better way?
            for indiv_record in indiv_taxon_data:
                if len(taxon_recent_records) < max_species_records_length:

                    #if the determiner is different to the recorder, set the
                    #determinater (!)
                    try:
                        if indiv_record[8] != indiv_record[9]:
    
                            detees = indiv_record[9].split(',')
                            deter = ''
    
                            for deter_name in sorted(detees):
                                if deter_name != '':
                                    deter = ','.join([deter, contrib_data[deter_name.strip()]])
    
                            if deter == 'Unknown' or deter == 'Unknown Unknown' or deter == '':
                                det = ' anon.'
                            else:
                                det = ''.join([' det. ', deter[1:]])
                        else:
                            det = ''
                    except AttributeError:
                        det = ''

                    if indiv_record[4] == 'Unknown':
                        date = '[date unknown]'
                    else:
                        date = indiv_record[4]


                    if indiv_record[8] == 'Unknown' or indiv_record[8] == 'Unknown Unknown' or indiv_record[8] == '':
                        rec = ' anon.'
                    else:
                        recs = indiv_record[8].split(',')
                        rec = ''

                        for recorder_name in sorted(recs):
                            rec = ','.join([rec, contrib_data[recorder_name.strip()]])

                    #limit grid reference to 100m
                    if len(indiv_record[2]) > 8:
                        grid = indiv_record[11]
                    else:
                        grid = indiv_record[2]

                    #taxon_recent_records = ''.join([taxon_recent_records, indiv_record[1], ' (VC', str(indiv_record[10]), ') ', grid, ' ', date.replace('/', '.'), ' (', rec[1:], det, '); '])

                    #substitute parameters for record values
                    #det and loc values can be empty - to remove empty spaces
                    #we check for preceeeding and trailing spaces first with
                    #these if they are empty strings
                    current_rec = self.dataset.config.get('Atlas', 'species_accounts_latest_format')

                    if indiv_record[1] == '':
                        current_rec = current_rec.replace(' %l', indiv_record[1])
                        current_rec = current_rec.replace('%l ', indiv_record[1])
                    else:
                        current_rec = current_rec.replace('%l', indiv_record[1])

                    if det == '':
                        current_rec = current_rec.replace(' %i', det)
                        current_rec = current_rec.replace('%i ', det)
                    else:
                        current_rec = current_rec.replace('%i', det)

                    current_rec = current_rec.replace('%v', str(indiv_record[10]))
                    current_rec = current_rec.replace('%g', grid)
                    current_rec = current_rec.replace('%d', date.replace('/', '.'))
                    current_rec = current_rec.replace('%r', rec[1:])

                    #append current record to the output
                    taxon_recent_records = ''.join([taxon_recent_records, current_rec, '; '])
                else:
                    remaining_records = remaining_records + 1

            #if any records remain, add a note to the output
            if remaining_records > 0:
                remaining_records_text = ''.join([' [+ ', str(remaining_records), ' more]'])
            else:
                remaining_records_text = ''

            #taxon blurb
            pdf.set_y(y_padding+12)
            pdf.set_x(x_padding+(((pdf.w / 2)-pdf.l_margin-pdf.r_margin)+3+5))
            pdf.set_font('Helvetica', '', 10)
            pdf.set_text_color(0)
            pdf.set_fill_color(255, 255, 255)
            pdf.set_line_width(0.1)

            if len(taxa_statistics[item[0]]['description']) > 0 and self.dataset.config.getboolean('Atlas', 'species_accounts_show_descriptions'):
                pdf.set_font('Helvetica', 'B', 10)
                pdf.multi_cell((((pdf.w / 2)-pdf.l_margin-pdf.r_margin)+12), 5, ''.join([taxa_statistics[item[0]]['description'], '\n\n']), 0, 'L', False)
                pdf.set_x(x_padding+(((pdf.w / 2)-pdf.l_margin-pdf.r_margin)+3+5))

            if self.dataset.config.getboolean('Atlas', 'species_accounts_show_latest'):
                pdf.set_font('Helvetica', '', 10)
                pdf.multi_cell((((pdf.w / 2)-pdf.l_margin-pdf.r_margin)+12), 5, ''.join(['Records (most recent first): ', taxon_recent_records[:-2], '.', remaining_records_text]), 0, 'L', False)

            #chart
            if self.dataset.config.getboolean('Atlas', 'species_accounts_show_phenology'):
                chart = Chart(self.dataset, item[0])
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

            temp_map_file = tempfile.NamedTemporaryFile(dir=self.dataset.temp_dir).name
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

            pdf.image(temp_map_file, x_padding+img_x_cent, y_padding+img_y_cent, int(pdfim_width), int(pdfim_height), 'PNG')


            #map container
            pdf.set_text_color(0)
            pdf.set_fill_color(255, 255, 255)
            pdf.set_line_width(0.1)
            pdf.rect(x_padding, 10+y_padding, ((pdf.w / 2)-pdf.l_margin-pdf.r_margin)+3, ((pdf.w / 2)-pdf.l_margin-pdf.r_margin)+3)

            if self.dataset.config.getboolean('Atlas', 'species_accounts_show_statistics'):

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

                #squares
                pdf.set_y((((pdf.w / 2)-pdf.l_margin-pdf.r_margin)+7)+y_padding)
                pdf.set_x((((pdf.w / 2)-pdf.l_margin-pdf.r_margin)-15)+x_padding)
                pdf.set_font('Helvetica', '', 12)
                pdf.set_text_color(0)
                pdf.set_fill_color(255, 255, 255)
                pdf.set_line_width(0.1)
                pdf.multi_cell(18, 5, ''.join(['S ', str(taxa_statistics[item[0]]['dist_count'])]), 0, 'R', False)

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


            rownum = rownum + 1

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
            if name != 'Unknown' and name != 'Unknown Unknown':
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
        try:
            pdf.output(self.save_in,'F')
        except IOError:
            md = gtk.MessageDialog(None,
                gtk.DIALOG_DESTROY_WITH_PARENT, gtk.MESSAGE_ERROR,
                gtk.BUTTONS_CLOSE, 'Unable to write to file. This usually means it''s open - close it and try again.')
            md.run()
            md.destroy()       



class Chart(gtk.Window):

    def __init__(self, dataset, item):
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
        self.mode = 'decades'
        self.chart = None

        self.temp_filename = None

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

            self.dataset.cursor.execute('SELECT MIN(data.decade), MAX(data.decade) \
                                        FROM data \
                                        WHERE data.decade IS NOT NULL')

            datas = self.dataset.cursor.fetchall()
            
            data = []
            for decade in range(datas[0][0], datas[0][1], 10):
                try:
                    data.append((decade, self.data[decade], decade))
                except KeyError:
                    data.append((decade, 0, decade))                
            
        try:
            barchart = bar_chart.BarChart()
            barchart.grid.set_visible(False)
            barchart.set_mode(bar_chart.MODE_VERTICAL)
            
            max_val = 1

            #HACK - Draw a hidden bar with a value, so we can have an 'empty' chart            
            bar = bar_chart.Bar('', max_val, '')
            bar.set_visible(False)
            barchart.add_bar(bar)
            
            for bar_info in data:
                if max_val < bar_info[1]:
                    max_val = bar_info[1]
                bar = bar_chart.Bar(*bar_info)
                bar.set_color(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'species_accounts_phenology_colour')))
                bar._label_object.set_property('size', 16)
                bar._value_label_object.set_property('size', 16)
                bar._label_object.set_property('weight', pango.WEIGHT_BOLD)
                bar._value_label_object.set_property('weight', pango.WEIGHT_BOLD)
                bar._label_object.set_property('color', gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'species_accounts_phenology_colour')))
                bar._value_label_object.set_property('color', gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'species_accounts_phenology_colour')))
                barchart.add_bar(bar)
            
            #HACK - Draw a hidden bar with a value, so we can have an 'empty' chart            
            bar = bar_chart.Bar('', max_val, '')
            bar.set_visible(False)
            barchart.add_bar(bar)
            
            self.chart = barchart
            self.chart.set_enable_mouseover(False)
            self.get_children()[0].pack_start(self.chart, True, True, 0)
            self.chart.show()
            self.temp_filename = tempfile.NamedTemporaryFile(dir=self.dataset.temp_dir).name
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

            self.data = {}

            for decade in decades:
                self.data[decade[0]] = decade[1]

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
