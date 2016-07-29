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
import csv
import cairo
from datetime import datetime
import sys
import os
import shutil
import glob 
import pango
import copy
from subprocess import call
import ConfigParser
import json

#dipper-stda
import cfg
import initialize
import atlas
import checklist
import singlespecies
import version
import dataset

class CellRendererClickablePixbuf(gtk.CellRendererPixbuf):

    __gsignals__ = {'clicked': (gobject.SIGNAL_RUN_LAST, gobject.TYPE_NONE,
                                (gobject.TYPE_STRING,))
                   }

    def __init__(self):
        gtk.CellRendererPixbuf.__init__(self)
        self.set_property('mode', gtk.CELL_RENDERER_MODE_ACTIVATABLE)

    def do_activate(self, event, widget, path, background_area, cell_area,
                    flags):
        self.emit('clicked', path)


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
                   'save_configuration':self.save_configuration,
                   'save_configuration_as':self.save_configuration_as,
                   'switch_update_title':self.switch_update_title,
                   'open_file':self.open_file,
                   'new_file':self.new_file,
                   'add_dateband':self.add_dateband,
                   'remove_dateband':self.remove_dateband,
                   'show_sql_parser':self.show_sql_parser,
                   'show_rarity_dialog':self.show_rarity_dialog,
                   'generate':self.generate,
                   'switch_sheet':self.switch_sheet,
                  }
        self.builder.connect_signals(signals)
        self.dataset = None
        
        self.builder.get_object('treeview5').get_selection().connect('changed', self.navigation_change)
        self.builder.get_object('treeview3').get_selection().connect('changed', self.list_family_selection_change)
        self.builder.get_object('treeview2').get_selection().connect('changed', self.atlas_family_selection_change)
        self.builder.get_object('treeview7').get_selection().connect('changed', self.single_species_species_selection_change)
        
        #reset navigation
        self.pre_generate = (None, None)
        self.navigate_to = (None, None)
        
        #hide tab headings
        self.builder.get_object('notebook1').set_show_tabs(False)
        self.builder.get_object('notebook2').set_show_tabs(False)
        self.builder.get_object('notebook3').set_show_tabs(False)
        self.builder.get_object('notebook4').set_show_tabs(False)
        
        #setup the filter for the cover image selectors
        initialize.setup_image_file_chooser(self.builder.get_object('filechooserbutton1'))
        initialize.setup_image_file_chooser(self.builder.get_object('filechooserbutton5'))

        #setup navigation treeview        
        treeview = self.builder.get_object('treeview5')
        treeview.set_rules_hint(False)
        treeview.get_selection().set_mode(gtk.SELECTION_SINGLE)

        rendererText = gtk.CellRendererText()
        column = gtk.TreeViewColumn("nav page", rendererText, text=0)
        treeview.append_column(column)

        store = gtk.TreeStore(str, int, int)
        treeview.set_model(store)

        #setup stats treeview        
        treeview = self.builder.get_object('treeview11')
        treeview.set_rules_hint(False)
        treeview.get_selection().set_mode(gtk.SELECTION_SINGLE)

        rendererText = gtk.CellRendererText()
        column = gtk.TreeViewColumn("Property", rendererText, text=0)
        treeview.append_column(column)
        
        rendererText = gtk.CellRendererText()
        column = gtk.TreeViewColumn("Value", rendererText, text=1)
        treeview.append_column(column)
                
        store = gtk.ListStore(str, str)
        treeview.set_model(store)

        #setup the atlas date bands treeview
        store = gtk.TreeStore(str, str, str, int, int)
           
        treeview = self.builder.get_object('treeview6')
        treeview.set_rules_hint(False)
        treeview.get_selection().set_mode(gtk.SELECTION_SINGLE)
        treeview.set_show_expanders(False)

        liststore = gtk.ListStore(gobject.TYPE_STRING)
        
        for i in range(len(cfg.markers)):
            liststore.append([cfg.markers[i]])

        renderer = gtk.CellRendererCombo()
        renderer.set_property("model", liststore)
        renderer.set_property("text_column", 0)
        renderer.set_property("editable", True)
        renderer.set_property("has-entry", False)
        renderer.connect("changed", self.combo_cell_edited, [store, 0, liststore])
        column = gtk.TreeViewColumn("Style", renderer, text=0)        
        column.set_sizing(gtk.TREE_VIEW_COLUMN_FIXED)
        column.set_fixed_width(120)
        treeview.append_column(column)
                
        renderer = CellRendererClickableText()
        renderer.connect("clicked", self.color_cell_edited, [store, 1])
        column = gtk.TreeViewColumn("Fill", renderer, markup=1)   
        column.set_fixed_width(75) 
        treeview.append_column(column)
       
        renderer = CellRendererClickableText()
        renderer.connect("clicked", self.color_cell_edited, [store, 2])
        column = gtk.TreeViewColumn("Border", renderer, markup=2)  
        column.set_fixed_width(75)
        treeview.append_column(column)
        
        adjustment = gtk.Adjustment(0, 0, 2050, 1, 1, 0)
        renderer = gtk.CellRendererSpin()
        renderer.set_property("editable", True)
        renderer.set_property("adjustment", adjustment)
        renderer.connect("edited", self.spin_cell_edited, [store, 3])      
        column = gtk.TreeViewColumn("From", renderer, text=3)      
        column.set_sizing(gtk.TREE_VIEW_COLUMN_FIXED)
        column.set_fixed_width(75)
        treeview.append_column(column)
        
        adjustment = gtk.Adjustment(0, 0, 2050, 1, 1, 0)
        renderer = gtk.CellRendererSpin()
        renderer.set_property("editable", True)
        renderer.set_property("adjustment", adjustment)
        renderer.connect("edited", self.spin_cell_edited, [store, 4])
        column = gtk.TreeViewColumn("To", renderer, text=4)        
        column.set_sizing(gtk.TREE_VIEW_COLUMN_FIXED)
        column.set_fixed_width(75)
        treeview.append_column(column)
       
        column = gtk.TreeViewColumn('')  
        treeview.append_column(column)
        
        treeview.set_model(store)
              
        #setup the single species map species treeview
        treeview = self.builder.get_object('treeview7')
        treeview.set_rules_hint(True)
        treeview.get_selection().set_mode(gtk.SELECTION_MULTIPLE)

        rendererText = gtk.CellRendererText()
        column = gtk.TreeViewColumn("Species", rendererText, text=0)
        treeview.append_column(column)
        
        #setup combo boxes        
        initialize.setup_combo_box(self.builder.get_object('combobox8'), cfg.paper_orientation)
        initialize.setup_combo_box(self.builder.get_object('combobox9'), cfg.paper_orientation)
        initialize.setup_combo_box(self.builder.get_object('combobox4'), cfg.paper_orientation)
        initialize.setup_combo_box(self.builder.get_object('combobox10'), cfg.paper_size)
        initialize.setup_combo_box(self.builder.get_object('combobox11'), cfg.paper_size)
        initialize.setup_combo_box(self.builder.get_object('combobox7'), cfg.paper_size)
        initialize.setup_combo_box(self.builder.get_object('combobox3'), cfg.grid_resolution)
        initialize.setup_combo_box(self.builder.get_object('combobox5'), cfg.grid_resolution)
        initialize.setup_combo_box(self.builder.get_object('combobox2'), cfg.grid_resolution)
        initialize.setup_combo_box(self.builder.get_object('combobox6'), cfg.markers)
        initialize.setup_combo_box(self.builder.get_object('combobox1'), cfg.grid_resolution)
        initialize.setup_combo_box(self.builder.get_object('combobox12'), cfg.phenology_types)
        initialize.setup_combo_box(self.builder.get_object('combobox16'), cfg.backgrounds)
        initialize.setup_combo_box(self.builder.get_object('combobox13'), cfg.markers)
        initialize.setup_combo_box(self.builder.get_object('combobox14'), cfg.grid_resolution)
        initialize.setup_combo_box(self.builder.get_object('combobox15'), cfg.grid_resolution)
        
        #setup vice-county selection treeviews
        initialize.setup_vice_county_treeview(self.builder.get_object('treeview1'))
        initialize.setup_vice_county_treeview(self.builder.get_object('treeview4'))
        initialize.setup_vice_county_treeview(self.builder.get_object('treeview8'))
        
        #setup family selection treeviews
        initialize.setup_family_treeview(self.builder.get_object('treeview2'))
        initialize.setup_family_treeview(self.builder.get_object('treeview3'))

        #setup gis mapping treeviews        
        initialize.setup_mapping_layers_treeview(self.builder.get_object('alignment2'))
                
        #if we have a filename provided, try to open it
        if filename != None:
            self.open_dataset(None, filename)
            
        #show the main window
        window = self.builder.get_object('window1')
        window.show()

    def add_dateband(self, widget):
        '''Add a date band to the date band atlas treeview.'''
        model = self.builder.get_object('treeview6').get_model()
        model.append(None, ['squares', '   <span background="#797979">      </span>   ', '   <span background="#797979">      </span>   ', 1980, 2030])
        self.builder.get_object('button3').set_sensitive(True)

    def remove_dateband(self, widget):
        '''Remove a date band to the date band atlas treeview.'''
        selection = self.builder.get_object('treeview6').get_selection()
        model, selected = selection.get_selected_rows()

        iters = [model.get_iter(path) for path in selected]
        for iter in iters:
            model.remove(iter)

        if len(model) == 1:
            self.builder.get_object('button3').set_sensitive(False)
       

    def color_cell_edited(self, widget, path, userdata):
        selection = self.builder.get_object('treeview6').get_selection()

        #hacky but it works        
        pre_colour = userdata[0][path][userdata[1]].split('"')
        
        if selection.path_is_selected(path):
            dialog = gtk.ColorSelectionDialog('Pick a Colour')
            dialog.colorsel.set_current_color(gtk.gdk.color_parse(pre_colour[1]))

            response = dialog.run()
            dialog.destroy()

            if response == gtk.RESPONSE_OK:
                color = dialog.colorsel.get_current_color()   
                userdata[0][path][userdata[1]] = ''.join(['   <span background="', str(color), '">      </span>   '])
         

    def combo_cell_edited(self, widget, path, new_iter, userdata):
        userdata[0][path][userdata[1]] = userdata[2].get_value(new_iter, 0)

    def spin_cell_edited(self, widget, path, value, userdata):
        userdata[0][path][userdata[1]] = int(value)

    def toggle_cell_edited(self, widget, path, userdata):
        selection = self.builder.get_object('treeview6').get_selection()
        if selection.path_is_selected(path):
            userdata[0][path][userdata[1]] = not widget.get_active()

    def generate(self, widget):
        if self.dataset.config.get('DEFAULT', 'type') == 'Atlas':
            self.generate_atlas(widget)
        elif self.dataset.config.get('DEFAULT', 'type') == 'Checklist':
            self.generate_list(widget)
        elif self.dataset.config.get('DEFAULT', 'type') == 'Single Species':
            self.generate_single_species_map(widget)
        
            
    def navigation_change(self, widget):
        if not self.pre_generate == (None, None):
            selection = self.builder.get_object('treeview5').get_selection()
            selection.select_iter(self.pre_generate[1])
            self.pre_generate = (None, None)
    
    
        if not self.navigate_to == (None, None):
            selection = self.builder.get_object('treeview5').get_selection()
            selection.select_path(self.navigate_to)
            self.navigate_to = (None, None)
        
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
            elif main_notebook_page == 2:
                self.builder.get_object('notebook4').set_current_page(sub_notebook_page)
        except TypeError:
            pass
            

    def show_about(self, widget):
        """Show the about dialog."""
        dialog = gtk.AboutDialog()
        dialog.set_name('dipper-stda\n')
        dialog.set_comments('An atlas & checklist generator')
        dialog.set_version(version.__version__)
        dialog.set_authors(['Charlie Barnes'])
        dialog.set_website('https://github.com/charlie-barnes/dipper-stda')
        dialog.set_license("This is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation; either version 2 of the Licence, or (at your option) any later version.\n\nThis program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.\n\nYou should have received a copy of the GNU General Public License along with this program; if not, write to the Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA")
        dialog.set_wrap_license(True)
        dialog.set_property('skip-taskbar-hint', True)
        dialog.run()
        dialog.destroy()
        
    def atlas_family_selection_change(self, selection):
        model, selected = selection.get_selected_rows()
        iters = [model.get_iter(path) for path in selected]

        if len(iters) > 0:    
            self.builder.get_object('hbox4').hide()
        else:
            self.builder.get_object('hbox4').show()


    def single_species_species_selection_change(self, selection):
        model, selected = selection.get_selected_rows()
        iters = [model.get_iter(path) for path in selected]

        if len(iters) > 0:    
            self.builder.get_object('hbox9').hide()
        else:
            self.builder.get_object('hbox9').show()

    def list_family_selection_change(self, selection):
        model, selected = selection.get_selected_rows()
        iters = [model.get_iter(path) for path in selected]

        if len(iters) > 0:    
            self.builder.get_object('hbox3').hide()
        else:
            self.builder.get_object('hbox3').show()

    def update_title(self, selection):
        """Update the title of the generated document based on the selection
        in the families list, appending the families to the current input.

        - if more than 2 are selected use a range approach.
        - if 2 are selected, append both.
        - if just 1, append it.

        """

        if self.builder.get_object('notebook1').get_current_page() == 0 and self.builder.get_object('checkbutton18').get_active():
            entry = self.builder.get_object('entry3')
        elif self.builder.get_object('notebook1').get_current_page() == 1 and self.builder.get_object('checkbutton17').get_active():
            entry = self.builder.get_object('entry4')
        elif self.builder.get_object('notebook1').get_current_page() == 2 and self.builder.get_object('checkbutton4').get_active():
            entry = self.builder.get_object('entry7')
            print "mark"

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

            elif len(iters) == 0:
                entry.set_text(''.join([orig_title,
                                        ]))
        except UnboundLocalError:
            pass



    def select_all_families(widget, treeview):
        """Select all families in the selection."""
        treeview.get_selection().select_all()

    def unselect_image(widget, filechoooserbutton):
        """Clear the cover image file selection."""
        filechoooserbutton.unselect_all()

    def open_dataset(self, widget, filename=None, type=None, config=None):
        """Open a data file."""
        self.builder.get_object('notebook1').set_sensitive(False)

        #if this isn't the first dataset we've opened this session,
        #delete the preceeding temp directory
        if not self.dataset == None:
            try:
                shutil.rmtree(self.dataset.temp_dir)
            except OSError:
                pass

        self.dataset = dataset.Dataset()

        #if we are opening a cfg file
        if config is not None:
            try:
                config[0].get('DEFAULT', 'type')
                self.dataset.config.filename = config[1]
                self.dataset.config = config[0]
                    
            except AttributeError:
                self.dataset.config.read([config])
                self.dataset.config.filename = config

                if filename is not None:
                    self.dataset.set_source(filename)
        else:
            self.dataset.set_type(type)
            self.dataset.set_source(filename)
        
        self.dataset.parse()
        
        store = self.builder.get_object('treeview5').get_model()
        store.clear()
        treeselection = self.builder.get_object('treeview5').get_selection()
        
        if self.dataset.config.get('DEFAULT', 'type') == 'Atlas':
            iter = store.append(None, ['Atlas', 0, 0])
            treeselection.select_iter(iter)
            store.append(iter, ['Taxa', 0, 1])
            store.append(iter, ['Vice-counties', 0, 2])
            store.append(iter, ['Page setup', 0, 3])
            store.append(iter, ['Table of contents', 0, 4])
            store.append(iter, ['Species density map', 0, 5])
            iter = store.append(iter, ['Species accounts', 0, 6])
            store.append(iter, ['Map extents', 0, 7])
            store.append(iter, ['Date bands', 0, 8])
        elif self.dataset.config.get('DEFAULT', 'type') == 'Checklist':
            iter = store.append(None, ['Checklist', 1, 0])
            treeselection.select_iter(iter)
            store.append(iter, ['Taxa', 1, 1])
            store.append(iter, ['Vice-counties', 1, 2])
            store.append(iter, ['Page setup', 1, 3])
        elif self.dataset.config.get('DEFAULT', 'type') == 'Single Species':
            iter = store.append(None, ['Single species map', 2, 0])
            treeselection.select_iter(iter)
            store.append(iter, ['Taxa', 2, 1])
            store.append(iter, ['Vice-counties', 2, 2])
            store.append(iter, ['Page setup', 2, 3])
        
        self.builder.get_object('treeview5').expand_all()
       
        try:

            if self.dataset.data_source.read() == True:
        
                self.update_widgets()

                self.dataset.builder = self.builder
                self.builder.get_object('treeview4').set_sensitive(self.dataset.use_vcs)
                
                if self.dataset.use_vcs:
                    self.builder.get_object('label61').set_markup('<i>Data will be grouped as one if no vice-counties are selected</i>')
                    self.builder.get_object('label37').set_markup('<i>Data will be grouped as one if no vice-counties are selected</i>')
                    self.builder.get_object('label38').set_markup('<i>Data will be grouped as one if no vice-counties are selected</i>')
                    self.builder.get_object('treeview1').set_sensitive(True)
                    self.builder.get_object('treeview4').set_sensitive(True)
                    self.builder.get_object('treeview8').set_sensitive(True)
                else:
                    self.builder.get_object('label61').set_markup('<i>Vice-county information is not present in the source file</i>')
                    self.builder.get_object('label37').set_markup('<i>Vice-county information is not present in the source file</i>')
                    self.builder.get_object('label38').set_markup('<i>Vice-county information is not present in the source file</i>')
                    self.builder.get_object('treeview1').set_sensitive(False)
                    self.builder.get_object('treeview4').set_sensitive(False)
                    self.builder.get_object('treeview8').set_sensitive(False)

                while gtk.events_pending():
                    gtk.main_iteration_do(True)

                try:
                    self.builder.get_object('window1').set_title(''.join([os.path.basename(self.dataset.config.filename), ' (',  os.path.dirname(self.dataset.config.filename), ') - dipper-stda',]) )
                except AttributeError:
                    self.builder.get_object('window1').set_title(' '.join(['Unsaved', self.dataset.config.get('DEFAULT', 'type'), '-','dipper-stda']))
                    
                self.builder.get_object('notebook1').set_sensitive(True)


                self.builder.get_object('hbox1').show()
                self.builder.get_object('menuitem7').set_sensitive(True)
                self.builder.get_object('menuitem8').set_sensitive(True)
                self.builder.get_object('toolbutton5').set_sensitive(True)
                self.builder.get_object('toolbutton3').set_sensitive(True)
                
                #populate the stats panel
                treeview = self.builder.get_object('treeview11')
                store = treeview.get_model()
                store.clear()

                for stat in [['Records', self.dataset.records], ['Species', len(self.dataset.species)], ['Families',  len(self.dataset.families)], ['Earliest', self.dataset.earliest], ['Latest', self.dataset.latest], ['Recorders', self.dataset.recorders], ['Determiners', self.dataset.determiners], ['Vice-counties', len(self.dataset.vicecounties)], ['Recorded 100km squares', len(self.dataset.occupied_squares['100km'])], ['Recorded 10km squares', len(self.dataset.occupied_squares['10km'])], ['Recorded 5km squares', len(self.dataset.occupied_squares['5km'])], ['Recorded 2km squares', len(self.dataset.occupied_squares['2km'])], ['Recorded 1km squares', len(self.dataset.occupied_squares['1km'])], ['Recorded 100m squares', len(self.dataset.occupied_squares['100m'])], ['Recorded 10m squares', len(self.dataset.occupied_squares['10m'])], ['Recorded 1m squares', len(self.dataset.occupied_squares['1m'])]   ]:
                    store.append(stat)
                
                #populate the sheet combobox                
                initialize.setup_combo_box(self.builder.get_object('combobox17'), ['-- all sheets --'] + self.dataset.available_sheets)
                
                try:
                    self.builder.get_object('combobox17').set_active(self.dataset.available_sheets.index(self.dataset.config.get('DEFAULT', 'sheets'))+1)
                except ValueError:
                    self.builder.get_object('combobox17').set_active(0)
                
            
        except AttributeError as e:
            self.builder.get_object('menuitem7').set_sensitive(False)
            self.builder.get_object('menuitem8').set_sensitive(False)
            self.builder.get_object('toolbutton5').set_sensitive(False)
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

    def generate_single_species_map(self, widget):
        """Process the data file and configuration."""
        if self.dataset is not None:
            self.update_config()

            vbox = self.builder.get_object('vbox2') 
            notebook = self.builder.get_object('notebook1')           

            #we can't produce a map without any geographic entity!
            if not len(self.dataset.config.get('Single Species', 'vice-counties'))>0:            
                self.navigate_to = (2, 1)
                self.builder.get_object('treeview5').get_selection().emit("changed")
            #we can't produce a map without any families selected!
            elif not len(self.dataset.config.get('Single Species', 'Single Species'))>0:            
                self.navigate_to = (2, 0)
                self.builder.get_object('treeview5').get_selection().emit("changed")               
            else:              

                vbox.set_sensitive(False)

                dialog = gtk.FileChooserDialog('Save As...',
                                               None,
                                               gtk.FILE_CHOOSER_ACTION_SAVE,
                                               (gtk.STOCK_CANCEL, gtk.RESPONSE_CANCEL,
                                                gtk.STOCK_SAVE, gtk.RESPONSE_OK))
                dialog.set_default_response(gtk.RESPONSE_OK)
                dialog.set_do_overwrite_confirmation(True)

                dialog.set_current_folder(os.path.dirname(os.path.abspath(self.dataset.config.get('DEFAULT', 'source'))))
                dialog.set_current_name(''.join([os.path.splitext(os.path.basename(self.dataset.config.get('DEFAULT', 'source')))[0], '_map.pdf']))

                dialog.set_property('skip-taskbar-hint', True)
        
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
                    spmapobj = singlespecies.SingleSpecies(self.dataset)
                    spmapobj.save_in = output

                    spmapobj.generate_base_map()

                    spmapobj.generate()
                    
                    if sys.platform == 'linux2':
                        call(["xdg-open", output])
                    else:
                        os.startfile(output)
                    
            selection = self.builder.get_object('treeview5').get_selection()
            self.pre_generate = selection.get_selected()
            
            vbox.set_sensitive(True)
            self.builder.get_object('window1').window.set_cursor(None)



    def generate_atlas(self, widget):
        """Process the data file and configuration."""
        if self.dataset is not None:
            self.update_config()

            vbox = self.builder.get_object('vbox2') 
            notebook = self.builder.get_object('notebook1')           

            #we can't produce an atlas without any geographic entity!
            if not len(self.dataset.config.get('Atlas', 'mapping_layers'))>0:            
                self.navigate_to = (0, 1)
                self.builder.get_object('treeview5').get_selection().emit("changed")
            #we can't produce an atlas or list without any families selected!
            elif not len(self.dataset.config.get('Atlas', 'families'))>0:            
                self.navigate_to = (0, 0)
                self.builder.get_object('treeview5').get_selection().emit("changed")               
            else:              

                vbox.set_sensitive(False)

                dialog = gtk.FileChooserDialog('Save As...',
                                               None,
                                               gtk.FILE_CHOOSER_ACTION_SAVE,
                                               (gtk.STOCK_CANCEL, gtk.RESPONSE_CANCEL,
                                                gtk.STOCK_SAVE, gtk.RESPONSE_OK))
                dialog.set_default_response(gtk.RESPONSE_OK)
                dialog.set_do_overwrite_confirmation(True)

                dialog.set_current_folder(os.path.dirname(os.path.abspath(self.dataset.config.get('DEFAULT', 'source'))))
                dialog.set_current_name(''.join([os.path.splitext(os.path.basename(self.dataset.config.get('DEFAULT', 'source')))[0], '_atlas.pdf']))

                dialog.set_property('skip-taskbar-hint', True)
                
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
                    atlasobj = atlas.Atlas(self.dataset)
                    atlasobj.save_in = output

                    atlasobj.generate_base_map()

                    if self.dataset.config.getboolean('Atlas', 'species_density_map_visible'):
                        atlasobj.generate_density_map()

                    atlasobj.generate()
                    
                    if sys.platform == 'linux2':
                        call(["xdg-open", output])
                    else:
                        os.startfile(output)
                    
            selection = self.builder.get_object('treeview5').get_selection()
            #self.pre_generate = selection.get_selected()
            
            vbox.set_sensitive(True)
            self.builder.get_object('window1').window.set_cursor(None)


    def generate_list(self, widget):
        """Process the data file and configuration."""
        if self.dataset is not None:
            self.update_config()

            vbox = self.builder.get_object('vbox2') 
            notebook = self.builder.get_object('notebook1')           

            #we can't produce an atlas without any geographic entity!
            if not len(self.dataset.config.get('Checklist', 'families'))>0:            
                self.navigate_to = (1, 0)
                self.builder.get_object('treeview5').get_selection().emit("changed")               
            else:              

                vbox.set_sensitive(False)

                dialog = gtk.FileChooserDialog('Save As...',
                                               None,
                                               gtk.FILE_CHOOSER_ACTION_SAVE,
                                               (gtk.STOCK_CANCEL, gtk.RESPONSE_CANCEL,
                                                gtk.STOCK_SAVE, gtk.RESPONSE_OK))
                dialog.set_default_response(gtk.RESPONSE_OK)
                dialog.set_do_overwrite_confirmation(True)

                dialog.set_current_folder(os.path.dirname(os.path.abspath(self.dataset.config.get('DEFAULT', 'source'))))
                dialog.set_current_name(''.join([os.path.splitext(os.path.basename(self.dataset.config.get('DEFAULT', 'source')))[0], '_checklist.pdf']))

                dialog.set_property('skip-taskbar-hint', True)
                
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
                    listing = checklist.Checklist(self.dataset)
                    listing.save_in = output
                    listing.generate()
                    
                    if sys.platform == 'linux2':
                        call(["xdg-open", output])
                    else:
                        os.startfile(output)

            selection = self.builder.get_object('treeview5').get_selection()
            self.pre_generate = selection.get_selected()
                
            vbox.set_sensitive(True)
            self.builder.get_object('window1').window.set_cursor(None)

    def update_widgets(self):

        #set up the atlas gui based on config settings
        if self.dataset.config.get('DEFAULT', 'type') == 'Atlas':

            #title
            self.builder.get_object('entry3').set_text(self.dataset.config.get('Atlas', 'title'))

            #author
            self.builder.get_object('entry2').set_text(self.dataset.config.get('Atlas', 'author'))

            #cover image
            if self.dataset.config.get('Atlas', 'cover_image') == '' or self.dataset.config.get('Atlas', 'cover_image') == None:
                self.unselect_image(self.builder.get_object('filechooserbutton1'))
            else:
                self.builder.get_object('filechooserbutton1').set_filename(self.dataset.config.get('Atlas', 'cover_image'))

            #inside cover
            self.builder.get_object('textview1').get_buffer().set_text(self.dataset.config.get('Atlas', 'inside_cover').replace('\n<nl>\n','\n\n'))

            #introduction
            self.builder.get_object('textview3').get_buffer().set_text(self.dataset.config.get('Atlas', 'introduction').replace('\n<nl>\n','\n\n'))

            #bibliography
            self.builder.get_object('textview2').get_buffer().set_text(self.dataset.config.get('Atlas', 'bibliography').replace('\n<nl>\n','\n\n'))

            #distribution unit
            self.builder.get_object('combobox3').set_active(cfg.grid_resolution.index(self.dataset.config.get('Atlas', 'distribution_unit')))

            #grid line style
            self.builder.get_object('combobox1').set_active(cfg.grid_resolution.index(self.dataset.config.get('Atlas', 'grid_lines_style')))

            #grid line colour
            self.builder.get_object('colorbutton1').set_color(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'grid_lines_colour')))

            #grid line visible
            self.builder.get_object('checkbutton2').set_active(self.dataset.config.getboolean('Atlas', 'grid_lines_visible'))

            #taxa
            try:
                selected_taxa = json.loads(self.dataset.config.get('Atlas', 'families'))
            except ValueError:
                selected_taxa = []
            
            store = self.builder.get_object('treeview2').get_model()
            store.clear()
            selection = self.builder.get_object('treeview2').get_selection()

            selection.unselect_all()
            
            if self.builder.get_object('treeview2').get_realized():
                self.builder.get_object('treeview2').scroll_to_point(0,0)

            selected_taxa_iter = []
            
            
            
            
            
            
            
            for kingdom in self.dataset.kingdoms.keys():
                parent_iter = None
                
                this_iter = store.append(parent_iter, [kingdom, 'kingdom'])
                self.dataset.kingdoms[kingdom][1] = this_iter
                            
                if ['kingdom', kingdom] in selected_taxa:
                    selected_taxa_iter.append([parent_iter, this_iter])
            
            for phylum in self.dataset.phyla.keys():
                try:
                    parent_iter = self.dataset.kingdoms[self.dataset.phyla[phylum][0]][1]
                except KeyError:
                    parent_iter = None
                    
                this_iter = store.append(parent_iter, [phylum, 'phylum'])
                self.dataset.phyla[phylum][1] = this_iter
                
                if ['phylum', phylum] in selected_taxa:
                    selected_taxa_iter.append([parent_iter, this_iter])
        
            for class_ in self.dataset.classes.keys():
                try:
                    parent_iter = self.dataset.phyla[self.dataset.classes[class_][0]][1]
                except KeyError:
                    parent_iter = None

                this_iter = store.append(parent_iter, [class_, 'class_'])
                self.dataset.classes[class_][1] = this_iter

                if ['class_', class_] in selected_taxa:
                    selected_taxa_iter.append([parent_iter, this_iter])

            for order in self.dataset.orders.keys():

                try:
                    parent_iter = self.dataset.classes[self.dataset.orders[order][0]][1]
                except KeyError:
                    parent_iter = None
                    
                this_iter = store.append(parent_iter, [order, 'order_'])
                self.dataset.orders[order][1] = this_iter

                if ['order_', order] in selected_taxa:
                    selected_taxa_iter.append([parent_iter, this_iter])

            for family in self.dataset.families.keys():
                try:
                    parent_iter = self.dataset.orders[self.dataset.families[family][0]][1]
                except KeyError:
                    parent_iter = None
                    
                this_iter = store.append(parent_iter, [family, 'family'])
                self.dataset.families[family][1] = this_iter
            
                if ['family', family] in selected_taxa:
                    selected_taxa_iter.append([parent_iter, this_iter])
    
            for genus in self.dataset.genera.keys():
                try:
                    parent_iter = self.dataset.families[self.dataset.genera[genus][0]][1]
                except KeyError:
                    parent_iter = None
                    
                this_iter = store.append(parent_iter, [genus, 'genus'])
                self.dataset.genera[genus][1] = this_iter
        
                if ['genus', genus] in selected_taxa:
                    selected_taxa_iter.append([parent_iter, this_iter])
  
            for species in self.dataset.specie.keys():
                try:
                    parent_iter = self.dataset.genera[self.dataset.specie[species][0]][1]
                except KeyError:
                    parent_iter = None
                    
                this_iter = store.append(parent_iter, [species, 'taxon'])
                self.dataset.specie[species][1] = this_iter

                if ['taxon', species] in selected_taxa:
                    selected_taxa_iter.append([parent_iter, this_iter])

            for parent, child in selected_taxa_iter:
                if parent is not None:
                    self.builder.get_object('treeview2').expand_to_path(store.get_path(parent))
                selection.select_iter(child)

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
            self.builder.get_object('combobox10').set_active(cfg.paper_size.index(self.dataset.config.get('Atlas', 'paper_size')))

            #paper orientation
            self.builder.get_object('combobox8').set_active(cfg.paper_orientation.index(self.dataset.config.get('Atlas', 'orientation')))

            #coverage style
            self.builder.get_object('combobox6').set_active(cfg.markers.index(self.dataset.config.get('Atlas', 'coverage_style')))

            #coverage colour
            self.builder.get_object('colorbutton4').set_color(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'coverage_colour')))

            #coverage visible
            self.builder.get_object('checkbutton1').set_active(self.dataset.config.getboolean('Atlas', 'coverage_visible'))

            #date band overlay visible
            self.builder.get_object('checkbutton3').set_active(self.dataset.config.getboolean('Atlas', 'date_band_overlay'))

            #species density visible
            self.builder.get_object('checkbutton19').set_active(self.dataset.config.getboolean('Atlas', 'species_density_map_visible'))

            #species density background visible
            self.builder.get_object('checkbutton20').set_active(self.dataset.config.getboolean('Atlas', 'species_density_map_background_visible'))

            #species density background
            try:
                self.builder.get_object('combobox16').set_active(cfg.backgrounds.index(self.dataset.config.get('Atlas', 'species_density_map_background')))
            except ValueError:
                self.builder.get_object('combobox16').set_active(0)

            #species density style
            self.builder.get_object('combobox13').set_active(cfg.markers.index(self.dataset.config.get('Atlas', 'species_density_map_style')))

            #species density unit
            self.builder.get_object('combobox14').set_active(cfg.grid_resolution.index(self.dataset.config.get('Atlas', 'species_density_map_unit')))

            #species density map low colour
            self.builder.get_object('colorbutton12').set_color(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'species_density_map_low_colour')))

            #species density map high colour
            self.builder.get_object('colorbutton13').set_color(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'species_density_map_high_colour')))

            #species density map grid line visible
            self.builder.get_object('checkbutton21').set_active(self.dataset.config.getboolean('Atlas', 'species_density_grid_lines_visible'))

            #species density map grid line style
            self.builder.get_object('combobox15').set_active(cfg.grid_resolution.index(self.dataset.config.get('Atlas', 'species_density_map_grid_lines_style')))

            #species density map grid line colour
            self.builder.get_object('colorbutton14').set_color(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'species_density_map_grid_lines_colour')))

            #atlas date bands
            store = self.builder.get_object('treeview6').get_model()
            store.clear()
            
            for row in json.loads(self.dataset.config.get('Atlas', 'date_bands')):
                fill_colour = ''.join(['   <span background="', row[1], '" >      </span>   '])
                border_colour = ''.join(['   <span background="', row[2], '" >      </span>   '])
                store.append(None, [row[0], fill_colour, border_colour, row[3], row[4]])

            #atlas mapping layers
            try:
                initialize.setup_mapping_layers_treeview(self.builder.get_object('alignment2'), json.loads(self.dataset.config.get('Atlas', 'mapping_layers')))
            except ValueError:
                initialize.setup_mapping_layers_treeview(self.builder.get_object('alignment2'))            

            #table of contents
            self.builder.get_object('checkbutton6').set_active(self.dataset.config.getboolean('Atlas', 'toc_show_families'))
            self.builder.get_object('checkbutton9').set_active(self.dataset.config.getboolean('Atlas', 'toc_show_species_names'))
            self.builder.get_object('checkbutton10').set_active(self.dataset.config.getboolean('Atlas', 'toc_show_common_names'))
            self.builder.get_object('checkbutton1110').set_active(self.dataset.config.getboolean('Atlas', 'toc_show_index'))
            self.builder.get_object('checkbutton1111').set_active(self.dataset.config.getboolean('Atlas', 'toc_show_contributors'))

            #species accounts
            self.builder.get_object('checkbutton12').set_active(self.dataset.config.getboolean('Atlas', 'species_accounts_show_descriptions'))
            self.builder.get_object('checkbutton13').set_active(self.dataset.config.getboolean('Atlas', 'species_accounts_show_latest'))
            self.builder.get_object('checkbutton14').set_active(self.dataset.config.getboolean('Atlas', 'species_accounts_show_statistics'))
            self.builder.get_object('checkbutton16').set_active(self.dataset.config.getboolean('Atlas', 'species_accounts_show_status'))
            self.builder.get_object('checkbutton15').set_active(self.dataset.config.getboolean('Atlas', 'species_accounts_show_phenology'))
            self.builder.get_object('combobox12').set_active(cfg.phenology_types.index(self.dataset.config.get('Atlas', 'species_accounts_phenology_type')))
            self.builder.get_object('colorbutton11').set_color(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'species_accounts_phenology_colour')))
            self.builder.get_object('combobox-entry3').set_text(self.dataset.config.get('Atlas', 'species_accounts_latest_format'))
            self.builder.get_object('checkbutton5').set_active(self.dataset.config.getboolean('Atlas', 'species_accounts_voucher_status'))

            #setup species account latest format combobox
            combobox = self.builder.get_object('combobox19')
            combobox.clear()

            model = gtk.ListStore(gobject.TYPE_STRING)

            for i in ['%l (VC%v) %g %d (%r %i)', '%l %g %d', '%l (VC%v) %g %d: %r %i']:
                model.append([i])

            cell = gtk.CellRendererText()
            combobox.pack_start(cell, True)
            combobox.add_attribute(cell, 'text', 0)
            
            combobox.set_model(model)
            
            #setup species explanation map species entry
            model = gtk.ListStore(gobject.TYPE_STRING)

            for i in sorted(self.dataset.species):
                model.append([i])
            
            completion = gtk.EntryCompletion()
            completion.set_model(model)
            completion.set_text_column(0)
            
            self.builder.get_object('entry5').set_completion(completion)
            self.builder.get_object('entry5').set_text(self.dataset.config.get('Atlas', 'species_accounts_explanation_species'))

        elif self.dataset.config.get('DEFAULT', 'type') == 'Checklist':
            #set up the list gui based on config settings
            #title
            self.builder.get_object('entry4').set_text(self.dataset.config.get('Checklist', 'title'))

            #author
            self.builder.get_object('entry1').set_text(self.dataset.config.get('Checklist', 'author'))

            #cover image
            if self.dataset.config.get('Checklist', 'cover_image') == '':
                self.unselect_image(self.builder.get_object('filechooserbutton5'))
            else:
                self.builder.get_object('filechooserbutton5').set_filename(self.dataset.config.get('Checklist', 'cover_image'))

            #inside cover
            self.builder.get_object('textview4').get_buffer().set_text(self.dataset.config.get('Checklist', 'inside_cover'))

            #introduction
            self.builder.get_object('textview6').get_buffer().set_text(self.dataset.config.get('Checklist', 'introduction'))

            #distribution unit
            self.builder.get_object('combobox5').set_active(cfg.grid_resolution.index(self.dataset.config.get('Checklist', 'distribution_unit')))

            #taxa
            selected_taxa = json.loads(self.dataset.config.get('Checklist', 'families'))

            store = self.builder.get_object('treeview3').get_model()
            store.clear()
            selection = self.builder.get_object('treeview3').get_selection()
            selection.unselect_all()
            
            if self.builder.get_object('treeview3').get_realized():
                self.builder.get_object('treeview3').scroll_to_point(0,0)

            selected_taxa_iter = []

            for kingdom in self.dataset.kingdoms:
                if kingdom == '':
                    kiter = None
                else:
                    kiter = store.append(None, [kingdom, 'kingdom'])
                                
                    if ['kingdom', kingdom] in selected_taxa:
                        selected_taxa_iter.append([None, kiter])
                
                for phylum in self.dataset.phyla[kingdom]:
                    if phylum == '':
                        piter = kiter
                    else:                        
                        piter = store.append(kiter, [phylum, 'phylum'])
                                
                        if ['phylum', phylum] in selected_taxa:
                            selected_taxa_iter.append([kiter, piter])
                
                    for class_ in self.dataset.classes[phylum]:
                        if class_ == '':
                            citer = piter
                        else:
                            citer = store.append(piter, [class_, 'class_'])
                                
                            if ['class_', class_] in selected_taxa:
                                selected_taxa_iter.append([piter, citer])
                
                        for order in self.dataset.orders[class_]:
                            if order == '':
                                oiter = citer
                            else:
                                oiter = store.append(citer, [order, 'order_'])
                                
                                if ['order_', order] in selected_taxa:
                                    selected_taxa_iter.append([citer, oiter])
                
                            for family in self.dataset.families[order]:
                                if family == '':
                                    fiter = oiter
                                else:
                                    fiter = store.append(oiter, [family, 'family'])
                                
                                    if ['family', family] in selected_taxa:
                                        selected_taxa_iter.append([oiter, fiter])

                                for genus in self.dataset.genera[family]:
                                    if genus == '':
                                        giter = fiter
                                    else:
                                        giter = store.append(fiter, [genus, 'genus'])
                                
                                        if ['genus', genus] in selected_taxa:
                                            selected_taxa_iter.append([fiter, giter])
                
                                    for species in self.dataset.specie[genus]:
                                        iter = store.append(giter, [species, 'taxon'])
                                
                                        if ['taxon', species] in selected_taxa:
                                            selected_taxa_iter.append([giter, iter])

            for parent, child in selected_taxa_iter:
                if parent is not None:
                    self.builder.get_object('treeview3').expand_to_path(store.get_path(parent))
                selection.select_iter(child)
                
            self.dataset.config.set('Checklist', 'families_update_title', str(self.builder.get_object('checkbutton17').get_active()))

            #vcs
            selection = self.builder.get_object('treeview4').get_selection()

            selection.unselect_all()
            
            if self.builder.get_object('treeview4').get_realized():
                self.builder.get_object('treeview4').scroll_to_point(0,0)
                
            try:
                for vc in self.dataset.config.get('Checklist', 'vice-counties').split(','):
                    selection.select_path(int(float(vc))-1)
                self.builder.get_object('treeview4').scroll_to_cell(int(float(self.dataset.config.get('Checklist', 'vice-counties').split(',')[0]))-1)
            except ValueError:
                pass

            #paper size
            self.builder.get_object('combobox11').set_active(cfg.paper_size.index(self.dataset.config.get('Checklist', 'paper_size')))

            #paper orientation
            self.builder.get_object('combobox9').set_active(cfg.paper_orientation.index(self.dataset.config.get('Checklist', 'orientation')))

        elif self.dataset.config.get('DEFAULT', 'type') == 'Single Species':
            #set up the single species map gui based on config settings
            #title
            self.builder.get_object('entry7').set_text(self.dataset.config.get('Single Species', 'title'))

            #author
            self.builder.get_object('entry6').set_text(self.dataset.config.get('Single Species', 'author'))

            #distribution unit
            self.builder.get_object('combobox2').set_active(cfg.grid_resolution.index(self.dataset.config.get('Single Species', 'distribution_unit')))

            #species
            store = gtk.ListStore(str)
            self.builder.get_object('treeview7').set_model(store)
            selection = self.builder.get_object('treeview7').get_selection()

            selection.unselect_all()
            
            if self.builder.get_object('treeview7').get_realized():
                self.builder.get_object('treeview7').scroll_to_point(0,0)

            for species in self.dataset.species:
                iter = store.append([species])
                if species.strip() in ''.join(self.dataset.config.get('Single Species', 'Single Species')).split(','):
                    selection.select_path(store.get_path((iter)))

            model, selected = selection.get_selected_rows()
            try:
                self.builder.get_object('treeview7').scroll_to_cell(selected[0])
            except IndexError:
                pass

            self.builder.get_object('checkbutton4').set_active(self.dataset.config.getboolean('Single Species', 'species_update_title'))
            
            #vcs
            selection = self.builder.get_object('treeview8').get_selection()

            selection.unselect_all()
            
            if self.builder.get_object('treeview8').get_realized():
                self.builder.get_object('treeview8').scroll_to_point(0,0)
                
            try:
                for vc in self.dataset.config.get('Single Species', 'vice-counties').split(','):
                    selection.select_path(int(float(vc))-1)
                self.builder.get_object('treeview8').scroll_to_cell(int(float(self.dataset.config.get('Single Species', 'vice-counties').split(',')[0]))-1)
            except ValueError:
                pass

            #paper size
            self.builder.get_object('combobox7').set_active(cfg.paper_size.index(self.dataset.config.get('Single Species', 'paper_size')))

            #paper orientation
            self.builder.get_object('combobox4').set_active(cfg.paper_orientation.index(self.dataset.config.get('Single Species', 'orientation')))


    def update_config(self):

        if self.dataset.config.get('DEFAULT', 'type') == 'Atlas':
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
            self.dataset.config.set('Atlas', 'inside_cover', inside_cover.replace('\n\n','\n<nl>\n'))

            buffer = self.builder.get_object('textview3').get_buffer()
            startiter, enditer = buffer.get_bounds()
            introduction = buffer.get_text(startiter, enditer, True)
            self.dataset.config.set('Atlas', 'introduction', introduction.replace('\n\n','\n<nl>\n'))

            buffer = self.builder.get_object('textview2').get_buffer()
            startiter, enditer = buffer.get_bounds()
            bibliography = buffer.get_text(startiter, enditer, True)
            self.dataset.config.set('Atlas', 'bibliography', bibliography.replace('\n\n','\n<nl>\n'))

            self.dataset.config.set('Atlas', 'distribution_unit', self.builder.get_object('combobox3').get_active_text())

            #grab a comma delimited list of families
            selection = self.builder.get_object('treeview2').get_selection()
            model, selected = selection.get_selected_rows()
            iters = [model.get_iter(path) for path in selected]

            families = []

            for iter in iters:
                families.append([model.get_value(iter, 1), model.get_value(iter, 0)])

            self.dataset.config.set('Atlas', 'families', json.dumps(families))
            self.dataset.config.set('Atlas', 'families_update_title', str(self.builder.get_object('checkbutton18').get_active()))

            #grab a comma delimited list of vcs
            selection = self.builder.get_object('treeview1').get_selection()
            model, selected = selection.get_selected_rows()
            iters = [model.get_iter(path) for path in selected]

            vcs = ''

            for iter in iters:
                vcs = ','.join([vcs, model.get_value(iter, 0)])

            self.dataset.config.set('Atlas', 'vice-counties', vcs[1:])

            #atlas date bands
            date_bands = []

            for row in self.builder.get_object('treeview6').get_model():
                fill_colour = row[1].split('"')[1]
                border_colour = row[2].split('"')[1]
                date_bands.append([row[0], fill_colour, border_colour,  row[3],  row[4]])

            self.dataset.config.set('Atlas', 'date_bands', json.dumps(date_bands))

            #mapping layers
            mapping_layers = []
            
            notebook = self.builder.get_object('alignment2').get_child()
            
            for page in range(0,notebook.get_n_pages()):
                treeview = notebook.get_nth_page(page).get_child()

                selection = treeview.get_selection()
                model, selected = selection.get_selected_rows()
                iters = [model.get_iter(path) for path in selected]
                
                for iter in iters:
                    category = notebook.get_tab_label(treeview.get_parent()).get_text()
                    
                    fill_colour = model.get_value(iter, 1).split('"')[1]
                    border_colour = model.get_value(iter, 2).split('"')[1]
                    
                    mapping_layers.append([category, model.get_value(iter, 0), fill_colour, border_colour])
            
            self.dataset.config.set('Atlas', 'mapping_layers', json.dumps(mapping_layers))

            #coverage
            self.dataset.config.set('Atlas', 'coverage_visible', str(self.builder.get_object('checkbutton1').get_active()))
            self.dataset.config.set('Atlas', 'coverage_style', self.builder.get_object('combobox6').get_active_text())
            self.dataset.config.set('Atlas', 'coverage_colour', str(self.builder.get_object('colorbutton4').get_color()))

            #date band overlay
            self.dataset.config.set('Atlas', 'date_band_overlay', str(self.builder.get_object('checkbutton3').get_active()))

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
            self.dataset.config.set('Atlas', 'toc_show_index', str(self.builder.get_object('checkbutton1110').get_active()))
            self.dataset.config.set('Atlas', 'toc_show_contributors', str(self.builder.get_object('checkbutton1111').get_active()))
            
            #species accounts
            self.dataset.config.set('Atlas', 'species_accounts_show_descriptions', str(self.builder.get_object('checkbutton12').get_active()))
            self.dataset.config.set('Atlas', 'species_accounts_show_latest', str(self.builder.get_object('checkbutton13').get_active()))
            self.dataset.config.set('Atlas', 'species_accounts_show_statistics', str(self.builder.get_object('checkbutton14').get_active()))
            self.dataset.config.set('Atlas', 'species_accounts_show_status', str(self.builder.get_object('checkbutton16').get_active()))
            self.dataset.config.set('Atlas', 'species_accounts_show_phenology', str(self.builder.get_object('checkbutton15').get_active()))
            self.dataset.config.set('Atlas', 'species_accounts_phenology_type', str(self.builder.get_object('combobox12').get_active_text()))
            self.dataset.config.set('Atlas', 'species_accounts_phenology_colour', str(self.builder.get_object('colorbutton11').get_color()))
            self.dataset.config.set('Atlas', 'species_accounts_latest_format', self.builder.get_object('combobox-entry3').get_text())
            self.dataset.config.set('Atlas', 'species_accounts_voucher_status', str(self.builder.get_object('checkbutton5').get_active()))
            self.dataset.config.set('Atlas', 'species_accounts_explanation_species', self.builder.get_object('entry5').get_text())

        elif self.dataset.config.get('DEFAULT', 'type') == 'Checklist':
            #list
            self.dataset.config.set('Checklist', 'title', self.builder.get_object('entry4').get_text())
            self.dataset.config.set('Checklist', 'author', self.builder.get_object('entry1').get_text())

            try:
                self.dataset.config.set('Checklist', 'cover_image', self.builder.get_object('filechooserbutton5').get_filename())
            except TypeError:
                self.dataset.config.set('Checklist', 'cover_image', '')

            buffer = self.builder.get_object('textview4').get_buffer()
            startiter, enditer = buffer.get_bounds()
            inside_cover = buffer.get_text(startiter, enditer, True)
            self.dataset.config.set('Checklist', 'inside_cover', inside_cover)

            buffer = self.builder.get_object('textview6').get_buffer()
            startiter, enditer = buffer.get_bounds()
            introduction = buffer.get_text(startiter, enditer, True)

            self.dataset.config.set('Checklist', 'introduction', introduction)
            self.dataset.config.set('Checklist', 'distribution_unit', self.builder.get_object('combobox5').get_active_text())

            #grab a comma delimited list of families
            selection = self.builder.get_object('treeview3').get_selection()
            model, selected = selection.get_selected_rows()
            iters = [model.get_iter(path) for path in selected]

            families = []

            for iter in iters:
                families.append([model.get_value(iter, 1), model.get_value(iter, 0)])

            self.dataset.config.set('Checklist', 'families', json.dumps(families))

            #grab a comma delimited list of vcs
            selection = self.builder.get_object('treeview4').get_selection()
            model, selected = selection.get_selected_rows()
            iters = [model.get_iter(path) for path in selected]

            vcs = ''

            for iter in iters:
                vcs = ','.join([vcs, model.get_value(iter, 0)])

            self.dataset.config.set('Checklist', 'vice-counties', vcs[1:])

            #page setup
            self.dataset.config.set('Checklist', 'paper_size', self.builder.get_object('combobox11').get_active_text())
            self.dataset.config.set('Checklist', 'orientation', self.builder.get_object('combobox9').get_active_text())
            
        elif self.dataset.config.get('DEFAULT', 'type') == 'Single Species':
            #single species map
            self.dataset.config.set('Single Species', 'title', self.builder.get_object('entry7').get_text())
            self.dataset.config.set('Single Species', 'author', self.builder.get_object('entry6').get_text())   
            self.dataset.config.set('Single Species', 'distribution_unit', self.builder.get_object('combobox2').get_active_text())     


            #grab a comma delimited list of species
            selection = self.builder.get_object('treeview7').get_selection()
            model, selected = selection.get_selected_rows()
            iters = [model.get_iter(path) for path in selected]

            species = ''

            for iter in iters:
                species = ','.join([species, model.get_value(iter, 0)])

            self.dataset.config.set('Single Species', 'Single Species', species[1:])
            
            self.dataset.config.set('Single Species', 'species_update_title', str(self.builder.get_object('checkbutton4').get_active()))

            #grab a comma delimited list of vcs
            selection = self.builder.get_object('treeview8').get_selection()
            model, selected = selection.get_selected_rows()
            iters = [model.get_iter(path) for path in selected]

            vcs = ''

            for iter in iters:
                vcs = ','.join([vcs, model.get_value(iter, 0)])

            self.dataset.config.set('Single Species', 'vice-counties', vcs[1:])

            #page setup
            self.dataset.config.set('Single Species', 'paper_size', self.builder.get_object('combobox7').get_active_text())
            self.dataset.config.set('Single Species', 'orientation', self.builder.get_object('combobox4').get_active_text())


    def switch_update_title(self, widget):
        if self.dataset.config.get('DEFAULT', 'type') == 'Atlas':
            self.dataset.config.set('Atlas', 'families_update_title', str(self.builder.get_object('checkbutton18').get_active()))
        elif self.dataset.config.get('DEFAULT', 'type') == 'Checklist':
            self.dataset.config.set('Checklist', 'families_update_title', str(self.builder.get_object('checkbutton17').get_active()))
        elif self.dataset.config.get('DEFAULT', 'type') == 'Single Species':
            self.dataset.config.set('Single Species', 'species_update_title', str(self.builder.get_object('checkbutton4').get_active()))

    def switch_sheet(self, widget):
        index = widget.get_active()
        store = widget.get_model()

        if self.dataset.config.get('DEFAULT', 'sheets') != store[index][0]:
            self.dataset.config.set('DEFAULT', 'sheets', store[index][0])

            #self.open_dataset(widget, filename=None, type=None, config=config_file)
            self.open_dataset(widget, self.dataset.config.get('DEFAULT', 'source'), self.dataset.config.get('DEFAULT', 'type'), [self.dataset.config, self.dataset.config.filename])

    def new_file_set(self, widget):
        '''Set the widget sensitivity to true'''
        widget.set_sensitive(True)
        
    def change_type(self, widget):
        '''Set the data source type'''  
        if widget.get_children()[0].get_active() == 1:
            widget.get_children()[1].set_visible(False)
            widget.get_children()[2].set_visible(True)
            widget.get_children()[3].set_visible(False)
        elif widget.get_children()[0].get_active() == 2:
            widget.get_children()[1].set_visible(True)
            widget.get_children()[2].set_visible(False)
            widget.get_children()[3].set_visible(False)
        else:    
            widget.get_children()[1].set_visible(False)
            widget.get_children()[2].set_visible(False)
            widget.get_children()[3].set_visible(True)
    
    def new_file(self, widget):
        '''Create a new file'''

        builder = gtk.Builder()
        builder.add_from_file('./gui/new_dialog.glade')
        dialog = builder.get_object('dialog')
          
        signals = {
                   'new_file_set':self.new_file_set,
                   'change_type':self.change_type,
                  }
        builder.connect_signals(signals)

        combobox = builder.get_object('combobox2')
        liststore = gtk.ListStore(str)
        cell = gtk.CellRendererText()
        combobox.pack_start(cell)
        combobox.add_attribute(cell, 'text', 0)
        combobox.set_model(liststore)        
        
        for option in ['File', 'MapMate', 'Recorder 6']:
            liststore.append([option])

        combobox.set_active(0)

        combobox = builder.get_object('combobox1')
        liststore = gtk.ListStore(str)
        cell = gtk.CellRendererText()
        combobox.pack_start(cell)
        combobox.add_attribute(cell, 'text', 0)
        combobox.set_model(liststore)         
        
        for option in ['Atlas', 'Checklist', 'Single Species']:
            liststore.append([option])

        combobox.set_active(0) 
        
        filechooserbutton = builder.get_object('filechooserbutton1')
        
        #filter for the data file filechooser
        filter = gtk.FileFilter()
        filter.set_name("Supported data files")
        filter.add_pattern("*.xls")
        filter.add_mime_type("application/vnd.ms-excel")
        filter.add_pattern("*.mdb")
        filter.add_mime_type("application/x-msaccess")
        filechooserbutton.add_filter(filter)
        filechooserbutton.set_filter(filter)
                
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
                filechooserbutton.add_filter(filter)
        except OSError:
            print "ssconvert isn't available - you're limited to reading XLS files. Install Gnumeric to make use of ssconvert."
            pass
        
        response = dialog.run()

        if response == 1:
            self.open_dataset(widget, filechooserbutton.get_filename(), combobox.get_active_text())

        dialog.destroy()        
    
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
        filter.set_name("dipper configuration files")
        filter.add_pattern("*.dcfg")
        dialog.add_filter(filter)
        dialog.set_filter(filter)
        
        response = dialog.run()
        config_file = dialog.get_filename()
        dialog.destroy()

        if response == -5:
            try:
                self.open_dataset(widget, filename=None, type=None, config=config_file)
            except IOError as e:
                self.dataset = None
                
                self.builder.get_object('menuitem7').set_sensitive(False)
                self.builder.get_object('menuitem8').set_sensitive(False)
                self.builder.get_object('toolbutton5').set_sensitive(False)
                self.builder.get_object('toolbutton3').set_sensitive(False)
                md = gtk.MessageDialog(None,
                    gtk.DIALOG_DESTROY_WITH_PARENT, gtk.MESSAGE_ERROR,
                    gtk.BUTTONS_OK, ''.join([str(e), '\n\nPlease locate it.']))
                md.run()
                md.destroy()
                    
                #if the file is missing, perhaps we moved it?
                tryagaindialog = gtk.FileChooserDialog('Locate missing data source...',
                               None,
                               gtk.FILE_CHOOSER_ACTION_OPEN,
                               (gtk.STOCK_CANCEL, gtk.RESPONSE_CANCEL,
                                gtk.STOCK_OPEN, gtk.RESPONSE_OK))
                tryagaindialog.set_default_response(gtk.RESPONSE_OK)
        
                #filter for the data file filechooser
                filter = gtk.FileFilter()
                filter.set_name("Supported data files")
                filter.add_pattern("*.xls")
                filter.add_mime_type("application/vnd.ms-excel")
                filter.add_pattern("*.mdb")
                filter.add_mime_type("application/x-msaccess")
                tryagaindialog.add_filter(filter)
                tryagaindialog.set_filter(filter)
                        
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
                        tryagaindialog.add_filter(filter)
                except OSError:
                    print "ssconvert isn't available - you're limited to reading XLS files. Install Gnumeric to make use of ssconvert."
                    pass
                
                response = tryagaindialog.run()
                datasource = tryagaindialog.get_filename()
                tryagaindialog.destroy()            
                         
                try:
                    self.open_dataset(widget, filename=datasource, type=None, config=config_file)
                except  IOError as e:
                    self.builder.get_object('menuitem7').set_sensitive(False)
                    self.builder.get_object('menuitem8').set_sensitive(False)
                    self.builder.get_object('toolbutton5').set_sensitive(False)
                    self.builder.get_object('toolbutton3').set_sensitive(False)
                    md = gtk.MessageDialog(None,
                        gtk.DIALOG_DESTROY_WITH_PARENT, gtk.MESSAGE_ERROR,
                        gtk.BUTTONS_CLOSE, ''.join(['Unable to open data file: ', str(e)]))
                    md.run()
                    md.destroy()



    def save_configuration(self, widget):
        '''Save a configuration file based on the current settings'''

        self.update_config()

        try:
            #write the config file
            with open(self.dataset.config.filename, 'wb') as configfile:
                self.dataset.config.write(configfile)
        except TypeError:
            dialog = gtk.FileChooserDialog('Save...',
                                           None,
                                           gtk.FILE_CHOOSER_ACTION_SAVE,
                                           (gtk.STOCK_CANCEL, gtk.RESPONSE_CANCEL,
                                            gtk.STOCK_SAVE, gtk.RESPONSE_OK))
            dialog.set_default_response(gtk.RESPONSE_OK)
            dialog.set_current_folder(os.path.dirname(os.path.abspath(self.dataset.config.get('DEFAULT', 'source'))))
            dialog.set_current_name(''.join(['Unsaved ', self.dataset.config.get('DEFAULT', 'type'), '.dcfg']))
            dialog.set_do_overwrite_confirmation(True)

            filter = gtk.FileFilter()
            filter.set_name("dipper configuration files")
            filter.add_pattern("*.dcfg")
            dialog.add_filter(filter)

            dialog.set_property('skip-taskbar-hint', True)
                
            response = dialog.run()

            self.dataset.config.filename = dialog.get_filename()
            dialog.destroy()
            
            if response == -5:
                with open(self.dataset.config.filename, 'wb') as configfile:
                    self.dataset.config.write(configfile)
                                          
        try:
            self.builder.get_object('window1').set_title(''.join([os.path.basename(self.dataset.config.filename), ' (',  os.path.dirname(self.dataset.config.filename), ') - dipper-stda',]) )
        except AttributeError:
            self.builder.get_object('window1').set_title(' '.join(['Unsaved', self.dataset.config.get('DEFAULT', 'type'), '-','dipper-stda']))


    def save_configuration_as(self, widget):
        '''Save a new configuration file based on the current settings'''

        self.update_config()

        dialog = gtk.FileChooserDialog('Save As...',
                                       None,
                                       gtk.FILE_CHOOSER_ACTION_SAVE,
                                       (gtk.STOCK_CANCEL, gtk.RESPONSE_CANCEL,
                                        gtk.STOCK_SAVE, gtk.RESPONSE_OK))
        dialog.set_default_response(gtk.RESPONSE_OK)
        dialog.set_current_folder(os.path.dirname(os.path.abspath(self.dataset.config.get('DEFAULT', 'source'))))
        dialog.set_current_name(''.join(['Unsaved ', self.dataset.config.get('DEFAULT', 'type'), '.dcfg']))
        dialog.set_do_overwrite_confirmation(True)

        filter = gtk.FileFilter()
        filter.set_name("dipper configuration files")
        filter.add_pattern("*.dcfg")
        dialog.add_filter(filter)

        dialog.set_property('skip-taskbar-hint', True)
            
        response = dialog.run()

        self.dataset.config.filename = dialog.get_filename()
        dialog.destroy()
        
        if response == -5:
            with open(self.dataset.config.filename, 'wb') as configfile:
                self.dataset.config.write(configfile)
                                          
        try:
            self.builder.get_object('window1').set_title(''.join([os.path.basename(self.dataset.config.filename), ' (',  os.path.dirname(self.dataset.config.filename), ') - dipper-stda',]) )
        except AttributeError:
            self.builder.get_object('window1').set_title(' '.join(['Unsaved', self.dataset.config.get('DEFAULT', 'type'), '-','dipper-stda']))
            
            
    def show_sql_parser(self, widget):
        builder = gtk.Builder()
        builder.add_from_file('./gui/sql_parser.glade')

        signals = {
                   'sql_parser_exit':self.sql_parser_exit,
                   'sql_parser_copy':self.sql_parser_copy,
                   'sql_parser_execute':self.sql_parser_execute,
                  }
        builder.connect_signals(signals)
        
        window = builder.get_object('window1')

        window.show()

  

    def sql_parser_exit(self, window, event=None):
        if window is not None:
            window.destroy()

    def sql_parser_copy(self, treeview):
        """Copy the contents of the textview to the clipboard."""
        
        deselect = False
        
        clipboard = gtk.clipboard_get(gtk.gdk.SELECTION_CLIPBOARD)
        treeselection = treeview.get_selection()
        (model, paths) = treeselection.get_selected_rows()

        if len(paths) < 1:
            deselect = True
            treeselection.select_all()
            (model, paths) = treeselection.get_selected_rows()
                    
        text = []

        ncolumns = range(model.get_n_columns())
        columns = []

        for column in treeview.get_columns():
            columns.append(column.get_title())
            
        text.append('\t'.join(columns))

        for path in paths:
            line = []
            for val in range(0, model.get_n_columns()):
                string = model.get_value(model.get_iter(path), val)
                if string is None:
                    string = ''
                line.append(string)
            text.append('\t'.join(line))

        clipboard.set_text('\n'.join(text))

        if deselect:           
            treeselection.unselect_all()

    def sql_parser_execute(self, treeview):
                                             
        self.dataset.cursor.execute('''select * from data''')
        itemlist = self.dataset.cursor.fetchall()


        self.dataset.cursor.execute('''PRAGMA table_info(data)''')

        colnames = [ i[1] for i in self.dataset.cursor.fetchall() ]
        self.dataset.cursor.execute('''PRAGMA table_info(data)''')

        coltypes = [ i[2] for i in self.dataset.cursor.fetchall() ]
        for index, item in enumerate(coltypes):
            if (item == 'TEXT'):
                coltypes[index] = str
            elif (item == 'NUMERIC'):
                coltypes[index] = str

        store = gtk.ListStore(*coltypes)
        for act in itemlist:
            store.append(act)

        for index, item in enumerate(colnames):
            rendererText = gtk.CellRendererText()
            column = gtk.TreeViewColumn(item, rendererText, text=index)
            column.set_sort_column_id(index)
            treeview.append_column(column)                                     

        treeview.get_selection().set_mode(gtk.SELECTION_MULTIPLE)
        treeview.set_model(store)




    def show_rarity_dialog(self, widget):
        builder = gtk.Builder()
        builder.add_from_file('./gui/rarity_dialog.glade')

        signals = {
                   'rarity_dialog_exit':self.rarity_dialog_exit,
                   'rarity_dialog_copy':self.rarity_dialog_copy,
                   'rarity_dialog_execute':self.rarity_dialog_execute,
                  }
        builder.connect_signals(signals)
        
        window = builder.get_object('dialog1')
        treeview = builder.get_object('treeview1')

        treeview.set_headers_visible(True)
                        
        cell = gtk.CellRendererText()
        column = gtk.TreeViewColumn('Taxon', cell, text=0)
        column.set_resizable(True)
        column.set_sort_column_id(0)
        treeview.append_column(column)
                        
        cell = gtk.CellRendererText()
        column = gtk.TreeViewColumn('Category', cell, text=1)
        column.set_resizable(False)
        column.set_sort_column_id(1)
        treeview.append_column(column)
                        
        cell = gtk.CellRendererText()
        column = gtk.TreeViewColumn('Sort Order', cell, text=2)
        column.set_resizable(False)
        column.set_visible(False)
        column.set_sort_column_id(2)
        treeview.append_column(column)

        window.show()        

    def rarity_dialog_exit(self, window, event=None):
        if window is not None:
            window.destroy()


    def rarity_dialog_copy(self, treeview):
        """Copy the contents of the textview to the clipboard."""
        
        deselect = False
        
        clipboard = gtk.clipboard_get(gtk.gdk.SELECTION_CLIPBOARD)
        treeselection = treeview.get_selection()
        (model, paths) = treeselection.get_selected_rows()

        if len(paths) < 1:
            deselect = True
            treeselection.select_all()
            (model, paths) = treeselection.get_selected_rows()
                    
        text = []

        ncolumns = range(model.get_n_columns())
        columns = []

        for column in treeview.get_columns():
            columns.append(column.get_title())
            
        text.append('\t'.join(columns))

        for path in paths:
            line = []
            for val in range(0, model.get_n_columns()):
                string = model.get_value(model.get_iter(path), val)
                if string is None:
                    string = ''
                line.append(string)
            text.append('\t'.join(line))

        clipboard.set_text('\n'.join(text))

        if deselect:           
            treeselection.unselect_all()

    def rarity_dialog_execute(self, treeview):    
        scrolledwindow1 = treeview.get_parent()
        dialog_vbox1 = scrolledwindow1.get_parent()
        grid = dialog_vbox1.get_children()[0]
        
        category_1_label = grid.get_children()[5].get_text()
        category_2_label = grid.get_children()[4].get_text()
        category_3_label = grid.get_children()[3].get_text()
        category_4_label = grid.get_children()[2].get_text()
        category_5_label = grid.get_children()[1].get_text()

        category_1_min = grid.get_children()[15].get_value()
        category_2_min = grid.get_children()[13].get_value()
        category_3_min = grid.get_children()[11].get_value()
        category_4_min = grid.get_children()[9].get_value()
        category_5_min = grid.get_children()[7].get_value()

        category_1_max = grid.get_children()[14].get_value()
        category_2_max = grid.get_children()[12].get_value()
        category_3_max = grid.get_children()[10].get_value()
        category_4_max = grid.get_children()[8].get_value()
        category_5_max = grid.get_children()[6].get_value()   
        
        max_age = 1980
                            
        self.dataset.cursor.execute('SELECT COUNT(DISTINCT(grid_2km)) \
                               FROM data \
                               WHERE data.year >= ' + str(max_age))

        data = self.dataset.cursor.fetchall()  
        total_coverage = float(data[0][0])
                   
        data = self.dataset.cursor.fetchall()  

        model = gtk.TreeStore(str, str, str)
        
        treeview.set_model(model)

        
        #grab the most widespread taxa from the last 30 years
        self.dataset.cursor.execute('SELECT COUNT(DISTINCT(grid_2km)) AS count \
                               FROM data \
                               WHERE data.year >= ' + str(max_age) + ' \
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
            if taxon[1] < max_age:## move this to the SQL
                model.append(None, [taxon[0], '[old]', taxon[2]])
                   
        #get the taxa that have been recorded in the last 30 years
        self.dataset.cursor.execute('SELECT data.taxon, COUNT(DISTINCT(grid_2km)), species_data.sort_order \
                               FROM data \
                               JOIN species_data ON data.taxon = species_data.taxon \
                               WHERE data.year >= ' + str(datetime.now().year-30) + ' \
                               GROUP BY data.taxon, species_data.sort_order')
                        
        data = self.dataset.cursor.fetchall() 
        
                         
        for taxon in data:
            percent = (float(taxon[1])/total_coverage)*100

            if percent > category_1_min and percent <= category_1_max:
                status = category_1_label
            elif percent > category_2_min and percent <= category_2_max:
                status = category_2_label
            elif percent > category_3_min and percent <= category_3_max:
                status = category_3_label
            elif percent > category_4_min and percent <= category_4_max:
                status = category_4_label
            elif percent > category_5_min and percent <= category_5_max:
                status = category_5_label
            else:
                status = ''
                
            self.dataset.cursor.execute('UPDATE species_data SET local_status=? WHERE taxon=?', (status, taxon[0]))        
            self.dataset.connection.commit()

            model.append(None, [taxon[0], status, taxon[2]])                     

        treeview.get_selection().set_mode(gtk.SELECTION_MULTIPLE)


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
