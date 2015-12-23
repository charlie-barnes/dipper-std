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
                   'save_config_as':self.save_config_as,
                   'save_configuration':self.save_configuration,
                   'load_config':self.load_config,
                   'switch_update_title':self.switch_update_title,
                   'open_file':self.open_file,
                   'navigation_change':self.navigation_change,
                   'list_family_selection_change':self.list_family_selection_change,
                   'atlas_family_selection_change':self.atlas_family_selection_change,
                   'single_species_species_selection_change':self.single_species_species_selection_change,
                   'add_dateband':self.add_dateband,
                   'remove_dateband':self.remove_dateband,
                  }
        self.builder.connect_signals(signals)
        self.dataset = None
        
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
        
        renderer_pix = CellRendererClickablePixbuf()
        column = gtk.TreeViewColumn("icon", renderer_pix, stock_id=3)
        treeview.append_column(column)        
        renderer_pix.connect('clicked', self.generate_from_treeview)
        
        treeselection = treeview.get_selection()
        
        store = gtk.TreeStore(str, int, int, str)
        self.builder.get_object('treeview5').set_model(store)
        iter = store.append(None, ['Atlas', 0, 0, gtk.STOCK_SAVE])
        treeselection.select_iter(iter)
        store.append(iter, ['Families', 0, 1, None])
        store.append(iter, ['Vice-counties', 0, 2, None])
        store.append(iter, ['Page Setup', 0, 3, None])
        store.append(iter, ['Table of Contents', 0, 4, None])
        store.append(iter, ['Species Density Map', 0, 5, None])
        iter = store.append(iter, ['Species Accounts', 0, 6, None])
        store.append(iter, ['Mapping', 0, 7, None])
        store.append(iter, ['Date bands', 0, 8, None])
        iter = store.append(None, ['Checklist', 1, 0, gtk.STOCK_SAVE])
        store.append(iter, ['Families', 1, 1, None])
        store.append(iter, ['Vice-counties', 1, 2, None])
        store.append(iter, ['Page Setup', 1, 3, None])
        iter = store.append(None, ['Single Species Map', 2, 0, gtk.STOCK_SAVE])
        store.append(iter, ['Species', 2, 1, None])
        store.append(iter, ['Vice-counties', 2, 2, None])
        store.append(iter, ['Page Setup', 2, 3, None])
        
        treeview.expand_all()

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

    def generate_from_treeview(self, widget, path):
        if path == '0':
            self.generate_atlas(widget)
        elif path == '1':
            self.generate_list(widget)
        elif path == '2':
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
        dialog.set_name('Atlas & Checklist Generator\n')
        dialog.set_version(''.join(['dipper-stda ', version.__version__]))
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
            self.dataset = dataset.Dataset(widget.get_filename())
        except AttributeError:
            self.dataset = dataset.Dataset(filename)

        try:

            if self.dataset.data_source.read() == True:

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
                self.builder.get_object('menuitem7').set_sensitive(True)
                self.builder.get_object('imagemenuitem3').set_sensitive(True)
                self.builder.get_object('imagemenuitem4').set_sensitive(True)
                self.builder.get_object('toolbutton5').set_sensitive(True)
                self.builder.get_object('toolbutton3').set_sensitive(True)
        except AttributeError as e:
            self.builder.get_object('menuitem7').set_sensitive(False)
            self.builder.get_object('imagemenuitem3').set_sensitive(False)
            self.builder.get_object('imagemenuitem4').set_sensitive(False)
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
            if not len(self.dataset.config.get('Species', 'vice-counties'))>0:            
                self.navigate_to = (2, 1)
                self.builder.get_object('treeview5').get_selection().emit("changed")
            #we can't produce a map without any families selected!
            elif not len(self.dataset.config.get('Species', 'species'))>0:            
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

                dialog.set_current_folder(os.path.dirname(os.path.abspath(self.dataset.filename)))
                dialog.set_current_name(''.join([os.path.splitext(os.path.basename(self.dataset.config.filename))[0], '_map.pdf']))

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
            if not len(self.dataset.config.get('Atlas', 'vice-counties'))>0:            
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
            self.pre_generate = selection.get_selected()
            
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
        self.builder.get_object('combobox3').set_active(cfg.grid_resolution.index(self.dataset.config.get('Atlas', 'distribution_unit')))

        #grid line style
        self.builder.get_object('combobox1').set_active(cfg.grid_resolution.index(self.dataset.config.get('Atlas', 'grid_lines_style')))

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

        #species accounts
        self.builder.get_object('checkbutton12').set_active(self.dataset.config.getboolean('Atlas', 'species_accounts_show_descriptions'))
        self.builder.get_object('checkbutton13').set_active(self.dataset.config.getboolean('Atlas', 'species_accounts_show_latest'))
        self.builder.get_object('checkbutton14').set_active(self.dataset.config.getboolean('Atlas', 'species_accounts_show_statistics'))
        self.builder.get_object('checkbutton16').set_active(self.dataset.config.getboolean('Atlas', 'species_accounts_show_status'))
        self.builder.get_object('checkbutton15').set_active(self.dataset.config.getboolean('Atlas', 'species_accounts_show_phenology'))
        self.builder.get_object('combobox12').set_active(cfg.phenology_types.index(self.dataset.config.get('Atlas', 'species_accounts_phenology_type')))
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
        self.builder.get_object('combobox5').set_active(cfg.grid_resolution.index(self.dataset.config.get('List', 'distribution_unit')))

        #families
        store = self.builder.get_object('treeview3').get_model()
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
        self.builder.get_object('combobox11').set_active(cfg.paper_size.index(self.dataset.config.get('List', 'paper_size')))

        #paper orientation
        self.builder.get_object('combobox9').set_active(cfg.paper_orientation.index(self.dataset.config.get('List', 'orientation')))

        #set up the single species map gui based on config settings
        #title
        self.builder.get_object('entry7').set_text(self.dataset.config.get('Species', 'title'))

        #author
        self.builder.get_object('entry6').set_text(self.dataset.config.get('Species', 'author'))

        #distribution unit
        self.builder.get_object('combobox2').set_active(cfg.grid_resolution.index(self.dataset.config.get('Species', 'distribution_unit')))

        #species
        store = gtk.ListStore(str)
        self.builder.get_object('treeview7').set_model(store)
        selection = self.builder.get_object('treeview7').get_selection()

        selection.unselect_all()
        
        if self.builder.get_object('treeview7').get_realized():
            self.builder.get_object('treeview7').scroll_to_point(0,0)

        for species in self.dataset.species:
            iter = store.append([species])
            if species.strip() in ''.join(self.dataset.config.get('Species', 'species')).split(','):
                selection.select_path(store.get_path((iter)))

        model, selected = selection.get_selected_rows()
        try:
            self.builder.get_object('treeview7').scroll_to_cell(selected[0])
        except IndexError:
            pass

        self.builder.get_object('checkbutton4').set_active(self.dataset.config.getboolean('Species', 'species_update_title'))
        
        #vcs
        selection = self.builder.get_object('treeview8').get_selection()

        selection.unselect_all()
        
        if self.builder.get_object('treeview8').get_realized():
            self.builder.get_object('treeview8').scroll_to_point(0,0)
            
        try:
            for vc in self.dataset.config.get('Species', 'vice-counties').split(','):
                selection.select_path(int(float(vc))-1)
            self.builder.get_object('treeview8').scroll_to_cell(int(float(self.dataset.config.get('Species', 'vice-counties').split(',')[0]))-1)
        except ValueError:
            pass

        #paper size
        self.builder.get_object('combobox7').set_active(cfg.paper_size.index(self.dataset.config.get('Species', 'paper_size')))

        #paper orientation
        self.builder.get_object('combobox4').set_active(cfg.paper_orientation.index(self.dataset.config.get('Species', 'orientation')))


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
                mapping_layers.append([category, model.get_value(iter, 0)])
        
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

        #species accounts
        self.dataset.config.set('Atlas', 'species_accounts_show_descriptions', str(self.builder.get_object('checkbutton12').get_active()))
        self.dataset.config.set('Atlas', 'species_accounts_show_latest', str(self.builder.get_object('checkbutton13').get_active()))
        self.dataset.config.set('Atlas', 'species_accounts_show_statistics', str(self.builder.get_object('checkbutton14').get_active()))
        self.dataset.config.set('Atlas', 'species_accounts_show_status', str(self.builder.get_object('checkbutton16').get_active()))
        self.dataset.config.set('Atlas', 'species_accounts_show_phenology', str(self.builder.get_object('checkbutton15').get_active()))
        self.dataset.config.set('Atlas', 'species_accounts_phenology_type', str(self.builder.get_object('combobox12').get_active_text()))
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
        

        #single species map
        self.dataset.config.set('Species', 'title', self.builder.get_object('entry7').get_text())
        self.dataset.config.set('Species', 'author', self.builder.get_object('entry6').get_text())   
        self.dataset.config.set('Species', 'distribution_unit', self.builder.get_object('combobox2').get_active_text())     


        #grab a comma delimited list of species
        selection = self.builder.get_object('treeview7').get_selection()
        model, selected = selection.get_selected_rows()
        iters = [model.get_iter(path) for path in selected]

        species = ''

        for iter in iters:
            species = ','.join([species, model.get_value(iter, 0)])

        self.dataset.config.set('Species', 'species', species[1:])
        
        self.dataset.config.set('Species', 'species_update_title', str(self.builder.get_object('checkbutton4').get_active()))

        #grab a comma delimited list of vcs
        selection = self.builder.get_object('treeview8').get_selection()
        model, selected = selection.get_selected_rows()
        iters = [model.get_iter(path) for path in selected]

        vcs = ''

        for iter in iters:
            vcs = ','.join([vcs, model.get_value(iter, 0)])

        self.dataset.config.set('Species', 'vice-counties', vcs[1:])

        #page setup
        self.dataset.config.set('Species', 'paper_size', self.builder.get_object('combobox7').get_active_text())
        self.dataset.config.set('Species', 'orientation', self.builder.get_object('combobox4').get_active_text())


    def switch_update_title(self, widget):
        self.dataset.config.set('Atlas', 'families_update_title', str(self.builder.get_object('checkbutton18').get_active()))
        self.dataset.config.set('List', 'families_update_title', str(self.builder.get_object('checkbutton17').get_active()))
        self.dataset.config.set('Species', 'species_update_title', str(self.builder.get_object('checkbutton4').get_active()))

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

        if response == -5:
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
