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
import xlrd
import os
import gtk
import json

from vaguedateparse import VagueDate
from geographiccoordinatesystem import Coordinate

class Read(gobject.GObject):

    def __init__(self, filename, dataset):
        gobject.GObject.__init__(self)

        self.filename = filename
        self.dataset = dataset


    def read(self):
        '''Read the file and insert the data into the sqlite database.'''

        book = xlrd.open_workbook(self.filename)

        text = ''.join(['Opening <b>', os.path.basename(self.filename) ,'</b>', ' from ', '<b>', os.path.dirname(os.path.abspath(self.filename)), '</b>'])

        temp_taxa_list = []

        #do we have a data sheet?
        for name in book.sheet_names():
            if name == '--data--':
                has_data = True
                
            if name[:2] != '--' and name [-2:] != '--':
                self.dataset.available_sheets.append(name)

        #load the data from the config'd sheets
        if self.dataset.config.get('DEFAULT', 'sheets') != '-- all sheets --':
            sheets = [(book.sheet_by_name(self.dataset.config.get('DEFAULT', 'sheets')))]
        else:
            sheets = []
            for sheet in self.dataset.available_sheets:
                sheets.append(book.sheet_by_name(sheet))

        try:
            #loop through the selected sheets of the workbook
            for sheet in sheets:
                #self.dataset.sheet = ' + '.join([self.dataset.sheet, sheet.name])
                
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

                rec_names = []
                det_names = []

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

                        if determiner not in det_names:
                            det_names.append(determiner)
                    except UnboundLocalError:
                        determiner = None
                    
                    vaguedate = VagueDate(date)
                    decade, year, month, day, decade_from, year_from, month_from, day_from, decade_to, year_to, month_to, day_to = vaguedate.decade, vaguedate.year, vaguedate.month,  vaguedate.day, vaguedate.decade_from, vaguedate.year_from, vaguedate.month_from,  vaguedate.day_from, vaguedate.decade_to, vaguedate.year_to, vaguedate.month_to,  vaguedate.day_to

                    #stats
                    if year > self.dataset.latest:
                        self.dataset.latest = year
                        
                    if year < self.dataset.earliest:
                        self.dataset.earliest = year

                    if recorder not in rec_names:
                        rec_names.append(recorder)

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

                        if vc not in self.dataset.vicecounties:
                            self.dataset.vicecounties.append(vc)
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

                    self.dataset.records = self.dataset.records + 1

                self.dataset.recorders = len(rec_names)
                self.dataset.determiners = len(det_names)

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
                            
                        if taxa not in self.dataset.species:
                            self.dataset.species.append(taxa)

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

