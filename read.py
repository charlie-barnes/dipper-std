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
from datetime import datetime
import mimetypes
import pyodbc
import tempfile
from subprocess import call

from vaguedateparse import VagueDate
from geographiccoordinatesystem import Coordinate

vcname2number = {
'West Cornwall with Scilly':1,
'East Cornwall':2,
'South Devon':3,
'North Devon':4,
'South Somerset':5,
'North Somerset':6,
'North Wiltshire':7,
'South Wiltshire':8,
'Dorset':9,
'Isle of Wight':10,
'South Hampshire':11,
'North Hampshire':12,
'West Sussex':13,
'East Sussex':14,
'East Kent':15,
'West Kent':16,
'Surrey':17,
'South Essex':18,
'North Essex':19,
'Hertfordshire':20,
'Middlesex':21,
'Berkshire':22,
'Oxfordshire':23,
'Buckinghamshire':24,
'East Suffolk':25,
'West Suffolk':26,
'East Norfolk':27,
'West Norfolk':28,
'Cambridgeshire':29,
'Bedfordshire':30,
'Huntingdonshire':31,
'Northamptonshire':32,
'East Gloucestershire':33,
'West Gloucestershire':34,
'Monmouthshire':35,
'Herefordshire':36,
'Worcestershire':37,
'Warwickshire':38,
'Staffordshire':39,
'Shropshire':40,
'Glamorganshire':41,
'Breconshire':42,
'Radnorshire':43,
'Carmarthenshire':44,
'Pembrokeshire':45,
'Cardiganshire':46,
'Montgomeryshire':47,
'Merionethshire':48,
'Caernarvonshire':49,
'Denbighshire':50,
'Flintshire':51,
'Anglesey':52,
'South Lincolnshire':53,
'North Lincolnshire':54,
'Leicestershire (with Rutland)':55,
'Leicestershire':55,
'Nottinghamshire':56,
'Derbyshire':57,
'Cheshire':58,
'South Lancashire':59,
'West Lancashire':60,
'South-east Yorkshire':61,
'North-east Yorkshire':62,
'South-west Yorkshire':63,
'Mid-west Yorkshire':64,
'North-west Yorkshire':65,
'County Durham':66,
'South Northumberland':67,
'North Northumberland':68,
'Westmorland (with Furness)':69,
'Cumberland':70,
'Isle of Man':71,
'Dumfriesshire':72,
'Kirkcudbrightshire':73,
'Wigtownshire':74,
'Ayrshire':75,
'Renfrewshire':76,
'Lanarkshire':77,
'Peeblesshire':78,
'Selkirkshire':79,
'Roxburghshire':80,
'Berwickshire':81,
'East Lothian':82,
'Midlothian':83,
'West Lothian':84,
'Fifeshire':85,
'Stirlingshire':86,
'West Perthshire':87,
'Mid Perthshire':88,
'East Perthshire':89,
'Angus':90,
'Kincardineshire':91,
'South Aberdeenshire':92,
'North Aberdeenshire':93,
'Banffshire':94,
'Moray':95,
'East Inverness-shire':96,
'West Inverness-shire':97,
'Argyllshire':98,
'Dunbartonshire':99,
'Clyde Isles':100,
'Kintyre':101,
'South Ebudes':102,
'Mid Ebudes':103,
'North Ebudes':104,
'West Ross & Cromarty':105,
'East Ross & Cromarty':106,
'East Sutherland':107,
'West Sutherland':108,
'Caithness':109,
'Outer Hebrides':110,
'Orkney':111,
'Shetland':112}

class Read(gobject.GObject):

    def __init__(self, source, dataset, progressbar):
        gobject.GObject.__init__(self)

        self.source = source
        self.dataset = dataset
        self.progressbar = progressbar
        self.progressbar.set_text('')
        self.progressbar.set_fraction(0.0)

    
    def read(self):
        if os.path.isfile(self.source):
            #guess the mimetype of the file   
            self.mime = mimetypes.guess_type(self.source)[0]
                    
            if self.mime == 'application/vnd.ms-excel':
                return self.query_xls()
            else:
                temp_file = tempfile.NamedTemporaryFile(dir=self.dataset.temp_dir).name
    
                try:
                    self.progressbar.set_text(''.join(['Converting ', os.path.basename(self.source), '...']))
    
                    while gtk.events_pending():
                        gtk.main_iteration(False)
                        
                    returncode = call(["ssconvert", self.source, ''.join([temp_file, '.xls'])])
    
                    if returncode == 0:
                        self.source = ''.join([temp_file, '.xls'])
                        return self.query_xls()
                except OSError:
                    pass
        elif os.path.isdir(self.source):
            return self.query_mapmate()
            
            #DBfile = self.source
            #conn = pyodbc.connect('DRIVER={Microsoft Access Driver (*.mdb)};DBQ='+DBfile)
            #cursor = conn.cursor()
            
        else:
            return self.query_sql()
        
            
    def query_mapmate(self):
        return False
        
        
    def query_sql(self):
        print self.source
    
    
    def query_xls(self):
        '''Read the file and insert the data into the sqlite database.'''
        book = xlrd.open_workbook(self.source)

        text = ''.join(['Opening <b>', os.path.basename(self.source) ,'</b>', ' from ', '<b>', os.path.dirname(os.path.abspath(self.source)), '</b>'])

        temp_taxa_list = []

        #do we have a data sheet?
        if '--data--' in book.sheet_names():
            has_data = True
        else:
            has_data = False
        
        #set available sheets
        for name in book.sheet_names():                
            if name[:2] != '--' and name [-2:] != '--':
                self.dataset.available_sheets.append(name)

        #load the data from the config'd sheets, else do all sheets
        try:
            sheets = [(book.sheet_by_name(self.dataset.config.get('DEFAULT', 'sheets')))]
        except xlrd.biffh.XLRDError:
            if self.dataset.config.get('DEFAULT', 'sheets') != '-- all sheets --':
                md = gtk.MessageDialog(None,
                    gtk.DIALOG_DESTROY_WITH_PARENT, gtk.MESSAGE_ERROR,
                    gtk.BUTTONS_CLOSE, ''.join(['Specified sheet \'', self.dataset.config.get('DEFAULT', 'sheets'), '\' not found in the source file. Defaulting to all sheets.']))
                md.run()
                md.destroy()
                
            sheets = []
            for sheet in self.dataset.available_sheets:
                sheets.append(book.sheet_by_name(sheet))

        try:
            sheetnum = 0
            totalrecords = 0
            row_count = 0
            
            for sheet in sheets:
                sheetnum = sheetnum + 1
                totalrecords = totalrecords + sheet.nrows
                
            #loop through the selected sheets of the workbook
            for sheet in sheets:

                self.dataset.cursor.execute("begin")
                
                self.progressbar.set_text(''.join(['Reading sheet ', sheet.name]))

                #self.dataset.sheet = ' + '.join([self.dataset.sheet, sheet.name])

                # try and match up the column headings
                for col_index in range(sheet.ncols):
                    if sheet.cell(0, col_index).value.lower().strip() in ['taxon name', 'taxon', 'recommended taxon name', 'species', 'taxonname']:
                        taxon_position = col_index
                    elif sheet.cell(0, col_index).value.lower().strip() in ['grid reference', 'grid ref', 'grid ref.', 'gridref', 'sample spatial reference', 'gridreference']:
                        grid_reference_position = col_index
                    elif sheet.cell(0, col_index).value.lower().strip() in ['date', 'sample date']:
                        date_position = col_index
                    elif sheet.cell(0, col_index).value.lower().strip() in ['location', 'site', 'site name', 'sitename']:
                        location_position = col_index
                    elif sheet.cell(0, col_index).value.lower().strip() in ['recorder', 'recorders']:
                        recorder_position = col_index
                    elif sheet.cell(0, col_index).value.lower().strip() in ['determiner', 'determiners']:
                        determiner_position = col_index
                    elif sheet.cell(0, col_index).value.lower().strip() in ['vc', 'vice-county', 'vice county', 'vicecounty']:
                        vc_position = col_index
                    elif sheet.cell(0, col_index).value.lower().strip() in ['voucher', 'voucher status']:
                        voucher_position = col_index
                    elif sheet.cell(0, col_index).value.lower().strip() in ['start date', 'startdate']:
                        start_date_position = col_index
                    elif sheet.cell(0, col_index).value.lower().strip() in ['end date', 'enddate']:
                        end_date_position = col_index
                    elif sheet.cell(0, col_index).value.lower().strip() in ['date type', 'datetype']:
                        date_type_position = col_index

                rec_names = []
                det_names = []

                #loop through each row, skipping the header (first) row
                for row_index in range(1, sheet.nrows):
                    row_count = row_count + 1
                    
                    if row_index % 10 == 0:
                        self.progressbar.set_fraction(float(row_count)/float(totalrecords))
                        self.progressbar.set_text(''.join(['Parsing row ', str(row_index), '/', str(sheet.nrows), ' of sheet ', sheet.name, ]))

                    while gtk.events_pending():
                        gtk.main_iteration(False)
                    
                    taxa = sheet.cell(row_index, taxon_position).value
                    location = sheet.cell(row_index, location_position).value
                    grid_reference = sheet.cell(row_index, grid_reference_position).value


                    #get the date; if it fails try and decode an NBN exchange format date instead
                    try:
                        parseddate = sheet.cell(row_index, date_position).value

                        #check to see if we have an excel date integer - if so convert to a date string
                        if sheet.cell(row_index, date_position).ctype == 3:                        
                            tupledate = xlrd.xldate_as_tuple(parseddate, book.datemode)
                            pdate = datetime(tupledate[0],tupledate[1],tupledate[2])
                            parseddate = pdate.strftime("%d/%m/%Y")
                    
                    #decode the NBN exchange format date        
                    except UnboundLocalError:
                        startdate = sheet.cell(row_index, start_date_position).value
                        enddate = sheet.cell(row_index, end_date_position).value
                        datetype = sheet.cell(row_index, date_type_position).value

                        if datetype.lower() in ('d', 'day'):      
                                
                            if sheet.cell(row_index, start_date_position).ctype == 3:   
                                tupledate = xlrd.xldate_as_tuple(startdate, book.datemode)                                
                                pdate = datetime(tupledate[0],tupledate[1],tupledate[2])
                                parseddate = pdate.strftime('%d/%m/%Y')
                            else:
                                parseddate = startdate            
                                         
                        elif datetype.lower() in ('dd', 'day range'):
                        
                            if sheet.cell(row_index, start_date_position).ctype == 3:                    
                                tupledate = xlrd.xldate_as_tuple(startdate, book.datemode)
                                pdate = datetime(tupledate[0],tupledate[1],tupledate[2])
                                startdate = pdate.strftime('%d/%m/%Y')
                                      
                            if sheet.cell(row_index, end_date_position).ctype == 3:       
                                tupledate = xlrd.xldate_as_tuple(enddate, book.datemode)
                                pdate = datetime(tupledate[0],tupledate[1],tupledate[2])
                                enddate = pdate.strftime('%d/%m/%Y')
                                                    
                            parseddate = '-'.join([startdate, endate])
                            
                        elif datetype.lower() in ('o', 'month'):        
                        
                            if sheet.cell(row_index, start_date_position).ctype == 3:      
                                tupledate = xlrd.xldate_as_tuple(startdate, book.datemode)
                                pdate = datetime.strptime(startdate, '%d/%m/%Y')
                                parseddate = pdate.strftime("%m %Y")
                            else:
                                pdate = datetime.strptime(startdate, '%d/%m/%Y')
                                parseddate = pdate.strftime("%m %Y")
                        
                        elif datetype.lower() in ('oo', 'month range'):  
                        
                            if sheet.cell(row_index, start_date_position).ctype == 3:      
                                tupledate = xlrd.xldate_as_tuple(startdate, book.datemode)
                                pdate = datetime.strptime(startdate, '%d/%m/%Y')
                                startdate = pdate.strftime("%m %Y")
                            else:
                                pdate = datetime.strptime(startdate, '%d/%m/%Y')
                                startdate = pdate.strftime("%m %Y")
                                
                            if sheet.cell(row_index, end_date_position).ctype == 3:      
                                tupledate = xlrd.xldate_as_tuple(enddate, book.datemode)
                                pdate = datetime.strptime(enddate, '%d/%m/%Y')
                                enddate = pdate.strftime("%m %Y")
                            else:
                                pdate = datetime.strptime(enddate, '%d/%m/%Y')
                                enddate = pdate.strftime("%m %Y")
                                                    
                            parseddate = '-'.join([startdate, endate])
                            
                        elif datetype.lower() in ('y', 'year'):  
                        
                            if sheet.cell(row_index, start_date_position).ctype == 3:      
                                tupledate = xlrd.xldate_as_tuple(startdate, book.datemode)
                                pdate = datetime.strptime(startdate, '%d/%m/%Y')
                                parseddate = pdate.strftime("%Y")
                            else:
                                pdate = datetime.strptime(startdate, '%d/%m/%Y')
                                parseddate = pdate.strftime("%Y")
                    
                        elif datetype.lower() in ('yy', 'year range'):  
                        
                            if sheet.cell(row_index, start_date_position).ctype == 3:      
                                tupledate = xlrd.xldate_as_tuple(startdate, book.datemode)
                                pdate = datetime.strptime(startdate, '%d/%m/%Y')
                                startdate = pdate.strftime("%Y")
                            else:
                                pdate = datetime.strptime(startdate, '%d/%m/%Y')
                                startdate = pdate.strftime("%Y")
                                
                            if sheet.cell(row_index, end_date_position).ctype == 3:      
                                tupledate = xlrd.xldate_as_tuple(enddate, book.datemode)
                                pdate = datetime.strptime(enddate, '%d/%m/%Y')
                                enddate = pdate.strftime("%Y")
                            else:
                                pdate = datetime.strptime(enddate, '%d/%m/%Y')
                                enddate = pdate.strftime("%Y")
                                                    
                            parseddate = '-'.join([startdate, endate])
                            
                        elif datetype.lower() in ('-y', 'before year'):  
                        
                            if sheet.cell(row_index, start_date_position).ctype == 3:   
                                tupledate = xlrd.xldate_as_tuple(enddate, book.datemode)
                                pdate = datetime.strptime(enddate, '%d/%m/%Y')
                                parseddate = pdate.strftime("%Y")
                            else:
                                pdate = datetime.strptime(enddate, '%d/%m/%Y')
                                parseddate = pdate.strftime("-%Y")
                                                                         
                        elif datetype.lower() in ('nd', 'u', 'no date', 'unknown'):
                            parseddate = 'Unknown'
                        

                    recorder = sheet.cell(row_index, recorder_position).value
                    
                    try:
                        voucher = sheet.cell(row_index, voucher_position).value
                    except UnboundLocalError:
                        voucher = None

                    #we can allow null determiners
                    try:
                        determiner = sheet.cell(row_index, determiner_position).value

                        if determiner not in det_names:
                            det_names.append(determiner)
                    except UnboundLocalError:
                        determiner = None
                    
                    vaguedate = VagueDate(parseddate)
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

                    if grid_100km not in self.dataset.occupied_squares['100km']:
                        self.dataset.occupied_squares['100km'].append(grid_100km)

                    if grid_10km not in self.dataset.occupied_squares['10km']:
                        self.dataset.occupied_squares['10km'].append(grid_10km)

                    if grid_5km not in self.dataset.occupied_squares['1km']:
                        self.dataset.occupied_squares['1km'].append(grid_1km)

                    if grid_2km not in self.dataset.occupied_squares['5km']:
                        self.dataset.occupied_squares['5km'].append(grid_5km)

                    if grid_1km not in self.dataset.occupied_squares['2km']:
                        self.dataset.occupied_squares['2km'].append(grid_2km)

                    if grid_100m not in self.dataset.occupied_squares['100m']:
                        self.dataset.occupied_squares['100m'].append(grid_100m)

                    if grid_10m not in self.dataset.occupied_squares['10m']:
                        self.dataset.occupied_squares['10m'].append(grid_10m)

                    if grid_1m not in self.dataset.occupied_squares['1m']:
                        self.dataset.occupied_squares['1m'].append(grid_1m)

                    #we can allow null vcs 
                    try:
                        #check to see if vc is a number
                        vc = sheet.cell(row_index, vc_position).value
                        try:
                            vc+0
                        except TypeError:
                            try:
                                vc = vcname2number[vc]
                            except KeyError:
                                vc = 'Unknown'
                            
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
                                                     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)',
                                                     [taxa,
                                                     location,
                                                     grid_reference,
                                                     grid_100km, grid_10km, grid_5km, grid_2km,
                                                     grid_1km, grid_100m, grid_10m, grid_1m,
                                                     easting, northing, accuracy, parseddate,
                                                     decade, year, month, day,
                                                     decade_from, year_from, month_from, day_from,
                                                     decade_to, year_to, month_to, day_to,
                                                     recorder,
                                                     determiner,
                                                     vc,
                                                     voucher])

                    self.dataset.records = self.dataset.records + 1

                self.dataset.recorders = len(rec_names)
                self.dataset.determiners = len(det_names)
                self.progressbar.set_text(''.join(['Committing data from sheet ', sheet.name]))
                self.dataset.cursor.execute("commit")                

            self.dataset.sheet = self.dataset.sheet[3:]
            
            #load the data sheet
            if has_data:
                sheet = book.sheet_by_name('--data--')

                # try and match up the column headings
                for col_index in range(sheet.ncols):

                    if sheet.cell(0, col_index).value.lower() in ['taxon', 'taxon name', 'species']:
                        taxon_position = col_index
                    elif sheet.cell(0, col_index).value.lower() in ['kingdom']:
                        kingdom_position = col_index
                    elif sheet.cell(0, col_index).value.lower() in ['phylum']:
                        phylum_position = col_index
                    elif sheet.cell(0, col_index).value.lower() in ['class']:
                        class_position = col_index
                    elif sheet.cell(0, col_index).value.lower() in ['order']:
                        order_position = col_index
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

                    try:
                        taxa = sheet.cell(row_index, taxon_position).value
                        genus = taxa.split()[0]
                        
                        if taxa in temp_taxa_list:
                            try:
                                kingdom = sheet.cell(row_index, kingdom_position).value
                                                            
                                if kingdom not in self.dataset.kingdoms.keys():
                                    self.dataset.kingdoms[kingdom] = ['', None]
                            except UnboundLocalError:
                                kingdom = ''
                                
                            try:
                                phylum = sheet.cell(row_index, phylum_position).value                            

                                if phylum not in self.dataset.phyla.keys():
                                    self.dataset.phyla[phylum] = [kingdom, None]
                            except UnboundLocalError:
                                phylum = ''
                                
                            try:
                                class_ = sheet.cell(row_index, class_position).value
                                
                                if class_ not in self.dataset.classes.keys():
                                    self.dataset.classes[class_] = [phylum, None]
                            except UnboundLocalError:
                                class_ = ''
                                
                            try:
                                order = sheet.cell(row_index, order_position).value
                                
                                if order not in self.dataset.orders.keys():
                                    self.dataset.orders[order] = [class_, None]                                
                            except UnboundLocalError:
                                order = ''
                                
                            try:
                                family = sheet.cell(row_index, family_position).value
                                
                                if family not in self.dataset.families.keys():
                                    self.dataset.families[family] = [order, None]
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

                            if genus not in self.dataset.genera.keys():
                                self.dataset.genera[genus] = [family, None]

                            if taxa not in self.dataset.specie.keys():
                                self.dataset.specie[taxa] = [genus, None]
                            ### some duplication here - clean it up
                            if taxa not in self.dataset.taxa.keys():
                                self.dataset.taxa[taxa] = {'kingdom': kingdom, 'phylum': phylum, 'class': class_, 'order': order, 'family': family}
                                
                            if taxa not in self.dataset.species:
                                self.dataset.species.append(taxa)

                            self.dataset.cursor.execute('INSERT INTO species_data \
                                                         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)',
                                                         [taxa,
                                                         sort_order,
                                                         nbn_key,
                                                         national_status,
                                                         None,
                                                         description,
                                                         common_name,
                                                         kingdom,
                                                         phylum,
                                                         class_,
                                                         order,
                                                         family,
                                                         genus])
                    except (IndexError, AttributeError):
                        pass

            else:
                #if we dont have species data sheet, just create the data using distinct taxa
                self.dataset.cursor.execute('SELECT DISTINCT(data.taxon) \
                                            FROM data \
                                            ORDER BY data.taxon')

                data = self.dataset.cursor.fetchall()

                for row in data:
                    self.dataset.cursor.execute('INSERT INTO species_data \
                                                 VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)',
                                                [row[0],
                                                None,
                                                None,
                                                None,
                                                None,
                                                None,
                                                None,
                                                None,
                                                None,
                                                None,
                                                None,
                                                None,
                                                None])
                                                

                for row_index in range(1, sheet.nrows):

                    taxa = sheet.cell(row_index, taxon_position).value      

                    if taxa not in self.dataset.specie.keys():
                        self.dataset.specie[taxa] = [None, None]                                          


            return True
        except UnboundLocalError as e:
            md = gtk.MessageDialog(None,
                gtk.DIALOG_DESTROY_WITH_PARENT, gtk.MESSAGE_ERROR,
                gtk.BUTTONS_CLOSE, ''.join(['Unable to open data file: ', str(e)]))
            md.run()
            md.destroy()
            return False

