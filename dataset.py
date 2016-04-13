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
import tempfile
import sqlite3
import mimetypes
import ConfigParser
import read
from subprocess import call

class Dataset(gobject.GObject):

    def __init__(self, filename):
        gobject.GObject.__init__(self)

        self.filename = filename
        self.mime = None
        self.builder = None

        self.records = None
        self.taxa = None
        self.families = []
        self.species = []

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
                                                     'bibliography': '',
                                                     'distribution_unit': '2km',
                                                     'families': '',
                                                     'species':'',
                                                     'families_update_title': 'True',
                                                     'vice-counties': '',
                                                     'date_bands': '[["squares", "#000", "#000", 1980, 2050], ["circles", "#a9a9a9", "#000", 1600, 1980]]',
                                                     'date_band_overlay': 'False',
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
                                                     'species_accounts_phenology_type': 'Months',
                                                     'species_update_title': 'True',
                                                     'mapping_layers': '',
                                                    })

            self.config.add_section('Atlas')
            self.config.add_section('List')
            self.config.add_section('Species')

        #guess the mimetype of the file
        self.mime = mimetypes.guess_type(self.filename)[0]

        if self.mime == 'application/vnd.ms-excel':
            self.data_source = read.Read(self.filename, self)
        else:
            temp_file = tempfile.NamedTemporaryFile(dir=self.temp_dir).name

            try:
                returncode = call(["ssconvert", self.filename, ''.join([temp_file, '.xls'])])

                if returncode == 0:
                    self.data_source = read.Read(''.join([temp_file, '.xls']), self)
            except OSError:
                pass


    def close(self):
        self.connection = None
        self.cursor = None

