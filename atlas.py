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
from datetime import datetime
import pdf
import version
import cfg
import shapefile
from PIL import Image
from PIL import ImageDraw
from PIL import ImageChops
import tempfile
from colour import Color
import math
import random
import chart
import os
import json

class Atlas(gobject.GObject):

    def __init__(self, dataset):
        gobject.GObject.__init__(self)
        self.dataset = dataset
        self.save_in = None
        self.page_unit = 'mm'
        self.base_map = None
        self.density_map_filename = None
        self.date_band_coverage = []
        self.increments = 14               
        
        
    def generate_density_map(self):

        #generate the base map
        scalefactor = 0.01

        layers = []
        for layer in json.loads(self.dataset.config.get('Atlas', 'mapping_layers')):
            layers.append(str('gis'+os.path.sep+layer[0]+os.path.sep+layer[1]+'.shp'))

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

            crop_left = int(math.ceil(bounds_bottom_x/100))
            crop_upper = int(math.ceil((1300000/100)-(bounds_top_y/100)))
            crop_right = int(math.ceil(bounds_top_x/100))
            crop_lower = int(math.ceil((1300000/100)-(bounds_bottom_y/100)))

            region = background_map.crop((crop_left, crop_upper, crop_right, crop_lower))
            region.save(temp_file, format='PNG')

            ###HACK: for some reason the crop of the background map isn't always the right
            #size. seems to be off by 1 pixel in some cases.
            (region_width, region_height) = region.size

            width_hack_diff = region_width-(int(xdist*scalefactor)+1)
            height_hack_diff = region_height-(int(ydist*scalefactor)+1)
            
            base_map.paste(region, (0, 0, (int(xdist*scalefactor)+1)+width_hack_diff, (int(ydist*scalefactor)+1)+height_hack_diff  ))

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

        grids = []

        gridsdict = {}
        max_count = 0

        for tup in data:
            grids.append(tup[0])
            gridsdict[tup[0]] = tup[1]
            if tup[1] > max_count:
                max_count = tup[1]

        #work out how many increments we need
        for ranges in cfg.gradation_ranges:
            ranges[2] = 0 # reset the swatch count
            if max_count >= ranges[0] and max_count <= ranges[1]:
                self.increments = cfg.gradation_ranges.index(ranges)+1

        self.grad_ranges = cfg.gradation_ranges[:self.increments]
        
        #calculate the colour gradient
        low = Color(self.dataset.config.get('Atlas', 'species_density_map_low_colour'))
        high = Color(self.dataset.config.get('Atlas', 'species_density_map_high_colour'))
        self.grad_fills = list(low.range_to(high, self.increments))
        
        r = shapefile.Reader('./markers/' + self.dataset.config.get('Atlas', 'species_density_map_style') + os.path.sep + self.dataset.config.get('Atlas', 'species_density_map_unit'))
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
                        
                        #increment the swatch count
                        swatch_ranges[2] = swatch_ranges[2] + 1
                        
                        
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
        self.scalefactor = 0.01

        if self.dataset.use_vcs:
            vcs_sql = ''.join(['WHERE data.vc IN (', self.dataset.config.get('Atlas', 'vice-counties'), ')'])
        else:
            vcs_sql = ''

        layers = []
        for layer in json.loads(self.dataset.config.get('Atlas', 'mapping_layers')):
            layers.append(str('gis'+os.path.sep+layer[0]+os.path.sep+layer[1]+'.shp'))

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
        #r = shapefile.Reader('./markers/' + self.dataset.config.get('Atlas', 'coverage_style') + os.path.sep + self.dataset.config.get('Atlas', 'distribution_unit'))
        #loop through each object in the shapefile
        #for obj in r.shapeRecords():
            #if the grid is in our coverage, extend the bounds to match
        #    if obj.record[0] in grids:
        #        if obj.shape.bbox[0] < self.bounds_bottom_x:
        #            self.bounds_bottom_x = obj.shape.bbox[0]#

        #        if obj.shape.bbox[1] < self.bounds_bottom_y:
        #            self.bounds_bottom_y = obj.shape.bbox[1]

        #        if obj.shape.bbox[2] > self.bounds_top_x:
        #            self.bounds_top_x = obj.shape.bbox[2]

        #        if obj.shape.bbox[3] > self.bounds_top_y:
        #            self.bounds_top_y = obj.shape.bbox[3]


        #loop through the date bands treeview
        #for row in self.dataset.builder.get_object('treeview6').get_model():
            # Read in the date band 1 grid ref shapefiles and extend the bounding box
        #    r = shapefile.Reader('./markers/' + row[0] + os.path.sep + self.dataset.config.get('Atlas', 'distribution_unit'))
            #loop through each object in the shapefile
        #    for obj in r.shapeRecords():
                #if the grid is in our coverage, extend the bounds to match
        #        if obj.record[0] in grids:
        #            if obj.shape.bbox[0] < self.bounds_bottom_x:
        #                self.bounds_bottom_x = obj.shape.bbox[0]

        #            if obj.shape.bbox[1] < self.bounds_bottom_y:
        #                self.bounds_bottom_y = obj.shape.bbox[1]

        #            if obj.shape.bbox[2] > self.bounds_top_x:
        #                self.bounds_top_x = obj.shape.bbox[2]

        #            if obj.shape.bbox[3] > self.bounds_top_y:
        #                self.bounds_top_y = obj.shape.bbox[3]            
        
        # Read in the layer shapefiles and extend the bounding box
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

        #loop through the date bands treeview
        for row in self.dataset.builder.get_object('treeview6').get_model():
            current_grids = []
                    
            r = shapefile.Reader('./markers/' + row[0] + os.path.sep + self.dataset.config.get('Atlas', 'distribution_unit'))
            #loop through each object in the shapefile
            for obj in r.shapeRecords():
                if obj.record[0] in grids:
                    current_grids.append(obj)
            
            self.date_band_coverage.append(current_grids)

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



        #add the coverage
        r = shapefile.Reader('./markers/' + self.dataset.config.get('Atlas', 'coverage_style') + os.path.sep + self.dataset.config.get('Atlas', 'distribution_unit'))
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

        #re-add each boundary shapefile, but don't fill (otherwise we loose the nice definitive outline of the boundary from the grids/coverage)
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
                self.base_map_draw.polygon(pixels, outline='rgb(' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'vice-counties_outline')).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'vice-counties_outline')).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Atlas', 'vice-counties_outline')).blue_float*255)) + ')')


        #mask off everything outside the boundary area
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
                #draw the final polygon (or the only one, if we have just the one)
                mask_draw.polygon(pixels, fill='rgb(0,0,0)')

        mask = ImageChops.invert(mask)
        self.base_map.paste(mask, (0,0), mask)

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
                                     WHERE ' + vcs_sql + ' \
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
                                     WHERE ' + vcs_sql + ' \
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
        doc = pdf.PDF(orientation=self.dataset.config.get('Atlas', 'orientation'),unit=self.page_unit,format=self.dataset.config.get('Atlas', 'paper_size'))
        doc.type = 'atlas'
        doc.toc_length = toc_length

        doc.col = 0
        doc.y0 = 0
        doc.set_title(self.dataset.config.get('Atlas', 'title'))
        doc.set_author(self.dataset.config.get('Atlas', 'author'))
        doc.set_creator(' '.join(['dipper-stda', version.__version__])) 
        doc.section = ''

        #title page
        doc.p_add_page()

        if self.dataset.config.get('Atlas', 'cover_image') is not None and os.path.isfile(self.dataset.config.get('Atlas', 'cover_image')):
            doc.image(self.dataset.config.get('Atlas', 'cover_image'), 0, 0, doc.w, doc.h)

        doc.set_text_color(0)
        doc.set_fill_color(255, 255, 255)
        doc.set_font('Helvetica', '', 28)
        doc.ln(15)
        doc.multi_cell(0, 10, doc.title, 0, 'L', False)

        doc.ln(20)
        doc.set_font('Helvetica', '', 18)
        doc.multi_cell(0, 10, ''.join([doc.author, '\n',datetime.now().strftime('%B %Y')]), 0, 'L', False)

        #inside cover
        doc.p_add_page()
        doc.set_font('Helvetica', '', 12)
        doc.multi_cell(0, 6, self.dataset.config.get('Atlas', 'inside_cover'), 0, 'J', False)

        doc.do_header = True

        #introduction
        if len(self.dataset.config.get('Atlas', 'introduction')) > 0:
            doc.p_add_page()
            doc.section = ('Introduction')
            doc.set_font('Helvetica', '', 20)
            doc.multi_cell(0, 20, 'Introduction', 0, 'J', False)
            doc.set_font('Helvetica', '', 12)
            doc.multi_cell(0, 6, self.dataset.config.get('Atlas', 'introduction'), 0, 'J', False)

        #species density map
        if self.dataset.config.getboolean('Atlas', 'species_density_map_visible'):
            doc.section = ('Introduction')
            doc.p_add_page()
            doc.set_font('Helvetica', '', 20)
            doc.multi_cell(0, 20, 'Species density', 0, 'J', False)

            im = Image.open(self.density_map_filename)

            width, height = im.size

            if self.dataset.config.get('Atlas', 'orientation')[0:1] == 'P':
                scalefactor = (doc.w-doc.l_margin-doc.r_margin)/width
                target_width = width*scalefactor
                target_height = height*scalefactor

                while target_height >= (doc.h-40-30):
                    target_height = target_height - 1

                scalefactor = target_height/height
                target_width = width*scalefactor

            elif self.dataset.config.get('Atlas', 'orientation')[0:1] == 'L':
                scalefactor = (doc.h-40-30)/height
                target_width = width*scalefactor
                target_height = height*scalefactor

                while target_width >= (doc.w-doc.l_margin-doc.r_margin):
                    target_width = target_width - 1

                scalefactor = target_width/width
                target_height = height*scalefactor

            centerer = ((doc.w-doc.l_margin-doc.r_margin)-target_width)/2

            doc.image(self.density_map_filename, doc.l_margin+centerer, 40, w=target_width, h=target_height, type='PNG')

            #add the colour swatches
            
            #scale down the marker to a sensible size
            if self.dataset.config.get('Atlas', 'species_density_map_unit') == '100km':
                scalefactor = 0.00025
            elif self.dataset.config.get('Atlas', 'species_density_map_unit') == '10km':
                scalefactor = 0.0025
            elif self.dataset.config.get('Atlas', 'species_density_map_unit') == '5km':
                scalefactor = 0.005
            elif self.dataset.config.get('Atlas', 'species_density_map_unit') == '2km':
                scalefactor = 0.0125
            elif self.dataset.config.get('Atlas', 'species_density_map_unit') == '1km':
                scalefactor = 0.025
                
            #draw each band marker in turn and save out
            #we always show date band 1
            r = shapefile.Reader('./markers/' + self.dataset.config.get('Atlas', 'species_density_map_style') + os.path.sep + self.dataset.config.get('Atlas', 'species_density_map_unit'))
            shapes = r.shapes()
            pixels = []

            #grab the first marker we come to - no need to be fussy
            for x, y in shapes[0].points:
                px = 2+int(float(x-shapes[0].bbox[0]) * scalefactor)
                py = 2+int(float(shapes[0].bbox[3]-y) * scalefactor)
                pixels.append((px,py))

            count = 0

            doc.set_font('Helvetica', '', 9)
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
    
                doc.image(swatch_temp_file, 10, (200 + (count*5) + ((14-self.increments)*5)), h=4, type='PNG')

                doc.set_y(200 + (count*5) + ((14-self.increments)*5))
                doc.cell(4)
                
                if swatch_ranges[0] == swatch_ranges[1]:
                    swatch_text = ''.join([str(swatch_ranges[0]), ' (', str(swatch_ranges[2]), ')'])
                else:
                    swatch_text = ''.join([str(swatch_ranges[0]), ' - ', str(swatch_ranges[1]), ' (', str(swatch_ranges[2]), ')'])
                    
                doc.cell(10, 5, swatch_text, 0, 1, 'L', True)
                
                count = count + 1         
                

        #explanation map
        if doc.orientation == 'Portrait':
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

            doc.section = ('Introduction')
            doc.p_add_page()
            doc.set_font('Helvetica', '', 20)
            doc.multi_cell(0, 20, 'Species account explanation', 0, 'J', False)

            y_padding = 39#######extra Y padding to centralize
            y_padding = (5 + (((doc.w / 2)-doc.l_margin-doc.r_margin)+3+10+y_padding) + ((((doc.w / 2)-doc.l_margin-doc.r_margin)+3)/3.75))/2
            x_padding = doc.l_margin

            #taxon heading
            doc.set_y(y_padding)
            doc.set_x(x_padding)
            doc.set_text_color(255)
            doc.set_fill_color(59, 59, 59)
            doc.set_line_width(0.1)
            doc.set_font('Helvetica', 'BI', 12)
            doc.cell(((doc.w)-doc.l_margin-doc.r_margin)/2, 5, random_species, 'TLB', 0, 'L', True)
            doc.set_font('Helvetica', 'B', 12)
            doc.cell(((doc.w)-doc.l_margin-doc.r_margin)/2, 5, common_name, 'TRB', 1, 'R', True)
            doc.set_x(x_padding)

            status_text = ''
            if self.dataset.config.getboolean('Atlas', 'species_accounts_show_status'):
                status_text = ''.join([designation])

            doc.multi_cell(((doc.w)-doc.l_margin-doc.r_margin), 5, status_text, 1, 'L', True)

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
            doc.set_y(y_padding+12)
            doc.set_x(x_padding+(((doc.w / 2)-doc.l_margin-doc.r_margin)+3+5))
            doc.set_font('Helvetica', '', 10)
            doc.set_text_color(0)
            doc.set_fill_color(255, 255, 255)
            doc.set_line_width(0.1)

            if len(taxa_statistics[random_species]['description']) > 0 and self.dataset.config.getboolean('Atlas', 'species_accounts_show_descriptions'):
                doc.set_font('Helvetica', '', 10)
                doc.multi_cell((((doc.w / 2)-doc.l_margin-doc.r_margin)+12), 5, ''.join([taxa_statistics[random_species]['description'], '\n\n']), 0, 'L', False)
                doc.set_x(x_padding+(((doc.w / 2)-doc.l_margin-doc.r_margin)+3+5))

            if self.dataset.config.getboolean('Atlas', 'species_accounts_show_latest'):
                doc.set_font('Helvetica', '', 10)
                doc.multi_cell((((doc.w / 2)-doc.l_margin-doc.r_margin)+12), 5, ''.join(['Records (most recent first): ', taxon_recent_records[:-2], '.', remaining_records_text]), 0, 'L', False)

            y_for_explanation = doc.get_y()

            #chart
            if self.dataset.config.getboolean('Atlas', 'species_accounts_show_phenology'):
                chartobj = chart.Chart(self.dataset, random_species, self.dataset.config.get('Atlas', 'species_accounts_phenology_type'))
                if chartobj.temp_filename != None:
                    doc.image(chartobj.temp_filename, x_padding, ((doc.w / 2)-doc.l_margin-doc.r_margin)+3+10+y_padding, ((doc.w / 2)-doc.l_margin-doc.r_margin)+3, (((doc.w / 2)-doc.l_margin-doc.r_margin)+3)/3.75, 'PNG')
                doc.rect(x_padding, ((doc.w / 2)-doc.l_margin-doc.r_margin)+3+10+y_padding, ((doc.w / 2)-doc.l_margin-doc.r_margin)+3, (((doc.w / 2)-doc.l_margin-doc.r_margin)+3)/3.75)


            current_map =  self.base_map.copy()
            current_map_draw = ImageDraw.Draw(current_map)

            all_grids = []

            #loop through each date band, grabbing the records
            count = 0
            
            #if we're overlaying the markers, we draw the date bands backwards (oldest first)
            if self.dataset.config.getboolean('Atlas', 'date_band_overlay'):
                treemodel_reversed = []
                           
                for row in self.dataset.builder.get_object('treeview6').get_model():
                    treemodel_reversed.insert(0, row)
                
                for row in treemodel_reversed:
                    fill_colour = row[1].split('"')[1]
                    border_colour = row[2].split('"')[1]
                    
                    self.dataset.cursor.execute('SELECT DISTINCT(grid_' + self.dataset.config.get('Atlas', 'distribution_unit') + ') AS grids \
                                                FROM data \
                                                WHERE data.taxon = "' + random_species + '" \
                                                AND data.year_to >= ' + str(row[3]) + '\
                                                AND data.year_to < ' + str(row[4]) + ' \
                                                AND data.year_from >= ' + str(row[3]) + ' \
                                                AND data.year_from < ' + str(row[4]))
    
                    date_b_grids = []
                    date_band_grids = self.dataset.cursor.fetchall()
                    for g in date_band_grids:
                        date_b_grids.append(g[0])
                        
                    #loop through each object in the coverage (to save having to loop through ALL objects)
                    for obj in self.date_band_coverage[count]:
                        #if the object is in our date band
                        if obj.record[0] in date_b_grids:
                            
                            pixels = []
                            #loop through each point in the object
                            for x,y in obj.shape.points:
                                px = (self.xdist * self.scalefactor)- (self.bounds_top_x - x) * self.scalefactor
                                py = (self.bounds_top_y - y) * self.scalefactor
                                pixels.append((px,py))
                            current_map_draw.polygon(pixels, fill='rgb(' + str(int(gtk.gdk.color_parse(fill_colour).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(fill_colour).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(fill_colour).blue_float*255)) + ')', outline='rgb(' + str(int(gtk.gdk.color_parse(border_colour).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(border_colour).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(border_colour).blue_float*255)) + ')')
    
                    count = count + 1
    
            #if we're not overlaying the markers, we draw the date bands forwards (newest first) and skip any markers that have already been drawn                        
            else:
                for row in self.dataset.builder.get_object('treeview6').get_model():
                    fill_colour = row[1].split('"')[1]
                    border_colour = row[2].split('"')[1]
                    
                    self.dataset.cursor.execute('SELECT DISTINCT(grid_' + self.dataset.config.get('Atlas', 'distribution_unit') + ') AS grids \
                                                FROM data \
                                                WHERE data.taxon = "' + random_species + '" \
                                                AND data.year_to >= ' + str(row[3]) + '\
                                                AND data.year_to < ' + str(row[4]) + ' \
                                                AND data.year_from >= ' + str(row[3]) + ' \
                                                AND data.year_from < ' + str(row[4]))
    
                    date_b_grids = []
                    date_band_grids = self.dataset.cursor.fetchall()
                    for g in date_band_grids:
                        date_b_grids.append(g[0])
                        
                    #loop through each object in the coverage (to save having to loop through ALL objects)
                    for obj in self.date_band_coverage[count]:
                        #if the object is in our date band and it's not already been drawn by a more recent date band
                        if (obj.record[0] in date_b_grids) and (obj.record[0] not in all_grids):
                            
                            pixels = []
                            #loop through each point in the object
                            for x,y in obj.shape.points:
                                px = (self.xdist * self.scalefactor)- (self.bounds_top_x - x) * self.scalefactor
                                py = (self.bounds_top_y - y) * self.scalefactor
                                pixels.append((px,py))
                            current_map_draw.polygon(pixels, fill='rgb(' + str(int(gtk.gdk.color_parse(fill_colour).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(fill_colour).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(fill_colour).blue_float*255)) + ')', outline='rgb(' + str(int(gtk.gdk.color_parse(border_colour).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(border_colour).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(border_colour).blue_float*255)) + ')')
    
                    count = count + 1
    
                    #add the current date band grids for reference next time round
                    for g in date_band_grids:
                        all_grids.append(g[0])









            temp_map_file = tempfile.NamedTemporaryFile(dir=self.dataset.temp_dir).name
            current_map.save(temp_map_file, format='PNG')

            (width, height) =  current_map.size

            max_width = (doc.w / 2)-doc.l_margin-doc.r_margin
            max_height = (doc.w / 2)-doc.l_margin-doc.r_margin

            if height > width:
                pdfim_height = max_width
                pdfim_width = (float(width)/float(height))*max_width
            else:
                pdfim_height = (float(height)/float(width))*max_width
                pdfim_width = max_width

            img_x_cent = ((max_width-pdfim_width)/2)+2
            img_y_cent = ((max_height-pdfim_height)/2)+12

            doc.image(temp_map_file, x_padding+img_x_cent, y_padding+img_y_cent, int(pdfim_width), int(pdfim_height), 'PNG')


            #map container
            doc.set_text_color(0)
            doc.set_fill_color(255, 255, 255)
            doc.set_line_width(0.1)
            doc.rect(x_padding, 10+y_padding, ((doc.w / 2)-doc.l_margin-doc.r_margin)+3, ((doc.w / 2)-doc.l_margin-doc.r_margin)+3)

            if self.dataset.config.getboolean('Atlas', 'species_accounts_show_statistics'):

                #from
                doc.set_y(11+y_padding)
                doc.set_x(1+x_padding)
                doc.set_font('Helvetica', '', 12)
                doc.set_text_color(0)
                doc.set_fill_color(255, 255, 255)
                doc.set_line_width(0.1)

                if str(taxa_statistics[random_species]['earliest']) == 'Unknown':
                    val = '?'
                else:
                    val = str(taxa_statistics[random_species]['earliest'])
                doc.multi_cell(18, 5, ''.join(['E ', val]), 0, 'L', False)

                #to
                doc.set_y(11+y_padding)
                doc.set_x((((doc.w / 2)-doc.l_margin-doc.r_margin)-15)+x_padding)
                doc.set_font('Helvetica', '', 12)
                doc.set_text_color(0)
                doc.set_fill_color(255, 255, 255)
                doc.set_line_width(0.1)

                if str(taxa_statistics[random_species]['latest']) == 'Unknown':
                    val = '?'
                else:
                    val = str(taxa_statistics[random_species]['latest'])
                doc.multi_cell(18, 5, ''.join(['L ', val]), 0, 'R', False)

                #records
                doc.set_y((((doc.w / 2)-doc.l_margin-doc.r_margin)+7)+y_padding)
                doc.set_x(1+x_padding)
                doc.set_font('Helvetica', '', 12)
                doc.set_text_color(0)
                doc.set_fill_color(255, 255, 255)
                doc.set_line_width(0.1)
                doc.multi_cell(18, 5, ''.join(['R ', str(taxa_statistics[random_species]['count'])]), 0, 'L', False)

                #squares
                doc.set_y((((doc.w / 2)-doc.l_margin-doc.r_margin)+7)+y_padding)
                doc.set_x((((doc.w / 2)-doc.l_margin-doc.r_margin)-15)+x_padding)
                doc.set_font('Helvetica', '', 12)
                doc.set_text_color(0)
                doc.set_fill_color(255, 255, 255)
                doc.set_line_width(0.1)
                doc.multi_cell(18, 5, ''.join(['S ', str(taxa_statistics[random_species]['dist_count'])]), 0, 'R', False)



            #### the explanations
            doc.set_font('Helvetica', '', 9)
            doc.set_draw_color(0,0,0)

            #species name
            doc.line(20,
                     50,
                     x_padding+7,
                     y_padding)

            doc.set_x(1+x_padding +15)
            doc.set_y(11+y_padding -50)
            doc.cell(1)
            doc.cell(10, 5, 'Species name', 0, 0, 'L', True)

            #common name
            doc.line(doc.w-doc.l_margin-30,
                     50,
                     doc.w-doc.l_margin-x_padding-10,
                     y_padding)

            doc.set_x(1+x_padding +15)
            doc.set_y(11+y_padding -48)
            doc.cell(150)
            doc.cell(10, 5, 'Common name', 0, 0, 'L', True)


            doc.set_y(y_padding+12)
            doc.set_x(x_padding+(((doc.w / 2)-doc.l_margin-doc.r_margin)+3+5))

            #taxon blurb
            if self.dataset.config.getboolean('Atlas', 'species_accounts_show_latest') or self.dataset.config.getboolean('Atlas', 'species_accounts_show_descriptions'):
                doc.line(x_padding+(((doc.w / 2)-doc.l_margin-doc.r_margin)+3+5)   +50, #x1
                         y_for_explanation,                                             #y1
                         170,                                                           #x2
                         240)                                                           #y2

                doc.set_x(10)
                doc.set_y(240)
                doc.cell(130)

                if self.dataset.config.getboolean('Atlas', 'species_accounts_show_latest') and self.dataset.config.getboolean('Atlas', 'species_accounts_show_descriptions'):
                    doc.cell(10, 5, 'Species description and most recent records', 0, 0, 'L', True)
                elif self.dataset.config.getboolean('Atlas', 'species_accounts_show_latest') and not self.dataset.config.getboolean('Atlas', 'species_accounts_show_descriptions'):
                    doc.cell(10, 5, 'Most recent records', 0, 0, 'L', True)
                elif not self.dataset.config.getboolean('Atlas', 'species_accounts_show_latest') and self.dataset.config.getboolean('Atlas', 'species_accounts_show_descriptions'):
                    doc.cell(10, 5, 'Species description', 0, 0, 'L', True)

            #phenology chart
            if self.dataset.config.getboolean('Atlas', 'species_accounts_show_phenology'):
                doc.line((   ((doc.w / 2)-doc.l_margin-doc.r_margin)+3     )       / 2, #x1
                         ((doc.w / 2)-doc.l_margin-doc.r_margin)+3+10+y_padding    +22, #y1
                         40,                                                            #x2
                         260)                                                           #y2

                doc.set_x(10)
                doc.set_y(260)
                doc.cell(15)
                doc.cell(10, 5, 'Records per month', 0, 0, 'L', True)

            #status
            if self.dataset.config.getboolean('Atlas', 'species_accounts_show_status'):
                doc.line(1+x_padding                                               +10, #x1
                         11+y_padding                                              -5,  #y1
                         1+x_padding                                               +40, #x2
                         11+y_padding                                              -37) #y2

                doc.set_x(10)
                doc.set_y(11+y_padding -43)
                doc.cell(30)
                doc.cell(10, 5, 'Species status', 0, 0, 'L', True)

            #statistics
            if self.dataset.config.getboolean('Atlas', 'species_accounts_show_statistics'):
                #earliest
                doc.line(1+x_padding                                               +15, #x1
                         11+y_padding                                              +2,  #y1
                         1+x_padding                                               +50, #x2
                         11+y_padding                                              -20) #y2

                doc.set_x(10)
                doc.set_y(11+y_padding -25)
                doc.cell(45)
                doc.cell(10, 5, 'Earliest record', 0, 0, 'L', True)

                #latest
                doc.line((((doc.w / 2)-doc.l_margin-doc.r_margin)-15)+x_padding    +10, #x1
                         11+y_padding,                                                  #y1
                         (((doc.w / 2)-doc.l_margin-doc.r_margin)-15)+x_padding    +40, #x2
                         11+y_padding -40)                                              #y2

                doc.set_x(10)
                doc.set_y(11+y_padding -45)
                doc.cell(100)
                doc.cell(10, 5, 'Latest record', 0, 0, 'L', True)

                #number of records
                doc.line(1+x_padding                                               +5,  #x1
                         (((doc.w / 2)-doc.l_margin-doc.r_margin)+7)+y_padding     +5,  #y1
                         20,                                                            #x2
                         220)                                                           #y2

                doc.set_x(10)
                doc.set_y(220)
                doc.cell(1)
                doc.cell(10, 5, 'Number of records', 0, 0, 'L', True)

                #number of squares
                doc.line((((doc.w / 2)-doc.l_margin-doc.r_margin)-15)+x_padding    +15, #x1
                         (((doc.w / 2)-doc.l_margin-doc.r_margin)+7)+y_padding     +5,  #y1
                         (((doc.w / 2)-doc.l_margin-doc.r_margin)-15)+x_padding    +30, #x2
                         240                                                       -20) #y2

                doc.set_x(10)
                doc.set_y(220)
                doc.cell(70)
                doc.cell(10, 5, ' '.join(['Number of', self.dataset.config.get('Atlas', 'distribution_unit'), 'squares the species occurs in']), 0, 0, 'L', True)


            # the date classes

            doc.line((((doc.w / 2)-doc.l_margin-doc.r_margin)-15)+x_padding    -30, #x1
                     (((doc.w / 2)-doc.l_margin-doc.r_margin)+7)+y_padding     -30, #y1
                     80, #x2
                     240                                                          ) #y2            
            
            doc.set_font('Helvetica', '', 9)
            doc.set_x(70)
            doc.set_y(240)
            doc.cell(60)
            doc.cell(10, 5, 'Date classes:', 0, 1, 'L', True)

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
            count = 0
            for row in self.dataset.builder.get_object('treeview6').get_model():
                fill_colour = row[1].split('"')[1]
                border_colour = row[2].split('"')[1]
                
                r = shapefile.Reader('./markers/' + row[0] + os.path.sep + self.dataset.config.get('Atlas', 'distribution_unit'))
                shapes = r.shapes()
                pixels = []

                #grab the first marker we come to - no need to be fussy
                for x, y in shapes[0].points:
                    px = 2+int(float(x-shapes[0].bbox[0]) * scalefactor)
                    py = 2+int(float(shapes[0].bbox[3]-y) * scalefactor)
                    pixels.append((px,py))
                
                date_band = Image.new('RGB',
                                     (  4+int(float((shapes[0].bbox[2]-shapes[0].bbox[0])) * scalefactor),
                                        4+int(float((shapes[0].bbox[3]-shapes[0].bbox[1])) * scalefactor)     ),
                                     'white')
                date_band_draw = ImageDraw.Draw(date_band)
                date_band_draw.polygon(pixels, fill='rgb(' + str(int(gtk.gdk.color_parse(fill_colour).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(fill_colour).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(fill_colour).blue_float*255)) + ')', outline='rgb(' + str(int(gtk.gdk.color_parse(border_colour).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(border_colour).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(border_colour).blue_float*255)) + ')')
                #and as a line so we can increase the thickness
                date_band_draw.line(pixels, width=3, fill='rgb(' + str(int(gtk.gdk.color_parse(border_colour).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(border_colour).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(border_colour).blue_float*255)) + ')')

                date_band_temp_file = tempfile.NamedTemporaryFile(dir=self.dataset.temp_dir).name
                date_band.save(date_band_temp_file, format='PNG')


                doc.image(date_band_temp_file, 80, 245+(count*5), h=4, type='PNG')
                doc.cell(75)
                
                if row[0] == '1600.0':
                    date_band_from_text = ' '.join(['before', str(int(float(row[3])))])
                else:
                    date_band_from_text = ' '.join([str(int(float(row[3]))), 'to', str(int(float(row[4])))])
                
                doc.cell(10, 5, date_band_from_text, 0, 1, 'L', True)
                
                count = count + 1

            #### end the explanations

            #end of explanation map code###############################

        taxon_count = 0
        genus_index = {}
        species_index = {}
        common_name_index = {}

        families = []
        rownum = 0

        if self.dataset.config.get('Atlas', 'paper_size') == 'A4' and self.dataset.config.get('Atlas', 'orientation') == 'Portrait':
            max_region_count = 2
        elif self.dataset.config.get('Atlas', 'paper_size') == 'A4' and self.dataset.config.get('Atlas', 'orientation') == 'Landscape':
            max_region_count = 1
                    
        region_count = 3

        #we should really use the selection & get unique taxa?
        for item in data:
            taxon_recent_records = ''

            doc.section = ''.join(['Family ', taxa_statistics[item[0]]['family']])

            if region_count > max_region_count:
                region_count = 1
                doc.startPageNums()
                doc.p_add_page()

            if taxa_statistics[item[0]]['family'] not in families:

                families.append(taxa_statistics[item[0]]['family'])

                if self.dataset.config.getboolean('Atlas', 'toc_show_families'):
                    doc.TOC_Entry(''.join(['Family ', taxa_statistics[item[0]]['family']]), level=0)

            if self.dataset.config.getboolean('Atlas', 'toc_show_species_names') and self.dataset.config.getboolean('Atlas', 'toc_show_common_names'):
                #only show the common name if we have one
                if len(taxa_statistics[item[0]]['common_name']) > 0:
                    doc.TOC_Entry(''.join([item[0], ' - ', taxa_statistics[item[0]]['common_name']]), level=1)
                else:
                    doc.TOC_Entry(''.join([item[0]]), level=1)
            elif not self.dataset.config.getboolean('Atlas', 'toc_show_species_names') and self.dataset.config.getboolean('Atlas', 'toc_show_common_names'):
                doc.TOC_Entry(taxa_statistics[item[0]]['common_name'], level=1)
            elif self.dataset.config.getboolean('Atlas', 'toc_show_species_names') and not self.dataset.config.getboolean('Atlas', 'toc_show_common_names'):
                doc.TOC_Entry(item[0], level=1)

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
                x_padding = doc.l_margin
            elif region_count == 2: # bottom taxon map
                y_padding = 5 + (((doc.w / 2)-doc.l_margin-doc.r_margin)+3+10+y_padding) + ((((doc.w / 2)-doc.l_margin-doc.r_margin)+3)/3.75)
                x_padding = doc.l_margin

            #taxon heading
            doc.set_y(y_padding)
            doc.set_x(x_padding)
            doc.set_text_color(255)
            doc.set_fill_color(59, 59, 59)
            doc.set_line_width(0.1)
            doc.set_font('Helvetica', 'BI', 12)
            doc.cell(((doc.w)-doc.l_margin-doc.r_margin)/2, 5, item[0], 'TLB', 0, 'L', True)
            doc.set_font('Helvetica', 'B', 12)
            doc.cell(((doc.w)-doc.l_margin-doc.r_margin)/2, 5, common_name, 'TRB', 1, 'R', True)
            doc.set_x(x_padding)

            status_text = ''
            if self.dataset.config.getboolean('Atlas', 'species_accounts_show_status'):
                status_text = ''.join([designation])

            doc.multi_cell(((doc.w)-doc.l_margin-doc.r_margin), 5, status_text, 1, 'L', True)

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
            doc.set_y(y_padding+12)
            doc.set_x(x_padding+(((doc.w / 2)-doc.l_margin-doc.r_margin)+3+5))
            doc.set_font('Helvetica', '', 10)
            doc.set_text_color(0)
            doc.set_fill_color(255, 255, 255)
            doc.set_line_width(0.1)

            if len(taxa_statistics[item[0]]['description']) > 0 and self.dataset.config.getboolean('Atlas', 'species_accounts_show_descriptions'):
                doc.set_font('Helvetica', '', 10)
                doc.multi_cell((((doc.w / 2)-doc.l_margin-doc.r_margin)+12), 5, ''.join([taxa_statistics[item[0]]['description'], '\n\n']), 0, 'L', False)
                doc.set_x(x_padding+(((doc.w / 2)-doc.l_margin-doc.r_margin)+3+5))

            if self.dataset.config.getboolean('Atlas', 'species_accounts_show_latest'):
                doc.set_font('Helvetica', '', 10)
                doc.multi_cell((((doc.w / 2)-doc.l_margin-doc.r_margin)+12), 5, ''.join(['Records (most recent first): ', taxon_recent_records[:-2], '.', remaining_records_text]), 0, 'L', False)

            #chart
            if self.dataset.config.getboolean('Atlas', 'species_accounts_show_phenology'):
                chartobj = chart.Chart(self.dataset, item[0], self.dataset.config.get('Atlas', 'species_accounts_phenology_type'))
                if chartobj.temp_filename != None:
                    doc.image(chartobj.temp_filename, x_padding, ((doc.w / 2)-doc.l_margin-doc.r_margin)+3+10+y_padding, ((doc.w / 2)-doc.l_margin-doc.r_margin)+3, (((doc.w / 2)-doc.l_margin-doc.r_margin)+3)/3.75, 'PNG')
                doc.rect(x_padding, ((doc.w / 2)-doc.l_margin-doc.r_margin)+3+10+y_padding, ((doc.w / 2)-doc.l_margin-doc.r_margin)+3, (((doc.w / 2)-doc.l_margin-doc.r_margin)+3)/3.75)


            #do the map

            current_map =  self.base_map.copy()
            current_map_draw = ImageDraw.Draw(current_map)


            #loop through each date band, grabbing the records
            count = 0
            all_grids = []
            
            #if we're overlaying the markers, we draw the date bands backwards (oldest first)
            if self.dataset.config.getboolean('Atlas', 'date_band_overlay'):
                for row in treemodel_reversed:
                    fill_colour = row[1].split('"')[1]
                    border_colour = row[2].split('"')[1]
                    
                    self.dataset.cursor.execute('SELECT DISTINCT(grid_' + self.dataset.config.get('Atlas', 'distribution_unit') + ') AS grids \
                                                FROM data \
                                                WHERE data.taxon = "' + item[0] + '" \
                                                AND data.year_to >= ' + str(row[3]) + '\
                                                AND data.year_to < ' + str(row[4]) + ' \
                                                AND data.year_from >= ' + str(row[3]) + ' \
                                                AND data.year_from < ' + str(row[4]))
    
                    date_b_grids = []
                    date_band_grids = self.dataset.cursor.fetchall()
                    for g in date_band_grids:
                        date_b_grids.append(g[0])
                        
                    #loop through each object in the coverage (to save having to loop through ALL objects)
                    for obj in self.date_band_coverage[count]:
                        #if the object is in our date band
                        if obj.record[0] in date_b_grids:
                            
                            pixels = []
                            #loop through each point in the object
                            for x,y in obj.shape.points:
                                px = (self.xdist * self.scalefactor)- (self.bounds_top_x - x) * self.scalefactor
                                py = (self.bounds_top_y - y) * self.scalefactor
                                pixels.append((px,py))
                            current_map_draw.polygon(pixels, fill='rgb(' + str(int(gtk.gdk.color_parse(fill_colour).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(fill_colour).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(fill_colour).blue_float*255)) + ')', outline='rgb(' + str(int(gtk.gdk.color_parse(border_colour).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(border_colour).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(border_colour).blue_float*255)) + ')')
    
                    count = count + 1
    
                    #add the current date band grids for reference next time round
                    for g in date_band_grids:
                        all_grids.append(g[0])

            #if we're not overlaying the markers, we draw the date bands forwards (newest first) and skip any markers that have already been drawn                        
            else:
                for row in self.dataset.builder.get_object('treeview6').get_model():
                    fill_colour = row[1].split('"')[1]
                    border_colour = row[2].split('"')[1]
                    
                    self.dataset.cursor.execute('SELECT DISTINCT(grid_' + self.dataset.config.get('Atlas', 'distribution_unit') + ') AS grids \
                                                FROM data \
                                                WHERE data.taxon = "' + item[0] + '" \
                                                AND data.year_to >= ' + str(row[3]) + '\
                                                AND data.year_to < ' + str(row[4]) + ' \
                                                AND data.year_from >= ' + str(row[3]) + ' \
                                                AND data.year_from < ' + str(row[4]))
    
                    date_b_grids = []
                    date_band_grids = self.dataset.cursor.fetchall()
                    for g in date_band_grids:
                        date_b_grids.append(g[0])
                        
                    #loop through each object in the coverage (to save having to loop through ALL objects)
                    for obj in self.date_band_coverage[count]:
                        #if the object is in our date band and it's not already been drawn by a more recent date band
                        if (obj.record[0] in date_b_grids) and (obj.record[0] not in all_grids):
                            
                            pixels = []
                            #loop through each point in the object
                            for x,y in obj.shape.points:
                                px = (self.xdist * self.scalefactor)- (self.bounds_top_x - x) * self.scalefactor
                                py = (self.bounds_top_y - y) * self.scalefactor
                                pixels.append((px,py))
                            current_map_draw.polygon(pixels, fill='rgb(' + str(int(gtk.gdk.color_parse(fill_colour).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(fill_colour).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(fill_colour).blue_float*255)) + ')', outline='rgb(' + str(int(gtk.gdk.color_parse(border_colour).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(border_colour).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(border_colour).blue_float*255)) + ')')
    
                    count = count + 1
    
                    #add the current date band grids for reference next time round
                    for g in date_band_grids:
                        all_grids.append(g[0])






            temp_map_file = tempfile.NamedTemporaryFile(dir=self.dataset.temp_dir).name
            current_map.save(temp_map_file, format='PNG')

            (width, height) =  current_map.size

            max_width = (doc.w / 2)-doc.l_margin-doc.r_margin
            max_height = (doc.w / 2)-doc.l_margin-doc.r_margin

            if height > width:
                pdfim_height = max_width
                pdfim_width = (float(width)/float(height))*max_width
            else:
                pdfim_height = (float(height)/float(width))*max_width
                pdfim_width = max_width

            img_x_cent = ((max_width-pdfim_width)/2)+2
            img_y_cent = ((max_height-pdfim_height)/2)+12

            doc.image(temp_map_file, x_padding+img_x_cent, y_padding+img_y_cent, int(pdfim_width), int(pdfim_height), 'PNG')


            #map container
            doc.set_text_color(0)
            doc.set_fill_color(255, 255, 255)
            doc.set_line_width(0.1)
            doc.rect(x_padding, 10+y_padding, ((doc.w / 2)-doc.l_margin-doc.r_margin)+3, ((doc.w / 2)-doc.l_margin-doc.r_margin)+3)

            if self.dataset.config.getboolean('Atlas', 'species_accounts_show_statistics'):

                #from
                doc.set_y(11+y_padding)
                doc.set_x(1+x_padding)
                doc.set_font('Helvetica', '', 12)
                doc.set_text_color(0)
                doc.set_fill_color(255, 255, 255)
                doc.set_line_width(0.1)

                if str(taxa_statistics[item[0]]['earliest']) == 'Unknown':
                    val = '?'
                else:
                    val = str(taxa_statistics[item[0]]['earliest'])
                doc.multi_cell(18, 5, ''.join(['E ', val]), 0, 'L', False)

                #to
                doc.set_y(11+y_padding)
                doc.set_x((((doc.w / 2)-doc.l_margin-doc.r_margin)-15)+x_padding)
                doc.set_font('Helvetica', '', 12)
                doc.set_text_color(0)
                doc.set_fill_color(255, 255, 255)
                doc.set_line_width(0.1)

                if str(taxa_statistics[item[0]]['latest']) == 'Unknown':
                    val = '?'
                else:
                    val = str(taxa_statistics[item[0]]['latest'])
                doc.multi_cell(18, 5, ''.join(['L ', val]), 0, 'R', False)

                #records
                doc.set_y((((doc.w / 2)-doc.l_margin-doc.r_margin)+7)+y_padding)
                doc.set_x(1+x_padding)
                doc.set_font('Helvetica', '', 12)
                doc.set_text_color(0)
                doc.set_fill_color(255, 255, 255)
                doc.set_line_width(0.1)
                doc.multi_cell(18, 5, ''.join(['R ', str(taxa_statistics[item[0]]['count'])]), 0, 'L', False)

                #squares
                doc.set_y((((doc.w / 2)-doc.l_margin-doc.r_margin)+7)+y_padding)
                doc.set_x((((doc.w / 2)-doc.l_margin-doc.r_margin)-15)+x_padding)
                doc.set_font('Helvetica', '', 12)
                doc.set_text_color(0)
                doc.set_fill_color(255, 255, 255)
                doc.set_line_width(0.1)
                doc.multi_cell(18, 5, ''.join(['S ', str(taxa_statistics[item[0]]['dist_count'])]), 0, 'R', False)

            taxon_parts = item[0].split(' ')

            if taxon_parts[0].lower() not in genus_index:
                genus_index[taxon_parts[0].lower()] = ['genus', doc.num_page_no()+doc.toc_length]

            if taxa_statistics[item[0]]['common_name'] not in common_name_index and taxa_statistics[item[0]]['common_name'] != 'None':
                common_name_index[taxa_statistics[item[0]]['common_name']] = ['common', doc.num_page_no()+doc.toc_length]

            if taxon_parts[len(taxon_parts)-1] == 'agg.':
                part = ' '.join([taxon_parts[len(taxon_parts)-2], taxon_parts[len(taxon_parts)-1]])
            else:
                part = taxon_parts[len(taxon_parts)-1]

            if ''.join([part, ' ', taxon_parts[0].lower()]) not in species_index:
                species_index[''.join([part, ', ', taxon_parts[0]]).lower()] = ['species', doc.num_page_no()+doc.toc_length]

            taxon_count = taxon_count + 1


            rownum = rownum + 1

            region_count = region_count + 1

        doc.section = ''
        #doc.do_header = False

        index = genus_index.copy()
        index.update(species_index)
        index.update(common_name_index)

        #index = sorted(index.iteritems())

        doc.p_add_page()

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

        doc.set_y(19)
        doc.set_font('Helvetica', 'I', 10)
        doc.multi_cell(0, 5, ' '.join([family_text,
                                       ' '.join([str(len(data)), taxa_text]),
                                       ' '.join(['mapped from', str(record_count), record_text]),]),
                       0, 'J', False)

        #doc.section = ''
        #doc.stopPageNums()
        doc.section = 'Index'
        doc.p_add_page()

        initial = ''

        doc.set_y(doc.y0+20)
        for taxon in sorted(index, key=lambda taxon: taxon.lower()):
            try:
                if taxon[0].upper() != initial:
                    if taxon[0].upper() != 'A':
                        doc.ln(3)
                    doc.set_font('Helvetica', 'B', 12)
                    doc.cell(0, 5, taxon[0].upper(), 0, 1, 'L', 0)
                    initial = taxon[0].upper()

                if index[taxon][0] == 'species':
                    pos = taxon.find(', ')
                    #capitalize the first letter of the genus
                    display_taxon = list(taxon)
                    display_taxon[pos+2] = display_taxon[pos+2].upper()
                    display_taxon = ''.join(display_taxon)
                    doc.set_font('Helvetica', '', 12)
                    doc.cell(0, 5, '  '.join([display_taxon, str(index[taxon][1])]), 0, 1, 'L', 0)
                elif index[taxon][0] == 'genus':
                    doc.set_font('Helvetica', '', 12)
                    doc.cell(0, 5, '  '.join([taxon.upper(), str(index[taxon][1])]), 0, 1, 'L', 0)
                elif index[taxon][0] == 'common':
                    doc.set_font('Helvetica', '', 12)
                    doc.cell(0, 5, '  '.join([taxon, str(index[taxon][1])]), 0, 1, 'L', 0)
            except IndexError:
                pass

        doc.setcol(0)

        doc.section = 'Contributors'
        doc.p_add_page()
        doc.set_font('Helvetica', '', 20)
        doc.multi_cell(0, 20, ''.join(['Contributors', ' (', str(len(contrib_data)), ')']), 0, 'J', False)
        doc.set_font('Helvetica', '', 8)

        contrib_blurb = []

        for name in sorted(contrib_data.keys()):
            if name != 'Unknown' and name != 'Unknown Unknown':
                contrib_blurb.append(''.join([name, ' (', contrib_data[name], ')']))

        doc.multi_cell(0, 5, ''.join([', '.join(contrib_blurb), '.']), 0, 'J', False)

        if len(self.dataset.config.get('Atlas', 'bibliography')) > 0:
            doc.section = ('References')
            doc.p_add_page()
            doc.set_font('Helvetica', '', 20)
            doc.multi_cell(0, 20, 'References', 0, 'J', False)
            doc.set_font('Helvetica', '', 10)
            doc.multi_cell(0, 6, self.dataset.config.get('Atlas', 'bibliography'), 0, 'J', False)
            
        doc.section = ''

        doc.set_y(-30)
        doc.set_font('Helvetica','',8)

        if doc.num_page_no() >= 4 and doc.section != 'Contents':
            doc.cell(0, 10, 'Generated in seconds using dipper-stda. For more information, see https://github.com/charlie-barnes/dipper-stda.', 0, 1, 'L')
            doc.cell(0, 10, ''.join(['Vice-county boundaries provided by the National Biodiversity Network. Contains Ordnance Survey data (C) Crown copyright and database right ', str(datetime.now().year), '.']), 0, 1, 'L')

        #doc.p_add_page()
        doc.section = ''

        #toc
        doc.insertTOC(3)

        #output
        try:
            doc.output(self.save_in,'F')
        except IOError:
            md = gtk.MessageDialog(None,
                gtk.DIALOG_DESTROY_WITH_PARENT, gtk.MESSAGE_ERROR,
                gtk.BUTTONS_CLOSE, 'Unable to write to file. This usually means it''s open - close it and try again.')
            md.run()
            md.destroy()       

