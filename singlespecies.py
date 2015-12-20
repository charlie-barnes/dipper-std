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
import chart

class SingleSpecies(gobject.GObject):

    def __init__(self, dataset):
        gobject.GObject.__init__(self)
        self.dataset = dataset
        self.save_in = None
        self.page_unit = 'mm'
        self.base_map = None
        self.date_band_coverage = []
        self.increments = 14               
        
        
    def generate_base_map(self):

        #generate the base map
        self.scalefactor = 0.0035

        if self.dataset.use_vcs:
            vcs_sql = ''.join(['WHERE data.vc IN (', self.dataset.config.get('Species', 'vice-counties'), ')'])
        else:
            vcs_sql = ''

        layers = []
        for vc in self.dataset.config.get('Species', 'vice-counties').split(','):
            layers.append(''.join(['./vice-counties/',cfg.vc_list[int(vc)-1][1],'.shp']))

        #add the total coverage & calc first and date band 2 grid arrays
        self.dataset.cursor.execute('SELECT DISTINCT(grid_' + self.dataset.config.get('Species', 'distribution_unit') + ') AS grids \
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
        r = shapefile.Reader('./markers/' + self.dataset.config.get('Species', 'coverage_style') + '/' + self.dataset.config.get('Species', 'distribution_unit'))
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


        #loop through the date bands treeview
        for row in self.dataset.builder.get_object('treeview6').get_model():
            # Read in the date band 1 grid ref shapefiles and extend the bounding box
            r = shapefile.Reader('./markers/' + row[0] + '/' + self.dataset.config.get('Species', 'distribution_unit'))
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
                    
            r = shapefile.Reader('./markers/' + row[0] + '/' + self.dataset.config.get('Species', 'distribution_unit'))
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
                        self.base_map_draw.polygon(pixels, fill='rgb(' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Species', 'vice-counties_fill')).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Species', 'vice-counties_fill')).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Species', 'vice-counties_fill')).blue_float*255)) + ')', outline='rgb(' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Species', 'vice-counties_outline')).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Species', 'vice-counties_outline')).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Species', 'vice-counties_outline')).blue_float*255)) + ')')
                        pixels = []
                    counter = counter + 1
                #draw the final polygon (or the only, if we have just the one)
                self.base_map_draw.polygon(pixels, fill='rgb(' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Species', 'vice-counties_fill')).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Species', 'vice-counties_fill')).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Species', 'vice-counties_fill')).blue_float*255)) + ')', outline='rgb(' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Species', 'vice-counties_outline')).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Species', 'vice-counties_outline')).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Species', 'vice-counties_outline')).blue_float*255)) + ')')



        #add the coverage
        r = shapefile.Reader('./markers/' + self.dataset.config.get('Species', 'coverage_style') + '/' + self.dataset.config.get('Species', 'distribution_unit'))
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
                    self.base_map_draw.polygon(pixels, fill='rgb(' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Species', 'coverage_colour')).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Species', 'coverage_colour')).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Species', 'coverage_colour')).blue_float*255)) + ')')


        #add the grid lines
        if self.dataset.config.getboolean('Atlas', 'grid_lines_visible'):
            r = shapefile.Reader('./markers/squares/' + self.dataset.config.get('Species', 'grid_lines_style'))
            #loop through each object in the shapefile
            for obj in r.shapes():
                pixels = []
                #loop through each point in the object
                for x,y in obj.points:
                    px = (self.xdist * self.scalefactor)- (self.bounds_top_x - x) * self.scalefactor
                    py = (self.bounds_top_y - y) * self.scalefactor
                    pixels.append((px,py))
                self.base_map_draw.polygon(pixels, outline='rgb(' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Species', 'grid_lines_colour')).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Species', 'grid_lines_colour')).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Species', 'grid_lines_colour')).blue_float*255)) + ')')

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
                        self.base_map_draw.polygon(pixels, fill='rgb(' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Species', 'vice-counties_fill')).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Species', 'vice-counties_fill')).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Species', 'vice-counties_fill')).blue_float*255)) + ')', outline='rgb(' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Species', 'vice-counties_outline')).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Species', 'vice-counties_outline')).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Species', 'vice-counties_outline')).blue_float*255)) + ')')
                        pixels = []
                    counter = counter + 1
                #draw the final polygon (or the only, if we have just the one)
                self.base_map_draw.polygon(pixels, outline='rgb(' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Species', 'vice-counties_outline')).red_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Species', 'vice-counties_outline')).green_float*255)) + ',' + str(int(gtk.gdk.color_parse(self.dataset.config.get('Species', 'vice-counties_outline')).blue_float*255)) + ')')


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
            vcs_sql = ''.join(['data.vc IN (', self.dataset.config.get('Species', 'vice-counties'), ') AND'])
        else:
            vcs_sql = ''
            
        species_sql = ''.join(['species_data.taxon = "', '","'.join(self.dataset.config.get('Species', 'species').split(',')), '"'])

        self.dataset.cursor.execute('SELECT data.taxon, species_data.family, species_data.national_status, species_data.local_status, COUNT(data.taxon), MIN(data.year), MAX(data.year), COUNT(DISTINCT(grid_' + self.dataset.config.get('Species', 'distribution_unit') + ')), \
                                   COUNT(DISTINCT(grid_' + self.dataset.config.get('Species', 'distribution_unit') + ')) AS squares, \
                                   COUNT(data.taxon) AS records, \
                                   MAX(data.year) AS year, \
                                   species_data.description, \
                                   species_data.common_name \
                                   FROM data \
                                   JOIN species_data ON data.taxon = species_data.taxon \
                                   WHERE ' + vcs_sql + ' ' + species_sql + ' \
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

        #the pdf
        doc = pdf.PDF(orientation=self.dataset.config.get('Species', 'orientation'),unit=self.page_unit,format=self.dataset.config.get('Species', 'paper_size'))
        doc.type = 'atlas'
        doc.toc_length = toc_length

        doc.col = 0
        doc.y0 = 0
        doc.set_title(self.dataset.config.get('Species', 'title'))
        doc.set_author(self.dataset.config.get('Species', 'author'))
        doc.set_creator(' '.join(['dipper-stda', version.__version__])) 
        doc.section = ''

        families = []
        rownum = 0

        if self.dataset.config.get('Species', 'paper_size') == 'A4' and self.dataset.config.get('Species', 'orientation') == 'Portrait':
            max_region_count = 2
        elif self.dataset.config.get('Species', 'paper_size') == 'A4' and self.dataset.config.get('Species', 'orientation') == 'Landscape':
            max_region_count = 1
                    
        region_count = 3

        #we should really use the selection & get unique taxa?
        for item in data:
            taxon_recent_records = ''

            if region_count > max_region_count:
                region_count = 1
                doc.startPageNums()
                doc.p_add_page()

            if taxa_statistics[item[0]]['family'] not in families:

                families.append(taxa_statistics[item[0]]['family'])

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
            self.dataset.cursor.execute('SELECT data.taxon, data.location, data.grid_native, data.grid_' + self.dataset.config.get('Species', 'distribution_unit') + ', data.date, data.decade_to, data.year_to, data.month_to, data.recorder, data.determiner, data.vc, data.grid_100m \
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
                    current_rec = self.dataset.config.get('Species', 'species_accounts_latest_format')

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
                doc.set_font('Helvetica', 'B', 10)
                doc.multi_cell((((doc.w / 2)-doc.l_margin-doc.r_margin)+12), 5, ''.join([taxa_statistics[item[0]]['description'], '\n\n']), 0, 'L', False)
                doc.set_x(x_padding+(((doc.w / 2)-doc.l_margin-doc.r_margin)+3+5))

            if self.dataset.config.getboolean('Atlas', 'species_accounts_show_latest'):
                doc.set_font('Helvetica', '', 10)
                doc.multi_cell((((doc.w / 2)-doc.l_margin-doc.r_margin)+12), 5, ''.join(['Records (most recent first): ', taxon_recent_records[:-2], '.', remaining_records_text]), 0, 'L', False)

            #chart
            if self.dataset.config.getboolean('Atlas', 'species_accounts_show_phenology'):
                chart = Chart(self.dataset, item[0], self.dataset.config.get('Species', 'species_accounts_phenology_type'))
                if chart.temp_filename != None:
                    doc.image(chart.temp_filename, x_padding, ((doc.w / 2)-doc.l_margin-doc.r_margin)+3+10+y_padding, ((doc.w / 2)-doc.l_margin-doc.r_margin)+3, (((doc.w / 2)-doc.l_margin-doc.r_margin)+3)/3.75, 'PNG')
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
                    
                    self.dataset.cursor.execute('SELECT DISTINCT(grid_' + self.dataset.config.get('Species', 'distribution_unit') + ') AS grids \
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
                    
                    self.dataset.cursor.execute('SELECT DISTINCT(grid_' + self.dataset.config.get('Species', 'distribution_unit') + ') AS grids \
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

            taxon_count = taxon_count + 1


            rownum = rownum + 1

            region_count = region_count + 1

        #output
        try:
            doc.output(self.save_in,'F')
        except IOError:
            md = gtk.MessageDialog(None,
                gtk.DIALOG_DESTROY_WITH_PARENT, gtk.MESSAGE_ERROR,
                gtk.BUTTONS_CLOSE, 'Unable to write to file. This usually means it''s open - close it and try again.')
            md.run()
            md.destroy()  
