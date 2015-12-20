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

import gtk
from pygtk_chart import bar_chart
import pango
import tempfile

class Chart(gtk.Window):

    def __init__(self, dataset, item, mode):
        gtk.Window.__init__(self)
        vbox = gtk.VBox()
        self.add(vbox)
        combo_box = gtk.combo_box_new_text()
        vbox.pack_end(combo_box, False, False, 0)

        self.set_decorated(False)
        self.item = item

        self.data = None
        self.division = None
        self.minimum_decade = None
        self.maximum_decade = None
        self.months = None
        self.decades = None
        self.mode = mode
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

    def __redraw__(self):
        # this should be done more efficently. init the bar chart with null values, then just update?

        if self.mode == 'Months':
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
        elif self.mode == 'Decades':

            self.dataset.cursor.execute('SELECT MIN(data.decade), MAX(data.decade) \
                                        FROM data \
                                        WHERE data.decade IS NOT NULL')

            datas = self.dataset.cursor.fetchall()
            
            data = []
            count = 1
            date_range = range(datas[0][0], datas[0][1], 10)
            
            for decade in date_range:
                if len(date_range) > 5:
                    if count % 3 == 0:
                        label = decade
                    else:
                        label = ''
                try:
                    data.append((label, self.data[decade], label))
                except KeyError:
                    data.append((label, 0, label))     
                
                count = count + 1           
            
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
                #bar._label_object.set_rotation(180)
                bar._label_object.set_property('wrap', False)
                barchart.add_bar(bar)
            
            #HACK - Draw a hidden bar with the max value, so we can have an 'empty' chart            
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

        if self.mode == 'Months':
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

        elif self.mode == 'Decades':

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
