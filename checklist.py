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
import os
import json

def repeat_to_length(string_to_expand, length):
   return (string_to_expand * ((length/len(string_to_expand))+1))[:length]
   
class Checklist(gobject.GObject):

    def __init__(self, dataset):
        gobject.GObject.__init__(self)
        self.dataset = dataset
        self.page_unit = 'mm'
        self.save_in = None

    def generate(self):

        taxa_statistics = {}
        taxon_list = []

        if len(self.dataset.config.get('Checklist', 'vice-counties')) > 0:
            self.dataset.use_vcs = True
        else:
            self.dataset.use_vcs = False
        
        if self.dataset.use_vcs:
            vcs_sql = ''.join(['data.vc IN (', self.dataset.config.get('Checklist', 'vice-counties'), ') AND'])
            vcs_sql_sel = 'data.vc'
        else:
            vcs_sql = ''
            vcs_sql_sel = '"00"'


        restriction_sql = ''                    
                    
        for restriction in json.loads(self.dataset.config.get('Checklist', 'families')):
            restriction_sql = ' or '.join([restriction_sql, ''.join(['species_data.' , restriction[0] , ' = "' , restriction[1] + '"'])])

        restriction_sql = restriction_sql[4:]

        self.dataset.cursor.execute('SELECT DISTINCT data.taxon \
                                   FROM data \
                                   JOIN species_data ON data.taxon = species_data.taxon \
                                   WHERE ' + vcs_sql + ' (' + restriction_sql + ')')

        data = self.dataset.cursor.fetchall()
        number_of_species = len(data)                                   

        self.dataset.cursor.execute('SELECT DISTINCT species_data.family \
                                   FROM data \
                                   JOIN species_data ON data.taxon = species_data.taxon \
                                   WHERE ' + vcs_sql + ' (' + restriction_sql + ')')

        data = self.dataset.cursor.fetchall()
        number_of_families = len(data)                                     

        self.dataset.cursor.execute('SELECT data.taxon, species_data.family, species_data.national_status, species_data.local_status, \
                                   COUNT(DISTINCT(grid_' + self.dataset.config.get('Checklist', 'distribution_unit') + ')) AS squares, \
                                   COUNT(data.taxon) AS records, \
                                   MAX(data.year) AS year, \
                                   ' + vcs_sql_sel + ' AS VC \
                                   FROM data \
                                   JOIN species_data ON data.taxon = species_data.taxon \
                                   WHERE ' + vcs_sql + ' (' + restriction_sql + ') \
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
        doc = pdf.PDF(orientation=self.dataset.config.get('Checklist', 'orientation'),unit=self.page_unit,format=self.dataset.config.get('Checklist', 'paper_size'))
        doc.type = 'Checklist'
        doc.do_header = False
        doc.dataset = self.dataset

        doc.col = 0
        doc.y0 = 0
        doc.set_title(self.dataset.config.get('Checklist', 'title'))
        doc.set_author(self.dataset.config.get('Checklist', 'author'))
        doc.set_creator(' '.join(['dipper-stda', version.__version__])) 
        doc.section = ''

        #title page
        doc.p_add_page()

        if self.dataset.config.get('Checklist', 'cover_image') is not None and os.path.isfile(self.dataset.config.get('Checklist', 'cover_image')):
            doc.image(self.dataset.config.get('Checklist', 'cover_image'), 0, 0, doc.w, doc.h)

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
        doc.do_header = True
        doc.set_font('Helvetica', '', 12)
        doc.multi_cell(0, 6, self.dataset.config.get('Checklist', 'inside_cover'), 0, 'J', False)

        #introduction
        if len(self.dataset.config.get('Checklist', 'introduction')) > 0:
            doc.section = ('Introduction')       
            doc.p_add_page() 
            doc.do_header = True
            doc.startPageNums()
            doc.set_font('Helvetica', '', 20)
            doc.cell(0, 20, 'Introduction', 0, 0, 'L', 0)
            doc.ln()
            doc.set_font('Helvetica', '', 12)
            doc.multi_cell(0, 6, self.dataset.config.get('Checklist', 'introduction'), 0, 0, 'L')
            doc.ln()
        else:
            doc.section = (' '.join(['Family', data[0][1].upper()]))     
            doc.p_add_page()
            doc.ln()
            

        #main heading            
        doc.set_font('Helvetica', '', 20)
        doc.cell(0, 15, 'Checklist', 0, 1, 'L', 0)

        col_width = 12.7#((self.w - self.l_margin - self.r_margin)/2)/7.5

        #vc headings
        doc.set_font('Helvetica', '', 10)
        doc.set_line_width(0.0)

        doc.set_x(doc.w-(7+col_width+(((col_width*3)+(col_width/4))*len(self.dataset.config.get('Checklist', 'vice-counties').split(',')))))

        doc.cell(col_width, 5, '', '0', 0, 'C', 0)
        
        if self.dataset.use_vcs:
            for vc in sorted(self.dataset.config.get('Checklist', 'vice-counties').split(',')):
                doc.vcs = self.dataset.config.get('Checklist', 'vice-counties').split(',')
                doc.cell((col_width*3), 5, ''.join(['VC',vc]), '0', 0, 'C', 0)
                doc.cell(col_width/4, 5, '', '0', 0, 'C', 0)
        else:
            doc.vcs = [None]
            doc.cell((col_width*3), 5, '', '0', 0, 'C', 0)
            doc.cell(col_width/4, 5, '', '0', 0, 'C', 0)                


        doc.ln()

        doc.set_x(doc.w-(7+col_width+(((col_width*3)+(col_width/4))*len(self.dataset.config.get('Checklist', 'vice-counties').split(',')))))
        doc.set_font('Helvetica', '', 8)
        doc.cell(col_width, 5, '', '0', 0, 'C', 0)

        for vc in sorted(doc.vcs):
            #colum headings
            doc.cell(col_width, 5, ' '.join([self.dataset.config.get('Checklist', 'distribution_unit'), 'sqs']), '0', 0, 'C', 0)
            doc.cell(col_width, 5, 'Records', '0', 0, 'C', 0)
            doc.cell(col_width, 5, 'Last in', '0', 0, 'C', 0)
            doc.cell(col_width/4, 5, '', '0', 0, 'C', 0)

        doc.doing_the_list = True
        doc.set_font('Helvetica', '', 8)
        
        if self.dataset.use_vcs:
            for vckey in sorted(self.dataset.config.get('Checklist', 'vice-counties').split(',')):
                #print vckey
    
                col = self.dataset.config.get('Checklist', 'vice-counties').split(',').index(vckey)+1
    
                doc.cell(col_width/col, 5, '', '0', 0, 'C', 0)
                doc.cell(col_width, 5, ' '.join([self.dataset.config.get('Checklist', 'distribution_unit'), 'sqs']), '0', 0, 'C', 0)
                doc.cell(col_width, 5, 'Records', '0', 0, 'C', 0)
                doc.cell(col_width, 5, 'Last in', '0', 0, 'C', 0)
        else:
            doc.cell(col_width/1, 5, '', '0', 0, 'C', 0)
            doc.cell(col_width, 5, ' '.join([self.dataset.config.get('Checklist', 'distribution_unit'), 'sqs']), '0', 0, 'C', 0)
            doc.cell(col_width, 5, 'Records', '0', 0, 'C', 0)
            doc.cell(col_width, 5, 'Last in', '0', 0, 'C', 0)                        

        taxon_count = 0

        record_count = 0
        families = []
        for key in taxon_list:
            if taxa_statistics[key]['family'] not in families:
                doc.section = (''.join(['Family ', taxa_statistics[key]['family'].upper()]))
                if doc.get_y() > (doc.h - 42):
                    doc.p_add_page()
                doc.set_font('Helvetica', 'B', 10)

                doc.set_fill_color(50, 50, 50)
                doc.set_text_color(255, 255, 255)

                doc.ln()
                doc.cell(0, 5, ''.join(['Family ', taxa_statistics[key]['family'].upper()]), 0, 1, 'L', 1)
                families.append(taxa_statistics[key]['family'])

                doc.set_fill_color(255, 255, 255)
                doc.set_text_color(0, 0, 0)

            elif doc.get_y() > (doc.h - 27):
                doc.section = (''.join(['Family ', taxa_statistics[key]['family'].upper()]))
                doc.p_add_page()
                doc.set_y(30)

            #species name
            doc.set_font('Helvetica', '', 10)
            strsize = doc.get_string_width(key)+2
            doc.cell(strsize, doc.font_size+2, key, '', 0, 'L', 0)

            #dots
            w = doc.w-(4+col_width+col_width+(((col_width*3)+(col_width/4))*len(doc.vcs))) - strsize
            nb = w/doc.get_string_width('.')

            dots = repeat_to_length('.', int(nb))
            doc.cell(w, doc.font_size+2, dots, 0, 0, 'R', 0)

            doc.set_font('Helvetica', '', 6)
            doc.cell(col_width, doc.font_size+3, taxa_statistics[key]['national_designation'], '', 0, 'L', 0)
            doc.set_font('Helvetica', '', 10)

            if self.dataset.use_vcs:
                for vckey in sorted(self.dataset.config.get('Checklist', 'vice-counties').split(',')):
                    #print vckey
    
                    doc.set_fill_color(230, 230, 230)
                    try:
    
                        if taxa_statistics[key]['vc'][vckey]['squares'] == '0':
                            squares = '-'
                        else:
                            squares = taxa_statistics[key]['vc'][vckey]['squares']
    
                        doc.cell(col_width, doc.font_size+2, squares, '', 0, 'L', 1)
                        doc.cell(col_width, doc.font_size+2, taxa_statistics[key]['vc'][vckey]['records'], '', 0, 'L', 1)
                        record_count = record_count + int(taxa_statistics[key]['vc'][vckey]['records'])
    
                        if taxa_statistics[key]['vc'][vckey]['year'] == 'None':
                            doc.cell(col_width, doc.font_size+2, '?', '', 0, 'L', 1)
                        else:
                            doc.cell(col_width, doc.font_size+2, taxa_statistics[key]['vc'][vckey]['year'], '', 0, 'C', 1)
    
                    except KeyError:
                        doc.cell(col_width, doc.font_size+2, '', '', 0, 'L', 1)
                        doc.cell(col_width, doc.font_size+2, '', '', 0, 'L', 1)
                        doc.cell(col_width, doc.font_size+2, '', '', 0, 'C', 1)
    
    
                    doc.set_fill_color(255, 255, 255)
    
                    doc.cell((col_width/4), doc.font_size+2, '', 0, 0, 'C', 0)
            else:
 
                  doc.set_fill_color(230, 230, 230)
                  try:
  
                      if taxa_statistics[key]['vc']['00']['squares'] == '0':
                          squares = '-'
                      else:
                          squares = taxa_statistics[key]['vc']['00']['squares']
  
                      doc.cell(col_width, doc.font_size+2, squares, '', 0, 'L', 1)
                      doc.cell(col_width, doc.font_size+2, taxa_statistics[key]['vc']['00']['records'], '', 0, 'L', 1)
                      record_count = record_count + int(taxa_statistics[key]['vc']['00']['records'])
  
                      if taxa_statistics[key]['vc']['00']['year'] == 'None':
                          doc.cell(col_width, doc.font_size+2, '?', '', 0, 'L', 1)
                      else:
                          doc.cell(col_width, doc.font_size+2, taxa_statistics[key]['vc']['00']['year'], '', 0, 'C', 1)
  
                  except KeyError:
                      doc.cell(col_width, doc.font_size+2, '', '', 0, 'L', 1)
                      doc.cell(col_width, doc.font_size+2, '', '', 0, 'L', 1)
                      doc.cell(col_width, doc.font_size+2, '', '', 0, 'C', 1)
  
  
                  doc.set_fill_color(255, 255, 255)
  
                  doc.cell((col_width/4), doc.font_size+2, '', 0, 0, 'C', 0)                                


            doc.ln()

            while gtk.events_pending():
                gtk.main_iteration()

            taxon_count = taxon_count + 1


        doc.section = ''
        doc.doing_the_list = False

        if len(families) > 1:
            family_text = ' '.join([str(number_of_families), 'families and'])
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

        doc.set_y(doc.get_y()+10)
        doc.set_font('Helvetica', 'I', 10)
        doc.multi_cell(0, 5, ' '.join([family_text,
                                       ' '.join([str(number_of_species), taxa_text]),
                                       ' '.join(['listed from', str(record_count), record_text]),]),
                       0, 'J', False)


        #output
        try:
            doc.output(self.save_in,'F')
        except IOError:
            md = gtk.MessageDialog(None,
                gtk.DIALOG_DESTROY_WITH_PARENT, gtk.MESSAGE_ERROR,
                gtk.BUTTONS_CLOSE, 'Unable to write to file. This usually means it''s open - close it and try again.')
            md.run()
            md.destroy()        


