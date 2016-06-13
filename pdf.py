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

def repeat_to_length(string_to_expand, length):
   return (string_to_expand * ((length/len(string_to_expand))+1))[:length]
   
try:
    from fpdf import FPDF
except ImportError:
    from pyfpdf import FPDF
    
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
        self.orientation = orientation

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
        if self.num_page_no() >= 4 and self.section != 'Contents' and self.section != 'Index' and self.section != 'Contributors' and self.section != 'References' and self.section != 'Introduction' and self.section != '':
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
            if (self.orientation == 'Portrait' and self.col == 0) or (self.orientation == 'Landscape' and (self.col == 0 or self.col == 1)) :
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

