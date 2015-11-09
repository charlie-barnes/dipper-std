#!/usr/bin/env python
#-*- coding: utf-8 -*-

### 2008-2010 Charlie Barnes.

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

### TODO: date formats: 05.vi.2010

from calendar import monthrange

class VagueDate(object):

    def __init__(self, s):
        """ Analyse the string to see if we are dealing with a single date or a
        range, and return a list of (year, month, day).
        
        """

        s = str(s)

        self.decade = None
        self.year = None
        self.month = None
        self.day = None

        self.decade_from = None
        self.year_from = None
        self.month_from = None
        self.day_from = None

        self.decade_to = None
        self.year_to = None
        self.month_to = None
        self.day_to = None
        
        if 'spring' in s.lower():
            s = ' '.join(['March', s[-4:], 'to', 'May', s[-4:]])
        elif 'summer' in s.lower():
            s = ' '.join(['June', s[-4:], 'to', 'August', s[-4:]])
        elif 'autumn' in s.lower():
            s = ' '.join(['September', s[-4:], 'to', 'November', s[-4:]])
        elif 'winter' in s.lower():
            try:
                s = ' '.join(['December', s[-4:], 'to', 'February', str(int(s[-4:])+1)])
            except ValueError:
                s = ' '.join(['December', 'to', 'February'])
                

        # are we dealing with a date range?
        if s.count(' - ') == 1:
            range_splitter = ' - '
        elif s.count(' to ') == 1:
            range_splitter = ' to '
        elif (s.count('-') == 1) and ((s.count('/') > 0) or
                                      ((s.count('-')-1) > 0) or
                                      (s.count('\\') > 0) or
                                      (s.count('.') > 0) or
                                      (s.count(' ') > 0)):
            range_splitter = '-'
        
        try:
            from_date, to_date = s.split(range_splitter)
            self.decade_from, self.year_from, self.month_from, self.day_from = self.parse(from_date)
            self.decade_to, self.year_to, self.month_to, self.day_to = self.parse(to_date)    
        except:
            self.decade, self.year, self.month, self.day = self.parse(s)
            self.decade_from, self.year_from, self.month_from, self.day_from = self.parse(s)
            self.decade_to, self.year_to, self.month_to, self.day_to = self.parse(s)

        if self.decade_from == self.decade_to:
            self.decade = self.decade_from
        if self.year_from == self.year_to:
            self.year = self.year_from
        if self.month_from == self.month_to:
            self.month = self.month_from          
        if self.day_from == self.day_to:
            self.day = self.day_from                        
        
        if self.month == None and self.month_from == None:
            self.day_from = 1
            self.month_from = 1
            self.day_to = 31
            self.month_to = 12
            
        if self.day == None and self.month is not None and self.year is not None:
            self.day_from = 1
            self.day_to = monthrange(self.year, self.month)[1]
            
        if self.day_from == None and self.month_from is not None and self.year_from is not None:
            self.day_from = 1
            
        if self.day_to == None and self.month_to is not None and self.year_to is not None:
            self.day_to = 1
            self.day_to = monthrange(self.year_to, self.month_to)[1]
                     
    def convert_month(self, s):
        """Convert a string month representation to an int month representation."""

        s = s.lower()

        if (s == "january") or (s == "jan") or (s == "jan."):
            month = 1
        elif (s == "february") or (s == "feb") or (s == "feb."):
            month = 2
        elif (s == "march") or (s == "mar") or (s == "mar."):
            month = 3
        elif (s == "april") or (s == "apr") or (s == "apr."):
            month = 4
        elif (s == "may") or (s == "may") or (s == "may."):
            month = 5
        elif (s == "june") or (s == "jun") or (s == "jun."):
            month = 6
        elif (s == "july") or (s == "jul") or (s == "jul."):
            month = 7
        elif (s == "august") or (s == "aug") or (s == "aug."):
            month = 8
        elif (s == "september") or (s == "sep") or (s == "sep."):
            month = 9
        elif (s == "october") or (s == "oct") or (s == "oct."):
            month = 10
        elif (s == "november") or (s == "nov") or (s == "nov."):
            month = 11
        elif (s == "december") or (s == "dec") or (s == "dec."):
            month = 12
            
        try:
            return month
        except UnboundLocalError:
            return None

    def parse(self, s):
        """Parse a string and try to convert to numerical year, month and day."""

        decade = None
        year = None
        month = None
        day = None

        if s.count('/') == 2:
            chunks = s.split('/')
        elif s.count('-') == 2:  
            chunks = s.split('-')
        elif s.count('\\') == 2:  
            chunks = s.split('\\')
        elif s.count('.') == 2:  
            chunks = s.split('.')
        elif s.count(' ') == 2:
            chunks = s.split(' ')

        try:
            try:
                if int(chunks[0]) > 31: # if chunk is greater than 31 we can assume it's the year
                    if (int(chunks[1]) <= 12) and (int(chunks[2]) <= 31): # year/month/day
                        year = int(chunks[0])
                        month = int(chunks[1])
                        day = int(chunks[2])
                else: # otherwise we assume the year is last
                    if (int(chunks[1]) <= 12) and (int(chunks[0]) <= 31): # day/month/year
                        year = int(chunks[2])
                        month = int(chunks[1])
                        day = int(chunks[0])
                    elif (int(chunks[0]) <= 12) and (int(chunks[1]) <= 31): # month/day/year
                        year = int(chunks[2])
                        month = int(chunks[0])
                        day = int(chunks[1])
                        
                if month < 1 or month > 12:
                    month = None
                    
                if day < 1 or day > 31:
                    day = None
            except ValueError: # chunks didn't cast to int - not a compact date format

                if (len(chunks[0]) < 5) and ((chunks[0][-2:] == "st") or (chunks[0][-2:] == "nd") or (chunks[0][-2:] == "rd") or (chunks[0][-2:] == "th")):
                    day = int(chunks[0][:-2]) # if chunk is less than 5 characters, and ends in st, nd, rd, th, it's a (ordinal) day
                    month = self.convert_month(chunks[1])
                else:
                    try: # if not, try casting to int (cardinal day)
                        day = int(chunks[0])
                        month = self.convert_month(chunks[1])
                        
                    except ValueError: # if the cast fails, assume we have a month
                        if (len(chunks[1]) < 5) and ((chunks[1][-2:] == "st") or (chunks[1][-2:] == "nd") or (chunks[1][-2:] == "rd") or (chunks[1][-2:] == "th")):
                            day = int(chunks[1][:-2]) # if chunk is less than 5 characters, and ends in st, nd, rd, th, it's a (ordinal) day
                        else:
                            try: # if not, try casting to int (cardinal day)
                                day = int(chunks[1])
                            except ValueError:
                                day = None
                    
                        month = self.convert_month(chunks[0])
                            
                try: # if s can be cast to an int, assume it's a year
                    year = int(chunks[2])
                    
                except ValueError:
                    year = None
                    
        except UnboundLocalError: # no chunks, splitting failed, try with 1 seperator (for e.g. May 2010)
            try:
                if s.count('/') == 1:  
                    chunks = s.split('/')
                elif s.count('-') == 1:  
                    chunks = s.split('-')
                elif s.count('\\') == 1:  
                    chunks = s.split('\\')
                elif s.count('.') == 1:  
                    chunks = s.split('.')
                elif s.count(' ') == 1:
                    chunks = s.split(' ')

                month = self.convert_month(chunks[0])

                if month == None:
                    try: # if s can be cast to an int, assume it's a year
                        year = int(chunks[0])
                        
                        try:
                            month = int(chunks[1])
                            if month < 1 or month > 12:
                                month = None
                        except ValueError:
                            month = None                
                    except ValueError:
                        year = None                
                else:
                    try: # if s can be cast to an int, assume it's a year
                        year = int(chunks[1])
                    except ValueError:
                        year = None
            except UnboundLocalError: # still no chunks, splitting failed, try with 0 seperators (for e.g. May)
                chunks = s
                
                try: # if s can be cast to an int, assume it's a year
                    year = int(chunks)
                except ValueError: # if s can't be cast as an int, assume it's a month
                    month = self.convert_month(chunks)
                    
        if year:
            decade = int(''.join([str(year)[:-1], '0']))
        
        return (decade, year, month, day)
        
if __name__ == "__main__":
    import sys
    
    try:
        parser = VagueDate(str(sys.argv[1]))
        print "date", parser.decade, parser.year, parser.month, parser.day
        print "date (from)", parser.decade_from, parser.year_from, parser.month_from, parser.day_from
        print "date (to)", parser.decade_to, parser.year_to, parser.month_to, parser.day_to
    except IndexError:
        exit()

