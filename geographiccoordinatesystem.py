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

class Coordinate(object):

    def __init__(self, ini=None):
        """Initialize the object variables and lookup tables."""

        self.e = None
        self.n = None
        self.grid_reference = None
        self.far_north = False
        
        self.width = None
        self.accuracy = None
        self.os_100km = None
        self.os_10km = None
        self.os_5km = None
        self.os_2km = None
        self.os_1km = None
        self.os_100m = None
        self.os_10m = None
        self.os_1m = None
    
        self.__to_en_matrix__ = {'HL':'012', 'HM':'112', 'HN':'212', 'HO':'312', 'HP':'412', 'JL':'512', 'JM':'612',
                                 'HQ':'011', 'HR':'111', 'HS':'211', 'HT':'311', 'HU':'411', 'JQ':'511', 'JR':'611',
                                 'HV':'000', 'HW':'110', 'HX':'210', 'HY':'310', 'HZ':'410', 'JV':'510', 'JW':'610',
                                 'NA':'09',  'NB':'19',  'NC':'29',  'ND':'39',  'NE':'49',  'OA':'59',  'OB':'69', 
                                 'NF':'08',  'NG':'18',  'NH':'28',  'NJ':'38',  'NK':'48',  'OF':'58',  'OG':'68',
                                 'NL':'07',  'NM':'17',  'NN':'27',  'NO':'37',  'NP':'47',  'OL':'57',  'OM':'67',
                                 'NQ':'06',  'NR':'16',  'NS':'26',  'NT':'36',  'NU':'46',  'OQ':'56',  'OR':'66',
                                 'NV':'05',  'NW':'15',  'NX':'25',  'NY':'35',  'NZ':'45',  'OV':'55',  'OW':'65',    
                                 'SA':'04',  'SB':'14',  'SC':'24',  'SD':'34',  'SE':'44',  'TA':'54',  'TB':'64',
                                 'SF':'03',  'SG':'13',  'SH':'23',  'SJ':'33',  'SK':'43',  'TF':'53',  'TG':'63',
                                 'SL':'02',  'SM':'12',  'SN':'22',  'SO':'32',  'SP':'42',  'TL':'52',  'TM':'62',
                                 'SQ':'01',  'SR':'11',  'SS':'21',  'ST':'31',  'SU':'41',  'TQ':'51',  'TR':'61', 
                                 'SV':'00',  'SW':'10',  'SX':'20',  'SY':'30',  'SZ':'40',  'TV':'50',  'TW':'60',}
                               
        self.__to_os_matrix__ = {'012':'HL', '112':'HM', '212':'HN', '312':'HO', '412':'HP', '512':'JL', '612':'JM',
                                 '011':'HQ', '111':'HR', '211':'HS', '311':'HT', '411':'HU', '511':'JQ', '611':'JR',
                                 '010':'HV', '110':'HW', '210':'HX', '310':'HY', '410':'HZ', '510':'JV', '610':'JW',                                                                                 
                                  '09':'NA',  '19':'NB',  '29':'NC',  '39':'ND',  '49':'NE',  '59':'OA',  '69':'OB',
                                  '08':'NF',  '18':'NG',  '28':'NH',  '38':'NJ',  '48':'NK',  '58':'OF',  '68':'OG',
                                  '07':'NL',  '17':'NM',  '27':'NN',  '37':'NO',  '47':'NP',  '57':'OL',  '67':'OM',
                                  '06':'NQ',  '16':'NR',  '26':'NS',  '36':'NT',  '46':'NU',  '56':'OQ',  '66':'OR', 
                                  '05':'NV',  '15':'NW',  '25':'NX',  '35':'NY',  '45':'NZ',  '55':'OV',  '65':'OW', 
                                  '04':'SA',  '14':'SB',  '24':'SC',  '34':'SD',  '44':'SE',  '54':'TA',  '64':'TB',
                                  '03':'SF',  '13':'SG',  '23':'SH',  '33':'SJ',  '43':'SK',  '53':'TF',  '63':'TG',
                                  '02':'SL',  '12':'SM',  '22':'SN',  '32':'SO',  '42':'SP',  '52':'TL',  '62':'TM', 
                                  '01':'SQ',  '11':'SR',  '21':'SS',  '31':'ST',  '41':'SU',  '51':'TQ',  '61':'TR',
                                  '00':'SV',  '10':'SW',  '20':'SX',  '30':'SY',  '40':'SZ',  '50':'TV',  '60':'TW',}

        self.__to_tetrad_matrix__ = {'08': 'E', '09': 'E', '18': 'E', '19': 'E', '28': 'J', '29': 'J', '38': 'J', '39': 'J', '48': 'P', '49': 'P', '58': 'P', '59': 'P', '68': 'U', '69': 'U', '78': 'U', '79': 'U', '88': 'Z', '89': 'Z', '98': 'Z', '99': 'Z',
                                     '06': 'D', '07': 'D', '16': 'D', '17': 'D', '26': 'I', '27': 'I', '36': 'I', '37': 'I', '46': 'N', '47': 'N', '56': 'N', '57': 'N', '66': 'T', '67': 'T', '76': 'T', '77': 'T', '86': 'Y', '87': 'Y', '96': 'Y', '97': 'Y',
                                     '04': 'C', '05': 'C', '14': 'C', '15': 'C', '24': 'H', '25': 'H', '34': 'H', '35': 'H', '44': 'M', '45': 'M', '54': 'M', '55': 'M', '64': 'S', '65': 'S', '74': 'S', '75': 'S', '84': 'X', '85': 'X', '94': 'X', '95': 'X',
                                     '02': 'B', '03': 'B', '12': 'B', '13': 'B', '22': 'G', '23': 'G', '32': 'G', '33': 'G', '42': 'L', '43': 'L', '52': 'L', '53': 'L', '62': 'R', '63': 'R', '72': 'R', '73': 'R', '82': 'W', '83': 'W', '92': 'W', '93': 'W',
                                     '00': 'A', '01': 'A', '10': 'A', '11': 'A', '20': 'F', '21': 'F', '30': 'F', '31': 'F', '40': 'K', '41': 'K', '50': 'K', '51': 'K', '60': 'Q', '61': 'Q', '70': 'Q', '71': 'Q', '80': 'V', '81': 'V', '90': 'V', '91': 'V',}

        self.__from_tetrad_matrix__ = {'E':'08', 'J':'28', 'P':'48', 'U':'68', 'Z':'88',
                                       'D':'06', 'I':'26', 'N':'46', 'T':'66', 'Y':'86',
                                       'C':'04', 'H':'24', 'M':'44', 'S':'64', 'X':'84',
                                       'B':'02', 'G':'22', 'L':'42', 'R':'62', 'W':'82',
                                       'A':'00', 'F':'20', 'K':'40', 'Q':'60', 'V':'80',}
                                       
        self.__to_5km_matrix__ = {'45':'NW', '46':'NW', '47':'NW', '48':'NW', '49':'NW', 
                                  '35':'NW', '36':'NW', '37':'NW', '38':'NW', '39':'NW', 
                                  '25':'NW', '26':'NW', '27':'NW', '28':'NW', '29':'NW', 
                                  '15':'NW', '16':'NW', '17':'NW', '18':'NW', '19':'NW', 
                                  '05':'NW', '06':'NW', '07':'NW', '08':'NW', '09':'NW',
                                  
                                  '95':'NE', '96':'NE', '97':'NE', '98':'NE', '99':'NE',
                                  '85':'NE', '86':'NE', '87':'NE', '88':'NE', '89':'NE',
                                  '75':'NE', '76':'NE', '77':'NE', '78':'NE', '79':'NE',
                                  '65':'NE', '66':'NE', '67':'NE', '68':'NE', '69':'NE',
                                  '55':'NE', '56':'NE', '57':'NE', '58':'NE', '59':'NE',
                                  
                                  '40':'SW', '41':'SW', '42':'SW', '43':'SW', '44':'SW',
                                  '30':'SW', '31':'SW', '32':'SW', '33':'SW', '34':'SW',
                                  '20':'SW', '21':'SW', '22':'SW', '23':'SW', '24':'SW',
                                  '10':'SW', '11':'SW', '12':'SW', '13':'SW', '14':'SW',
                                  '00':'SW', '01':'SW', '02':'SW', '03':'SW', '04':'SW',
                                  
                                  '90':'SE', '91':'SE', '92':'SE', '93':'SE', '94':'SE', 
                                  '80':'SE', '81':'SE', '82':'SE', '83':'SE', '84':'SE', 
                                  '70':'SE', '71':'SE', '72':'SE', '73':'SE', '74':'SE', 
                                  '60':'SE', '61':'SE', '62':'SE', '63':'SE', '64':'SE', 
                                  '50':'SE', '51':'SE', '52':'SE', '53':'SE', '54':'SE',}
                                    
        self.__from_5km_matrix__ = {'NW':'05', 'NE':'55',
                                    'SW':'00', 'SE':'50',}
        
        if type(ini) is tuple and len(ini) == 2:   
            self.set_en(ini[0], ini[1])
            if self.e is not None and self.n is not None:
                self.__to_os__()

        if type(ini) is str:  
            self.set_grid_reference(ini)

    def set_en(self, e, n):
        """Set the easting/northings and convert to OS grid reference."""

        if int(e) < 0 or int(e) > 700000:
            self.e = None
        else:
            self.e = str(int(e)).rjust(6, '0')
            
        if int(n) < 0 or int(n) > 1300000:
            self.n = None
        else:
            if int(n) >= 1000000: # greater than _or_ equal to
                self.far_north = True
                self.n = str(int(n)).rjust(7, '0')
            else:
                self.far_north = False
                self.n = str(int(n)).rjust(6, '0')
            
        self.accuracy = None
        
        if self.e is not None and self.n is not None:
            self.__to_os__()
            self.__convert_os__()

    def set_grid_reference(self, grid_reference):
        """Set the grid reference and convert to easting/northings."""
        self.grid_reference = grid_reference
        
        if self.grid_reference[0:2].upper() in ('HL', 'HM', 'HN', 'HO', 'HP',
                                                'JL', 'JM', 'HQ', 'HR', 'HS',
                                                'HT', 'HU', 'JQ', 'JR', 'HV',
                                                'HW', 'HX', 'HY', 'HZ', 'JV',
                                                'JW',):
            self.far_north = True
        else:
            self.far_north = False
        
        try:
            self.__convert_os__()
            self.__to_en__()
        except:
            self.e = None
            self.n = None
            self.grid_reference = None
            self.far_north = False
            
            self.accuracy = None
            self.os_100km = None
            self.os_10km = None
            self.os_5km = None
            self.os_2km = None
            self.os_1km = None
            self.os_100m = None
            self.os_10m = None
            self.os_1m = None

    def __convert_os__(self):
        """Convert the grid reference to other accuracies"""
        if self.grid_reference is not None:

            if self.grid_reference[-2:] in self.__from_5km_matrix__.keys():
                self.accuracy = 1
                self.os_5km = self.grid_reference
                self.width = 5000
            elif self.grid_reference[-1:] in self.__from_tetrad_matrix__.keys():
                self.accuracy = (len(self.grid_reference)-2)/2
                self.width = 2000
                self.os_2km = self.grid_reference
                #################need to get 5km from tetrad. self.os_5km = self.grid_reference
            else:
                self.accuracy = (len(self.grid_reference)-2)/2

            self.os_100km = self.grid_reference [:2]
            x_nums = self.grid_reference[2:self.accuracy+2]
            y_nums = self.grid_reference[self.accuracy+2:]

            if self.accuracy >= 0:
                self.os_100km = ''.join([self.os_100km, x_nums[:0], y_nums[:0]])
                self.width = 100000
                
            if self.accuracy >= 1:
                self.os_10km = ''.join([self.os_100km, x_nums[:1], y_nums[:1]])
                self.width = 10000
                
            if self.accuracy >= 2:
                self.os_5km = ''.join([self.os_100km, x_nums[:1], y_nums[:1], self.__to_5km_matrix__[''.join([x_nums[1], y_nums[1]])]])
                self.os_2km = ''.join([self.os_100km, x_nums[:1], y_nums[:1], self.__to_tetrad_matrix__[''.join([x_nums[1], y_nums[1]])]])
                self.os_1km = ''.join([self.os_100km, x_nums[:2], y_nums[:2]])
                
            if self.accuracy >= 3:
                self.os_100m = ''.join([self.os_100km, x_nums[:3], y_nums[:3]])
                self.width = 100
                
            if self.accuracy >= 4:
                self.os_10m = ''.join([self.os_100km, x_nums[:4], y_nums[:4]])
                self.width = 10
                
            if self.accuracy >= 5:
                self.os_1m = ''.join([self.os_100km, x_nums[:5], y_nums[:5]])
                self.width = 1
                              
    def __to_os__(self):
        """Convert easting/northing to OS grid reference."""
        try:
            if self.far_north:
                self.grid_reference = ''.join([self.__to_os_matrix__[''.join([str(self.e)[0], str(self.n)[0:2]])],
                                               str(self.e)[1:6],
                                               str(self.n)[2:(6+1)]],)
            else:                                   
                self.grid_reference = ''.join([self.__to_os_matrix__[''.join([str(self.e)[0], str(self.n)[0]])],
                                               str(self.e)[1:6],
                                               str(self.n)[1:6]],)      
        except KeyError:
            self.grid_reference = None


    def __to_en__(self):
        """Convert OS grid reference to easting/northing."""
        if self.grid_reference is not None:

            try:
                #implment dinty/5km for far north!!!!!!!
                if self.far_north:
                    self.e = int((self.__to_en_matrix__[self.grid_reference[:2].upper()][:1] + \
                                  self.grid_reference[2:][:len(self.grid_reference[2:]) /2]).ljust(6, '0'))
                    self.n = int((self.__to_en_matrix__[self.grid_reference[:2].upper()][1:] + \
                                  self.grid_reference[2:][len(self.grid_reference[2:]) /2:]).ljust(7, '0')) 
                else:
                    if self.grid_reference[-2:] in ('NW', 'NE', 'SW', 'SE'):
                        self.e = int((self.__to_en_matrix__[self.grid_reference[:2].upper()][:1] + \
                                      self.grid_reference[2] + \
                                      self.__from_5km_matrix__[self.grid_reference[-2:]][0]).ljust(6, '0'))
                        
                        self.n = int((self.__to_en_matrix__[self.grid_reference[:2].upper()][1:] + \
                                      self.grid_reference[3] + \
                                      self.__from_5km_matrix__[self.grid_reference[-2:]][1]).ljust(6, '0')) 
                    elif len(self.grid_reference) == 5:
                        self.e = int((self.__to_en_matrix__[self.grid_reference[:2].upper()][:1] + \
                                      self.grid_reference[2] + \
                                      self.__from_tetrad_matrix__[self.grid_reference[4].upper()][0]).ljust(6, '0'))
                        self.n = int((self.__to_en_matrix__[self.grid_reference[:2].upper()][1:] + \
                                      self.grid_reference[3] + \
                                      self.__from_tetrad_matrix__[self.grid_reference[4].upper()][1]).ljust(6, '0'))
                                                      
                    else:
                        self.e = int((self.__to_en_matrix__[self.grid_reference[:2].upper()][:1] + \
                                      self.grid_reference[2:][:len(self.grid_reference[2:]) /2]).ljust(6, '0'))
                        self.n = int((self.__to_en_matrix__[self.grid_reference[:2].upper()][1:] + \
                                      self.grid_reference[2:][len(self.grid_reference[2:]) /2:]).ljust(6, '0'))        
                 
            except KeyError:
                self.e = None
                self.n = None      
            except ValueError:
                self.e = None
                self.n = None
                
        if self.e < 0 or self.e > 700000:
            self.e = None
            
        if self.n < 0 or self.n > 1300000:
            self.n = None
        
if __name__ == "__main__":
    import sys
    
    try:
        reference = Coordinate((sys.argv[1], sys.argv[2]))
        
        if reference.grid_reference is None:
            print 'Invalid easting/northing'
        else:
            print reference.grid_reference
    except IndexError:
        try:
            reference = Coordinate(sys.argv[1])
            if reference.e is None or reference.n is None: 
                print 'Invalid grid reference'
            else:
                print reference.e, reference.n   
        except IndexError:
            exit()
