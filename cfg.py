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

import os

#vice county list - number & filename
vc_list = [[1, 'West Cornwall'],
         [2, 'East Cornwall'],
         [3, 'South Devon'],
         [4, 'North Devon'],
         [5, 'South Somerset'],
         [6, 'North Somerset'],
         [7, 'North Wiltshire'],
         [8, 'South Wiltshire'],
         [9, 'Dorset'],
         [10, 'Isle of Wight'],
         [11, 'South Hampshire'],
         [12, 'North Hampshire'],
         [13, 'West Sussex'],
         [14, 'East Sussex'],
         [15, 'East Kent'],
         [16, 'West Kent'],
         [17, 'Surrey'],
         [18, 'South Essex'],
         [19, 'North Essex'],
         [20, 'Hertfordshire'],
         [21, 'Middlesex'],
         [22, 'Berkshire'],
         [23, 'Oxfordshire'],
         [24, 'Buckinghamshire'],
         [25, 'East Suffolk'],
         [26, 'West Suffolk'],
         [27, 'East Norfolk'],
         [28, 'West Norfolk'],
         [29, 'Cambridgeshire'],
         [30, 'Bedfordshire'],
         [31, 'Huntingdonshire'],
         [32, 'Northamptonshire'],
         [33, 'East Gloucestershire'],
         [34, 'West Gloucestershire'],
         [35, 'Monmouthshire'],
         [36, 'Herefordshire'],
         [37, 'Worcestershire'],
         [38, 'Warwickshire'],
         [39, 'Stafford'],
         [40, 'Shropshire'],
         [41, 'Glamorgan'],
         [42, 'Brecon'],
         [43, 'Radnorshire'],
         [44, 'Carmarthenshire'],
         [45, 'Pembrokeshire'],
         [46, 'Cardiganshire'],
         [47, 'Montgomeryshire'],
         [48, 'Merionethshire'],
         [49, 'Caernarvonshire'],
         [50, 'Denbighshire'],
         [51, 'Flintshire'],
         [52, 'Anglesey'],
         [53, 'South Lincolnshire'],
         [54, 'North Lincolnshire'],
         [55, 'Leicestershire'],
         [56, 'Nottinghamshire'],
         [57, 'Derbyshire'],
         [58, 'Cheshire'],
         [59, 'South Lancashire'],
         [60, 'West Lancashire'],
         [61, 'South East Yorkshire'],
         [62, 'North East Yorkshire'],
         [63, 'South West Yorkshire'],
         [64, 'Mid West Yorkshire'],
         [65, 'North West Yorkshire'],
         [66, 'Durham'],
         [67, 'South Northumberland'],
         [68, 'North Northumberland'],
         [69, 'Westmorland'],
         [70, 'Cumberland'],
         [71, 'Isle of Man'],
         [72, 'Dumfrieshire'],
         [73, 'Kirkcudbrightshire'],
         [74, 'Wigtownshire'],
         [75, 'Ayrshire'],
         [76, 'Renfrewshire'],
         [77, 'Lanarkshire'],
         [78, 'Peebles'],
         [79, 'Selkirk'],
         [80, 'Roxburgh'],
         [81, 'Berwickshire'],
         [82, 'East Lothian'],
         [83, 'Midlothian'],
         [84, 'West Lothian'],
         [85, 'Fifeshire'],
         [86, 'Stirling'],
         [87, 'West Perth'],
         [88, 'Mid Perth'],
         [89, 'East Perth'],
         [90, 'Angus'],
         [91, 'Kincardine'],
         [92, 'South Aberdeen'],
         [93, 'North Aberdeen'],
         [94, 'Banff'],
         [95, 'Moray'],
         [96, 'East Inverness'],
         [97, 'West Inverness'],
         [98, 'Argyllshire'],
         [99, 'Dunbarton'],
         [100, 'Clyde Islands'],
         [101, 'Kintyre'],
         [102, 'South Ebudes'],
         [103, 'Mid Ebudes'],
         [104, 'North Ebudes'],
         [105, 'West Ross'],
         [106, 'East Ross'],
         [107, 'East Sutherland'],
         [108, 'West Sutherland'],
         [109, 'Caithness'],
         [110, 'Outer Hebrides'],
         [111, 'Orkney'],
         [112, 'Shetland']]
         

#min, max, count
gradation_ranges = [[1,1, 0],
                    [2,2, 0],
                    [3,3, 0],
                    [4,5, 0],
                    [6,10, 0],
                    [11,15, 0],
                    [16,20, 0],
                    [21,35, 0],
                    [36,50, 0],
                    [51,75, 0],
                    [76,100, 0],
                    [101,250, 0],
                    [251,500, 0],
                    [501,1000, 0]]
                    

#grid resolution list
grid_resolution = ['100km', '10km', '5km', '2km', '1km',]

#paper sizes list
paper_size = ['A4',]

#paper orientation list
paper_orientation = ['Portrait', 'Landscape',]

#phenology chart types
phenology_types = ['Months', 'Decades']

#walk the markers directory searching for GIS markers
markers = []

for style in os.listdir('markers/'):
    markers.append(style)

#walk the backgrounds directory searching for GIS markers
backgrounds = []

for background in os.listdir('backgrounds/'):
    backgrounds.append(background)                     
