# -*- coding: utf-8 -*-
"""
Created on Mon Apr 22 14:57:36 2019

@author: RDCRLDDH
"""

import pandas as pd
import xlwings as xl

class Bridge(object):
    
    def __init__(self):
        
        self.name = None
        self.station = None
        self.low_chord = None
        self.points = []
        self.n_val = None
        self.break_points = []
        self.bridge_type = None
        self.bridge_width = None
        self.embank_slope = None
        self.embank_elev = None
        self.ww_angle = None
        self.ww_width = None
        self.ent_radius = None
        self.pier_elev = []
        self.pier_width = []
        
        
    @staticmethod
    def test(line):
        if line[:2] == 'BR':
            return True
        return False

    def import_geo(self, line, wspro_file):
        self.name = line.split()[1]
        #print(self.name)
        self.station, self.low_chord = line.split()[-1].split(',')
        
        line = next(wspro_file)
        
        while line[:2] == 'GR':
            vals = []
            for i in line[11:].split():
                vals.append(i)
                  
            #vals = [tuple(i.split(',')) for i in vals]
            for i in vals:
                self.points.append((float(i.split(',')[0]), float(i.split(',')[1])))
            line = next(wspro_file)
            
        #Parse manning n values
        while line[:1] == 'N':
            self.n_val = line.split()[-1]
            line = next(wspro_file)
            
        if line[:2] == 'SA':
            line = next(wspro_file)
#            if line == 'SA':
#                line = next(wspro_file)
#            else:
#                try:
#                    sa_vals = [float(i) for i in line.split()[1].split(',')]
#                    for i in sa_vals:
#                        self.break_points.append(i)
#                    line = next(wspro_file)
#                except:
#                    break
            
        #bridge 
        if line[:2] == 'CD':
            vals = line.split()[-1]
            print(self.name)
            print(vals)
            if len(vals.split(',')) == 7:
                self.bridge_type, self.bridge_width, self.embank_slope, self.embank_elev, \
                self.ww_angle, self.ww_width, self.ent_radius= vals.split(',')
                line = next(wspro_file)
            elif len(vals.split(',')) == 4:
                self.bridge_type, self.bridge_width, self.embank_slope, self.embank_elev = vals.split(',')
                line = next(wspro_file)
                
        #pier/pile Dta        
        if line[:2] == 'PW':
            
            vals = line.split()[1:]
            for i in vals:
                self.pier_elev.append(i.split(',')[0])
                self.pier_width.append(i.split(',')[1])
            line = next(wspro_file)
            
        return line
            
            
class Approach_CrossSetion(object):
            
    def __init__(self):
        
        self.name = None
        self.distance = None
        self.points = []
        self.n = []
        self.break_points = []
        
        
    @staticmethod
    def test(line):
        if line[:2] == 'AS':
            return True
        return False
    
    
    def import_geo(self, line, wspro_file):
        
        self.name = line[5:9].strip()
        self.distance = line[10:].strip()
        self.q = None
        
        line = next(wspro_file)
        
        
        while line[:2] == 'GR':
            vals = []
            for i in line[11:].split():
                vals.append(i)
                  
            #vals = [tuple(i.split(',')) for i in vals]
            for i in vals:
                self.points.append((int(i.split(',')[0]), float(i.split(',')[1])))
            line = next(wspro_file)
        
        #Parse manning n values
        while line[:1] == 'N':
            n_vals = [float(i) for i in line.split()[1].split(',')]
            for i in n_vals:
                self.n.append(i)
            line = next(wspro_file)
            
        while line[:2] == 'SA':           
            try:
                sa_vals = [float(i) for i in line.split()[1].split(',')]
                for i in sa_vals:
                    self.break_points.append(i)
                line = next(wspro_file)
            except:
                break
                
            
        #Look for cross section flow paramater    
        if line[:1] == 'Q':
            self.q = float(line.split()[1])
            line = next(wspro_file)
            
#        print(self.n)    
#        print(self.break_points)
        #assert(len(self.n) == len(self.break_points))
        return line          
        

class CrossSection(object):
    
    def __init__(self):
        
        self.name = None
        self.distance = None
        self.points = []
        self.n = []
        self.break_points = []
        
        
    @staticmethod
    def test(line):
        if line[:2] == 'XS':
            return True
        return False
    
    
    def import_geo(self, line, wspro_file):
        
        self.name = line[5:9].strip()
        self.distance = line[10:].strip()
        self.q = None
        
        line = next(wspro_file)
        
        
        while line[:2] == 'GR':
            vals = []
            for i in line[11:].split():
                vals.append(i)
                  
            #vals = [tuple(i.split(',')) for i in vals]
            for i in vals:
                self.points.append((int(i.split(',')[0]), float(i.split(',')[1])))
            line = next(wspro_file)
        
        #Parse manning n values
        while line[:1] == 'N':
            n_vals = [float(i) for i in line.split()[1].split(',')]
            for i in n_vals:
                self.n.append(i)
            line = next(wspro_file)
            
        while line[:2] == 'SA':           
            try:
                sa_vals = [float(i) for i in line.split()[1].split(',')]
                for i in sa_vals:
                    self.break_points.append(i)
                line = next(wspro_file)
            except:
                break
                
            
        #Look for cross section flow paramater    
        if line[:1] == 'Q':
            self.q = float(line.split()[1])
            line = next(wspro_file)
            
#        print(self.n)    
#        print(self.break_points)
        #assert(len(self.n) == len(self.break_points))
        return line   
        

class Header(object):
    def __init__(self):
        self.reach_name = None
        self.Q = None
        self.WaterSurfaceElev = None
        self.slope = None
        
    @staticmethod
    def test(line):
        if line[:2] == 'T1':
            return True
        return False

    def import_geo(self, line, wspro_file):
        
        while line[:2] == 'T1':
            #print(line)
            self.reach_name = line[11:]
            
            line = next(wspro_file)
        while line[:1] == 'Q':
            self.Q = line.split()[1]
            line = next(wspro_file)
            
        if line[:2] == 'WS':
            self.WaterSurfaceElev = line.split()[1]
            line = next(wspro_file)
        elif line[:2] == 'SK':
            self.slope = line.split()
            line = next(wspro_file)
            #Job Parametrs
            #Not sure what these mean
        while line[:1] == 'J':
            line = next(wspro_file)
        
        
        return line
        

class ParseWSPRO(object):
    
    def __init__(self, wspro_filename):
        self.title = None
        self.geo_list = []
        if wspro_filename == '' or wspro_filename is None:
            raise AttributeError('Filename passed is blank.')

        with open(wspro_filename, 'rt') as wspro_file:
            for line in wspro_file:
                if Header.test(line):
                    head = Header()
                    head.import_geo(line,wspro_file)
                    self.title = head.reach_name.strip()
                    self.geo_list.append(head)
                elif CrossSection.test(line):
                    xs = CrossSection()
                    xs.import_geo(line, wspro_file)
                    self.geo_list.append(xs)
                elif Approach_CrossSetion.test(line):
                    a_xs = Approach_CrossSetion()
                    a_xs.import_geo(line, wspro_file)
                    self.geo_list.append(a_xs)
                elif Bridge.test(line):
                    br = Bridge()
                    br.import_geo(line,wspro_file)
                    self.geo_list.append(br)
                
                
def bridge_data(item):
    
    bridge_type = {1:'Vertical embankments AND vertical abutments, with or without wingwalls',
    2 :'Sloping embankments AND vertical abutments without wingwalls',
    3 :'Sloping embankments AND sloping spillthrough abutments',
    4 :'Sloping embankments AND vertical abutments with wingwalls'}
    return {'Name':item.name,
    'Low Chord':item.low_chord,
    'N Val':item.n_val,
    'Bridge Type':bridge_type[int(item.bridge_type)],
    'Bridge Width':item.bridge_width,
    'Embankment Slope':item.embank_slope,
    'Embankment Elevation':item.embank_elev,
    'Wing Wall Angle':item.ww_angle,
    'Wing Wall Width':item.ww_width,
    'Entrance Radius': item.ent_radius,
    'Pier Elevation':item.pier_elev,
    'Pier Width':item.pier_width}


           
    
                
def main():
#    east_filename = r"D:\NarraguagusRiver_CherryfieldME\Handoff\WSPRO\cherryfield.east"            
#    geo_east = ParseWSPRO(east_filename)
#    
#    west_filename = r"D:\NarraguagusRiver_CherryfieldME\Handoff\WSPRO\cherryfield.west"            
#    geo_west = ParseWSPRO(west_filename)
    
    main_filename = r"C:\workspace\NarraguagusRiver_CherryfieldME\Handoff\WSPRO\cherryfield"
    geo_main = ParseWSPRO(main_filename)
    
#    
#    east_branch_book = pd.ExcelWriter(r"C:\workspace\Narraguagus_River\GIS\analysis\Narraguagus_River_East_Branch.xlsx")
#    west_branch_book = pd.ExcelWriter(r"C:\workspace\Narraguagus_River\GIS\analysis\Narraguagus_River_West_Branch.xlsx")
    main_branch_book = pd.ExcelWriter(r"C:\workspace\Narraguagus_River\GIS\analysis\Narraguagus_River_Mainstem.xlsx")
    
    #bridge_list = [item for item in geo_main.geo_list if isinstance(item, Bridge)]
    
    for item in geo_main.geo_list:
        if isinstance(item, CrossSection):
            print(item.name)
            pd.DataFrame(item.points, columns=['Station','Elevation']).to_excel(main_branch_book, sheet_name=item.name,index=False)
        elif isinstance(item,Bridge):
            print(item.name)
            sht_name = item.name + '_Bridge'
            pd.DataFrame(item.points, columns=['Station','Elevation']).to_excel(main_branch_book, sheet_name=sht_name,index=False,
                        startrow=14)
            pd.DataFrame.from_dict(bridge_data(item), orient='index').to_excel(main_branch_book, sheet_name=sht_name)
        elif isinstance(item, Approach_CrossSetion):
            print(item.name)
            pd.DataFrame(item.points, columns=['Station','Elevation']).to_excel(main_branch_book, sheet_name=item.name,index=False)

#    for item in geo_west.geo_list:
#        if isinstance(item, CrossSection):
#            print(item.name)
#            pd.DataFrame(item.points, columns=['Station','Elevation']).to_excel(west_branch_book, sheet_name=item.name,index=False)
#                
#    for item in geo_east.geo_list:
#        if isinstance(item, CrossSection):
#            print(item.name)
#            pd.DataFrame(item.points, columns=['Station','Elevation']).to_excel(east_branch_book, sheet_name=item.name,index=False)
#
    main_branch_book.save()
#    west_branch_book.save()
#    east_branch_book.save()
    
if __name__ == '__main__':
    
    
    main()
                