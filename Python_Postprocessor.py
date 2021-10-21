##########################################################################################################################################################################################
########################################     Pipeline Stress POSTPROCESSOR    ###############################################################################
########################################     Subject:    Abaqus FEA Postprocessing   ###############################################################################
########################################     Author :    Engr.Jesurobo Collins       #####################################################################################    #################################################################################################
########################################     Project:    Personal project            ##############################################################################################
########################################     Tools used: Python,Abaqus,Excel         ##############################################################################################
########################################     Email:      collins4engr@yahoo.com      ##############################################################################################
#########################################################################################################################################################################################
import os
from abaqus import*
from abaqusConstants import*
import xlsxwriter
import glob
import numpy as np


# CHANGE TO CURRENT WORKING DIRECTORY
os.chdir('C:/temp/Pipeline Parametric studies/Stresses')

###CREATE EXCEL WORKBOOK, SHEETS AND ITS PROPERTIES####
execFile = 'Results.xlsx'
workbook = xlsxwriter.Workbook(execFile)
workbook.set_properties({
    'title':    'This is Abaqus postprocessing',
    'subject':  'Pipe Stress Postprocessing',   
    'author':   'Collins Jesurobo',
    'company':  'Personal Project',
    'comments': 'Created with Python and XlsxWriter'})

SHEET1 = workbook.add_worksheet()
SHEET1.center_horizontally()
SHEET1.fit_to_pages(1, 1)
SHEET1.set_column(0,6,19)
SHEET1.set_column(6,7,21)
SHEET1.merge_range('A1:G1', 'MAXIMUM AND MINIMUM AXIAL STRESS WITH CORRESPONDING WORST LOADCASE,WORST LOAD STEP AND NODE WHERE IT OCCURS')
SHEET1.merge_range('A7:H7', 'WORST STRESS FOR EACH LOADCASE (FROM EACH ODB) VERSUS WALL THICKNESS STUDIED')

SHEET2 = workbook.add_worksheet()
SHEET2.center_horizontally()
SHEET2.fit_to_pages(1, 1)
SHEET2.set_column(0,1,20)
SHEET2.set_column(2,9,12)
SHEET2.merge_range('A1:J1', 'STRESS RESULTS FOR ALL PIPELINE NODES FOR EACH LOADCASE AND CORRESPONDING LOAD STEP')


# DEFINES THE WORKSHEET FORMATTING (font name, size, cell colour etc.)
format_title = workbook.add_format()
format_title.set_bold('bold')
format_title.set_align('center')
format_title.set_align('vcenter')
format_title.set_bg_color('#F2F2F2')
format_title.set_font_size(10)
format_title.set_font_name('Arial')
format_table_headers = workbook.add_format()
format_table_headers.set_align('center')
format_table_headers.set_align('vcenter')
format_table_headers.set_text_wrap('text_wrap')
format_table_headers.set_bg_color('#F2F2F2')
format_table_headers.set_border()
format_table_headers.set_font_size(10)
format_table_headers.set_font_name('Arial')

###WRITING THE TITLES TO SHEET1,SHEET2###
SHEET1.write_row('B2',['S11(MPa)','S22(MPa)','S33(MPa)','S12(MPa)','Smises(MPa)','WorstLoadcase','LoadStep','element'],format_title)
SHEET1.write('A3', 'Max value',format_title)
SHEET1.write('A4', 'Min value',format_title)
SHEET1.write('A5', 'Absolute Max value',format_title)
SHEET1.write_row('A8', ['Loadcase','Thickness(mm)','S11_max(MPa)','S11_min(MPa)','S11_abs(MPa)',
                         'Smises_max(MPa)','Smises_min(MPa)','Smises_abs(MPa)'],format_title)
SHEET2.write_row('A2',['Loadcase','LoadStep','element','S11(MPa)','S22(Pa)','S33(MPa)','S12(MPa)',
                       'Smises(MPa)','Sallow(MPa)','Acceptance criteria' ],format_title)

# ALLOWABLE YIELD STRESS
Df=0.96         # design factor
SMYS=450        # specified minimum yield strength
Sallow = Df * SMYS           

###LOOP THROUGH THE ODBs, LOOP THROUGH EACH STEPS AND EXTRACT STRESS RESULTS FOR ALL PIPELINE NODES###
row=1
col=0
for i in glob.glob('*.odb'):     # loop  to access all odbs in the folder
        odb = session.openOdb(i) # open each odb
        step = odb.steps.keys()  # probe the content of the steps object in odb, steps object is a dictionary, so extract the step names with keys()
        instances = odb.rootAssembly.instances.keys() # probe instances
        elementset = odb.rootAssembly.instances[instances[0]].elementSets.keys()      # probe nodeset
        section = odb.rootAssembly.instances[instances[0]].elementSets[elementset[0]] # extract section for pipeline nodeset
        ###DEFINE RESULT OUTPUT####
        for k in range(len(step)):
                S = odb.steps[step[k]].frames[-1].fieldOutputs['S'].getSubset(region=section).values   # results for all displacements U 
                for Stress in S:
                        S11 = Stress.data[0]*10**-6            # extract S11 (axial stress) 
                        S22 = Stress.data[1]*10**-6           # extract S22 (hoop stress)
                        S33 = Stress.data[2]*10**-6            # extract S33 (radial stress) 
                        S12 = Stress.data[3]*10**-6            # extract S12 (Shear stress)
                        Smises = Stress.mises*10**-6
                        n1 =  Stress.elementLabel       # extract node numbers 
                        ### WRITE OUT MAIN RESULT OUTPUT####
                        SHEET2.write(row+1,col,i.split('.')[0],format_table_headers)  # loadcases
                        SHEET2.write(row+1,col+1,step[k],format_table_headers)        # steps in odb
                        SHEET2.write(row+1,col+2,n1,format_table_headers)             # all nodes in the pipeline
                        SHEET2.write(row+1,col+3,S11,format_table_headers)            # write axial stresses to excel, in MPa        
                        SHEET2.write(row+1,col+4,S22,format_table_headers)            # write hoop stresses to excel, in MPa
                        SHEET2.write(row+1,col+5,S33,format_table_headers)            # write radial stresses to excel,in MPa
                        SHEET2.write(row+1,col+6,S12,format_table_headers)            # write shear stresses to excel,in MPa
                        SHEET2.write(row+1,col+7,Smises,format_table_headers)
                        SHEET2.write(row+1,col+8,Sallow,format_table_headers)
                        # COMPARE THE VMISES STRESSES TO ALLOWABLE YIELD STRESS
                        if Smises > Sallow and abs(S11) > Sallow:
                                SHEET2.write(row+1,col+9,'Fail',format_table_headers)
                        else:
                                SHEET2.write(row+1,col+9,'Pass',format_table_headers)
                        row+=1

### GET THE MAXIMUM AND MINIMUM, AND ABSOLUTE MAXIMUM VALUES OF STRESSES AND AXIAL STRESS AND WRITE THEM INTO SUMMARY SHEET(SHEET1) 
def output2():
        SHEET1.write('B3', '=max(Sheet2!D3:D200000)',format_table_headers)            # Maximum axial stress
        SHEET1.write('C3', '=max(Sheet2!E3:E200000)',format_table_headers)            # Maximum hoop stress
        SHEET1.write('D3', '=max(Sheet2!F3:F200000)',format_table_headers)            # Maximum radial stress
        SHEET1.write('E3', '=max(Sheet2!G3:G200000)',format_table_headers)            # Maximum shear stress
        SHEET1.write('F3', '=max(Sheet2!H3:H200000)',format_table_headers)            # Maimum Vonmises stress
        
        
        SHEET1.write('B4', '=min(Sheet2!D3:D200000)',format_table_headers)            # Minimum axial stress
        SHEET1.write('C4', '=min(Sheet2!E3:E200000)',format_table_headers)            # Minimum hoop stress
        SHEET1.write('D4', '=min(Sheet2!F3:F200000)',format_table_headers)            # Minimum radial stress
        SHEET1.write('E4', '=min(Sheet2!G4:G200000)',format_table_headers)            # Minimum shear stress
        SHEET1.write('F4', '=min(Sheet2!H3:H200000)',format_table_headers)            # Maximum Vonmises stress
        
        SHEET1.write('B5','=IF(ABS(B3)>ABS(B4),ABS(B3),ABS(B4))',format_table_headers) # Absolute maximum axial stress
        SHEET1.write('C5','=IF(ABS(C3)>ABS(C4),ABS(C3),ABS(C4))',format_table_headers) # Absolute maximum hoop stress
        SHEET1.write('D5','=IF(ABS(D3)>ABS(D4),ABS(D3),ABS(D4))',format_table_headers) # Absolute maximum radial stress
        SHEET1.write('E5','=IF(ABS(E3)>ABS(E4),ABS(E3),ABS(E4))',format_table_headers) # Absolute maximum shear stress
        SHEET1.write('F5','=IF(ABS(F3)>ABS(F4),ABS(F3),ABS(F4))',format_table_headers) # Absolute maximum Vonmises stress
        

        ### WORST LOADCASE AND LOADSTEP CORRESPONDING TO MAXIMUM AND MINIMUM STRESSES VALUES
        SHEET1.write('G3','=INDEX(Sheet2!A3:A200000,MATCH(MAX(Sheet2!D3:D200000),Sheet2!D3:D200000,0))',format_table_headers)# worst loadcase for maximum stress
        SHEET1.write('H3','=INDEX(Sheet2!B3:B200000,MATCH(MAX(Sheet2!D3:D200000),Sheet2!D3:D200000,0))',format_table_headers)# worst loadstep for maximum stress
        SHEET1.write('I3','=INDEX(Sheet2!C3:C200000,MATCH(MAX(Sheet2!D3:D200000),Sheet2!D3:D200000,0))',format_table_headers)# worst element number for maximum stress
        SHEET1.write('G4','=INDEX(Sheet2!A3:A200000,MATCH(MIN(Sheet2!D3:D200000),Sheet2!D3:D200000,0))',format_table_headers)# worst loadcase for minimum stress
        SHEET1.write('H4','=INDEX(Sheet2!B3:B200000,MATCH(MIN(Sheet2!D3:D200000),Sheet2!D3:D200000,0))',format_table_headers)# worst loadstep for minimum stress
        SHEET1.write('I4','=INDEX(Sheet2!C3:C200000,MATCH(MIN(Sheet2!D3:D200000),Sheet2!D3:D200000,0))',format_table_headers)# worst element number for minimum stress

output2()
### PROBE THE ODB AND GET THE NAMES OF LOADCASES AND WRITE THEM INTO SHEET1
def output3():
        row=0
        col=0
        for LC in glob.glob('*.odb'):
                SHEET1.write(row+8,col,LC.split('.')[0],format_table_headers)
                row+=1

# WRITE THE COLUMN FOR WALL THICKNESSES THAT WAS USED IN THE PARAMETRIC STUDIES
Thick_data = [15.9,19.1,22.3,25.1,27.1,30.2]         # varied thickness in mm
SHEET1.write_column('B9',Thick_data,format_table_headers)

### WORST LONGITUDINAL STRESS VALUES
SHEET1.write('C9', '{=MAX(IF(Sheet2!A3:A200000=Sheet1!A9,  Sheet2!D3:D200000))}',format_table_headers)
SHEET1.write('C10','{=MAX(IF(Sheet2!A3:A200000=Sheet1!A10, Sheet2!D3:D200000))}',format_table_headers)
SHEET1.write('C11','{=MAX(IF(Sheet2!A3:A200000=Sheet1!A11, Sheet2!D3:D200000))}',format_table_headers)
SHEET1.write('C12','{=MAX(IF(Sheet2!A3:A200000=Sheet1!A12, Sheet2!D3:D200000))}',format_table_headers)
SHEET1.write('C13','{=MAX(IF(Sheet2!A3:A200000=Sheet1!A13, Sheet2!D3:D200000))}',format_table_headers)
SHEET1.write('C14','{=MAX(IF(Sheet2!A3:A200000=Sheet1!A14, Sheet2!D3:D200000))}',format_table_headers)

SHEET1.write('D9', '{=MIN(IF(Sheet2!A3:A200000=Sheet1!A9,  Sheet2!D3:D200000))}',format_table_headers)
SHEET1.write('D10','{=MIN(IF(Sheet2!A3:A200000=Sheet1!A10, Sheet2!D3:D200000))}',format_table_headers)
SHEET1.write('D11','{=MIN(IF(Sheet2!A3:A200000=Sheet1!A11, Sheet2!D3:D200000))}',format_table_headers)
SHEET1.write('D12','{=MIN(IF(Sheet2!A3:A200000=Sheet1!A12, Sheet2!D3:D200000))}',format_table_headers)
SHEET1.write('D13','{=MIN(IF(Sheet2!A3:A200000=Sheet1!A13, Sheet2!D3:D200000))}',format_table_headers)
SHEET1.write('D14','{=MIN(IF(Sheet2!A3:A200000=Sheet1!A14, Sheet2!D3:D200000))}',format_table_headers)

SHEET1.write('E9','=IF(ABS(C9)>ABS(D9),ABS(C9),ABS(D9))',   format_table_headers) 
SHEET1.write('E10','=IF(ABS(C10)>ABS(D9),ABS(C10),ABS(D10))',format_table_headers) 
SHEET1.write('E11','=IF(ABS(C11)>ABS(D9),ABS(C11),ABS(D11))',format_table_headers) 
SHEET1.write('E12','=IF(ABS(C12)>ABS(D9),ABS(C12),ABS(D12))',format_table_headers) 
SHEET1.write('E13','=IF(ABS(C13)>ABS(D9),ABS(C13),ABS(D13))',format_table_headers) 
SHEET1.write('E14','=IF(ABS(C14)>ABS(D9),ABS(C14),ABS(D14))',format_table_headers)

### WORST VONMISES STRESS VALUES
SHEET1.write('F9', '{=MAX(IF(Sheet2!A3:A200000=Sheet1!A9,  Sheet2!H3:H200000))}',format_table_headers)
SHEET1.write('F10','{=MAX(IF(Sheet2!A3:A200000=Sheet1!A10, Sheet2!H3:H200000))}',format_table_headers)
SHEET1.write('F11','{=MAX(IF(Sheet2!A3:A200000=Sheet1!A11, Sheet2!H3:H200000))}',format_table_headers)
SHEET1.write('F12','{=MAX(IF(Sheet2!A3:A200000=Sheet1!A12, Sheet2!H3:H200000))}',format_table_headers)
SHEET1.write('F13','{=MAX(IF(Sheet2!A3:A200000=Sheet1!A13, Sheet2!H3:H200000))}',format_table_headers)
SHEET1.write('F14','{=MAX(IF(Sheet2!A3:A200000=Sheet1!A14, Sheet2!H3:H200000))}',format_table_headers)

SHEET1.write('G9', '{=MIN(IF(Sheet2!A3:A200000=Sheet1!A9,  Sheet2!H3:H200000))}',format_table_headers)
SHEET1.write('G10','{=MIN(IF(Sheet2!A3:A200000=Sheet1!A10, Sheet2!H3:H200000))}',format_table_headers)
SHEET1.write('G11','{=MIN(IF(Sheet2!A3:A200000=Sheet1!A11, Sheet2!H3:H200000))}',format_table_headers)
SHEET1.write('G12','{=MIN(IF(Sheet2!A3:A200000=Sheet1!A12, Sheet2!H3:H200000))}',format_table_headers)
SHEET1.write('G13','{=MIN(IF(Sheet2!A3:A200000=Sheet1!A13, Sheet2!H3:H200000))}',format_table_headers)
SHEET1.write('G14','{=MIN(IF(Sheet2!A3:A200000=Sheet1!A14, Sheet2!H3:H200000))}',format_table_headers)

SHEET1.write('H9', '=IF(ABS(F9)>ABS(G9),ABS(F9),ABS(G9))',format_table_headers) 
SHEET1.write('H10','=IF(ABS(F10)>ABS(G10),ABS(F10),ABS(G10))',format_table_headers) 
SHEET1.write('H11','=IF(ABS(F11)>ABS(G11),ABS(F11),ABS(G11))',format_table_headers) 
SHEET1.write('H12','=IF(ABS(F12)>ABS(G12),ABS(F12),ABS(G12))',format_table_headers) 
SHEET1.write('H13','=IF(ABS(F13)>ABS(G13),ABS(F13),ABS(G13))',format_table_headers) 
SHEET1.write('H14','=IF(ABS(F14)>ABS(G14),ABS(F14),ABS(G14))',format_table_headers)

# CREATE A PLOT OF VONMISES AND LONGITUDINAL STRESSES VERSUS PIPE WALL THICKNESS
chart1 = workbook.add_chart({'type': 'line'})
# Add a series to the chart.

chart1.add_series({
        'name': 'Vonmises stress ',
        'categories':'=SHEET1!$B$9:$B$14',                     #Thickness in x-axis
        'values': '=SHEET1!$H$9:$H$14'})                       #Vonmises stress in y-axis

chart1.add_series({
        'name': ' Longitudinal stress ',
        'categories':'=SHEET1!$B$9:$B$14',                     #Thickness in x-axis
        'values': '=SHEET1!$E$9:$E$14'})                       #Longitudinal stress in y-axis

chart1.set_title({'name': 'Wall Thickness versus Stress',})
chart1.set_x_axis({'name': 'Pipeline Wall Thickness(mm)',})
chart1.set_y_axis({'name': 'max Long stress(MPa)',})
chart1.set_style(9)

# Insert the chart into the worksheet
SHEET1.insert_chart('E15', chart1)
output3()
workbook.close()

# opens the resultant spreadsheet
os.startfile(execFile)

# parameteric study completed























