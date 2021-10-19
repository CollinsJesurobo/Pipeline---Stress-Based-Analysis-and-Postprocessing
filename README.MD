                 Stress Based Design - Parametric study of the influence of wall thickness on pipeline stress.

Tools used:
 -  Abaqus
 -  Python
 - xlsxwriter


Description:

The allowable stress in a pipeline is determined using the  Allowable Stress Design (ASD) methodology in accordance with DNVGLSTF101 taking into consideration the design factors corresponding to a safety class.
The stress response on the pipeline is compared to the allowable values to determine the acceptance criteria (pass/fail).

This tool has been prepared to extract stresses and compare to them to the limiting yield value, and to verify the stress behaviour due to effect of  wall thickness variation.
A basecase abaqus input file of the pipeline resting on seabed was first prepared . The wall thickness in the input file was parameterized in such a way as to allow variation of the thickness value .
Six thickness values has been evaluated in this study using the python parametric study method for Abaqus to generate 6 input files and automatically batch runs to obtain 6odbs which is called LOADCASES in this analysis. The thickness values are :[15.9mm,19.1mm,22.3mm,25.1mm,27.1mm, 30.2mm]. The parametric study script  used here is called "study.psf".

Upon successful completion of the preprocessing, the 6odbs are generated in the project folder. Each odbs contains 6 steps and 999 elements with stress extracted from each element having a default of 4 integration points .
Thus, the total loadcase data  is 6 x 6 x 999*4 =143856 loadcases for each variable.

The postprocessing is done in the Python scripts called "Python postprocessor". The script open the odbs one after the other ,  loop through each loading steps and extract the results for  stresses S11,S22,S33,S12 and Smises,
then write the result to excel sheet2 with help of xlsxwriter, then open sheet1 and write the maximum, minimum values of S11,S22,S33,S12 and Smises and their corresponding element, loadstep were it occurs.
Thus with a table of values of worst S11 and Smises for each loadcase and their corresponding thickness values, a plot is made as shown in the attached excel.
From the plot, one can see that as the wall thickness increases, the stresses reduces which depicts an inverse  relationship.

1.Preprocessing
- prepare a working basecase abaqus model of the pipeline
- parametrize the variable for study in the input file
- prepare parametric study file (.psf) that contains the values to be varied 
- run the the.psf file until successful completion 

2. Postprocessing
- prepare the python script for postprocessibg
- open all the odbs, loop through each of them
- loop through the steps in the odb
- extract results for stresses S11,S22,S33,S12 and Smises,
- compare stresses to Allowable stress value
- determine the worst Stresses and associated worstloadcase
- create excel workbook and worksheets using xlsxwriter library
- write results to the sheets 
- make a plot of axial and vonmises stresses 7versus wall thickness
- parametric study completed





