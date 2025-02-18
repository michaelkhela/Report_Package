# Automation of Narrative (MSEL + PLS + VABS)

# Use this code to create study visit reports that can be sent to the parent/guardian of participants.

# Please note the following before running the code:
# - ALL INPUT DOCUMENTS HAVE TO BE CLOSED

# Script created by Michael Khela on 10/20/2023, reorganized 01/27/2024

import sys 
import os

#user inputs below

# INSERT SUBJECT ID BELOW
sub_id = "12012"

# INSERT PLS ADMIN (enter _ for ____)
pls_admin = "1"

# INSERT MSEL ADMIN (enter _ for ____)
msel_admin = "1"

# INSERT ADMIN 1 (enter _ for ____)
admin_1 = "1"

# INSERT ADMIN 2 (enter _ for ____)
admin_2 = "1"

# INSERT YOUR FILE PATH TO "Report_Package"
root_filepath = r"/Users/michaelkhela/Desktop/Report_Package/"

# INSERT NAME OF REDCAP EXPORT (add .csv if not present already)
redcap_filename = "0_REDCap_export.csv"

#-----------DO NOT EDIT BELOW-------------------------------------------------------------------------------------------
sys.path.append(os.path.abspath(root_filepath + "Inputs/"))
from Report_Automation import report_fcn

report_fcn(root_filepath, sub_id, pls_admin, msel_admin, admin_1, admin_2, redcap_filename)
