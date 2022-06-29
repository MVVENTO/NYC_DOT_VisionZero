#
#
# 
#  
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import fonts

data = {

    # New speed humps installed
    "5": {
        "Metric Name": "New speed humps installed",
		"January 2022 (spatial)": 0,
		"February 2022 (spatial)": "N/A",
		"March 2022 (spatial)": 0,
		"April 2022 (spatial)":0,
		"May 2022 (spatial)":0,
        # dynamically add months to python script
		"YTD Total 2022 (spatial)":0, # sum of Columns C2, D2, E2, F2, G2,
        "Overall Target" : 250,
        "Primary Contacts" : "Ann Marie Doherty/ Alicia Posner/ Seth Hostetter",
        "Data Contacts" : "Seth Hostetter/Arthur Getman"

    },
    # Safety projects (SIPs) completed _ VZ priority geographies only
    "7": {
        "Metric Name": "Safety projects (SIPs) completed _ VZ priority geographies only",
		"January 2022 (spatial)": 0,
		"February 2022 (spatial)": "N/A",
		"March 2022 (spatial)": "N/A",
		"April 2022 (spatial)":"N/A",
		"May 2022 (spatial)":"N/A",
        # dynamically add months to python script
		"YTD Total 2022 (spatial)":"N/A",  
        "Overall Target" : 50,
        "Primary Contacts" : "Jeff Malamy/ Sean Quinn",
        "Data Contacts" : "Seth Hostetter/Arthur Getman"
    },
    
    "8": {
        "Metric Name": "Safety projects (SIPs) completed _ all projects ",
		"January 2022 (spatial)": 3, # sum of SIP_ Corridors + SIP_Intersections 
		"February 2022 (spatial)": 1, # sum of SIP_ Corridors + SIP_Intersections
		"March 2022 (spatial)": 0, # sum of SIP_ Corridors + SIP_Intersections
		"April 2022 (spatial)":2, # sum of SIP_ Corridors + SIP_Intersections
		"May 2022 (spatial)":0,
        # dynamically add months to python script
		"YTD Total 2022 (spatial)":6, # sum of Columns C4, D4, E4, F4, G4,
        "Overall Target" : 250,
        "Primary Contacts" : "Ann Marie Doherty/ Alicia Posner/ Seth Hostetter",
        "Data Contacts" : "Seth Hostetter/Arthur Getman"
    },

}