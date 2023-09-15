# Parsing-xml-file-in-Excel
Parsing database from xml to Excel (Python/VBA)


# AVEVA ISM ISM Standard Xml Class Library format.xml
The AVEVA ISM ISM Standard Xml Class Library format.xml file is a database that needs to be populated in Excel

# XML_Python_Excel.py
The XML_Python_Excel.py file is an implementation of Excel database parsing using Python.
Implemented:
 -individual tab creation
 -bold column names
 -correct output
The resulting Excel file is saved under the name data_3.xlsx.

# XML_vba_Excel.txt
The XML_vba_Excel.txt file is an implementation of database parsing in Excel using VBA.
The Microsoft XML library, v 6.0, must be patched in Visual Basic for the macro to work correctly
Implemented:
 -selection of the file to be processed
 -creation of separate tabs
 -bolding of names
 -column sorting
 -correct output of information
 -displaying information by cell size
 -display a message when the macro has stopped working
 -optimising the macro process by disabling screen refreshes 
The resulting Excel file is saved under the name Date_(day.month.year).xlsx.

# Macro.xlsm
The Macro.xlsm file is an implementation of the XML_vba_Excel.txt macro in Excel.
The Microsoft XML library, v 6.0, must be patched in Visual Basic for the macro to work correctly
Implemented:
 -button to start the macro
 -selection of the file to be processed
 -creation of separate tabs
 -bolding of names
 -column sorting
 -correct output of information
 -displaying information by cell size
 -display a message when the macro has stopped working
 -optimising the macro process by disabling screen refreshes 
The resulting Excel file is saved under the name Date_(day.month.year).xlsx.
