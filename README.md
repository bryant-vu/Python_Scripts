Creates PowerPoint decks. Formatting and refreshing of data is handled in Excel/VBA; Python just strings together these workbooks so the entire process can run without the need for user interjection. 

GUI Script:
  - Calls on Master_Script_v2.py
  - Allows for user input including date, option to refresh data, option to choose which (or all) decks to run
  
Master_Script:
  - Calls on various Excel.xlsm workbooks connected to OLAP cube(s) and runs macros to export slides to PowerPoint
  
Memo: Pyinstaller is leveraged for user accessibility (shortcut to .exe file is located on desktop so user only needs to double click this icon to initiate deck creation process.)
