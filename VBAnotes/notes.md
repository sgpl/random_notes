# VBA lets us create (UDFs) User Defined Functions to automate stuff. 

### two options
	- hardcode directly into excel workbook
	- create an addin accessible by other users


### need to know

	- record a macro
	- add modules
	- browse objects
	- get some background on declaring variables


	- record a macro
		1. from tools >macro >record a macro
		2. from the developer tab (shortcut can be defined from console)

		## note : need to figure out mac specific code for executing code (as opposed to just the shortcut or pressing the run button)

		## also need to save in .xlsm format > macro-enabled workbook

	- add modules

	- browse objects
		click on 'object browser' icon on the toolbar
			- greek icon > methods
			- hand pointing icon > properties

	- get some background on declaring variables
##

	- find 'immediate window' from view > immediate window

## immediate window is critical for debugging code. 


----------------------------------------
	- loops 
	- dimension and instantiate objects
	- create sub routines and user defined fxns

----------------------------------------
loops
----------------------------------------

# For-Next loops

For counter = start To end [Step step]
	[statement]
[Exit For]
	[statement]
Next [counter]

eg: 

Sub ForNextLoopExample() # start of program
    Dim i As Integer # Dimensions (declares) i as an integer
    Dim iCounter As Integer # see above
    
    For i = 1 To 100 # loops from 1 to 100 ; could also write
    	# For i = 1 To 100 Step 5
        iCounter = iCounter + 1 # 
    Next i
    MsgBox iCounter
    
End Sub

# declaring Double instead of integer == float so 0.00
----------------------------------------

# For-Each-Next Loop

Syntax: 

For Each element In group
	[statement]
[Exit For]
	[statement]
Next [element]


eg: 



----------------------------------------

Do-While and Do-Until loops

Syntax: 

Do [{While | Until} condition]
	[statement]
[Exit Do]
	[statement]
Loop




