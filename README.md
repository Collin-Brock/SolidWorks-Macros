# What is a Macro?
⦁	A macro is a sequence of actions executed with a single input.

# How to add a Macro bound to a Keyboard button in SolidWorks

⦁	Open any document or create a new document.

⦁	Go to Settings Icon > Customize.
 
⦁	Select the Keyboard tab.
 
⦁	Change the Category dropdown to Macros.
 
⦁	Click the three dots next to the New Macro Button command to open the Customize Macro Button window.
 
⦁	Click the three dots to open the file explorer.

⦁	Navigate to and open the macro .swp file
 
⦁	Click Ok to set the macro button

⦁	Assign a keyboard key to the Shortcut(s) column by clicking the empty box and selecting a key
 
# How to add a Macro bound to a Toolbar in SolidWorks

⦁	Open any document in SolidWorks.

⦁	Go to Settings Icon > Customize.
 
⦁	Select the Commands tab.
 
⦁	Change the Toolbars: option to Macro.
 
⦁	Drag the  	Icon to any Toolbar.

⦁	Click the three dots to open the file explorer.

⦁	Navigate to and open the macro .swp file.

⦁	Click Ok to set the macro button or follow the same procedure for the icon choose image to set an image for the button
How to add a Macro bound to a Shortcuts Bar in SolidWorks:

⦁	Open any document in SolidWorks.

⦁	Go to Settings Icon > Customize.
 
⦁	Select the Shortcut Bars tab.
 
⦁	Change the Toolbars: option to Macro.
 
⦁	Select a Shortcut Toolbar that is shown for each different file or state
This toolbar is activated by the “S” key by default

⦁	Drag the Macro Icon to any of the Shortcut Tool Bars.

⦁	Click the three dots to open the file explorer.

⦁	Navigate to and open the macro .swp file.

⦁	Click Ok to set the macro button or follow the same procedure for the icon choose image to set an image for the button
Where to Find Existing Macros: [SolidWorks-Macros](https://github.com/Collin-Brock/SolidWorks-Macros/tree/main/PlainText)

#  Icons

⦁	Icons are 16x16 pixel .bmp file that can be created in Microsoft Paint

# Creating a Macro
⦁	Right-click on the Command Manager at the top of the SolidWorks window.
⦁	Go to Toolbars and select Macro Toolbar.
⦁	You will now see this toolbar
 
⦁	Use this toolbar to:
⦁	Run a macro
⦁	Stop recording a macro
⦁	Start/Pause recording a macro
⦁	Edit a macro

# Recording a Macro

⦁	Click Start Record to record actions. Note that not all actions may be recorded, and the generated code might be suboptimal.
Editing a Macro:

⦁	You can view the existing macro code by selecting Edit. 

⦁	Macros are in the Vault in the Engineering\Engineering Macros folder and are read-only, including comments explaining their functions.

# Macros Available

# B_Autoformat_REMOVES_NOTES.swp

*Will remove notes added to the Sheet Format

⦁	This macro will take the current open drawing on the SolidWorks window and edit the sheet properties template to be formatted with the template in the template path location

⦁	Removes text containing insp: from dimensions

⦁	Removes text containing REF from dimensions, then adds parentheses to that dimension
Carbide Tile Ballon Macro.swp

⦁	First, select a single carbide tile inside a drawing file. Then this macro will add a balloon where you selected has no leader or circular border.
Center_Dimensions.swp

⦁	This macro will get every dimension in the current drawing sheet and center them if the text is within the leader lines

# Custom Properties Fix.swp

*Deleted data in the custom properties inside of configuration cannot be undone with CTRL+Z

*Use only on assembly and part files

⦁	This macro will delete configuration-specific custom properties and consolidate them into the main custom properties, and retain only the weight property in the configuration-specific custom properties.
Fractional Base 32nd.swp

⦁	This macro converts all dimensions on the current drawing sheet to a fractional format, rounded to the nearest 32nd.
Reattach Macro

*Either disable Context toolbars -> Show on Selection in the Customize window under the Toolbars Tab, or move the mouse after selection to dismiss the dimension context toolbar

⦁	First, select a dangling dimension, then move your mouse to dismiss the toolbar. Then this macro will run the Reattach command, which will indicate which leader to reattach with a red X. Select the geometry to reattach to.
Add REV Symbol

⦁	First right right-click anywhere in a SolidWorks Drawing. Then this macro will create a triangle symbol with the current Revision number as defined in the drawing’s properties
Page Title From Sheet Name

⦁	This macro will take the sheet name and create a note in the upper right corner of the drawing containing this name in bold, underlined, italicized, 24 pt font. Only will correctly place the note in the B sheet format
Make Dim Reference

⦁	First, select any number of dimensions in a drawing file. Then this macro will now toggle the parentheses for the selected dimensions.
CUST DRAWING

⦁	First, open a drawing to make the customer print from

⦁	Creates a new drawing as drawingname-CUST

⦁	Deletes all but the last sheet

⦁	Removes all notes

⦁	Removes Tolerance and Revision Blocks

⦁	Removes Detail and Section Views

⦁	Removes welding marks

⦁	Removes Material

⦁	Removes BOM and Balloons

⦁	Removes Heat Treat and Finish

⦁	Adds CONFIDENTIAL DRAWING background for GET and BTI with CUST_DRAWING_BTI

# Image Generator

Creates a PNG and JPG file in the image path using the Filename-png and Filename-jpg, respectively

⦁	First Set Options -> Export File Format -> TIF/PSD/JPG/PNG-> 

⦁	Remove background -> Checked

⦁	Print Capture -> Selected

⦁	DPI -> 400 min

⦁	Paper Size -> B for consistency

⦁Then open an assembly or part

⦁	Then rotate the part to where the image is to be generated

⦁	Creates a PNG and JPG file in the macro path location using the Filename-png and Filename-jpg, respectively

⦁	Sets Feature Manager Tab width to zero

⦁	Zooms to fit

⦁	Removes Shadows

⦁	Hides all Types

⦁	Turns off all spotlights

⦁	Turns up Ambient light level to .75

⦁	Changes the color of any component at any level of the assembly with CAR in the name to R=215 G=200 B=76

⦁	Changes any other color to R=217 G=217 B=217

⦁	Sets Quality to highest on

⦁	Propagates Quality to all components in an assembly

⦁	Sets Quality back to fine after image generation

⦁	Sets Feature Manager Width back to the same as before

# Image Generator BV

Creates a PNG and JPG file in the macro path location using the Filename-png and Filename-jpg, respectively

⦁	First Set Options -> Export File Format -> TIF/PSD/JPG/PNG-> 

⦁	Remove background -> Checked

⦁	Print Capture -> Selected

⦁	DPI -> 400 min

⦁	Paper Size -> B for consistency

⦁ Then open an assembly or part

⦁	Then rotate the part to where the image is to be generated

⦁	Creates a PNG and JPG file in the macro path location using the Filename-png and Filename-jpg, respectively

⦁	Sets Feature Manager Tab width to zero

⦁	Zooms to fit

⦁	Removes Shadows

⦁	Hides all Types

⦁	Turns off all spotlights

⦁	Turns up Ambient light level to .75

⦁	Changes the color of any component at any level of the assembly with CAR in the name to R=100 G=100 B=100

⦁	Changes any other color to R=255 G=0 B=0

⦁	Sets Quality to highest on

⦁	Propagates Quality to all components in an assembly

⦁	Sets Quality back to fine after image generation

⦁	Sets Feature Manager Width back to the same as before

# PIB Creation

⦁	First Open Part or Assembly

⦁	Then rotate the part to where the image is to be generated

⦁	Creates a new drawing document

⦁	Saves as Filename-PIB

⦁	Removes all items from the sheet format

⦁	Adds View as Current View

⦁	Scales the View 2x automatically generated scale

# StepFilePerConfiguration.swp

⦁	Saves each configuration of a part or assembly as a STEP file

⦁	Saves to step file path location

# Drawing Creation .swp

⦁	Adds a sheet for each configuration

⦁	Sets the sheet format provided in the format path location

⦁	Adds Bold, Underlined, and Italicized title consisting of the configuration's name to the  upper right-hand corner

⦁	Renames the sheet to the configuration's name

⦁	Inserts a view of the top and right in the left-hand side respectively

# Selected Frac Base 16th.swp 

⦁	Takes all selected entities and changes them to a fractional base of 16th if they are a display dimension. 

# Selected Frac Base 32nd.swp 

⦁	Takes all selected entities and changes them to a fractional base of 32nd if they are a display dimension.

# Selected Dec DocPLCS.swp 

⦁	Takes all selected entities and changes them to decimal with the places of the document settings if they are a display dimension

# +PrimaryPrecision.swp

⦁	Takes all selected entities and increases the precision by one if they are display dimensions.

# -PrimaryPrecision.swp

⦁	Takes all selected entities and decreases the precision by one if they are display dimensions.

# Writing VBA Scripts

Comments

⦁	Comments are indicated by an apostrophe (') and are typically displayed in green.

The main procedure is defined using Sub and End Sub statements:

Sub Main()

' Code goes here
  Exit Sub 'Exits the current Sub (Optional)
End Sub

Functions

⦁	Functions are defined using Function and End Function statements:

Function FunctionName(inputs(Optional)) As ReturnType(Optional)

' Code goes here 

Exit Function 'Exits the current Function (Optional) 

End Function

About VBA

⦁	VBA is an object-oriented programming language. This means that software is developed by interacting with objects and modifying their properties.

⦁	The challenging aspect is identifying the correct object and property to modify.

⦁	Use the Object Browser   to search for objects and their properties.

⦁	Properties or other objects can be accessed by using the object.object/property format

Declaring Variables

⦁	Declare variables and define their type:

Dim variableName As DataType

⦁	Common data types include Long, Boolean, Variant, SldWorks.X, and others.

Setting a variable to an object

Set variableName = object

⦁	Be sure to put Option Explicit in the beginning of any program so you are forced by the compiler to define every variable type.

⦁	 If you are changing a property, you do not need to use the Set command and in fact it will invoke an error if you do.

Finally, there are built-in commands you can use to control your programs, including:

If variable_name = value

End If

While variable_name = value

Wend

For i = 0 to UBound(array_variable)

Next I 
