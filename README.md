1. Overview

The IFC Parameter Automation add-in is developed to support automatic creation and binding of IFC-related shared parameters in Autodesk Revit.

The tool:

Loads IFC Property Set definitions from JSON files

Creates shared parameters automatically

Binds parameters to selected Revit categories

Supports Windows, Doors, and Walls

Export null Parameters to an excel report

This tool was developed as part of a Project Work.

2. System Requirements

Autodesk Revit 2024

Windows 10 or Windows 11

3. Installation Instructions
Step 1 – Extract Files

Extract the provided ZIP file to: C:\ProgramData\Autodesk\Revit\Addins\2024

Step 2 – Start Revit

Open Autodesk Revit 2024

Go to the Add-Ins tab

In External Tools Click IFC Parameter

4. How to Use

Click Add IFC Parameters

Select an element (Window, Door, or Wall)

The plugin loads the corresponding IFC Property Set

Select the desired parameters

Click Confirm

The parameters will be automatically created and bound

Click Check IFC Parameters

A save as dialogue pops up to save the report

5. Technical Notes

Shared parameters JSON Files are stored in:

%ProgramData%\Autodesk\Revit\Addins\2024\Psets\

6. Limitations

Currently supports:

Windows

Doors

Walls

Revit 2024 only

7. Author

Developed by: Karim Ishak Fahmy Guirguis
Project Work – TU Dresden
Year: 2026