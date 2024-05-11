# New Project Template VBA

## Overview

This VBA project template provides a base for creating macros in Excel. 

## Purpose

The purpose of this template is to provide a starting point with basic components, such as message boxes, timing, screen updating, and formatting.


## Usage

1. Copy and paste this code into your VBA editor.
2. Modify the parameters passed to the `Formatting_Example` macro in the `Template_Macro` subroutine (range to edit and table name to use). 
- Example: `Formatting_Example "A1:C10", "tableName"`
3. Customize the `Formatting_Example` subroutine to perform the desired tasks.

## Subroutines

- **Template_Macro:** Main subroutine for executing the macro. 
- **Confirmation_MsgBox:** Displays a confirmation message box before proceeding with the macro execution.
- **Withdrawal_MsgBox:** Displays a message box indicating that the macro was not run.
- **Success_MsgBox:** Displays a message box with the execution time if the macro runs successfully.
- **ScreenUpdating:** Controls the screen updating feature to improve macro performance.
- **Formatting_Example:** Example subroutine for formatting Excel data. 

## Notes

- Feel free to expand upon or modify this template to accommodate additional functionality or specific project needs.
