Attribute VB_Name = "DEV_i_pM_Globals_VCDG"
' Package: VersionControlDataGenerator
'============================================================================================
'   NAME:     i_pM_Globals_VersionControlDG
'============================================================================================
'   Purpose:  global settings required for version control data geneator
'   Access:   Private
'   Type:     Module
'   Author:   Günther Lehner
'   Contact:  guleh@pm.me
'   GitHubID: gueleh
'   Required:
'   Usage: please refer to the guidance document and to the guidance directly
'     in the code
'--------------------------------------------------------------------------------------------
'   VERSION HISTORY
'   Version    Date    Developer    Changes
'   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 0.2.0    25.10.2023    gueleh    Added globals and function for range observation worksheet
'   0.1.0    24.10.2023    gueleh    Initially created
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit
Option Private Module

Private Const s_m_COMPONENT_NAME As String = "i_pM_Globals_VersionControlDG"

'DECLARATIONS

'CONTRACT DEFINITION SETTINGS
      'These settings inform the code about the expected structure of the settings
      'sheets and the range definition sheets - this means, you have to comply
      'with these settings when setting up the respective sheets.
      
      'Adapt the constant values to meet your requirements.

   'Settings Sheet

   'Range Definitions Sheet
      'The Const values are set in method bExportRangeContentData of
      'DEV_i_C_VersionControlExport as they are only required there
      'If you need to change them, do it there

'DATA GENERATION SCOPE SETTINGS
   'These settings inform the code about which version control data do have to
   'be generated. Adapt the constant values to meet your requirements.

   'Export code modules for version control?
      'PLEASE NOTE: This requires the reference
      '"Microsoft Visual Basic for Applications Extensibility 5.3".
      'This library allows VBA to access the object model of the visual basic editor, i.e.
      'with it you can read and change code with code.
      'Your security settings might not allow
      'to activate or use this libary. In this case please set the const to False.
   Public Const b_i_p_VCDG_EXPORT_CODE_MODULES As Boolean = True
   
   'Export data of defined names of the workbook (also worksheet scope names)?
   Public Const b_i_p_VCDG_EXPORT_DEFINED_NAME_DATA As Boolean = True

   'Export meta data of worksheets?
   Public Const b_i_p_VCDG_EXPORT_WORKSHEET_META_DATA As Boolean = True

   'Export settings stored in settings sheets?
      'PLEASE NOTE: For this to work the sheets have to meet the
      'contractual requirements (see guidance) and you also need to add these
      'worksheets to the function below which returns a collection
      'with the settings sheets.
   Public Const b_i_p_VCDG_EXPORT_SETTINGS_SHEET_CONTENTS As Boolean = True

   'Export contents of named ranges to be included into version control?
      'PLEASE NOTE: For this to work the sheet and the named ranges have to meet the
      'contractual requirements (see guidance) and you also need to add the
      'worksheet containing the parameters to the function below
   Public Const b_i_p_VCDG_EXPORT_NAMED_RANGE_CONTENTS As Boolean = True
   
   'Export data of references to used libraries
   Public Const b_i_p_VCDG_EXPORT_PROJECT_REFERENCES As Boolean = True
   
' Procedure Name: oCol_i_p_VCDG_SettingsSheets
' Purpose: builds and returns a collection with worksheets which are treated as "settings sheets" during version control data generation
' Procedure Kind: Function
' Procedure Access: Public
' Return Type: Collection
' Author: Günther Lehner
' Contact:  guleh@pm.me
' GitHubID: gueleh
' Requires:
' Usage: see guidance document and code comments
'------------------------------------------------------------------------------------
' Version   Date      Developer   Changes
' 0.1.0    24.10.2023    gueleh  Initially created
'------------------------------------------------------------------------------------
' Backlog:
' None
'------------------------------------------------------------------------------------
Public Function oCol_i_p_VCDG_SettingsSheets() As Collection
'Do not change
   Dim oCol As New Collection
   With oCol
'REMOVE THE DEMO ENTRIES
      .Add wksDemoRangeDefSheet
      .Add wksDemoSettingsSheet
'ADD YOUR SETTINGS SHEETS TO THE COLLECTION HERE


'Do not change
   End With
   Set oCol_i_p_VCDG_SettingsSheets = oCol
End Function

' Purpose: provides the worksheet object for the sheet containing the parameters
' for inclusion of named ranges into version control data
' 0.2.0    25.10.2023    gueleh    Initially created
Public Function oWks_i_o_VCDG_RangeSettings() As Worksheet
'REPLACE THE DEMO SHEET WITH YOUR SHEET CONTAINING THE PARAMETERS
'FOR INCLUSION OF RANGE CONTENTS IN VERSION CONTROL DATA
   Set oWks_i_o_VCDG_RangeSettings = wksDemoRangeDefSheet
End Function
