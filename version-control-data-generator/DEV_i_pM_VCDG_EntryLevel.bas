Attribute VB_Name = "DEV_i_pM_VCDG_EntryLevel"
' Package: VersionControlDataGenerator
'============================================================================================
'   NAME:     DEV_i_pM_VCDG_EntryLevel
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
' 0.2.0    25.10.2023    gueleh    Imported from FF2 and adapted
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================

Option Explicit
Option Private Module

Public Sub DEV_i_p_ExportDataForVersionControl()

   Const sERROR_MESSAGE_PREFIX As String = "Following errors occurred:" & vbNewLine
   Const sERROR_MESSAGE_SUFFIX As String = vbNewLine & "Please look at the entries printed to the direct windows for more details."
   Dim bHasErrors As Boolean
   Dim sErrorMessage As String
   Dim iCalcSetting As Integer
   
   With Application
      iCalcSetting = .Calculation
      .ScreenUpdating = False
      .Calculation = xlCalculationManual
      .EnableEvents = False
   End With
   
Try:
   On Error GoTo Catch
   
   Dim oC_VersionControlExport As New DEV_i_C_VCDG_Export
   
   If b_i_p_VCDG_EXPORT_CODE_MODULES Then
         If Not _
      oC_VersionControlExport.bExportAllComponents() _
         Then Err.Raise 9999, , "Export of components failed."
   End If
   
   If b_i_p_VCDG_EXPORT_DEFINED_NAME_DATA Then
         If Not _
      oC_VersionControlExport.bExportNameData() _
         Then Err.Raise 9999, , "Export of defined name data failed."
   End If
   
   If b_i_p_VCDG_EXPORT_WORKSHEET_META_DATA Then
         If Not _
      oC_VersionControlExport.bExportWorksheetNameData() _
         Then Err.Raise 9999, , "Export of worksheet meta data failed."
   End If
   
   If b_i_p_VCDG_EXPORT_PROJECT_REFERENCES Then
         If Not _
      oC_VersionControlExport.bExportReferenceData() _
         Then Err.Raise 9999, , "Export of project reference data failed."
   End If

   If b_i_p_VCDG_EXPORT_SETTINGS_SHEET_CONTENTS Then
         If Not _
      oC_VersionControlExport.bExportSettingsSheetData() _
         Then Err.Raise 9999, , "Export of settings sheets data failed."
   End If

   If b_i_p_VCDG_EXPORT_NAMED_RANGE_CONTENTS Then
         If Not _
      oC_VersionControlExport.bExportRangeContentData() _
         Then Err.Raise 9999, , "Export of settings sheets data failed."
   End If

Finally:
   On Error Resume Next
   If bHasErrors Then
      MsgBox sERROR_MESSAGE_PREFIX & sErrorMessage, vbCritical
   End If

   With Application
      .Calculation = iCalcSetting
      .EnableEvents = True
      .Calculate
      .ScreenUpdating = True
   End With
   
   Exit Sub

Catch:
'we want to execute as many calls as possible and not stop when one
'execution fails
   bHasErrors = True
   sErrorMessage = sErrorMessage & Err.Description & vbNewLine
   Resume Next
End Sub
