VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "i_C_SETTINGS_Sheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Package: SETTINGS (allowing to work with a settings sheet)
'============================================================================================
'   NAME:     i_C_SETTINGS_Sheet
'============================================================================================
'   Purpose:  gets settings from a settings sheet which fulfills the structural contract
'   Access:   Public
'   Type:     Class Module
'   Author:   Günther Lehner
'   Contact:  guleh@pm.me
'   GitHubID: gueleh
'   Required:
'   Usage:
'--------------------------------------------------------------------------------------------
'   VERSION HISTORY
'   Version    Date    Developer    Changes
'   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 0.2.0    26.10.2023    gueleh    Imported from FF2 and renamed
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit

Private oWks_m_Settings As Worksheet
Private l_m_RowStart As Long
Private l_m_ColID As Long
Private l_m_ColName As Long
Private l_m_ColValue As Long

Public Property Get oWks_prop_r_SettingsSheet() As Worksheet
   Set oWks_prop_r_SettingsSheet = oWks_m_Settings
End Property


' Procedure Name: bConstruct
' Purpose: sets up the parameters needed to work with the settings sheet
' Procedure Kind: Function
' Procedure Access: Public
' Parameter oWks (Worksheet): the worksheet object of a settings sheet
' Parameter lRowStart (Long): the row containing the first setting
' Parameter lColID (Long): the column containing the settings ID, e.g. the defined name
' Parameter lColName (Long): the column containing the name of a setting
' Parameter lColValue (Long): the column containing the value of a setting
' Return Type: Boolean
' Author: Günther Lehner
' Contact:  guleh@pm.me
' GitHubID: gueleh
' Requires: a worksheet fulfilling the contract, i.e. having one setting per row
'  with a unique ID, a name and a value in a dedicated column of the worksheet
' Usage:
'------------------------------------------------------------------------------------
' Version   Date      Developer   Changes
' 0.3.0    26.10.2023    gueleh    Adapted from FF2 code
'------------------------------------------------------------------------------------
' Backlog:
' None
'------------------------------------------------------------------------------------
Public Function bConstruct _
( _
   ByRef oWks As Worksheet, _
   ByVal lRowStart As Long, _
   ByVal lColID As Long, _
   ByVal lColName As Long, _
   ByVal lColValue As Long _
) As Boolean

   Dim bHasError As Boolean

Try: On Error GoTo Catch

   Set oWks_m_Settings = oWks
   l_m_RowStart = lRowStart
   l_m_ColID = lColID
   l_m_ColName = lColName
   l_m_ColValue = lColValue

Finally: On Error Resume Next
   bConstruct = Not bHasError
   Exit Function

Catch:
   bHasError = True
   Debug.Print Err.Number & ", " & Err.Description
   Resume Finally
End Function

' Procedure Name: bGetSettingsFromSettingsSheet
' Purpose: building and returning a collection for all setting rows in a settings sheet
' Procedure Kind: Function
' Procedure Access: Public
' Parameter oCol_Settings (Collection): collection with f_C_Setting instances, one for each row/setting - comes in empty, is returned
' Return Type: Boolean
' Author: Günther Lehner
' Contact:  guleh@pm.me
' GitHubID: gueleh
' Requires:
' Usage:
'------------------------------------------------------------------------------------
' Version   Date      Developer   Changes
' 0.3.0    26.10.2023    gueleh    Adapted from FF2 code
'------------------------------------------------------------------------------------
' Backlog:
' None
'------------------------------------------------------------------------------------
Public Function bGetSettingsFromSettingsSheet _
( _
   ByRef oCol_Settings As Collection _
) As Boolean

   Dim oCSetting As i_C_SETTINGS_Item
   Dim lRow As Long
   Dim sID, sName As String

   Dim bHasError As Boolean

Try: On Error GoTo Catch

   For lRow = l_m_RowStart To oWks_m_Settings.UsedRange.Rows.Count + l_m_RowStart
      sID = CStr(oWks_m_Settings.Cells(lRow, l_m_ColID).FormulaLocal)
      sName = CStr(oWks_m_Settings.Cells(lRow, l_m_ColName).FormulaLocal)
      If Len(sID) > 0 And Len(sName) > 0 Then
         Set oCSetting = New i_C_SETTINGS_Item
         oCSetting.sID = sID
         oCSetting.sName = sName
         oCSetting.sValue = CStr(oWks_m_Settings.Cells(lRow, l_m_ColValue).Value2)
         oCSetting.sFormulaLocal = CStr(oWks_m_Settings.Cells(lRow, l_m_ColValue).FormulaLocal)
         oCol_Settings.Add oCSetting
      End If
   Next lRow
   
Finally: On Error Resume Next
   bGetSettingsFromSettingsSheet = Not bHasError
   Exit Function

Catch:
   bHasError = True
   Debug.Print Err.Number & ", " & Err.Description
   Resume Finally
End Function



