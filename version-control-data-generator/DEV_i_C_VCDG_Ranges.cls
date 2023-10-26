VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DEV_i_C_VCDG_Ranges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Package: VersionControlDataGenerator
'============================================================================================
'   NAME:     DEV_i_C_VCDG_Ranges
'============================================================================================
'   Purpose:  exports all data relevant for convenient version control
'   Access:   Private
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
' 0.2.0    25.10.2023    gueleh    Initially created, based on code imported from Flow Framework 2
'--------------------------------------------------------------------------------------------
'   BACKLOG
'   ''''''''''''''''''''
'   none
'============================================================================================
Option Explicit

Private oCol_m_Ranges As New Collection
Private oWks_m_RangeInfos As Worksheet
Private l_m_RowStart As Long
Private l_m_ColName As Long
Private l_m_ColDefinedName As Long

Public Sub Construct _
( _
   ByRef oWksRangeInfos As Worksheet, _
   ByVal lRowStart As Long, _
   ByVal lColName As Long, _
   ByVal lColDefinedName As Long _
)
   Set oWks_m_RangeInfos = oWksRangeInfos
   l_m_RowStart = lRowStart
   l_m_ColDefinedName = lColDefinedName
   l_m_ColName = lColName
End Sub

Public Property Get oCol_prop_r_VersionControlRanges() As Collection
   Set oCol_prop_r_VersionControlRanges = oCol_m_Ranges
End Property

Public Function bBuildCollectionWithRanges() As Boolean

   Dim bHasError As Boolean

Try:
   On Error GoTo Catch

   Dim oC As DEV_i_C_VCDG_Range
   Dim lRow As Long
   Dim bOneNotValidEntryFound As Boolean
   Dim sDefinedName As String
   Dim oRng As Range
   
   If oWks_m_RangeInfos Is Nothing Then
      Err.Raise _
         9999, , _
         "Construct has not been called, thus no worksheet object set."
   End If
      
   With oWks_m_RangeInfos
      For lRow = l_m_RowStart To .UsedRange.Rows.Count + l_m_RowStart
         sDefinedName = .Cells(lRow, l_m_ColDefinedName).Value2
         If Len(sDefinedName) > 0 Then
            On Error Resume Next
            Set oRng = ThisWorkbook.Names(sDefinedName).RefersToRange
            If Err.Number > 0 Then
               Err.Clear
               Debug.Print "Defined Name " & sDefinedName & " is not valid."
               Set oRng = Nothing
               bOneNotValidEntryFound = True
            End If
            On Error GoTo Catch
            If Not oRng Is Nothing Then
               Set oC = New DEV_i_C_VCDG_Range
               oC.Construct .Cells(lRow, l_m_ColName).Value2, sDefinedName, oRng
               oCol_m_Ranges.Add oC
            End If
         End If
      Next lRow
   End With
      
'End of your code <<<<<<<

'Fixed, don't change
Finally:
   On Error Resume Next

   If bOneNotValidEntryFound Then
      MsgBox "At least one entry for version control of range contents is not valid, " & _
         "please look into the direct window after processing to find out more.", vbExclamation
   End If
   
   bBuildCollectionWithRanges = Not bHasError
   Exit Function

Catch:
   bHasError = True
   Debug.Print Err.Number & ", " & Err.Description
   Resume Finally

End Function



