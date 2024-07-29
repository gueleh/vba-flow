Attribute VB_Name = "i_MDemoUsage"
'   NAME:     i_MDemoUsage
'============================================================================================
'   Purpose:  demonstrates usage of i_C_DriveMapper for provided valid network path
'   Access:   Public
'   Type:     Modul
'   Author:   Günther Lehner
'   Contact:  guleh@pm.me
'   GitHubID: gueleh
'   Required: i_C_DriveMapper
'   Usage:
'--------------------------------------------------------------------------------------------
'   VERSION HISTORY
'   Version    Date    Developer    Changes
'   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   1.0.0    29.07.2024    gueleh    Initially created
'============================================================================================
Option Explicit

Private Const s_m_COMPONENT_NAME As String = "i_MDemoUsage"
Private Const sADDRESS_PATH As String = "B1"
Private Const sADDRESS_FILES As String = "B6"
Private Const sADDRESS_FOLDERS As String = "C6"

Public Sub DemonstrateDriveMapper()
    Dim oCDriveMapper As New i_C_DriveMapper
    Dim oRootFolder As Scripting.Folder
    Dim oFile As Scripting.File
    Dim oFolder As Scripting.Folder
    Dim bHasError As Boolean
    Dim l As Long
    
    On Error GoTo Catch
    
    Set oRootFolder = oCDriveMapper.oRootFolder(CStr(wksTestDrive.Range(sADDRESS_PATH).Value2))
    wksTestDrive.Range(sADDRESS_FILES).Resize(wksTestDrive.UsedRange.Rows.Count).EntireRow.ClearContents
    l = 0
    For Each oFile In oRootFolder.Files
        wksTestDrive.Range(sADDRESS_FILES).Offset(l).Value2 = oFile.Name
        l = l + 1
    Next oFile
    l = 0
    For Each oFolder In oRootFolder.SubFolders
        wksTestDrive.Range(sADDRESS_FOLDERS).Offset(l).Value2 = oFolder.Name
    Next oFolder
    
    
Finally:
    On Error Resume Next
    If bHasError Then MsgBox "Yikes! Some' gone wrong", vbCritical
    Exit Sub
Catch:
    bHasError = True
    Resume Finally
End Sub



