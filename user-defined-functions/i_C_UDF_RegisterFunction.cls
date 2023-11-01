VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "i_C_UDF_RegisterFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function bRegisterUserDefinedFunction(ByVal sFunctionName As String, ByVal sDESCRIPTION As String, ByRef saArgs() As String, Optional ByVal sCategory As String = "Other User Defined Functions") As Boolean
   Dim bHasError As Boolean
   
   On Error GoTo Catch
   Application.MacroOptions _
      Macro:=sFunctionName, _
      Description:=sDESCRIPTION, _
      ArgumentDescriptions:=saArgs, _
      Category:=sCategory
   
Finally:
   On Error Resume Next
   bRegisterUserDefinedFunction = Not bHasError
   Exit Function
   
Catch:
   bHasError = True
   Resume Finally
End Function


