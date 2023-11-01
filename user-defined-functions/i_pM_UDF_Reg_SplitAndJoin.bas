Attribute VB_Name = "i_pM_UDF_Reg_SplitAndJoin"
Option Explicit
Option Private Module

Public Function b_i_p_RegisterJoin() As Boolean
   Const sFUNCTION_NAME As String = "s_i_g_Join"
   Const sDESCRIPTION As String = "Concatenates the cell contents of a range, " _
      & "delimited by a provided delimiter, optionally skipping empty cells, optionally being not a volatile function."
   
   Dim oC As New i_C_UDF_RegisterFunction
   Dim saArgs(1 To 4) As String
   Dim bHasError As Boolean
   
   On Error GoTo Catch
   saArgs(1) = "Range with contents to join"
   saArgs(2) = "Delimiter to use"
   saArgs(3) = CStr(True) & " to skip empty cells (optional)"
   saArgs(4) = CStr(False) & " for function not being volatile, i.e. not being automatically updated, e.g. to increase performance of calculation (optional)"
   If Not oC.bRegisterUserDefinedFunction(sFUNCTION_NAME, sDESCRIPTION, saArgs) Then Err.Raise 9999

Finally:
   On Error Resume Next
   b_i_p_RegisterJoin = Not bHasError
   Exit Function
   
Catch:
   bHasError = True
   If Err.Number = 9999 Then
      Debug.Print s_i_p_ERROR_TEXT_REGISTER_METHOD & sFUNCTION_NAME
   Else
      Debug.Print "b_i_p_RegisterJoin()" & s_i_p_ERROR_TEXT_REGISTER & sFUNCTION_NAME
   End If
   Resume Finally
End Function

Public Function b_i_p_RegisterSplit() As Boolean
   Const sFUNCTION_NAME As String = "sa_i_g_Split"
   Const sDESCRIPTION As String = "Takes a seed string with elements that are separated by a delimiter, " _
      & "splits it into its elements and returns the elements in an array formula " _
      & ". Defaults to being volatile, this can be turned off."
   
   Dim oC As New i_C_UDF_RegisterFunction
   Dim saArgs(1 To 4) As String
   Dim bHasError As Boolean
   
   On Error GoTo Catch
   saArgs(1) = "String with elements to be split"
   saArgs(2) = "Delimiter to use"
   saArgs(3) = CStr(True) & " to print results to rows instead of columns (optional)"
   saArgs(4) = CStr(False) & " for function not being volatile, i.e. not being automatically updated, e.g. to increase performance of calculation (optional)"
   If Not oC.bRegisterUserDefinedFunction(sFUNCTION_NAME, sDESCRIPTION, saArgs) Then Err.Raise 9999

Finally:
   On Error Resume Next
   b_i_p_RegisterSplit = Not bHasError
   Exit Function
   
Catch:
   bHasError = True
   If Err.Number = 9999 Then
      Debug.Print s_i_p_ERROR_TEXT_REGISTER_METHOD & sFUNCTION_NAME
   Else
      Debug.Print "b_i_p_RegisterSplit()" & s_i_p_ERROR_TEXT_REGISTER & sFUNCTION_NAME
   End If
   Resume Finally
End Function

Public Function b_i_p_RegisterSplitCountNeededCells() As Boolean
   Const sFUNCTION_NAME As String = "s_i_g_SplitCountNeededCells"
   Const sDESCRIPTION As String = "Takes a seed string with elements that are separated by a delimiter, " _
      & "calculates how many cells will be needed when using the function sa_i_g_Split on it."
   
   Dim oC As New i_C_UDF_RegisterFunction
   Dim saArgs(1 To 2) As String
   Dim bHasError As Boolean
   
   On Error GoTo Catch
   saArgs(1) = "String with elements to be split"
   saArgs(2) = "Delimiter to use"
   If Not oC.bRegisterUserDefinedFunction(sFUNCTION_NAME, sDESCRIPTION, saArgs) Then Err.Raise 9999

Finally:
   On Error Resume Next
   b_i_p_RegisterSplitCountNeededCells = Not bHasError
   Exit Function
   
Catch:
   bHasError = True
   If Err.Number = 9999 Then
      Debug.Print s_i_p_ERROR_TEXT_REGISTER_METHOD & sFUNCTION_NAME
   Else
      Debug.Print "b_i_p_RegisterSplitCountNeededCells()" & s_i_p_ERROR_TEXT_REGISTER & sFUNCTION_NAME
   End If
   Resume Finally
End Function


