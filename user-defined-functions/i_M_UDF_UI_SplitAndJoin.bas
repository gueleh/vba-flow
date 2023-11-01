Attribute VB_Name = "i_M_UDF_UI_SplitAndJoin"
Option Explicit

Public Function s_i_g_Join _
( _
   ByRef o_Range As Range, _
   ByVal s_Delimiter As String, _
   Optional ByVal b_SkipEmptyCells As Boolean = False, _
   Optional ByVal b_IsVolatileFunction As Boolean = True _
) As String
Attribute s_i_g_Join.VB_Description = "Concatenates the cell contents of a range, delimited by a provided delimiter, optionally skipping empty cells, optionally being not a volatile function."
Attribute s_i_g_Join.VB_ProcData.VB_Invoke_Func = " \n19"

   Dim sReturnValue As String
   Dim oRngCell As Range
   
   If b_IsVolatileFunction Then Application.Volatile
   On Error GoTo Catch
   For Each oRngCell In o_Range
      If Not b_SkipEmptyCells _
      Or (b_SkipEmptyCells And Len(oRngCell.Value2) > 0) _
      Then
         sReturnValue = sReturnValue + CStr(oRngCell.Value2) + s_Delimiter
      End If
   Next oRngCell
   sReturnValue = Left$(sReturnValue, Len(sReturnValue) - Len(s_Delimiter))
   
Finally:
   On Error Resume Next
   s_i_g_Join = sReturnValue
   Exit Function
Catch:
   sReturnValue = s_i_p_ERROR_TEXT
   Resume Finally
   
End Function

Public Function sa_i_g_Split _
( _
   ByRef s_StringToSplit As String, _
   ByVal s_Delimiter As String, _
   Optional ByVal b_TransposeAndPrintToRows As Boolean = False, _
   Optional ByVal b_IsVolatileFunction As Boolean = True _
)
Attribute sa_i_g_Split.VB_Description = "Takes a seed string with elements that are separated by a delimiter, splits it into its elements and returns the elements in an array formula . Defaults to being volatile, this can be turned off."
Attribute sa_i_g_Split.VB_ProcData.VB_Invoke_Func = " \n19"

   Dim saValues() As String
   Dim vaReturnValues() As Variant
   Dim lIndex As Long
   
   If b_IsVolatileFunction Then Application.Volatile
   On Error GoTo Catch
      
   saValues = Split(s_StringToSplit, s_Delimiter)
   
   If b_TransposeAndPrintToRows Then
      ReDim vaReturnValues(1 To (UBound(saValues) + 1), 1 To 1)
   Else
      ReDim vaReturnValues(1 To 1, 1 To (UBound(saValues) + 1))
   End If
   
   For lIndex = 0 To UBound(saValues)
      If b_TransposeAndPrintToRows Then
         vaReturnValues(lIndex + 1, 1) = saValues(lIndex)
      Else
         vaReturnValues(1, lIndex + 1) = saValues(lIndex)
      End If
   Next lIndex
      
   
Finally:
   On Error Resume Next
   sa_i_g_Split = vaReturnValues
   Exit Function
Catch:
   ReDim vaReturnValues(1 To 1, 1 To 1)
   vaReturnValues(1, 1) = s_i_p_ERROR_TEXT
   Resume Finally
   
End Function

Public Function s_i_g_SplitCountNeededCells _
( _
   ByRef s_StringToSplit As String, _
   ByVal s_Delimiter As String, _
   Optional ByVal b_IsVolatileFunction As Boolean = True _
) As String
Attribute s_i_g_SplitCountNeededCells.VB_Description = "Takes a seed string with elements that are separated by a delimiter, calculates how many cells will be needed when using the function sa_i_g_Split on it."
Attribute s_i_g_SplitCountNeededCells.VB_ProcData.VB_Invoke_Func = " \n19"

   Dim saValues() As String
   Dim sReturnValue As String
   
   If b_IsVolatileFunction Then Application.Volatile
   On Error GoTo Catch
      
   saValues = Split(s_StringToSplit, s_Delimiter)
   
   sReturnValue = UBound(saValues) + 1
      
   
Finally:
   On Error Resume Next
   s_i_g_SplitCountNeededCells = sReturnValue
   Exit Function
Catch:
   sReturnValue = s_i_p_ERROR_TEXT
   Resume Finally
   
End Function


