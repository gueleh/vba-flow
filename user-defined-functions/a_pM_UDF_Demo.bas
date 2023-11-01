Attribute VB_Name = "a_pM_UDF_Demo"
Option Explicit

Public Sub Auto_Open()
   Dim bHasError As Boolean
   If Not i_pM_UDF_Reg_SplitAndJoin.b_i_p_RegisterJoin() Then bHasError = True
   If Not i_pM_UDF_Reg_SplitAndJoin.b_i_p_RegisterSplitCountNeededCells() Then bHasError = True
   If Not i_pM_UDF_Reg_SplitAndJoin.b_i_p_RegisterSplit() Then bHasError = True
   
   If bHasError Then MsgBox "There was at least one error in a_pM_UDF_Demo.Auto_Open(), please check the direct window.", vbCritical
   
End Sub

