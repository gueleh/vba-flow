Attribute VB_Name = "i_M_UDF_UI_CRC32"
Option Explicit

Public Function s_i_g_CRC32HashFromString(ByVal sInput As String) As String
   Dim sReturnValue As String
   On Error GoTo Catch
   
   Application.Volatile
   sReturnValue = sCRC32Hash(sInput)
   
Finally:
   On Error Resume Next
   s_i_g_CRC32HashFromString = sReturnValue
   Exit Function
Catch:
   sReturnValue = "Error"
   Resume Finally
End Function
