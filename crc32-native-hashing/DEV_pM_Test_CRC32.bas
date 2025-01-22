Attribute VB_Name = "DEV_pM_Test_CRC32"
Option Explicit
Option Private Module

Private Sub mTestCRC32()
   Dim sHash As String
   Dim sHash2 As String
   Dim sHash3 As String
   sHash = sCRC32Hash("Hello World!")
   sHash2 = sCRC32Hash("Hello World!")
   sHash3 = sCRC32Hash("I am different")
   Debug.Print "Reproduce hash value: " & (sHash = sHash2)
   Debug.Print "Different seed, different hash: " & (sHash <> sHash3)
End Sub

