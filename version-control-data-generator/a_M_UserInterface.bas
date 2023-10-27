Attribute VB_Name = "a_M_UserInterface"
Option Explicit

' "Public" is used temporarily so that the button can be set to run this sub
' conveniently - please note that in a public module like this one the code
' will still be run when the sub is "Private" when clicking the button

'Public Sub click_ExportDataForVersionControl()
Private Sub click_ExportDataForVersionControl()
   DEV_i_pM_VCDG_EntryLevel.DEV_i_p_ExportDataForVersionControl
End Sub
