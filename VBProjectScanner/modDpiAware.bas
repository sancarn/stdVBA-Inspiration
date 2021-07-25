Attribute VB_Name = "modDpiAware"
Option Explicit

Dim cFormLoader As clsDpiPmFormLoader

Private Sub Main()
    '/// must be activated before 1st form is displayed
    '/// project must begin with Sub Main
    Set cFormLoader = New clsDpiPmFormLoader
    cFormLoader.Activate    ' start hook listener
    
    '/// change to your startup form, as needed
    frmMain.Show
End Sub

