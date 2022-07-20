Attribute VB_Name = "IPOpenClose"
Option Explicit
Option Private Module

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: This routine is run every time the application is
'           opened. It registers all UDFs in the workbook.
'
Public Sub Auto_Open()
    ' Register the UDFs.
    HandleFunctionRegistration True
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: This routine is run every time the application is
'           closed. It unregisters all UDFs in the workbook.
'
Public Sub Auto_Close()
    ' Unregister the UDFs
    HandleFunctionRegistration False
End Sub

Private Sub Workbook_AddInInstall()
'run when addin is first installed
    MsgBox "IP Functions AddIn Installed Successfully."
End Sub

Private Sub SetOptions()
'   Set options for the IPADD function
    Application.MacroOptions Macro:="IPADD", _
        Description:="Adds a number to an IP Address", _
        Category:=16, _
        HelpContextID:=6, _
        HelpFile:=ThisWorkbook.Path & "\ip functions help.chm"

End Sub
