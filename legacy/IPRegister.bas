Attribute VB_Name = "IPRegister"
'
' Dexcription:  Based on a technique first proposed by Laurent Longre.
'               NOTE: You must modify the LoadFunctionData procedure
'               in order to use this module to register your UDFs.
'
' Authors:      Rob Bovey, www.appspro.com
'               Stephen Bullen, www.oaltd.co.uk
'
'
' Chapter Change Overview
' Ch#   Comment
' --------------------------------------------------------------
' 05    Initial version
'
Option Explicit
Option Private Module

' This can be any loaded Windows DLL that contains exported API functions.
Private Const msMODULE_TEXT As String = """user32.dll"""

' This type structure holds the complete description for a UDF.
Private Type REG_ARGS
    sFuncName As String
    lNumArgs As Long
    sArgNames As String
    sCategory As String
    sFuncDescr As String
    sArgDescr As String
    sProc As String
End Type


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: This procedure is called from Auto_Open to register
'           the UDFs in the function library and by Auto_Close
'           to unregister them.
'
' Arguments:    bRegister       Pass True to register the list
'                               of UDFs in this workbook, False
'                               to unregister them.
'
' Date          Developer       Chap    Action
' --------------------------------------------------------------
' 06/01/08      Rob Bovey       Ch05    Initial version
'
Public Sub HandleFunctionRegistration(ByVal bRegister As Boolean)

    Dim lCount As Long
    Dim auFuncData() As REG_ARGS
    
    ' Load the data for each UDF.
    LoadFunctionData auFuncData()
    
    ' Loop the array of UDF descriptions and register or unregister
    ' them as specified by the bRegister argument.
    For lCount = LBound(auFuncData) To UBound(auFuncData)
        If bRegister Then
            CallRegister auFuncData(lCount)
        Else
            CallUnregister auFuncData(lCount)
        End If
    Next lCount
    
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: This procedure holds the data for all of the UDFs in
'           this workbook. This data is used to register and
'           un-register the functions with the Excel function
'           wizard. You must manually modify this procedure to add
'           one element to the auFuncData() array (and complete all
'           of its information) for each UDF that you want to register.
'
' Arguments:    auFuncData()    Returned by this procedure.
'                               The array holds one completed
'                               REG_ARGS UDT for each UDF that
'                               needs to be registered.
'
' Date          Developer       Chap    Action
' --------------------------------------------------------------
' 06/01/08      Rob Bovey       Ch05    Initial version
'
Private Sub LoadFunctionData(ByRef auFuncData() As REG_ARGS)

    ' You must set this to the number of functions you're registering.
    Const lNUM_FUNCTIONS As Long = 20

    ' This array will hold one user-defined type for each of your UDFs.
    ReDim auFuncData(1 To lNUM_FUNCTIONS)

    ' **********************************************************
    ' IFERROR Function Data.
    ' **********************************************************
    ' You must fill out one user-defined type for each UDF.
'    auFuncData(1).sFuncName = "IFERROR"
'    ' This must be the number of function arguments plus the return value.
'    auFuncData(1).lNumArgs = 3
'    ' A comma-delimited list of argument names
'    auFuncData(1).sArgNames = "ToEvaluate,Default"
'    ' The category that you want this function to appear under in the function wizard.
'    auFuncData(1).sCategory = "SampleUDF"
'    ' The description that will appear in the function wizard for this function.
'    auFuncData(1).sFuncDescr = "A replacement for:" & vbLf & "=IF(ISERROR(<function>),<default>,<function>)"
'    ' The descriptions for each function argument. These must be surrounded in double quotes and comma-delimited
'    auFuncData(1).sArgDescr = """The expression to exaluate."",""The expression to return if ToEvaluate is an error. """
'    ' This must be the name of a unique API function in the library specified by the msMODULE_TEXT constant.
'    ' It can be any API function, but you must use a diferent one for each of your UDFs.
'    auFuncData(1).sProc = "CharNextA"
    
    ' **********************************************************
    ' Additional Function Data.
    ' **********************************************************
    
    auFuncData(1).sFuncName = "IPADD"
    auFuncData(1).lNumArgs = 2
    auFuncData(1).sArgNames = "ip,number"
    auFuncData(1).sCategory = "IP Functions"
    auFuncData(1).sFuncDescr = "Adds a number to an IP address"
    auFuncData(1).sArgDescr = """An ip address in dotted decimal notation"",""Number to add to the IP """
    auFuncData(1).sProc = "CharLowerA"
    
    auFuncData(2).sFuncName = "IPADDR"
    auFuncData(2).lNumArgs = 2
    auFuncData(2).sArgNames = "ip,octets"
    auFuncData(2).sCategory = "IP Functions"
    auFuncData(2).sFuncDescr = "Returns the IP address of an ip/mask"
    auFuncData(2).sArgDescr = """An ip address in dotted decimal notation with a subnet mask /x"",""Number of octets to return """
    auFuncData(2).sProc = "AdjustWindowRect"
    
    auFuncData(3).sFuncName = "IPBIN2DD"
    auFuncData(3).lNumArgs = 1
    auFuncData(3).sArgNames = "ip"
    auFuncData(3).sCategory = "IP Functions"
    auFuncData(3).sFuncDescr = "Convert from 32 bit binary to dotted decimal notation"
    auFuncData(3).sArgDescr = """An ip address in 32 bit binary format """
    auFuncData(3).sProc = "AdjustWindowRectEx"
    
    auFuncData(4).sFuncName = "IPBROADCAST"
    auFuncData(4).lNumArgs = 1
    auFuncData(4).sArgNames = "ip"
    auFuncData(4).sCategory = "IP Functions"
    auFuncData(4).sFuncDescr = "Returns the broadcast address of an ip address/subnet string"
    auFuncData(4).sArgDescr = """An ip address in dotted decimal notation with a subnet mask /x """
    auFuncData(4).sProc = "AlignRects"

    auFuncData(5).sFuncName = "IPDD2BIN"
    auFuncData(5).lNumArgs = 1
    auFuncData(5).sArgNames = "ip"
    auFuncData(5).sCategory = "IP Functions"
    auFuncData(5).sFuncDescr = "Convert from dotted decimal to binary notation"
    auFuncData(5).sArgDescr = """An ip address in dotted decimal notation """
    auFuncData(5).sProc = "AllowForegroundActivation"
    
    auFuncData(6).sFuncName = "IPDD2DEC"
    auFuncData(6).lNumArgs = 1
    auFuncData(6).sArgNames = "ip"
    auFuncData(6).sCategory = "IP Functions"
    auFuncData(6).sFuncDescr = "Convert from dotted decimal to decimal notation"
    auFuncData(6).sArgDescr = """An ip address in dotted decimal notation """
    auFuncData(6).sProc = "AllowSetForegroundWindow"
    
    auFuncData(7).sFuncName = "IPDD2HEX"
    auFuncData(7).lNumArgs = 1
    auFuncData(7).sArgNames = "ip"
    auFuncData(7).sCategory = "IP Functions"
    auFuncData(7).sFuncDescr = "Convert from dotted decimal to Hex notation"
    auFuncData(7).sArgDescr = """An ip address in dotted decimal notation """
    auFuncData(7).sProc = "AnimateWindow"
    
    auFuncData(8).sFuncName = "IPDEC2DD"
    auFuncData(8).lNumArgs = 1
    auFuncData(8).sArgNames = "ip"
    auFuncData(8).sCategory = "IP Functions"
    auFuncData(8).sFuncDescr = "Convert from decimal to dotted decimal notation"
    auFuncData(8).sArgDescr = """An ip address in decimal notation """
    auFuncData(8).sProc = "AnyPopup"
    
    auFuncData(9).sFuncName = "IPHOSTS"
    auFuncData(9).lNumArgs = 2
    auFuncData(9).sArgNames = "ip,include_net"
    auFuncData(9).sCategory = "IP Functions"
    auFuncData(9).sFuncDescr = "Returns the number of host address in the subnet"
    auFuncData(9).sArgDescr = """An ip address in dotted decimal notation"",""TRUE to include network and broadcast addresses in host count """
    auFuncData(9).sProc = "AppendMenuA"

    auFuncData(10).sFuncName = "IPISIN"
    auFuncData(10).lNumArgs = 2
    auFuncData(10).sArgNames = "ip1,ip2"
    auFuncData(10).sCategory = "IP Functions"
    auFuncData(10).sFuncDescr = "Checks to see if ip1 is contained in ip2"
    auFuncData(10).sArgDescr = """An ip address in dotted decimal notation"",""An ip address in dotted decimal notation with subnet /x """
    auFuncData(10).sProc = "AppendMenuW"
    
    auFuncData(11).sFuncName = "IPISNETWORK"
    auFuncData(11).lNumArgs = 1
    auFuncData(11).sArgNames = "ip"
    auFuncData(11).sCategory = "IP Functions"
    auFuncData(11).sFuncDescr = "Checks to see if the ip is the network address with the given mask"
    auFuncData(11).sArgDescr = """An ip address in dotted decimal notation with subnet mask a.b.c.d/x """
    auFuncData(11).sProc = "ArrangeIconicWindows"
    
    auFuncData(12).sFuncName = "IPMASKVAL"
    auFuncData(12).lNumArgs = 1
    auFuncData(12).sArgNames = "ip1"
    auFuncData(12).sCategory = "IP Functions"
    auFuncData(12).sFuncDescr = "Returns the dotted decimal mask notation"
    auFuncData(12).sArgDescr = """An ip address in dotted decimal notation a.b.c.d/x """
    auFuncData(12).sProc = "AttachThreadInput"
    
    auFuncData(13).sFuncName = "IPMASKWILD"
    auFuncData(13).lNumArgs = 1
    auFuncData(13).sArgNames = "ip"
    auFuncData(13).sCategory = "IP Functions"
    auFuncData(13).sFuncDescr = "Returns the dotted decimal wildcard mask of a subnet"
    auFuncData(13).sArgDescr = """An ip address in dotted decimal notation a.b.c.d/x """
    auFuncData(13).sProc = "BeginDeferWindowPos"
    
    auFuncData(14).sFuncName = "IPNETWORK"
    auFuncData(14).lNumArgs = 1
    auFuncData(14).sArgNames = "ip"
    auFuncData(14).sCategory = "IP Functions"
    auFuncData(14).sFuncDescr = "Returns the network address ip an ip address/subnet string"
    auFuncData(14).sArgDescr = """An ip address in dotted decimal notation a.b.c.d/x """
    auFuncData(14).sProc = "BeginPaint"
    
    auFuncData(15).sFuncName = "IPNEXTNET"
    auFuncData(15).lNumArgs = 1
    auFuncData(15).sArgNames = "ip"
    auFuncData(15).sCategory = "IP Functions"
    auFuncData(15).sFuncDescr = "Returns the next subnet of the same size"
    auFuncData(15).sArgDescr = """An ip address in dotted decimal notation a.b.c.d/x """
    auFuncData(15).sProc = "BlockInput"
    
    auFuncData(16).sFuncName = "IPOCTET"
    auFuncData(16).lNumArgs = 2
    auFuncData(16).sArgNames = "ip,octet"
    auFuncData(16).sCategory = "IP Functions"
    auFuncData(16).sFuncDescr = "Returns the specified octet from an IP address"
    auFuncData(16).sArgDescr = """An ip address in dotted decimal notation"",""Which octet to return 1 through 4 """
    auFuncData(16).sProc = "BringWindowToTop"
    
    auFuncData(17).sFuncName = "IPVALID"
    auFuncData(17).lNumArgs = 1
    auFuncData(17).sArgNames = "ip"
    auFuncData(17).sCategory = "IP Functions"
    auFuncData(17).sFuncDescr = "Returns true if valid IP Address"
    auFuncData(17).sArgDescr = """An ip address in dotted decimal notation """
    auFuncData(17).sProc = "BuildReasonArray"
    
    auFuncData(18).sFuncName = "IP2DD"
    auFuncData(18).lNumArgs = 2
    auFuncData(18).sArgNames = "ip,mask"
    auFuncData(18).sCategory = "IP Functions"
    auFuncData(18).sFuncDescr = "Returns the DD/mask notation from ip and mask values"
    auFuncData(18).sArgDescr = """An ip address in dotted decimal notation"",""A mask in dotted decimal notation """
    auFuncData(18).sProc = "CallMsgFilter"
    
    auFuncData(19).sFuncName = "IPCLASS"
    auFuncData(19).lNumArgs = 1
    auFuncData(19).sArgNames = "ip"
    auFuncData(19).sCategory = "IP Functions"
    auFuncData(19).sFuncDescr = "Returns the class of the address"
    auFuncData(19).sArgDescr = """An ip address in dotted decimal notation """
    auFuncData(19).sProc = "CallNextHookEx"
    
    auFuncData(20).sFuncName = "IPRANGE"
    auFuncData(20).lNumArgs = 2
    auFuncData(20).sArgNames = "ip,sep"
    auFuncData(20).sCategory = "IP Functions"
    auFuncData(20).sFuncDescr = "Returns a string of the start-end addresses"
    auFuncData(20).sArgDescr = """An ip address in dotted decimal notation"",""Separator string """
    auFuncData(20).sProc = "CallWindowProcA"

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: This procedure uses the XLM REGISTER function to
'           register the specified UDF with the Excel function
'           wizard. A discussion of the XLM REGISTER function is
'           beyond the scope of this book. You can get detailed
'           information on this function from the Excel 97 SDK on
'           the Microsoft web site. Go to the URL:
'               http://msdn.microsoft.com/library/
'           and look under:
'                   Office Solutions Development
'                       Microsoft Office
'                           Microsoft Office 97
'                               Product Documentation
'                                   Excel
'
' Arguments:    uData           A REG_ARGS UDT that contains a
'                               description of the UDF that will
'                               be registered.
'
' Date          Developer       Chap    Action
' --------------------------------------------------------------
' 06/01/08      Rob Bovey       Ch05    Initial version
'
Private Sub CallRegister(ByRef uData As REG_ARGS)

    Dim szArgString As String
    
    ' Build the REGISTER function string that we'll pass to the ExecuteExcel4Macro function.
    szArgString = "REGISTER(" & msMODULE_TEXT & ",""" & uData.sProc & """,""" & String$(uData.lNumArgs + 1, "P") _
        & """,""" & uData.sFuncName & """,""" & uData.sArgNames & """," & 1 & ",""" & uData.sCategory _
        & """,,,""" & uData.sFuncDescr & """," & uData.sArgDescr & ")"
    
    ' The total length of the argument string passed to the ExecuteExcel4Macro function
    ' cannot exceed 255 characters.
    If Len(szArgString) <= 255 Then
        Application.ExecuteExcel4Macro szArgString
        ' Define the function name as an Excel global defined name.
        szArgString = "SET.NAME(""" & uData.sFuncName & """,0)"
        Application.ExecuteExcel4Macro szArgString
    Else
        MsgBox "The argument string for the " & uData.sFuncName & _
            " function was more than 255 characters long.", vbExclamation, "Error!"
    End If
    
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: This procedure uses the XLM REGISTER and UNREGISTER
'           functions to unregister the specified UDF. A discussion
'           of these functions is beyond the scope of this book.
'           See the comment for the CallRegister function for how
'           to obtain more information.
'
' Arguments:    uData           A REG_ARGS UDT that contains a
'                               description of the UDF that will
'                               be registered.
'
' Date          Developer       Chap    Action
' --------------------------------------------------------------
' 06/01/08      Rob Bovey       Ch05    Initial version
'
Private Sub CallUnregister(ByRef uData As REG_ARGS)

    ' First, re-register the function as a hidden function to remove
    ' its name from the function wizard. This must be done because of
    ' a bug in the XLM UNREGISTER function.
    Application.ExecuteExcel4Macro "REGISTER(" & msMODULE_TEXT & _
        ",""" & uData.sProc & """,""P"",""" & uData.sFuncName & """,,0)"
        
    ' Next, unregister the function.
    Application.ExecuteExcel4Macro "UNREGISTER(" & uData.sFuncName & ")"
    Application.ExecuteExcel4Macro "SET.NAME(""" & uData.sFuncName & """)"
    
End Sub



'ActivateKeyboardLayout AddClipboardFormatListener AdjustWindowRect
'AdjustWindowRectEx AlignRects AllowForegroundActivation AllowSetForegroundWindow
'AnimateWindow AnyPopup AppendMenuA AppendMenuW ArrangeIconicWindows AttachThreadInput
'BeginDeferWindowPos BeginPaint BlockInput BringWindowToTop BroadcastSystemMessage
'BroadcastSystemMessageA BroadcastSystemMessageExA BroadcastSystemMessageExW
'BroadcastSystemMessageW BuildReasonArray CalcMenuBar CalculatePopupWindowPosition
'CallMsgFilter CallMsgFilterA CallMsgFilterW CallNextHookEx CallWindowProcA
'CallWindowProcW CancelShutdown CascadeChildWindows CascadeWindows ChangeClipboardChain
'ChangeDisplaySettingsA ChangeDisplaySettingsExA ChangeDisplaySettingsExW
'ChangeDisplaySettingsW ChangeMenuA ChangeMenuW ChangeWindowMessageFilter
'ChangeWindowMessageFilterEx CharLowerA CharLowerBuffA CharLowerBuffW CharLowerW
'CharNextA CharNextExA CharNextW CharPrevA CharPrevExA CharPrevW CharToOemA
'CharToOemBuffA CharToOemBuffW CharToOemW CharUpperA CharUpperBuffA CharUpperBuffW Char
