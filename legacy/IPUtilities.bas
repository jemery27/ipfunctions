Attribute VB_Name = "IPUtilities"
Option Explicit
Option Compare Text
Option Private Module

  'Custom data type for undoing
    Type SaveRange
        Val As Variant
        addr As String
        NumberFormat As String
    End Type
    
'   Stores info about current selection
    Public OldWorkbook As Workbook
    Public OldSheet As Worksheet
    Public OldSelection() As SaveRange

Public Declare PtrSafe Function RIBBONHANDLER Lib "User32" (control As IRibbonControl)
'Handle user clicks on the ribbon
    Select Case control.ID
        Case "ipSort"
            Call IPSORT
        Case "ipFill"
            Call IPFILL
        Case "ipFillSubnet"
            Call IPFILLSUBNET
        Case "ipValues"
            Call IPVALUES
        Case "ipHelp"
            Call SHOWHELP
    End Select

End Function

Public Declare PtrSafe Function IPFILL Lib "User32" ()
'   fill a range with consecutive IP's
    Dim i As Long
    
'   Abort if a range isn't selected
    If TypeName(Selection) <> "Range" Then Exit Function
    
'   Abort if too many cells are selected
    If Selection.Count > 50000 Then
        MsgBox "You selected too many cells.", vbCritical
        Exit Function
    End If
    
'   save range for undo
    Call SaveForUndo
    
    Application.ScreenUpdating = False
    
    If Selection.Count = 1 Then     'if a single cell is selected then fill the subnet
        For i = 2 To (IPDD2DEC(IPBROADCAST(Selection.Value)) - IPDD2DEC(Selection.Value) + 1) 'old - IPHOSTS(Selection.Value)
            Selection.Cells(i, 1) = IPADD(Selection.Cells(1, 1), i - 1)
        Next
    ElseIf Selection.Rows.Count > 1 Then    'if more than one row is selected then fill rows
        For i = 2 To Selection.Rows.Count
            Selection.Cells(i, 1) = IPADD(Selection.Cells(1, 1), i - 1)
        Next
    Else
        For i = 2 To Selection.Columns.Count    'otherwise fill cells across
            Selection.Cells(1, i) = IPADD(Selection.Cells(1, 1), i - 1)
        Next
    End If

'   Specify the Undo Sub
    Application.OnUndo "Undo the IP Fill Command", "UndoIPUtilities"
End Function


Public Declare PtrSafe Function IPFILLSUBNET Lib "User32" ()
'   fill a range with consecutive subnets
    Dim i As Long
    
'   Abort if a range isn't selected
    If TypeName(Selection) <> "Range" Then Exit Function
    
'   Abort if too many cells are selected
    If Selection.Count > 50000 Then
        MsgBox "You selected too many cells.", vbCritical
        Exit Function
    End If
    
'   save range for undo
    Call SaveForUndo
    
    Application.ScreenUpdating = False

    If Selection.Rows.Count > 1 Then
        For i = 2 To Selection.Rows.Count
            Selection.Cells(i, 1) = IPNEXTNET(Selection.Cells(i - 1, 1))
        Next
    Else
        For i = 2 To Selection.Columns.Count
            Selection.Cells(1, i) = IPNEXTNET(Selection.Cells(1, i - 1))
        Next
    End If
    
'   Specify the Undo Sub
    Application.OnUndo "Undo the IP Fill Command", "UndoIPUtilities"
End Function



Public Declare PtrSafe Function IPSORT Lib "User32" ()
'   Sorts a selection of ip addresses in ip order
    Dim i As Long
    'Dim myrange As Range
    
'   Abort if a range isn't selected
    If TypeName(Selection) <> "Range" Then Exit Function
    
'   Abort if too many cells are selected
    If Selection.Count > 50000 Then
        MsgBox "You selected too many cells.", vbCritical
        Exit Function
    End If
    
'   save range for undo
    Call SaveForUndo
    
    Application.ScreenUpdating = False
    
    For i = 1 To Selection.Rows.Count         'convert DD to decimal
        Selection.Cells(i, 1) = IPDD2DEC(Selection.Cells(i, 1))
    Next
    
    Selection.Sort Key1:=Selection              'sort
    
    For i = 1 To Selection.Rows.Count         'convert back to DD
        Selection.Cells(i, 1) = IPDEC2DD(Selection.Cells(i, 1))
    Next
    
'   Specify the Undo Sub
    Application.OnUndo "Undo the IP Sort Command", "UndoIPUtilities"
'   need to figure out how to keep the masks if possible
End Function


Public Declare PtrSafe Function IPVALUES Lib "User32" ()
'   convert formulas to string values
    Dim strValue As String
    Dim cell As Range
    
'   Abort if a range isn't selected
    If TypeName(Selection) <> "Range" Then Exit Function
    
'   Abort if too many cells are selected
    If Selection.Count > 50000 Then
        MsgBox "You selected too many cells.", vbCritical
        Exit Function
    End If
    
'   save range for undo
    Call SaveForUndo
    
    Application.ScreenUpdating = False
    
'   Replace each cell with a formula with it's string value
    For Each cell In Selection
        If cell.HasFormula Then
            strValue = CStr(cell.Value)
            cell.NumberFormat = "@"
            cell.Value = strValue
        End If
    Next
    
'   Specify the Undo Sub
    Application.OnUndo "Undo the IP Convert Command", "UndoIPUtilities"
End Function

Sub SaveForUndo()
'saves the range for undo
    Dim i As Long
    Dim cell As Range

'   The next block of statements
'   Save the current values for undoing
    ReDim OldSelection(Selection.Count)
    Set OldWorkbook = ActiveWorkbook
    Set OldSheet = ActiveSheet
    i = 0
    For Each cell In Selection
        i = i + 1
        OldSelection(i).addr = cell.Address
        OldSelection(i).Val = cell.Formula
        OldSelection(i).NumberFormat = cell.NumberFormat
    Next cell
            
End Sub


Sub UndoIPUtilities()
'   Undoes the effect of the ZeroRange sub
    Dim i As Long
    
'   Tell user if a problem occurs
    On Error GoTo Problem

    Application.ScreenUpdating = False
    
'   Make sure the correct workbook and sheet are active
    OldWorkbook.Activate
    OldSheet.Activate
    
'   Restore the saved information
    For i = 1 To UBound(OldSelection)
        Range(OldSelection(i).addr).NumberFormat = OldSelection(i).NumberFormat
        Range(OldSelection(i).addr).Formula = OldSelection(i).Val
    Next i
    Exit Sub

'   Error handler
Problem:
    MsgBox "Can't undo", vbCritical
End Sub

Sub SHOWHELP()
'Show the help file
    Application.Help ThisWorkbook.Path & "\ip functions help.chm"
End Sub
