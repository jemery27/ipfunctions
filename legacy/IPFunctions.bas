Attribute VB_Name = "IPFunctions"
'Option Explicit
'Option Compare Text
'Option Private Module

'Dim WorkRange As Range
'Dim UndoError As Boolean, HaltOperation As Boolean

Public Function IPNETWORK(ip As String) As String
'returns the network address ip an ip address/subnet string
    Dim octet As Integer
    Dim mask As String
    
    mask = IPMASKVAL(ip)
    
    For octet = 1 To 4
        IPNETWORK = IPNETWORK & (IPOCTET(ip, octet) And IPOCTET(mask, octet)) & "."
    Next
    
    IPNETWORK = Left(IPNETWORK, Len(IPNETWORK) - 1)

End Function

Public Function IPBROADCAST(ip As String) As String
'returns the broadcast address of an ip address/subnet string

    Dim octet As Integer
    Dim masklen As String, bcast As String, IPNET As String
    
    masklen = Right(ip, Len(ip) - InStr(1, ip, "/"))
    
    ip = IPDD2BIN(ip)
    
    IPBROADCAST = IPBIN2DD(Left(ip, masklen) & WorksheetFunction.Rept("1", 32 - masklen))
    
End Function

Public Function IPISIN(ip1 As String, ip2 As String) As Boolean
'check to see if ip1 is contained in the network of ip2
    Dim ip1masklen As String, ip2masklen
    Dim ip2net As String, ip1net As String
    
'    Dim ip1 As String, ip2 As String
'    ip1 = "10.20.30.40"
'    ip2 = "10.20.30.12/24"
    
    ip2masklen = Right(ip2, Len(ip2) - InStr(1, ip2, "/"))
    ip2net = IPDD2BIN(IPNETWORK(ip2))
    
    If InStr(1, ip1, "/") Then                      'if ip1 has a mask
        ip1masklen = Right(ip1, Len(ip1) - InStr(1, ip1, "/"))
        ip1net = IPDD2BIN(IPNETWORK(ip1))           'then find it's network address
    Else
        ip1masklen = 32
        ip1net = IPDD2BIN(IPNETWORK(ip2 & "/32"))    'else it's network is it's ip/32
    End If
    
    If ip2masklen < ip1masklen Then
        If Left(ip2net, ip2masklen) = Left(ip1net, ip2masklen) Then
            IPISIN = True
        Else
            IPISIN = False
        End If
    Else
        IPISIN = False
    End If
    
End Function

Public Function IPMASKVAL(ip As String) As String
'returns the dotted decimal notation of a mask length
    Dim mask
    
    mask = Right(ip, Len(ip) - InStr(1, ip, "/"))
    
    With WorksheetFunction
        IPMASKVAL = .Rept(1, mask) & .Rept(0, 32 - mask)
    End With
    
    IPMASKVAL = IPBIN2DD(IPMASKVAL)
    
End Function

Public Function IPMASKWILD(ip As String) As String
'returns the dotted decimal wildcard notation of a mask length
    Dim mask
    
    mask = Right(ip, Len(ip) - InStr(1, ip, "/"))
    
    With WorksheetFunction
        IPMASKWILD = .Rept(0, mask) & .Rept(1, 32 - mask)
    End With
    
    IPMASKWILD = IPBIN2DD(IPMASKWILD)
End Function

Public Function IPDD2BIN(ip As String) As String
'convert IP Address from dotted decimal to decimal
    Dim octet As Integer
    Dim IPHEX As String

    For octet = 1 To 4
        IPHEX = Right("00" & CStr(Hex(IPOCTET(ip, octet))), 2)
        IPDD2BIN = IPDD2BIN & Application.WorksheetFunction.Hex2Bin(IPHEX, 8)
    Next
  
End Function

Public Function IPBIN2DD(ip As String) As String
'convert from binary to dotted decimal notation
    
    With Application.WorksheetFunction
    If Len(ip) < 32 Then
        ip = .Rept(0, Len(ip) - 32) & ip
    ElseIf Len(ip) > 32 Then
        ip = Right(ip, 32)
    End If
    
    IPBIN2DD = .Bin2Dec(Left(ip, 8)) & "." & .Bin2Dec(Mid(ip, 9, 8)) & "." _
                & .Bin2Dec(Mid(ip, 17, 8)) & "." & .Bin2Dec(Right(ip, 8))
    
    End With
End Function

Public Function IPDD2DEC(ip As String) As String
'convert IP Address from dotted decimal to decimal
    Dim octet As Integer
    Dim IPHEX As String

    For octet = 1 To 4
        IPHEX = IPHEX & Right("00" & CStr(Hex(IPOCTET(ip, octet))), 2)
    Next
    
    IPDD2DEC = IPDD2DEC & Application.WorksheetFunction.Hex2Dec(IPHEX)
    
End Function

Public Function IPDD2HEX(ip As String) As String
'convert IP Address from dotted decimal to decimal
    Dim octet As Integer
    Dim IPHEX As String

    For octet = 1 To 4
        IPDD2HEX = IPDD2HEX & Right("00" & CStr(Hex(IPOCTET(ip, octet))), 2)
    Next
    
End Function

Public Function IPISNETWORK(ip As String) As Boolean
'returns true if the ip address is the correct network address for the mask
    If IPADDR(ip) = IPNETWORK(ip) Then
        IPISNETWORK = True
    Else
        IPISNETWORK = False
    End If

End Function

Public Function IPADD(ip As String, x As Double) As String
'add a number to an ip address
    ip = IPDD2DEC(IPADDR(ip))
    IPADD = IPDEC2DD(ip + x)
End Function

Public Function IPHOSTS(ip As String, Optional net As Boolean = True) As Double
'returns the number of host ip's in the subnet
    Dim masklen As Integer
    
    masklen = Right(ip, Len(ip) - InStr(1, ip, "/"))

    IPHOSTS = 2 ^ (32 - masklen)
    
    If Not net Then
        IPHOSTS = IPHOSTS - 2
    End If
    
End Function

Public Function IPNEXTNET(ip As String) As String
'returns the next subnet of the same size
    Dim masklen As Integer, mask As String
    
    mask = Right(ip, Len(ip) - InStr(1, ip, "/"))
    
    IPNEXTNET = IPDEC2DD(IPDD2DEC(IPNETWORK(ip)) + IPHOSTS(ip)) & _
        "/" & mask
End Function


Public Function IPDEC2DD(ip As String) As String
'convert IP address from decimal number to dotted decimal
    Dim IPHEX As String
    Dim octet As Integer
    
    IPDEC2DD = ""

    With Application.WorksheetFunction
    
    IPHEX = Right("00000000" & .Dec2Hex(ip), 8)         'pad the hex with 00's for 8 chars
    
    For octet = 1 To 4                                  'convert each hex octet to dec
        IPDEC2DD = IPDEC2DD & .Hex2Dec(Left(IPHEX, 2)) & "."
        IPHEX = Right(IPHEX, Len(IPHEX) - 2)
    Next
    
    IPDEC2DD = Left(IPDEC2DD, Len(IPDEC2DD) - 1)        'trim the trailing dot
       
    End With
End Function



Public Function IPOCTET(ip As String, octet As Integer) As String
'returns the specified octet from an IP address

    ip = IPADDR(ip)
    
    If octet < 0 Or octet > 4 Or Not IPVALID(ip) Then
        MsgBox octet
        IPOCTET = -1
        Return
    End If

    With Application.WorksheetFunction
    
    Select Case octet
        Case 1
            IPOCTET = Left(ip, .Find(".", ip) - 1)
        Case 2 To 3
            IPOCTET = Mid(ip, .Find("@", .Substitute(ip, ".", "@", octet - 1)) + 1, _
                .Find("#", .Substitute(ip, ".", "#", octet)) - .Find("@", _
                .Substitute(ip, ".", "@", octet - 1)) - 1)
        Case 4
            IPOCTET = Right(ip, .Find(".", StrReverse(ip)) - 1)
    End Select

    End With
End Function


Public Function IPVALID(ip As String) As Boolean
'returns true if valid IP Address
    Dim regex As Object
    
    Set regex = New RegExp
    regex.Pattern = "^([0-9]{1,3}\.){3}[0-9]{1,3}(\/([0-9]|[1-2][0-9]|3[0-2]))?$"

    IPVALID = regex.Test(ip)
    
End Function


Public Function IPADDR(ip As String, Optional octets As Integer) As String
'returns the ip address of an ip/mask

    If IPVALID(ip) And InStr(1, ip, "/") <> 0 Then
        IPADDR = Left(ip, InStr(1, ip, "/") - 1)
    Else
        IPADDR = ip
    End If
    
    If 1 <= octets And octets <= 3 Then
        IPADDR = Left(IPADDR, InStr(1, WorksheetFunction.Substitute(IPADDR, ".", "$", octets), "$"))
    End If
    
End Function

Public Function IP2DD(ip As String, mask As String) As String
'returns the DD/mask notation from ip and mask values
    Dim masklen As Integer
    Dim maskbin As String
        
    maskbin = IPDD2BIN(mask)
    
    If Left(maskbin, 1) = 1 Then
        masklen = 32 - Len(Replace(maskbin, "1", ""))
    Else
        masklen = 32 - Len(Replace(maskbin, "0", ""))
    End If
    
    If IPDD2DEC(mask) = 0 Then masklen = 0
    If IPDD2HEX(mask) = "FFFFFFFF" Then masklen = 32
    
    IP2DD = ip & "/" & CStr(masklen)
    
End Function

Public Function IPCLASS(ip As String) As String
'returns the class of an IP Address
    Dim ipbin As String
    
    ipbin = IPDD2BIN(ip)
    
    If Left(ipbin, 1) = "0" Then
        IPCLASS = "A"
    ElseIf Left(ipbin, 2) = "10" Then IPCLASS = "B"
    ElseIf Left(ipbin, 3) = "110" Then IPCLASS = "C"
    ElseIf Left(ipbin, 4) = "1110" Then IPCLASS = "D"
    Else: IPCLASS = "E"
    End If
    
End Function

Public Function IPRANGE(ip As String, Optional sep As String = "-") As String
'returns a string of the start-end ip addresses of the subnet
    Dim myIP As String
    myIP = ip

    IPRANGE = CStr(IPNETWORK(myIP)) & sep & CStr(IPBROADCAST(ip))

End Function


