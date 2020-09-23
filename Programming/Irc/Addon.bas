Attribute VB_Name = "Addon"
Private Declare Function inet_addr Lib "wsock32.dll" (ByVal cp As String) As Long
Private Declare Function htonl Lib "wsock32.dll" (ByVal hostlong As Long) As Long

Function IrcGetLongIP(ByVal AscIp$) As String
    On Error GoTo IrcGetLongIpError:              ' If there is an error go down
    Dim inn&                                      ' Declare variable
    inn = htonl(inet_addr(AscIp))                 ' Use api to take the asc ip and convert to long
    If inn < 0 Then                               ' Self Explanitory
        IrcGetLongIP = CVar(inn + 4294967296#)    ' Return a CVar of the number + the given
    Else                                          ' Self Explanitory
        IrcGetLongIP = CVar(inn)                  ' Just retturn the CVar of inn
    End If                                        ' Self Explanitory
    Exit Function                                 ' Exit Function
IrcGetLongIpError:                                ' Label from above
    IrcGetLongIP = "0"                            ' There was an Error Return 0
    Exit Function                                 ' Exit Function
    Resume                                        ' Resume
End Function




