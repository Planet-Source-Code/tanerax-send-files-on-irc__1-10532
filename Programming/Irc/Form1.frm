VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9075
   LinkTopic       =   "Form1"
   ScaleHeight     =   8445
   ScaleWidth      =   9075
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtTo 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3840
      TabIndex        =   14
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox TxtSize 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   1920
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Insert"
      Height          =   375
      Left            =   3360
      TabIndex        =   11
      Top             =   2640
      Width           =   2415
   End
   Begin VB.TextBox TxtIP 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3000
      TabIndex        =   9
      Top             =   1560
      Width           =   3135
   End
   Begin VB.TextBox TxtFileName 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2640
      TabIndex        =   7
      Top             =   1200
      Width           =   3735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock DCC 
      Left            =   7680
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Send Text"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   4560
      Width           =   8895
      Begin VB.CommandButton Command1 
         Caption         =   "SEND"
         Height          =   255
         Left            =   7920
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   7575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Temporary Log DATA"
      Height          =   3135
      Left            =   0
      TabIndex        =   1
      Top             =   5280
      Width           =   9015
      Begin VB.TextBox Text1 
         Height          =   2775
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   240
         Width           =   8775
      End
   End
   Begin MSWinsockLib.Winsock IDent 
      Left            =   8160
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton CmdConnect 
      Caption         =   "Connect"
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   0
      Width           =   5295
   End
   Begin MSWinsockLib.Winsock Irc 
      Left            =   8640
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4200
      Top             =   6360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label4 
      Caption         =   "To:"
      Height          =   255
      Left            =   3480
      TabIndex        =   15
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "FileSize:"
      Height          =   255
      Left            =   2520
      TabIndex        =   13
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "IP Address"
      Height          =   255
      Left            =   2040
      TabIndex        =   10
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "FileName:"
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   1200
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// Programming By: Tanerax [Tanerax@nbnet.nb.ca] [50298342]
'// Program: DCC Sending v1.0
'// Comments: Complete
'// This is a working example of filesending. This example connects to
'// irc for the purpose of sending a file. Includes joining server, ident
'// server, ctcp for sending a file and also sending a file
'// code is easy to understand and is completly commented.
'// ********************************************************
'// If You Use This At All In A Program Please Mention Me
'// If You Can Optimize This Please Send Me A Optimized Version
'// Via E-Mail
'// Thank You.

Dim Connected As Boolean
Dim IDentName As String
Sub Connect()
    Irc.Close                                             '// Close Preveious Connect
    Irc.Connect "hebron.in.us.dal.net", 6667              '// Connect to Dal.net
    Connected = True                                      '// Make Connect = true
End Sub

Private Sub CmdConnect_Click()
    Connect                                               '// Run Connect sub
End Sub

Private Sub Command1_Click()
     Irc.SendData Text2.Text & vbCrLf                     '// Send Any Message in Text2
End Sub

Private Sub Command2_Click()
    Irc.Close                                             '// Close Irc Connection
    IDent.Close                                           '// Close Ident
    DCC.Close                                             '// Close DCC
    End                                                   '// End
End Sub

Private Sub Command3_Click()
    Text2.Text = "Privmsg " & TxtTo.Text & " :DCC SEND " & TxtFileName.Text & " " & IrcGetLongIP(TxtIP.Text) & " 1560 " & TxtSize.Text & "" '// Set Text Parameters
End Sub

Private Sub DCC_ConnectionRequest(ByVal requestID As Long)
    Dim Retval As Boolean                          '// Make RetVal a Boolean
    If DCC.State <> sckClosed Then DCC.Close       '// If The State Is Not Close Close It
    DCC.Accept requestID                           '// Accept The Request ID
    Retval = SendFile(TxtFileName.Tag)             '// Begin FileSend
End Sub

Private Sub Form_Load()
    Connected = False                             '// Preset Connected to False
    IDentName = "Test"                            '// Set Ident Server
    IDent.LocalPort = 113                         '// Set Ident Port
    IDent.Listen                                  '// Set Ident Server to Listen
    DCC.LocalPort = 1560                          '// Sets Port to 1560
    DCC.Listen                                    '// Sets Winsock To Listen
End Sub

Private Sub IDent_ConnectionRequest(ByVal requestID As Long)
    IDent.Close                                        '// Close Ident Server
    IDent.Accept requestID                             '// Accept Request
    IDent.SendData "113, 133:USERID:UNIX:" & IDentName '// Send Information
End Sub

Private Sub Irc_Connect()
    Irc.SendData "USER Test.net a a a a" & vbCrLf                 '// Send USER
    Irc.SendData "NICK SendTest" & vbCrLf                         '// Send Nick
    Form1.Caption = "Status [ Hebron.In.Us.Dal.Net ] [ SendTest ]" '// Change Caption
End Sub

Private Sub Irc_DataArrival(ByVal bytesTotal As Long)
    Dim temp As String
    Irc.GetData temp, vbString                             '// Get Incoming
    Text1.Text = Text1.Text & temp & vbNewLine             '// Write to Text Box
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbEnter Then                     '// On the enter key
        Command1_Click                            '// Run the Command1_click
    Else
End Sub

Function SendFile(sFile As String) As Boolean
Dim I, fLength, ret                               '// Declare Variables
Dim Buffer As String                              '// Declare Buffer
Dim bSize As Long                                 '// Declare BufferSize
bSize = 1024                                      '// Set BufferSize
I = FreeFile                                      '// Set I As FreeFile
    Open sFile For Binary Access Read As I        '// Open File For Binary Read
    fLength = LOF(I)                              '// Gets The File Length
        Do Until EOF(I)                           '// Begin A Loop Until EOF
            If fLength - Loc(I) <= bSize Then     '// If The Buffer Is Larger Than
                bSize = fLength - Loc(I)          '// The Rest Of the File. Make The
            End If                                '// New Buffer Size The Rest Of The
                                                  '// File
            If bSize = 0 Then Exit Do             '// If Buffer Size Is 0 Send Done
            bytesent = bytesent + bSize           '// Adds The Buffer To Bytes Sent
            Buffer = Space$(bSize)                '// Get The Buffer From The BlockSize
            Get I, , Buffer                       '// Take Block From File
            DoEvents                              '// DoEvnets
            ret = DoEvents()                      '// Do Events Return
            DCC.SendData Buffer                   '// Send Block
        Loop                                      '// Loop
    Close I                                       '// Close File
DCC.Close                                         '// Close The Connection
SendFile = True                                   '// Return A True
End Function

Private Sub TxtFileName_DblClick()
    CommonDialog1.ShowOpen                        '// show open Dialog
    Open CommonDialog1.filename For Input As #1   '// Open the selected file
        TxtSize.Text = LOF(1)                     '// make the text the total filesize
    Close #1                                      '// close the file
    TxtFileName.Text = CommonDialog1.filename     '// make the text the filename
    TxtFileName.Tag = CommonDialog1.filename      '// make the tag the filename as well
End Sub
