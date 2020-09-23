VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Send'er"
   ClientHeight    =   2790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   ScaleHeight     =   2790
   ScaleWidth      =   4320
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   180
      TabIndex        =   1
      Text            =   "C:\file\file.exe"
      Top             =   225
      Width           =   3390
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send File"
      Height          =   285
      Left            =   180
      TabIndex        =   0
      Top             =   540
      Width           =   1545
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   360
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      Height          =   240
      Left            =   1575
      TabIndex        =   3
      Top             =   2115
      Width           =   420
   End
   Begin VB.Label Label1 
      Caption         =   "Sent:          bytes"
      Height          =   195
      Left            =   1170
      TabIndex        =   2
      Top             =   2115
      Width           =   1590
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'code by: high tech
'email: techx@mailwire.net
'aim: telnetguru
'website a) http://customsoftware.cjb.net
'website b) http://heyyouvisitme.cjb.net

Private Can_Send_Next_Chunk As Boolean 'highly self explainable
Private Sub Command1_Click()
Dim File_Chunk As String * 4000 'the string that will hold the data from
                                'the file, 4000 bytes (4kb)
'attempt connection
Winsock1.RemoteHost = Winsock1.LocalIP
Winsock1.RemotePort = 123
Winsock1.Connect
'wait for connection
Do
DoEvents
Loop Until Winsock1.State = sckConnected
'send the name of file to the server
FileName = "blah.txt" 'you can add code to send correct filename
'and use BOF as a switch
Winsock1.SendData "BOF" & FileName
DoEvents 'let winsock do its business
Open Text1 For Binary As #1
Do While Not EOF(1)
Get #1, , File_Chunk
Winsock1.SendData File_Chunk
Can_Send_Next_Chunk = False
Do
DoEvents
Loop Until Can_Send_Next_Chunk = True
Label2 = Label2 + 4000
Loop
Close #1
DoEvents
Winsock1.SendData "EOF"
'and thats it, the files been sent
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim Data As String
Winsock1.GetData Data
If Data = "CON" Then
Can_Send_Next_Chunk = True
End If
End Sub

