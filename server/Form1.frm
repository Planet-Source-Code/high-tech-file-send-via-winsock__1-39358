VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2655
      Top             =   1125
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      Height          =   240
      Left            =   1260
      TabIndex        =   1
      Top             =   810
      Width           =   465
   End
   Begin VB.Label Label1 
      Caption         =   "received:           bytes"
      Height          =   195
      Left            =   585
      TabIndex        =   0
      Top             =   855
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
Private Sub Form_Load()
Winsock1.LocalPort = 123
Winsock1.Listen
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
Winsock1.Close
Winsock1.Accept requestID

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim Data As String
Winsock1.GetData Data
If Left(Data, 3) = "BOF" Then
Open App.Path & "\" & Right(Data, Len(Data) - 3) For Binary As #1
ElseIf Data = "EOF" Then
Close #1
MsgBox "File Done."
Else
Put #1, , Data
Label2 = Label2 + 4000
Winsock1.SendData "CON"
End If
End Sub

