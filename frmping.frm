VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pinger - By - Joker - |P|h|r|o|z|e|n| Entertainment"
   ClientHeight    =   360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   360
   ScaleWidth      =   3135
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      MaxLength       =   24
      TabIndex        =   1
      Top             =   0
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ping"
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim ECHO As ICMP_ECHO_REPLY
   Dim pos As Integer
   
  'ping an ip address, passing the
  'address and the ECHO structure
   Call Ping(Text1.Text, ECHO)
   
  'display the results from the ECHO structure
   Form1.Print GetStatusCode(ECHO.status)
   Form1.Print ECHO.Address
   Form1.Print ECHO.RoundTripTime & " ms"
   Form1.Print ECHO.DataSize & " bytes"
   
   If Left$(ECHO.Data, 1) <> Chr$(0) Then
      pos = InStr(ECHO.Data, Chr$(0))
      Form1.Print Left$(ECHO.Data, pos - 1)
   End If

   Form1.Print ECHO.DataPointer
End Sub

Private Sub Label1_Click()
End Sub

Private Sub Form_Load()

End Sub
