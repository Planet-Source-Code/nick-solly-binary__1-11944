VERSION 5.00
Begin VB.Form frmBinary 
   Caption         =   "Binary Tester"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5700
   Icon            =   "frmBinary.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   5700
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdEnter 
      Caption         =   "Enter"
      Default         =   -1  'True
      Height          =   735
      Left            =   1200
      TabIndex        =   10
      Top             =   3000
      Width           =   2895
   End
   Begin VB.TextBox txtno 
      Height          =   285
      Left            =   2520
      TabIndex        =   8
      Top             =   2520
      Width           =   3015
   End
   Begin VB.Label lblAsc 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   14
      Top             =   1920
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label11 
      Caption         =   "Asc. Figure:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   13
      Top             =   1920
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblbinary 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   12
      Top             =   1200
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label10 
      Caption         =   "Binary No:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   11
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Enter a number between 1-255:"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   2520
      Width           =   2415
   End
   Begin VB.Image Img1on 
      Height          =   480
      Left            =   5160
      Picture         =   "frmBinary.frx":0442
      Top             =   360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Img2on 
      Height          =   480
      Left            =   4440
      Picture         =   "frmBinary.frx":0884
      Top             =   360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Img4on 
      Height          =   480
      Left            =   3720
      Picture         =   "frmBinary.frx":0CC6
      Top             =   360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Img8on 
      Height          =   480
      Left            =   3000
      Picture         =   "frmBinary.frx":1108
      Top             =   360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Img16on 
      Height          =   480
      Left            =   2280
      Picture         =   "frmBinary.frx":154A
      Top             =   360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Img32on 
      Height          =   480
      Left            =   1560
      Picture         =   "frmBinary.frx":198C
      Top             =   360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Img64on 
      Height          =   480
      Left            =   840
      Picture         =   "frmBinary.frx":1DCE
      Top             =   360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Img128on 
      Height          =   480
      Left            =   120
      Picture         =   "frmBinary.frx":2210
      Top             =   360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label8 
      Caption         =   "1"
      Height          =   255
      Left            =   5280
      TabIndex        =   7
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "2"
      Height          =   255
      Left            =   4560
      TabIndex        =   6
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "4"
      Height          =   255
      Left            =   3840
      TabIndex        =   5
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "8"
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "16"
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "32"
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "64"
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "128"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   375
   End
   Begin VB.Image Img1off 
      Height          =   480
      Left            =   5160
      Picture         =   "frmBinary.frx":2652
      Top             =   360
      Width           =   480
   End
   Begin VB.Image Img2off 
      Height          =   480
      Left            =   4440
      Picture         =   "frmBinary.frx":2A94
      Top             =   360
      Width           =   480
   End
   Begin VB.Image Img4off 
      Height          =   480
      Left            =   3720
      Picture         =   "frmBinary.frx":2ED6
      Top             =   360
      Width           =   480
   End
   Begin VB.Image Img8off 
      Height          =   480
      Left            =   3000
      Picture         =   "frmBinary.frx":3318
      Top             =   360
      Width           =   480
   End
   Begin VB.Image Img16off 
      Height          =   480
      Left            =   2280
      Picture         =   "frmBinary.frx":375A
      Top             =   360
      Width           =   480
   End
   Begin VB.Image Img32off 
      Height          =   480
      Left            =   1560
      Picture         =   "frmBinary.frx":3B9C
      Top             =   360
      Width           =   480
   End
   Begin VB.Image Img64off 
      Height          =   480
      Left            =   840
      Picture         =   "frmBinary.frx":3FDE
      Top             =   360
      Width           =   480
   End
   Begin VB.Image Img128off 
      Height          =   480
      Left            =   120
      Picture         =   "frmBinary.frx":4420
      Top             =   360
      Width           =   480
   End
End
Attribute VB_Name = "frmBinary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdEnter_Click()
Dim intpress As Integer
Dim intNo As Integer
Dim strBinary As String

If txtno = "" Or Val(txtno) > 255 Or Val(txtno) <= 0 Then
    intpress = MsgBox("Please enter a valid number", vbOKOnly, "Error")
    txtno = ""
Else
intNo = Val(txtno)
If intNo - 128 >= 0 Then
    Img128on.Visible = True
    intNo = intNo - 128
End If
If intNo - 64 >= 0 Then
    Img64on.Visible = True
    intNo = intNo - 64
End If
If intNo - 32 >= 0 Then
    Img32on.Visible = True
    intNo = intNo - 32
End If
If intNo - 16 >= 0 Then
    Img16on.Visible = True
    intNo = intNo - 16
End If
If intNo - 8 >= 0 Then
    Img8on.Visible = True
    intNo = intNo - 8
End If
If intNo - 4 >= 0 Then
    Img4on.Visible = True
    intNo = intNo - 4
End If
If intNo - 2 >= 0 Then
    Img2on.Visible = True
    intNo = intNo - 2
End If
If intNo - 1 >= 0 Then
    Img1on.Visible = True
    intNo = intno1
End If
End If
strBinary = ""
If Img128on.Visible = True Then
    strBinary = strBinary & "1"
    Else
    strBinary = strBinary & "0"
End If
If Img64on.Visible = True Then
    strBinary = strBinary & "1"
Else
    strBinary = strBinary & "0"
End If
If Img32on.Visible = True Then
    strBinary = strBinary & "1"
Else
    strBinary = strBinary & "0"
End If
If Img16on.Visible = True Then
    strBinary = strBinary & "1"
Else
    strBinary = strBinary & "0"
End If
If Img8on.Visible = True Then
    strBinary = strBinary & "1"
Else
    strBinary = strBinary & "0"
End If
If Img4on.Visible = True Then
    strBinary = strBinary & "1"
Else
    strBinary = strBinary & "0"
End If
If Img2on.Visible = True Then
    strBinary = strBinary & "1"
Else
    strBinary = strBinary & "0"
End If
If Img1on.Visible = True Then
    strBinary = strBinary & "1"
Else
    strBinary = strBinary & "0"
End If
Label10.Visible = True
lblbinary.Visible = True
Label11.Visible = True
lblAsc.Visible = True
lblAsc.Caption = Chr(txtno)
lblbinary = strBinary
Label11.Visible = True
lblAsc.Visible = True
txtno.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub txtno_Change()
Img128on.Visible = False
Img64on.Visible = False
Img32on.Visible = False
Img16on.Visible = False
Img8on.Visible = False
Img4on.Visible = False
Img2on.Visible = False
Img1on.Visible = False
strBinary = ""
Label10.Visible = False
lblbinary = ""
lblbinary.Visible = False
Label11.Visible = False
lblAsc = ""
lblAsc.Visible = False
End Sub
