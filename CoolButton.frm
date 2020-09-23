VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cool Button"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   2310
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Click Me"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Cool Button"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   430
      TabIndex        =   3
      Top             =   0
      Width           =   1440
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Example"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   600
      TabIndex        =   2
      Top             =   240
      Width           =   1125
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Click Me"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1155
      Width           =   1575
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000000&
      Visible         =   0   'False
      X1              =   1920
      X2              =   1920
      Y1              =   1080
      Y2              =   1440
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      Visible         =   0   'False
      X1              =   360
      X2              =   1920
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      Visible         =   0   'False
      X1              =   1920
      X2              =   360
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   360
      X2              =   360
      Y1              =   1080
      Y2              =   1440
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
'Normal button
MsgBox "Cool Button Example"
End Sub

Private Sub Form_Load()
'Set the colors of the lines to look like border
'Also can be done in the Properties Window
Line1.BorderColor = &HE0E0E0
Line2.BorderColor = &HE0E0E0
Line3.BorderColor = &H808080
Line4.BorderColor = &H808080
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Hide the the borders of the button or lines when the mouse
'is not over the label
Line1.Visible = False
Line2.Visible = False
Line3.Visible = False
Line4.Visible = False
End Sub

Private Sub Label1_Click()
'Test click
MsgBox "Cool Button Example"
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Inverts the colors of the lines to simulate a pushed button
Line1.BorderColor = &H808080
Line2.BorderColor = &H808080
Line3.BorderColor = &HE0E0E0
Line4.BorderColor = &HE0E0E0
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Show the border of the button or lines when the mouse
'is over the label
Line1.Visible = True
Line2.Visible = True
Line3.Visible = True
Line4.Visible = True
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Set colors of lines back to default colors
Call Form_Load
End Sub
