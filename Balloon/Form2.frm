VERSION 5.00
Object = "{46AD6921-79B8-11D4-A217-0050046EACC3}#2.0#0"; "Balloon_TIP.ocx"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3165
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   4650
   LinkTopic       =   "Form2"
   ScaleHeight     =   3165
   ScaleWidth      =   4650
   StartUpPosition =   3  'Windows Default
   Begin Balloon_TIP.BalloonTip BalloonTip1 
      Left            =   1080
      Top             =   600
      _ExtentX        =   423
      _ExtentY        =   423
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   975
      Left            =   1920
      TabIndex        =   0
      ToolTipText     =   "Tooltip"
      Top             =   1200
      Width           =   1095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
