VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Registry File Editor"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7035
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   7035
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "E&xample 2"
      Height          =   375
      Left            =   5640
      TabIndex        =   14
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Example"
      Height          =   375
      Left            =   5640
      TabIndex        =   13
      Top             =   240
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Save in file:"
      Height          =   855
      Left            =   120
      TabIndex        =   10
      Top             =   3480
      Width           =   5295
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Text            =   "C:\Windows\Desktop\Myregistry.reg"
         Top             =   360
         Width           =   3855
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   4200
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Registry options"
      Height          =   3015
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5295
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   240
         TabIndex        =   5
         Text            =   "HKEY_CLASSES_ROOT"
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Top             =   1800
         Width           =   4815
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Top             =   2400
         Width           =   4815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Text            =   "\"
         Top             =   1200
         Width           =   4815
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Key:    (two or more keys you must separated wit this character ""\"")"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   4875
      End
      Begin VB.Label Label3 
         Caption         =   "Value:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Value data:"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   2160
         Width           =   825
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Chose head key"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1155
      End
   End
   Begin VB.TextBox Navodnik 
      Height          =   195
      Left            =   6480
      TabIndex        =   0
      Text            =   """"
      Top             =   6600
      Width           =   150
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Send me e-mail"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   5760
      MousePointer    =   99  'Custom
      TabIndex        =   15
      ToolTipText     =   "Send mail"
      Top             =   2880
      Width           =   1080
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSave_Click()
DeleteFile Text4.Text
FileSave "REGEDIT4", Text4.Text
WriteRegFile Text4.Text, Combo1.Text & Text1.Text, Navodnik.Text & Text2.Text & Navodnik.Text, Navodnik.Text & Text3.Text & Navodnik.Text
MsgBox "Run this file and add information in registry"
End Sub


Private Sub Command1_Click()
Combo1.ListIndex = 1
Text1.Text = "\Software\Toyo"
Text2.Text = "FormBackColor"
Text3.Text = "Red"
End Sub

Private Sub Command4_Click()
End
End Sub



Private Sub Command2_Click()
Combo1.ListIndex = 2
Text1.Text = "\Software\Toyo\Options"
Text2.Text = "FontName"
Text3.Text = "Tahoma"
End Sub

Private Sub Label1_Click()
ShellExecute 0, "open", "mailto:boras@vip.hr ; franjo.boras@sb.tel.hr", "", "", 0
End Sub
