VERSION 5.00
Begin VB.MDIForm menu 
   BackColor       =   &H8000000C&
   Caption         =   "MENU UTAMA"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9960
   LinkTopic       =   "MDIForm1"
   Picture         =   "menu.frx":0000
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   9960
      TabIndex        =   0
      Top             =   0
      Width           =   9960
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "KELUAR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "CARI DATA"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "INPUT DATA"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         Height          =   615
         Left            =   3960
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   1815
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         Height          =   615
         Left            =   2040
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   1815
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         FillColor       =   &H00E0E0E0&
         Height          =   615
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   1815
      End
   End
   Begin VB.Menu inp 
      Caption         =   "&INPUT DATA"
   End
   Begin VB.Menu car 
      Caption         =   "&CARI DATA"
   End
   Begin VB.Menu cetak 
      Caption         =   "CE&TAK DATA"
   End
   Begin VB.Menu us 
      Caption         =   "&USER"
   End
   Begin VB.Menu keluar 
      Caption         =   "&KELUAR"
   End
End
Attribute VB_Name = "menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub car_Click()
cari.Show
End Sub

Private Sub cetak_Click()
karyawan.Show
End Sub

Private Sub Command1_Click()
input_data.Show
End Sub

Private Sub Command2_Click()
cari.Show
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub inp_Click()
input_data.Show
End Sub

Private Sub keluar_Click()
End
End Sub

Private Sub MDIForm_Load()
inp.Enabled = False
car.Enabled = False
cetak.Enabled = False
us.Enabled = False
keluar.Enabled = False
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
End Sub

Private Sub us_Click()
user.Show
End Sub
