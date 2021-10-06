VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7455
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame fraMainFrame 
      Height          =   4470
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   7380
      Begin VB.PictureBox Picture1 
         Height          =   1980
         Left            =   240
         Picture         =   "frmSplash.frx":162F72
         ScaleHeight     =   1920
         ScaleWidth      =   1920
         TabIndex        =   8
         Top             =   840
         Width           =   1980
      End
      Begin VB.Label lblStatus2 
         Alignment       =   2  'Center
         Caption         =   "File: "
         Height          =   255
         Left            =   720
         TabIndex        =   9
         Top             =   4080
         Width           =   5895
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         Caption         =   "Gathering Data - Please Wait"
         Height          =   255
         Left            =   720
         TabIndex        =   7
         Top             =   3840
         Width           =   5895
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "Create Daily Oven Report from Omega RD9900"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   2520
         TabIndex        =   6
         Tag             =   "Product"
         Top             =   1200
         Width           =   3915
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "Norlake Manufacturing Company"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2505
         TabIndex        =   5
         Tag             =   "CompanyProduct"
         Top             =   765
         Width           =   4650
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6915
         TabIndex        =   4
         Tag             =   "Platform"
         Top             =   2400
         Width           =   90
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6240
         TabIndex        =   3
         Tag             =   "Version"
         Top             =   2760
         Width           =   885
      End
      Begin VB.Label lblCompany 
         Alignment       =   1  'Right Justify
         Caption         =   "2021 Norlake Manufacturing Company"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   2
         Tag             =   "Company"
         Top             =   3360
         Width           =   2895
      End
      Begin VB.Label lblCopyright 
         Alignment       =   1  'Right Justify
         Caption         =   "Copyright"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   1
         Tag             =   "Copyright"
         Top             =   3120
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
End Sub
