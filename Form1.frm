VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CBExtender 1.1"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5970
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   5970
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "CBExtender Options"
      Height          =   3195
      Left            =   3180
      TabIndex        =   2
      Top             =   150
      Width           =   2685
      Begin VB.CheckBox chkProps 
         Caption         =   "Show Common Folders"
         Height          =   285
         Index           =   5
         Left            =   150
         TabIndex        =   8
         Top             =   2070
         Width           =   2145
      End
      Begin VB.CheckBox chkProps 
         Caption         =   "Auto Select"
         Height          =   285
         Index           =   1
         Left            =   150
         TabIndex        =   7
         Top             =   690
         Value           =   1  'Checked
         Width           =   1185
      End
      Begin VB.CheckBox chkProps 
         Caption         =   "Restrict Items"
         Height          =   285
         Index           =   4
         Left            =   150
         TabIndex        =   6
         Top             =   1710
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox chkProps 
         Caption         =   "List Directories"
         Height          =   285
         Index           =   3
         Left            =   150
         TabIndex        =   5
         Top             =   1350
         Value           =   1  'Checked
         Width           =   1425
      End
      Begin VB.CheckBox chkProps 
         Caption         =   "Auto Size List"
         Height          =   285
         Index           =   2
         Left            =   150
         TabIndex        =   4
         Top             =   1020
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox chkProps 
         Caption         =   "Auto Complete"
         Height          =   285
         Index           =   0
         Left            =   150
         TabIndex        =   3
         Top             =   330
         Width           =   1425
      End
   End
   Begin VB.PictureBox pic16 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3780
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   570
      Visible         =   0   'False
      Width           =   240
   End
   Begin MSComctlLib.ImageList iml16 
      Left            =   3570
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483633
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   "dft"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageCombo ImageCombo1 
      Height          =   360
      Left            =   120
      TabIndex        =   9
      Top             =   240
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   635
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblInfo 
      Caption         =   "- Double Click a Drive to see the first level folders"
      Height          =   435
      Index           =   0
      Left            =   150
      TabIndex        =   11
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label lblInfo 
      Caption         =   "- Double Click the Browse Icon to get a directory dialog"
      Height          =   435
      Index           =   1
      Left            =   150
      TabIndex        =   10
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Label lblAction 
      AutoSize        =   -1  'True
      Caption         =   "Path:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   1
      Top             =   3330
      Width           =   420
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mExtender    As clsCBExtender
Attribute mExtender.VB_VarHelpID = -1


Private Sub chkProps_Click(Index As Integer)

    With mExtender
        Select Case Index
        '/* auto complete
        Case 0
            .p_AutoComplete = IIf(chkProps(0).Value = 1, True, False)
        '/* auto select
        Case 1
            .p_AutoSelect = IIf(chkProps(1).Value = 1, True, False)
        '/* auto size
        Case 2
            .p_AutoSize = IIf(chkProps(2).Value = 1, True, False)
        '/* list directories
        Case 3
            .p_DirectoryList = IIf(chkProps(3).Value = 1, True, False)
        '/* restrict user selection
        Case 4
            .p_RestrictItems = IIf(chkProps(4).Value = 1, True, False)
        '/* show common folders
        Case 5
            .ListDrives IIf(chkProps(5).Value = 1, 1, 0)
        End Select
    End With
    
End Sub

Private Sub Form_Load()

    '/* init
    Set mExtender = New clsCBExtender
    'Set ImageCombo1.ImageList = iml16
    With mExtender
        .SetReference ImageCombo1
        .p_RestrictItems = False
        Set .p_ImageList = iml16
        Set .p_Picture = pic16
        '/* select display detail
        '.ListDrives Special_Folders
        .ListDrives Drives_Only
        '/* auto select entry
        .p_AutoSelect = True
        '/* autosize list to entries (wd)
        .p_AutoSize = True
        '/* enable directory listing
        .p_DirectoryList = True
        '/* deny user entry
        .p_RestrictItems = True
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set mExtender = Nothing

End Sub

Private Sub mExtender_eFilePath(sPath As String)

    lblAction.Caption = ""
    lblAction.Caption = sPath
    lblAction.Refresh
    
End Sub
