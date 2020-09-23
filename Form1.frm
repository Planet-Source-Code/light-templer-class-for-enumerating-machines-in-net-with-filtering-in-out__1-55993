VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Example to class  'Enum machines in network'   by LightTempler"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   6555
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnAllMachines 
      BackColor       =   &H00E0B4A3&
      Caption         =   "Enum DFS"
      Height          =   360
      Index           =   9
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4590
      Width           =   1860
   End
   Begin VB.CommandButton btnAllMachines 
      BackColor       =   &H00E0B4A3&
      Caption         =   "Enum Time Server"
      Height          =   360
      Index           =   8
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4027
      Width           =   1860
   End
   Begin VB.CommandButton btnAllMachines 
      BackColor       =   &H00E0B4A3&
      Caption         =   "Enum SQL Server"
      Height          =   360
      Index           =   7
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3471
      Width           =   1860
   End
   Begin VB.CommandButton btnAllMachines 
      BackColor       =   &H00E0B4A3&
      Caption         =   "Enum Domain Controller"
      Height          =   360
      Index           =   6
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2915
      Width           =   1860
   End
   Begin VB.CommandButton btnAllMachines 
      BackColor       =   &H00E0B4A3&
      Caption         =   "Enum Print Server"
      Height          =   360
      Index           =   5
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2359
      Width           =   1860
   End
   Begin VB.CommandButton btnAllMachines 
      BackColor       =   &H00E0B4A3&
      Caption         =   "Enum Unix Server"
      Height          =   360
      Index           =   4
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1803
      Width           =   1860
   End
   Begin VB.CommandButton btnAllMachines 
      BackColor       =   &H00E0B4A3&
      Caption         =   "Enum terminal server"
      Height          =   360
      Index           =   3
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1247
      Width           =   1860
   End
   Begin VB.CommandButton btnAllMachines 
      BackColor       =   &H00E0B4A3&
      Caption         =   "Enum windows comps"
      Height          =   360
      Index           =   2
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   691
      Width           =   1860
   End
   Begin VB.TextBox txtDomain 
      BackColor       =   &H00FAC5AD&
      Height          =   300
      Left            =   3270
      TabIndex        =   7
      Text            =   "NPG"
      Top             =   1080
      Width           =   2895
   End
   Begin VB.ListBox lbComps 
      BackColor       =   &H00E6FCFD&
      Height          =   2985
      Left            =   2475
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   1995
      Width           =   3690
   End
   Begin VB.TextBox txtFilterOut 
      BackColor       =   &H00FAC5AD&
      Height          =   300
      Left            =   3270
      TabIndex        =   4
      ToolTipText     =   "The OUT Filter is Left-aligned!"
      Top             =   630
      Width           =   2895
   End
   Begin VB.TextBox txtFilterIn 
      BackColor       =   &H00FAC5AD&
      Height          =   300
      Left            =   3270
      TabIndex        =   3
      ToolTipText     =   "Compared with this VB LIKE statement."
      Top             =   195
      Width           =   2895
   End
   Begin VB.CommandButton btnAllMachines 
      BackColor       =   &H00E0B4A3&
      Caption         =   "Enum all machines"
      Height          =   360
      Index           =   1
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   135
      Width           =   1860
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Machines found:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   2595
      TabIndex        =   17
      Top             =   1740
      Width           =   1275
   End
   Begin VB.Label Label 
      BackStyle       =   0  'Transparent
      Caption         =   "... add more to your needs easily."
      Height          =   420
      Left            =   165
      TabIndex        =   16
      Top             =   5055
      Width           =   2370
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Domain"
      Height          =   255
      Index           =   0
      Left            =   2475
      TabIndex        =   6
      Top             =   1110
      Width           =   660
   End
   Begin VB.Label lblFilterOut 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Filter Out"
      Height          =   255
      Left            =   2385
      TabIndex        =   2
      Top             =   645
      Width           =   750
   End
   Begin VB.Label lblFilterIn 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Filter In"
      Height          =   255
      Left            =   2505
      TabIndex        =   1
      Top             =   240
      Width           =   630
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'
'   frmMain.frm
'

' Demo form for clsEnumComps

Option Explicit

Private WithEvents oCOMPS As clsEnumComps
Attribute oCOMPS.VB_VarHelpID = -1
'
'
'

Private Sub btnAllMachines_Click(Index As Integer)
    
    Dim lFoundComps As Long
    
    lbComps.Clear
    Set oCOMPS = New clsEnumComps
    
    With oCOMPS
        .FilterIn = txtFilterIn.Text
        .FilterOut = txtFilterOut.Text
        
        lFoundComps = .EnumComps(txtDomain.Text, Index)
        
        If lFoundComps = -1 Then
            MsgBox "Error enum machines online in net!", vbExclamation, " Abort"
        Else
            MsgBox "Machines matching the filters: " & lFoundComps, vbInformation, " Found:"
        End If
        
    End With
    
End Sub


Private Sub oCOMPS_CompFound(sCompName As String)
    
    lbComps.AddItem sCompName
    
End Sub


Private Sub oCOMPS_Error(sError As String)
    
    MsgBox sError, vbExclamation, " Error enum comps:"
    
End Sub

' #*#
