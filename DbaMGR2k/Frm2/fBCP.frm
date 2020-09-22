VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form fBCP 
   Appearance      =   0  'Flat
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Visual BCP"
   ClientHeight    =   3450
   ClientLeft      =   1980
   ClientTop       =   1710
   ClientWidth     =   7380
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fBCP.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   Tag             =   "10"
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Passaggio 3"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2850
      Index           =   2
      Left            =   0
      TabIndex        =   8
      Tag             =   "2004"
      Top             =   0
      Visible         =   0   'False
      Width           =   7155
      Begin VB.Frame fra 
         Caption         =   "File Info"
         Height          =   2775
         Index           =   3
         Left            =   3120
         TabIndex        =   32
         Top             =   0
         Width           =   3975
         Begin VB.CheckBox chk 
            Alignment       =   1  'Right Justify
            Caption         =   "Carica Definizione Formato da Tabella"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   38
            Top             =   1560
            Value           =   1  'Checked
            Width           =   3135
         End
         Begin VB.CommandButton cmd 
            Caption         =   ".."
            Height          =   300
            Index           =   4
            Left            =   3480
            TabIndex        =   41
            Top             =   2400
            Width           =   375
         End
         Begin VB.TextBox tDati 
            Height          =   285
            Index           =   8
            Left            =   120
            TabIndex        =   40
            Text            =   "errormsg.log"
            Top             =   2400
            Width           =   3255
         End
         Begin VB.CommandButton cmd 
            Caption         =   ".."
            Height          =   300
            Index           =   3
            Left            =   3480
            TabIndex        =   37
            Top             =   1200
            Width           =   375
         End
         Begin VB.TextBox tDati 
            Height          =   285
            Index           =   7
            Left            =   120
            TabIndex        =   36
            Text            =   "bcp.fmt"
            Top             =   1200
            Width           =   3255
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   33
            Text            =   "Combo1"
            Top             =   480
            Width           =   2295
         End
         Begin VB.Label lbl 
            Caption         =   "File Errori (Vuoto per nessun File)"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   39
            Top             =   2160
            Width           =   2535
         End
         Begin VB.Label lbl 
            Caption         =   "File Formato"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   35
            Top             =   960
            Width           =   2055
         End
         Begin VB.Label lbl 
            Caption         =   "Formato Input/Output"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame fra 
         Caption         =   "Options"
         Height          =   2775
         Index           =   2
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Width           =   3015
         Begin VB.CheckBox chk 
            Alignment       =   1  'Right Justify
            Caption         =   "Truncate Log on Import"
            Enabled         =   0   'False
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   31
            Top             =   2400
            Width           =   2295
         End
         Begin VB.CheckBox chk 
            Alignment       =   1  'Right Justify
            Caption         =   "Compatibile SQL 6.x"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   30
            Top             =   2040
            Width           =   2295
         End
         Begin VB.CheckBox chk 
            Alignment       =   1  'Right Justify
            Caption         =   "Mantieni valori IDENTITY"
            Enabled         =   0   'False
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   29
            Top             =   1680
            Width           =   2295
         End
         Begin VB.TextBox tDati 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   6
            Left            =   1680
            MaxLength       =   10
            MultiLine       =   -1  'True
            TabIndex        =   28
            Text            =   "fBCP.frx":0442
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox tDati 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   5
            Left            =   1680
            MaxLength       =   10
            MultiLine       =   -1  'True
            TabIndex        =   27
            Text            =   "fBCP.frx":044A
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox tDati 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   1680
            MaxLength       =   10
            MultiLine       =   -1  'True
            TabIndex        =   26
            Text            =   "fBCP.frx":044E
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox tDati 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   1680
            MaxLength       =   10
            MultiLine       =   -1  'True
            TabIndex        =   25
            Text            =   "fBCP.frx":0452
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lbl 
            Caption         =   ".."
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   54
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label lbl 
            Caption         =   ".."
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   53
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label lbl 
            Caption         =   ".."
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   52
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label lbl 
            Caption         =   ".."
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   51
            Top             =   240
            Width           =   1335
         End
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Passaggio 4"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2800
      Index           =   4
      Left            =   0
      TabIndex        =   9
      Tag             =   "2006"
      Top             =   0
      Visible         =   0   'False
      Width           =   7155
      Begin RichTextLib.RichTextBox rtf 
         Height          =   2415
         Left            =   120
         TabIndex        =   47
         Top             =   360
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   4260
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"fBCP.frx":0457
      End
      Begin VB.Label lbl 
         Caption         =   "Stato Trasferimento"
         Height          =   255
         Index           =   9
         Left            =   480
         TabIndex        =   48
         Top             =   120
         Width           =   2055
      End
   End
   Begin VB.PictureBox picNav 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   7380
      TabIndex        =   0
      Top             =   2880
      Width           =   7380
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Fine"
         Height          =   360
         Index           =   4
         Left            =   5910
         MaskColor       =   &H00000000&
         TabIndex        =   4
         Tag             =   "104"
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "Ava&nti >"
         Height          =   360
         Index           =   3
         Left            =   4455
         MaskColor       =   &H00000000&
         TabIndex        =   3
         Tag             =   "103"
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "< &Indietro"
         Height          =   360
         Index           =   2
         Left            =   3360
         MaskColor       =   &H00000000&
         TabIndex        =   2
         Tag             =   "102"
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Annulla"
         Height          =   360
         Index           =   1
         Left            =   1920
         MaskColor       =   &H00000000&
         TabIndex        =   1
         Tag             =   "101"
         Top             =   120
         Width           =   1092
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   108
         X2              =   7012
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   240
         X2              =   7144
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Passaggio 5"
      Enabled         =   0   'False
      Height          =   2800
      Index           =   3
      Left            =   0
      TabIndex        =   42
      Top             =   0
      Visible         =   0   'False
      Width           =   7155
      Begin VB.TextBox tFlex 
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   1080
         TabIndex        =   46
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cbo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid Flex 
         Height          =   2055
         Left            =   120
         TabIndex        =   44
         Top             =   0
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   3625
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
      End
      Begin VB.Label lblHint 
         Height          =   495
         Left            =   1200
         TabIndex        =   50
         Top             =   2160
         Width           =   5880
      End
      Begin VB.Label lblStep 
         Caption         =   "Indicazioni:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   120
         TabIndex        =   43
         Top             =   2160
         Width           =   960
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Passaggio 1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2800
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Tag             =   "2000"
      Top             =   0
      Width           =   7155
      Begin VB.Frame fra 
         Caption         =   "Direzione"
         Height          =   2295
         Index           =   1
         Left            =   3960
         TabIndex        =   17
         Top             =   240
         Width           =   2415
         Begin VB.CommandButton cmd 
            Caption         =   "Out"
            Height          =   855
            Index           =   1
            Left            =   120
            Picture         =   "fBCP.frx":04D2
            Style           =   1  'Graphical
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton cmd 
            Caption         =   "In"
            Height          =   855
            Index           =   0
            Left            =   120
            Picture         =   "fBCP.frx":0914
            Style           =   1  'Graphical
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame fra 
         Caption         =   "Connessione"
         Height          =   2415
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   3375
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   0
            Left            =   960
            TabIndex        =   15
            Text            =   "Combo1"
            Top             =   1800
            Width           =   2295
         End
         Begin VB.CheckBox chk 
            Caption         =   "Usa Questa Connessione"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   12
            Top             =   600
            Value           =   1  'Checked
            Width           =   2415
         End
         Begin VB.TextBox tDati 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   960
            MaxLength       =   20
            TabIndex        =   14
            Top             =   1440
            Width           =   2295
         End
         Begin VB.TextBox tDati 
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   960
            MaxLength       =   20
            TabIndex        =   13
            Top             =   1080
            Width           =   2295
         End
         Begin VB.CheckBox chk 
            Caption         =   "Connessione Trusted NT"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label lbl 
            Caption         =   "UserName   Password        Server"
            Height          =   1095
            Index           =   0
            Left            =   120
            TabIndex        =   16
            Top             =   1080
            Width           =   735
         End
      End
   End
   Begin VB.Frame fraStep 
      BorderStyle     =   0  'None
      Caption         =   "Passaggio 2"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2800
      Index           =   1
      Left            =   0
      TabIndex        =   6
      Tag             =   "2002"
      Top             =   0
      Visible         =   0   'False
      Width           =   7155
      Begin VB.CommandButton cmd 
         Caption         =   ".."
         Height          =   300
         Index           =   2
         Left            =   6360
         TabIndex        =   23
         Top             =   2400
         Width           =   375
      End
      Begin VB.TextBox tDati 
         Height          =   285
         Index           =   2
         Left            =   3240
         TabIndex        =   21
         Top             =   2400
         Width           =   3015
      End
      Begin MSComctlLib.TreeView tView 
         Height          =   2775
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   4895
         _Version        =   393217
         Indentation     =   176
         LabelEdit       =   1
         Style           =   7
         Appearance      =   1
      End
      Begin MSComctlLib.ImageList imgDatabase 
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   8
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fBCP.frx":0D56
               Key             =   "OpenFolder"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fBCP.frx":11AA
               Key             =   "Folder"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fBCP.frx":15FE
               Key             =   "Backup"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fBCP.frx":1A52
               Key             =   "Database"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fBCP.frx":1D6E
               Key             =   "ServerGroup"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fBCP.frx":208A
               Key             =   "Server"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fBCP.frx":23A6
               Key             =   "Users"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fBCP.frx":26C2
               Key             =   "Logs"
            EndProperty
         EndProperty
      End
      Begin VB.Label lblTabe 
         Caption         =   "Nome File Dati"
         Height          =   255
         Left            =   3240
         TabIndex        =   49
         Top             =   1800
         Width           =   3615
      End
      Begin VB.Label lbl 
         Caption         =   "Nome File Dati"
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   22
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label lblStep 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Selezione della Tabella"
         ForeColor       =   &H80000008&
         Height          =   1575
         Index           =   2
         Left            =   3240
         TabIndex        =   7
         Tag             =   "2003"
         Top             =   120
         Width           =   3720
      End
   End
End
Attribute VB_Name = "fBCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const NUM_STEPS = 5

Const BTN_CANCEL = 1
Const BTN_BACK = 2
Const BTN_NEXT = 3
Const BTN_FINISH = 4

Const STEP_0 = 0
Const STEP_1 = 1
Const STEP_2 = 2
Const STEP_3 = 3
Const STEP_FINISH = 4

Const DIR_NONE = 0
Const DIR_BACK = 1
Const DIR_NEXT = 2

Const FRM_TITLE = "Visual BCP "

Private mNode As Node
Private WithEvents bcpServer As SQLDMO.BulkCopy
Attribute bcpServer.VB_VarHelpID = -1
Private bBcpIN As Boolean
Private sDbName As String
Private sTB_View As String
Private sFileN As String
Private sFileF As String    'file Formato
Private miRow As Integer, miCol As Integer
Private m_blnAnnullaMod As Boolean
Private m_lngVSBarW  As Long
Private m_lWidth As Long, m_lHeight As Long

Private bcpCol As Collection
Private bFMTDirty As Boolean

'Variabili a livello di modulo
Private mnCurStep       As Integer
Private mbFinishOK      As Boolean

Private Sub bcpServer_BatchImported(ByVal Message As String)
    'rtf.Text = rtf.Text & vbLf & Message
    rtf.Text = rtf.Text & IIf(Len(rtf.Text), vbCrLf, "") & Message
    rtf.SelStart = Len(rtf.Text)
    
End Sub

Private Sub bcpServer_RowsCopied(ByVal Message As String, ByVal Rows As Long)
    
    rtf.Text = rtf.Text & vbLf & MyLoadResString(k_Transferred_Rows) & Format$(Rows, kLNum)
    rtf.SelStart = Len(rtf.Text)
End Sub

Private Sub cbo_GotFocus()
    m_blnAnnullaMod = False
    If Flex.Col = 1 Then lblHint = MyLoadResString(k_Column_Type_modif_if_compatible)
End Sub

Private Sub chk_Click(Index As Integer)
    If Index = 1 Then
        tDati(0).Enabled = 1 - chk(Index).Value
        tDati(1).Enabled = 1 - chk(Index).Value
        Combo1(0).Enabled = 1 - chk(Index).Value
    End If
End Sub

Private Sub cmd_Click(Index As Integer)

    TestActiveControl
    Select Case Index
        Case 0, 1
            bBcpIN = Index = 0
            Me.Caption = FRM_TITLE & "- " & MyLoadResString(IIf(bBcpIN, k_Import, k_Export)) & " Data"
        Case 2  'File di Import/Export
            GetFileName 0
        Case 3
            GetFileName 1
        Case 4
            GetFileName 2
        Case Else
    End Select
    
End Sub

Private Sub cmdNav_Click(Index As Integer)
    Dim nAltStep As Integer
    
    Dim bGO As Boolean
    
    TestActiveControl
    Select Case Index
        Case BTN_CANCEL
            Unload Me
          
        Case BTN_BACK
            'Inserire qui altre istruzioni Case per
            'la gestione di altri passaggi
            If mnCurStep = STEP_FINISH Then
                If Combo1(1).ListIndex <> 0 Then
                    fraStep(mnCurStep).Visible = False
                    'fraStep(mnCurStep).Left = -10000
                    mnCurStep = mnCurStep - 1
                End If
            End If
            
            nAltStep = mnCurStep - 1
            SetStep nAltStep, DIR_BACK
          
        Case BTN_NEXT
            'Inserire qui altre istruzioni Case per
            'la gestione di altri passaggi
            Select Case mnCurStep
                Case STEP_0
                    bGO = True
                Case STEP_1
                    bGO = (Len(sDbName) <> 0) And (Len(sTB_View) <> 0) And (Len(sFileN) <> 0)
                Case Else
                    bGO = True
            End Select
            If bGO Then
                If mnCurStep = STEP_2 Then
                    If Combo1(1).ListIndex <> 0 Then
                        fraStep(mnCurStep).Visible = True
                        'fraStep(mnCurStep).Left = -10000
                        mnCurStep = mnCurStep + 1
                    End If
                End If
                nAltStep = mnCurStep + 1
                
                SetStep nAltStep, DIR_NEXT
            Else
                If mnCurStep > 0 Then IncompleteData mnCurStep
            End If
        Case BTN_FINISH
            'Inserire qui il codice per creare la creazione guidata
            bGO = True
            If Combo1(1).ListIndex = 0 Then 'FormatFile
                If Not bBcpIN Or bFMTDirty Then bGO = CreateFMT()
            End If
            If bGO Then BCP_Execute
    End Select
    
End Sub
Private Function CreateFMT() As Boolean
    
    Dim iFile As Integer, i As Integer
    Dim sBuf As String, sTmp As String
    Dim lErr As Long, sErr As String
    
    iFile = FreeFile
    On Local Error GoTo ErrCreate
    Open sFileF For Output As #iFile
    
    Print #iFile, IIf(chk(3).Value, "6.0", "8.0")
    Print #iFile, Format$(bcpCol.Count)
    For i = 1 To bcpCol.Count
        With bcpCol.Item(i)
            sBuf = sPad(False, Format$(i), 8)
            sTmp = GetBCPType(.oType)
            If Len(sTmp) Then
                sBuf = sBuf & sPad(False, sTmp, 14)
            Else
                Err.Raise kErrBCP_TypeFMT
            End If
            
            
            sBuf = sBuf & sPad(False, IIf(StrComp(sTmp, "sqlvariant", vbTextCompare) = 0, "4", "0"), 8)
            
            If .IsString Then
                sTmp = .oSize2
            ElseIf StrComp(sTmp, "sqlvariant", vbTextCompare) = 0 Then
                sTmp = "0"
            Else
                sTmp = .oSize
            End If
            sBuf = sBuf & sPad(False, sTmp, 8)
            sTmp = Chr$(34) & .Separ & Chr$(34)
            
            sBuf = sBuf & sPad(False, sTmp, 26)
            sBuf = sBuf & sPad(False, Format$(.iPos), 6)
            sBuf = sBuf & sPad(False, .oName, k_MaxColumnLen + 2)
            
            
            'sBuf = sBuf & Space$(10) & """" & """" 'collation
            
            If .IsString() Then
                sBuf = sBuf & .Collation()
            Else
                sBuf = sBuf & """" & """"  'collation
            End If
        End With
        Print #iFile, sBuf
    Next
    
    Close #iFile
    CreateFMT = True
    bFMTDirty = False
    
ExitHere:

    Exit Function
ErrCreate:
    lErr = Err.Number
    sErr = Err.Description
    Close #iFile
    Beep
    MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr, vbOKOnly, MyLoadResString(k_Error_FormatFile)
    Resume ExitHere
    
End Function
Private Function GetBCPType(ByVal sVal As String) As String
    
    Dim sBuf As String
    Dim i As Integer
    
    i = InStr(sVal, "{")
    If i <> 0 Then
        sVal = Mid$(sVal, i + 1)
        sVal = Replace(sVal, "}", "")
    End If
    sVal = Trim$(sVal)
    Select Case UCase$(sVal)
        Case "SQL_VARIANT"
            sBuf = "SQLVARIANT"
        Case "BIGINT"
            sBuf = "SQLBIGINT"
        Case "CHAR", "VARCHAR", "TEXT", "SYSNAME"
            sBuf = "SQLCHAR"
        Case "NCHAR", "NVARCHAR", "NTEXT"
            sBuf = "SQLNCHAR"
        Case "BINARY", "VARBINARY", "IMAGE", "TIMESTAMP"
            sBuf = "SQLBINARY"
        Case "DATETIME"
            sBuf = "SQLDATETIME"
        Case "SMALLDATETIME"
            sBuf = "SQLDATETIM4"
        Case "DECIMAL"
            sBuf = "SQLDECIMAL"
        Case "NUMERIC"
            sBuf = "SQLNUMERIC"
        Case "FLOAT"
            sBuf = "SQLFLT8"
        Case "REAL"
            sBuf = "SQLFLT4"
        Case "INT"
            sBuf = "SQLINT"
        Case "SMALLINT"
            sBuf = "SQLSMALLINT"
        Case "TINYINT"
            sBuf = "SQLTINYINT"
        Case "MONEY"
            sBuf = "SQLMONEY"
        Case "SMALLMONEY"
            sBuf = "SQLMONEY4"
        Case "BIT"
            sBuf = "SQLBIT"
        Case "UNIQUEIDENTIFIER"
            sBuf = "SQLUNIQUEID"
        Case Else
    End Select
    GetBCPType = sBuf
    
End Function
Private Function GetSQLType(ByVal sVal As String) As String
    
    Dim sBuf As String
    Select Case UCase$(Trim$(sVal))
        Case "SQLBIGINT"
            sBuf = "bigint"
        Case "SQLVARIANT"
            sBuf = "sql_variant"
        Case "SQLCHAR"
            sBuf = "char"
        Case "SQLNCHAR"
            sBuf = "nchar"
        Case "SQLBINARY"
            sBuf = "binary"
        Case "SQLDATETIME"
            sBuf = "datetime"
        Case "SQLDATETIM4"
            sBuf = "smalldatetime"
        Case "SQLDECIMAL"
            sBuf = "decimal"
        Case "SQLNUMERIC"
            sBuf = "numeric"
        Case "SQLFLT8"
            sBuf = "float"
        Case "SQLFLT4"
            sBuf = "real"
        Case "SQLINT"
            sBuf = "int"
        Case "SQLSMALLINT"
            sBuf = "smallint"
        Case "SQLTINYINT"
            sBuf = "tinyint"
        Case "SQLMONEY"
            sBuf = "money"
        Case "SQLMONEY4"
            sBuf = "smallmoney"
        Case "SQLBIT"
            sBuf = "bit"
        Case "SQLUNIQUEID"
            sBuf = "uniqueidentifier"
        Case Else
    End Select
    GetSQLType = sBuf
    
End Function

Private Sub BCP_Execute()

    Dim lErr As Long, sErr As String
    Dim sTempFile As String
    Dim lType As Long, lRows As Long
    Dim oSVR As SQLDMO.SQLServer2
    Dim oObj As Object
        
    Screen.MousePointer = vbHourglass
    Select Case Combo1(1).ListIndex
        Case 0
            lType = SQLDMODataFile_UseFormatFile
        Case 1
            lType = SQLDMODataFile_NativeFormat
        Case 2
            lType = SQLDMODataFile_TabDelimitedChar 'SQLDMODataFile_Default
        Case 3
            lType = SQLDMODataFile_CommaDelimitedChar
    End Select
        

    Set bcpServer = New SQLDMO.BulkCopy
    
    rtf.Text = ""
    
    
    With bcpServer
        .DataFilePath = sFileN
        .ErrorFilePath = Trim$(tDati(8).Text)
        .DataFileType = lType
        If lType = SQLDMODataFile_UseFormatFile Then .FormatFilePath = sFileF
        sTempFile = GetTempFile("tmp")
        .LogFilePath = sTempFile
        .MaximumErrorsBeforeAbort = ValInt(tDati(3))
        .FirstRow = ValInt(tDati(4))
        
        If ValInt(tDati(5)) <> 0 Then .LastRow = ValInt(tDati(5))
        If ValInt(tDati(6)) <> 0 Then .ImportRowsPerBatch = ValInt(tDati(6))
        Debug.Print .ImportRowsPerBatch

        .UseBulkCopyOption = True
        
        .Use6xCompatible = chk(3).Value
        If chk(2).Enabled Then .IncludeIdentityValues = chk(2).Value
        If chk(4).Enabled Then .TruncateLog = chk(4).Value
    End With
    
    lErr = ErrConnectBCPServer(oSVR, sErr)
            
    If lErr = 0 Then
        Debug.Print oSVR.EnableBcp = True
        bcpServer.UseExistingConnection = chk(1).Value = vbUnchecked
        
        Set oObj = GetTbByName(oSVR, sDbName, sTB_View)
        If oObj Is Nothing Then
            Set oObj = GetViewByName(oSVR, sDbName, sTB_View)
        End If
        If oObj Is Nothing Then
            lErr = kErr_NotFound           'NOT FOUND
            sErr = ReplaceMsg(MyLoadResString(k_RES_Object_Not_Found), Array("1%", "2%"), Array(sDbName, sTB_View))
        End If
        
'        On Local Error Resume Next
'        Set oObj = oSVR.Databases(sDbName).Tables("[" & sTB_View & "]")
'        lErr = Err.Number
'        If lErr = -2147199728 Then          'NOT FOUND
'            Err.Clear
'            lErr = 0
'            sOwner = GetViewOwner(sDbName, sTB_View)
'            Set oObj = oSVR.Databases(sDbName).Views("[" & sTB_View & "]")
'            lErr = Err.Number
'        End If
'        sErr = Err.Description
'        On Local Error GoTo 0
    End If
    If lErr = 0 Then
        Debug.Print oSVR.EnableBcp = True
        rtf.Text = MyLoadResString(k_Beginning_Data_Transfer) & vbLf & String$(20, "-")
        On Local Error Resume Next
        If bBcpIN Then
            lRows = oObj.ImportData(bcpServer)
        Else
            lRows = oObj.ExportData(bcpServer)
        End If
        lErr = Err.Number
        sErr = Err.Description
        On Local Error GoTo 0
        Debug.Print sErr
        rtf.Text = rtf.Text & vbLf & MyLoadResString(k_Transferred_Rows) & Format$(lRows, kLNum)
        ReadFromFile sTempFile
        
    End If
    Screen.MousePointer = vbDefault
    
    If lErr <> 0 Then
        Beep
        rtf.Text = rtf.Text & vbLf & MyLoadResString(kMsgBoxError) & ": " & lErr & vbTab & sErr
    Else
        mbFinishOK = False
        SetNavBtns mnCurStep
    End If
    rtf.SelStart = Len(rtf.Text)
    
    Set bcpServer = Nothing
    Set oSVR = Nothing
    Set oObj = Nothing
    
End Sub
Private Sub ReadFromFile(ByVal sTempFile As String)

    Dim numOfBytes As Long
    Dim buf As String
    Dim hFile As Long
    
    On Local Error Resume Next
    hFile = lopen(sTempFile, 0)
    numOfBytes = GetFileSize(hFile, ByVal 0&)
    buf = Space(numOfBytes)
    If Err.Number <> 0 Then
        buf = MyLoadResString(k_Error_reading_System_Handle) & sTempFile _
                    & vbCrLf & Err.Number & "-" & Err.Description
    Else
        lread hFile, buf, numOfBytes
    End If
    ' restituisce l'output
    rtf.Text = rtf.Text & vbLf & buf
    Debug.Print buf
    
    lclose (hFile)
    DeleteFile (sTempFile)
    On Local Error GoTo 0

End Sub

Private Sub Combo1_Click(Index As Integer)
    If Index = 1 Then
        cmd(3).Enabled = Combo1(1).ListIndex = 0
        tDati(7).Enabled = Combo1(1).ListIndex = 0
        chk(5).Enabled = Combo1(1).ListIndex = 0 And bBcpIN
        chk(5).Value = vbChecked
    End If

End Sub

Private Sub Flex_DblClick()
    
    If bcpCol.Count <> 0 Then
        If Flex.MouseRow > 0 Then FlexEdit vbKeySpace
    End If

End Sub
Private Sub Flex_GotFocus()
    TestActiveControl
    cbo.Visible = False
    tFlex.Visible = False
    ShowFlexHint
End Sub
Private Sub Flex_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyF2 Then
        If bcpCol.Count <> 0 Then
            If Flex.Row > 0 Then FlexEdit vbKeySpace
        End If
    End If
End Sub
Private Sub Flex_KeyPress(KeyAscii As Integer)

    Dim iCode As Integer
    
    iCode = KeyAscii
    KeyAscii = 0
    If bcpCol.Count Then
        If iCode <> vbKeyReturn And iCode <> vbKeyEscape Then FlexEdit iCode
    End If

End Sub

Private Sub Flex_RowColChange()

    ShowFlexHint
    
End Sub

Private Sub Flex_Scroll()
    tFlex.Visible = False
    cbo.Visible = False
End Sub
Private Sub ShowFlexHint()
    
    Dim sBuf As String
    Select Case Flex.Col
        Case 0
            sBuf = MyLoadResString(k_Column_Name)
        Case 1
            sBuf = MyLoadResString(k_Column_Type) '(k_Column_Name)
        Case 2
            sBuf = MyLoadResString(k_Actual_Column_Size)
        Case 3
            sBuf = MyLoadResString(k_Column_Size_to_get_only_CHAR)
        Case 4
            sBuf = MyLoadResString(k_Field_Terminator_4_BCP)
        Case 5
            sBuf = MyLoadResString(k_Column_Position_in_current_T_V)
        Case 6
            sBuf = MyLoadResString(k_NewDbSort)
    End Select
    
    lblHint.Caption = sBuf

End Sub
Private Sub Form_Load()
    
    Dim i As Integer
    Dim oApplication As SQLDMO.Application
    Dim oNameList As SQLDMO.NameList
    

    m_lWidth = Me.Width
    m_lHeight = Me.Height
    
    m_lngVSBarW = Screen.TwipsPerPixelX * GetSystemMetrics(SM_CYVTHUMB)

    fra(0).Caption = MyLoadResString(k_Connection)
    fra(1).Caption = MyLoadResString(k_Direction)
    chk(0).Caption = MyLoadResString(k_Trusted_NT)
    chk(1).Caption = MyLoadResString(k_Use_this_connection)
    cmdNav(1).Caption = MyLoadResString(k_Cancel)
    cmdNav(2).Caption = MyLoadResString(k_Previous)
    cmdNav(3).Caption = MyLoadResString(k_Next)
    cmdNav(4).Caption = MyLoadResString(k_Done)
    lblTabe.Caption = MyLoadResString(k_Selected_Object)
    lbl(1).Caption = MyLoadResString(k_Data_File_Name)
    lblStep(2).Caption = MyLoadResString(k_Table_View_selection)
    
    lbl(2).Caption = MyLoadResString(k_Max_Error_Number)
    lbl(3).Caption = MyLoadResString(k_First_Row)
    lbl(4).Caption = MyLoadResString(k_Last_Row)
    lbl(5).Caption = MyLoadResString(k_Batch_Size)
        
    chk(2).Caption = MyLoadResString(k_Preserve_Identity)
    chk(3).Caption = MyLoadResString(k_Compatible_SQL_6x)
    
    lbl(6).Caption = MyLoadResString(k_IO_Format)
    lbl(7).Caption = MyLoadResString(k_Format_File)
    chk(5).Caption = MyLoadResString(k_Load_Data_Definition_Format_from_Table)
    lbl(8).Caption = MyLoadResString(k_Error_File_Empty_for_no_file)
    
    lblStep(1).Caption = MyLoadResString(k_Hint)
    
    lbl(9).Caption = MyLoadResString(k_Transfer_Status)
    
    bBcpIN = True
    Set bcpCol = New Collection
    
    Flex.RowHeightMin = cbo.Height
   
 
'    With cbo
'        .AddItem "bit"
'        .AddItem "binary"
'        .AddItem "char"
'        .AddItem "datetime"
'        .AddItem "decimal"
'        .AddItem "float"
'        .AddItem "image"
'        .AddItem "int"
'        .AddItem "money"
'        .AddItem "nchar"
'        .AddItem "numeric"
'        .AddItem "ntext"
'        .AddItem "nvarchar"
'        .AddItem "real"
'        .AddItem "smalldatetime"
'        .AddItem "smallint"
'        .AddItem "smallmoney"
'        .AddItem "sysname"
'        .AddItem "timestamp"
'        .AddItem "tinyint"
'        .AddItem "text"
'        .AddItem "uniqueidentifier"
'        .AddItem "varbinary"
'        .AddItem "varchar"
'    End With
    
    Combo1(0).Clear
    Combo1(0).Enabled = False
    Combo1(1).Clear
    With Combo1(1)
        .AddItem MyLoadResString(k_Custom) '"Custom"
        .AddItem MyLoadResString(k_Native)  '"Native"
        .AddItem MyLoadResString(k_Character)  '"Character"
        .AddItem MyLoadResString(k_CommaDel) '"Comma Delimited"
        .ListIndex = 0
    End With

    Combo1(0).AddItem objServer.Name
    Set oApplication = New Application
    With oApplication
        Set oNameList = .ListAvailableSQLServers
        With oNameList
            For i = 1 To .Count
                Combo1(0).AddItem (oNameList.Item(i))
            Next
        End With
    End With
    Set oApplication = Nothing
    
    Combo1(0).ListIndex = 0
    tView.ImageList = imgDatabase
    'Inizializza tutte le variabili
    mbFinishOK = False
    
    For i = 0 To NUM_STEPS - 1
      fraStep(i).Left = -10000
    Next
    
    'Determina il primo passaggio:
    SetStep 0, DIR_NONE

End Sub

Private Sub SetStep(nStep As Integer, nDirection As Integer)
  
    Dim bVal As Boolean
    Dim i As Integer
    For i = 0 To 4
        fraStep(i).Visible = i = nStep
    Next
    fraStep(nStep).ZOrder
    
    Select Case nStep
        Case STEP_0
            Set bcpCol = New Collection
            Me.Caption = FRM_TITLE & "- " & MyLoadResString(IIf(bBcpIN, k_Import, k_Export)) & " Data"
        Case STEP_1
            Me.Caption = FRM_TITLE & "- " & MyLoadResString(IIf(bBcpIN, k_Import, k_Export)) & " Data"
            If nDirection = DIR_NEXT Then
                Screen.MousePointer = vbHourglass
                lblTabe.Caption = MyLoadResString(k_Selected_Object)
                tDati(2).Text = ""
                sDbName = ""
                sTB_View = ""
                sFileN = ""
                LoadDatabaseNodes
                Screen.MousePointer = vbDefault
            End If
      
        Case STEP_2
            Me.Caption = FRM_TITLE & "- " & MyLoadResString(IIf(bBcpIN, k_Import, k_Export)) & " Data"
            If nDirection = DIR_NEXT Then
                tDati(3).Text = 10
                tDati(4).Text = 1
                tDati(5).Text = 0
                tDati(6).Text = "1.000"

                Screen.MousePointer = vbHourglass
                bVal = Test4Identity()
                
                chk(2).Enabled = bVal
                chk(2).Value = Abs(bVal)
                chk(3).Value = vbUnchecked
                
                chk(4).Enabled = bBcpIN
                chk(4).Value = vbUnchecked
                
                sFileF = ""
                tDati(7).Text = "bcp.fmt"
                Screen.MousePointer = vbDefault
                chk(5).Enabled = Combo1(1).ListIndex = 0 And bBcpIN
                chk(5).Value = vbChecked
                
            End If
            mbFinishOK = False
            
        Case STEP_3
            If nDirection = DIR_NEXT Then
                Screen.MousePointer = vbHourglass
                If LoadColDefin() Then
                    LoadGrid
                Else
                    SetStep nStep - 1, DIR_BACK
                    Exit Sub
                End If
                Screen.MousePointer = vbDefault
            End If
            mbFinishOK = False
      
        Case STEP_FINISH
            Me.Caption = FRM_TITLE & "- " & MyLoadResString(IIf(bBcpIN, k_Import, k_Export)) & " Data - " & MyLoadResString(k_Collected_All_Information)
            mbFinishOK = True
        
    End Select
    
    'Visualizza il passaggio successivo
    fraStep(mnCurStep).Enabled = False
    fraStep(nStep).Left = 0
    If nStep <> mnCurStep Then
        fraStep(mnCurStep).Left = -10000
    End If
    fraStep(nStep).Enabled = True

    SetNavBtns nStep
  
End Sub

Private Sub SetNavBtns(nStep As Integer)
    mnCurStep = nStep
    
    If mnCurStep = 0 Then
        cmdNav(BTN_BACK).Enabled = False
        cmdNav(BTN_NEXT).Enabled = True
    ElseIf mnCurStep = NUM_STEPS - 1 Then
        cmdNav(BTN_NEXT).Enabled = False
        cmdNav(BTN_BACK).Enabled = True
    Else
        cmdNav(BTN_BACK).Enabled = True
        cmdNav(BTN_NEXT).Enabled = True
    End If
    
    If mbFinishOK Then
        cmdNav(BTN_FINISH).Enabled = True
    Else
        cmdNav(BTN_FINISH).Enabled = False
    End If
End Sub

'===========================================================
'Questa routine visualizza un messaggio di errore se
'l'utente non ha inserito i dati sufficienti per continuare.
'===========================================================
Sub IncompleteData(nIndex As Integer)
    
    Dim sTmp As String
    
    Select Case nIndex
        Case STEP_1
            If Len(sDbName) = 0 Then
                sTmp = MyLoadResString(k_Select_Database)
            ElseIf Len(sTB_View) = 0 Then
                sTmp = MyLoadResString(k_Select_Table_or_View)
            ElseIf Len(sFileN) = 0 Then
                sTmp = MyLoadResString(k_Select_File_for) & MyLoadResString(IIf(bBcpIN, k_Import, k_Export))
            End If
        Case Else
    End Select
    Beep
    MsgBox sTmp, vbInformation
End Sub

Private Sub Form_Resize()

    Static b As Boolean
    If Not Me.WindowState = vbMinimized Then
        If Not b Then
            b = True
            TestActiveControl
            On Error Resume Next
            If Me.Width < m_lWidth Then Me.Width = m_lWidth
            If Me.Height <> m_lHeight Then Me.Height = m_lHeight
            SizeControls
            b = False
        End If
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set bcpServer = Nothing
    Set mNode = Nothing
    Set fBCP = Nothing
End Sub

Private Sub LoadDatabaseNodes()
    
    Dim dmoDatabase As New SQLDMO.Database
    Dim dmoTable As SQLDMO.Table
    Dim dmoView As SQLDMO.View
    Dim sDbNode As String
    Dim oColErrDB As New Collection
    Dim i As Integer
    Dim lErr As Long, sErr As String
     
        
    tView.Nodes.Clear
    objServer.DisConnect
    objServer.ReConnect

    For Each dmoDatabase In objServer.Databases     'Do for each Database in the Server
        sDbNode = ""
        On Local Error GoTo ErrorHandle
        sDbNode = dmoDatabase.Name
        Set mNode = tView.Nodes.Add(, , sDbNode)
        mNode.Text = sDbNode
        mNode.Image = "Database"  'Database icon
                    
        Set mNode = tView.Nodes.Add(sDbNode, tvwChild, dmoDatabase.Name & "Tables")
        mNode.Text = "Tables"
        mNode.Image = "Folder"  'Folder icon
        Set dmoTable = New SQLDMO.Table
        For Each dmoTable In dmoDatabase.Tables
            If dmoTable.SystemObject = 0 Then
                Set mNode = tView.Nodes.Add(sDbNode & "Tables", tvwChild)
                mNode.Text = dmoTable.Name
            End If
        Next
        Set dmoTable = Nothing
        
        If Not bBcpIN Then
            Set mNode = tView.Nodes.Add(dmoDatabase.Name, tvwChild, dmoDatabase.Name & "Views")
            mNode.Text = "Views"
            mNode.Image = "Folder"  'Folder icon
            Set dmoView = New SQLDMO.View
            For Each dmoView In dmoDatabase.Views
                If dmoView.SystemObject = 0 Then
                    Set mNode = tView.Nodes.Add(dmoDatabase.Name & "Views", tvwChild)
                    mNode.Text = dmoView.Name
                End If
            Next
Continua:
            Set dmoView = Nothing
            Set dmoTable = Nothing
            Set dmoDatabase = Nothing

        End If
        On Local Error GoTo 0
    Next
    Set dmoView = Nothing
    Set dmoTable = Nothing
    Set dmoDatabase = Nothing
    
    
    If oColErrDB.Count Then
        For i = 1 To oColErrDB.Count
            On Local Error Resume Next
            tView.Nodes.Remove oColErrDB.Item(i)
            On Local Error GoTo 0
        Next
    End If
    Set oColErrDB = Nothing

    Exit Sub
  
ErrorHandle:
    lErr = Err.Number
    sErr = Err.Description
    If Len(sDbNode) <> 0 Then
        oColErrDB.Add sDbNode
        sDbNode = ""
    End If
    
    Debug.Print Err.Number; Err.Description
    If lErr <> kErr_NotValidUser Then 'not a valid user in database 'xxx'.
        MsgBox sErr, vbCritical, MyLoadResString(kMsgBoxError)
    End If
    Resume Continua
    
End Sub
Private Function ErrConnectBCPServer(oSVR As SQLDMO.SQLServer2, ByRef sErr As String) As Long
      
    Dim sServer As String
    Dim sUser As String
    Dim vPwd As Variant
    Dim lErr As Long
    Dim iTimeOut As Integer
    
    If chk(1).Value Then
        Set oSVR = objServer
        oSVR.EnableBcp = True
    Else
        sServer = Combo1(0).Text
        sUser = Trim$(tDati(0).Text)
        vPwd = Trim$(tDati(1).Text)
        If Len(vPwd) = 0 Then vPwd = Null
        
        iTimeOut = Val(Trim$(GetIni("CONNECT", "TIMEOUT", App.Path & "\" & App.EXEName & ".INI")))
        If iTimeOut = 0 Then iTimeOut = 90
        'If StrComp(sServer, "<local>", vbTextCompare) = 0 Then sServer = NomeComputer()
        If StrComp(sServer, NomeComputer(), vbTextCompare) = 0 Then sServer = "(local)"
        
        Screen.MousePointer = vbHourglass
        
        Set oSVR = New SQLDMO.SQLServer2
        oSVR.LoginTimeout = iTimeOut
        
        If chk(0).Value Then objServer.LoginSecure = True
        oSVR.EnableBcp = True
        
        On Error Resume Next
        oSVR.Connect sServer, sUser, vPwd

        lErr = Err.Number
        sErr = Err.Description
        On Local Error GoTo 0
      
        Screen.MousePointer = vbDefault
    End If
    ErrConnectBCPServer = lErr
  
End Function

Private Sub tDati_Change(Index As Integer)
    Select Case Index
        Case 2
            sFileN = Trim$(tDati(Index).Text)
        Case 7
            sFileF = Trim$(tDati(Index).Text)
    End Select
End Sub

Private Sub tDati_GotFocus(Index As Integer)
    tDati(Index).SelStart = 0
End Sub

Private Sub tDati_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index > 2 And Index < 7 Then
        Digit_KDown tDati(Index), KeyCode
    Else
        If KeyCode = vbKeyUP Or KeyCode = vbKeyDOWN Then KeyCode = 0
    End If
End Sub

Private Sub tDati_KeyPress(Index As Integer, KeyAscii As Integer)
    
    Select Case KeyAscii
        Case vbKeyEscape
            KeyAscii = 0
            SendKeys kPREV
        Case vbKeyReturn
            KeyAscii = 0
            SendKeys kNEXT
        Case vbKeyBACK
            If Index > 2 And Index < 7 Then Digit_KPress tDati(Index), KeyAscii
        Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95, 92, 58
            If Index > 2 And Index < 7 Then Digit_KPress tDati(Index), KeyAscii
        Case Else
            KeyAscii = 0
    End Select

End Sub

Private Sub tDati_LostFocus(Index As Integer)
    Dim lVal As Long, lVal2 As Long
    
    If Index = 4 Or Index = 5 Then
        lVal = ValInt(tDati(4).Text)
        lVal2 = ValInt(tDati(5).Text)
        If lVal > lVal2 Then tDati(5).Text = 0
    End If
End Sub

Private Sub tFlex_GotFocus()
    m_blnAnnullaMod = False
End Sub

Private Sub tFlex_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then NoPopUpMenu tFlex
End Sub

Private Sub tView_NodeClick(ByVal Node As MSComctlLib.Node)

    Dim sBuf As String, sOrg As String
    
    On Local Error Resume Next
    sBuf = tView.SelectedItem.Parent.Parent
    On Local Error GoTo 0
    
    If Len(sBuf) Then
        On Local Error Resume Next
        sOrg = tView.SelectedItem
        On Local Error GoTo 0
        If Len(sOrg) Then
            sDbName = sBuf
            sTB_View = sOrg
                        
            lblTabe.Caption = MyLoadResString(k_Selected_Object) & sOrg
        End If
    End If

End Sub
Private Sub GetFileName(ByVal iType As Integer)
    
    Dim lErr As Long, sErr As String
    Dim sBuf As String

    Dim obj As cOpen
    Set obj = New cOpen
    
    Select Case iType
        Case 0
            obj.Title = MyLoadResString(k_Select_File_for) & MyLoadResString(IIf(bBcpIN, k_Import, k_Export))
            obj.Def_Dir() = objServer.Registry.SQLRootPath
            obj.Filter() = MyLoadResString(k_All_Files) & Chr$(0) & "*.*" & Chr$(0) & Chr$(0)
        Case 1
            obj.Title = MyLoadResString(k_Select_File_for) & MyLoadResString(k_Format_File) & MyLoadResString(IIf(bBcpIN, k_Import, k_Export))
            obj.Def_Dir() = objServer.Registry.SQLRootPath
            obj.Filter() = MyLoadResString(k_Select_File_for) & " (*.fmt)" & Chr$(0) & "*.fmt" & Chr$(0) & MyLoadResString(k_All_Files) & Chr$(0) & "*.*" & Chr$(0) & Chr$(0)
        Case 2
            obj.Title = MyLoadResString(k_Select_File_for) & MyLoadResString(k_Error_File)
            obj.Def_Dir() = objServer.Registry.ErrorLogPath
            obj.Filter() = MyLoadResString(k_Error_File) & " (*.log)" & Chr$(0) & "*.log" & Chr$(0) & MyLoadResString(k_All_Files) & Chr$(0) & "*.*" & Chr$(0) & Chr$(0)
    End Select
        
    obj.formHwnd() = Me.hWnd

    Select Case iType
        Case 0, 1
            If bBcpIN Then
                obj.Do_Open
            Else
                obj.Do_Save
            End If
        Case 2
            obj.Do_Save
    End Select

    lErr = obj.ErrCode()
    If lErr = 0 Then
        sBuf = obj.FullPath()
        Select Case iType
            Case 0
                If Len(sBuf) Then
                    sFileN = sBuf
                    tDati(2).Text = sFileN
                End If
            Case 1
                If Len(sBuf) Then
                    sFileF = sBuf
                    tDati(7).Text = sFileF
                End If
            Case 2
                If Len(sBuf) Then tDati(8).Text = sBuf
        End Select
    Else
        If lErr = 1 Then
            lErr = 0
        Else
            sErr = obj.ErrMsg()
        End If
    End If
    Set obj = Nothing
    
    If lErr <> 0 Then
        Beep
        MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr, vbOKOnly, App.EXEName
    End If

End Sub
Private Function Test4Identity() As Boolean

    Dim bVal As Boolean
    Dim i As Integer, lErr As Long
    Dim oObj As Object
    Dim sErr As String
    
    
    Set oObj = GetTbByName(objServer, sDbName, sTB_View)
    If oObj Is Nothing Then
        Set oObj = GetViewByName(objServer, sDbName, sTB_View)
    End If
    If oObj Is Nothing Then
        lErr = kErr_NotFound           'NOT FOUND
        sErr = ReplaceMsg(MyLoadResString(k_RES_Object_Not_Found), Array("1%", "2%"), Array(sDbName, sTB_View))
    End If
    
    If lErr = 0 Then
        With oObj
            For i = 1 To .Columns.Count
                If .Columns(i).Identity Then
                    bVal = True
                    Exit For
                End If
            Next
        End With
    Else
        MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr, vbOKOnly, App.EXEName
    End If
    Set oObj = Nothing
    Test4Identity = bVal
    
End Function
Private Function LoadColDefin() As Boolean

    Dim i As Integer, lErr As Long
    Dim sBuf As String, sTmp As String
    Dim iFile As Integer, iAct As Integer, iRow As Integer, iPos As Integer
    Dim bRet As Boolean
    Dim oObj As Object
    Dim sObj As String
    Dim objCol As cBcpRow
    Dim oType As SQLDMO.SystemDatatype2
    Dim bUDT As Boolean
    Dim sErr As String
    
    bFMTDirty = False
    
    bRet = True
    Set bcpCol = New Collection
    
    If chk(5).Value Then
        sObj = "[" & sTB_View & "]"
        
        Set oObj = GetTbByName(objServer, sDbName, sTB_View)
        If oObj Is Nothing Then
            Set oObj = GetViewByName(objServer, sDbName, sTB_View)
        End If
        If oObj Is Nothing Then
            lErr = kErr_NotFound           'NOT FOUND
        End If
        
        
'        On Local Error Resume Next
'        Set oObj = objServer.Databases(sDbName).Tables(sObj)
'        lErr = Err.Number
'        Debug.Print Err.Description
'        If lErr = -2147199728 Then          'NOT FOUND
'            Err.Clear
'            lErr = 0
'            Set oObj = objServer.Databases(sDbName).Views(sObj)
'            lErr = Err.Number
'        End If
'        On Local Error GoTo 0
        
        If lErr = 0 Then
            With oObj
                For i = 1 To .Columns.Count
                    
                    On Local Error Resume Next
                    Set oType = objServer.Databases(sDbName).SystemDatatypes(.Columns(i).DataType)
                    bUDT = Err.Number <> 0
                    On Local Error GoTo 0
                
                    Set objCol = New cBcpRow
                    objCol.iPos = i
                    objCol.oName = .Columns(i).Name
                    If Not bUDT Then
                        objCol.oType = .Columns(i).DataType
                    Else
                        objCol.oType = .Columns(i).DataType & " {" & objServer.Databases(sDbName).UserDefinedDatatypes(.Columns(i).DataType).BaseType & "}"
                    End If
                    objCol.oSize = .Columns(i).Length
                                        
                    If objCol.IsString Then objCol.oSize2 = .Columns(i).Length
                    objCol.Collation() = .Columns(i).Collation
                    
                    bcpCol.Add objCol
                Next
            End With
            bRet = True
        Else
            On Local Error GoTo ErrTrap
            Err.Raise lErr
            On Local Error GoTo 0
        End If
    Else        'Load From File
        i = 0
        iFile = FreeFile
        On Local Error GoTo ErrTrap
        Open sFileF For Input Shared As #iFile
            
        Do While Not EOF(iFile)
            Line Input #iFile, sBuf
            If iAct = 0 Then
                If chk(3).Value Then
                    If sBuf <> "6.0" Then Err.Raise kErrBCP_NotValidFMT
                Else
                    If sBuf <> "8.0" Then Err.Raise kErrBCP_NotValidFMT
                End If
            ElseIf iAct = 1 Then
                iRow = Val(sBuf)
                If iRow = 0 Then Err.Raise kErrBCP_InvRec
            Else
                i = i + 1
                Set objCol = New cBcpRow
                
                sTmp = ""
                iPos = InStr(UCase$(sBuf), "SQL")
                If iPos Then
                    sTmp = GetSQLType(Mid$(sBuf, iPos, 12))     'Tipo
                    sBuf = Mid$(sBuf, iPos + 22)
                End If
                If Len(sTmp) Then
                    objCol.oType = sTmp
                Else
                    Err.Raise kErrBCP_TypeFMT
                End If
                
                iPos = InStr(sBuf, Chr$(34))                    'Size
                sTmp = Left$(sBuf, iPos - 1)
                objCol.oSize = Val(sTmp)
                If objCol.IsString Then objCol.oSize2 = objCol.oSize
                                    
                
                sTmp = ""                                       'Separator
                iPos = InStr(sBuf, Chr$(34))
                If iPos Then
                    sTmp = Mid$(sBuf, iPos + 1)
                    sBuf = sTmp
                    iPos = InStr(sTmp, Chr$(34))
                    If iPos Then
                        If iPos > 1 Then
                            sTmp = Left$(sTmp, iPos)
                        Else
                            sTmp = ""
                        End If
                        sBuf = Mid$(sBuf, iPos + 1)
                    Else
                        Err.Raise kErrBCP_TypeFMT
                        'sTmp = ""
                    End If
                    If Len(sTmp) Then objCol.Separ = sTmp
                Else
                    Err.Raise kErrBCP_TypeFMT
                End If
                
                sBuf = Trim$(sBuf)                              'Pos
                sTmp = Val(sBuf)
                If Val(sTmp) Then objCol.iPos = Val(sTmp)
                
                iPos = InStr(sBuf, Format$(Val(sTmp)))          'Nome Col
                If iPos Then
                    sBuf = Mid(sBuf, iPos + Len(sTmp))
                    objCol.oName = Trim$(sBuf)
                Else
                    Err.Raise kErrBCP_TypeFMT
                End If
                
                bcpCol.Add objCol
            End If
            iAct = iAct + 1
        Loop
        
        If i <> iRow Then Err.Raise kErrBCP_InvRec
        
        Close #iFile
        
    End If
    
    Me.Caption = FRM_TITLE & "- " & MyLoadResString(IIf(bBcpIN, k_Import, k_Export)) & " Data" & IIf(chk(5).Value, "", MyLoadResString(k_EXTERNAL_Format_File_Definition))
    LoadColDefin = bRet
    
ExitHere:
    Set oType = Nothing
    Set oObj = Nothing
    Set objCol = Nothing

    Exit Function

ErrTrap:
    lErr = Err.Number
    sErr = Err.Description
    Close #iFile
    Screen.MousePointer = vbDefault
    Beep
    MsgBox MyLoadResString(kMsgBoxError) & ": " & lErr & vbLf & sErr, vbOKOnly, MyLoadResString(k_Error_FormatFile)
    Resume ExitHere
        
End Function
Private Sub LoadGrid()

    Dim i As Integer
    Dim sBuf As String
    
    With Flex
        .Redraw = False
        .Cols = 7
        .Rows = 1
        .Clear
        .TextMatrix(0, 0) = MyLoadResString(k_Column_Name)
        .ColWidth(0) = 600
        .TextMatrix(0, 1) = MyLoadResString(k_Type)
        .ColWidth(1) = 1500
        .TextMatrix(0, 2) = MyLoadResString(k_Dimension)
        .ColWidth(2) = 800
        .TextMatrix(0, 3) = MyLoadResString(k_Dimension_to_reach)
        .ColWidth(3) = 800
        .TextMatrix(0, 4) = MyLoadResString(k_Separator) '"Separ."
        .ColWidth(4) = 850
        .TextMatrix(0, 5) = MyLoadResString(k_Object_Sequence) '"Seq."
        .ColWidth(5) = 600
        .TextMatrix(0, 6) = MyLoadResString(k_NewDbSort) 'Collation
        .ColWidth(6) = 1800
        
        
        For i = 1 To bcpCol.Count
            With bcpCol.Item(i)
                sBuf = .oName & vbTab & .oType & vbTab & .oSize & vbTab & .oSize2 & vbTab & .Separ & vbTab & .iPos & vbTab & .Collation()
            End With
            .AddItem sBuf
            .Row = .Rows - 1
            .Col = 1
            Set Flex.CellPicture = LoadResPicture(k_ResCombo, vbResBitmap)
            .CellPictureAlignment = flexAlignRightTop
            
            .Col = .Cols - 1
            Set Flex.CellPicture = LoadResPicture(k_ResCombo, vbResBitmap)
            .CellPictureAlignment = flexAlignRightTop
            .CellBackColor = IIf(bcpCol.Item(i).IsString, vbWindowBackground, vbInactiveBorder)
            sBuf = ""
        Next
        If .Rows = 1 Then .AddItem ""
        .FixedRows = 1
        .Redraw = True
        
        SettaLargFlex Flex, 0, .ColWidth(0)
        .Row = 1
        .Col = 0
    End With
    
End Sub
Private Sub cbo_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim iCode As Integer
    
    If Shift = 0 Then
        Select Case KeyCode
            Case vbKeyRIGHT, vbKeyLEFT
                iCode = KeyCode
                KeyCode = 0
                cbo.Visible = False
                Flex.Col = Flex.Col - (38 - iCode)
        End Select
    End If
End Sub

Private Sub cbo_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case vbKeyEscape
            KeyAscii = 0
            m_blnAnnullaMod = True
            Flex.SetFocus
        Case vbKeyReturn
            KeyAscii = 0
            Flex.SetFocus
    End Select

End Sub

Private Sub cbo_LostFocus()
    
    Dim sBuf As String
    
    If Not m_blnAnnullaMod Then
        With bcpCol.Item(miRow)
            If miCol = 1 Then
                .oType() = cbo.List(cbo.ListIndex)
                sBuf = .oType()
                Flex.Col = Flex.Cols - 1
                Flex.TextMatrix(miRow, Flex.Col) = .Collation()
                Flex.CellBackColor = IIf(.IsString, vbWindowBackground, vbInactiveBorder)
                Flex.Col = 1
            Else
                .Collation() = cbo.List(cbo.ListIndex)
                sBuf = .Collation()
            End If
            Flex.TextMatrix(miRow, miCol) = sBuf
        End With
        bFMTDirty = True
    End If
    
    cbo.Visible = False
    Flex.SetFocus
    'If Not m_blnAnnullaMod Then Flex.Col = 1
    
    m_blnAnnullaMod = True

End Sub

Private Sub FlexEdit(KeyAscii As Integer)
    
    Dim bEdit As Boolean
    Dim sBuf As String
    Dim iLen As Integer, iRow As Integer, iCol As Integer
    Dim iTmpType As iPropG
    Dim blnIsCombo As Boolean
    
    iRow = Flex.Row
    iCol = Flex.Col
    sBuf = Flex.Text
    
    If bcpCol.Count = 0 Then Exit Sub

    Select Case iCol
        Case 0, 2
        Case 1, 6
             blnIsCombo = True
             bEdit = True
        Case 3
            If bcpCol.Item(iRow).IsString Then
                bEdit = True
                iTmpType = am_Long
                iLen = 3
            End If
        Case 4
            bEdit = True
            iTmpType = am_String
            iLen = 15
        Case 5
            bEdit = True
            iTmpType = am_Long
            iLen = 3
    End Select
        
    If Flex.CellBackColor = vbInactiveBorder Then bEdit = False
    
    If bEdit Then
        miRow = iRow
        miCol = iCol
        Select Case KeyAscii
            Case vbKeyEscape, vbKeyReturn
                bEdit = False
        End Select
    End If

    If bEdit Then
        If blnIsCombo Then
            If bcpCol.Count Then
                LoadCboEdit iCol = 1
                
                With Flex
                    cbo.ListIndex = 0
                    cbo.ListIndex = GetItem(.TextMatrix(.Row, .Col), cbo)
                    cbo.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth ', .CellHeight
                    cbo.Visible = True
                    cbo.SetFocus
                End With
            End If
        Else
            tFlex.MaxLength = 0
            tFlex.Text = ""
            tFlex.MaxLength = iLen
            tFlex.Text = sBuf
            Debug.Print iTmpType = am_Long
            tFlex.SelStart = 0
            With Flex
                tFlex.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth, .CellHeight
            End With
            tFlex.Visible = True
            tFlex.ZOrder
            tFlex.SetFocus
        End If
        If (iTmpType And am_String) = am_String Then
            If KeyAscii = vbKeySpace Then KeyAscii = 0
'            If KeyAscii Then
'                If Len(sBuf) Then
'                    sBuf = Chr$(KeyAscii) & Mid$(sBuf, 2, iLen - 2)
'                Else
'                    sBuf = Chr$(KeyAscii)
'                End If
'                iLen = 1
'            Else
'                iLen = 0
'            End If
'            tFlex = sBuf
'            tFlex.SelStart = iLen
'        Else
'            tFlex_KeyPress KeyAscii
        End If
        On Local Error Resume Next
        SendKeys Chr$(KeyAscii)
        On Local Error GoTo 0
    End If

End Sub
Private Sub tFlex_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim iTmpType As iPropG
  
    Select Case miCol
        Case 0, 1
        Case 2, 3, 5
            iTmpType = am_Long
        Case 4
            iTmpType = am_String
    End Select
    
    If KeyCode = vbKeyUP Then
        KeyCode = 0
    ElseIf KeyCode = vbKeyDOWN Then
    Else
        If (iTmpType And am_Dec) = am_Dec Then
            Decimal_kDown tFlex, KeyCode
        ElseIf (iTmpType And am_Long) = am_Long Then
            Digit_KDown tFlex, KeyCode
        'ElseIf (iTmpType And am_Date) = am_Date Then
            'Data_KDown tFlex, KeyCode, Left$(kYEAR, Elem(miRow).iLen)
        '    Data_KDown tFlex, KeyCode, Left$(kYEAR, iLen)
        ElseIf (iTmpType And am_String) = am_String Then
        End If
    End If
End Sub
Private Sub tFlex_KeyPress(KeyAscii As Integer)

    Dim iTmpType As iPropG
    
    Select Case miCol
        Case 0, 1
        Case 2, 3, 5
            iTmpType = am_Long
        Case 4
            iTmpType = am_String
    End Select
    
    Select Case KeyAscii
        Case 60, 64, 124, 91, 93    '<>|[]
            If miCol <> 4 Then KeyAscii = 0
        Case vbKeyEscape
            m_blnAnnullaMod = True
            tFlex = Flex.Text
            Flex.SetFocus
            KeyAscii = 0
        Case vbKeyReturn
            Flex.SetFocus
            KeyAscii = 0
        Case vbKeyBACK
            If (iTmpType And am_Dec) = am_Dec Then
                Decimal_kPress tFlex, KeyAscii
            ElseIf (iTmpType And am_Long) = am_Long Then
                Digit_KPress tFlex, KeyAscii
            'ElseIf (iTmpType And am_Date) = am_Date Then
                'Data_KPress tFlex, KeyAscii, Left$(kYEAR, Elem(iElem).iLen)
            '    Data_KPress tFlex, KeyAscii, Left$(kYEAR, iLen)
            ElseIf (iTmpType And am_String) = am_String Then

            End If
        Case vbKeySpace
            If miCol <> 4 Then KeyAscii = 0
        Case Else
            If (iTmpType And am_Dec) = am_Dec Then
                Decimal_kPress tFlex, KeyAscii
            ElseIf (iTmpType And am_Long) = am_Long Then
                Digit_KPress tFlex, KeyAscii
            'ElseIf (iTmpType And am_Date) = am_Date Then
                'Data_KPress tFlex, KeyAscii, Left$(kYEAR, Elem(iElem).iLen)
            '    Data_KPress tFlex, KeyAscii, Left$(kYEAR, iLen)
            ElseIf (iTmpType And am_String) = am_String Then
                'OverStrike tFlex
            End If
    End Select

End Sub
Private Sub tFlex_LostFocus()
        
    If Not m_blnAnnullaMod Then
        tFlex = Trim$(tFlex)
        Select Case miCol
'            Case 2
'                bcpCol.Item(miRow).oSize = ValInt(tFlex.Text)
            Case 3
                bcpCol.Item(miRow).oSize2 = ValInt(tFlex.Text)
            Case 4
                bcpCol.Item(miRow).Separ = tFlex.Text
            Case 5
                bcpCol.Item(miRow).iPos = ValInt(tFlex.Text)
        End Select
        bFMTDirty = True
        Flex.TextMatrix(miRow, miCol) = tFlex
    End If

    m_blnAnnullaMod = True
    tFlex.Visible = False
    Flex.SetFocus
    
End Sub
Private Sub TestActiveControl()

    If cbo.Visible Then
        cbo_LostFocus
    ElseIf tFlex.Visible Then
        tFlex_LostFocus
    End If

   On Local Error Resume Next
   If TypeOf Me.ActiveControl Is TextBox Then
      If Me.ActiveControl Is tFlex Then
         tFlex_LostFocus
      End If
   ElseIf TypeOf Me.ActiveControl Is ComboBox Then
      If Me.ActiveControl Is cbo Then
         cbo_LostFocus
      End If
   End If
   On Local Error GoTo 0

End Sub

Private Sub LoadCboEdit(ByVal bDataType As Boolean)

    Dim i As Integer
    Dim oType As SQLDMO.SystemDatatype2
    Dim oDb As SQLDMO.Database2
    Dim vVar As Variant
    Dim oQry As SQLDMO.QueryResults
    
    Set oDb = objServer.Databases(sDbName)
    
    cbo.Clear
    If bDataType Then
        For Each oType In oDb.SystemDatatypes
            cbo.AddItem oType.Name
        Next
        Set oType = Nothing
        
        For i = 1 To oDb.UserDefinedDatatypes.Count
            cbo.AddItem oDb.UserDefinedDatatypes(i).Name & " (" & oDb.UserDefinedDatatypes(i).BaseType & ")"
            cbo.ItemData(cbo.NewIndex) = True
        Next
    Else
        Set oQry = objServer.EnumCollations()
        If Not oQry Is Nothing Then vVar = Qry2Var(oQry)
        If IsArray(vVar) Then
            With cbo
                For i = 0 To UBound(vVar, 2)
                    .AddItem vVar(0, i)
                Next
            End With
        End If
    End If
    Set oDb = Nothing
    
End Sub
Private Sub SizeControls()
    Dim i As Integer
    
    picNav.Top = Me.ScaleHeight - picNav.Height
    
    cmdNav(4).Left = picNav.ScaleWidth - k_MargineFRM - cmdNav(4).Width
    cmdNav(3).Left = cmdNav(4).Left - k_MargineFRM - cmdNav(3).Width
    cmdNav(2).Left = cmdNav(3).Left - cmdNav(2).Width - 12
    cmdNav(1).Left = cmdNav(2).Left - k_MargineFRM - cmdNav(1).Width
        
    Line1(0).X1 = k_FlexBevel / 2
    Line1(0).X2 = picNav.ScaleWidth - (k_FlexBevel / 2)
    Line1(1).X1 = k_FlexBevel / 2
    Line1(1).X2 = picNav.ScaleWidth - (k_FlexBevel / 2)
    For i = 0 To 4
        fraStep(i).Width = Me.ScaleWidth - k_FlexBevel
        fraStep(i).Height = picNav.Top - 12
        
        fraStep(i).Left = 0
    Next
    tView.Left = (k_FlexBevel / 2)
    tView.Height = fraStep(1).Height - k_FlexBevel
    tView.Width = fraStep(1).Width - tDati(2).Width - cmd(2).Width - (2 * k_FlexBevel)
    tDati(2).Left = tView.Width + (k_FlexBevel / 2)
    tDati(2).Top = fraStep(1).Height - (3 * tDati(2).Height)
    cmd(2).Left = tDati(2).Left + tDati(2).Width + (k_FlexBevel / 2)
    cmd(2).Top = tDati(2).Top + (cmd(2).Height - tDati(2).Height)
    lbl(1).Top = tDati(2).Top - k_FlexBevel - lbl(1).Height '(2 * lbl(1).Height)
    lblTabe.Top = fraStep(1).Height \ 3
    lblStep(2).Left = tView.Width + k_FlexBevel
    lblTabe.Left = tView.Width + k_FlexBevel
    lbl(1).Left = tView.Width + k_FlexBevel
    
    fra(2).Left = k_FlexBevel
    fra(3).Left = Me.ScaleWidth - fra(3).Width - k_FlexBevel
    Flex.Width = fraStep(3).Width - (2 * Flex.Left)
    SettaLargFlex Flex, 0, 200
    
    rtf.Width = fraStep(4).Width - (2 * rtf.Left)
    
End Sub
