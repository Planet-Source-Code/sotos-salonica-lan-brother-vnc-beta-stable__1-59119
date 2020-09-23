VERSION 5.00
Object = "{49633AB0-0296-4F1A-897F-19017A9AE174}#1.0#0"; "VNCX.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{77EBD0B1-871A-4AD1-951A-26AEFE783111}#2.1#0"; "vbalExpBar6.ocx"
Object = "{9DC93C3A-4153-440A-88A7-A10AEDA3BAAA}#3.5#0"; "vbalDTab6.ocx"
Begin VB.Form FRMSympan 
   Caption         =   "Lan Brother 0.2 VNC (beta stable)"
   ClientHeight    =   10470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14850
   Icon            =   "RemoteLan.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   698
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   990
   StartUpPosition =   2  'CenterScreen
   Begin RemoteLan.Socket vncCheck 
      Index           =   0
      Left            =   8010
      Top             =   6210
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.PictureBox PicLanBrother 
      BackColor       =   &H00000000&
      Height          =   3345
      Index           =   0
      Left            =   3990
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   219
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   275
      TabIndex        =   1
      Top             =   750
      Visible         =   0   'False
      Width           =   4185
   End
   Begin VNCXLibCtl.VNCViewer RemoteDesk 
      Height          =   4485
      Index           =   0
      Left            =   3990
      OleObjectBlob   =   "RemoteLan.frx":628A
      TabIndex        =   6
      Top             =   750
      Width           =   4155
   End
   Begin VNCXLibCtl.VNCViewer VNCViewer1 
      Height          =   1875
      Index           =   0
      Left            =   4410
      OleObjectBlob   =   "RemoteLan.frx":635A
      TabIndex        =   10
      Top             =   6270
      Width           =   2685
   End
   Begin VB.PictureBox PicPanel 
      BackColor       =   &H00FF8080&
      Height          =   405
      Left            =   3990
      ScaleHeight     =   345
      ScaleWidth      =   9645
      TabIndex        =   17
      Top             =   0
      Width           =   9705
      Begin VB.TextBox txtpswd 
         Height          =   315
         Left            =   6540
         TabIndex        =   26
         Top             =   0
         Width           =   1065
      End
      Begin VB.CommandButton Command3 
         Caption         =   "c&Lipboard"
         Enabled         =   0   'False
         Height          =   285
         Left            =   4260
         TabIndex        =   25
         Top             =   30
         Width           =   885
      End
      Begin VB.CommandButton CmdCtrlAltDel 
         Caption         =   "Ctrl-Alt-&Del"
         Height          =   285
         Left            =   3300
         TabIndex        =   24
         Top             =   30
         Width           =   915
      End
      Begin VB.CommandButton CmdRefresh 
         Caption         =   "&REFRESH"
         Height          =   285
         Left            =   2310
         TabIndex        =   23
         Top             =   30
         Width           =   945
      End
      Begin VB.CheckBox ChkControl 
         Caption         =   "&CONTROL"
         Height          =   285
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Ðëçñçò Ýëåí÷ïò Þ áðëü View."
         Top             =   30
         Value           =   1  'Checked
         Width           =   825
      End
      Begin VB.ComboBox Combo2 
         Enabled         =   0   'False
         Height          =   315
         Left            =   7860
         TabIndex        =   21
         ToolTipText     =   "×ñþìáôá"
         Top             =   0
         Width           =   525
      End
      Begin VB.CheckBox Check2 
         Caption         =   "AL&T"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1620
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "ÐáôçìÝíï ALT"
         Top             =   30
         Width           =   555
      End
      Begin VB.CheckBox ChkCtrl 
         Caption         =   "&CTRL"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1020
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "ÐáôçìÝíï Control"
         Top             =   30
         Width           =   555
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   8460
         TabIndex        =   18
         ToolTipText     =   "Êùäéêïðïßçóç"
         Top             =   0
         Width           =   645
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "VNC Password*"
         Height          =   285
         Left            =   5340
         TabIndex        =   27
         Top             =   30
         Width           =   1155
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FF0000&
         X1              =   7680
         X2              =   7680
         Y1              =   0
         Y2              =   360
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FF0000&
         X1              =   2250
         X2              =   2250
         Y1              =   0
         Y2              =   360
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         X1              =   960
         X2              =   960
         Y1              =   0
         Y2              =   360
      End
   End
   Begin VB.Timer TimChecktabs 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   13860
      Top             =   0
   End
   Begin VB.VScrollBar ScrlLan 
      Height          =   3345
      Left            =   8760
      TabIndex        =   16
      Top             =   780
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox picConnecting 
      BackColor       =   &H00FFC0C0&
      Height          =   1095
      Left            =   4440
      ScaleHeight     =   1035
      ScaleWidth      =   4275
      TabIndex        =   13
      Top             =   1560
      Visible         =   0   'False
      Width           =   4335
      Begin VB.Label lblConnecting2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "PLEASE WAIT...."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   720
         TabIndex        =   15
         Top             =   540
         Width           =   2775
      End
      Begin VB.Label lblConnecting 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Connecting to Workstation."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   270
         TabIndex        =   14
         Top             =   150
         Width           =   3825
      End
   End
   Begin VB.CheckBox chkViewOnly 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Check1"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   4410
      TabIndex        =   12
      Top             =   5940
      Visible         =   0   'False
      Width           =   195
   End
   Begin vbalDTab6.vbalDTabControl tabMulty 
      Height          =   315
      Left            =   3990
      TabIndex        =   7
      Top             =   450
      Visible         =   0   'False
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty SelectedFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12640511
   End
   Begin VB.CommandButton cmdBarSizer 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   12015
      Left            =   3900
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   105
   End
   Begin VB.CheckBox chkLanInfo 
      BackColor       =   &H00C0FFFF&
      Caption         =   "LAN Thorough Check. (Slow)"
      Height          =   255
      Left            =   510
      TabIndex        =   4
      Top             =   8400
      Width           =   255
   End
   Begin MSComctlLib.ListView lstLan 
      Height          =   3345
      Left            =   480
      TabIndex        =   3
      Top             =   4410
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   5900
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2012
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "VNC"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "IP"
         Object.Width           =   2487
      EndProperty
   End
   Begin VB.CommandButton cmdLanRefresh 
      Caption         =   "Refresh"
      Height          =   285
      Left            =   480
      TabIndex        =   2
      Top             =   7980
      Width           =   285
   End
   Begin vbalExplorerBarLib6.vbalExplorerBarCtl ExpBar 
      Height          =   12060
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   21273
      BackColorEnd    =   8388608
      BackColorStart  =   16761024
      Begin VB.PictureBox PicGlobe 
         Height          =   465
         Left            =   330
         Picture         =   "RemoteLan.frx":642A
         ScaleHeight     =   405
         ScaleWidth      =   555
         TabIndex        =   8
         Top             =   8880
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdSendServer 
         Caption         =   "Send VNCServer"
         Enabled         =   0   'False
         Height          =   285
         Left            =   780
         TabIndex        =   9
         Top             =   7980
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.Label lblvncName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Caption         =   "no Connection"
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   0
      Left            =   4410
      TabIndex        =   11
      Top             =   6030
      Visible         =   0   'False
      Width           =   2685
   End
End
Attribute VB_Name = "FRMSympan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tvWidth(100) As Integer
Dim tvHeight(100) As Integer
Dim ReHeight(100) As Integer
Dim ReWidth(100) As Integer
Dim StoredTop As Integer
Dim StoredLeft As Integer
Dim LastConnection As String
Dim CurrentDeskView As Integer
Dim LanBroOnTop As Boolean
Dim TVitems As Integer
Dim TVitemsWorking As Integer
Dim TVitemNames(100) As String
Dim TVItemsPerRow As Integer

Private Sub Form_Load()
    
    Call InitExpBar
    Call UpdateMachines
End Sub

Private Sub CmdCtrlAltDel_Click()
'ÁÐÏÓÔÏËÇ CONTROL-ALT-DELETE ÓÔÏ PC
If ChkControl.Value = vbChecked Then
    RemoteDesk(CurrentDeskView).SendKeyPress (46), , 6
End If
End Sub

Private Sub cmdLanRefresh_Click()
    Call UpdateMachines
End Sub

Private Sub cmdBarSizer_Click()
'ôï do
'äéáöïñåôéêÞ ìðáñá ïôáí ìéêñÝíåé.
    ExpBar.Redraw = False
    RemWidth = RemoteDesk(CurrentDeskView).Width
    ExpBar.Bars.Clear
    If ExpBar.Width > 100 Then
        f = -11
        For x = ExpBar.Width To 59 Step -10
            f = f + 10
            ExpBar.Width = x
            cmdBarSizer.Left = ExpBar.Width
            RemoteDesk(CurrentDeskView).Left = cmdBarSizer.Left + 8
            RemoteDesk(CurrentDeskView).Width = RemWidth + f
            PicLanBrother(0).Left = cmdBarSizer.Left + 8
            PicLanBrother(0).Width = RemWidth + f
        ' DoEvents
        Next x
    Else
        f = -11
        For x = 59 To 259 Step 10
            f = f + 10
            ExpBar.Width = x
            cmdBarSizer.Left = ExpBar.Width + 1
            RemoteDesk(CurrentDeskView).Left = cmdBarSizer.Left + 8
            RemoteDesk(CurrentDeskView).Width = RemWidth - f
            PicLanBrother(0).Left = cmdBarSizer.Left + 8
            PicLanBrother(0).Width = RemWidth - f
        ' DoEvents
        Next x
        Call InitExpBar
    End If
    ExpBar.Redraw = True
    Call Form_Resize
End Sub

Private Sub CmdRefresh_Click()
RemoteDesk(CurrentDeskView).Refresh
End Sub

Private Sub cmdSendServer_Click()
' ôï do

    Debug.Print lstLan.SelectedItem.Text
End Sub





Private Sub ExpBar_ItemClick(itm As vbalExplorerBarLib6.cExplorerBarItem)
Dim tabX As cTab
'Debug.Print itm.Text
    Select Case itm.Text
        Case "Connect"
            Call lstLan_DblClick
        Case "Add to LanBrother"
            If PicLanBrother(0).Visible = True Then
                Call LanTv(lstLan.SelectedItem.Text)
            End If
        Case "Create LanBrother"
            If PicLanBrother(0).Visible = True Then ' åëåí÷ïò áí õðáñ÷åé áëëïò lanbrother
                PicLanBrother(0).ZOrder 0 ' öåñå ôïí ìðñïóôÜ
                LanBroOnTop = True
                BringfrontRemote ("LanBrother")
                Exit Sub
            End If
            PicLanBrother(0).Visible = True
            PicLanBrother(0).ZOrder 0
            If tabMulty.Tabs.Count = 0 Then
                Set tabX = tabMulty.Tabs.Add("LanBrother", , "LanBrother")
                tabMulty.Visible = True
             Else
                Set tabX = tabMulty.Tabs.Add("LanBrother", LastConnection, "LanBrother")
                tabX.Selected = True
            End If
            TVitems = 0
            TVitemsWorking = 0
            LastConnection = "LanBrother"
            LanBroOnTop = True
            BringfrontRemote ("LanBrother")
    End Select
End Sub



Private Sub InitExpBar()
Dim cBar As cExplorerBar
Dim cItem As cExplorerBarItem

    ExpBar.UseExplorerStyle = False

    Set cBar = ExpBar.Bars.Add(, , "ÅÑÃÁÓÉÅÓ")
    cBar.BackColor = &HC0FFFF '&HFFC0C0
    cBar.IsSpecial = True
    cBar.WatermarkMode = eWaterMarkColourise
    cBar.WatermarkPicture = PicGlobe.Picture
    Set cItem = cBar.Items.Add(, , "New Connection.", , eItemLink)
    cItem.Bold = True
    cItem.TextColor = vbBlack
    Set cItem = cBar.Items.Add(, , "Create LanBrother", , eItemLink)
    cItem.Bold = True
    Set cItem = cBar.Items.Add(, , "Open Connection.", , eItemLink)
    cItem.Bold = True
    cItem.TextColor = vbBlack
    Set cItem = cBar.Items.Add(, , "Save Connection.", , eItemLink)
    cItem.Bold = True
    cItem.TextColor = vbBlack
    Set cItem = cBar.Items.Add(, , "Properties.", , eItemLink)
    cItem.Bold = True
    cItem.TextColor = vbBlack
    Set cItem = cBar.Items.Add(, , "Exit.", , eItemLink)
    cItem.Bold = True
    cItem.TextColor = vbBlack
    
    Set cBar = ExpBar.Bars.Add(, , "LOCAL NETWORK")
    Set cItem = cBar.Items.Add(, , "plusLan", , eItemControlPlaceHolder)
    cItem.Control = chkLanInfo
    Set cItem = cBar.Items.Add(, , "Name                   VNC    Address - IP", , eItemText)
    cBar.BackColor = &HC0FFFF
    Set cItem = cBar.Items.Add(, , "", , eItemControlPlaceHolder)
    cItem.Control = lstLan
    Set cItem = cBar.Items.Add(, , "RefLan", , eItemControlPlaceHolder)
    cItem.Control = cmdLanRefresh
    Set cItem = cBar.Items.Add(, , "", , eItemText)
    
    Set cItem = cBar.Items.Add(, , " ", , eItemText)
    'Set cItem = cBar.Items.Add(, , " ", , eItemControlPlaceHolder)
    'cItem.Control = cmdSendServer
    'Set cItem = cBar.Items.Add(, , " ", , eItemText)
    Set cItem = cBar.Items.Add(, , "Connect", , eItemLink)
    cItem.Bold = True
    Set cItem = cBar.Items.Add(, , "Deploy VNC Server", , eItemLink)
    cItem.Bold = True
    cItem.TextColor = vbBlack
    Set cItem = cBar.Items.Add(, , "Add to LanBrother", , eItemLink)
    cItem.Bold = True

    Set cBar = ExpBar.Bars.Add(, "info", "Information")
    'cBar.State = eBarCollapsed
    'Call CustomColours
End Sub

Private Sub LanTv(Rname As String)
'ÓÅ ÁÕÔÇÍ ÔÇÍ ÑÏÕÔÉÍÁ ÈÁ ÐÑÏÓÈÅÔÏ ÔÁ tvcontorls óôï PICTUREBOX PIC LAN BROTHER
Dim LevelHeight As Integer
Dim LevelLeft As Integer
'    TVItemsPerRow = CInt(Val(txtPerRow.Text))
    'Debug.Print lblvncName(0).Container.Name
    If TVitems = 0 Then
        Set lblvncName(0).Container = PicLanBrother(0)
        Set VNCViewer1(0).Container = PicLanBrother(0)
        Set chkViewOnly(0).Container = PicLanBrother(0)
        'Set ScrlLan.Container = PicLanBrother(0)
        lblvncName(0).Top = 0
        lblvncName(0).Left = 0
        lblvncName(0).Caption = Rname
        VNCViewer1(0).Top = lblvncName(0).Height
        VNCViewer1(0).Left = 0
        VNCViewer1(0).BackColor = vbGreen
        VNCViewer1(0).ForeColor = vbGreen
        chkViewOnly(0).Top = 1
        chkViewOnly(0).Left = 0
        TVitemNames(0) = Rname
        VNCViewer1(0).Connect Rname, 5900, txtpswd
        lblvncName(0).Visible = True
        VNCViewer1(0).Visible = True
        chkViewOnly(0).Visible = True
        TVitems = 1
        Exit Sub
    End If
    For x = 0 To TVitems
        If TVitemNames(x) = Rname Then
           MsgBox "Workstation is Allready Loaded.", , "Ê.Å.Ó"
           Exit Sub
        End If
    Next x
    On Error Resume Next
    Load lblvncName(TVitems)
    Load VNCViewer1(TVitems)
    Load chkViewOnly(TVitems)
    Set lblvncName(TVitems).Container = PicLanBrother(0)
    Set VNCViewer1(TVitems).Container = PicLanBrother(0)
    Set chkViewOnly(TVitems).Container = PicLanBrother(0) 'lblvncName(TVitems) ' PicLanBrother(0)
    lblvncName(TVitems).Caption = Rname
    TVitemNames(TVitems) = Rname
    lblvncName(TVitems).Visible = True
    VNCViewer1(TVitems).Visible = True
    chkViewOnly(TVitems).Visible = True
    VNCViewer1(TVitems).Connect Rname, 5900, txtpswd
    TVitems = TVitems + 1
    Call ReArrangeTv
    'TVitems = TVitems + 1
    Exit Sub
    
    Select Case TVitems
      Case Is <= (TVItemsPerRow - 1)
        LevelHeight = 0
        LevelLeft = (TVitems) * (lblvncName(0).Left + lblvncName(0).Width)
      Case Is <= (TVItemsPerRow * 2) - 1
        LevelHeight = lblvncName(0).Height + VNCViewer1(0).Height + 1
        LevelLeft = (((TVitems - TVItemsPerRow) * (lblvncName(0).Left + lblvncName(0).Width))) + 1
      Case Is <= (TVItemsPerRow * 3) - 1
        LevelHeight = ((lblvncName(0).Height + VNCViewer1(0).Height + 1) * 2) + 1
        LevelLeft = (((TVitems - (TVItemsPerRow * 2)) * (lblvncName(0).Left + lblvncName(0).Width))) + 1
      Case Is <= (TVItemsPerRow * 4) - 1
        LevelHeight = ((lblvncName(0).Height + VNCViewer1(0).Height + 1) * 3) + 1
        LevelLeft = (((TVitems - (TVItemsPerRow * 3)) * (lblvncName(0).Left + lblvncName(0).Width))) + 1
      Case ID <= (TVItemsPerRow * 5) - 1
        LevelHeight = ((lblvncName(0).Height + VNCViewer1(0).Height + 1) * 4) + 1
        LevelLeft = (((TVitems - (TVItemsPerRow * 4)) * (lblvncName(0).Left + lblvncName(0).Width))) + 1
    End Select
End Sub

Private Sub ReArrangeTv()
Dim LevelHeight As Integer
Dim LevelLeft As Integer
Dim level As Integer
Dim x As Integer
Dim PeRow As String

    PeRow = Left(CStr(PicLanBrother(0).Width / lblvncName(0).Width), 1)
    If PeRow = "0" Then PeRow = "1"
    TVItemsPerRow = CInt(Val(PeRow))
    level = 1
    LevelLeft = lblvncName(0).Width
    
    For x = 1 To TVitems - 1
        If level = TVItemsPerRow Then
            LevelHeight = LevelHeight + (lblvncName(0).Height + VNCViewer1(0).Height)
            level = 0
        End If
        
        lblvncName(x).Top = LevelHeight
        lblvncName(x).Left = level * LevelLeft
        VNCViewer1(x).Top = LevelHeight + lblvncName(0).Height
        VNCViewer1(x).Left = level * LevelLeft
        chkViewOnly(x).Top = LevelHeight + 1
        chkViewOnly(x).Left = level * LevelLeft
        lblvncName(x).ZOrder 0
        level = level + 1
    Next
    If VNCViewer1(TVitems - 1).Top + VNCViewer1(TVitems - 1).Height > PicLanBrother(0).Height Then
        'äçìéïõñãßá scrollbar
        PicLanBrother(0).Height = VNCViewer1(TVitems - 1).Top + VNCViewer1(TVitems - 1).Height
        ScrlLan.Top = PicLanBrother(0).Top
        ScrlLan.Left = FRMSympan.ScaleWidth - ScrlLan.Width - 1
        ScrlLan.Height = FRMSympan.ScaleHeight - PicLanBrother(0).Top
        ScrlLan.Visible = True
        ScrlLan.ZOrder 0
    Else
        'áðïêñõøç scrollbar
        ScrlLan.Visible = False
    End If
        
End Sub

Private Sub Form_Resize()
'******************************************
'ÄÅÍ ÈÅËÙ RESIZE ÓÅ ÐÏËÕ ÌÉÊÑÅÓ ÄÉÁÓÔÁÓÅÉÓ*
'******************************************

On Error Resume Next
    If FRMSympan.ScaleHeight < 300 Or FRMSympan.ScaleWidth < 400 Then
        'ìçí êáíåéò áëëï resize (to do)
        Exit Sub
    End If
    ExpBar.Redraw = False
    ExpBar.Height = FRMSympan.ScaleHeight
    cmdBarSizer.Height = FRMSympan.ScaleHeight
    
    If LanBroOnTop = True Then
        PicLanBrother(0).Height = FRMSympan.ScaleHeight - PicLanBrother(0).Top
        PicLanBrother(0).Width = FRMSympan.ScaleWidth - (cmdBarSizer.Left + cmdBarSizer.Width) + 1
        If TVitems > 1 Then Call ReArrangeTv
        tabMulty.Left = PicLanBrother(0).Left + 2
        tabMulty.Width = PicLanBrother(0).Width
        ExpBar.Redraw = True
        Call ServerInfo("LanBrother")
        Exit Sub
    End If
    
    RemoteDesk(CurrentDeskView).Height = FRMSympan.ScaleHeight - RemoteDesk(CurrentDeskView).Top
    RemoteDesk(CurrentDeskView).Width = FRMSympan.ScaleWidth - (cmdBarSizer.Left + cmdBarSizer.Width) + 1
    With RemoteDesk(CurrentDeskView)
        .Stretch = False
        .StretchY 10, (ReHeight(CurrentDeskView) / .Height) * 10.5
        .StretchX 10, (ReWidth(CurrentDeskView) / .Width) * 10.5
        .StretchMode = vncxStretchModeHalftone
        .Stretch = True
    End With
    ExpBar.Redraw = True
    tabMulty.Left = RemoteDesk(CurrentDeskView).Left + 2
    tabMulty.Width = RemoteDesk(CurrentDeskView).Width
    PicPanel.Left = RemoteDesk(CurrentDeskView).Left + 2
    PicPanel.Width = RemoteDesk(CurrentDeskView).Width
    Call ServerInfo(Str(CurrentDeskView))
End Sub

Private Sub Form_Unload(Cancel As Integer)
    For xcount = 0 To TVitems - 1
        If VNCViewer1(xcount).IsConnected Then
            VNCViewer1(xcount).Stop
            VNCViewer1(xcount).Disconnect
        End If
    Next xcount
    For xcount = 0 To RemoteDesk.Count - 1
        If RemoteDesk(xcount).IsConnected Then
            RemoteDesk(xcount).Stop
            RemoteDesk(xcount).Disconnect
        End If
    Next xcount
End Sub

Private Sub ImgShrink_Click()
    VNCViewer1(picCTRL.Tag).Top = StoredTop
    VNCViewer1(picCTRL.Tag).Left = StoredLeft
    VNCViewer1(picCTRL.Tag).Width = 205
    VNCViewer1(picCTRL.Tag).Height = 155
    VNCViewer1(picCTRL.Tag).ScrollBars = vncxScrollbarNone
    sldZoom.Value = 20
    picCTRL.Visible = False
End Sub

Private Sub lblvncName_DblClick(Index As Integer)
' äéðëï êëéê óå ìðáñá LanTVitem

  If VNCViewer1(Index).IsConnected Then
        StoredTop = VNCViewer1(Index).Top
        StoredLeft = VNCViewer1(Index).Left
        VNCViewer1(Index).ZOrder 0
        VNCViewer1(Index).Left = 5
        VNCViewer1(Index).Width = 620
        VNCViewer1(Index).Top = 64
        VNCViewer1(Index).Height = 506
        picCTRL.Tag = Index
        picCTRL.Visible = True
        chkView.Value = chkViewOnly(Index).Value
        lblname.Caption = lblvncName(Index).Caption
        VNCViewer1(picCTRL.Tag).ScrollBars = vncxScrollbarBoth
        sldZoom.Value = 60
  End If
  
End Sub

Private Sub lstLan_BeforeLabelEdit(Cancel As Integer)
' ÃÉÁ ÍÁ ÌÇÍ ÅÐÉÔÑÅÐÅÉ EDITING ÔÉÓ ËÉÓÔÁÓ
    Cancel = True
End Sub

Private Sub lstLan_DblClick()
' Ï ×ÑÇÓÔÇÓ ÅÊÁÍÅ ÄÉÐËÏ ÊËÉÊ ÓÅ ÊÁÐÏÉÏ PC ÁÐÏ ÔÇÍ ËÉÓÔÁ ÔÏÕ ÄÉÊÔ¾ÏÕ
Dim DesksOpen As Integer
Dim NewIndex As Integer
    On Error Resume Next 'Åëåí÷ïò áí õðáñ÷åé Þäç áíïéêôü ðáñÜèçñï ìå ôïí óõãêåêñéìåíï õðïëïãéóôç
    If tabMulty.Tabs.Count > 0 Then
        For x = 0 To tabMulty.Tabs.Count
            'áí õðáñ÷åé çäç áíïéêôï ðáñáèõñï ìå ôïí óõãêåêñéìåíï Ç/Õ öåñå ôï ìðñïóôÜ
            If IsError(tabMulty.Tabs.Item("0" + Str(x)).Caption = lstLan.SelectedItem) Then
                'Resume Next
              Else
                If tabMulty.Tabs.Item("0" + Str(x)).Caption = lstLan.SelectedItem Then
                    BringfrontRemote ("0" + Str(x))
                    Exit Sub
                End If
            End If
        Next x
    End If
  
    NewIndex = RemoteDesk.Count
    For x = 0 To RemoteDesk.Count    'êáíù åëåí÷ï áí êáðïéá áðï ôá öïñôïìåíá controls åéíáé áðïóõíäåìåíï
        If RemoteDesk(x).IsConnected = False Then
            NewIndex = x   'áí å÷ù control á÷ñçóéìïðïßçôï ôïôå öïñôùíù ôçí óõíäåóç ó'áõôï
            Exit For
        End If
    Next x
    If IsError(RemoteDesk(NewIndex).IsConnected) Then
        Load RemoteDesk(NewIndex) 'áí ôåëéùóåé ôï ðáñáðÜíù loop êáé äåí å÷ù âñåé control , äçìéïõñãù åíá
    End If
    picConnecting.Visible = True 'åìöíáíßæù Ìýíçìá áðïðåéñáò óýíäåóçò
    picConnecting.ZOrder 0       'êáé ôï öÝñíù on top
    RemoteDesk(NewIndex).Connect lstLan.SelectedItem, 5900, txtpswd 'êáé êÜíù ôçí óýíäåóç
End Sub

Private Function BringfrontRemote(Index As String)
Dim sMsg As String
Dim x As Integer
    'RemoteDesk(CurrentDeskView).RequestFrames = False
    If Index = "LanBrother" Then
       LanBroOnTop = True
       PicLanBrother(0).ZOrder 0
       CurrentDeskView = 0 'íá äïêéìáóù -1
       For x = 0 To RemoteDesk.Count - 1
            RemoteDesk(x).RequestFrames = False
       Next x
       tabMulty.Tabs.Item(Index).Selected = True
       Call Form_Resize
       Exit Function
    End If
    LanBroOnTop = False
    x = Val(Index)
    CurrentDeskView = x 'ÃÉÁ ÔÏ RESIZE
    tabMulty.Tabs.Item(Index).Selected = True 'öåñíù ìðñïóôá ôï óùóôï tab
    RemoteDesk(x).ZOrder 0 'ôï óùóôï remote
    RemoteDesk(x).RequestFrames = True 'æçôáù áíáíåùóåéò ïèïíçò
    RemoteDesk(x).Refresh
    Call Form_Resize 'êáé êáíù resize
    
End Function
Private Sub lstLan_ItemClick(ByVal Item As MSComctlLib.ListItem)

'Debug.Print Item
'Åäù ìðïñù íá âáëù íá åìöáíéæïíôáé Ý÷ôñá ðëçñïöïñéåò ãéá ôïí ç/õ ðïõ åðéëå÷èçêå

End Sub

'1***********************************************************
Private Sub RemoteDesk_BeforeKeyDown(Index As Integer, ByVal Modifiers As VNCXLibCtl.VNCXModifiersEnum, ByVal KeyCode As Long, Action As VNCXLibCtl.VNCXActionsEnum)
If ChkControl.Value = vbUnchecked Then
  Action = vncxActionNone
End If
End Sub

Private Sub RemoteDesk_BeforeKeyUp(Index As Integer, ByVal Modifiers As VNCXLibCtl.VNCXModifiersEnum, ByVal KeyCode As Long, Action As VNCXLibCtl.VNCXActionsEnum)
If ChkControl.Value = vbUnchecked Then
  Action = vncxActionNone
End If
End Sub

Private Sub RemoteDesk_BeforeMouseDown(Index As Integer, ByVal Buttons As VNCXLibCtl.VNCXButtonsEnum, ByVal Modifiers As VNCXLibCtl.VNCXModifiersEnum, ByVal x As Long, ByVal y As Long, Action As VNCXLibCtl.VNCXActionsEnum)
If ChkControl.Value = vbUnchecked Then
  Action = vncxActionNone
End If
End Sub

Private Sub RemoteDesk_BeforeMouseMove(Index As Integer, ByVal Buttons As VNCXLibCtl.VNCXButtonsEnum, ByVal Modifiers As VNCXLibCtl.VNCXModifiersEnum, ByVal x As Long, ByVal y As Long, Action As VNCXLibCtl.VNCXActionsEnum)
If ChkControl.Value = vbUnchecked Then
  Action = vncxActionNone
End If
End Sub

Private Sub RemoteDesk_BeforeMouseUp(Index As Integer, ByVal Buttons As VNCXLibCtl.VNCXButtonsEnum, ByVal Modifiers As VNCXLibCtl.VNCXModifiersEnum, ByVal x As Long, ByVal y As Long, Action As VNCXLibCtl.VNCXActionsEnum)
If ChkControl.Value = vbUnchecked Then
  Action = vncxActionNone
End If
End Sub

Private Sub RemoteDesk_BeforeMouseWheel(Index As Integer, ByVal Buttons As VNCXLibCtl.VNCXButtonsEnum, ByVal Rotation As Long, ByVal Modifiers As VNCXLibCtl.VNCXModifiersEnum, ByVal x As Long, ByVal y As Long, Action As VNCXLibCtl.VNCXActionsEnum)
If ChkControl.Value = vbUnchecked Then
  Action = vncxActionNone
End If
End Sub

Private Sub RemoteDesk_Connected(Index As Integer, ServerName As String)
    Dim tabX As cTab
    
    RemoteDesk(CurrentDeskView).RequestFrames = False 'óôáìáôþ ôï refresh ôïõ ðñïçãïýìåíïõ remote ãéá ëïãïõò bandwidth
    
    With RemoteDesk(Index)
        .ScrollBars = vncxScrollbarNone
       ' .ProgressiveDrawing = True
        ReHeight(Index) = .FrameHeight
        ReWidth(Index) = .FrameWidth
        '.Display
        .StretchY 10, (ReHeight(Index) / .Height) * 10.5
        .StretchX 10, (ReWidth(Index) / .Width) * 10.5
        .StretchMode = vncxStretchModeHalftone
        .Stretch = True
        .Start
        tabMulty.Visible = True
        .Visible = True
        .ZOrder 0
        If tabMulty.Tabs.Count = 0 Then
          Set tabX = tabMulty.Tabs.Add("0" + Str(Index), , UCase(.ServerName))
          'tabX.Selected = True
         Else
          Set tabX = tabMulty.Tabs.Add("0" + Str(Index), LastConnection, UCase(.ServerName))
          tabX.Selected = True
        End If
        picConnecting.Visible = False
        LastConnection = "0" + Str(Index)
        CurrentDeskView = Index 'ÃÉÁ ÔÏ RESIZE
        LanBroOnTop = False
        '.BitsPerPixel
        .RequestFrames = True 'æçôáù updates ãéá ôï íÝï remote
        '.Refresh
    End With
    
Call Form_Resize

End Sub
Private Sub ServerInfo(ByRef DeskNo As String)
Dim DsN As Integer
    If DeskNo = "LanBrother" Then
        ExpBar.Bars("info").Items.Clear
        sMsg = "ÕðïëïãéóôÝò: " & Str(TVitems) & vbCrLf & _
               "Óå ëåéôïõñãßá:" & Str(TVitemsWorking)
        ExpBar.Bars("info").Items.Add "ServerINFO", , sMsg, , eItemText
        Exit Sub
    End If
    DsN = Val(DeskNo)
    On Error Resume Next
    sMsg = "¼íïìá: " & TrimNull(RemoteDesk(DsN).ServerName) & vbCrLf & _
                "Port: " & TrimNull(RemoteDesk(DsN).Port) & vbCrLf & _
                "Åêäïóç Server: " & TrimNull(RemoteDesk(DskN).ServerVersion) & vbCrLf & _
                "ÁíÜëõóç:     " & TrimNull(RemoteDesk(DskN).FrameWidth) & "x" & TrimNull(RemoteDesk(DsN).FrameHeight) & vbCrLf & _
                "Ôïðéêç áíÜëõóç: " & TrimNull(RemoteDesk(DsN).ViewWidth) & "x" & TrimNull(RemoteDesk(DsN).ViewHeight) & vbCrLf & _
                "Bit áíá Pixel: " & TrimNull(RemoteDesk(DsN).BitsPerPixel)
    ExpBar.Bars("info").Items.Clear
    If RemoteDesk(DsN).ServerName <> "" Then
        ExpBar.Bars("info").Items.Add "ServerINFO", , sMsg, , eItemText
    End If
    
End Sub

Private Sub Slider1_Change()
    RemoteDesk(0).StretchY 11, Slider1.Value
End Sub

Private Sub RemoteDesk_Error(Index As Integer, ByVal Number As VNCXLibCtl.VNCXErrorsEnum, Message As String, ByVal v1 As Variant)
    picConnecting.Visible = False
    Debug.Print Message
End Sub

Private Sub sldZoom_change()
    A = sldZoom.Value / 10
    VNCViewer1(picCTRL.Tag).StretchX A, 10
    VNCViewer1(picCTRL.Tag).StretchY A, 10
End Sub

Private Sub tabMulty_TabClick(theTab As vbalDTab6.cTab, ByVal iButton As MouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Single, ByVal y As Single)
  '********* ÐÑÅÐÅÉ ÍÁ ÂÑ¿ ÔÑ¼ÐÏ ÍÁ ÐÅÑÍÁÙ ÌÐÑÏÓÔÁ ÔÏ ÅÐÉËÅÃÌ¸ÍÏ ÁÍÔÉÊÉÌÅÍÏ
    'Debug.Print theTab.Key
    'Debug.Print theTab.Index
    BringfrontRemote (theTab.Key)
End Sub

Private Sub tabMulty_TabClose(theTab As vbalDTab6.cTab, bCancel As Boolean)
Dim x As Integer
    If theTab.Key = "LanBrother" Then
        For x = 1 To VNCViewer1.Count - 1
            VNCViewer1(x).Stop
            VNCViewer1(x).Disconnect
            Unload VNCViewer1(x)
            Unload lblvncName(x)
            Unload chkViewOnly(x)
        Next x
        lblvncName(0).Caption = ""
        lblvncName(0).Visible = False
        chkViewOnly(0).Visible = False
        VNCViewer1(0).Visible = False
        VNCViewer1(0).Stop
        VNCViewer1(0).Disconnect
        TVitems = 0
        TVitemsWorking = 0
        TimChecktabs.Enabled = True
        PicLanBrother(0).Visible = False
        Exit Sub
    Else
        x = Val(theTab.Key)
        RemoteDesk(x).Stop
        RemoteDesk(x).Disconnect
        RemoteDesk(x).Visible = False
        TimChecktabs.Enabled = True
        If x <> 0 Then
          Unload RemoteDesk(x)
        End If
    End If
End Sub

Private Sub TimChecktabs_Timer()
    If tabMulty.Tabs.Count = 0 Then
        tabMulty.Visible = False
        CurrentDeskView = 0
        RemoteDesk(0).Visible = True
        RemoteDesk(0).ZOrder 0
        TimChecktabs.Enabled = False
       Else
        TimChecktabs.Enabled = False
        BringfrontRemote (tabMulty.SelectedTab.Key)
    End If
End Sub






Private Sub vncCheck_OnConnect(Index As Integer)
    Debug.Print Index & "  socket connected"
End Sub

Private Sub VNCViewer1_BeforeKeyDown(Index As Integer, ByVal Modifiers As VNCXLibCtl.VNCXModifiersEnum, ByVal KeyCode As Long, Action As VNCXLibCtl.VNCXActionsEnum)
    If chkViewOnly(Index).Value = vbUnchecked Then
        Action = vncxActionNone
    End If
End Sub

Private Sub VNCViewer1_BeforeKeyUp(Index As Integer, ByVal Modifiers As VNCXLibCtl.VNCXModifiersEnum, ByVal KeyCode As Long, Action As VNCXLibCtl.VNCXActionsEnum)
    If chkViewOnly(Index).Value = vbUnchecked Then
        Action = vncxActionNone
    End If
End Sub

Private Sub VNCViewer1_BeforeMouseDown(Index As Integer, ByVal Buttons As VNCXLibCtl.VNCXButtonsEnum, ByVal Modifiers As VNCXLibCtl.VNCXModifiersEnum, ByVal x As Long, ByVal y As Long, Action As VNCXLibCtl.VNCXActionsEnum)
    If chkViewOnly(Index).Value = vbUnchecked Then
        Action = vncxActionNone
    End If
End Sub

Private Sub VNCViewer1_BeforeMouseMove(Index As Integer, ByVal Buttons As VNCXLibCtl.VNCXButtonsEnum, ByVal Modifiers As VNCXLibCtl.VNCXModifiersEnum, ByVal x As Long, ByVal y As Long, Action As VNCXLibCtl.VNCXActionsEnum)
    If chkViewOnly(Index).Value = vbUnchecked Then
        Action = vncxActionNone
    End If
End Sub

Private Sub VNCViewer1_BeforeMouseUp(Index As Integer, ByVal Buttons As VNCXLibCtl.VNCXButtonsEnum, ByVal Modifiers As VNCXLibCtl.VNCXModifiersEnum, ByVal x As Long, ByVal y As Long, Action As VNCXLibCtl.VNCXActionsEnum)
    If chkViewOnly(Index).Value = vbUnchecked Then
        Action = vncxActionNone
    End If
End Sub

Private Sub VNCViewer1_BeforeMouseWheel(Index As Integer, ByVal Buttons As VNCXLibCtl.VNCXButtonsEnum, ByVal Rotation As Long, ByVal Modifiers As VNCXLibCtl.VNCXModifiersEnum, ByVal x As Long, ByVal y As Long, Action As VNCXLibCtl.VNCXActionsEnum)
    If chkViewOnly(Index).Value = vbUnchecked Then
        Action = vncxActionNone
    End If
End Sub

Private Sub VNCViewer1_Connected(Index As Integer, ServerName As String)
    
    With VNCViewer1(Index)
        
        .ScrollBars = vncxScrollbarNone
        '.ProgressiveDrawing = True
        If tvHeight(Index) = 0 Then
            tvHeight(Index) = .ViewHeight
            tvWidth(Index) = .ViewWidth
        End If
        .StretchX 10, tvWidth(Index) / .Width * 10.5
        .StretchY 10, tvHeight(Index) / .Height * 10.5
        .StretchMode = vncxStretchModeHalftone
        .Stretch = True
        .Start
        .BitsPerPixel = 8
        .Refresh
        TVitemsWorking = TVitemsWorking + 1
    End With
    
    lblvncName(Index).BackColor = vbGreen
    lblvncName(Index).Caption = UCase(ServerName)
    lblvncName(Index).FontBold = True
    chkViewOnly(Index).Enabled = True
    'VNCViewer1(Index).Refresh
End Sub

Private Sub VNCViewer1_Error(Index As Integer, ByVal Number As VNCXLibCtl.VNCXErrorsEnum, Message As String, ByVal v1 As Variant)
    'Debug.Print "Error index :" & Index & " No:  " & Number
    If Number = 5 Then
    'lblvncName(Index).Caption = "Áäýíáôç óýíäåóç"
    VNCViewer1(Index).Stop
    VNCViewer1(Index).Disconnect
    VNCViewer1(Index).BackColor = vbBlack
    VNCViewer1(Index).ForeColor = vbBlack
End If
End Sub

Private Sub VNCViewer1_CommandNotify(Index As Integer, ByVal v1 As Variant, ByVal v2 As Variant, ByVal v3 As Variant, ByVal v4 As Variant)
    Debug.Print Index & " Command Notify"
End Sub

Private Sub VNCViewer1_DataTransfer(Index As Integer, ByVal TotalRead As Long, ByVal TotalWritten As Long)
    Debug.Print Index & " DataTransfer"
End Sub

Private Sub VNCViewer1_DragDrop(Index As Integer, Source As Control, x As Single, y As Single)
    Debug.Print Index & " DragDrop"
End Sub

Private Sub VNCViewer1_DragOver(Index As Integer, Source As Control, x As Single, y As Single, State As Integer)
    Debug.Print Index & " DrogOver"
End Sub

Private Sub VNCViewer1_FrameEdgeDetected(Index As Integer, ByVal Edge As VNCXLibCtl.VNCXEdgesEnum)
    Debug.Print Index & " FrameEdgeDetected"
End Sub

Private Sub VNCViewer1_FrameUpdated(Index As Integer, ByVal x As Long, ByVal y As Long, ByVal W As Long, ByVal H As Long)
    Debug.Print Index & " FrameUpdated"
End Sub

Private Sub VNCViewer1_LocalClipboard(Index As Integer, ByVal DataType As VNCXLibCtl.VNCXDataTypesEnum, ByVal data As Variant, Action As VNCXLibCtl.VNCXActionsEnum)
    Debug.Print Index & " LocalClipboard"
End Sub

Private Sub VNCViewer1_MouseTrap(Index As Integer, ByVal Trapped As Boolean)
    Debug.Print Index & " MouseTrap"
End Sub

Private Sub VNCViewer1_Password(Index As Integer)
    Debug.Print Index & " Password"
End Sub

Private Sub VNCViewer1_ScreenEdgeDetected(Index As Integer, ByVal Edge As VNCXLibCtl.VNCXEdgesEnum)
    Debug.Print Index & " ScreenEdgeDetected"
End Sub

Private Sub VNCViewer1_ServerBell(Index As Integer, Action As VNCXLibCtl.VNCXBellActionsEnum)
    Debug.Print Index & " ServerBell"
End Sub

Private Sub VNCViewer1_ServerClipboard(Index As Integer, ByVal DataType As VNCXLibCtl.VNCXDataTypesEnum, ByVal data As Variant, Action As VNCXLibCtl.VNCXActionsEnum)
    Debug.Print Index & " ServerClipboard"
End Sub

Private Sub VNCViewer1_Validate(Index As Integer, Cancel As Boolean)
    Debug.Print Index & "Validate"
End Sub

Private Sub VNCViewer1_ViewEdgeDetected(Index As Integer, ByVal Edge As VNCXLibCtl.VNCXEdgesEnum)
    Debug.Print Index & " ViewEdgeDetected"
End Sub

Private Sub VNCViewer1_Warning(Index As Integer, ByVal Number As VNCXLibCtl.VNCXWarningsEnum, Message As String, ByVal v1 As Variant)
    Debug.Print Index & " Warning"
End Sub

Private Sub CustomColours()
Dim i As Long
Dim j As Long
    With ExpBar
         .Redraw = False
         .BackColorStart = RGB(255, 239, 154)
         .BackColorEnd = RGB(137, 129, 93)
         For i = 1 To .Bars.Count
            With .Bars(i)
               If (.IsSpecial) Then
                  .TitleBackColorLight = RGB(137, 129, 93)
                  .TitleBackColorDark = RGB(89, 84, 61)
                  .TitleForeColor = RGB(255, 255, 230)
                  .TitleForeColorOver = RGB(255, 239, 154)
                  .BackColor = RGB(255, 253, 245)
               Else
                  .TitleBackColorLight = RGB(255, 255, 230)
                  .TitleBackColorDark = RGB(255, 239, 154)
                  .TitleForeColor = RGB(89, 84, 61)
                  .TitleForeColorOver = RGB(137, 129, 93)
                  .BackColor = RGB(255, 249, 225)
               End If
               For j = 1 To .Items.Count
                  With .Items(j)
                     .TextColor = RGB(89, 84, 61)
                     .TextColorOver = RGB(170, 163, 130)
                  End With
               Next j
            End With
         Next i
         .Redraw = True
    End With
End Sub

Private Sub UpdateMachines()
'On Error Resume Next
Dim ServerCollection As Collection, i As Long
Dim ipAddress As Collection
Dim FirstCount As Integer
Dim pcOn As Boolean
    
    Me.MousePointer = vbArrowHourglass
    lstLan.ListItems.Clear
    Set ServerCollection = GetServers(NTBased)
        For i = 1 To ServerCollection.Count
            lstLan.ListItems.Add i, , ServerCollection.Item(i)
            lstLan.ListItems.Item(i).ForeColor = vbBlue
            If chkLanInfo.Value = vbChecked Then
               'pcOn = Ping(ServerCollection.Item(i))
                Set ipAddress = GetIPAddresses(ServerCollection.Item(i))
                If ipAddress.Count = 0 Then
                    lstLan.ListItems.Item(i + FirstCount).ListSubItems.Add , , ""
                    lstLan.ListItems.Item(i + FirstCount).ListSubItems.Add , , "Êëåéóôü"
                Else
                    ChkForVncR ipAddress(1), CInt(i)
                    lstLan.ListItems.Item(i).ListSubItems.Add , , "" 'check for vnc
                    lstLan.ListItems.Item(i).ListSubItems.Add , , ipAddress(1)
                    lstLan.ListItems.Item(i).ListSubItems(2).ForeColor = vbBlue
                End If
                
            End If
        Next i
    FirstCount = (ServerCollection.Count)
    Set ServerCollection = GetServers(Windows9x)
        For i = 1 To ServerCollection.Count
            lstLan.ListItems.Add i + FirstCount, , ServerCollection.Item(i)
            If chkLanInfo.Value = vbChecked Then
                Set ipAddress = GetIPAddresses(ServerCollection.Item(i))
                If ipAddress.Count = 0 Then
                    lstLan.ListItems.Item(i + FirstCount).ListSubItems.Add , , ""
                    lstLan.ListItems.Item(i + FirstCount).ListSubItems.Add , , "Êëåéóôü"
                Else
                    ChkForVncR ipAddress(1), CInt(i + FirstCount)
                    lstLan.ListItems.Item(i + FirstCount).ListSubItems.Add , , "" ' check for vnc
                    lstLan.ListItems.Item(i + FirstCount).ListSubItems.Add , , ipAddress(1)
                End If
                
            End If
        Next i
    Me.MousePointer = vbArrow

End Sub

Private Function ChkForVncR(ipAddress As String, Index As Integer)
    Load vncCheck(Index)
   ' Set vncCheck(Index) = New CSocket
    vncCheck(Index).Connect ipAddress, 5900
    
End Function


Private Sub VncCheck_onDataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim data As String, Position As Long
Dim IsVnc As String

    vncCheck(Index).GetData data
    Debug.Print vncCheck(Index).RemoteHost & " -- " & data & vbCrLf
    IsVnc = Left(data, 3)
    If IsVnc = "RFB" Then
        lstLan.ListItems.Item(Index).ListSubItems(1).Text = "N"
        lstLan.ListItems.Item(Index).ListSubItems(1).ForeColor = vbGreen
    End If
    vncCheck(Index).CloseSocket
    Unload vncCheck(Index)
    
End Sub

Private Sub vncCheck_onError(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    lstLan.ListItems.Item(Index).ListSubItems(1).Text = "--"
    lstLan.ListItems.Item(Index).ListSubItems(1).ForeColor = vbRed
    
    Debug.Print vncCheck(Index).RemoteHost & " -" & Str(Number) & "-" & Description
    vncCheck(Index).CloseSocket
    Unload vncCheck(Index)
    
End Sub
