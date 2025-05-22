VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Begin VB.Form frmSegments 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Segment Data"
   ClientHeight    =   8595
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   8040
   HelpContextID   =   3
   Icon            =   "frmSegments.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      HelpContextID   =   3
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Select Segment to Be Edited"
      Top             =   840
      Width           =   3495
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6375
      HelpContextID   =   3
      Left            =   240
      TabIndex        =   7
      Top             =   2040
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   11245
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Morphometry"
      TabPicture(0)   =   "frmSegments.frx":09AA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label3(5)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3(4)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblSegment(6)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblSegment(5)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblSegment(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblSegment(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblSegments(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblSegments(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblSegment(7)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblSegment(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblSegments(20)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblSegments(22)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblSegments(33)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtSegment(2)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtSegment(10)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtSegment(9)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtSegment(8)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtSegment(7)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtSegment(6)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtSegment(5)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtSegment(4)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtSegment(3)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtSegment(1)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Combo2"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).ControlCount=   24
      TabCaption(1)   =   "Observed WQ"
      TabPicture(1)   =   "frmSegments.frx":09C6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label3(3)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label3(2)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblSegments(4)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblSegments(12)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblSegments(11)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lblSegments(10)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lblSegments(9)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lblSegments(8)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "lblSegments(7)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "lblSegments(6)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "lblSegments(5)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "lblSegments(34)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "lblSegments(36)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label1"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "lblSegments(35)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "lblSegments(3)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "txtSegment(30)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "txtSegment(28)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "txtSegment(26)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "txtSegment(24)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "txtSegment(22)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "txtSegment(20)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "txtSegment(18)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "txtSegment(16)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "txtSegment(14)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "txtSegment(13)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "txtSegment(11)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "txtSegment(12)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "txtSegment(29)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "txtSegment(17)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "txtSegment(19)"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "txtSegment(21)"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "txtSegment(23)"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "txtSegment(25)"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "txtSegment(27)"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "txtSegment(15)"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).ControlCount=   36
      TabCaption(2)   =   "Calibration Factors"
      TabPicture(2)   =   "frmSegments.frx":09E2
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label3(1)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lblSegments(17)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "lblSegments(16)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "lblSegments(15)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "lblSegments(14)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "lblSegments(13)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "lblSegments(2)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Label3(0)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "lblSegments(21)"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "lblSegments(25)"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "lblSegments(26)"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "lblSegments(27)"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "lblSegments(29)"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "lblSegments(28)"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "txtSegment(33)"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "txtSegment(32)"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "txtSegment(31)"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "txtSegment(34)"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "txtSegment(35)"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "txtSegment(36)"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "txtSegment(37)"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "txtSegment(38)"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "txtSegment(39)"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).Control(23)=   "txtSegment(40)"
      Tab(2).Control(23).Enabled=   0   'False
      Tab(2).Control(24)=   "txtSegment(41)"
      Tab(2).Control(24).Enabled=   0   'False
      Tab(2).Control(25)=   "txtSegment(42)"
      Tab(2).Control(25).Enabled=   0   'False
      Tab(2).Control(26)=   "txtSegment(44)"
      Tab(2).Control(26).Enabled=   0   'False
      Tab(2).Control(27)=   "txtSegment(43)"
      Tab(2).Control(27).Enabled=   0   'False
      Tab(2).Control(28)=   "txtSegment(46)"
      Tab(2).Control(28).Enabled=   0   'False
      Tab(2).Control(29)=   "txtSegment(45)"
      Tab(2).Control(29).Enabled=   0   'False
      Tab(2).Control(30)=   "txtSegment(48)"
      Tab(2).Control(30).Enabled=   0   'False
      Tab(2).Control(31)=   "txtSegment(47)"
      Tab(2).Control(31).Enabled=   0   'False
      Tab(2).ControlCount=   32
      TabCaption(3)   =   "Internal Load"
      TabPicture(3)   =   "frmSegments.frx":09FE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtSegment(49)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "txtSegment(50)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "txtSegment(51)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "txtSegment(52)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "txtSegment(53)"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "txtSegment(54)"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "lblSegments(32)"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "lblSegments(31)"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "lblSegments(30)"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "lblSegments(18)"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "lblSegments(19)"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "lblSegments(23)"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "Label3(6)"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "Label3(7)"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).Control(14)=   "Label3(8)"
      Tab(3).Control(14).Enabled=   0   'False
      Tab(3).ControlCount=   15
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   15
         Left            =   -70920
         TabIndex        =   111
         Text            =   "0"
         Top             =   2505
         Width           =   1212
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   27
         Left            =   -70920
         TabIndex        =   110
         Text            =   "0"
         Top             =   5295
         Width           =   1212
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   25
         Left            =   -70920
         TabIndex        =   109
         Text            =   "0"
         Top             =   4830
         Width           =   1212
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   23
         Left            =   -70920
         TabIndex        =   108
         Text            =   "0"
         Top             =   4365
         Width           =   1212
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   21
         Left            =   -70920
         TabIndex        =   107
         Text            =   "0"
         Top             =   3900
         Width           =   1212
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   19
         Left            =   -70920
         TabIndex        =   106
         Text            =   "0"
         Top             =   3435
         Width           =   1212
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   17
         Left            =   -70920
         TabIndex        =   105
         Text            =   "0"
         Top             =   2970
         Width           =   1212
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         HelpContextID   =   49
         Index           =   29
         Left            =   -70920
         TabIndex        =   104
         Text            =   "0"
         Top             =   5760
         Width           =   1212
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   12
         Left            =   -69480
         TabIndex        =   103
         Text            =   "0"
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         HelpContextID   =   302
         Index           =   11
         Left            =   -70920
         TabIndex        =   102
         Text            =   "0"
         Top             =   1080
         Width           =   1212
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   13
         Left            =   -70920
         TabIndex        =   101
         Text            =   "0"
         Top             =   2040
         Width           =   1212
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         HelpContextID   =   304
         Index           =   47
         Left            =   3720
         TabIndex        =   88
         Text            =   "0"
         Top             =   5160
         Width           =   1212
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         HelpContextID   =   303
         Index           =   48
         Left            =   5160
         TabIndex        =   87
         Text            =   "0"
         Top             =   5160
         Width           =   855
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         HelpContextID   =   304
         Index           =   45
         Left            =   3720
         TabIndex        =   86
         Text            =   "0"
         Top             =   4680
         Width           =   1212
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         HelpContextID   =   303
         Index           =   46
         Left            =   5160
         TabIndex        =   85
         Text            =   "0"
         Top             =   4680
         Width           =   855
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         HelpContextID   =   304
         Index           =   43
         Left            =   3720
         TabIndex        =   80
         Text            =   "0"
         Top             =   4200
         Width           =   1212
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         HelpContextID   =   303
         Index           =   44
         Left            =   5160
         TabIndex        =   79
         Text            =   "0"
         Top             =   4200
         Width           =   855
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         HelpContextID   =   3
         Left            =   -72120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Segment Downstream of Current Segment ( 0 = Out of Reservoir )"
         Top             =   1440
         Width           =   3255
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         HelpContextID   =   41
         Index           =   49
         Left            =   -71520
         TabIndex        =   66
         Text            =   "0"
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         HelpContextID   =   303
         Index           =   50
         Left            =   -70320
         TabIndex        =   65
         Text            =   "0"
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         HelpContextID   =   41
         Index           =   51
         Left            =   -71520
         TabIndex        =   64
         Text            =   "0"
         Top             =   2730
         Width           =   855
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         HelpContextID   =   303
         Index           =   52
         Left            =   -70320
         TabIndex        =   63
         Text            =   "0"
         Top             =   2730
         Width           =   855
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         HelpContextID   =   41
         Index           =   53
         Left            =   -71520
         TabIndex        =   62
         Text            =   "0"
         Top             =   3300
         Width           =   855
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         HelpContextID   =   303
         Index           =   54
         Left            =   -70320
         TabIndex        =   61
         Text            =   "0"
         Top             =   3300
         Width           =   855
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         HelpContextID   =   303
         Index           =   42
         Left            =   5160
         TabIndex        =   52
         Text            =   "0"
         Top             =   3720
         Width           =   855
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         HelpContextID   =   304
         Index           =   41
         Left            =   3720
         TabIndex        =   51
         Text            =   "0"
         Top             =   3720
         Width           =   1212
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         HelpContextID   =   303
         Index           =   40
         Left            =   5160
         TabIndex        =   50
         Text            =   "0"
         Top             =   3240
         Width           =   855
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         HelpContextID   =   304
         Index           =   39
         Left            =   3720
         TabIndex        =   49
         Text            =   "0"
         Top             =   3240
         Width           =   1212
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         HelpContextID   =   303
         Index           =   38
         Left            =   5160
         TabIndex        =   48
         Text            =   "0"
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         HelpContextID   =   304
         Index           =   37
         Left            =   3720
         TabIndex        =   47
         Text            =   "0"
         Top             =   2760
         Width           =   1212
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         HelpContextID   =   303
         Index           =   36
         Left            =   5160
         TabIndex        =   46
         Text            =   "0"
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         HelpContextID   =   304
         Index           =   35
         Left            =   3720
         TabIndex        =   45
         Text            =   "0"
         Top             =   2280
         Width           =   1212
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         HelpContextID   =   303
         Index           =   34
         Left            =   5160
         TabIndex        =   44
         Text            =   "0"
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         HelpContextID   =   304
         Index           =   31
         Left            =   3720
         TabIndex        =   43
         Text            =   "0"
         Top             =   1320
         Width           =   1212
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         HelpContextID   =   303
         Index           =   32
         Left            =   5160
         TabIndex        =   42
         Text            =   "0"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         HelpContextID   =   304
         Index           =   33
         Left            =   3720
         TabIndex        =   41
         Text            =   "0"
         Top             =   1800
         Width           =   1212
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   14
         Left            =   -69480
         TabIndex        =   29
         Text            =   "0"
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   16
         Left            =   -69480
         TabIndex        =   28
         Text            =   "0"
         Top             =   2505
         Width           =   855
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   18
         Left            =   -69480
         TabIndex        =   27
         Text            =   "0"
         Top             =   2970
         Width           =   855
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   20
         Left            =   -69480
         TabIndex        =   26
         Text            =   "0"
         Top             =   3435
         Width           =   855
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   22
         Left            =   -69480
         TabIndex        =   25
         Text            =   "0"
         Top             =   3900
         Width           =   855
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   24
         Left            =   -69480
         TabIndex        =   24
         Text            =   "0"
         Top             =   4365
         Width           =   855
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   26
         Left            =   -69480
         TabIndex        =   23
         Text            =   "0"
         Top             =   4830
         Width           =   855
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   28
         Left            =   -69480
         TabIndex        =   22
         Text            =   "0"
         Top             =   5295
         Width           =   855
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   30
         Left            =   -69480
         TabIndex        =   21
         Text            =   "0"
         Top             =   5760
         Width           =   855
      End
      Begin VB.TextBox txtSegment 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   -72120
         TabIndex        =   1
         Text            =   "0"
         Top             =   960
         Width           =   3255
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         HelpContextID   =   63
         Index           =   3
         Left            =   -72120
         TabIndex        =   3
         Text            =   "0"
         Top             =   2040
         Width           =   732
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   4
         Left            =   -71280
         TabIndex        =   4
         Text            =   "0"
         Top             =   2895
         Width           =   1095
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   5
         Left            =   -71280
         TabIndex        =   5
         Text            =   "0"
         Top             =   3390
         Width           =   1095
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         HelpContextID   =   85
         Index           =   6
         Left            =   -71280
         TabIndex        =   6
         Text            =   "0"
         Top             =   3900
         Width           =   1095
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         HelpContextID   =   83
         Index           =   7
         Left            =   -71280
         TabIndex        =   11
         Text            =   "0"
         Top             =   4395
         Width           =   1095
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         HelpContextID   =   303
         Index           =   8
         Left            =   -69840
         TabIndex        =   10
         Text            =   "0"
         Top             =   4395
         Width           =   855
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         HelpContextID   =   84
         Index           =   9
         Left            =   -71280
         TabIndex        =   9
         Text            =   "0"
         Top             =   5400
         Width           =   1095
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         HelpContextID   =   303
         Index           =   10
         Left            =   -69840
         TabIndex        =   8
         Text            =   "0"
         Top             =   5400
         Width           =   855
      End
      Begin VB.TextBox txtSegment 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         HelpContextID   =   63
         Index           =   2
         Left            =   -71760
         TabIndex        =   75
         Text            =   "0"
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Label lblSegments 
         Alignment       =   1  'Right Justify
         Caption         =   "Non-Algal Turb. (1/m):"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   -74280
         TabIndex        =   113
         Top             =   1080
         Width           =   3255
      End
      Begin VB.Label lblSegments 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   35
         Left            =   -70920
         TabIndex        =   112
         ToolTipText     =   "Non-algal turbidity estimated from observed chl-a & secchi"
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "(min = 0.08/m)"
         Height          =   255
         Left            =   -68640
         TabIndex        =   100
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label lblSegments 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   36
         Left            =   -69480
         TabIndex        =   99
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label lblSegments 
         Alignment       =   1  'Right Justify
         Caption         =   "Non-Algal Turb Est. (1/Secchi - 0.025*Chl):"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   34
         Left            =   -74760
         TabIndex        =   98
         Top             =   1560
         Width           =   3735
      End
      Begin VB.Label lblSegments 
         Alignment       =   2  'Center
         Caption         =   "0.12"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   33
         Left            =   -69840
         TabIndex        =   97
         Top             =   4920
         Width           =   855
      End
      Begin VB.Label lblSegments 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   22
         Left            =   -71280
         TabIndex        =   96
         ToolTipText     =   "Mixed Layer Depth Predicted from Mean Depth"
         Top             =   4920
         Width           =   1095
      End
      Begin VB.Label lblSegments 
         Alignment       =   1  'Right Justify
         Caption         =   "Estimated Mixed Depth (m):"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   20
         Left            =   -74040
         TabIndex        =   95
         Top             =   4920
         Width           =   2655
      End
      Begin VB.Label lblSegments 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   32
         Left            =   -71520
         TabIndex        =   94
         Top             =   3840
         Width           =   855
      End
      Begin VB.Label lblSegments 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   31
         Left            =   -70320
         TabIndex        =   93
         Top             =   3840
         Width           =   855
      End
      Begin VB.Label lblSegments 
         Alignment       =   1  'Right Justify
         Caption         =   "Defaults:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   30
         Left            =   -74520
         TabIndex        =   92
         Top             =   3840
         Width           =   2895
      End
      Begin VB.Label lblSegments 
         Alignment       =   2  'Center
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   28
         Left            =   3720
         TabIndex        =   91
         Top             =   5640
         Width           =   1215
      End
      Begin VB.Label lblSegments 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   29
         Left            =   5160
         TabIndex        =   90
         Top             =   5640
         Width           =   855
      End
      Begin VB.Label lblSegments 
         Alignment       =   1  'Right Justify
         Caption         =   "Defaults:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   27
         Left            =   720
         TabIndex        =   89
         Top             =   5640
         Width           =   2895
      End
      Begin VB.Label lblSegments 
         Alignment       =   1  'Right Justify
         Caption         =   "Metalimnetic O2 Depletion:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   26
         Left            =   720
         TabIndex        =   84
         Top             =   5160
         Width           =   2895
      End
      Begin VB.Label lblSegments 
         Alignment       =   1  'Right Justify
         Caption         =   "Hypolimnetic O2 Depletion:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   25
         Left            =   720
         TabIndex        =   83
         Top             =   4680
         Width           =   2895
      End
      Begin VB.Label lblSegments 
         Alignment       =   1  'Right Justify
         Caption         =   "Particulate (Total P - Ortho P):"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   21
         Left            =   720
         TabIndex        =   81
         Top             =   4200
         Width           =   2895
      End
      Begin VB.Label lblSegment 
         Alignment       =   1  'Right Justify
         Caption         =   "Outflow Segment:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   -73920
         TabIndex        =   74
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label lblSegments 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Phosphorus:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   18
         Left            =   -74040
         TabIndex        =   72
         Top             =   2760
         Width           =   2415
      End
      Begin VB.Label lblSegments 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Nitrogen:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   19
         Left            =   -73920
         TabIndex        =   71
         Top             =   3330
         Width           =   2295
      End
      Begin VB.Label lblSegments 
         Alignment       =   1  'Right Justify
         Caption         =   "Conservative Substance:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   23
         Left            =   -74040
         TabIndex        =   70
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Mean"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   -71520
         TabIndex        =   69
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "CV"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   -70320
         TabIndex        =   68
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Units:  mg/m2-day"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   -72240
         TabIndex        =   67
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "CV"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   5160
         TabIndex        =   60
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblSegments 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Phosphorus:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   720
         TabIndex        =   59
         Top             =   1800
         Width           =   2895
      End
      Begin VB.Label lblSegments 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Nitrogen:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   13
         Left            =   720
         TabIndex        =   58
         Top             =   2280
         Width           =   2895
      End
      Begin VB.Label lblSegments 
         Alignment       =   1  'Right Justify
         Caption         =   "Chlorophyll-a:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   14
         Left            =   720
         TabIndex        =   57
         Top             =   2760
         Width           =   2895
      End
      Begin VB.Label lblSegments 
         Alignment       =   1  'Right Justify
         Caption         =   "Secchi Depth:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   15
         Left            =   720
         TabIndex        =   56
         Top             =   3240
         Width           =   2895
      End
      Begin VB.Label lblSegments 
         Alignment       =   1  'Right Justify
         Caption         =   "Organic Nitrogen:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   16
         Left            =   720
         TabIndex        =   55
         Top             =   3720
         Width           =   2895
      End
      Begin VB.Label lblSegments 
         Alignment       =   1  'Right Justify
         Caption         =   "Dispersion Rate:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   17
         Left            =   720
         TabIndex        =   54
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Mean"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   3840
         TabIndex        =   53
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblSegments 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Nitrogen  (ppb):"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   -73920
         TabIndex        =   40
         Top             =   2505
         Width           =   2895
      End
      Begin VB.Label lblSegments 
         Alignment       =   1  'Right Justify
         Caption         =   "Chlorophyll-a  (ppb):"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   -73920
         TabIndex        =   39
         Top             =   2970
         Width           =   2895
      End
      Begin VB.Label lblSegments 
         Alignment       =   1  'Right Justify
         Caption         =   "Secchi Depth (m):"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   -73920
         TabIndex        =   38
         Top             =   3435
         Width           =   2895
      End
      Begin VB.Label lblSegments 
         Alignment       =   1  'Right Justify
         Caption         =   "Organic Nitrogen (ppb):"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   -74280
         TabIndex        =   37
         Top             =   3900
         Width           =   3255
      End
      Begin VB.Label lblSegments 
         Alignment       =   1  'Right Justify
         Caption         =   "Total P - Ortho P (ppb):"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   9
         Left            =   -73920
         TabIndex        =   36
         Top             =   4365
         Width           =   2895
      End
      Begin VB.Label lblSegments 
         Alignment       =   1  'Right Justify
         Caption         =   "Hypolimnetic O2 Depletion (ppb/d):"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   10
         Left            =   -74520
         TabIndex        =   35
         Top             =   4830
         Width           =   3495
      End
      Begin VB.Label lblSegments 
         Alignment       =   1  'Right Justify
         Caption         =   "Metalimnetic O2 Depletion (ppb/d):"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   11
         Left            =   -74520
         TabIndex        =   34
         Top             =   5295
         Width           =   3495
      End
      Begin VB.Label lblSegments 
         Alignment       =   1  'Right Justify
         Caption         =   "Conservative Substance (ppb):"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   12
         Left            =   -74760
         TabIndex        =   33
         Top             =   5760
         Width           =   3735
      End
      Begin VB.Label lblSegments 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Phosphorus (ppb):"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   -73920
         TabIndex        =   32
         Top             =   2040
         Width           =   2895
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "CV"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   -69480
         TabIndex        =   31
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Mean"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   -70800
         TabIndex        =   30
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblSegment 
         Alignment       =   1  'Right Justify
         Caption         =   "Segment Name:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   -73920
         TabIndex        =   20
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblSegments 
         Alignment       =   1  'Right Justify
         Caption         =   "Mixed Layer Depth (m):"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   -74040
         TabIndex        =   19
         Top             =   4440
         Width           =   2655
      End
      Begin VB.Label lblSegments 
         Alignment       =   1  'Right Justify
         Caption         =   "Hypolimnetic Thickness (m):"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   -74160
         TabIndex        =   18
         Top             =   5400
         Width           =   2775
      End
      Begin VB.Label lblSegment 
         Alignment       =   1  'Right Justify
         Caption         =   "Segment Group:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   -73920
         TabIndex        =   17
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label lblSegment 
         Alignment       =   1  'Right Justify
         Caption         =   "Surface Area (km2):"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   -73560
         TabIndex        =   16
         Top             =   3000
         Width           =   2175
      End
      Begin VB.Label lblSegment 
         Alignment       =   1  'Right Justify
         Caption         =   "Mean Depth (m):"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   -73320
         TabIndex        =   15
         Top             =   3480
         Width           =   1935
      End
      Begin VB.Label lblSegment 
         Alignment       =   1  'Right Justify
         Caption         =   " Length (km):"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   -72840
         TabIndex        =   14
         Top             =   3960
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "CV"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   -69840
         TabIndex        =   13
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Mean"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   -71160
         TabIndex        =   12
         Top             =   2400
         Width           =   855
      End
   End
   Begin VB.TextBox txtSegment 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   1440
      TabIndex        =   76
      Text            =   "0"
      Top             =   840
      Width           =   732
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   77
      Top             =   0
      Width           =   8040
      _ExtentX        =   14182
      _ExtentY        =   1111
      ButtonWidth     =   1032
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "List"
            Key             =   "btnList"
            Object.ToolTipText     =   "List segment & tributary network"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add"
            Key             =   "btnAdd"
            Object.ToolTipText     =   "Add a new segment at end of the segment list"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Insert"
            Key             =   "btnInsert"
            Object.ToolTipText     =   "Insert a new segment after the selected one"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            Key             =   "btnDelete"
            Object.ToolTipText     =   "Delete selected segment"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Clear"
            Key             =   "bthClear"
            Object.ToolTipText     =   "Assign default values to all input values for selected segment"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Undo"
            Key             =   "btnUndo"
            Object.ToolTipText     =   "Restore initial values for all segments"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Help"
            Key             =   "btnHelp"
            Object.ToolTipText     =   "Get Help"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancel"
            Key             =   "btnCancel"
            Object.ToolTipText     =   "Ignore edits & return to menu"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "OK"
            Key             =   "btnOK"
            Object.ToolTipText     =   "Keep edits & return to program menu"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Label lblSegments 
      Alignment       =   1  'Right Justify
      Caption         =   "Hypol. O2 Depletion (ppb/day):"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   24
      Left            =   840
      TabIndex        =   82
      Top             =   6600
      Width           =   2895
   End
   Begin VB.Label lblCount 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   78
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label lblDefinitions 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   960
      TabIndex        =   73
      Top             =   1560
      Width           =   6015
   End
End
Attribute VB_Name = "frmSegments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private jseG As Integer
Private Const nboX As Integer = 54

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim s As String
Dim j As Long

Select Case Button

Case "List"
    UpdateSegmentValues (2)
    ContextId = frmMenu.mnuListNetwork.HelpContextID
    Call List_Tree
    ViewSheet ("Segment Network")

Case "Add"
    If Nseg = NSMAX - 1 Then
                MsgBox "Too many segments", vbExclamation
            Else
    If MsgBox("Add New Segment ?", vbYesNo) = vbYes Then
            UpdateSegmentValues (2)
            If CheckValues Then
                Call SegmentEdit(Nseg, 1)
                jseG = Nseg
                UpdateCombos
                UpdateSegmentValues (1)
                End If
            End If
      End If
            
Case "Insert"
    If Nseg = NSMAX - 1 Then
            MsgBox "Too many segments", vbExclamation
            Else
    If MsgBox("Insert New Segment ?", vbYesNo) = vbYes Then
            UpdateSegmentValues (2)
            If CheckValues Then
                Call SegmentEdit(jseG, 1)
                jseG = jseG + 1
                UpdateCombos
                UpdateSegmentValues (1)
                End If
            End If
            End If

Case "Delete"
    If Nseg <= 1 Then
            MsgBox "You can't delete the only segment", vbExclamation
            Else
    If MsgBox("Delete Segment " & jseG & " " & SegName(jseG) & " ?", vbYesNo) = vbYes Then
            Call SegmentEdit(jseG, 0)
            jseG = 1
            UpdateCombos
            UpdateSegmentValues (1)
            End If
            End If
            

Case "Clear"
    If MsgBox("Clear All Input Values for Selected Segment?", vbYesNo) = vbYes Then
            s = SegName(jseG)
            j = Iout(jseG)
            Call SegZero(jseG)
            Iout(jseG) = j
            SegName(jseG) = s
            UpdateCombos
            UpdateSegmentValues (1)
            End If

Case "Undo"
        jseG = 1
        Backup (1)
        UpdateCombos
        UpdateSegmentValues (1)

Case "Help"
        ShowHelp (HelpContextID)
        
Case "Cancel"
        Backup (1)
        Unload Me
        
Case "OK"
        UpdateSegmentValues (2)
        If CheckValues Then
            Icalc = 0
            Unload Me
            End If

End Select

End Sub
Private Sub Combo2_Click()
'select outflow segment
'store outflow segment number in hidden text box behind combo box

    'If Combo2.ListIndex = jseG Then
    '    MsgBox "The outflow segment cannot equal the selected segment", vbExclamation
    '    Else
'      txtSegment(2) = Combo2.ListIndex
    '    End If
        
End Sub

Private Sub Form_Load()
    jseG = 1
    UpdateCombos
    SetToolTips
    UpdateSegmentValues (1)
End Sub

Private Sub Combo1_click()
'load new segment
    j = Combo1.ListIndex + 1
    If j <> jseG Then
    UpdateSegmentValues (2)
    If CheckValues Then jseG = j
    Combo1.ListIndex = jseG - 1
    UpdateSegmentValues (1)
    End If
'    UpdateCombos
    
End Sub

Sub UpdateCombos()

lblCount.Caption = "Number of Segments = " & Nseg
    
    With Combo1
    .Clear
    For i = 1 To Nseg
        .AddItem Format(i, "00") & " " & SegName(i)
        Next i
    .ListIndex = jseG - 1
    End With
    
    With Combo2
    .Clear
    .AddItem "Out of Reservoir"
    For i = 1 To Nseg
       .AddItem Format(i, "00") & " " & SegName(i)
       Next i
      .ListIndex = Iout(jseG)
    End With
      
    End Sub

Sub UpdateSegmentValues(io)
'io=1 copy source values to temporary array
'io=2 copy from temporary array to original
'io=3 copy from text boxes to temporary array
'io=4 copy from temporary array to text boxes
'io=5 clear temporary array to default values

Select Case io

Case 1   'fill temporary array

txtSegment(0) = jseG         'segment number
txtSegment(1) = SegName(jseG)  'segment label
'txtSegment(2) = Iout(jseG)
txtSegment(3) = Iag(jseG)
txtSegment(4) = Area(jseG)
txtSegment(5) = Zmn(jseG)
txtSegment(6) = Slen(jseG)
txtSegment(7) = Zmxi(jseG)
txtSegment(8) = CvZmxi(jseG)
txtSegment(9) = Zhyp(jseG)
txtSegment(10) = CvZhyp(jseG)
txtSegment(11) = Turbi(jseG)
txtSegment(12) = CvTurbi(jseG)
lblSegments(22).Caption = Format(ZmixEst(Zmn(jseG)), "0.0")

k = 11
For i = 2 To 9
k = k + 2
    txtSegment(k) = Cobs(jseG, i)    'tp
    txtSegment(k + 1) = CvCobs(jseG, i)
    Next i
txtSegment(29) = Cobs(jseG, 1)      'conserv
txtSegment(30) = CvCobs(jseG, 1)

AssignTurb
k = 29
For i = 1 To 9
k = k + 2
    txtSegment(k) = Cal(jseG, i)
    txtSegment(k + 1) = CvCal(jseG, i)
    Next i
For i = 1 To 3
k = k + 2
    txtSegment(k) = InternalLoad(jseG, i)
    txtSegment(k + 1) = CvInternalLoad(jseG, i)
    Next i
Combo2.ListIndex = Iout(jseG)

'Changed = False

Case 2    'copy from temporary array back to original values


Combo1.List(jseG - 1) = Format(jseG, "00") & " " & txtSegment(1)
Iout(jseG) = Combo2.ListIndex
jseG = txtSegment(0)
SegName(jseG) = txtSegment(1)
'Iout(jseG) = txtSegment(2)
Iag(jseG) = txtSegment(3)
Area(jseG) = txtSegment(4)
Zmn(jseG) = txtSegment(5)
Slen(jseG) = txtSegment(6)
Zmxi(jseG) = txtSegment(7)
CvZmxi(jseG) = txtSegment(8)
Zhyp(jseG) = txtSegment(9)
CvZhyp(jseG) = txtSegment(10)

Turbi(jseG) = txtSegment(11)
CvTurbi(jseG) = txtSegment(12)
lblSegments(22).Caption = Format(ZmixEst(Zmn(jseG)), "0.0")

k = 11
For i = 2 To 9
k = k + 2
    Cobs(jseG, i) = txtSegment(k)
    CvCobs(jseG, i) = txtSegment(k + 1)
    Next i
Cobs(jseG, 1) = txtSegment(29)
CvCobs(jseG, 1) = txtSegment(30)
AssignTurb
k = 29
For i = 1 To 9
k = k + 2
    Cal(jseG, i) = txtSegment(k)
    CvCal(jseG, i) = txtSegment(k + 1)
    Next i
For i = 1 To 3
k = k + 2
    InternalLoad(jseG, i) = txtSegment(k)
    CvInternalLoad(jseG, i) = txtSegment(k + 1)
    Next i

Case 3    'set defaults
'    Changed = True
'    For j = 2 To nboX
'        txtSegment(j).Text = 0
'        Next j
'    txtSegment(2) = 0
'    txtSegment(3) = 1     'calibration factors
'    txtSegment(31) = 1
'    txtSegment(33) = 1
'    txtSegment(35) = 1
'    txtSegment(37) = 1
'    txtSegment(39) = 1
'    txtSegment(41) = 1
    
Case default

End Select

End Sub

Private Function CheckValues() As Boolean

CheckValues = False
If Iout(jseG) = jseG Then
        MsgBox "The outflow segment cannot equal the selected segment", vbExclamation
    ElseIf Area(jseG) <= 0 Or Slen(jseG) <= 0 Or Zmn(jseG) <= 0 Then
        MsgBox "Positive values required for area, length, & depth", vbExclamation
    Else
        CheckValues = True
    End If

End Function

Private Sub SetToolTips()
            For i = 3 To nboX
                txtSegment(i).TabIndex = i
                Next i

'tab 0 morphometry
             txtSegment(0).ToolTipText = "Segment Number"
             txtSegment(1).ToolTipText = "Name of the Segment"
             txtSegment(2).ToolTipText = "Outflow Segment Number (0 = Reservoir Outflow)"
             txtSegment(3).ToolTipText = "Segment Group Number (=1 if all segments are in same reservoir)"
             txtSegment(4).ToolTipText = "Surface area (km2)"
             txtSegment(5).ToolTipText = "Mean depth (meters)"
             txtSegment(6).ToolTipText = "Segment length (km2)"
             txtSegment(7).ToolTipText = "Mean Depth of Mixed Layer (meters)"
             txtSegment(8).ToolTipText = "Coefficient of Variation for Mixed Layer"
             txtSegment(9).ToolTipText = "Mean Hypolimnetic Depth (meters)"
             txtSegment(10).ToolTipText = "Coefficient of Variation for Hypolimnetic Depth"
'tab 1 observed water quality
             txtSegment(11).ToolTipText = "Non-Algal Turbidity = 1/secchi - .025x chla"
             txtSegment(12).ToolTipText = "Coefficient of Variation for Non-Agal Turbidity"
             txtSegment(13).ToolTipText = "Observed Mean Total Phosphorus (ppb)"
             txtSegment(14).ToolTipText = "Coefficient of Variation for Observed Mean Phosphorus"
             txtSegment(15).ToolTipText = "Observed Mean Total Nitrogen(ppb)"
             txtSegment(16).ToolTipText = "Coefficient of Variation for Observed Mean Nitrogen"
             txtSegment(17).ToolTipText = "Observed Mean Chlorophyll-A (ppb)"
             txtSegment(18).ToolTipText = "Coefficient of Variation for Observed Mean Chlorophyll-a"
             txtSegment(19).ToolTipText = "Observed Mean Secchi Depth (meters)"
             txtSegment(20).ToolTipText = "Coefficient of Variation for Observed Mean Secchi"
             txtSegment(21).ToolTipText = "Observed Mean Organic Nitrogen(ppb)"
             txtSegment(22).ToolTipText = "Coefficient of Variation for Observed Mean Organic Nitrogen"
             txtSegment(23).ToolTipText = "Observed Total P - Ortho P (ppb)"
             txtSegment(24).ToolTipText = "Coefficient of Variation for Observed Total P - Ortho P"
             txtSegment(25).ToolTipText = "Hypolimnetic Oxygen Depletion (mg/m3 - day)"
             txtSegment(26).ToolTipText = "Coefficient of Variation for Hypolimnetic Oxygen Depletion"
             txtSegment(27).ToolTipText = "Metalimnetic Oxygen Depletion (mg/m3 - day)"
             txtSegment(28).ToolTipText = "Coefficient of Variation for Metalimnetic Oxygen Depletion"
             txtSegment(29).ToolTipText = "Observed Mean Conservative Substance Conc (-)"
             txtSegment(30).ToolTipText = "Coefficient of Variation for Observed Mean Conservative Substance"

'tab 3 calibration factors
             txtSegment(31).ToolTipText = "Dispersion Calibration Factor (-)"
             txtSegment(32).ToolTipText = "Coefficient of Variation for Dispersion Calibration Factor"
             txtSegment(33).ToolTipText = "Calibration Factor for Total Phosphorus (1.0)"
             txtSegment(34).ToolTipText = "Coefficient of Variation for Phosphorus Calibration Factor"
             txtSegment(35).ToolTipText = "Calibration Factor for Total Nitrogen (1.0)"
             txtSegment(36).ToolTipText = "Coefficient of Variation for Nitrogen Calibration Factor"
             txtSegment(37).ToolTipText = "Calibration Factor for Chlorophyll-A (1.0)"
             txtSegment(38).ToolTipText = "Coefficient of Variation for Chlorophyll-A Calibration Factor"
             txtSegment(39).ToolTipText = "Calibration Factor for Secchi Depth (1.0)"
             txtSegment(40).ToolTipText = "Coefficient of Variation for Secchi Depth Calibration Factor"
             txtSegment(41).ToolTipText = "Calibration Factor for Organic Nitrogen (1.0)"
             txtSegment(42).ToolTipText = "Coefficient of Variation for Organic Nitrogen Calibration Factor"
             txtSegment(43).ToolTipText = "Calibration Factor for Total P - Ortho P (1.0)"
             txtSegment(44).ToolTipText = "Coefficient of Variation for Total P - Ortho P Calibration Factor"
             txtSegment(45).ToolTipText = "Calibration Factor for Hypolimnetic Oxygen Depletion (1.0)"
             txtSegment(46).ToolTipText = "Coefficient of Variation for Oxygen Depletion Calibration Factor"
             txtSegment(47).ToolTipText = "Calibration Factor for Metalimnetic Oxygen Depletion (1.0)"
             txtSegment(48).ToolTipText = "Coefficient of Variation for Oxygen Depletion Calibration Factor"

'tab4 - internal loads
             txtSegment(49).ToolTipText = "Internal Load of Conservative Substance (mg/m2-day)"
             txtSegment(50).ToolTipText = "Coefficient of Variation for Internal Load of Conservative Substance"
             txtSegment(51).ToolTipText = "Internal Load for Total Phosphorus (mg/m2-day)"
             txtSegment(52).ToolTipText = "Coefficient of Variation for Internal Load of Total Phosphorus"
             txtSegment(53).ToolTipText = "Internal Load for Total Nitrogen (mg/m2-day)"
             txtSegment(54).ToolTipText = "Coefficient of Variation for Internal Load of Total Nitrogen"
             
             
            End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblDefinitions.Caption = ""
End Sub

Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblDefinitions.Caption = ""
End Sub
Private Sub SSTab2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
     lblDefinitions.Caption = ""
End Sub
Private Sub SSTab3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblDefinitions.Caption = ""
End Sub
Private Sub SSTab4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblDefinitions.Caption = ""
End Sub

Private Sub txtSegment_Change(index As Integer)
    lblDefinitions.Caption = txtSegment(index).ToolTipText
    If index <> 1 Then
        If Not VerifyPositive(txtSegment(index).Text) Then txtSegment(index).Text = ""
        End If
End Sub

Private Sub txtSegment_lostfocus(index As Integer)
    If index <> 1 Then If txtSegment(index).Text = "" Then txtSegment(index).Text = 0
    If index = 5 Then lblSegments(22).Caption = Format(ZmixEst(Val(txtSegment(index).Text)), "0.0")
    If index > 16 And index < 21 Then AssignTurb
    
End Sub

Private Sub txtSegment_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    lblDefinitions.Caption = txtSegment(index).ToolTipText
    'WhatsThisHelp = txtSegment(Index).HelpContextID
   
End Sub

Sub AssignTurb()
'estimate non-algal turbidity
    x(1) = Val(txtSegment(17).Text)   'chla mean & cv
    x(2) = Val(txtSegment(18).Text)
    x(3) = Val(txtSegment(19).Text)   'secchi mean & cv
    x(4) = Val(txtSegment(20).Text)
    Call TurbEst(x(1), x(2), x(3), x(4), x(5), x(6))
    lblSegments(35).Caption = Format(x(5), "0.00")
    lblSegments(36).Caption = Format(x(6), "0.00")
End Sub
