VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Begin VB.Form frmTribs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Tributary Data"
   ClientHeight    =   9285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7005
   HelpContextID   =   4
   Icon            =   "frmTributaries.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9285
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   51
      Top             =   0
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   1111
      ButtonWidth     =   1032
      ButtonHeight    =   953
      Appearance      =   1
      HelpContextID   =   4
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
            Object.ToolTipText     =   "Add a new tributary at end of the current list"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Insert"
            Key             =   "btnInsert"
            Object.ToolTipText     =   "Insert a new tributary after the currently selected one"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            Key             =   "btnDelete"
            Object.ToolTipText     =   "Delete the selected tributary"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Clear"
            Key             =   "btnClear"
            Object.ToolTipText     =   "Assign default values to selected tributary"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Undo"
            Key             =   "btnUndo"
            Object.ToolTipText     =   "Restore initial values"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Help"
            Key             =   "btnHelp"
            Object.ToolTipText     =   "Get help"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancel"
            Key             =   "btnCancel"
            Object.ToolTipText     =   "Ignore edits & return to program menu"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "OK"
            Key             =   "btnOK"
            Object.ToolTipText     =   "Save edits & return to program menu"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6855
      HelpContextID   =   4
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   12091
      _Version        =   393216
      Tabs            =   2
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
      TabCaption(0)   =   "Monitored Inputs"
      TabPicture(0)   =   "frmTributaries.frx":09AA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblLabel(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblLabel(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblLabel(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblLabel(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblLabel(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblLabel(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblLabel(6)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblLabel(7)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblLabel(8)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblLabel(9)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblLabel(10)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblLabel(11)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtText(22)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmbTribSeg"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtText(0)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtText(1)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtText(2)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtText(3)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtText(4)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtText(5)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtText(6)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtText(7)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtText(8)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtText(9)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtText(10)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtText(11)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtText(12)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtText(13)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtText(23)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "cmbTribType"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).ControlCount=   30
      TabCaption(1)   =   "Land Uses"
      TabPicture(1)   =   "frmTributaries.frx":09C6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblLabel(13)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblLabel(14)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblLabel(15)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblLabel(12)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblLabel(17)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lblLabel(18)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lblLabel(19)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lblLabel(16)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "lblLabel(20)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "lblLabel(21)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txtText(14)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "txtText(15)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "txtText(16)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "txtText(17)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "txtText(18)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "txtText(19)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "txtText(20)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "txtText(21)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).ControlCount=   18
      Begin VB.ComboBox cmbTribType 
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
         HelpContextID   =   301
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Select Tributary Type"
         Top             =   1800
         Width           =   3495
      End
      Begin VB.TextBox txtText 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   23
         Left            =   2280
         TabIndex        =   49
         Text            =   "Text1"
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox txtText 
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
         Height          =   375
         Index           =   21
         Left            =   -71760
         TabIndex        =   45
         Text            =   "Text1"
         Top             =   5760
         Width           =   1215
      End
      Begin VB.TextBox txtText 
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
         Height          =   375
         Index           =   20
         Left            =   -71760
         TabIndex        =   43
         Text            =   "Text1"
         Top             =   5160
         Width           =   1215
      End
      Begin VB.TextBox txtText 
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
         Height          =   375
         Index           =   19
         Left            =   -71760
         TabIndex        =   40
         Text            =   "Text1"
         Top             =   4560
         Width           =   1215
      End
      Begin VB.TextBox txtText 
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
         Height          =   375
         Index           =   18
         Left            =   -71760
         TabIndex        =   39
         Text            =   "Text1"
         Top             =   3960
         Width           =   1215
      End
      Begin VB.TextBox txtText 
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
         Height          =   375
         Index           =   17
         Left            =   -71760
         TabIndex        =   37
         Text            =   "Text1"
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox txtText 
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
         Height          =   375
         Index           =   16
         Left            =   -71760
         TabIndex        =   35
         Text            =   "Text1"
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox txtText 
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
         Height          =   375
         Index           =   15
         Left            =   -71760
         TabIndex        =   32
         Text            =   "Text1"
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox txtText 
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
         Height          =   375
         Index           =   14
         Left            =   -71760
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox txtText 
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
         Left            =   2280
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   840
         Width           =   3495
      End
      Begin VB.TextBox txtText 
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
         Height          =   375
         HelpContextID   =   303
         Index           =   12
         Left            =   4920
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   6240
         Width           =   800
      End
      Begin VB.TextBox txtText 
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
         Height          =   375
         Index           =   11
         Left            =   3480
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   6240
         Width           =   1215
      End
      Begin VB.TextBox txtText 
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
         Height          =   375
         HelpContextID   =   303
         Index           =   10
         Left            =   4920
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   5664
         Width           =   800
      End
      Begin VB.TextBox txtText 
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
         Height          =   375
         Index           =   9
         Left            =   3480
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   5660
         Width           =   1215
      End
      Begin VB.TextBox txtText 
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
         Height          =   375
         HelpContextID   =   303
         Index           =   8
         Left            =   4920
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   5088
         Width           =   800
      End
      Begin VB.TextBox txtText 
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
         Height          =   375
         Index           =   7
         Left            =   3480
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   5080
         Width           =   1215
      End
      Begin VB.TextBox txtText 
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
         Height          =   375
         HelpContextID   =   303
         Index           =   6
         Left            =   4920
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   4512
         Width           =   800
      End
      Begin VB.TextBox txtText 
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
         Height          =   375
         Index           =   5
         Left            =   3480
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   4500
         Width           =   1215
      End
      Begin VB.TextBox txtText 
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
         Height          =   375
         HelpContextID   =   303
         Index           =   4
         Left            =   4920
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   3936
         Width           =   800
      End
      Begin VB.TextBox txtText 
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
         Height          =   375
         Index           =   3
         Left            =   3480
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   3920
         Width           =   1215
      End
      Begin VB.TextBox txtText 
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
         Height          =   375
         HelpContextID   =   303
         Index           =   2
         Left            =   4920
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   3360
         Width           =   800
      End
      Begin VB.TextBox txtText 
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
         Height          =   375
         Index           =   1
         Left            =   3480
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   3340
         Width           =   1215
      End
      Begin VB.TextBox txtText 
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
         Height          =   375
         Index           =   0
         Left            =   3480
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   2760
         Width           =   1215
      End
      Begin VB.ComboBox cmbTribSeg 
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
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Select Segment Associated with This Tributary"
         Top             =   1320
         Width           =   3495
      End
      Begin VB.TextBox txtText 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   22
         Left            =   2280
         TabIndex        =   48
         Text            =   "Text1"
         Top             =   1320
         Width           =   3255
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         Caption         =   " Landuse Category"
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
         Index           =   21
         Left            =   -73920
         TabIndex        =   50
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label lblLabel 
         Alignment       =   2  'Center
         Caption         =   " Drainage Area (km2)"
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
         Index           =   20
         Left            =   -72000
         TabIndex        =   47
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Category 5:"
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
         Left            =   -74640
         TabIndex        =   46
         Top             =   4080
         Width           =   2655
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Category 8:"
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
         Left            =   -74640
         TabIndex        =   44
         Top             =   5760
         Width           =   2655
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Category 7:"
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
         Left            =   -74640
         TabIndex        =   42
         Top             =   5160
         Width           =   2655
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Category 6:"
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
         Left            =   -74640
         TabIndex        =   41
         Top             =   4560
         Width           =   2655
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Category 1:"
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
         Index           =   12
         Left            =   -74640
         TabIndex        =   38
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Category 4:"
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
         Left            =   -74640
         TabIndex        =   36
         Top             =   3480
         Width           =   2655
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Category 3:"
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
         Left            =   -74640
         TabIndex        =   34
         Top             =   2880
         Width           =   2655
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Category 2:"
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
         Left            =   -74640
         TabIndex        =   33
         Top             =   2280
         Width           =   2655
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Tributary Name:"
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
         Index           =   11
         Left            =   360
         TabIndex        =   30
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Tributary Type:"
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
         Index           =   10
         Left            =   360
         TabIndex        =   29
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Segment:"
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
         Index           =   9
         Left            =   360
         TabIndex        =   28
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label lblLabel 
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
         Index           =   8
         Left            =   4920
         TabIndex        =   27
         ToolTipText     =   "Coefficient of Variation = Standard Error / Mean"
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label lblLabel 
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
         Index           =   7
         Left            =   3600
         TabIndex        =   26
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Conservative Subst. Conc (ppb):"
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
         Index           =   6
         Left            =   240
         TabIndex        =   23
         Top             =   6240
         Width           =   3015
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Inorganic N Conc (ppb):"
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
         Index           =   5
         Left            =   600
         TabIndex        =   20
         Top             =   5660
         Width           =   2655
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Total N Conc (ppb):"
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
         Index           =   4
         Left            =   600
         TabIndex        =   17
         Top             =   5080
         Width           =   2655
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Ortho P Conc (ppb):"
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
         Index           =   3
         Left            =   600
         TabIndex        =   15
         Top             =   4500
         Width           =   2655
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Total P Conc (ppb):"
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
         Left            =   600
         TabIndex        =   14
         Top             =   3920
         Width           =   2655
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Annual Flow Rate (hm3/yr):"
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
         Left            =   600
         TabIndex        =   9
         Top             =   3340
         Width           =   2655
      End
      Begin VB.Label lblLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Watershed Area (km2):"
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
         Left            =   600
         TabIndex        =   6
         Top             =   2760
         Width           =   2655
      End
   End
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
      HelpContextID   =   2
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Select Tributary to be Edited"
      Top             =   840
      Width           =   3375
   End
   Begin VB.Label lblCount 
      Caption         =   "Total Number of Tribs. ="
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
      Left            =   3840
      TabIndex        =   52
      Top             =   840
      Width           =   2775
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
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   1560
      Width           =   6375
   End
End
Attribute VB_Name = "frmTribs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private jtriB As Integer
Private Const nboX As Integer = 23

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button

Case "List"
    UpdateTribValues (2)
    ContextId = frmMenu.mnuListNetwork.HelpContextID
    Call List_Tree
    ViewSheet ("Segment Network")

Case "Add"
    If NTrib = NTMAX Then
            MsgBox "Too many tributaries", vbExclamation
    
    Else
    If MsgBox("Add New Tributary ?", vbYesNo) = vbYes Then
            UpdateTribValues (2)
            Call TribEdit(NTrib, 1)
            jtriB = NTrib
            UpdateCombo1
            UpdateTribValues (1)
            End If
    End If

Case "Insert"

   If NTrib = NTMAX Then
            MsgBox "Too many tributaries", vbExclamation
   Else
    If MsgBox("Insert New Tributary ?", vbYesNo) = vbYes Then
            UpdateTribValues (2)
            Call TribEdit(jtriB, 1)
            jtriB = jtriB + 1
            UpdateCombo1
            UpdateTribValues (1)
            End If
    End If
    
Case "Delete"
    If MsgBox("Delete Tributary " & TribName(jseG) & " ?", vbYesNo) = vbYes Then
            Call TribEdit(jtriB, 0)
            jtriB = 1
            UpdateCombo1
            UpdateTribValues (1)
            End If

Case "Clear"
    If MsgBox("Clear All Input Values for Selected Tributary?", vbYesNo) = vbYes Then
            s = TribName(jtriB)
            i = Iseg(jtriB)
            j = Icode(jtriB)
            Call TribZero(jtriB) 'exclude name, segment number, & type
            TribName(jtriB) = s
            Iseg(jtriB) = i
            Icode(jtriB) = j
            UpdateCombo1
            UpdateTribValues (1)
            End If

Case "Undo"
       Backup (1)
       jtriB = 1
       UpdateCombo1
       UpdateTribValues (1)
    
Case "Help"
        ShowHelp (HelpContextID)
        
Case "Cancel"
        Backup (1)
        Unload Me
        
Case "OK"
        UpdateTribValues (2)
        Icalc = 0
        Unload Me

End Select
End Sub

'Private Sub cmbTribSeg_Click()
'changed segment assignment
'    txtText(22) = cmbTribSeg.ListIndex
'End Sub

'Private Sub cmbTribType_Click()
'changed tributary type code
'    txtText(23) = cmbTribType.ListIndex + 1
'End Sub

Private Sub Form_Load()
    Backup (0)
    jtriB = 1
    UpdateCombo1
    UpdateSegBox
    UpdateTribTypeBox
    SetToolTips
    UpdateTribValues (1)

'land use category labels
    For i = 1 To NCAT
        lblLabel(i + 11).Caption = LandUseName(i)
        Next i
       
    End Sub

Sub UpdateSegBox()
    If Iseg(jtriB) = 0 Then MsgBox "segment=0"
    With cmbTribSeg
        .Clear
    .AddItem "No Segment - Trib. Ignored"
    For i = 1 To Nseg
        .AddItem Format(i, "00") & " " & SegName(i)
        Next i
    .ListIndex = Iseg(jtriB)
    End With
    End Sub

Sub UpdateTribTypeBox()
    With cmbTribType
        .Clear
    For i = 1 To N_Type_Codes
        .AddItem Format(i, "00") & " " & Type_Code(i)
        Next i
    .ListIndex = Icode(jtriB) - 1
    End With
    End Sub
 
Private Sub Combo1_click()
'load new Trib
    j = Combo1.ListIndex + 1
    If j <> jtriB Then UpdateTribValues (2)
    jtriB = j
    Combo1.ListIndex = jtriB - 1
    UpdateSegBox
    UpdateTribTypeBox
    UpdateTribValues (1)
End Sub

Sub UpdateCombo1()
    lblCount.Caption = "Number of Tributaries = " & NTrib
    With Combo1
    .Clear
    For i = 1 To NTrib
        .AddItem Format(i, "00") & " " & TribName(i)
        Next i
    .ListIndex = jtriB - 1
    End With
    End Sub

Sub UpdateTribValues(io)
'io=1 copy source values to txttextorary array
'io=2 copy from txttextorary array to original
'io=3 copy from text boxes to txttextorary array
'io=4 copy from txttextorary array to text boxes
'io=5 clear txttextorary array to default values

Select Case io

Case 1   'fill txttextorary array

txtText(13) = TribName(jtriB)
txtText(0) = Darea(jtriB)           'Trib number
txtText(1) = Flow(jtriB)    'Trib label
txtText(2) = CvFlow(jtriB)
txtText(3) = Conci(jtriB, 2)
txtText(4) = CvCi(jtriB, 2)
txtText(5) = Conci(jtriB, 4)
txtText(6) = CvCi(jtriB, 4)
txtText(7) = Conci(jtriB, 3)
txtText(8) = CvCi(jtriB, 3)
txtText(9) = Conci(jtriB, 5)
txtText(10) = CvCi(jtriB, 5)
txtText(11) = Conci(jtriB, 1)
txtText(12) = CvCi(jtriB, 1)

txtText(14) = Warea(jtriB, 1)
txtText(15) = Warea(jtriB, 2)
txtText(16) = Warea(jtriB, 3)
txtText(17) = Warea(jtriB, 4)
txtText(18) = Warea(jtriB, 5)
txtText(19) = Warea(jtriB, 6)
txtText(20) = Warea(jtriB, 7)
txtText(21) = Warea(jtriB, 8)
'txtText(22) = Iseg(jtriB)
'txtText(23) = Icode(jtriB)
cmbTribSeg.ListIndex = Iseg(jtriB)
cmbTribType.ListIndex = Icode(jtriB) - 1

'Changed = False

Case 2   'copy from text array back to original values

Combo1.List(jtriB - 1) = Format(jtriB - 1, "00") & " " & txtText(13)
TribName(jtriB) = txtText(13)
Darea(jtriB) = txtText(0)
Flow(jtriB) = txtText(1)
CvFlow(jtriB) = txtText(2)
Conci(jtriB, 2) = txtText(3)
CvCi(jtriB, 2) = txtText(4)
Conci(jtriB, 4) = txtText(5)
CvCi(jtriB, 4) = txtText(6)
Conci(jtriB, 3) = txtText(7)
CvCi(jtriB, 3) = txtText(8)
Conci(jtriB, 5) = txtText(9)
CvCi(jtriB, 5) = txtText(10)
Conci(jtriB, 1) = txtText(11)
CvCi(jtriB, 1) = txtText(12)

Warea(jtriB, 1) = txtText(14)
Warea(jtriB, 2) = txtText(15)
Warea(jtriB, 3) = txtText(16)
Warea(jtriB, 4) = txtText(17)
Warea(jtriB, 5) = txtText(18)
Warea(jtriB, 6) = txtText(19)
Warea(jtriB, 7) = txtText(20)
Warea(jtriB, 8) = txtText(21)

'Iseg(jtriB) = txtText(22)
'Icode(jtriB) = txtText(23)
Iseg(jtriB) = cmbTribSeg.ListIndex
Icode(jtriB) = cmbTribType.ListIndex + 1

Case 3    'set defaults
 '   For j = 0 To nboX
 '       If j <> 13 And j <> 22 And j <> 23 Then txtText(j) = 0
 '       Next j
 '   'txtText(22) = 1
 '   'txtText(23) = 1
 '   Changed = True
        
Case default

End Select

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblDefinitions.Caption = ""
End Sub

Private Sub SetToolTips()
        
        For i = 0 To 21
            txtText(i).TabIndex = i
            Next i

'tab 0 monitored inputs
             txtText(13).ToolTipText = "Tributary Name"
             txtText(22).ToolTipText = "Segment Name"
             txtText(23).ToolTipText = "Tributary Type"
                          
             txtText(0).ToolTipText = "Total Drainage Area (km2)"
             txtText(1).ToolTipText = "Flow (hm3/yr)"
             txtText(2).ToolTipText = "Coefficient of Variation for Total Flow"
             txtText(3).ToolTipText = "Total Phosphorus Conc (ppb)"
             txtText(4).ToolTipText = "Coefficient of Variation for Total P Concentration"
             
             txtText(5).ToolTipText = "Ortho Phosphorus Conc (ppb)"
             txtText(6).ToolTipText = "Coefficient of Variation for Ortho P Concentration"
             txtText(7).ToolTipText = "Total Nitrogen Conc (ppb)"
             txtText(8).ToolTipText = "Coefficient of Variation for Total P Nitrogen"
             txtText(9).ToolTipText = "Inorganic Nitrogen Conc (ppb)"
             txtText(10).ToolTipText = "Coefficient of Variation for Inorganic N Concentration"
             txtText(11).ToolTipText = "Conservative Substance Conc (ppb)"
             txtText(12).ToolTipText = "Coefficient of Variation for Conservative Substance Concentration"
             
'tab 1 nonpoint watersheds
             txtText(14).ToolTipText = "Drainage Area in LandUse Category 1"
             txtText(15).ToolTipText = "Drainage Area in LandUse Category 2"
             txtText(16).ToolTipText = "Drainage Area in LandUse Category 3"
             txtText(17).ToolTipText = "Drainage Area in LandUse Category 4"
             txtText(18).ToolTipText = "Drainage Area in LandUse Category 5"
             txtText(19).ToolTipText = "Drainage Area in LandUse Category 6"
             txtText(20).ToolTipText = "Drainage Area in LandUse Category 7"
             txtText(21).ToolTipText = "Drainage Area in LandUse Category 8"
            
            End Sub


Private Sub txtText_Change(index As Integer)

    lblDefinitions.Caption = txtText(index).ToolTipText

'allow numeric entries only
    If index <> 13 And index <> 22 And index <> 23 Then
        If Not VerifyPositive(txtText(index).Text) Then txtText(index).Text = ""
        End If
       
End Sub
Private Sub txtText_lostfocus(index As Integer)

    If index <> 13 And index <> 22 And index <> 23 Then
        If txtText(index).Text = "" Then txtText(index).Text = 0
        End If
       
End Sub

Private Sub txtText_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    lblDefinitions.Caption = txtText(index).ToolTipText
End Sub

Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 lblDefinitions.Caption = ""
End Sub
Private Sub SSTab2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 lblDefinitions.Caption = ""
End Sub

