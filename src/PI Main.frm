VERSION 5.00
Object = "{33337113-F789-11CE-86F8-0020AFD8C6DB}#1.0#0"; "ipport40.ocx"
Object = "{33337143-F789-11CE-86F8-0020AFD8C6DB}#1.0#0"; "netcod40.ocx"
Object = "{33337153-F789-11CE-86F8-0020AFD8C6DB}#1.0#0"; "ipinfo40.ocx"
Object = "{33337183-F789-11CE-86F8-0020AFD8C6DB}#1.0#0"; "http40.ocx"
Object = "{33337283-F789-11CE-86F8-0020AFD8C6DB}#1.0#0"; "mime40.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmPI 
   AutoRedraw      =   -1  'True
   Caption         =   "Private Idaho 32  Version 5"
   ClientHeight    =   7065
   ClientLeft      =   405
   ClientTop       =   825
   ClientWidth     =   11265
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "PI Main.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7065
   ScaleWidth      =   11265
   Begin ComctlLib.ListView lvwAttachments 
      Height          =   1125
      Left            =   90
      TabIndex        =   21
      Top             =   4530
      Visible         =   0   'False
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   1984
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      OLEDropMode     =   1
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   0
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   735
      Left            =   120
      TabIndex        =   37
      Top             =   3720
      Visible         =   0   'False
      Width           =   8415
      ExtentX         =   14843
      ExtentY         =   1296
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10320
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   9360
      Top             =   4680
   End
   Begin RichTextLib.RichTextBox txtCC 
      Height          =   315
      Left            =   870
      TabIndex        =   35
      Top             =   1080
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   556
      _Version        =   393217
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"PI Main.frx":030A
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
   Begin RichTextLib.RichTextBox txtTo 
      Height          =   315
      Left            =   870
      TabIndex        =   34
      ToolTipText     =   "Double Click to see properties"
      Top             =   540
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   556
      _Version        =   393217
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"PI Main.frx":038C
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
   Begin VB.TextBox txtCCAddresses 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
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
      Height          =   255
      Left            =   8220
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   33
      Text            =   "PI Main.frx":040E
      Top             =   1110
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.TextBox txtToAddresses 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
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
      Height          =   255
      Left            =   8280
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   32
      Text            =   "PI Main.frx":041F
      Top             =   510
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.PictureBox picGenericIcon 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   9300
      Picture         =   "PI Main.frx":0430
      ScaleHeight     =   32
      ScaleMode       =   0  'User
      ScaleWidth      =   32
      TabIndex        =   25
      Top             =   3150
      Visible         =   0   'False
      Width           =   495
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   8040
      TabIndex        =   24
      Top             =   6780
      Visible         =   0   'False
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.TextBox txtsubject 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   870
      TabIndex        =   23
      Top             =   1440
      Width           =   6705
   End
   Begin VB.CommandButton btnCC 
      Caption         =   "Cc..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   90
      TabIndex        =   19
      Top             =   1140
      Width           =   615
   End
   Begin VB.CommandButton btnTo 
      Caption         =   "To..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   540
      Width           =   615
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   6690
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   10583
            MinWidth        =   10583
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   5292
            MinWidth        =   5292
            Picture         =   "PI Main.frx":0872
            Text            =   "Attachment Encryption Options"
            TextSave        =   "Attachment Encryption Options"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Encryption Options for Attachment"
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   5292
            MinWidth        =   5292
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Selected Remailer"
         EndProperty
      EndProperty
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
   Begin Threed.SSPanel SSPanel1 
      Align           =   1  'Align Top
      Height          =   465
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   820
      _Version        =   131074
      PictureBackgroundStyle=   2
      PictureBackground=   "PI Main.frx":0984
      Begin VB.CommandButton Command1 
         Caption         =   "To..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10620
         TabIndex        =   36
         Top             =   30
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox cmbRemailerSelect 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "PI Main.frx":112AE
         Left            =   6630
         List            =   "PI Main.frx":112B0
         TabIndex        =   14
         Text            =   "cmbRemailerSelect"
         Top             =   60
         Width           =   3405
      End
      Begin Threed.SSRibbon SSRibbon1 
         Height          =   345
         Index           =   13
         Left            =   5040
         TabIndex        =   38
         ToolTipText     =   "Display HTML content"
         Top             =   60
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "PI Main.frx":112B2
         ButtonStyle     =   3
      End
      Begin Threed.SSRibbon SSRibbon1 
         Height          =   345
         Index           =   12
         Left            =   4230
         TabIndex        =   26
         ToolTipText     =   "Forward Message"
         Top             =   60
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         Enabled         =   0   'False
         PictureUseMask  =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "PI Main.frx":114CA
         Alignment       =   4
         ButtonStyle     =   3
         PictureAlignment=   1
      End
      Begin Threed.SSRibbon SSRibbon1 
         Height          =   345
         Index           =   11
         Left            =   4710
         TabIndex        =   17
         ToolTipText     =   "Decode a MIME message"
         Top             =   60
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "PI Main.frx":115DC
         ButtonStyle     =   3
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Send Options: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5610
         TabIndex        =   15
         Top             =   120
         Width           =   1035
      End
      Begin Threed.SSRibbon SSRibbon1 
         Height          =   345
         Index           =   20
         Left            =   9960
         TabIndex        =   13
         ToolTipText     =   "Send the message"
         Top             =   0
         Visible         =   0   'False
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "PI Main.frx":117FE
         ButtonStyle     =   3
      End
      Begin Threed.SSRibbon SSRibbon1 
         Height          =   345
         Index           =   21
         Left            =   10230
         TabIndex        =   12
         ToolTipText     =   "Reply to sender"
         Top             =   30
         Visible         =   0   'False
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "PI Main.frx":11D04
         ButtonStyle     =   3
      End
      Begin Threed.SSRibbon SSRibbon1 
         Height          =   345
         Index           =   10
         Left            =   3810
         TabIndex        =   11
         ToolTipText     =   "Reply to sender"
         Top             =   60
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         Enabled         =   0   'False
         PictureUseMask  =   -1  'True
         Picture         =   "PI Main.frx":11E6C
         ButtonStyle     =   3
      End
      Begin Threed.SSRibbon SSRibbon1 
         Height          =   345
         Index           =   9
         Left            =   3390
         TabIndex        =   10
         ToolTipText     =   "Encrypt the message"
         Top             =   60
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "PI Main.frx":11FAA
         ButtonStyle     =   3
      End
      Begin Threed.SSRibbon SSRibbon1 
         Height          =   345
         Index           =   7
         Left            =   3030
         TabIndex        =   9
         ToolTipText     =   "Add an attachment"
         Top             =   60
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "PI Main.frx":120F6
         ButtonStyle     =   3
      End
      Begin Threed.SSRibbon SSRibbon1 
         Height          =   345
         Index           =   6
         Left            =   2670
         TabIndex        =   8
         ToolTipText     =   "Prepare Remailer Message"
         Top             =   60
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "PI Main.frx":1223E
         ButtonStyle     =   3
      End
      Begin Threed.SSRibbon SSRibbon1 
         Height          =   345
         Index           =   0
         Left            =   0
         TabIndex        =   7
         ToolTipText     =   "Send the message"
         Top             =   60
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "PI Main.frx":12366
         ButtonStyle     =   3
      End
      Begin Threed.SSRibbon SSRibbon1 
         Height          =   345
         Index           =   5
         Left            =   420
         TabIndex        =   6
         ToolTipText     =   "Save message"
         Top             =   60
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "PI Main.frx":1286C
         ButtonStyle     =   3
      End
      Begin Threed.SSRibbon SSRibbon1 
         Height          =   345
         Index           =   4
         Left            =   1200
         TabIndex        =   5
         ToolTipText     =   "Open Message"
         Top             =   45
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "PI Main.frx":12BBE
         ButtonStyle     =   3
      End
      Begin Threed.SSRibbon SSRibbon1 
         Height          =   345
         Index           =   3
         Left            =   870
         TabIndex        =   4
         ToolTipText     =   "New Message"
         Top             =   60
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "PI Main.frx":12F10
         ButtonStyle     =   3
      End
      Begin Threed.SSRibbon SSRibbon1 
         Height          =   345
         Index           =   2
         Left            =   2250
         TabIndex        =   3
         ToolTipText     =   "Paste"
         Top             =   60
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "PI Main.frx":13262
         ButtonStyle     =   3
      End
      Begin Threed.SSRibbon SSRibbon1 
         Height          =   345
         Index           =   1
         Left            =   1920
         TabIndex        =   2
         ToolTipText     =   "Copy"
         Top             =   60
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "PI Main.frx":135B4
         ButtonStyle     =   3
      End
      Begin Threed.SSRibbon SSRibbon1 
         Height          =   345
         Index           =   8
         Left            =   1590
         TabIndex        =   1
         ToolTipText     =   "Cut"
         Top             =   60
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         _Version        =   131074
         PictureFrames   =   1
         BackStyle       =   1
         PictureUseMask  =   -1  'True
         Picture         =   "PI Main.frx":13906
         ButtonStyle     =   3
      End
   End
   Begin RichTextLib.RichTextBox MessageArea 
      DragIcon        =   "PI Main.frx":13C58
      Height          =   1980
      Left            =   120
      TabIndex        =   20
      Top             =   1800
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   3493
      _Version        =   393217
      ScrollBars      =   3
      Appearance      =   0
      OLEDropMode     =   1
      TextRTF         =   $"PI Main.frx":13F62
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
   Begin MIMELibCtl.MIME MIME1 
      Left            =   9960
      Top             =   3840
      Boundary        =   ""
      ContentType     =   ""
      ContentTypeAttr =   ""
      Message         =   ""
      MessageHeaders  =   ""
   End
   Begin VB.Label lblFrom 
      BackStyle       =   0  'Transparent
      Caption         =   "lblFrom"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   870
      TabIndex        =   31
      Top             =   840
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Label lblFrom 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "From: "
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
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   30
      Top             =   840
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label lblcc 
      BackStyle       =   0  'Transparent
      Caption         =   "Cc: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   29
      Top             =   1140
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label lblSubject 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject"
      Height          =   195
      Left            =   7740
      TabIndex        =   28
      Top             =   1470
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label lblTo 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7650
      TabIndex        =   27
      Top             =   540
      Visible         =   0   'False
      Width           =   525
   End
   Begin ComctlLib.ImageList imglstLarge 
      Left            =   9300
      Top             =   1590
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PI Main.frx":13FE4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imglstSmall 
      Left            =   9300
      Top             =   2310
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PI Main.frx":142FE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Subject:"
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
      Height          =   240
      Index           =   0
      Left            =   90
      TabIndex        =   22
      Top             =   1440
      Width           =   720
   End
   Begin HTTPLibCtl.HTTP HTTP1 
      Left            =   10680
      Top             =   3750
      Accept          =   ""
      LocalFile       =   ""
      Password        =   ""
      ProxyPort       =   80
      ProxyServer     =   ""
      URL             =   ""
      User            =   ""
      UserAgent       =   "devSoft's HTTP Control"
      WinsockLoaded   =   -1  'True
   End
   Begin IPINFOLibCtl.IPInfo IPInfo1 
      Left            =   10590
      Top             =   3150
      PendingRequests =   0
      ServiceName     =   ""
      ServicePort     =   0
      ServiceProtocol =   ""
      WinsockLoaded   =   -1  'True
   End
   Begin NETCODELibCtl.NetCode NetCode1 
      Left            =   10620
      Top             =   2520
      DecodedData     =   ""
      EncodedData     =   ""
      FileName        =   ""
      Format          =   0
      IntelliCode     =   -1  'True
      MaxFileSize     =   0
      Mode            =   "0755"
      Overwrite       =   0   'False
      ProgressStep    =   1
   End
   Begin IPPORTLibCtl.IPPort IPPort1 
      Left            =   10620
      Top             =   1920
      EOL             =   ""
      InBufferSize    =   2048
      KeepAlive       =   0   'False
      Linger          =   -1  'True
      LocalPort       =   0
      MaxLineLength   =   2048
      OutBufferSize   =   2048
      RemoteHost      =   ""
      RemotePort      =   0
      WinsockLoaded   =   -1  'True
   End
   Begin VB.Menu FileMenu 
      Caption         =   "&File"
      Begin VB.Menu mFile_Save 
         Caption         =   "Save"
      End
      Begin VB.Menu FileExport 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu FileNul2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveAttachmentAs 
         Caption         =   "Save Attachment As"
      End
      Begin VB.Menu mAttachmentsNull 
         Caption         =   "-"
      End
      Begin VB.Menu mEncodeFile 
         Caption         =   "E&ncode a File into Message Area"
      End
      Begin VB.Menu mDecodeFile 
         Caption         =   "D&ecode Data in Message Area"
      End
      Begin VB.Menu mFile_Import 
         Caption         =   "&Import File or Message"
      End
      Begin VB.Menu mAttachfile 
         Caption         =   "Attach a File or Object"
      End
      Begin VB.Menu FileNull1 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mSplitFile 
         Caption         =   "Split a File"
      End
      Begin VB.Menu mMergeFile 
         Caption         =   "Merge a Split File"
      End
      Begin VB.Menu FileNull 
         Caption         =   "-"
      End
      Begin VB.Menu FileAddress 
         Caption         =   "&Address book..."
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu FileSave 
         Caption         =   "&Save settings"
      End
      Begin VB.Menu mSetMixmasterPath 
         Caption         =   "Set Mixmaster Path"
      End
      Begin VB.Menu mSetPGPKeysPath 
         Caption         =   "Set PGPKeys/Tools Path"
      End
      Begin VB.Menu mUsePGP 
         Caption         =   "Disable Utilities and PGP"
      End
      Begin VB.Menu FileNull2 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu PrintSetup 
         Caption         =   "Print Setup..."
      End
      Begin VB.Menu FilePrintM 
         Caption         =   "Print Message"
      End
      Begin VB.Menu FilePage 
         Caption         =   "Page setup..."
         Visible         =   0   'False
      End
      Begin VB.Menu FileNull3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu FileConnect 
         Caption         =   "&Connect"
         Visible         =   0   'False
      End
      Begin VB.Menu FileDisconnect 
         Caption         =   "&Disconnect"
         Visible         =   0   'False
      End
      Begin VB.Menu FileNull4 
         Caption         =   "-"
      End
      Begin VB.Menu FileExit 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu EditMenu 
      Caption         =   "&Edit"
      Begin VB.Menu EditCut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu EditCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu EditPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu EditSep 
         Caption         =   "-"
      End
      Begin VB.Menu EditClrAll 
         Caption         =   "Clea&r all"
         Shortcut        =   ^R
      End
      Begin VB.Menu EditClrMsg 
         Caption         =   "C&lear message"
         Shortcut        =   ^L
      End
      Begin VB.Menu EditCopyMsg 
         Caption         =   "C&opy message"
         Shortcut        =   ^O
      End
      Begin VB.Menu EditPasteMsg 
         Caption         =   "Paste &message"
         Shortcut        =   ^M
      End
      Begin VB.Menu EditSelectAll 
         Caption         =   "&Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu EditSetFont 
         Caption         =   "&Font"
         Shortcut        =   ^F
      End
      Begin VB.Menu EditSep1 
         Caption         =   "-"
      End
      Begin VB.Menu EditHeader 
         Caption         =   "Insert &header"
         Shortcut        =   ^H
      End
      Begin VB.Menu EditSig 
         Caption         =   "Insert s&ignature"
         Shortcut        =   ^I
      End
      Begin VB.Menu EditSep3 
         Caption         =   "-"
      End
      Begin VB.Menu EditOptions 
         Caption         =   "Options..."
      End
   End
   Begin VB.Menu mView 
      Caption         =   "&View"
      Begin VB.Menu mLargeAttachmentIcons 
         Caption         =   "Attachments with Large Icons"
      End
      Begin VB.Menu mSmallAttachmentIcons 
         Caption         =   "Attachments with List View "
      End
      Begin VB.Menu ViewSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mViewCypherPunksRemailerList 
         Caption         =   "View CypherPunks Remailer  List"
      End
      Begin VB.Menu mViewMixmasterRemailerList 
         Caption         =   "View Mixmaster Remailer  List"
      End
   End
   Begin VB.Menu mPGP 
      Caption         =   "&PGP"
      Begin VB.Menu KeyMenu 
         Caption         =   "&Keys"
         WindowList      =   -1  'True
         Begin VB.Menu KeyCreate 
            Caption         =   "&Create key pair"
         End
         Begin VB.Menu KeyCertify 
            Caption         =   "&Sign or Certify a Key"
         End
         Begin VB.Menu KeyEditTrust 
            Caption         =   "Change level of &trust in a key"
         End
         Begin VB.Menu keySep 
            Caption         =   "-"
         End
         Begin VB.Menu PGPDeleteKey 
            Caption         =   "&Delete key..."
         End
         Begin VB.Menu PGPAddKey 
            Caption         =   "&Import a key/keys from message"
         End
         Begin VB.Menu KeySep1 
            Caption         =   "-"
         End
         Begin VB.Menu PGPInsertKey 
            Caption         =   "&Insert key in message..."
         End
         Begin VB.Menu mKeyRingIDs 
            Caption         =   "&View Keys on Keyring"
         End
         Begin VB.Menu keySep3 
            Caption         =   "-"
         End
         Begin VB.Menu KeySubmit 
            Caption         =   "Submit &key to server"
         End
         Begin VB.Menu PGPGetKey 
            Caption         =   "&Get key from server"
         End
         Begin VB.Menu SelectKeyServer 
            Caption         =   "Select a Key Server"
         End
         Begin VB.Menu keySep4 
            Caption         =   "-"
         End
         Begin VB.Menu KeyOptions 
            Caption         =   "&Options..."
         End
      End
      Begin VB.Menu mSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu PGPEncrypt 
         Caption         =   "&Encrypt message"
         Shortcut        =   ^E
      End
      Begin VB.Menu PGPEncryptToSelf 
         Caption         =   "Encrypt to &Self"
         Shortcut        =   ^S
      End
      Begin VB.Menu PGPEnSign 
         Caption         =   "Encrypt and &sign message"
      End
      Begin VB.Menu PGPInSertDetachedSignature 
         Caption         =   "Create a  &Detached Signature"
         Shortcut        =   ^Z
      End
      Begin VB.Menu PGPClearSign 
         Caption         =   "&Clear sign message"
         Shortcut        =   ^G
      End
      Begin VB.Menu PGPDecrypt 
         Caption         =   "Decrypt or &Verify message"
         Shortcut        =   ^D
      End
      Begin VB.Menu mPGPEstimatePassPhrase 
         Caption         =   "Estimate Quality of Passphrase "
      End
      Begin VB.Menu mAnalyseMessage 
         Caption         =   "Analyse Message"
      End
      Begin VB.Menu pgpSEP 
         Caption         =   "Encryption Options"
         Begin VB.Menu PGPMultiple 
            Caption         =   "Use multiple keys"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu PGPSelf 
            Caption         =   "Encrypt to self"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu PGPEyes 
            Caption         =   "Eyes only"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu PGPConvent 
            Caption         =   "Use Conventional Encryption"
         End
         Begin VB.Menu mAdvancedEncryption 
            Caption         =   "Advanced Algorithms"
         End
         Begin VB.Menu PGPObscurity 
            Caption         =   "Obscurity"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu PGPWrap 
            Caption         =   "Word wrap on encrypt/sign"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu PGPFile 
            Caption         =   "File Operations"
         End
         Begin VB.Menu mAttachmentEncryptionOptions 
            Caption         =   "Attachment Encryption Options"
            Begin VB.Menu mDontEncryptAttachment 
               Caption         =   "No Encryption"
               Enabled         =   0   'False
               Visible         =   0   'False
            End
            Begin VB.Menu mEncryptAttachmentWithKey 
               Caption         =   "Encrypt with Key"
               Checked         =   -1  'True
            End
            Begin VB.Menu mConventionallyEncryptAttachment 
               Caption         =   "Conventionally Encrypt"
            End
         End
      End
      Begin VB.Menu PGPMin 
         Caption         =   "Run PGP minimized"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu PGPVerify 
         Caption         =   "&Verify PGP Distribution"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu PGPSep3 
         Caption         =   "-"
      End
      Begin VB.Menu PGPOptions 
         Caption         =   "PGP &Options..."
      End
   End
   Begin VB.Menu mNewsgroups 
      Caption         =   "&News"
      Begin VB.Menu USENETGate 
         Caption         =   "mail2news"
         Begin VB.Menu Prepare_Usenet_Nym 
            Caption         =   "Prepare Mail2News Message using Nym"
         End
         Begin VB.Menu Prepare_usenet_standard 
            Caption         =   "Prepare Mail2News Message"
         End
      End
      Begin VB.Menu mGetNews 
         Caption         =   "&Newsgroup Poster"
      End
      Begin VB.Menu NewsNull2 
         Caption         =   "-"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mFingerOps 
      Caption         =   "&Finger Operations"
      Begin VB.Menu GetFinger 
         Caption         =   "Finger"
      End
      Begin VB.Menu PI_Test_Click 
         Caption         =   "Special Test"
         Visible         =   0   'False
      End
      Begin VB.Menu DoConnect 
         Caption         =   "Connect"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mMessage 
      Caption         =   "&Message"
      Begin VB.Menu TransferReply 
         Caption         =   "&Insert reply markers"
      End
      Begin VB.Menu mTransferXHeaders 
         Caption         =   "&Mail-Headers..."
         Begin VB.Menu mAddMailHeaders 
            Caption         =   "View Mail Headers"
         End
         Begin VB.Menu mEnableMailHeaders 
            Caption         =   "Enable Mail Headers"
         End
      End
   End
   Begin VB.Menu mTools 
      Caption         =   "Tools"
      Begin VB.Menu mToolsAddressBook 
         Caption         =   "Address Book"
      End
   End
   Begin VB.Menu mNym 
      Caption         =   "N&ym"
      Begin VB.Menu TransferPrepare 
         Caption         =   "&Send Message Using Nym "
         Enabled         =   0   'False
         Shortcut        =   ^Y
         Visible         =   0   'False
      End
      Begin VB.Menu TransferEncrypt 
         Caption         =   "&Encrypt nym message"
         Visible         =   0   'False
      End
      Begin VB.Menu mFile_DecryptNymMessage 
         Caption         =   "Decrypt a Nym Message"
         Visible         =   0   'False
      End
      Begin VB.Menu NymNull1 
         Caption         =   "-"
      End
      Begin VB.Menu TransferNym 
         Caption         =   "&Create nym..."
      End
      Begin VB.Menu mDeleteNym 
         Caption         =   "&Delete a nym local and from server"
      End
      Begin VB.Menu NymReplyChange 
         Caption         =   "Change nym &reply block..."
      End
      Begin VB.Menu mShowNyms 
         Caption         =   "Show Nyms"
      End
      Begin VB.Menu NymShow 
         Caption         =   "&Show nym server stats..."
      End
   End
   Begin VB.Menu mHelp 
      Caption         =   "&Help"
      Begin VB.Menu HelpAbout 
         Caption         =   "&About..."
      End
      Begin VB.Menu HelpInfo 
         Caption         =   "&Information..."
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mRegistration 
         Caption         =   "Enter registration details"
      End
      Begin VB.Menu SendSysInfo 
         Caption         =   "PI &Diagnostics"
      End
      Begin VB.Menu HelpStep 
         Caption         =   "&Step me through..."
         Begin VB.Menu StepEncrypt 
            Caption         =   "Encrypting a message"
         End
         Begin VB.Menu StepDecrypt 
            Caption         =   "Decrypting a message"
         End
         Begin VB.Menu StepSign 
            Caption         =   "Signing a message"
         End
         Begin VB.Menu StepSend 
            Caption         =   "Sending a message"
         End
         Begin VB.Menu StepAttach 
            Caption         =   "Sending an attachment and using MIME"
         End
         Begin VB.Menu StepSep1 
            Caption         =   "-"
         End
         Begin VB.Menu StepSendKey 
            Caption         =   "Giving my PGP key to someone"
         End
         Begin VB.Menu StepGetKey 
            Caption         =   "Getting/sending MIT server keys"
         End
         Begin VB.Menu StepAddKey 
            Caption         =   "Adding a PGP key to my public key ring"
         End
         Begin VB.Menu StepDelete 
            Caption         =   "Deleting a PGP key from my key ring"
         End
         Begin VB.Menu StepCreateKey 
            Caption         =   "Creating a new PGP key pair"
         End
         Begin VB.Menu StepSep2 
            Caption         =   "-"
         End
         Begin VB.Menu StepRemailer 
            Caption         =   "Sending an anonymous message"
         End
         Begin VB.Menu StepUSENET 
            Caption         =   "Posting an anonymous USENET article"
         End
         Begin VB.Menu StepUpdateInfo 
            Caption         =   "Updating remailer information"
         End
         Begin VB.Menu StepSep3 
            Caption         =   "-"
         End
         Begin VB.Menu StepNym 
            Caption         =   "Creating a nym"
         End
         Begin VB.Menu StepNymSend 
            Caption         =   "Sending a nym message"
         End
         Begin VB.Menu StepNymPass 
            Caption         =   "Changing a nym password"
         End
         Begin VB.Menu StepNymReply 
            Caption         =   "Changing or creating a nym reply block"
         End
         Begin VB.Menu StepNymDelete 
            Caption         =   "Deleting a nym"
         End
         Begin VB.Menu StepSep4 
            Caption         =   "-"
         End
         Begin VB.Menu StepWeb 
            Caption         =   "Anonymously accessing a Web page"
         End
         Begin VB.Menu StepInfo 
            Caption         =   "Getting Internet information"
         End
      End
      Begin VB.Menu HelpSep 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu HelpSend 
         Caption         =   "&Send feedback"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu HelpSys 
         Caption         =   "Add system info"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuAttachmentOptions 
      Caption         =   "AttachmentOptionsPopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu popupSaveAttachmentAs 
         Caption         =   "Save Attachment As"
      End
   End
End
Attribute VB_Name = "frmPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public AddressList As New cEmailAddressList
Public AttachmentFileName As String
Public MessageID As Long
Public bDisplayMessageMode As Boolean
Private Win As New CWindow
'Public FileIconsImageList As New FileIconImageList
Private m_BusyCancel As Boolean
Private m_ControlKey As Boolean
Private m_fToFieldChanged As Boolean
Private m_fCCFieldChanged As Boolean
'===================================
'Declarations for FileIconImageList
'====================================
'Private mFileTypes() As String
'Private mImglstLarge As ImageList
'Private mImglstSmall As ImageList
Private mLastImageNum As Integer
Public gTemporaryFile As String

'===============Instance index========
Private iLocalInstanceReference As Integer


Private Sub btnCC_Click()
ShowStatus 1, ""
gComposeMode = True
GetRecipient
End Sub

Private Sub btnTo_Click()
ShowStatus 1, ""
gComposeMode = True
GetRecipient
End Sub


Private Sub cmbRemailerSelect_Click()
Select Case cmbRemailerSelect.ListIndex
        Case STANDARD_EMAIL
            DontUseRemailer
            SSRibbon1(7).Enabled = True
            SSRibbon1(0).Picture = SSRibbon1(20).Picture
        
        Case REMAILER_CYPHERPUNK
            UseCypherPunk
            'frmRemailerList.Caption = "Cypherpunk Remailer List"
            WriteProfile "Remailer Info", "EncryptionToRemailers", "True"
        
        Case REMAILER_MIX
            gMixPath = ReadProfile("Remailer Info", "MixmasterPath")
            If Not iFileExists(gMixPath & "\mixmaste.exe") Then
                MsgBox "Can't find file 'Mixmaste.exe' in the path you have selected from the file menu.", vbApplicationModal + vbCritical, "Mixmaster Error"
                DontUseRemailer
                cmbRemailerSelect.ListIndex = 0
                Exit Sub
            End If
            ' Check for shortcuts as well
            SetUseOfMixmaster
            
        Case SEND_MESSAGES_USING_NYM
            Unload frmRemailerList
            SSRibbon1(0).Picture = SSRibbon1(21).Picture
            gRemailerType = SEND_MESSAGES_USING_NYM
        
        Case ENCRYPT_BEFORE_SENDING_MESSAGE
            Unload frmRemailerList
            SSRibbon1(0).Picture = SSRibbon1(21).Picture
            gRemailerType = ENCRYPT_BEFORE_SENDING_MESSAGE
            
         Case ENCRYPT_AND_SIGN_BEFORE_SENDING_MESSAGE
            Unload frmRemailerList
            SSRibbon1(0).Picture = SSRibbon1(21).Picture
            gRemailerType = ENCRYPT_AND_SIGN_BEFORE_SENDING_MESSAGE
            
        Case SIGN_BEFORE_SENDING_MESSAGE
            Unload frmRemailerList
            SSRibbon1(0).Picture = SSRibbon1(21).Picture
            gRemailerType = SIGN_BEFORE_SENDING_MESSAGE
            
    End Select
SaveSettings ' this will save the options
End Sub

Private Sub cmbRemailerSelect_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub


Private Sub Command1_Click()
Dim buffer As String * 1024
Dim ret1 As Long
ret1 = spgpKeyGenerate("john <test@asdflkjtest.com>" & Chr(0), "12345678" & Chr(0), buffer & Chr(0), 1, 0, 1024, 0, 0, 0, Me.hWnd)

End Sub

Private Sub DoConnect_Click()
    CheckConnection
End Sub

Private Sub EditClrAll_Click()
    txtTo.Text = ""
    txtsubject.Text = ""
    txtCC.Text = ""
    MessageArea.Text = ""
    lvwAttachments.ListItems.Clear
    Form_Resize
End Sub


Private Sub EditClrMsg_Click()
    MessageArea.Text = ""
End Sub

Private Sub EditCopy_Click()
    EditPerform WM_COPY
End Sub

Private Sub EditCopyMsg_Click()
    If Len(MessageArea) > 0 Then
        Clipboard.Clear
        Clipboard.SetText MessageArea.Text 'SelStart 'Mid(MessageArea, 1, Length)
    End If
End Sub

Private Sub EditCut_Click()
  EditPerform WM_CUT
End Sub

Private Sub EditHeader_Click()
    MessageArea.SelText = gHeader
End Sub

Private Sub EditOptions_Click()
    Form17.Show 1
End Sub

Private Sub EditPaste_Click()
  EditPerform WM_PASTE
End Sub

Private Sub EditPasteMsg_Click()
    MessageArea.Text = Clipboard.GetText()
End Sub

Public Sub EditSelectAll_Click()
    MessageArea.SetFocus
    MessageArea.SelStart = 0
    MessageArea.SelLength = Len(MessageArea)
End Sub

Private Sub EditSig_Click()
    MessageArea.SetFocus
    MessageArea.SelStart = Len(MessageArea.Text)
    MessageArea.SelText = vbCrLf & vbCrLf & gSig
End Sub
Private Sub EmailWrap_Click()
'turn this off for the present
MessageArea.Text = InsertCRLFs()
End Sub


Private Sub FileExit_Click()
    Unload Me
End Sub

Private Sub FileExport_Click()

Dim FileNum As Integer

    On Error GoTo ExportError
    
    '---------------------------------------------
    'prepare the file save as dialog
    '---------------------------------------------
    CommonDialog1.DialogTitle = "Save message as"
    CommonDialog1.Flags = &H2& + &H4&
    CommonDialog1.Filter = "Text Files (*.txt)|*.txt"
    CommonDialog1.FilterIndex = 1
    CommonDialog1.CancelError = True
    CommonDialog1.InitDir = App.Path
    CommonDialog1.FileName = Left(txtsubject.Text, 64)  '& ".txt"
    CommonDialog1.DefaultExt = "txt"
    CommonDialog1.Action = 2
    FileNum = FreeFile
    Open CommonDialog1.FileName For Output As FileNum
    Print #FileNum, MessageArea.Text
    Close #FileNum
    '---------------------------------------------
    'switch back to normal settings
    '---------------------------------------------
    ChDrive Mid$(App.Path, 1, 3)
    ChDir App.Path
ExportError:
    Exit Sub
End Sub

Private Sub FilePage_Click()
    MsgBox "Not implemented yet."
End Sub

Private Sub FilePrintM_Click()
    Dim buffer As String
    Dim ndx As Integer
    Dim foo As Integer
    On Error Resume Next
    If Printers.Count = 0 Then
        MsgBox "No printers are installed.  Can't continue", vbCritical, "Fatal Error."
        Exit Sub
    End If
    CommonDialog1.ShowPrinter
   ' Printer.Print ""
    MessageArea.SelPrint (Printer.hDC)
     Printer.EndDoc
End Sub


Private Sub FileSave_Click()
    SaveSettings
End Sub


Private Sub Form_Activate()
'Dim Win As New CWindow
'Win.OnTop(Me) = False
Dim SigData As TSig_Data

ShowAttachmentContainer
gActivePIInstance = iLocalInstanceReference
If gPGPVersion = NoPGP Then Exit Sub
If mEncryptAttachmentWithKey.Checked Then ShowStatus 2, "Encrypt Attachments with recipient's PublicKey"
If mConventionallyEncryptAttachment.Checked Then ShowStatus 2, "Conventionally Encrypt Attachment"
If Not bDisplayMessageMode Then Exit Sub

Select Case spgpAnalyseMessage(MessageArea.Text)
    Case PGPAnalyze_Encrypted, PGPAnalyze_EncryptedConventional
        If spgpDecryptMessage = vbCancel Then bDisplayMessageMode = False
    Case PGPAnalyze_EncryptedNoKeys ' Key data
        MsgBox "You don't have the keys on your keyring to decrypt this message.", vbApplicationModal + vbCritical, "PGP Decrypt/Verify"
    Case Else
        bDisplayMessageMode = False
End Select

End Sub

Private Sub Form_Click()
ShowStatus 1, ""
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'see VBPJ dec 97
If KeyCode = vbKeyControl Then
    m_ControlKey = True
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyControl Then
    m_ControlKey = False
End If
End Sub

Private Sub Form_Load()
    Dim FileNum As Integer
    Dim where As Integer
    Dim First As Integer
    Dim strBuffer As String
    Dim Tmp As Integer
    Dim Stimer As Single
    Dim SectionName As String
    Dim cmd As String
    Dim Ecount As Integer
    Dim iResult As Integer
    Dim tmpstr As String
    Dim tmpstr2 As String
    Dim lVersion As Long
    Dim i As Integer
    Dim sPGPStatus As String
    
    Dim App As New CApplication
    On Error GoTo LoadError
    iLocalInstanceReference = gFormInstance
    Win.Center Me, Null
    'Win.OnTop(Me) = True
    '---------------------------------------------
    'init stuff
   ' SubClass (Me.hWnd) 'don't allow them to use the system exit
    '---------------------------------------------
    'Set PI version here
    '---------------------------------------------
    Me.Caption = "Private Idaho E-Mail (" & App.Version & ") for Win9x/NT/2000/XP"
   '--------------------------------------
   ' FileIconsImageList = FileIconImageList

    mLastImageNum = 1
    Set lvwAttachments.Icons = imglstLarge
    Set lvwAttachments.SmallIcons = imglstSmall
    lvwAttachments.ListItems.Clear
    
    txtTo.SelStart = 0
    txtTo.Text = ""
    
    'This is the default
    MailHeader(0).ID = ""
    MailHeader(1).ID = "Newsgroups: "
    MailHeader(2).ID = "Subject: "
    MailHeader(3).ID = "Message-ID: "
    MailHeader(4).ID = "Reference: "
    MailHeader(5).ID = "X-Header: "
    StatusBar.Height = TextHeight("Test") * 1.8
    
    
    gObscurity = 0
    gRemailerType = STANDARD_EMAIL
    gNewsgroupType = 0
    gMinState = 1
    gCutStr = ""
    gLatentStr = ""
    gCRLF = vbCrLf
   ' gSMTPLog = 0
    gNymState = gNYM_IDLE
    gPassPhrase = ""
    gPiStr = "Private i Mail"
    gExit = 0
    gPGPTempFile = "pvtidaho"
    'gPGPResponse.Count = 0
        
       
    '--------------------------------------------
    'Check for PGP Version
    '--------------------------------------------
   ' On Error Resume Next
               
    SectionName = "Options"
    sPGPStatus = ReadProfile(SectionName, "PGPStatus")
    gPGPVersion = sPGPStatus
    'Check the licence

    If Not sPGPStatus = NoPGP Or sPGPStatus = "" Then
        '----------------------------------------------
        'NoPGP means the user does not want to use PGP
        'PGP Not found means it hasn't been found
        ' -------------------------------------------------
        If Not PGP_SDKPresent Then
            gPGPVersion = PGPNotFound
        Else
            'lVersion = spgpVersion
            gPGPVersion = PGP5x
        End If
      
    Select Case gPGPVersion
        Case PGPNotFound
                frmPGPUseOptions.Show vbModal 'this will set gPGPversion and write to registry
                
        'Case NoPGP
                
        Case PGP5x
                On Error Resume Next
                lVersion = spgpVersion
                If Not lVersion > 0 Then
                    gPGPVersion = PGPNotFound
                Else
                    ShowStatus 1, "SPGP Version: " & lVersion & " found!"
                End If
                WriteProfile SectionName, "PGPStatus", gPGPVersion
                   
    End Select
    
 End If
 RestoreSettings
 If gPGPVersion = PGP5x Then
        EnablePGPMenuItems
        mUsePGP.Caption = "Disable Utilities and PGP"
            '
    'Set up the remailer options
    '
    'cmbRemailerSelect.AddItem "Standard Email", 0
   ' cmbRemailerSelect.AddItem "Send via remailer Type 1(Cypherpunk)", 1
   ' cmbRemailerSelect.AddItem "Send via remailer Type 2(Mixmaster)", 2
   ' cmbRemailerSelect.AddItem "Encrypt before Sending Message", 3
   ' cmbRemailerSelect.AddItem "Encrypt and Sign before Sending Message", 4
   ' cmbRemailerSelect.AddItem "Just Sign before Sending Messages", 5
   ' cmbRemailerSelect.AddItem "Send Messages using Nym", 6
   '  cmbRemailerSelect.ListIndex = 0
    Else
        DisablePGPMenuItems
        mUsePGP.Caption = "Enable Utilities and PGP"
            'Set up the remailer options
    '
   ' cmbRemailerSelect.AddItem "Standard Email", 0
    
    ' cmbRemailerSelect.ListIndex = 0
       ' cmbRemailerSelect.ListIndex = -1 '0
    End If
    cmbRemailerSelect.Visible = True
    'gPGPKeyID = ReadProfile(SectionName, "KeyID")
   

   
    'restore settings stored in the ini
    
    SetAttachmentEncryptionOptions
    MessageArea.SelBold = False
    'mDontEncryptAttachment.Checked = True
    'ShowAttachmentOptions ("Don't Encrypt Attachments")
    '
    'Set up the remailer options
    '
    'cmbRemailerSelect.AddItem "Standard Email", 0
   ' cmbRemailerSelect.AddItem "Send via remailer Type 1(Cypherpunk)", 1
  '  cmbRemailerSelect.AddItem "Send via remailer Type 2(Mixmaster)", 2
   ' cmbRemailerSelect.AddItem "Encrypt before Sending Message", 3
   ' cmbRemailerSelect.AddItem "Encrypt and Sign before Sending Message", 4
   ' cmbRemailerSelect.AddItem "Just Sign before Sending Messages", 5
    'cmbRemailerSelect.AddItem "Send Messages using Nym", 6
    ' cmbRemailerSelect.ListIndex = 0
    ' cmbRemailerSelect.Visible = True
   
    'cmbRemailerSelect.AddItem "Anonymous Post to NewsGroup", 5
    'cmbRemailerSelect.AddItem "Anonymous Post to NewsGroup using Nym", 6
   ' If frmMain.CheckLicenceExpired Then
     '   cmbRemailerSelect.Locked = True
     '  mPGP.Enabled = False
     '   mNym.Enabled = False
     '   mSetMixmasterPath.Enabled = False
     '   cmbRemailerSelect.Enabled = False
   ' End If
    
    
    If gRemailerType = REMAILER_CYPHERPUNK Then cmbRemailerSelect.ListIndex = 1
    If gRemailerType = REMAILER_MIX Then cmbRemailerSelect.ListIndex = 2
    If gRemailerType = STANDARD_EMAIL Then cmbRemailerSelect.ListIndex = 0
    If gRemailerType = SEND_MESSAGES_USING_NYM Then cmbRemailerSelect.ListIndex = 4
    If gRemailerType = ENCRYPT_BEFORE_SENDING_MESSAGE Then cmbRemailerSelect.ListIndex = 3
    
    
    ' House Keeping
    Set App = Nothing
    Set Win = Nothing
    AddressList.Initialise
    MousePointer = vbDefault
  '  If frmMain.CheckLicenceExpired Then
   '     ShowStatus 1, "Trial period expired."
   ' Else
    '    ShowStatus 1, ""
   ' End If
    Exit Sub

LoadError:
     MousePointer = vbArrow
    MsgBox Err.Description & " Main Form Load " & Str$(where) & "-" + Str$(Err)
    Unload frmSplash
    Err.Clear
End Sub




Private Sub Form_Resize()
Dim MessageAreaHeight As Integer
Dim MessageAreaWidth As Integer
Dim BottomMargin As Integer
'Dim LeftMargin As Integer
'Static StatusTop As Long
'DoEvents
On Error Resume Next
  ' BottomMargin = 800
   'LeftMargin = 200
  ' DoEvents
   If WindowState <> 1 Then
        
        
        DoEvents
        If lvwAttachments.Visible Then
            lvwAttachments.Top = ScaleHeight - lvwAttachments.Height - StatusBar.Height    '- 1500
            lvwAttachments.Width = Width - MessageArea.Left - 180
        End If
                
        If lvwAttachments.Visible Then
            MessageArea.Height = lvwAttachments.Top - MessageArea.Top
            MessageArea.Width = Width - MessageArea.Left - 180
        Else
            MessageArea.Height = ScaleHeight - MessageArea.Top - StatusBar.Height
            MessageArea.Width = Width - MessageArea.Left - 180
        End If
        MessageAreaHeight = MessageArea.Height
        MessageAreaWidth = MessageArea.Width
        
        WebBrowser1.Top = MessageArea.Top
        WebBrowser1.Height = MessageAreaHeight
        WebBrowser1.Width = MessageAreaWidth
        
       ' StatusBar.Panels(0).Left = MessageArea.Left
        StatusBar.Panels(1).Width = MessageAreaWidth / 3 '- StatusBar.Panels(1).Left
        StatusBar.Panels(2).Width = MessageAreaWidth / 3 '- StatusBar.Panels(2).Left
        StatusBar.Panels(3).Width = MessageAreaWidth / 3 '- StatusBar.Panels(3).Left
        
        SSPanel1.Width = Width - 8
        
        txtCC.Width = MessageAreaWidth - txtCC.Left
        txtsubject.Width = MessageAreaWidth - txtsubject.Left
               
        txtTo.Width = MessageAreaWidth - txtTo.Left
        lblFrom(1).Width = txtTo.Width
        'Enabled = True
        'End If
    End If

End Sub

Private Sub Form_Terminate()
Unload Me
Set frmPI = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
   ' Call UnSubClass(Me.hWnd)
    'Unload frmRemailerList
    Dim Form As Form
For Each Form In Forms
    If Form.name <> Me.name And Form.name <> "frmMain" Then
        Unload Form
        Set Form = Nothing
    End If
Next Form
'InstanceNumber=InstanceNumber-1
'Clear this
On Error Resume Next
If Not gTemporaryFile = "" Then Kill gTemporaryFile

'Set AddressList = Nothing
gComposeMode = False
'Wipe any temporary web files
WipeFile (App.Path & "\temp.html")
End Sub


Private Sub GetFinger_Click()
    If CheckConnection Then
        frmFingerCommand.Show  'SendFinger
    End If
End Sub

Private Sub HelpAbout_Click()
  frmAbout.Show vbModal
End Sub
Private Sub HelpInfo_Click()
    ShowWebHelp
End Sub


Private Sub HelpSys_Click()
    HelpSys.Checked = Not HelpSys.Checked
End Sub
Private Sub HTTP1_EndTransfer(Direction As Integer)
Dim bRes As Boolean
Dim iReturnedRemailers As Integer
   
   'Initialise this to non zero
   iReturnedRemailers = 1
   
   If Direction = 0 Then Exit Sub '0 = client
    
    On Error GoTo HTTPEndTransferErr
    Me.MousePointer = vbHourglass
    Select Case gWebState
    Case GETREMAILERUPDATE
        DoEvents
        'Convert to txt format by adding crlfs etc
        gWebPage = ConvertStr(gWebPage)
        bRes = PutFileText(App.Path & "\remailer.htm", gWebPage)
        gTotalRemailers = 0
        frmRemailerList.lblStatus = "Filling type 1 remailer list..."
        frmRemailerList.InitializeRemailers (App.Path + "\remailer.htm")
        iReturnedRemailers = gTotalRemailers
        frmRemailerList.lblStatus = "Filling private remailer list..."
        frmRemailerList.InitializeRemailers (App.Path + "\private.txt")
        DoEvents
         'This will fill the matched remailer list
        frmRemailerList.SortRemailers
        frmRemailerList.FillRemailerList
    Case MIXUPDATE
        DoEvents
        gWebPage = ConvertStr(gWebPage)
        WriteStrToFile App.Path & "\mixmaster.htm", gWebPage
        gTotalRemailers = 0
        frmRemailerList.InitializeRemailers (App.Path & "\mixmaster.htm")
        frmRemailerList.lblStatus = "Filling Remailer List..."
        iReturnedRemailers = gTotalRemailers
        DoEvents
        frmRemailerList.SortRemailers
        frmRemailerList.FillRemailerList
    Case TYPE2UPDATE
        DoEvents
        gWebPage = ConvertStr(gWebPage)
        frmRemailerList.lblStatus = "Writing data to Mixmaster Type2.lis file."
        'bRes = PutFileText(App.Path & "\type2.lis", gWebPage)
        bRes = PutFileText(gMixPath & "\type2.lis", gWebPage)
        'WriteStrToFile App.Path & "\type2.lis", gWebPage
    Case PUBRINGUPDATE
        DoEvents
        gWebPage = ConvertStr(gWebPage)
        frmRemailerList.lblStatus = "Updating " & gMixPath & "\pubring.mix file."
       ' bRes = PutFileText(App.Path & "\pubring.mix", gWebPage)
        bRes = PutFileText(gMixPath & "\pubring.mix", gWebPage)
    Case GETSERVERKEY
        gWebPage = ConvertStr(gWebPage)
        MessageArea.SelText = GetKeyBlock(gWebPage)
        MessageArea.SetFocus
    Case GETREMAILERKEYS
        DoEvents
        gWebPage = ConvertStr(gWebPage)
        PGPAddKeyFromString (gWebPage)
        'MessageArea.SelText = gWebPage
        'MessageArea.SetFocus
    End Select
    On Error Resume Next
    Me.MousePointer = vbDefault
    HTTP1.Action = 0
    'HTTP1.WinsockLoaded = False
    gWebState = HTTPIDLE
    ShowStatus 1, ""
    If iReturnedRemailers = 0 Then
        frmRemailerList.lblStatus.ForeColor = vbRed
        frmRemailerList.lblStatus = "Old data displayed because no remailers info found at this site. Check URL."
    Else
        frmRemailerList.lblStatus = "Updated as of: " & Now
    End If
    ProgressBar1.Visible = False
    DoEvents
    Exit Sub

HTTPEndTransferErr:
    Me.MousePointer = vbDefault
    MsgBox "Error in HTTP Transfer: " & Err.Description
    HTTP1.WinsockLoaded = False
    HTTP1.Action = a_Idle
    ProgressBar1.Visible = False
    ShowStatus 1, ""
    Err.Clear
End Sub

Private Sub HTTP1_Error(ErrorCode As Integer, Description As String)
    'gWebState = IdleState
    HTTP1.Action = a_Idle
    HTTP1.WinsockLoaded = False
    MsgBox "HTTP error " & Format$(ErrorCode) & ".  " _
            & Description, 16, "HTTP Error"
End Sub

Private Sub HTTP1_Header(Field As String, Value As String)
Dim s As String
InitialiseProgressBar
Select Case Field
    Case "Content-Length"
        ProgressBar1.Max = Value
End Select

End Sub

Private Sub HTTP1_StartTransfer(Direction As Integer)
    
    If Direction = 0 Then Exit Sub '0 is client
    '---------------------------------------------
    'gWebPage will hold the text contents of the downloaded web page
    '---------------------------------------------
    gWebPage = ""
    frmRemailerList.lblStatus = "Download started..."
    'Screen.MousePointer = vbHourglass
End Sub

Private Sub HTTP1_Transfer(Direction As Integer, BytesTransferred As Long, Text As String)

    On Error GoTo HTTPTransferErr
    If Direction = 0 Then Exit Sub '0 = client
    If BytesTransferred > ProgressBar1.Max Then
        ProgressBar1.Max = BytesTransferred * 2
        ProgressBar1.Value = BytesTransferred
    End If
    gWebPage = gWebPage & Text
    Exit Sub

HTTPTransferErr:
    MsgBox Err.Description & " : - > HTTPTransfer", vbApplicationModal, App.Title
    HTTP1.Action = a_Idle
    HTTP1.WinsockLoaded = False
    Err.Clear
End Sub


Private Sub IPPort1_DataIn(Text As String, EOL As Boolean)
Dim Done As Boolean
Dim Prechopped As Boolean
Dim Pos As Integer
Dim foo As String
Dim MyText As String
Dim TLen As Integer
    TLen = Len(Text)
    Done = False
    Prechopped = True
    MyText = ""
    Pos = 1
    Do While Not Done
        foo = Mid(Text, Pos, 1) '
        If foo = vbLf Then
            MessageArea.SelText = MyText & vbCrLf
            MyText = ""
            Prechopped = False
        ElseIf foo = vbCrLf Then
            MessageArea.SelText = MyText & vbCrLf
            MyText = ""
            Prechopped = False
       Else
           MyText = MyText + foo
        End If
        Pos = Pos + 1
        If Pos > TLen Then
            Done = True
            If Prechopped Then
                MessageArea.Text = MessageArea.Text + Text + vbCrLf
            End If
        End If
    Loop
End Sub
Private Sub IPPort1_Disconnected(StatusCode As Integer, Description As String)

    ShowStatus 1, "Status " & Description
End Sub

Private Sub IPPort1_Error(ErrorCode As Integer, Description As String)

    MessageArea.Text = Err.Number & ": " & Err.Description
End Sub

Private Sub IPPort1_ReadyToSend()
Dim foo As Integer
Dim Storestring1 As String
Dim Storestring2 As String

Dim cmd As String
Dim iResult As Integer

    On Error GoTo KeyCertifyError
    '[??]
    Exit Sub
    gKeyID = ""
    gCancelAction = False
    
    '---------------------------------------------
    'display the list of personal keys to sign with
    '---------------------------------------------
    'z$ = Form26.Label1.Caption
    Form26.Label1.Caption = "Select a key to use for signing the public key."
    CheckMultipleKey
    
    '---------------------------------------------
    'display the public key ring
    '---------------------------------------------
    Storestring1 = frmSelectUserID.Label2.Caption
    Storestring2 = frmSelectUserID.Label1.Caption
    frmSelectUserID.Label2.Caption = "Select a key to sign"
    frmSelectUserID.Label1.Caption = "from the public key ring."
    frmSelectUserID.Show 1
    frmSelectUserID.Label2.Caption = Storestring1
    frmSelectUserID.Label1.Caption = Storestring2
    Unload frmSelectUserID
    
    If Not gCancelAction Then
        '---------------------------------------------
        'User selected okay
        '---------------------------------------------
        cmd = ""
       ' cmd = gPGPPath & "\PGP -ks " + Chr$(34) + gKeyID + Chr$(34) + " -u " + Chr$(34) + gPGPKeyID + Chr$(34)
        CheckLen (cmd)
        ExecCmd (cmd)
        'UpdatePublicKeysFile
    Else
        '---------------------------------------------
        'User hit cancel, or failed to select a key
        '---------------------------------------------
        gCancelAction = False
    End If
Exit Sub
    
KeyCertifyError:
    If Err.Number = 32755 Then
        MsgBox "Could not decrypt the message.  Suggest you open a DOS session and type pgp " + gPGPFile + ".out at the the command prompt in your PGP directory to find out what happened."
    Else
        Err.Number = 53  'no pgpfooFile, it means PGP command did not complete, just abort
    End If
    Err.Clear
    
End Sub



Private Sub KeyCertify_Click()
Dim msg As String
Call KeyCreate_Click

End Sub
Private Sub KeyCreate_Click()

'Dim FileFunctions As New cFileFunctions
'Dim objFiles As New Collection
'Dim objFile As File
'Dim iResponse As Long
Dim SectionName As String
Dim pid As Long
Dim sPath As String

SectionName = "PGP Options"
On Error GoTo BadKeys

sPath = ReadProfile(SectionName, "PGPKeys.exe Location")
If sPath = "" Then
    frmFindPGPKeys.Show vbModal
    'Try again
    sPath = ReadProfile(SectionName, "PGPKeys.exe Location")
    If sPath = "" Then
        Exit Sub
    Else
        pid = Shell(sPath, vbMinimizedFocus)
    End If
Else
     
   Dim sDir As String
   On Error Resume Next
   'Shell out but don't wait for termination
   pid = Shell(sPath, vbMinimizedFocus)
   DoEvents
     
End If
Exit Sub
BadKeys:
    Me.MousePointer = vbDefault
    MsgBox "Error detected while trying to find or execute the PGPKeys.exe file.  Error reported as: " & Err.Description, vbApplicationModal + vbCritical, "File Location Error"
    Err.Clear
End Sub

Private Sub KeyEditTrust_Click()
Call KeyCreate_Click

End Sub


Private Sub KeyOptions_Click()
    gMultiType = 1
    Form25.Show vbModal
End Sub

Private Sub KeySubmit_Click()
Dim SectionName As String
    If CheckConnection Then
        'is there a key there?
        If InStr(1, MessageArea.Text, "PUBLIC KEY BLOCK") = 0 Then
            MsgBox "Please enter the PGP public key you'd like to submit to the key server in the Message box."
            Beep
            Exit Sub
        End If
        If gSubKeyURL = "" Then
            SectionName = "Net Info"
            gSubKeyURL = ReadProfile(SectionName, "SubmitKeyURL")
            If gSubKeyURL = "" Then frmSelectKeyServer.Show vbModal
            If gSubKeyURL = "" Then
                gSubKeyURL = "pgp-public-keys@pgp.ai.mit.edu"
            End If
        End If
        txtTo.Text = gSubKeyURL
        'Tell the server to add the key.
        txtsubject.Text = "add"
        ShowStatus 1, ""
        SendToOutBox
        'SendMailMessage
    End If
End Sub





Private Sub lblFrom_Change(Index As Integer)
lblFrom(0).Visible = True
lblFrom(1).Visible = True
End Sub

Private Sub lvwAttachments_DblClick()
Dim res As Long
Dim obj As Object
Dim Attachment As String
Dim FileToLaunch As String
Dim msg As String
Dim iRes As Integer

On Error Resume Next
If Not gTemporaryFile = "" Then Kill gTemporaryFile

On Error GoTo BadAttachmentLaunch

Attachment = lvwAttachments.SelectedItem.Text 'App.Path & "\mailbox\attachments\" & lvwAttachments.SelectedItem.Text
'
'If it is asc then decode it within this app.
'

If GetExt(Attachment) = "asc" Then
    frmAttachmentOptions.lblFileName = Attachment
    frmAttachmentOptions.Show vbModal
    iRes = frmAttachmentOptions.iRes
    Set frmAttachmentOptions = Nothing
    Select Case iRes
        Case vbYes
            
            vb2spgpContext.Initialise
            vb2spgpContext.FileIn = TempPathLocation & Attachment '
            vb2spgpContext.FileOut = TempPathLocation & IIf(InStr(1, StripExt(Attachment), ".") = 0, StripExt(Attachment) & ".htm", StripExt(Attachment)) 'TempPathLocation & StripExt(lvwAttachments.SelectedItem.Text)
            spgpDecryptFile
            FileToLaunch = vb2spgpContext.FileOut
        Case vbNo
            CommonDialog1.DialogTitle = "Decrypt " & Attachment & " and Save file as:"
            CommonDialog1.Flags = &H2& + &H4&
            CommonDialog1.Filter = "File Type (*." & GetExt(StripExt(Attachment)) & ")"
            CommonDialog1.FilterIndex = 1
            CommonDialog1.CancelError = True
            CommonDialog1.InitDir = App.Path
            CommonDialog1.FileName = IIf(InStr(1, StripExt(Attachment), ".") = 0, StripExt(Attachment) & ".htm", StripExt(Attachment)) 'TempPathLocation & StripExt(lvwAttachments.SelectedItem.Text)
            CommonDialog1.DefaultExt = GetExt(StripExt(Attachment))
            CommonDialog1.Action = 2
            ChDrive Mid(App.Path, 1, 3)
            ChDir App.Path
            vb2spgpContext.Initialise
            vb2spgpContext.FileIn = Attachment
            vb2spgpContext.FileOut = CommonDialog1.FileName '& CommonDialog1.DefaultExt
            spgpDecryptFile
            ShowStatus 1, "Decrypted file was saved successfully"
            Exit Sub
        Case vbCancel
            Exit Sub
        End Select
 
 Else
    FileToLaunch = Attachment
 
 End If
'
'Now try and launch it if we can
'
    DoEvents
    gTemporaryFile = FileToLaunch
    res = ShellExecute(Me.hWnd, "open", FileToLaunch, vbNullString, TempPathLocation, SW_SHOW) 'App.Path & "\temp"
    DoEvents
    If res < 32 Then
             Kill FileToLaunch 'TempPathLocation & lvwAttachments.SelectedItem.Text
             Err.Raise "Error was encountered launching the application associated with this attachment.   Please check your " & TempPathLocation & " directory to makes sure there are no plain text (decrypted) files there."
    End If
    'Need these otherwise the temp file be deleted before the application is lauched
    'ShowStatus ("Temporary file: " & vb2spgpContext.FileOut & " exists!")
    'Kill vb2spgpContext.FileOut
Exit Sub
BadAttachmentLaunch:
    MsgBox "An error has been encountered: " & Err.Description, vbApplicationModal + vbCritical, "Attachment Launch"
    Err.Clear
End Sub



Private Sub lvwAttachments_ItemClick(ByVal Item As ComctlLib.ListItem)
Dim i As String
'Text1.Text = Item
i = lvwAttachments.SelectedItem.Text
End Sub

Private Sub lvwAttachments_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyDelete Then
    lvwAttachments.ListItems.Remove (lvwAttachments.SelectedItem.Index)
    If lvwAttachments.ListItems.Count = 0 Then
        lvwAttachments.Visible = False
        Form_Resize
    End If
End If
End Sub



Private Sub lvwAttachments_LostFocus()
On Error Resume Next
If Not gTemporaryFile = "" Then Kill gTemporaryFile
End Sub

Private Sub lvwAttachments_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = 2 Then   ' Check if left mouse button
                       ' was clicked.
      PopupMenu mnuAttachmentOptions  ' Display the File menu as a
                        ' pop-up menu.
End If
 

End Sub

Private Sub lvwAttachments_OLEDragOver(Data As ComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Effect = vbDropEffectNone
End Sub

Private Sub mAddMailHeaders_Click()
frmMailHeader.Show
End Sub

Private Sub mAdvancedEncryption_Click()
frmAdvancedEncryptionOptions.Show vbModal
End Sub

Private Sub mAnalyseMessage_Click()
Select Case spgpAnalyseMessage(MessageArea.Text)
        Case PGPAnalyze_Encrypted, PGPAnalyze_EncryptedConventional
           'spgpDecryptMessage
           ' If Not SignatureProperties.Status = "SIGNED_NOT" Then
            ShowStatus 1, "Message is encrypted."   '"Signed by: " & SignatureProperties.UserID
               ' ShowStatus 2, "Signature Status: " & SignatureProperties.Status
           ' End If
        Case PGPAnalyze_Unknown
                ShowStatus 1, "The message area does not contain an encrypted message"
        Case PGPAnalyze_Signed, PGPAnalyze_DetachedSignature
                ShowStatus 1, "The message has been signed...."
                DoEvents
                'spgpVerifyMessage
        Case PGPAnalyze_Key
                ShowStatus 1, "Message contains a key"
        Case PGPAnalyze_EncryptedNoKeys ' Key data
                ShowStatus 1, "You don't have the keys on your keyring to decrypt this message."
            End Select
End Sub

Private Sub mAttachFile_Click()
AddAttachment
End Sub

Private Sub mConventionallyEncryptAttachment_Click()
mEncryptAttachmentWithKey.Checked = False
'mDontEncryptAttachment.Checked = False
mConventionallyEncryptAttachment.Checked = True
ShowStatus 2, "Conventionally Encrypt Attachment"
End Sub

Private Sub mDecodeFile_Click()
Dim EncodedData As String
Dim FileNum As Integer
Dim FileName As String
Dim TextLine As String
Dim NumBytes As Long
Dim msg As String
Dim foo As Long
Dim i As Integer
Dim FileSize As Long
Dim LineCount As Long
On Error GoTo ImportError
If Len(MessageArea.Text) = 0 Then
    MsgBox "The is nothing in the message area.", vbExclamation + vbApplicationModal, "Empty Message Area"
    DoEvents
    Exit Sub
End If
'ShowStatus ("Looking for beginning of file...")
'foo = InStr(1, MessageArea, "name=")
'i = 0
'If Not foo = 0 Then
  '  foo = foo + Len("name=""")
   ' Do While i < 128
       ' TextLine = Mid(MessageArea, foo + i, 1)
       ' If TextLine = vbCr Or TextLine = vbLf Or TextLine = """" Then Exit Do
        'FileName = FileName & TextLine
        'i = i + 1
    'Loop
    'ShowStatus ("Found file: " & FileName)
   ' DoEvents
'End If
'If Not InStr(1, MessageArea, "base64") = 0 Then
'    NetCode1.Format = f_BASE64
'Else
 '   NetCode1.Format = f_UUEncode
'End If
    
NetCode1.MaxFileSize = 0
NetCode1.Overwrite = True
ShowStatus 1, "Decoding file: " & FileName
DoEvents
NetCode1.FileName = App.Path & "\" & FileName
On Error GoTo ImportError
NetCode1.EncodedData = MessageArea.Text
NetCode1.Action = 3 'Decode to file
NetCode1.Action = 0

DoEvents
MousePointer = vbDefault

CommonDialog1.DialogTitle = "Save file"

If Not FileName = "" Then
    CommonDialog1.Filter = GetExt(FileName)
    CommonDialog1.FileName = App.Path & "\" & FileName
Else
    CommonDialog1.Filter = GetExt(NetCode1.FileName)
    CommonDialog1.FileName = NetCode1.FileName
End If

CommonDialog1.InitDir = App.Path
CommonDialog1.ShowSave
FileNum = FreeFile
Open CommonDialog1.FileName & "." & CommonDialog1.Filter For Output As FileNum
Print #FileNum, NetCode1.DecodedData
Close FileNum
ShowStatus 1, "Saved file at: " & CommonDialog1.FileName & "." & CommonDialog1.Filter
ChDir App.Path

Exit Sub
ImportError:
    Reset
    Beep
    MsgBox Err.Description, vbApplicationModal, App.Title
    MousePointer = vbDefault
    ChDir App.Path
    Err.Clear
End Sub

Private Sub mDeleteNym_Click()
'We need to do this to have control over the To box
    gComposeMode = False
    gCancelAction = False
    AddressList.Initialise
    RemoveUnderlineFromToBox
    RemoveUnderlineFromccBox

    If Not gRemailerType = REMAILER_CYPHERPUNK Then
        cmbRemailerSelect.ListIndex = 1
    End If
    'gNymState = gNYMDEL
    ShowStatus 1, ""
    frmMultiNyms.Show vbModal
    If gCancelAction Then Exit Sub
    ProcessNymCommand gNYMDEL, Nym.ListIndex
    gNymState = gNYM_IDLE
    ShowStatus 1, "You can now send the change request."
End Sub

Private Sub mDontEncryptAttachment_Click()
mEncryptAttachmentWithKey.Checked = False
mDontEncryptAttachment.Checked = True
ShowStatus 2, "Don't Encrypt Attachment"
mConventionallyEncryptAttachment.Checked = False
End Sub

Private Sub mEnableMailHeaders_Click()
mEnableMailHeaders.Checked = Not mEnableMailHeaders.Checked
If mEnableMailHeaders.Checked = True Then
    MailHeader(0).ID = "ExtendedHeaders"
Else
    MailHeader(0).ID = ""
End If
End Sub

Private Sub mEncodeFile_Click()
Set frmEncodeFile.MessageArea = Me.MessageArea
frmEncodeFile.Show vbModal
End Sub

Private Sub mEncryptAttachmentWithKey_Click()
ShowStatus 2, "Encrypt Attachment with Key"
mEncryptAttachmentWithKey.Checked = True
'mDontEncryptAttachment.Checked = False
mConventionallyEncryptAttachment.Checked = False
End Sub

Private Sub MessageArea_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim J As Integer
Dim k As String
Dim lListItem As ListItem
Dim sFileName As String

'Set lvwAttachments.Icons = FileIconsImageList1.Icons
'Set lvwAttachments.SmallIcons = FileIconsImageList1.SmallIcons

If Data.GetFormat(vbCFFiles) Then
    
    Dim vFN
    For Each vFN In Data.Files
        sFileName = StripFileName(CStr(vFN))
        J = GetFileIconNum(sFileName)
        Set lListItem = lvwAttachments.ListItems.Add(, , sFileName, J, J)
        
        'Create copy in temporary file location
        FileCopy CStr(vFN), TempPathLocation & sFileName
        
        'lListItem.SubItems(1) = FileSize("", mExtractCab.CompressedFiles.Item(i).FileSize)
        'lListItem.SubItems(2) = GetFileType(mExtractCab.CompressedFiles.Item(i).FileCabName)
        'lListItem.SubItems(3) = FileTime("", mExtractCab.CompressedFiles.Item(i).FileDate, mExtractCab.CompressedFiles.Item(i).FileTime)
        ' lvwAttachments MessageArea = vFN
        ' MessageArea = Data.GetData(vbCFFiles)
    Next vFN
    ShowAttachmentContainer
End If
'Effect = 9
End Sub


Private Sub mFile_DecryptNymMessage_Click()
gNymState = gNYM_DECRYPT
frmMultiNyms.Show vbModal
gNymState = gNYM_IDLE
End Sub

Private Sub mFile_Import_Click()
    ImportFile
End Sub

Private Sub mFile_Save_Click()
    SaveMessage
End Sub



Private Sub IPPort1_Connected(StatusCode As Integer, Description As String)

    If 0 = StatusCode Then    'OK
        'ShowStatus("")
    Else
        MsgBox "Connection failed: " & Description, vbAbortRetryIgnore, App.Title
    End If
End Sub

Private Sub mGetNews_Click()
frmNewsReader.Show
End Sub

Public Sub mKeyRingIDs_Click()
Dim BufferOut As String
Dim i As Long
Dim Count As Long
Dim BufferLen As Long
Dim spgperr As String * 256
    
    'Note vb2spgpContext.SelectPrivateKeys must be set to show private keys
    BufferLen = spgpKeyRingCount() * 1024
    BufferOut = String(BufferLen, Chr(0))
    i = spgpKeyRingID(BufferOut, BufferLen)
    If i = 0 Then
        'vb2spgpContext.SelectPrivateKeys = False
        frmViewKeyRing.lblContext = "This is a list of all keys on your keyring."
        frmViewKeyRing.Show vbModal
    Else
        Call spgpGetErrorString(i, spgperr)
        MsgBox "An error occured:  The error that was returned is: " & spgperr, vbApplicationModal + vbCritical, "View Keyring"
    End If
    
    
End Sub

Private Sub mLargeAttachmentIcons_Click()
lvwAttachments.View = lvwIcon
End Sub

Private Sub mMergeFile_Click()
frmMergeFile.Show
End Sub

Private Sub mnuSaveAttachmentAs_Click()
If Not lvwAttachments.Visible Then
    MsgBox "There are no attachments to save.", vbApplicationModal + vbCritical, "No attachments"
Else
    popupSaveAttachmentAs_Click
End If
End Sub

Private Sub mPGPEstimatePassPhrase_Click()
Dim PPQuality As Long
Dim sResponse As String
sResponse = myInputBox("Input the passphrase you wish to test in the text box below. ", "Passphrase quality check.")
PPQuality = spgpEstimatePassphraseQuality(sResponse)
MsgBox "A value less than 100 should be considered as being capable of improvement. " & vbCrLf & vbCrLf & "Your phassphrase has been determined as having a quality value of " & PPQuality, vbInformation + vbApplicationModal
End Sub

Private Sub mRegistration_Click()
frmLicence.Show vbModal
End Sub
Private Sub mSetMixmasterPath_Click()
SetMixMasterPath
End Sub

Private Sub mSetPGPKeysPath_Click()
frmFindPGPKeys.Show vbModal
End Sub

Private Sub mShowNyms_Click()
frmNymsList.Show
End Sub

Private Sub mSmallAttachmentIcons_Click()
lvwAttachments.View = lvwList
lvwAttachments.Refresh
End Sub

Private Sub mSplitFile_Click()
frmFileSplitter.Show
End Sub

Private Sub mToolsAddressBook_Click()
frmEditAddressBook.Show vbModal
End Sub



Private Sub mUsePGP_Click()
Dim SectionName As String

SectionName = "Options"
gPGPVersion = ReadProfile(SectionName, "PGPStatus")
If gPGPVersion = PGPNotFound Then
    MsgBox "One of the PGP crucial libraries cannot be found.  Try re-starting Private Idaho.  If this does not work you may need to re-install PGP or Private Idaho.", vbApplicationModal + vbCritical, "PGP not found"
    Exit Sub
End If
'If gPGPVersion = NoPGP Then frmPGPUseOptions.Show vbModal
If gPGPVersion = NoPGP Then
    gPGPVersion = PGP5x
    mUsePGP.Caption = "Disable Utilities and PGP"
Else
    gPGPVersion = NoPGP
    mUsePGP.Caption = "Enable Utilities and PGP"
End If

If gPGPVersion = PGP5x Then
    EnablePGPMenuItems
Else
    DisablePGPMenuItems
End If
WriteProfile SectionName, "PGPStatus", gPGPVersion
lvwAttachments.Visible = True
'EnablePGPMenuItems
End Sub





Private Sub mViewCypherPunksRemailerList_Click()
gRemailerType = REMAILER_CYPHERPUNK
frmRemailerList.Show
frmRemailerList.Caption = "Cypherpunk Remailer List"
End Sub

Private Sub mViewMixmasterRemailerList_Click()
gRemailerType = REMAILER_MIX
frmRemailerList.Show
frmRemailerList.Caption = "Mixmaster Remailer List"
End Sub

Private Sub NymReplyChange_Click()
    Dim Vbresponse As Integer
    Dim msg As String
    
    'We need to do this to have control over the To box
   
    'gComposeMode = False
    AddressList.Initialise
    txtTo.Text = ""
    txtCC.Text = ""
    'RemoveUnderlineFromToBox
    'RemoveUnderlineFromccBox
    
   
    'gNymState = gNYMRPLYCHANGE
    Nym.create = False
    frmMultiNyms.Show vbModal
    
    'Note must set this
    gNymState = gNYMRPLYCHANGE
    If Not gCancelAction Then ProcessNymCommand gNymState, Nym.ListIndex
    If gCancelAction Then
        gCancelAction = False
        txtTo.Text = ""
    Else
        AddressList.Initialise
        RemoveUnderlineFromToBox
        ScanTextForContacts txtTo, CONTACT_TO_LIST
        SendNymMessage
        Unload Me
    End If
AddressList.Initialise
End Sub

Private Sub NymShow_Click()
    gShowNymStatus = 1
    frmNymServerStats.Command1.Enabled = False
    frmNymServerStats.Show
    gShowNymStatus = 0
End Sub


Private Sub PGPAddKey_Click()
PGPAddKeyFromString (MessageArea.Text)
End Sub


Private Sub PGPConvent_Click()
    PGPConvent.Checked = Not PGPConvent.Checked
    If PGPConvent.Checked Then
        PGPEnSign.Enabled = False
    Else
        PGPEnSign.Enabled = True
    End If
End Sub
Private Sub PGPDecrypt_Click()

    Dim ClipText As String
    Dim FileNum As Integer
    Dim iResult As Integer
    Dim Ecount As Integer
    Dim TextLine As String
    Dim Cyphertext As String
    Dim TheFileName As String
    'Dim SigData As TSig_Data
   ' Dim foo As Integer
        
    On Error GoTo DecryptError
    
    If Not PGPFile.Checked Then
            Select Case spgpAnalyseMessage(MessageArea.Text)
                Case PGPAnalyze_Encrypted, PGPAnalyze_EncryptedConventional
                    spgpDecryptMessage
                    If Not SignatureProperties.Status = "SIGNED_NOT" Then
                        ShowStatus 1, "Signed by: " & SignatureProperties.UserID
                        ShowStatus 2, "Signature Status: " & SignatureProperties.Status
                    End If
                Case PGPAnalyze_Unknown
                    MsgBox "The message area does not contain an encrypted message", vbApplicationModal + vbCritical, "Analyse Message"
                Case PGPAnalyze_Signed, PGPAnalyze_DetachedSignature
                    ShowStatus 1, "The message has been signed...."
                    DoEvents
                    spgpVerifyMessage
                Case PGPAnalyze_Key
                    MsgBox "Message contains a key", vbApplicationModal + vbCritical, "Decrypt Verify"
                Case PGPAnalyze_EncryptedNoKeys ' Key data
                    MsgBox "You don't have the keys on your keyring to decrypt this message.", vbApplicationModal + vbCritical, "PGP Decrypt/Verify"
            End Select
        
    Else
            CommonDialog1.DialogTitle = "Select File to decrypt"
            CommonDialog1.Flags = &H2& + &H4&
            CommonDialog1.Filter = "All Files (*.asc)|*.asc"
            CommonDialog1.FilterIndex = 1
            CommonDialog1.CancelError = True
            CommonDialog1.InitDir = App.Path
            CommonDialog1.Action = 1
            ChDrive Mid$(App.Path, 1, 3)
            ChDir App.Path
            TheFileName = CommonDialog1.FileName
            
            CommonDialog1.DialogTitle = "Save decrypted file as:"
            CommonDialog1.Flags = &H2& + &H4&
            CommonDialog1.Filter = "All Files (*.*)|*.*"
            CommonDialog1.FilterIndex = 1
            CommonDialog1.CancelError = True
            CommonDialog1.InitDir = App.Path
            CommonDialog1.Action = 2
            ChDrive Mid$(App.Path, 1, 3)
            ChDir App.Path
                vb2spgpContext.Initialise
                vb2spgpContext.FileIn = TheFileName
                If Not InStr(CommonDialog1.FileName, ".asc") = 0 Then
                    vb2spgpContext.FileOut = StripExt(CommonDialog1.FileName)
                Else
                    vb2spgpContext.FileOut = CommonDialog1.FileName
                End If
                spgpDecryptFile
                ShowStatus 1, "File successfully decrypted and save as: " & CommonDialog1.FileName
        End If
 '   if s
    Exit Sub

DecryptError:
    MsgBox "Could not decrypt or verify the message.  Following error was returned by PGP: " & Err.Description
    Err.Clear
End Sub

Private Sub PGPDeleteKey_Click()

    Dim iResponse As Long
        
        '---------------------------------------------
    'present the list of public keys
    '---------------------------------------------
    vb2spgpContext.Initialise
    vb2spgpContext.SelectPrivateKeys = False
    frmViewKeyRing.lblContext = "This is a list of all keys on your keyring."
    frmViewKeyRing.Show vbModal
    If Not gCancelAction Then
    'Delete the selected key
        iResponse = spgpKeyRemove(Key.UserID)
        If iResponse = 0 Then
            ShowStatus 1, "The key " & Key.UserID & " was successfully removed."
        Else
            ShowStatus 1, "Removal of the the key " & Key.UserID & " was unsuccessfull."
        End If
        gCancelAction = False
    Else
        ShowStatus 1, "Action cancelled or too many keys chosen!"
    End If

End Sub
Private Sub PGPEncrypt_Click()
Dim sToList As String
Dim sCCList As String
Dim iNumPGPEntries As Integer
Dim iNumNonPGPEntries As Integer

'This will scan the Tx and To boxes just in case the timer did not go off
Call Timer1_Timer

MousePointer = vbHourglass

'Save original address lists
sToList = txtTo.Text
sCCList = txtCC.Text

ConvertMailGroupsToAddressList (CONTACT_TO_LIST)
ConvertMailGroupsToAddressList (CONTACT_CC_LIST)

iNumPGPEntries = AddressList.GetNumberPGPContacts
iNumNonPGPEntries = AddressList.GetNumberNonPGPContacts

If Not iNumNonPGPEntries = 0 Then
    MsgBox "There are addresses in either your TO and CC list that are not on your PGP keyring.  Can't process encryption request.", vbCritical + vbApplicationModal
Else
    txtCC.Text = AddressList.GetListAllContacts(CONTACT_CC_LIST)
    txtTo.Text = AddressList.GetListAllContacts(CONTACT_TO_LIST)
    PGPEncryptMessage False, False 'Sign and clearsign

End If

'Replace original
txtTo.Text = sToList
txtCC.Text = sCCList
'This clever little trick will push the original address fields back into the address list
RemoveUnderlineFromToBox
RemoveUnderlineFromccBox

MousePointer = vbDefault
End Sub

Private Sub PGPEncryptToSelf_Click()
   
    On Error GoTo FileEnSError
    gCancelAction = False
    
    'First set the sign parameters as this is common for both routines
     vb2spgpContext.Initialise
    vb2spgpContext.ConventionalEncrypt = 0
    vb2spgpContext.KeyEncrypt = 1
      
    gPGPKeyID = ReadProfile("PGP Options", "Default Key ID")
    If gPGPKeyID = "" Then
        MsgBox "The default key is not set.  Set in PGP:Options", vbApplicationModal + vbCritical, "Encrypt to Self"
        Exit Sub
    End If
    vb2spgpContext.CryptKeyID = gPGPKeyID
    'vb2spgpContext.SignKeyID = ""
    
    vb2spgpContext.TextMode = 0

    vb2spgpContext.Armor = 1
    spgpEncryptMessage
        
gCancelAction = False
Exit Sub
FileEnSError:
        gCancelAction = False
        MsgBox "There was an error.  The reason given by the operating system is: " & Err.Description, vbApplicationModal, App.Title
        Err.Clear
End Sub

Private Sub PGPEnSign_Click()
Dim sToList As String
Dim sCCList As String
Dim iNumPGPEntries As Integer
Dim iNumNonPGPEntries As Integer

'This will scan the Tx and To boxes just in case the timer did not go off
Call Timer1_Timer

MousePointer = vbHourglass
'Save original address lists
sToList = txtTo.Text
sCCList = txtCC.Text

ConvertMailGroupsToAddressList (CONTACT_TO_LIST)
ConvertMailGroupsToAddressList (CONTACT_CC_LIST)

iNumPGPEntries = AddressList.GetNumberPGPContacts
iNumNonPGPEntries = AddressList.GetNumberNonPGPContacts

If Not iNumNonPGPEntries = 0 Then
    MsgBox "There are addresses in either your TO and CC list that are not on your PGP keyring.  Can't process enryption request.", vbCritical + vbApplicationModal
Else
    txtCC.Text = AddressList.GetListAllContacts(CONTACT_CC_LIST)
    txtTo.Text = AddressList.GetListAllContacts(CONTACT_TO_LIST)
    PGPEncryptMessage True, False 'Sign and clearsign

End If

'Replace original
txtTo.Text = sToList
txtCC.Text = sCCList
'This clever little trick will push the original address fields back into the address list
txtTo.SelStart = 0
txtTo.SelLength = Len(txtTo.Text)
txtTo.SelUnderline = False
txtTo.SelLength = 0
MousePointer = vbDefault

End Sub
Private Sub PGPEyes_Click()

    PGPEyes.Checked = Not PGPEyes.Checked
End Sub
Private Sub PGPFile_Click()
    If Not PGPFile.Checked Then
        '---------------------------------------------
        'is the file option is not checked, this is the menu
        '---------------------------------------------
        PGPEncrypt.Caption = "&Encrypt file..."
        PGPDecrypt.Caption = "&Decrypt or verify file..."
        PGPEnSign.Caption = "Encrypt and &sign file..."
        PGPClearSign.Caption = "&Clear sign file..."
    Else
        '---------------------------------------------
        'is the file option is checked, this is the menu
        '---------------------------------------------
        PGPEncrypt.Caption = "&Encrypt message"
        PGPDecrypt.Caption = "&Decrypt or verify message"
        PGPEnSign.Caption = "Encrypt and &sign message"
        PGPClearSign.Caption = "&Clear sign message"
    End If
    '---------------------------------------------
    'toggle the state
    '---------------------------------------------
    PGPFile.Checked = Not PGPFile.Checked
End Sub
Private Sub PGPGetKey_Click()
    Dim foo As String
    Dim SectionName As String
    
    On Error GoTo GetKeyError
    If CheckConnection Then
        If Len(MessageArea) > 0 Then
            '---------------------------------------------
            'test first to see if message area is empty
            '---------------------------------------------
            foo = MsgBox("The message area contains text.  Is it okay to clear it?", vbYesNo, "Get Key From Server")
            If foo = vbYes Then
                MessageArea.Text = ""
            Else
                Exit Sub
            End If
        End If
        If Len(txtTo.Text) = 0 Then
            '---------------------------------------------
            'the user did not specify a recipient in the "to:" box
            '---------------------------------------------
            MessageArea.Text = "No user specified in the To: box." & vbCrLf
            MessageArea.Text = MessageArea.Text + "Please enter a valid e-mail address," & vbCrLf
            MessageArea.Text = MessageArea.Text + "or click on the right arrow button in the" & vbCrLf
            MessageArea.Text = MessageArea.Text + "To: box to choose a name from your address book."
            Exit Sub
        End If
        MailConnector.ServerState = HTTPSTATE
        gWebState = GETSERVERKEY
        If Not HTTP1.WinsockLoaded Then HTTP1.WinsockLoaded = True

        
        If gGetKeyURL = "" Then
            SectionName = "Net Info"
            gGetKeyURL = ReadProfile(SectionName, "GetKeyURL")
            If gGetKeyURL = "" Then
            'frmSelectKeyServer.Show vbModal
            gGetKeyURL = "http://pgp5.ai.mit.edu:11371/pks/lookup?op=get&exact=on&search="
            SectionName = "Net Info"
            WriteProfile SectionName, "GetKeyURL", gGetKeyURL
            'gGetKeyURL = "http://pgp5.ai.mit.edu:11371/pks/lookup?op=get&exact=on&search="
            'WriteProfile SectionName, "GetKeyURL", gGetKeyURL
            End If
        End If
        ShowStatus 1, ""
        ShowStatus 1, "Requesting key from server at " & gGetKeyURL & txtTo.Text
        DoEvents
        GetWebURL (gGetKeyURL & txtTo.Text)
        MailConnector.ServerState = 0
    End If
    Exit Sub
GetKeyError:
    HideStatus
    MsgBox Err.Description & " (in PGPGetKey)"
    MailConnector.ServerState = 0
    Err.Clear
End Sub

Private Sub PGPInsertDetachedSignature_Click()

Dim bRes As Boolean
'Dim sResponse As String
Dim sFileName As String
Dim sOutPutFileName As String

 '------------------------------------------------------------
    'First check if this is a file command
    '------------------------------------------------------------
    
    On Error GoTo SigError
    If Not PGPFile.Checked Then
         '------------------------
        'Just sign the message
        '--------------------------
        sFileName = GetTemporaryFile()
        Call PutFileText(sFileName, MessageArea.Text)
        ShowStatus 1, "Creating Detached Signature for Message Area."
    Else
        ShowStatus 1, "Creating Detached Signature for a File."
        CommonDialog1.DialogTitle = "Specify the file you wish to use to generate the 'Detached Signature'"
        CommonDialog1.Flags = &H2& + &H4&
        CommonDialog1.Filter = "All Files (*.*)|*.*|Document Files (*.doc)|*.doc"
        CommonDialog1.FilterIndex = 1
        CommonDialog1.CancelError = True
        CommonDialog1.InitDir = App.Path
        CommonDialog1.ShowOpen
        'If Not CommonDialog1.CancelError Then Exit Sub
        sFileName = CommonDialog1.FileName
        ChDrive Mid$(App.Path, 1, 3)
        ChDir App.Path
    End If

    '---------------------------------------------
    'handle case of saving signature into the message then saving to file
    '---------------------------------------------
    CommonDialog1.DialogTitle = "Specify the file to save the Detached Signature to."
    CommonDialog1.Flags = &H2& + &H4&
    CommonDialog1.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
    CommonDialog1.DefaultExt = ".txt"
    CommonDialog1.FilterIndex = 1
    CommonDialog1.CancelError = True
    CommonDialog1.InitDir = App.Path
           ' CommonDialog1.Action = 1
    CommonDialog1.ShowSave
    sOutPutFileName = CommonDialog1.FileName
    ChDrive Mid$(App.Path, 1, 3)
    ChDir App.Path
    bRes = PutFileText(sOutPutFileName, sPGPInsertDetachedSignature(sFileName))
    ShowStatus 1, "Detached signature saved in " & sFileName
    KillTemporaryFiles
Exit Sub

SigError:
 Err.Clear
 KillTemporaryFiles
 Exit Sub
End Sub

Private Sub PGPInsertKey_Click()
    InsertKey (IDKEY)
End Sub

Private Sub PGPMin_Click()

    PGPMin.Checked = Not PGPMin.Checked
    If PGPMin.Checked = True Then
        gMinState = 2
    Else
        gMinState = 1
    End If
End Sub

Private Sub PGPMultiple_Click()
    PGPMultiple.Checked = Not PGPMultiple.Checked
End Sub

Private Sub PGPObscurity_Click()

    If PGPObscurity.Checked = True Then
        PGPObscurity.Checked = False
        'gEncryptToRemailer = True
        gObscurity = 0
    Else
        PGPObscurity.Checked = True
        'gEncryptToRemailer = True
        gObscurity = 1
    End If
End Sub

Private Sub PGPOptions_Click()
    frmPGPOptions.Show vbModal
End Sub

Private Sub PGPSelf_Click()
    PGPSelf.Checked = Not PGPSelf.Checked
End Sub
Private Sub PGPClearSign_Click()
Dim sToList As String
Dim sCCList As String
Dim iNumPGPEntries As Integer
Dim iNumNonPGPEntries As Integer

'This will scan the Tx and To boxes just in case the timer did not go off
Call Timer1_Timer
MousePointer = vbHourglass

'Save original address lists
sToList = txtTo.Text
sCCList = txtCC.Text

ConvertMailGroupsToAddressList (CONTACT_TO_LIST)
ConvertMailGroupsToAddressList (CONTACT_CC_LIST)

iNumPGPEntries = AddressList.GetNumberPGPContacts
iNumNonPGPEntries = AddressList.GetNumberNonPGPContacts

If iNumNonPGPEntries > 0 Or iNumPGPEntries > 1 Then
    MsgBox "There are either multiple addresses in your TO list or addresses that are not on your PGP Keyring.  Can't process encryption request.", vbCritical + vbApplicationModal
    'MousePointer = vbDefault
    'Exit Sub
Else
    txtCC.Text = AddressList.GetListAllContacts(CONTACT_CC_LIST)
    txtTo.Text = AddressList.GetListAllContacts(CONTACT_TO_LIST)
    PGPEncryptMessage False, True 'Sign and clearsign

End If

'Replace original
txtTo.Text = sToList
txtCC.Text = sCCList
'This clever little trick will push the original address fields back into the address list
txtTo.SelStart = 0
txtTo.SelLength = Len(txtTo.Text)
txtTo.SelUnderline = False
txtTo.SelLength = 0
MousePointer = vbDefault



End Sub

Private Sub PGPWrap_Click()
    PGPWrap.Checked = Not PGPWrap.Checked
End Sub
Private Sub PrepareUseNetNymMessage()
Dim sMsg As String
'Dim RemailerOption As Integer
If MessageArea.Text = "" Then
    
    sMsg = "The messages area is blank. " & vbCrLf & vbCrLf
    sMsg = sMsg & "What you need to do is first add the news/message or attachment to the Message Area and then select this menu item."
    MsgBox sMsg, vbApplicationModal + vbCritical, "Message Blank"
    Exit Sub
End If

'We need to do this to have control over the To box
gCancelAction = False
gComposeMode = False
txtTo.Text = ""
txtCC.Text = ""
AddressList.Initialise

frmUSENETGateways.Show vbModal

'Need to have access to remailer data

frmMultiNyms.Show vbModal
ProcessNymCommand gNYM_USENET_PREPARE, Nym.ListIndex
RemoveUnderlineFromToBox
'RemoveUnderlineFromccBox

If gCancelAction Then
    gCancelAction = False
Else
    SendPIMessage
    If Not gCancelAction Then
        Unload Me
    End If
End If
gNymState = gNYM_IDLE
gCancelAction = False
End Sub

Private Sub popupSaveAttachmentAs_Click()
Dim Attachment As String
Dim SourceFile As String
Attachment = lvwAttachments.SelectedItem.Text 'App.Path & "\mailbox\attachments\" & lvwAttachments.SelectedItem.Text
            On Error Resume Next
            CommonDialog1.CancelError = True
            SourceFile = App.Path & "\temp\" & Attachment
            CommonDialog1.DialogTitle = "Save attachment as " & Attachment & " and Save file as:"
            CommonDialog1.Flags = &H2& + &H4&
            CommonDialog1.Filter = "File Type (*." & GetExt(Attachment) & ")"
            CommonDialog1.FilterIndex = 1
            'CommonDialog1.CancelError = True
            CommonDialog1.InitDir = App.Path
            CommonDialog1.FileName = Attachment 'IIf(InStr(1, StripExt(Attachment), ".") = 0, StripExt(Attachment) & ".htm", StripExt(Attachment)) 'TempPathLocation & StripExt(lvwAttachments.SelectedItem.Text)
            CommonDialog1.DefaultExt = GetExt(Attachment)
            CommonDialog1.Action = 2
            If Not Err.Number = 0 Then Exit Sub
            FileCopy SourceFile, CommonDialog1.FileName
            ChDrive Mid(App.Path, 1, 3)
            ChDir App.Path
           'vb2spgpContext.Initialise
           ' vb2spgpContext.FileIn = Attachment
            'vb2spgpContext.FileOut = CommonDialog1.FileName '& CommonDialog1.DefaultExt
            'spgpDecryptFile
            'ShowStatus 1, "Decrypted file was saved successfully"
End Sub

Private Sub Prepare_Usenet_Nym_Click()
PrepareUseNetNymMessage
End Sub
Private Sub Prepare_usenet_standard_Click()
Dim iRes As Integer
Dim msg As String
'Dim sMsg As String
If MessageArea.Text = "" Then
    
    msg = "The messages area is blank. " & vbCrLf & vbCrLf
    msg = msg & "What you need to do is first add the news/message or attachment to the Message Area and then select this menu item."
    MsgBox msg, vbApplicationModal + vbCritical, "Message Blank"
    Exit Sub
End If
gComposeMode = False
AddressList.Initialise
txtCC.Text = ""
txtTo.Text = ""


frmUSENETGateways.Show vbModal
'If Not gCancelAction Then
   ' msg = "The USENET message has now been prepared with the the mail headers containing the required IDs. " & vbCrLf & vbCrLf
    'msg = msg & "Do you wish to send the news through a Remailer, or directly to the mail2news service.  If you send if directly, there is a good chance your identity can be traced.  Also, the 'From: ' details will be taken from the setting in your email options area." & vbCrLf & vbCrLf
    'msg = msg & "If you wish to use a Remailer, click 'Yes'." & vbCrLf & vbCrLf
   ' msg = msg & "If you wish to send it directly, then click 'No'." & vbCrLf & vbCrLf
   ' msg = msg & "If you wish to cancel all together, then click 'Cancel'."
   ' iRes = MsgBox(msg, vbYesNoCancel + vbQuestion, "Send USENET Message")
   ' Select Case iRes
       ' Case vbNo
            'MailHeader(0).ID = "USENET"
            'MailHeader(0).Value = ""
           ' gRemailerType = STANDARD_EMAIL ' DontUseRemailer
            'SendPIMessage
        
        'Case vbYes
            gComposeMode = False
            AddressList.Initialise
            MailHeader(0).ID = "USENET"
            gNewsgroupType = USENET
           ' gRemailerType = REMAILER_CYPHERPUNK 'Load frmRemailerList
            SendPIMessage
            gNewsgroupType = 0

       ' Case vbCancel
   ' End Select
'End If
MailHeader(0).ID = ""
If Not gCancelAction Then Unload Me
gCancelAction = False
End Sub

Private Sub PrintSetup_Click()
CommonDialog1.Flags = &H40&
CommonDialog1.Action = 5
End Sub

Private Sub RemailerAppend_Click()
    
    'gEncryptToRemailer = False
    'USENETFi.Checked = False
    'UseNetSoda.Checked = False
    'USENETGate.Checked = False
    'USENETNone.Checked = True
    gNewsgroupType = 0
    'TransferAES.Checked = True
End Sub

Public Sub RemailerKeys()
    Dim tmpstr As String
    
    On Error GoTo RemailerKeysError
    
        
    If CheckConnection Then
        gWebState = GETREMAILERKEYS
        If Not HTTP1.WinsockLoaded Then HTTP1.WinsockLoaded = True
        '
        frmRemailerList.lblStatus = "Getting current remailer PGP keys."
        'ShowStatus 1, "Getting current remailer PGP keys."
        DoEvents
        If gPGPKeysURL = "" Then
            Dim SectionName As String
            SectionName = "Net Info"
            gPGPKeysURL = ReadProfile(SectionName, "PGPKeysURL")
            If gPGPKeysURL = "" Then
                'ShowStatus 1, "Invalid URL"
                frmRemailerList.lblStatus = "Invalid URL"
            End If
        End If
        DoEvents
        GetWebURL (gPGPKeysURL)
        HideStatus
        DoEvents
    End If
    Exit Sub
RemailerKeysError:
    HideStatus
    MsgBox Err.Description + "RemailerKeys"
    Err.Clear
End Sub




Private Sub RemailersCP_Click()

 UseCypherPunk
End Sub

Private Sub RemailersMix_Click()
SetUseOfMixmaster
End Sub

Public Sub RemailerUpdate()
Dim Response As String
    On Error GoTo RemailerError
    
    frmRemailerList.lblStatus.ForeColor = vbBlack
    If CheckConnection Then
        MailConnector.ServerState = HTTPSTATE
        '---------------------------------------------
        'fetch the URL for obtaining remailer data
        '---------------------------------------------
       ' If gRemailerInfoURL = "" Then
        
            Dim SectionName As String
            SectionName = "Net Info"
            If gRemailerType = REMAILER_MIX Then
                gMixListURL = ReadProfile(SectionName, "MixListURL")
                If Len(gMixListURL) = 0 Then
                   ' ShowStatus 1, "Not a valid URL."
                   frmRemailerList.lblStatus = "Not a valid URL."
                End If
                gMixType2URL = ReadProfile(SectionName, "MixType2URL")
                If Len(gMixType2URL) = 0 Then
                    'ShowStatus 1, "Not a valid URL."
                    frmRemailerList.lblStatus = "Not a valid URL."
                End If
                gMixPubRingURL = ReadProfile(SectionName, "MixPubRingURL")
                If Len(gMixPubRingURL) = 0 Then
                    'ShowStatus 1, "Not a valid URL."
                    frmRemailerList.lblStatus = "Not a valid URL."
                End If
            Else
                gRemailerInfoURL = ReadProfile(SectionName, "RemailerInfoURL")
                If Len(gRemailerInfoURL) = 0 Then
                    'ShowStatus 1, "Not a valid URL."
                    frmRemailerList.lblStatus = "Not a valid URL."
                End If
            End If
        'End If
        '---------------------------------------------
        'mixmaster option selected on menu
        '---------------------------------------------
        MousePointer = vbHourglass
        If gRemailerType = REMAILER_MIX Then
            gWebState = MIXUPDATE
            If Not HTTP1.WinsockLoaded Then HTTP1.WinsockLoaded = True
            frmRemailerList.lblStatus = "Downloading mixmaster list from " & gMixListURL
            DoEvents
            GetWebURL (gMixListURL)
            gWebState = TYPE2UPDATE
            If Not HTTP1.WinsockLoaded Then HTTP1.WinsockLoaded = True
            frmRemailerList.lblStatus = "Downloading mixmaster Type2.lis from " & gMixType2URL
            DoEvents
            GetWebURL (gMixType2URL)
            gWebState = PUBRINGUPDATE
            If Not HTTP1.WinsockLoaded Then HTTP1.WinsockLoaded = True
            frmRemailerList.lblStatus = "Downloading Mixmaster Pubring.mix file from " & gMixPubRingURL
            DoEvents
            GetWebURL (gMixPubRingURL)
        Else
            '---------------------------------------------
            'mixmaster is not checked on the menu
            '---------------------------------------------
            gWebState = GETREMAILERUPDATE
            If Not HTTP1.WinsockLoaded Then HTTP1.WinsockLoaded = True
            frmRemailerList.lblStatus = "Downloading remailer info from " & gRemailerInfoURL
            DoEvents
            GetWebURL (gRemailerInfoURL)
        End If
        If gCancelAction Then
            gCancelAction = False
            Exit Sub
        End If
       
        
        'And do the private stuff as well
       ' If iFileExists(App.Path & "\private.txt") Then
        '    frmRemailerList.InitializeRemailers (App.Path & "\private.txt")
        'End If
        '---------------------------------------------
        'set the newsgroup menu for cp state
        '---------------------------------------------
        'USENETGate.Visible = True
       ' USENETFi.Visible = False
        'UseNetSoda.Visible = False
        '------------------------------
        ' This will sort the remailers as well and fill the matched remailers list
        '  (strange place to put it...need to fix
        '-----------------------------------------
        'SortRemailers
        'FillRemailerList
    End If
    MousePointer = vbDefault
    Exit Sub
RemailerError:
    HideStatus
    MousePointer = vbDefault
    MsgBox Err.Description & "(in RemailerUpdate)"
    MailConnector.ServerState = 0
    Err.Clear
End Sub





Private Sub SelectKeyServer_Click()
frmSelectKeyServer.Show
End Sub

Private Sub SendSysInfo_Click()
   PrepareFeedback
End Sub

Private Sub SSRibbon1_Click(Index As Integer, Value As Integer)
Static UnloadForm As Boolean
'Dim bEncryptToSelf As String
'Dim SectionName As String

On Error Resume Next
SSRibbon1(Index).Value = False

DoEvents
If Value = 0 Then
    UnloadForm = False
    ShowStatus 1, ""
    'Unload frmRemailerList
    'Do this to ensure the addressee is picked up
    txtTo_LostFocus
    txtCC_LostFocus
      Select Case Index
        Case 0
            Select Case gRemailerType
                Case SEND_MESSAGES_USING_NYM
                    PrepareNymMessage
                
                Case ENCRYPT_BEFORE_SENDING_MESSAGE
                    Me.MousePointer = vbHourglass
                    gComposeMode = True
                   ' Load frmRemailerList
                  '  frmRemailerList.Caption = "Cypherpunk Remailer List"
                    SendPIMessage
                    Me.MousePointer = vbDefault
                
                Case SIGN_BEFORE_SENDING_MESSAGE
                    Me.MousePointer = vbHourglass
                    gComposeMode = True
                   ' Load frmRemailerList
                   ' frmRemailerList.Caption = "Cypherpunk Remailer List"
                    SendPIMessage
                    Me.MousePointer = vbDefault
            
                 Case ENCRYPT_AND_SIGN_BEFORE_SENDING_MESSAGE
                    Me.MousePointer = vbHourglass
                    gComposeMode = True
                   ' Load frmRemailerList
                   ' frmRemailerList.Caption = "Cypherpunk Remailer List"
                   SendPIMessage
                    Me.MousePointer = vbDefault
                    
                Case REMAILER_CYPHERPUNK
                    Me.MousePointer = vbHourglass
                    gComposeMode = True
                   ' Load frmRemailerList
                   ' frmRemailerList.Caption = "Cypherpunk Remailer List"
                    SendPIMessage
                    Me.MousePointer = vbDefault
                
                Case REMAILER_MIX
                    Me.MousePointer = vbHourglass
                    gComposeMode = True
                  '  Load frmRemailerList
                  '  frmRemailerList.Caption = "Mixmaster Remailer List"
                    SendPIMessage
                    Me.MousePointer = vbDefault
                
                Case STANDARD_EMAIL
                    'SendToOutBox
                    Me.MousePointer = vbHourglass
                    gComposeMode = True
                    SendToOutBox
                    Me.MousePointer = vbDefault
            
            End Select
            
            If gCancelAction Then
                gCancelAction = False
            Else
             '   Unload frmRemailerList
                UnloadForm = True 'Unload Me
            End If
                        
        Case 1
            EditPerform WM_COPY
            
        Case 2
            EditPerform WM_PASTE
        Case 3
            txtTo.Text = ""
            txtsubject.Text = ""
            txtCC.Text = ""
            MessageArea.Text = ""
            
        Case 4
            ImportFile
        Case 5
            SaveMessage
        Case 6
            If (gRemailerType = REMAILER_MIX) Or (gRemailerType = REMAILER_CYPHERPUNK) Then AppendInfo
        Case 7
            AddAttachment
        Case 8
            EditPerform WM_CUT
        Case 9
            
                PGPEncrypt_Click

        Case 10
            gComposeMode = True
            frmMain.ReplyToSender (iLocalInstanceReference)
            
        Case 11
            DecodeMessage
        Case 12
            gComposeMode = True
            frmMain.ForwardMessage (iLocalInstanceReference)
        Case 13
            Dim bRes As Boolean
           ' WipeFile (TempPathLocation & "temp.html")
            If PIForm(iLocalInstanceReference).WebBrowser1.Visible = True Then Exit Sub
            bRes = PutFileText(TempPathLocation & "temp.htm", Trim(PIForm(gActivePIInstance).MessageArea.Text))
            PIForm(iLocalInstanceReference).MessageArea.Visible = False
            PIForm(iLocalInstanceReference).WebBrowser1.Visible = True
            DoEvents
            PIForm(iLocalInstanceReference).WebBrowser1.Navigate (TempPathLocation & "temp.htm")
            'WipeFile (TempPathLocation & "temp.html")

    End Select
    'MousePointer = vbDefault
Else
    If UnloadForm Then Unload Me
End If
End Sub



Private Sub StepAddKey_Click()

    Form32.Caption = "Adding a public key"
    'Form32.Text1.Text = "Before you can send someone an encrypted message, you need a copy of their public key.  To add a copy of your key:" + gCRLF + gCRLF + "1. - Copy the public key (either from an e-mail message or key server) and paste it into the Message text box." + gCRLF + gCRLF + "2. - From the Keys menu, choose the 'Add key from message' command." + gCRLF + gCRLF + "3. - PGP will run in the DOS window.  If you are running Windows 95, click the icon in the taskbar.  Certify the key." + gCRLF + gCRLF + "The key is added to your public key ring."
    Form32.RichTextBox1 = GetFileText(App.Path & "\HelpText8.rtf")
    Form32.Show
End Sub

Private Sub StepAttach_Click()

    Form32.Caption = "Sending an attachment"
    Form32.RichTextBox1 = GetFileText(App.Path & "\HelpText5.rtf")
    'Form32.Text1.Text = "You can attach a file to a message sent from Private i Mail.  To send a message with an attachment:" + gCRLF + gCRLF + "1. - Compose the message in the Message text box." + gCRLF + gCRLF + "2. - In the To: text box, enter the e-mail address of the person you'll be sending the message to." + gCRLF + gCRLF + "3. - Make sure you have a connection to the Internet if you are using the 16-bit version." + gCRLF + gCRLF + "4. - Check the Attachment checkbox and specify the file to attach." + gCRLF + gCRLF + "5. - In the drop-down box, specify if you'd like the file encrypted, you may also want to encrypt the text message before you send it." + gCRLF + gCRLF + "6. - Select the Send button (or choose 'Send' from the Message menu)." + gCRLF + gCRLF + "Note: The file is Base64 encoded to a MIME compliant attachment." + gCRLF + gCRLF + "Private i Mail currently doesn't support sending attachments through remailers."
    Form32.Show
End Sub

Private Sub StepCreateKey_Click()

    Form32.Caption = "Creating a PGP key pair"
    'Form32.Text1.Text = "If you'd like to create a PGP secret and public key to use with a Nym.ID:" + gCRLF + gCRLF + "1. - From the Keys menu, choose the 'Create key pair' command." + gCRLF + gCRLF + "2. - PGP will run in the DOS window.  Follow the steps for creating a key." + gCRLF + gCRLF + "Hint: Use a key size of 1024 bits or higher."
    Form32.RichTextBox1 = GetFileText(App.Path & "\HelpText10.rtf")
    Form32.Show
End Sub

Private Sub StepDecrypt_Click()

    Form32.Caption = "Decrypting a message"
    Form32.RichTextBox1 = GetFileText(App.Path & "\HelpText2.rtf")
    Form32.Show
End Sub

Private Sub StepDelete_Click()

    Form32.Caption = "Deleting a public key"
    'Form32.Text1.Text = "To remove a key from your public key ring:" + gCRLF + gCRLF + "1. - From the Keys menu, choose the 'Delete key' command." + gCRLF + gCRLF + "2. - Select the key to remove and click OK." + gCRLF + gCRLF + "3. - PGP will run in the DOS window.  If you are running Windows 95, click the icon in the taskbar.  Verify you want to remove the key." + gCRLF + gCRLF + "The key is removed from your public key ring."
    Form32.RichTextBox1 = GetFileText(App.Path & "\HelpText9.rtf")
    Form32.Show
End Sub

Private Sub StepEncrypt_Click()

    Form32.Caption = "Encrypting a message"
    Form32.RichTextBox1 = GetFileText(App.Path & "\HelpText1.rtf")
    Form32.Show
End Sub

Private Sub StepGetKey_Click()

    'Form32.Caption = "Sending/getting MIT server keys"
    Form32.RichTextBox1 = GetFileText(App.Path & "\HelpText7.rtf")
    Form32.Show
End Sub

Private Sub StepInfo_Click()

    Form32.Caption = "Internet privacy info"
    'Form32.Text1.Text = "Private i Mail has links to a variety of Internet privacy sources.  To access them:" + gCRLF + gCRLF + "1. - Ensure Private i Mail can communicate with your Web browser.  The default is Netscape Navigator.  If you're using another browser, choose the Options command in the Web menu." + gCRLF + gCRLF + "2. - You should be connected to the Internet with the browser running and not minimized." + gCRLF + gCRLF + "3. - From the Web menu, choose the information you'd like to access." + gCRLF + gCRLF + "Your browser will display the associated Web page."
    Form32.RichTextBox1 = GetFileText(App.Path & "\HelpText20.rtf")
    Form32.Show
End Sub

Private Sub StepNym_Click()
   Form32.Caption = "Creating a Nym"
    'Form32.Text1.Text = "A Nym(as in ano'Nym'ous) is an alias used for private communications.  Once you've created a Nym.ID account, you can send messages through it to people.  Unlike anonymous remailers, they can reply back to you without knowing your identity.  Various free servers are available for setting up Nym.ID accounts.  These are much more secure than using anon.penet.fi.  To create a Nym.ID:" + gCRLF + gCRLF + "1. - From the Nym.ID menu, choose the 'Create Nym.ID' command." + gCRLF + gCRLF + "This command steps you through the entire Nym.ID creation process with a series of easy to follow dialog boxes." + gCRLF + gCRLF + "Refer to the on-line help for additional information."
    Form32.RichTextBox1 = GetFileText(App.Path & "\HelpText14.rtf")
    Form32.Show
End Sub

Private Sub StepNymDelete_Click()

    Form32.Caption = "Deleting a Nym"
    'Form32.Text1.Text = "To delete a Nym:" + gCRLF + gCRLF + "1. - From the Nym menu, choose the 'Delete Nym' command." + gCRLF + gCRLF + "2. - Select the Nym account to delete and click OK." + gCRLF + gCRLF + "3. - Send the prepared message to the Nym server.  Your Nym will be deleted."
    Form32.RichTextBox1 = GetFileText(App.Path & "\HelpText18.rtf")
    Form32.Show
End Sub

Private Sub StepNymPass_Click()

    Form32.Caption = "Changing a Nym password"
    'Form32.Text1.Text = "If you want to change your Nym password:" + gCRLF + gCRLF + "1. - From the Nym menu, choose the 'Change Nym password' command." + gCRLF + gCRLF + "2. - Select the Nym.ID to change and click OK." + gCRLF + gCRLF + "3. - In the Message text box enter your old and new passwords." + gCRLF + gCRLF + "4. - From the Nym.ID menu, choose the 'Encrypt Nym message' command." + gCRLF + gCRLF + "5. Send the message."
    Form32.RichTextBox1 = GetFileText(App.Path & "\HelpText16.rtf")
    Form32.Show
End Sub

Private Sub StepNymReply_Click()


    Form32.Caption = "Changing a Nym reply block"
    Form32.RichTextBox1 = GetFileText(App.Path & "\HelpText20.rtf")
    Form32.Show
End Sub

Private Sub StepNymSend_Click()


    Form32.Caption = "Sending a Nym message"
    'Form32.Text1.Text = "Once you've created a Nym.ID account, you can send messages through it.  To do so:" + gCRLF + gCRLF + "1. - Compose the message in the Message text box." + gCRLF + gCRLF + "2. - In the To: text box, enter the e-mail address of the person you'll be sending the message to." + gCRLF + gCRLF + "3. - From the Nym.ID menu, choose the 'Prepare Nym.ID message' command." + gCRLF + gCRLF + "4. - Select the Nym.ID account to use and click OK." + gCRLF + gCRLF + "5. - For alias type nyms, from the Nym.ID menu, choose the 'Encrypt Nym.ID message' command." + gCRLF + gCRLF + "7. Send the message."
    Form32.RichTextBox1 = GetFileText(App.Path & "\HelpText15.rtf")
    Form32.Show
End Sub

Private Sub StepRemailer_Click()


    Form32.Caption = "Anonymous messages"
    'Form32.Text1.Text = "You can send some a message without revealing your identity by using an anonymous remailer.  To send an anonymous message:" + gCRLF + gCRLF + "1. - Compose the message in the Message text box." + gCRLF + gCRLF + "2. - In the To: text box, enter the e-mail address of the person you'll be sending the message to." + gCRLF + gCRLF + "3. - From the Remailers menu, choose the type of remailer to use." + gCRLF + gCRLF + "4. - Select the remailer to send the message through from the Remailer drop-down list.  Selecting 'chain' routes the message through several remailers." + gCRLF + gCRLF + "5. - From the Message menu, choose the 'Append info' command.  This formats the message for sending through a remailer.  If you selected 'chain,' a dialog box will prompt you for the remailers to use." + gCRLF + gCRLF + "6. - Send the message." + gCRLF + gCRLF + "Note: The Cypherpunk type remailers support a variety of advanced features.  Refer to the on-line help."
    Form32.RichTextBox1 = GetFileText(App.Path & "\HelpText11.rtf")
    Form32.Show
End Sub

Private Sub StepSend_Click()

    Form32.Caption = "Sending a message from Private i Mail"
    Form32.RichTextBox1 = GetFileText(App.Path & "\HelpText4.rtf")
    'Form32.Text1.Text = "If you have a connection to the Internet you can send a message directly from Private i Mail.  To send a message:" + gCRLF + gCRLF + "1. - Compose the message in the Message text box." + gCRLF + gCRLF + "2. - In the To: text box, enter the e-mail address of the person you'll be sending the message to." + gCRLF + gCRLF + "3. - Make sure you have a connection to the Internet." + gCRLF + gCRLF + "4. - Select the Send button (or choose 'Send' from the Message menu)." + gCRLF + gCRLF + "Note: Before sending a message, you need to provide information about your e-mail server.  From the File menu, choose the 'Options' command." + gCRLF + gCRLF + "If you can't send e-mail directly from Private i Mail, you can easily transfer the message to an e-mail application such as Eudora or Pegasus.  Refer to the on-line help."
    Form32.Show
End Sub

Private Sub StepSendKey_Click()

    Form32.Caption = "Sending your public key"
    'Form32.Text1.Text = "Before someone can send you an encrypted message, they need a copy of your public key.  To send a copy of your key:" + gCRLF + gCRLF + "1. - From the Keys menu, choose the 'Insert key in message' command." + gCRLF + gCRLF + "2. - Select your key from the user ID dialog box and click OK." + gCRLF + gCRLF + "3. - PGP will run and fetch the key from your public key ring." + gCRLF + gCRLF + "4. - The public key is inserted in the Message text box.  You can now send the key to someone you want to privately correspond with."
    Form32.RichTextBox1 = GetFileText(App.Path & "\HelpText6.rtf")
    Form32.Show
End Sub

Private Sub StepSign_Click()


    Form32.Caption = "Signing a message"
    Form32.RichTextBox1 = GetFileText(App.Path & "\HelpText3.rtf")
    'Form32.Text1.Text = "There are two ways to sign a message.  One way is to clear-sign a message.  This method leaves the text alone, but wraps it with a signature.  Signing the message must be the last step as the signature depends on the contents of the message.  The second way is to sign an encrypted file.  Either way, the intent is to let the recipient know that you are the one who created the message because it requires your secret key and password to compute the signature." + gCRLF + gCRLF + "To sign a message:" + gCRLF + gCRLF + "1. - Compose the message in the Message text box." + gCRLF + gCRLF + "2. - From the PGP menu, choose the 'Clear sign message' or 'Encrypt and Sign' command." + gCRLF + gCRLF + "3. - PGP will run in the DOS window.  Enter your passphrase." + gCRLF + gCRLF + "A signature is attached to the message."
    Form32.Show
End Sub

Private Sub StepUpdateInfo_Click()
    Form32.Caption = "Updating remailer info"
    'Form32.Text1.Text = "To update remailer info:" + gCRLF + gCRLF + "1. - Make sure you have a Net connection." + gCRLF + gCRLF + "2. - From the Remailers menu, choose the 'Update remailer info' command." + gCRLF + gCRLF + "Private i Mail will connect to Raph's Web page and download the current remailer data and update the remailer list box.  If Cypherpunk is checked in the Remailer menu, Cypherpunk info is updated.  If Mixmaster is checked, Mixmaster info is updated."
    Form32.RichTextBox1 = GetFileText(App.Path & "\HelpText13.rtf")
    Form32.Show
End Sub

Private Sub StepUSENET_Click()
Dim zap As String
    Form32.Caption = "USENET articles"
    zap = "You can post a USENET article without revealing your identity by using a remailer. "
    zap = zap & "To post an anonymous article:" + gCRLF + gCRLF
    zap = zap & "1. - Compose the message in the Message text box." + gCRLF + gCRLF
    zap = zap & "2. - Click on 'Newsgroups' and fill in the required info. " & vbCrLf & vbCrLf
    zap = zap & "3. - Click on 'Prepare USENET Message' to creat the Nym message (you would have previously created a Nym." & vbCrLf & vbCrLf
    zap = zap & "4. - Send using a remailer"
    '
    'Form32.Text1.Text = zap
    Form32.RichTextBox1 = GetFileText(App.Path & "\HelpText12.rtf")
    Form32.Show
End Sub

Private Sub StepWeb_Click()


    Form32.Caption = "Anonymous Web page access"
    'Form32.Text1.Text = "It's very easy for someone to log Web pages you visit.  "
    'Form32.Text1.Text = Form32.Text1.Text & " Community Connexion (c2.org), has a free Web anonymizer service. "
    'Form32.Text1.Text = Form32.Text1.Text & "To use it from Private Idaho Mail:" + gCRLF + gCRLF + "1. - Type the Web page URL you want to anonymously visit in the Message text box and select (highlight) it." + gCRLF + gCRLF + "2. Ensure Private i Mail can communicate with your Web browser.  The default is Netscape Navigator.  If you're using another browser, choose the Options command in the Web menu." + gCRLF + gCRLF + "3. - You should be connected to the Internet with the browser running and not minimized." + gCRLF + gCRLF + "4. - From the Web menu, choose the 'Anonymous jump to URL' command." + gCRLF + gCRLF + "Your browser will anonymously access the Web page." + gCRLF + gCRLF + "Hint: Enter frequently accessed Web pages in Private i Mail's address book."
    Form32.RichTextBox1 = GetFileText(App.Path & "\HelpText19.rtf")
    Form32.Show
End Sub
Private Sub Timer1_Timer()
If m_fToFieldChanged Then ScanTextForContacts txtTo, CONTACT_TO_LIST
m_fToFieldChanged = False
If m_fCCFieldChanged Then ScanTextForContacts txtCC, CONTACT_CC_LIST
m_fCCFieldChanged = False
Timer1.Enabled = False
If frmMain.CheckLicenceExpired Then ShowStatus 1, "Trail period expired."
End Sub

Private Sub TransferApp1_Click()

    Dim SectionName As String
    Dim AppName As String
    Dim AppScript As String
    Dim WindApp As String
    
    SectionName = "Options"
    WindApp$ = ReadProfile(SectionName, "App1Wind")
    AppScript = ReadProfile(SectionName, "App1Script")
    TransferInfo WindApp$, AppScript
    
End Sub

Private Sub TransferApp2_Click()


    Dim SectionName As String
    Dim AppName
    Dim AppScript As String
    Dim WindApp As String
    
    SectionName = "Options"
    WindApp = ReadProfile(SectionName, "App2Wind")
    AppScript = ReadProfile(SectionName, "App2Script")
    TransferInfo WindApp, AppScript
    
End Sub

Private Sub TransferApp3_Click()


    Dim SectionName As String
    Dim AppName
    Dim AppScript As String
    Dim WindApp As String
    
    SectionName = "Options"
    WindApp$ = ReadProfile(SectionName, "App3Wind")
    AppScript = ReadProfile(SectionName, "App3Script")
    TransferInfo WindApp$, AppScript
    
End Sub

Private Sub TransferApp4_Click()


    Dim WindApp As String
    Dim SectionName As String
    Dim AppName
    Dim AppScript As String
    
    SectionName = "Options"
    WindApp$ = ReadProfile(SectionName, "App4Wind")
    AppScript = ReadProfile(SectionName, "App4Script")
    TransferInfo WindApp$, AppScript
End Sub


   

'Private Sub TransferAT_Click()
   ' If Not gRemailerType = STANDARD_EMAIL Then
    '    AppendInfo
    '    TransferInfo gEmailer, gtranScript
   ' End If
'End Sub

Private Sub TransferEncrypt_Click()
    EncryptMessage txtTo.Text, "-eatw"
End Sub

Private Sub TransferEu_Click()
    TransferInfo gEmailer, gtranScript
End Sub

Private Sub TransferNym_Click()
    gComposeMode = False
    AddressList.Initialise
    txtCC.Text = ""
    txtTo.Text = ""
    cmbRemailerSelect.ListIndex = 1
    Nym.create = True
    frmCreateNymStep1.Show
    cmbRemailerSelect.ListIndex = 0
End Sub
Private Sub PrepareNymMessage() 'TransferPrepare_Click()
Dim TempFile As String
Dim i As Integer
Dim Vbresponse As Integer
Dim msg As String

'We need to do this to have control over the To box
'gComposeMode = False
InsertAttachmentsIntoMessage
AddressList.Initialise
RemoveUnderlineFromToBox
RemoveUnderlineFromccBox
ScanTextForContacts txtTo, CONTACT_TO_LIST
ScanTextForContacts txtCC, CONTACT_CC_LIST
ConvertMailGroupsToAddressList (CONTACT_TO_LIST)
ConvertMailGroupsToAddressList (CONTACT_CC_LIST)
txtTo.Text = AddressList.GetListAllContacts(CONTACT_TO_LIST)
txtCC.Text = AddressList.GetListAllContacts(CONTACT_CC_LIST)

frmMultiNyms.Show vbModal
If gCancelAction Then Exit Sub
gNymState = gNYMPREPARE
ProcessNymCommand gNymState, Nym.ListIndex
'AddressList.Initialise
If gCancelAction Then
    txtTo.Text = ""
Else
    'Note: this will be set as an option later...
    Nym.DontUseRemailer = True
    
    'Not only To is valid to send to nym server
    'AddressList.Initialise
    'ScanTextForContacts txtTo, CONTACT_TO_LIST
    AddressList.Initialise
    RemoveUnderlineFromToBox
    ScanTextForContacts txtTo, CONTACT_TO_LIST
    SendNymMessage
    If Not gCancelAction Then Unload Me
End If

End Sub

Private Sub TransferReply_Click()
Dim foo As String
Dim fie As String
Dim tmpChar As String
Dim i As Integer
    MousePointer = vbHourglass
    If MessageArea.Text <> "" Then
        fie = ">"
        foo = InsertCRLFs()
        For i = 1 To Len(foo)
            tmpChar = Mid(foo, i, 1)
            fie = fie + tmpChar
            If tmpChar = Chr(10) Then
                fie = fie + ">"
            End If
        Next
        MessageArea.Text = fie
        txtTo.Text = gMessageRecord.From
    End If
    MousePointer = vbDefault
End Sub

Private Sub EditSetFont_Click()

    '---------------------------------------------
    'display the common font dialog
    '---------------------------------------------
    CommonDialog1.Flags = cdlCFBoth
    CommonDialog1.ShowFont
    
    '---------------------------------------------
    'set the font for the message area
    '---------------------------------------------
    MessageArea.SelFontName = CommonDialog1.FontName
    MessageArea.SelBold = CommonDialog1.FontBold
    MessageArea.SelItalic = CommonDialog1.FontItalic
    MessageArea.SelStrikeThru = CommonDialog1.FontStrikethru
    MessageArea.SelFontSize = CommonDialog1.FontSize
    
    '---------------------------------------------
    'set the font for the to: box
    '---------------------------------------------
    txtTo.Font = CommonDialog1.FontName
    txtTo.SelBold = CommonDialog1.FontBold
    txtTo.SelItalic = CommonDialog1.FontItalic
    txtTo.SelStrikeThru = CommonDialog1.FontStrikethru

    '---------------------------------------------
    'set the font for the subject: box
    '---------------------------------------------
    txtsubject.Font = CommonDialog1.FontName
    txtsubject.FontBold = CommonDialog1.FontBold
    txtsubject.FontItalic = CommonDialog1.FontItalic
    txtsubject.FontStrikethru = CommonDialog1.FontStrikethru

    '---------------------------------------------
    'set the font for the cc: box
    '---------------------------------------------
    txtCC.Font = CommonDialog1.FontName
    txtCC.SelBold = CommonDialog1.FontBold
    txtCC.SelItalic = CommonDialog1.FontItalic
    txtCC.SelStrikeThru = CommonDialog1.FontStrikethru

    '---------------------------------------------
    'set the font for the bcc: box
    '---------------------------------------------
    'Text4.Font = CommonDialog1.FontName
   ' Text4.FontBold = CommonDialog1.FontBold
    'Text4.FontItalic = CommonDialog1.FontItalic
   ' Text4.FontStrikethru = CommonDialog1.FontStrikethru

End Sub

Public Property Get BusyCancel() As Boolean
    BusyCancel = m_BusyCancel
End Property

Public Property Let BusyCancel(ByVal bBusyCancel As Boolean)
    m_BusyCancel = bBusyCancel
End Property

Public Sub InitialiseDisplay()

If iFileExists(App.Path + "\remailer.txt") Then
    frmRemailerList.InitializeRemailers (App.Path + "\remailer.txt")
Else
    If iFileExists(App.Path + "\remailer.htm") Then frmRemailerList.InitializeRemailers (App.Path + "\remailer.htm")
End If
End Sub

Private Sub SendPIMessage()
Dim sAddressList As String
Dim VarArray As Variant
Dim MemberArray As Variant
Dim iResponse As Integer
Dim i As Integer
Dim iNumPGPEntries As Integer
Dim iNumNonPGPEntries As Integer
Dim sMember As String
Dim bMixedAddressTypes As Boolean
Dim sResult As String
Dim SectionName As String
Dim bEncryptToSelf As String


On Error GoTo BadSend
'Get Encrypt to self option
SectionName = "PGP Options"
bEncryptToSelf = ReadProfile("PGP Options", "EncryptToSelf")
If bEncryptToSelf = "true" Then EncryptToSelftandSendToSentBox
 
    'If not compose mode - we are probably here due to nym, mail2news or remailer activity
    If Not gComposeMode Then
        AddressList.Initialise
        RemoveUnderlineFromToBox
        RemoveUnderlineFromccBox
       ScanTextForContacts txtTo, CONTACT_TO_LIST
        ScanTextForContacts txtCC, CONTACT_CC_LIST
    End If
    gCancelAction = False
    MessageArea.SetFocus
    
    'Convert MailGroups to address list
    ConvertMailGroupsToAddressList (CONTACT_TO_LIST)
    ConvertMailGroupsToAddressList (CONTACT_CC_LIST)
               
    iNumPGPEntries = AddressList.GetNumberPGPContacts
    iNumNonPGPEntries = AddressList.GetNumberNonPGPContacts

    DoEvents
    If (gRemailerType = ENCRYPT_BEFORE_SENDING_MESSAGE) Or (gRemailerType = ENCRYPT_AND_SIGN_BEFORE_SENDING_MESSAGE) Then
        If iNumPGPEntries = 0 Then
            MsgBox "You have selected the option to encrypt using public keys.  PGP can't find the key for the recepient that you have specified." & vbCrLf & vbCrLf & "Either modify your send option or select a recepient on your keyring.", vbApplicationModal + vbCritical, "Public Key not Found"
            gCancelAction = True
            Exit Sub
        End If
    
        If iNumPGPEntries > 0 And iNumNonPGPEntries > 0 Then
            iResponse = MsgBox("You have selected some addressees from your public keyring and others that are not on your public keyring.  Private Idaho will generally automatically encrypt messages if the addressees are all taken from you public keyring. " & vbCrLf & vbCrLf & "Note: If you have selected a remailer then the message is not encrypted with the public keys contained in your addressee list." & vbCrLf & vbCrLf & "If you proceed, then the message will not be encrypted with any of the public keys.", vbApplicationModal + vbOKCancel + vbQuestion, App.Title)
            If iResponse = vbCancel Then
                gCancelAction = True
                Exit Sub
            End If
            bMixedAddressTypes = True
        End If
    End If
    If (gRemailerType = REMAILER_CYPHERPUNK) Or (gRemailerType = REMAILER_MIX) Then
        txtTo.Text = AddressList.GetListAllContacts(CONTACT_TO_LIST)
        txtCC.Text = AddressList.GetListAllContacts(CONTACT_CC_LIST)
        InsertAttachmentsIntoMessage
        AppendInfo True ' Only bring up the dialog box the first time, false does not bring it up
        'if user cancelled remailers then don't send
        If Not gCancelAction Then
            ShowStatus 1, "Sending to PI OutBox folder."
            DoEvents
            If Not gMixSent Then
               ' If bEncryptToSelf = "true" Then EncryptToSelftandSendToSentBox
                SendToOutBox
            Else
                gMixSent = False
            End If
        Else
            ShowStatus 1, "Process cancelled - Message not sent."
            Beep
        End If
    Else
        If bMixedAddressTypes Then
            'Okay prepare the address lists for non PGP stuff
                txtTo.Text = AddressList.GetListAllContacts(CONTACT_TO_LIST)
                txtCC.Text = AddressList.GetListAllContacts(CONTACT_CC_LIST)
                If Not txtTo.Text = "" Then
                   ' If bEncryptToSelf = "true" Then EncryptToSelftandSendToSentBox
                    SendToOutBox
                End If
        Else
                
                 'Now get Mail Group members
                txtTo.Text = AddressList.GetListAllNonPGPContacts(CONTACT_TO_LIST)
                txtCC.Text = AddressList.GetListAllNonPGPContacts(CONTACT_CC_LIST)
                If Not txtTo.Text = "" Then
                    SendToOutBox
                End If
                'Now send to PGP Addressees
                If Not iNumPGPEntries = 0 Then
                                       
                    txtTo.Text = AddressList.GetListAllPGPContacts(CONTACT_TO_LIST)
                    txtCC.Text = AddressList.GetListAllPGPContacts(CONTACT_CC_LIST)
                     
                     If gRemailerType = SIGN_BEFORE_SENDING_MESSAGE Then
                         PGPEncryptMessage False, True
                    End If
                    
                    If gRemailerType = ENCRYPT_AND_SIGN_BEFORE_SENDING_MESSAGE Then
                        PGPEncryptMessage True, False 'Sign and clearsign
                    End If
                    If gRemailerType = ENCRYPT_BEFORE_SENDING_MESSAGE Then
                        PGPEncryptMessage False, False 'Sign and clearsign
                    End If
                   ' If bEncryptToSelf = "true" Then EncryptToSelftandSendToSentBox True
                    SendToOutBox True 'Assumes attachments have been encrypted when sending to outbox
                End If
        End If
            KillTemporaryFiles
        End If
    Exit Sub
BadSend:
    gCancelAction = False
    MsgBox "A serious error has occured trying to send the message to your 'OutBox'.  The process will abort here.  For your information, the following error occured: " & Err.Description, vbApplicationModal + vbCritical, App.Title
    Err.Clear
    Me.MousePointer = vbDefault
    Exit Sub
    
End Sub

Private Sub ImportFile()
Dim FileNum As Integer
    Dim TextLine As String
    Dim NumBytes As Long
    Dim msg As String
    Dim foo As Long
    Dim FileSize As Long
    Dim LineCount As Long
    
    On Error GoTo ImportError
    CommonDialog1.DialogTitle = "Open message text file"
    CommonDialog1.Flags = &H2& + &H4&
    CommonDialog1.Filter = "Text Files (*.txt)|*.txt|Asc Files (*.asc)|*.asc|All Files (*.*)|*.*"
    CommonDialog1.FilterIndex = 1
    CommonDialog1.CancelError = True
    CommonDialog1.InitDir = App.Path
    CommonDialog1.Action = 1
    FileNum = FreeFile
    Open CommonDialog1.FileName For Input As FileNum
        If Len(MessageArea.Text) > 0 Then
            foo = MsgBox("The message area contains text.  Is it okay to insert into the current location?", vbYesNo, "Send Feedback")
            If foo = vbNo Then
                Exit Sub
            End If
        End If
       ' MessageArea.SelStart = 0
       ' MessageArea.Text = ""
        msg = ""
        MousePointer = vbHourglass
        ShowStatus 1, "Loading file " & CommonDialog1.FileName & "..."
        DoEvents
        FileSize = FileLen(CommonDialog1.FileName)
        frmBusy.Style = 1
        frmBusy.AllowCancel = True
       ' frmBusy.BarCaption =
        frmBusy.CallingForm = frmPI
        frmBusy.Message = "Loading file " & CommonDialog1.FileName & ".  Please wait."
        frmBusy.Show
        Me.BusyCancel = False
        
        Do While Not EOF(FileNum)
            Line Input #FileNum, TextLine
            LineCount = LineCount + 1
            msg = msg & TextLine & vbCrLf
            NumBytes = NumBytes + Len(TextLine)
            frmBusy.BarPercent = 100 * NumBytes / FileSize
            frmBusy.BarCaption = "Processing line " & LineCount
            If Me.BusyCancel Then Exit Do
            DoEvents
        Loop
    MessageArea.SelText = msg
    MousePointer = vbDefault
    Unload frmBusy
    'End If
    Close FileNum
    ChDir App.Path
    Exit Sub

ImportError:
    Reset
    MsgBox Err.Description & " in FileImport", vbApplicationModal, App.Title
    Err.Clear
End Sub

Private Sub SaveMessage()
Dim bRes As Boolean
Dim rs As Recordset
Dim MessageFileName As String
Dim lFolderID As Long
    
    'Write to datbase here....
    '
    On Error GoTo WriteMessError
    'First Find the Inbox folder
    Set rs = DB.OpenRecordset("Folders", dbOpenDynaset)
    rs.FindFirst "[Folder] =" & "'" & "Drafts" & "'"
    If rs.NoMatch Then
        Err.Raise 1002, , "Draft folder missing from database."
    End If
    lFolderID = rs("folder id")
    Set rs = DB.OpenRecordset("Messages", dbOpenDynaset)
    'Look for current message number in Drafts, if there write over
    If Not MessageID = 0 Then
        'Okay, this is a message that came from the draft areas
        rs.FindFirst "[Message ID] =" & MessageID '"'" & MessageID & "'"
        If rs.NoMatch Then
            rs.AddNew
            MessageID = 0
        Else
            rs.Edit
        End If
    'If rs("Message ID") = MessageID Then
    Else
        rs.AddNew
        'This flags that the message has been saved
        MessageID = rs("Message ID")
    End If
    'rs.AddNew
    rs("Folder ID") = lFolderID
    rs("To") = txtTo.Text
    rs("From") = MailConnector.EmailAddress
    rs("Date Sent") = CDate(Now())
    MessageFileName = DatePart("d", Now())
    MessageFileName = MessageFileName & DatePart("m", Now())
    MessageFileName = MessageFileName & Mid$(DatePart("yyyy", Now()), 3, 2)
    MessageFileName = MessageFileName & rs("Message ID") & ".pim"
    rs("Subject") = txtsubject.Text
    rs("CC") = txtCC.Text
    bRes = PutFileText(App.Path & "\mailbox\" & MessageFileName, MessageArea.Text)
    rs("Attachment") = False
    rs("MIME Message") = MessageFileName
    
    'This is crude but last thing is to strip out PGP message
    rs("Message Read") = False
    rs("Incoming Message") = False
    rs("Message Deleted") = False
    rs("Message Sent") = False
    
    If gPGPVersion = PGP5x Then
        rs("Message Status") = spgpAnalyseMessage(MessageArea.Text)
    Else
        rs("Message Status") = PGPAnalyze_Unknown
    End If
    rs.Update
    rs.Close
   ' Unload Me
    Exit Sub

WriteMessError:
    ShowStatus 1, "Following error occurred in Save Message: " & Err.Description
    Err.Clear
End Sub



Private Sub EditPerform(EditFunction As Integer)
If TypeOf Me.ActiveControl Is TextBox Then
    Call SendMessage(Me.ActiveControl.hWnd, EditFunction, 0, 0&)
ElseIf TypeOf Me.ActiveControl Is RichTextBox Then
    If m_ControlKey = False Then
        Call SendMessage(Me.ActiveControl.hWnd, EditFunction, 0, 0&)
    End If
Else
    Beep
End If
End Sub

Public Sub UseCypherPunk()
   gRemailerType = REMAILER_CYPHERPUNK
   'gEncryptToRemailer = True
    'SSRibbon1(7).Enabled = True
    'SetFocus
    SSRibbon1(6).Enabled = True
    SSRibbon1(0).Picture = SSRibbon1(21).Picture
    'cmbRemailerSelect.ListIndex = 1 'need remailer
    'frmRemailerList.Show
   Exit Sub
   
    
End Sub

Private Sub SetUseOfMixmaster()
Dim SectionName As String
    gRemailerType = REMAILER_MIX
    SSRibbon1(6).Enabled = False
    SSRibbon1(0).Picture = SSRibbon1(21).Picture
End Sub

Public Sub DontUseRemailer()
gRemailerType = STANDARD_EMAIL
End Sub

Private Sub AddAttachment()
Dim lListItem As ListItem
Dim sFileName As String
Dim J As Integer

   On Error GoTo AttachError

        '---------------------------------------------
        'File Open/Save Dialog Box Flags
        
        'Do this for PGP only - PGP2.6.3 can't handle long file names
        'CommonDialog1.FileTitle = cdlOFNNoLongNames
    
        CommonDialog1.DialogTitle = "Open file to attach"
        CommonDialog1.Flags = &H2& Or &H4& Or &H40000 'cdlOFNNoLongNames
        CommonDialog1.Filter = "All Files (*.*)|*.*"
        CommonDialog1.FilterIndex = 1
        CommonDialog1.CancelError = True
        CommonDialog1.InitDir = App.Path
        CommonDialog1.Action = 1
        
        sFileName = StripFileName(CommonDialog1.FileName)
        FileCopy CommonDialog1.FileName, TempPathLocation & sFileName
        J = GetFileIconNum(sFileName)
        Set lListItem = lvwAttachments.ListItems.Add(, , sFileName, J, J)
        ShowAttachmentContainer
        
        ChDrive Mid$(App.Path, 1, 3)
        ChDir App.Path
       
    Exit Sub

AttachError:
    ShowStatus 1, "Error processing attachment - The following error occurred: " & Err.Description
    Beep
    Err.Clear
End Sub

Private Sub EncryptMessageArea(KeyID As String)
Dim PGPCmdString As String
Dim TheFileName As String
Dim tmpstr As String
Dim FileNum As Long
    
    On Error GoTo PGPEncryptError
    vb2spgpContext.Initialise
    gCancelAction = False
   ' PGPCmdString = ""
    'gPGPKeyID = ""
    If PGPConvent.Checked = False Then
        If KeyID = "" Then
            MsgBox "Need to specify a recipient.  Click on the 'To:' button to select a receipient.  If you wish to encrypt then the recipient's public key must be on your keyring.", vbApplicationModal + vbCritical, "No Recipient"
            Exit Sub
        End If
        If Not KeyOnKeyRing(KeyID) Then
            Beep
            MsgBox "Recipient is not on your keyring.  To fix this obtain the person's public key and then add to your keyring using the PI Key->Add Key ring function.", vbApplicationModal + vbCritical, "No Recipient"
            Exit Sub
        End If
       ' gKeyID = txtTo.Text
    End If
    
    '---------------------------------------------
    'handle case of saving a message to file, then encrypting
    '---------------------------------------------
    If PGPFile.Checked Then
        If Not Len(MessageArea.Text) = 0 Then
         tmpstr = MsgBox("The message area has text in it - would you like to encrypt the message area and save it in a file?  If you select 'No' you will then be given the opportunity to specify the file you would like to encrypt.", vbYesNo + vbQuestion + vbApplicationModal, "File Encrypt")
         If tmpstr = vbYes Then
            CommonDialog1.DialogTitle = "Specify the name of the file you wish use."
            CommonDialog1.Flags = &H2& + &H4&
            CommonDialog1.Filter = "PGP .asc Files (*.asc)|*.asc"
            CommonDialog1.FilterIndex = 1
            CommonDialog1.CancelError = True
            CommonDialog1.InitDir = App.Path
            CommonDialog1.Action = 1
            TheFileName = CommonDialog1.FileName
            ChDrive Mid$(App.Path, 1, 3)
            ChDir App.Path
            FileNum = FreeFile
            Open CommonDialog1.FileName For Output As FileNum
                Print #FileNum, MessageArea.Text
            Close #FileNum
         End If
        End If
        '---------------------------------------------
        'no, it is a file encrypt
        '---------------------------------------------
        If Len(TheFileName) = 0 Then
            '---------------------------------------------
            'user has not already specified a file name from above
            '---------------------------------------------
            CommonDialog1.DialogTitle = "Specify the file you wish to encrypt"
            CommonDialog1.Flags = &H2& + &H4&
            CommonDialog1.Filter = "All Files (*.*)|*.*"
            CommonDialog1.FilterIndex = 1
            CommonDialog1.CancelError = True
            CommonDialog1.InitDir = App.Path
            CommonDialog1.Action = 1
            TheFileName = CommonDialog1.FileName
            ChDrive Mid$(App.Path, 1, 3)
            ChDir App.Path
    
        End If
        'If PGPEyes.Checked = True Then
         '   PGPCmdString = "m"
        'End If
        MousePointer = vbHourglass
        ShowStatus 1, "Encrypting the file " & TheFileName & "..."
        'vb2spgpContext.TextMode = 0
        '[AC 2/3/2002]
        vb2spgpContext.TextMode = 1
        vb2spgpContext.Armor = 1
        vb2spgpContext.FileIn = TheFileName
        vb2spgpContext.FileOut = StripExt(TheFileName) & ".asc"
        If PGPConvent.Checked Then
            vb2spgpContext.ConventionalEncrypt = 1
            spgpEncryptfile '"", "-cat" + PGPCmdString, TheFileName
        Else
            vb2spgpContext.ConventionalEncrypt = 0
            spgpEncryptfile 'txtTo.Text, "-eat" + PGPCmdString, TheFileName
        End If
        MousePointer = vbDefault
        ShowStatus 1, "Encrypted file saved as " & vb2spgpContext.FileOut & ".  Orignal file unchanged."
    Else
    '---------------------------------------------
    'is this a eyes, conventional, other encrypt?
    '---------------------------------------------
        If PGPEyes.Checked = True Then
            PGPCmdString = "m"
        End If
        'vb2spgpContext.TextMode = 0
        '[AC 2/3/2002]
        vb2spgpContext.TextMode = 1
        vb2spgpContext.Armor = 1
        MousePointer = vbHourglass
        If PGPConvent.Checked = True Then
            vb2spgpContext.ConventionalEncrypt = 1
            vb2spgpContext.KeyEncrypt = 0
            EncryptMessage "", "-catw" + PGPCmdString
        Else
            vb2spgpContext.ConventionalEncrypt = 0
            vb2spgpContext.KeyEncrypt = 1
            EncryptMessage KeyID, "-eatw" + PGPCmdString
        End If
        MousePointer = vbDefault
    End If
    
    Exit Sub
PGPEncryptError:
     MousePointer = vbDefault
     MsgBox Err.Description & " in PGP Encryption", vbCritical + vbApplicationModal, App.Title
     Err.Clear
     
End Sub

Public Sub CreateReplyBlock()
    MessageArea = ""
    'Okay set up the To address for Appendnyminfo
    If Nym.UseNewsGroupReply Then
        txtTo.Text = Nym.NewsGroupReplyEmail
    Else
        txtTo.Text = Nym.EmailAddress
    End If
    AppendNymInfo
    If gCancelAction Then Exit Sub
    MessageArea.SelStart = 0
    MessageArea.SelText = "Reply-Block:" & vbCrLf
    MessageArea.SelText = "::" & vbCrLf
    If Nym.UseNewsGroupReply Then
        MessageArea.SelText = "Anon-To: " & Nym.NewsGroupReplyEmail & vbCrLf
    Else
        MessageArea.SelText = "Anon-To: " & txtTo.Text & vbCrLf 'Nym.EmailAddress & vbCrLf
    End If
    MessageArea.SelText = "Latent-Time: " & Nym.LatentTime & vbCrLf
           
    If Not Nym.PassPhrase(1) = "" Then
        MessageArea.SelText = "Encrypt-Key: " & Nym.PassPhrase(1) & vbCrLf & vbCrLf
    Else
        MessageArea.SelText = vbCrLf
    End If
    If Nym.UseNewsGroupReply Then
        MessageArea.SelText = "##" & vbCrLf
        MessageArea.SelText = "Newsgroups: " & Nym.NewsGroupReplyGroup & vbCrLf
        MessageArea.SelText = "Subject: " & Nym.NewsGroupReplySubject & vbCrLf & vbCrLf
    End If
'AppendNymInfo
End Sub



Public Sub DisablePGPMenuItems()
'Exit Sub
mPGP.Visible = False
DontUseRemailer
cmbRemailerSelect.Visible = False
'cmbRemailerSelect.ListIndex = 0
'mNewsgroups.Visible = False
USENETGate.Enabled = False
SSRibbon1(9).Enabled = False
cmbRemailerSelect.Clear
cmbRemailerSelect.AddItem "Standard Email", 0
cmbRemailerSelect.ListIndex = 0
cmbRemailerSelect.Visible = True
mNym.Visible = False
End Sub

Public Sub EnablePGPMenuItems()
'Exit Sub
mPGP.Visible = True
'mRemailers.Visible = True
'mNewsgroups.Visible = True
USENETGate.Enabled = True
cmbRemailerSelect.Visible = True
SSRibbon1(9).Enabled = True
mNym.Visible = True
cmbRemailerSelect.Clear
    cmbRemailerSelect.AddItem "Standard Email", 0
    cmbRemailerSelect.AddItem "Send via remailer Type 1(Cypherpunk)", 1
    cmbRemailerSelect.AddItem "Send via remailer Type 2(Mixmaster)", 2
    cmbRemailerSelect.AddItem "Encrypt before Sending Message", 3
    cmbRemailerSelect.AddItem "Encrypt and Sign before Sending Message", 4
    cmbRemailerSelect.AddItem "Just Sign before Sending Messages", 5
    cmbRemailerSelect.AddItem "Send Messages using Nym", 6
    cmbRemailerSelect.Visible = True
     cmbRemailerSelect.ListIndex = 0
End Sub



Public Sub GetRecipient()
Dim foostr As String
    Dim x As String
    Dim pos1 As Integer
    Dim pos2 As Integer
    Dim i As Integer
        
    '---------------------------------------------
    'clear any current name from the "to:" box
    '---------------------------------------------
   ' txtTo.Text = ""
   ' gKeyID = ""
    '---------------------------------------------
    'present the list of public keys
    '---------------------------------------------
    frmSelectUserID.Label2 = ""
    frmSelectUserID.Label1 = "Select a recipient from either your public key ring, your personal address book or from an existing mail group."
    frmSelectUserID.Show vbModal
End Sub

Public Sub ShowStatus(Panel As Integer, Status As String)
StatusBar.Style = sbrNormal
If TextWidth(Status & "  ") > StatusBar.Panels(Panel).Width Then
    StatusBar.Panels(Panel).Width = TextWidth(Status & "    ")
End If
StatusBar.Panels.Item(Panel) = Status
End Sub

Private Function CheckConnection()
    Dim ReturnValue As Boolean
    Dim WaitAWhile As Variant
    Dim Req1 As Long
    Dim Response As String

    On Error GoTo BadConnection
    If IPInfo1.WinsockLoaded Then
        CheckConnection = True
    Else
        IPInfo1.WinsockLoaded = True
        CheckConnection = True
    End If
    DoEvents
    Exit Function
    
BadConnection:
    CheckConnection = False
    Beep
    MsgBox "Can't create connection", vbApplicationModal + vbCritical
    Err.Clear
End Function

Public Sub ShowAttachmentOptions(Attachment As String)
'StatusBar.Style = sbrNormal
'If TextWidth(Attachment & "  ") > StatusBar.Panels(2).Width _
'    Then StatusBar.Panels(2).Width = TextWidth(Attachment & " ")
'StatusBar.Panels.Item(2) = Attachment
End Sub

Public Sub ShowRemailer(Remailer As String)
StatusBar.Style = sbrNormal
If TextWidth(Remailer & "  ") > StatusBar.Panels(3).Width _
    Then StatusBar.Panels(3).Width = TextWidth(Remailer & " ")
StatusBar.Panels.Item(3) = Remailer
End Sub

Private Sub SetAttachmentEncryptionOptions()
   ' mDontEncryptAttachment.Checked = False
    mEncryptAttachmentWithKey.Checked = True
    mConventionallyEncryptAttachment.Checked = False
End Sub

Private Sub ShowAttachmentContainer()
If lvwAttachments.ListItems.Count = 0 Then
    lvwAttachments.Visible = False
Else
    lvwAttachments.Visible = True
End If
Form_Resize
DoEvents
End Sub

Public Sub InsertAttachmentsIntoMessage()
Dim i As Integer
On Error GoTo BadInsert
'Exit Sub
    MIME1.Action = a_ResetData
    MIME1.PartCount = lvwAttachments.ListItems.Count + 1
    MIME1.PartDecodedString(0) = MessageArea.Text
    MIME1.PartEncoding(0) = pe_7Bit '1 'pe_8Bit
    MIME1.PartContentType(0) = "text/plain"
        For i = 1 To MIME1.PartCount - 1
            MIME1.PartDecodedFile(i) = TempPathLocation & lvwAttachments.ListItems.Item(i)
        Next i
    MIME1.Message = MessageArea.Text  'GetTemporaryFile
    MIME1.Action = 4
    MessageArea.Text = ""
    MessageArea.SelStart = 0
    'MessageArea.Text = MIME1.MessageHeaders & vbCrLf & GetFileText(MIME1.Message)   'MIME1.PartDecodedString(0)
    If MIME1.PartCount > 1 Then
       ' Clipboard.Clear
       ' Clipboard.SetText (MIME1.MessageHeaders)
        MessageArea.Text = MIME1.MessageHeaders & vbCrLf & vbCrLf & MIME1.Message
    Else
        MessageArea.Text = MIME1.Message
    End If
    '
    'Now remove the attachments
    '
    lvwAttachments.ListItems.Clear
    lvwAttachments.Visible = False
    Form_Resize
    DoEvents
    Exit Sub
BadInsert:
    MsgBox "Error inserting the messages into the message area.  Error returned was: " & Err.Description, vbApplicationModal + vbCritical, "Insertion Error"
    Err.Clear
End Sub
Public Sub DecodeMessage()
Dim i As Integer
Dim J As Long
Dim bRes As Boolean
Dim lListItem As ListItem

    
    On Error GoTo BadDecode
    If InStr(MessageArea.Text, "boundary=") = 0 Then Exit Sub
    MIME1.Action = 5 'a_ResetData
    MIME1.Message = MessageArea.Text
    MIME1.Action = 2
    'MIME1.ContentType = MIME1.ContentType
    MessageArea.Text = MIME1.PartDecodedString(0)
    
    For i = 1 To MIME1.PartCount - 1
        If MIME1.PartName(i) = "" Then
            MIME1.PartName(i) = "Unknown.htm"
        End If
        'Note TempPathLocation is used for all attachments - rather than putting it into the path lis
        'in the attachment list.  AC 23/2/2002
        AttachmentFileName = MIME1.PartName(i)  'TempPathLocation & MIME1.PartName(i)
        bRes = PutFileText(App.Path & "\temp\" & AttachmentFileName, MIME1.PartDecodedString(i))
        J = GetFileIconNum(AttachmentFileName) 'MIME1.PartName(i))
        Set lListItem = lvwAttachments.ListItems.Add(, , AttachmentFileName, J, J)
        On Error Resume Next
        'Do this to render the message in frmMain.DisplayMessage
        AttachmentFileName = App.Path & "\temp\" & AttachmentFileName
    Next
    lvwAttachments.Visible = True
    Form_Resize
    
    Exit Sub
BadDecode:
    MsgBox "Error decoding the message. Error returned was: " & Err.Description, vbCritical + vbApplicationModal, "Decoding Error"
    Err.Clear
End Sub

Public Sub PGPAddKeyFromString(sKeyString As String)
Dim ClipText As String
Dim cmd As String
Dim iResult As Integer
Dim Ecount As Integer
Dim FileNum As Integer
Dim foo As String

Dim i As Long
Dim BufferIn As String
Dim spgperr As String * 256
Dim KeyProps As String
On Error GoTo AddKeyError


    BufferIn = sKeyString & Chr(0)
    KeyProps = String(4096, Chr(0))
    
    ' keyprops takes either key id(s) or user id(s)
    ' and returns the key's properties
  '  MsgBox ("importing keys" & BufferIn)
    i = spgpKeyImport(BufferIn, KeyProps, Len(KeyProps), 1, 0)
     '  MsgBox ("Beep" & KeyProps)
    If Not i = 0 Then
        Beep
       '     MsgBox ("Beep")
        Call spgpGetErrorString(i, spgperr)
        ShowStatus 1, "Error occurred importing keys.  Error code: " & spgperr
       'Err.Raise 1000, "spgpKeyImport", spgperr
    Else
        ShowStatus 1, "Keys imported successfully..."
    End If
    

    ' parse the returned property-string into a TKey_Data record
    'NOTE CALL this only if I call spgpKey_Import function.  the above does NOT return all the data
    ' used by parseKeyData   Key = ParseKeyData(KeyProps)
   

Exit Sub
AddKeyError:
    MsgBox Err.Description & " (in PGPAddKey)", vbCritical + vbApplicationModal, App.Title
    Err.Clear
End Sub

Private Sub InitialiseProgressBar()
ProgressBar1.Top = StatusBar.Top + 10
ProgressBar1.Left = StatusBar.Panels.Item(3).Left + 10
ProgressBar1.Height = 0.9 * StatusBar.Height
ProgressBar1.Width = StatusBar.Panels.Item(3).Width - 10
ProgressBar1.Visible = True
ProgressBar1.Min = 0
ProgressBar1.Max = 100
ProgressBar1.Value = 0
End Sub


'Functions and Subs
Public Function GetFileIconNum(psPathName As String) As Integer
    Dim i As Integer
    Dim lsExt As String
    Dim lsKey As String
    Dim shFileInfoStruct As SHFILEINFO
    
    If Trim(psPathName & "") <> "" Then
        lsExt = "." & GetExt(psPathName)
        lsKey = "a" & lsExt
        For i = 1 To imglstLarge.ListImages.Count
            If lsKey = imglstLarge.ListImages.Item(i).Key Then
                GetFileIconNum = i
                Exit Function
            End If
        Next
    Else
        GetFileIconNum = 1
        Exit Function
    End If
    
    If GetFileIconNum = 0 And lsExt <> "" Then
        'add an icon for it
        Call SHGetFileInfo(lsExt, 0, shFileInfoStruct, LenB(shFileInfoStruct), SHGFI_SYSICONINDEX Or SHGFI_ICON Or SHGFI_USEFILEATTRIBUTES)
        If shFileInfoStruct.hIcon > 0 Then
            mLastImageNum = mLastImageNum + 1
            imglstLarge.ListImages.Add , lsKey, picGenericIcon.Image
            i = ImageList_ReplaceIcon(imglstLarge.hImageList, mLastImageNum - 1, shFileInfoStruct.hIcon)
            Call SHGetFileInfo(lsExt, 0, shFileInfoStruct, LenB(shFileInfoStruct), SHGFI_ICON Or SHGFI_SMALLICON Or SHGFI_USEFILEATTRIBUTES)
            imglstSmall.ListImages.Add , lsKey, picGenericIcon.Image
            Call ImageList_ReplaceIcon(imglstSmall.hImageList, mLastImageNum - 1, shFileInfoStruct.hIcon)
            GetFileIconNum = mLastImageNum
        End If
    Else
        GetFileIconNum = 1
    End If
End Function


Private Function GetExt(psFileName As String) As String
    Dim llPos As Long
    Dim lbFound As Boolean
    
    If Len(psFileName) <= 1 Then
        GetExt = ""
        Exit Function
    End If
    
    lbFound = False
    For llPos = Len(psFileName) - 1 To 1 Step -1
        If Mid$(psFileName, llPos, 1) = "." Then
            lbFound = True
            GetExt = Mid$(psFileName, llPos + 1)
            Exit For
        End If
    Next llPos
    
    If Not lbFound Then
        GetExt = ""
    End If
End Function


Public Sub EnableMessageFields()
On Error Resume Next
SSRibbon1(0).Enabled = True ' don't allow them to send
SSRibbon1(6).Enabled = True ' don't allow them to send
SSRibbon1(9).Enabled = True ' don't allow them to send
SSRibbon1(4).Enabled = True ' don't allow them to send
SSRibbon1(7).Enabled = True ' don't allow them to send
         
lblFrom(0).Visible = False
lblFrom(1).Visible = False
lblcc.Visible = False
lblTo.Visible = False
txtToAddresses.Visible = False
txtCCAddresses.Visible = False
lblSubject.Visible = False

txtCC.Visible = True
txtTo.Visible = True
btnTo.Visible = True
btnCC.Visible = True
txtsubject.Visible = True


End Sub

Private Sub txtCC_Change()
If gComposeMode Then
    Timer1.Enabled = False
    m_fCCFieldChanged = True
    Timer1.Enabled = True
Else
    Timer1.Enabled = False
    m_fCCFieldChanged = False
End If
End Sub

Private Sub txtCC_DblClick()
SelectEntryAtCursor txtCC
DoEvents
frmShowAddressBookProperties.txtEmailAddress = AddressList.GetContactEmailAddress(txtCC.SelText, CONTACT_CC_LIST)
frmShowAddressBookProperties.txtDisplayName = AddressList.GetContactDisplayName(txtCC.SelText, CONTACT_CC_LIST)
frmShowAddressBookProperties.Show vbModal
End Sub

Private Sub txtCC_KeyDown(KeyCode As Integer, Shift As Integer)
If txtCC.SelLength = 0 Then Exit Sub
If KeyCode = vbKeyDelete Then
    AddressList.RemoveContact txtCC.SelText, CONTACT_CC_LIST 'Remove Type 1
    FillCCBoxWithList
     KeyCode = 0
End If
End Sub

Private Sub txtCC_KeyPress(KeyAscii As Integer)
Dim toAddressee As String
gComposeMode = True
Select Case KeyAscii
    Case Asc(",")
        KeyAscii = Asc(";")
    Case 13
        KeyAscii = 0
End Select

txtCC.SelUnderline = False

End Sub

Private Sub txtCC_LostFocus()
If Not gComposeMode Then Exit Sub
If m_fCCFieldChanged Then ScanTextForContacts txtCC, CONTACT_CC_LIST
m_fCCFieldChanged = False

End Sub

Private Sub txtCC_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
SelectEntryAtCursor txtCC
End Sub

Private Sub txtCCAddresses_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub txtTo_Change()
'reset it to allow user to type info in
If gComposeMode Then
    Timer1.Enabled = False
    m_fToFieldChanged = True
    Timer1.Enabled = True
Else
    Timer1.Enabled = False
    m_fToFieldChanged = False
End If

End Sub

Private Sub txtTo_DblClick()
SelectEntryAtCursor txtTo
DoEvents
frmShowAddressBookProperties.txtEmailAddress = AddressList.GetContactEmailAddress(txtTo.SelText, CONTACT_TO_LIST)
frmShowAddressBookProperties.txtDisplayName = AddressList.GetContactDisplayName(txtTo.SelText, CONTACT_TO_LIST)
frmShowAddressBookProperties.Show vbModal

End Sub
Private Sub txtTo_KeyDown(KeyCode As Integer, Shift As Integer)
If txtTo.SelLength = 0 Then Exit Sub
If KeyCode = vbKeyDelete Then
    AddressList.RemoveContact txtTo.SelText, CONTACT_TO_LIST 'Remove Type 1
    FillToBoxWithList
    KeyCode = 0 'This will stop the beep
End If

End Sub

Private Sub txtTo_KeyPress(KeyAscii As Integer)
Dim toAddressee As String
gComposeMode = True
Select Case KeyAscii
    Case Asc(",")
        KeyAscii = Asc(";")
    Case 13
        KeyAscii = 0
End Select

txtTo.SelUnderline = False


End Sub

Private Sub txtTo_LostFocus()
If Not gComposeMode Then Exit Sub
If m_fToFieldChanged Then ScanTextForContacts txtTo, CONTACT_TO_LIST
m_fToFieldChanged = False

End Sub

Private Sub txtTo_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
SelectEntryAtCursor txtTo
End Sub



Private Sub txtToAddresses_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Public Sub FillToBoxWithList()
Dim VarArray As Variant
Dim i As Long

txtTo.SelStart = 0
txtTo.Text = ""
VarArray = AddressList.GetAllContacts(CONTACT_TO_LIST)
For i = 1 To VarArray(0, CONTACT_TO_LIST, 0)
    txtTo.SelUnderline = False
    If i > 1 Then
            txtTo.SelText = "; "
            txtTo.SelUnderline = True
            txtTo.SelText = VarArray(1, CONTACT_TO_LIST, i)
    Else
            txtTo.SelUnderline = True
            txtTo.SelText = VarArray(1, CONTACT_TO_LIST, i)
            
    End If
Next
txtTo.SelUnderline = False
End Sub
Public Sub FillCCBoxWithList()
Dim VarArray As Variant
Dim i As Long

txtTo.SelStart = 0
txtCC.Text = ""
VarArray = AddressList.GetAllContacts(CONTACT_CC_LIST)
For i = 1 To VarArray(0, CONTACT_CC_LIST, 0)
    txtTo.SelUnderline = False
    If i > 1 Then
            txtCC.SelText = "; "
            txtCC.SelUnderline = True
            txtCC.SelText = VarArray(1, CONTACT_CC_LIST, i)
    Else
            txtCC.SelUnderline = True
            txtCC.SelText = VarArray(1, CONTACT_CC_LIST, i)
            
    End If
Next
txtCC.SelUnderline = False
End Sub
'
'Return "" if not found in database
'
Public Function LookUpContactRecord(ContactID As String, iAddresseeList As Integer) As String

Dim rs As Recordset
Dim ContactEmailAddress As String
Dim ContactFullName As String
Dim VarArray As Variant
Dim lIndex As Long

On Error GoTo BadUpdate
' Strip out <, quotes etc
LookUpContactRecord = ""
ContactEmailAddress = StripEMailAddress(ContactID)
ContactFullName = StripFullName(ContactID)
If ContactEmailAddress = "" Then
    LookUpContactRecord = ""
    Exit Function
End If
Set rs = DB.OpenRecordset("Contacts", dbOpenDynaset)
rs.FindFirst "[Contact Email] =" & "'" & ContactEmailAddress & "'"

If rs.NoMatch Then
    'If there is no match then okay, to add to address list
    VarArray = Array(ContactID, ContactFullName, ContactEmailAddress, CONTACT_NEW)
    lIndex = AddressList.AddContact(VarArray, iAddresseeList)
    LookUpContactRecord = ""
Else
    'rs("Contact Email") = ContactEmailAddress
    LookUpContactRecord = IIf(IsNull(rs("Contact Name")), rs("Contact Email"), rs("Contact Name"))
    'The first argument will be the one displayed in the To box.
    VarArray = Array(LookUpContactRecord, rs("Contact Name"), rs("Contact Email"), CONTACT_IN_DB)
    lIndex = AddressList.AddContact(VarArray, iAddresseeList)
    rs.Close
End If
Exit Function
BadUpdate:
    'rs.Close
    MsgBox "Could not obtain your contacts from the database. Error was: " & Err.Description, vbCritical + vbApplicationModal, "Contact Update Error"
    Err.Clear
End Function
'
'Return "" if not found in database
'
Public Function LookUpPGPKeyring(ContactID As String, iAddresseeList As Integer) As String

Dim ContactEmailAddress As String
Dim ContactFullName As String
Dim VarArray As Variant
Dim lIndex As Long

On Error Resume Next
If gPGPVersion = NoPGP Then Exit Function

' Strip out <, quotes etc
LookUpPGPKeyring = ""
ContactEmailAddress = StripEMailAddress(ContactID)
ContactFullName = StripFullName(ContactID)
If ContactEmailAddress = "" Then
    LookUpPGPKeyring = ""
    Exit Function
End If
'ContactEmailAddress = "<" & ContactEmailAddress & ">"
If spgpKeyIsOnRing(ContactEmailAddress) = 0 Then
    'If there is no match then okay, to add to address list
   ' VarArray = Array(ContactID, ContactFullName, ContactEmailAddress, CONTACT_NEW)
   ' lIndex = AddressList.AddContact(VarArray, iAddresseeList)
   ' LookUpPGPKeyring = IIf(IsNull(ContactFullName), ContactEmailAddress, ContactFullName)
'Else
     'The first argument will be the one displayed in the To box.
    VarArray = Array(ContactFullName & " " & ContactEmailAddress, ContactFullName, ContactEmailAddress, CONTACT_ON_PGPKEYRING)
    lIndex = AddressList.AddContact(VarArray, CONTACT_TO_LIST)
    LookUpPGPKeyring = IIf(IsNull(ContactFullName), ContactEmailAddress, ContactFullName)
End If

End Function
Public Sub ScanTextForContacts(rtbAddressList As RichTextBox, iContactType As Integer)
Dim bEOL As Boolean
Dim lSelstart As Long
Dim lSelEnd As Long
Dim sAddressee As String
Dim bDone As Boolean

'First clear the underline and scan from scratch
rtbAddressList.SelStart = 0
lSelstart = rtbAddressList.SelStart

'First check if we are over an entry or between or at the end
'First look back to see if we are the first entry ie eol
rtbAddressList.Span vbCrLf, True, True
If rtbAddressList.SelLength = 0 Then
    Exit Sub
Else
    '--------------------------------------------------------------------------
    'First let's skip over any entries that are in the contact list already
    'These are entries with an underline
    '---------------------------------------------------------------------------
    rtbAddressList.SelStart = 0
    rtbAddressList.SelUnderline = False
    lSelstart = 0
  Do 'Loop until line is parsed
    rtbAddressList.SelStart = lSelstart
    rtbAddressList.SelUnderline = False
    lSelEnd = lSelstart
    Do
        lSelEnd = lSelstart + rtbAddressList.SelLength 'Note this line must be first
        rtbAddressList.SelLength = rtbAddressList.SelLength + 1
        
    Loop While rtbAddressList.SelUnderline = True And lSelEnd < Len(rtbAddressList.Text)
        lSelstart = lSelEnd
        Select Case lSelEnd
                    
            Case Is >= Len(rtbAddressList.Text)
                bDone = True
            Case Else
            'Okay we have found a delimter of sorts - so move forward until next alpha char
            rtbAddressList.SelStart = lSelstart
            rtbAddressList.Span " ,;" & vbCrLf, True, False
            lSelstart = rtbAddressList.SelStart + rtbAddressList.SelLength
            rtbAddressList.SelStart = lSelstart
            rtbAddressList.SelLength = 1
            'Now see if we there is an underline char - if not process as email address
            If Not rtbAddressList.SelUnderline = True Then
                'First look for end of address
                rtbAddressList.Span ",;" & vbCrLf, True, True
                If rtbAddressList.SelLength = 0 Then
                    bDone = True
                Else
                    '.sellength is now pointing to , ; or vbclrlf
                    lSelEnd = rtbAddressList.SelLength
                    'Check to see it the field has an '@' character ie is it an emal
                    If InStr(1, rtbAddressList.SelText, "@") = 0 Then 'rtbAddressList.Span "@", True, True
                        'Possible group or list, check it
                        If Not LookUpGroupName(rtbAddressList.SelText) Then
                            MsgBox "Mail list '" & rtbAddressList.SelText & "' was not found.", vbApplicationModal + vbCritical, "Scan for Addresses"
                        Else
                            'List group details....
                            Dim sMember As String
                            Dim DisplayName As String
                            Dim MemberArray As Variant
                            Dim VarArray As Variant
                            'Dim ContactFull
                            Dim i As Integer
                            Dim lIndex As Long
                            'MemberArray = GetMembersOfGroup(rtbAddressList.SelText)
                            'For i = 1 To CInt(MemberArray(0))
                               ' sMember = MemberArray(i)
                                DisplayName = rtbAddressList.SelText
                                'ContactEmailAddress = StripEMailAddress(sMember)
                                'ContactFullName = StripFullName(sMember)
                                'The first item is the name of the group
                                'Note last entry past data holds the type of contact, eg PGP etc
                                VarArray = Array(DisplayName, "", "", CONTACT_IN_MAILGROUP)
                                lIndex = AddressList.AddContact(VarArray, iContactType)
                            'Next
                            rtbAddressList.SelUnderline = True
                        End If
                        lSelstart = rtbAddressList.SelStart + rtbAddressList.SelLength 'lSelEnd
                    Else
                        rtbAddressList.SelLength = lSelEnd
                        sAddressee = AddressList.GetContactDisplayName(rtbAddressList.SelText, iContactType)
                        'If empty  string then it is not in the database but it has been added to the Contact List
                        If sAddressee = "" Then
                            sAddressee = LookUpPGPKeyring(rtbAddressList.SelText, iContactType)
                            If sAddressee = "" Then
                                sAddressee = LookUpContactRecord(rtbAddressList.SelText, iContactType)
                            End If
      
                            rtbAddressList.SelUnderline = True
                        Else
                            rtbAddressList.SelUnderline = True
                        End If
                        lSelstart = rtbAddressList.SelStart + rtbAddressList.SelLength
                    End If
                End If
                
            End If
            
        End Select
Loop Until bDone = True
    
     
End If

End Sub
Public Sub SelectEntryAtCursor(rtbAddressBox As RichTextBox)
Dim lSelstart As Long
Dim lSelEnd As Long
Dim lrtbText As Long

'First look back to see if we are at the begining or end of the line
If rtbAddressBox.SelStart >= Len(rtbAddressBox.Text) Then Exit Sub
If rtbAddressBox.SelStart = 0 Then Exit Sub
rtbAddressBox.Span ";", False, True

'Check if we are the beginning of the line
If rtbAddressBox.SelStart = 0 Then
    rtbAddressBox.Span ";" & vbCrLf, True, True
Else
    'We have found a ";" at txt.selstart
    ' We need to do this as the we need to bump forward past the comma and spaces
    rtbAddressBox.Span "; ", True, False
    rtbAddressBox.Span ";" & vbCrLf, True, True
    'rtbAddressBox.SelText = rtbAddressBox.SelText
End If

End Sub

Public Sub PGPEncryptMessage(SignMessage As Boolean, ClearSign As Boolean)
Dim TheFileName As String
    Dim bRes As Boolean
    Dim foo As String
    Dim FileNum As Integer
    Dim VarArray As Variant
    Dim msg As String
    Dim i As Integer
    Dim bSignatureToFile As Boolean
    
    
    On Error GoTo FileEnSError
    gCancelAction = False
    
    'First set the sign parameters as this is common for both routines
     vb2spgpContext.Initialise
     If PGPConvent.Checked = True Then
        vb2spgpContext.ConventionalEncrypt = 1
        vb2spgpContext.KeyEncrypt = 0
    Else
        vb2spgpContext.ConventionalEncrypt = 0
        vb2spgpContext.KeyEncrypt = 1
   End If
     
     'vb2spgpContext.KeyEncrypt = 1
     If ClearSign Then
        vb2spgpContext.Clear = 1
        vb2spgpContext.Sign = 1
        vb2spgpContext.KeyEncrypt = 0
     End If
     If SignMessage Then
        vb2spgpContext.Sign = 1
     End If
        '---------------------------------------------
        'Find the  Key to use
        'if key is selected, it is in global gPGPKeyID
        '---------------------------------------------
    
    If vb2spgpContext.KeyEncrypt = 1 Then
        VarArray = AddressList.GetAllPGPContacts(CONTACT_ALL_LIST)
        If VarArray(0, CONTACT_ALL_LIST, 0) >= 1 Then
            For i = 1 To VarArray(0, CONTACT_ALL_LIST, 0)
                If i = 1 Then
                    gPGPKeyID = VarArray(3, CONTACT_ALL_LIST, i)
                Else
                    gPGPKeyID = gPGPKeyID & "," & VarArray(3, CONTACT_ALL_LIST, i)
                End If
            Next
        Else
            vb2spgpContext.SelectPrivateKeys = False
            frmViewKeyRing.lblContext = "You need to select a public key to encrypt the message with.  Select from this list or private keys which key you would like to sign the message with"
            frmViewKeyRing.Caption = "Select a Public Key encrypt the message with"
            frmViewKeyRing.Show vbModal
        
            If gPGPKeyID = "" Then
                Beep
                Exit Sub
            End If
        End If
        vb2spgpContext.CryptKeyID = gPGPKeyID
    End If
    '----------------------------------
    ' Now check if we need to sign or ask for a key
    '-----------------------------------
     If SignMessage Or ClearSign Then
        gPGPKeyID = ReadProfile("PGP Options", "Default Key ID")
        If gPGPKeyID = "" Then
            vb2spgpContext.SelectPrivateKeys = True
            frmViewKeyRing.lblContext = "You need to sign this message.  Please select a key from private key ring."
            frmViewKeyRing.Caption = "Select Key to sign the message"
            frmViewKeyRing.Show vbModal
        End If
            
        If gPGPKeyID = "" Then
            Beep
            Exit Sub
        Else
            vb2spgpContext.SignKeyID = gPGPKeyID
            'vb2spgpContext.SignKeyID = ""
        End If
    
    End If
    
    '-----------------------------
    'We have a key to encrypt with now
    '-----------------------------
   
    '------------------------------------------------------------
    'First check if this is a file command
    '------------------------------------------------------------
    If PGPFile.Checked Then
        bSignatureToFile = False
        If Not MessageArea.Text = "" Then
            If ClearSign Then
                frmFileEncryptionOption.optFile(0).Caption = "ClearSign the message area and save to a file."
                frmFileEncryptionOption.optFile(1).Caption = "Clear Sign an existing file."
            Else
                frmFileEncryptionOption.optFile(0).Caption = "Encrypt/Sign the message area and save to a file."
                frmFileEncryptionOption.optFile(1).Caption = "Encrypt/Sign an existing file."
            End If
            frmFileEncryptionOption.Show vbModal
            bSignatureToFile = frmFileEncryptionOption.EncryptToFile
            Set frmFileEncryptionOption = Nothing
        End If
        If bSignatureToFile Then
            
    '---------------------------------------------
    'handle case of saving a message to file, then encrypting
    '---------------------------------------------
            CommonDialog1.DialogTitle = "Specify the file to save to."
            CommonDialog1.Flags = &H2& + &H4&
            CommonDialog1.Filter = "PGP .asc Files (*.asc)|*.asc"
            CommonDialog1.FilterIndex = 1
            CommonDialog1.CancelError = True
            CommonDialog1.InitDir = App.Path
           ' CommonDialog1.Action = 1
            CommonDialog1.ShowSave
            TheFileName = CommonDialog1.FileName
            If InStr(1, TheFileName, ".asc") = 0 Then
                TheFileName = TheFileName & ".asc"
            End If
            
            ChDrive Mid$(App.Path, 1, 3)
            ChDir App.Path
            
            vb2spgpContext.TextMode = 0
            vb2spgpContext.Armor = 1
            
            If PGPConvent.Checked = True Then
                vb2spgpContext.ConventionalEncrypt = 1
                spgpEncryptMessage
            Else
                vb2spgpContext.ConventionalEncrypt = 0
                spgpEncryptMessage
            End If
            bRes = PutFileText(TheFileName, MessageArea.Text)
            If bRes = True Then
                ShowStatus 1, "Success - Message saved at " & TheFileName
            Else
                ShowStatus 1, "Error detected - check " & TheFileName
            End If
        Else
        '---------------------------------------------
        'no, it is a file encrypt
        '---------------------------------------------
            CommonDialog1.DialogTitle = "Specify the file you wish to encrypt"
            CommonDialog1.Flags = &H2& + &H4&
            CommonDialog1.Filter = "All Files (*.*)|*.*"
            CommonDialog1.FilterIndex = 1
            CommonDialog1.CancelError = True
            CommonDialog1.InitDir = App.Path
            'CommonDialog1.Action = 1
            CommonDialog1.ShowOpen
            TheFileName = CommonDialog1.FileName
            If Not InStr(TheFileName, ".asc") = 0 Then
                MsgBox "Please rename the file to not have the extension '.asc' as this will be the extension of the save file.", vbApplicationModal + vbCritical, "Check for filename extension."
                gCancelAction = False
                Exit Sub
            End If
            ChDrive Mid$(App.Path, 1, 3)
            ChDir App.Path
            vb2spgpContext.TextMode = 0
            vb2spgpContext.Armor = 1
            vb2spgpContext.FileIn = TheFileName
            vb2spgpContext.FileOut = TheFileName & ".asc"
            If PGPConvent.Checked Then
                vb2spgpContext.ConventionalEncrypt = 1
            Else
                vb2spgpContext.ConventionalEncrypt = 0
            End If
            spgpEncryptfile
            ShowStatus 1, "Encrypted file save as " & vb2spgpContext.FileOut
        End If
    Else
        '------------------------
        'Just encrypt and sign the message
        '--------------------------

        vb2spgpContext.TextMode = 0
        vb2spgpContext.Armor = 1
        spgpEncryptMessage
        If Not ClearSign Then EncryptAttachments (gPGPKeyID)
    End If
        
        
gCancelAction = False
Exit Sub
FileEnSError:
        gCancelAction = False
        MsgBox "There was an error.  The reason given by the operating system is: " & Err.Description, vbApplicationModal, App.Title
        Err.Clear
End Sub

Public Function spgpVerifyMessage()
' all strings must be of fixed length
  Dim BufferIn As String
  Dim BufferOut As String   ' passes clear-text, receives cipher-text
  Dim pPass As String * 256
  Dim SigProps As String * 256
  Dim spgperr As String * 256
  Dim UserID As String * 256
  Dim Sig As TSig_Data
  'Dim Key As TKey_Data
  Dim AnalysisData As Long
  Dim i As Long
  Dim J As Integer
 
  BufferIn = String(Len(MessageArea.Text & Chr(0)), Chr(0)) ' final null is required
  BufferIn = MessageArea.Text & Chr(0)
  BufferOut = String(Len(BufferIn) + 1024, Chr(0))
  SigProps = "" & Chr(0)         ' output buffers should be initialised & terminated
  'pPass = GetPassPhrase(0, BufferIn)
  'If Trim(pPass) = "" Then Exit Function
  i = spgpDecode(BufferIn, BufferOut, Len(BufferOut), pPass, SigProps)
  If Not i = 0 Then
    Call spgpGetErrorString(i, spgperr)
    Err.Raise 2000, "An error has occurred verifying the message: ", spgperr
    Beep
    ShowStatus 1, Err.Description
  Else
    Sig = ParseSigData(SigProps)
    Select Case Sig.Status
        Case Is = "SIGNED_BAD"
            Err.Raise 1001, "Bad signature.", "Bad signature detected."
            Beep
            ShowStatus 1, Err.Description
        Case Is = "SIGNED_GOOD"
            ShowStatus 1, "The message is okay and signature is good."
            spgpDecryptMessage
            MessageArea.SelStart = 0
            MessageArea.Text = StripNulls(BufferOut)
    End Select
End If
End Function

Public Sub RemoveUnderlineFromToBox()
txtTo.SelStart = 0
txtTo.SelLength = Len(txtTo.Text)
txtTo.SelUnderline = False
txtTo.SelLength = 0

End Sub
Public Sub RemoveUnderlineFromccBox()
txtCC.SelStart = 0
txtCC.SelLength = Len(txtCC.Text)
txtCC.SelUnderline = False
txtCC.SelLength = 0
End Sub

Public Sub AddMembersToAddressList(MemberArray As Variant, iAddressListType As Integer)
Dim VarArray As Variant
Dim sMember As String
Dim i As Integer
Dim lIndex As Long
Dim sResult As Boolean
'Now add the members to the list
    For i = 1 To MemberArray(0)
        sMember = MemberArray(i)
       ' DisplayName = sMember
        'ContactEmailAddress = StripEMailAddress(sMember)
        'ContactFullName = StripFullName(sMember)
        'The first item is the name of the group
        VarArray = Array(StripFullName(sMember), _
                    StripFullName(sMember), _
                    StripEMailAddress(sMember), _
                    MemberArray(MemberArray(0) + 1)) 'This is whether in is in DB or PGP
        'The last value in the MemberArray+1 holds the type of address list
        lIndex = AddressList.AddContact(VarArray, iAddressListType)
    Next
End Sub

Public Sub ConvertMailGroupsToAddressList(iAddressListType As Integer)
Dim VarArray As Variant
Dim i As Integer
Dim sResult As String
Dim MemberArray As Variant
Dim sMember As String

If iAddressListType = CONTACT_TO_LIST Then
    VarArray = AddressList.GetAllMailGroups(CONTACT_TO_LIST)
        If Not VarArray(0) = 0 Then
            For i = 1 To VarArray(0)
                sMember = VarArray(i)
                MemberArray = GetMembersOfGroup(sMember)
                'Note: the last member of the array +1 determines where the contact came from eg DB, PGP
                AddMembersToAddressList MemberArray, CONTACT_TO_LIST
                'Now Remove the MailGroup Name from the List as it is no longer needed
                sResult = AddressList.RemoveContact(sMember, CONTACT_TO_LIST)
            Next
        End If
End If

If iAddressListType = CONTACT_CC_LIST Then
        VarArray = AddressList.GetAllMailGroups(CONTACT_CC_LIST)
        If Not VarArray(0) = 0 Then
            For i = 1 To VarArray(0)
                sMember = VarArray(i)
                MemberArray = GetMembersOfGroup(sMember)
                'Note: the last member of the array +1 holds the type of address list
                AddMembersToAddressList MemberArray, CONTACT_CC_LIST
                sResult = AddressList.RemoveContact(sMember, CONTACT_CC_LIST)
            Next
        End If
End If
End Sub

Public Sub ProcessNymCommand(NymState As Integer, ListIndex As Integer)
Dim i As Integer
    Dim J As Integer
    Dim msg As String
    Dim Vbresponse As Integer
    Dim FileNum As Integer
    Dim CopyFileX As Integer
    Dim tmpstr1 As String
    Dim tmpstr3 As String
    Dim tmpstr2 As String
    Dim tmpNym As String
    Dim tmpNymServer As String
    Dim tmpFullNym As String
    Dim rs As Recordset
    Dim TextLine As String
    
    On Error GoTo SetNymError
   gNymState = NymState
   If Not NymState = gNYMCONFIG Then
   
        Set rs = DB.OpenRecordset("Nyms", dbOpenDynaset)
        If rs.EOF Or ListIndex = -1 Then
            Beep
            ShowStatus 1, "No Nyms selected or no Nyms in database."
            Exit Sub
        End If
        rs.MoveFirst
        For i = 0 To ListIndex - 1
            rs.MoveNext
        Next
    
        'NOTE rs is now used throughout
        Nym.ID = rs("Nym Email")
        Nym.name = rs("Nym Full Name")
        Nym.Server = rs("Nym Server")
        'Nym.name = rs("Name")
        Nym.acksend = rs("acksend")
        Nym.signsend = rs("signsend")
        Nym.cryptrecv = rs("cryptrecv")
        Nym.fixedsize = rs("fixedsize")
        Nym.disable = rs("disable")
        Nym.fingerkey = rs("fingerkey")
        Nym.EmailAddress = rs("EmailAddress")
        Nym.NewsGroupReplyEmail = rs("NewsGroupReplyEmail")
        Nym.NewsGroupReplyGroup = rs("NewsGroupReplyGroup")
        Nym.NewsGroupReplySubject = rs("NewsGroupReplySubject")
        Nym.UseNewsGroupReply = rs("UseNewsGroupReply")
    
        Nym.LatentTime = rs("latenttime")

        If IsNull(rs("Nym Passphrases")) Then
            Nym.PassPhrase(0) = ""
        Else
            Nym.PassPhrase(0) = rs("Nym Passphrases")
        End If
    End If
    Select Case NymState
            Case gNYM_DECRYPT
                '''
                If MessageArea.Text = "" Then Exit Sub
                FileNum = FreeFile
                Open gPGPPath + "\" + gPGPFile For Output As FileNum
                Print #FileNum, MessageArea.Text
                Close #FileNum
                Nym.PassPhrase(0) = "r3181147"
                ExecCmd (gPGPPath & "\pgp " + gPGPFile + " -o " + gPGPPath + "\" + gPGPFile & ".out" & " -z " + Chr$(34) + Nym.PassPhrase(0) + Chr$(34))
                Nym.PassPhrase(0) = "r3181147"
                ExecCmd (gPGPPath & "\pgp " + gPGPFile + " -o " + gPGPPath + "\" + gPGPFile & ".out" & " -z " + Chr$(34) + Nym.PassPhrase(0) + Chr$(34))
             
               ' ExecCmd (cmd)
                FileNum = FreeFile
                Open gPGPPath + "\" + gPGPFile & ".out" For Output As FileNum
                MessageArea.Text = ""
                MessageArea.SelStart = 0
                While Not EOF(FileNum)
                    Line Input #FileNum, TextLine
                    MessageArea.SelText = TextLine & vbCrLf
                Wend
                Close #FileNum
                Kill gPGPPath + "\" + gPGPFile
                WipeFile (gPGPPath + "\" + gPGPFile)
            Case gNYMDEL
                rs.Delete
                rs.Close
                txtTo.Text = Nym.Server
                'If chkServer.Value = vbChecked Then
                    MessageArea.SelStart = 0
                    i = InStr(1, Nym.Server, "@")
                    tmpstr1 = Mid$(Nym.Server, i, (Len(Nym.Server) - i) + 1)
    
                    MessageArea.SelText = "Config: " + vbCrLf + "From: " + Nym.ID + vbCrLf + "Nym-Commands: delete" & vbCrLf & vbCrLf
                    'frmMultiNyms.Hide
                    txtTo.Text = Nym.Server
                    vb2spgpContext.Initialise
                    vb2spgpContext.TextMode = 1
                    vb2spgpContext.Armor = 1
                    vb2spgpContext.KeyEncrypt = 1
                    vb2spgpContext.Sign = 1
                    vb2spgpContext.SignKeyID = Nym.ID
                    vb2spgpContext.CryptKeyID = txtTo.Text
                    spgpEncryptMessage
                'End If
            Case gNYMRPLYCHANGE
                    'Okay get updated data
                                       
                    frmCreateNymStep5.Caption = "Change Reply Block"
                    frmCreateNymStep5.lblPrompt = "Enter changes to your existing Reply Block"
                    frmCreateNymStep5.Show vbModal
                    'frmCreateNymStep5.Hide
                    Set frmCreateNymStep5 = Nothing
                    If gCancelAction Then
                        rs.Close
                        Exit Sub
                    End If
                    'Now save in the databse
                    rs.Edit
                    rs("Nym Email") = Nym.ID
    
                    rs("Nym Server") = Nym.Server
                    rs("acksend") = Nym.acksend
                    rs("signsend") = Nym.signsend
                    rs("cryptrecv") = Nym.cryptrecv
                    rs("fixedsize") = Nym.fixedsize
                    rs("disable") = Nym.disable
                    rs("fingerkey") = Nym.fingerkey
                    'If it is a change and Nym.name = "" then don't alter in database
                    If Not Nym.name = "" Then
                        'rs("Full Name") = IIf(Nym.name = "", " ", Nym.name)
                        rs("Nym Full Name") = IIf(Nym.name = "", " ", Nym.name)
                    End If
                   
                    rs("NewsGroupReplyEmail") = IIf(Nym.NewsGroupReplyEmail = "", " ", Nym.NewsGroupReplyEmail)
                    rs("NewsGroupReplyGroup") = IIf(Nym.NewsGroupReplyGroup = "", " ", Nym.NewsGroupReplyGroup)
                    rs("NewsGroupReplySubject") = IIf(Nym.NewsGroupReplySubject = "", " ", Nym.NewsGroupReplySubject)
                    rs("UseNewsGroupReply") = Nym.UseNewsGroupReply
                    
                    rs("latenttime") = Nym.LatentTime
                    rs("Nym PassPhrases") = IIf(Nym.PassPhrase(0) = "", Null, Nym.PassPhrase(0))
                    rs.Update
                    rs.Close
                    If NymState = gNYM_IDLE Then Exit Sub
                    
                    'See if we need to use remailers
                    If Nym.UseNewsGroupReply Then
                        gRemailerType = STANDARD_EMAIL
                    Else
                        gRemailerType = REMAILER_CYPHERPUNK
                    End If
                    
                    CreateReplyBlock
                    If gCancelAction Then Exit Sub
                    NymAliasConfig
                    'DoEvents
                    txtTo.Text = Nym.Server
                    vb2spgpContext.Initialise
                    vb2spgpContext.TextMode = 1 '**MUST BE 1 for NYM.ALIAS to work
                    vb2spgpContext.Armor = 1
                    vb2spgpContext.KeyEncrypt = 1
                    vb2spgpContext.Sign = 1
                    vb2spgpContext.SignKeyID = Nym.ID
                    vb2spgpContext.CryptKeyID = txtTo.Text
                    spgpEncryptMessage
                    ShowStatus 1, "You can now send the Reply Block Change request."
            Case gNYMPREPARE
                ShowStatus 1, "Preparing Nym message..."
                MessageArea.SelStart = 0
                rs.Close
                i = InStr(1, Nym.Server, "@")
                tmpstr1 = Mid$(Nym.Server, i, (Len(Nym.Server) - i) + 1)
                'If Not IsNewNym(Nym.Server) Then
                  '  MessageArea.SelText = "From: " & Nym.ID & tmpstr1 & " (" + Nym.name + ")" & vbCrLf & "Password: " & vbCrLf + "To: " & txtTo.Text & vbCrLf & "Subject: " & txtsubject.Text & vbCrLf & vbCrLf
                   ' txtTo.Text = Nym.Server
                   ' txtsubject.Text = ""
                   ' MsgBox "Add your Nym passsword to the message, then encrypt the message before sending."
                'Else
                    MessageArea.SelText = "From: " & Nym.ID & vbCrLf
                    MessageArea.SelText = "To: " & txtTo.Text & vbCrLf
                    MessageArea.SelText = "Subject: " & txtsubject.Text & vbCrLf
                    MessageArea.SelText = "CC: " & txtCC.Text & vbCrLf
                    MessageArea.SelText = vbCrLf
                    txtsubject.Text = ""
                    vb2spgpContext.Initialise
                    vb2spgpContext.Armor = 1
                    vb2spgpContext.TextMode = 1 'was 0
                    vb2spgpContext.KeyEncrypt = 1
                    vb2spgpContext.Sign = 1
                    vb2spgpContext.SignKeyID = Nym.ID
                    vb2spgpContext.CryptKeyID = Nym.Server
                    ShowStatus 1, "Encrypting Nym message..."
                    spgpEncryptMessage
                    If gCancelAction Then
                        ShowStatus 1, "Message cancelled."
                    Else
                        ShowStatus 1, "Copying file to message area....."
                        DoEvents
                        txtTo.Text = "send" & Mid(Nym.Server, InStr(1, Nym.Server, "@"), (Len(Nym.Server)))
                        txtCC = ""
                        ShowStatus 1, "Nym message successfully created."
                    End If
               ' End If
                'KillTemporaryFiles
            Case gNYM_USENET_PREPARE
                rs.Close
                MessageArea.SelStart = 0
                i = InStr(1, Nym.Server, "@")
                tmpstr1 = Mid$(Nym.Server, i, (Len(Nym.Server) - i) + 1)
               ' If Not IsNewNym(Nym.Server) Then
                  '  MessageArea.SelText = "From: " & Nym.ID & tmpstr1 & " (" + Nym.Name + ")" & vbCrLf & "Password: " & vbCrLf + "To: " & txtTo.Text & vbCrLf & "Subject: " & txtsubject.Text & vbCrLf & vbCrLf
                  ' txtTo.Text = Nym.Server
                   ' txtsubject.Text = ""
                   ' MsgBox "Add your Nym passsword to the message, then encrypt the message before sending."
                 'Else
                    PrepareUSENETMessage (Nym.ID)
                    vb2spgpContext.Initialise
                    vb2spgpContext.Armor = 1
                    vb2spgpContext.TextMode = 0
                   ''
                    vb2spgpContext.KeyEncrypt = 0
                    vb2spgpContext.Sign = 1
                    vb2spgpContext.SignKeyID = Nym.ID
                    vb2spgpContext.CryptKeyID = Nym.Server
                    spgpEncryptMessage
                    txtTo.Text = "send" & Mid(Nym.Server, InStr(1, Nym.Server, "@"), (Len(Nym.Server)))
               ' End If
                ShowStatus 1, "You can now send the message."
            Case gNYMCONFIG
                Set rs = DB.OpenRecordset("Nyms", dbOpenDynaset)
                rs.AddNew
                rs("Nym Email") = Nym.ID
                rs("Nym Full Name") = Nym.name
                rs("Nym Server") = Nym.Server
                rs("acksend") = Nym.acksend
                rs("signsend") = Nym.signsend
                rs("cryptrecv") = Nym.cryptrecv
                rs("fixedsize") = Nym.fixedsize
                rs("disable") = Nym.disable
                rs("fingerkey") = Nym.fingerkey
                rs("EmailAddress") = IIf(Nym.EmailAddress = "", " ", Nym.EmailAddress)
    
                rs("NewsGroupReplyEmail") = IIf(Nym.NewsGroupReplyEmail = "", " ", Nym.NewsGroupReplyEmail)
                rs("NewsGroupReplyGroup") = IIf(Nym.NewsGroupReplyGroup = "", " ", Nym.NewsGroupReplyGroup)
                rs("NewsGroupReplySubject") = IIf(Nym.NewsGroupReplySubject = "", " ", Nym.NewsGroupReplySubject)
                rs("UseNewsGroupReply") = Nym.UseNewsGroupReply
    
                If Nym.ChangeName Then rs("Nym Full Name") = Nym.name 'update
                rs("latenttime") = Nym.LatentTime
                tmpstr1 = ""
                For i = 1 To UBound(Nym.PassPhrase())
                    tmpstr1 = tmpstr1 & Nym.PassPhrase(i) & ","
                Next
                rs("Nym PassPhrases") = tmpstr1
                rs.Update
                rs.Close
                If Nym.UseNewsGroupReply Then
                    DontUseRemailer
                Else
                    UseCypherPunk
                End If
                CreateReplyBlock
                NymAliasConfig
                txtTo.Text = Nym.Server
                If Not gManualEncrypt Then
                    vb2spgpContext.Initialise
                    vb2spgpContext.Armor = 1
                    vb2spgpContext.TextMode = 1 '***THIS MUST BE 1 FOR NYMS TO WORK ***
                    vb2spgpContext.Sign = 1
                    vb2spgpContext.SignKeyID = Nym.ID
                    vb2spgpContext.KeyEncrypt = 1
                    vb2spgpContext.CryptKeyID = Nym.Server
                    spgpEncryptMessage  'EncryptMessage , "-seatws"
                End If
            End Select
    KillTemporaryFiles
    'gCancelAction = False
    Exit Sub

SetNymError:
    KillTemporaryFiles
    Reset
    MsgBox "Trouble processing the Nym command.  Error was: " & Err.Description, vbApplicationModal + vbCritical, App.Title
    Unload frmMultiNyms
    Exit Sub
End Sub

Public Function IsNewNym(theName As String) As Boolean
Dim i As Integer

    IsNewNym = False
    For i = 1 To gTotalRemailers
        If gRemailerArray(i).name = theName Then
            If gRemailerArray(i).newnym = 1 Then IsNewNym = True
            Exit For
        End If
Next
End Function


Public Sub NymAliasConfig()
Dim NymConfigStr As String
    MessageArea.SelStart = 0
    MessageArea.SelText = "Config:" & vbCrLf
    MessageArea.SelText = "From: " & Nym.ID & vbCrLf
    If gNymState = gNYMCONFIG Then
        NymConfigStr = IIf(Nym.create, " create", "")
        NymConfigStr = NymConfigStr & IIf(Nym.acksend, " +acksend", " -acksend")
        NymConfigStr = NymConfigStr & IIf(Nym.signsend, " +signsend", " -signsend")
        NymConfigStr = NymConfigStr & IIf(Nym.cryptrecv, " +cryptrecv", " -cryptrecv")
        NymConfigStr = NymConfigStr & IIf(Nym.disable, " +disable", " -disable")
        NymConfigStr = NymConfigStr & IIf(Nym.fingerkey, " +fingerkey", " -fingerkey")
        NymConfigStr = NymConfigStr & IIf(Nym.name = "", "", " name=" & Chr$(34) & Nym.name & Chr$(34))
        MessageArea.SelText = "Nym-Commands:" & NymConfigStr & vbCrLf
       ' MessageArea.SelText = "Nym-Commands: create +acksend +signsend name=" & Chr$(34) & Nym.name & Chr$(34) & vbCrLf
        MessageArea.SelText = "Public-Key:" & vbCrLf
        InsertKey (NYMKEY)
    Else
        If gNymState = gNYMRPLYCHANGE Then
        NymConfigStr = IIf(Nym.acksend, " +acksend", " -acksend")
        NymConfigStr = NymConfigStr & IIf(Nym.signsend, " +signsend", " -signsend")
        NymConfigStr = NymConfigStr & IIf(Nym.cryptrecv, " +cryptrecv", " -cryptrecv")
        NymConfigStr = NymConfigStr & IIf(Nym.disable, " +disable", " -disable")
        NymConfigStr = NymConfigStr & IIf(Nym.fingerkey, " +fingerkey", " -fingerkey")
        NymConfigStr = NymConfigStr & IIf(Nym.ChangeName, " name=" & Chr$(34) & Nym.name & Chr$(34), "")
        
        MessageArea.SelText = "Nym-Commands:" & NymConfigStr & vbCrLf
       ' MessageArea.SelText = "Nym-Commands: create +acksend +signsend name=" & Chr$(34) & Nym.name & Chr$(34) & vbCrLf
        MessageArea.SelText = "Public-Key:" & vbCrLf
        InsertKey (NYMKEY)
        End If
    End If
End Sub

Private Sub PrepareUSENETMessage(From As String)
MessageArea.SelStart = 0
MessageArea.SelText = "From: " & From & vbLf
'tmpstr1 = Mid$(txtTo.Text, i, (Len(txtTo.Text) - i) + 1)
MessageArea.SelText = "To: " & txtTo.Text & vbLf
MessageArea.SelText = MailHeader(1).ID & MailHeader(1).Value & vbLf
MessageArea.SelText = MailHeader(2).ID & MailHeader(2).Value & vbLf
If Not MailHeader(3).Value = "" Then MessageArea.SelText = MailHeader(3).ID & "<" & MailHeader(3).Value & ">" & vbLf
If Not MailHeader(4).Value = "" Then MessageArea.SelText = MailHeader(4).ID & "<" & MailHeader(4).Value & ">" & vbLf
If Not MailHeader(5).Value = "" Then MessageArea.SelText = MailHeader(5).ID & MailHeader(5).Value & vbLf

MessageArea.SelText = vbLf
txtsubject.Text = ""
'i = InStr(1, txtTo.Text, "@")
txtTo.Text = "send" & Mid$(txtTo.Text, InStr(1, txtTo.Text, "@"), (Len(txtTo.Text)))
End Sub

Public Sub SendNymMessage()
Dim msg As String
Dim Vbresponse As Integer

    gCancelAction = False
    DoEvents
    'If Nym.DontUseRemailer Then
      '  DontUseRemailer
    'Else
     '   UseCypherPunk
   ' End If
 
    'If (gRemailerType = SEND_MESSAGES_USING_NYM) Or (gRemailerType = STANDARD_EMAIL) Then
    ' If Nym.DontUseRemailer Then
        msg = "Your Nym Message has been prepared.  You must decide whether or not to send this Nym Message/Nym Configuration request through a remailer. "
        msg = msg & "You can send it directly from your email account if you wish however it is not the safest method as the source of the Nym traffic may be able to be determined as coming from your e-mail account. " & vbCrLf & vbCrLf
        msg = msg & "If you want to send this Nym Message via a standard Remailer, click 'Yes' " & vbCrLf & vbCrLf
        msg = msg & "If you want to send this Nym Message directly to the Nym server for processing, click 'No' " & vbCrLf & vbCrLf
        msg = msg & "If you wish to Cancel out, then click  'Cancel'"
        Vbresponse = MsgBox(msg, vbYesNoCancel + vbApplicationModal + vbQuestion, "Use Remailer?")
        Select Case Vbresponse
            Case vbCancel
                gCancelAction = True
            Case vbYes
                gRemailerType = REMAILER_CYPHERPUNK
                AddressList.Initialise
                RemoveUnderlineFromToBox
                RemoveUnderlineFromccBox
                ScanTextForContacts txtTo, CONTACT_TO_LIST
                ScanTextForContacts txtCC, CONTACT_CC_LIST
                SendPIMessage
            Case vbNo
                SendToOutBox
        End Select
End Sub


