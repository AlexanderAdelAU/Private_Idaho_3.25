VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCreateNymStep5 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Creating a Nym - Step 5"
   ClientHeight    =   6765
   ClientLeft      =   465
   ClientTop       =   810
   ClientWidth     =   9885
   ControlBox      =   0   'False
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
   LinkTopic       =   "Form21"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6765
   ScaleWidth      =   9885
   Begin VB.CommandButton Command3 
      Caption         =   "Help"
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
      Left            =   6300
      TabIndex        =   56
      Top             =   6240
      Width           =   1035
   End
   Begin TabDlg.SSTab tabNymOptions 
      Height          =   3375
      Left            =   120
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2730
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   5953
      _Version        =   393216
      Style           =   1
      Tabs            =   9
      Tab             =   6
      TabsPerRow      =   9
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "acksend"
      TabPicture(0)   =   "C2ORG5.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "acksend(0)"
      Tab(0).Control(1)=   "Label1(3)"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "signsend"
      TabPicture(1)   =   "C2ORG5.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "acksend(2)"
      Tab(1).Control(1)=   "Label1(1)"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "cryptrecv"
      TabPicture(2)   =   "C2ORG5.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cryptrecv(0)"
      Tab(2).Control(1)=   "Label1(4)"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "fixedsize"
      TabPicture(3)   =   "C2ORG5.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "acksend(4)"
      Tab(3).Control(1)=   "Label1(5)"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "disable"
      TabPicture(4)   =   "C2ORG5.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "disable(5)"
      Tab(4).Control(1)=   "Label1(6)"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "fingerkey"
      TabPicture(5)   =   "C2ORG5.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "fingerkey(6)"
      Tab(5).Control(1)=   "Label1(7)"
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "name (Alias)"
      TabPicture(6)   =   "C2ORG5.frx":00A8
      Tab(6).ControlEnabled=   -1  'True
      Tab(6).Control(0)=   "Label1(8)"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "lblnamestatus"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "name(7)"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).ControlCount=   3
      TabCaption(7)   =   "Latent-time"
      TabPicture(7)   =   "C2ORG5.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "acksend(1)"
      Tab(7).Control(1)=   "Label1(2)"
      Tab(7).ControlCount=   2
      TabCaption(8)   =   "E-Mail Address"
      TabPicture(8)   =   "C2ORG5.frx":00E0
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Label1(9)"
      Tab(8).Control(1)=   "lblNote"
      Tab(8).Control(2)=   "fReplyBlock(1)"
      Tab(8).Control(3)=   "fReplyBlock(0)"
      Tab(8).Control(4)=   "optReplyBlock(0)"
      Tab(8).Control(5)=   "optReplyBlock(1)"
      Tab(8).ControlCount=   6
      Begin VB.OptionButton optReplyBlock 
         Caption         =   "Use NewsGroup Reply Address"
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
         Left            =   -73950
         TabIndex        =   58
         Top             =   1980
         Width           =   2535
      End
      Begin VB.OptionButton optReplyBlock 
         Caption         =   "Use Email Reply Address"
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
         Index           =   0
         Left            =   -73950
         TabIndex        =   57
         Top             =   1590
         Value           =   -1  'True
         Width           =   2385
      End
      Begin VB.Frame fReplyBlock 
         Caption         =   "Use your Emal Address as your Reply Block"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Index           =   0
         Left            =   -70170
         TabIndex        =   54
         Top             =   1110
         Width           =   4425
         Begin VB.TextBox txtEmailAddress 
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
            Left            =   120
            TabIndex        =   4
            Top             =   510
            Width           =   3525
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Your real email address"
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
            Height          =   225
            Index           =   10
            Left            =   120
            TabIndex        =   55
            Top             =   270
            Width           =   2535
         End
      End
      Begin VB.Frame fReplyBlock 
         Caption         =   "Use Newsgroup as your Reply Block"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2115
         Index           =   1
         Left            =   -70170
         TabIndex        =   48
         Top             =   1140
         Visible         =   0   'False
         Width           =   4425
         Begin VB.TextBox txtNewsGroupEmail 
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
            Index           =   2
            Left            =   120
            TabIndex        =   53
            Text            =   "Something you recognise"
            Top             =   1710
            Width           =   3645
         End
         Begin VB.TextBox txtNewsGroupEmail 
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
            Index           =   1
            Left            =   120
            TabIndex        =   51
            Text            =   "alt.anonymous.messages"
            Top             =   1140
            Width           =   3645
         End
         Begin VB.TextBox txtNewsGroupEmail 
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
            Index           =   0
            Left            =   120
            TabIndex        =   50
            Text            =   "mail2news@anon.lcs.mit.edu"
            Top             =   510
            Width           =   3645
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Reply Block Subject"
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
            Height          =   225
            Index           =   13
            Left            =   120
            TabIndex        =   59
            Top             =   1500
            Width           =   2535
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "NewsGroup Name"
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
            Height          =   225
            Index           =   12
            Left            =   120
            TabIndex        =   52
            Top             =   900
            Width           =   2535
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Mail2News Service"
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
            Height          =   225
            Index           =   11
            Left            =   120
            TabIndex        =   49
            Top             =   270
            Width           =   2535
         End
      End
      Begin VB.Frame name 
         Caption         =   "name="
         Height          =   735
         Index           =   7
         Left            =   5130
         TabIndex        =   43
         Top             =   420
         Width           =   3105
         Begin VB.TextBox txtName 
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
            Left            =   240
            TabIndex        =   45
            Top             =   240
            Width           =   2775
         End
      End
      Begin VB.Frame fingerkey 
         Caption         =   "fingerkey"
         Height          =   735
         Index           =   6
         Left            =   -70080
         TabIndex        =   39
         Top             =   660
         Width           =   2145
         Begin VB.OptionButton optfingerkey 
            Caption         =   "No"
            Height          =   315
            Index           =   1
            Left            =   1290
            TabIndex        =   41
            Top             =   300
            Value           =   -1  'True
            Width           =   795
         End
         Begin VB.OptionButton optfingerkey 
            Caption         =   "Yes"
            Height          =   315
            Index           =   0
            Left            =   390
            TabIndex        =   40
            Top             =   300
            Width           =   795
         End
      End
      Begin VB.Frame disable 
         Caption         =   "disable"
         Height          =   735
         Index           =   5
         Left            =   -70080
         TabIndex        =   35
         Top             =   660
         Width           =   2145
         Begin VB.OptionButton optdisable 
            Caption         =   "No"
            Height          =   315
            Index           =   1
            Left            =   1290
            TabIndex        =   37
            Top             =   300
            Value           =   -1  'True
            Width           =   795
         End
         Begin VB.OptionButton optdisable 
            Caption         =   "Yes"
            Height          =   315
            Index           =   0
            Left            =   390
            TabIndex        =   36
            Top             =   300
            Width           =   795
         End
      End
      Begin VB.Frame acksend 
         Caption         =   "fixedsize"
         Height          =   735
         Index           =   4
         Left            =   -70080
         TabIndex        =   31
         Top             =   660
         Width           =   2145
         Begin VB.OptionButton optfixedsize 
            Caption         =   "No"
            Height          =   315
            Index           =   1
            Left            =   1290
            TabIndex        =   33
            Top             =   300
            Value           =   -1  'True
            Width           =   795
         End
         Begin VB.OptionButton optfixedsize 
            Caption         =   "Yes"
            Height          =   315
            Index           =   0
            Left            =   390
            TabIndex        =   32
            Top             =   300
            Width           =   795
         End
      End
      Begin VB.Frame cryptrecv 
         Caption         =   "cryptrecv"
         Height          =   735
         Index           =   0
         Left            =   -70080
         TabIndex        =   27
         Top             =   660
         Width           =   2145
         Begin VB.OptionButton optcryptrecv 
            Caption         =   "No"
            Height          =   315
            Index           =   1
            Left            =   1290
            TabIndex        =   29
            Top             =   300
            Width           =   795
         End
         Begin VB.OptionButton optcryptrecv 
            Caption         =   "Yes"
            Height          =   315
            Index           =   0
            Left            =   390
            TabIndex        =   28
            Top             =   300
            Value           =   -1  'True
            Width           =   795
         End
      End
      Begin VB.Frame acksend 
         Caption         =   "signsend"
         Height          =   735
         Index           =   2
         Left            =   -70080
         TabIndex        =   23
         Top             =   660
         Width           =   2145
         Begin VB.OptionButton optsignsend 
            Caption         =   "No"
            Height          =   315
            Index           =   1
            Left            =   1290
            TabIndex        =   25
            Top             =   300
            Value           =   -1  'True
            Width           =   795
         End
         Begin VB.OptionButton optsignsend 
            Caption         =   "Yes"
            Height          =   315
            Index           =   0
            Left            =   390
            TabIndex        =   24
            Top             =   300
            Width           =   795
         End
      End
      Begin VB.Frame acksend 
         Caption         =   "Latent-Time"
         Height          =   735
         Index           =   1
         Left            =   -70080
         TabIndex        =   19
         Top             =   660
         Width           =   2145
         Begin VB.TextBox txtLatentTime 
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
            Left            =   480
            TabIndex        =   20
            Text            =   "+0:00"
            Top             =   240
            Width           =   705
         End
      End
      Begin VB.Frame acksend 
         Caption         =   "acksend"
         Height          =   735
         Index           =   0
         Left            =   -70080
         TabIndex        =   15
         Top             =   660
         Width           =   2145
         Begin VB.OptionButton optacksend 
            Caption         =   "Yes"
            Height          =   315
            Index           =   0
            Left            =   390
            TabIndex        =   17
            Top             =   300
            Width           =   795
         End
         Begin VB.OptionButton optacksend 
            Caption         =   "No"
            Height          =   315
            Index           =   1
            Left            =   1320
            TabIndex        =   16
            Top             =   300
            Value           =   -1  'True
            Width           =   795
         End
      End
      Begin VB.Label lblNote 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Note: If your reply block is to a newsgroup then you don't need a remailer in your reply block."
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
         Height          =   705
         Left            =   -73890
         TabIndex        =   60
         Top             =   2400
         Width           =   3165
      End
      Begin VB.Label lblnamestatus 
         Caption         =   "lblnamestatus"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1155
         Left            =   5130
         TabIndex        =   47
         Top             =   1200
         Width           =   3255
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   $"C2ORG5.frx":00FC
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
         Height          =   675
         Index           =   9
         Left            =   -74910
         TabIndex        =   46
         Top             =   450
         Width           =   9315
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   $"C2ORG5.frx":0264
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
         Height          =   1155
         Index           =   8
         Left            =   120
         TabIndex        =   44
         Top             =   540
         Width           =   4695
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   """'fingerkey""' will allow people to obtain your nym's PGP public key by fingering your E-Mail address."
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
         Height          =   735
         Index           =   7
         Left            =   -74520
         TabIndex        =   42
         Top             =   780
         Width           =   3975
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   $"C2ORG5.frx":03A1
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
         Height          =   975
         Index           =   6
         Left            =   -74520
         TabIndex        =   38
         Top             =   660
         Width           =   3975
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   $"C2ORG5.frx":0493
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
         Height          =   1095
         Index           =   5
         Left            =   -74400
         TabIndex        =   34
         Top             =   660
         Width           =   4095
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   $"C2ORG5.frx":0567
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
         Height          =   915
         Index           =   4
         Left            =   -74640
         TabIndex        =   30
         Top             =   660
         Width           =   4335
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   $"C2ORG5.frx":0612
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
         Height          =   1095
         Index           =   1
         Left            =   -74640
         TabIndex        =   26
         Top             =   660
         Width           =   4335
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   $"C2ORG5.frx":0708
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
         Height          =   795
         Index           =   2
         Left            =   -74250
         TabIndex        =   21
         Top             =   660
         Width           =   3705
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   $"C2ORG5.frx":0799
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
         Height          =   1095
         Index           =   3
         Left            =   -74640
         TabIndex        =   18
         Top             =   660
         Width           =   4215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Remailer Passphrases (Only required if you will use a Remailer in your reply block.):"
      ClipControls    =   0   'False
      Height          =   1845
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   9615
      Begin VB.TextBox txtPassPhrase 
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
         Height          =   330
         Index           =   0
         Left            =   2340
         TabIndex        =   0
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox txtPassPhrase 
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
         Height          =   330
         Index           =   1
         Left            =   6780
         TabIndex        =   1
         Top             =   810
         Width           =   2445
      End
      Begin VB.TextBox txtPassPhrase 
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
         Height          =   330
         Index           =   2
         Left            =   2340
         TabIndex        =   2
         Top             =   1290
         Width           =   2295
      End
      Begin VB.TextBox txtPassPhrase 
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
         Height          =   330
         Index           =   3
         Left            =   6780
         TabIndex        =   3
         Top             =   1260
         Width           =   2415
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Passphrase  for Nym Server:"
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
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "Click for Help"
         Top             =   870
         Width           =   2175
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   $"C2ORG5.frx":0852
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
         Height          =   435
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   8055
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Passphrase for Remailer 1:"
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
         Height          =   255
         Index           =   1
         Left            =   4770
         TabIndex        =   11
         ToolTipText     =   "Click for Help"
         Top             =   870
         Width           =   1935
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Passphrase for Remailer 2:"
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
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "Click for Help"
         Top             =   1350
         Width           =   2175
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Passphrase for Remailer 3:"
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
         Height          =   255
         Index           =   3
         Left            =   4770
         TabIndex        =   9
         ToolTipText     =   "Click for Help"
         Top             =   1350
         Width           =   1965
      End
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Next"
      Default         =   -1  'True
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
      Left            =   7590
      TabIndex        =   5
      Top             =   6210
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
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
      Left            =   8790
      TabIndex        =   7
      Top             =   6210
      Width           =   975
   End
   Begin VB.Label lblWarning 
      BackStyle       =   0  'Transparent
      Caption         =   "Note:  If you typed in phasephrases for remailers above, you will need to select that many remailers in the next step!"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   480
      TabIndex        =   61
      Top             =   6180
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Options"
      Height          =   255
      Left            =   150
      TabIndex        =   22
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label lblPrompt 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Create your Reply Block."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   30
      Width           =   6975
   End
End
Attribute VB_Name = "frmCreateNymStep5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim i As Integer
    Dim rs As Recordset
    Dim msg As String
    
    If txtEmailAddress = "" And txtNewsGroupEmail(0) = "" Then
        msg = "For some reason your E-Mail address is blank. " & vbCrLf & vbCrLf
        msg = msg & "This could be due to a database update. Just fill in the correct address and press 'Done'.  " & vbCrLf & vbCrLf
        msg = msg & "Note: This is the address that the nym server will send your emails to after encryption."
        MsgBox msg, vbApplicationModal + vbCritical, "Missing Email Address"
        Exit Sub
    End If
    Nym.name = txtName
    Nym.LatentTime = txtLatentTime
    'Check options
    Nym.acksend = False
    Nym.signsend = False
    Nym.cryptrecv = False
    Nym.fixedsize = False
    Nym.disable = False
    Nym.fingerkey = False
    If optacksend(0).Value = True Then Nym.acksend = True
    If optsignsend(0).Value = True Then Nym.signsend = True
    If optcryptrecv(0).Value = True Then Nym.cryptrecv = True
    If optfixedsize(0).Value = True Then Nym.fixedsize = True
    If optdisable(0).Value = True Then Nym.disable = True
    If optfingerkey(0).Value = True Then Nym.fingerkey = True
    
    Nym.EmailAddress = txtEmailAddress
    Nym.NewsGroupReplyEmail = txtNewsGroupEmail(0)
    Nym.NewsGroupReplyGroup = txtNewsGroupEmail(1)
    Nym.NewsGroupReplySubject = txtNewsGroupEmail(2)
    
      
    If optReplyBlock(1).Value = True Then
        Nym.UseNewsGroupReply = True
    Else
         Nym.UseNewsGroupReply = False
    End If
    
    For i = 0 To txtPassPhrase.Count - 1
        Nym.PassPhrase(i + 1) = txtPassPhrase(i).Text
        Nym.PassPhrase(0) = Nym.PassPhrase(0) & txtPassPhrase(i).Text
    Next
    If Not gNymState = gNYMRPLYCHANGE Then
        Unload Me
        frmCreateNymStep6.Show
    End If
    Unload Me
End Sub

Private Sub Command2_Click()
  gNymState = gNYM_IDLE
  gCancelAction = True
  Unload Me
End Sub

Private Sub Command3_Click()
 On Error Resume Next
 Form32.Caption = "Creating a Reply Block"
    Form32.RichTextBox1 = GetFileText(App.Path & "\HelpText20.rtf")
    Form32.Show vbModal
End Sub

Private Sub Form_Load()
'Select Case gNymState
 '   Case gNYMRPLYCHANGE, gNYMPREPARE, gNYMCONFIG
      '  tabNymOptions.Enabled = True
'End Select
'If gRemailerType = STANDARD_EMAIL Then Frame1.Enabled = False
    

    Dim Win As New CWindow
    Win.Center Me, Null
    'Win.OnTop(Me) = True
    
    txtName = Nym.name
    txtEmailAddress = Nym.EmailAddress
    txtNewsGroupEmail(0) = IIf(Nym.NewsGroupReplyEmail = "", txtNewsGroupEmail(0), Nym.NewsGroupReplyEmail)
    txtNewsGroupEmail(1) = IIf(Nym.NewsGroupReplyGroup = "", txtNewsGroupEmail(1), Nym.NewsGroupReplyGroup)
    txtNewsGroupEmail(2) = IIf(Nym.NewsGroupReplySubject = "", txtNewsGroupEmail(2), Nym.NewsGroupReplySubject)
    If Nym.UseNewsGroupReply Then
        fReplyBlock(0).Visible = False
        fReplyBlock(1).Visible = True
        optReplyBlock(1).Value = True
        'PIForm(gActivePIInstance).cmbRemailerSelect.ListIndex = 0 'don't need remailer
    Else
        'fReplyBlock(1).Visible = False
        fReplyBlock(0).Visible = True
        fReplyBlock(1).Visible = False
        optReplyBlock(0).Value = True
        'PIForm(gActivePIInstance).cmbRemailerSelect.ListIndex = 1 'need remailer
    End If
    
    Nym.create = True
    Nym.ChangeName = False
    If Not gIsNewNym Then
        Command1.Caption = "Done"
        frmCreateNymStep5.Caption = "Select Remailer Chain"
        Nym.create = False
        If Nym.acksend = True Then
            optacksend(0).Value = True
        Else
            optacksend(1).Value = True
        End If
        If Nym.signsend = True Then
            optsignsend(0).Value = True
        Else
            optsignsend(1).Value = True
        End If
        'optsignsend(0).Value = Nym.signsend
        If Nym.cryptrecv = True Then
            optcryptrecv(0).Value = True
        Else
            optcryptrecv(1).Value = True
        End If
        'optcryptrecv(0).Value = Nym.cryptrecv
        If Nym.fixedsize = True Then
            optfixedsize(0).Value = True
        Else
            optfixedsize(1).Value = True
        End If
        'optfixedsize(0).Value = Nym.fixedsize
        If Nym.disable = True Then
            optdisable(0).Value = True
        Else
            optdisable(1).Value = True
        End If
        'optdisable(0).Value = Nym.disable
        If Nym.fingerkey = True Then
            optfingerkey(0).Value = True
        Else
            optfingerkey(1).Value = True
        End If
        'optfingerkey(0).Value = Nym.fingerkey
        txtLatentTime = Nym.LatentTime
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmCreateNymStep5 = Nothing
End Sub

Private Sub Label2_Click(Index As Integer)
Dim sHelp As String

sHelp = "1. The 'Passphrase for Nym Server' is used by the nym to encrypt any messages that anyone sends to your nym." & vbCrLf
sHelp = sHelp & "2. If you choose a remailer, then the 'Passphrase for Remailer N' is used by these Remailers to send the message from the Nym either directly to you or another Remailer, and so on." & vbCrLf & vbCrLf
sHelp = sHelp & "You need to remember these as you will use them to decrypt the encryption each remailer has applied to your message, finishing with the Nym server." & vbCrLf & vbCrLf
sHelp = sHelp & "In summary, Encryption will be applied by Nym->Remailer1->Remailer2 etc, so you need to decrypt in the reverse order, as you will receive your reply from Remailer 2.  Which would have received its message from Remailer 1, and it would have received its message from the Nym server."
MsgBox sHelp, vbApplicationModal + vbQuestion, "Phase Phrase help"

End Sub

Private Sub optReplyBlock_Click(Index As Integer)
Select Case Index
    Case 0
        fReplyBlock(0).Visible = True
        fReplyBlock(1).Visible = False
        Nym.UseNewsGroupReply = False
        Frame1.Enabled = True
        'PIForm(gActivePIInstance).cmbRemailerSelect.ListIndex = 1 'need remailer
    Case 1
        fReplyBlock(1).Visible = True
        fReplyBlock(0).Visible = False
        Nym.UseNewsGroupReply = True
        Frame1.Enabled = False
        'PIForm(gActivePIInstance).cmbRemailerSelect.ListIndex = 0 'don't need remailer
        
End Select
End Sub

Private Sub txtName_Change()
If Nym.name = txtName Then
    Nym.ChangeName = False
    txtName.ForeColor = vbBlack
    lblnamestatus = ""
Else
    If Not Nym.create Then
        txtName.ForeColor = vbRed
        Nym.ChangeName = True
        lblnamestatus = "Note: You appear to be changing your reply block and your Alias name.  This is okay provided you have checked to see that the nym does not already exist.  If it does you well have to resend the change request."
    End If
End If
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    Nym.ChangeName = True
End Sub

Private Sub txtNewsGroupEmail_Change(Index As Integer)
Nym.ChangeName = False
txtNewsGroupEmail(Index).ForeColor = vbBlack
lblnamestatus = ""
Select Case Index
    Case 0
        If Nym.NewsGroupReplyEmail = txtNewsGroupEmail(Index) Then Exit Sub
    Case 1
        If Nym.NewsGroupReplyGroup = txtNewsGroupEmail(Index) Then Exit Sub
    Case 2
        If Nym.NewsGroupReplySubject = txtNewsGroupEmail(Index) Then Exit Sub
End Select
'If Nym.NewsGroupReplyEmail = txtNewsGroupEmail(Index) Then
  '  Nym.ChangeName = False
   ' txtNewsGroupEmail(Index).ForeColor = vbBlack
   ' lblnamestatus = ""
'End If
If Not Nym.create Then
    txtNewsGroupEmail(Index).ForeColor = vbRed
    Nym.ChangeName = True
    lblnamestatus = "Note: You appear to be changing your reply block and your Alias name.  This is okay provided you have checked to see that the nym does not already exist.  If it does you well have to resend the change request."
End If

End Sub

Private Sub txtPassPhrase_Change(Index As Integer)
If Index > 0 Then lblWarning.Visible = True
End Sub
