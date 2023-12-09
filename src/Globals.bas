Attribute VB_Name = "Globals"
Option Explicit
'Globals
Global Const MAX_NUMBER_OF_PI_INSTANCES As Integer = 20
Global gFormInstance As Integer
Global PIForm(MAX_NUMBER_OF_PI_INSTANCES) As frmPI ' Allow for up to 10 instances
Global gActivePIInstance As Integer

'Memory Allocation Data
Type MemoryBuffer
    Address As Long
    StartAddr As Long
    BufferSize As Long
    hMem As Long
End Type
Global gRemailerSelectCaption As String ' Use by frmMultiSelect Nyms as the caption
Global gBuffer As MemoryBuffer
Global Const INVALID As Integer = -9999
'Used for the AddressList
Global Const CONTACT_ALL_LIST = 0
Global Const CONTACT_TO_LIST = 1
Global Const CONTACT_CC_LIST = 2
Global Const CONTACT_NEW = 0

'Descriptors for selected addressees
Global Const CONTACT_IN_DB = 1
Global Const CONTACT_ON_PGPKEYRING = 2
Global Const CONTACT_IN_MAILGROUP = 3
'Values 3 to 10 are reserved for groups and are dynamic
Global Const CONTACT_ARRAY_LIMIT = 4


Global gFullRelease As Integer
'Project Level Constant Declaration Section

Global Const Cypherpunk = 1
Global Const Soda = 2
Global Const penet = 3
Global Const USENET = 4
Global Const mix = 5
Global Const NYMKEY = 2
Global Const IDKEY = 1

'Email Scan Interval
Global giEmailScanInterval As Integer
Global giTimerCounter As Integer

'Email Scan Options
Global Const SCAN_PGP_ONLY As Integer = 0
Global Const SCAN_PLAIN_ONLY As Integer = 3
Global Const SCAN_ALL As Integer = 1
Global Const SCAN_NONE As Integer = 2

Global giEmailScanOption As Integer

'Folder constants and IDs

'Global Const gInFolderID As Long = 3
'Global Const gDeletedFolderID As Long = 2
'Global gFolderID As Long

'Remailer Globals
Type RemailerType
    ShortName As String
    name As String
    Address As String
    history As String
    latency As String
    uptime As String
    cpunk As Integer
    eric As Integer
    mix As Integer
    penet As Integer
    alpha As Integer
    newnym As Integer
    Encrypt As Integer
    post As Integer
    latent As Integer
    hash As Integer
    cut As Integer
    Reserved1 As Integer
    Reserved2 As Integer
End Type
'Public Nym.latenttime As String

Type NymInfo
    acksend As Boolean
    fixedsize As Boolean
    disable As Boolean
    fingerkey As Boolean
    LatentTime As String
    name As String
    signsend As Boolean
    cryptrecv As Boolean
    ID As String
    EmailAddress As String
    Server As String
    ChangeName As Boolean
    PassPhrase(0 To 5) As String
    UseNewsGroupReply As Boolean
    NewsGroupReplyEmail As String
    NewsGroupReplyGroup As String
    NewsGroupReplySubject As String
    create As Boolean
    ListIndex As Integer 'this is the selected nym from the multinyms list
    NymState As Integer ' this is the nyms status, ie create, delete, idle, replychange etc
    DontUseRemailer As Boolean
End Type
Global Nym As NymInfo

Type EMailerInfo
    ButtonName As String
    MailName As String
    Script As String
End Type

Global gRemailerArray(512) As RemailerType
Global gTotalRemailers As Integer
Global gMatchedRemailers(512) As RemailerType
Global gTotalMatchedRemailers As Integer
Global gSortRemailer As RemailerType
Global gShowNymStatus As Integer
Global gIsNewNym As Boolean
Global gNymState As Integer


'PGP Response Strings
Type PGPResponse
    Count As Integer
    res(3) As String
End Type
'Global gPGPResponse As PGPResponse

Global Const gNYM_IDLE = 0
Global Const gNYMDEL = 2
Global Const gNYMRPLYCHANGE = 3
Global Const gNYMPREPARE = 4
Global Const gNYMCONFIG = 5
Global Const gNYM_DECRYPT = 6
Global Const gNYM_USENET_PREPARE = 7


'Project Level Variable Declaration Section


Global gnumRemailers As Integer
Global numInfo As Integer
'Global Remailers(50) As String * 30
Global Remailers(50) As String
Global RemailInfo(50) As String
Global gwhichRemailer As Integer
Global gPGPPath As String
Global gMixPath As String
Global gPGPFile As String
Global gPGPKeyID As String
Global gKeyID As String
'Global gEncryptToRemailer As Boolean
Global gManualEncrypt As Boolean
'---------------------------------------------
'new 1.5 declarations
'---------------------------------------------
Global gEmailer As String
Global gPGPTempFile As String
Global gMailerInfo(12) As EMailerInfo
Global gRemailerType As Integer
Public Const STANDARD_EMAIL As Integer = 0
Public Const REMAILER_CYPHERPUNK As Integer = 1
Public Const REMAILER_MIX As Integer = 2
Public Const ENCRYPT_BEFORE_SENDING_MESSAGE As Integer = 3
Public Const ENCRYPT_AND_SIGN_BEFORE_SENDING_MESSAGE As Integer = 4
Public Const SIGN_BEFORE_SENDING_MESSAGE  As Integer = 5
Public Const SEND_MESSAGES_USING_NYM As Integer = 6

Global gRemailerTypeURL As Integer
Global gNewsgroupType As Integer
'Global gUSENETMailer As String
Global gObscurity As Integer
'Global gPIPIF As String
Global gMinState As Integer
Global gLatentStr As String
Global gSubStr As String
Global gCutStr As String
'Global gAppINI As String
'Global ViaCrypt As String
'Global gOS2 As String
Global gBrowserPath As String
Global gBrowserString As String
Global gURLStart As String
Global gURLEnd As String
Global gCRLF As String
Global gHeader As String
Global gSig As String

Global gMultiType As Integer
Global gtranScript As String
Global gc2WWWAnon As String
Global gwrapNum As Integer

Global gPiStr As String
Global gPassPhrase As String

Global gExit As Integer
Global gDebugString As String
'---------------------------------------------
'new v2.8c declarations
'---------------------------------------------
Global gCancelAction As Boolean
Global gMixSent As Boolean

'*********POP Globals ***********************
Type MailMessage
    PGP As Integer
    zap As Integer
    Read As Integer
    SentDate As String
    To As String
    CC As String
    ReplyTo As String
    From As String
    Subject As String
    MessageID As String
    Received As String
    Header As String
    Contents As String
    EndTransfer As Boolean
    Attachment As Boolean
    MessageSize As Long
    MessageNumber As Long
    ReturnPath As String
End Type

Type XHeader
    ID As String
    Value As String
End Type


Type MailIndex
    RecordNumber As Integer
    FilePos As Long
End Type
Global gMessageRecord As MailMessage
Global MasterIndex As MailIndex

'Mailheader(0) holds the combined header for email
'Mailheader(1) to Mailheader(5) holds the NewsGroup posting headers
'Mailheader(6) to Mailheader(9) holds the Email Extra Headers
Global MailHeader(9) As XHeader

Public Const CONNECT_POP3 As Integer = 0
Public Const CONNECT_IMAP4 As Integer = 1

Type MailConnectType
    EmailAddress As String
    RealName As String
    SMTPServerName As String
    NNTPServerName As String
    ReplyEmailAddress As String
    POPPort As Integer
    SMTPPort As Integer
    ServerConnected As Boolean
    ServerState As Integer
    ConnectUsing As Integer
    MailServerName As String
    AccountName As String
    NumMessages As Integer
    AccountPassword As String
    DNSServerName As String
    AuthenticationRequired As Boolean
End Type
Public MailConnector As MailConnectType
Public gMessage As String

'Used when displaying messages rather than composing
Global gComposeMode As Boolean


Global gFoundMessages As String
Global gMessagesToBeDeleted As String
Global Const POPDECRYPT = 1
'Governs out,in and deleted presentation
Global giMailbox As Integer
Global Const SHOW_IN_MAILBOX As Integer = 0
Global Const SHOW_OUT_MAILBOX As Integer = 1
Global Const SHOW_DELETED_MAILBOX As Integer = 2


Global DB As Database

'*********HTTP Globals *********
Global gWebPage As String
Global gWebState As Integer
Global gRemailerInfoURL As String
Global gMixListURL As String
Global gMixType2URL As String
Global gMixPubRingURL As String
Global gPGPKeysURL As String
Global gSubKeyURL As String
Global gGetKeyURL As String


Global Const SMTPSTATE = 1
Global Const POPSTATE = 2
Global Const HTTPSTATE = 3
Global Const IMAPSTATE = 4

Global Const HTTPIDLE = 0
Global Const GETREMAILERUPDATE = 1
Global Const GETSERVERKEY = 2
Global Const GETREMAILERKEYS = 3
Global Const MIXUPDATE = 4
Global Const PUBRINGUPDATE = 5
Global Const TYPE2UPDATE = 6

'--PGP Variables
Public gPGPVersion As String
Public Const NoPGP = "0"
Public Const PGP26x = "1"
Public Const PGP5x = "2"
Public Const PGPNotFound = "3"

Public vb2spgpContext As New spgpContext
'Public MessageType As New cMessageType
Public RemailerContext As New cRemailer

'Used in frmMain
'Public Tree(2) As New cTree
Public Grid As New cGrid

'==================
'Signature and Key Data
'==================
Public SignatureProperties As TSig_Data
