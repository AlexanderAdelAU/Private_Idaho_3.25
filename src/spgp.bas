Attribute VB_Name = "spgp"
'SPGP.BAS
'
'version 2.5.0.0, 28 June 2000
'added spgpGetPreferences, spgpSetPreferences & related constants
'added spgpKeyPropADK & related constants
'added spgpUISelectKeysDialog, spgpPreferencesDialog
'added spgpKeyPropRevocationKey
'--------------------------------------------------------
'version 2.4.1.0, 29 February 2000
'changed spgpSubKeyGenerate parameters, changed alias names for
'spgpEncode, spgpEncodeFile, spgpKeyImport, spgpKeyImportFile
'--------------------------------------------------------
'version 2.4.0.0, 24 February 2000
'added functions spgpKeyGenerate, spgpSubKeyGenerate, spgpSdkApiVersion
'--------------------------------------------------------
'version 2.3.0.0, 31 December 99
'added functions spgpUIEncode & spgpUIEncodefile
'added function spgpEstimatePassphraseQuality
'added UI dialog functions, ordinal values
'--------------------------------------------------------
'version 2.2.3.3, 14 November 99
'added function spgpKeyRemove
'--------------------------------------------------------
'version 2.2.3.0, 24 Septempber 99
'added functions spgp_KeyImport, spgp_KeyImportFile spgpKeyPassChange,
'spgpKeyEnable, spgpKeyDisable, spgp_Encode, spgp_EncodeFile
'--------------------------------------------------------
'version 2.2, 14 July 99
'added functions spgpKeySign, spgpKeySigRemove, spgpKeyPropUserID,
'spgpKeyPropSig, spgpKeyRingToFile
'--------------------------------------------------------
'version 2.1.1.0, 7 June 99
'added functions spgpDetachedSigCreate & spgpDetachedSigVerify
'added function spgpKeyRingCount
'added default hash for signing, and related global constant
'--------------------------------------------------------
'version 2.1.0.0, 1 June 99
'added function spgpVersion
'added functions spgpAnalyzeEx, spgpAnalyzeFileEx & related Constants
'fixed keyimportfile, again
'pursuant to the above, added 'ByRef' to As Long params where needed
'introduced confusing new version numbering, just what we need
'--------------------------------------------------------
'version 2.0e, 7 Mar 99
'added Preferences functions and constants
'--------------------------------------------------------
'version 2.0, 21 Feb 99
'now supports PGP 6.x
'added ciphering algorithms
'fixed keyimportfile broken in 1.1
'abbreviated keyprops in keyimport functions
'changed some key/signature properties to global constants
'keyexport functions can use compatible-format mode
'--------------------------------------------------------
'version 1.1, 4 Feb 99
'fixed armor/clearsigning problem
'primary Key ID no longer required
'added hashing algorithms
'added spgpAnalyzeFile()
'--------------------------------------------------------
'
' note that while some of these constants have the same name as
' variables found in the pgp *.h files, they do not all have the same
' values -- these are not meant to be translations of those files!

'public key algorithms
Global Const PGPPublicKeyAlgorithm_Invalid = 0 'this (0) may not be correct, but how to test it? anyone got an invalid algorithm laying around?
Global Const PGPPublicKeyAlgorithm_RSA = 1
Global Const PGPPublicKeyAlgorithm_RSAEncryptOnly = 2
Global Const PGPPublicKeyAlgorithm_RSASignOnly = 3
Global Const PGPPublicKeyAlgorithm_ElGamal = 4 'A.K.A. Diffie-Hellman
Global Const PGPPublicKeyAlgorithm_DSA = 5

'symmetric ciphers
Global Const PGPCipherAlgorithm_IDEA = 1
Global Const PGPCipherAlgorithm_3DES = 2
Global Const PGPCipherAlgorithm_CAST5 = 3

'hashing algorithms
Global Const PGPHashAlgorithm_Default = 0
Global Const PGPHashAlgorithm_MD5 = 1
Global Const PGPHashAlgorithm_SHA = 2
Global Const PGPHashAlgorithm_RIPEMD160 = 3
Global Const PGPHashAlgorithm_SHADouble = 4 ' not available in PGP 6

'trust levels
Global Const PGPKeyTrust_Undefined = 0
Global Const PGPKeyTrust_Unknown = 1
Global Const PGPKeyTrust_Never = 2
Global Const PGPKeyTrust_Marginal = 5
Global Const PGPKeyTrust_Complete = 6
Global Const PGPKeyTrust_Ultimate = 7

'validity levels
Global Const PGPValidity_Unknown = 0
Global Const PGPValidity_Invalid = 1
Global Const PGPValidity_Marginal = 2
Global Const PGPValidity_Complete = 3

'analysis results
Global Const PGPAnalyze_Encrypted = 0            ' Encrypted message
Global Const PGPAnalyze_Signed = 1               ' Signed message
Global Const PGPAnalyze_DetachedSignature = 2    ' Detached signature
Global Const PGPAnalyze_Key = 3                  ' Key data
Global Const PGPAnalyze_Unknown = 4              ' Non-pgp message
' these are for spgpAnalyzeEx only
Global Const PGPAnalyze_EncryptedConventional = 5 ' Like it says
Global Const PGPAnalyze_EncryptedNoKeys = 6      ' Key-encrypted to keys not on local ring

'signature status
Global Const SIGNED_GOOD = 0
Global Const SIGNED_NOT = 1
Global Const SIGNED_BAD = 2
Global Const SIGNED_NO_KEY = 3

'preferences, old style
Global Const PGPPref_PublicKeyring = 0
Global Const PGPPref_PrivateKeyring = 1
Global Const PGPPref_RandomSeedFile = 2
Global Const PGPPref_DefaultKeyID = 3

'----------------------------------------------------------------------------------------------------
' preferences, jeune ecole
Public Const spgpPref_PublicKeyring As Long = 1
Public Const spgpPref_PrivateKeyring As Long = 2
Public Const spgpPref_RandomSeedFile As Long = 4
Public Const spgpPref_DefaultKeyID As Long = 8
Public Const spgpPref_GroupsFile As Long = 16

' for spgpPreferencesDialog, these select which page to open
Public Const spgpPrefsPage_GeneralPrefs = 0
Public Const spgpPrefsPage_KeyringPrefs = 1
Public Const spgpPrefsPage_EmailPrefs = 2
Public Const spgpPrefsPage_HotkeyPrefs = 3
Public Const spgpPrefsPage_KeyserverPrefs = 4
Public Const spgpPrefsPage_CAPrefs = 5
Public Const spgpPrefsPage_AdvancedPrefs = 6
    
' Key Properties Flags
  ' string properties
Public Const spgpKeyProp_KeyID            As Long = &H1       ' 1
Public Const spgpKeyProp_UserID           As Long = &H2       ' 2
Public Const spgpKeyProp_Fingerprint      As Long = &H4       ' 4
Public Const spgpKeyProp_CreationTimeStr  As Long = &H8       ' 8
Public Const spgpKeyProp_ExpirationTimeStr As Long = &H10     ' 16
  ' numeric properties
Public Const spgpKeyProp_Keybits          As Long = &H80      ' 128
Public Const spgpKeyProp_KeyAlg           As Long = &H100     ' 256
Public Const spgpKeyProp_Trust            As Long = &H200     ' 512
Public Const spgpKeyProp_Validity         As Long = &H400     ' 1024
Public Const spgpKeyProp_CreationTime     As Long = &H800     ' 2048
Public Const spgpKeyProp_ExpirationTime   As Long = &H1000    ' 4096
  ' boolean properties
Public Const spgpKeyProp_IsSecret         As Long = &H8000    ' & cetera
Public Const spgpKeyProp_IsAxiomatic      As Long = &H10000
Public Const spgpKeyProp_IsRevoked        As Long = &H20000
Public Const spgpKeyProp_IsDisabled       As Long = &H40000
Public Const spgpKeyProp_IsExpired        As Long = &H80000
Public Const spgpKeyProp_IsSecretShared   As Long = &H100000
Public Const spgpKeyProp_CanEncrypt       As Long = &H200000
Public Const spgpKeyProp_CanDecrypt       As Long = &H400000
Public Const spgpKeyProp_CanSign          As Long = &H800000
Public Const spgpKeyProp_CanVerify        As Long = &H1000000
Public Const spgpKeyProp_HasRevoker       As Long = &H2000000
Public Const spgpKeyProp_HasADK           As Long = &H4000000
Public Const spgpKeyProp_HasSubKey        As Long = &H8000000

Public Const MAX_PATH As Long = 260

Public Type spgpPreferenceRec
  PublicKeyring As String * MAX_PATH
  PrivateKeyring As String * MAX_PATH
  RandomSeedFile As String * MAX_PATH
  GroupsFile As String * MAX_PATH
  DefaultKeyID As String * 10
End Type

'----------------------------------------------------------------------------------------------------
' function names exported from the dll are case-sensitive.
' encrypt/decrypt
Declare Function spgpEncode Lib "spgp.dll" Alias "spgp_encode" (ByVal BufferIn As String, ByVal BufferOut As String, ByVal BufferOutLen As Long, ByVal Encrypt As Long, ByVal Sign As Long, ByVal SignAlg As Long, ByVal Conventional As Long, ByVal ConventionalAlg As Long, ByVal Armor As Long, ByVal TextMode As Long, ByVal Clear As Long, ByVal Compress As Long, ByVal EyesOnly As Long, ByVal MIME As Long, ByVal CryptKeyID As String, ByVal SignKeyID As String, ByVal SignKeyPass As String, ByVal ConventionalPass As String, ByVal Comment As String, ByVal MIMESeparator As String) As Long
Declare Function spgpEncodeFile Lib "spgp.dll" Alias "spgp_encodefile" (ByVal FileIn As String, ByVal FileOut As String, ByVal Encrypt As Long, ByVal Sign As Long, ByVal SignAlg As Long, ByVal Conventional As Long, ByVal ConventionalAlg As Long, ByVal Armor As Long, ByVal TextMode As Long, ByVal Clear As Long, ByVal Compress As Long, ByVal EyesOnly As Long, ByVal MIME As Long, ByVal CryptKeyID As String, ByVal SignKeyID As String, ByVal SignKeyPass As String, ByVal ConventionalPass As String, ByVal Comment As String, ByVal MIMESeparator As String) As Long
' old versions
' Declare Function spgpEncode Lib "spgp.dll" Alias "spgpencode" (ByVal BufferIn As String, ByVal BufferOut As String, ByVal BufferOutLen As Long, ByVal Encrypt As Long, ByVal Sign As Long, ByVal SignAlg As Long, ByVal Conventional As Long, ByVal ConventionalAlg As Long, ByVal Armor As Long, ByVal TextMode As Long, ByVal Clear As Long, ByVal Compress As Long, ByVal EyesOnly As Long, ByVal MIME As Long, ByVal CryptKeyID As String, ByVal SignKeyID As String, ByVal SignKeyPass As String, ByVal ConventionalPass As String, ByVal Comment As String, ByVal MIMESeparator As String) As Long
' Declare Function spgpEncodeFile Lib "spgp.dll" Alias "spgpencodefile" (ByVal FileIn As String, ByVal FileOut As String, ByVal Encrypt As Long, ByVal Sign As Long, ByVal SignAlg As Long, ByVal Conventional As Long, ByVal ConventionalAlg As Long, ByVal Armor As Long, ByVal TextMode As Long, ByVal Clear As Long, ByVal Compress As Long, ByVal EyesOnly As Long, ByVal MIME As Long, ByVal CryptKeyID As String, ByVal SignKeyID As String, ByVal SignKeyPass As String, ByVal ConventionalPass As String, ByVal Comment As String, ByVal MIMESeparator As String) As Long
Declare Function spgpDecode Lib "spgp.dll" Alias "spgpdecode" (ByVal BufferIn As String, ByVal BufferOut As String, ByVal BufferOutLen As Long, ByVal Pass As String, ByVal SigProps As String) As Long
Declare Function spgpDecodeFile Lib "spgp.dll" Alias "spgpdecodefile" (ByVal FileIn As String, ByVal FileOut As String, ByVal Pass As String, ByVal SigProps As String) As Long

' key import/export
Declare Function spgpKeyExport Lib "spgp.dll" Alias "spgpkeyexport" (ByVal KeyID As String, ByVal BufferOut As String, ByVal BufferOutLen As Long, ByVal ExportPrivate As Long, ByVal ExportCompatible As Long) As Long
Declare Function spgpKeyExportFile Lib "spgp.dll" Alias "spgpkeyexportfile" (ByVal KeyID As String, ByVal FileOut As String, ByVal ExportPrivate As Long, ByVal ExportCompatible As Long) As Long
Declare Function spgpKeyImport Lib "spgp.dll" Alias "spgp_keyimport" (ByVal BufferIn As String, ByVal KeyProps As String, ByVal KeyPropsLen As Long, ByVal Import As Long, ByVal AllProps As Long) As Long
Declare Function spgpKeyImportFile Lib "spgp.dll" Alias "spgp_keyimportfile" (ByVal FileIn As String, ByVal KeyProps As String, ByVal KeyPropsLen As Long, ByVal Import As Long, ByVal AllProps As Long) As Long
' old versions
' Declare Function spgpKeyImport Lib "spgp.dll" Alias "spgpkeyimport" (ByVal BufferIn As String, ByVal KeyProps As String, ByVal KeyPropsLen As Long) As Long
' Declare Function spgpKeyImportFile Lib "spgp.dll" Alias "spgpkeyimportfile" (ByVal FileIn As String, ByVal KeyProps As String, ByVal KeyPropsLen As Long) As Long

' key properties
Declare Function spgpKeyProps Lib "spgp.dll" Alias "spgpkeyprops" (ByVal KeyID As String, ByVal KeyProps As String, ByVal KeyPropsLen As Long) As Long
Declare Function spgpKeyRingID Lib "spgp.dll" Alias "spgpkeyringid" (ByVal BufferOut As String, ByVal BufferOutLen As Long) As Long
Declare Function spgpKeyRingCount Lib "spgp.dll" Alias "spgpkeyringcount" () As Long
Declare Function spgpKeyRingToFile Lib "spgp.dll" Alias "spgpkeyringtofile" (ByVal FileOut As String) As Long
Declare Function spgpKeyPropUserID Lib "spgp.dll" Alias "spgpkeypropuserid" (ByVal KeyID As String, ByVal BufferOut As String, ByVal BufferOutLen As Long) As Long
Declare Function spgpKeyPropSig Lib "spgp.dll" Alias "spgpkeypropsig" (ByVal UserID As String, ByVal BufferOut As String, ByVal BufferOutLen As Long) As Long
Declare Function spgpKeyPropADK Lib "spgp.dll" Alias "spgpkeypropadk" (ByVal KeyHexID As String, ByVal ADKeyProps As String, ByVal ADKeyPropsLen As Long, ADKeyCount As Long, ByVal Flags As Long) As Long
Declare Function spgpKeyPropRevocationKey Lib "spgp.dll" Alias "spgpkeyproprevocationkey" (ByVal KeyHexID As String, ByVal RevKeyProps As String, ByVal RevKeyPropsLen As Long, RevKeyCount As Long, ByVal Flags As Long) As Long

' error strings
Declare Function spgpGetErrorString Lib "spgp.dll" Alias "spgpgeterrorstring" (ByVal theError As Long, ByVal BufferOut As String) As Long

' analyze
Declare Function spgpAnalyze Lib "spgp.dll" Alias "spgpanalyze" (ByVal BufferIn As String) As Long
Declare Function spgpAnalyzeFile Lib "spgp.dll" Alias "spgpanalyzefile" (ByVal FileIn As String) As Long
Declare Function spgpAnalyzeEx Lib "spgp.dll" Alias "spgpanalyzeex" (ByVal BufferIn As String, ByVal BufferOut As String, ByVal BufferOutLen As Long) As Long
Declare Function spgpAnalyzeFileEx Lib "spgp.dll" Alias "spgpanalyzefileex" (ByVal FileIn As String, ByVal BufferOut As String, ByVal BufferOutLen As Long) As Long

' prefs
Declare Function spgpSetPreference Lib "spgp.dll" Alias "spgpsetpreference" (ByVal Preference As Long, ByVal BufferIn As String) As Long
Declare Function spgpGetPreference Lib "spgp.dll" Alias "spgpgetpreference" (ByVal Preference As Long, ByVal BufferOut As String) As Long
Declare Function spgpSetPreferences Lib "spgp.dll" Alias "spgpsetpreferences" (Prefs As spgpPreferenceRec, ByVal Flags As Long) As Long
Declare Function spgpGetPreferences Lib "spgp.dll" Alias "spgpgetpreferences" (Prefs As spgpPreferenceRec, ByVal Flags As Long) As Long
Declare Function spgpPreferencesDialog Lib "spgp.dll" Alias "spgppreferencesdialog" (ByVal ShowPage As Long, ByVal WindowHandle As Long) As Long

' misc
Declare Function spgpKeyIsOnRing Lib "spgp.dll" Alias "spgpkeyisonring" (ByVal KeyID As String) As Long
Declare Function spgpVersion Lib "spgp.dll" Alias "spgpversion" () As Long
Declare Function spgpSdkApiVersion Lib "spgp.dll" Alias "spgpsdkapiversion" () As Long
Declare Function spgpEstimatePassphraseQuality Lib "spgp.dll" Alias "spgpestimatepassphrasequality" (ByVal PassPhrase As String) As Long
'Declare Function spgpPGPPath Lib "spgp.dll" Alias "spgppgppath" (ByVal Path As String) As Long

' key manipulation
Declare Function spgpKeyChange Lib "spgp.dll" Alias "spgpkeypasschange" (ByVal KeyID As String, ByVal OldPhrase As String, ByVal NewPhrase As String) As Long
Declare Function spgpKeyEnable Lib "spgp.dll" Alias "spgpkeyenable" (ByVal KeyID As String) As Long
Declare Function spgpKeyDisable Lib "spgp.dll" Alias "spgpkeydisable" (ByVal KeyID As String) As Long
Declare Function spgpKeyRemove Lib "spgp.dll" Alias "spgpkeyremove" (ByVal KeyID As String) As Long
Declare Function spgpKeySigRemove Lib "spgp.dll" Alias "spgpkeysigremove" (ByVal KeyHexID As String, ByVal UserID As String, ByVal SignHexID As String) As Long
Declare Function spgpKeySign Lib "spgp.dll" Alias "spgpkeysign" (ByVal KeyHexID As String, ByVal UserID As String, ByVal SignKeyID As String, ByVal SignKeyPass As String, ByVal Expires As Long, ByVal Exportable As Long, ByVal Trust As Long, ByVal Validity As Long) As Long
Declare Function spgpKeyGenerate Lib "spgp.dll" Alias "spgpkeygenerate" (ByVal UserID As String, ByVal PassPhrase As String, ByVal NewKeyHexID As String, ByVal KeyAlg As Long, ByVal CipherAlg As Long, ByVal Size As Long, ByVal Expires As Long, ByVal FastGeneration As Long, ByVal FailWithoutEntropy As Long, ByVal WinHandle As Long) As Long
Declare Function spgpSubKeyGenerate Lib "spgp.dll" Alias "spgpsubkeygenerate" (ByVal MasterKeyHexID As String, ByVal MasterKeyPass As String, ByVal NewSubKeyHexID As String, ByVal KeyAlg As Long, ByVal Size As Long, ByVal ExpiresIn As Long, ByVal FastGeneration As Long, ByVal FailWithoutEntropy As Long, ByVal WinHandle As Long) As Long

' signatures
Declare Function spgpDetachedSigCreate Lib "spgp.dll" Alias "spgpdetachedsigcreate" (ByVal FileIn As String, ByVal SigFile As String, ByVal SignKeyID As String, ByVal SignKeyPass As String, ByVal Comment As String, ByVal SignAlg As Long, ByVal Armor As Long) As Long
Declare Function spgpDetachedSigVerify Lib "spgp.dll" Alias "spgpdetachedsigverify" (ByVal SigFile As String, ByVal SignedFile As String, ByVal SigProps As String) As Long

' User Interface
Declare Function spgpUIEncode Lib "spgp.dll" Alias "spgpuiencode" (ByVal BufferIn As String, ByVal BufferOut As String, ByVal BufferOutLen As Long, ByVal Encrypt As Long, ByVal Sign As Long, ByVal SignAlg As Long, ByVal Conventional As Long, ByVal ConventionalAlg As Long, ByVal Clear As Long, ByVal Compress As Long, ByVal EyesOnly As Long, ByVal MIME As Long, ByVal Comment As String, ByVal MIMESeparator As String, ByVal WindowHandle As Long) As Long
Declare Function spgpUIEncodeFile Lib "spgp.dll" Alias "spgpuiencodefile" (ByVal FileIn As String, ByVal FileOut As String, ByVal Encrypt As Long, ByVal Sign As Long, ByVal SignAlg As Long, ByVal ConventionalEncrypt As Long, ByVal ConventionalAlg As Long, ByVal Armor As Long, ByVal TextMode As Long, ByVal Clear As Long, ByVal Compress As Long, ByVal EyesOnly As Long, ByVal MIME As Long, ByVal Comment As String, ByVal MIMESeparator As String, ByVal WindowHandle As Long) As Long

Declare Function spgpUIRecipientsDialog Lib "spgp.dll" Alias "spgpuirecipientsdialog" (ByVal RecipientHexID As String, ByVal RecipientHexIDLen As Long, ByVal Caption As String, ByVal Reserved1 As String, ByVal DisplayMarginalValidity As Long, ByVal Reserved2 As Long, ByVal Reserved3 As Long, ByVal Reserved4 As Long, ByVal Reserved5 As Long, ByVal WindowHandle As Long) As Long
Declare Function spgpUISigningPassphraseDialog Lib "spgp.dll" Alias "spgpuisigningpassphrasedialog" (ByVal SelectedKeyHexID As String, ByVal SelectedKeyPass As String, ByVal DefaultKeyID As String, ByVal FindMatchingKey As Long, ByVal WindowHandle As Long) As Long
Declare Function spgpUIConfirmationPassphraseDialog Lib "spgp.dll" Alias "spgpuiconfirmationpassphrasedialog" (ByVal PassPhrase As String, ByVal ShowPassphraseQuality As Long, ByVal MinimumPassphraseQuality As Long, ByVal MinimumPassphraseLength As Long, ByVal WindowHandle As Long) As Long
Declare Function spgpUIKeyPassphraseDialog Lib "spgp.dll" Alias "spgpuikeypassphrasedialog" (ByVal KeyID As String, ByVal PassPhrase As String, ByVal WindowHandle As Long) As Long
Declare Function spgpUISelectKeysDialog Lib "spgp.dll" Alias "spgpuiselectkeysdialog" (ByVal KeyID As String, ByVal KeyProps As String, ByVal Prompt As String, ByVal KeyPropsLen As Long, ByVal ShowKeyRing As Long, ByVal Flags As Long, ByVal WinHandle As Long) As Long

'----------------------------------------------------------------------
' Function ordinal values, for those who would rather import by ordinal
'----------------------------------------------------------------------

'  Function Name      Ordinal

'  spgpencode             1
'  spgpencodefile         2
'  spgp_encode            3
'  spgp_encodefile        4

'  spgpdecode             5
'  spgpdecodefile         6

'  spgpkeyexport          7
'  spgpkeyexportfile      8
'  spgpkeyimport          9
'  spgp_keyimport         10
'  spgpkeyimportfile      11
'  spgp_keyimportfile     12

'  spgpkeyprops           13
'  spgpkeyringid          14
'  spgpkeyringtofile      15

'  spgpgeterrorstring     16

'  spgpanalyze            17
'  spgpanalyzeex          18
'  spgpanalyzefile        19
'  spgpanalyzefileex      20

'  spgpkeyisonring        21
'  spgpsetpreference      22
'  spgpgetpreference      23
'  spgpversion            24
'  spgpdetachedsigcreate  25
'  spgpdetachedsigverify  26
'  spgpkeyringcount       27
'  spgpkeyenable          28
'  spgpkeydisable         29
'  spgpkeypasschange      30
'  spgpkeysign            31
'  spgpkeypropuserid      32
'  spgpkeypropsig         33
'  spgpkeyremove          34
'  spgpkeysigremove       35
  
'  spgpuiencode           36
'  spgpuiencodefile       37
'  spgpuirecipientsdialog 38
'  spgpuisigningpassphrasedialog      39
'  spgpuiconfirmationpassphrasedialog 40
'  spgpuikeypassphrasedialog          41
'  spgpestimatepassphrasequality      42

'  spgpkeygenerate        43
'  spgpsubkeygenerate     44

'  spgpsdkapiversion      45

'  spgpsetpreferences     46
'  spgpgetpreferences     47
'  spgppreferencesdialog  48
  
'  spgpuiselectkeysdialog 49
'  spgpkeypropadk         50
'  spgpkeyproprevocationkey 51

