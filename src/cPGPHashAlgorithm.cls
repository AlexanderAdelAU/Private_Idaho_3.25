VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPGPHashAlgorithm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const m_PGPHashAlgorithm_Default = 0
Private Const m_PGPHashAlgorithm_MD5 = 1
Private Const m_PGPHashAlgorithm_SHA = 2
Private Const m_PGPHashAlgorithm_RIPEMD160 = 3
Private Const m_PGPHashAlgorithm_SHADouble = 4 ' not available in PGP 6
Public Property Get PGPPublicKeyAlgorithm_MD5() As Integer
PGPPublicKeyAlgorithm_MD5 = m_PGPHashAlgorithm_MD5
End Property

Public Property Get PGPHashAlgorithm_SHA() As Integer
PGPHashAlgorithm_SHA = m_PGPHashAlgorithm_SHA
End Property

Public Property Get PGPPublicKeyAlgorithm_Default() As Integer
PGPPublicKeyAlgorithm_Default = m_PGPHashAlgorithm_MD5
End Property

Public Property Get PGPHashAlgorithm_RIPEMD160() As Integer
PGPHashAlgorithm_RIPEMD160 = m_PGPHashAlgorithm_RIPEMD160
End Property
