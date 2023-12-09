VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPGPCipherAlgorithm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const m_PGPCipherAlgorithm_IDEA = 1
Private Const m_PGPCipherAlgorithm_3DES = 2
Private Const m_PGPCipherAlgorithm_CAST5 = 3

Public Property Get PGPCipherAlgorithm_IDEA() As Variant
PGPCipherAlgorithm_IDEA = m_PGPCipherAlgorithm_IDEA
End Property


Public Property Get PGPCipherAlgorithm_3DES() As Variant
PGPCipherAlgorithm_3DES = m_PGPCipherAlgorithm_3DES
End Property
Public Property Get PGPCipherAlgorithm_CAST5() As Variant
PGPCipherAlgorithm_CAST5 = m_PGPCipherAlgorithm_CAST5
End Property
Public Property Get PGPCipherAlgorithm_Default() As Variant
PGPCipherAlgorithm_Default = m_PGPCipherAlgorithm_IDEA
End Property
