VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBCbLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public docLogType      As Integer
Public log             As String
Public procType        As String
Public procMemo        As String
Public regDT           As String
Public ip              As String

Public Sub fromDictionary(dic As Variant)

    docLogType = dic.Item("docLogType")
    log = dic.Item("log")
    procType = dic.Item("procType")
    procMemo = dic.Item("procMemo")
    
    regDT = dic.Item("regDT")
    ip = dic.Item("ip")

End Sub

