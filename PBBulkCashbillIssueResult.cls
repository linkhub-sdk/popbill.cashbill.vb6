VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBBulkCashbillIssueResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public mgtKey           As String
Public code             As Long
Public message          As String
Public confirmNum       As String
Public tradeDate        As String
Public tradeDT        As String

Public Sub fromDictinary(dic As Variant)
    mgtKey = dic.Item("mgtKey")
    code = dic.Item("code")
    message = dic.Item("message")
    confirmNum = dic.Item("confirmNum")
    tradeDate = dic.Item("tradeDate")
    tradeDT = dic.Item("tradeDT")
End Sub


