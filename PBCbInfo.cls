VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBCbInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public itemKey As String
Public mgtKey As String
Public tradeDate As String
Public issueDT As String
Public customerName As String
Public itemName As String
Public identityNum As String
Public taxationType As String

Public totalAmount As String
Public tradeUsage As String
Public tradeOpt As String
Public tradeType As String
Public stateMemo As String
Public stateCode As Integer
Public stateDT As String
Public printYN As Boolean

Public confirmNum As String
Public orgTradeDate As String
Public orgConfirmNum As String

Public ntssendDT As String
Public ntsresult As String
Public ntsresultDT As String
Public ntsresultCode As String
Public ntsresultMessage As String

Public regDT As String

Public Sub fromDictionary(dic As Variant)
    itemKey = dic.Item("itemKey")
    mgtKey = dic.Item("mgtKey")
    tradeDate = dic.Item("tradeDate")
    issueDT = dic.Item("issueDT")
    customerName = dic.Item("customerName")
    itemName = dic.Item("itemName")
    identityNum = dic.Item("identityNum")
    taxationType = dic.Item("taxationType")
    
    totalAmount = dic.Item("totalAmount")
    tradeUsage = dic.Item("tradeUsage")
    tradeOpt = dic.Item("tradeOpt")
    tradeType = dic.Item("tradeType")
    stateMemo = dic.Item("stateMemo")
    stateCode = dic.Item("stateCode")
    stateDT = dic.Item("stateDT")
    printYN = dic.Item("printYN")
    
    confirmNum = dic.Item("confirmNum")
    orgTradeDate = dic.Item("orgTradeDate")
    orgConfirmNum = dic.Item("orgConfirmNum")
    
    ntssendDT = dic.Item("ntssendDT")
    ntsresult = dic.Item("ntsresult")
    ntsresultDT = dic.Item("ntsresultDT")
    ntsresultCode = dic.Item("ntsresultCode")
    ntsresultMessage = dic.Item("ntsresultMessage")
    regDT = dic.Item("regDT")
    
End Sub
