VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBCashbill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public mgtKey As String
Public tradeDate As String
Public tradeUsage As String
Public tradeType As String

Public taxationType As String
Public supplyCost As String
Public tax As String
Public serviceFee As String
Public totalAmount As String

Public franchiseCorpNum As String
Public franchiseCorpName As String
Public franchiseCEOName As String
Public franchiseAddr As String
Public franchiseTEL As String

Public identityNum As String
Public customerName As String
Public itemName As String
Public orderNumber As String

Public email As String
Public hp As String
Public fax As String
Public smssendYN As Boolean
Public faxsendYN As Boolean

Public confirmNum As String

Public orgConfirmNum As String
Public orgTradeDate As String



Public Function toDictionary() As Variant
Set toDictionary = CreateObject("Scripting.Dictionary")
    
    toDictionary.Add "mgtKey", mgtKey
    toDictionary.Add "tradeDate", tradeDate
    toDictionary.Add "tradeUsage", tradeUsage
    toDictionary.Add "tradeType", tradeType
    
    toDictionary.Add "taxationType", taxationType
    toDictionary.Add "supplyCost", supplyCost
    toDictionary.Add "tax", tax
    toDictionary.Add "serviceFee", serviceFee
    toDictionary.Add "totalAmount", totalAmount
    
    toDictionary.Add "franchiseCorpNum", franchiseCorpNum
    toDictionary.Add "franchiseCorpName", franchiseCorpName
    toDictionary.Add "franchiseCEOName", franchiseCEOName
    toDictionary.Add "franchiseAddr", franchiseAddr
    toDictionary.Add "franchiseTEL", franchiseTEL
    
    toDictionary.Add "identityNum", identityNum
    toDictionary.Add "customerName", customerName
    toDictionary.Add "itemName", itemName
    toDictionary.Add "orderNumber", orderNumber
    
    toDictionary.Add "email", email
    toDictionary.Add "hp", hp
    toDictionary.Add "fax", fax
    toDictionary.Add "smssendYN", smssendYN
    toDictionary.Add "faxsendYN", faxsendYN
    
    toDictionary.Add "confirmNum", confirmNum
    
    toDictionary.Add "orgConfirmNum", orgConfirmNum
    toDictionary.Add "orgTradeDate", orgTradeDate
    
    
End Function


Public Sub fromDictionary(dic As Variant)

    mgtKey = dic.Item("mgtKey")
    tradeDate = dic.Item("tradeDate")
    tradeUsage = dic.Item("tradeUsage")
    tradeType = dic.Item("tradeType")
    
    taxationType = dic.Item("taxationType")
    supplyCost = dic.Item("supplyCost")
    tax = dic.Item("tax")
    serviceFee = dic.Item("serviceFee")
    totalAmount = dic.Item("totalAmount")
    
    franchiseCorpNum = dic.Item("franchiseCorpNum")
    franchiseCorpName = dic.Item("franchiseCorpName")
    franchiseCEOName = dic.Item("franchiseCEOName")
    franchiseAddr = dic.Item("franchiseAddr")
    franchiseTEL = dic.Item("franchiseTEL")
    
    identityNum = dic.Item("identityNum")
    customerName = dic.Item("customerName")
    itemName = dic.Item("itemName")
    orderNumber = dic.Item("orderNumber")
    
    email = dic.Item("email")
    hp = dic.Item("hp")
    fax = dic.Item("fax")
    smssendYN = dic.Item("smssendYN")
    faxsendYN = dic.Item("faxsendYN")
    
    confirmNum = dic.Item("confirmNum")
    
    orgConfirmNum = dic.Item("orgConfirmNum")
    orgTradeDate = dic.Item("orgTradeDate")
    
End Sub
