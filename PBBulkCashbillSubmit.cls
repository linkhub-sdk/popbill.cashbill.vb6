VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBBulkCashbillSubmit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public cashbills As New Collection


Public Function toDictionary() As Variant
    Set toDictionary = CreateObject("Scripting.Dictionary")
    
    If (cashbills Is Nothing) = False Then
        Dim t_cashbillCollection As New Collection
        Dim row As PBCashbill
        For Each row In cashbills
            t_cashbillCollection.Add row.toDictionary()
        Next
        If t_cashbillCollection.Count > 0 Then
            toDictionary.Add "cashbills", t_cashbillCollection
        End If
    End If
End Function


