VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PBCBService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'=================================================================================
' Class Module for base module for Popbill API SDK. It include base functionality for
' RESTful web service request and parse json result. It uses Linkhub class module
' to accomplish authentication APIs.
'
' This module uses advapi32.dll and crypt32.dllm. most of windows system has it
' on system32 by default.
'
' http://www.popbill.com
' Author : Kim Seongjun (pallet027@gmail.com)
' Written : 2014-09-29
'
' Thanks for your interest.
'=================================================================================
Private Const ServiceID_REAL = "POPBILL"
Private Const ServiceID_TEST = "POPBILL_TEST"
Private Const ServiceURL_REAL = "https://popbill.linkhub.co.kr"
Private Const ServiceURL_TEST = "https://popbill_test.linkhub.co.kr"
Private Const APIVersion = "1.0"

Private m_IsTest As Boolean

Private m_token_Dic
Private m_Linkhub As Linkhub
Private m_scope As New Collection
Private m_LastErrCode As Long
Private m_LastErrMessage As String

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Sub Class_Initialize()
    Set m_token_Dic = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Class_Terminate()
    m_token_Dic.RemoveAll
    Set m_token_Dic = Nothing
End Sub
Private Function UniStrToUTF8(UniString) As Byte()
   ' Convert a Unicode string to a byte stream of UTF-8
   Dim BArray() As Byte
   Dim TempB() As Byte
   Dim i As Long
   Dim k As Long
   Dim TLen As Long
   Dim b1 As Byte
   Dim b2 As Byte
   Dim UTF16 As Long
   Dim j
   TLen = Len(UniString)  ' Obtain length of Unicode input string
   If TLen = 0 Then Exit Function  ' get out if there's nothing to convert
   k = 0
   For i = 1 To TLen
      ' Work out the UTF16 value of the Unicode character
      CopyMemory b1, ByVal StrPtr(UniString) + ((i - 1) * 2), 1
      CopyMemory b2, ByVal StrPtr(UniString) + ((i - 1) * 2) + 1, 1
      ' Combine the 2 bytes into the Unicode UTF-16
      UTF16 = b2  ' assign b2 to UTF16 before multiplying by 256 to avoid overflow
      UTF16 = UTF16 * 256 + b1
      ' Convert UTF-16 to 2 or 3 bytes of UTF-8
      TempB = ToUTF8(UTF16)
      ' Copy the resultant bytes to BArray
      For j = 0 To UBound(TempB)
         ReDim Preserve BArray(k)
         BArray(k) = TempB(j): k = k + 1
      Next
      ReDim TempB(0)
   Next
   UniStrToUTF8 = BArray  ' Return the resultant UTF-8 byte array
End Function

Private Function ToUTF8(ByVal UTF16 As Long) As Byte()
   ' Convert a 16bit UTF-16BE to 2 or 3 UTF-8 bytes
   Dim BArray() As Byte
   If UTF16 < &H80 Then
      ReDim BArray(0)  ' one byte UTF-8
      BArray(0) = UTF16  ' Use number as is
   ElseIf UTF16 < &H800 Then
      ReDim BArray(1)  ' two byte UTF-8
      BArray(1) = &H80 + (UTF16 And &H3F)  ' Least Significant 6 bits
      UTF16 = UTF16 \ &H40  ' Shift UTF16 number right 6 bits
      BArray(0) = &HC0 + (UTF16 And &H1F)  ' Use 5 remaining bits
   Else
      ReDim BArray(2)  ' three byte UTF-8
      BArray(2) = &H80 + (UTF16 And &H3F)  ' Least Significant 6 bits
      UTF16 = UTF16 \ &H40  ' Shift UTF16 number right 6 bits
      BArray(1) = &H80 + (UTF16 And &H3F)  ' Use next 6 bits
      UTF16 = UTF16 \ &H40  ' Shift UTF16 number right 6 bits again
      BArray(0) = &HE0 + (UTF16 And &HF)  ' Use 4 remaining bits
   End If
   ToUTF8 = BArray  ' Return UTF-8 bytes in an array
End Function
'�ֱ� �����ڵ� Ȯ��
Public Property Get LastErrCode() As Long
    LastErrCode = m_LastErrCode
End Property
'�ֱٿ����޽��� Ȯ��
Public Property Get LastErrMessage() As String
    LastErrMessage = m_LastErrMessage
End Property
'�׽�Ʈ �÷���
Public Property Let IsTest(ByVal value As Boolean)
    m_IsTest = value
End Property
'�ʱ�ȭ
Public Sub Initialize(linkID As String, SecretKey As String)
    Set m_Linkhub = New Linkhub
    
    m_Linkhub.linkID = linkID
    m_Linkhub.SercetKey = SecretKey
    
    m_scope.Add "member"
    ''���ݿ�����
    m_scope.Add "140"
    
    m_IsTest = False
    
End Sub
Private Function getSession_token(CorpNum As String) As String
    Dim refresh As Boolean
    
    refresh = False
    Dim m_Token As LinkhubToken
    
    Set m_Token = Nothing
    
    If m_token_Dic.Exists(CorpNum) Then
        Set m_Token = m_token_Dic.Item(CorpNum)
    End If
    
    If m_Token Is Nothing Then
        refresh = True
    Else
        Dim utcnow As String
        utcnow = Format(m_Linkhub.UTCTime, "yyyy-MM-ddTHH:mm:ss.sssZ")
        refresh = m_Token.expiration < utcnow
    End If
    
    If refresh Then
    
        Set m_Token = m_Linkhub.getToken(IIf(m_IsTest, ServiceID_TEST, ServiceID_REAL), CorpNum, m_scope)
        
        m_token_Dic.Add CorpNum, m_Token
        
        If m_Token Is Nothing Then
            Err.Raise m_Linkhub.LastErrCode, "POPBILL", m_Linkhub.LastErrMessage
            Exit Function
        End If
        
    End If
    
    getSession_token = m_Token.session_token
End Function

Private Function httpGET(url As String, BearerToken As String, Optional UserID As String) As Variant
     
    Dim winhttp1
    
    Set winhttp1 = CreateObject("WinHttp.WinHttpRequest.5.1")
    Call winhttp1.Open("GET", IIf(m_IsTest, ServiceURL_TEST, ServiceURL_REAL) + url)
    
    Call winhttp1.setRequestHeader("Authorization", "Bearer " + BearerToken)
    Call winhttp1.setRequestHeader("x-pb-version", APIVersion)
    
    If UserID <> "" Then
        Call winhttp1.setRequestHeader("x-pb-userid", UserID)
    End If
    
    
    winhttp1.Send
    winhttp1.WaitForResponse
    
    Dim Response As String
    Response = winhttp1.responseText
       
    Dim parsedDic As Object
    
    If winhttp1.Status <> 200 Then
    
        Set parsedDic = m_Linkhub.parse(Response)
        
        Err.Raise parsedDic.Item("code"), "POPBILL", parsedDic.Item("message")
        
        Exit Function
        
    End If
    
    Set winhttp1 = Nothing
    
    Set httpGET = m_Linkhub.parse(Response)

End Function

Private Function httpPOST(url As String, BearerToken As String, postdata As String, Optional UserID As String) As Variant
    Set httpPOST = httpPOST_override(url, BearerToken, "", postdata, UserID)
    
End Function

Private Function httpPOST_override(url As String, BearerToken As String, override As String, postdata As String, Optional UserID As String) As Variant
     
    Dim winhttp1
    
    Set winhttp1 = CreateObject("WinHttp.WinHttpRequest.5.1")
    Call winhttp1.Open("POST", IIf(m_IsTest, ServiceURL_TEST, ServiceURL_REAL) + url)
    Call winhttp1.setRequestHeader("x-pb-version", APIVersion)
    
    
    If BearerToken <> "" Then
        Call winhttp1.setRequestHeader("Authorization", "Bearer " + BearerToken)
    End If
    
    If override <> "" Then
        Call winhttp1.setRequestHeader("X-HTTP-Method-Override", override)
    End If
    
    If UserID <> "" Then
        Call winhttp1.setRequestHeader("x-pb-userid", UserID)
    End If
    
    Call winhttp1.setRequestHeader("Content-Type", "Application/json")
    
    winhttp1.Send (postdata)
    winhttp1.WaitForResponse
    
    Dim Response As String
    Response = winhttp1.responseText
       
    Dim parsedDic As Object
    
    If winhttp1.Status <> 200 Then
    
        Set parsedDic = m_Linkhub.parse(Response)
        
        Err.Raise parsedDic.Item("code"), "POPBILL", IIf(IsNull(parsedDic.Item("message")), "Null Exception", parsedDic.Item("message"))
        
        Exit Function
        
    End If
    
    Set winhttp1 = Nothing
    
    Set httpPOST_override = m_Linkhub.parse(Response)

End Function

Private Function GetOnlyFileName(ByVal FilePath As String) As String

     Dim Temp() As String
     Temp = Split(FilePath, "\")
     GetOnlyFileName = Split(FilePath, "\")(UBound(Temp))

End Function

Private Function httpPOST_File(url As String, BearerToken As String, postdata As String, FilePath As String, Optional UserID As String) As Variant
     
    Dim winhttp1
    Dim boundary As String
    boundary = "---------------------popbill"
    
    Set winhttp1 = CreateObject("WinHttp.WinHttpRequest.5.1")
    Call winhttp1.Open("POST", IIf(m_IsTest, ServiceURL_TEST, ServiceURL_REAL) + url)
    Call winhttp1.setRequestHeader("x-pb-version", APIVersion)
    
    
    If BearerToken <> "" Then
        Call winhttp1.setRequestHeader("Authorization", "Bearer " + BearerToken)
    End If
    
    If UserID <> "" Then
        Call winhttp1.setRequestHeader("x-pb-userid", UserID)
    End If
    
    Call winhttp1.setRequestHeader("Content-Type", "multipart/form-data; boundary=" + boundary)
    
    
    If postdata <> "" Then
        'TODO ÷������ ����.
    
    End If
    
    Dim head, tail As String
    
    head = "--" & boundary & vbCrLf & _
           "Content-Disposition: form-data; name=""Filedata""; filename=""" & GetOnlyFileName(FilePath) & """" + vbCrLf & _
           "Content-Type: application/octet-stream" & vbCrLf & vbCrLf
           
    tail = vbCrLf & "--" & boundary & "--" & vbCrLf
    
    '### �̹��� ���̳ʸ��� ��´�.
    Dim binary() As Byte
    Open FilePath For Binary Access Read As #1
         ReDim binary(LOF(1))
         Get #1, , binary
    Close #1
    
    '### ������ �����͸� �����.
    Dim BHeader1() As Byte, BHeader2() As Byte, Buffer() As Byte
    Dim Size1 As Long, Size2 As Long, Size3 As Long
    BHeader1 = UniStrToUTF8(head)
    BHeader2 = StrConv(tail, vbFromUnicode)
    
    Size1 = UBound(BHeader1)
    Size2 = UBound(binary)
    Size3 = UBound(BHeader2)
    
    ReDim Buffer(Size1 + Size2 + Size3 + 2)
    CopyMemory Buffer(0), BHeader1(0), Size1 + 1
    CopyMemory Buffer(Size1 + 1), binary(0), Size2 + 1
    CopyMemory Buffer(Size1 + Size2 + 2), BHeader2(0), Size3 + 1
    
    winhttp1.Send (Buffer)
    winhttp1.WaitForResponse
    
    Dim Response As String
    Response = winhttp1.responseText
       
    Dim parsedDic As Object
    
    If winhttp1.Status <> 200 Then
    
        Set parsedDic = m_Linkhub.parse(Response)
        
        Err.Raise parsedDic.Item("code"), "POPBILL", IIf(IsNull(parsedDic.Item("message")), "Null Exception", parsedDic.Item("message"))
        
        Exit Function
        
    End If
    
    Set winhttp1 = Nothing
    
    Set httpPOST_File = m_Linkhub.parse(Response)

End Function
'ȸ���ܾ���ȸ
Public Function GetBalance(CorpNum As String) As Double
    On Error GoTo ErrHandler
    
    GetBalance = m_Linkhub.GetBalance(getSession_token(CorpNum), IIf(m_IsTest, ServiceID_TEST, ServiceID_REAL))
    
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    GetBalance = m_LastErrCode
End Function
'��Ʈ�� �ܾ���ȸ
Public Function GetPartnerBalance(CorpNum As String) As Double
    On Error GoTo ErrHandler
    
    GetPartnerBalance = m_Linkhub.GetPartnerBalance(getSession_token(CorpNum), IIf(m_IsTest, ServiceID_TEST, ServiceID_REAL))
    
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    GetPartnerBalance = m_LastErrCode
End Function
'�˺� �⺻ URL
Public Function GetPopbillURL(CorpNum As String, UserID As String, TOGO As String) As String
On Error GoTo ErrHandler
    Dim result As Variant
    
    Set result = httpGET("/?TG=" + TOGO, getSession_token(CorpNum), UserID)
    
    GetPopbillURL = result.Item("url")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    GetPopbillURL = ""
End Function
'ȸ������ ����
Public Function CheckIsMember(CorpNum As String, linkID As String) As PBResponse
On Error GoTo ErrHandler
    Dim result As Variant
    
    Set result = httpGET("/Join?CorpNum=" + CorpNum + "&LID=" + linkID, "")
    
    Set CheckIsMember = New PBResponse
    CheckIsMember.code = result.Item("code")
    CheckIsMember.message = result.Item("message")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set CheckIsMember = Nothing
End Function
'ȸ������
Public Function JoinMember(JoinInfo As PBJoinForm) As PBResponse
On Error GoTo ErrHandler
    Set JoinMember = Nothing
    
    Dim postdata As String
    Dim tmp
    
    Set tmp = CreateObject("Scripting.Dictionary")
    
    tmp.Add "LinkID", JoinInfo.linkID
    tmp.Add "CorpNum", JoinInfo.CorpNum
    tmp.Add "CorpName", JoinInfo.CorpName
    tmp.Add "CEOName", JoinInfo.CEOName
    tmp.Add "Addr", JoinInfo.Addr
    tmp.Add "ZipCode", JoinInfo.ZipCode
    tmp.Add "BizClass", JoinInfo.BizClass
    tmp.Add "BizType", JoinInfo.BizType
    tmp.Add "ContactName", JoinInfo.ContactName
    tmp.Add "ContactEmail", JoinInfo.ContactEmail
    tmp.Add "ContactFAX", JoinInfo.ContactFAX
    tmp.Add "ContactHP", JoinInfo.ContactHP
    tmp.Add "ContactTEL", JoinInfo.ContactTEL
    tmp.Add "ID", JoinInfo.ID
    tmp.Add "PWD", JoinInfo.PWD
    
    postdata = m_Linkhub.toString(tmp)
    
    Dim result As Variant
    
    Set result = httpPOST("/Join", "", postdata)
    
    Set JoinMember = New PBResponse
    
    JoinMember.code = result.Item("code")
    JoinMember.message = result.Item("message")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
End Function
'''''''''''''''''''''''''' End Of PopbillBase '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'���ݿ����� URLȮ��
Public Function GetURL(CorpNum As String, UserID As String, TOGO As String) As String
On Error GoTo ErrHandler
    Dim result As Variant
    
    Set result = httpGET("/Cashbill?TG=" + TOGO, getSession_token(CorpNum), UserID)
    
    GetURL = result.Item("url")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    GetURL = ""
End Function
'����ܰ� Ȯ��
Public Function GetUnitCost(CorpNum As String) As Double
On Error GoTo ErrHandler
    Dim result As Variant
    
    Set result = httpGET("/Cashbill?cfg=UNITCOST", getSession_token(CorpNum))
    
    GetUnitCost = result.Item("unitCost")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    GetUnitCost = m_LastErrCode
End Function

'������ȣ ��뿩�� Ȯ��
Public Function CheckMgtKeyInUse(CorpNum As String, mgtKey As String) As PBResponse
On Error GoTo ErrHandler
    If mgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "������ȣ�� �Էµ��� �ʾҽ��ϴ�."
    End If
    
    Dim result As Variant
    
    Set result = httpGET("/Cashbill/" + mgtKey, getSession_token(CorpNum))
    
    Set CheckMgtKeyInUse = New PBResponse
    
    CheckMgtKeyInUse.code = IIf(IsNull(result.Item("itemKey")), 0, 1)
    CheckMgtKeyInUse.message = IIf(IsNull(result.Item("itemKey")), "�̻����", "�����")
    
    Exit Function
ErrHandler:
    If Err.Number = -11000005 Then
        Set CheckMgtKeyInUse = New PBResponse
        CheckMgtKeyInUse.code = 0
        CheckMgtKeyInUse.message = "�̻����"
        Exit Function
    End If
    
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set CheckMgtKeyInUse = Nothing
End Function
'�ӽ�����
Public Function Register(CorpNum As String, Cashbill As PBCashbill, Optional UserID As String = "") As PBResponse
    On Error GoTo ErrHandler

    Dim result As Variant
    Dim tmpDic As Variant
    
    Set tmpDic = Cashbill.toDictionary
    
    Dim postdata As String
    
    postdata = m_Linkhub.toString(tmpDic)
    
    Set result = httpPOST("/Cashbill", getSession_token(CorpNum), postdata, UserID)
    
    Set Register = New PBResponse
    
    Register.code = result.Item("code")
    Register.message = result.Item("message")
    
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set Register = Nothing

End Function
'����
Public Function Update(CorpNum As String, mgtKey As String, Cashbill As PBCashbill, Optional UserID As String = "") As PBResponse
On Error GoTo ErrHandler
    If mgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "������ȣ�� �Էµ��� �ʾҽ��ϴ�."
    End If

    Dim result As Variant
    
    Dim tmpDic As Variant
    
    Set tmpDic = Cashbill.toDictionary
    
    Dim postdata As String
    
    postdata = m_Linkhub.toString(tmpDic)
    
    Set result = httpPOST_override("/Cashbill/" + mgtKey, _
                                    getSession_token(CorpNum), _
                                    "PATCH", _
                                    postdata, _
                                    UserID)
    
    Set Update = New PBResponse
    
    Update.code = result.Item("code")
    Update.message = result.Item("message")
    
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set Update = Nothing

End Function
'����
Public Function Delete(CorpNum As String, mgtKey As String, Optional UserID As String) As PBResponse
On Error GoTo ErrHandler
    If mgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "������ȣ�� �Էµ��� �ʾҽ��ϴ�."
    End If
    
    Dim result As Variant
    
    Set result = httpPOST_override("/Cashbill/" + mgtKey, getSession_token(CorpNum), "DELETE", "", UserID)
    
    Set Delete = New PBResponse
    
    Delete.code = result.Item("code")
    Delete.message = result.Item("message")
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set Delete = Nothing
End Function
'����
Public Function Issue(CorpNum As String, mgtKey As String, Optional Memo As String, Optional UserID As String) As PBResponse
On Error GoTo ErrHandler
    If mgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "������ȣ�� �Էµ��� �ʾҽ��ϴ�."
    End If
    
    Dim result As Variant
    
    Dim postdata As String
    Dim tmp As Variant
    Set tmp = CreateObject("Scripting.Dictionary")
    
    tmp.Add "memo", Memo
    
    postdata = m_Linkhub.toString(tmp)
    
    Set result = httpPOST_override("/Cashbill/" + mgtKey, getSession_token(CorpNum), "ISSUE", postdata, UserID)
    
    Set Issue = New PBResponse
    
    Issue.code = result.Item("code")
    Issue.message = result.Item("message")
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
End Function

'���� ���
Public Function CancelIssue(CorpNum As String, mgtKey As String, Optional Memo As String, Optional UserID As String) As PBResponse
On Error GoTo ErrHandler
    If mgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "������ȣ�� �Էµ��� �ʾҽ��ϴ�."
    End If
    
    Dim result As Variant
    
    Dim postdata As String
    
    postdata = "{""memo"":""" + Memo + """}"
    
    Set result = httpPOST_override("/Cashbill/" + mgtKey, _
                                getSession_token(CorpNum), "CANCELISSUE", postdata, UserID)
    
    Set CancelIssue = New PBResponse
    
    CancelIssue.code = result.Item("code")
    CancelIssue.message = result.Item("message")
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
End Function

'�̸��� ������
Public Function SendEmail(CorpNum As String, mgtKey As String, Receiver As String, Optional UserID As String) As PBResponse
On Error GoTo ErrHandler
    If mgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "������ȣ�� �Էµ��� �ʾҽ��ϴ�."
    End If
    
    Dim result As Variant
    
    Dim postdata As String
    
    postdata = "{""receiver"":""" + Receiver + """}"
    
    Set result = httpPOST_override("/Cashbill/" + mgtKey, _
                                getSession_token(CorpNum), "EMAIL", postdata, UserID)
    
    Set SendEmail = New PBResponse
    
    SendEmail.code = result.Item("code")
    SendEmail.message = result.Item("message")
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
End Function
'�����缱��
Public Function SendSMS(CorpNum As String, mgtKey As String, Sender As String, Receiver As String, Contents As String, Optional UserID As String) As PBResponse
On Error GoTo ErrHandler
    If mgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "������ȣ�� �Էµ��� �ʾҽ��ϴ�."
    End If
    
    Dim result As Variant
    
    Dim postdata As String
    
    Dim T As Variant
    
    Set T = CreateObject("Scripting.Dictionary")
    
    T.Add "receiver", Receiver
    T.Add "sender", Sender
    T.Add "contents", Contents
    
    postdata = m_Linkhub.toString(T)
    
    Set result = httpPOST_override("/Cashbill/" + mgtKey, _
                                getSession_token(CorpNum), "SMS", postdata, UserID)
    
    Set SendSMS = New PBResponse
    
    SendSMS.code = result.Item("code")
    SendSMS.message = result.Item("message")
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
End Function
'�ѽ� ������
Public Function SendFax(CorpNum As String, mgtKey As String, Sender As String, Receiver As String, Optional UserID As String) As PBResponse
On Error GoTo ErrHandler
    If mgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "������ȣ�� �Էµ��� �ʾҽ��ϴ�."
    End If
    
    Dim result As Variant
    
    Dim postdata As String
    
    Dim T As Variant
    
    Set T = CreateObject("Scripting.Dictionary")
    
    T.Add "receiver", Receiver
    T.Add "sender", Sender
    
    postdata = m_Linkhub.toString(T)
    
    Set result = httpPOST_override("/Cashbill/" + mgtKey, _
                                getSession_token(CorpNum), "FAX", postdata, UserID)
    
    Set SendFax = New PBResponse
    
    SendFax.code = result.Item("code")
    SendFax.message = result.Item("message")
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
End Function
'����/������� Ȯ��
Public Function GetInfo(CorpNum As String, mgtKey As String, Optional UserID As String) As PBCbInfo
On Error GoTo ErrHandler
    If mgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "������ȣ�� �Էµ��� �ʾҽ��ϴ�."
    End If
    
    Dim result As Variant
    
    Set result = httpGET("/Cashbill/" + mgtKey, _
                                getSession_token(CorpNum), UserID)
    
    Set GetInfo = New PBCbInfo
    
    
    GetInfo.fromDictionary result
    
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set GetInfo = Nothing
End Function
'����/�俪���� �ٷ�(�ִ�1000��)Ȯ��
Public Function GetInfos(CorpNum As String, MgtKeyList As Collection, Optional UserID As String) As Collection
On Error GoTo ErrHandler
    If MgtKeyList Is Nothing Or MgtKeyList.Count = 0 Then
        Err.Raise -99999999, "POPBILL", "������ȣ�� �Էµ��� �ʾҽ��ϴ�."
    End If
    
    Dim result As Variant
    
    Dim postdata As String
    
    postdata = m_Linkhub.toString(MgtKeyList)
    
    Set result = httpPOST("/Cashbill/States", _
                                getSession_token(CorpNum), postdata, UserID)
    
    Set GetInfos = New Collection
    
    Dim T As Variant
    Dim info As PBCbInfo
    
    For Each T In result
        Set info = New PBCbInfo
        
        info.fromDictionary T
        
        GetInfos.Add info
    
    Next
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set GetInfos = Nothing
End Function
'������ Ȯ��
Public Function GetDetailInfo(CorpNum As String, mgtKey As String, Optional UserID As String) As PBCashbill
On Error GoTo ErrHandler
    If mgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "������ȣ�� �Էµ��� �ʾҽ��ϴ�."
    End If
    
    Dim result As Variant
    
    Set result = httpGET("/Cashbill/" + mgtKey + "?Detail", _
                                getSession_token(CorpNum), UserID)
    
    Set GetDetailInfo = New PBCashbill
    
    
    GetDetailInfo.fromDictionary result
    
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set GetDetailInfo = Nothing
End Function
'�˾�URL Ȯ��
Public Function GetPopUpURL(CorpNum As String, mgtKey As String, UserID As String) As String
On Error GoTo ErrHandler
    If mgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "������ȣ�� �Էµ��� �ʾҽ��ϴ�."
    End If
    
    Dim result As Variant
    
    Set result = httpGET("/Cashbill/" + mgtKey + "?TG=POPUP", _
                                getSession_token(CorpNum), UserID)
    
    GetPopUpURL = result.Item("url")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    GetPopUpURL = ""
End Function
'�μ� URLȮ��
Public Function GetPrintURL(CorpNum As String, mgtKey As String, UserID As String) As String
On Error GoTo ErrHandler
    If mgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "������ȣ�� �Էµ��� �ʾҽ��ϴ�."
    End If
    
    Dim result As Variant
    
    Set result = httpGET("/Cashbill/" + mgtKey + "?TG=PRINT", _
                                getSession_token(CorpNum), UserID)
    
    GetPrintURL = result.Item("url")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    GetPrintURL = ""
End Function
'���޹޴��� �μ� URLȮ��
Public Function GetEPrintURL(CorpNum As String, mgtKey As String, UserID As String) As String
On Error GoTo ErrHandler
    If mgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "������ȣ�� �Էµ��� �ʾҽ��ϴ�."
    End If
    
    Dim result As Variant
    
    Set result = httpGET("/Cashbill/" + mgtKey + "?TG=EPRINT", _
                                getSession_token(CorpNum), UserID)
    
    GetEPrintURL = result.Item("url")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    GetEPrintURL = ""
End Function

'���� URLȮ��
Public Function GetMailURL(CorpNum As String, mgtKey As String, UserID As String) As String
On Error GoTo ErrHandler
    If mgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "������ȣ�� �Էµ��� �ʾҽ��ϴ�."
    End If
    
    Dim result As Variant
    
    Set result = httpGET("/Cashbill/" + mgtKey + "?TG=MAIL", _
                                getSession_token(CorpNum), UserID)
    
    GetMailURL = result.Item("url")
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    GetMailURL = ""
End Function
'�ٷ��μ� URLȮ��
Public Function GetMassPrintURL(CorpNum As String, MgtKeyList As Collection, UserID As String) As String
On Error GoTo ErrHandler
    If MgtKeyList Is Nothing Or MgtKeyList.Count = 0 Then
        Err.Raise -99999999, "POPBILL", "������ȣ�� �Էµ��� �ʾҽ��ϴ�."
    End If
    
    Dim result As Variant
    
    Dim postdata As String
    
    postdata = m_Linkhub.toString(MgtKeyList)
    
    Set result = httpPOST("/Cashbill/Prints", _
                                getSession_token(CorpNum), postdata, UserID)
    
    GetMassPrintURL = result.Item("url")
    
    Exit Function
    
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    GetMassPrintURL = ""
End Function

'�����̷�Ȯ��
Public Function GetLogs(CorpNum As String, mgtKey As String, Optional UserID As String) As Collection
On Error GoTo ErrHandler
    If mgtKey = "" Then
        Err.Raise -99999999, "POPBILL", "������ȣ�� �Էµ��� �ʾҽ��ϴ�."
    End If
    
    Dim result As Variant
    
    Set result = httpGET("/Cashbill/" + mgtKey + "/Logs", _
                                getSession_token(CorpNum), UserID)
    
    Set GetLogs = New Collection
    
    Dim T As Variant
    Dim info As PBCbLog
    
    For Each T In result
        Set info = New PBCbLog
        
        info.fromDictionary T
        
        GetLogs.Add info
    
    Next
    
    Exit Function
ErrHandler:
    m_LastErrCode = Err.Number
    m_LastErrMessage = Err.Description
    Set GetLogs = Nothing
End Function