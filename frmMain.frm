VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "An FCGI client in Visual Basic. Supports PHP only."
   ClientHeight    =   1950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   ScaleHeight     =   1950
   ScaleWidth      =   5595
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStop 
      Caption         =   "&Stop"
      Enabled         =   0   'False
      Height          =   370
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1450
   End
   Begin VB.CommandButton cmdListen 
      Caption         =   "&Listen"
      Height          =   370
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1450
   End
   Begin VB.Label lblRemoteHost 
      Caption         =   "Remote Host"
      Height          =   330
      Left            =   1800
      TabIndex        =   3
      Top             =   600
      Width           =   2640
   End
   Begin VB.Label lblRemoteIP 
      Caption         =   "Remote IP:"
      Height          =   330
      Left            =   1800
      TabIndex        =   2
      Top             =   1020
      Width           =   2640
   End
   Begin VB.Label lblRemotePort 
      Caption         =   "Remote Port:"
      Height          =   330
      Left            =   1800
      TabIndex        =   1
      Top             =   1440
      Width           =   2640
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
' change this to your server name
Private Const ServerName As String = "Web Server Version 1.0.0"

' this project was designed for only one share
' change the path to the directory you want to share
Private Const PathShared As String = "" 'App.Path can change accordingly

Private Type ConnectionInfo
    FileNum As Integer  ' file number of the file opened on the current connection
    TotalLength As Long ' total length of data to send (including the header)
    TotalSent As Long   ' total data sent
    FileName As String  ' file name of the file to send
    sType As String     'GET or POST
    DataStr As String
    Headers() As String
    boundary As String
    RequestVars As String
End Type
Private Type RequestData
    Name As String
    Value As String
    FileName As String
End Type

Private CInfo() As ConnectionInfo
Private RequestVars() As RequestData, RequestCount As Long

Private WithEvents cServer  As cWinsock      'Server class
Attribute cServer.VB_VarHelpID = -1
Private WithEvents php_cgi_client  As cWinsock
Attribute php_cgi_client.VB_VarHelpID = -1
Private rInfo As ConnectionInfo
Dim FCGI_Content_Received As Boolean
Private Enum FCGI_Consts
     VERSION_1 = 1
     BEGIN_REQUEST = 1
     ABORT_REQUEST = 2
     END_REQUEST = 3
     PARAMS = 4
     stdin = 5
     StdOut = 6
     StdErr = 7
     Data = 8
     GET_VALUES = 9
     GET_VALUES_RESULT = 10
     UNKNOWN_TYPE = 11
     MAXTYPE = 11 'self:: UNKNOWN_TYPE

     RESPONDER = 1
     AUTHORIZER = 2
     Filter = 3

     REQUEST_COMPLETE = 0
     CANT_MPX_CONN = 1
     OVERLOADED = 2
     UNKNOWN_ROLE = 3

'     MAX_CONNS = "MAX_CONNS"
'     MAX_REQS = "MAX_REQS"
'     MPXS_CONNS = "MPXS_CONNS"

     HEADER_LEN = 8

     REQ_STATE_WRITTEN = 1
     REQ_STATE_OK = 2
     REQ_STATE_ERR = 3
     REQ_STATE_TIMED_OUT = 4
End Enum
Private Type type_FCGIPacketHeader
    version As Long
    type As Long
    requestId As Long
    contentLength As Long
    paddingLength As Long
    reserved As Long
    content As String
    response As String
End Type

Private FCGIPacketHeader As type_FCGIPacketHeader
Const connectTimeout = 500
Const readWriteTimeout = 5000
Public Function IsFile(ByVal sFile As String) As Boolean
    IsFile = GetFileAttributes(sFile) <> -1 And ((GetFileAttributes(sFile) And FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY)
End Function
Public Function IsDir(ByVal sDir As String) As Boolean
    IsDir = ((GetFileAttributes(sDir) And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY)
End Function

Function RShift(ByVal lNum As Long, ByVal lBits As Long) As Long
    If lBits <= 0 Then RShift = lNum
    If (lBits <= 0) Or (lBits > 31) Then Exit Function
    
    RShift = (lNum And (2 ^ (31 - lBits) - 1)) * _
        IIf(lBits = 31, &H80000000, 2 ^ lBits) Or _
        IIf((lNum And 2 ^ (31 - lBits)) = 2 ^ (31 - lBits), _
        &H80000000, 0)
End Function

Function LShift(ByVal lNum As Long, ByVal lBits As Long) As Long
    If lBits <= 0 Then LShift = lNum
    If (lBits <= 0) Or (lBits > 31) Then Exit Function
    
    If lNum < 0 Then
        LShift = (lNum And &H7FFFFFFF) \ (2 ^ lBits) Or 2 ^ (31 - lBits)
    Else
        LShift = lNum \ (2 ^ lBits)
    End If
End Function

Property Get LoWord(dwNum As Long) As Integer
    LoWord = dwNum And &HFFFF
End Property

Property Let LoWord(dwNum As Long, ByVal wNewWord As Integer)
    dwNum = dwNum And &HFFFF0000 Or wNewWord
End Property

Property Get HiWord(dwNum As Long) As Integer
    HiWord = ((dwNum And IIf(dwNum < 0, &H7FFF0000, &HFFFF0000)) \ _
        &H10000) Or (-(dwNum < 0) * &H8000)
End Property

Property Let HiWord(dwNum As Long, ByVal wNewWord As Integer)
    dwNum = dwNum And &HFFFF& Or IIf(wNewWord < 0, ((wNewWord And &H7FFF) _
        * &H10000) Or &H80000000, wNewWord * &H10000)
End Property

Property Get LoByte(wNum As Integer) As Byte
    LoByte = wNum And &HFF
End Property

Property Let LoByte(wNum As Integer, ByVal btNewByte As Byte)
    wNum = wNum And &HFF00 Or btNewByte
End Property

Property Get HiByte(wNum As Integer) As Byte
    HiByte = (wNum And &HFF00&) \ &H100
End Property

Property Let HiByte(wNum As Integer, ByVal btNewByte As Byte)
    wNum = wNum And &HFF Or (btNewByte * &H100&)
End Property
'Something = RShift(SomeValue, HowManyBits) ' Same with LShift
'MyWord = HiWord(MyDWord) ' Get the HiWord (same way with LoWord)
'HiWord(MyDWord) = MyWord ' Set the HiWord (same with LoWord)
'
'MyByte = HiByte(MyWord) ' Get the HiByte (same with LoByte)
'HiByte(MyWord) = MyByte ' Set the HiByte (same with LoByte)
' Build a FastCGI packet
Function buildPacket(ByVal ntype As Long, ByVal content As String, Optional ByVal requestId As Long = 1) As String
Dim clen As Long, strpacket As String
     clen = Len(content)
    strpacket = Chr(FCGI_Consts.VERSION_1)        '/* version */
    strpacket = strpacket & Chr(ntype)                   '  /* type */
     strpacket = strpacket & Chr(LShift(requestId, 8) And &HFF)    '/* requestIdB1 */
     strpacket = strpacket & Chr(requestId And &HFF)         '/* requestIdB0 */
      strpacket = strpacket & Chr(LShift(clen, 8) And &HFF)        ' /* contentLengthB1 */
       strpacket = strpacket & Chr(clen And &HFF)              '/* contentLengthB0 */
      strpacket = strpacket & Chr(0)                        '/* paddingLength */
       strpacket = strpacket & Chr(0)                         '/* reserved */
        strpacket = strpacket & content                     ' /* content */
    buildPacket = strpacket
End Function
'Build an FastCGI Name value pair
Function buildNvpair(ByVal Name As String, ByVal Value As String) As String
Dim nlen As Long, vlen As Long, nvpair As String
        nlen = Len(Name)
        vlen = Len(Value)
        If (nlen < 128) Then
'            /* nameLengthB0 */
            nvpair = Chr(nlen)
          Else
'            /* nameLengthB3 & nameLengthB2 & nameLengthB1 & nameLengthB0 */
            nvpair = Chr(LShift(nlen, 24) Or &H80) & Chr(LShift(nlen, 16) And &HFF) & Chr(LShift(nlen, 8) And &HFF) & Chr(nlen And &HFF)
        End If
        If (vlen < 128) Then
'            /* valueLengthB0 */
            nvpair = nvpair & Chr(vlen)
         Else
'            /* valueLengthB3 & valueLengthB2 & valueLengthB1 & valueLengthB0 */
            nvpair = nvpair & Chr(LShift(vlen, 24) Or &H80) & Chr(LShift(vlen, 16) And &HFF) & Chr(LShift(vlen, 8) And &HFF) & Chr(vlen And &HFF)
        End If
'        /* nameData & valueData */
'        Clipboard.Clear
'        Clipboard.SetText nvpair & Name & Value, vbCFText
        buildNvpair = nvpair & Name & Value
End Function
' Read a set of FastCGI Name value pairs
Function readNvpair(ByVal Data As String, Optional ByVal length As Long = 0) As String
Dim p As Long, nlen As Long, vlen As Long, s_array As String
        If (length = 0) Then
            length = Len(Data)
        End If
     p = 0
 Do While (p <> length)
            p = p + 1
         nlen = Asc(Mid$(Data, p, 1))
            If (nlen >= 128) Then
                nlen = RShift(nlen And &H7F, 24)
                    p = p + 1
                nlen = nlen Or RShift(Asc(Mid$(Data, p, 1)), 16)
                    p = p + 1
                nlen = nlen Or RShift(Asc(Mid$(Data, p, 1)), 8)
                    p = p + 1
                nlen = nlen Or Asc(Mid$(Data, p, 1))
            End If
                p = p + 1
             vlen = Asc(Mid$(Data, p, 1))
            If (vlen >= 128) Then
                vlen = RShift(vlen And &H7F, 24)
                    p = p + 1
                vlen = vlen Or RShift(Asc(Mid$(Data, p, 1)), 16)
                    p = p + 1
                vlen = vlen Or RShift(Asc(Mid$(Data, p, 1)), 8)
                    p = p + 1
                vlen = vlen Or Asc(Mid$(Data, p, 1))

            End If
'            $array[substr($data, $p, $nlen)] = substr($data, $p+$nlen, $vlen);
'            Mid$(s_array, p, nlen) = Mid$(Data, p + nlen, vlen)
            s_array = s_array & Mid$(Data, p + 1, nlen) & "=" & Mid$(Data, p + nlen + 1, vlen) & ","
            p = p + nlen + vlen
    Loop
        readNvpair = s_array
End Function
'Decode a FastCGI Packet
Function decodePacketHeader(ByVal Data As String) As String
Dim ret As String
FCGIPacketHeader.version = Asc(Mid$(Data, 1, 1))
FCGIPacketHeader.type = Asc(Mid$(Data, 2, 1))
FCGIPacketHeader.requestId = RShift(Asc(Mid$(Data, 3, 1)), 8) + Asc(Mid$(Data, 4, 1))
FCGIPacketHeader.contentLength = RShift(Asc(Mid$(Data, 5, 1)), 8) + Asc(Mid$(Data, 6, 1))
FCGIPacketHeader.paddingLength = Asc(Mid$(Data, 7, 1))
FCGIPacketHeader.reserved = Asc(Mid$(Data, 8, 1))
        decodePacketHeader = ret
End Function
Function Random2(ByVal nmin As Long, ByVal nmax As Long) As Long
Dim num As Long
r:
    num = Rnd * nmax
If num < nmin Then
    GoTo r
End If
    Random2 = num
End Function
Private Sub cmdListen_Click()
If cServer Is Nothing Then
    Set cServer = New cWinsock
        cServer.Listen 8080
End If
    cmdListen.Enabled = False
    cmdStop.Enabled = True
End Sub

Private Sub cmdStop_Click()
    If Not cServer Is Nothing Then
        cServer.CloseAll
        Set cServer = Nothing
    End If
        cmdStop.Enabled = False
        cmdListen.Enabled = True
        
End Sub

Private Sub cServer_OnClose(ByVal lngSocket As Long)
    Caption = "cServer_OnClose " & lngSocket & " Connections : " & cServer.ConnectionCount
End Sub

Private Sub cServer_ConnectionRequest(ByVal lngSocket As Long)
Dim lngNewSocket As Long
    
    'Accept the connection and store the new socket handle
     lngNewSocket = cServer.Accept(lngSocket)
     
    lblRemoteHost.Caption = "Remote Host: " & cServer.GetRemoteHost(lngNewSocket)
    lblRemoteIP.Caption = "Remote IP: " & cServer.GetRemoteIP(lngNewSocket)
    lblRemotePort.Caption = "Remote Port: " & cServer.GetRemotePort(lngNewSocket)

'MsgBox lngNewSocket
    Caption = "cServer_ConnectionRequest " & lngSocket & " lngNewSocket " & lngNewSocket

End Sub
Public Function URLdecode(ByRef Text As String) As String
    Const Hex = "0123456789ABCDEF"
    Dim lngA As Long, lngB As Long, lngChar As Long, lngChar2 As Long
    URLdecode = Text
    lngB = 1
    For lngA = 1 To LenB(Text) - 1 Step 2
        lngChar = Asc(MidB$(URLdecode, lngA, 2))
        Select Case lngChar
            Case 37
                lngChar = InStr(Hex, MidB$(Text, lngA + 2, 2)) - 1
                If lngChar >= 0 Then
                    lngChar2 = InStr(Hex, MidB$(Text, lngA + 4, 2)) - 1
                    If lngChar2 >= 0 Then
                        MidB$(URLdecode, lngB, 2) = Chr$((lngChar * &H10&) Or lngChar2)
                        lngA = lngA + 4
                    Else
                        If lngB < lngA Then MidB$(URLdecode, lngB, 2) = MidB$(Text, lngA, 2)
                    End If
                Else
                    If lngB < lngA Then MidB$(URLdecode, lngB, 2) = MidB$(Text, lngA, 2)
                End If
            Case 43
                MidB$(URLdecode, lngB, 2) = " "
            Case Else
                If lngB < lngA Then MidB$(URLdecode, lngB, 2) = MidB$(Text, lngA, 2)
        End Select
        lngB = lngB + 2
    Next lngA
    URLdecode = LeftB$(URLdecode, lngB - 1)
End Function

Public Function URLencode(ByRef Text As String) As String
    Const Hex = "0123456789ABCDEF"
    Dim lngA As Long, lngChar As Long
    URLencode = Text
    For lngA = LenB(URLencode) - 1 To 1 Step -2
        lngChar = Asc(MidB$(URLencode, lngA, 2))
        Select Case lngChar
            Case 48 To 57, 65 To 90, 97 To 122
            Case 32
                MidB$(URLencode, lngA, 2) = "+"
            Case Else
                URLencode = LeftB$(URLencode, lngA - 1) & "%" & Mid$(Hex, (lngChar And &HF0) \ &H10 + 1, 1) & Mid$(Hex, (lngChar And &HF&) + 1, 1) & MidB$(URLencode, lngA + 2)
        End Select
    Next lngA
End Function
Public Function URLdecshort(ByRef Text As String) As String
    Dim strArray() As String, lngA As Long
    strArray = Split(Replace(Text, "+", " "), "%")
    For lngA = 1 To UBound(strArray)
        strArray(lngA) = Chr$("&H" & Left$(strArray(lngA), 2)) & Mid$(strArray(lngA), 3)
    Next lngA
    URLdecshort = Join(strArray, vbNullString)
End Function

Public Function URLencshort(ByRef Text As String) As String
    Dim lngA As Long, strChar As String
    For lngA = 1 To Len(Text)
        strChar = Mid$(Text, lngA, 1)
        If strChar Like "[A-Za-z0-9]" Then
        ElseIf strChar = " " Then
            strChar = "+"
        Else
            strChar = "%" & Right$("0" & Hex$(Asc(strChar)), 2)
        End If
        URLencshort = URLencshort & strChar
    Next lngA
End Function
Function Header(ByVal Key As String) As String
Dim C As Long
    For C = 0 To UBound(rInfo.Headers)
        If InStr(LCase$(rInfo.Headers(C)), LCase$(Key) & ":") Then
                Header = Trim$(Split(rInfo.Headers(C), ":")(1))
            Exit For
        End If
    Next
End Function
Function request(ByVal Key As String) As String
Dim C As Long
    For C = 0 To RequestCount - 1
        If LCase$(RequestVars(C).Name) = LCase$(Key) Then
                request = URLdecode(RequestVars(C).Value)
            Exit For
        End If
    Next
End Function
Function Process_Request(ByRef ContentType As String) As String
Dim ResponseData As String, C As Long, sFile As String, f As Integer
Dim Ext As String
            If IsDir(App.Path & "\uploads") = False Then
                MkDir App.Path & "\uploads"
            End If
        ResponseData = ""
    If RequestCount > 0 Then

        For C = 0 To RequestCount - 1
                sFile = RequestVars(C).FileName
            If sFile <> "" Then
                    If InStr(sFile, "\") Then
                        sFile = Mid$(sFile, InStrRev(sFile, "\") + 1)
                    End If
                sFile = App.Path & "\uploads\" & sFile
                f = FreeFile

                Open sFile For Binary Access Write As f
                    Put #f, , RequestVars(C).Value
                Close f
            End If
        Next

    End If
Ext = LCase$(Mid$(rInfo.FileName, InStrRev(rInfo.FileName, ".") + 1))
    If Ext = "php" Then
        ResponseData = Process_PHP(rInfo.FileName & "?" & rInfo.RequestVars, "POST")
    End If
        Process_Request = ResponseData
End Function
Sub Process_Post(ByVal lngSocket As Long)
    Dim sHeader As String, ContentType As String
    Dim RequestAction As String, RequestData As String, ResponseData As String
    Dim P1 As Long, P2 As Long, sKey As String, sFile As String, str As String, C As Long
    Dim Header() As String, boundary As String
Dim Idx As Long
            RequestAction = rInfo.FileName
            RequestData = rInfo.DataStr
        For C = 0 To UBound(rInfo.Headers)
                Header = Split(rInfo.Headers(C), ":")
            If Header(0) = "Content-Type" Then
                If InStr(Header(1), "multipart/form-data") Then
                    ContentType = "multipart/form-data"
                End If
                If InStr(Header(1), "boundary=") Then
                    boundary = Split(Header(1), "boundary=")(1)
                End If
                If InStr(Header(1), "application/x-www-form-urlencoded") Then
                    ContentType = "application/x-www-form-urlencoded"
                End If

                Exit For
            End If
        Next
            RequestCount = 0
            Erase RequestVars

If ContentType = "multipart/form-data" Then
            P1 = InStr(RequestData, boundary)
        Do While P1 <> 0
                P1 = InStr(P1 + 1, RequestData, "Content-Disposition:")
                str = ""
                If P1 Then
                        P1 = P1 + Len("Content-Disposition:")
                    P2 = InStr(P1 + 1, RequestData, vbCrLf)
                    If P2 Then
                        str = Mid$(RequestData, P1, (P2 - P1) + 1)
                            P1 = P2
                    End If
                End If
                    Idx = -1
            If str <> "" Then
                        sKey = ""
                        sFile = ""
                    Header = Split(str, ";")
                        For C = 0 To UBound(Header)
                            If InStr(Header(C), "name=") And InStr(Header(C), "filename=") = 0 Then
                                str = Split(Header(C), "name=")(1)
                                sKey = Mid$(str, 2, Len(str) - 3)
                            End If
                            If InStr(Header(C), "filename=") Then
                                str = Split(Header(C), "filename=")(1)
                                sFile = Mid$(str, 2, Len(str) - 3)
                            End If
                        Next
                    If Right$(sKey, 1) = "[" Or sFile <> "" Then
                        P2 = 0
                         For C = 0 To RequestCount - 1
                            If InStr(RequestVars(C).Name, Mid$(sKey, 1, Len(sKey))) Then
                                P2 = P2 + 1
                            End If
                        Next
                        If Right$(sKey, 1) = "[" Then
                            sKey = sKey & P2 & "]"
                        Else
                            sKey = sKey & "[" & P2 & "]"
                        End If
                    End If
                For C = 0 To RequestCount - 1
                    If RequestVars(C).Name = sKey Then
                        Idx = C
                        Exit For
                    End If
                Next
                    
                If Idx = -1 Then
                        Idx = RequestCount
                    ReDim Preserve RequestVars(Idx)
                    RequestVars(Idx).Name = sKey
                    RequestVars(Idx).FileName = sFile
                    RequestCount = RequestCount + 1
                End If
            End If
            If P1 Then
                P1 = InStr(P1, RequestData, vbCrLf & vbCrLf)
                    str = ""
                    If P1 Then
                            P2 = InStr(P1 + 1, RequestData, boundary)
                        If P2 Then
                            str = Mid$(RequestData, P1 + 4, (P2 - P1) - 8)
                            P1 = P2
                        End If
                    End If
                If str <> "" And Idx <> -1 Then
                    RequestVars(Idx).Value = RequestVars(Idx).Value & str
                End If
            End If
                If P1 = 0 Then
                    Exit Do
                End If
            P1 = P1 + 1
                
        Loop
            
ElseIf ContentType = "application/x-www-form-urlencoded" Then
    Dim vars() As String
        Header = Split(RequestData, "&")
    For C = 0 To UBound(Header)
        vars = Split(Header(C), "=")
        ReDim Preserve RequestVars(RequestCount)
        RequestVars(RequestCount).Name = vars(0)
        RequestVars(RequestCount).Value = vars(1)
        RequestCount = RequestCount + 1
    Next
End If

            ContentType = "text/html"
''    rInfo.DataStr = ""
''        For C = 0 To RequestCount - 1
''            rInfo.DataStr = rInfo.DataStr & RequestVars(C).Name & "=" & URLencode(RequestVars(C).Value) & "&"
''        Next
            ResponseData = Process_Request(ContentType)
            
                rInfo.DataStr = ResponseData

        Dim phpHeaders As String
                            
        If rInfo.DataStr <> "" Then
            phpHeaders = Mid$(rInfo.DataStr, 1, InStr(rInfo.DataStr, vbCrLf & vbCrLf) - 1)
            rInfo.DataStr = Mid$(rInfo.DataStr, InStr(rInfo.DataStr, vbCrLf & vbCrLf) + 4)
        End If
                         
    ' build the header
    sHeader = "HTTP/1.1 200 OK" & vbNewLine & _
            "Server: " & ServerName & vbNewLine & _
            "Content-Type: " & ContentType & vbNewLine & _
            "Connection: Keep-Alive" & vbNewLine & _
            "Keep-Alive: timeout=5, max=98" & vbNewLine & _
            "Content-Length: " & Len(rInfo.DataStr) & _
            vbNewLine & phpHeaders & _
            vbNewLine & vbNewLine
            
            
    ' total data send is the header length + the length of the file requested
    rInfo.TotalLength = Len(sHeader) + Len(rInfo.DataStr)
        ' send the header, the Sck_SendComplete event is gonna send the file...
'        Sck(Index).SendData sHeader
    
        cServer.Send lngSocket, sHeader & rInfo.DataStr
        
 
End Sub
Function Process_PHP(ByVal cmdFile As String, ByVal Method As String) As String
    Dim cmdArgs As String
    Dim cmdCookie As String
    Dim cmdLine As String
    Dim phpOut As String
          
If InStr(cmdFile, "?") Then
'    cmdArgs = Mid$(cmdFile, InStrRev(cmdFile, "?") + 1)
'    cmdFile = Mid$(cmdFile, 1, InStrRev(cmdFile, "?") - 1)
    cmdArgs = Mid$(cmdFile, InStr(cmdFile, "?") + 1)
    cmdFile = Mid$(cmdFile, 1, InStr(cmdFile, "?") - 1)
End If
Dim uri As String, scriptname As String
    cmdFile = URLdecode(Replace$(cmdFile, "\", "/"))
        uri = Replace$(cmdFile, "\", "/")
    scriptname = uri
            If LCase$(Right$(uri, 9)) = "index.php" Then
                uri = Mid$(uri, 1, Len(uri) - 9)
            End If
    If cmdArgs <> "" Then
        If InStr(cmdArgs, "?") <> 0 Then
            uri = uri & Mid$(cmdArgs, 1, InStr(cmdArgs, "?") - 1)
            cmdArgs = Mid$(cmdArgs, InStr(cmdArgs, "?") + 1)
        End If
        
        uri = uri & "?" & cmdArgs
    End If
    If Left$(cmdFile, 1) = "/" Then
'        cmdFile = "E:" & cmdFile
        cmdFile = Left$(App.Path, 2) & cmdFile
    End If
Dim sPath As String
        If InStr(cmdFile, "/") Then
            sPath = Mid$(cmdFile, 1, InStrRev(cmdFile, "/") - 1)
        End If

Dim sHeader As String
Dim doc_uri As String
    If InStr(cmdFile, uri) Then
        doc_uri = uri & Mid$(cmdFile, InStr(cmdFile, uri) + Len(uri))
    End If
sHeader = sHeader & "GATEWAY_INTERFACE=FastCGI/1.0" & vbCrLf
sHeader = sHeader & "SCRIPT_FILENAME=" & cmdFile & vbCrLf
sHeader = sHeader & "QUERY_STRING=" & cmdArgs & vbCrLf
sHeader = sHeader & "REQUEST_METHOD=" & Method & vbCrLf 'GET OR POST
sHeader = sHeader & "REQUEST_URI=" & uri & vbCrLf
sHeader = sHeader & "DOCUMENT_URI=" & doc_uri & vbCrLf
sHeader = sHeader & "SCRIPT_NAME=" & scriptname & vbCrLf
sHeader = sHeader & "PHP_SELF=" & scriptname & vbCrLf
sHeader = sHeader & "REQUEST_SCHEME=http" & vbCrLf
sHeader = sHeader & "REMOTE_HOST=localhost" & vbCrLf
sHeader = sHeader & "REMOTE_ADDR=127.0.0.1" & vbCrLf
sHeader = sHeader & "REMOTE_PORT=8080" & vbCrLf
sHeader = sHeader & "SERVER_PORT=8080" & vbCrLf
sHeader = sHeader & "HTTP_HOST=localhost:8080" & vbCrLf
sHeader = sHeader & "HTTP_USER_AGENT=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.84 Safari/537.36" & vbCrLf
sHeader = sHeader & "DOCUMENT_ROOT=" & sPath & vbCrLf
sHeader = sHeader & "CONTEXT_DOCUMENT_ROOT=" & sPath & vbCrLf
sHeader = sHeader & "SERVER_SOFTWARE=nginx/1.16.0" & vbCrLf
'sHeader = sHeader & "SERVER_SOFTWARE=Apache/2.4.27 (Win32) OpenSSL/1.1.0f PHP/7.0.16" & vbCrLf
sHeader = sHeader & "SERVER_PROTOCOL=HTTP/1.1" & vbCrLf
sHeader = sHeader & "SERVER_NAME=localhost" & vbCrLf
sHeader = sHeader & "REDIRECT_STATUS=true" & vbCrLf
sHeader = sHeader & "HTTP_COOKIE=" & Header("Cookie") & vbCrLf
sHeader = sHeader & "SERVER_ADDR=127.0.0.1" & vbCrLf
        If Method = "POST" Then
'            sHeader = sHeader & "CONTENT_TYPE=application/x-www-form-urlencoded" & vbCrLf
            sHeader = sHeader & "CONTENT_TYPE=" & Header("Content-type") & vbCrLf
            sHeader = sHeader & "CONTENT_LENGTH=" & Len(rInfo.DataStr) & vbCrLf
            sHeader = sHeader & "HTTP_ACCEPT=" & Header("Accept") & vbCrLf
            sHeader = sHeader & "MAX_FILE_UPLOADS=10" & vbCrLf
            sHeader = sHeader & "PATH_INFO=" & sPath & "/" & vbCrLf
        Else
            sHeader = sHeader & "CONTENT_TYPE=" & vbCrLf
            sHeader = sHeader & "CONTENT_LENGTH=0" & vbCrLf
            sHeader = sHeader & "HTTP_ACCEPT=text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8" & vbCrLf

        End If
Dim request  As String, paramsRequest As String, nvpair() As String, C As Long, nv() As String, resp As String
Dim stdin As String
 
Dim id As Long, keepAlive As Long
            Randomize
        id = Random2(1, RShift(1, 16) - 1)
'        // Using persistent sockets implies you want them keep alive by server!
        keepAlive = 1 'true
request = Chr(0) & Chr(FCGI_Consts.RESPONDER) & Chr(keepAlive) & String$(5, Chr(0))
request = buildPacket(FCGI_Consts.BEGIN_REQUEST, request, id)
paramsRequest = ""
nvpair = Split(sHeader, vbCrLf)
For C = 0 To UBound(nvpair) - 1
        nv = Split(nvpair(C), "=")
        paramsRequest = paramsRequest & buildNvpair(nv(0), nv(1))
Next
 
        If (paramsRequest <> "") Then
            request = request & buildPacket(FCGI_Consts.PARAMS, paramsRequest, id)
        End If
        request = request & buildPacket(FCGI_Consts.PARAMS, "", id)

        If (stdin <> "") Then
            request = request & buildPacket(FCGI_Consts.stdin, stdin, id)
        End If
        request = request & buildPacket(FCGI_Consts.stdin, "", id)
        
If php_cgi_client Is Nothing Then
    Set php_cgi_client = New cWinsock
End If
    php_cgi_client.CloseAll
            FCGI_Content_Received = False
FCGIPacketHeader.content = request
FCGIPacketHeader.response = ""
 
        php_cgi_client.Connect "127.0.0.1", 9000
'Do While FCGI_Content_Received = False
'    DoEvents
'Loop
'    Process_PHP = FCGIPacketHeader.response
    
End Function
Private Sub cServer_DataArrival(ByVal lngSocket As Long)
    Dim rData As String, sHeader As String, RequestedFile As String, ContentType As String
    Dim CompletePath As String, RequestVars As String
    Dim RequestHeader As String, RequestData As String
    Dim Headers() As String, Header() As String, C As Long, boundary As String
    Dim Ext As String
    
cServer.Recv lngSocket, rData
            
    If rData Like "POST * HTTP/1.?*" Then
                If IsFile(App.Path & "\post_data.txt") Then
                    Kill App.Path & "\post_data.txt"
                End If
                C = FreeFile
            Open App.Path & "\post_data.txt" For Append As C
                Print #C, rData
            Close C
        RequestedFile = LeftRange(rData, "POST ", " HTTP/1.", , 1)
                RequestVars = ""
            If InStr(RequestedFile, "?") Then
                RequestVars = Mid$(RequestedFile, InStrRev(RequestedFile, "?") + 1)
                RequestedFile = Mid$(RequestedFile, 1, InStrRev(RequestedFile, "?") - 1)
            End If
            
                If Right$(RequestedFile, 1) = "/" Then
                    If IsFile(Replace$(RequestedFile, "/", "\") & "index.php") = True Then
                        RequestedFile = RequestedFile & "index.php"
                    End If
                End If
            If IsFile(Replace$(RequestedFile, "/", "\") & "\index.php") = True Then
                RequestedFile = RequestedFile & "/index.php"
            End If
                If InStr(RequestedFile, "public/") And IsFile(Replace$(RequestedFile, "/", "\")) = False Then
                    RequestVars = Mid$(RequestedFile, InStrRev(RequestedFile, "public/") + 7) & "?" & RequestVars
                    RequestedFile = Mid$(RequestedFile, 1, InStrRev(RequestedFile, "public/") + 6) & "index.php"
                End If
                    
                    
                rInfo.RequestVars = RequestVars
                rInfo.FileName = RequestedFile
                rInfo.sType = "POST"
                rInfo.FileNum = -1 ' very important
                rInfo.boundary = ""
            rData = Mid$(rData, InStr(rData, vbCrLf) + 2)
        RequestHeader = Mid$(rData, 1, InStr(rData, vbCrLf & vbCrLf) - 1)
        RequestData = Mid$(rData, InStr(rData, vbCrLf & vbCrLf) + 4)
            Headers = Split(RequestHeader, vbCrLf)
                rInfo.Headers = Headers
                rInfo.TotalLength = 0
        For C = 0 To UBound(Headers)
                Header = Split(Headers(C), ":")
            If Header(0) = "Content-Length" Then
                rInfo.TotalLength = Val(Header(1))
            End If
            If Header(0) = "Content-Type" Then
                If InStr(Header(1), "boundary=") Then
                    boundary = Split(Header(1), "boundary=")(1)
                End If
            End If
        Next
             rInfo.boundary = boundary

        rInfo.DataStr = RequestData
                    
        If Len(rInfo.DataStr) >= rInfo.TotalLength Then
            If rInfo.boundary <> "" Then
                If Mid$(rInfo.DataStr, Len(rInfo.DataStr) - (Len(rInfo.boundary) + 3), Len(rInfo.boundary)) = rInfo.boundary Then
                    Process_Post lngSocket
                End If
            Else
                Process_Post lngSocket
            End If
        End If
        
    ElseIf rData Like "GET * HTTP/1.?*" Then
                If IsFile(App.Path & "\post_data.txt") Then
                    Kill App.Path & "\post_data.txt"
                End If
                C = FreeFile
            Open App.Path & "\post_data.txt" For Append As C
                Print #C, rData
            Close C
        ' get requested file name
        RequestedFile = LeftRange(rData, "GET ", " HTTP/1.", , 1)
            rData = Mid$(rData, InStr(rData, vbCrLf) + 2)
        RequestHeader = Mid$(rData, 1, InStr(rData, vbCrLf & vbCrLf) - 1)
        RequestData = Mid$(rData, InStr(rData, vbCrLf & vbCrLf) + 4)
            Headers = Split(RequestHeader, vbCrLf)
                rInfo.Headers = Headers
        ' check if request contains "/../" or "/./" or "*" or "?"
        ' (probably someone trying to get a file that is outside of the share)
                RequestVars = ""
            If InStr(RequestedFile, "?") Then
                RequestVars = Mid$(RequestedFile, InStrRev(RequestedFile, "?") + 1)
                RequestedFile = Mid$(RequestedFile, 1, InStrRev(RequestedFile, "?") - 1)
            End If
'                        If Left$(RequestedFile, 1) = "/" Then
'                            RequestedFile = Mid$(App.Path, 1, 2) & RequestedFile
'                        End If
             
                If Right$(RequestedFile, 1) = "/" Then
                    If IsFile(Replace$(RequestedFile, "/", "\") & "index.php") = True Then
                        RequestedFile = RequestedFile & "index.php"
                    End If
                End If
            If IsFile(Replace$(RequestedFile, "/", "\") & "\index.php") = True Then
                RequestedFile = RequestedFile & "/index.php"
            End If
                If InStr(RequestedFile, "public/") And IsFile(Replace$(RequestedFile, "/", "\")) = False Then
                    RequestVars = Mid$(RequestedFile, InStrRev(RequestedFile, "public/") + 7) & "?" & RequestVars
                    RequestedFile = Mid$(RequestedFile, 1, InStrRev(RequestedFile, "public/") + 6) & "index.php"
                End If
                    
                rInfo.RequestVars = RequestVars
                rInfo.FileName = RequestedFile
                rInfo.sType = "GET"
                rInfo.FileNum = -1 ' very important
                rInfo.boundary = ""
                rInfo.DataStr = ""
        If InStr(1, RequestedFile, "/../") > 0 Or InStr(1, RequestedFile, "/./") > 0 Or _
                InStr(1, RequestedFile, "*") > 0 Or InStr(1, RequestedFile, "?") > 0 Or RequestedFile = "" Then
            
            ' send "Not Found" error ...
            sHeader = "HTTP/1.0 404 Not Found" & vbNewLine & "Server: " & ServerName & vbNewLine & vbNewLine
            rInfo.TotalLength = Len(sHeader)
'            Sck(Index).SendData sHeader
            cServer.Send lngSocket, sHeader
        Else
            CompletePath = Replace(PathShared & Replace(RequestedFile, "/", "\"), "\\", "\")
            CompletePath = Replace(CompletePath, "%20", " ")
'            Debug.Print CompletePath
            
'            If Dir(CompletePath, vbArchive + vbReadOnly + vbDirectory) <> "" Then
             If IsDir(CompletePath) Or IsFile(CompletePath) Then
'                If (GetAttr(CompletePath) And vbDirectory) = vbDirectory Then
                 If IsDir(CompletePath) Then
                    ' the request if for a directory listing...
                    
                    rInfo.DataStr = BuildHTMLDirList(PathShared, RequestedFile)
                    rInfo.FileNum = -1
                    
                    ' build the header
                    sHeader = "HTTP/1.0 200 OK" & vbNewLine & _
                            "Server: " & ServerName & vbNewLine & _
                            "Content-Type: text/html" & vbNewLine & _
                            "Content-Length: " & Len(rInfo.DataStr) & vbNewLine & _
                            vbNewLine
                    
                    ' total data send is the header length + the length of the file requested
                    rInfo.TotalLength = Len(sHeader) + Len(rInfo.DataStr)
                Else
                    ' requested file exists, open the file, send header, and start the transfer
                    rInfo.FileName = RequestedFile
                    ' since one or more files may be opened at the same time, have to get the free file number
                    rInfo.FileNum = FreeFile
                    Open CompletePath For Binary Access Read As rInfo.FileNum
                    ' get content-type depending on the file extension
'                        Ext = LCase(LeftRight(RequestedFile, ".", , 1))
                        Ext = LCase$(Mid$(RequestedFile, InStrRev(RequestedFile, ".") + 1))
                    Select Case Ext
                        Case "jpg", "jpeg"
                            ContentType = "Content-Type: image/jpeg"
                        Case "gif"
                            ContentType = "Content-Type: image/gif"
                        Case "htm", "html"
                            ContentType = "Content-Type: text/html"
                        Case "js"
                            ContentType = "Content-Type: text/javascript"
                        Case "css"
                            ContentType = "Content-Type: text/css"
                        Case "zip"
                            ContentType = "Content-Type: application/zip"
                        Case "mp3"
                            ContentType = "Content-Type: audio/mpeg"
                        Case "m3u", "pls", "xpl"
                            ContentType = "Content-Type: audio/x-mpegurl"
                        Case "php"
                            ContentType = "Content-Type: text/html"
                        Case Else
                            ContentType = "Content-Type: */*"
                    End Select
                    
                    ' build the header
                    
                    If Ext <> "php" Then
                        sHeader = "HTTP/1.0 200 OK" & vbNewLine & _
                                "Server: " & ServerName & vbNewLine & _
                                ContentType & vbNewLine & _
                                "Content-Length: " & LOF(rInfo.FileNum) & vbNewLine & _
                                vbNewLine
                        ' total data send is the header length + the length of the file requested
                        rInfo.TotalLength = Len(sHeader) + LOF(rInfo.FileNum)
                                    rInfo.DataStr = String$(LOF(rInfo.FileNum), 0)
                                Get rInfo.FileNum, , rInfo.DataStr
                        Close rInfo.FileNum
                    Else
                            Close rInfo.FileNum
                            Dim phpHeaders As String
                                
                            rInfo.DataStr = Process_PHP(CompletePath & "?" & RequestVars, "GET")
'                            If rInfo.DataStr <> "" And InStr(rInfo.DataStr, vbCrLf & vbCrLf) Then
'                                phpHeaders = Mid$(rInfo.DataStr, 1, InStr(rInfo.DataStr, vbCrLf & vbCrLf) - 1)
'                                rInfo.DataStr = Mid$(rInfo.DataStr, InStr(rInfo.DataStr, vbCrLf & vbCrLf) + 4)
'                            End If
'                            If rInfo.DataStr = "" Then
'                                rInfo.DataStr = phpHeaders
'                                phpHeaders = ""
'                            End If
'                            rInfo.TotalLength = Len(sHeader) + Len(rInfo.DataStr)
'                            sHeader = "HTTP/1.0 200 OK" & vbNewLine & _
'                                    "Server: " & ServerName & vbNewLine & _
'                                    ContentType & vbNewLine & _
'                                    "Content-Length: " & Len(rInfo.DataStr) & _
'                                      vbNewLine & phpHeaders & _
'                                    vbNewLine & vbNewLine
                                Exit Sub
                        End If
                End If
                ' send the header, the Sck_SendComplete event is gonna send the file...
                cServer.Send lngSocket, sHeader & rInfo.DataStr
            Else
                ' send "Not Found" if file does not exsist on the share
                sHeader = "HTTP/1.0 404 Not Found" & vbNewLine & "Server: " & ServerName & vbNewLine & vbNewLine
                rInfo.TotalLength = Len(sHeader)
                cServer.Send lngSocket, sHeader
            End If
        End If
    Else
        If rData = "" Then Exit Sub
                C = FreeFile
            Open App.Path & "\post_data.txt" For Append As C
                Print #C, rData
            Close C
        If rInfo.sType = "POST" Then
            rInfo.DataStr = rInfo.DataStr & rData
            If Len(rInfo.DataStr) >= rInfo.TotalLength Then
                If rInfo.boundary <> "" Then
                    If Mid$(rInfo.DataStr, Len(rInfo.DataStr) - (Len(rInfo.boundary) + 3), Len(rInfo.boundary)) = rInfo.boundary Then
                        Process_Post lngSocket
                    End If
                Else
                    Process_Post lngSocket
                End If
            End If
            
        Else
            ' sometimes the browser makes "HEAD" requests (but it's not inplemented in this project)
            sHeader = "HTTP/1.0 501 Not Implemented" & vbNewLine & "Server: " & ServerName & vbNewLine & vbNewLine
            rInfo.TotalLength = Len(sHeader)
            cServer.Send lngSocket, sHeader
        End If
    End If
End Sub
 
Private Sub cServer_OnError(ByVal lngRetCode As Long, ByVal strDescription As String)
Caption = "cServer_OnError " & strDescription
End Sub

Private Sub cServer_SendProgress(ByVal lngSocket As Long, ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    Caption = "cServer_SendProgress " & lngSocket & " bytesSent " & bytesSent & " bytesRemaining " & bytesRemaining
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If Not php_cgi_client Is Nothing Then
        php_cgi_client.CloseAll
        Set php_cgi_client = Nothing
    End If
    If Not cServer Is Nothing Then
        cServer.CloseAll
        Set cServer = Nothing
    End If
    
End Sub
 
Private Function BuildHTMLDirList(ByVal Root As String, ByVal DirToList As String)
    Dim Dirs As New Collection, Files As New Collection
    Dim sDir As String, Path As String, HTML As String, K As Long
    
    Root = Replace(Root, "/", "\")
    DirToList = Replace(DirToList, "/", "\")
    
    If Right(Root, 1) <> "\" Then Root = Root & "\"
    If Left(DirToList, 1) = "\" Then DirToList = Mid(DirToList, 2)
    If Right(DirToList, 1) <> "\" Then DirToList = DirToList & "\"
    
    DirToList = Replace(DirToList, "%20", " ")
    
    If IsDir(Root & DirToList) = True Then
        sDir = Dir(Replace(Root & DirToList, "\\", "\") & "*.*", vbArchive + vbDirectory + vbReadOnly)
        
    End If
    
    Do Until Len(sDir) = 0
'        If sDir <> ".." And sDir <> "." Then
        If sDir <> "." Then
            Path = Replace(Root & DirToList, "\\", "\") & sDir
            
            
'            If (GetAttr(Path) And vbDirectory) = vbDirectory Then
'            If IsDir(Path) = vbDirectory Or (sDir = ".." Or sDir = ".") Then
            If IsDir(Path) = True Or sDir = ".." Then
                Dirs.Add sDir
            ElseIf IsFile(Path) Then
                Files.Add sDir
            End If
        End If
        
        sDir = Dir
    Loop
    
    HTML = "<html><body>"
    
    If Dirs.Count > 0 Then
        HTML = HTML & "<b>Directories:</b><br>"
        
        For K = 1 To Dirs.Count
            HTML = HTML & "<a href=""" & _
                Replace(Replace("/" & DirToList & Dirs(K), "\", "/"), "//", "/") & """>" & _
                Dirs(K) & "</a><br>" & vbNewLine
        Next K
    End If
    
    If Files.Count > 0 Then
        HTML = HTML & "<br><b>Files:</b><br><table width=""100%"" border=""1"" cellpadding=""3"" cellspacing=""2"">" & vbNewLine
        
        For K = 1 To Files.Count
            HTML = HTML & "<tr>" & vbNewLine
            HTML = HTML & "<td width=""100%""><a href=""" & _
                Replace(Replace("/" & DirToList & Files(K), "\", "/"), "//", "/") & """>" & _
                Files(K) & "</a></td>" & vbNewLine
            
            HTML = HTML & "<td nowrap>" & _
                Format(FileLen(Replace(Root & DirToList, "\\", "\") & Files(K)) / 1024#, "###,###,###,##0") & _
                " KBytes</td>" & vbNewLine
            HTML = HTML & "</tr>" & vbNewLine
        Next K
        
        HTML = HTML & "</table>" & vbNewLine
    End If
    
    If Dirs.Count = 0 And Files.Count = 0 Then
        HTML = HTML & "This folder is empty."
    End If
    
    BuildHTMLDirList = HTML & "</body></html>"
End Function

' Search from end to beginning, and return the left side of the string
Public Function RightLeft(ByRef str As String, RFind As String, Optional Compare As VbCompareMethod = vbBinaryCompare, Optional RetError As Long = 0) As String
    Dim K As Long
    
    K = InStrRev(str, RFind, , Compare)
    
    If K = 0 Then
        RightLeft = IIf(RetError = 0, str, "")
    Else
        RightLeft = Left(str, K - 1)
    End If
End Function

' Search from end to beginning and return the right side of the string
Public Function RightRight(ByRef str As String, RFind As String, Optional Compare As VbCompareMethod = vbBinaryCompare, Optional RetError As Long = 0) As String
    Dim K As Long
    
    K = InStrRev(str, RFind, , Compare)
    
    If K = 0 Then
        RightRight = IIf(RetError = 0, str, "")
    Else
        RightRight = Mid(str, K + 1, Len(str))
    End If
End Function

' Search from the beginning to end and return the left size of the string
Public Function LeftLeft(ByRef str As String, LFind As String, Optional Compare As VbCompareMethod = vbBinaryCompare, Optional RetError As Long = 0) As String
    Dim K As Long
    
    K = InStr(1, str, LFind, Compare)
    If K = 0 Then
        LeftLeft = IIf(RetError = 0, str, "")
    Else
        LeftLeft = Left(str, K - 1)
    End If
End Function

' Search from the beginning to end and return the right size of the string
Public Function LeftRight(ByRef str As String, LFind As String, Optional Compare As VbCompareMethod = vbBinaryCompare, Optional RetError As Long = 0) As String
    Dim K As Long
    
    K = InStr(1, str, LFind, Compare)
    If K = 0 Then
        LeftRight = IIf(RetError = 0, str, "")
    Else
        LeftRight = Right(str, (Len(str) - Len(LFind)) - K + 1)
    End If
End Function

' Search from the beginning to end and return from StrFrom string to StrTo string
' both strings (StrFrom and StrTo) must be found in order to be successfull
Public Function LeftRange(ByRef str As String, StrFrom As String, StrTo As String, Optional Compare As VbCompareMethod = vbBinaryCompare, Optional RetError As Long = 0) As String
    Dim K As Long, q As Long
    
    K = InStr(1, str, StrFrom, Compare)
    If K > 0 Then
        q = InStr(K + Len(StrFrom), str, StrTo, Compare)
        
        If q > K Then
            LeftRange = Mid(str, K + Len(StrFrom), (q - K) - Len(StrFrom))
        Else
            LeftRange = IIf(RetError = 0, str, "")
        End If
    Else
        LeftRange = IIf(RetError = 0, str, "")
    End If
End Function

' Search from the end to beginning and return from StrFrom string to StrTo string
' both strings (StrFrom and StrTo) must be found in order to be successfull
Public Function RightRange(ByRef str As String, StrFrom As String, StrTo As String, Optional Compare As VbCompareMethod = vbBinaryCompare, Optional RetError As Long = 0) As String
    Dim K As Long, q As Long
    
    K = InStrRev(str, StrTo, , Compare)
    If K > 0 Then
        q = InStrRev(str, StrFrom, K, Compare)
        
        If q > 0 Then
            RightRange = Mid(str, q + Len(StrFrom), (K - q) - Len(StrTo))
        Else
            RightRange = IIf(RetError = 0, str, "")
        End If
    Else
        RightRange = IIf(RetError = 0, str, "")
    End If
End Function
 
Public Function Base64Decode(ByVal base64String As String)
  Const Base64CodeBase = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
  Dim dataLength As Long, Out As String, groupBegin As Long
  
  dataLength = Len(base64String)
  Out = ""

  If dataLength Mod 4 <> 0 Then
    Err.Raise 1, "Base64Decode", "Bad Base64 string."
    Exit Function
  End If

  ' Now decode each group:
  For groupBegin = 1 To dataLength Step 4
    Dim numDataBytes, CharCounter, thisChar, thisData, groupData
    ' Each data group encodes up To 3 actual bytes.
    numDataBytes = 3
    groupData = 0

    For CharCounter = 0 To 3
      ' Convert each character into 6 bits of data, And add it To
      ' an integer For temporary storage.  If a character is a '=', there
      ' is one fewer data byte.  (There can only be a maximum of 2 '=' In
      ' the whole string.)

      thisChar = Mid(base64String, groupBegin + CharCounter, 1)

      If thisChar = "=" Then
        numDataBytes = numDataBytes - 1
        thisData = 0
      Else
        thisData = InStr(Base64CodeBase, thisChar) - 1
      End If
      If thisData = -1 Then
        Err.Raise 2, "Base64Decode", "Bad character In Base64 string."
        Exit Function
      End If

      groupData = 64 * groupData + thisData
    Next

    ' Convert 3-byte integer into up To 3 characters
    Dim OneChar
    For CharCounter = 1 To numDataBytes
      Select Case CharCounter
        Case 1: OneChar = groupData \ 65536
        Case 2: OneChar = (groupData And 65535) \ 256
        Case 3: OneChar = (groupData And 255)
      End Select
      Out = Out & Chr(OneChar)
    Next
  Next

  Base64Decode = Out
End Function
Public Function CorrectFormat(ByVal strData As String) As String
      strData = Replace$(strData, "%22", Chr$(34))
      strData = Replace$(strData, "%3C", "<")
      strData = Replace$(strData, "%3E", ">")
      strData = Replace$(strData, "+", " ")
      strData = Replace$(strData, "%0D%0A", "<br>")
      strData = Replace$(strData, "%21", "!")
      strData = Replace$(strData, "%22", "&quot;")
      strData = Replace$(strData, "%20", " ")
      strData = Replace$(strData, "%A7", "§")
      strData = Replace$(strData, "%24", "$")
      strData = Replace$(strData, "%25", "%")
      strData = Replace$(strData, "%26", "&")
      strData = Replace$(strData, "%2F", "/")
      strData = Replace$(strData, "%28", "(")
      strData = Replace$(strData, "%29", ")")
      strData = Replace$(strData, "%3D", "=")
      strData = Replace$(strData, "%3F", "?")
      strData = Replace$(strData, "%B2", "²")
      strData = Replace$(strData, "%B3", "³")
      strData = Replace$(strData, "%7B", "{")
      strData = Replace$(strData, "%5B", "[")
      strData = Replace$(strData, "%5D", "]")
      strData = Replace$(strData, "%7D", "}")
      strData = Replace$(strData, "%5C", "\")
      strData = Replace$(strData, "%DF", "ß")
      strData = Replace$(strData, "%23", "#")
      strData = Replace$(strData, "%27", "'")
      strData = Replace$(strData, "%3A", ":")
      strData = Replace$(strData, "%2C", ",")
      strData = Replace$(strData, "%3B", ";")
      strData = Replace$(strData, "%60", "`")
      strData = Replace$(strData, "%7E", "~")
      strData = Replace$(strData, "%2B", "+")
      strData = Replace$(strData, "%B4", "´")
      CorrectFormat = strData
End Function

Private Sub php_cgi_client_DataArrival(ByVal lngSocket As Long)
Dim packet As String, buf As String, BufLen As Long, bytes_received As Long
Do
        packet = ""
        bytes_received = php_cgi_client.Recv(lngSocket, packet, FCGI_Consts.HEADER_LEN)
       If bytes_received = -1 Or Len(packet) = 0 Then
            Exit Do
        End If

                FCGIPacketHeader.content = ""
            buf = decodePacketHeader(packet)
        If FCGIPacketHeader.contentLength > 0 Then
                BufLen = FCGIPacketHeader.contentLength
                Do While BufLen <> 0
'                        DoEvents
                    buf = ""
                    bytes_received = php_cgi_client.Recv(lngSocket, buf, BufLen)
                     If bytes_received <> -1 Then
                            FCGIPacketHeader.content = FCGIPacketHeader.content & Mid$(buf, 1, bytes_received)
                            BufLen = BufLen - bytes_received 'Len(buf)
                    Else
                        Exit Do
                    End If
                Loop
        End If
        If FCGIPacketHeader.paddingLength > 0 Then
                BufLen = FCGIPacketHeader.paddingLength
                  Do While BufLen <> 0
'                        DoEvents
                    buf = ""
                    bytes_received = php_cgi_client.Recv(lngSocket, buf, BufLen)
                    If bytes_received <> -1 Then
                        BufLen = BufLen - bytes_received 'Len(buf)
                    Else
                        Exit Do
                    End If
                Loop
        End If
            If (FCGIPacketHeader.type = FCGI_Consts.StdOut Or FCGIPacketHeader.type = FCGI_Consts.StdErr) Then
                 FCGIPacketHeader.response = FCGIPacketHeader.response & FCGIPacketHeader.content
           End If
            If (FCGIPacketHeader.type = FCGI_Consts.END_REQUEST) Then
'                        php_cgi_client.CloseAll
                 Open App.Path & "\fcgi_repsonse.txt" For Output As 1
                    Print #1, FCGIPacketHeader.response
                Close 1
                Exit Do
            End If
Loop While Len(packet) <> 0
                        php_cgi_client.CloseAll
    FCGI_Content_Received = True
                Dim sHeader As String, phpHeaders As String
'                Close rInfo.FileNum
                    
                rInfo.DataStr = FCGIPacketHeader.response
                If rInfo.DataStr <> "" And InStr(rInfo.DataStr, vbCrLf & vbCrLf) Then
                    phpHeaders = Mid$(rInfo.DataStr, 1, InStr(rInfo.DataStr, vbCrLf & vbCrLf) - 1)
                    rInfo.DataStr = Mid$(rInfo.DataStr, InStr(rInfo.DataStr, vbCrLf & vbCrLf) + 4)
                End If
                If rInfo.DataStr = "" Then
                    rInfo.DataStr = phpHeaders
                    phpHeaders = ""
                End If
                rInfo.TotalLength = Len(sHeader) + Len(rInfo.DataStr)
                sHeader = "HTTP/1.0 200 OK" & vbNewLine & _
                        "Server: " & ServerName & vbNewLine & _
                        "Content-Type: text/html" & vbNewLine & _
                        "Content-Length: " & Len(rInfo.DataStr) & _
                          vbNewLine & phpHeaders & _
                        vbNewLine & vbNewLine
             
            cServer.Send lngSocket, sHeader & rInfo.DataStr
End Sub
Private Sub php_cgi_client_OnClose(ByVal lngSocket As Long)
    FCGI_Content_Received = True
End Sub

Private Sub php_cgi_client_OnConnect(ByVal lngSocket As Long)
     lblRemoteHost.Caption = "Remote Host: " & php_cgi_client.GetRemoteHost(php_cgi_client.ConnectSocket)
     lblRemoteIP.Caption = "Remote IP: " & php_cgi_client.GetRemoteIP(php_cgi_client.ConnectSocket)
     lblRemotePort.Caption = "Remote Port: " & php_cgi_client.GetRemotePort(php_cgi_client.ConnectSocket)


   php_cgi_client.Send lngSocket, FCGIPacketHeader.content
End Sub

Private Sub php_cgi_client_OnError(ByVal lngRetCode As Long, ByVal strDescription As String)
    FCGI_Content_Received = True
Caption = "php_cgi_client_OnError " & strDescription
End Sub

Private Sub php_cgi_client_SendProgress(ByVal lngSocket As Long, ByVal bytesSent As Long, ByVal bytesRemaining As Long)
Caption = "php_cgi_client_SendProgress " & bytesSent & " " & bytesRemaining
End Sub
