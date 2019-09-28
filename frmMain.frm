VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "An FCGI client in Visual Basic. Supports PHP only."
   ClientHeight    =   5700
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   9975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdListen 
      Caption         =   "Listen"
      Height          =   370
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1450
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (dst As Any, ByVal iLen&)

Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Const FILE_ATTRIBUTE_DIRECTORY = &H10
' change this to your server name
Private Const ServerName As String = "Web Server Version 1.0.0"

' this project was designed for only one share
' change the path to the directory you want to share
Private Const PathShared As String = ""


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
Private start_time As Date, end_time As Date

Private WithEvents cServer  As cWinsock      'Server class
Attribute cServer.VB_VarHelpID = -1
Private WithEvents cClient  As cWinsock
Attribute cClient.VB_VarHelpID = -1
Private WithEvents php_cgi_client  As cWinsock
Attribute php_cgi_client.VB_VarHelpID = -1
Private rInfo As ConnectionInfo
Dim Dos As New DOSOutputs
Dim Header_Received As Boolean
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
