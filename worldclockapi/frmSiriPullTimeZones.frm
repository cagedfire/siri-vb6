VERSION 5.00
Begin VB.Form frmSiriPullTimeZones 
   Caption         =   "Siri - Time Zone Pull Request #1"
   ClientHeight    =   5355
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   8775
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtQuery 
      BeginProperty Font 
         Name            =   "@Adobe Gothic Std B"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1140
      TabIndex        =   1
      Top             =   720
      Width           =   6495
   End
   Begin VB.CommandButton cmdGetTime 
      Caption         =   "Process"
      BeginProperty Font 
         Name            =   "Roboto Condensed"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2580
      TabIndex        =   0
      Top             =   1800
      Width           =   3615
   End
   Begin VB.Label lblCurrentTime 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Roboto Condensed"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   480
      TabIndex        =   3
      Top             =   3000
      Width           =   7935
   End
   Begin VB.Label lblQueryBox 
      BackStyle       =   0  'Transparent
      Caption         =   "Query:"
      BeginProperty Font 
         Name            =   "Roboto"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmSiriPullTimeZones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public url As String
Public city As String

' A revised sassy commenting scheme.

' BOOM. Get that time.
Private Sub cmdGetTime_Click()

GetCity (txtQuery.Text)

End Sub

Public Function GetCity(query)
' ahaha no. shush i name it what i like.
Dim pointerChar As Integer

query = Replace$(query, ".", "")
query = Replace$(query, "?", "")
query = Replace$(query, ",", "")
query = Replace$(query, "!", "")
query = Replace$(query, "(", "")
query = Replace$(query, ")", "")

pointerChar = 0
city = "Sydney"

' nup. you have to process for all the eventuality's
query = Trim$(query)
pointerChar = InStr(1, query, "in", vbTextCompare)
query = Mid$(query, pointerChar + 3)

If query <> "" Then
    city = query
End If

city = StrConv(city, vbProperCase)

url = "http://api.openweathermap.org/data/2.5/weather?q=" & city & "&mode=xml&units=metric"

ProcessResponse

End Function

Public Function ProcessResponse()
Dim XMLFile As New DOMDocument30
Dim XMLNode As IXMLDOMNode
Dim gLong As String
Dim gLat As String
Dim coords As String
Dim dSuccess As Boolean
Dim time As Double
Dim country As String

' Preempt lag.
lblCurrentTime.Caption = "Loading..."

DoEvents

XMLFile.async = False

' DOWN THAT FILE.
dSuccess = XMLFile.Load(url)

If XMLFile.parseError.errorCode <> 0 Or dSuccess = False Then
    lblCurrentTime.Caption = "Failed to grab the time. Error: " & XMLFile.parseError.reason
    Exit Function
End If

' Grab the right nodules.
Set XMLNode = XMLFile.selectSingleNode("//current/city/coord")

If Not XMLNode Is Nothing Then
    gLong = XMLNode.Attributes.getNamedItem("lon").Text
    gLat = XMLNode.Attributes.getNamedItem("lat").Text
    coords = gLong & "," & gLat
    url = "http://api.timezonedb.com/?lat=" & gLat & "&lng=" & gLong & "&key=Q0Y8QVG2SGUT"
    ' DOWN THAT FILE.
    dSuccess = XMLFile.Load(url)
    
    If XMLFile.parseError.errorCode <> 0 Or dSuccess = False Then
        lblCurrentTime.Caption = "Failed to grab the time. Error: " & XMLFile.parseError.reason
        Exit Function
    End If
    
    Set XMLNode = XMLFile.selectSingleNode("//result")
    
    time = XMLNode.selectSingleNode("timestamp").Text
    country = XMLNode.selectSingleNode("countryCode").Text
    
    ' IN HH:MM:SS FORMAT YAS I DID IT GO GORDON.
    lblCurrentTime.Caption = "The time is currently " & _
        Format$( _
        DateAdd("s", time, DateSerial(1970, 1, 1)), _
        "hh:mm:ss AMPM" _
        ) & " in " & _
        city & _
        " (" & _
        country & ")."
    
    Exit Function
End If

lblCurrentTime = "Failed."

End Function



