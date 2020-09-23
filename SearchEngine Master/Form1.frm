VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   Caption         =   "SearchEngineMaster"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11025
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   11025
   StartUpPosition =   3  'Windows Default
   Begin Proyecto1.InternetFile web 
      Left            =   7200
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox SearchBox 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "VB code"
      Top             =   360
      Width           =   4935
   End
   Begin SHDocVwCtl.WebBrowser Browser 
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   10095
      ExtentX         =   17806
      ExtentY         =   11033
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label Label1 
      Caption         =   "Type your KeyWords here:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   50
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------'
'                      SEARCH ENGINE MASTER by Ricardo Arias                      '
'---------------------------------------------------------------------------------'
Private Sub Command1_Click()
Open App.Path & "\Results.html" For Output As #2
Print #2, "<html><head></head><title>SearchEngineMaster</title><body background='back.jpg' >"
Print #2, "<center><h1><font color=#ff0000>SearchEngineMaster</font></h1><p>"
Print #2, "General search results for " & Date & " : " & Time & "</center><p>"
Print #2, "<p><hr><p>"
GetForWebResults 'go for the Search Engines results!!

End Sub

Sub ClosePage() 'this continues and ends the web page
Print #2, "<p><hr><p>"
Print #2, "realtime generated page using SearchEngineMaster by Ricardo Arias<br>"
Print #2, "<a href=" & Chr(34) & "mailto:ricadoarias@yahoo.com" & Chr(34) & ">ricardoarias@yahoo.com</a><p><font color=#ff0000 size=+3><b>dont forget to vote!</b></font>"
Print #2, "</body></html>"
Close #2
Browser.Navigate App.Path & "\Results.html"
End Sub

Private Sub Form_Load()
web.ConnectType = Direct 'configuring the Downloader
web.LocalFile = App.Path & "\data.000" ' just for hidding to the users eyes but are html
Me.Width = Screen.Width
Me.Height = Screen.Height
Me.Top = 0
Me.Left = 0
Browser.Width = Screen.Width - 150
Browser.Navigate App.Path & "\searchenginemaster.html"
End Sub

Private Sub web_DownloadComplete()
ParseIt 'when the file is on your PC has to been readed and parsed to get results...
End Sub

Sub GetForWebResults() 'This one goes one by one in the search engines
Static WhichOne As Integer ' this static let you go one by one
WhichOne = WhichOne + 1
Select Case WhichOne 'depending on the value determine which engine is used

Case 1
web.Url = "http://www.google.com/search?hl=en&q=" & SearchBox.Text
Case 2
web.Url = "http://search.yahoo.com/bin/search?p=" & SearchBox.Text
Case 3
web.Url = "http://search.msn.com/results.asp?RS=CHECKED&FORM=MSNH&v=1&q=" & SearchBox.Text
Case Else
ClosePage
WhichOne = 0
End Select

web.StartDownload 'start downloading the search engine results
End Sub

Sub ParseIt()
On Error Resume Next ' if something fails
Dim Data As String, Engine As String, Result As String, Description As String
Dim Position1 As Long, Position2 As Long, Position3 As Long, Position4 As Long
Open App.Path & "\data.000" For Binary Access Read As #1
Data = Space$(LOF(1))
Get #1, , Data
Close #1
Kill App.Path & "\data.000"


If InStr(Data, "<title>Google Search:") <> 0 Then ' here we detect if the result is from Google
Engine = "Google"
ElseIf InStr(Data, "<title>Yahoo! Search Results") <> 0 Then ' Yahoo
Engine = "Yahoo"
ElseIf InStr(Data, ">MSN Search:") <> 0 Then ' MSN
Engine = "MSN"
End If
Select Case Engine

Case "Google"
Print #2, "<center><img src='google.gif'></center>"
Print #2, "<p><h3>Results from Google</h3><p>"
For i = 1 To 10 ' Check one by one the ocurrences
DoEvents
Position1 = InStr(Position1 + 1, Data, "<p><a href")
Position2 = InStr(Position1 + 1, Data, "</a>")
Result = Mid(Data, Position1 + 3, (Position2 - 1) - (Position1 + 2))
Print #2, "<p>" & Str(i) & ".- " & Result & "</a><br>"
Position3 = InStr(Position2 + 1, Data, "</span>")
Position4 = InStr(Position3 + 1, Data, "<br>")
Description = Mid(Data, Position3 + 7, (Position4 - 1) - (Position3 + 7))
Print #2, Description & "<p>"
Next i

Case "Yahoo"
Print #2, "</a><p><hr><p><center><img src='yahoo.gif'></center><p>"
Print #2, "<h3>Results from Yahoo</h3><p>"
For i = 1 To 20
Position1 = InStr(Position1 + 1, Data, "<li><big>")
Position1 = InStr(Position1 + 1, Data, "*http")
Position2 = InStr(Position1 + 1, Data, ">")
Result = Mid(Data, Position1 + 1, (Position2 - 1) - (Position1 + 1))
Print #2, "<p>" & Str(i) & ".- <a href=" & Chr(34) & Result & Chr(34) & ">" & Result & "</a><br>"
Position3 = InStr(Position2 + 1, Data, "</a>")
Description = Mid(Data, Position2 + 1, (Position3 - 1) - (Position2))
Print #2, Description & "<p>"
Next i
End Select


GetForWebResults 'goes again for the other page

End Sub

Private Sub web_DownloadProgress(lBytesRead As Long)
DoEvents 'just for dont freeze during download
End Sub
