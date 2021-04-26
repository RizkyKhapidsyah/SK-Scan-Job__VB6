VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   10605
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "index"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8260
      TabIndex        =   3
      Top             =   0
      Width           =   2295
   End
   Begin SHDocVwCtl.WebBrowser web 
      Height          =   6615
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   10575
      ExtentX         =   18653
      ExtentY         =   11668
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
   Begin VB.CommandButton Command1 
      Caption         =   "&Search"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   0
      Width           =   7215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag As Boolean
Dim iDoc As IHTMLDocument2
Dim el As Object
Dim total As String

Dim fso As New FileSystemObject
Dim sTxt As TextStream
Private Const conSwNormal = 1

Private Sub Command1_Click()
Command2.Enabled = False
Text1 = Replace(Text1, " ", "%20")
flag = False
web.Navigate "http://www.optioncarriere.com/cgi-bin/nw/search.cgi?s=" & Text1
Do While Not flag
    DoEvents
    DoEvents
    DoEvents
    DoEvents
Loop
getList
Launch
End Sub

Private Sub getList()
Dim obj As IHTMLElement
Dim oo As Object
Dim tbl As HTMLTable

Set iDoc = web.Document
' je vais localiser le tableau où se trouve les annonces
Set oo = iDoc.All.tags("FONT")

For scan = 0 To oo.length - 1
    If InStr(oo.Item(scan).innerText, "Les annonces") > 0 Then
        'ok, le tableau suivant contient les annonces
        posA = oo.Item(scan).sourceIndex
        Exit For
    End If
Next

For g = posA To iDoc.All.length - 1
    If iDoc.All.Item(g).tagName = "TABLE" Then
    posB = iDoc.All.Item(g).sourceIndex
    Exit For
    End If
Next

Set oo = iDoc.All.tags("TABLE")

For g = 0 To oo.length - 1
    If oo.Item(g).sourceIndex = posB Then
        'ok j'ai le tableau
        Set tbl = oo.Item(g)
        Exit For
    End If
Next



For l = 0 To tbl.All.length - 1
    If tbl.All.Item(l).tagName = "A" Then
        Set el = tbl.All.Item(l)
        strContent = strContent & "<a href='" & el.href & "'>" & el.innerText & "</a><br>" & vbCrLf
    End If
Next
Set sTxt = fso.CreateTextFile(App.Path & "\jobResult.htm")
header = "<html><head></head><body>" & vbCrLf
total = header & strContent & "</body></html>"
sTxt.write header & strContent & "</body></html>"
sTxt.Close
Set fso = Nothing

End Sub

Private Sub Launch()
web.Document.write (total & "<br>Ce fichier est disponible sous le nom jobResult.htm dans le répertoire de l'application")
Command2.Enabled = True
End Sub

Private Sub Command2_Click()
web.Navigate App.Path & "\jobResult.htm"

End Sub

Private Sub Form_Resize()
web.Width = Me.Width - 151
If (Me.Width - 151 - Command2.Width - 10 > 0) Then Command2.Left = Me.Width - 151 - Command2.Width - 10
If Me.Width - 3500 > 0 Then Text1.Width = Me.Width - 3500
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command1_Click


End Sub

Private Sub web_DocumentComplete(ByVal pDisp As Object, URL As Variant)
flag = True
End Sub

