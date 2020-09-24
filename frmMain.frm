VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "LinkCrawl"
   ClientHeight    =   5295
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4590
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   4590
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   5040
      Width           =   4590
      _ExtentX        =   8096
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cmbAddress 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Text            =   "Type Url Here"
      Top             =   600
      Width           =   4335
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4590
      _ExtentX        =   8096
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "images"
      DisabledImageList=   "images"
      HotImageList    =   "images"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbCrawl"
            Object.ToolTipText     =   "Crawl"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbSave"
            Object.ToolTipText     =   "Save List"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbClear"
            Object.ToolTipText     =   "Clear List"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList images 
      Left            =   1200
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0D1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":103A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.ListBox Links 
      Height          =   3570
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   4335
   End
   Begin SHDocVwCtl.WebBrowser web 
      Height          =   3375
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   4335
      ExtentX         =   7646
      ExtentY         =   5953
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
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   225
      Left            =   120
      TabIndex        =   4
      Top             =   4800
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnucrawl 
         Caption         =   "&Crawl"
      End
      Begin VB.Menu mnusep 
         Caption         =   "-"
      End
      Begin VB.Menu mnusave 
         Caption         =   "&Save Links"
      End
      Begin VB.Menu mnuclear 
         Caption         =   "&Clear List"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
   Begin VB.Menu mnupopup 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnucopy 
         Caption         =   "&Copy Link"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbAddress_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
web.Navigate2 cmbAddress.Text
End If
End Sub


Private Sub Form_Load()

ProgressBar1.Min = 0
ProgressBar1.Max = 1
End Sub



Private Sub Links_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button <> 2 Then
Exit Sub
Else
Call PopupMenu(mnupopup)
End If

End Sub

Private Sub mnuAbout_Click()
MsgBox "Link Crawl Program by xTrin", vbInformation, "About"
End Sub

Private Sub mnuclear_Click()
Links.Clear

End Sub

Private Sub mnucopy_Click()
    Clipboard.Clear
    Clipboard.SetText Links.List(Links.ListIndex)
End Sub

Private Sub mnusave_Click()
Links.AddItem vbCrLf, 0
Links.AddItem Date & " - " & cmbAddress.Text, 1
For i = 0 To Links.ListCount - 1
  Open App.Path & "\links.dat" For Append As #1
  Print #1, Links.List(i)
  Close #1
Next i
End Sub

Private Sub Web_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
If ProgressMax <> 0 And Progress > -1 And Progress <= ProgressMax Then
    ProgressBar1.Value = Progress / ProgressMax
End If
If ProgressBar1.Value = 1 Then
   ProgressBar1.Value = 0
Else
   ProgressBar1.Visible = True
End If
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case Is = "tbCrawl":
Links.Clear
Call Crawl
StatusBar1.Panels.Item(1).Text = Links.ListCount & " Link(s)"
Case Is = "tbSave":
Call mnusave_Click
Case Is = "tbClear":
Links.Clear

End Select

End Sub

Private Sub Crawl()

Dim i As Long
Dim xDoc As IHTMLDocument2
Dim xMainWnd As IHTMLWindow2
Dim xWnd As IHTMLWindow2
Dim xFrames As IHTMLFramesCollection2
Dim yFrames As Long

'set xDoc equal to the Webpage in the browser control
Set xDoc = web.Document
'set xWnd equal to the main window if there are frames
Set xWnd = xDoc.parentWindow
'set xMainWnd equal to the top frame
Set xMainWnd = xWnd.Top
'gather all the frames
Set xFrames = xMainWnd.frames
'get number of frames
yFrames = xFrames.length
'loop through the frames if there are any and gather the links
If yFrames > 1 Then

  For i = 0 To yFrames - 1
  
   Dim xFrameDoc As IHTMLDocument2
   Dim xFrameWnd As IHTMLWindow2
   
   Set xFrameWnd = xFrames.Item(i)
   Set xFrameDoc = xFrameWnd.Document
   
   Call GetLinks(xFrameDoc)
   
   Set xFrameDoc = Nothing
   Set xFrameWnd = Nothing
   
  Next i

Else
'if there are no frames gather the links
Call GetLinks(xDoc)

End If

Set xFrames = Nothing
Set xMainWnd = Nothing
Set xWnd = Nothing
Set xDoc = Nothing

   
End Sub

Private Sub GetLinks(doc As IHTMLDocument2)

On Error GoTo GetLinks_Err

Dim i, j, yElements As Long
Dim xElements As IHTMLElementCollection
Dim xElement As IHTMLElement
Dim yPos As Integer
Dim xUrl As String

'get the <body> </body> section of the page
Set xElement = doc.body
Set xElements = xElement.All

yElements = xElements.length

For i = 0 To yElements - 1

  Dim sTag As String
  
  Set xElement = xElements.Item(i)
  
'Check every "anchor" for file type
        sTag = UCase(xElement.tagName)
        
        If sTag = "A" Then
            Dim xAnchor As IHTMLAnchorElement
            Dim shref As String
            
            Set xAnchor = xElement
            shref = xAnchor.href
            sUrl = ParseURL(shref)
            Links.AddItem sUrl
                
            Set xAnchor = Nothing
            
        End If

Next i

GetLinks_Cleanup:
    
    Set xElement = Nothing
    Set xElements = Nothing
    
    Exit Sub

GetLinks_Err:
MsgBox "Error", vbCritical
GoTo GetLinks_Cleanup


End Sub

Private Function ParseURL(url As String)

    Dim nTok As Integer
    Dim sTempUrl As String
    Dim sProt As String
    
    ParseURL = url
    
    nTok = InStrRev(url, "http://")
    If nTok Then
        'Check for URL's embedded in other URLS
        url = Right(url, Len(url) - nTok - 6)
        sProt = "http://"
    Else
        nTok = InStrRev(url, "ftp://")
        If nTok Then
            url = Right(url, Len(url) - nTok - 5)
            sProt = "ftp://"
        End If
    End If

    nTok = InStrRev(url, "http://")
    If nTok Then
        ParseURL = Right(url, Len(url) - nTok + 1)
        Exit Function
    Else
        nTok = InStrRev(url, "ftp://")
        If nTok Then
            ParseURL = Right(url, Len(url) - nTok + 1)
            Exit Function
        End If
    End If
    
    ParseURL = sProt & url
    
End Function



