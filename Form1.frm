VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "VB Mica & Fluent UI Demo"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   9375
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.PictureBox PicWV 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7545
      Left            =   0
      ScaleHeight     =   7545
      ScaleWidth      =   11775
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Initializing mica..."
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   1
         Top             =   270
         Width           =   1725
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private WithEvents WV As cWebView2
Attribute WV.VB_VarHelpID = -1

Private Sub Form_Load()
    Me.Visible = True
    Dim StartTime As Long, EndTime As Long
    StartTime = GetTickCount
    'Load WebView2 binding
    Set WV = New_c.WebView2
    If WV.BindTo(PicWV.hWnd) = 0 Then MsgBox "Couldn't initialize WebView-Binding": Exit Sub
    DoEvents
    Dim PixelStep As Integer, Pixels As Integer, R As Integer, G As Integer, B As Integer, TimeElapsedText As String
    PixelStep = 175
    CalculateMica PixelStep, Pixels, R, G, B
    EndTime = GetTickCount
    EndTime = GetTickCount
    TimeElapsedText = "Iterated over " & Pixels & " pixels with step " & PixelStep & ", took " & EndTime - StartTime & "ms to initialize mica."
    Label1.Visible = False
    WV.AddScriptToExecuteOnDocumentCreated "function SubmitButtonClicked(){vbH().RaiseMessageEvent('SubmitButtonClick',document.getElementById('textbox').value)}"
    WV.AddScriptToExecuteOnDocumentCreated "function SecondaryButtonClicked(){vbH().RaiseMessageEvent('SecondaryButtonClick','')}"
    WV.NavigateToString "<!DOCTYPE html> <html> <head> <meta charset=""UTF-8""> <meta name=""viewport"" content=""width=device-width, initial-scale=1.0"">" & _
        "<title>" & Me.Caption & "</title>" & _
        "<style> .mica {width: 100%; height: 100%; padding: 0 1rem; position: fixed; backdrop-filter: blur(40px); background-color: rgb(" & R & ", " & G & ", " & B & ");}" & _
        "body {margin: 0; padding: 0; font-size: 87.5%}" & _
        ".btn-info {color: #fff; background-color: #0078d4; border: 1px solid #2b73aa; transition: filter 0.2s ease-out;}" & _
        ".btn-secondary {color: #000; background-color: #f8fafd; border: 1px solid #c4cad0; transition: filter 0.2s ease-out;}" & _
        ".btn:hover {filter: brightness(1.1);}" & _
        ".btn {cursor: pointer; display: inline-block; text-align: center; vertical-align: middle; user-select: none; padding: .5rem; font-size: inherit; border-radius: .25rem; }" & _
        "input {padding: 0.5rem 0.5rem; line-height: 1; box-sizing: border-box; border-radius: 4px; border: 1px solid #e1e1e3; background-color: #fafbfd; color: #000; border-bottom: 1px solid #0078d4;}" & _
        "</style></head> <body>" & _
        "<div class=""mica"">" & _
        "<h1>VB Mica & Fluent UI Demo</h1>" & _
        "<p>Maybe a brand new way to create GUIs in VB6, powered by Microsoft Edge WebView2.<br>" & _
        "The GUI is created using HTML, JS and CSS, and can call native VB code. The background color is extracted from the current wallpaper.</p>" & _
        "<p id=""loadtime""></p>" & _
        "<p><input id=""textbox"" type=""text"" placeholder=""Textbox""></p>" & _
        "<button onclick=""SubmitButtonClicked()"" type=""button"" class=""btn btn-info"">Submit button</button> &nbsp;" & _
        "<button onclick=""SecondaryButtonClicked()"" type=""button"" class=""btn btn-secondary"">Secondary button</button>" & _
        "</div></body>"
    WV.ExecuteScript "document.getElementById('loadtime').innerHTML = '" & TimeElapsedText & "';"
End Sub

Private Sub WV_JSMessage(ByVal sMsg As String, ByVal sMsgContent As String, oJSONContent As cCollection)
  Select Case sMsg
    Case "SubmitButtonClick"
        MsgBox "Submit button was clicked. Textbox's value is: " & sMsgContent
    Case "SecondaryButtonClick"
        MsgBox "Secondary button was clicked."
  End Select
End Sub

Private Sub Form_Resize()
    PicWV.Height = Me.Height
    PicWV.Width = Me.Width
    If Not WV Is Nothing Then WV.SyncSizeToHostWindow
End Sub

