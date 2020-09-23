VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database 2 HTML"
   ClientHeight    =   5325
   ClientLeft      =   150
   ClientTop       =   855
   ClientWidth     =   10215
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   10215
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   255
      Left            =   1680
      TabIndex        =   9
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      Format          =   22740993
      CurrentDate     =   36972
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3975
      Left            =   5760
      TabIndex        =   8
      Top             =   480
      Width           =   4215
      ExtentX         =   7435
      ExtentY         =   7011
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
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9000
      TabIndex        =   4
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton cmdWebPage 
      Caption         =   "WebPage"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7440
      TabIndex        =   3
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton cmdTable 
      Caption         =   "Table"
      Height          =   375
      Left            =   5880
      TabIndex        =   2
      Top             =   4800
      Width           =   975
   End
   Begin RichTextLib.RichTextBox Source 
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   3625
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":0442
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MyGrid 
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   3625
      _Version        =   393216
      BackColor       =   12648447
      FixedCols       =   0
      WordWrap        =   -1  'True
      AllowUserResizing=   1
      BandDisplay     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0C0&
      Height          =   4215
      Left            =   5640
      TabIndex        =   12
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label lblGames 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3240
      TabIndex        =   11
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Select a date:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Preview:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   5640
      TabIndex        =   7
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Source Code:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Database View:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   1935
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************
'* Project: Database 2 HTML                                                    *
'* Programmer: Gnu Kemist GnuKemist@yahoo.com                *
'* Version: 0.0.1 (as of March. 22, 2001)                                   *
'* Known Bugs: None                                                                *
'*************************************************

Private Sub cmdPreview_Click()
'# Opens the most recent HTML file created and displays it on _
    browser control.                                                                    #
WebBrowser1.Navigate App.Path & "\schedule.html"
End Sub

Private Sub cmdTable_Click()
'# Variables used to store the number of rows and column of grid. _
    This information tells us how many fields (columns) and how many _
    records (rows) were returned by the recordset.  In essence, this _
    part of the code can be used with any recordset, no matter what _
    tables are accessed!                                                                    #
Dim MaxCol, MaxRows As Integer

'# Clears Source before populating it with new HTML source code #
Source.Text = ""

rstGrid.MoveFirst

'# Stores number of rows and columns #
MaxCol = MyGrid.Cols
MaxRows = rstGrid.RecordCount
'# Initial part of any HTML page. You can be fancy here!!! #
Source.Text = "<html>" & vbCrLf
Source.Text = Source.Text & "<body>" & vbCrLf
Source.Text = Source.Text & "<center>" & vbCrLf
Source.Text = Source.Text & "<p><b><font color=#0000FF>Scheduled Games For: </font></b><font color=#008000><b>" & DTPicker1.Value & "</font></b></p>" & vbCrLf
Source.Text = Source.Text & "<table border=" & "0" & " cellpadding=" & "2" & " cellspacing=" & "0" & ">" & vbCrLf
Source.Text = Source.Text & "<tr>"
'# This loop creates a top(first) row and adds the names of all fields returned _
    by the recordset into separate columns #
For x = 0 To MaxCol - 1
    MyGrid.Row = 0  '# Row header of grid #
    MyGrid.Col = x
    Source.Text = Source.Text & "<td align=center bgcolor=#C0C0C0 style=" & Chr(34) _
    & "font-weight: bold" & Chr(34) & ">" & MyGrid.Text & "</td>" & vbCrLf
Next x
Source.Text = Source.Text & "</tr>"
'# This loop now populates an html table with all records returned by the recordset into _
    the body of the new table. #
For Row = 1 To MaxRows
    Source.Text = Source.Text & "<tr>" & vbCrLf
    For Col = 0 To MaxCol - 1
        MyGrid.Row = Row
        MyGrid.Col = Col
        '# This IF statement checkes whether the active row is an odd or even row. _
            If ODD, then the background is white for entire row; if EVEN, the background _
            is a light green! #
        If Row Mod 2 = 0 Then
            Source.Text = Source.Text & "<td bgcolor=#FFFFFF> " & MyGrid.Text & " </td>" & vbCrLf
        Else
            Source.Text = Source.Text & "<td bgcolor=#99FFCC> " & MyGrid.Text & " </td>" & vbCrLf
        End If
    Next Col
    Source.Text = Source.Text & "</tr>" & vbCrLf
Next Row

'# Closing tags for the HTML page #
Source.Text = Source.Text & "</table>" & vbCrLf
Source.Text = Source.Text & "</center>" & vbCrLf
Source.Text = Source.Text & "</body>" & vbCrLf
Source.Text = Source.Text & "</html>" & vbCrLf

'# Now that we have the source code, enable button to save source as a HTML file #
cmdWebPage.Enabled = True

End Sub

Private Sub cmdWebPage_Click()
'# Creates or Replaces a HTML file called schedule.html with the source code from Source #
Open App.Path & "\schedule.html" For Output As #1
Print #1, Source.Text
Close #1

'# Now that have saved the HTML file, allow user to view it #
cmdPreview.Enabled = True

End Sub

Private Sub DTPicker1_Change()
'# First, we clear the contents of Source #
Source.Text = ""

'# Closes the recordset if it is open #
If rstGrid.State <> adStateClosed Then
    rstGrid.Close
End If

'# Query database for games played on selected date #
Set cmdGrid.ActiveConnection = cnnHockey
cmdGrid.CommandText = "SELECT HomeTeam,VisitorTeam,Date,Time" _
    & " FROM SeasonSchedule WHERE Date=#" & DTPicker1.Value & "# ORDER BY Time"
rstGrid.CursorLocation = adUseClient
rstGrid.Open cmdGrid, , adOpenStatic, adLockBatchOptimistic

'# Binds grid to 'new' recordset #
Set MyGrid.DataSource = rstGrid
MyGrid.Refresh

'# Sub responsible for the layout of the grid #
GridLayout

'# Updates a label with the number of records returned by recordset #
lblGames.Caption = rstGrid.RecordCount & "  Games Scheduled"

'# Desables the WebPage and Preview buttons #
cmdWebPage.Enabled = False
cmdPreview.Enabled = False

End Sub

Private Sub DTPicker1_KeyPress(KeyAscii As Integer)
'# Prevents user from typing into the control #
KeyAscii = 0
End Sub

Private Sub Form_Load()
'# Used to query database for first and last day of season #
Dim cmd As New ADODB.Command
Dim rst As New ADODB.Recordset

'# Connection to the databasew #
Call Constructor

'# Query database for first and last day of season #
Set cmd.ActiveConnection = cnnHockey
cmd.CommandText = "SELECT MIN(Date) AS FirstDay, MAX(Date) AS LastDay FROM SeasonSchedule"
rst.CursorLocation = adUseClient
rst.Open cmd, , adOpenStatic, adLockBatchOptimistic

'# Assigns values to the DTPicker1 control.  The first date displayed is _
    the first day of the season BY DEFAULT                                      #
DTPicker1.MinDate = rst!FirstDay
DTPicker1.MaxDate = rst!LastDay
DTPicker1.Value = DTPicker1.MinDate

'# Closes objects created in the process above #
Set cmd = Nothing
rst.Close

'# Responsible for requerying the database everytime a new date is selected #
Call DTPicker1_Change

End Sub

Private Sub Form_Unload(Cancel As Integer)
'# Clean up after objects created #
Call Destructor
Set cmdGrid = Nothing
If rstGrid.State <> adStateClosed Then
    rstGrid.Close
End If

End Sub

Public Sub GridLayout()
'# Responsible for Changing the Column Widths #
MyGrid.ColWidth(0) = 1575
MyGrid.ColWidth(1) = 1575
MyGrid.ColWidth(2) = 800

End Sub

Private Sub mnuAbout_Click()
frmAbout.Show 1
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub
