VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmWrap 
   Caption         =   "LynxGrid Demo"
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   9480
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtEdit 
      BorderStyle     =   0  'None
      Height          =   765
      Left            =   6900
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   4050
      Visible         =   0   'False
      Width           =   2325
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7920
      Top             =   150
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483643
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Wrap.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Wrap.frx":5F9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Wrap.frx":7D5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Wrap.frx":D54E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Wrap.frx":137E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Wrap.frx":19A82
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Wrap.frx":1F6A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Wrap.frx":2563E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Wrap.frx":268C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Wrap.frx":2C4E2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin LynxGridTest.LynxGrid LynxGrid 
      Height          =   4875
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   9315
      _extentx        =   16431
      _extenty        =   8599
      gridlines       =   0   'False
      allowuserresizing=   1
      autosizerow     =   -1  'True
      editable        =   -1  'True
      font            =   "Wrap.frx":2C644
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   "Label1"
      Height          =   30
      Left            =   60
      TabIndex        =   1
      Top             =   5100
      Visible         =   0   'False
      Width           =   825
   End
End
Attribute VB_Name = "frmWrap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim lRow As Long

    With LynxGrid
        .Redraw = False
        
        .EditTrigger = lgEnterKey
        .FullRowSelect = False
        
        .ExpandRowImage = 10
        .FocusRectColor = &HFF0000
        .FocusRectMode = lgRow
        .FocusRectStyle = lgFRHeavy
        
        .RowHeightMin = 315
        
        .ImageList = ImageList1
        
        .AddColumn "", 350, , , , , lgAlignCenterTop
        .AddColumn "", 350, , , , , lgAlignCenterTop
        .AddColumn "", 350, lgAlignCenterTop, lgBoolean
        .AddColumn "Text", 4000, lgAlignLeftTop, , , , lgAlignRightTop, True
        .AddColumn "Schedule", 1950, lgAlignLeftTop, lgDate, , , lgAlignLeftTop
        .AddColumn "Created", 1950, lgAlignLeftTop, lgDate, , , lgAlignRightTop
        
        .BindControl 3, txtEdit
        
        lRow = .AddItem()
        .CellImage(lRow, 0) = 6
        .CellImage(lRow, 1) = 2
        .CellChecked(lRow, 2) = True
        .CellText(lRow, 3) = "Meeting with Bank Manager." & vbCrLf & "Meeting with Mortgage Advisor"
        .CellForeColor(lRow, 3) = vbGreen
        .CellText(lRow, 4) = "07.04.2006 10:00"
        .CellText(lRow, 5) = "05.04.2006 20:02"
        
        lRow = .AddItem()
        .CellImage(lRow, 0) = 7
        .CellImage(lRow, 1) = 3
        .CellText(lRow, 3) = "Appointment with Dentist." & vbCrLf & "Collect Train Ticket." & vbCrLf & "Post utility bills to Solicitor."
        .CellForeColor(lRow, 3) = vbRed
        .CellImage(lRow, 4) = 8
        
        .CellText(lRow, 4) = "08.04.2006 09:15"
        .CellText(lRow, 5) = "05.04.2006 20:03"
        
        lRow = .AddItem()
        .CellImage(lRow, 0) = 6
        .CellImage(lRow, 1) = 4
        .CellText(lRow, 3) = "Book Taxi for evening." & vbCrLf & "Meet Peter at Cinema"
        .CellText(lRow, 4) = "09.04.2006 12:15"
        .CellText(lRow, 5) = "05.04.2006 20:05"
                
        lRow = .AddItem()
        .CellImage(lRow, 0) = 7
        .CellChecked(lRow, 2) = True
        .CellText(lRow, 3) = "Visit Tourist Office to collect brochures." & vbCrLf & "Pick-up suit from cleaners." & vbCrLf & "Get groceries!"
        .CellForeColor(lRow, 3) = vbGreen
        
        .CellImage(lRow, 4) = 9
        .CellText(lRow, 4) = "10.04.2006 09:35"
        .CellText(lRow, 5) = "06.04.2006 09:20"
        
        lRow = .AddItem()
        .CellImage(lRow, 0) = 6
        .CellText(lRow, 3) = "Backup Digital Photos."
        .CellText(lRow, 4) = "11.04.2006 09:15"
        .CellText(lRow, 5) = "06.04.2006 09:23"
        
        lRow = .AddItem()
        .CellImage(lRow, 0) = 4
        .CellText(lRow, 3) = "Catalog DVD samples" & vbCrLf & "Print Contents pages" & vbCrLf & "Duplicate Orders"
        .CellText(lRow, 4) = "12.04.2006 09:15"
        .CellText(lRow, 5) = "06.04.2006 09:30"
        
        'For lRow = 5 To 100
        '    .AddItem
        '    .CellText(lRow, 3) = "Text for test " & lRow
        'Next lRow
        
        .Redraw = True
    End With
End Sub


Private Sub Form_Resize()
    LynxGrid.Height = Me.Height - 795
    LynxGrid.Width = Me.Width - 285
End Sub


Private Sub LynxGrid_CellImageClick(ByVal Row As Long, ByVal Col As Long)
    If Col = 4 Then
        MsgBox "You clicked on a Cell Image in Column 4", vbInformation
    End If
End Sub

Private Sub LynxGrid_RequestEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 3 Then
        txtEdit.Text = LynxGrid.CellText(Row, Col)
    End If
End Sub


Private Sub LynxGrid_RequestUpdate(ByVal Row As Long, ByVal Col As Long, NewValue As String, Cancel As Boolean)
    If Col = 3 Then
        NewValue = txtEdit.Text
    End If
End Sub


