VERSION 5.00
Begin VB.Form frmFormatting 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LynxGrid Demo"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdStartStop 
      Caption         =   "Start/Stop"
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   4740
      Width           =   1125
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4140
      Top             =   4290
   End
   Begin LynxGridTest.LynxGrid LynxGrid1 
      Height          =   1665
      Left            =   120
      TabIndex        =   1
      Top             =   510
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   2937
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColorHdr    =   255
      FocusRectMode   =   2
      FullRowSelect   =   0   'False
   End
   Begin LynxGridTest.LynxGrid LynxGrid2 
      Height          =   1665
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   2937
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColorHdr    =   12582912
      FocusRectMode   =   2
      FullRowSelect   =   0   'False
   End
   Begin LynxGridTest.LynxGrid LynxGrid3 
      Height          =   1665
      Left            =   120
      TabIndex        =   5
      Top             =   4740
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   2937
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FocusRectMode   =   2
      FullRowSelect   =   0   'False
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Example of ProgressBars:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   4
      Top             =   4410
      Width           =   3120
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Example of Column formatting:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   2
      Top             =   2250
      Width           =   3705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Example of per-cell formatting:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   3690
   End
End
Attribute VB_Name = "frmFormatting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdStartStop_Click()
    Timer1.Enabled = Not Timer1.Enabled
    
    If Timer1.Enabled Then
         With LynxGrid3
            .Redraw = False
            
            .CellProgressValue(0, 2) = 0
            .CellText(0, 3) = "Copying..."
            .CellProgressValue(1, 2) = 0
            .CellText(1, 3) = "Copying..."
            .CellProgressValue(2, 2) = 0
            .CellText(2, 3) = "Copying..."
        
            .Redraw = True
        End With
    End If
End Sub

Private Sub Form_Load()
    Dim lRow As Long

    With LynxGrid1
        .Redraw = False
        
        .AddColumn "BackColor", 1000
        .AddColumn "ForeColor", 1000
        .AddColumn "Alignment", 1000
        .AddColumn "Style", 1000
        
        '########################################################################################
        'AddItem returns the Index of the new Item making it simple to reference the new Cells
        
        lRow = .AddItem("Yellow" & vbTab & "Blue" & vbTab & "Left" & vbTab & "None")
        .CellBackColor(lRow, 0) = vbYellow
        .CellForeColor(lRow, 1) = vbBlue
        
        lRow = .AddItem("Blue" & vbTab & "Yellow" & vbTab & "Right" & vbTab & "Bold")
        .CellBackColor(lRow, 0) = vbBlue
        .CellForeColor(lRow, 1) = vbYellow
        .CellAlignment(lRow, 2) = lgAlignRightCenter
        .CellFontBold(lRow, 3) = True
        
        lRow = .AddItem("Green" & vbTab & "Red" & vbTab & "Centre" & vbTab & "Italic")
        .CellBackColor(lRow, 0) = vbGreen
        .CellForeColor(lRow, 1) = vbRed
        .CellAlignment(lRow, 2) = lgAlignCenterCenter
        .CellFontItalic(lRow, 3) = True
        
        .Redraw = True
    End With
    
    With LynxGrid2
        .Redraw = False
        
        '########################################################################################
        'Format masks used to format data as it is displayed
        .AddColumn "ItemCode", 1000
        .AddColumn "Due Date", 1500, , , "ddd dd mmm yy"
        .AddColumn "Price", 1000, lgAlignRightCenter, lgNumeric, ".00"
        
        '########################################################################################
        'AddItem can be called without any cell data to create an empty row
        .AddItem
        
        '########################################################################################
        'CellText allows any Cell to be indivudually set
        .CellText(0, 0) = "ABCD"
        .CellText(0, 1) = "10/02/06"
        .CellText(0, 2) = "45"
        
        .AddItem
        .CellText(1, 0) = "CVRT"
        .CellText(1, 1) = "22/04/06"
        .CellText(1, 2) = "70"
        
        .AddItem
        .CellText(2, 0) = "BCFF"
        .CellText(2, 1) = "01/05/06"
        .CellText(2, 2) = "52.60"
        
        .Redraw = True
    End With
    
    With LynxGrid3
        .Redraw = False
        
        .AddColumn "Filename", 1000
        .AddColumn "Size", 800, lgAlignRightCenter, lgNumeric
        .AddColumn "Progress", 1200, , lgProgressBar
        .AddColumn "Status", 1100, lgAlignCenterCenter
        
        .AddItem "abc.txt" & vbTab & "100KB"
        .AddItem "def.txt" & vbTab & "56KB"
        .AddItem "blue.bmp" & vbTab & "2400KB"
        
        .Redraw = True
    End With
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
    With LynxGrid3
        .Redraw = False
        
        'NOTE: CellProgressValue is an Integer between 0 and 100.
        
        If .CellProgressValue(0, 2) < 100 Then
            .CellProgressValue(0, 2) = .CellProgressValue(0, 2) + RandomInt(5, 10)
        ElseIf .CellText(0, 3) <> "Complete" Then
            .CellText(0, 3) = "Complete"
        End If
        
        If .CellProgressValue(1, 2) < 100 Then
            .CellProgressValue(1, 2) = .CellProgressValue(1, 2) + RandomInt(10, 20)
        ElseIf .CellText(1, 3) <> "Complete" Then
            .CellText(1, 3) = "Complete"
        End If
        
        If .CellProgressValue(2, 2) < 100 Then
            .CellProgressValue(2, 2) = .CellProgressValue(2, 2) + RandomInt(1, 5)
        ElseIf .CellText(2, 3) <> "Complete" Then
            .CellText(2, 3) = "Complete"
        End If
        
        .Redraw = True
    End With
End Sub


