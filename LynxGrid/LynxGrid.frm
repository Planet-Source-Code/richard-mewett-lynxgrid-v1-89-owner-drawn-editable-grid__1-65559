VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "LynxGrid Tester © 2006 Richard Mewett"
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8100
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCellWrap 
      Caption         =   "Cell Wrap..."
      Height          =   345
      Left            =   10590
      TabIndex        =   26
      Top             =   6570
      Width           =   1215
   End
   Begin VB.CommandButton cmdBinding 
      Caption         =   "Binding..."
      Height          =   345
      Left            =   10590
      TabIndex        =   25
      Top             =   6150
      Width           =   1215
   End
   Begin VB.CommandButton cmdFormatting 
      Caption         =   "Formatting..."
      Height          =   345
      Left            =   10590
      TabIndex        =   24
      Top             =   5730
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   7470
      TabIndex        =   50
      Top             =   5280
      Width           =   3075
      Begin VB.CommandButton cmdRangeFormat 
         Caption         =   "Range Format"
         Height          =   345
         Left            =   1740
         TabIndex        =   22
         Top             =   1410
         Width           =   1245
      End
      Begin VB.CommandButton cmdChangedCells 
         Caption         =   "Changed Cells..."
         Height          =   525
         Left            =   1740
         TabIndex        =   23
         Top             =   2100
         Width           =   1245
      End
      Begin VB.CommandButton cmdSort 
         Caption         =   "Sort by Name"
         Height          =   345
         Left            =   1740
         TabIndex        =   21
         Top             =   1020
         Width           =   1245
      End
      Begin VB.CommandButton cmdAddItems 
         Caption         =   "Add Items"
         Height          =   345
         Left            =   1740
         TabIndex        =   19
         Top             =   240
         Width           =   1245
      End
      Begin VB.CommandButton cmdRemoveItem 
         Caption         =   "RemoveItem"
         Height          =   345
         Left            =   1740
         TabIndex        =   20
         Top             =   630
         Width           =   1245
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Items"
         Height          =   195
         Left            =   90
         TabIndex        =   61
         Top             =   2310
         Width           =   375
      End
      Begin VB.Label lblItemCount 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   840
         TabIndex        =   62
         Top             =   2250
         Width           =   795
      End
      Begin VB.Label lblMouseRowCol 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   840
         TabIndex        =   53
         Top             =   450
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mouse"
         Height          =   195
         Left            =   90
         TabIndex        =   52
         Top             =   510
         Width           =   480
      End
      Begin VB.Label lblSelectedCount 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   840
         TabIndex        =   58
         Top             =   1530
         Width           =   795
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Selected"
         Height          =   195
         Left            =   90
         TabIndex        =   57
         Top             =   1590
         Width           =   630
      End
      Begin VB.Label lblCheckedCount 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   840
         TabIndex        =   60
         Top             =   1890
         Width           =   795
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Checked"
         Height          =   195
         Left            =   90
         TabIndex        =   59
         Top             =   1950
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Row/Col:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   51
         Top             =   180
         Width           =   810
      End
      Begin VB.Label lblRowCol 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   840
         TabIndex        =   55
         Top             =   840
         Width           =   795
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Current"
         Height          =   195
         Left            =   90
         TabIndex        =   54
         Top             =   930
         Width           =   510
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Count:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   56
         Top             =   1290
         Width           =   570
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8820
      Top             =   330
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9420
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483643
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LynxGrid.frx":0000
            Key             =   "MALE1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LynxGrid.frx":059A
            Key             =   "MALE2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LynxGrid.frx":0B34
            Key             =   "MALE3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LynxGrid.frx":10CE
            Key             =   "FEMALE1"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LynxGrid.frx":1668
            Key             =   "FEMALE2"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "LynxGrid.frx":1C02
            Key             =   "FEMALE3"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtHelp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   645
      Left            =   90
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   63
      Text            =   "LynxGrid.frx":219C
      Top             =   7410
      Width           =   7305
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      Height          =   630
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   11820
      TabIndex        =   27
      Top             =   0
      Width           =   11880
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Single File solution"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   1
         Left            =   1260
         TabIndex        =   30
         Top             =   300
         Width           =   1620
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Owner-drawn editable Grid UserControl"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   1
         Left            =   1260
         TabIndex        =   29
         Top             =   60
         Width           =   3315
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   28
         Top             =   60
         Width           =   1035
      End
   End
   Begin VB.Frame fraProperties 
      Caption         =   "Properties"
      Height          =   4635
      Left            =   7470
      TabIndex        =   31
      Top             =   630
      Width           =   4365
      Begin VB.CheckBox chkApplySelectionToImages 
         Caption         =   "Apply Selection To Images"
         Height          =   195
         Left            =   1920
         TabIndex        =   14
         Top             =   2760
         Width           =   2295
      End
      Begin VB.CheckBox chkHideSelection 
         Caption         =   "HideSelection"
         Height          =   195
         Left            =   90
         TabIndex        =   9
         Top             =   2118
         Width           =   1425
      End
      Begin VB.CheckBox chkAlphaBlendSelection 
         Caption         =   "Alpha Blend Selection"
         Height          =   195
         Left            =   1920
         TabIndex        =   13
         Top             =   2520
         Width           =   1935
      End
      Begin VB.ComboBox cboThemeStyle 
         Height          =   315
         Left            =   1470
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   4230
         Width           =   1695
      End
      Begin VB.ComboBox cboThemeColor 
         Height          =   315
         Left            =   1470
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   3840
         Width           =   1695
      End
      Begin VB.ComboBox cboFocusRectMode 
         Height          =   315
         Left            =   1470
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   3060
         Width           =   1005
      End
      Begin VB.CheckBox chkFullRowSelect 
         Caption         =   "FullRowSelect"
         Height          =   195
         Left            =   90
         TabIndex        =   7
         Top             =   1656
         Width           =   1425
      End
      Begin VB.ComboBox cboFocusRectStyle 
         Height          =   315
         Left            =   1470
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   3420
         Width           =   1005
      End
      Begin VB.CheckBox chkGridLines 
         Caption         =   "GridLines"
         Height          =   195
         Left            =   90
         TabIndex        =   8
         Top             =   1887
         Width           =   1425
      End
      Begin VB.CheckBox chkHotHeaderTracking 
         Caption         =   "HotHeaderTracking"
         Height          =   195
         Left            =   90
         TabIndex        =   10
         Top             =   2349
         Width           =   1815
      End
      Begin VB.CheckBox chkScrollTrack 
         Caption         =   "ScrollTrack"
         Height          =   195
         Left            =   90
         TabIndex        =   12
         Top             =   2820
         Width           =   1425
      End
      Begin VB.CheckBox chkEditable 
         Caption         =   "Editable"
         Height          =   195
         Left            =   90
         TabIndex        =   6
         Top             =   1425
         Width           =   1425
      End
      Begin VB.CheckBox chkColumnSort 
         Caption         =   "ColumnSort"
         Height          =   195
         Left            =   90
         TabIndex        =   4
         Top             =   963
         Width           =   1425
      End
      Begin VB.CheckBox chkCheckBoxes 
         Caption         =   "CheckBoxes"
         Height          =   195
         Left            =   90
         TabIndex        =   1
         Top             =   270
         Width           =   1425
      End
      Begin VB.CheckBox chkMultiSelect 
         Caption         =   "MultiSelect"
         Height          =   195
         Left            =   90
         TabIndex        =   11
         Top             =   2580
         Width           =   1425
      End
      Begin VB.CheckBox chkColumnResize 
         Caption         =   "ColumnResize"
         Height          =   195
         Left            =   90
         TabIndex        =   3
         Top             =   732
         Width           =   1425
      End
      Begin VB.CheckBox chkDisplayEllipsis 
         Caption         =   "DisplayEllipsis"
         Height          =   195
         Left            =   90
         TabIndex        =   5
         Top             =   1194
         Width           =   1425
      End
      Begin VB.CheckBox chkColumnDrag 
         Caption         =   "ColumnDrag"
         Height          =   195
         Left            =   90
         TabIndex        =   2
         Top             =   501
         Width           =   1425
      End
      Begin VB.Label lblThemeStyle 
         AutoSize        =   -1  'True
         Caption         =   "ThemeStyle"
         Height          =   195
         Left            =   90
         TabIndex        =   49
         Top             =   4260
         Width           =   840
      End
      Begin VB.Label lblThemeColor 
         AutoSize        =   -1  'True
         Caption         =   "ThemeColor"
         Height          =   195
         Left            =   90
         TabIndex        =   48
         Top             =   3900
         Width           =   855
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "FocusRectMode"
         Height          =   195
         Left            =   90
         TabIndex        =   46
         Top             =   3120
         Width           =   1185
      End
      Begin VB.Label lblColor 
         AutoSize        =   -1  'True
         Caption         =   "FocusRectColor"
         Height          =   195
         Index           =   5
         Left            =   1920
         TabIndex        =   42
         Top             =   1860
         Width           =   1140
      End
      Begin VB.Label lblViewColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   3480
         TabIndex        =   43
         Top             =   1830
         Width           =   795
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "FocusRectStyle"
         Height          =   195
         Left            =   90
         TabIndex        =   47
         Top             =   3480
         Width           =   1125
      End
      Begin VB.Label lblViewColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   3480
         TabIndex        =   41
         Top             =   1530
         Width           =   795
      End
      Begin VB.Label lblColor 
         AutoSize        =   -1  'True
         Caption         =   "ForeColorSel"
         Height          =   195
         Index           =   4
         Left            =   1920
         TabIndex        =   40
         Top             =   1560
         Width           =   900
      End
      Begin VB.Label lblColor 
         AutoSize        =   -1  'True
         Caption         =   "ForeColorEdit"
         Height          =   195
         Index           =   3
         Left            =   1920
         TabIndex        =   38
         Top             =   1260
         Width           =   945
      End
      Begin VB.Label lblViewColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   3480
         TabIndex        =   39
         Top             =   1230
         Width           =   795
      End
      Begin VB.Label lblViewColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   3480
         TabIndex        =   35
         Top             =   630
         Width           =   795
      End
      Begin VB.Label lblColor 
         AutoSize        =   -1  'True
         Caption         =   "BackColorEdit"
         Height          =   195
         Index           =   1
         Left            =   1920
         TabIndex        =   34
         Top             =   660
         Width           =   1005
      End
      Begin VB.Label lblColor 
         AutoSize        =   -1  'True
         Caption         =   "BackColorBkg"
         Height          =   195
         Index           =   0
         Left            =   1920
         TabIndex        =   32
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label lblViewColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   3480
         TabIndex        =   33
         Top             =   330
         Width           =   795
      End
      Begin VB.Label lblColor 
         AutoSize        =   -1  'True
         Caption         =   "BackColorSel"
         Height          =   195
         Index           =   2
         Left            =   1920
         TabIndex        =   36
         Top             =   960
         Width           =   960
      End
      Begin VB.Label lblViewColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   3480
         TabIndex        =   37
         Top             =   930
         Width           =   795
      End
      Begin VB.Label lblColor 
         AutoSize        =   -1  'True
         Caption         =   "GridColor"
         Height          =   195
         Index           =   6
         Left            =   1920
         TabIndex        =   44
         Top             =   2175
         Width           =   645
      End
      Begin VB.Label lblViewColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   3480
         TabIndex        =   45
         Top             =   2130
         Width           =   795
      End
   End
   Begin LynxGridTest.LynxGrid LynxGrid 
      Height          =   6675
      Left            =   90
      TabIndex        =   0
      Top             =   690
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   11774
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AlphaBlendSelection=   -1  'True
      FocusRectMode   =   2
      AllowUserResizing=   1
      Checkboxes      =   -1  'True
      ColumnDrag      =   -1  'True
      ColumnSort      =   -1  'True
      Editable        =   -1  'True
   End
   Begin VB.Label lblExamples 
      AutoSize        =   -1  'True
      Caption         =   "Examples:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   10590
      TabIndex        =   64
      Top             =   5400
      Width           =   870
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub SetColors()
    With LynxGrid
        lblViewColor(0).BackColor = .BackColorBkg
        lblViewColor(1).BackColor = .BackColorEdit
        lblViewColor(2).BackColor = .BackColorSel
        lblViewColor(3).BackColor = .ForeColorEdit
        lblViewColor(4).BackColor = .ForeColorSel
        lblViewColor(5).BackColor = .FocusRectColor
        lblViewColor(6).BackColor = .GridColor
   End With
End Sub

Private Sub cboFocusRectMode_Click()
    LynxGrid.FocusRectMode = cboFocusRectMode.ListIndex
End Sub

Private Sub cboFocusRectStyle_Click()
    LynxGrid.FocusRectStyle = cboFocusRectStyle.ListIndex
End Sub

Private Sub cboThemeColor_Click()
    LynxGrid.ThemeColor = cboThemeColor.ListIndex
    SetColors
End Sub

Private Sub cboThemeStyle_Click()
    LynxGrid.ThemeStyle = cboThemeStyle.ListIndex
End Sub


Private Sub chkAlphaBlendSelection_Click()
    LynxGrid.AlphaBlendSelection = chkAlphaBlendSelection.Value
End Sub

Private Sub chkApplySelectionToImages_Click()
    LynxGrid.ApplySelectionToImages = chkApplySelectionToImages.Value
End Sub

Private Sub chkCheckBoxes_Click()
    LynxGrid.CheckBoxes = chkCheckBoxes.Value
End Sub

Private Sub chkColumnDrag_Click()
    LynxGrid.ColumnDrag = chkColumnDrag.Value
End Sub


Private Sub chkColumnResize_Click()
    If chkColumnResize.Value Then
        LynxGrid.AllowUserResizing = lgResizeCol
    Else
        LynxGrid.AllowUserResizing = lgResizeNone
    End If
End Sub

Private Sub chkColumnSort_Click()
    LynxGrid.ColumnSort = chkColumnSort.Value
End Sub

Private Sub chkDisplayEllipsis_Click()
    LynxGrid.DisplayEllipsis = chkDisplayEllipsis.Value
End Sub

Private Sub chkEditable_Click()
     LynxGrid.Editable = chkEditable.Value
End Sub

Private Sub chkFullRowSelect_Click()
    LynxGrid.FullRowSelect = chkFullRowSelect.Value
End Sub

Private Sub chkGridLines_Click()
    LynxGrid.GridLines = chkGridLines.Value
End Sub

Private Sub chkHideSelection_Click()
    LynxGrid.HideSelection = chkHideSelection.Value
End Sub

Private Sub chkHotHeaderTracking_Click()
    LynxGrid.HotHeaderTracking = chkHotHeaderTracking.Value
End Sub

Private Sub chkMultiSelect_Click()
    LynxGrid.MultiSelect = chkMultiSelect.Value
End Sub

Private Sub chkScrollTrack_Click()
    LynxGrid.ScrollTrack = chkScrollTrack.Value
End Sub

Private Sub cmdAddItems_Click()
    LoadDemoData
End Sub

Private Sub cmdBinding_Click()
    frmBoundControls.Show vbModeless
End Sub

Private Sub cmdCellWrap_Click()
    frmWrap.Show vbModeless
End Sub

Private Sub cmdChangedCells_Click()
    'NOTE:
    'When editing a Cell directly in the Grid the CellChanged Property for the Cell is set to True
    'The TrackEdits Property controls whether programmatic edits (CellText) change the CellChanged Property

    Dim lCol As Long
    Dim lRow As Long
    Dim lChanges As Long
    
    Screen.MousePointer = vbHourglass
    
    With LynxGrid
        For lRow = 0 To .ItemCount - 1
            For lCol = 0 To .Cols - 1
                If .CellChanged(lRow, lCol) Then
                    lChanges = lChanges + 1
                End If
            Next lCol
        Next lRow
    End With
    
    Screen.MousePointer = vbDefault
    
    If lChanges > 0 Then
        MsgBox lChanges & " Cells have have been changed", vbInformation
    Else
        MsgBox "No Cells have have been changed", vbInformation
    End If
End Sub

Private Sub cmdFormatting_Click()
    frmFormatting.Show vbModeless
End Sub

Private Sub cmdRangeFormat_Click()
    With LynxGrid
        .Redraw = False
        
         'Change Font-Style on Forename Column
        .FormatCells 0, .ItemCount - 1, 2, 2, lgCFFontBold, True
        
        'Change Colours on Surname Column
        .FormatCells 0, .ItemCount - 1, 3, 3, lgCFForeColor, vbYellow
        .FormatCells 0, .ItemCount - 1, 3, 3, lgCFBackColor, vbBlue
        
        'Change Font on Job Title Column
        .FormatCells 0, .ItemCount - 1, 4, 4, lgCFFontName, "Times"
        
        .Redraw = True
    End With
End Sub

Private Sub cmdRemoveItem_Click()
    With LynxGrid
        If .Row >= 0 Then
            .RemoveItem .Row
        End If
    End With
End Sub

Private Sub cmdSort_Click()
    'Sort the Grid by Columns 2 (Forename) & 3 (Surname)
    'NOTE: No Sort Order is specified so clicking button again automatically
    'reverses Sort Order
    
    LynxGrid.Sort 2, , 3
End Sub

Private Sub CreateGrid()
    '###############################################################
    'Columns can be defined in two ways
    
    '1.) By creating columns with a FormatString and then applying properties
    '        .FormatString = "<Code|<Description|>Value"
    '        .ColWidth(0) = 1000
    '        .ColWidth(1) = 3000
    '        .ColWidth(2) = 1500
    '        .ColType(2) = dgNumeric
    '        .ColFormat(2) = "£###.00"
    
    '2.) By calling the AddColumn function (returns the Column Index)
    '        .AddColumn "Code", 1000
    '        .AddColumn "Description", 3000, lgAlignLeftCenter
    '        .AddColumn "Value", 1500, lgAlignRightCenter, lgNumeric, "£###.00"
    '###############################################################
    
    'Notes:
    'Columns have an InputFilter property. This is a string which defines which
    'characters are allowed via keyboard entry using the internal TextBox editor.
    ' < = lowercase
    ' > = UPPERCASE
    ' 1234567890 = Allow only numbers
    
    'Date and Numeric columns default an InputFilter if one is not specified. You
    'can use the EditKeyPress Event for further control.
    
    With LynxGrid
        'Setting Redraw to False stops the Grid redrawing when Items/Cells are
        'changed which makes adding data much faster (and stops application flickering)
        .Redraw = False
        
        'EditTrigger defines which actions toggle Cell Edits. You can use multiple
        'Triggers by using "Or" as below
        
        'Trigger on Enter or any non special purpose key
        '.EditTrigger = lgEnterKey Or lgAnyKey
        'Trigger on Enter of DoubleClick
        .EditTrigger = lgEnterKey Or lgMouseDblClick
        
        'The height used for each Row
        .RowHeightMin = 315
        
        'Set ImageList to provide Item Images
        .ImageList = ImageList1
        
        'Create the Columns
        .AddColumn "Code", 1000, , , , ">" 'Allow Only UPPERCASE
        .AddColumn "G", 250
        .AddColumn "Forename", 1500
        .AddColumn "Surname", 1500
        .AddColumn "Job Title", 1500
        .AddColumn "Pension", 1000, lgAlignCenterCenter, lgBoolean
        .AddColumn "DOB", 1000, , lgDate
        .AddColumn "Premium", 1000, lgAlignRightCenter, lgNumeric, "#.00"
        .AddColumn "Notes", 5000
        
        'Change position of Column in display
        '.ColPosition(2) = 1
        '.ColPosition(6) = 2
        
        'Tell the grid to Draw again!
        .Redraw = True
    End With
End Sub


Private Sub Form_Load()
    With cboFocusRectMode
        .AddItem "None"
        .AddItem "Row"
        .AddItem "Col"
    End With
    
    With cboFocusRectStyle
        .AddItem "Light"
        .AddItem "Heavy"
    End With
    
    With cboThemeColor
        .AddItem "Custom"
        .AddItem "Default"
        .AddItem "Blue"
        .AddItem "Green"
    End With
    
    With cboThemeStyle
        .AddItem "Windows3D"
        .AddItem "WindowsFlat"
        .AddItem "WindowsXP"
        .AddItem "OfficeXP"
    End With
    
    'Set the controls to demo Properties
    With LynxGrid
        .Redraw = False
        
        chkCheckBoxes.Value = Abs(.CheckBoxes)
        chkColumnDrag.Value = Abs(.ColumnDrag)
        chkColumnResize.Value = Abs(.AllowUserResizing = lgResizeCol)
        chkColumnSort.Value = Abs(.ColumnSort)
        chkDisplayEllipsis.Value = Abs(.DisplayEllipsis)
        chkEditable.Value = Abs(.Editable)
        chkFullRowSelect.Value = Abs(.FullRowSelect)
        chkGridLines.Value = Abs(.GridLines)
        chkHideSelection.Value = Abs(.HideSelection)
        chkHotHeaderTracking.Value = Abs(.HotHeaderTracking)
        chkMultiSelect.Value = Abs(.MultiSelect)
        chkScrollTrack.Value = Abs(.ScrollTrack)
        
        cboFocusRectMode.ListIndex = .FocusRectMode
        cboFocusRectStyle.ListIndex = .FocusRectStyle
        
        cboThemeColor.ListIndex = .ThemeColor
        cboThemeStyle.ListIndex = .ThemeStyle
        
        chkAlphaBlendSelection.Value = Abs(.AlphaBlendSelection)
        chkApplySelectionToImages.Value = Abs(.ApplySelectionToImages)
        
        SetColors
         
        .Redraw = True
    End With
    
    CreateGrid
    LoadDemoData
End Sub

Private Function GetColor(NewValue As Long) As Long
    On Local Error GoTo SetBCError

    With CommonDialog1
        .Flags = cdlCCRGBInit
        .Color = NewValue
        .ShowColor
        
        GetColor = .Color
    End With
    Exit Function
    
SetBCError:
    GetColor = NewValue
    Exit Function
End Function

Private Sub lblViewColor_Click(Index As Integer)
    lblViewColor(Index).BackColor = GetColor(lblViewColor(Index).BackColor)
    
    With LynxGrid
        Select Case Index
            Case 0: .BackColorBkg = lblViewColor(0).BackColor
            Case 1: .BackColorEdit = lblViewColor(1).BackColor
            Case 2: .BackColorSel = lblViewColor(2).BackColor
            Case 3: .ForeColorEdit = lblViewColor(3).BackColor
            Case 4: .ForeColorSel = lblViewColor(4).BackColor
            Case 5: .FocusRectColor = lblViewColor(5).BackColor
            Case 6: .GridColor = lblViewColor(6).BackColor
        End Select
    End With
End Sub

Private Sub LoadDemoData()
    Dim lCount As Long
    Dim lRow As Long
    
    With LynxGrid
        'Setting Redraw to False stops the Grid redrawing when Items/Cells are
        'changed which makes adding data much faster (and stops application flickering)
        .Redraw = False
        
        'Add some random data
        For lCount = 1 To 49
            lRow = .AddItem(Format$("XD" & Format$(.ItemCount, "000")))
            
            'Simple method to specify Gender!
            If RandomInt(0, 1) = 0 Then
                'Set the Key for the ImageList Image (can use text Key or Index)
                .ItemImage(lRow) = "MALE" & RandomInt(1, 3)
                .CellText(lRow, 1) = "M"
                .CellText(lRow, 2) = GetForename(ntMale)
                
                'The grid supports per cell formatting but provides Item
                'formatting options for simplicity when only per Row formatting
                'is required (Item formatting reformats all Cells in the Row).
                .ItemForeColor(lRow) = vbBlue
            Else
                .ItemImage(lRow) = RandomInt(3, 6)
                .CellText(lRow, 1) = "F"
                .CellText(lRow, 2) = GetForename(ntFemale)
                .ItemForeColor(lRow) = vbRed
            End If
            
            .CellText(lRow, 3) = GetSurname()
            .CellText(lRow, 4) = GetJobName()
            .CellChecked(lRow, 5) = (RandomInt(0, 1) = 0)
            .CellText(lRow, 6) = DateSerial(RandomInt(1930, 1990), RandomInt(1, 12), RandomInt(1, 28))
            .CellText(lRow, 7) = Round(100 + (Rnd * 100), 2)
        Next lCount
        
        '.ColVisible(0) = False
        
         'Tell the grid to Draw again!
        .Redraw = True
    End With
End Sub

Private Sub LynxGrid_ItemChecked(Row As Long)
    lblCheckedCount.Caption = LynxGrid.CheckedCount
    lblCheckedCount.Refresh
End Sub

Private Sub LynxGrid_ItemCountChanged()
    lblItemCount.Caption = LynxGrid.ItemCount
    lblItemCount.Refresh
End Sub

Private Sub LynxGrid_MouseLeave()
    lblMouseRowCol.Caption = "-"
End Sub

Private Sub LynxGrid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMouseRowCol.Caption = LynxGrid.MouseRow & "," & LynxGrid.MouseCol
    lblMouseRowCol.Refresh
End Sub

Private Sub LynxGrid_RequestEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    'Is the Edit allowed?
    Select Case Col
        Case 1 'Gender Column
            Cancel = True
    End Select
End Sub

Private Sub LynxGrid_RowColChanged()
    lblRowCol.Caption = LynxGrid.Row & "," & LynxGrid.Col
    lblRowCol.Refresh
End Sub

Private Sub LynxGrid_SelectionChanged()
    lblSelectedCount.Caption = LynxGrid.SelectedCount()
    lblSelectedCount.Refresh
End Sub

