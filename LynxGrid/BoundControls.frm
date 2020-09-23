VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBoundControls 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LynxGrid Demo"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   645
      Left            =   4110
      TabIndex        =   6
      Top             =   5460
      Visible         =   0   'False
      Width           =   1755
      Begin VB.TextBox txtName 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   1155
      End
   End
   Begin VB.TextBox txtTest 
      Height          =   315
      Left            =   1110
      TabIndex        =   5
      Top             =   5760
      Width           =   1515
   End
   Begin MSComCtl2.DTPicker dtpDOB 
      Height          =   315
      Left            =   3870
      TabIndex        =   3
      Top             =   3420
      Visible         =   0   'False
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      _Version        =   393216
      Format          =   24051713
      CurrentDate     =   38875
   End
   Begin VB.ComboBox cboJob 
      Height          =   315
      Left            =   2070
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   3150
      Visible         =   0   'False
      Width           =   1425
   End
   Begin LynxGridTest.LynxGrid LynxGrid 
      Height          =   5175
      Left            =   90
      TabIndex        =   1
      Top             =   450
      Width           =   6195
      _extentx        =   10927
      _extenty        =   9128
      focusrectmode   =   2
      allowuserresizing=   1
      columnsort      =   -1  'True
      editable        =   -1  'True
      font            =   "BoundControls.frx":0000
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Test Control"
      Height          =   195
      Left            =   90
      TabIndex        =   4
      Top             =   5790
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Example of binding external Edit Controls"
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
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4965
   End
End
Attribute VB_Name = "frmBoundControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub LoadDemoData()
    Dim lCount As Long
    
    With cboJob
        For lCount = 0 To JobCount()
            .AddItem GetJobName(lCount)
        Next lCount
    End With
        
    With LynxGrid
        'Setting Redraw to False stops the Grid redrawing when Items/Cells are
        'changed which makes adding data much faster (and stops application flickering)
        .Redraw = False
        
        'EditTrigger defines which actions toggle Cell Edits. You can use multiple
        'Triggers by using "Or" as below
        .EditTrigger = lgEnterKey Or lgMouseDblClick
        
        'The height used for each Row
        .RowHeightMin = 315
        
        'Create the Columns
        .AddColumn "Forename", 1500
        .AddColumn "Surname", 1500
        .AddColumn "Job Title", 1500
        .AddColumn "DOB", 1250, , lgDate
        
        'Bind the external Controls to the Column
        .BindControl 1, txtName  'Defaults to automatically changing Left, Top, Height & Width
        .BindControl 2, cboJob, lgBCLeft Or lgBCTop Or lgBCWidth
        .BindControl 3, dtpDOB, lgBCLeft Or lgBCTop Or lgBCWidth
        
        'Add some random data
        For lCount = 0 To 15
            .AddItem
            
            'Simple method to specify Gender!
            If RandomInt(0, 1) = 0 Then
                .CellText(lCount, 0) = GetForename(ntMale)
            Else
                .CellText(lCount, 0) = GetForename(ntFemale)
            End If
            
            .CellText(lCount, 1) = GetSurname()
            .CellText(lCount, 2) = GetJobName()
            .CellText(lCount, 3) = DateSerial(RandomInt(1930, 1990), RandomInt(1, 12), RandomInt(1, 28))
        Next lCount
        
        'Tell the grid to Draw again!
        .Redraw = True
    End With
End Sub



Private Sub Form_Load()
    LoadDemoData
End Sub


Private Sub LynxGrid_RequestEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
     'This Event is fired before an Edit begins. Row & Col identify the Cell that
     'will be edited. Setting Cancel=True will abort the Edit
     
     '############################################################################################################
     'NOTE:
     'If manual processing of the the Edit Control Size/Position is required then you can use the
     'following:
     
     '.ColLeft(Col)     - The Left position of the Cell
     '.ColWidth(Col)    - The Width of the Cell
     '.RowTop(Row)      - The Top position of the Cell
     '.RowHeightMin     - The Height of the Cell
     
     'The MoveControl setting on BindColumn defines what combination (if any) of Left, Top, Height & Width
     'will be adjusted
     '############################################################################################################
     
     Debug.Print "LynxGrid_RequestEdit"
     
     Select Case Col
        Case 1 'Surname
            txtName.Text = LynxGrid.CellText(Row, Col)
        
        Case 2 'Job Title
            If Len(LynxGrid.CellText(Row, Col)) > 0 Then
                cboJob.Text = LynxGrid.CellText(Row, Col)
            Else
                cboJob.ListIndex = -1
            End If
            
        Case 3 'DOB
            dtpDOB.Value = CDate(LynxGrid.CellText(Row, Col))
            
    End Select
End Sub

Private Sub LynxGrid_RequestUpdate(ByVal Row As Long, ByVal Col As Long, NewValue As String, Cancel As Boolean)
    'This Event is fired before an Edit commits. Row & Col identify the Cell that
    'will be updated. Setting Cancel=True will abort the Update
    
    'NewValue is used to get the data that will be used to update the Cell. For
    'columns that are using the internal textbox this will be populated automatically

    Debug.Print "LynxGrid_RequestUpdate"

    Select Case Col
        Case 1 'Surname
            NewValue = txtName.Text
    
        Case 2 'Job Title
            NewValue = cboJob.Text
            
        Case 3 'DOB
            NewValue = dtpDOB.Value
            
    End Select
End Sub

