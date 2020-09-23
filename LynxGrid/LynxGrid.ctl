VERSION 5.00
Begin VB.UserControl LynxGrid 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2835
   KeyPreview      =   -1  'True
   ScaleHeight     =   92
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   189
   ToolboxBitmap   =   "LynxGrid.ctx":0000
End
Attribute VB_Name = "LynxGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'####################################################################################
'Title:     LynxGrid
'Function:  Owner-drawn editable Grid
'Author:    Richard Mewett
'Created:   01/08/05
'Version:   1.89 (10 May 2007)
'
'Copyright © 2005 Richard Mewett. All rights reserved.

'Provides a combination of MSFlexGrid and ListView (Report Style) functionality.

'####################################################################################
'Credits:   Paul Caton - Subclassing
'           Gary Noble (Phantom Man)- API Scroll Bar Code
'           Heriberto Mantilla Santamaría - XP Theme API + Alpha Blend
'           Matthew R. Usner - DrawArrow + Beta testing
'           LaVolpe - Bug fixes & numerous suggestions
'           Riccardo Cohen - Bug reports & ownerdrawn XP/Office ThemeStyles
'           Thierry Calu - ComboBox Height adjustment & automatic DropDown
'           John Underhill (Steppenwolfe) - Unicode suggestions / ReturnAddr patch
            
'####################################################################################
'This software is provided "as-is," without any express or implied warranty.
'In no event shall the author be held liable for any damages arising from the
'use of this software.
'If you do not agree with these terms, do not install "LynxGrid". Use of
'the program implicitly means you have agreed to these terms.
'
'Permission is granted to anyone to use this software for any purpose,
'including commercial use, and to alter and redistribute it, provided that
'the following conditions are met:
'
'1. All redistributions of source code files must retain all copyright
'   notices that are currently in place, and this list of conditions without
'   any modification.
'
'2. All redistributions in binary form must retain all occurrences of the
'   above copyright notice and web site addresses that are currently in
'   place (for example, in the About boxes).
'
'3. Modified versions in source or binary form must be plainly marked as
'   such, and must not be misrepresented as being the original software.

'################################################################
'API Declarations
Private Declare Function IsCharAlphaNumeric Lib "USER32" Alias "IsCharAlphaNumericA" (ByVal cChar As Byte) As Long
Private Declare Function IsWindowUnicode Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SetParent Lib "USER32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SendMessageAsLong Lib "USER32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function SetCapture Lib "USER32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "USER32" () As Long

Private Declare Function SetRect Lib "USER32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetRectRgn Lib "gdi32" (ByVal hRgn As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long

Private Declare Function DrawTextA Lib "USER32" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawTextW Lib "USER32" (ByVal hdc As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function MoveTo Lib "gdi32" Alias "MoveToEx" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Any) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DrawFocusRect Lib "USER32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function DrawFrameControl Lib "USER32" (ByVal hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Private Declare Function FillRect Lib "USER32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "USER32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GradientFill Lib "msimg32" (ByVal hdc As Long, pVertex As Any, ByVal dwNumVertex As Long, pMesh As Any, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long

'XP
Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
Private Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal lhDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, pClipRect As RECT) As Long
Private Declare Function DrawThemeEdge Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, ByVal iStateId As Long, pDestRect As RECT, ByVal uEdge As Long, ByVal uFlags As Long, pContentRect As RECT) As Long
Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hwnd As Long, ByVal pszClassList As Long) As Long

Private Const CLR_INVALID = &HFFFF

Private Const CB_SETITEMHEIGHT = &H153
Private Const CB_SHOWDROPDOWN = &H14F
Private Const CB_GETDROPPEDSTATE = &H157

Private Const DT_BOTTOM = &H8
Private Const DT_CENTER = &H1
Private Const DT_LEFT = &H0
Private Const DT_RIGHT = &H2
Private Const DT_TOP = &H0
Private Const DT_VCENTER = &H4
Private Const DT_WORD_ELLIPSIS = &H40000
Private Const DT_SINGLELINE = &H20
Private Const DT_WORDBREAK = &H10
Private Const DT_CALCRECT = &H400

Private Const BDR_RAISEDINNER = &H4
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKENINNER = &H8
Private Const BDR_SUNKENOUTER = &H2
Private Const BF_BOTTOM = &H8
Private Const BF_LEFT = &H1
Private Const BF_RIGHT = &H4
Private Const BF_TOP = &H2
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)

Private Const DFC_BUTTON        As Long = &H4

Private Const DFCS_FLAT         As Long = &H4000
Private Const DFCS_BUTTONCHECK  As Long = &H0
Private Const DFCS_BUTTONPUSH   As Long = &H10
Private Const DFCS_CHECKED      As Long = &H400
Private Const DFCS_PUSHED = &H200
Private Const DFCS_TRANSPARENT = &H800 ' Win98/2000 only
Private Const DFCS_HOT = &H1000

Private Const VER_PLATFORM_WIN32_NT = 2

Private Const GRADIENT_FILL_RECT_H    As Long = &H0
Private Const GRADIENT_FILL_RECT_V    As Long = &H1
Private Const GRADIENT_FILL_TRIANGLE  As Long = &H2
Private GRADIENT_FILL_RECT_DIRECTION  As Long

Private Const GWL_STYLE = (-16)
Private Const ES_UPPERCASE As Long = &H8&
Private Const ES_LOWERCASE As Long = &H10&

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type TRIVERTEX
   X As Long
   Y As Long
   Red As Integer
   Green As Integer
   Blue As Integer
   Alpha As Integer
End Type

Private Type GRADIENT_RECT
   UPPERLEFT As Long
   LOWERRIGHT As Long
End Type

'################################################################
'Subclassing
Private Enum eMsgWhen
    [MSG_AFTER] = 1                                  'Message calls back after the original (previous) WndProc
    [MSG_BEFORE] = 2                                 'Message calls back before the original (previous) WndProc
    [MSG_BEFORE_AND_AFTER] = MSG_AFTER Or MSG_BEFORE 'Message calls back before and after the original (previous) WndProc
End Enum

Private Const ALL_MESSAGES     As Long = -1          'All messages added or deleted
Private Const CODE_LEN         As Long = 200         'Length of the machine code in bytes
Private Const GWL_WNDPROC      As Long = -4          'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04         As Long = 88          'Table B (before) address patch offset
Private Const PATCH_05         As Long = 93          'Table B (before) entry count patch offset
Private Const PATCH_08         As Long = 132         'Table A (after) address patch offset
Private Const PATCH_09         As Long = 137         'Table A (after) entry count patch offset

Private Type tSubData                                'Subclass data type
    hwnd                       As Long               'Handle of the window being subclassed
    nAddrSub                   As Long               'The address of our new WndProc (allocated memory).
    nAddrOrig                  As Long               'The address of the pre-existing WndProc
    nMsgCntA                   As Long               'Msg after table entry count
    nMsgCntB                   As Long               'Msg before table entry count
    aMsgTblA()                 As Long               'Msg after table array
    aMsgTblB()                 As Long               'Msg Before table array
End Type

Private sc_aSubData()          As tSubData           'Subclass data array
Private sc_aBuf(1 To CODE_LEN) As Byte               'Code buffer byte array
Private sc_pCWP                As Long               'Address of the CallWindowsProc
Private sc_pEbMode             As Long               'Address of the EbMode IDE break/stop/running function
Private sc_pSWL                As Long               'Address of the SetWindowsLong function

Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GetWindowLongA Lib "USER32" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowLongW Lib "USER32" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLongA Lib "USER32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowLongW Lib "USER32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long

'################################################################
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function TrackMouseEvent Lib "USER32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function TrackMouseEventComCtl Lib "Comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function LoadLibraryW Lib "kernel32" (ByVal lpLibFileName As String) As Long

Private Enum TRACKMOUSEEVENT_FLAGS
    TME_HOVER = &H1&
    TME_LEAVE = &H2&
    TME_QUERY = &H40000000
    TME_CANCEL = &H80000000
End Enum

Private Const WM_SETFOCUS       As Long = &H7
Private Const WM_KILLFOCUS      As Long = &H8
Private Const WM_MOUSELEAVE     As Long = &H2A3
Private Const WM_MOUSEMOVE      As Long = &H200
Private Const WM_MOUSEHOVER     As Long = &H2A1
Private Const WM_MOUSEWHEEL     As Long = &H20A
Private Const WM_VSCROLL        As Long = &H115
Private Const WM_HSCROLL        As Long = &H114
Private Const WM_THEMECHANGED   As Long = &H31A
Private Const WM_ACTIVATE       As Long = &H6
Private Const WM_ACTIVATEAPP    As Long = &H1C
                                
Private Type TRACKMOUSEEVENT_STRUCT
  cbSize          As Long
  dwFlags         As TRACKMOUSEEVENT_FLAGS
  hwndTrack       As Long
  dwHoverTime     As Long
End Type

Private bTrack                       As Boolean
Private bTrackUser32                 As Boolean

'################################################################
'API Scroll Bars
Private Declare Function InitialiseFlatSB Lib "comctl32.dll" Alias "InitializeFlatSB" (ByVal lhWnd As Long) As Long
Private Declare Function SetScrollInfo Lib "USER32" (ByVal hwnd As Long, ByVal n As Long, lpcScrollInfo As SCROLLINFO, ByVal BOOL As Boolean) As Long
Private Declare Function GetScrollInfo Lib "USER32" (ByVal hwnd As Long, ByVal n As Long, LPSCROLLINFO As SCROLLINFO) As Long
Private Declare Function EnableScrollBar Lib "USER32" (ByVal hwnd As Long, ByVal wSBflags As Long, ByVal wArrows As Long) As Long
Private Declare Function ShowScrollBar Lib "USER32" (ByVal hwnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long
Private Declare Function FlatSB_EnableScrollBar Lib "comctl32.dll" (ByVal hwnd As Long, ByVal int2 As Long, ByVal UINT3 As Long) As Long
Private Declare Function FlatSB_ShowScrollBar Lib "comctl32.dll" (ByVal hwnd As Long, ByVal code As Long, ByVal fRedraw As Boolean) As Long
Private Declare Function FlatSB_GetScrollInfo Lib "comctl32.dll" (ByVal hwnd As Long, ByVal code As Long, LPSCROLLINFO As SCROLLINFO) As Long
Private Declare Function FlatSB_SetScrollInfo Lib "comctl32.dll" (ByVal hwnd As Long, ByVal code As Long, LPSCROLLINFO As SCROLLINFO, ByVal fRedraw As Boolean) As Long
Private Declare Function FlatSB_SetScrollProp Lib "comctl32.dll" (ByVal hwnd As Long, ByVal Index As Long, ByVal NewValue As Long, ByVal fRedraw As Boolean) As Long
Private Declare Function UninitializeFlatSB Lib "comctl32.dll" (ByVal hwnd As Long) As Long

Public Enum ScrollBarOrienationEnum
    Scroll_Horizontal
    Scroll_Vertical
    Scroll_Both
End Enum

Public Enum ScrollBarStyleEnum
    Style_Regular = 1& ' FSB_REGULAR_MODE
    Style_Flat = 0& 'FSB_FLAT_MODE
End Enum

Public Enum EFSScrollBarConstants
    efsHorizontal = 0 'SB_HORZ
    efsVertical = 1 'SB_VERT
End Enum

Private Const SB_BOTTOM = 7
Private Const SB_ENDSCROLL = 8
Private Const SB_HORZ = 0
Private Const SB_LEFT = 6
Private Const SB_LINEDOWN = 1
Private Const SB_LINELEFT = 0
Private Const SB_LINERIGHT = 1
Private Const SB_LINEUP = 0
Private Const SB_PAGEDOWN = 3
Private Const SB_PAGELEFT = 2
Private Const SB_PAGERIGHT = 3
Private Const SB_PAGEUP = 2
Private Const SB_RIGHT = 7
Private Const SB_THUMBTRACK = 5
Private Const SB_TOP = 6
Private Const SB_VERT = 1

Private Const SIF_RANGE = &H1
Private Const SIF_PAGE = &H2
Private Const SIF_POS = &H4
Private Const SIF_TRACKPOS = &H10
Private Const SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)

Private Const ESB_DISABLE_BOTH = &H3
Private Const ESB_ENABLE_BOTH = &H0
Private Const MK_CONTROL = &H8
Private Const WSB_PROP_VSTYLE = &H100&
Private Const WSB_PROP_HSTYLE = &H200&
Private Const FSB_FLAT_MODE = 1&
Private Const FSB_REGULAR_MODE = 0&

Private Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type

Private m_bInitialised      As Boolean
Private m_eOrientation      As ScrollBarOrienationEnum
Private m_eStyle            As ScrollBarStyleEnum
Private m_hWnd              As Long
Private m_lSmallChangeHorz  As Long
Private m_lSmallChangeVert  As Long
Private m_bEnabledHorz      As Boolean
Private m_bEnabledVert      As Boolean
Private m_bVisibleHorz      As Boolean
Private m_bVisibleVert      As Boolean
Private m_bNoFlatScrollBars As Boolean

'################################################################
Private Enum lgFlagsEnum
    lgFLChecked = 2
    lgFLSelected = 4
    lgFLChanged = 8
    lgFLFontBold = 16
    lgFLFontItalic = 32
    lgFLFontUnderline = 64
    lgFLWordWrap = 128
End Enum

Private Enum lgHeaderStateEnum
    lgNormal = 1
    lgHot = 2
    lgDown = 3
End Enum

Private Enum lgRectTypeEnum
    lgRTColumn = 0
    lgRTCheckBox = 1
    lgRTImage = 2
End Enum

Public Enum lgAllowUserResizingEnum
    lgResizeNone = 0
    lgResizeCol = 1
    'lgResizeRow = 2
    lgResizeBoth = 4
End Enum

Public Enum lgAlignmentEnum
    lgAlignLeftTop = DT_LEFT Or DT_TOP
    lgAlignLeftCenter = DT_LEFT Or DT_VCENTER
    lgAlignLeftBottom = DT_LEFT Or DT_BOTTOM
    lgAlignCenterTop = DT_CENTER Or DT_TOP
    lgAlignCenterCenter = DT_CENTER Or DT_VCENTER
    lgAlignCenterBottom = DT_CENTER Or DT_BOTTOM
    lgAlignRightTop = DT_RIGHT Or DT_TOP
    lgAlignRightCenter = DT_RIGHT Or DT_VCENTER
    lgAlignRightBottom = DT_RIGHT Or DT_BOTTOM
End Enum

Public Enum lgBorderStyleEnum
    lgNone = 0
    lgSingle = 1
End Enum

Public Enum lgCellFormatEnum
    lgCFBackColor = 1
    lgCFForeColor = 2
    lgCFImage = 2
    lgCFFontName = 3
    lgCFFontBold = 4
    lgCFFontItalic = 5
    lgCFFontUnderline = 6
End Enum

Public Enum lgDataTypeEnum
    lgString = 0
    lgNumeric = 1
    lgDate = 2
    lgBoolean = 3
    lgProgressBar = 4
    lgCustom = 5
End Enum

Public Enum lgEditTriggerEnum
    lgNone = 0
    lgEnterKey = 2
    lgF2Key = 4
    lgMouseClick = 8
    lgMouseDblClick = 16
    lgAnyKey = 32
End Enum

Public Enum lgFocusRectModeEnum
    lgNone = 0
    lgRow = 1
    lgCol = 2
End Enum

Public Enum lgFocusRectStyleEnum
    lgFRLight = 0
    lgFRHeavy = 1
End Enum

Public Enum lgMoveControlEnum
    lgBCNone = 0
    lgBCHeight = 1
    lgBCWidth = 2
    lgBCLeft = 4
    lgBCTop = 8
End Enum

Public Enum lgSearchModeEnum
    lgSMEqual = 0
    lgSMGreaterEqual = 1
    lgSMLike = 2
    lgSMNavigate = 4
End Enum

Public Enum lgSortTypeEnum
    lgSTAscending = 0
    lgSTDescending = 1
End Enum

Public Enum lgThemeColorEnum
    lgTCCustom = 0
    lgTCDefault = 1
    lgTCBlue = 2
    lgTCGreen = 3
End Enum

Public Enum lgThemeStyleEnum
    lgTSWindows3D = 0
    lgTSWindowsFlat = 1
    lgTSWindowsXP = 2
    lgTSOfficeXP = 3
End Enum

#If False Then
    Private lgFLChecked, lgFLSelected, lgFLChanged, lgFLFontBold, lgFLFontItalic, lgFLFontUnderline, lgFLWordWrap
    Private lgNormal, lgHot, lgDown
    Private lgResizeNone, lgResizeCol, lgResizeRow, lgResizeBoth
    Private lgAlignLeftTop, lgAlignLeftCenter, lgAlignLeftBottom, lgAlignCenterTop, lgAlignCenterCenter, lgAlignCenterBottom, lgAlignRightTop, lgAlignRightCenter, lgAlignRightBottom
    Private lgCFBackColor, lgCFForeColor, lgCFImage, lgCFFontName, lgCFFontBold, lgCFFontItalic, lgCFFontUnderline
    Private lgNone, lgSingle
    Private lgString, lgNumeric, lgDate, lgBoolean, lgProgressBar, lgCustom
    Private lgNone, lgEnterKey, lgF2Key, lgMouseClick, lgMouseDblClick, lgAnyKey
    Private lgNone, lgRow, lgCol
    Private lgFRLight, lgFRHeavy
    Private lgSMEqual, lgSMGreaterEqual, lgSMLike, lgSMNavigate
    Private lgSTAscending, lgSTDescending
    Private lgTCCustom, lgTCDefault, lgTCBlue, lgTCGreen
    Private lgTSWindows3D, lgTSWindowsFlat, lgTSWindowsXP, lgTSOfficeXP
#End If

Private Const ROW_HEIGHT                As Long = 16

Private Const DEF_ALLOWUSERRESIZING         As Long = lgAllowUserResizingEnum.lgResizeNone
Private Const DEF_ALPHABLENDSELECTION       As Boolean = False
Private Const DEF_APPLYSELECTIONTOIMAGES    As Boolean = True
Private Const DEF_AUTOSIZEROW               As Boolean = False
Private Const DEF_BACKCOLOR                 As Long = vbWindowBackground
Private Const DEF_BACKCOLORBKG              As Long = &H808080
Private Const DEF_BACKCOLOREDIT             As Long = &HC0FFFF
Private Const DEF_BACKCOLORFIXED            As Long = vbButtonFace
Private Const DEF_BACKCOLORSEL              As Long = vbHighlight
Private Const DEF_BORDERSTYLE               As Long = lgBorderStyleEnum.lgSingle
Private Const DEF_CACHEINCREMENT            As Long = 10
Private Const DEF_CHECKBOXES                As Boolean = False
Private Const DEF_COLUMNDRAG                As Boolean = False
Private Const DEF_COLUMNHEADERS             As Boolean = True
Private Const DEF_COLUMNSORT                As Boolean = False
Private Const DEF_DISPLAYELLIPSIS           As Boolean = True
Private Const DEF_EDITABLE                  As Boolean = False
Private Const DEF_EDITTRIGGER               As Long = lgEditTriggerEnum.lgEnterKey
Private Const DEF_ENABLED                   As Boolean = True
Private Const DEF_FOCUSRECTCOLOR            As Long = &HFFFF&
Private Const DEF_FOCUSRECTMODE             As Long = lgFocusRectModeEnum.lgRow
Private Const DEF_FOCUSRECTSTYLE            As Long = lgFocusRectStyleEnum.lgFRHeavy
Private Const DEF_FORECOLOR                 As Long = vbWindowText
Private Const DEF_FORECOLOREDIT             As Long = vbWindowText
Private Const DEF_FORECOLORFIXED            As Long = vbButtonText
Private Const DEF_FORECOLORHDR              As Long = vbWindowText
Private Const DEF_FORECOLORSEL              As Long = vbHighlightText
Private Const DEF_FORECOLORTOTALS           As Long = vbRed
Private Const DEF_FORMATSTRING              As String = vbNullString
Private Const DEF_FULLROWSELECT             As Boolean = True
Private Const DEF_GRIDCOLOR                 As Long = &HC0C0C0
Private Const DEF_GRIDLINES                 As Boolean = True
Private Const DEF_GRIDLINEWIDTH             As Long = 1
Private Const DEF_HIDESELECTION             As Boolean = True
Private Const DEF_HOTHEADERTRACKING         As Boolean = True
Private Const DEF_LOCKED                    As Boolean = False
Private Const DEF_MULTISELECT               As Boolean = False
Private Const DEF_PROGRESSBARCOLOR          As Long = &H8080FF
Private Const DEF_REDRAW                    As Boolean = True
Private Const DEF_ROWHEIGHTMAX              As Long = 0
Private Const DEF_ROWHEIGHTMIN              As Long = 0
Private Const DEF_SCALEUNITS                As Integer = vbTwips
Private Const DEF_SCROLLTRACK               As Boolean = True
Private Const DEF_SEARCHCOLUMN              As Long = 0
Private Const DEF_THEMECOLOR                As Long = lgThemeColorEnum.lgTCCustom
Private Const DEF_THEMESTYLE                As Long = lgThemeStyleEnum.lgTSWindowsXP
Private Const DEF_TRACKEDITS                As Boolean = False

Private Const NULL_RESULT               As Long = -1
Private Const AUTOSCROLL_TIMEOUT        As Long = 25
Private Const SIZE_VARIANCE             As Long = 4

Private Const SCROLL_NONE               As Long = 0
Private Const SCROLL_UP                 As Long = 1
Private Const SCROLL_DOWN               As Long = 2

'##########################################
'For Rendering
Private Const MAX_CHECKBOXSIZE          As Long = 16
Private Const SIZE_SORTARROW            As Long = 8

Private Const HEADER_LEFT               As Long = 3
Private Const TEXT_SPACE                As Long = 3
Private Const ARROW_SPACE               As Long = 5

Private Const DEFAULT_LEFTTEXT          As Long = 3
Private Const RIGHT_CHECKBOX            As Long = 15
'##########################################

Private Type udtColumn
    EditCtrl As Object
    dCustomWidth As Single
    lWidth As Long
    lX As Long
    nAlignment As lgAlignmentEnum
    nImageAlignment As lgAlignmentEnum
    nSortOrder As lgSortTypeEnum
    nType As Integer
    nFlags As Integer
    MoveControl As Integer
    bVisible As Boolean
    sCaption As String
    sFormat As String
    sInputFilter As String
    sTag As String
End Type

Private Type udtCell
    nAlignment As Integer
    nFormat As Integer
    nFlags As Integer
    sValue As String
End Type

Private Type udtItem
    lHeight As Long
    lImage As Long
    lItemData As Long
    nFlags As Integer
    sTag As String
    Cell() As udtCell
End Type

Private Type udtFormat
    lBackColor As Long
    lForeColor As Long
    nImage As Integer
    sFontName As String
    lRefCount As Long
End Type

Private Type udtRender
    DTFlag As Long
    CheckBoxSize As Long
    ImageSpace As Long
    ImageHeight As Long
    ImageWidth As Long
    LeftImage As Long
    LeftText As Long
    HeaderHeight As Long
    TextHeight As Long
End Type

Private WithEvents txtEdit As TextBox
Attribute txtEdit.VB_VarHelpID = -1

'################################################################
'Data & Columns
Private mCols() As udtColumn
Private mItems() As udtItem
Private mColPtr() As Long
Private mRowPtr() As Long
Private mCF() As udtFormat

Private mItemCount As Long
Private mItemsVisible As Long
Private mSortColumn As Long
Private mSortSubColumn As Long

Private mEditCol As Long
Private mEditRow As Long
Private mCol As Long
Private mRow As Long
Private mMouseCol As Long
Private mMouseRow As Long
Private mMouseDownCol As Long
Private mMouseDownRow As Long

Private mMouseDownX As Long

Private mSelectedRow As Long

Private mR As udtRender
Private mEditPending As Boolean
Private mMouseDown As Boolean
Private mDragCol As Long
Private mResizeCol As Long
Private mEditParent As Long

'################################################################
'Appearance Properties
Private mApplySelectionToImages As Boolean
Private mBackColor As Long
Private mBackColorBkg As Long
Private mBackColorEdit As Long
Private mBackColorFixed As Long
Private mBackColorSel As Long
Private mForeColor As Long
Private mForeColorEdit As Long
Private mForeColorFixed As Long
Private mForeColorHdr As Long
Private mForeColorSel As Long
Private mForeColorTotals As Long

Private mFocusRectColor As Long
Private mGridColor As Long
Private mProgressBarColor As Long

Private mAlphaBlendSelection As Boolean
Private mBorderStyle As lgBorderStyleEnum
Private mDisplayEllipsis As Boolean
Private mFocusRectMode As lgFocusRectModeEnum
Private mFocusRectStyle As lgFocusRectStyleEnum
Private mFont As Font
Private mGridLines As Boolean
Private mGridLineWidth As Long
Private mThemeColor As lgThemeColorEnum
Private mThemeStyle As lgThemeStyleEnum

'################################################################
'Behaviour Properties
Private mAllowUserResizing As lgAllowUserResizingEnum
Private mAutoSizeRow As Boolean
Private mCheckboxes As Boolean
Private mColumnDrag As Boolean
Private mColumnHeaders As Boolean
Private mColumnSort As Boolean
Private mEditable As Boolean
Private mEditTrigger As lgEditTriggerEnum
Private mFullRowSelect As Boolean
Private mHideSelection As Boolean
Private mHotHeaderTracking As Boolean
Private mMultiSelect As Boolean
Private mRedraw As Boolean
Private mScrollTrack As Boolean
Private mTrackEdits As Boolean

'################################################################
'Miscellaneous Properties
Private mCacheIncrement As Long
Private mEnabled As Boolean
Private mExpandRowImage As Integer
Private mFormatString As String
Private mLocked As Boolean
Private mRowHeightMax As Long
Private mRowHeightMin As Long
Private mScaleUnits As ScaleModeConstants
Private mSearchColumn As Long

Private mImageList As Object
Private mImageListScaleMode As Integer

'################################################################
'Control State Variables
Private mInCtrl As Boolean
Private mInFocus As Boolean
Private mWinNT As Boolean
Private mWinXP As Boolean
Private mLockFocusDraw As Boolean

Private mPendingRedraw As Boolean
Private mPendingScrollBar As Boolean

Private mTextBoxStyle As Long
Private mClipRgn As Long
Private hTheme As Long
Private mScrollAction As Long
Private mScrollTick As Long
Private mHotColumn As Long
Private mIgnoreKeyPress As Boolean

'################################################################
'Events - Standard VB
Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Events - Control Specific
Public Event CellImageClick(ByVal Row As Long, ByVal Col As Long)
Public Event ColumnClick(Col As Long)
Public Event ColumnSizeChanged(Col As Long, MoveControl As lgMoveControlEnum)
Public Event CustomSort(Ascending As Boolean, Col As Long, Value1 As String, Value2 As String, Swap As Boolean)
Public Event ItemChecked(Row As Long)
Public Event ItemCountChanged()
Public Event MouseEnter()
Public Event MouseLeave()
Public Event RowColChanged()
Public Event Scroll()
Public Event SelectionChanged()
Public Event SortComplete()
Public Event ThemeChanged()

Public Event EditKeyPress(ByVal Col As Long, KeyAscii As Integer)
Public Event EnterCell()
Public Event RequestEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Public Event RequestUpdate(ByVal Row As Long, ByVal Col As Long, NewValue As String, Cancel As Boolean)

Private Function IsColumnTruncated(Col As Long) As Boolean
    If (mR.LeftText > DEFAULT_LEFTTEXT) And (Col = 0) Then
        IsColumnTruncated = True
    End If
End Function

'Subclass handler
Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)
    Dim eBar As EFSScrollBarConstants
    Dim lV As Long, lSC As Long
    Dim lScrollCode As Long
    Dim tSI As SCROLLINFO
    Dim zDelta As Long
    Dim lHSB As Long
    Dim lVSB As Long
    Dim bRedraw As Boolean
    
    'Debug.Print "zSubclass_Proc " & Timer
    
    Select Case uMsg
        Case WM_VSCROLL, WM_HSCROLL, WM_MOUSEWHEEL
            lScrollCode = (wParam And &HFFFF&)
            
            lHSB = SBValue(efsHorizontal)
            lVSB = SBValue(efsVertical)
    
            Select Case uMsg
            
                Case WM_HSCROLL ' Get the scrollbar type
                    eBar = efsHorizontal
                    
                Case WM_VSCROLL
                    eBar = efsVertical
                    
                Case Else     'WM_MOUSEWHEEL
                    eBar = IIf(lScrollCode And MK_CONTROL, efsHorizontal, efsVertical)
                    lScrollCode = IIf(wParam / 65536 < 0, SB_LINEDOWN, SB_LINEUP)
                    
            End Select
            
            bRedraw = True
    
            Select Case lScrollCode
            
                Case SB_THUMBTRACK
                    ' Is vertical/horizontal?
                    pSBGetSI eBar, tSI, SIF_TRACKPOS
                    SBValue(eBar) = tSI.nTrackPos
                    
                    bRedraw = mScrollTrack
    
                Case SB_LEFT, SB_BOTTOM
                     SBValue(eBar) = IIf(lScrollCode = 7, SBMax(eBar), SBMin(eBar))
    
                Case SB_RIGHT, SB_TOP
                     SBValue(eBar) = SBMin(eBar)
    
                Case SB_LINELEFT, SB_LINEUP
                
                    If SBVisible(eBar) Then
                    
                        lV = SBValue(eBar)
                        If (eBar = efsHorizontal) Then
                            lSC = m_lSmallChangeHorz
                        Else
                            lSC = m_lSmallChangeVert
                        End If
                        
                        If (lV - lSC < SBMin(eBar)) Then
                             SBValue(eBar) = SBMin(eBar)
                        Else
                             SBValue(eBar) = lV - lSC
                        End If
                        
                    End If
    
                Case SB_LINERIGHT, SB_LINEDOWN
                    If SBVisible(eBar) Then
            
                        lV = SBValue(eBar)
                        
                        If (eBar = efsHorizontal) Then
                            lSC = m_lSmallChangeHorz
                        Else
                            lSC = m_lSmallChangeVert
                        End If
                        
                        If (lV + lSC > SBMax(eBar)) Then
                             SBValue(eBar) = SBMax(eBar)
                        Else
                             SBValue(eBar) = lV + lSC
                        End If
                    End If
    
                Case SB_PAGELEFT, SB_PAGEUP
                     SBValue(eBar) = SBValue(eBar) - SBLargeChange(eBar)
    
                Case SB_PAGERIGHT, SB_PAGEDOWN
                     SBValue(eBar) = SBValue(eBar) + SBLargeChange(eBar)
    
                Case SB_ENDSCROLL
                    If Not mScrollTrack Then
                        DrawGrid True
                    End If
    
            End Select
            
            If (lHSB <> SBValue(efsHorizontal)) Or (lVSB <> SBValue(efsVertical)) Then
                UpdateCell
                
                If bRedraw Then
                    DrawGrid True
                End If
                
                RaiseEvent Scroll
            End If
        
        Case WM_MOUSEWHEEL
                
        Case WM_MOUSEMOVE
            If Not mInCtrl Then
                mInCtrl = True
                Call TrackMouseLeave(lng_hWnd)
                RaiseEvent MouseEnter
            End If
    
        Case WM_MOUSELEAVE
            If mInCtrl Then
                mInCtrl = False
                DrawHeaderRow
                UserControl.Refresh
                RaiseEvent MouseLeave
            End If
            
        Case WM_SETFOCUS
             If mEnabled Then
                If Not mInFocus Then
                    mInFocus = True
                        
                    If Not mLockFocusDraw Then
                        'Debug.Print "WM_SETFOCUS"
                        DrawGrid True
                    End If
                End If
             End If
    
        Case WM_KILLFOCUS
            If lng_hWnd = UserControl.hwnd Then
                If mEnabled Then
                    If mInFocus Then
                        mInFocus = False
                           
                        If Not mLockFocusDraw Then
                            'Debug.Print "WM_KILLFOCUS"
                            DrawGrid True
                        End If
                    End If
                End If
            ElseIf Not mInCtrl Then
                UpdateCell
            End If
    
        Case WM_THEMECHANGED
            DrawGrid True
            RaiseEvent ThemeChanged

    End Select
End Sub

Public Function AddColumn(Optional Caption As String, Optional Width As Single, Optional Alignment As lgAlignmentEnum = lgAlignLeftCenter, Optional DataType As lgDataTypeEnum = lgString, Optional Format As String, Optional InputFilter As String, Optional ImageAlignment As lgAlignmentEnum = lgAlignLeftCenter, Optional WordWrap As Boolean, Optional Index As Long = 0) As Long
    '#############################################################################################################################
    'Purpose: Add a Column to the Grid
    
    'Caption        - The text that appears on the Header
    'Width          - The Width!
    'Alignment      - The Alignment!
    'DataType       - Allows the control to determine proper Sort Sequence when Sorting
    'Format         - Format Mask applied to Cell data before it is displayed (i.e. "#.00")
    'InputFilter    - Characters allowed in TextBox entry
    'ImageAlignment - Image Alignment!
    'WordWrap       - Enable Word-Wrap
    'Index          - Allows a new Column to be Inserted before an existing one
    
    'mColPtr() is used as an Index to the Columns (a bit like an array of "pointers")
    '#############################################################################################################################
    
    Dim lCount As Long
    Dim lNewCol As Long
    
    If mCols(0).nAlignment <> 0 Then
        lNewCol = UBound(mCols) + 1
        ReDim Preserve mCols(lNewCol)
        ReDim Preserve mColPtr(lNewCol)
    End If
    
    If (Index > 0) And (Index < lNewCol) Then
        If lNewCol > 1 Then
            For lCount = lNewCol To Index + 1 Step -1
                mColPtr(lCount) = mColPtr(lCount - 1)
            Next lCount
            mColPtr(Index) = lNewCol
        End If
        
        AddColumn = Index
    Else
        mColPtr(lNewCol) = lNewCol
        AddColumn = lNewCol
    End If
 
    With mCols(lNewCol)
        .sCaption = Caption
        .dCustomWidth = Width
        
        'lWidth is always Pixels (because thats what API functions require) and
        'is calculated to prevent repeated Width Scaling calculations
        .lWidth = ScaleX(.dCustomWidth, mScaleUnits, vbPixels)
        
        .nAlignment = Alignment
        .nImageAlignment = ImageAlignment
        .nSortOrder = lgSTAscending
        .nType = DataType
        .sFormat = Format
        
        If Len(InputFilter) = 0 Then
            Select Case DataType
                Case lgNumeric
                    .sInputFilter = "1234567890.,-"
                Case lgDate
                    .sInputFilter = "1234567890./-"
            End Select
        Else
            .sInputFilter = InputFilter
        End If
        
        If WordWrap Then
            .nFlags = lgFLWordWrap
        End If
        
        .bVisible = True
    End With
    
    DisplayChange
End Function

Public Function AddItem(Optional ByVal Item As String, Optional Index As Long = 0, Optional Checked As Boolean) As Long
    '#############################################################################################################################
    'Purpose: Add an Item (new Row) to the Grid
    
    'Item       - This contains the data for the Cells in the new Row. You can pass multiple
    '           Cells by using a Delimiter between Cell data
    'Index      - Allows a new Item to be Inserted before an existing one
    'Checked    - Default Checked state of the new Item
    
    'mItems() is an array of the Items in the Grid
    'mRowPtr() is used as an Index to the Items (a bit like an array of "pointers")
    
    'The Index technique is used to allow faster Inserts & Sorts since we only need to swap a Long (4 bytes)
    'rather than a large data structure (a UDT in this case)
    
    'The mItems() is resized incrementally to reduce the Redim Preserve overhead. The default mCacheIncrement
    'is 10 but this can be increased to a higher value to increase performance if adding thousands of Items
    '#############################################################################################################################

    Dim lCol As Long
    Dim lCount As Long
    Dim sText() As String
    
    mItemCount = mItemCount + 1
    If mItemCount > UBound(mItems) Then
        ReDim Preserve mItems(mItemCount + mCacheIncrement)
        ReDim Preserve mRowPtr(mItemCount + mCacheIncrement)
    End If
    
    If (Index > 0) And (Index < mItemCount) Then
        If mItemCount > 1 Then
            For lCount = mItemCount To Index + 1 Step -1
                mRowPtr(lCount) = mRowPtr(lCount - 1)
            Next lCount
            mRowPtr(Index) = mItemCount
        End If
        
        AddItem = Index
    Else
        mRowPtr(mItemCount) = mItemCount
        AddItem = mItemCount
    End If
    
    If mRowHeightMin > 0 Then
        mItems(mItemCount).lHeight = ScaleY(mRowHeightMin, mScaleUnits, vbPixels)
    Else
        mItems(mItemCount).lHeight = ROW_HEIGHT
    End If
    
    ReDim mItems(mItemCount).Cell(UBound(mCols))
        
    For lCount = LBound(mCols) To UBound(mCols)
        With mItems(mItemCount).Cell(lCount)
            .nAlignment = mCols(lCount).nAlignment
            .nFormat = -1
            .nFlags = mCols(lCount).nFlags
        End With
        
        ApplyCellFormat mItemCount, lCount, lgCFBackColor, mBackColor
        ApplyCellFormat mItemCount, lCount, lgCFForeColor, mForeColor
        ApplyCellFormat mItemCount, lCount, lgCFFontName, mFont.Name
    Next lCount
    
    If UBound(mCols) > 0 Then
        lCol = 0
        sText() = Split(Item, vbTab)
        For lCount = LBound(sText) To UBound(sText)
            With mItems(mItemCount).Cell(lCol)
                .sValue = sText(lCount)
            End With
            
            lCol = lCol + 1
            If lCol > UBound(mCols) Then
                Exit For
            End If
        Next lCount
    Else
        mItems(mItemCount).Cell(0).sValue = Item
    End If
    
    If Checked Then
        SetFlag mItems(mItemCount).nFlags, lgFLChecked, True
    End If
    
    DisplayChange
    
    RaiseEvent ItemCountChanged
End Function

Public Property Get AllowUserResizing() As lgAllowUserResizingEnum
Attribute AllowUserResizing.VB_ProcData.VB_Invoke_Property = ";Behavior"
    AllowUserResizing = mAllowUserResizing
End Property

Public Property Let AllowUserResizing(ByVal NewValue As lgAllowUserResizingEnum)
    mAllowUserResizing = NewValue
    
    PropertyChanged "AllowUserResizing"
End Property

Public Property Let ApplySelectionToImages(ByVal NewValue As Boolean)
    mApplySelectionToImages = NewValue
    DrawGrid mRedraw
    
    PropertyChanged "ApplySelectionToImages"
End Property

Public Property Get ApplySelectionToImages() As Boolean
Attribute ApplySelectionToImages.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ApplySelectionToImages = mApplySelectionToImages
End Property

Private Sub ApplyCellFormat(ByVal Row As Long, ByVal Col As Long, Apply As lgCellFormatEnum, NewValue As Variant)
    '#############################################################################################################################
    'Purpose: Apply formatting to a Cell. Attempts to find a matching entry in the
    'Format Table and creates a new entry if a match is not found.
    
    'In any "normal" use the grid will only have a few specifically formatted cells
    '(such as Red forecolor in a financial column to indicate negative). It is therefore
    'wasteful for each cell to store these properties. This system significantly reduces
    'the memory used by the cells in a large Grid at the cost of slightly reduced perfomance.
    
    'The Format element is an Integer allowing 32767 combinations. It could be a
    'long for more combinations - however the aim is to keep the Cell UDT as small as possible!
    
    Dim lBackColor As Long
    Dim lForeColor As Long
    Dim nImage As Integer
    Dim sFontName As String
    
    Dim lCount As Long
    Dim nIndex As Integer
    Dim nFreeIndex As Integer
    Dim nNewIndex As Integer
    Dim bMatch As Boolean
    
    nIndex = mItems(Row).Cell(Col).nFormat
    
    If nIndex >= 0 Then
        'Get current properties
        With mCF(nIndex)
            lBackColor = .lBackColor
            lForeColor = .lForeColor
            nImage = .nImage
            sFontName = .sFontName
        End With
    Else
        'Set default properties
        lBackColor = mBackColor
        lForeColor = mForeColor
        sFontName = mFont.Name
    End If
        
    Select Case Apply
        Case lgCFBackColor
            lBackColor = NewValue
        
        Case lgCFForeColor
            lForeColor = NewValue
            
        Case lgCFImage
            nImage = NewValue
            
        Case lgCFFontName
            sFontName = NewValue
            
    End Select
    
    'Search Format Table for matching entry
    nFreeIndex = -1
    For lCount = 0 To UBound(mCF)
        If (mCF(lCount).lBackColor = lBackColor) And (mCF(lCount).lForeColor = lForeColor) And (mCF(lCount).nImage = nImage) And (mCF(lCount).sFontName = sFontName) Then
            'Existing Entry matches what we required
            bMatch = True
            nNewIndex = lCount
            Exit For
        ElseIf (mCF(lCount).lRefCount = 0) And (nFreeIndex = -1) Then
            'An unused entry
            nFreeIndex = lCount
        End If
    Next lCount
    
    'No existing matches
    If Not bMatch Then
        'Is there an unused Entry?
        If nFreeIndex >= 0 Then
            nNewIndex = nFreeIndex
        Else
            nNewIndex = UBound(mCF) + 1
            ReDim Preserve mCF(nNewIndex + 9)
        End If
        
        With mCF(nNewIndex)
            .lBackColor = lBackColor
            .lForeColor = lForeColor
            .nImage = nImage
            .sFontName = sFontName
        End With
    End If
    
    'Has the Format Entry Index changed?
    If (nIndex <> nNewIndex) Then
        'Increment reference count for new entry
        mCF(nNewIndex).lRefCount = mCF(nNewIndex).lRefCount + 1
           
        If nIndex >= 0 Then
            'Decrement reference count for previous entry
            mCF(nIndex).lRefCount = mCF(nIndex).lRefCount - 1
        End If
    End If
        
    mItems(Row).Cell(Col).nFormat = nNewIndex
End Sub

Public Property Get AutoSizeRow() As Boolean
Attribute AutoSizeRow.VB_ProcData.VB_Invoke_Property = ";Behavior"
    AutoSizeRow = mAutoSizeRow
End Property

Public Property Let AutoSizeRow(ByVal NewValue As Boolean)
    mAutoSizeRow = NewValue
    DrawGrid mRedraw
    
    PropertyChanged "AutoSizeRow"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackColor = mBackColor
End Property

Public Property Let BackColor(ByVal NewValue As OLE_COLOR)
    mBackColor = NewValue
    DrawGrid mRedraw
    
    PropertyChanged "BackColor"
End Property

Public Property Get BackColorBkg() As OLE_COLOR
Attribute BackColorBkg.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackColorBkg = mBackColorBkg
End Property

Public Property Let BackColorBkg(ByVal NewValue As OLE_COLOR)
    mBackColorBkg = NewValue
    UserControl.BackColor = mBackColorBkg
    DisplayChange
    
    PropertyChanged "BackColorBkg"
End Property

Public Property Get BackColorEdit() As OLE_COLOR
Attribute BackColorEdit.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackColorEdit = mBackColorEdit
End Property

Public Property Let BackColorEdit(ByVal lNewValue As OLE_COLOR)
    mBackColorEdit = lNewValue
    
    PropertyChanged "BackColorEdit"
End Property

Public Property Get BackColorFixed() As OLE_COLOR
Attribute BackColorFixed.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackColorFixed = mBackColorFixed
End Property

Public Property Let BackColorFixed(ByVal NewValue As OLE_COLOR)
    mBackColorFixed = NewValue
    
    PropertyChanged "BackColorFixed"
End Property

Public Property Get BackColorSel() As OLE_COLOR
    BackColorSel = mBackColorSel
End Property

Public Property Let BackColorSel(ByVal NewValue As OLE_COLOR)
    mBackColorSel = NewValue
    DisplayChange
    
    PropertyChanged "BackColorSel"
End Property

Public Sub BindControl(ByVal Col As Long, Ctrl As Object, Optional MoveControl As lgMoveControlEnum = lgBCHeight Or lgBCLeft Or lgBCTop Or lgBCWidth)
    '#############################################################################################################################
    'Purpose: Bind an external Control to a Column
    
    'Col    - Column Index
    'Ctrl   - The Control!
    'Resize - Specify how the Control Size should be modified
    '#############################################################################################################################

    Set mCols(Col).EditCtrl = Ctrl
    mCols(Col).MoveControl = MoveControl
End Sub

Private Function BlendColor(ByVal oColorFrom As OLE_COLOR, ByVal oColorTo As OLE_COLOR, Optional ByVal Alpha As Long = 128) As Long
    Dim lCFrom As Long
    Dim lCTo   As Long
    Dim lSrcR  As Long
    Dim lSrcG  As Long
    Dim lSrcB  As Long
    Dim lDstR  As Long
    Dim lDstG  As Long
    Dim lDstB  As Long
 
    lCFrom = oColorFrom
    lCTo = oColorTo
    lSrcR = lCFrom And &HFF
    lSrcG = (lCFrom And &HFF00&) \ &H100&
    lSrcB = (lCFrom And &HFF0000) \ &H10000
    lDstR = lCTo And &HFF
    lDstG = (lCTo And &HFF00&) \ &H100&
    lDstB = (lCTo And &HFF0000) \ &H10000
    BlendColor = RGB(((lSrcR * Alpha) / 255) + ((lDstR * (255 - Alpha)) / 255), ((lSrcG * Alpha) / 255) + ((lDstG * (255 - Alpha)) / 255), ((lSrcB * Alpha) / 255) + ((lDstB * (255 - Alpha)) / 255))
End Function


Public Property Get BorderStyle() As lgBorderStyleEnum
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BorderStyle = mBorderStyle
End Property

Public Property Let BorderStyle(ByVal NewValue As lgBorderStyleEnum)
    mBorderStyle = NewValue
    UserControl.BorderStyle = mBorderStyle
    
    PropertyChanged "BorderStyle"
End Property

Public Property Get CacheIncrement() As Long
    CacheIncrement = mCacheIncrement
End Property

Public Property Let CacheIncrement(ByVal NewValue As Long)
    If NewValue < 0 Then
        mCacheIncrement = 1
    Else
        mCacheIncrement = NewValue
    End If
    
    PropertyChanged "CacheIncrement"
End Property

Public Property Let CellAlignment(ByVal Row As Long, ByVal Col As Long, NewValue As lgAlignmentEnum)
    mItems(mRowPtr(Row)).Cell(mColPtr(Col)).nAlignment = NewValue
    DrawGrid mRedraw
End Property

Public Property Get CellAlignment(ByVal Row As Long, ByVal Col As Long) As lgAlignmentEnum
    CellAlignment = mItems(mRowPtr(Row)).Cell(mColPtr(Col)).nAlignment
End Property

Public Property Let CellBackColor(ByVal Row As Long, ByVal Col As Long, NewValue As Long)
    ApplyCellFormat Row, Col, lgCFBackColor, NewValue
    
    DrawGrid mRedraw
End Property

Public Property Get CellBackColor(ByVal Row As Long, ByVal Col As Long) As Long
    CellBackColor = mCF(mItems(mRowPtr(Row)).Cell(mColPtr(Col)).nFormat).lBackColor
End Property

Public Property Let CellChecked(ByVal Row As Long, ByVal Col As Long, NewValue As Boolean)
    SetFlag mItems(mRowPtr(Row)).Cell(mColPtr(Col)).nFlags, lgFLChecked, NewValue
    DrawGrid mRedraw
End Property

Public Property Get CellChecked(ByVal Row As Long, ByVal Col As Long) As Boolean
    CellChecked = mItems(mRowPtr(Row)).Cell(mColPtr(Col)).nFlags And lgFLChecked
End Property

Public Property Let CellChanged(ByVal Row As Long, ByVal Col As Long, NewValue As Boolean)
    SetFlag mItems(mRowPtr(Row)).Cell(mColPtr(Col)).nFlags, lgFLChanged, NewValue
End Property

Public Property Get CellChanged(ByVal Row As Long, ByVal Col As Long) As Boolean
    CellChanged = mItems(mRowPtr(Row)).Cell(mColPtr(Col)).nFlags And lgFLChanged
End Property

Public Property Let CellFontBold(ByVal Row As Long, ByVal Col As Long, NewValue As Boolean)
    SetFlag mItems(mRowPtr(Row)).Cell(mColPtr(Col)).nFlags, lgFLFontBold, NewValue
    DrawGrid mRedraw
End Property

Public Property Get TrackEdits() As Boolean
    TrackEdits = mTrackEdits
End Property

Public Property Let TrackEdits(ByVal NewValue As Boolean)
    mTrackEdits = NewValue
    
    PropertyChanged "TrackEdits"
End Property

Public Property Get CellFontBold(ByVal Row As Long, ByVal Col As Long) As Boolean
    CellFontBold = mItems(mRowPtr(Row)).Cell(mColPtr(Col)).nFlags And lgFLFontBold
End Property

Public Property Let CellFontItalic(ByVal Row As Long, ByVal Col As Long, NewValue As Boolean)
    SetFlag mItems(mRowPtr(Row)).Cell(mColPtr(Col)).nFlags, lgFLFontItalic, NewValue
    DrawGrid mRedraw
End Property

Public Property Get CellFontItalic(ByVal Row As Long, ByVal Col As Long) As Boolean
    CellFontItalic = mItems(mRowPtr(Row)).Cell(mColPtr(Col)).nFlags And lgFLFontItalic
End Property

Public Property Let CellFontUnderline(ByVal Row As Long, ByVal Col As Long, NewValue As Boolean)
    SetFlag mItems(mRowPtr(Row)).Cell(mColPtr(Col)).nFlags, lgFLFontUnderline, NewValue
    DrawGrid mRedraw
End Property

Public Property Get CellFontUnderline(ByVal Row As Long, ByVal Col As Long) As Boolean
    CellFontUnderline = mItems(mRowPtr(Row)).Cell(mColPtr(Col)).nFlags And lgFLFontUnderline
End Property

Public Property Let CellForeColor(ByVal Row As Long, ByVal Col As Long, NewValue As Long)
    ApplyCellFormat Row, Col, lgCFForeColor, NewValue
    DrawGrid mRedraw
End Property

Public Property Get CellForeColor(ByVal Row As Long, ByVal Col As Long) As Long
    CellForeColor = mCF(mItems(mRowPtr(Row)).Cell(mColPtr(Col)).nFormat).lForeColor
End Property

Public Property Let CellFontName(ByVal Row As Long, ByVal Col As Long, NewValue As String)
    ApplyCellFormat Row, Col, lgCFFontName, NewValue
    DrawGrid mRedraw
End Property

Public Property Get CellFontName(ByVal Row As Long, ByVal Col As Long) As String
    CellFontName = mCF(mItems(mRowPtr(Row)).Cell(mColPtr(Col)).nFormat).sFontName
End Property

Public Property Let CellImage(ByVal Row As Long, ByVal Col As Long, NewValue As Variant)
    Dim nImage As Integer
    
    On Local Error GoTo ItemImageError
    
    If IsNumeric(NewValue) Then
        nImage = NewValue
    Else
        nImage = -mImageList.ListImages(NewValue).Index
    End If
    
    ApplyCellFormat Row, Col, lgCFImage, nImage
    DrawGrid mRedraw
    Exit Property
    
ItemImageError:
    ApplyCellFormat Row, Col, lgCFImage, 0
End Property

Public Property Get CellImage(ByVal Row As Long, ByVal Col As Long) As Variant
    Dim nImage As Integer
    
    nImage = mCF(mItems(mRowPtr(Row)).Cell(mColPtr(Col)).nFormat).nImage
    
    If nImage >= 0 Then
        CellImage = nImage
    Else
        CellImage = mImageList.ListImages(Abs(nImage)).Key
    End If
End Property

Public Property Let CellProgressValue(ByVal Row As Long, ByVal Col As Long, NewValue As Integer)
    If mCols(Col).nType = lgProgressBar Then
        If NewValue > 100 Then
            NewValue = 100
        ElseIf NewValue < 0 Then
            NewValue = 0
        End If
        
        mItems(mRowPtr(Row)).Cell(mColPtr(Col)).nFlags = NewValue
        DrawGrid mRedraw
    End If
End Property

Public Property Get CellProgressValue(ByVal Row As Long, ByVal Col As Long) As Integer
    If mCols(Col).nType = lgProgressBar Then
        CellProgressValue = mItems(mRowPtr(Row)).Cell(mColPtr(Col)).nFlags
    End If
End Property

Public Property Let CellText(ByVal Row As Long, ByVal Col As Long, NewValue As String)
    mItems(mRowPtr(Row)).Cell(mColPtr(Col)).sValue = NewValue
    SetRowSize Row
    
    If mTrackEdits Then
        CellChanged(Row, Col) = True
    End If
    
    DrawGrid mRedraw
End Property

Public Property Get CellText(ByVal Row As Long, ByVal Col As Long) As String
    CellText = mItems(mRowPtr(Row)).Cell(mColPtr(Col)).sValue
End Property

Public Property Let CellWordWrap(ByVal Row As Long, ByVal Col As Long, NewValue As Boolean)
    SetFlag mItems(mRowPtr(Row)).Cell(mColPtr(Col)).nFlags, lgFLWordWrap, NewValue
    DrawGrid mRedraw
End Property

Public Property Get CellWordWrap(ByVal Row As Long, ByVal Col As Long) As Boolean
    CellWordWrap = mItems(mRowPtr(Row)).Cell(mColPtr(Col)).nFlags And lgFLFontItalic
End Property

Public Property Get CheckBoxes() As Boolean
Attribute CheckBoxes.VB_ProcData.VB_Invoke_Property = ";Behavior"
    CheckBoxes = mCheckboxes
End Property

Public Property Let CheckBoxes(ByVal NewValue As Boolean)
    mCheckboxes = NewValue
    DisplayChange
    
    PropertyChanged "CheckBoxes"
End Property

Public Function CheckedCount() As Long
    '#############################################################################################################################
    'Purpose: Return Count of Checked Items
    '#############################################################################################################################
    
    Dim lCount As Long
    
    For lCount = LBound(mItems) To UBound(mItems)
        If mItems(lCount).nFlags And lgFLChecked Then
            CheckedCount = CheckedCount + 1
        End If
    Next lCount
End Function

Public Sub Clear()
    '#############################################################################################################################
    'Purpose: Remove all Items from the Grid. Does not affect Column Headers
    '#############################################################################################################################
  
    ReDim mItems(0)
    ReDim mRowPtr(0)
    ReDim mCF(0)
    
    mMouseDownCol = NULL_RESULT
    mMouseDownRow = NULL_RESULT
    
    mCol = NULL_RESULT
    mRow = NULL_RESULT
    mSelectedRow = NULL_RESULT
    
    mHotColumn = NULL_RESULT
    mDragCol = NULL_RESULT
    mResizeCol = NULL_RESULT
    
    mSortColumn = NULL_RESULT
    mSortSubColumn = NULL_RESULT
    
    mScrollAction = SCROLL_NONE
    mItemCount = -1
    
    DrawGrid True
End Sub

Public Property Get Col() As Long
    Col = mCol
End Property

Public Property Let Col(ByVal NewValue As Long)
    If SetRowCol(mRow, NewValue) Then
        DrawGrid mRedraw
    End If
End Property

Public Property Get ColAlignment(ByVal Index As Long) As lgAlignmentEnum
    ColAlignment = mCols(Index).nAlignment
End Property

Public Property Let ColAlignment(ByVal Index As Long, ByVal NewValue As lgAlignmentEnum)
    mCols(Index).nAlignment = NewValue
    
    DrawGrid mRedraw
End Property

Public Property Get ColImageAlignment(ByVal Index As Long) As lgAlignmentEnum
    ColImageAlignment = mCols(Index).nImageAlignment
End Property

Public Property Let ColImageAlignment(ByVal Index As Long, ByVal NewValue As lgAlignmentEnum)
    mCols(Index).nImageAlignment = NewValue
    
    DrawGrid mRedraw
End Property

Public Property Get ColFormat(ByVal Index As Long) As String
    ColFormat = mCols(Index).sFormat
End Property

Public Property Let ColFormat(ByVal Index As Long, ByVal NewValue As String)
    mCols(Index).sFormat = NewValue
    
    DrawGrid mRedraw
End Property

Public Property Get ColHeading(ByVal Index As Long) As String
    ColHeading = mCols(Index).sCaption
End Property

Public Property Let ColHeading(ByVal Index As Long, ByVal NewValue As String)
    mCols(Index).sCaption = NewValue
    
    DrawGrid mRedraw
End Property

Public Property Get ColInputFilter(ByVal Index As Long) As String
    ColInputFilter = mCols(Index).sInputFilter
End Property

Public Property Let ColInputFilter(ByVal Index As Long, ByVal NewValue As String)
    mCols(Index).sInputFilter = NewValue
End Property

Public Function ColLeft(ByVal Index As Long) As Long
    Dim R As RECT
    
    SetColRect Index, R
    ColLeft = R.Left
End Function

Public Property Get ColPosition(ByVal Index As Long) As Long
    ColPosition = mColPtr(Index)
End Property

Public Property Let ColPosition(ByVal Index As Long, ByVal NewValue As Long)
    Dim lTemp As Long
     
    If (mColPtr(Index) <> NewValue) Then
        lTemp = mColPtr(Index)
        mColPtr(Index) = NewValue
        mColPtr(NewValue) = lTemp
        
        DrawGrid mRedraw
    End If
End Property

Public Property Get Cols() As Long
    Cols = UBound(mCols)
End Property

Public Property Let Cols(ByVal NewValue As Long)
    ReDim mCols(NewValue)
End Property

Public Property Get ColType(ByVal Index As Long) As lgDataTypeEnum
    ColType = mCols(Index).nType
End Property

Public Property Let ColType(ByVal Index As Long, ByVal NewValue As lgDataTypeEnum)
    mCols(Index).nType = NewValue
End Property

Public Property Get ColWordWrap(ByVal Index As Long) As Boolean
    ColWordWrap = mCols(Index).nFlags And lgFLWordWrap
End Property

Public Property Let ColWordWrap(ByVal Index As Long, ByVal NewValue As Boolean)
    SetFlag mCols(Index).nFlags, lgFLWordWrap, NewValue
End Property

Public Property Get ColumnDrag() As Boolean
    ColumnDrag = mColumnDrag
End Property

Public Property Let ColumnDrag(ByVal NewValue As Boolean)
    mColumnDrag = NewValue
    
    PropertyChanged "ColumnDrag"
End Property

Public Property Get ColumnSort() As Boolean
    ColumnSort = mColumnSort
End Property

Public Property Let ColumnSort(ByVal NewValue As Boolean)
    mColumnSort = NewValue
    
    PropertyChanged "ColumnSort"
End Property

Public Property Get ColTag(ByVal Index As Long) As String
    ColTag = mCols(Index).sTag
End Property

Public Property Let ColTag(ByVal Index As Long, ByVal NewValue As String)
    mCols(Index).sTag = NewValue
End Property

Public Property Get ColVisible(ByVal Index As Long) As Boolean
    ColVisible = mCols(Index).bVisible
End Property

Public Property Let ColVisible(ByVal Index As Long, ByVal NewValue As Boolean)
    mCols(Index).bVisible = NewValue
    
    DrawGrid mRedraw
End Property

Public Property Get ColWidth(ByVal Index As Long) As Single
    ColWidth = mCols(Index).dCustomWidth
End Property

Public Property Let ColWidth(ByVal Index As Long, ByVal NewValue As Single)
    'dCustomWidth is in the Units the Control is operating in
    mCols(Index).dCustomWidth = NewValue
    mCols(Index).lWidth = ScaleX(NewValue, mScaleUnits, vbPixels)
       
    DrawGrid mRedraw
End Property

Private Sub CreateRenderData()
    '#############################################################################################################################
    'Purpose: Calculates rendering parameters & sets display options. Used
    'to prevent unneccesary recalculations when redrawing the Grid
    '#############################################################################################################################
   
    Dim lCount As Long
    Dim lSize As Long
    
    With mR
        lSize = ScaleY(mRowHeightMin, mScaleUnits, vbPixels)
        If lSize > MAX_CHECKBOXSIZE Then
            .CheckBoxSize = MAX_CHECKBOXSIZE
        Else
            .CheckBoxSize = lSize - 4
        End If

        If mCheckboxes Then
            .LeftText = .CheckBoxSize + 2
        Else
            .LeftImage = 0
            .LeftText = DEFAULT_LEFTTEXT
        End If
        
        .LeftImage = .LeftText
        
        If mImageList Is Nothing Then
            .ImageSpace = 0
        Else
            .ImageSpace = ((GetRowHeight() - mImageList.ImageHeight) / 2)
            .ImageHeight = mImageList.ImageHeight
            .ImageWidth = mImageList.ImageWidth
            For lCount = LBound(mItems) To UBound(mItems)
                If mItems(lCount).lImage <> 0 Then
                    .LeftText = .LeftText + mImageList.ImageWidth + 2
                    Exit For
                End If
            Next lCount
        End If
        
        .HeaderHeight = GetColumnHeadingHeight()
        .TextHeight = UserControl.TextHeight("A")
        
        If mDisplayEllipsis Then
            .DTFlag = DT_SINGLELINE Or DT_WORD_ELLIPSIS
        Else
            .DTFlag = DT_SINGLELINE
        End If
    End With
End Sub

Private Sub DisplayChange()
    If mRedraw Then
        Refresh
    Else
        mPendingRedraw = True
        mPendingScrollBar = True
    End If
End Sub

Public Property Get DisplayEllipsis() As Boolean
Attribute DisplayEllipsis.VB_ProcData.VB_Invoke_Property = ";Appearance"
    DisplayEllipsis = mDisplayEllipsis
End Property

Public Property Let DisplayEllipsis(ByVal NewValue As Boolean)
    mDisplayEllipsis = NewValue
    DisplayChange
    
    PropertyChanged "DisplayEllipsis"
End Property

Public Sub Sort(Optional Sort As Long = -1, Optional SortType As lgSortTypeEnum = -1, Optional SubSort As Long = -1, Optional SubSortType As lgSortTypeEnum = -1)
    '#############################################################################################################################
    'Purpose: Sort Grid based on current Sort Columns.
    '#############################################################################################################################
    
    Dim lCount As Long
    Dim lRowIndex As Long
    
    If UpdateCell() Then
        'Set new Columns if specified
        If Sort <> -1 Then
            mSortColumn = Sort
        End If
        
        If SubSort <> -1 Then
            mSortSubColumn = SubSort
        End If
        
        'Validate Sort Columns
        If (mSortColumn = NULL_RESULT) And (mSortSubColumn <> NULL_RESULT) Then
            mSortColumn = mSortSubColumn
            mSortSubColumn = NULL_RESULT
        ElseIf mSortColumn = mSortSubColumn Then
            mSortSubColumn = NULL_RESULT
        End If
        
        'Set Sort Order if specified - otherwise inverse last Sort Order
        With mCols(mSortColumn)
            If SortType = -1 Then
                If .nSortOrder = lgSTAscending Then
                    .nSortOrder = lgSTDescending
                Else
                    .nSortOrder = lgSTAscending
                End If
            Else
                .nSortOrder = SortType
            End If
        End With
        
        If mSortSubColumn <> NULL_RESULT Then
            With mCols(mSortSubColumn)
                If SubSortType = -1 Then
                    If .nSortOrder = lgSTAscending Then
                        .nSortOrder = lgSTDescending
                    Else
                        .nSortOrder = lgSTAscending
                    End If
                Else
                    .nSortOrder = SubSortType
                End If
            End With
        End If
        
        'Note previously selected Row
        If mRow > NULL_RESULT Then
            lRowIndex = mRowPtr(mRow)
        End If
        
        SortArray LBound(mItems), mItemCount, mSortColumn, mCols(mSortColumn).nSortOrder
        SortSubList
        
        For lCount = LBound(mRowPtr) To mItemCount
            If mRowPtr(lCount) = lRowIndex Then
                mRow = lCount
                Exit For
            End If
        Next lCount
        
        DrawGrid True
        
        RaiseEvent SortComplete
    End If
End Sub



Private Sub DrawGrid(bRedraw As Boolean)
    '#############################################################################################################################
    'Purpose: The Primary Rendering routine. Draws Columns & Rows
    '#############################################################################################################################

    Dim IR As RECT
    Dim R As RECT
    Dim lX As Long
    Dim lY As Long
    
    Dim lCol As Long
    Dim lRow As Long
    Dim lMaxRow As Long
    Dim lStartCol As Long
    Dim lColumnsWidth As Long
    Dim lBottomEdge As Long
    Dim lGridColor As Long
    Dim lImageLeft As Long
    Dim lRowWrapSize As Long
    Dim lStart As Long
    Dim lValue As Long
    Dim nImage As Integer
    Dim bLockColor As Boolean
    Dim sText As String
    Dim bBold As Boolean
    Dim bItalic As Boolean
    Dim bUnderLine As Boolean
    Dim sFontName As String
    
    If bRedraw Then
        lStartCol = SBValue(efsHorizontal)
        lGridColor = TranslateColor(mGridColor)
        
        lY = mR.HeaderHeight
        mItemsVisible = ItemsVisible()
        lRowWrapSize = (mR.TextHeight * 2)
        
        With UserControl
            .Cls
            
            bBold = .FontBold
            bItalic = .FontItalic
            bUnderLine = .FontUnderline
            sFontName = .FontName
            
            lColumnsWidth = DrawHeaderRow()
            
            lMaxRow = (SBValue(efsVertical) + mItemsVisible)
            If lMaxRow > mItemCount Then
                lMaxRow = mItemCount
            End If
                
            lStart = SBValue(efsVertical)
            For lRow = lStart To lMaxRow
                If (mMultiSelect Or mFullRowSelect) And (mItems(mRowPtr(lRow)).nFlags And lgFLSelected) Then
                    bLockColor = True
                    If lStartCol = 0 Then ' ensure 1st column is visible
                        If mCols(0).lWidth < mR.LeftText Then
                            SetRect R, 0, lY + 1, mCols(0).lWidth, lY + (mItems(mRowPtr(lRow)).lHeight) = 1
                        Else
                            SetRect R, 0, lY + 1, mR.LeftText, lY + (mItems(mRowPtr(lRow)).lHeight) + 1
                        End If
                        DrawRect .hdc, R, TranslateColor(mBackColor), True
                    Else
                        R.Right = 0
                    End If
                    
                    SetRect R, R.Right, lY + 1, lColumnsWidth, lY + (mItems(mRowPtr(lRow)).lHeight) + 1
                    
                    If mAlphaBlendSelection Then
                       lValue = BlendColor(TranslateColor(mBackColorSel), TranslateColor(mBackColor), 120)
                    Else
                       lValue = TranslateColor(mBackColorSel)
                    End If
                    
                    DrawRect .hdc, R, lValue, True
                    
                    .ForeColor = mForeColorSel
                Else
                    bLockColor = False
                    SetRect R, 0, lY + 1, lColumnsWidth, lY + (mItems(mRowPtr(lRow)).lHeight) + 1
                    DrawRect .hdc, R, TranslateColor(mBackColor), True
                End If
                
                lX = 0
                For lCol = lStartCol To UBound(mCols)
                    If mCols(mColPtr(lCol)).bVisible Then
                        SetRectRgn mClipRgn, lX, lY, lX + mCols(mColPtr(lCol)).lWidth, lY + mItems(mRowPtr(lRow)).lHeight
                        SelectClipRgn .hdc, mClipRgn
                        
                        Call SetRect(R, lX, lY, lX + mCols(mColPtr(lCol)).lWidth, lY + mItems(mRowPtr(lRow)).lHeight)
                        
                        If Not bLockColor Then
                            If mCF(mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).nFormat).lBackColor <> mBackColor Then
                                DrawRect .hdc, R, TranslateColor(mCF(mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).nFormat).lBackColor), True
                            End If
                            .ForeColor = mCF(mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).nFormat).lForeColor
                        End If
                        
                        If lCol = 0 Then
                            If mCheckboxes Then
                                Call SetRect(R, 3, lY, mR.CheckBoxSize, lY + mItems(mRowPtr(lRow)).lHeight)
                                
                                If mItems(mRowPtr(lRow)).nFlags And lgFLChecked Then
                                    lValue = 5
                                Else
                                    lValue = 0
                                End If
                                
                                If Not DrawTheme("Button", 3, lValue, R) Then
                                    If mItems(mRowPtr(lRow)).nFlags And lgFLChecked Then
                                        Call DrawFrameControl(.hdc, R, DFC_BUTTON, DFCS_BUTTONCHECK Or DFCS_CHECKED Or DFCS_FLAT)
                                    Else
                                        Call DrawFrameControl(.hdc, R, DFC_BUTTON, DFCS_BUTTONCHECK Or DFCS_FLAT)
                                    End If
                                End If
                            End If
                            
                            If mR.ImageSpace > 0 Then
                                'If we have an Image Index then Draw it
                                If mItems(mRowPtr(lRow)).lImage <> 0 Then
                                    'Calculate Image offset (using ScaleMode of ImageList)
                                    If lImageLeft = 0 Then
                                        lImageLeft = ScaleX(mR.LeftImage, vbPixels, mImageListScaleMode)
                                    End If
                                    
                                    If bLockColor And mApplySelectionToImages Then
                                        mImageList.ListImages(Abs(mItems(mRowPtr(lRow)).lImage)).Draw .hdc, lImageLeft, ScaleY(lY + mR.ImageSpace, vbPixels, mImageListScaleMode), 2
                                    Else
                                        mImageList.ListImages(Abs(mItems(mRowPtr(lRow)).lImage)).Draw .hdc, lImageLeft, ScaleY(lY + mR.ImageSpace, vbPixels, mImageListScaleMode), 1
                                    End If
                                End If
                            End If
                            
                            Call SetRect(R, mR.LeftText + TEXT_SPACE, lY, (lX + mCols(mColPtr(lCol)).lWidth) - TEXT_SPACE, lY + mItems(mRowPtr(lRow)).lHeight)
                        Else
                            Call SetRect(R, lX + TEXT_SPACE, lY, (lX + mCols(mColPtr(lCol)).lWidth) - TEXT_SPACE, lY + mItems(mRowPtr(lRow)).lHeight)
                        End If
                       
                        Select Case mCols(mColPtr(lCol)).nType
                            Case lgBoolean
                                SetItemRect mRowPtr(lRow), mColPtr(lCol), lY, R, lgRTCheckBox
                                
                                If mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).nFlags And lgFLChecked Then
                                    lValue = 5
                                Else
                                    lValue = 0
                                End If
                                
                                If Not DrawTheme("Button", 3, lValue, R) Then
                                    If mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).nFlags And lgFLChecked Then
                                        Call DrawFrameControl(.hdc, R, DFC_BUTTON, DFCS_BUTTONCHECK Or DFCS_CHECKED Or DFCS_FLAT)
                                    Else
                                        Call DrawFrameControl(.hdc, R, DFC_BUTTON, DFCS_BUTTONCHECK Or DFCS_FLAT)
                                    End If
                                End If
                                   
                            Case lgProgressBar
                                If mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).nFlags > 0 Then
                                    lValue = ((mCols(mColPtr(lCol)).lWidth - 2) / 100) * mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).nFlags
                                
                                    SetRect R, lX + 2, lY + 2, lX + lValue, (lY + mItems(mRowPtr(lRow)).lHeight) - 2
                                    DrawRect .hdc, R, TranslateColor(mProgressBarColor), True
                                End If
                            
                            Case Else
                                UserControl.FontName = mCF(mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).nFormat).sFontName
                                
                                With mItems(mRowPtr(lRow)).Cell(mColPtr(lCol))
                                    UserControl.FontBold = .nFlags And lgFLFontBold
                                    UserControl.FontItalic = .nFlags And lgFLFontItalic
                                    UserControl.FontUnderline = .nFlags And lgFLFontUnderline
                                    
                                    If Len(mCols(mColPtr(lCol)).sFormat) > 0 Then
                                        sText = Format$(.sValue, mCols(mColPtr(lCol)).sFormat)
                                    Else
                                        sText = .sValue
                                    End If
                                    
                                    If mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).nFlags And lgFLWordWrap Then
                                        lValue = .nAlignment Or DT_WORDBREAK
                                        
                                        IR.Left = 0
                                        IR.Right = mCols(mColPtr(lCol)).lWidth
                                        Call DrawText(UserControl.hdc, sText, Len(sText), IR, DT_CALCRECT Or DT_WORDBREAK)

                                        If (IR.Bottom - IR.Top) > mR.TextHeight Then
                                            nImage = mExpandRowImage
                                        Else
                                            nImage = 0
                                        End If
                                        
                                        If mItems(mRowPtr(lRow)).lHeight < lRowWrapSize Then
                                            'Truncate Rect to prevent wrapped text from showing
                                            R.Bottom = R.Top + mR.TextHeight
                                        End If
                                    Else
                                        lValue = .nAlignment Or mR.DTFlag
                                        nImage = mCF(mItems(mRowPtr(lRow)).Cell(mColPtr(lCol)).nFormat).nImage
                                    End If
                                    
                                    If nImage <> 0 Then
                                        SetItemRect mRowPtr(lRow), mColPtr(lCol), lY, IR, lgRTImage
                                        
                                        If IR.Left >= 0 Then
                                            If bLockColor And mApplySelectionToImages Then
                                                mImageList.ListImages(Abs(nImage)).Draw UserControl.hdc, ScaleX(IR.Left, vbPixels, mImageListScaleMode), ScaleY(lY + mR.ImageSpace, vbPixels, mImageListScaleMode), 2
                                            Else
                                                mImageList.ListImages(Abs(nImage)).Draw UserControl.hdc, ScaleX(IR.Left, vbPixels, mImageListScaleMode), ScaleY(lY + mR.ImageSpace, vbPixels, mImageListScaleMode), 1
                                            End If
                                        End If
                                        
                                        'Adjust Text Rect
                                        Select Case mCols(mColPtr(lCol)).nImageAlignment
                                            Case lgAlignLeftTop, lgAlignLeftCenter, lgAlignLeftBottom
                                                R.Left = R.Left + (IR.Right - IR.Left)

                                            Case lgAlignRightTop, lgAlignRightCenter, lgAlignRightBottom
                                                R.Right = R.Right - (IR.Right - IR.Left)
                                        End Select
                                    End If
                                    
                                    Call DrawText(UserControl.hdc, sText, -1, R, lValue)
                                End With
                        End Select
                        
                        lX = lX + mCols(mColPtr(lCol)).lWidth
                    End If
                Next lCol
                
                SelectClipRgn .hdc, 0&
                
                'Display Horizontal Lines
                If mGridLines Then
                    DrawLine .hdc, 0, lY, lColumnsWidth, lY, lGridColor, mGridLineWidth
                End If
                
                lY = lY + mItems(mRowPtr(lRow)).lHeight
            Next lRow
            
            '#############################################################################################################################
            'Display Vertical Lines
            If mGridLines Then
                lBottomEdge = R.Bottom
            
                lX = 0
                For lCol = lStartCol To UBound(mCols)
                    If mCols(mColPtr(lCol)).bVisible Then
                        DrawLine .hdc, lX, mR.HeaderHeight, lX, lBottomEdge, lGridColor, mGridLineWidth
                    
                        lX = lX + mCols(mColPtr(lCol)).lWidth
                    End If
                Next lCol
            End If
            
            '#############################################################################################################################
            'Display Focus Rectangle
            If (mFocusRectMode <> lgFocusRectModeEnum.lgNone) And (mRow >= 0) Then
                If Not mHideSelection Or mInFocus Then
                    lY = RowTop(mRow)
                    If lY >= 0 Then
                        R.Right = 0
                        If mFocusRectMode = lgCol Then
                            SetColRect mCol, R
                            R.Top = lY + 1
                            R.Bottom = lY + mItems(mRowPtr(mRow)).lHeight
                        Else 'If mFullRowSelect Then
                            SetRect R, mR.LeftText, lY + 1, lColumnsWidth, lY + mItems(mRowPtr(mRow)).lHeight
                        End If
                        
                        If R.Right > 0 Then
                            Select Case mFocusRectStyle
                                Case lgFRLight
                                    Call DrawFocusRect(.hdc, R)
                    
                                Case lgFRHeavy
                                    DrawRect .hdc, R, TranslateColor(mFocusRectColor), False
                            End Select
                        End If
                    End If
                End If
            End If
            
            .Refresh
            
            .FontBold = bBold
            .FontItalic = bItalic
            .FontUnderline = bUnderLine
            .FontName = sFontName
        End With
        
        'Debug.Print "Drawgrid mRedraw " & Timer
        
        mPendingRedraw = False
    Else
        mPendingRedraw = True
    End If
End Sub

Public Sub FormatCells(ByVal RowFrom As Long, ByVal RowTo As Long, ByVal ColFrom As Long, ByVal ColTo As Long, Mode As lgCellFormatEnum, NewValue As Variant)
    Dim lCol As Long
    Dim lRow As Long
    
    Dim lValue As Long
    Dim bValue As Boolean
    Dim sValue As String
    
    Select Case Mode
        Case lgCFBackColor, lgCFForeColor, lgCFImage
            lValue = CLng(NewValue)
        Case lgCFFontName
            sValue = CStr(NewValue)
        Case lgCFFontBold, lgCFFontItalic, lgCFFontUnderline
            bValue = CBool(NewValue)
    End Select
    
    For lRow = RowFrom To RowTo
        For lCol = ColFrom To ColTo
            Select Case Mode
                Case lgCFBackColor
                    CellBackColor(lRow, lCol) = lValue
                Case lgCFForeColor
                    CellForeColor(lRow, lCol) = lValue
                Case lgCFImage
                    CellImage(lRow, lCol) = lValue
                Case lgCFFontName
                    CellFontName(lRow, lCol) = sValue
                Case lgCFFontBold
                    CellFontBold(lRow, lCol) = bValue
                Case lgCFFontItalic
                    CellFontItalic(lRow, lCol) = bValue
                Case lgCFFontUnderline
                    CellFontUnderline(lRow, lCol) = bValue
            End Select
        Next lCol
    Next lRow
End Sub

Private Function LongToSignedShort(dwUnsigned As Long) As Integer
   If dwUnsigned < 32768 Then
      LongToSignedShort = CInt(dwUnsigned)
   Else
      LongToSignedShort = CInt(dwUnsigned - &H10000)
   End If
End Function
Private Sub FillGradient(lhDC As Long, rRect As RECT, ByVal clrFirst As OLE_COLOR, ByVal clrSecond As OLE_COLOR, Optional ByVal bVertical As Boolean)
    Dim pVert(0 To 1)   As TRIVERTEX
    Dim pGradRect       As GRADIENT_RECT
    
    With pVert(0)
        .X = rRect.Left
        .Y = rRect.Top
        .Red = LongToSignedShort((clrFirst And &HFF&) * 256)
        .Green = LongToSignedShort(((clrFirst And &HFF00&) / &H100&) * 256)
        .Blue = LongToSignedShort(((clrFirst And &HFF0000) / &H10000) * 256)
        .Alpha = 0
    End With
    
    With pVert(1)
        .X = rRect.Right
        .Y = rRect.Bottom
        .Red = LongToSignedShort((clrSecond And &HFF&) * 256)
        .Green = LongToSignedShort(((clrSecond And &HFF00&) / &H100&) * 256)
        .Blue = LongToSignedShort(((clrSecond And &HFF0000) / &H10000) * 256)
        .Alpha = 0
    End With
    
    With pGradRect
        .UPPERLEFT = 0
        .LOWERRIGHT = 1
    End With
    
    GradientFill lhDC, pVert(0), 2, pGradRect, 1, IIf(Not bVertical, GRADIENT_FILL_RECT_H, GRADIENT_FILL_RECT_V)
End Sub
    
Private Sub DrawHeader(lCol As Long, State As lgHeaderStateEnum)
    '#############################################################################################################################
    'Purpose: Renders a Column Header. This involves drawing the Border, displaying
    'the Caption and optionally Sort Arrows
    '#############################################################################################################################

    Dim R As RECT
    
    If lCol > NULL_RESULT Then
        With UserControl
            .ForeColor = mForeColorHdr
            
            'Draw the Column Headers
            Call SetRect(R, mCols(mColPtr(lCol)).lX, 0, mCols(mColPtr(lCol)).lX + mCols(mColPtr(lCol)).lWidth + 1, mR.HeaderHeight)
            DrawRect .hdc, R, TranslateColor(BackColorFixed), True
            
            Select Case mThemeStyle
                Case lgTSWindows3D
                    Select Case State
                         Case lgNormal
                             Call DrawFrameControl(.hdc, R, DFC_BUTTON, DFCS_BUTTONPUSH)
                         Case lgHot
                             Call DrawFrameControl(.hdc, R, DFC_BUTTON, DFCS_BUTTONPUSH Or DFCS_HOT)
                         Case lgDown
                             Call DrawFrameControl(.hdc, R, DFC_BUTTON, DFCS_BUTTONPUSH Or DFCS_PUSHED)
                     End Select
             
                Case lgTSWindowsFlat
                    Select Case State
                         Case lgNormal
                             Call DrawFrameControl(.hdc, R, DFC_BUTTON, DFCS_BUTTONPUSH Or DFCS_FLAT)
                         Case lgHot
                             Call DrawFrameControl(.hdc, R, DFC_BUTTON, DFCS_BUTTONPUSH Or DFCS_HOT)
                         Case lgDown
                             Call DrawFrameControl(.hdc, R, DFC_BUTTON, DFCS_BUTTONPUSH Or DFCS_PUSHED)
                     End Select
                
                Case lgTSWindowsXP
                    'Try XP Theme API
                    If Not DrawTheme("Header", 1, State, R) Then
                        'Use XP emulation
                        DrawXPHeader .hdc, R, State
                    End If
                
                Case lgTSOfficeXP
                    DrawOfficeXPHeader .hdc, R, State
                   
            End Select
            
            'Render Sort Arrows
            If mCols(mColPtr(lCol)).lWidth > SIZE_SORTARROW Then
                If mColPtr(lCol) = mSortColumn Then
                    DrawSortArrow (mCols(mColPtr(lCol)).lX + mCols(mColPtr(lCol)).lWidth) - 12, 6, 9, 5, mCols(mColPtr(lCol)).nSortOrder
                    
                    Call SetRect(R, mCols(mColPtr(lCol)).lX + HEADER_LEFT, 0, (mCols(mColPtr(lCol)).lX + mCols(mColPtr(lCol)).lWidth) - (ARROW_SPACE + SIZE_SORTARROW), mR.HeaderHeight)
                ElseIf mColPtr(lCol) = mSortSubColumn Then
                    DrawSortArrow (mCols(mColPtr(lCol)).lX + mCols(mColPtr(lCol)).lWidth) - 12, 6, 6, 3, mCols(mColPtr(lCol)).nSortOrder
                    
                    Call SetRect(R, mCols(mColPtr(lCol)).lX + HEADER_LEFT, 0, (mCols(mColPtr(lCol)).lX + mCols(mColPtr(lCol)).lWidth) - (ARROW_SPACE + SIZE_SORTARROW), mR.HeaderHeight)
                Else
                    Call SetRect(R, mCols(mColPtr(lCol)).lX + HEADER_LEFT, 0, (mCols(mColPtr(lCol)).lX + mCols(mColPtr(lCol)).lWidth) - (HEADER_LEFT * 2), mR.HeaderHeight)
                End If
            Else
                Call SetRect(R, mCols(mColPtr(lCol)).lX + HEADER_LEFT, 0, (mCols(mColPtr(lCol)).lX + mCols(mColPtr(lCol)).lWidth) - (HEADER_LEFT * 2), mR.HeaderHeight)
            End If
            
            Call DrawText(.hdc, mCols(mColPtr(lCol)).sCaption, -1, R, mCols(mColPtr(lCol)).nAlignment Or mR.DTFlag)
            
        End With
    End If
End Sub

Public Property Get HideSelection() As Boolean
Attribute HideSelection.VB_ProcData.VB_Invoke_Property = ";Behavior"
    HideSelection = mHideSelection
End Property

Public Property Let HideSelection(ByVal NewValue As Boolean)
    mHideSelection = NewValue
    DisplayChange
    
    PropertyChanged "HideSelection"
End Property

Private Function DrawHeaderRow() As Long
    '#############################################################################################################################
    'Purpose: Renders all Column Headers
    '#############################################################################################################################
    
    Dim lCol As Long
    Dim lX As Long
    
    mHotColumn = NULL_RESULT
    
    For lCol = SBValue(efsHorizontal) To UBound(mCols)
         If mCols(mColPtr(lCol)).bVisible Then
            mCols(mColPtr(lCol)).lX = lX
            DrawHeader lCol, lgNormal
            lX = lX + mCols(mColPtr(lCol)).lWidth
        End If
    Next lCol
    
    DrawHeaderRow = lX
End Function

Private Function InvertThisColor(oInsColor As OLE_COLOR)
    '#############################################################################################################################
    'Source: Riccardo Cohen
    '#############################################################################################################################
    
    Dim lROut As Long, lGOut As Long, lBOut As Long
    Dim lRGB As Long
   
    lRGB = TranslateColor(oInsColor)
    
    lROut = (255 - (lRGB And &HFF&))
    lGOut = (255 - ((lRGB And &HFF00&) / &H100))
    lBOut = (255 - ((lRGB And &HFF0000) / &H10000))
    InvertThisColor = RGB(lROut, lGOut, lBOut)
End Function


Private Sub DrawLine(hdc As Long, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, lcolor As Long, lWidth As Long)
    Dim PT As POINTAPI
    Dim hPen As Long
    Dim hPenOld As Long
    
    hPen = CreatePen(0, lWidth, lcolor)
    hPenOld = SelectObject(hdc, hPen)
    MoveToEx hdc, X1, Y1, PT
    LineTo hdc, X2, Y2
    SelectObject hdc, hPenOld
    DeleteObject hPen
End Sub

Private Sub DrawOfficeXPHeader(lhDC As Long, rRect As RECT, State As lgHeaderStateEnum)
    '#############################################################################################################################
    'Purpose:   Draw a Column Header in Office XP Style
    'Notes:     Created from original source by Riccardo Cohen
    '#############################################################################################################################
    
    With rRect
        Select Case State
            Case lgNormal
                Call FillGradient(lhDC, rRect, &HFCE1CB, &HE0A57D, True)
                
                DrawLine lhDC, .Left, .Top, .Right, .Top, &H9C613B, 1
                DrawLine lhDC, .Left, .Bottom - 1, .Right, .Bottom - 1, &H9C613B, 1
                
                DrawLine lhDC, .Right - 2, .Top + 2, .Right - 2, .Bottom - 4, &HCB8C6A, 1
                DrawLine lhDC, .Left, .Top + 3, .Left, .Bottom - 3, &HFFFFFF, 1

            Case lgHot
                .Right = .Right - 1
                Call FillGradient(lhDC, rRect, &HDCFFFF, &H5BC0F7, True)
                
                DrawLine lhDC, .Left, .Top, .Right, .Top, &H9C613B, 1
                DrawLine lhDC, .Left, .Bottom - 1, .Right, .Bottom - 1, &H9C613B, 1
                
                DrawLine lhDC, .Left, .Top + 3, .Left, .Bottom - 3, &HFFFFFF, 1

            Case lgDown
                .Right = .Right - 1
                Call FillGradient(lhDC, rRect, &H87FE8, &H7CDAF7, True)
                
                DrawLine lhDC, .Left, .Top, .Right, .Top, &H9C613B, 1
                DrawLine lhDC, .Left, .Bottom - 1, .Right, .Bottom - 1, &H9C613B, 1
                
                DrawLine lhDC, .Left, .Top + 3, .Left, .Bottom - 3, &HFFFFFF, 1
                
        End Select
    End With
End Sub

Private Sub DrawXPHeader(lhDC As Long, rRect As RECT, State As lgHeaderStateEnum)
    '#############################################################################################################################
    'Purpose:   Draw a Column Header in XP Style
    'Notes:     Created from original source by Riccardo Cohen
    '#############################################################################################################################
    
    Dim TempColor As OLE_COLOR

    With rRect
        Select Case State
            Case lgNormal
                DrawRect lhDC, rRect, TranslateColor(vbButtonFace), True
        
                DrawLine lhDC, .Left, .Bottom - 1, .Right, .Bottom - 1, &HB2C2C5, 1
                DrawLine lhDC, .Left, .Bottom - 2, .Right, .Bottom - 2, &HBECFD2, 1
                DrawLine lhDC, .Left, .Bottom - 3, .Right, .Bottom - 3, &HC8D8DC, 1
                
                DrawLine lhDC, .Right - 2, .Top + 2, .Right - 2, .Bottom - 4, &H99A8AC, 1
                DrawLine lhDC, .Left, .Top + 2, .Left, .Bottom - 4, &HFFFFFF, 1
                
            Case lgHot
                DrawRect lhDC, rRect, &HF3F8FA, True
                
                DrawLine lhDC, .Left + 2, .Bottom - 1, .Right - 2, .Bottom - 1, &H19B1F9, 1
                DrawLine lhDC, .Left + 1, .Bottom - 2, .Right - 1, .Bottom - 2, &H47C2FC, 1
                DrawLine lhDC, .Left, .Bottom - 3, .Right, .Bottom - 3, 43512, 1

            Case lgDown
                TempColor = ForeColor
                
                UserControl.ForeColor = InvertThisColor(TempColor)
                .Bottom = .Bottom - 1
                DrawRect lhDC, rRect, &H0&, True
                
                DrawLine lhDC, .Left, .Bottom - 1, .Right, .Bottom - 1, InvertThisColor(&HB2C2C5), 1
                DrawLine lhDC, .Left, .Bottom - 2, .Right, .Bottom - 2, InvertThisColor(&HBECFD2), 1
                DrawLine lhDC, .Left, .Bottom - 3, .Right, .Bottom - 3, InvertThisColor(&HC8D8DC), 1
                DrawLine lhDC, .Right - 2, .Top + 2, .Right - 2, .Bottom - 4, InvertThisColor(&H99A8AC), 1
                DrawLine lhDC, .Left, .Top + 2, .Left, .Bottom - 4, InvertThisColor(&HFFFFFF), 1
        End Select
    End With
End Sub


Private Sub DrawRect(hdc As Long, rc As RECT, lcolor As Long, bFilled As Boolean)
    Dim lNewBrush As Long
  
    lNewBrush = CreateSolidBrush(lcolor)
    
    If bFilled Then
        Call FillRect(hdc, rc, lNewBrush)
    Else
        Call FrameRect(hdc, rc, lNewBrush)
    End If

    Call DeleteObject(lNewBrush)
End Sub

Private Sub DrawSortArrow(lX As Long, lY As Long, lWidth As Long, lStep As Long, nOrientation As lgSortTypeEnum)
    '#############################################################################################################################
    'Purpose: Renders the Sort/Sub-Sort arrows
    '#############################################################################################################################
   
    Dim hPenOld As Long
    Dim hPen As Long
    Dim lCount As Long
    Dim lVerticalChange As Long
    Dim X1 As Long
    Dim X2 As Long
    Dim Y1 As Long
    
    hPen = CreatePen(0, 1, TranslateColor(vbButtonShadow))
    hPenOld = SelectObject(hdc, hPen)
    
    If nOrientation = lgSTDescending Then
        lVerticalChange = -1
        lY = lY + lStep - 1
    Else
        lVerticalChange = 1
    End If
    
    X1 = lX
    X2 = lWidth
    Y1 = lY
        
    MoveTo hdc, X1, Y1, ByVal 0&
    
    For lCount = 1 To lStep
        LineTo hdc, X1 + X2, Y1
        X1 = X1 + 1
        Y1 = Y1 + lVerticalChange
        X2 = X2 - 2
        MoveTo hdc, X1, Y1, ByVal 0&
    Next lCount
    
    Call SelectObject(hdc, hPenOld)
    Call DeleteObject(hPen)
End Sub

Private Sub DrawText(ByVal hdc As Long, ByVal lpString As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long)
    '#############################################################################################################################
    'Purpose: Renders the Text for Column Headers & Cells. On Windows NT/2000/XP
    '(or better) the Control supports Unicode
    '#############################################################################################################################
   
    If mWinNT Then
        DrawTextW hdc, StrPtr(lpString), nCount, lpRect, wFormat
    Else
        DrawTextA hdc, lpString, nCount, lpRect, wFormat
    End If
End Sub

Private Function DrawTheme(sClass As String, ByVal iPart As Long, ByVal iState As Long, rtRect As RECT, Optional ByVal CloseTheme As Boolean = False) As Boolean
    '#############################################################################################################################
    'Purpose: On Windows XP allows certain elements of the Grid to be drawn using
    'the current Windows Theme
    '#############################################################################################################################
    
    Dim lResult As Long
    
    On Error GoTo DrawThemeError
    
    If mWinXP Then
        If (mThemeStyle = lgTSWindowsXP) Or (mThemeStyle = lgTSOfficeXP) Then
            hTheme = OpenThemeData(UserControl.hwnd, StrPtr(sClass))
            If (hTheme) Then
                lResult = DrawThemeBackground(hTheme, UserControl.hdc, iPart, iState, rtRect, rtRect)
                DrawTheme = (lResult = 0)
            Else
                DrawTheme = False
            End If
            
            If CloseTheme Then
                Call CloseThemeData(hTheme)
            End If
        End If
    End If
    Exit Function

DrawThemeError:
    DrawTheme = False
End Function

Public Property Get Editable() As Boolean
Attribute Editable.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Editable = mEditable
End Property

Public Property Let Editable(ByVal NewValue As Boolean)
    mEditable = NewValue
    
    PropertyChanged "Editable"
End Property

Public Sub EditCell(ByVal Row As Long, ByVal Col As Long)
    '#############################################################################################################################
    'Purpose: Used to start an Edit. Note the RequestEdit event. This event allows
    'the Edit to be cancelled before anything visible occurs by setting the Cancel
    'flag.
    '#############################################################################################################################
   
    Dim bCancel As Boolean
    
    If mEditPending Then
        If Not UpdateCell() Then
            Exit Sub
        End If
    End If
    
    If IsEditable() And (mCols(mColPtr(Col)).nType <> lgBoolean) Then
        RaiseEvent RequestEdit(Row, Col, bCancel)
        If Not bCancel Then
            mEditCol = Col
            mEditRow = Row
            
            MoveEditControl mCols(mColPtr(mEditCol)).MoveControl
            
            'Check if an external Control is used.
            If mCols(mColPtr(mEditCol)).EditCtrl Is Nothing Then
                'Using internal TextBox
                With txtEdit
                    Select Case mItems(mRowPtr(mEditRow)).Cell(mColPtr(mEditCol)).nAlignment
                        Case lgAlignCenterBottom, lgAlignCenterCenter, lgAlignCenterTop
                            .Alignment = vbCenter
                        Case lgAlignLeftBottom, lgAlignLeftCenter, lgAlignLeftTop
                            .Alignment = vbLeftJustify
                        Case Else
                            .Alignment = vbRightJustify
                    End Select
                    
                    If mWinNT Then
                        Select Case mCols(mColPtr(mEditCol)).sInputFilter
                            Case "<"
                                Call SetWindowLongW(.hwnd, GWL_STYLE, mTextBoxStyle Or ES_LOWERCASE)
                            Case ">"
                                Call SetWindowLongW(.hwnd, GWL_STYLE, mTextBoxStyle Or ES_UPPERCASE)
                            Case Else
                                Call SetWindowLongW(.hwnd, GWL_STYLE, mTextBoxStyle)
                        End Select
                    Else
                        Select Case mCols(mColPtr(mEditCol)).sInputFilter
                            Case "<"
                                Call SetWindowLongA(.hwnd, GWL_STYLE, mTextBoxStyle Or ES_LOWERCASE)
                            Case ">"
                                Call SetWindowLongA(.hwnd, GWL_STYLE, mTextBoxStyle Or ES_UPPERCASE)
                            Case Else
                                Call SetWindowLongA(.hwnd, GWL_STYLE, mTextBoxStyle)
                        End Select
                    End If
                    
                    .BackColor = mBackColorEdit
                    .FontBold = mItems(mRowPtr(mEditRow)).Cell(mColPtr(mEditCol)).nFlags And lgFLFontBold
                    .FontItalic = mItems(mRowPtr(mEditRow)).Cell(mColPtr(mEditCol)).nFlags And lgFLFontItalic
                    .FontUnderline = mItems(mRowPtr(mEditRow)).Cell(mColPtr(mEditCol)).nFlags And lgFLFontUnderline
                    
                    .Text = mItems(mRowPtr(mEditRow)).Cell(mColPtr(mEditCol)).sValue
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .Visible = True
                    .SetFocus
                End With
            Else
                On Local Error Resume Next
                
                With mCols(mColPtr(mEditCol)).EditCtrl
                    If UserControl.ContainerHwnd <> .Container.hwnd Then
                        mEditParent = UserControl.ContainerHwnd
                        SetParent .hwnd, UserControl.ContainerHwnd
                    Else
                        mEditParent = 0
                    End If
                    .Enabled = True
                    .Visible = True
                    .ZOrder
                    
                    Subclass_Start .hwnd
                    Call Subclass_AddMsg(.hwnd, WM_KILLFOCUS, MSG_AFTER)
                    
                    If TypeOf mCols(mColPtr(mEditCol)).EditCtrl Is VB.ComboBox Then
                        SendMessageAsLong mCols(mColPtr(mEditCol)).EditCtrl.hwnd, CB_SHOWDROPDOWN, 1&, 0&
                    End If
                    
                    .SetFocus
                End With
                
                On Local Error GoTo 0
            End If
            
            mEditPending = True
        End If
    End If
End Sub

Public Property Get EditTrigger() As lgEditTriggerEnum
Attribute EditTrigger.VB_ProcData.VB_Invoke_Property = ";Behavior"
    EditTrigger = mEditTrigger
End Property

Public Property Let EditTrigger(ByVal NewValue As lgEditTriggerEnum)
    mEditTrigger = NewValue
    
    PropertyChanged "EditTrigger"
End Property

Public Function FindItem(ByVal SearchText As String, Optional ByVal SearchColumn As Long = -1, Optional SearchMode As lgSearchModeEnum = lgSMEqual, Optional MatchCase As Boolean) As Long
    '#############################################################################################################################
    'Purpose: Search the specified Column for a Cell that matches the search text
    
    'SearchText     - The text to look for
    'SearchColumn   - The Column to search in (defaults to the SearchColumn property if not specified)
    'SearchMode     - The type of search required. The lgSMNavigate mode is used by the Grid internally
    '               when searching for an entry that matches the keys the user is pressing.
    
    'MatchCase      - Specify a case sensitive or case insensitive search
    
    Dim lCount As Long
    Dim sCellText As String
    
    FindItem = NULL_RESULT
    
    If SearchColumn = -1 Then
        SearchColumn = mSearchColumn
    End If
    
    If (SearchColumn >= 0) And (Len(SearchText) > 0) Then
        If Not MatchCase Then
            SearchText = UCase$(SearchText)
        End If
        
        For lCount = LBound(mItems) To mItemCount
            If MatchCase Then
                sCellText = mItems(mRowPtr(lCount)).Cell(SearchColumn).sValue
            Else
                sCellText = UCase$(mItems(mRowPtr(lCount)).Cell(SearchColumn).sValue)
            End If
            
            Select Case SearchMode
                Case lgSMEqual
                    If sCellText = SearchText Then
                        FindItem = lCount
                        Exit For
                    End If
                
                Case lgSMGreaterEqual
                    If sCellText >= SearchText Then
                        FindItem = lCount
                        Exit For
                    End If
                
                Case lgSMLike
                    If sCellText Like SearchText & "*" Then
                        FindItem = lCount
                        Exit For
                    End If
                
                Case lgSMNavigate
                    If Len(sCellText) > 0 Then
                        If (sCellText >= SearchText) And ((Mid$(sCellText, 1, 1)) = Mid$(SearchText, 1, 1)) Then
                            FindItem = lCount
                            Exit For
                        End If
                    End If
    
            End Select
            
        Next lCount
    End If
End Function

Public Property Let ExpandRowImage(NewValue As Variant)
    On Local Error GoTo ExpandRowImageError
    
    If IsNumeric(NewValue) Then
        mExpandRowImage = NewValue
    Else
        mExpandRowImage = -mImageList.ListImages(NewValue).Index
    End If
    
    DrawGrid mRedraw
    Exit Property
    
ExpandRowImageError:
    Exit Property
End Property

Public Property Get ExpandRowImage() As Variant
Attribute ExpandRowImage.VB_MemberFlags = "400"
    If mExpandRowImage >= 0 Then
        ExpandRowImage = mExpandRowImage
    Else
        ExpandRowImage = mImageList.ListImages(Abs(mExpandRowImage)).Key
    End If
End Property

Public Property Let FocusRectColor(ByVal NewValue As OLE_COLOR)
    mFocusRectColor = NewValue
    
    PropertyChanged "FocusRectColor"
End Property

Public Property Get FocusRectColor() As OLE_COLOR
Attribute FocusRectColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    FocusRectColor = mFocusRectColor
End Property

Public Property Get FocusRectMode() As lgFocusRectModeEnum
Attribute FocusRectMode.VB_ProcData.VB_Invoke_Property = ";Appearance"
    FocusRectMode = mFocusRectMode
End Property

Public Property Let FocusRectMode(ByVal NewValue As lgFocusRectModeEnum)
    mFocusRectMode = NewValue
    DisplayChange
    
    PropertyChanged "FocusRectMode"
End Property

Public Property Get FocusRectStyle() As lgFocusRectStyleEnum
Attribute FocusRectStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    FocusRectStyle = mFocusRectStyle
End Property

Public Property Let FocusRectStyle(ByVal NewValue As lgFocusRectStyleEnum)
    mFocusRectStyle = NewValue
    DisplayChange
    
    PropertyChanged "FocusRectStyle"
End Property

Public Property Get Font() As Font
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Appearance"
   Set Font = mFont
End Property

Public Property Set Font(ByVal NewValue As StdFont)
    Set mFont = NewValue
    Set UserControl.Font = mFont
    
    PropertyChanged "Font"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ForeColor = mForeColor
End Property

Public Property Let ForeColor(ByVal NewValue As OLE_COLOR)
    mForeColor = NewValue
    
    PropertyChanged "ForeColor"
End Property

Public Property Get ForeColorEdit() As OLE_COLOR
Attribute ForeColorEdit.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ForeColorEdit = mForeColorEdit
End Property

Public Property Let ForeColorEdit(ByVal lNewValue As OLE_COLOR)
    mForeColorEdit = lNewValue
    
    PropertyChanged "ForeColorEdit"
End Property

Public Property Get ForeColorFixed() As OLE_COLOR
Attribute ForeColorFixed.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ForeColorFixed = mForeColorFixed
End Property

Public Property Let ForeColorFixed(ByVal lNewValue As OLE_COLOR)
    mForeColorFixed = lNewValue
    
    PropertyChanged "ForeColorFixed"
End Property

Public Property Get ForeColorSel() As OLE_COLOR
Attribute ForeColorSel.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ForeColorSel = mForeColorSel
End Property

Public Property Let ForeColorSel(ByVal lNewValue As OLE_COLOR)
    mForeColorSel = lNewValue
    DisplayChange
    
    PropertyChanged "ForeColorSel"
End Property

Public Property Get ForeColorTotals() As OLE_COLOR
Attribute ForeColorTotals.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ForeColorTotals = mForeColorTotals
End Property

Public Property Let ForeColorTotals(ByVal NewValue As OLE_COLOR)
    mForeColorTotals = NewValue
    DisplayChange
    
    PropertyChanged "ForeColorTotals"
End Property

Public Property Get FormatString() As String
    FormatString = mFormatString
End Property

Public Property Let FormatString(ByVal NewValue As String)
    '#############################################################################################################################
    'Purpose: Used to create multiple Columns with one string
    
    'Each Column is seperated by a "|" char. The Alignment can be specified by
    'using "^" for Centre, "<" for right an ">" for left (default)
    '#############################################################################################################################
    
    Dim lCol As Long
    Dim sCols() As String
    
    mFormatString = NewValue
    
    If Len(mFormatString) > 0 Then
        sCols() = Split(NewValue, "|")
        If UBound(sCols()) > UBound(mCols) Then
            Cols = UBound(sCols()) + 1
        End If
        
        For lCol = LBound(sCols) To UBound(sCols)
            Select Case Mid$(sCols(lCol), 1, 1)
                Case "^"
                    mCols(mColPtr(lCol)).sCaption = Mid$(sCols(lCol), 2)
                    mCols(mColPtr(lCol)).nAlignment = lgAlignCenterCenter
                Case "<"
                    mCols(mColPtr(lCol)).sCaption = Mid$(sCols(lCol), 2)
                    mCols(mColPtr(lCol)).nAlignment = lgAlignLeftCenter
                Case ">"
                    mCols(mColPtr(lCol)).sCaption = Mid$(sCols(lCol), 2)
                    mCols(mColPtr(lCol)).nAlignment = lgAlignRightCenter
                Case Else
                    mCols(mColPtr(lCol)).sCaption = sCols(lCol)
            End Select
            
            mCols(mColPtr(lCol)).dCustomWidth = 1000
            mCols(mColPtr(lCol)).lWidth = ScaleX(mCols(mColPtr(lCol)).dCustomWidth, mScaleUnits, vbPixels)
            mCols(mColPtr(lCol)).bVisible = True
        Next lCol
    Else
        ReDim mCols(0)
        Clear
    End If
    
    DisplayChange
    
    PropertyChanged "FormatString"
End Property

Public Property Get FullRowSelect() As Boolean
Attribute FullRowSelect.VB_ProcData.VB_Invoke_Property = ";Behavior"
    FullRowSelect = mFullRowSelect
End Property

Public Property Let FullRowSelect(ByVal NewValue As Boolean)
    mFullRowSelect = NewValue
    DisplayChange
    
    PropertyChanged "FullRowSelect"
End Property

Private Function GetColFromX(X As Single) As Long
    '#############################################################################################################################
    'Purpose: Return Column from mouse position
    '#############################################################################################################################
    
    Dim lX As Long
    Dim lCol As Long
    
    GetColFromX = -1
    
    For lCol = SBValue(efsHorizontal) To UBound(mCols)
        With mCols(mColPtr(lCol))
            If .bVisible Then
                If (X > lX) And (X <= lX + .lWidth) Then
                    GetColFromX = lCol
                    Exit For
                End If
                
                lX = lX + .lWidth
            End If
        End With
    Next lCol
End Function

Private Function GetColumnHeadingHeight() As Long
    '#############################################################################################################################
    'Purpose: Return Height of Header Row
    '#############################################################################################################################
    
    Dim lHeight As Long
    
    With UserControl
        lHeight = .TextHeight("A") + 4
        If GetRowHeight() > lHeight Then
            GetColumnHeadingHeight = GetRowHeight()
        Else
            GetColumnHeadingHeight = lHeight
        End If
    End With
End Function

Private Function GetFlag(ByVal nFlags As Integer, nFlag As lgFlagsEnum) As Boolean
    '#############################################################################################################################
    'Purpose: Gets information by bit flags
    '#############################################################################################################################
    
    If nFlags And nFlag Then
        GetFlag = True
    End If
End Function

Private Function GetRowFromY(Y As Single) As Long
    '#############################################################################################################################
    'Purpose: Return Row from mouse position
    '#############################################################################################################################

    Dim lColumnHeadingHeight As Long
    Dim lRow As Long
    Dim lStart As Long
    Dim lY As Long
    
    'Are we below Header?
    If mColumnHeaders Then
        lColumnHeadingHeight = GetColumnHeadingHeight()
        If Y <= lColumnHeadingHeight Then
            GetRowFromY = -1
            Exit Function
        End If
    End If
            
    lY = lColumnHeadingHeight
    lStart = SBValue(efsVertical)
   
    For lRow = lStart To mItemCount
        lY = lY + mItems(mRowPtr(lRow)).lHeight
        
        If lY >= Y Then
            Exit For
        End If
    Next lRow
    
    If lRow <= mItemCount Then
        GetRowFromY = lRow
    Else
        GetRowFromY = -1
    End If
End Function


Private Function GetRowHeight() As Long
    '#############################################################################################################################
    'Purpose: Return Row Height
    '#############################################################################################################################
    
    With UserControl
        If mRowHeightMin > 0 Then
            GetRowHeight = .ScaleY(mRowHeightMin, mScaleUnits, vbPixels)
        Else
            GetRowHeight = ROW_HEIGHT
        End If
    End With
End Function

Public Property Get GridColor() As OLE_COLOR
Attribute GridColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    GridColor = mGridColor
End Property

Public Property Let GridColor(ByVal NewValue As OLE_COLOR)
    mGridColor = NewValue
    DrawGrid mRedraw
        
    PropertyChanged "GridColor"
End Property

Public Property Get ProgressBarColor() As OLE_COLOR
Attribute ProgressBarColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ProgressBarColor = mProgressBarColor
End Property
Private Function ReturnAddr(ByVal sDLL As String, _
                            ByVal sProc As String) As Long
'Return the address of the specified DLL/procedure

    'Get the specified procedure address
    If mWinNT Then
        ReturnAddr = GetProcAddress(GetModuleHandleW(StrPtr(sDLL)), sProc)
    Else
        ReturnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)
    End If
    'In the IDE, validate that the procedure address was located
    Debug.Assert ReturnAddr
  
End Function


Public Property Let ProgressBarColor(ByVal NewValue As OLE_COLOR)
    mProgressBarColor = NewValue
    DrawGrid mRedraw
        
    PropertyChanged "ProgressBarColor"
End Property

Public Property Get GridLines() As Boolean
Attribute GridLines.VB_ProcData.VB_Invoke_Property = ";Appearance"
    GridLines = mGridLines
End Property

Public Property Let GridLines(ByVal NewValue As Boolean)
    mGridLines = NewValue
    DisplayChange
    
    PropertyChanged "GridLines"
End Property

Public Property Let GridLineWidth(NewValue As Long)
    mGridLineWidth = NewValue
    DrawGrid mRedraw
    
    PropertyChanged "GridLineWidth"
End Property

Public Property Get GridLineWidth() As Long
Attribute GridLineWidth.VB_ProcData.VB_Invoke_Property = ";Appearance"
    GridLineWidth = mGridLineWidth
End Property

Public Property Get ForeColorHdr() As OLE_COLOR
    ForeColorHdr = mForeColorHdr
End Property

Public Property Let ForeColorHdr(ByVal NewValue As OLE_COLOR)
    mForeColorHdr = NewValue
    
    PropertyChanged "ForeColorHdr"
End Property

Public Property Get HotHeaderTracking() As Boolean
Attribute HotHeaderTracking.VB_ProcData.VB_Invoke_Property = ";Behavior"
    HotHeaderTracking = mHotHeaderTracking
End Property
Private Function IsValidRowCol(Row As Long, Col As Long) As Boolean
    IsValidRowCol = (Row > NULL_RESULT) And (Col > NULL_RESULT)
End Function
Public Property Let HotHeaderTracking(ByVal NewValue As Boolean)
    mHotHeaderTracking = NewValue
    
    If Not NewValue Then
        DrawHeaderRow
    End If
    
    PropertyChanged "HotHeaderTracking"
End Property

Public Property Get ImageList() As Object
    Set ImageList = mImageList
End Property

Public Property Let ImageList(ByVal NewValue As Object)
    Set mImageList = NewValue
    If Not mImageList Is Nothing Then
        mImageListScaleMode = mImageList.Parent.ScaleMode
    End If
    
    DisplayChange
End Property

Private Function IsEditable() As Boolean
    If Not mLocked And mEditable Then
        IsEditable = (mItemCount >= 0)
    End If
End Function

Private Function IsFunctionExported(ByVal sFunction As String, ByVal sModule As String) As Boolean
    Dim lModule As Long

    If mWinNT Then
        lModule = GetModuleHandleW(StrPtr(sModule))
        If (lModule = 0) Then
            lModule = LoadLibraryW(StrPtr(sModule))
        End If
    Else
        lModule = GetModuleHandleA(sModule)
        If (lModule = 0) Then
            lModule = LoadLibraryA(sModule)
        End If
    End If
    If Not (lModule = 0) Then
        If GetProcAddress(lModule, StrPtr(sFunction)) Then
            IsFunctionExported = True
        End If
        FreeLibrary lModule
    End If
End Function


Public Property Let ItemBackColor(ByVal Index As Long, ByVal NewValue As Long)
    Dim lCol As Long
    
    For lCol = LBound(mCols) To UBound(mCols)
        CellBackColor(Index, lCol) = NewValue
    Next lCol
    
    DrawGrid mRedraw
End Property

Public Property Get ItemChecked(ByVal Index As Long) As Boolean
    ItemChecked = mItems(mRowPtr(Index)).nFlags And lgFLChecked
End Property

Public Property Let ItemChecked(ByVal Index As Long, ByVal NewValue As Boolean)
    SetFlag mItems(mRowPtr(Index)).nFlags, lgFLChecked, NewValue
    DrawGrid mRedraw
End Property

Public Property Get ItemCount() As Long
    ItemCount = mItemCount + 1
End Property

Public Property Get ItemData(ByVal Index As Long) As Long
    ItemData = mItems(mRowPtr(Index)).lItemData
End Property

Public Property Let ItemData(ByVal Index As Long, ByVal NewValue As Long)
    mItems(mRowPtr(Index)).lItemData = NewValue
End Property

Public Property Let ItemFontBold(ByVal Index As Long, ByVal NewValue As Boolean)
    Dim lCol As Long
    
    For lCol = LBound(mCols) To UBound(mCols)
        CellFontBold(Index, lCol) = NewValue
    Next lCol
    
    DrawGrid mRedraw
End Property

Public Property Let ItemForeColor(ByVal Index As Long, ByVal NewValue As Long)
    Dim lCol As Long
    
    For lCol = LBound(mCols) To UBound(mCols)
        CellForeColor(Index, lCol) = NewValue
    Next lCol
    
    DrawGrid mRedraw
End Property

Public Property Let ItemImage(ByVal Index As Long, NewValue As Variant)
    On Local Error GoTo ItemImageError
    
    If IsNumeric(NewValue) Then
        mItems(mRowPtr(Index)).lImage = NewValue
    Else
        mItems(mRowPtr(Index)).lImage = -mImageList.ListImages(NewValue).Index
    End If
    
    DrawGrid mRedraw
    Exit Property
    
ItemImageError:
    mItems(mRowPtr(Index)).lImage = 0
End Property

Public Property Get ItemImage(ByVal Index As Long) As Variant
    If mItems(mRowPtr(Index)).lImage >= 0 Then
        ItemImage = mItems(mRowPtr(Index)).lImage
    Else
        ItemImage = mImageList.ListImages(Abs(mItems(mRowPtr(Index)).lImage)).Key
    End If
End Property

Public Property Get ItemSelected(ByVal Index As Long) As Boolean
    ItemSelected = mItems(mRowPtr(Index)).nFlags And lgFLSelected
End Property

Public Property Let ItemSelected(ByVal Index As Long, ByVal NewValue As Boolean)
    SetFlag mItems(mRowPtr(Index)).nFlags, lgFLSelected, NewValue
    DrawGrid mRedraw
End Property

Public Property Get ItemTag(ByVal Index As Long) As String
    ItemTag = mItems(mRowPtr(Index)).sTag
End Property

Public Property Let ItemTag(ByVal Index As Long, ByVal NewValue As String)
    mItems(mRowPtr(Index)).sTag = NewValue
End Property

Public Function ItemsVisible() As Long
    Dim lBorderWidth As Long
    
    If mBorderStyle = lgSingle Then
        lBorderWidth = 2
    End If

    With UserControl
        ItemsVisible = (.ScaleHeight - GetColumnHeadingHeight() - (lBorderWidth * 2)) / GetRowHeight()
    End With
End Function

Public Property Get MouseCol() As Long
    MouseCol = mMouseCol
End Property

Public Property Get MouseRow() As Long
    MouseRow = mMouseRow
End Property

Private Sub MoveEditControl(ByVal MoveControl As lgMoveControlEnum)
    '#############################################################################################################################
    'Purpose: Used to position and optionally resize the Edit control.
    '#############################################################################################################################
   
    Dim R As RECT
    Dim lBorderWidth As Long
    Dim nScaleMode As ScaleModeConstants
    Dim lHeight As Long
    
    SetColRect mEditCol, R
    
    If Not IsColumnTruncated(mEditCol) Then
        R.Left = R.Left + mGridLineWidth
    End If
    
    On Local Error Resume Next
    
    'Check if an external Control is used.
    If mCols(mColPtr(mEditCol)).EditCtrl Is Nothing Then
        'Using internal TextBox
        With txtEdit
            .Left = R.Left
            .Top = RowTop(mEditRow) + mGridLineWidth
            .Height = mItems(mRowPtr(mEditRow)).lHeight - mGridLineWidth
            .Width = (R.Right - R.Left)
        End With
    Else
        nScaleMode = UserControl.Parent.ScaleMode
        If mBorderStyle = lgSingle Then
            lBorderWidth = 2
        End If
                    
        If (TypeOf mCols(mColPtr(mEditCol)).EditCtrl Is VB.ComboBox) Then
            With mCols(mColPtr(mEditCol)).EditCtrl
                If mCols(mColPtr(mEditCol)).MoveControl And lgBCLeft Then
                    .Left = ScaleX(R.Left + lBorderWidth, vbPixels, nScaleMode) + UserControl.Extender.Left
                End If
                If mCols(mColPtr(mEditCol)).MoveControl And lgBCTop Then
                    .Top = ScaleY(RowTop(mEditRow) + lBorderWidth, vbPixels, nScaleMode) + UserControl.Extender.Top
                End If
                If mCols(mColPtr(mEditCol)).MoveControl And lgBCWidth Then
                    .Width = ScaleX((R.Right - R.Left), vbPixels, nScaleMode)
                End If
                If mCols(mColPtr(mEditCol)).MoveControl And lgBCHeight Then
                    lHeight = mRowHeightMin / Screen.TwipsPerPixelX - mGridLineWidth - 4
                    Call SendMessageAsLong(.hwnd, CB_SETITEMHEIGHT, -1, ByVal lHeight)
                    Call SendMessageAsLong(.hwnd, CB_SETITEMHEIGHT, 0, ByVal lHeight)
                End If
            End With
        Else
            With mCols(mColPtr(mEditCol)).EditCtrl
                If mCols(mColPtr(mEditCol)).MoveControl And lgBCLeft Then
                    .Left = ScaleX(R.Left + lBorderWidth, vbPixels, nScaleMode) + UserControl.Extender.Left
                End If
                If mCols(mColPtr(mEditCol)).MoveControl And lgBCTop Then
                    .Top = ScaleY(RowTop(mEditRow) + mGridLineWidth + lBorderWidth, vbPixels, nScaleMode) + UserControl.Extender.Top
                End If
                If mCols(mColPtr(mEditCol)).MoveControl And lgBCHeight Then
                    .Height = ScaleY(mItems(mRowPtr(mEditRow)).lHeight - mGridLineWidth, vbPixels, nScaleMode)
                End If
                If mCols(mColPtr(mEditCol)).MoveControl And lgBCWidth Then
                    .Width = ScaleX((R.Right - R.Left), vbPixels, nScaleMode)
                End If
            End With
        End If
    End If
    
    On Local Error GoTo 0
End Sub

Public Property Get MultiSelect() As Boolean
Attribute MultiSelect.VB_ProcData.VB_Invoke_Property = ";Behavior"
    MultiSelect = mMultiSelect
End Property

Public Property Let MultiSelect(ByVal NewValue As Boolean)
    mMultiSelect = NewValue
    
    If Not NewValue Then
        SetSelection False
        DisplayChange
    End If
    
    PropertyChanged "MultiSelect"
End Property

Private Function NavigateDown() As Long
    If mRow < mItemCount Then
        NavigateDown = mRow + 1
    Else
        NavigateDown = mRow
    End If
End Function

Private Function NavigateLeft() As Long
    If mCol > 0 Then
        NavigateLeft = mCol - 1
    Else
        NavigateLeft = mCol
    End If
End Function

Private Function NavigateRight() As Long
    If mCol < UBound(mCols) Then
        NavigateRight = mCol + 1
    Else
        NavigateRight = mCol
    End If
End Function

Private Function NavigateUp() As Long
    If mRow > 0 Then
        NavigateUp = mRow - 1
    Else
        NavigateUp = mRow
    End If
End Function

Private Property Get Orientation() As ScrollBarOrienationEnum
    SBOrientation = m_eOrientation
End Property

Private Sub pSBClearUp()
    If m_hWnd <> 0 Then
        On Error Resume Next
        ' Stop flat scroll bar if we have it:
        If Not (m_bNoFlatScrollBars) Then
            UninitializeFlatSB m_hWnd
        End If

        On Error GoTo 0
    End If
    m_hWnd = 0
    m_bInitialised = False
End Sub

Private Sub pSBCreateScrollBar()
    Dim lR As Long
    Dim hParent As Long

    On Error Resume Next
    lR = InitialiseFlatSB(m_hWnd)
    If (Err.Number <> 0) Then
        'Can't find DLL entry point InitializeFlatSB in COMCTL32.DLL
        ' Means we have version prior to 4.71
        ' We get standard scroll bars.
        m_bNoFlatScrollBars = True
    Else
        SBStyle = m_eStyle
    End If
End Sub

Private Sub pSBGetSI(ByVal eBar As EFSScrollBarConstants, ByRef tSI As SCROLLINFO, ByVal fMask As Long)
    Dim Lo As Long

    Lo = eBar
    tSI.fMask = fMask
    tSI.cbSize = LenB(tSI)

    If (m_bNoFlatScrollBars) Then
        GetScrollInfo m_hWnd, Lo, tSI
    Else
        FlatSB_GetScrollInfo m_hWnd, Lo, tSI
    End If

End Sub

Private Sub pSBLetSI(ByVal eBar As EFSScrollBarConstants, ByRef tSI As SCROLLINFO, ByVal fMask As Long)
    Dim Lo As Long

    Lo = eBar
    tSI.fMask = fMask
    tSI.cbSize = LenB(tSI)

    If (m_bNoFlatScrollBars) Then
        SetScrollInfo m_hWnd, Lo, tSI, True
    Else
        FlatSB_SetScrollInfo m_hWnd, Lo, tSI, True
    End If

End Sub

Private Sub pSBSetOrientation()
    ShowScrollBar m_hWnd, SB_HORZ, Abs((m_eOrientation = Scroll_Both) Or (m_eOrientation = Scroll_Horizontal))
    ShowScrollBar m_hWnd, SB_VERT, Abs((m_eOrientation = Scroll_Both) Or (m_eOrientation = Scroll_Vertical))
End Sub

Public Property Get Redraw() As Boolean
Attribute Redraw.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Redraw = mRedraw
End Property

Public Property Let Redraw(ByVal NewValue As Boolean)
    mRedraw = NewValue
    
    If mRedraw Then
        If mPendingScrollBar Then
            SetScrollBars
        End If
        If mPendingRedraw Then
            CreateRenderData
            DrawGrid mRedraw
        End If
    Else
        mPendingScrollBar = False
        mPendingRedraw = False
    End If
    
    PropertyChanged "Redraw"
End Property

Public Sub Refresh()
    CreateRenderData
    SetScrollBars
    DrawGrid True
End Sub

Public Sub RemoveItem(ByVal Index As Long)
    Dim lCount As Long
    Dim lPosition As Long
    Dim bSelected As Boolean
   
    '#############################################################################################################################
    'See AddItem for details of the Arrays used
    '#############################################################################################################################
    
    'Note selected state before deletion
    bSelected = mItems(mRowPtr(Index)).nFlags And lgFLSelected
    
    'Decrement the reference count on each cells format Entry
    If mItemCount >= 0 Then
        For lCount = 0 To UBound(mCols)
            If mItems(Index).Cell(Count).nFormat >= 0 Then
                mCF(mItems(Index).Cell(lCount).nFormat).lRefCount = mCF(mItems(Index).Cell(lCount).nFormat).lRefCount - 1
            End If
        Next lCount
    End If
    
    lPosition = mRowPtr(Index)
    
    'Reset Item Data
    For lCount = mRowPtr(Index) To mItemCount - 1
        mItems(lCount) = mItems(lCount + 1)
    Next lCount
    
    'Adjust Index
    For lCount = Index To mItemCount - 1
        mRowPtr(lCount) = mRowPtr(lCount + 1)
    Next lCount
    
    'Validate Indexes for Items after deleted Item
    For lCount = 0 To mItemCount - 1
        If mRowPtr(lCount) > lPosition Then
            mRowPtr(lCount) = mRowPtr(lCount) - 1
        End If
    Next lCount
    
    mItemCount = mItemCount - 1
     
    If mItemCount < 0 Then
        Clear
    Else
        If (mItemCount + mCacheIncrement) < UBound(mItems) Then
            ReDim Preserve mItems(mItemCount)
            ReDim Preserve mRowPtr(mItemCount)
        End If
 
        If bSelected Then
            If mMultiSelect Then
                RaiseEvent SelectionChanged
            ElseIf Index > mItemCount Then
                SetFlag mItems(mRowPtr(mItemCount)).nFlags, lgFLSelected, True
            ElseIf mItemCount >= 0 Then
                SetFlag mItems(mRowPtr(Index)).nFlags, lgFLSelected, True
            End If
        End If
        
        If Index > mItemCount Then
            SetRowCol mRow - 1, mCol
        End If
    End If
    
    DisplayChange
    
    RaiseEvent ItemCountChanged
End Sub

Public Property Get Row() As Long
    Row = mRow
End Property

Public Property Let Row(ByVal NewValue As Long)
    If SetRowCol(NewValue, mCol) Then
        DrawGrid mRedraw
    End If
End Property

Public Property Get RowHeight(ByVal Index As Long) As Single
    RowHeight = ScaleY(mItems(mRowPtr(Index)).lHeight, vbPixels, mScaleUnits)
End Property

Public Property Let RowHeight(ByVal Index As Long, ByVal NewValue As Single)
    If NewValue = -1 Then
        SetRowSize Row
    Else
        mItems(mRowPtr(Index)).lHeight = ScaleY(NewValue, mScaleUnits, vbPixels)
    End If
    
    SetScrollBars
    DrawGrid mRedraw
End Property

'Public Sub DebugFormatTable()
'    Dim lCount As Long
'    Dim lTotalRef As Long
'
'    For lCount = LBound(mCF) To UBound(mCF)
'        With mCF(lCount)
'            If .lRefCount > 0 Then
'                Debug.Print lCount; .lRefCount; .lBackColor; .lForeColor; .sFontName
'                lTotalRef = lTotalRef + .lRefCount
'            End If
'        End With
'    Next lCount
'
'    Debug.Print " = " & lTotalRef
'End Sub

Public Property Get RowHeightMin() As Long
    RowHeightMin = mRowHeightMin
End Property

Public Property Let RowHeightMin(ByVal NewValue As Long)
    mRowHeightMin = NewValue
    DisplayChange
       
    PropertyChanged "RowHeightMin"
End Property

Public Property Get RowHeightMax() As Long
    RowHeightMax = mRowHeightMax
End Property

Public Property Let RowHeightMax(ByVal NewValue As Long)
    mRowHeightMax = NewValue
    DisplayChange
       
    PropertyChanged "RowHeightMax"
End Property

Public Function RowTop(Index As Long) As Long
    Dim lRow As Long
    Dim lStart As Long
    Dim lY As Long
    
    lStart = SBValue(efsVertical)
    
    If Index >= lStart Then
        lY = GetColumnHeadingHeight()
        
        For lRow = lStart To Index - 1
            lY = lY + mItems(mRowPtr(lRow)).lHeight
        Next lRow
    Else
        lY = NULL_RESULT
    End If
    
    RowTop = lY
End Function

Private Property Get SBCanBeFlat() As Boolean
    SBCanBeFlat = Not (m_bNoFlatScrollBars)
End Property

Private Sub SBCreate(ByVal hWndA As Long)
    pSBClearUp
    m_hWnd = hWndA
    pSBCreateScrollBar
End Sub

Private Property Get SBEnabled(ByVal eBar As EFSScrollBarConstants) As Boolean
    If (eBar = efsHorizontal) Then
        SBEnabled = m_bEnabledHorz
    Else
        SBEnabled = m_bEnabledVert
    End If
End Property

Public Property Get AlphaBlendSelection() As Boolean
    AlphaBlendSelection = mAlphaBlendSelection
End Property

Public Property Let AlphaBlendSelection(ByVal NewValue As Boolean)
    mAlphaBlendSelection = NewValue
    DisplayChange
     
    PropertyChanged "AlphaBlendSelection"
End Property

Private Property Let SBEnabled(ByVal eBar As EFSScrollBarConstants, ByVal bEnabled As Boolean)
    Dim Lo As Long
    Dim lF As Long

    Lo = eBar
    If (bEnabled) Then
        lF = ESB_ENABLE_BOTH
    Else
        lF = ESB_DISABLE_BOTH
    End If
    If (m_bNoFlatScrollBars) Then
        EnableScrollBar m_hWnd, Lo, lF
    Else
        FlatSB_EnableScrollBar m_hWnd, Lo, lF
    End If

End Property

Private Property Get SBLargeChange(ByVal eBar As EFSScrollBarConstants) As Long
    Dim tSI As SCROLLINFO
    pSBGetSI eBar, tSI, SIF_PAGE
    SBLargeChange = tSI.nPage
End Property

Private Property Let SBLargeChange(ByVal eBar As EFSScrollBarConstants, ByVal iLargeChange As Long)
    Dim tSI As SCROLLINFO

    pSBGetSI eBar, tSI, SIF_ALL
    tSI.nMax = tSI.nMax - tSI.nPage + iLargeChange
    tSI.nPage = iLargeChange
    pSBLetSI eBar, tSI, SIF_PAGE Or SIF_RANGE
End Property

Private Property Get SBMax(ByVal eBar As EFSScrollBarConstants) As Long
    Dim tSI As SCROLLINFO
    pSBGetSI eBar, tSI, SIF_RANGE Or SIF_PAGE
    SBMax = tSI.nMax                                  ' - tSI.nPage
End Property

Private Property Let SBMax(ByVal eBar As EFSScrollBarConstants, ByVal iMax As Long)
    Dim tSI As SCROLLINFO
    tSI.nMax = iMax + SBLargeChange(eBar)
    tSI.nMin = SBMin(eBar)
    pSBLetSI eBar, tSI, SIF_RANGE
End Property

Private Property Get SBMin(ByVal eBar As EFSScrollBarConstants) As Long
    Dim tSI As SCROLLINFO
    pSBGetSI eBar, tSI, SIF_RANGE
    SBMin = tSI.nMin
End Property

Private Property Let SBMin(ByVal eBar As EFSScrollBarConstants, ByVal iMin As Long)
    Dim tSI As SCROLLINFO
    tSI.nMin = iMin
    tSI.nMax = SBMax(eBar) + SBLargeChange(eBar)
    pSBLetSI eBar, tSI, SIF_RANGE
End Property

Private Property Let SBOrientation(ByVal eOrientation As ScrollBarOrienationEnum)
    m_eOrientation = eOrientation
    pSBSetOrientation
End Property

Private Sub SBRefresh()
    EnableScrollBar m_hWnd, SB_VERT, ESB_ENABLE_BOTH
End Sub

Private Property Get SBSmallChange(ByVal eBar As EFSScrollBarConstants) As Long
    If (eBar = efsHorizontal) Then
        SBSmallChange = m_lSmallChangeHorz
    Else
        SBSmallChange = m_lSmallChangeVert
    End If
End Property

Private Property Let SBSmallChange(ByVal eBar As EFSScrollBarConstants, ByVal lSmallChange As Long)
    If (eBar = efsHorizontal) Then
        m_lSmallChangeHorz = lSmallChange
    Else
        m_lSmallChangeVert = lSmallChange
    End If
End Property

Private Property Get SBStyle() As ScrollBarStyleEnum
    SBStyle = m_eStyle
End Property

Private Property Let SBStyle(ByVal eStyle As ScrollBarStyleEnum)
    Dim lR As Long
    If (m_bNoFlatScrollBars) Then
        ' can't do it..
        'Debug.Print "Can't set non-regular style mode on this system - COMCTL32.DLL version < 4.71."
        Exit Property
    Else
        If (m_eOrientation = Scroll_Horizontal) Or (m_eOrientation = Scroll_Both) Then
            lR = FlatSB_SetScrollProp(m_hWnd, WSB_PROP_HSTYLE, eStyle, True)
        End If
        If (m_eOrientation = Scroll_Vertical) Or (m_eOrientation = Scroll_Both) Then
            lR = FlatSB_SetScrollProp(m_hWnd, WSB_PROP_VSTYLE, eStyle, True)
        End If
        'Debug.Print lR
        m_eStyle = eStyle
    End If

End Property

Private Property Get SBValue(ByVal eBar As EFSScrollBarConstants) As Long
    Dim tSI As SCROLLINFO
    pSBGetSI eBar, tSI, SIF_POS
    SBValue = tSI.nPos
End Property

Private Property Let SBValue(ByVal eBar As EFSScrollBarConstants, ByVal iValue As Long)
    Dim tSI As SCROLLINFO
    
    If SBVisible(eBar) Then
        If (iValue <> SBValue(eBar)) Then
            tSI.nPos = iValue
            pSBLetSI eBar, tSI, SIF_POS
        End If
    End If
End Property

Private Property Get SBVisible(ByVal eBar As EFSScrollBarConstants) As Boolean
    If (eBar = efsHorizontal) Then
        SBVisible = m_bVisibleHorz
    Else
        SBVisible = m_bVisibleVert
    End If
End Property

Private Property Let SBVisible(ByVal eBar As EFSScrollBarConstants, ByVal bState As Boolean)
    If (eBar = efsHorizontal) Then
        m_bVisibleHorz = bState
    Else
        m_bVisibleVert = bState
    End If
    If (m_bNoFlatScrollBars) Then
        ShowScrollBar m_hWnd, eBar, Abs(bState)
    Else
        FlatSB_ShowScrollBar m_hWnd, eBar, Abs(bState)
    End If
End Property

Public Property Get ScaleUnits() As ScaleModeConstants
    ScaleUnits = mScaleUnits
End Property

Public Property Let ScaleUnits(ByVal NewValue As ScaleModeConstants)
    mScaleUnits = NewValue
    
    PropertyChanged "ScaleUnits"
End Property

Private Sub ScrollList(nDirection As Integer)
    '#############################################################################################################################
    'Purpose: Used to automatically scroll the list up or down when the mouse
    'is dragged out of the Control
    '#############################################################################################################################

    Dim lCount As Long
    Dim lItemsVisible As Long

    mScrollAction = nDirection
      
    Do While mScrollAction = nDirection
        mScrollTick = GetTickCount()
        
        If nDirection = SCROLL_UP Then
            If SBValue(efsVertical) > SBMin(efsVertical) Then
                SBValue(efsVertical) = SBValue(efsVertical) - 1
                If mMultiSelect Then
                    SetFlag mItems(mRowPtr(SBValue(efsVertical))).nFlags, lgFLSelected, True
                Else
                    mRow = SBValue(efsVertical)
                    SetSelection False
                    SetSelection True, mRow, mRow
                End If
                
                RaiseEvent RowColChanged
            Else
                Exit Do
            End If
        Else
            If SBValue(efsVertical) < SBMax(efsVertical) Then
                lItemsVisible = ItemsVisible()
                
                SBValue(efsVertical) = SBValue(efsVertical) + 1
                If mMultiSelect Then
                    For lCount = SBValue(efsVertical) To SBValue(efsVertical) + lItemsVisible
                        If lCount > mItemCount Then
                            Exit For
                        Else
                            SetFlag mItems(mRowPtr(lCount)).nFlags, lgFLSelected, True
                        End If
                    Next lCount
                Else
                    mRow = SBValue(efsVertical) + (lItemsVisible - 1)
                    If mRow > mItemCount Then
                        mRow = mItemCount
                    End If
                    SetSelection False
                    SetSelection True, mRow, mRow
                End If
                
                RaiseEvent RowColChanged
            Else
                Exit Do
            End If
        End If
        
        RaiseEvent SelectionChanged
        DrawGrid mRedraw
        RaiseEvent Scroll
        
        Sleep AUTOSCROLL_TIMEOUT
        DoEvents
    Loop
End Sub

Public Property Get ScrollTrack() As Boolean
Attribute ScrollTrack.VB_ProcData.VB_Invoke_Property = ";Behavior"
    ScrollTrack = mScrollTrack
End Property

Public Property Let ScrollTrack(ByVal NewValue As Boolean)
    mScrollTrack = NewValue
    
    PropertyChanged "ScrollTrack"
End Property

Public Property Get SearchColumn() As Long
    SearchColumn = mSearchColumn
End Property

Public Property Let SearchColumn(ByVal NewValue As Long)
    mSearchColumn = NewValue
    
    PropertyChanged "SearchColumn"
End Property

Public Function SelectedCount() As Long
    '#############################################################################################################################
    'Purpose: Return Count of Selected Items
    '#############################################################################################################################
    
    Dim lCount As Long
    
    For lCount = LBound(mItems) To UBound(mItems)
        If mItems(lCount).nFlags And lgFLSelected Then
            SelectedCount = SelectedCount + 1
        End If
    Next lCount
End Function

Private Function SetColRect(ByVal Index As Long, R As RECT)
    '#############################################################################################################################
    'Purpose: Set the drawing boundary for a Column
    '#############################################################################################################################
    
    Dim lCol As Long
    Dim lCount As Long
    Dim lScrollValue As Long
    Dim lX As Long
    
    lScrollValue = SBValue(efsHorizontal)
    
    If Index < lScrollValue Then
        R.Left = -1
    Else
        For lCol = lScrollValue To Index - 1
            If mCols(mColPtr(lCol)).bVisible Then
                lX = lX + mCols(mColPtr(lCol)).lWidth
                lCount = lCount + 1
            End If
        Next lCol
        
        If IsColumnTruncated(Index) Then
            R.Left = mR.LeftText
            R.Right = R.Left + (mCols(mColPtr(Index)).lWidth - mR.LeftText)
        Else
            R.Left = lX
            R.Right = R.Left + mCols(mColPtr(Index)).lWidth
        End If
    End If
End Function

Private Sub SetItemRect(ByVal Row As Long, ByVal Col As Long, lY As Long, R As RECT, ItemType As lgRectTypeEnum)
    Dim lHeight As Long
    Dim lWidth As Long
    Dim lLeft As Long
    Dim lTop As Long
    Dim nAlignment As lgAlignmentEnum
    
    Select Case ItemType
        Case lgRTColumn
            nAlignment = mCols(Col).nAlignment
            
        Case lgRTCheckBox
            nAlignment = mCols(Col).nAlignment
            lHeight = mR.CheckBoxSize
            lWidth = mR.CheckBoxSize

        Case lgRTImage
            nAlignment = mCols(Col).nImageAlignment
            lHeight = mR.ImageHeight
            lWidth = mR.ImageWidth
    End Select
    
    Select Case nAlignment
        Case lgAlignLeftTop
            lLeft = mCols(Col).lX + 1
            lTop = lY + 2
        Case lgAlignLeftCenter
            lLeft = mCols(Col).lX + 1
            lTop = (lY + (mItems(mRowPtr(Row)).lHeight) / 2) - (lHeight / 2)
        Case lgAlignLeftBottom
            lLeft = mCols(Col).lX + 1
            lTop = (lY + (mItems(mRowPtr(Row)).lHeight)) - (lHeight + 2)
        
        Case lgAlignCenterTop
            lLeft = (mCols(Col).lX + (mCols(Col).lWidth) / 2) - (lWidth / 2)
            lTop = lY + 2
        Case lgAlignCenterCenter
            lLeft = (mCols(Col).lX + (mCols(Col).lWidth) / 2) - (lWidth / 2)
            lTop = (lY + (mItems(mRowPtr(Row)).lHeight) / 2) - (lHeight / 2)
        Case lgAlignCenterBottom
            lLeft = (mCols(Col).lX + (mCols(Col).lWidth) / 2) - (lWidth / 2)
            lTop = (lY + (mItems(mRowPtr(Row)).lHeight)) - (lHeight + 2)
        
        Case lgAlignRightTop
            lLeft = (mCols(Col).lX + mCols(Col).lWidth) - (lWidth + 1)
            lTop = lY + 2
        Case lgAlignRightCenter
            lLeft = (mCols(Col).lX + mCols(Col).lWidth) - (lWidth + 1)
            lTop = (lY + (mItems(mRowPtr(Row)).lHeight) / 2) - (lHeight / 2)
        Case lgAlignRightBottom
            lLeft = (mCols(Col).lX + mCols(Col).lWidth) - (lWidth + 1)
            lTop = (lY + (mItems(mRowPtr(Row)).lHeight)) - (lHeight + 2)
        
    End Select
    
    Call SetRect(R, lLeft, lTop, lLeft + lWidth, lTop + lHeight)
End Sub

Private Sub SetFlag(nFlags As Integer, nFlag As lgFlagsEnum, bValue As Boolean)
    If bValue Then
        nFlags = (nFlags Or nFlag)
    Else
        nFlags = (nFlags And Not (nFlag))
    End If
End Sub

Private Sub SetRedrawState(bState As Boolean)
    '#############################################################################################################################
    'Purpose: Used to prevent Internal Redraws while preserving User Controlled Redraw state
    '
    'bDrawLocked used to prevent nested Calls to Lock Redraw
    '#############################################################################################################################
   
    Static bDrawLocked As Boolean
    Static bOriginalRedraw As Boolean
    
    If bState Then
        bDrawLocked = False
        mRedraw = bOriginalRedraw
    ElseIf Not bDrawLocked Then
        bDrawLocked = True
        bOriginalRedraw = mRedraw
        mRedraw = False
    End If
End Sub


Private Function SetRowCol(lRow As Long, lCol As Long, Optional bSetScroll As Boolean) As Boolean
    '#############################################################################################################################
    'Purpose: To update current Row/Col and fire Events if necessary
    '#############################################################################################################################
    
    Dim R As RECT
    Dim lCount As Long
    
    If (mCol <> lCol) Or (mRow <> lRow) Then
        mCol = lCol
        mRow = lRow
        
        RaiseEvent RowColChanged
        
        'Do we need to change Bars?
        If bSetScroll Then
            SetColRect mCol, R
            
            'Scroll to make Column visible
            If R.Left < 0 Then
                 For lCount = SBValue(efsHorizontal) To SBMin(efsHorizontal) Step -1
                    If R.Left > 0 Then
                        Exit For
                    End If
                    
                    SBValue(efsHorizontal) = SBValue(efsHorizontal) - 1
                    SetColRect mCol, R
                Next lCount
            Else
                For lCount = SBValue(efsHorizontal) To SBMax(efsHorizontal)
                    If R.Left + mCols(mCol).lWidth < UserControl.ScaleWidth Then
                        Exit For
                    End If
                    
                    SBValue(efsHorizontal) = SBValue(efsHorizontal) + 1
                    SetColRect mCol, R
                Next lCount
            End If
            
            If SBValue(efsHorizontal) = SBMin(efsHorizontal) Then
                SetScrollBars
            End If
            
            If mRow < SBValue(efsVertical) Then
                SBValue(efsVertical) = SBValue(efsVertical) - 1
            ElseIf mRow > SBValue(efsVertical) + (ItemsVisible() - 1) Then
                SBValue(efsVertical) = SBValue(efsVertical) + 1
            End If
            
            RaiseEvent Scroll
        End If
        
        SetRowCol = True
    End If
End Function

Private Sub SetScrollBars()
    '#############################################################################################################################
    'Purpose: Sets the visibilty of scroll bars and sets max scroll values
    '#############################################################################################################################

    Dim lCol As Long
    Dim lRow As Long
    Dim lHeight As Long
    Dim lWidth As Long
    Dim lVSB As Long
    Dim bHVisible As Boolean
    Dim bVVisible As Boolean
    
    If m_hWnd <> 0 Then
        '#############################################################################################################################
        'Calculate total width of columns
        For lCol = LBound(mCols) To UBound(mCols)
            If mCols(mColPtr(lCol)).bVisible Then
                lWidth = lWidth + mCols(mColPtr(lCol)).lWidth
            End If
        Next lCol
        
        If (lWidth > UserControl.ScaleWidth) Then
            SBMax(efsHorizontal) = UBound(mCols) - 1
            bHVisible = True
        Else
            SBMax(efsHorizontal) = UBound(mCols)
            bHVisible = (SBValue(efsHorizontal) > SBMin(efsHorizontal))
        End If
        
        '#############################################################################################################################
        'Calculate total height of rows
        lHeight = GetColumnHeadingHeight()
        For lRow = LBound(mItems) To mItemCount
            lHeight = lHeight + mItems(mRowPtr(lRow)).lHeight
        Next lRow
   
        If lHeight > UserControl.ScaleHeight Then
            'Adjust scrollbar to best-fit Rows to Grid
            lHeight = GetColumnHeadingHeight()
            For lRow = mItemCount To LBound(mItems) Step -1
                lHeight = lHeight + mItems(mRowPtr(lRow)).lHeight
                
                If lHeight > UserControl.ScaleHeight Then
                    Exit For
                End If
                
                lVSB = lVSB + 1
            Next lRow
        
            SBMax(efsVertical) = mItemCount - lVSB
            bVVisible = True
        Else
            SBMax(efsVertical) = mItemCount
        End If
        
        '#############################################################################################################################
        'If SBVisible(efsHorizontal) <> bHVisible Then
            SBVisible(efsHorizontal) = bHVisible
        'End If
        'If SBVisible(efsVertical) <> bVVisible Then
            SBVisible(efsVertical) = bVVisible
        'End If
    End If
End Sub

Private Function SetSelection(bState As Boolean, Optional lFromRow As Long = -1, Optional lToRow As Long = -1) As Boolean
    Dim lCount As Long
    Dim lStep As Long
    Dim bSelectionChanged As Boolean
    
    If lFromRow = -1 Then
        lFromRow = LBound(mItems)
    End If
    
    If lToRow = -1 Then
        lToRow = UBound(mItems)
    End If
    
    If lFromRow >= lToRow Then
        lStep = -1
    Else
        lStep = 1
    End If
    
    For lCount = lFromRow To lToRow Step lStep
        If (mItems(mRowPtr(lCount)).nFlags And lgFLSelected) <> bState Then
            SetFlag mItems(mRowPtr(lCount)).nFlags, lgFLSelected, bState
            bSelectionChanged = True
        End If
    Next lCount
    
    SetSelection = bSelectionChanged
End Function

Private Sub SortArrayString(ByVal lFirst As Long, ByVal lLast As Long, lSortColumn As Long, ByVal nSortType As Integer)
    '#############################################################################################################################
    'Purpose: A simple data-type aware quick-sort method to Sort Grid Rows
    '#############################################################################################################################

    Dim lBoundary As Long
    Dim lIndex As Long
    Dim bSwap As Boolean
    
    If lLast <= lFirst Then Exit Sub

    SwapLng mRowPtr(lFirst), mRowPtr((lFirst + lLast) / 2)
    
    lBoundary = lFirst

    For lIndex = lFirst + 1 To lLast
        bSwap = False
        If nSortType = 0 Then
            bSwap = mItems(mRowPtr(lIndex)).Cell(lSortColumn).sValue > mItems(mRowPtr(lFirst)).Cell(lSortColumn).sValue
        Else
            bSwap = mItems(mRowPtr(lIndex)).Cell(lSortColumn).sValue < mItems(mRowPtr(lFirst)).Cell(lSortColumn).sValue
        End If
        
        If bSwap Then
            lBoundary = lBoundary + 1
            SwapLng mRowPtr(lBoundary), mRowPtr(lIndex)
        End If
    Next lIndex

    SwapLng mRowPtr(lFirst), mRowPtr(lBoundary)
    SortArrayString lFirst, lBoundary - 1, lSortColumn, nSortType
    SortArrayString lBoundary + 1, lLast, lSortColumn, nSortType
End Sub

Private Sub SortArrayDate(ByVal lFirst As Long, ByVal lLast As Long, lSortColumn As Long, ByVal nSortType As Integer)
    '#############################################################################################################################
    'Purpose: A simple data-type aware quick-sort method to Sort Grid Rows
    '#############################################################################################################################

    Dim lBoundary As Long
    Dim lIndex As Long
    Dim bIsDate(1) As Boolean
    Dim bSwap As Boolean
    
    If lLast <= lFirst Then Exit Sub

    SwapLng mRowPtr(lFirst), mRowPtr((lFirst + lLast) / 2)
    
    lBoundary = lFirst

    For lIndex = lFirst + 1 To lLast
        bIsDate(0) = IsDate(mItems(mRowPtr(lIndex)).Cell(lSortColumn).sValue)
        bIsDate(1) = IsDate(mItems(mRowPtr(lFirst)).Cell(lSortColumn).sValue)
        
        If nSortType = 0 Then
            If Not bIsDate(0) Then
                bSwap = False
            ElseIf Not bIsDate(1) Then
                bSwap = True
            Else
                bSwap = CDate(mItems(mRowPtr(lIndex)).Cell(lSortColumn).sValue) > CDate(mItems(mRowPtr(lFirst)).Cell(lSortColumn).sValue)
            End If
        Else
            If Not bIsDate(0) Then
                bSwap = True
            ElseIf Not bIsDate(1) Then
                bSwap = False
            Else
                bSwap = CDate(mItems(mRowPtr(lIndex)).Cell(lSortColumn).sValue) < CDate(mItems(mRowPtr(lFirst)).Cell(lSortColumn).sValue)
            End If
        End If
        
        If bSwap Then
            lBoundary = lBoundary + 1
            SwapLng mRowPtr(lBoundary), mRowPtr(lIndex)
        End If
    Next lIndex

    SwapLng mRowPtr(lFirst), mRowPtr(lBoundary)
    SortArrayDate lFirst, lBoundary - 1, lSortColumn, nSortType
    SortArrayDate lBoundary + 1, lLast, lSortColumn, nSortType
End Sub


Private Sub SortArrayNumeric(ByVal lFirst As Long, ByVal lLast As Long, lSortColumn As Long, ByVal nSortType As Integer)
    '#############################################################################################################################
    'Purpose: A simple data-type aware quick-sort method to Sort Grid Rows
    '#############################################################################################################################

    Dim lBoundary As Long
    Dim lIndex As Long
    Dim bSwap As Boolean
    
    If lLast <= lFirst Then Exit Sub

    SwapLng mRowPtr(lFirst), mRowPtr((lFirst + lLast) / 2)
    
    lBoundary = lFirst

    For lIndex = lFirst + 1 To lLast
        bSwap = False
        If nSortType = 0 Then
            bSwap = Val(mItems(mRowPtr(lIndex)).Cell(lSortColumn).sValue) > Val(mItems(mRowPtr(lFirst)).Cell(lSortColumn).sValue)
        Else
            bSwap = Val(mItems(mRowPtr(lIndex)).Cell(lSortColumn).sValue) < Val(mItems(mRowPtr(lFirst)).Cell(lSortColumn).sValue)
        End If
        
        If bSwap Then
            lBoundary = lBoundary + 1
            SwapLng mRowPtr(lBoundary), mRowPtr(lIndex)
        End If
    Next lIndex

    SwapLng mRowPtr(lFirst), mRowPtr(lBoundary)
    SortArrayNumeric lFirst, lBoundary - 1, lSortColumn, nSortType
    SortArrayNumeric lBoundary + 1, lLast, lSortColumn, nSortType
End Sub


Private Sub SortArrayCustom(ByVal lFirst As Long, ByVal lLast As Long, lSortColumn As Long, ByVal nSortType As Integer)
    '#############################################################################################################################
    'Purpose: A simple data-type aware quick-sort method to Sort Grid Rows
    '#############################################################################################################################

    Dim lBoundary As Long
    Dim lIndex As Long
    Dim bSwap As Boolean
    
    If lLast <= lFirst Then Exit Sub

    SwapLng mRowPtr(lFirst), mRowPtr((lFirst + lLast) / 2)
    
    lBoundary = lFirst

    For lIndex = lFirst + 1 To lLast
        bSwap = False
        If nSortType = 0 Then
            RaiseEvent CustomSort(True, lSortColumn, mItems(mRowPtr(lIndex)).Cell(lSortColumn).sValue, mItems(mRowPtr(lFirst)).Cell(lSortColumn).sValue, bSwap)
        Else
            RaiseEvent CustomSort(False, lSortColumn, mItems(mRowPtr(lIndex)).Cell(lSortColumn).sValue, mItems(mRowPtr(lFirst)).Cell(lSortColumn).sValue, bSwap)
        End If
        
        If bSwap Then
            lBoundary = lBoundary + 1
            SwapLng mRowPtr(lBoundary), mRowPtr(lIndex)
        End If
    Next lIndex

    SwapLng mRowPtr(lFirst), mRowPtr(lBoundary)
    SortArrayCustom lFirst, lBoundary - 1, lSortColumn, nSortType
    SortArrayCustom lBoundary + 1, lLast, lSortColumn, nSortType
End Sub


Private Sub SortArrayBool(ByVal lFirst As Long, ByVal lLast As Long, lSortColumn As Long, ByVal nSortType As Integer)
    '#############################################################################################################################
    'Purpose: A simple data-type aware quick-sort method to Sort Grid Rows
    '#############################################################################################################################

    Dim lBoundary As Long
    Dim lIndex As Long
    Dim bSwap As Boolean
    
    If lLast <= lFirst Then Exit Sub

    SwapLng mRowPtr(lFirst), mRowPtr((lFirst + lLast) / 2)
    
    lBoundary = lFirst

    For lIndex = lFirst + 1 To lLast
        bSwap = False
        If nSortType = 0 Then
            bSwap = GetFlag(mItems(mRowPtr(lIndex)).Cell(lSortColumn).nFlags, lgFLChecked) > GetFlag(mItems(mRowPtr(lFirst)).Cell(lSortColumn).nFlags, lgFLChecked)
        Else
            bSwap = GetFlag(mItems(mRowPtr(lIndex)).Cell(lSortColumn).nFlags, lgFLChecked) < GetFlag(mItems(mRowPtr(lFirst)).Cell(lSortColumn).nFlags, lgFLChecked)
        End If
        
        If bSwap Then
            lBoundary = lBoundary + 1
            SwapLng mRowPtr(lBoundary), mRowPtr(lIndex)
        End If
    Next lIndex

    SwapLng mRowPtr(lFirst), mRowPtr(lBoundary)
    SortArrayBool lFirst, lBoundary - 1, lSortColumn, nSortType
    SortArrayBool lBoundary + 1, lLast, lSortColumn, nSortType
End Sub


Private Sub SortArray(ByVal lFirst As Long, ByVal lLast As Long, lSortColumn As Long, ByVal nSortType As Integer)
    '#############################################################################################################################
    'Purpose: A simple data-type aware quick-sort method to Sort Grid Rows
    '#############################################################################################################################
    
    Select Case mCols(lSortColumn).nType
        Case lgBoolean
            SortArrayBool lFirst, lLast, lSortColumn, nSortType
        Case lgDate
            SortArrayDate lFirst, lLast, lSortColumn, nSortType
        Case lgNumeric
            SortArrayNumeric lFirst, lLast, lSortColumn, nSortType
        Case lgCustom
            SortArrayCustom lFirst, lLast, lSortColumn, nSortType
        Case Else
            SortArrayString lFirst, lLast, lSortColumn, nSortType
    End Select
End Sub
Private Sub SortSubList()
    '#############################################################################################################################
    'Purpose: Used to sort by a secondary Column after a Sort
    '#############################################################################################################################
    
    Dim lCount As Long
    Dim lStartSort As Long
    Dim bDifferent As Boolean
    Dim sMajorSort As String

    If mSortSubColumn > NULL_RESULT Then
        'Re-Sort the Items by a secondary column, preserving the sort sequence of the
        'primary sort
        
        lStartSort = LBound(mItems)
        For lCount = LBound(mItems) To mItemCount
            bDifferent = mItems(mRowPtr(lCount)).Cell(mSortColumn).sValue <> sMajorSort
            If bDifferent Or lCount = mItemCount Then
                If lCount > 1 Then
                    If lCount - lStartSort > 1 Then
                        If lCount = mItemCount And Not bDifferent Then
                            SortArray lStartSort, lCount, mSortSubColumn, mCols(mSortSubColumn).nSortOrder
                        Else
                            SortArray lStartSort, lCount - 1, mSortSubColumn, mCols(mSortSubColumn).nSortOrder
                        End If
                    End If
                    lStartSort = lCount
                End If
                
                sMajorSort = mItems(mRowPtr(lCount)).Cell(mSortColumn).sValue
            End If
        Next lCount
    End If
End Sub

Private Sub SwapLng(Value1 As Long, Value2 As Long)
    Static lTemp As Long

    lTemp = Value1
    Value1 = Value2
    Value2 = lTemp
End Sub

Private Function ToggleEdit() As Boolean
    '#############################################################################################################################
    'Purpose: Used to start a new Edit or commit a pending one
    '#############################################################################################################################
    
    If IsEditable() Then
        ToggleEdit = True
        
        If mEditPending Then
            UpdateCell
        ElseIf (mRow <> NULL_RESULT) And (mCol <> NULL_RESULT) Then
            EditCell mRow, mCol
        End If
    End If
End Function

Public Property Let TopRow(ByVal NewValue As Long)
    If NewValue > SBMax(efsVertical) Then
        SBValue(efsVertical) = SBMax(efsVertical)
    Else
        SBValue(efsVertical) = NewValue
    End If
    
    SetRowCol NewValue, mCol, True
    DrawGrid mRedraw
End Property

Private Sub TrackMouseLeave(ByVal lng_hWnd As Long)
    Dim tme As TRACKMOUSEEVENT_STRUCT
    
    If bTrack Then
        With tme
            .cbSize = Len(tme)
            .dwFlags = TME_LEAVE
            .hwndTrack = lng_hWnd
        End With
        
        If bTrackUser32 Then
            Call TrackMouseEvent(tme)
        Else
            Call TrackMouseEventComCtl(tme)
        End If
    End If
End Sub

Public Property Let ThemeColor(NewValue As lgThemeColorEnum)
    mThemeColor = NewValue
    SetColors
    DrawGrid True
    
    PropertyChanged "ThemeColor"
End Property

Public Property Get ThemeColor() As lgThemeColorEnum
    ThemeColor = mThemeColor
End Property

Private Sub SetColors()
    Select Case mThemeColor
        Case lgTCDefault
            mBackColor = DEF_BACKCOLOR
            mForeColor = DEF_FORECOLOR
            mBackColorSel = DEF_BACKCOLORSEL
            mForeColorSel = DEF_FORECOLORSEL
            
            mFocusRectColor = DEF_FOCUSRECTCOLOR
            mGridColor = DEF_GRIDCOLOR
            
        Case lgTCBlue
            mBackColor = DEF_BACKCOLOR
            mForeColor = DEF_FORECOLOR
            mBackColorSel = &HF1D8C9
            mForeColorSel = &H9C613B
            
            mFocusRectColor = &H9C613B
            mGridColor = &HEBEBEB
            
        Case lgTCGreen
            mBackColor = DEF_BACKCOLOR
            mForeColor = DEF_FORECOLOR
            mBackColorSel = &H8FC5B5
            mForeColorSel = &HE1F9F7
            
            mFocusRectColor = &H385D3F
            mGridColor = &HC0FFC0
           
    End Select
End Sub

Public Property Let ThemeStyle(NewValue As lgThemeStyleEnum)
    mThemeStyle = NewValue
    DrawGrid True
    
    PropertyChanged "ThemeStyle"
End Property

Public Property Get ThemeStyle() As lgThemeStyleEnum
    ThemeStyle = mThemeStyle
End Property


Private Function TranslateColor(ByVal clrColor As OLE_COLOR, Optional hPalette As Long = 0) As Long
    If OleTranslateColor(clrColor, hPalette, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function

Private Function UpdateCell() As Boolean
    '#############################################################################################################################
    'Purpose: Used to commit an Edit. Note the RequestUpate event. This event allows
    'the Upate to be cancelled by setting the Cancel flag.
    '#############################################################################################################################
   
    Dim bCancel As Boolean
    Dim bRequestUpdate As Boolean
    Dim sNewValue As String
    
    If mEditPending Then
        If mCols(mColPtr(mEditCol)).EditCtrl Is Nothing Then
            bRequestUpdate = (mItems(mRowPtr(mEditRow)).Cell(mColPtr(mEditCol)).sValue <> txtEdit.Text)
            sNewValue = txtEdit.Text
        Else
            bRequestUpdate = True
        End If
        
        If bRequestUpdate Then
            RaiseEvent RequestUpdate(mEditRow, mEditCol, sNewValue, bCancel)
        End If
        
        If Not bCancel Then
            SetRedrawState False
        
            If mCols(mColPtr(mEditCol)).EditCtrl Is Nothing Then
                txtEdit.Visible = False
            Else
                On Local Error Resume Next
                
                With mCols(mColPtr(mEditCol)).EditCtrl
                    If mEditParent <> 0 Then
                        SetParent .hwnd, mEditParent
                    End If
                    
                    Subclass_Stop .hwnd
                    
                    .Visible = False
                End With
                
                On Local Error GoTo 0
            End If
            
            mEditPending = False
            
            If bRequestUpdate Then
                mItems(mRowPtr(mEditRow)).Cell(mColPtr(mEditCol)).sValue = sNewValue
                SetRowSize mEditRow
                
                SetFlag mItems(mRowPtr(mEditRow)).Cell(mColPtr(mEditCol)).nFlags, lgFLChanged, True
                
                DisplayChange
            End If
            
            SetRedrawState True
            DrawGrid True
        End If
    End If
    
    UpdateCell = Not bCancel
End Function

Private Sub SetRowSize(ByVal Row As Long)
    Dim R As RECT
    Dim lCol As Long
    Dim lHeight As Long
    Dim sText As String
    
    If mAutoSizeRow Then
        For lCol = LBound(mCols) To UBound(mCols)
            sText = mItems(mRowPtr(Row)).Cell(lCol).sValue
            
            If mItems(mRowPtr(Row)).Cell(lCol).nFlags And lgFLWordWrap Then
                SetRect R, 0, 2, mCols(lCol).lWidth, 0
                DrawText UserControl.hdc, sText, Len(sText), R, DT_CALCRECT Or DT_WORDBREAK
            Else
                SetRect R, 0, 0, mCols(lCol).lWidth, 0
                DrawText UserControl.hdc, sText, Len(sText), R, DT_CALCRECT
            End If
            
            If R.Bottom > lHeight Then
                lHeight = R.Bottom
            End If
        Next lCol
        
        If lHeight < ScaleY(mRowHeightMin, mScaleUnits, vbPixels) Then
            mItems(mRowPtr(Row)).lHeight = ScaleY(mRowHeightMin, mScaleUnits, vbPixels)
        ElseIf (mRowHeightMax > 0) And (lHeight > ScaleY(mRowHeightMax, mScaleUnits, vbPixels)) Then
            mItems(mRowPtr(Row)).lHeight = ScaleY(mRowHeightMax, mScaleUnits, vbPixels)
        Else
            mItems(mRowPtr(Row)).lHeight = lHeight
        End If
    End If
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    Select Case mCols(mColPtr(mEditCol)).sInputFilter
        Case vbNullString
            'No Filter
        Case "<", ">"
            'lowercase / UPPERCASE
        Case Else
            Select Case KeyAscii
                Case vbKeyBack, vbKeyDelete
                    'Do not restrict!
                
                Case Else
                    If InStr(mCols(mColPtr(mEditCol)).sInputFilter, Chr$(KeyAscii)) = 0 Then
                        KeyAscii = 0
                    End If
            End Select
    End Select

    RaiseEvent EditKeyPress(mEditCol, KeyAscii)
End Sub


Private Sub UserControl_Click()
    If (mEditTrigger And lgMouseClick) And (mMouseRow > NULL_RESULT) Then
        ToggleEdit
    End If
    
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    If (mEditTrigger And lgMouseDblClick) And (mMouseRow > NULL_RESULT) Then
        ToggleEdit
    End If
    
    RaiseEvent DblClick
End Sub

Private Sub UserControl_Initialize()
    Dim OS As OSVERSIONINFO
    
    mClipRgn = CreateRectRgn(0, 0, 0, 0)
    
    mLockFocusDraw = IsWindowUnicode(UserControl.hwnd)
    
    OS.dwOSVersionInfoSize = Len(OS)
    Call GetVersionEx(OS)
    
    mWinNT = ((OS.dwPlatformId And VER_PLATFORM_WIN32_NT) = VER_PLATFORM_WIN32_NT)
    
    If (OS.dwMajorVersion > 5) Then
        mWinXP = True
    ElseIf (OS.dwMajorVersion = 5) And (OS.dwMinorVersion >= 1) Then
        mWinXP = True
    End If
    
    Set txtEdit = UserControl.Controls.Add("VB.TextBox", "txtEdit")
    With txtEdit
        .BorderStyle = 0
        .Visible = False
        
        If mWinNT Then
            mTextBoxStyle = GetWindowLongW(.hwnd, GWL_STYLE)
        Else
            mTextBoxStyle = GetWindowLongA(.hwnd, GWL_STYLE)
        End If
    End With

    ReDim mCols(0)
    ReDim mColPtr(0)
    Clear
End Sub

Private Sub UserControl_InitProperties()
    Set mFont = Ambient.Font

    '################################################################################
    'Appearance Properties
    mApplySelectionToImages = DEF_APPLYSELECTIONTOIMAGES
    mBackColor = DEF_BACKCOLOR
    mBackColorBkg = DEF_BACKCOLORBKG
    mBackColorEdit = DEF_BACKCOLOREDIT
    mBackColorFixed = DEF_BACKCOLORFIXED
    mBackColorSel = DEF_BACKCOLORSEL
    mForeColor = DEF_FORECOLOR
    mForeColorEdit = DEF_FORECOLOREDIT
    mForeColorFixed = DEF_FORECOLORFIXED
    mForeColorHdr = DEF_FORECOLORHDR
    mForeColorSel = DEF_FORECOLORSEL
    mForeColorTotals = DEF_FORECOLORTOTALS
    
    mFocusRectColor = DEF_FOCUSRECTCOLOR
    mGridColor = DEF_GRIDCOLOR
    mProgressBarColor = DEF_PROGRESSBARCOLOR
    
    mAlphaBlendSelection = DEF_ALPHABLENDSELECTION
    mDisplayEllipsis = DEF_DISPLAYELLIPSIS
    mFocusRectMode = DEF_FOCUSRECTMODE
    mFocusRectStyle = DEF_FOCUSRECTSTYLE
    mGridLines = DEF_GRIDLINES
    mGridLineWidth = DEF_GRIDLINEWIDTH
    mThemeColor = DEF_THEMECOLOR
    mThemeStyle = DEF_THEMESTYLE
    
    '################################################################################
    'Behaviour Properties
    mAllowUserResizing = DEF_ALLOWUSERRESIZING
    mAutoSizeRow = DEF_AUTOSIZEROW
    mBorderStyle = DEF_BORDERSTYLE
    mCheckboxes = DEF_CHECKBOXES
    mColumnDrag = DEF_COLUMNDRAG
    mColumnHeaders = DEF_COLUMNHEADERS
    mColumnSort = DEF_COLUMNSORT
    mEditable = DEF_EDITABLE
    mEditTrigger = DEF_EDITTRIGGER
    mFullRowSelect = DEF_FULLROWSELECT
    mHideSelection = DEF_HIDESELECTION
    mHotHeaderTracking = DEF_HOTHEADERTRACKING
    mMultiSelect = DEF_MULTISELECT
    mRedraw = DEF_REDRAW
    mScrollTrack = DEF_SCROLLTRACK
    mTrackEdits = DEF_TRACKEDITS
    
    '################################################################################
    'Miscellaneous Properties
    mCacheIncrement = DEF_CACHEINCREMENT
    mEnabled = DEF_ENABLED
    mFormatString = DEF_FORMATSTRING
    mLocked = DEF_LOCKED
    mRowHeightMax = DEF_ROWHEIGHTMAX
    mRowHeightMin = DEF_ROWHEIGHTMIN
    mScaleUnits = DEF_SCALEUNITS
    mSearchColumn = DEF_SEARCHCOLUMN
    
    '################################################################################
    'Apply Settings
    With UserControl
        .BackColor = mBackColorBkg
        .BorderStyle = mBorderStyle
    End With
    
    CreateRenderData
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lNewCol As Long
    Dim lNewRow As Long
    Dim bClearSelection As Boolean
    Dim bRedraw As Boolean

    lNewCol = mCol
    lNewRow = mRow
    
    SetRedrawState False
    
    'Used to determine if selected Items need to be cleared
    bClearSelection = True

    Select Case KeyCode
        Case vbKeyReturn, vbKeyEscape 'Allow escape to abort editing
           bClearSelection = False

           If (mEditTrigger And lgEnterKey) Then
               If KeyCode = vbKeyEscape Then
                   txtEdit.Visible = False
                   mEditPending = False
               Else
                   If ToggleEdit() Then KeyCode = 0
               End If
           End If
            
        Case vbKeyF2
            bClearSelection = False
            
            If (mEditTrigger And lgF2Key) Then
                If ToggleEdit() Then
                    KeyCode = 0
                End If
            End If
            
        Case vbKeySpace
            bClearSelection = False
        
            If mCheckboxes Then
                mIgnoreKeyPress = True
                
                bRedraw = True
                SetFlag mItems(mRowPtr(mRow)).nFlags, lgFLChecked, Not GetFlag(mItems(mRowPtr(mRow)).nFlags, lgFLChecked)
                RaiseEvent ItemChecked(mRow)
                
                KeyCode = 0
            End If
            
        Case vbKeyA
            bClearSelection = False
            
            If (Shift And vbCtrlMask) And mMultiSelect Then
                mIgnoreKeyPress = True
                
                SetSelection True
                RaiseEvent SelectionChanged
                KeyCode = 0
            End If
    
        Case vbKeyUp
            If (Shift And vbShiftMask) And mMultiSelect Then
                bClearSelection = False
            End If
            
            If UpdateCell() Then
                lNewRow = NavigateUp()
                
                KeyCode = 0
            End If
            
        Case vbKeyDown
            If (Shift And vbShiftMask) And mMultiSelect Then
                bClearSelection = False
            End If
        
            If UpdateCell() Then
                lNewRow = NavigateDown()
                
                KeyCode = 0
            End If
            
        Case vbKeyLeft
            If Not mEditPending Then
                lNewCol = NavigateLeft()
                KeyCode = 0
            End If
            
        Case vbKeyRight
            If Not mEditPending Then
                lNewCol = NavigateRight()
                KeyCode = 0
            End If
            
        Case vbKeyPageUp
            If UpdateCell() Then
                If mRow > 0 Then
                    lNewRow = (mRow - ItemsVisible()) + 1
                    If lNewRow < 0 Then
                        lNewRow = 0
                    End If
                    
                    SBValue(efsVertical) = lNewRow
                End If
                
                KeyCode = 0
            End If
        
        Case vbKeyPageDown
            If UpdateCell() Then
                If mRow < mItemCount Then
                    lNewRow = (mRow + ItemsVisible()) - 1
                    If lNewRow > mItemCount Then
                        lNewRow = mItemCount
                    End If
         
                    SBValue(efsVertical) = lNewRow
                End If
                
                KeyCode = 0
            End If
        
        Case vbKeyHome
            If Shift And vbShiftMask Then
                If UpdateCell() Then
                    If mMultiSelect Then
                        bClearSelection = False
          
                        SetSelection False
                        SetSelection True, 1, mRow
                        RaiseEvent SelectionChanged
                    End If
                    
                    lNewRow = 0
                    
                    SBValue(efsVertical) = SBMin(efsVertical)
                    KeyCode = 0
                End If
            ElseIf Shift And vbCtrlMask Then
                If UpdateCell() Then
                    lNewRow = 0
                    
                    SBValue(efsVertical) = SBMin(efsVertical)
                    KeyCode = 0
                End If
            ElseIf Not mEditPending Then
                lNewCol = 0
                
                SBValue(efsHorizontal) = SBMin(efsHorizontal)
                KeyCode = 0
            End If
            
        Case vbKeyEnd
            If Shift And vbShiftMask Then
                If UpdateCell() Then
                    If mMultiSelect Then
                        bClearSelection = False
          
                        SetSelection False
                        SetSelection True, mRow, mItemCount
                        RaiseEvent SelectionChanged
                    End If
                    
                    lNewRow = mItemCount
                    
                    SBValue(efsVertical) = SBMax(efsVertical)
                    KeyCode = 0
                End If
            ElseIf Shift And vbCtrlMask Then
                If UpdateCell() Then
                    lNewRow = mItemCount
                    
                    SBValue(efsVertical) = SBMax(efsVertical)
                    KeyCode = 0
                End If
            ElseIf Not mEditPending Then
                lNewCol = UBound(mCols)
                
                SBValue(efsHorizontal) = SBMax(efsHorizontal)
                KeyCode = 0
            End If
            
        Case Else
            If Not mEditPending Then
                If (mEditTrigger And lgAnyKey) Then
                    bClearSelection = False
                
                    If ToggleEdit() Then
                        KeyCode = 0
                    End If
                End If
            End If
               
    End Select
    
    SetRedrawState True
    
    If KeyCode = 0 Then
        'Do we want to clear selection?
        If bClearSelection And (mRow <> lNewRow) Then
            bRedraw = SetSelection(False)
        End If
        
        If lNewRow > NULL_RESULT Then
            If Not mItems(mRowPtr(lNewRow)).nFlags And lgFLSelected Then
                bRedraw = True
                SetFlag mItems(mRowPtr(lNewRow)).nFlags, lgFLSelected, True
                RaiseEvent SelectionChanged
            End If
        End If
        
        If bRedraw Or SetRowCol(lNewRow, lNewCol, True) Then
            DrawGrid mRedraw
        End If
    Else
        RaiseEvent KeyDown(KeyCode, Shift)
    End If
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    '#############################################################################################################################
    'Purpose: This will find the Item that contains a Cell with text that is >= to the text typed. Each
    'character entered is appended to the previous one if the time interval is less than 1 second.
    
    'Key searching is disabled if the Grid is Disabled, and Edit is in progress or the KeyPress event is
    'in an Ignore State (setting the SearchColumn to -1 will also prevent searches).
    
    Static lTime As Long
    Static sCode As String
    
    Dim lResult As Long
    Dim bEatKey As Boolean
   
    If mEnabled Then
        'Used to prevent a beep
        If (mEditTrigger And lgEnterKey) And (KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape) Then
            KeyAscii = 0
            bEatKey = True
        ElseIf Not mIgnoreKeyPress And Not mEditPending Then
            If IsCharAlphaNumeric(KeyAscii) Then
                If (GetTickCount() - lTime) < 1000 Then
                    sCode = sCode & Chr$(KeyAscii)
                Else
                    sCode = Chr$(KeyAscii)
                End If
                
                lTime = GetTickCount()
                
                lResult = FindItem(sCode, mSearchColumn, lgSMNavigate)
                If lResult > NULL_RESULT Then
                    TopRow = lResult
                End If
            End If
        End If
        
        If Not bEatKey Then RaiseEvent KeyPress(KeyAscii)
    End If
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    mIgnoreKeyPress = False
    
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim R As RECT
    Dim lMCol As Long
    Dim bCancel As Boolean
    Dim bProcessed As Boolean
    Dim bRedraw As Boolean
    Dim bSelectionChanged As Boolean
    Dim bState As Boolean
    
    If Not mLocked And (Button <> 0) And (mItemCount >= 0) Then
        mScrollAction = SCROLL_NONE
            
        lMCol = GetColFromX(X)
        
        mMouseDownRow = GetRowFromY(Y)
        mMouseDownX = X
            
        If Button = vbLeftButton Then
            Call SetCapture(UserControl.hwnd)
            mMouseDown = True
            
            If Y < mR.HeaderHeight Then
                If (UserControl.MousePointer <> vbSizeWE) Then
                    mMouseDownCol = lMCol
                    If mMouseDownCol <> NULL_RESULT Then
                        With UserControl
                            DrawHeader mMouseCol, lgDown
                            .Refresh
                        End With
                    End If
                End If
            ElseIf mMouseDownRow > NULL_RESULT Then
                If UpdateCell() Then
                    If mCheckboxes And (X <= RIGHT_CHECKBOX) Then
                        bRedraw = True
                        mMouseDown = False
                        
                        SetFlag mItems(mRowPtr(mMouseDownRow)).nFlags, lgFLChecked, Not GetFlag(mItems(mRowPtr(mMouseDownRow)).nFlags, lgFLChecked)
                        
                        RaiseEvent ItemChecked(mMouseDownRow)
                    Else
                        If lMCol > NULL_RESULT Then
                            If IsEditable() And mCols(mColPtr(lMCol)).nType = lgBoolean Then
                                SetItemRect mRowPtr(mMouseDownRow), mColPtr(lMCol), RowTop(mMouseDownRow), R, lgRTCheckBox
                                
                                If (X >= R.Left) And (Y >= R.Top) And (X <= R.Left + mR.CheckBoxSize) And (Y <= R.Top + mR.CheckBoxSize) Then
                                    bRedraw = True
                                    RaiseEvent RequestEdit(mMouseDownRow, lMCol, bCancel)
                                    
                                    If Not bCancel Then
                                        bState = (mItems(mRowPtr(mMouseDownRow)).Cell(mColPtr(lMCol)).nFlags And lgFLChecked)
                                        SetFlag mItems(mRowPtr(mMouseDownRow)).Cell(mColPtr(lMCol)).nFlags, lgFLChecked, Not bState
                                    End If
                                End If
                            End If
                        End If
                        
                        If Not bProcessed Then
                            bState = (mItems(mRowPtr(mMouseDownRow)).nFlags And lgFLSelected)
                            
                            If mMultiSelect Then
                                If (Shift And vbShiftMask) Then
                                    bSelectionChanged = SetSelection(False) Or SetSelection(True, mRow, mMouseDownRow)
                                ElseIf Shift And vbCtrlMask Then
                                    SetFlag mItems(mRowPtr(mMouseDownRow)).nFlags, lgFLSelected, Not bState
                                    bSelectionChanged = True
                                Else
                                    SetSelection False
                                    SetFlag mItems(mRowPtr(mMouseDownRow)).nFlags, lgFLSelected, True
                                    bSelectionChanged = True
                                End If
                            Else
                                If Shift And vbCtrlMask Then
                                    SetSelection False
                                    SetFlag mItems(mRowPtr(mMouseDownRow)).nFlags, lgFLSelected, Not bState
                                    bSelectionChanged = True
                                ElseIf Not bState Then
                                    SetSelection False
                                    SetFlag mItems(mRowPtr(mMouseDownRow)).nFlags, lgFLSelected, True
                                    bSelectionChanged = True
                                End If
                            End If
                        End If
                        
                        bRedraw = bRedraw Or SetRowCol(mMouseDownRow, lMCol)
                    End If
                    
                    If bRedraw Then
                        DrawGrid mRedraw
                    End If
                End If
            End If
        Else ' Right Button
            If mMouseDownRow > NULL_RESULT Then
                If UpdateCell() Then
                    SetRowCol mMouseDownRow, lMCol
                    bSelectionChanged = SetSelection(False) Or SetSelection(True, mMouseDownRow, mMouseDownRow)
                    DrawGrid mRedraw
                End If
            End If
        End If
        
        If bSelectionChanged Then
            RaiseEvent SelectionChanged
        End If
    End If
    
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static lResizeX As Long
    
    Dim R As RECT
    Dim lCount As Long
    Dim lWidth As Long
    Dim nMove As lgMoveControlEnum
    Dim nPointer As Integer
    Dim bSelectionChanged As Boolean
    
    If Not mLocked And (mItemCount >= 0) Then
        mMouseCol = GetColFromX(X)
        mMouseRow = GetRowFromY(Y)
        
        '####################################################################################
        'Header button tracking
        If mMouseDownCol <> NULL_RESULT Then
            If (mMouseDownCol = mMouseCol) And (MouseRow = NULL_RESULT) Then
                DrawHeader mMouseCol, lgDown
            Else
                DrawHeader mMouseDownCol, lgNormal
            End If
            UserControl.Refresh
        End If
        
        'Hot tracking
        If mHotHeaderTracking And (Button = 0) Then
            If Y < mR.HeaderHeight Then
                'Do we need to draw a new "hot" header?
                If (mMouseCol <> mHotColumn) Then
                    DrawHeaderRow
                    DrawHeader mMouseCol, lgHot
                    mHotColumn = mMouseCol
                End If
            ElseIf (mHotColumn <> NULL_RESULT) Then
                'We have a previous "hot" header to clear
                DrawHeaderRow
            End If
        End If
    
        '####################################################################################
        If (Button = vbLeftButton) Then
            If (mResizeCol >= 0) Then
                'We are resizing a Column
                lWidth = (X - lResizeX)
                If lWidth > 1 Then
                    mCols(mColPtr(mResizeCol)).lWidth = lWidth
                    mCols(mColPtr(mResizeCol)).dCustomWidth = ScaleX(mCols(mColPtr(mResizeCol)).lWidth, vbPixels, mScaleUnits)
                    
                    DrawGrid mRedraw
                    
                    nMove = mCols(mColPtr(mResizeCol)).MoveControl
                    RaiseEvent ColumnSizeChanged(mResizeCol, nMove)
                    
                    If mEditPending Then
                        MoveEditControl nMove
                    End If
                End If
            ElseIf (mMouseDownRow = NULL_RESULT) Then
                If mColumnDrag Then
                    DrawHeaderRow
                    
                    If (mMouseDownCol > NULL_RESULT) And (mDragCol = NULL_RESULT) Then
                        mDragCol = mMouseDownCol
                    End If
                    If (mDragCol <> NULL_RESULT) Then
                        mCols(mColPtr(mDragCol)).lX = mCols(mColPtr(mDragCol)).lX - (mMouseDownX - X)
                    End If
                End If
            Else
                If mMouseDown And Y < 0 Then
                    'Mouse has been dragged off off the control
                    ScrollList SCROLL_UP
                ElseIf mMouseDown And Y > UserControl.ScaleHeight Then
                    'Mouse has been dragged off off the control
                    ScrollList SCROLL_DOWN
                ElseIf mMouseDown And (Shift = 0) And (mMouseRow > NULL_RESULT) Then
                    If mScrollAction = SCROLL_NONE Then
                        bSelectionChanged = SetSelection(False)
                        
                        If mMultiSelect Then
                            SetSelection True, mMouseDownRow, mMouseRow
                        Else
                            SetSelection True, mMouseRow, mMouseRow
                        End If
                        
                        If SetRowCol(mMouseRow, mMouseCol) Then
                            RaiseEvent SelectionChanged
                            DrawGrid mRedraw
                        End If
                    Else
                        mScrollAction = SCROLL_NONE
                    End If
                End If
            End If
        ElseIf (Button = 0) Then
            nPointer = vbDefault
                
            'Only check for resize cursor if no buttons depressed
            If (mMouseRow = NULL_RESULT) Then
                lResizeX = 0
                mResizeCol = NULL_RESULT
                
                If (mAllowUserResizing = lgResizeCol) Or (mAllowUserResizing = lgResizeBoth) Then
                    For lCount = SBValue(efsHorizontal) To UBound(mCols)
                        If mCols(mColPtr(lCount)).bVisible Then
                            lWidth = lWidth + mCols(mColPtr(lCount)).lWidth
                            
                            If (X < lWidth + SIZE_VARIANCE) And (X > lWidth - SIZE_VARIANCE) Then
                                nPointer = vbSizeWE
                                mResizeCol = lCount
                                Exit For
                            End If
                            
                            lResizeX = lResizeX + mCols(mColPtr(lCount)).lWidth
                        End If
                    Next lCount
                End If
            End If
        
            With UserControl
                If .MousePointer <> nPointer Then
                    .MousePointer = nPointer
                End If
            End With
        End If
    End If
    
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim R As RECT
    Dim lCurrentMouseCol As Long
    Dim lCurrentMouseRow As Long
    Dim lTemp As Long
    
    If (Button = vbLeftButton) Then
        Call ReleaseCapture
        
        lCurrentMouseCol = GetColFromX(X)
        lCurrentMouseRow = GetRowFromY(Y)
        
        If (mDragCol >= 0) Then
            'We moved a Column
            If lCurrentMouseCol > NULL_RESULT Then
                lTemp = mColPtr(mDragCol)
                mColPtr(mDragCol) = mColPtr(lCurrentMouseCol)
                mColPtr(lCurrentMouseCol) = lTemp
            End If
            DrawGrid True
        ElseIf (mResizeCol >= 0) Then
            'We resized a Column so reset Scrollbars
            SetScrollBars
            DrawGrid mRedraw
            
            UserControl.MousePointer = vbDefault
        ElseIf (lCurrentMouseRow = NULL_RESULT) Then
            'Sort requested from Column Header click
            If (lCurrentMouseCol = mMouseDownCol) And (mMouseDownCol <> NULL_RESULT) Then
                If mColumnSort Then
                    If (Shift And vbCtrlMask) And (mSortColumn <> NULL_RESULT) Then
                        If mSortSubColumn <> mColPtr(mMouseDownCol) Then
                            mCols(mColPtr(mMouseDownCol)).nSortOrder = lgSTAscending
                        End If
                        mSortSubColumn = mColPtr(mMouseDownCol)
                        
                        Sort , mCols(mColPtr(mSortColumn)).nSortOrder
                    Else
                        If mSortColumn <> mColPtr(mMouseDownCol) Then
                            mCols(mColPtr(mMouseDownCol)).nSortOrder = lgSTAscending
                            mSortSubColumn = NULL_RESULT
                        End If
                        mSortColumn = mColPtr(mMouseDownCol)
                        
                        If mSortSubColumn <> NULL_RESULT Then
                            Sort , , , mCols(mColPtr(mSortSubColumn)).nSortOrder
                        Else
                            Sort
                        End If
                    End If
                Else
                    DrawHeaderRow
                    RaiseEvent ColumnClick(mMouseDownCol)
                End If
            End If
        ElseIf mMouseDownRow > NULL_RESULT Then
            If IsValidRowCol(mMouseRow, mMouseCol) Then
                If SetRowCol(mMouseRow, mMouseCol) Then
                    DrawGrid mRedraw
                End If
                
                If mCF(mItems(mRowPtr(mMouseRow)).Cell(mColPtr(mMouseCol)).nFormat).nImage <> 0 Then
                    SetItemRect mRowPtr(mMouseRow), mMouseCol, RowTop(mMouseRow), R, lgRTImage
                    If (X >= R.Left) And (Y >= R.Top) And (X <= R.Left + mR.ImageWidth) And (Y <= R.Top + mR.ImageHeight) Then
                        RaiseEvent CellImageClick(mMouseRow, mMouseCol)
                    End If
                ElseIf mItems(mRowPtr(mMouseRow)).Cell(mColPtr(mMouseCol)).nFlags And lgFLWordWrap Then
                    If mExpandRowImage > 0 Then
                        SetItemRect mRowPtr(mMouseRow), mMouseCol, RowTop(mMouseRow), R, lgRTImage
                        If (X >= R.Left) And (Y >= R.Top) And (X <= R.Left + mR.ImageWidth) And (Y <= R.Top + mR.ImageHeight) Then
                            If RowHeight(mMouseRow) = RowHeightMin Then
                                RowHeight(Row) = -1
                            Else
                                RowHeight(Row) = RowHeightMin
                            End If
                        End If
                    End If
                End If
            End If
        Else
            DrawHeaderRow
        End If
    End If
    
    mMouseDown = False
    mMouseDownCol = NULL_RESULT
    
    mDragCol = NULL_RESULT
    mResizeCol = NULL_RESULT
    
    mScrollAction = SCROLL_NONE
    
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        '################################################################################
        'Appearance Properties
        mApplySelectionToImages = .ReadProperty("ApplySelectionToImages", DEF_APPLYSELECTIONTOIMAGES)
        mBackColor = .ReadProperty("BackColor", DEF_BACKCOLOR)
        mBackColorBkg = .ReadProperty("BackColorBkg", DEF_BACKCOLORBKG)
        mBackColorEdit = .ReadProperty("BackColorEdit", DEF_BACKCOLOREDIT)
        mBackColorFixed = .ReadProperty("BackColorFixed", DEF_BACKCOLORFIXED)
        mBackColorSel = .ReadProperty("BackColorSel", DEF_BACKCOLORSEL)
        mForeColor = .ReadProperty("ForeColor", DEF_FORECOLOR)
        mForeColorEdit = .ReadProperty("ForeColorEdit", DEF_FORECOLOREDIT)
        mForeColorFixed = .ReadProperty("ForeColorFixed", DEF_FORECOLORFIXED)
        mForeColorHdr = .ReadProperty("ForeColorHdr", DEF_FORECOLORHDR)
        mForeColorSel = .ReadProperty("ForeColorSel", DEF_FORECOLORSEL)
        mForeColorTotals = .ReadProperty("ForeColorTotals", DEF_FORECOLORTOTALS)
        
        mGridColor = .ReadProperty("GridColor", DEF_GRIDCOLOR)
        mProgressBarColor = .ReadProperty("ProgressBarColor", DEF_PROGRESSBARCOLOR)
        
        mAlphaBlendSelection = .ReadProperty("AlphaBlendSelection", DEF_ALPHABLENDSELECTION)
        mBorderStyle = .ReadProperty("BorderStyle", DEF_BORDERSTYLE)
        mDisplayEllipsis = .ReadProperty("DisplayEllipsis", DEF_DISPLAYELLIPSIS)
        mFocusRectColor = .ReadProperty("FocusRectColor", DEF_FOCUSRECTCOLOR)
        mFocusRectMode = .ReadProperty("FocusRectMode", DEF_FOCUSRECTMODE)
        mFocusRectStyle = .ReadProperty("FocusRectStyle", DEF_FOCUSRECTSTYLE)
        mGridLines = .ReadProperty("GridLines", DEF_GRIDLINES)
        mGridLineWidth = .ReadProperty("GridLineWidth", DEF_GRIDLINEWIDTH)
        mThemeColor = .ReadProperty("ThemeColor", DEF_THEMECOLOR)
        mThemeStyle = .ReadProperty("ThemeStyle", DEF_THEMESTYLE)
        
        '################################################################################
        'Behaviour Properties
        mAllowUserResizing = .ReadProperty("AllowUserResizing", DEF_ALLOWUSERRESIZING)
        mAutoSizeRow = .ReadProperty("AutoSizeRow", DEF_AUTOSIZEROW)
        mCheckboxes = .ReadProperty("Checkboxes", DEF_CHECKBOXES)
        mColumnDrag = .ReadProperty("ColumnDrag", DEF_COLUMNDRAG)
        mColumnHeaders = .ReadProperty("ColumnHeaders", DEF_COLUMNHEADERS)
        mColumnSort = .ReadProperty("ColumnSort", DEF_COLUMNSORT)
        mEditable = .ReadProperty("Editable", DEF_EDITABLE)
        mEditTrigger = .ReadProperty("EditTrigger", DEF_EDITTRIGGER)
        mFullRowSelect = .ReadProperty("FullRowSelect", DEF_FULLROWSELECT)
        mHideSelection = .ReadProperty("HideSelection", DEF_HIDESELECTION)
        mHotHeaderTracking = .ReadProperty("HotHeaderTracking", DEF_HOTHEADERTRACKING)
        mMultiSelect = .ReadProperty("MultiSelect", DEF_MULTISELECT)
        mRedraw = .ReadProperty("Redraw", DEF_REDRAW)
        mScrollTrack = .ReadProperty("ScrollTrack", DEF_SCROLLTRACK)
        mTrackEdits = .ReadProperty("TrackEdits", DEF_TRACKEDITS)
        
        '################################################################################
        'Miscellaneous Properties
        mCacheIncrement = .ReadProperty("CacheIncrement", DEF_CACHEINCREMENT)
        mEnabled = .ReadProperty("Enabled", DEF_ENABLED)
        mFormatString = .ReadProperty("FormatString", DEF_FORMATSTRING)
        mLocked = .ReadProperty("Locked", DEF_LOCKED)
        mRowHeightMax = .ReadProperty("RowHeightMax", DEF_ROWHEIGHTMAX)
        mRowHeightMin = .ReadProperty("RowHeightMin", DEF_ROWHEIGHTMIN)
        mScaleUnits = .ReadProperty("ScaleUnits", DEF_SCALEUNITS)
        mSearchColumn = .ReadProperty("SearchColumn", DEF_SEARCHCOLUMN)
    
        Set Font = .ReadProperty("Font", Ambient.Font)
    End With
    
    '################################################################################
    'Apply Settings
    
    With UserControl
        .BackColor = mBackColorBkg
        .BorderStyle = mBorderStyle
    End With
    
    FormatString = mFormatString
    CreateRenderData
    SetColors
    
    '#############################################################################################################################
    'Subclassing
    If Ambient.UserMode Then
        bTrack = True
        bTrackUser32 = IsFunctionExported("TrackMouseEvent", "User32")
        
        If Not bTrackUser32 Then
            If Not IsFunctionExported("_TrackMouseEvent", "Comctl32") Then
                bTrack = False
            End If
        End If
        
        With UserControl
            Call Subclass_Start(.hwnd)
            Call Subclass_AddMsg(.hwnd, WM_KILLFOCUS, MSG_AFTER)
            Call Subclass_AddMsg(.hwnd, WM_SETFOCUS, MSG_AFTER)
            Call Subclass_AddMsg(.hwnd, WM_MOUSEWHEEL, MSG_AFTER)
            Call Subclass_AddMsg(.hwnd, WM_MOUSEMOVE, MSG_AFTER)
            Call Subclass_AddMsg(.hwnd, WM_MOUSELEAVE, MSG_AFTER)
            Call Subclass_AddMsg(.hwnd, WM_MOUSEHOVER, MSG_AFTER)
            Call Subclass_AddMsg(.hwnd, WM_HSCROLL, MSG_AFTER)
            Call Subclass_AddMsg(.hwnd, WM_VSCROLL, MSG_AFTER)
            
            If mWinXP Then
                Call Subclass_AddMsg(.hwnd, WM_THEMECHANGED)
            End If
        End With
        
        SBCreate UserControl.hwnd
        SBStyle = Style_Regular
        
        SBLargeChange(efsHorizontal) = 5
        SBSmallChange(efsHorizontal) = 1
        
        SBLargeChange(efsVertical) = 5
        SBSmallChange(efsVertical) = 1
     End If
End Sub

Private Sub UserControl_Resize()
    If m_hWnd <> 0 Then
        Refresh
    End If
End Sub

Private Sub UserControl_Terminate()
    On Local Error GoTo UserControl_TerminateError
    
    If Not mClipRgn = 0 Then DeleteObject mClipRgn
    
    pSBClearUp
    Call Subclass_Stop(UserControl.hwnd)

UserControl_TerminateError:
    Exit Sub
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("Font", mFont, Ambient.Font)
        
        '################################################################################
        'Appearance Properties
        Call .WriteProperty("ApplySelectionToImages", mApplySelectionToImages, DEF_APPLYSELECTIONTOIMAGES)
        Call .WriteProperty("BackColor", mBackColor, DEF_BACKCOLOR)
        Call .WriteProperty("BackColorBkg", mBackColorBkg, DEF_BACKCOLORBKG)
        Call .WriteProperty("BackColorEdit", mBackColorEdit, DEF_BACKCOLOREDIT)
        Call .WriteProperty("BackColorFixed", mBackColorFixed, DEF_BACKCOLORFIXED)
        Call .WriteProperty("BackColorSel", mBackColorSel, DEF_BACKCOLORSEL)
        Call .WriteProperty("ForeColor", mForeColor, DEF_FORECOLOR)
        Call .WriteProperty("ForeColorEdit", mForeColorEdit, DEF_FORECOLOREDIT)
        Call .WriteProperty("ForeColorFixed", mForeColorFixed, DEF_FORECOLORFIXED)
        Call .WriteProperty("ForeColorHdr", mForeColorHdr, DEF_FORECOLORHDR)
        Call .WriteProperty("ForeColorSel", mForeColorSel, DEF_FORECOLORSEL)
        Call .WriteProperty("ForeColorTotals", mForeColorTotals, DEF_FORECOLORTOTALS)
        
        Call .WriteProperty("GridColor", mGridColor, DEF_GRIDCOLOR)
        Call .WriteProperty("ProgressBarColor", mProgressBarColor, DEF_PROGRESSBARCOLOR)
        
        Call .WriteProperty("AlphaBlendSelection", mAlphaBlendSelection, DEF_ALPHABLENDSELECTION)
        Call .WriteProperty("BorderStyle", mBorderStyle, DEF_BORDERSTYLE)
        Call .WriteProperty("DisplayEllipsis", mDisplayEllipsis, DEF_DISPLAYELLIPSIS)
        Call .WriteProperty("FocusRectMode", mFocusRectMode, DEF_FOCUSRECTMODE)
        Call .WriteProperty("FocusRectColor", mFocusRectColor, DEF_FOCUSRECTCOLOR)
        Call .WriteProperty("FocusRectStyle", mFocusRectStyle, DEF_FOCUSRECTSTYLE)
        Call .WriteProperty("GridLines", mGridLines, DEF_GRIDLINES)
        Call .WriteProperty("GridLineWidth", mGridLineWidth, DEF_GRIDLINEWIDTH)
        Call .WriteProperty("ThemeColor", mThemeColor, DEF_THEMECOLOR)
        Call .WriteProperty("ThemeStyle", mThemeStyle, DEF_THEMESTYLE)
        
        '################################################################################
        'Behaviour Properties
        Call .WriteProperty("AllowUserResizing", mAllowUserResizing, DEF_ALLOWUSERRESIZING)
        Call .WriteProperty("AutoSizeRow", mAutoSizeRow, DEF_AUTOSIZEROW)
        Call .WriteProperty("Checkboxes", mCheckboxes, DEF_CHECKBOXES)
        Call .WriteProperty("ColumnDrag", mColumnDrag, DEF_COLUMNDRAG)
        Call .WriteProperty("ColumnHeaders", mColumnHeaders, DEF_COLUMNHEADERS)
        Call .WriteProperty("ColumnSort", mColumnSort, DEF_COLUMNSORT)
        Call .WriteProperty("Editable", mEditable, DEF_EDITABLE)
        Call .WriteProperty("EditTrigger", mEditTrigger, DEF_EDITTRIGGER)
        Call .WriteProperty("FullRowSelect", mFullRowSelect, DEF_FULLROWSELECT)
        Call .WriteProperty("HideSelection", mHideSelection, DEF_HIDESELECTION)
        Call .WriteProperty("HotHeaderTracking", mHotHeaderTracking, DEF_HOTHEADERTRACKING)
        Call .WriteProperty("MultiSelect", mMultiSelect, DEF_MULTISELECT)
        Call .WriteProperty("Redraw", mRedraw, DEF_REDRAW)
        Call .WriteProperty("ScrollTrack", mScrollTrack, DEF_SCROLLTRACK)
        Call .WriteProperty("TrackEdits", mTrackEdits, DEF_TRACKEDITS)
        
        '################################################################################
        'Miscellaneous Properties
        Call .WriteProperty("CacheIncrement", mCacheIncrement, DEF_CACHEINCREMENT)
        Call .WriteProperty("Enabled", mEnabled, DEF_ENABLED)
        Call .WriteProperty("FormatString", mFormatString, DEF_FORMATSTRING)
        Call .WriteProperty("Locked", mLocked, DEF_LOCKED)
        Call .WriteProperty("RowHeightMax", mRowHeightMax, DEF_ROWHEIGHTMAX)
        Call .WriteProperty("RowHeightMin", mRowHeightMin, DEF_ROWHEIGHTMIN)
        Call .WriteProperty("ScaleUnits", mScaleUnits, DEF_SCALEUNITS)
        Call .WriteProperty("SearchColumn", mSearchColumn, DEF_SEARCHCOLUMN)
    End With
End Sub

'========================================================================================
' Subclass code - The programmer may call any of the following Subclass_??? routines
'========================================================================================

Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)

'Add a message to the table of those that will invoke a callback. You should Subclass_Start first and then add the messages
'Parameters:
'   lng_hWnd - The handle of the window for which the uMsg is to be added to the callback table
'   uMsg     - The message number that will invoke a callback. NB Can also be ALL_MESSAGES, ie all messages will callback
'   When     - Whether the msg is to callback before, after or both with respect to the the default (previous) handler

    With sc_aSubData(zIdx(lng_hWnd))
        If (When And eMsgWhen.MSG_BEFORE) Then
            Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
        End If
        If (When And eMsgWhen.MSG_AFTER) Then
            Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
        End If
    End With

End Sub

'Private Sub Subclass_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
''Delete a message from the table of those that will invoke a callback.
''Parameters:
''   lng_hWnd - The handle of the window for which the uMsg is to be removed from the callback table
''   uMsg     - The message number that will be removed from the callback table. NB Can also be ALL_MESSAGES, ie all messages will callback
''   When     - Whether the msg is to be removed from the before, after or both callback tables
'
'    With sc_aSubData(zIdx(lng_hWnd))
'        If (When And eMsgWhen.MSG_BEFORE) Then
'            Call zDelMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
'        End If
'        If (When And eMsgWhen.MSG_AFTER) Then
'            Call zDelMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
'        End If
'    End With
'End Sub

Private Function Subclass_InIDE() As Boolean

'Return whether we're running in the IDE.

    Debug.Assert zSetTrue(Subclass_InIDE)

End Function

Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long

'Start subclassing the passed window handle
'Parameters:
'   lng_hWnd - The handle of the window to be subclassed
'Returns;
'   The sc_aSubData() index

Dim i                        As Long                       'Loop index
Dim j                        As Long                       'Loop index
Dim nSubIdx                  As Long                       'Subclass data index
Dim sSubCode                 As String                     'Subclass code string

Const PUB_CLASSES            As Long = 0                   'The number of UserControl public classes
Const GMEM_FIXED             As Long = 0                   'Fixed memory GlobalAlloc flag
Const PAGE_EXECUTE_READWRITE As Long = &H40&               'Allow memory to execute without violating XP SP2 Data Execution Prevention
Const PATCH_01               As Long = 18                  'Code buffer offset to the location of the relative address to EbMode
Const PATCH_02               As Long = 68                  'Address of the previous WndProc
Const PATCH_03               As Long = 78                  'Relative address of SetWindowsLong
Const PATCH_06               As Long = 116                 'Address of the previous WndProc
Const PATCH_07               As Long = 121                 'Relative address of CallWindowProc
Const PATCH_0A               As Long = 186                 'Address of the owner object
Const FUNC_CWPA               As String = "CallWindowProcA" 'We use CallWindowProc to call the original WndProc
Const FUNC_CWPW               As String = "CallWindowProcW" 'We use CallWindowProc to call the original WndProc
Const FUNC_EBM               As String = "EbMode"          'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
Const FUNC_SWLA               As String = "SetWindowLongA"  'SetWindowLongA allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
Const FUNC_SWLW               As String = "SetWindowLongW"  'SetWindowLongA allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
Const MOD_USER               As String = "user32"          'Location of the SetWindowLongA & CallWindowProc functions
Const MOD_VBA5               As String = "vba5"            'Location of the EbMode function if running VB5
Const MOD_VBA6               As String = "vba6"            'Location of the EbMode function if running VB6

'If it's the first time through here..

    If (sc_aBuf(1) = 0) Then

        'Build the hex pair subclass string
        sSubCode = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & _
                   "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & _
                   "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & _
                   "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90" & _
                   Hex$(&HA4 + (PUB_CLASSES * 12)) & "070000C3"

        'Convert the string from hex pairs to bytes and store in the machine code buffer
        i = 1
        Do While j < CODE_LEN
            j = j + 1
            sc_aBuf(j) = CByte("&H" & Mid$(sSubCode, i, 2))                       'Convert a pair of hex characters to an eight-bit value and store in the static code buffer array
            i = i + 2
        Loop                                                                      'Next pair of hex characters

        'Get API function addresses
        If (Subclass_InIDE) Then                                                  'If we're running in the VB IDE
            sc_aBuf(16) = &H90                                                    'Patch the code buffer to enable the IDE state code
            sc_aBuf(17) = &H90                                                    'Patch the code buffer to enable the IDE state code
            sc_pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)                            'Get the address of EbMode in vba6.dll
            If (sc_pEbMode = 0) Then                                              'Found?
                sc_pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)                        'VB5 perhaps
            End If
        End If

        Call zPatchVal(VarPtr(sc_aBuf(1)), PATCH_0A, ObjPtr(Me))                  'Patch the address of this object instance into the static machine code buffer
        'If IsWindowUnicode(lng_hWnd) Then
        If mWinNT Then
            sc_pCWP = zAddrFunc(MOD_USER, FUNC_CWPW)                                   'Get the address of the CallWindowsProc function
            sc_pSWL = zAddrFunc(MOD_USER, FUNC_SWLW)
        Else
            sc_pCWP = zAddrFunc(MOD_USER, FUNC_CWPA)                                   'Get the address of the CallWindowsProc function
            sc_pSWL = zAddrFunc(MOD_USER, FUNC_SWLA)
        End If
        'Get the address of the SetWindowLongA function
        ReDim sc_aSubData(0 To 0) As tSubData                                     'Create the first sc_aSubData element

    Else
        nSubIdx = zIdx(lng_hWnd, True)
        If (nSubIdx = -1) Then                                                    'If an sc_aSubData element isn't being re-cycled
            nSubIdx = UBound(sc_aSubData()) + 1                                   'Calculate the next element
            ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData                  'Create a new sc_aSubData element
        End If

        Subclass_Start = nSubIdx
    End If

    With sc_aSubData(nSubIdx)

        .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)                             'Allocate memory for the machine code WndProc
        Call VirtualProtect(ByVal .nAddrSub, CODE_LEN, PAGE_EXECUTE_READWRITE, i) 'Mark memory as executable
        Call RtlMoveMemory(ByVal .nAddrSub, sc_aBuf(1), CODE_LEN)                 'Copy the machine code from the static byte array to the code array in sc_aSubData

        .hwnd = lng_hWnd
        'Store the hWnd
        'If IsWindowUnicode(lng_hWnd) Then
        If mWinNT Then
            .nAddrOrig = SetWindowLongW(.hwnd, GWL_WNDPROC, .nAddrSub)
        Else
            .nAddrOrig = SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrSub)
        End If
        'Set our WndProc in place

        Call zPatchRel(.nAddrSub, PATCH_01, sc_pEbMode)                           'Patch the relative address to the VBA EbMode api function, whether we need to not.. hardly worth testing
        Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)                           'Original WndProc address for CallWindowProc, call the original WndProc
        Call zPatchRel(.nAddrSub, PATCH_03, sc_pSWL)                              'Patch the relative address of the SetWindowLongA api function
        Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)                           'Original WndProc address for SetWindowLongA, unsubclass on IDE stop
        Call zPatchRel(.nAddrSub, PATCH_07, sc_pCWP)                              'Patch the relative address of the CallWindowProc api function
    End With

End Function

Private Sub Subclass_StopAll()

'Stop all subclassing

Dim i As Long

    i = UBound(sc_aSubData())                                                     'Get the upper bound of the subclass data array
    Do While i >= 0                                                               'Iterate through each element
        With sc_aSubData(i)
            If (.hwnd <> 0) Then                                                  'If not previously Subclass_Stop'd
                Call Subclass_Stop(.hwnd)                                         'Subclass_Stop
            End If
        End With

        i = i - 1                                                                 'Next element
    Loop

End Sub

Private Sub Subclass_Stop(ByVal lng_hWnd As Long)

'Stop subclassing the passed window handle
'Parameters:
'   lng_hWnd - The handle of the window to stop being subclassed

    With sc_aSubData(zIdx(lng_hWnd))
        'If IsWindowUnicode(.hwnd) Then
        If mWinNT Then
            Call SetWindowLongW(.hwnd, GWL_WNDPROC, .nAddrOrig)
        Else
            Call SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrOrig)
        End If
        'Restore the original WndProc
        Call zPatchVal(.nAddrSub, PATCH_05, 0)                                    'Patch the Table B entry count to ensure no further 'before' callbacks
        Call zPatchVal(.nAddrSub, PATCH_09, 0)                                    'Patch the Table A entry count to ensure no further 'after' callbacks
        Call GlobalFree(.nAddrSub)                                                'Release the machine code memory
        .hwnd = 0                                                                 'Mark the sc_aSubData element as available for re-use
        .nMsgCntB = 0                                                             'Clear the before table
        .nMsgCntA = 0                                                             'Clear the after table
        Erase .aMsgTblB                                                           'Erase the before table
        Erase .aMsgTblA                                                           'Erase the after table
    End With

End Sub

'----------------------------------------------------------------------------------------
'These z??? routines are exclusively called by the Subclass_??? routines.
'----------------------------------------------------------------------------------------

Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)

'Worker sub for Subclass_AddMsg

Dim nEntry  As Long                                                             'Message table entry index
Dim nOff1   As Long                                                             'Machine code buffer offset 1
Dim nOff2   As Long                                                             'Machine code buffer offset 2

    If (uMsg = ALL_MESSAGES) Then                                                 'If all messages
        nMsgCnt = ALL_MESSAGES                                                    'Indicates that all messages will callback
    Else                                                                        'Else a specific message number
        Do While nEntry < nMsgCnt                                                 'For each existing entry. NB will skip if nMsgCnt = 0
            nEntry = nEntry + 1

            If (aMsgTbl(nEntry) = 0) Then                                         'This msg table slot is a deleted entry
                aMsgTbl(nEntry) = uMsg                                            'Re-use this entry
                Exit Sub                                                          'Bail
            ElseIf (aMsgTbl(nEntry) = uMsg) Then                                  'The msg is already in the table!
                Exit Sub                                                          'Bail
            End If
        Loop                                                                      'Next entry

        nMsgCnt = nMsgCnt + 1                                                     'New slot required, bump the table entry count
        ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long                              'Bump the size of the table.
        aMsgTbl(nMsgCnt) = uMsg                                                   'Store the message number in the table
    End If

    If (When = eMsgWhen.MSG_BEFORE) Then                                          'If before
        nOff1 = PATCH_04                                                          'Offset to the Before table
        nOff2 = PATCH_05                                                          'Offset to the Before table entry count
    Else                                                                        'Else after
        nOff1 = PATCH_08                                                          'Offset to the After table
        nOff2 = PATCH_09                                                          'Offset to the After table entry count
    End If

    If (uMsg <> ALL_MESSAGES) Then
        Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))                          'Address of the msg table, has to be re-patched because Redim Preserve will move it in memory.
    End If
    Call zPatchVal(nAddr, nOff2, nMsgCnt)                                         'Patch the appropriate table entry count

End Sub

Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long

'Return the memory address of the passed function in the passed dll

    zAddrFunc = ReturnAddr(sDLL, sProc)
    Debug.Assert zAddrFunc                                                        'You may wish to comment out this line if you're using vb5 else the EbMode GetProcAddress will stop here everytime because we look for vba6.dll first

End Function

'Private Sub zDelMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
''Worker sub for Subclass_DelMsg
'
'  Dim nEntry As Long
'
'    If (uMsg = ALL_MESSAGES) Then                                                 'If deleting all messages
'        nMsgCnt = 0                                                               'Message count is now zero
'        If When = eMsgWhen.MSG_BEFORE Then                                        'If before
'            nEntry = PATCH_05                                                     'Patch the before table message count location
'          Else                                                                    'Else after
'            nEntry = PATCH_09                                                     'Patch the after table message count location
'        End If
'        Call zPatchVal(nAddr, nEntry, 0)                                          'Patch the table message count to zero
'      Else                                                                        'Else deleteting a specific message
'        Do While nEntry < nMsgCnt                                                 'For each table entry
'            nEntry = nEntry + 1
'            If (aMsgTbl(nEntry) = uMsg) Then                                      'If this entry is the message we wish to delete
'                aMsgTbl(nEntry) = 0                                               'Mark the table slot as available
'                Exit Do                                                           'Bail
'            End If
'        Loop                                                                      'Next entry
'    End If
'End Sub

Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long

'Get the sc_aSubData() array index of the passed hWnd
'Get the upper bound of sc_aSubData() - If you get an error here, you're probably Subclass_AddMsg-ing before Subclass_Start

    zIdx = UBound(sc_aSubData)
    Do While zIdx >= 0                                                            'Iterate through the existing sc_aSubData() elements
        With sc_aSubData(zIdx)
            If (.hwnd = lng_hWnd) Then                                            'If the hWnd of this element is the one we're looking for
                If (Not bAdd) Then                                                'If we're searching not adding
                    Exit Function                                                 'Found
                End If
            ElseIf (.hwnd = 0) Then                                               'If this an element marked for reuse.
                If (bAdd) Then                                                    'If we're adding
                    Exit Function                                                 'Re-use it
                End If
            End If
        End With
        zIdx = zIdx - 1                                                           'Decrement the index
    Loop

    If (Not bAdd) Then
        Debug.Assert False                                                        'hWnd not found, programmer error
    End If

    'If we exit here, we're returning -1, no freed elements were found

End Function

Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)

'Patch the machine code buffer at the indicated offset with the relative address to the target address.

    Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)

End Sub

Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)

'Patch the machine code buffer at the indicated offset with the passed value

    Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)

End Sub

Private Function zSetTrue(ByRef bValue As Boolean) As Boolean

'Worker function for Subclass_InIDE

    zSetTrue = True
    bValue = True

End Function



