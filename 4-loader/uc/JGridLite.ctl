VERSION 5.00
Begin VB.UserControl JGridLite 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6600
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   440
   ToolboxBitmap   =   "JGridLite.ctx":0000
End
Attribute VB_Name = "JGridLite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'Autor          : J. Elihu
'----------------------------------------------------
'Version        : 1.2
'Requirements   : None
'Comments       : Based on the JGrid 2.6

Option Explicit

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal length As Long)
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32.dll" (ByRef lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32.dll" (ByRef lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long

Private Const MEM_COMMIT                As Long = &H1000
Private Const PAGE_EXECUTE_READWRITE    As Long = &H40
Private Const MEM_RELEASE               As Long = &H8000&


'/SCROLL
Private Type SCROLLINFO
  cbSize    As Long
  fMask     As Long
  nMin      As Long
  nMax      As Long
  nPage     As Long
  nPos      As Long
  nTrackPos As Long
End Type

Private Type TRIVERTEX
  x     As Long
  Y     As Long
  Red   As Integer
  Green As Integer
  Blue  As Integer
  Alpha As Integer
End Type


Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function SetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal n As Long, lpcScrollInfo As SCROLLINFO, ByVal bool As Boolean) As Long
Private Declare Function GetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal n As Long, lpScrollInfo As SCROLLINFO) As Long

Private Const SM_CXVSCROLL                  As Long = 2

'/Draw
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal Hdc As Long, ByVal x As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByVal lColorRef As Long) As Long
Private Declare Function OleTranslateColor2 Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal Hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal Hdc As Long, ByVal crColor As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal Hdc As Long, ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal Hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal Hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal Hdc As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal Hdc As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal Hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal Hdc As Long) As Long
Private Declare Function SetPixelV Lib "gdi32.dll" (ByVal Hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32.dll" (ByVal hBitmap As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal Hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal Hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Integer) As Long
Private Declare Function StretchBlt Lib "gdi32.dll" (ByVal Hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GdiTransparentBlt Lib "gdi32.dll" (ByVal Hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Private Declare Function SetStretchBltMode Lib "gdi32.dll" (ByVal Hdc As Long, ByVal nStretchMode As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal Hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Private Declare Function GradientFill Lib "msimg32" (ByVal Hdc As Long, ByRef Vertex As TRIVERTEX, ByVal nVertex As Long, ByRef Mesh As POINTAPI, ByVal nMesh As Long, ByVal mode As Long) As Long

Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal Y As Long) As Long

Private Declare Function TrackMouseEvent Lib "user32.dll" (ByRef lpEventTrack As Long) As Long

'?Border
Private Declare Function GetWindowRect& Lib "user32" (ByVal hWnd As Long, lpRect As RECT)
Private Declare Function ExcludeClipRect Lib "gdi32.dll" (ByVal Hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function GetWindowDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal Hdc As Long) As Long


Private Type POINTAPI
    x               As Long
    Y               As Long
End Type

Private Type RECT
    Left            As Long
    Top             As Long
    Right           As Long
    Bottom          As Long
End Type

Private Const WM_MOUSEWHEEL             As Long = &H20A
Private Const WM_VSCROLL                As Long = &H115
Private Const WM_HSCROLL                As Long = &H114
Private Const WM_NCPAINT                As Long = &H85
Private Const WM_DESTROY                As Long = &H2
Private Const WM_NCCALCSIZE             As Long = &H83
Private Const WM_MOUSELEAVE             As Long = &H2A3&

Private Const GWL_WNDPROC               As Long = -4
Private Const GWL_STYLE                 As Long = (-16)

Private Const WS_VSCROLL                As Long = &H200000
Private Const WS_HSCROLL                As Long = &H100000

Private Const SB_HORZ                   As Long = 0
Private Const SB_VERT                   As Long = 1
Private Const SB_BOTH                   As Long = 3
Private Const SB_LINEDOWN               As Long = 1
Private Const SB_LINEUP                 As Long = 0
Private Const SB_PAGEDOWN               As Long = 3
Private Const SB_PAGEUP                 As Long = 2
Private Const SB_THUMBTRACK             As Long = 5
Private Const SB_ENDSCROLL              As Long = 8
Private Const SB_LEFT                   As Long = 6
Private Const SB_RIGHT                  As Long = 7

Private Const SIF_ALL                   As Long = &H17
Private Const SM_CYBORDER               As Long = 6

Private Type tHeader
    Text    As String
    Width   As Long
    Aling   As Integer
    Left    As Long
End Type

Private Type tSubItem
    Text    As String
    Icon    As Long
End Type
Private Type tItem
    Item()  As tSubItem
    Tag     As String
End Type

Private pASMWrapper     As Long
Private PrevWndProc     As Long
Private hSubclassedWnd  As Long

Event SelectionChanged(ByVal Item As Long)


Private m_HeaderColor   As OLE_COLOR
Private m_ForeColor     As OLE_COLOR
Private m_HeaderBack    As OLE_COLOR
Private m_BorderColor   As OLE_COLOR
Private m_SelColor      As OLE_COLOR

Private m_Cols()        As tHeader
Private m_Items()       As tItem
Private m_smllChange(1) As Long
Private mouTrack(3)     As Long

Private m_ItemH         As Long
Private m_RowH          As Long
Private m_HeaderH       As Long
Private m_GridW         As Long
Private m_SelRow        As Long

Private m_HeaderRct     As RECT

Private mbTrack         As Boolean
Private mbResizeCol     As Boolean
Private mbNoDraw        As Boolean
Private mlColHit        As Long

Public Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    WindowProc = CallWindowProc(PrevWndProc, hWnd, uMsg, wParam, lParam)

    Select Case uMsg
        
        Case WM_DESTROY
            'Call StopSubclassing

        Case WM_HSCROLL, WM_VSCROLL
        
            Dim lSB     As Long
            Dim tSI     As SCROLLINFO
            Dim lSBCode As Long
            Dim lv      As Long
            Dim lSC     As Long
            
            lSB = IIf(uMsg = WM_VSCROLL, SB_VERT, SB_HORZ)
            lSBCode = (wParam And &HFFFF&) 'LoWord(wParam)
            Select Case lSBCode
                Case SB_LINEDOWN

                    lv = sb_value(lSB)
                    lSC = m_smllChange(lSB)
                    
                    If (lv + lSC > sb_max(lSB)) Then
                        sb_value(lSB) = sb_max(lSB)
                    Else
                        sb_value(lSB) = lv + lSC
                    End If
                    'pRaiseEvent lSB, False
                    
                Case SB_LINEUP
                    
                    lv = sb_value(lSB)
                    lSC = m_smllChange(lSB)
                    
                    If (lv - lSC < sb_min(lSB)) Then
                        sb_value(lSB) = sb_min(lSB)
                    Else
                        sb_value(lSB) = lv - lSC
                    End If
                    'pRaiseEvent lSB, False
                    
                Case SB_PAGEDOWN
                    sb_value(lSB) = sb_value(lSB) + sb_large_change(lSB)
                    'pRaiseEvent lSB, False
            
                Case SB_PAGEUP
                    sb_value(lSB) = sb_value(lSB) - sb_large_change(lSB)
                    'pRaiseEvent lSB, False
            
                Case SB_THUMBTRACK
                
                    pSBGetSI lSB, tSI, &H10
                    sb_value(lSB) = tSI.nTrackPos 'HiWord(wParam)
                    'pRaiseEvent lSB, True
                Case SB_ENDSCROLL
                    '
                Case SB_LEFT
                    sb_value(lSB) = sb_min(lSB)
                    'pRaiseEvent lSB, False
                    
                Case SB_RIGHT
                    sb_value(lSB) = sb_max(lSB)
                    'pRaiseEvent lSB, False
            End Select
            DrawGrid True

        Case WM_MOUSEWHEEL
        
            Dim m_lWheelScrollLines As Long
            Dim m_lSmallChangeVert As Long
            
            lSB = SB_VERT
            If sb_max(lSB) = 0 Then Exit Function
            
            If wParam < 0 Then
                sb_value(lSB) = sb_value(lSB) + m_smllChange(lSB)
            Else
                sb_value(lSB) = sb_value(lSB) - m_smllChange(lSB)
            End If
            DrawGrid True
            
            'm_lWheelScrollLines = 3
            'm_lSmallChangeVert = 1
            
            'lDelta = (zDelta \ 120) * m_lSmallChangeVert * m_lWheelScrollLines
            'lSB = SB_VERT
            'RaiseEvent MouseWheel(eBar, lDelta)
            'If Not (lDelta = 0) And sb_value(lSB) Then
                'If sb_value(lSB) Then sb_value(lSB) = sb_value(lSB) + lDelta
            'End If
            
        Case WM_NCPAINT
        
            If UserControl.BorderStyle <> vbFixedSingle Then Exit Function

            Dim Rct     As RECT
            Dim Rct1    As RECT
            Dim dvc     As Long
            Dim lSize   As Long
            
            
            dvc = GetWindowDC(hWnd)
            GetWindowRect hWnd, Rct
            lSize = GetSystemMetrics(&H6) 'SM_CYBORDER
            
            Rct.Right = Rct.Right - Rct.Left
            Rct.Bottom = Rct.Bottom - Rct.Top
            Rct.Left = 0
            Rct.Top = 0
            
            ExcludeClipRect dvc, lSize + 1, lSize + 1, Rct.Right - lSize - 1, Rct.Bottom - lSize - 1
            
            Dim hPen        As Long
            Dim OldPen      As Long
                                
            hPen = CreatePen(0, lSize, m_BorderColor)
            OldPen = SelectObject(dvc, hPen)
            Rectangle dvc, Rct.Left, Rct.Top, Rct.Right, Rct.Bottom
            Call SelectObject(dvc, OldPen)
            DeleteObject hPen
                               
            ReleaseDC hWnd, dvc
            
        Case WM_MOUSELEAVE
            
            mlColHit = -1
            DrawGrid
            mbTrack = False
   
    End Select

End Function


Private Sub UserControl_Initialize()

    Dim ASM(0 To 104) As Byte
    Dim pVar As Long
    Dim ThisClass As Long
    Dim CallbackFunction As Long
    Dim pVirtualFree
    Dim sCode As String
    Dim i As Long
        
    pASMWrapper = VirtualAlloc(ByVal 0&, 104, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    If pASMWrapper <> 0 Then
    

        ThisClass = ObjPtr(Me)
        Call CopyMemory(pVar, ByVal ThisClass, 4)
        Call CopyMemory(CallbackFunction, ByVal (pVar + 1956), 4)

        pVirtualFree = GetProcAddress(GetModuleHandle("kernel32.dll"), "VirtualFree")
        
        sCode = "90FF05000000006A0054FF742418FF742418FF742418FF7424186800000000B800000000FFD0FF0D00000000A10000000085C075" & _
                "0458C21000A10000000085C0740458C2100058595858585868008000006A00680000000051B800000000FFE00000000000000000"
                
        For i = 0 To Len(sCode) - 1 Step 2
            ASM(i / 2) = CByte("&h" & Mid$(sCode, i + 1, 2))
        Next
        
        
        Call CopyMemory(ASM(3), pASMWrapper + 96, 4)
        Call CopyMemory(ASM(40), pASMWrapper + 96, 4)
        Call CopyMemory(ASM(58), pASMWrapper + 96, 4)
        Call CopyMemory(ASM(45), pASMWrapper + 100, 4)
        Call CopyMemory(ASM(84), pASMWrapper, 4)
        Call CopyMemory(ASM(27), ThisClass, 4)
        Call CopyMemory(ASM(32), CallbackFunction, 4)
        Call CopyMemory(ASM(90), pVirtualFree, 4)
        Call CopyMemory(ByVal pASMWrapper, ASM(0), 104)
    End If
    
    m_SelRow = -1
    mlColHit = -1
End Sub

Private Sub UserControl_InitProperties()
    m_HeaderBack = &HF0F0F0
    m_HeaderColor = 0 'vbWhite
    UserControl.BackColor = vbWhite
    m_BorderColor = &H908782  '&HB2ACA5
    m_HeaderH = 26
    m_SelColor = &HFF6600  '&HDDAC84
    
End Sub


Private Sub UserControl_Terminate()
    Dim Counter As Long
    Dim Flag As Long

    If pASMWrapper <> 0 Then
        Call StopSubclassing
        Call CopyMemory(Counter, ByVal (pASMWrapper + 104), 4)
        If Counter = 0 Then
            Call VirtualFree(ByVal pASMWrapper, 0, MEM_RELEASE)
        Else
            Flag = 1
            Call CopyMemory(ByVal (pASMWrapper + 108), Flag, 4)
        End If
    End If

End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag

        .WriteProperty "BackColor", UserControl.BackColor
        .WriteProperty "HeaderH", m_HeaderH
        .WriteProperty "ForeColor", m_ForeColor
        .WriteProperty "HeaderColor", m_HeaderColor
        .WriteProperty "HeaderBack", m_HeaderBack
        .WriteProperty "BorderColor", m_BorderColor
        .WriteProperty "SelColor", m_SelColor
        
    End With
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
    
        UserControl.BackColor = .ReadProperty("BackColor", vbWhite)
        m_HeaderH = .ReadProperty("HeaderH", 26)
        
        m_ForeColor = .ReadProperty("ForeColor", 0)
        m_HeaderColor = .ReadProperty("HeaderColor", 0)
        m_HeaderBack = .ReadProperty("HeaderBack", &HF0F0F0)
        m_BorderColor = .ReadProperty("BorderColor", &HB2ACA5)
        m_SelColor = .ReadProperty("SelColor", &HFF6600)

    End With
    
    If Ambient.UserMode Then
        SetSubclassing UserControl.hWnd
        
        m_smllChange(SB_HORZ) = 20
        m_smllChange(SB_VERT) = 16
        
        SetRect m_HeaderRct, 0, 0, UserControl.ScaleWidth, m_HeaderH
        UpateValues1
        
        mouTrack(0) = 16&
        mouTrack(1) = &H2
        mouTrack(2) = UserControl.hWnd
    End If
    
    
    DrawGrid True
End Sub
Private Sub UserControl_Resize()
    SetRect m_HeaderRct, 0, 0, UserControl.ScaleWidth, m_HeaderH
    UpdateScrollV
    UpdateScrollH
    DrawGrid True
End Sub
Private Sub UserControl_Show()
    DrawGrid True
End Sub
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
Dim iRow As Long
Dim iCol As Long

    'RaiseEvent KeyDown(KeyCode, Shift)
    If ItemCount = 0 Then Exit Sub
    
    Select Case KeyCode
        Case vbKeyDown
            If m_SelRow < ItemCount - 1 Then ChangeSelection m_SelRow + 1
        Case vbKeyUp
            If m_SelRow > 0 Then ChangeSelection m_SelRow - 1
        Case vbKeyRight, vbKeyTab
            '
        Case vbKeyLeft
            '
        Case vbKeyEnd, vbKeyHome
            If KeyCode = vbKeyEnd Then ChangeSelection ItemCount - 1
            If KeyCode = vbKeyHome Then ChangeSelection 0
        Case vbKeyPageDown, vbKeyPageUp
            If KeyCode = vbKeyPageDown Then sb_value(SB_VERT) = sb_value(SB_VERT) + sb_large_change(SB_VERT)
            If KeyCode = vbKeyPageUp Then sb_value(SB_VERT) = sb_value(SB_VERT) - sb_large_change(SB_VERT)
        Case Else
            
            On Error Resume Next
            Dim j           As Long
            Dim lStart      As Long
            Dim pChar       As String
            Dim iChar       As String
            Dim bFound      As Boolean
            Dim lCol        As Long
        
            lStart = m_SelRow + 1
            lCol = 0
            If lStart > ItemCount - 1 Then lStart = 0
            pChar = Chr(KeyCode)
            If pChar = "" Then Exit Sub
            
            For j = lStart To ItemCount - 1
                iChar = UCase(Left(m_Items(j).Item(lCol).Text, 1))
                If iChar <> "" And pChar = iChar Then
                    ChangeSelection j
                    bFound = True
                    Exit For
                End If
            Next
            If Not bFound And lStart > 0 Then
                For j = 0 To lStart '- 1
                    iChar = UCase(Left(m_Items(j).Item(lCol).Text, 1))
                    If iChar <> "" And pChar = iChar Then
                        ChangeSelection j
                        Exit For
                    End If
                Next
            End If
            
    End Select
End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim lRow As Long
    
    If PtInRect(m_HeaderRct, x, Y) And Button = 1 And UserControl.MousePointer = vbSizeWE Then
        mbResizeCol = True
    End If
    
    If x > m_GridW Then Exit Sub
    If Y <= m_ItemH Then Exit Sub
    
    lRow = GetRowFromY(Y)
    If lRow = -1 Then Exit Sub
    'If lRow <> m_SelRow Then
        ChangeSelection lRow
    'End If
    
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim lPointer    As Long
Dim lx1         As Long
Dim lx2         As Long
Dim i           As Long
    
    If Not mbTrack Then
        TrackMouseEvent mouTrack(0)
        mbTrack = True
    End If

    If PtInRect(m_HeaderRct, x, Y) = 0 Then
        If mbResizeCol = 0 Then GoTo Items
    End If

    lPointer = vbDefault
    lx1 = x + get_scroll(SB_HORZ)
    
    Select Case Button
        Case 0
            
            mbResizeCol = False
            For i = 0 To ColumnCount - 1
                lx2 = lx2 + m_Cols(i).Width
                If lx1 >= lx2 - m_Cols(i).Width And lx1 <= lx2 Then
                    mlColHit = i
                    If (lx1 < lx2 + 2) And (lx1 > lx2 - 2) Then lPointer = vbSizeWE ': mbResizeCol = True
                    Exit For
                End If
            Next
            DrawGrid
        Case 1:
            If mbResizeCol Then
                lPointer = vbSizeWE
                lx2 = lx1 - m_Cols(mlColHit).Left
                If lx2 < 3 Then lx2 = 3
                m_GridW = (m_GridW - m_Cols(mlColHit).Width) + lx2
                m_Cols(mlColHit).Width = lx2
                For i = mlColHit + 1 To ColumnCount - 1
                    m_Cols(i).Left = m_Cols(i - 1).Left + m_Cols(i - 1).Width
                Next
                UpdateScrollH
                DrawGrid
            Else
                Debug.Print "Here " & Timer
            End If
    End Select
    
    If UserControl.MousePointer <> lPointer Then UserControl.MousePointer = lPointer
    Exit Sub
Items:
    If UserControl.MousePointer <> lPointer Then UserControl.MousePointer = lPointer
    If mlColHit <> -1 Then mlColHit = -1: DrawGrid
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If mbResizeCol Then
        mbResizeCol = False
        If PtInRect(m_HeaderRct, x, Y) = 0 Then UserControl.MousePointer = vbDefault
        UpdateScrollH
        DrawGrid
    End If
End Sub


Public Sub AddColumn(Text As String, Optional ByVal Width As Integer = 100, Optional ByVal Alignment As AlignmentConstants)
Dim l As Long

    l = ColumnCount
    ReDim Preserve m_Cols(l)
    With m_Cols(l)
        .Text = Text
        .Width = Width
        .Aling = Alignment
        .Left = m_GridW
    End With
    m_GridW = m_GridW + (Width)
    UpdateScrollH
End Sub

Public Sub AddRow(ParamArray Fields() As Variant)
On Error Resume Next
Dim i   As Long
Dim l   As Long
Dim Tw  As Long
Dim lW  As Long

    l = ItemCount
    ReDim Preserve m_Items(l)

    With m_Items(l)
        ReDim .Item(ColumnCount - 1)
        For i = 0 To UBound(Fields)
            If i > ColumnCount - 1 Then Exit For
            .Item(i).Text = CStr(Fields(i))
        Next
    End With
    If mbNoDraw Then Exit Sub
    UpdateScrollV
    'If IsVisibleRow(l) Then DrawGrid
    
End Sub
Public Sub ClearTable()
    Call Clear
    Erase m_Cols
    UpdateScrollH
    m_GridW = 0
    DrawGrid True
End Sub
Public Sub Clear()
    Erase m_Items
    m_SelRow = -1
    UpdateScrollV
    DrawGrid True
End Sub
Public Sub Refresh()
   'Draw
End Sub

Property Get ColumnCount() As Long
On Error GoTo e
    ColumnCount = UBound(m_Cols) + 1
e:
End Property

Property Get ItemCount() As Long
On Error GoTo e
    ItemCount = UBound(m_Items) + 1
e:
End Property

Property Get BorderColor() As OLE_COLOR: BorderColor = m_BorderColor: End Property
Property Let BorderColor(ByVal Value As OLE_COLOR)
    m_BorderColor = Value
    PropertyChanged "BorderColor"
End Property
Property Get ForeColor() As OLE_COLOR: ForeColor = m_ForeColor: End Property
Property Let ForeColor(ByVal Value As OLE_COLOR)
    m_ForeColor = Value
    PropertyChanged "ForeColor"
    DrawGrid True
End Property
Property Get HeaderForeColor() As OLE_COLOR: HeaderForeColor = m_HeaderColor: End Property
Property Let HeaderForeColor(ByVal Value As OLE_COLOR)
    m_HeaderColor = Value
    PropertyChanged "HeaderColor"
    DrawGrid True
End Property
Property Get HeaderBackColor() As OLE_COLOR: HeaderBackColor = m_HeaderBack: End Property
Property Let HeaderBackColor(ByVal Value As OLE_COLOR)
    m_HeaderBack = Value
    PropertyChanged "HeaderBack"
    DrawGrid True
End Property
Property Get SelectionColor() As OLE_COLOR: SelectionColor = m_SelColor: End Property
Property Let SelectionColor(ByVal Value As OLE_COLOR)
    m_SelColor = Value
    DrawGrid
    PropertyChanged "SelColor"
End Property
Property Get HeaderHeight() As Long: HeaderHeight = m_HeaderH: End Property
Property Let HeaderHeight(ByVal Value As Long)
    m_HeaderH = Value
    UpdateScrollV
    DrawGrid
    PropertyChanged "HeaderH"
End Property

Property Let NoDraw(Value As Boolean)
    mbNoDraw = Value
    If Not Value Then UpdateScrollV: DrawGrid
End Property


Property Get ItemText(ByVal Item As Long, Optional ByVal Column As Long) As String
On Local Error Resume Next
    ItemText = m_Items(Item).Item(Column).Text
End Property
Property Let ItemText(ByVal Item As Long, Optional ByVal Column As Long, Value As String)
On Local Error Resume Next
    If m_Items(Item).Item(Column).Text = Value Then Exit Property
    m_Items(Item).Item(Column).Text = Value
    DrawGrid
End Property

Property Get ItemTag(ByVal Item As Long) As String
On Local Error Resume Next
    ItemTag = m_Items(Item).Tag
End Property
Property Let ItemTag(ByVal Item As Long, ByVal Value As String)
On Local Error Resume Next
    If m_Items(Item).Tag = Value Then Exit Property
    m_Items(Item).Tag = Value
End Property

Property Get SelectedItem() As Long: SelectedItem = m_SelRow: End Property
Property Let SelectedItem(ByVal Value As Long)
    If Value < 0 Then Value = -1
    If Value > ItemCount - 1 Then Value = -1
    If m_SelRow <> Value Then
        ChangeSelection Value
    End If
End Property

Property Get ScrollBar(ByVal Index As Long) As Long: ScrollBar = sb_value(Index): End Property
Property Let ScrollBar(ByVal Index As Long, ByVal Value As Long)
    sb_value(Index) = Value
End Property



Private Property Get lHeaderH() As Long
    'lHeaderH = IIf(m_Header, m_HeaderH * dpiScale, 0)
    lHeaderH = m_HeaderH
End Property
Private Property Get lGridH() As Long
    lGridH = UserControl.ScaleHeight - lHeaderH
End Property



'/? Private Subs

Private Sub UpateValues1()
Dim Th As Integer
Dim Px As Long
    
    Px = 6 '* e_Scale
    Th = UserControl.TextHeight("ÀJ")
    If Th + Px > m_ItemH Then m_ItemH = Th + Px
    m_RowH = m_ItemH
    m_smllChange(SB_VERT) = m_RowH
    
End Sub

Private Sub UpdateScrollV()
On Error Resume Next
Dim lHeight     As Long
Dim lProportion As Long

    lHeight = ((ItemCount * m_RowH) + 5) - (UserControl.ScaleHeight - m_HeaderH)
    If (lHeight > 0) Then
      lProportion = lHeight \ ((UserControl.ScaleHeight - m_HeaderH) + 1)
      sb_max(SB_VERT) = lHeight
      sb_large_change(SB_VERT) = lHeight \ lProportion
    Else
      sb_large_change(SB_VERT) = 0
      sb_max(SB_VERT) = 0
    End If

End Sub

Private Sub UpdateScrollH()
On Error Resume Next
Dim lWidth      As Long
Dim lProportion As Long
    
    lWidth = m_GridW - (UserControl.ScaleWidth - (GetSystemMetrics(SM_CXVSCROLL)))
    If (lWidth > 0) Then
        lProportion = lWidth \ (UserControl.ScaleWidth) + 1
        sb_large_change(SB_HORZ) = lWidth \ lProportion
        sb_max(SB_HORZ) = lWidth
    Else
        sb_large_change(SB_HORZ) = 0
        sb_max(SB_HORZ) = 0
    End If
    
End Sub

Private Sub ChangeSelection(eRow As Long)

    'If Not IsCompleteVisibleItem(eRow, eCol) Then SetVisibleItem eRow, eCol
    If eRow = m_SelRow Then
        If Not IsCompleteVisibleRow(eRow) Then SetVisibleRow eRow
        Exit Sub
    End If
    
    m_SelRow = eRow

    If m_SelRow = -1 Then
        DrawGrid
        GoTo Evt
    End If
    
    If Not IsCompleteVisibleRow(eRow) Then
        SetVisibleRow eRow
    Else
        DrawGrid
    End If
    
Evt:
    RaiseEvent SelectionChanged(eRow)
End Sub
Private Sub SetVisibleRow(eRow As Long)
On Error GoTo e
Dim lx  As Integer
Dim ly  As Integer
Dim Rct     As RECT
    
    If eRow = -1 Then Exit Sub
    ly = eRow * m_RowH
    
    '?Vertical
    If (ly + m_RowH) - (UserControl.ScaleHeight - m_HeaderH) > sb_value(SB_VERT) Then
        sb_value(SB_VERT) = ((ly + m_RowH) + 2) - (UserControl.ScaleHeight - m_HeaderH)
    ElseIf ly < sb_value(SB_VERT) Then
        sb_value(SB_VERT) = ly
    End If
e:
    DrawGrid
End Sub
Private Function get_scroll(lSB As Long) As Long
    get_scroll = sb_value(lSB)
End Function

Private Function IsVisibleRow(ByVal eRow As Long) As Boolean
On Error Resume Next
Dim Y As Long
    If sb_max(SB_VERT) = 0 Then IsVisibleRow = True: Exit Function
    Y = (eRow * m_RowH) - sb_value(1)
    IsVisibleRow = (Y + m_ItemH > 0) And Y <= (UserControl.ScaleHeight - m_HeaderH)
End Function
Private Function IsCompleteVisibleRow(eRow As Long) As Boolean
On Local Error Resume Next
Dim Y       As Long
Dim bRow    As Boolean

    Y = (eRow * m_RowH) - sb_value(SB_VERT)
    'bRow = (Y >= 0) And (Y + m_ItemH <= m_GridH)
    IsCompleteVisibleRow = bRow
    
End Function

Private Sub DrawGrid(Optional ByVal bForce As Boolean)
On Local Error Resume Next
Dim lCol    As Long
Dim lRow    As Long
Dim ly      As Long
Dim lx      As Long
Dim lSx     As Long 'Start X
Dim lSCol   As Long 'Start Col
Dim lColW   As Long
Dim dvc     As Long
Dim IRct    As RECT
Dim tRct    As RECT
Dim lPx     As Long
Dim lPx2    As Long


    UserControl.AutoRedraw = True
    UserControl.Cls
    

    lCol = 0
    lRow = 0
    
    lx = -get_scroll(SB_HORZ)
    ly = -get_scroll(SB_VERT)
    
    dvc = UserControl.Hdc

    ly = ly + m_HeaderH
    lSx = lx
    lSCol = -1
    
    
    Do While lRow <= ItemCount - 1 And ly < UserControl.ScaleHeight
    
        If ly + m_RowH > 0 Then '?Visible
        
            If lRow = m_SelRow Then
            
                lPx = -sb_value(SB_HORZ)
                
                SetRect IRct, lPx, ly, lPx + m_GridW, ly + m_ItemH
                DrawBack dvc, BlendColor(m_SelColor, UserControl.BackColor, 80), IRct
                DrawBorder dvc, BlendColor(SysColor(m_SelColor), UserControl.BackColor, 125), lPx, ly, m_GridW, m_ItemH
                
            End If

            Do While lCol < ColumnCount And lx < UserControl.ScaleWidth
            
                lColW = m_Cols(lCol).Width
                
                If lx + lColW > 0 Then
                    
                    If (lSCol = -1) Then
                        lSCol = lCol
                        lSx = lx
                    End If
                    
                    SetRect IRct, lx, ly, lx + lColW, ly + m_ItemH
                    
                    If Trim(m_Items(lRow).Item(lCol).Text) <> vbNullString Then
                        SetRect tRct, lx + (4), ly, lx + lColW - (2), ly + m_ItemH
                        
                        If tRct.Right < tRct.Left Then tRct.Right = tRct.Left
                        'If tRct.Right - tRct.Left > 0 Then
                        UserControl.ForeColor = m_ForeColor
                        DrawText dvc, m_Items(lRow).Item(lCol).Text, Len(m_Items(lRow).Item(lCol).Text), tRct, GetTextFlag(lCol)
                        
                    End If
                    
                End If
                
                lx = lx + m_Cols(lCol).Width
                lCol = lCol + 1
            Loop
            
            '?Reset to Scroll Position
            lCol = lSCol
            lx = lSx
        End If
    
        ly = ly + m_RowH
        lRow = lRow + 1
    Loop
    DrawHeader
    UserControl.AutoRedraw = False
End Sub
Private Sub DrawHeader()
Dim lCol    As Long
Dim lColW   As Long
Dim lx      As Long
Dim IRct    As RECT
Dim lPx     As Long

    '/? Header
    SetRect IRct, 0, 0, UserControl.ScaleWidth, m_HeaderH
    DrawBack UserControl.Hdc, UserControl.BackColor, IRct
    SetRect IRct, 0, 0, UserControl.ScaleWidth, m_HeaderH - 1
    FillGradient UserControl.Hdc, IRct, BlendColor(SysColor(m_HeaderBack), vbWhite, 125), SysColor(m_HeaderBack), True
    DrawLine UserControl.Hdc, 0, m_HeaderH - 1, UserControl.ScaleWidth, m_HeaderH - 1, BlendColor(m_HeaderColor, m_HeaderBack, 75)

    lx = -get_scroll(SB_HORZ)
    UserControl.ForeColor = m_HeaderColor
    
    Do While lCol < ColumnCount And lx < UserControl.ScaleWidth
        lColW = m_Cols(lCol).Width
        If lx + lColW > 0 Then
            

            If mlColHit = lCol Then
                SetRect IRct, lx + 1, 0, lx + lColW, m_HeaderH
                FillGradient UserControl.Hdc, IRct, ShiftColor(&HFBEFCC, 15), ShiftColor(&HFBEFCC, 5), True
                DrawBorder UserControl.Hdc, &HE0A57D, lx, -1, lColW, m_HeaderH + 1
            Else
                DrawLine UserControl.Hdc, lx + lColW, 3, lx + lColW, m_HeaderH - 5, BlendColor(m_HeaderColor, m_HeaderBack, 50)
            End If
            
            SetRect IRct, lx + (4), 0, lx + lColW - (2), m_HeaderH
            If IRct.Right - IRct.Left < 3 Then IRct.Right = IRct.Left
            DrawText UserControl.Hdc, m_Cols(lCol).Text, Len(m_Cols(lCol).Text), IRct, GetTextFlag(lCol)
        End If
        lx = lx + m_Cols(lCol).Width
        lCol = lCol + 1
    Loop
    
    
End Sub

Private Sub DrawBack(lpDC As Long, Color As Long, Rct As RECT)
Dim hBrush  As Long

    hBrush = CreateSolidBrush(Color)
    Call FillRect(lpDC, Rct, hBrush)
    Call DeleteObject(hBrush)
    
End Sub
Private Sub DrawBorder(lpDC As Long, Color As Long, x As Long, Y As Long, W As Long, H As Long)
Dim hPen As Long
    hPen = CreatePen(0, 1, Color)
    Call SelectObject(lpDC, hPen)
    RoundRect lpDC, x, Y, x + W, Y + H, 0, 0
    DeleteObject hPen
End Sub
Private Sub DrawLine(lpDC As Long, x As Long, Y As Long, x2 As Long, y2 As Long, Color As Long)
Dim PT      As POINTAPI
Dim hPen    As Long
Dim hPenOld As Long

    hPen = CreatePen(0, 1, Color)
    hPenOld = SelectObject(lpDC, hPen)
    Call MoveToEx(lpDC, x, Y, PT)
    Call LineTo(lpDC, x2, y2)
    Call SelectObject(lpDC, hPenOld)
    Call DeleteObject(hPen)
    
End Sub

Private Sub FillGradient(dvc As Long, Rct As RECT, ByVal Color1 As OLE_COLOR, ByVal Color2 As OLE_COLOR, Optional ByVal FillVertical As Boolean)
Dim TRVRT(1)   As TRIVERTEX
Dim PT         As POINTAPI
'Dim pGradRect   As GRADIENT_RECT
    
    With TRVRT(0)
        .x = Rct.Left
        .Y = Rct.Top
        .Red = LongToSignedShort(Color1 And &HFF& * 256)
        .Green = LongToSignedShort(((Color1 And &HFF00&) / &H100&) * 256)
        .Blue = LongToSignedShort(((Color1 And &HFF0000) / &H10000) * 256)
        .Alpha = 0
    End With
    
    With TRVRT(1)
        .x = Rct.Right
        .Y = Rct.Bottom
        .Red = LongToSignedShort((Color2 And &HFF&) * 256)
        .Green = LongToSignedShort(((Color2 And &HFF00&) / &H100&) * 256)
        .Blue = LongToSignedShort(((Color2 And &HFF0000) / &H10000) * 256)
        .Alpha = 0
    End With
    PT.Y = 1
    GradientFill dvc, TRVRT(0), 2, PT, 1, Abs(FillVertical)
End Sub

Private Function GetTextFlag(Col As Long) As Long
    GetTextFlag = &H4 Or &H20 Or &H40000
    Select Case m_Cols(Col).Aling
        Case 1: GetTextFlag = GetTextFlag Or &H2
        Case 2: GetTextFlag = GetTextFlag Or &H1
    End Select
End Function

Private Function LongToSignedShort(dwUnsigned As Long) As Integer
   If dwUnsigned < 32768 Then
      LongToSignedShort = CInt(dwUnsigned)
   Else
      LongToSignedShort = CInt(dwUnsigned - &H10000)
   End If
End Function


'/Scaroll Bars
Private Property Get sb_max(lSB As Long) As Long
Dim tSI As SCROLLINFO
   pSBGetSI lSB, tSI, &H1 Or &H2 'SIF_RANGE Or SIF_PAGE
   sb_max = tSI.nMax
End Property
Private Property Let sb_max(lSB As Long, ByVal Value As Long)
Dim tSI As SCROLLINFO
   tSI.nMax = Value + sb_large_change(lSB)
   tSI.nMin = 0 'SBMin(eBar)
   pSBLetSI lSB, tSI, &H1 'SIF_RANGE
End Property
Private Property Get sb_value(lSB As Long) As Long
Dim tSI As SCROLLINFO
   pSBGetSI lSB, tSI, &H4 'SIF_POS
   sb_value = tSI.nPos
End Property
Private Property Let sb_value(lSB As Long, Value As Long)
Dim tSI As SCROLLINFO
    If Not Value = sb_value(lSB) Then
        tSI.nPos = Value
        pSBLetSI lSB, tSI, &H4 'SIF_POS
    End If
End Property

Private Property Get sb_large_change(lSB As Long) As Long
Dim tSI As SCROLLINFO
    pSBGetSI lSB, tSI, &H2 'SIF_PAGE
    sb_large_change = tSI.nPage
End Property
Private Property Let sb_large_change(lSB As Long, ByVal Value As Long)
Dim tSI As SCROLLINFO
   pSBGetSI lSB, tSI, SIF_ALL
   tSI.nMax = tSI.nMax - tSI.nPage + Value
   tSI.nPage = Value
   pSBLetSI lSB, tSI, &H2 Or &H1 'SIF_PAGE Or SIF_RANGE
End Property
Private Property Get sb_small_change(lSB As Long) As Long
    sb_small_change = m_smllChange(lSB)
End Property
Private Property Let sb_small_change(lSB As Long, ByVal Value As Long)
    m_smllChange(lSB) = Value
End Property
Private Property Get sb_min(lSB As Long) As Long
Dim tSI As SCROLLINFO
   pSBGetSI lSB, tSI, &H1 'SIF_RANGE
   sb_min = tSI.nMin
End Property
Private Property Let sb_min(lSB As Long, ByVal Value As Long)
Dim tSI As SCROLLINFO
   tSI.nMin = Value
   tSI.nMax = sb_max(lSB) + sb_large_change(lSB)
   pSBLetSI lSB, tSI, &H1 'SIF_RANGE
End Property
Private Sub pSBGetSI(lSB As Long, ByRef tSI As SCROLLINFO, ByVal lMask As Long)
    tSI.fMask = lMask: tSI.cbSize = LenB(tSI)
    GetScrollInfo UserControl.hWnd, lSB, tSI
End Sub
Private Sub pSBLetSI(lSB As Long, ByRef tSI As SCROLLINFO, ByVal lMask As Long)
   tSI.fMask = lMask: tSI.cbSize = LenB(tSI)
   SetScrollInfo UserControl.hWnd, lSB, tSI, True
End Sub

Private Function GetRowFromY(ByVal Y As Long) As Long
    Y = Y + sb_value(SB_VERT) - m_HeaderH
    GetRowFromY = Y \ m_RowH
    If GetRowFromY >= ItemCount Then GetRowFromY = -1
End Function

Private Function BlendColor(ByVal clrFirst As Long, ByVal clrSecond As Long, ByVal lAlpha As Long) As Long
Dim clrFore(3)      As Byte
Dim clrBack(3)      As Byte

    OleTranslateColor clrFirst, 0, VarPtr(clrFore(0))
    OleTranslateColor clrSecond, 0, VarPtr(clrBack(0))
    
    clrFore(0) = (clrFore(0) * lAlpha + clrBack(0) * (255 - lAlpha)) / 255
    clrFore(1) = (clrFore(1) * lAlpha + clrBack(1) * (255 - lAlpha)) / 255
    clrFore(2) = (clrFore(2) * lAlpha + clrBack(2) * (255 - lAlpha)) / 255
    CopyMemory BlendColor, clrFore(0), 4
End Function
Private Function ShiftColor(ByVal vlngColor As Long, ByVal vlngValue As Long) As Long

  '// this function will add or remove a certain Color quantity and return the result
  Dim lngRed   As Long
  Dim lngBlue  As Long
  Dim lngGreen As Long
  Const C_MAX  As Long = &HFF

   lngBlue = ((vlngColor \ &H10000) Mod &H100) + vlngValue
   lngGreen = ((vlngColor \ &H100) Mod &H100) + vlngValue
   lngRed = (vlngColor And &HFF) + vlngValue

   '// values will overflow a byte only in one direction
   '// eg: if we added 32 to our color, then only a > 255 overflow can occurr.
   If vlngValue > 0 Then
      If lngRed > C_MAX Then lngRed = C_MAX
      If lngGreen > C_MAX Then lngGreen = C_MAX
      If lngBlue > C_MAX Then lngBlue = C_MAX

   ElseIf vlngValue < 0 Then
      If lngRed < 0 Then lngRed = 0
      If lngGreen < 0 Then lngGreen = 0
      If lngBlue < 0 Then lngBlue = 0
   End If

   '// more optimization by replacing the RGB function by its correspondent calculation
   ShiftColor = lngRed + 256& * lngGreen + 65536 * lngBlue

End Function
Private Function SysColor(oColor As Long) As Long: OleTranslateColor2 oColor, 0, SysColor: End Function




' ActiveVB
Private Function SetSubclassing(ByVal hWnd As Long) As Boolean

    'Setzt Subclassing, sofern nicht schon gesetzt
    
    If PrevWndProc = 0 Then
        If pASMWrapper <> 0 Then
            PrevWndProc = SetWindowLong(hWnd, GWL_WNDPROC, pASMWrapper)
            If PrevWndProc <> 0 Then
                hSubclassedWnd = hWnd
                SetSubclassing = True
            End If
        End If
    End If

End Function

' ActiveVB
Private Function StopSubclassing() As Boolean

    'Stopt Subclassing, sofern gesetzt

    If hSubclassedWnd <> 0 Then
        If PrevWndProc <> 0 Then
            Call SetWindowLong(hSubclassedWnd, GWL_WNDPROC, PrevWndProc)
            hSubclassedWnd = 0
            PrevWndProc = 0
            StopSubclassing = True
        End If
    End If

End Function









