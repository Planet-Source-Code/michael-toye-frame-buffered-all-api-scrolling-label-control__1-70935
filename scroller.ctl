VERSION 5.00
Begin VB.UserControl ctlScroller 
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1380
   ScaleHeight     =   420
   ScaleWidth      =   1380
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   120
      Top             =   0
   End
End
Attribute VB_Name = "ctlScroller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Boolean, ByVal fdwUnderline As Boolean, ByVal fdwStrikeOut As Boolean, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function StretchDIBits& Lib "gdi32" (ByVal hdc&, ByVal X&, ByVal Y&, ByVal dx&, ByVal dy&, ByVal SrcX&, ByVal SrcY&, ByVal Srcdx&, ByVal Srcdy&, Bits As Any, BInf As Any, ByVal Usage&, ByVal Rop&)
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long

Private Type LOGBRUSH
    lbStyle As Long
    lbColor As Long
    lbHatch As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private mCaption As String
Const LOGPIXELSY = 90
Const COLOR_WINDOW = 5
Const Message = "Hello !"
Const OPAQUE = 2
Const TRANSPARENT = 1
Const FW_DONTCARE = 0
Const FW_THIN = 100
Const FW_EXTRALIGHT = 200
Const FW_LIGHT = 300
Const FW_NORMAL = 400
Const FW_MEDIUM = 500
Const FW_SEMIBOLD = 600
Const FW_BOLD = 700
Const FW_EXTRABOLD = 800
Const FW_HEAVY = 900
Const FW_BLACK = FW_HEAVY
Const FW_DEMIBOLD = FW_SEMIBOLD
Const FW_REGULAR = FW_NORMAL
Const FW_ULTRABOLD = FW_EXTRABOLD
Const FW_ULTRALIGHT = FW_EXTRALIGHT
'used with fdwCharSet
Const ANSI_CHARSET = 0
Const DEFAULT_CHARSET = 1
Const SYMBOL_CHARSET = 2
Const SHIFTJIS_CHARSET = 128
Const HANGEUL_CHARSET = 129
Const CHINESEBIG5_CHARSET = 136
Const OEM_CHARSET = 255
'used with fdwOutputPrecision
Const OUT_CHARACTER_PRECIS = 2
Const OUT_DEFAULT_PRECIS = 0
Const OUT_DEVICE_PRECIS = 5
'used with fdwClipPrecision
Const CLIP_DEFAULT_PRECIS = 0
Const CLIP_CHARACTER_PRECIS = 1
Const CLIP_STROKE_PRECIS = 2
'used with fdwQuality
Const DEFAULT_QUALITY = 0
Const DRAFT_QUALITY = 1
Const PROOF_QUALITY = 2
'used with fdwPitchAndFamily
Const DEFAULT_PITCH = 0
Const FIXED_PITCH = 1
Const VARIABLE_PITCH = 2

Private HiLite&
Private HiLite2&
Private LoLite&
Private Greyed&
Private Shadow&

Private bufferDC&, bufferBM&
Private blankDC&, blankBM&
Private MessageDC&, MessageBM&
Private MessageMaskDC&, MessageMaskBM&
Private mStyle&, bLastbit As Boolean
Private gT&, gL&, gW&, gH&, mW&, mH&, scroll_pos&, mBackColor&
Private mFontSize&, mFontFace$, mScrollDisabled As Boolean
Public Property Let BackColor(n&)
Dim R&, G&, B&, c&

    If n < 0 Then
        OleTranslateColor n, 0, c
        B = (c \ 65536) And &HFF
        G = (c \ 256) And &HFF
        R = c And &HFF
        mBackColor = RGB(R, G, B)
    Else
        mBackColor = n
    End If
PropertyChanged BackColor
GradientToBlank
PaintControl
UserControl.Refresh
End Property
Public Property Get BackColor() As Long
BackColor = mBackColor
End Property
Public Property Let FontFace(s$)
mFontFace = s
End Property
Public Property Get FontFace() As String
PropertyChanged FontFace
FontFace = mFontFace
ResetMessageDC
PaintControl
UserControl.Refresh
End Property
Public Property Let FontSize(n&)
PropertyChanged FontSize
mFontSize = n
ResetMessageDC
PaintControl
UserControl.Refresh
End Property
Public Property Get FontSize() As Long
FontSize = mFontSize
End Property

Public Property Get BackStyle() As Long
    BackStyle = mStyle
    
End Property

Public Property Let BackStyle(StyleNo As Long)
mStyle = StyleNo
If mStyle > 3 Then mStyle = 3
GradientToBlank
PaintControl
PropertyChanged BackStyle
UserControl.Refresh
End Property
Public Property Let Caption(s As String)
mCaption = s
ResetMessageDC
PaintControl

PropertyChanged Caption

End Property
Public Property Get Caption() As String
Caption = mCaption
End Property

Function CreateMyFont(nSize&, sFontFace$, bBold As Boolean, bItalic As Boolean) As Long
    CreateMyFont = CreateFont(-MulDiv(nSize, GetDeviceCaps(GetDC(0), LOGPIXELSY), 72), 0, 0, 0, _
                              IIf(bBold, FW_BOLD, FW_NORMAL), bItalic, False, False, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, _
                              CLIP_DEFAULT_PRECIS, PROOF_QUALITY, DEFAULT_PITCH, sFontFace)
End Function
Private Sub SplitRGB(ByVal clr&, R&, G&, B&)
    R = clr And &HFF: G = (clr \ &H100&) And &HFF: B = (clr \ &H10000) And &HFF
End Sub
Private Sub Gradient(DC&, X&, Y&, dx&, dy&, ByVal c1&, ByVal c2&, v As Boolean)
Dim r1&, G1&, B1&, r2&, G2&, B2&, B() As Byte
Dim i&, lR!, lG!, lB!, dR!, dG!, dB!, BI&(9), xx&, yy&, dd&, hRPen&
    If dx = 0 Or dy = 0 Then Exit Sub
    If v Then xx = 1: yy = dy: dd = dy Else xx = dx: yy = 1: dd = dx
    SplitRGB c1, r1, G1, B1: SplitRGB c2, r2, G2, B2: ReDim B(dd * 4 - 1)
    dR = (r2 - r1) / (dd - 1): lR = r1: dG = (G2 - G1) / (dd - 1): lG = G1: dB = (B2 - B1) / (dd - 1): lB = B1
    For i = 0 To (dd - 1) * 4 Step 4: B(i + 2) = lR: lR = lR + dR: B(i + 1) = lG: lG = lG + dG: B(i) = lB: lB = lB + dB: Next
    BI(0) = 40: BI(1) = xx: BI(2) = -yy: BI(3) = 2097153: StretchDIBits DC, X, Y, dx, dy, 0, 0, xx, yy, B(0), BI(0), 0, vbSrcCopy
End Sub

Sub CreateDCandBitmap(ByRef DC&, ByRef BM&, w&, h&)
    DC = CreateCompatibleDC(GetDC(0))
    BM = CreateCompatibleBitmap(GetDC(0), w, h)
    SelectObject DC, BM
    SetBkMode DC, TRANSPARENT
End Sub

Private Sub Timer1_Timer()
    PaintControl

    scroll_pos = scroll_pos + IIf(bLastbit, -1, 1)
    If (mW - scroll_pos) <= gW And Not bLastbit Then
        bLastbit = True
    ElseIf scroll_pos <= 0 And bLastbit Then
        Timer1.Enabled = False
        bLastbit = False
        scroll_pos = 0
        PaintControl
    End If
End Sub

Private Sub UserControl_Click()
If Timer1.Enabled = True Then
    mScrollDisabled = True
    Timer1.Enabled = False
    scroll_pos = 0: PaintControl
End If
End Sub

Private Sub UserControl_DblClick()
If mScrollDisabled Then mScrollDisabled = False

End Sub

Private Sub UserControl_Initialize()

    gT = 0: gL = 0: gW = UserControl.Width \ Screen.TwipsPerPixelX: gH = UserControl.Height \ Screen.TwipsPerPixelY

    HiLite = RGB(215, 215, 215)
    HiLite2 = RGB(255, 255, 255)
    LoLite = RGB(165, 165, 165)
    Shadow = RGB(150, 150, 150)
    Greyed = RGB(190, 190, 190)

    CreateDCandBitmap bufferDC, bufferBM, gW, gH
    CreateDCandBitmap blankDC, blankBM, gW, gH
    
    mFontFace = "Tahoma"
    mFontSize = 10
    
    ResetMessageDC
    
    GradientToBlank
    PaintControl
    
    UserControl.Refresh
    
End Sub
Sub ResetMessageDC()
Dim lB As LOGBRUSH, f&, R As RECT
    UserControl.FontName = mFontFace: UserControl.FontSize = mFontSize
    mW = UserControl.TextWidth(mCaption) \ Screen.TwipsPerPixelX: mH = UserControl.TextHeight(mCaption) \ Screen.TwipsPerPixelY
    mW = mW + 2: mH = mH + 2
    CreateDCandBitmap MessageDC, MessageBM, mW, mH
    CreateDCandBitmap MessageMaskDC, MessageMaskBM, mW, mH
    lB.lbColor = vbWhite
    lB.lbStyle = 0
    lB.lbHatch = 1
    f = CreateBrushIndirect(lB)
    SetRect R, 0, 0, mW, mH
    FillRect MessageMaskDC, R, f
    DeleteObject f
    SetFont MessageDC, mFontFace, mFontSize
    SetTextColor MessageDC, 0
    SetFont MessageMaskDC, mFontFace, mFontSize
    SetTextColor MessageMaskDC, 0
    WriteToMessageDC mCaption
    
End Sub
Sub WriteToMessageDC(s$)
    TextOut MessageDC, 0, 0, s, Len(s)
    TextOut MessageMaskDC, 0, 0, s, Len(s)
End Sub
Sub SetFont(DC&, sFace$, nSize&)
    DeleteObject SelectObject(DC, CreateMyFont(nSize, sFace, False, False))
End Sub
Sub GradientToBlank()
    If mStyle = 0 Then
        Gradient blankDC, 0, 0, gW, gH, mBackColor, mBackColor, True
    ElseIf mStyle = 1 Then
        Gradient blankDC, 0, 0, gW, gH, vbWhite, HiLite, True
    ElseIf mStyle = 2 Then
        Gradient blankDC, 0, 0, gW, gH / 2, HiLite, HiLite2, True
        Gradient blankDC, 0, gH / 2, gW, gH, HiLite, LoLite, True
    Else
        Gradient blankDC, 0, 0, gW, gH, Greyed, HiLite, True
    End If
End Sub
Sub BlankToBuffer()
    BitBlt bufferDC, gL, gT, gW, gH, blankDC, 0, 0, vbSrcCopy
End Sub
Sub BufferToScreen()
    BitBlt UserControl.hdc, gL, gT, gW, gH, bufferDC, 0, 0, vbSrcCopy
End Sub
Sub messageToBuffer(o&)
    BitBlt bufferDC, gL, gT, mW, mH, MessageMaskDC, o, 0, vbSrcAnd
    BitBlt bufferDC, gL, gT, mW, mH, MessageDC, o, 0, vbSrcPaint
    
End Sub
Private Sub PaintControl()
    BlankToBuffer
    messageToBuffer scroll_pos
    If scroll_pos = 0 And Not Timer1.Enabled And mW > gW Then
        SetPixel bufferDC, gW - 2, gH - 2, 0
        SetPixel bufferDC, gW - 4, gH - 2, 0
        SetPixel bufferDC, gW - 6, gH - 2, 0
    End If
    BufferToScreen
    
End Sub

Private Sub UserControl_InitProperties()
mCaption = UserControl.Name
mStyle = 1
mFontFace = "Tahoma"
mFontSize = 10
mBackColor = vbWhite
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (mW + 2) <= gW Or (X \ Screen.TwipsPerPixelX) < (gW - 10) Then Exit Sub
If Timer1.Enabled Or mScrollDisabled Then Exit Sub
scroll_pos = 0
Timer1.Enabled = True

End Sub

Private Sub UserControl_Paint()
    PaintControl
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
         mCaption = .ReadProperty("Caption", UserControl.Name)
         mStyle = .ReadProperty("Style", 1)
         mFontFace = .ReadProperty("Face", "Tahoma")
         mFontSize = .ReadProperty("fntSize", 10)
         mBackColor = .ReadProperty("BackColor", vbWhite)
    End With
GradientToBlank
PaintControl
End Sub

Private Sub UserControl_Resize()

    DeleteDC bufferDC: DeleteObject bufferBM
    DeleteDC blankDC: DeleteObject blankBM
    
    gT = 0: gL = 0: gW = UserControl.Width \ Screen.TwipsPerPixelX: gH = UserControl.Height \ Screen.TwipsPerPixelY

    CreateDCandBitmap bufferDC, bufferBM, gW, gH
    CreateDCandBitmap blankDC, blankBM, gW, gH

    GradientToBlank
    PaintControl
    
End Sub

Private Sub UserControl_Show()
    GradientToBlank
    ResetMessageDC
    PaintControl

End Sub

Private Sub UserControl_Terminate()
    DeleteDC bufferDC: DeleteObject bufferBM
    DeleteDC blankDC: DeleteObject blankBM
    DeleteDC MessageDC: DeleteObject MessageBM
    DeleteDC MessageMaskDC: DeleteObject MessageMaskBM
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
         .WriteProperty "Caption", mCaption
         .WriteProperty "Style", mStyle
         .WriteProperty "Face", mFontFace
         .WriteProperty "fntSize", mFontSize
         .WriteProperty "BackColor", mBackColor
    End With
End Sub
