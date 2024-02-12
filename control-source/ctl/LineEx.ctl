VERSION 5.00
Begin VB.UserControl LineEx 
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   ClipBehavior    =   0  'None
   HasDC           =   0   'False
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "LineEx.ctx":0000
   Windowless      =   -1  'True
   Begin VB.Timer tmrPainting 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   480
      Top             =   240
   End
   Begin VB.Timer tmrSetExtenderPosSize 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   960
      Top             =   240
   End
End
Attribute VB_Name = "LineEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IBSSubclass

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type T_MSG
    hWnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As T_MSG, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Const PM_REMOVE = &H1

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function GetClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long) As Long
Private Declare Function GetRgnBox Lib "gdi32" (ByVal hRgn As Long, lpRect As RECT) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

Private Enum FillModeConstants
    FillModeAlternate = &H0
    FillModeWinding = &H1
End Enum

Private Const UnitPixel = 2
Private Const QualityModeLow As Long = 1
Private Const SmoothingModeAntiAlias As Long = &H4

Private Const WM_USER As Long = &H400
Private Const WM_INVALIDATE As Long = WM_USER + 11 ' custom message

Private Type POINTL
    X As Long
    Y As Long
End Type

Private Type GDIPlusStartupInput
    GdiPlusVersion                      As Long
    DebugEventCallback                  As Long
    SuppressBackgroundThread            As Long
    SuppressExternalCodecs              As Long
End Type
 
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, ByRef graphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As Long) As Long
Private Declare Function GdiplusStartup Lib "gdiplus" (ByRef token As Long, ByRef lpInput As GDIPlusStartupInput, Optional ByRef lpOutput As Any) As Long
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mSmoothingMode As Long) As Long
Private Declare Function GdipCreatePen1 Lib "gdiplus" (ByVal Color As Long, ByVal Width As Single, ByVal unit As Long, pen As Long) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long
Private Declare Function GdipSetPenDashStyle Lib "gdiplus" (ByVal pen As Long, ByVal dStyle As Long) As Long
Private Declare Function GdipDrawPolygonI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByRef pPoints As Any, ByVal Count As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "GdiPlus.dll" (ByVal mColor As Long, ByRef mBrush As Long) As Long
Private Declare Function GdipFillPolygonI Lib "gdiplus" (ByVal graphics As Long, ByVal Brush As Long, ByRef pPoints As Any, ByVal Count As Long, ByVal FillMode As Long) As Long
Private Declare Function GdipDeleteBrush Lib "GdiPlus.dll" (ByVal mBrush As Long) As Long
Private Declare Function GdipDrawLineI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function InvalidateRectAsNull Lib "user32" Alias "InvalidateRect" (ByVal hWnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long
Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, Col As Long) As Long

Public Enum LEStyleConstants
    leStyleNormal
    leStyleArrow
    leStyleDoubleArrow
    leStyleArrow2
    leStyleArrow3
End Enum

Public Enum LEDirectionConstants
    leRight
    leLeft
    leUp
    leDown
    leRightUp
    leRightDown
    leLeftUp
    leLeftDown
End Enum

Private Enum LESetFromConstants
    leUCPosSize
    leAngle
    leLength
    leXY
End Enum

Private Const Pi = 3.14159265358979

Private Const mdef_BorderColor = vbWindowText
Private Const mdef_BorderStyle = vbBSSolid
Private Const mdef_BorderWidth = 1
Private Const mdef_Quality = seQualityHigh
Private Const mdef_Opacity = 100
Private Const mdef_Direction = leRightDown
Private Const mdef_Style = leStyleNormal
Private Const mdef_ArrowLength = 15
Private Const mdef_ArrowThickness = 5

Private mX1 As Single
Private mX2 As Single
Private mY1 As Single
Private mY2 As Single
Private mLength As Single
Private mAngle As Single
Private mDirection As LEDirectionConstants

Private mOpacity As Single
Private mBorderColor As Long
Private mBorderStyle  As BorderStyleConstants
Private mBorderWidth  As Integer
Private mQuality As SEQualityConstants
Private mStyle As LEStyleConstants
Private mArrowLength As Long
Private mArrowThickness As Long

Private mGdipToken As Long
Private mContainerHwnd As Long
Private mUserMode As Boolean
Private mSetFrom As LESetFromConstants
Private mChangingUCPosSize As Boolean
Private mLastExtenderLeft As Single
Private mLastExtenderTop As Single
Private mLastExtenderWidth As Single
Private mLastExtenderHeight As Single
Private mContainerScaleMode As ScaleModeConstants
Private mDrawingOutsideUC As Boolean
Private mInvalidateMsgPosted As Boolean
Private mSubclassed As Boolean
Private mSetDesignTimeDirection As Boolean
Private mChangingPosSize As Boolean

Private Function IBSSubclass_MsgResponse(ByVal hWnd As Long, ByVal iMsg As Long) As Long
    IBSSubclass_MsgResponse = emrPreprocess
End Function

Private Sub IBSSubclass_UnsubclassIt()
    Unsubclass
End Sub

Private Function IBSSubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, wParam As Long, lParam As Long, bConsume As Boolean) As Long
    Select Case iMsg
        Case WM_INVALIDATE
            Dim iMessage As T_MSG
            
            PeekMessage iMessage, hWnd, WM_INVALIDATE, WM_INVALIDATE, PM_REMOVE  ' remove posted message, if any
            mInvalidateMsgPosted = False
            InvalidateRectAsNull hWnd, 0&, 1&
    End Select
End Function

Private Sub tmrPainting_Timer()
    tmrPainting.Enabled = False
End Sub

Private Sub tmrSetExtenderPosSize_Timer()
    tmrSetExtenderPosSize.Enabled = False
    SetExtenderPosSize
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    If PropertyName = "ScaleUnits" Then
        mContainerScaleMode = ControlFromParent.Container.ScaleMode
        PropertyChanged "X1"
        PropertyChanged "Y1"
        PropertyChanged "X2"
        PropertyChanged "Y2"
        PropertyChanged "Length"
    End If
End Sub

Private Sub UserControl_InitProperties()
    mBorderColor = mdef_BorderColor
    mBorderStyle = mdef_BorderStyle
    mBorderWidth = mdef_BorderWidth
    mQuality = mdef_Quality
    mOpacity = mdef_Opacity
    mDirection = mdef_Direction
    mSetFrom = leUCPosSize
    mStyle = mdef_Style
    mArrowLength = mdef_ArrowLength
    mArrowThickness = mdef_ArrowThickness
    
    mContainerScaleMode = ControlFromParent.Container.ScaleMode
    
    mLastExtenderLeft = UserControl.Extender.Left
    mLastExtenderTop = UserControl.Extender.Top
    mLastExtenderWidth = UserControl.Extender.Width
    mLastExtenderHeight = UserControl.Extender.Height
    
    On Error Resume Next
    mContainerHwnd = UserControl.ContainerHwnd
    mUserMode = Ambient.UserMode
    On Error GoTo 0
    If (Not mUserMode) Then
        mStyle = Val(GetSetting(App.Title, TypeName(Me), "DefStyle", mdef_Style))
        mSetDesignTimeDirection = True
    End If
    Subclass
End Sub

Private Sub SetDesignTimeDirection()
    Dim iPt As POINTAPI
    Dim iMouseMovedX As Long
    Dim iMouseMovedY As Long
    Dim iParenthWnd As Long
    Const cTolerancePos As Long = 20
    Const cToleranceSize As Long = 20
    
    mSetDesignTimeDirection = False
    On Error Resume Next
    iParenthWnd = UserControl.Parent.hWnd
    On Error GoTo 0
    If iParenthWnd = 0 Then Exit Sub
    
    GetCursorPos iPt
    ScreenToClient iParenthWnd, iPt
    iPt.X = iPt.X - UserControl.ScaleX(GetUserControlAbsoluteLeftTwips, vbTwips, vbPixels)
    iPt.Y = iPt.Y - UserControl.ScaleX(GetUserControlAbsoluteTopTwips, vbTwips, vbPixels)
    If Abs(iPt.X - UserControl.ScaleWidth) < cTolerancePos Then
        ' mouse moved left-to-right
        iMouseMovedX = vbAlignRight
    ElseIf Abs(iPt.X) < cTolerancePos Then
        ' mouse moved right-to-left
        iMouseMovedX = vbAlignLeft
    End If
    If Abs(iPt.Y) < cTolerancePos Then
        ' mouse moved bottom-to-top
        iMouseMovedY = vbAlignTop
    ElseIf Abs(iPt.Y - UserControl.ScaleHeight) < cTolerancePos Then
        ' mouse moved top-to-bottom
        iMouseMovedY = vbAlignBottom
    End If
    
    If (iMouseMovedX = vbAlignRight) And (iMouseMovedY = vbAlignTop) Then
        If UserControl.ScaleHeight < cToleranceSize Then
            mDirection = leRight
        ElseIf UserControl.ScaleWidth < cToleranceSize Then
            mDirection = leUp
        Else
            mDirection = leRightUp
        End If
    ElseIf (iMouseMovedX = vbAlignRight) And (iMouseMovedY = vbAlignBottom) Then
        If UserControl.ScaleHeight < cToleranceSize Then
            mDirection = leRight
        ElseIf UserControl.ScaleWidth < cToleranceSize Then
            mDirection = leDown
        Else
            mDirection = leRightDown
        End If
    ElseIf (iMouseMovedX = vbAlignLeft) And (iMouseMovedY = vbAlignTop) Then
        If UserControl.ScaleHeight < cTolerancePos Then
            mDirection = leLeft
        ElseIf UserControl.ScaleWidth < cToleranceSize Then
            mDirection = leUp
        Else
            mDirection = leLeftUp
        End If
    ElseIf (iMouseMovedX = vbAlignLeft) And (iMouseMovedY = vbAlignBottom) Then
        If UserControl.ScaleHeight < cTolerancePos Then
            mDirection = leRight
        ElseIf UserControl.ScaleWidth < cToleranceSize Then
            mDirection = leDown
        Else
            mDirection = leLeftDown
        End If
    End If

End Sub

Private Function GetUserControlAbsoluteLeftTwips() As Long
    Dim iScaleMode As Long
    Dim iContainer As Object
    Dim iNextContainer As Object
    Dim iParent As Object
    
    iScaleMode = vbTwips
    On Error Resume Next
    iScaleMode = UserControl.Extender.Container.ScaleMode
    Set iContainer = UserControl.Extender.Container
    Set iParent = UserControl.Parent
    On Error GoTo 0
    GetUserControlAbsoluteLeftTwips = ScaleX(UserControl.Extender.Left, iScaleMode, vbTwips)
    
    Do Until iContainer Is iParent
        Set iNextContainer = Nothing
        iScaleMode = vbTwips
        On Error Resume Next
        Set iNextContainer = iContainer.Container
        iScaleMode = iNextContainer.ScaleMode
        On Error GoTo 0
        If iNextContainer Is Nothing Then Exit Do
        GetUserControlAbsoluteLeftTwips = GetUserControlAbsoluteLeftTwips + ScaleX(iContainer.Left, iScaleMode, vbTwips)
        Set iContainer = iNextContainer
    Loop
End Function


Private Function GetUserControlAbsoluteTopTwips() As Long
    Dim iScaleMode As Long
    Dim iContainer As Object
    Dim iNextContainer As Object
    Dim iParent As Object
    
    iScaleMode = vbTwips
    On Error Resume Next
    iScaleMode = UserControl.Extender.Container.ScaleMode
    Set iContainer = UserControl.Extender.Container
    Set iParent = UserControl.Parent
    On Error GoTo 0
    GetUserControlAbsoluteTopTwips = ScaleY(UserControl.Extender.Top, iScaleMode, vbTwips)
    
    Do Until iContainer Is iParent
        Set iNextContainer = Nothing
        iScaleMode = vbTwips
        On Error Resume Next
        Set iNextContainer = iContainer.Container
        iScaleMode = iNextContainer.ScaleMode
        On Error GoTo 0
        If iNextContainer Is Nothing Then Exit Do
        GetUserControlAbsoluteTopTwips = GetUserControlAbsoluteTopTwips + ScaleY(iContainer.Top, iScaleMode, vbTwips)
        Set iContainer = iNextContainer
    Loop
End Function

Private Sub UserControl_Paint()
    Dim hRgnExpand As Long
    Dim hRgn As Long
    Dim rgnRect As RECT
    Dim iAuxExpand As Long
    Static sTheLastTimeWasExpanded As Boolean
    Static sLastPosWas00 As Boolean
    Dim iMsgPosted As Boolean
    
    If Not mChangingUCPosSize Then
        If (mLastExtenderLeft <> UserControl.Extender.Left) Or (mLastExtenderTop <> UserControl.Extender.Top) Then
            PosSizeChange
            sLastPosWas00 = True
        ElseIf (mLastExtenderLeft = 0) And (mLastExtenderTop = 0) Then
            If Not sLastPosWas00 Then
                PosSizeChange
                sLastPosWas00 = True
            End If
        Else
            sLastPosWas00 = True
        End If
    Else
        sLastPosWas00 = True
    End If
    
    If mBorderWidth > 1 Then
        iAuxExpand = Round(mBorderWidth / 2 + 0.005)
    End If
    If mStyle <> leStyleNormal Then
        If (mArrowThickness) > iAuxExpand Then
            iAuxExpand = mArrowThickness
        End If
    End If
    
    mDrawingOutsideUC = False
    hRgn = CreateRectRgn(0, 0, 0, 0)
    If GetClipRgn(UserControl.hDC, hRgn) = 0& Then  ' hDc is one passed to Paint
        DeleteObject hRgn: hRgn = 0
    Else
        GetRgnBox hRgn, rgnRect             ' get its bounds & adjust our region accordingly (i.e.,expand 1 pixel)
        If (iAuxExpand > 0) Or sTheLastTimeWasExpanded Then
            rgnRect.Left = rgnRect.Left - iAuxExpand
            rgnRect.Top = rgnRect.Top - iAuxExpand
            rgnRect.Right = rgnRect.Right + iAuxExpand
            rgnRect.Bottom = rgnRect.Bottom + iAuxExpand
            hRgnExpand = CreateRectRgn(rgnRect.Left, rgnRect.Top, rgnRect.Right, rgnRect.Bottom)
        
            SelectClipRgn UserControl.hDC, hRgnExpand
            DeleteObject hRgnExpand
            mDrawingOutsideUC = True
            If Not tmrPainting.Enabled Then
                PostInvalidateMsg
                iMsgPosted = True
            End If
        End If
    End If
    sTheLastTimeWasExpanded = (iAuxExpand <> 0)
    tmrPainting.Enabled = False
    tmrPainting.Enabled = True
    If Not iMsgPosted Then Draw
    
    If hRgnExpand <> 0 Then SelectClipRgn UserControl.hDC, hRgn  ' restore original clip region
    If hRgn <> 0 Then DeleteObject hRgn
End Sub

Private Sub PostInvalidateMsg()
    Static sTime As Double
    
    If mDrawingOutsideUC Then
        If (mContainerHwnd <> 0) And mSubclassed Then
            If (Not mInvalidateMsgPosted) Or ((Timer - sTime) > 1) Or (Timer < sTime) Then
                PostMessage mContainerHwnd, WM_INVALIDATE, 0&, 0&
                mInvalidateMsgPosted = True
                sTime = Timer
            End If
        End If
    End If
End Sub

Private Sub Draw()
    Dim iGraphics As Long
    Dim hPen As Long
    Dim iPointsL(1) As POINTL
    Dim iPointsA() As POINTL
    Dim iLeft As Single
    Dim iTop As Single
    Dim iSlope As Single
    
    If mBorderStyle = vbTransparent Then Exit Sub
    'Debug.Print "Draw " & Ambient.DisplayName
    
    If mGdipToken = 0 Then InitGDI
    If GdipCreateFromHDC(UserControl.hDC, iGraphics) = 0 Then
        
        On Error GoTo Err_Exit
        iLeft = ControlFromParent.Container.ScaleX(UserControl.Extender.Left, mContainerScaleMode, vbPixels)
        iTop = ControlFromParent.Container.ScaleY(UserControl.Extender.Top, mContainerScaleMode, vbPixels)
        On Error GoTo 0
        
        iPointsL(0).X = UserControl.ScaleX(mX1, vbHimetric, vbPixels) - iLeft
        iPointsL(0).Y = UserControl.ScaleY(mY1, vbHimetric, vbPixels) - iTop
        iPointsL(1).X = UserControl.ScaleX(mX2, vbHimetric, vbPixels) - iLeft
        iPointsL(1).Y = UserControl.ScaleY(mY2, vbHimetric, vbPixels) - iTop
        
        If GdipCreatePen1(ConvertColor(mBorderColor, mOpacity), mBorderWidth, UnitPixel, hPen) = 0 Then
            If ((mBorderStyle = vbBSSolid) Or (mBorderStyle = vbBSInsideSolid)) Or (mBorderWidth > 1) Then
                Call GdipSetSmoothingMode(iGraphics, SmoothingMode)
            Else
                Call GdipSetSmoothingMode(iGraphics, QualityModeLow)
            End If
            If ((mBorderStyle > vbBSSolid) And (mBorderStyle < vbBSInsideSolid)) Then
                Call GdipSetPenDashStyle(hPen, mBorderStyle - 1)
            End If
            GdipDrawLineI iGraphics, hPen, iPointsL(0).X, iPointsL(0).Y, iPointsL(1).X, iPointsL(1).Y
            Call GdipDeletePen(hPen)
            
            ReDim iPointsA(2)
            If (mStyle <> leStyleNormal) And (mArrowLength > 0) And (mArrowThickness > 0) Then
                If mBorderWidth = 1 Then
                    iPointsA(1).X = iPointsL(1).X
                    iPointsA(1).Y = iPointsL(1).Y
                Else
                    iPointsA(1) = GetArrowPointCenter(iPointsL(1).X, iPointsL(1).Y, iPointsL(0).X, iPointsL(0).Y, -mBorderWidth)
                End If
                iPointsA(0) = GetArrowPointRight(iPointsL(1).X, iPointsL(1).Y, iPointsL(0).X, iPointsL(0).Y, mArrowThickness, mArrowLength)
                iPointsA(2) = GetArrowPointLeft(iPointsL(1).X, iPointsL(1).Y, iPointsL(0).X, iPointsL(0).Y, mArrowThickness, mArrowLength)
                FillPolygon iGraphics, iPointsA
                    
                If mStyle = leStyleDoubleArrow Then
                    If mBorderWidth = 1 Then
                        iPointsA(1).X = iPointsL(0).X
                        iPointsA(1).Y = iPointsL(0).Y
                    Else
                        iPointsA(1) = GetArrowPointCenter(iPointsL(0).X, iPointsL(0).Y, iPointsL(1).X, iPointsL(1).Y, -mBorderWidth)
                    End If
                    iPointsA(0) = GetArrowPointRight(iPointsL(0).X, iPointsL(0).Y, iPointsL(1).X, iPointsL(1).Y, mArrowThickness, mArrowLength)
                    iPointsA(2) = GetArrowPointLeft(iPointsL(0).X, iPointsL(0).Y, iPointsL(1).X, iPointsL(1).Y, mArrowThickness, mArrowLength)
                    FillPolygon iGraphics, iPointsA
                ElseIf mStyle = leStyleArrow2 Then
                    ReDim iPointsA(3)
                    iPointsA(0) = GetArrowPointRight(iPointsL(0).X, iPointsL(0).Y, iPointsL(1).X, iPointsL(1).Y, mArrowThickness, mArrowLength / 4)
                    iPointsA(1) = GetArrowPointLeft(iPointsL(0).X, iPointsL(0).Y, iPointsL(1).X, iPointsL(1).Y, mArrowThickness, mArrowLength / 4)
                    iPointsA(2) = GetArrowPointLeft(iPointsL(0).X, iPointsL(0).Y, iPointsL(1).X, iPointsL(1).Y, mArrowThickness, 0)
                    iPointsA(3) = GetArrowPointRight(iPointsL(0).X, iPointsL(0).Y, iPointsL(1).X, iPointsL(1).Y, mArrowThickness, 0)
                    FillPolygon iGraphics, iPointsA
                ElseIf mStyle = leStyleArrow3 Then
                    ReDim iPointsA(5)
                    iPointsA(0) = GetArrowPointCenter(iPointsL(0).X, iPointsL(0).Y, iPointsL(1).X, iPointsL(1).Y, mArrowLength / 4)
                    iPointsA(1) = GetArrowPointLeft(iPointsL(0).X, iPointsL(0).Y, iPointsL(1).X, iPointsL(1).Y, mArrowThickness, 0)
                    iPointsA(2) = GetArrowPointLeft(iPointsL(0).X, iPointsL(0).Y, iPointsL(1).X, iPointsL(1).Y, mArrowThickness, mArrowLength)
                    iPointsA(3) = GetArrowPointCenter(iPointsL(0).X, iPointsL(0).Y, iPointsL(1).X, iPointsL(1).Y, mArrowLength * 1.25)
                    iPointsA(4) = GetArrowPointRight(iPointsL(0).X, iPointsL(0).Y, iPointsL(1).X, iPointsL(1).Y, mArrowThickness, mArrowLength)
                    iPointsA(5) = GetArrowPointRight(iPointsL(0).X, iPointsL(0).Y, iPointsL(1).X, iPointsL(1).Y, mArrowThickness, 0)
                    FillPolygon iGraphics, iPointsA
                End If
            End If
        End If
        
Err_Exit:
        Call GdipDeleteGraphics(iGraphics)
    End If
End Sub

Private Function GetArrowPointRight(ByVal mX1 As Double, ByVal mY1 As Double, ByVal mX2 As Double, ByVal mY2 As Double, ByVal D1 As Double, ByVal D2 As Double) As POINTL
    Dim dist As Double
    Dim X As Double
    Dim Y As Double
    
    dist = Sqr((mX2 - mX1) ^ 2 + (mY2 - mY1) ^ 2)
    
    If dist <> 0 Then
        X = (mX1 * (dist - D2) + mX2 * D2) / dist
        Y = (mY1 * (dist - D2) + mY2 * D2) / dist
        GetArrowPointRight.X = X + D1 * (mY2 - mY1) / dist '+ 0.005
        GetArrowPointRight.Y = Y - D1 * (mX2 - mX1) / dist '- 0.005
    End If
End Function

Private Function GetArrowPointLeft(ByVal mX1 As Double, ByVal mY1 As Double, ByVal mX2 As Double, ByVal mY2 As Double, ByVal D1 As Double, ByVal D2 As Double) As POINTL
    Dim dist As Double
    Dim X As Double
    Dim Y As Double
    
    dist = Sqr((mX2 - mX1) ^ 2 + (mY2 - mY1) ^ 2)
    If dist <> 0 Then
        X = (mX1 * (dist - D2) + mX2 * D2) / dist
        Y = (mY1 * (dist - D2) + mY2 * D2) / dist
        GetArrowPointLeft.X = X - D1 * (mY2 - mY1) / dist '- 0.005
        GetArrowPointLeft.Y = Y + D1 * (mX2 - mX1) / dist '+ 0.005
    End If
End Function

Private Function GetArrowPointCenter(ByVal mX1 As Double, ByVal mY1 As Double, ByVal mX2 As Double, ByVal mY2 As Double, ByVal D2 As Double) As POINTL
    Dim dist As Double
    Dim X As Double
    Dim Y As Double
    
    dist = Sqr((mX2 - mX1) ^ 2 + (mY2 - mY1) ^ 2)
    If dist <> 0 Then
        X = (mX1 * (dist - D2) + mX2 * D2) / dist
        Y = (mY1 * (dist - D2) + mY2 * D2) / dist
        GetArrowPointCenter.X = X
        GetArrowPointCenter.Y = Y
    End If
End Function

Private Sub InitGDI()
    Dim GdipStartupInput As GDIPlusStartupInput
    
    GdipStartupInput.GdiPlusVersion = 1&
    Call GdiplusStartup(mGdipToken, GdipStartupInput, ByVal 0)
End Sub

Private Sub TerminateGDI()
    Call GdiplusShutdown(mGdipToken)
    mGdipToken = 0
End Sub

Private Function ConvertColor(nColor As Long, nOpacity As Single) As Long
    Dim BGRA(0 To 3) As Byte
    Dim iColor As Long
    
    TranslateColor nColor, 0&, iColor
    
    BGRA(3) = CByte((nOpacity / 100) * 255)
    BGRA(0) = ((iColor \ &H10000) And &HFF)
    BGRA(1) = ((iColor \ &H100) And &HFF)
    BGRA(2) = (iColor And &HFF)
    CopyMemory ConvertColor, BGRA(0), 4&
End Function

Private Property Get SmoothingMode() As Long
    If mQuality = seQualityHigh Then
        SmoothingMode = SmoothingModeAntiAlias
    Else
        SmoothingMode = QualityModeLow
    End If
End Property

Private Sub FillPolygon(ByVal nGraphics As Long, Points() As POINTL, Optional nFillMode As FillModeConstants = FillModeAlternate)
    Dim hBrush As Long
    Dim iRet As Long
    
    iRet = GdipCreateSolidFill(ConvertColor(mBorderColor, mOpacity), hBrush)
    If iRet = 0 Then
        Call GdipSetSmoothingMode(nGraphics, SmoothingMode)
        GdipFillPolygonI nGraphics, hBrush, Points(0), UBound(Points) + 1, nFillMode
        Call GdipDeleteBrush(hBrush)
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mBorderColor = PropBag.ReadProperty("BorderColor", mdef_BorderColor)
    mBorderStyle = PropBag.ReadProperty("BorderStyle", mdef_BorderStyle)
    mBorderWidth = PropBag.ReadProperty("BorderWidth", mdef_BorderWidth)
    mQuality = PropBag.ReadProperty("Quality", mdef_Quality)
    mOpacity = PropBag.ReadProperty("Opacity", mdef_Opacity)
    mDirection = PropBag.ReadProperty("Direction", mdef_Direction)
    mStyle = PropBag.ReadProperty("Style", mdef_Style)
    mArrowLength = PropBag.ReadProperty("ArrowLength", mdef_ArrowLength)
    mArrowThickness = PropBag.ReadProperty("ArrowThickness", mdef_ArrowThickness)
    
    mContainerScaleMode = ControlFromParent.Container.ScaleMode
    
    If (PropBag.ReadProperty("X1", 0) <> 0) Or (PropBag.ReadProperty("Y1", 0) <> 0) Or (PropBag.ReadProperty("X2", 0) <> 0) Or (PropBag.ReadProperty("Y2", 0) <> 0) Then
        mX1 = ControlFromParent.Container.ScaleX(PropBag.ReadProperty("X1", 0), ControlFromParent.Container.ScaleMode, vbHimetric)
        mY1 = ControlFromParent.Container.ScaleY(PropBag.ReadProperty("Y1", 0), ControlFromParent.Container.ScaleMode, vbHimetric)
        mX2 = ControlFromParent.Container.ScaleX(PropBag.ReadProperty("X2", 0), ControlFromParent.Container.ScaleMode, vbHimetric)
        mY2 = ControlFromParent.Container.ScaleY(PropBag.ReadProperty("Y2", 0), ControlFromParent.Container.ScaleMode, vbHimetric)
        mSetFrom = leXY
        SetAngleAndLengthFromXY
        SetExtenderPosSize
        tmrSetExtenderPosSize.Enabled = True
    Else
        mX1 = PropBag.ReadProperty("mX1", 0)
        mY1 = PropBag.ReadProperty("mY1", 0)
        mX2 = PropBag.ReadProperty("mX2", 0)
        mY2 = PropBag.ReadProperty("mY2", 0)
        mLength = PropBag.ReadProperty("Length", 0)
        mAngle = PropBag.ReadProperty("Angle", 0)
        mDirection = PropBag.ReadProperty("Direction", leRightDown)
        If (mX1 = 0) And (mY1 = 0) And (mX2 = 0) And (mY2 = 0) Then
            mSetFrom = leUCPosSize
            mDirection = leRightDown
        Else
            mSetFrom = PropBag.ReadProperty("SetFrom", leXY)
        End If
    End If
    
    If mSetFrom = leAngle Then
        mAngle = mAngle - 1
        Angle = mAngle + 1
    ElseIf mSetFrom = leLength Then
        mLength = mLength - 1
        Length = mLength + 1
    ElseIf mSetFrom = leXY Then
        SetAngleAndLengthFromXY
   ElseIf mSetFrom = leUCPosSize Then
        PosSizeChange
        SetAngleAndLengthFromXY
    End If
    
    mLastExtenderLeft = UserControl.Extender.Left
    mLastExtenderTop = UserControl.Extender.Top
    mLastExtenderWidth = UserControl.Extender.Width
    mLastExtenderHeight = UserControl.Extender.Height
    
    On Error Resume Next
    mContainerHwnd = UserControl.ContainerHwnd
    mUserMode = Ambient.UserMode
    On Error GoTo 0
    Subclass
End Sub

Private Sub UserControl_Resize()
    If mSetDesignTimeDirection Then
        SetDesignTimeDirection
    End If
    If Not mChangingUCPosSize Then
        If (UserControl.Extender.Width <> mLastExtenderWidth) Or (UserControl.Extender.Height <> mLastExtenderHeight) Then
            PosSizeChange
            SetAngleAndLengthFromXY
            mLastExtenderLeft = UserControl.Extender.Left
            mLastExtenderTop = UserControl.Extender.Top
            mLastExtenderWidth = UserControl.Extender.Width
            mLastExtenderHeight = UserControl.Extender.Height
        End If
    End If
End Sub

Private Sub PosSizeChange()
    Dim iLeft As Single
    Dim iTop As Single
    Dim iWidth As Single
    Dim iHeight As Single
    
    If mChangingPosSize Then Exit Sub
    mChangingPosSize = True
    
    iLeft = UserControl.Extender.Left
    iTop = UserControl.Extender.Top
    iWidth = UserControl.Extender.Width
    iHeight = UserControl.Extender.Height
    
    'X1 , Y1, X2, Y2
    Select Case mDirection
        Case leUp, leLeftUp, leRightUp
            mY1 = iTop + iHeight
            mY2 = iTop
        Case leRight, leLeft
            mY1 = iTop
            mY2 = iTop
            UserControl.Extender.Height = ControlFromParent.Container.ScaleY(mBorderWidth, vbPixels, ControlFromParent.Container.ScaleMode)
        Case Else
            mY1 = iTop
            mY2 = iTop + iHeight
    End Select
    Select Case mDirection
        Case leLeft, leLeftDown, leLeftUp
            mX1 = iLeft + iWidth
            mX2 = iLeft
        Case leUp, leDown
            mX1 = iLeft
            mX2 = iLeft
            UserControl.Extender.Width = ControlFromParent.Container.ScaleX(mBorderWidth, vbPixels, ControlFromParent.Container.ScaleMode)
        Case Else
            mX1 = iLeft
            mX2 = iLeft + iWidth
    End Select
    
    mX1 = ControlFromParent.Container.ScaleX(mX1, mContainerScaleMode, vbHimetric)
    mY1 = ControlFromParent.Container.ScaleX(mY1, mContainerScaleMode, vbHimetric)
    mX2 = ControlFromParent.Container.ScaleX(mX2, mContainerScaleMode, vbHimetric)
    mY2 = ControlFromParent.Container.ScaleX(mY2, mContainerScaleMode, vbHimetric)
    
    PropertyChanged "X1"
    PropertyChanged "Y1"
    PropertyChanged "X2"
    PropertyChanged "Y2"
    PropertyChanged "Length"
    PropertyChanged "Angle"
    PropertyChanged "Direction"
    
    mSetFrom = leUCPosSize
    mLastExtenderLeft = UserControl.Extender.Left
    mLastExtenderTop = UserControl.Extender.Top
    Me.Refresh
    mChangingPosSize = False
End Sub

Private Sub SetAngleAndLengthFromXY()
    Dim iSlope As Single
    
    If mChangingPosSize Then Exit Sub
    
    ' Length
    mLength = Sqr((mX1 - mX2) ^ 2 + (mY1 - mY2) ^ 2)
    
    ' Angle
    If (mX2 - mX1) = 0 Then
        If mY2 < mY1 Then
            mAngle = 270
            mDirection = leUp
        Else
            mAngle = 90
            mDirection = leDown
        End If
    Else
        iSlope = (mY2 - mY1) / (mX2 - mX1)
        mAngle = 180 * Atn(iSlope) / Pi
        If mX1 > mX2 Then
            mAngle = mAngle + 180
        ElseIf mY2 < mY1 Then
            mAngle = mAngle + 360
        End If
        If (mY2 - mY1) = 0 Then
            If mX2 < mX1 Then
                mDirection = leLeft
            Else
                mDirection = leRight
            End If
        Else
            If mAngle < 90 Then
                mDirection = leRightDown
            ElseIf mAngle < 180 Then
                mDirection = leLeftDown
            ElseIf mAngle < 270 Then
                mDirection = leLeftUp
            Else
                mDirection = leRightUp
            End If
        End If
    End If
End Sub

Private Sub UserControl_Terminate()
    Unsubclass
    If mGdipToken <> 0 Then TerminateGDI
    If (mBorderWidth > 1) Or mDrawingOutsideUC Then InvalidateRectAsNull mContainerHwnd, 0&, 1&  ' paint the container when the control is deleted if the BorderWidth is greater than 1
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "BorderColor", mBorderColor, mdef_BorderColor
    PropBag.WriteProperty "BorderStyle", mBorderStyle, mdef_BorderStyle
    PropBag.WriteProperty "BorderWidth", mBorderWidth, mdef_BorderWidth
    PropBag.WriteProperty "Quality", mQuality, mdef_Quality
    PropBag.WriteProperty "Opacity", mOpacity, mdef_Opacity
    PropBag.WriteProperty "Direction", mDirection, mdef_Direction
    PropBag.WriteProperty "Style", mStyle, mdef_Style
    PropBag.WriteProperty "ArrowLength", mArrowLength, mdef_ArrowLength
    PropBag.WriteProperty "ArrowThickness", mArrowThickness, mdef_ArrowThickness
    
    PropBag.WriteProperty "mX1", mX1, 0
    PropBag.WriteProperty "mY1", mY1, 0
    PropBag.WriteProperty "mX2", mX2, 0
    PropBag.WriteProperty "mY2", mY2, 0
    PropBag.WriteProperty "Length", mLength, 0
    PropBag.WriteProperty "Angle", mAngle, 0
    PropBag.WriteProperty "Direction", mDirection, leRightDown
    PropBag.WriteProperty "SetFrom", mSetFrom, leXY
End Sub


Public Property Get BorderColor() As OLE_COLOR
    BorderColor = mBorderColor
End Property

Public Property Let BorderColor(ByVal nValue As OLE_COLOR)
    If nValue <> mBorderColor Then
        If Not IsValidOLE_COLOR(nValue) Then Err.Raise 380: Exit Property
        mBorderColor = nValue
        PostInvalidateMsg
        Me.Refresh
        PropertyChanged "BorderColor"
    End If
End Property


Public Property Get BorderStyle() As BorderStyleConstants
    BorderStyle = mBorderStyle
End Property

Public Property Let BorderStyle(ByVal nValue As BorderStyleConstants)
    If nValue <> mBorderStyle Then
        If (nValue < vbTransparent) Or (nValue > vbBSInsideSolid) Then Err.Raise 380, TypeName(Me): Exit Property
        mBorderStyle = nValue
        PostInvalidateMsg
        Me.Refresh
        PropertyChanged "BorderStyle"
    End If
End Property


Public Property Get BorderWidth() As Long
    BorderWidth = mBorderWidth
End Property

Public Property Let BorderWidth(ByVal nValue As Long)
    If nValue < 1 Then
        nValue = 1
    End If
    If nValue <> mBorderWidth Then
        PostInvalidateMsg
        mBorderWidth = nValue
        Me.Refresh
        PropertyChanged "BorderWidth"
    End If
End Property


Public Property Get Quality() As SEQualityConstants
    Quality = mQuality
End Property

Public Property Let Quality(ByVal nValue As SEQualityConstants)
    If nValue <> mQuality Then
        If (nValue < seQualityLow) Or (nValue > seQualityHigh) Then Err.Raise 380, TypeName(Me): Exit Property
        mQuality = nValue
        PostInvalidateMsg
        Me.Refresh
        PropertyChanged "Quality"
    End If
End Property


Public Property Get Style() As LEStyleConstants
    Style = mStyle
End Property

Public Property Let Style(ByVal nValue As LEStyleConstants)
    If nValue <> mStyle Then
        If (nValue < leStyleNormal) Or (nValue > leStyleArrow3) Then Err.Raise 380, TypeName(Me): Exit Property
        mStyle = nValue
        If (Not mUserMode) Then
            SaveSetting App.Title, TypeName(Me), "DefStyle", mStyle
        End If
        PostInvalidateMsg
        Me.Refresh
        PropertyChanged "Style"
    End If
End Property


Public Property Get ArrowLength() As Long
    ArrowLength = mArrowLength
End Property

Public Property Let ArrowLength(ByVal nValue As Long)
    If nValue <> mArrowLength Then
        If (nValue < 0) Then nValue = 0
        If (nValue > 100) Then nValue = 100
        mArrowLength = nValue
        PostInvalidateMsg
        Me.Refresh
        PropertyChanged "ArrowLength"
    End If
End Property


Public Property Get ArrowThickness() As Long
    ArrowThickness = mArrowThickness
End Property

Public Property Let ArrowThickness(ByVal nValue As Long)
    If nValue <> mArrowThickness Then
        If (nValue < 0) Then nValue = 0
        If (nValue > 100) Then nValue = 100
        mArrowThickness = nValue
        PostInvalidateMsg
        Me.Refresh
        PropertyChanged "ArrowThickness"
    End If
End Property


Public Property Get Opacity() As Single
    Opacity = mOpacity
End Property

Public Property Let Opacity(ByVal nValue As Single)
    If nValue <> mOpacity Then
        If nValue > 100 Then
            nValue = 100
        ElseIf nValue < 0 Then
            nValue = 0
        End If
        If nValue <> mOpacity Then
            mOpacity = nValue
            PostInvalidateMsg
            Me.Refresh
            PropertyChanged "Opacity"
        End If
    End If
End Property


Public Property Get Direction() As LEDirectionConstants
    Direction = mDirection
End Property

Public Property Let Direction(ByVal nValue As LEDirectionConstants)
    Dim iWidth As Single
    Dim iHeight As Single
    Dim iLengthPrev As Single
    
    If nValue <> mDirection Then
        If (nValue < leRight) Or (nValue > leLeftDown) Then Err.Raise 380, TypeName(Me): Exit Property
        
        mDirection = nValue
        
        If nValue = leDown Then
            mX2 = mX1
            mY2 = mY1 + mLength
        ElseIf nValue = leUp Then
            mX2 = mX1
            mY2 = mY1 - mLength
        ElseIf nValue = leRight Then
            mX2 = mX1 + mLength
            mY2 = mY1
        ElseIf nValue = leLeft Then
            mX2 = mX1 - mLength
            mY2 = mY1
        Else
            iWidth = ControlFromParent.Container.ScaleX(UserControl.Extender.Width, mContainerScaleMode, vbHimetric)
            iHeight = ControlFromParent.Container.ScaleX(UserControl.Extender.Height, mContainerScaleMode, vbHimetric)
            
            If nValue = leRightDown Then
                mX2 = mX1 + iWidth
                mY2 = mY1 + iHeight
            ElseIf nValue = leLeftDown Then
                mX2 = mX1 - iWidth
                mY2 = mY1 + iHeight
            ElseIf nValue = leRightUp Then
                mX2 = mX1 + iWidth
                mY2 = mY1 - iHeight
            ElseIf nValue = leLeftUp Then
                mX2 = mX1 - iWidth
                mY2 = mY1 - iHeight
            End If
        End If
        
        iLengthPrev = mLength
        SetAngleAndLengthFromXY
        mLength = iLengthPrev

        SetExtenderPosSize
            
        PostInvalidateMsg
        Me.Refresh
        PropertyChanged "Direction"
        PropertyChanged "Angle"
        PropertyChanged "X2"
        PropertyChanged "Y2"
    End If
End Property


Public Property Get Angle() As Single
    Angle = mAngle
End Property

Public Property Let Angle(ByVal nValue As Single)
    Dim iAngle As Single
    Dim iSlope As Single
    Dim iPrev As Single
    Dim iAdj As Single
    
    If nValue <> mAngle Then
        If nValue > 360 Then
            nValue = nValue Mod 360
        ElseIf nValue < 0 Then
            nValue = nValue Mod 360
            nValue = 360 - nValue
        End If
        If nValue <> mAngle Then
            mAngle = nValue
            iAngle = mAngle
            If (iAngle = 270) Then
                mX2 = mX1
                mDirection = leUp
                mY2 = mY1 - mLength
            ElseIf (iAngle = 90) Then
                mX2 = mX1
                mDirection = leDown
                mY2 = mY1 + mLength
            ElseIf (iAngle = 0) Then
                mY2 = mY1
                mDirection = leRight
                mX2 = mX1 + mLength
            ElseIf (iAngle = 180) Then
                mY2 = mY1
                mDirection = leLeft
                mX2 = mX1 - mLength
            Else
                If iAngle > 270 Then
                    iAngle = iAngle - 360
                ElseIf iAngle > 90 Then
                    iAngle = iAngle - 180
                End If
                iSlope = Tan(iAngle * Pi / 180)
                ' from you.com
                If (mAngle > 90) And (mAngle < 270) Then
                    iPrev = mX1
                    mX1 = mX2 + (mLength * Cos(Atn(iSlope)))
                    iAdj = iPrev - mX1
                    mX1 = mX1 + iAdj
                    mX2 = mX2 + iAdj
                    mY2 = mY1 + (mLength * Sin(Atn(iSlope))) * -1
                Else
                    mX2 = mX1 + (mLength * Cos(Atn(iSlope)))
                    mY2 = mY1 + (mLength * Sin(Atn(iSlope)))
                End If
                ' end from
                
                If mX2 < mX1 Then
                    If mY2 < mY1 Then
                        mDirection = leLeftUp
                    Else
                        mDirection = leLeftDown
                    End If
                Else
                    If mY2 < mY1 Then
                        mDirection = leRightUp
                    Else
                        mDirection = leRightDown
                    End If
                End If
                
                SetExtenderPosSize
            End If
            
            mSetFrom = leAngle
            
            PostInvalidateMsg
            Me.Refresh
            PropertyChanged "Angle"
            PropertyChanged "X1"
            PropertyChanged "Y1"
            PropertyChanged "X2"
            PropertyChanged "Y2"
            PropertyChanged "Direction"
        End If
    End If
End Property


Private Sub SetExtenderPosSize()
    Dim iLeft As Single
    Dim iTop As Single
    Dim iWidth As Single
    Dim iHeight As Single
    
    If x1 < x2 Then
        iLeft = x1
    Else
        iLeft = x2
    End If
    If y1 < y2 Then
        iTop = y1
    Else
        iTop = y2
    End If
    iWidth = Abs(x1 - x2)
    iHeight = Abs(y1 - y2)
    
    mChangingUCPosSize = True
    UserControl.Extender.Left = iLeft
    UserControl.Extender.Top = iTop
    UserControl.Extender.Width = iWidth
    UserControl.Extender.Height = iHeight
    mChangingUCPosSize = False
End Sub

Public Property Get Length() As Single
    Length = ControlFromParent.Container.ScaleX(mLength, vbHimetric, mContainerScaleMode)
End Property

Private Property Get ControlFromParent() As Object
    Dim iName As String
    Dim iIndex As Integer
    Dim iPos As Long
    
    iName = Ambient.DisplayName
    iPos = InStr(iName, "(")
    If iPos > 0 Then
        iIndex = Val(Mid$(iName, iPos + 1))
        Set ControlFromParent = UserControl.Parent.Controls(Left$(iName, iPos - 1), iIndex)
    Else
        Set ControlFromParent = UserControl.Parent.Controls(iName)
    End If
End Property

Public Property Let Length(ByVal nValue As Single)
    Dim iAngle As Single
    Dim iSlope As Single
    Dim iPrev As Single
    Dim iAdj As Single
    
    If nValue < 0 Then nValue = 0
    If nValue <> mLength Then
        mLength = ControlFromParent.Container.ScaleY(nValue, mContainerScaleMode, vbHimetric)
        
        iAngle = mAngle
        If iAngle > 270 Then
            iAngle = iAngle - 360
        ElseIf iAngle > 90 Then
            iAngle = iAngle - 180
        End If
        iSlope = Tan(iAngle * Pi / 180)
        ' from you.com
        If (mAngle > 90) And (mAngle < 270) Then
            iPrev = mX1
            mX1 = mX2 + (mLength * Cos(Atn(iSlope)))
            iAdj = iPrev - mX1
            mX1 = mX1 + iAdj
            mX2 = mX2 + iAdj
            mY2 = mY1 + (mLength * Sin(Atn(iSlope))) * -1
        Else
            mX2 = mX1 + (mLength * Cos(Atn(iSlope)))
            mY2 = mY1 + (mLength * Sin(Atn(iSlope)))
        End If
        ' end from
        
        SetExtenderPosSize
        
        PostInvalidateMsg
        Me.Refresh
        PropertyChanged "Length"
        PropertyChanged "X1"
        PropertyChanged "Y1"
        PropertyChanged "X2"
        PropertyChanged "Y2"
        
        mSetFrom = leLength
    End If
End Property


Public Property Get x1() As Single
    x1 = ControlFromParent.Container.ScaleX(mX1, vbHimetric, mContainerScaleMode)
End Property

Public Property Let x1(ByVal nValue As Single)
    If nValue <> mX1 Then
        mX1 = ControlFromParent.Container.ScaleX(nValue, mContainerScaleMode, vbHimetric)
        
        SetAngleAndLengthFromXY
        SetExtenderPosSize
        mSetFrom = leXY
        mLastExtenderLeft = UserControl.Extender.Left
        mLastExtenderTop = UserControl.Extender.Top
        mLastExtenderWidth = UserControl.Extender.Width
        mLastExtenderHeight = UserControl.Extender.Height
        
        PostInvalidateMsg
        Me.Refresh
        PropertyChanged "X1"
        PropertyChanged "Length"
        PropertyChanged "Angle"
        PropertyChanged "Direction"
    End If
End Property


Public Property Get y1() As Single
    y1 = ControlFromParent.Container.ScaleY(mY1, vbHimetric, mContainerScaleMode)
End Property

Public Property Let y1(ByVal nValue As Single)
    If nValue <> mY1 Then
        mY1 = ControlFromParent.Container.ScaleY(nValue, mContainerScaleMode, vbHimetric)
        
        SetAngleAndLengthFromXY
        SetExtenderPosSize
        mSetFrom = leXY
        mLastExtenderLeft = UserControl.Extender.Left
        mLastExtenderTop = UserControl.Extender.Top
        mLastExtenderWidth = UserControl.Extender.Width
        mLastExtenderHeight = UserControl.Extender.Height
        
        PostInvalidateMsg
        Me.Refresh
        PropertyChanged "Y1"
        PropertyChanged "Length"
        PropertyChanged "Angle"
        PropertyChanged "Direction"
    End If
End Property


Public Property Get x2() As Single
    x2 = ControlFromParent.Container.ScaleX(mX2, vbHimetric, mContainerScaleMode)
End Property

Public Property Let x2(ByVal nValue As Single)
    If nValue <> mX2 Then
        mX2 = ControlFromParent.Container.ScaleX(nValue, mContainerScaleMode, vbHimetric)
        
        SetAngleAndLengthFromXY
        SetExtenderPosSize
        mSetFrom = leXY
        mLastExtenderLeft = UserControl.Extender.Left
        mLastExtenderTop = UserControl.Extender.Top
        mLastExtenderWidth = UserControl.Extender.Width
        mLastExtenderHeight = UserControl.Extender.Height
        
        PostInvalidateMsg
        Me.Refresh
        PropertyChanged "X2"
        PropertyChanged "Length"
        PropertyChanged "Angle"
        PropertyChanged "Direction"
    End If
End Property


Public Property Get y2() As Single
    y2 = ControlFromParent.Container.ScaleY(mY2, vbHimetric, mContainerScaleMode)
End Property

Public Property Let y2(ByVal nValue As Single)
    If nValue <> mY2 Then
        mY2 = ControlFromParent.Container.ScaleY(nValue, mContainerScaleMode, vbHimetric)
        
        SetAngleAndLengthFromXY
        SetExtenderPosSize
        mSetFrom = leXY
        mLastExtenderLeft = UserControl.Extender.Left
        mLastExtenderTop = UserControl.Extender.Top
        mLastExtenderWidth = UserControl.Extender.Width
        mLastExtenderHeight = UserControl.Extender.Height
        
        PostInvalidateMsg
        Me.Refresh
        PropertyChanged "Y2"
        PropertyChanged "Length"
        PropertyChanged "Angle"
        PropertyChanged "Direction"
    End If
End Property

Private Function IsValidOLE_COLOR(ByVal nColor As Long) As Boolean
    Const S_OK As Long = 0
    IsValidOLE_COLOR = (TranslateColor(nColor, 0, nColor) = S_OK)
End Function

Public Sub Refresh()
    UserControl.Refresh
End Sub

Private Sub Subclass()
    If mContainerHwnd <> 0 Then
        AttachMessage Me, mContainerHwnd, WM_INVALIDATE
        mSubclassed = True
    End If
End Sub

Private Sub Unsubclass()
    If mSubclassed Then
        DetachMessage Me, mContainerHwnd, WM_INVALIDATE
        mSubclassed = False
    End If
End Sub

' Extender properties and methods
Public Property Get Name() As String
    Name = Ambient.DisplayName
End Property

Public Property Get Tag() As String
    Tag = Extender.Tag
End Property

Public Property Let Tag(ByVal Value As String)
    Extender.Tag = Value
End Property

Public Property Get Parent() As Object
    Set Parent = UserControl.Parent
End Property

Public Property Get Container() As Object
    Set Container = Extender.Container
End Property

Public Property Set Container(ByVal Value As Object)
    Set Extender.Container = Value
End Property

Public Property Get Left() As Single
    Left = Extender.Left
End Property

Public Property Let Left(ByVal Value As Single)
    Extender.Left = Value
End Property

Public Property Get Top() As Single
    Top = Extender.Top
End Property

Public Property Let Top(ByVal Value As Single)
    Extender.Top = Value
End Property

Public Property Get Width() As Single
    Width = Extender.Width
End Property

Public Property Let Width(ByVal Value As Single)
    Extender.Width = Value
End Property

Public Property Get Height() As Single
    Height = Extender.Height
End Property

Public Property Let Height(ByVal Value As Single)
    Extender.Height = Value
End Property

Public Property Get Visible() As Boolean
    Visible = Extender.Visible
End Property

Public Property Let Visible(ByVal Value As Boolean)
    Extender.Visible = Value
End Property

Public Property Get ToolTipText() As String
    ToolTipText = Extender.ToolTipText
End Property

Public Property Let ToolTipText(ByVal Value As String)
    Extender.ToolTipText = Value
End Property

Public Property Get DragIcon() As IPictureDisp
    Set DragIcon = Extender.DragIcon
End Property

Public Property Let DragIcon(ByVal Value As IPictureDisp)
    Extender.DragIcon = Value
End Property

Public Property Set DragIcon(ByVal Value As IPictureDisp)
    Set Extender.DragIcon = Value
End Property

Public Property Get DragMode() As Integer
    DragMode = Extender.DragMode
End Property

Public Property Let DragMode(ByVal Value As Integer)
    Extender.DragMode = Value
End Property

Public Sub Drag(Optional ByRef Action As Variant)
    If IsMissing(Action) Then Extender.Drag Else Extender.Drag Action
End Sub

Public Sub ZOrder(Optional ByRef Position As Variant)
    If IsMissing(Position) Then Extender.ZOrder Else Extender.ZOrder Position
End Sub

