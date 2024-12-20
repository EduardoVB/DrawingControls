VERSION 5.00
Begin VB.UserControl ShapeEx 
   BackStyle       =   0  'Transparent
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   ClipBehavior    =   0  'None
   HasDC           =   0   'False
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "ShapeEx.ctx":0000
   Windowless      =   -1  'True
   Begin VB.Timer tmrPainting 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "ShapeEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IBSSubclass

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hDC As Long) As Long

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

Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As T_MSG, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Const PM_REMOVE = &H1

Private Type XFORM
    eM11 As Single
    eM12 As Single
    eM21 As Single
    eM22 As Single
    eDx As Single
    eDy As Single
End Type

Private Declare Function SetGraphicsMode Lib "gdi32" (ByVal hDC As Long, ByVal iMode As Long) As Long
Private Declare Function SetWorldTransform Lib "gdi32" (ByVal hDC As Long, lpXform As XFORM) As Long
Private Declare Function ModifyWorldTransform Lib "gdi32" (ByVal hDC As Long, lpXform As XFORM, ByVal iMode As Long) As Long
Private Const MWT_IDENTITY = 1
Private Const MWT_LEFTMULTIPLY = 2
'Private Const MWT_RIGHTMULTIPLY = 3

Private Const GM_ADVANCED = 2
'Private Const GM_COMPATIBLE = 1

Private Const Pi = 3.14159265358979

Private Const WM_USER As Long = &H400
Private Const WM_INVALIDATE As Long = WM_USER + 11 ' custom message

Private Declare Function GetClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long) As Long
Private Declare Function GetRgnBox Lib "gdi32" (ByVal hRgn As Long, lpRect As RECT) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function InvalidateRectAsNull Lib "user32" Alias "InvalidateRect" (ByVal hWnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long
'Private Declare Function GetUpdateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByVal lColorRef As Long) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type GDIPlusStartupInput
    GdiPlusVersion                      As Long
    DebugEventCallback                  As Long
    SuppressBackgroundThread            As Long
    SuppressExternalCodecs              As Long
End Type
 
Private Type POINTL
    X As Long
    Y As Long
End Type

Private Enum WrapMode
    WrapModeTile
    WrapModeTileFlipX
    WrapModeTileFlipY
    WrapModeTileFlipXY
    WrapModeClamp
End Enum

Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, ByRef graphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As Long) As Long
Private Declare Function GdiplusStartup Lib "gdiplus" (ByRef token As Long, ByRef lpInput As GDIPlusStartupInput, Optional ByRef lpOutput As Any) As Long
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mSmoothingMode As Long) As Long
Private Declare Function GdipDeleteBrush Lib "GdiPlus.dll" (ByVal mBrush As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "GdiPlus.dll" (ByVal mColor As Long, ByRef mBrush As Long) As Long
Private Declare Function GdipFillEllipseI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipFillRectangleI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipDrawEllipseI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipCreatePen1 Lib "gdiplus" (ByVal Color As Long, ByVal Width As Single, ByVal unit As Long, pen As Long) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long
Private Declare Function GdipDrawArcI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal StartAngle As Single, ByVal SweepAngle As Single) As Long
Private Declare Function GdipDrawLineI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function GdipFillPieI Lib "gdiplus" (ByVal graphics As Long, ByVal Brush As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal StartAngle As Single, ByVal SweepAngle As Single) As Long
Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, Col As Long) As Long
Private Declare Function GdipDrawPolygonI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByRef pPoints As Any, ByVal Count As Long) As Long
Private Declare Function GdipFillPolygonI Lib "gdiplus" (ByVal graphics As Long, ByVal Brush As Long, ByRef pPoints As Any, ByVal Count As Long, ByVal FillMode As Long) As Long
Private Declare Function GdipDrawClosedCurve2I Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, Points As Any, ByVal Count As Long, ByVal tension As Single) As Long
Private Declare Function GdipFillClosedCurve2I Lib "gdiplus" (ByVal graphics As Long, ByVal Brush As Long, Points As Any, ByVal Count As Long, ByVal tension As Single, ByVal FillMode As Long) As Long
Private Declare Function GdipSetPenDashStyle Lib "gdiplus" (ByVal pen As Long, ByVal dStyle As Long) As Long
Private Declare Function GdipCreatePath Lib "GdiPlus.dll" (ByVal mBrushMode As Long, ByRef mpath As Long) As Long
Private Declare Function GdipAddPathEllipseI Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipDeletePath Lib "GdiPlus.dll" (ByVal mpath As Long) As Long
Private Declare Function GdipSetPathGradientCenterColor Lib "GdiPlus.dll" (ByVal mBrush As Long, ByVal mColors As Long) As Long
Private Declare Function GdipSetPathGradientSurroundColorsWithCount Lib "GdiPlus.dll" (ByVal mBrush As Long, ByRef mColor As Long, ByRef mCount As Long) As Long
Private Declare Function GdipCreatePathGradientFromPath Lib "GdiPlus.dll" (ByVal mpath As Long, ByRef mPolyGradient As Long) As Long
Private Declare Function GdipCreatePathGradientI Lib "gdiplus" (Points As POINTL, ByVal Count As Long, ByVal WrapMd As Long, polyGradient As Long) As Long
Private Declare Function GdipAddPathArcI Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdipClosePathFigure Lib "GdiPlus.dll" (ByVal mpath As Long) As Long
Private Declare Function GdipDrawPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mpath As Long) As Long
Private Declare Function GdipFillPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mpath As Long) As Long
Private Declare Function GdipDrawPieI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal StartAngle As Single, ByVal SweepAngle As Single) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hbm As Long, ByVal hpal As Long, bitmap As Long) As Long
Private Declare Function GdipCreateTexture Lib "gdiplus" (ByVal Image As Long, ByVal WrapMd As WrapMode, texture As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long

Private Enum FillModeConstants
    FillModeAlternate = &H0
    FillModeWinding = &H1
End Enum

Private Const UnitPixel = 2
Private Const QualityModeLow As Long = 1
Private Const SmoothingModeAntiAlias As Long = &H4

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Event Click()
Attribute Click.VB_UserMemId = -600
Attribute Click.VB_MemberFlags = "200"
Event DblClick()
Attribute DblClick.VB_UserMemId = -601
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_UserMemId = -605
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_UserMemId = -606
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_UserMemId = -604

Public Enum SEShapeConstants
    seShapeRectangle = vbShapeRectangle ' 0
    seShapeSquare = vbShapeSquare ' 1
    seShapeOval = vbShapeOval ' 2
    seShapeCircle = vbShapeCircle ' 3
    seShapeRoundedRectangle = vbShapeRoundedRectangle ' 4
    seShapeRoundedSquare = vbShapeRoundedSquare ' 5
    seShapeTriangleEquilateral = 6
    seShapeTriangleIsosceles = 7
    seShapeTriangleScalene = 8
    seShapeTriangleRight = 9
    seShapeRhombus = 10
    seShapeKite = 11
    seShapeDiamond = 12
    seShapeTrapezoid = 13
    seShapeParalellogram = 14
    seShapeSemicircle = 15
    seShapeRegularPolygon = 16
    seShapeStar = 17
    seShapeJaggedStar = 18
    seShapeHeart = 19
    seShapeArrow = 20
    seShapeCrescent = 21
    seShapeDrop = 22
    seShapeEgg = 23
    seShapeLocation = 24
    seShapeSpeaker = 25
    seShapeCloud = 26
    seShapeTalk = 27
    seShapeShield = 28
    seShapePie = 29
End Enum

Public Enum SEBackStyleConstants
    seTransparent = 0
    seOpaque = 1
End Enum

Public Enum SEFillStyleConstants
    seFSSolid = vbFSSolid
    seFSTransparent = vbFSTransparent
    seFSTexture = 8
End Enum

Public Enum SEQualityConstants
    seQualityLow = 0
    seQualityHigh = 1
End Enum

Public Enum SEStyle3DConstants
    seStyle3DNone = 0
    seStyle3DLight = 1
    seStyle3DShadow = 2
    seStyle3DBoth = 3
End Enum

Public Enum SEStyle3DEffectConstants
    seStyle3EffectAuto = 0
    seStyle3EffectDiffuse = 1
    seStyle3EffectGem = 2
End Enum

Public Enum SEFlippedConstants
    seFlippedNo = 0
    seFlippedHorizontally = 1
    seFlippedVertically = 2
    seFlippedBoth = 3
End Enum

Public Enum SESubclassingConstants
    seSCNo = 0 ' never ' In most cases subclassing is not necessary (and without subclassing is safer for running in the IDE), but in some special cases when the figure needs to be painted outside the control bounds, it may experience glitches.
    seSCYes = 1 ' always
    seSCNotInIDE = 2 ' compiled will use subclassing
    seSCNotInIDEDesignTime = 3 ' IDE run time and compiled will use subclassing
End Enum

Public Enum SEClickModeConstants
    seClickDisabled = 0
    seClickShape = 1
    seClickControl = 2
End Enum

' Property defaults
Private Const mdef_BackColor = vbWindowBackground
Private Const mdef_BackStyle = seTransparent
Private Const mdef_BorderColor = vbWindowText
Private Const mdef_Shape = seShapeRectangle
Private Const mdef_FillColor = vbBlack
Private Const mdef_FillStyle = vbFSTransparent
Private Const mdef_BorderStyle = vbBSSolid
Private Const mdef_BorderWidth = 1
Private Const mdef_Quality = seQualityHigh
Private Const mdef_RotationDegrees = 0
Private Const mdef_Opacity = 100
Private Const mdef_Shift = 0
Private Const mdef_Vertices = 5
Private Const mdef_CurvingFactor = 0
Private Const mdef_Flipped = seFlippedNo
Private Const mdef_MousePointer = vbDefault
Private Const mdef_Style3D = seStyle3DNone
Private Const mdef_Style3DEffect = seStyle3EffectAuto
Private Const mdef_UseSubclassing = seSCNotInIDEDesignTime ' seSCYes
Private Const mdef_ClickMode = seClickShape ' seClickControl

' Properties
Private mBackColor  As Long
Private mBackStyle As SEBackStyleConstants
Private mBorderColor As Long
Private mShape As SEShapeConstants
Private mFillColor As Long
Private mFillStyle  As Long
Private mBorderStyle  As BorderStyleConstants
Private mBorderWidth  As Integer
Private mQuality As SEQualityConstants
Private mRotationDegrees As Single
Private mOpacity As Single
Private mShift As Single
Private mVertices As Integer
Private mCurvingFactor As Integer
Private mFlipped As SEFlippedConstants
Private mMousePointer As Integer
Private mMouseIcon As StdPicture
Private mStyle3D As SEStyle3DConstants
Private mStyle3DEffect As SEStyle3DEffectConstants
Private mUseSubclassing As SESubclassingConstants
Private mFillTexture As StdPicture
Private mClickMode As SEClickModeConstants

Private mGdipToken As Long
Private mContainerHwnd As Long
Private mAttached As Boolean
Private mShiftPutAutomatically As Single
Private mCurvingFactor2 As Single
Private mUserMode As Boolean
Private mSubclassed As Boolean
Private mDrawingOutsideUC As Boolean
Private mInvalidateMsgPosted As Boolean
Private mTextureBrush As Long

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

Private Sub UserControl_AmbientChanged(PropertyName As String)
    If PropertyName = "UserMode" Then mUserMode = Ambient.UserMode
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_HitTest(X As Single, Y As Single, HitResult As Integer)
    If mUserMode Then
        If mClickMode <> seClickDisabled Then
            If mClickMode = seClickControl Then
                HitResult = vbHitResultHit
            Else
                Dim iColor As Long
                Dim iDC As Long
                
                TranslateColor mBorderColor, 0, iColor
                If UserControl.Point(X, Y) = iColor Then
                    HitResult = vbHitResultHit
                ElseIf mFillStyle = vbFSSolid Then
                    TranslateColor mFillColor, 0, iColor
                    If UserControl.Point(X, Y) = iColor Then
                        HitResult = vbHitResultHit
                    End If
                ElseIf mFillStyle = seFSTexture Then
                    iDC = GetDC(UserControl.ContainerHwnd)
                    iColor = GetBkColor(iDC)
                    ReleaseDC UserControl.ContainerHwnd, iDC
                    TranslateColor iColor, 0, iColor
                    If UserControl.Point(X, Y) <> iColor Then
                        HitResult = vbHitResultHit
                    End If
                ElseIf mFillStyle = seFSTransparent Then
                    If mBackStyle = seOpaque Then
                        TranslateColor mBackColor, 0, iColor
                        If UserControl.Point(X, Y) = iColor Then
                            HitResult = vbHitResultHit
                        End If
                    End If
                End If
            End If
        End If
    Else
        HitResult = vbHitResultHit
    End If
End Sub

Private Sub UserControl_InitProperties()
    mBackColor = mdef_BackColor
    mBackStyle = mdef_BackStyle
    mBorderColor = mdef_BorderColor
    mShape = mdef_Shape
    mFillColor = mdef_FillColor
    mFillStyle = mdef_FillStyle
    mBorderStyle = mdef_BorderStyle
    mBorderWidth = mdef_BorderWidth
    mQuality = mdef_Quality
    mRotationDegrees = mdef_RotationDegrees
    mOpacity = mdef_Opacity
    mShift = mdef_Shift
    mVertices = mdef_Vertices
    mCurvingFactor = mdef_CurvingFactor
    mFlipped = mdef_Flipped
    mMousePointer = mdef_MousePointer
    Set mMouseIcon = Nothing
    mStyle3D = mdef_Style3D
    mStyle3DEffect = mdef_Style3DEffect
    mUseSubclassing = mdef_UseSubclassing
    mClickMode = mdef_ClickMode
    
    On Error Resume Next
    mContainerHwnd = UserControl.ContainerHwnd
    mUserMode = Ambient.UserMode
    On Error GoTo 0
    SetCurvingFactor2
    Subclass
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    If mClickMode <> seClickDisabled Then
        If (KeyAscii = vbKeySpace) Then
            UserControl_Click
        End If
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Paint()
    Dim hRgn As Long
    Dim rgnRect As RECT
    Dim hRgnExpand As Long
    Dim iExpandForPen As Long
    Dim iGMPrev As Long
    Dim mtx1 As XFORM, mtx2 As XFORM, c As Single, s As Single
    Dim iExpandOutsideForAngle As Long
    Dim iExpandOutsideForFigure As Long
    Dim iExpandOutsideForCurve As Long
    Dim iLng As Long
    Dim iLng2 As Long
    Dim iShift As Long
    Static sTheLastTimeWasExpanded As Boolean
    Dim iAuxExpand As Long
    
    If (mRotationDegrees > 0) Or (mFlipped <> seFlippedNo) Then
        iGMPrev = SetGraphicsMode(UserControl.hDC, GM_ADVANCED)
        ModifyWorldTransform UserControl.hDC, mtx1, MWT_IDENTITY
        If mRotationDegrees = 0 Then
            c = 1
            s = 0
        Else
            c = Cos(-mRotationDegrees / 360 * 2 * Pi)
            s = Sin(-mRotationDegrees / 360 * 2 * Pi)
        End If
        mtx1.eM11 = c: mtx1.eM12 = s: mtx1.eM21 = -s: mtx1.eM22 = c: mtx1.eDx = (UserControl.ScaleWidth - 1) / 2: mtx1.eDy = (UserControl.ScaleHeight - 1) / 2
        If mFlipped = seFlippedHorizontally Then
            mtx2.eM11 = -1: mtx2.eM22 = 1: mtx2.eDx = (UserControl.ScaleWidth - 1) / 2: mtx2.eDy = -(UserControl.ScaleHeight - 1) / 2
        ElseIf mFlipped = seFlippedVertically Then
            mtx2.eM11 = 1: mtx2.eM22 = -1: mtx2.eDx = -(UserControl.ScaleWidth - 1) / 2: mtx2.eDy = (UserControl.ScaleHeight - 1) / 2
        ElseIf mFlipped = seFlippedBoth Then
            mtx2.eM11 = -1: mtx2.eM22 = -1: mtx2.eDx = (UserControl.ScaleWidth - 1) / 2: mtx2.eDy = (UserControl.ScaleHeight - 1) / 2
        Else
            mtx2.eM11 = 1: mtx2.eM22 = 1: mtx2.eDx = -(UserControl.ScaleWidth - 1) / 2: mtx2.eDy = -(UserControl.ScaleHeight - 1) / 2
        End If
    End If
        
    iExpandForPen = mBorderWidth / 2
    If mCurvingFactor > 0 Then
        iAuxExpand = UserControl.ScaleWidth / UserControl.ScaleHeight * (1 + mCurvingFactor / 2)
        If iAuxExpand > iExpandForPen Then
            iExpandForPen = iAuxExpand
        End If
    End If
    If (mShape > seShapeRoundedSquare) Then
        If UserControl.ScaleWidth > UserControl.ScaleHeight Then
            iAuxExpand = UserControl.ScaleWidth / UserControl.ScaleHeight * mBorderWidth
        Else
            iAuxExpand = UserControl.ScaleHeight / UserControl.ScaleWidth * mBorderWidth
        End If
        If (mShape = seShapeStar) Or (mShape = seShapeJaggedStar) Then
            iExpandForPen = iExpandForPen * mVertices / 6
        End If
        If iAuxExpand > iExpandForPen Then
            iExpandForPen = iAuxExpand
        End If
    End If
    If ShapeHasShift(mShape) Then
        If UserControl.ScaleWidth > UserControl.ScaleHeight Then
            iShift = mShift * UserControl.ScaleWidth / 100
        Else
            iShift = mShift * UserControl.ScaleHeight / 100
        End If
        iLng = UserControl.ScaleWidth / 2 - iShift * 1.3
        If iLng < 0 Then
            iExpandOutsideForFigure = Abs(iLng)
        Else
            iLng = UserControl.ScaleWidth / 2 + iShift - UserControl.ScaleWidth
            If iLng > 0 Then
                iExpandOutsideForFigure = iLng
            Else
                iLng = UserControl.ScaleWidth / 2 + iShift
                If iLng < 0 Then
                    iExpandOutsideForFigure = Abs(iLng)
                End If
                If iExpandOutsideForFigure < Abs(iShift) Then
                    iExpandOutsideForFigure = Abs(iShift)
                End If
            End If
        End If
    End If
    If mCurvingFactor <> 0 Then
        iExpandOutsideForCurve = (UserControl.Width ^ 2 + UserControl.Height ^ 2) ^ 0.5 * 1.2
    End If
    If (mRotationDegrees <> 0) Then
        If (mShape <> seShapeCircle) And (mShape <> seShapeStar) And (mShape <> seShapeJaggedStar) Then
            iLng = Abs((UserControl.Width - UserControl.Height) / 2)
            iLng2 = (UserControl.Width ^ 2 + UserControl.Height ^ 2) ^ 0.5
            If iLng < iLng2 Then
                iExpandOutsideForAngle = iLng
            Else
                iExpandOutsideForAngle = iLng2
            End If
        End If
    End If
    
    iAuxExpand = iExpandForPen + iExpandOutsideForAngle + iExpandOutsideForFigure
    mDrawingOutsideUC = False
    hRgn = CreateRectRgn(0, 0, 0, 0)
    If GetClipRgn(UserControl.hDC, hRgn) = 0& Then  ' hDc is one passed to Paint
        DeleteObject hRgn: hRgn = 0
    Else
        GetRgnBox hRgn, rgnRect             ' get its bounds & adjust our region accordingly (i.e.,expand 1 pixel)
        If (iAuxExpand <> 0) Or sTheLastTimeWasExpanded Then
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
            End If
        End If
    End If
    sTheLastTimeWasExpanded = (iAuxExpand <> 0)
    tmrPainting.Enabled = False
    tmrPainting.Enabled = True
    
    If (mRotationDegrees > 0) Or (mFlipped <> seFlippedNo) Then
        SetWorldTransform UserControl.hDC, mtx1
        ModifyWorldTransform UserControl.hDC, mtx2, MWT_LEFTMULTIPLY
    End If
    
    Draw
    
    If hRgnExpand <> 0 Then SelectClipRgn UserControl.hDC, hRgn  ' restore original clip region
    If hRgn <> 0 Then DeleteObject hRgn
    
    If (mRotationDegrees > 0) Or (mFlipped <> seFlippedNo) Then
        SetGraphicsMode UserControl.hDC, iGMPrev
    End If
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

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mBackColor = PropBag.ReadProperty("BackColor", mdef_BackColor)
    mBackStyle = PropBag.ReadProperty("BackStyle", mdef_BackStyle)
    mBorderColor = PropBag.ReadProperty("BorderColor", mdef_BorderColor)
    mShape = PropBag.ReadProperty("Shape", mdef_Shape)
    mFillColor = PropBag.ReadProperty("FillColor", mdef_FillColor)
    mFillStyle = PropBag.ReadProperty("FillStyle", mdef_FillStyle)
    mBorderStyle = PropBag.ReadProperty("BorderStyle", mdef_BorderStyle)
    mBorderWidth = PropBag.ReadProperty("BorderWidth", mdef_BorderWidth)
    mQuality = PropBag.ReadProperty("Quality", mdef_Quality)
    mRotationDegrees = PropBag.ReadProperty("RotationDegrees", mdef_RotationDegrees)
    mOpacity = PropBag.ReadProperty("Opacity", mdef_Opacity)
    mShift = PropBag.ReadProperty("Shift", mdef_Shift)
    mShiftPutAutomatically = PropBag.ReadProperty("ShiftPutAutomatically", 0)
    mVertices = PropBag.ReadProperty("Vertices", mdef_Vertices)
    mCurvingFactor = PropBag.ReadProperty("CurvingFactor", mdef_CurvingFactor)
    mFlipped = PropBag.ReadProperty("Flipped", mdef_Flipped)
    mMousePointer = PropBag.ReadProperty("MousePointer", mdef_MousePointer)
    Set mMouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    mStyle3D = PropBag.ReadProperty("Style3D", mdef_Style3D)
    mStyle3DEffect = PropBag.ReadProperty("Style3DEffect", mdef_Style3DEffect)
    mUseSubclassing = PropBag.ReadProperty("UseSubclassing", mdef_UseSubclassing)
    Set mFillTexture = PropBag.ReadProperty("FillTexture", Nothing)
    mClickMode = PropBag.ReadProperty("ClickMode", mdef_ClickMode)
    
    UserControl.MousePointer = mMousePointer
    Set UserControl.MouseIcon = mMouseIcon
    
    On Error Resume Next
    mContainerHwnd = UserControl.ContainerHwnd
    mUserMode = Ambient.UserMode
    On Error GoTo 0
    SetCurvingFactor2
    Subclass
End Sub

Private Sub UserControl_Terminate()
    Unsubclass
    If mGdipToken <> 0 Then TerminateGDI
    If (mBorderWidth > 1) Or (mRotationDegrees > 0) Or (mCurvingFactor > 0) Or mDrawingOutsideUC Then InvalidateRectAsNull mContainerHwnd, 0&, 1& ' paint the container when the control is deleted if the BorderWidth is greater than 1 or the control is rotated (if it painted outside its bounds)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "BackColor", mBackColor, mdef_BackColor
    PropBag.WriteProperty "BackStyle", mBackStyle, mdef_BackStyle
    PropBag.WriteProperty "BorderColor", mBorderColor, mdef_BorderColor
    PropBag.WriteProperty "Shape", mShape, mdef_Shape
    PropBag.WriteProperty "FillColor", mFillColor, mdef_FillColor
    PropBag.WriteProperty "FillStyle", mFillStyle, mdef_FillStyle
    PropBag.WriteProperty "BorderStyle", mBorderStyle, mdef_BorderStyle
    PropBag.WriteProperty "BorderWidth", mBorderWidth, mdef_BorderWidth
    PropBag.WriteProperty "Quality", mQuality, mdef_Quality
    PropBag.WriteProperty "RotationDegrees", mRotationDegrees, mdef_RotationDegrees
    PropBag.WriteProperty "Opacity", mOpacity, mdef_Opacity
    PropBag.WriteProperty "Shift", mShift, mdef_Shift
    PropBag.WriteProperty "ShiftPutAutomatically", mShiftPutAutomatically, 0
    PropBag.WriteProperty "Vertices", mVertices, mdef_Vertices
    PropBag.WriteProperty "CurvingFactor", mCurvingFactor, mdef_CurvingFactor
    PropBag.WriteProperty "Flipped", mFlipped, mdef_Flipped
    PropBag.WriteProperty "MousePointer", mMousePointer, mdef_MousePointer
    PropBag.WriteProperty "MouseIcon", mMouseIcon, Nothing
    PropBag.WriteProperty "Style3D", mStyle3D, mdef_Style3D
    PropBag.WriteProperty "Style3DEffect", mStyle3DEffect, mdef_Style3DEffect
    PropBag.WriteProperty "UseSubclassing", mUseSubclassing, mdef_UseSubclassing
    PropBag.WriteProperty "FillTexture", mFillTexture, Nothing
    PropBag.WriteProperty "ClickMode", mClickMode, mdef_ClickMode
End Sub


Public Property Get BorderColor() As OLE_COLOR
Attribute BorderColor.VB_Description = "Returns/sets the color of an object's border."
Attribute BorderColor.VB_ProcData.VB_Invoke_Property = ";Apariencia"
Attribute BorderColor.VB_UserMemId = -503
    BorderColor = mBorderColor
End Property

Public Property Let BorderColor(ByVal nValue As OLE_COLOR)
    If nValue <> mBorderColor Then
        If Not IsValidOLE_COLOR(nValue) Then Err.Raise 380: Exit Property
        mBorderColor = nValue
        Me.Refresh
        PropertyChanged "BorderColor"
    End If
End Property


Public Property Get Shape() As SEShapeConstants
Attribute Shape.VB_Description = "Returns/sets a value indicating the appearance of a control."
Attribute Shape.VB_ProcData.VB_Invoke_Property = ";Apariencia"
    Shape = mShape
End Property

Public Property Let Shape(ByVal nValue As SEShapeConstants)
    If nValue <> mShape Then
        If (nValue < seShapeRectangle) Or (nValue > seShapePie) Then Err.Raise 380, TypeName(Me): Exit Property
        If ShapeHasShift(mShape) Then
            If mShift = mShiftPutAutomatically Then
                mShift = 0
                mShiftPutAutomatically = 0
            End If
        End If
        mShape = nValue
        If ShapeHasShift(mShape) Then
            If mShift = 0 Then
                mShift = 20
                mShiftPutAutomatically = mShift
            End If
        End If
        Me.Refresh
        PropertyChanged "Shape"
    End If
End Property

Private Function ShapeHasShift(nShape As SEShapeConstants) As Boolean
    Select Case nShape
        Case seShapeTriangleScalene, seShapeKite, seShapeDiamond, seShapeTrapezoid, seShapeParalellogram, seShapeArrow, seShapeStar, seShapeJaggedStar, seShapeTalk, seShapeCrescent
            ShapeHasShift = True
    End Select
End Function

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Apariencia"
Attribute BackColor.VB_UserMemId = -501
    BackColor = mBackColor
End Property

Public Property Let BackColor(ByVal nValue As OLE_COLOR)
    If nValue <> mBackColor Then
        If Not IsValidOLE_COLOR(nValue) Then Err.Raise 380: Exit Property
        mBackColor = nValue
        Me.Refresh
        PropertyChanged "BackColor"
    End If
End Property


Public Property Get BackStyle() As SEBackStyleConstants
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
Attribute BackStyle.VB_ProcData.VB_Invoke_Property = ";Apariencia"
Attribute BackStyle.VB_UserMemId = -502
    BackStyle = mBackStyle
End Property

Public Property Let BackStyle(ByVal nValue As SEBackStyleConstants)
    If nValue <> mBackStyle Then
        If (nValue < seTransparent) Or (nValue > seOpaque) Then Err.Raise 380, TypeName(Me): Exit Property
        mBackStyle = nValue
        Me.Refresh
        PropertyChanged "BackStyle"
    End If
End Property


Public Property Get FillColor() As OLE_COLOR
Attribute FillColor.VB_Description = "Returns/sets the color used to fill in shapes, circles, and boxes."
Attribute FillColor.VB_ProcData.VB_Invoke_Property = ";Apariencia"
Attribute FillColor.VB_UserMemId = -510
    FillColor = mFillColor
End Property

Public Property Let FillColor(ByVal nValue As OLE_COLOR)
    If nValue <> mFillColor Then
        If Not IsValidOLE_COLOR(nValue) Then Err.Raise 380: Exit Property
        mFillColor = nValue
        Me.Refresh
        PropertyChanged "FillColor"
    End If
End Property


Public Property Get FillStyle() As SEFillStyleConstants
Attribute FillStyle.VB_Description = "Returns/sets the fill style of a shape."
Attribute FillStyle.VB_ProcData.VB_Invoke_Property = ";Apariencia"
Attribute FillStyle.VB_UserMemId = -511
    FillStyle = mFillStyle
End Property

Public Property Let FillStyle(ByVal nValue As SEFillStyleConstants)
    If nValue <> mFillStyle Then
        If (nValue < seFSSolid) Or (nValue > seFSTexture) Then Err.Raise 380, TypeName(Me): Exit Property
        mFillStyle = nValue
        Me.Refresh
        PropertyChanged "FillStyle"
    End If
End Property


Public Property Get BorderStyle() As BorderStyleConstants
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = ";Apariencia"
Attribute BorderStyle.VB_UserMemId = -504
    BorderStyle = mBorderStyle
End Property

Public Property Let BorderStyle(ByVal nValue As BorderStyleConstants)
    If nValue <> mBorderStyle Then
        If (nValue < vbTransparent) Or (nValue > vbBSInsideSolid) Then Err.Raise 380, TypeName(Me): Exit Property
        mBorderStyle = nValue
        Me.Refresh
        PropertyChanged "BorderStyle"
    End If
End Property


Public Property Get BorderWidth() As Long
Attribute BorderWidth.VB_ProcData.VB_Invoke_Property = ";Apariencia"
Attribute BorderWidth.VB_UserMemId = -505
    BorderWidth = mBorderWidth
End Property

Public Property Let BorderWidth(ByVal nValue As Long)
    If nValue < 1 Then
        nValue = 1
    End If
    If nValue <> mBorderWidth Then
        mBorderWidth = nValue
        Me.Refresh
        PropertyChanged "BorderWidth"
    End If
End Property


Public Property Get Quality() As SEQualityConstants
Attribute Quality.VB_ProcData.VB_Invoke_Property = ";Apariencia"
    Quality = mQuality
End Property

Public Property Let Quality(ByVal nValue As SEQualityConstants)
    If nValue <> mQuality Then
        If (nValue < seQualityLow) Or (nValue > seQualityHigh) Then Err.Raise 380, TypeName(Me): Exit Property
        mQuality = nValue
        Me.Refresh
        PropertyChanged "Quality"
    End If
End Property


Public Property Get RotationDegrees() As Single
    RotationDegrees = mRotationDegrees
End Property

Public Property Let RotationDegrees(ByVal nValue As Single)
    Dim iFraction As Single
    
    If nValue <> mRotationDegrees Then
        iFraction = nValue - Round(nValue)
        nValue = nValue Mod 360
        If nValue < 0 Then nValue = nValue + 360
        nValue = nValue + iFraction
        If nValue >= 360 Then
            nValue = nValue - 360
        ElseIf nValue < 0 Then
            nValue = nValue + 360
        End If
        If nValue <> mRotationDegrees Then
            mRotationDegrees = nValue
            Me.Refresh
            PostInvalidateMsg
            PropertyChanged "RotationDegrees"
        End If
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
            Me.Refresh
            PostInvalidateMsg
            PropertyChanged "Opacity"
        End If
    End If
End Property


Public Property Get Shift() As Single
    Shift = mShift
End Property

Public Property Let Shift(ByVal nValue As Single)
    If nValue <> mShift Then
        mShift = nValue
        Me.Refresh
        PropertyChanged "Shift"
    End If
End Property


Public Property Get Vertices() As Integer
    Vertices = mVertices
End Property

Public Property Let Vertices(ByVal nValue As Integer)
    If nValue <> mVertices Then
        mVertices = nValue
        If mVertices < 2 Then mVertices = 2
        If mVertices > 100 Then mVertices = 100
        Me.Refresh
        PropertyChanged "Vertices"
    End If
End Property


Public Property Get CurvingFactor() As Integer
    CurvingFactor = mCurvingFactor
End Property

Public Property Let CurvingFactor(ByVal nValue As Integer)
    If nValue <> mCurvingFactor Then
        mCurvingFactor = nValue
        If mCurvingFactor < -100 Then mCurvingFactor = -100
        If mCurvingFactor > 100 Then mCurvingFactor = 100
        SetCurvingFactor2
        Me.Refresh
        PostInvalidateMsg
        PropertyChanged "CurvingFactor"
    End If
End Property


Public Property Get Flipped() As SEFlippedConstants
    Flipped = mFlipped
End Property

Public Property Let Flipped(ByVal nValue As SEFlippedConstants)
    If nValue <> mFlipped Then
        If (nValue < seFlippedNo) Or (nValue > seFlippedBoth) Then Err.Raise 380, TypeName(Me): Exit Property
        mFlipped = nValue
        Me.Refresh
        PropertyChanged "Flipped"
    End If
End Property


Public Property Get MousePointer() As VBRUN.MousePointerConstants
    MousePointer = mMousePointer
End Property

Public Property Let MousePointer(ByVal nValue As VBRUN.MousePointerConstants)
    If nValue <> mMousePointer Then
        mMousePointer = nValue
        UserControl.MousePointer = mMousePointer
        PropertyChanged "MousePointer"
    End If
End Property


Public Property Get MouseIcon() As StdPicture
    Set MouseIcon = mMouseIcon
End Property

Public Property Let MouseIcon(ByVal nValue As StdPicture)
    Set MouseIcon = nValue
End Property

Public Property Set MouseIcon(ByVal nValue As StdPicture)
    If Not nValue Is mMouseIcon Then
        Set mMouseIcon = nValue
        Set UserControl.MouseIcon = mMouseIcon
        PropertyChanged "MouseIcon"
    End If
End Property


Public Property Get Style3D() As SEStyle3DConstants
    Style3D = mStyle3D
End Property

Public Property Let Style3D(ByVal nValue As SEStyle3DConstants)
    If nValue <> mStyle3D Then
        If (nValue < seStyle3DNone) Or (nValue > seStyle3DBoth) Then Err.Raise 380, TypeName(Me): Exit Property
        mStyle3D = nValue
        Me.Refresh
        PropertyChanged "Style3D"
    End If
End Property


Public Property Get Style3DEffect() As SEStyle3DEffectConstants
    Style3DEffect = mStyle3DEffect
End Property

Public Property Let Style3DEffect(ByVal nValue As SEStyle3DEffectConstants)
    If nValue <> mStyle3DEffect Then
        If (nValue < seStyle3EffectAuto) Or (nValue > seStyle3EffectGem) Then Err.Raise 380, TypeName(Me): Exit Property
        mStyle3DEffect = nValue
        Me.Refresh
        PropertyChanged "Style3DEffect"
    End If
End Property


Public Property Get UseSubclassing() As SESubclassingConstants
    UseSubclassing = mUseSubclassing
End Property

Public Property Let UseSubclassing(ByVal nValue As SESubclassingConstants)
    Dim iMessage As T_MSG
    
    If nValue <> mUseSubclassing Then
        If (nValue < seSCNo) Or (nValue > seSCNotInIDEDesignTime) Then Err.Raise 380, TypeName(Me): Exit Property
        If Not mUserMode Then
            If nValue = seSCNo Then
                MsgBox "In most cases subclassing is not necessary (and without subclassing is safer for running in the IDE), but in some special cases when the figure needs to be painted outside the control bounds, it may experience glitches.", vbInformation
            End If
        End If
        mUseSubclassing = nValue
        If mUseSubclassing <> seSCNo Then
            If mSubclassed Then Unsubclass
            Subclass
            If mSubclassed Then
                PostInvalidateMsg
            Else
                If mContainerHwnd <> 0 Then
                    PeekMessage iMessage, mContainerHwnd, WM_INVALIDATE, WM_INVALIDATE, PM_REMOVE  ' remove posted message, if any
                End If
            End If
        Else
            If mSubclassed Then Unsubclass
            If mContainerHwnd <> 0 Then
                PeekMessage iMessage, mContainerHwnd, WM_INVALIDATE, WM_INVALIDATE, PM_REMOVE  ' remove posted message, if any
            End If
        End If
        PropertyChanged "UseSubclassing"
    End If
End Property


Public Property Get FillTexture() As StdPicture
    Set FillTexture = mFillTexture
End Property


Public Property Set FillTexture(nImage As StdPicture)
    If Not mFillTexture Is nImage Then
        If Not nImage Is Nothing Then
            If nImage.Type <> vbPicTypeBitmap Then
                Err.Raise 380, TypeName(Me), "Texture image must be type bitmap."
            End If
        End If
        Set mFillTexture = nImage
        If mTextureBrush <> 0 Then CreateTextureBrush
        Me.Refresh
        PropertyChanged "FillTexture"
    End If
End Property

Public Property Let FillTexture(nImage As StdPicture)
    Set FillTexture = nImage
End Property


Public Property Get ClickMode() As SEClickModeConstants
    ClickMode = mClickMode
End Property

Public Property Let ClickMode(ByVal nValue As SEClickModeConstants)
    If nValue <> mClickMode Then
        If (nValue < seClickDisabled) Or (nValue > seClickControl) Then Err.Raise 380, TypeName(Me): Exit Property
        mClickMode = nValue
        PropertyChanged "ClickMode"
    End If
End Property


Public Property Get hWnd() As Long
Attribute hWnd.VB_UserMemId = -515
    hWnd = UserControl.hWnd
End Property
    
Public Sub Refresh()
Attribute Refresh.VB_UserMemId = -550
    UserControl.Refresh
End Sub

Private Sub Draw()
    Dim iDiameter As Long
    Dim iGraphics As Long
    Dim iFillColor As Long
    Dim iFilled As Boolean
    Dim iUseTexture As Boolean
    Dim iHeight As Long
    Dim iRoundSize As Long
    Dim iPts() As POINTL
    Dim iEdge As Long
    Dim iUCWidth As Long
    Dim iUCHeight As Long
    Dim iLng As Long
    Dim c As Long
    Dim iPts2() As POINTL
    Dim iPts3() As POINTL
    Dim iShift As Long
    Dim iHalfBorderWidth As Long
    
    If mGdipToken = 0 Then InitGDI
    If GdipCreateFromHDC(UserControl.hDC, iGraphics) = 0 Then
        iUseTexture = False
        If mFillStyle = seFSSolid Then
            iFilled = True
            iFillColor = mFillColor
        ElseIf (mFillStyle = seFSTexture) And (Not (mFillTexture Is Nothing)) Then
            If mFillTexture.Handle <> 0 Then
                iFilled = True
                iUseTexture = True
            End If
        ElseIf mBackStyle = seOpaque Then
            iFilled = True
            iFillColor = mBackColor
        End If
        If iUseTexture Then
            If mTextureBrush = 0 Then CreateTextureBrush
        Else
            If mTextureBrush <> 0 Then DestroyTextureBrush
        End If
        
        iUCWidth = UserControl.ScaleWidth - 1
        iUCHeight = UserControl.ScaleHeight - 1
        
        If mShape = seShapeOval Then
            If iFilled Then
                FillEllipse iGraphics, iFillColor, 0, 0, iUCWidth, iUCHeight
            End If
            If mBorderStyle <> vbTransparent Then
                DrawEllipse iGraphics, mBorderColor, mBorderWidth, 0, 0, iUCWidth, iUCHeight
            End If
        ElseIf mShape = seShapeCircle Then
            If iUCWidth < iUCHeight Then
                iDiameter = iUCWidth
            Else
                iDiameter = iUCHeight
            End If
            If iFilled Then
                FillEllipse iGraphics, iFillColor, iUCWidth / 2 - iDiameter / 2, iUCHeight / 2 - iDiameter / 2, iDiameter, iDiameter
            End If
            If mBorderStyle <> vbTransparent Then
                DrawEllipse iGraphics, mBorderColor, mBorderWidth, iUCWidth / 2 - iDiameter / 2, (iUCHeight / 2 - iDiameter / 2), iDiameter - 0.51, iDiameter - 0.51
            End If
        ElseIf mShape = seShapeSquare Then
            If UserControl.ScaleWidth < UserControl.ScaleHeight Then
                iHeight = iUCWidth
            Else
                iHeight = iUCHeight
            End If
            
            ReDim iPts(3)
            iPts(0).X = iUCWidth / 2 - iHeight / 2
            iPts(0).Y = iUCHeight / 2 - iHeight / 2
            iPts(1).X = iUCWidth / 2 - iHeight / 2
            iPts(1).Y = iUCHeight / 2 - iHeight / 2 + iHeight
            iPts(2).X = iUCWidth / 2 - iHeight / 2 + iHeight
            iPts(2).Y = iUCHeight / 2 - iHeight / 2 + iHeight
            iPts(3).X = iUCWidth / 2 - iHeight / 2 + iHeight
            iPts(3).Y = iUCHeight / 2 - iHeight / 2
            
            If mBorderStyle = vbBSInsideSolid Then
                iHalfBorderWidth = mBorderWidth / 2
                iPts(0).X = iPts(0).X + iHalfBorderWidth
                iPts(0).Y = iPts(0).Y + iHalfBorderWidth
                iPts(1).X = iPts(1).X + iHalfBorderWidth
                iPts(1).Y = iPts(1).Y - iHalfBorderWidth
                iPts(2).X = iPts(2).X - iHalfBorderWidth
                iPts(2).Y = iPts(2).Y - iHalfBorderWidth
                iPts(3).X = iPts(3).X - iHalfBorderWidth
                iPts(3).Y = iPts(3).Y + iHalfBorderWidth
            End If
            
            If iFilled Then
                FillPolygon iGraphics, iFillColor, iPts
            End If
            If mBorderStyle <> vbTransparent Then
                DrawPolygon iGraphics, mBorderColor, mBorderWidth, iPts
            End If
        ElseIf mShape = seShapeRoundedRectangle Then
            If UserControl.ScaleWidth < UserControl.ScaleHeight Then
                iRoundSize = UserControl.ScaleWidth * 0.125
            Else
                iRoundSize = UserControl.ScaleHeight * 0.125
            End If
            If iFilled Then
                FillRoundRect iGraphics, iFillColor, 0, 0, iUCWidth, iUCHeight, iRoundSize
            End If
            If mBorderStyle <> vbTransparent Then
                DrawRoundRect iGraphics, mBorderColor, mBorderWidth, 0, 0, iUCWidth, iUCHeight, iRoundSize
            End If
        ElseIf mShape = seShapeRoundedSquare Then
            If UserControl.ScaleWidth < UserControl.ScaleHeight Then
                iHeight = iUCWidth
            Else
                iHeight = iUCHeight
            End If
            iRoundSize = iHeight * 0.125
            If iFilled Then
                FillRoundRect iGraphics, iFillColor, UserControl.ScaleWidth / 2 - iHeight / 2, UserControl.ScaleHeight / 2 - iHeight / 2, iHeight - 1, iHeight - 1, iRoundSize
            End If
            If mBorderStyle <> vbTransparent Then
                DrawRoundRect iGraphics, mBorderColor, mBorderWidth, UserControl.ScaleWidth / 2 - iHeight / 2, UserControl.ScaleHeight / 2 - iHeight / 2, iHeight - 1, iHeight - 1, iRoundSize
            End If
        ElseIf mShape = seShapeTriangleEquilateral Then
            ReDim iPts(2)
            
            If UserControl.ScaleWidth < UserControl.ScaleHeight Then
                iEdge = iUCWidth
            Else
                iEdge = iUCHeight
            End If
            
'            iEdge = iHeight * 2 / 3 ^ 0.5
            iHeight = (3 ^ 0.5 * iEdge) / 2
            iPts(0).X = iUCWidth / 2
            iPts(0).Y = iUCHeight / 2 - iHeight / 2
            iPts(1).X = iUCWidth / 2 - iEdge / 2
            iPts(1).Y = iUCHeight / 2 + iHeight / 2
            iPts(2).X = iUCWidth / 2 + iEdge / 2
            iPts(2).Y = iUCHeight / 2 + iHeight / 2
            
            If mBorderStyle = vbBSInsideSolid Then
                iHalfBorderWidth = mBorderWidth / 2
                iPts(0).Y = iPts(0).Y + iHalfBorderWidth
                iPts(1).X = iPts(1).X + iHalfBorderWidth / 2
                iPts(2).X = iPts(2).X - iHalfBorderWidth / 2
            End If
                
            If iFilled Then
                FillPolygon iGraphics, iFillColor, iPts
            End If
            If mBorderStyle <> vbTransparent Then
                DrawPolygon iGraphics, mBorderColor, mBorderWidth, iPts
            End If
            
        ElseIf mShape = seShapeTriangleIsosceles Then
            ReDim iPts(2)
            
            iPts(0).X = iUCWidth / 2
            iPts(0).Y = 0
            iPts(1).X = 0
            iPts(1).Y = iUCHeight
            iPts(2).X = iUCWidth
            iPts(2).Y = iUCHeight
            
            If mBorderStyle = vbBSInsideSolid Then
                iHalfBorderWidth = mBorderWidth / 2
                iPts(0).Y = iPts(0).Y + iHalfBorderWidth
                iPts(1).X = iPts(1).X + iHalfBorderWidth / 2
                iPts(2).X = iPts(2).X - iHalfBorderWidth / 2
            End If
            
            If iFilled Then
                FillPolygon iGraphics, iFillColor, iPts
            End If
            If mBorderStyle <> vbTransparent Then
                DrawPolygon iGraphics, mBorderColor, mBorderWidth, iPts
            End If
        ElseIf mShape = seShapeTriangleScalene Then
            ReDim iPts(2)
            
            iPts(0).X = iUCWidth / 2 - (iUCWidth / 100 * mShift)
            iPts(0).Y = 0
            iPts(1).X = 0
            iPts(1).Y = iUCHeight
            iPts(2).X = iUCWidth
            iPts(2).Y = iUCHeight
            
            If mBorderStyle = vbBSInsideSolid Then
                iHalfBorderWidth = mBorderWidth / 2
                iPts(0).Y = iPts(0).Y + iHalfBorderWidth
                iPts(1).X = iPts(1).X + iHalfBorderWidth / 2
                iPts(2).X = iPts(2).X - iHalfBorderWidth
            End If
            
            If iFilled Then
                FillPolygon iGraphics, iFillColor, iPts
            End If
            If mBorderStyle <> vbTransparent Then
                DrawPolygon iGraphics, mBorderColor, mBorderWidth, iPts
            End If
        ElseIf mShape = seShapeTriangleRight Then
            ReDim iPts(2)
            
            iPts(0).X = 0
            iPts(0).Y = 0
            iPts(1).X = 0
            iPts(1).Y = iUCHeight
            iPts(2).X = iUCWidth
            iPts(2).Y = iUCHeight
            
            If mBorderStyle = vbBSInsideSolid Then
                iHalfBorderWidth = mBorderWidth / 2
                iPts(0).Y = iPts(0).Y + iHalfBorderWidth
                iPts(1).X = iPts(1).X + iHalfBorderWidth / 2
                iPts(2).X = iPts(2).X - iHalfBorderWidth
            End If
            
            If iFilled Then
                FillPolygon iGraphics, iFillColor, iPts
            End If
            If mBorderStyle <> vbTransparent Then
                DrawPolygon iGraphics, mBorderColor, mBorderWidth, iPts
            End If
        ElseIf mShape = seShapeRhombus Then
            ReDim iPts(3)
            
            iPts(0).X = iUCWidth / 2
            iPts(0).Y = 0
            iPts(1).X = 0
            iPts(1).Y = iUCHeight / 2
            iPts(2).X = iUCWidth / 2
            iPts(2).Y = iUCHeight
            iPts(3).X = iUCWidth
            iPts(3).Y = iUCHeight / 2
             
            If mBorderStyle = vbBSInsideSolid Then
                iHalfBorderWidth = mBorderWidth / 2
                iPts(0).Y = iPts(0).Y + iHalfBorderWidth
                iPts(1).X = iPts(1).X + iHalfBorderWidth
                iPts(2).Y = iPts(2).Y - iHalfBorderWidth
                iPts(3).X = iPts(3).X - iHalfBorderWidth
            End If
             
            If iFilled Then
                FillPolygon iGraphics, iFillColor, iPts
            End If
            If mBorderStyle <> vbTransparent Then
                DrawPolygon iGraphics, mBorderColor, mBorderWidth, iPts
            End If
        ElseIf mShape = seShapeKite Then
            ReDim iPts(3)
            
            iLng = iUCHeight / 2 - (iUCHeight / 100 * mShift / 20 * 15)
            
            iPts(0).X = iUCWidth / 2
            iPts(0).Y = 0
            iPts(1).X = 0
            iPts(1).Y = iLng
            iPts(2).X = iUCWidth / 2
            iPts(2).Y = iUCHeight
            iPts(3).X = iUCWidth
            iPts(3).Y = iLng
             
            If mBorderStyle = vbBSInsideSolid Then
                iHalfBorderWidth = mBorderWidth / 2
                iPts(0).Y = iPts(0).Y + iHalfBorderWidth
                iPts(1).X = iPts(1).X + iHalfBorderWidth
                iPts(2).Y = iPts(2).Y - iHalfBorderWidth
                iPts(3).X = iPts(3).X - iHalfBorderWidth
            End If
             
            If iFilled Then
                FillPolygon iGraphics, iFillColor, iPts
            End If
            If mBorderStyle <> vbTransparent Then
                DrawPolygon iGraphics, mBorderColor, mBorderWidth, iPts
            End If
        ElseIf mShape = seShapeDiamond Then
            ReDim iPts(4)
            
            iLng = iUCHeight / 2 - (iUCHeight / 100 * mShift / 20 * 15)
            
            iPts(0).X = iUCWidth * 0.33
            iPts(0).Y = 0
            iPts(1).X = 0
            iPts(1).Y = iLng
            iPts(2).X = iUCWidth / 2
            iPts(2).Y = iUCHeight
            iPts(3).X = iUCWidth
            iPts(3).Y = iLng
            iPts(4).X = iUCWidth * 0.66
            iPts(4).Y = 0
             
            If mBorderStyle = vbBSInsideSolid Then
                iHalfBorderWidth = mBorderWidth / 2
                iPts(0).Y = iPts(0).Y + iHalfBorderWidth
                iPts(1).X = iPts(1).X + iHalfBorderWidth
                iPts(2).Y = iPts(2).Y - iHalfBorderWidth
                iPts(3).X = iPts(3).X - iHalfBorderWidth
                iPts(4).Y = iPts(4).Y + iHalfBorderWidth
            End If
             
            If iFilled Then
                FillPolygon iGraphics, iFillColor, iPts
            End If
            If mBorderStyle <> vbTransparent Then
                DrawPolygon iGraphics, mBorderColor, mBorderWidth, iPts
            End If
        ElseIf mShape = seShapeTrapezoid Then
            ReDim iPts(3)
            
            iLng = (iUCWidth / 100 * mShift)
            If iLng > iUCWidth / 2 Then
                iLng = iUCWidth / 2
            End If
            iPts(0).X = iLng
            iPts(0).Y = 0
            iPts(1).X = 0
            iPts(1).Y = iUCHeight
            iPts(2).X = iUCWidth
            iPts(2).Y = iUCHeight
            iPts(3).X = iUCWidth - iLng
            iPts(3).Y = 0
             
            If mBorderStyle = vbBSInsideSolid Then
                iHalfBorderWidth = mBorderWidth / 2
                iPts(0).Y = iPts(0).Y + iHalfBorderWidth
                iPts(1).X = iPts(1).X + iHalfBorderWidth
                iPts(1).Y = iPts(1).Y - iHalfBorderWidth
                iPts(2).X = iPts(2).X - iHalfBorderWidth
                iPts(2).Y = iPts(2).Y - iHalfBorderWidth
                iPts(3).Y = iPts(3).Y + iHalfBorderWidth
            End If
            
            If iFilled Then
                FillPolygon iGraphics, iFillColor, iPts
            End If
            If mBorderStyle <> vbTransparent Then
                DrawPolygon iGraphics, mBorderColor, mBorderWidth, iPts
            End If
        ElseIf mShape = seShapeParalellogram Then
            ReDim iPts(3)
            
            iLng = (iUCWidth / 100 * mShift)
            If iLng > iUCWidth Then
                iLng = iUCWidth
            End If
            iPts(0).X = iLng
            iPts(0).Y = 0
            iPts(1).X = 0
            iPts(1).Y = iUCHeight
            iPts(2).X = iUCWidth - iLng
            iPts(2).Y = iUCHeight
            iPts(3).X = iUCWidth
            iPts(3).Y = 0
             
            If mBorderStyle = vbBSInsideSolid Then
                iHalfBorderWidth = mBorderWidth / 2
                iPts(0).X = iPts(0).X + iHalfBorderWidth
                iPts(0).Y = iPts(0).Y + iHalfBorderWidth
                iPts(1).X = iPts(1).X + iHalfBorderWidth
                iPts(1).Y = iPts(1).Y - iHalfBorderWidth
                iPts(2).X = iPts(2).X - iHalfBorderWidth
                iPts(2).Y = iPts(2).Y - iHalfBorderWidth
                iPts(3).X = iPts(3).X - iHalfBorderWidth
                iPts(3).Y = iPts(3).Y + iHalfBorderWidth
            End If
            
            If iFilled Then
                FillPolygon iGraphics, iFillColor, iPts
            End If
            If mBorderStyle <> vbTransparent Then
                DrawPolygon iGraphics, mBorderColor, mBorderWidth, iPts
            End If
        ElseIf mShape = seShapeSemicircle Then
            If iFilled Then
                FillSemicircle iGraphics, iFillColor, 0, 0, iUCWidth, iUCHeight
            End If
            If mBorderStyle <> vbTransparent Then
                DrawSemicircle iGraphics, mBorderColor, mBorderWidth, 0, 0, iUCWidth, iUCHeight
            End If
        ElseIf mShape = seShapeRegularPolygon Then
            If UserControl.ScaleWidth < UserControl.ScaleHeight Then
                iHeight = iUCWidth
            Else
                iHeight = iUCHeight
            End If
            
            ReDim iPts(mVertices - 1)
            
            If mBorderStyle = vbBSInsideSolid Then
                iHeight = iHeight - mBorderWidth / 2
                If iHeight < mBorderWidth Then iHeight = mBorderWidth
            End If
            
            For c = 0 To mVertices - 1
                iPts(c).X = (iHeight / 2) * Cos(2 * Pi * (c + 1) / mVertices) + iUCWidth / 2
                iPts(c).Y = (iHeight / 2) * Sin(2 * Pi * (c + 1) / mVertices) + iUCHeight / 2
            Next c
            
            If iFilled Then
                FillPolygon iGraphics, iFillColor, iPts
            End If
            If mBorderStyle <> vbTransparent Then
                DrawPolygon iGraphics, mBorderColor, mBorderWidth, iPts
            End If
        ElseIf (mShape = seShapeStar) Then
            If UserControl.ScaleWidth < UserControl.ScaleHeight Then
                iHeight = iUCWidth
            Else
                iHeight = iUCHeight
            End If
            
            ReDim iPts(mVertices * 2 - 1)
            
            If mBorderStyle = vbBSInsideSolid Then
                iHeight = iHeight - mBorderWidth / 2
                If iHeight < mBorderWidth Then iHeight = mBorderWidth
            End If
            
            For c = 0 To mVertices * 2 - 1
                iPts(c).X = (iHeight / 2) * Cos(2 * Pi * (c + 1) / (mVertices * 2)) + iUCWidth / 2
                iPts(c).Y = (iHeight / 2) * Sin(2 * Pi * (c + 1) / (mVertices * 2)) + iUCHeight / 2
            Next c
            
            ReDim iPts2(mVertices - 1)
            iShift = (iHeight / 100 * mShift / 3) + 10
            
            For c = 0 To mVertices - 1
                iPts2(c).X = (iHeight / 2 - iShift) * Cos(2 * Pi * (c + 1) / mVertices) + iUCWidth / 2
                iPts2(c).Y = (iHeight / 2 - iShift) * Sin(2 * Pi * (c + 1) / mVertices) + iUCHeight / 2
            Next c
            
            ReDim iPts3(mVertices * 2 - 1)
            For c = 0 To mVertices * 2 - 1
                If c Mod 2 = 0 Then
                    iPts3(c).X = iPts2(c / 2).X
                    iPts3(c).Y = iPts2(c / 2).Y
                Else
                    iPts3(c).X = iPts((c + 1) Mod (UBound(iPts) + 1)).X
                    iPts3(c).Y = iPts((c + 1) Mod (UBound(iPts) + 1)).Y
                End If
            Next c
            
            If iFilled Then
                FillPolygon iGraphics, iFillColor, iPts3, FillModeWinding
            End If
            If mBorderStyle <> vbTransparent Then
                DrawPolygon iGraphics, mBorderColor, mBorderWidth, iPts3
            End If
        ElseIf (mShape = seShapeJaggedStar) Then
            If UserControl.ScaleWidth < UserControl.ScaleHeight Then
                iHeight = iUCWidth
            Else
                iHeight = iUCHeight
            End If
            
            ReDim iPts(mVertices * 2 - 1)
            
            If mBorderStyle = vbBSInsideSolid Then
                iHeight = iHeight - mBorderWidth / 2
                If iHeight < mBorderWidth Then iHeight = mBorderWidth
            End If
            
            For c = 0 To mVertices * 2 - 1
                iPts(c).X = (iHeight / 2) * Cos(2 * Pi * (c + 1) / (mVertices * 2)) + iUCWidth / 2
                iPts(c).Y = (iHeight / 2) * Sin(2 * Pi * (c + 1) / (mVertices * 2)) + iUCHeight / 2
            Next c
            
            ReDim iPts2(mVertices - 1)
            iShift = (iHeight / 100 * mShift / 3) + 10
            
            For c = 0 To mVertices - 1
                iPts2(c).X = (iHeight / 2 - iShift) * Cos(2 * Pi * (c + 1) / mVertices) + iUCWidth / 2
                iPts2(c).Y = (iHeight / 2 - iShift) * Sin(2 * Pi * (c + 1) / mVertices) + iUCHeight / 2
            Next c
            
            ReDim iPts3(mVertices * 2 - 1)
            For c = 0 To mVertices * 2 - 1
                If c Mod 2 = 0 Then
                    iPts3(c).X = iPts2(c / 2).X
                    iPts3(c).Y = iPts2(c / 2).Y
                Else
                    iPts3(c).X = iPts(c).X
                    iPts3(c).Y = iPts(c).Y
                End If
            Next c
            
            If iFilled Then
                FillPolygon iGraphics, iFillColor, iPts3, FillModeWinding
            End If
            If mBorderStyle <> vbTransparent Then
                DrawPolygon iGraphics, mBorderColor, mBorderWidth, iPts3
            End If
        ElseIf mShape = seShapeHeart Then
            ReDim iPts(13)
            
            If mBorderStyle = vbBSInsideSolid Then
                iUCWidth = iUCWidth - mBorderWidth
                iUCHeight = iUCHeight - mBorderWidth
            End If
            
            iPts(0).X = iUCWidth * 0.5
            iPts(0).Y = iUCHeight * 0.19
            iPts(1).X = iUCWidth * 0.35
            iPts(1).Y = iUCHeight * 0.04
            iPts(2).X = iUCWidth * 0.15
            iPts(2).Y = iUCHeight * 0.03
            iPts(3).X = iUCWidth * 0.005
            iPts(3).Y = iUCHeight * 0.2
            iPts(4).X = iUCWidth * 0.02
            iPts(4).Y = iUCHeight * 0.45
            iPts(5).X = iUCWidth * 0.2 ''''
            iPts(5).Y = iUCHeight * 0.7 '''
            iPts(6).X = iUCWidth * 0.49
            iPts(6).Y = iUCHeight * 0.99
            iPts(7).X = iUCWidth * 0.51
            iPts(7).Y = iUCHeight * 0.99
            iPts(8).X = iUCWidth * 0.8 '''
            iPts(8).Y = iUCHeight * 0.7 '''
            iPts(9).X = iUCWidth * 0.98
            iPts(9).Y = iUCHeight * 0.45
            iPts(10).X = iUCWidth * 0.995
            iPts(10).Y = iUCHeight * 0.2
            iPts(11).X = iUCWidth * 0.85
            iPts(11).Y = iUCHeight * 0.03
            iPts(12).X = iUCWidth * 0.65
            iPts(12).Y = iUCHeight * 0.04
            iPts(13).X = iUCWidth * 0.5
            iPts(13).Y = iUCHeight * 0.19
            
            If mBorderStyle = vbBSInsideSolid Then
                iHalfBorderWidth = mBorderWidth / 2
                For c = 0 To UBound(iPts)
                    iPts(c).X = iPts(c).X + iHalfBorderWidth
                    iPts(c).Y = iPts(c).Y + iHalfBorderWidth
                Next
            End If
                
            If iFilled Then
                FillClosedCurve iGraphics, iFillColor, iPts, 0.45
            End If
            If mBorderStyle <> vbTransparent Then
                DrawClosedCurve iGraphics, mBorderColor, mBorderWidth, iPts, 0.45
            End If
        ElseIf mShape = seShapeArrow Then
            ReDim iPts(6)
            
            iLng = iUCWidth * (0.75 - mShift / 100 * 0.75 / 20 * 15)
            If iLng > iUCWidth * 0.95 Then iLng = iUCWidth * 0.95
            
            iPts(0).X = iUCWidth * 0.005
            iPts(0).Y = iUCHeight * 0.25
            iPts(1).X = iLng
            iPts(1).Y = iUCHeight * 0.25
            iPts(2).X = iLng
            iPts(2).Y = iUCHeight * 0.005
            iPts(3).X = iUCWidth * 0.995
            iPts(3).Y = iUCHeight / 2
            iPts(4).X = iLng
            iPts(4).Y = iUCHeight * 0.995
            iPts(5).X = iLng
            iPts(5).Y = iUCHeight * 0.75
            iPts(6).X = iUCWidth * 0.005
            iPts(6).Y = iUCHeight * 0.75
            
            If mBorderStyle = vbBSInsideSolid Then
                iHalfBorderWidth = mBorderWidth / 2
                iPts(0).X = iPts(0).X + iHalfBorderWidth
                iPts(2).Y = iPts(2).Y + iHalfBorderWidth
                iPts(3).X = iPts(3).X - iHalfBorderWidth
                iPts(4).Y = iPts(4).Y - iHalfBorderWidth
                iPts(6).X = iPts(6).X + iHalfBorderWidth
            End If
            
            If iFilled Then
                FillPolygon iGraphics, iFillColor, iPts
            End If
            If mBorderStyle <> vbTransparent Then
                DrawPolygon iGraphics, mBorderColor, mBorderWidth, iPts
            End If
        ElseIf mShape = seShapeCrescent Then
            
            ReDim iPts(11)
            iLng = iUCWidth * (0.2 + mShift / 50)
            
            If mBorderStyle = vbBSInsideSolid Then
                iUCWidth = iUCWidth - mBorderWidth
                iUCHeight = iUCHeight - mBorderWidth
            End If
            
            ' top
            iPts(0).X = iUCWidth * 0.25 + iLng
            iPts(0).Y = iUCHeight * 0.005
            iPts(1).X = iUCWidth * 0.245 + iLng * 0.52
            iPts(1).Y = iUCHeight * 0.04
            ' left
            iPts(2).X = iUCWidth * 0.24
            iPts(2).Y = iUCHeight * 0.2
            iPts(3).X = iUCWidth * 0.1
            iPts(3).Y = iUCHeight * 0.5
            iPts(4).X = iUCWidth * 0.24
            iPts(4).Y = iUCHeight * 0.8
            ' bottom
            iPts(5).X = iUCWidth * 0.245 + iLng * 0.52
            iPts(5).Y = iUCHeight * 0.96
            iPts(6).X = iUCWidth * 0.25 + iLng
            iPts(6).Y = iUCHeight * 0.995
            ' right
            iPts(7).X = iUCWidth * 0.25 + iLng * 0.72
            iPts(7).Y = iUCHeight * 0.92
            iPts(8).X = iUCWidth * 0.25 + iLng * 0.44
            iPts(8).Y = iUCHeight * 0.77
            iPts(9).X = iUCWidth * 0.25 + iLng * 0.3
            iPts(9).Y = iUCHeight * 0.5
            iPts(10).X = iUCWidth * 0.25 + iLng * 0.44
            iPts(10).Y = iUCHeight * 0.23
            iPts(11).X = iUCWidth * 0.25 + iLng * 0.72
            iPts(11).Y = iUCHeight * 0.08
            
            If mBorderStyle = vbBSInsideSolid Then
                iHalfBorderWidth = mBorderWidth / 2
                For c = 0 To UBound(iPts)
                    iPts(c).X = iPts(c).X + iHalfBorderWidth
                    iPts(c).Y = iPts(c).Y + iHalfBorderWidth
                Next
            End If
            
            If iFilled Then
                FillClosedCurve iGraphics, iFillColor, iPts, 0.5
            End If
            If mBorderStyle <> vbTransparent Then
                DrawClosedCurve iGraphics, mBorderColor, mBorderWidth, iPts, 0.5
            End If
        
'            UserControl.DrawWidth = 10
'            On Error Resume Next
'            For c = 0 To UBound(iPts)
'                UserControl.PSet (iPts(c).X, iPts(c).Y), vbRed
'            Next
'            On Error GoTo 0
        
        ElseIf mShape = seShapeDrop Then
            ReDim iPts(11)
            
            If mBorderStyle = vbBSInsideSolid Then
                iUCWidth = iUCWidth - mBorderWidth
                iUCHeight = iUCHeight - mBorderWidth
            End If
            
            iPts(0).X = iUCWidth * 0.49
            iPts(0).Y = iUCHeight * 0.005
            iPts(1).X = iUCWidth * 0.25
            iPts(1).Y = iUCHeight * 0.23
            
            iPts(2).X = iUCWidth * 0.05
            iPts(2).Y = iUCHeight * 0.5
            iPts(3).X = iUCWidth * 0.05
            iPts(3).Y = iUCHeight * 0.75
            
            iPts(4).X = iUCWidth * 0.2
            iPts(4).Y = iUCHeight * 0.9
            iPts(5).X = iUCWidth * 0.4
            iPts(5).Y = iUCHeight * 0.98
            iPts(6).X = iUCWidth * 0.6
            iPts(6).Y = iUCHeight * 0.98
            iPts(7).X = iUCWidth * 0.8
            iPts(7).Y = iUCHeight * 0.9
            
            iPts(8).X = iUCWidth * 0.95
            iPts(8).Y = iUCHeight * 0.75
            iPts(9).X = iUCWidth * 0.95
            iPts(9).Y = iUCHeight * 0.5
            
            iPts(10).X = iUCWidth * 0.75
            iPts(10).Y = iUCHeight * 0.23
            iPts(11).X = iUCWidth * 0.51
            iPts(11).Y = iUCHeight * 0.005
            
            If mBorderStyle = vbBSInsideSolid Then
                iHalfBorderWidth = mBorderWidth / 2
                For c = 0 To UBound(iPts)
                    iPts(c).X = iPts(c).X + iHalfBorderWidth
                    iPts(c).Y = iPts(c).Y + iHalfBorderWidth
                Next
            End If
            
            If iFilled Then
                FillClosedCurve iGraphics, iFillColor, iPts, 0.5
            End If
            If mBorderStyle <> vbTransparent Then
                DrawClosedCurve iGraphics, mBorderColor, mBorderWidth, iPts, 0.5
            End If
        ElseIf mShape = seShapeEgg Then
            ReDim iPts(11)
            
            If mBorderStyle = vbBSInsideSolid Then
                iUCWidth = iUCWidth - mBorderWidth
                iUCHeight = iUCHeight - mBorderWidth
            End If
            
            iPts(0).X = iUCWidth * 0.4
            iPts(0).Y = iUCHeight * 0.1
            iPts(1).X = iUCWidth * 0.2
            iPts(1).Y = iUCHeight * 0.26
            
            iPts(2).X = iUCWidth * 0.05
            iPts(2).Y = iUCHeight * 0.53
            iPts(3).X = iUCWidth * 0.05
            iPts(3).Y = iUCHeight * 0.75
            
            iPts(4).X = iUCWidth * 0.18
            iPts(4).Y = iUCHeight * 0.92
            iPts(5).X = iUCWidth * 0.4
            iPts(5).Y = iUCHeight * 0.99
            iPts(6).X = iUCWidth * 0.6
            iPts(6).Y = iUCHeight * 0.99
            iPts(7).X = iUCWidth * 0.82
            iPts(7).Y = iUCHeight * 0.92
            
            iPts(8).X = iUCWidth * 0.95
            iPts(8).Y = iUCHeight * 0.75
            iPts(9).X = iUCWidth * 0.95
            iPts(9).Y = iUCHeight * 0.53
            
            iPts(10).X = iUCWidth * 0.8
            iPts(10).Y = iUCHeight * 0.26
            iPts(11).X = iUCWidth * 0.6
            iPts(11).Y = iUCHeight * 0.1
            
            If mBorderStyle = vbBSInsideSolid Then
                iHalfBorderWidth = mBorderWidth / 2
                For c = 0 To UBound(iPts)
                    iPts(c).X = iPts(c).X + iHalfBorderWidth
                    iPts(c).Y = iPts(c).Y + iHalfBorderWidth
                Next
            End If
            
            If iFilled Then
                FillClosedCurve iGraphics, iFillColor, iPts, 0.5
            End If
            If mBorderStyle <> vbTransparent Then
                DrawClosedCurve iGraphics, mBorderColor, mBorderWidth, iPts, 0.5
            End If
            
'            UserControl.DrawWidth = 10
'            On Error Resume Next
'            For c = 0 To UBound(iPts)
'                UserControl.PSet (iPts(c).X, iPts(c).Y), IIf(c = 4, vbGreen, IIf(c = 13, vbBlue, vbRed))
'            Next
'            On Error GoTo 0

        ElseIf mShape = seShapeLocation Then
            Dim iUCWidthOrig As Long
            Dim iUCHeightOrig As Long
            
            iUCWidthOrig = iUCWidth
            iUCHeightOrig = iUCHeight
            
            If mBorderStyle = vbBSInsideSolid Then
                iUCWidth = iUCWidth - mBorderWidth
                iUCHeight = iUCHeight - mBorderWidth
            End If
            
            If iFilled Then
                ReDim iPts(24)
                
                ' start going from bottom middle to left
                iPts(0).X = iUCWidth * 0.49
                iPts(0).Y = iUCHeight * 0.98
                iPts(1).X = iUCWidth * 0.28
                iPts(1).Y = iUCHeight * 0.77
                ' outer left
                iPts(2).X = iUCWidth * 0.05
                iPts(2).Y = iUCHeight * 0.5
                iPts(3).X = iUCWidth * 0.05
                iPts(3).Y = iUCHeight * 0.25
                ' outer top
                iPts(4).X = iUCWidth * 0.23
                iPts(4).Y = iUCHeight * 0.097
                iPts(5).X = iUCWidth * 0.4
                iPts(5).Y = iUCHeight * 0.05
                iPts(6).X = iUCWidth * 0.6
                iPts(6).Y = iUCHeight * 0.05
                iPts(7).X = iUCWidth * 0.77
                iPts(7).Y = iUCHeight * 0.097
                ' outer right
                iPts(8).X = iUCWidth * 0.95
                iPts(8).Y = iUCHeight * 0.25
                iPts(9).X = iUCWidth * 0.95
                iPts(9).Y = iUCHeight * 0.5
                ' going from right to bottom
                iPts(10).X = iUCWidth * 0.72
                iPts(10).Y = iUCHeight * 0.77
                ' at the bottom
                iPts(11).X = iUCWidth * 0.51
                iPts(11).Y = iUCHeight * 0.98
                iPts(12).X = iUCWidth * 0.5
                iPts(12).Y = iUCHeight * 0.97
                ' go inside, bottom of circle
                iPts(13).X = iUCWidth * 0.5
                iPts(13).Y = iUCHeight * 0.641
                iPts(14).X = iUCWidth * 0.47
                iPts(14).Y = iUCHeight * 0.591
                ' inner right of circle
                iPts(15).X = iUCWidth * 0.65
                iPts(15).Y = iUCHeight * 0.52
                iPts(16).X = iUCWidth * 0.73
                iPts(16).Y = iUCHeight * 0.38
                ' inner top of circle
                iPts(17).X = iUCWidth * 0.62
                iPts(17).Y = iUCHeight * 0.23
                iPts(18).X = iUCWidth * 0.38
                iPts(18).Y = iUCHeight * 0.23
                ' inner left of circle
                iPts(19).X = iUCWidth * 0.26
                iPts(19).Y = iUCHeight * 0.38
                iPts(20).X = iUCWidth * 0.34
                iPts(20).Y = iUCHeight * 0.52
                ' again in bottom of circle
                iPts(21).X = iUCWidth * 0.48
                iPts(21).Y = iUCHeight * 0.581
                iPts(22).X = iUCWidth * 0.48
                iPts(22).Y = iUCHeight * 0.601
                iPts(23).X = iUCWidth * 0.5
                iPts(23).Y = iUCHeight * 0.641
                ' go to outer bottom (to join the start)
                iPts(24).X = iUCWidth * 0.5
                iPts(24).Y = iUCHeight * 0.945
                
                If mBorderStyle = vbBSInsideSolid Then
                    iHalfBorderWidth = mBorderWidth / 2
                    For c = 0 To UBound(iPts)
                        iPts(c).X = iPts(c).X + iHalfBorderWidth
                        iPts(c).Y = iPts(c).Y + iHalfBorderWidth
                    Next
                End If
                
                FillClosedCurve iGraphics, iFillColor, iPts, 0.55, FillModeWinding
                If mBorderStyle <> vbTransparent Then
                    DrawEllipse iGraphics, mBorderColor, mBorderWidth, iUCWidthOrig / 2 - iUCWidthOrig * 0.47 / 2, iUCHeightOrig * 0.202, iUCWidthOrig * 0.47, iUCHeightOrig * 0.372
                End If
                If mBorderStyle <> vbTransparent Then
                    ReDim iPts(11)
                    
                    ' start going from bottom middle to left
                    iPts(0).X = iUCWidth * 0.49
                    iPts(0).Y = iUCHeight * 0.98
                    iPts(1).X = iUCWidth * 0.28
                    iPts(1).Y = iUCHeight * 0.77
                    ' outer left
                    iPts(2).X = iUCWidth * 0.05
                    iPts(2).Y = iUCHeight * 0.5
                    iPts(3).X = iUCWidth * 0.05
                    iPts(3).Y = iUCHeight * 0.25
                    ' outer top
                    iPts(4).X = iUCWidth * 0.23
                    iPts(4).Y = iUCHeight * 0.097
                    iPts(5).X = iUCWidth * 0.4
                    iPts(5).Y = iUCHeight * 0.05
                    iPts(6).X = iUCWidth * 0.6
                    iPts(6).Y = iUCHeight * 0.05
                    iPts(7).X = iUCWidth * 0.77
                    iPts(7).Y = iUCHeight * 0.097
                    ' outer right
                    iPts(8).X = iUCWidth * 0.95
                    iPts(8).Y = iUCHeight * 0.25
                    iPts(9).X = iUCWidth * 0.95
                    iPts(9).Y = iUCHeight * 0.5
                    ' going from right to bottom
                    iPts(10).X = iUCWidth * 0.72
                    iPts(10).Y = iUCHeight * 0.77
                    iPts(11).X = iUCWidth * 0.51
                    iPts(11).Y = iUCHeight * 0.98
                    
                    If mBorderStyle = vbBSInsideSolid Then
                        iHalfBorderWidth = mBorderWidth / 2
                        For c = 0 To UBound(iPts)
                            iPts(c).X = iPts(c).X + iHalfBorderWidth
                            iPts(c).Y = iPts(c).Y + iHalfBorderWidth
                        Next
                    End If
                    
                    DrawClosedCurve iGraphics, mBorderColor, mBorderWidth, iPts, 0.5
                End If
            ElseIf mBorderStyle <> vbTransparent Then
                ReDim iPts(11)
                
                ' start going from bottom middle to left
                iPts(0).X = iUCWidth * 0.49
                iPts(0).Y = iUCHeight * 0.98
                iPts(1).X = iUCWidth * 0.28
                iPts(1).Y = iUCHeight * 0.77
                ' outer left
                iPts(2).X = iUCWidth * 0.05
                iPts(2).Y = iUCHeight * 0.5
                iPts(3).X = iUCWidth * 0.05
                iPts(3).Y = iUCHeight * 0.25
                ' outer top
                iPts(4).X = iUCWidth * 0.23
                iPts(4).Y = iUCHeight * 0.097
                iPts(5).X = iUCWidth * 0.4
                iPts(5).Y = iUCHeight * 0.05
                iPts(6).X = iUCWidth * 0.6
                iPts(6).Y = iUCHeight * 0.05
                iPts(7).X = iUCWidth * 0.77
                iPts(7).Y = iUCHeight * 0.097
                ' outer right
                iPts(8).X = iUCWidth * 0.95
                iPts(8).Y = iUCHeight * 0.25
                iPts(9).X = iUCWidth * 0.95
                iPts(9).Y = iUCHeight * 0.5
                ' going from right to bottom
                iPts(10).X = iUCWidth * 0.72
                iPts(10).Y = iUCHeight * 0.77
                iPts(11).X = iUCWidth * 0.51
                iPts(11).Y = iUCHeight * 0.98
                
                If mBorderStyle = vbBSInsideSolid Then
                    iHalfBorderWidth = mBorderWidth / 2
                    For c = 0 To UBound(iPts)
                        iPts(c).X = iPts(c).X + iHalfBorderWidth
                        iPts(c).Y = iPts(c).Y + iHalfBorderWidth
                    Next
                End If
                
                DrawClosedCurve iGraphics, mBorderColor, mBorderWidth, iPts, 0.5
                
                'DrawEllipse iGraphics, mBorderColor, mBorderWidth, iUCWidth / 2 - iUCWidth * 0.47 / 2, iUCHeight * 0.205, iUCWidth * 0.47, iUCHeight * 0.365
                DrawEllipse iGraphics, mBorderColor, mBorderWidth, iUCWidth / 2 - iUCWidth * 0.47 / 2, iUCHeight * 0.202, iUCWidth * 0.47, iUCHeight * 0.372
            End If
        ElseIf mShape = seShapeSpeaker Then
            ReDim iPts(5)
            
            iPts(0).X = 0
            iPts(0).Y = iUCHeight * 0.28
            iPts(1).X = iUCWidth * 0.37
            iPts(1).Y = iUCHeight * 0.28
            iPts(2).X = iUCWidth
            iPts(2).Y = 0
            iPts(3).X = iUCWidth
            iPts(3).Y = iUCHeight
            iPts(4).X = iUCWidth * 0.37
            iPts(4).Y = iUCHeight * 0.72
            iPts(5).X = 0
            iPts(5).Y = iUCHeight * 0.72
            
            If mBorderStyle = vbBSInsideSolid Then
                iHalfBorderWidth = mBorderWidth / 2
                iPts(0).X = iPts(0).X + iHalfBorderWidth
                iPts(2).X = iPts(2).X - iHalfBorderWidth
                iPts(2).Y = iPts(2).Y + iHalfBorderWidth
                iPts(3).X = iPts(3).X - iHalfBorderWidth
                iPts(3).Y = iPts(3).Y - iHalfBorderWidth
                iPts(5).X = iPts(5).X + iHalfBorderWidth
            End If
            
            If iFilled Then
                FillPolygon iGraphics, iFillColor, iPts
            End If
            If mBorderStyle <> vbTransparent Then
                DrawPolygon iGraphics, mBorderColor, mBorderWidth, iPts
            End If
        ElseIf mShape = seShapeCloud Then
            ReDim iPts(19)
            
            If mBorderStyle = vbBSInsideSolid Then
                iUCWidth = iUCWidth - mBorderWidth
                iUCHeight = iUCHeight - mBorderWidth
            End If
            
            ' bottom, starting at the middle and going left
            iPts(0).X = iUCWidth * 0.49
            iPts(0).Y = iUCHeight * 0.995
            iPts(1).X = iUCWidth * 0.2
            iPts(1).Y = iUCHeight * 0.995
            ' left
            iPts(2).X = iUCWidth * 0.015
            iPts(2).Y = iUCHeight * 0.85
            iPts(3).X = iUCWidth * 0.015
            iPts(3).Y = iUCHeight * 0.6
            ' left middle
            iPts(4).X = iUCWidth * 0.11
            iPts(4).Y = iUCHeight * 0.45
            ' point pushing inside
            iPts(5).X = iUCWidth * 0.22
            iPts(5).Y = iUCHeight * 0.4
            ' going up
            iPts(6).X = iUCWidth * 0.25
            iPts(6).Y = iUCHeight * 0.2
            iPts(7).X = iUCWidth * 0.29
            iPts(7).Y = iUCHeight * 0.12
            ' top
            iPts(8).X = iUCWidth * 0.35
            iPts(8).Y = iUCHeight * 0.07
            iPts(9).X = iUCWidth * 0.5
            iPts(9).Y = iUCHeight * 0.1
            iPts(10).X = iUCWidth * 0.63
            iPts(10).Y = iUCHeight * 0.3
            ' going down, new part
            iPts(11).X = iUCWidth * 0.63
            iPts(11).Y = iUCHeight * 0.3
            iPts(12).X = iUCWidth * 0.72
            iPts(12).Y = iUCHeight * 0.27
            iPts(13).X = iUCWidth * 0.78
            iPts(13).Y = iUCHeight * 0.37
            iPts(14).X = iUCWidth * 0.8
            iPts(14).Y = iUCHeight * 0.56
            iPts(15).X = iUCWidth * 0.8
            iPts(15).Y = iUCHeight * 0.56
            ' to the right
            iPts(16).X = iUCWidth * 0.9
            iPts(16).Y = iUCHeight * 0.7
            iPts(17).X = iUCWidth * 0.9
            iPts(17).Y = iUCHeight * 0.9
            iPts(18).X = iUCWidth * 0.8
            iPts(18).Y = iUCHeight * 0.995
            iPts(19).X = iUCWidth * 0.51
            iPts(19).Y = iUCHeight * 0.995

            If mBorderStyle = vbBSInsideSolid Then
                iHalfBorderWidth = mBorderWidth / 2
                For c = 0 To UBound(iPts)
                    iPts(c).X = iPts(c).X + iHalfBorderWidth
                    iPts(c).Y = iPts(c).Y + iHalfBorderWidth
                Next
            End If

            If iFilled Then
                FillClosedCurve iGraphics, iFillColor, iPts, 0.5
            End If
            If mBorderStyle <> vbTransparent Then
                DrawClosedCurve iGraphics, mBorderColor, mBorderWidth, iPts, 0.5
            End If
        ElseIf mShape = seShapeTalk Then
            iLng = mShift
            If iLng < 0 Then iLng = 0
            
            If mBorderStyle = vbBSInsideSolid Then
                iUCWidth = iUCWidth - mBorderWidth
                iUCHeight = iUCHeight - mBorderWidth
            End If
            
            iShift = iUCWidth / 100 * (mShift - 18) * 0.5
            If iShift > 300 Then iShift = 300
            If iShift < -300 Then iShift = -300
            If iLng > 0 Then
                ReDim iPts(16)
            Else
                ReDim iPts(9)
            End If
            
            ' left
            If iLng > 0 Then
                iPts(0).X = iUCWidth * 0.09
                iPts(0).Y = iUCHeight * 0.74
            Else
                iPts(0).X = iUCWidth * 0.15
                iPts(0).Y = iUCHeight * 0.78
            End If
            iPts(1).X = iUCWidth * 0.05
            iPts(1).Y = iUCHeight * 0.65
            iPts(2).X = iUCWidth * 0.05
            iPts(2).Y = iUCHeight * 0.27
            ' top
            iPts(3).X = iUCWidth * 0.15
            iPts(3).Y = iUCHeight * 0.1
            iPts(4).X = iUCWidth * 0.5
            iPts(4).Y = iUCHeight * 0.05
            iPts(5).X = iUCWidth * 0.85
            iPts(5).Y = iUCHeight * 0.1
            ' right
            iPts(6).X = iUCWidth * 0.99
            iPts(6).Y = iUCHeight * 0.26
            iPts(7).X = iUCWidth * 0.965
            iPts(7).Y = iUCHeight * 0.66
            ' bottom
            iPts(8).X = iUCWidth * 0.78
            iPts(8).Y = iUCHeight * 0.77
            iPts(9).X = iUCWidth * 0.4
            iPts(9).Y = iUCHeight * 0.78
            If iLng > 0 Then
                ' bottom left, the following is the start of the spike
                iPts(10).X = iUCWidth * 0.31
                iPts(10).Y = iUCHeight * 0.78 + iShift * 0.035
                iPts(11).X = iPts(10).X
                iPts(11).Y = iPts(10).Y
                iPts(12).X = iUCWidth * 0.25
                iPts(12).Y = iUCHeight * 0.81 + iShift * 0.04
                ' bottom left, the following is the point spike
                iPts(13).X = iUCWidth * 0.01 - iShift
                iPts(13).Y = iUCHeight * 0.99 + iShift * 0.5
                iPts(14).X = iUCWidth * 0.115
                iPts(14).Y = iUCHeight * 0.85
                iPts(15).X = iUCWidth * 0.14
                iPts(15).Y = iUCHeight * 0.81
                iPts(16).X = iUCWidth * 0.15
                iPts(16).Y = iUCHeight * 0.77
            End If
            
            If mBorderStyle = vbBSInsideSolid Then
                iHalfBorderWidth = mBorderWidth / 2
                For c = 0 To UBound(iPts)
                    iPts(c).X = iPts(c).X + iHalfBorderWidth
                    iPts(c).Y = iPts(c).Y + iHalfBorderWidth
                Next
            End If
            
            If iFilled Then
                FillClosedCurve iGraphics, iFillColor, iPts, 0.5
            End If
            
'            UserControl.DrawWidth = 10
'            On Error Resume Next
'            For c = 0 To UBound(iPts)
'                UserControl.PSet (iPts(c).X, iPts(c).Y), IIf(c = 10, vbGreen, IIf(c = 13, vbBlue, vbRed))
'            Next
'            On Error GoTo 0

            If mBorderStyle <> vbTransparent Then
                DrawClosedCurve iGraphics, mBorderColor, mBorderWidth, iPts, 0.5
            End If
            
            If mShift < 0 Then
                iShift = iShift * -1
                If iFilled Then
                    FillEllipse iGraphics, iFillColor, iUCWidth * 0.24 - iShift * 0.4, iUCHeight * 0.79 + iShift * 0.05, iUCWidth * 0.05 + iUCWidth * 0.05 * iShift / 150, iUCHeight * 0.1 + iUCHeight * 0.1 * iShift / 150
                    FillEllipse iGraphics, iFillColor, iUCWidth * 0.23 - iShift * 0.7, iUCHeight * 0.84 + iShift * 0.16, iUCWidth * 0.035 + iUCWidth * 0.035 * iShift / 150, iUCHeight * 0.07 + iUCHeight * 0.07 * iShift / 150
                    FillEllipse iGraphics, iFillColor, iUCWidth * 0.18 - iShift * 0.9, iUCHeight * 0.92 + iShift * 0.22, iUCWidth * 0.025 + iUCWidth * 0.025 * iShift / 150, iUCHeight * 0.05 + iUCHeight * 0.05 * iShift / 150
                End If
                DrawEllipse iGraphics, mBorderColor, mBorderWidth, iUCWidth * 0.24 - iShift * 0.4, iUCHeight * 0.79 + iShift * 0.05, iUCWidth * 0.05 + iUCWidth * 0.05 * iShift / 150, iUCHeight * 0.1 + iUCHeight * 0.1 * iShift / 150
                DrawEllipse iGraphics, mBorderColor, mBorderWidth, iUCWidth * 0.23 - iShift * 0.7, iUCHeight * 0.84 + iShift * 0.16, iUCWidth * 0.035 + iUCWidth * 0.035 * iShift / 150, iUCHeight * 0.07 + iUCHeight * 0.07 * iShift / 150
                DrawEllipse iGraphics, mBorderColor, mBorderWidth, iUCWidth * 0.18 - iShift * 0.9, iUCHeight * 0.92 + iShift * 0.22, iUCWidth * 0.025 + iUCWidth * 0.025 * iShift / 150, iUCHeight * 0.05 + iUCHeight * 0.05 * iShift / 150
            End If
            
        ElseIf mShape = seShapeShield Then
            ReDim iPts(22)
            
            If mBorderStyle = vbBSInsideSolid Then
                iUCWidth = iUCWidth - mBorderWidth
                iUCHeight = iUCHeight - mBorderWidth
            End If
            
            'point top
            iPts(0).X = iUCWidth * 0.51
            iPts(0).Y = iUCHeight * 0.005
            iPts(1).X = iUCWidth * 0.49
            iPts(1).Y = iUCHeight * 0.005
            ' side top-left
            iPts(2).X = iUCWidth * 0.4
            iPts(2).Y = iUCHeight * 0.07 '
            iPts(3).X = iUCWidth * 0.275
            iPts(3).Y = iUCHeight * 0.14 '
            iPts(4).X = iUCWidth * 0.14
            iPts(4).Y = iUCHeight * 0.202  '
            iPts(5).X = iUCWidth * 0.047
            iPts(5).Y = iUCHeight * 0.237
            ' point left
            iPts(6).X = iUCWidth * 0.005
            iPts(6).Y = iUCHeight * 0.252
            iPts(7).X = iUCWidth * 0.007
            iPts(7).Y = iUCHeight * 0.262
            iPts(8).X = iUCWidth * 0.01
            iPts(8).Y = iUCHeight * 0.28
            ' side bottom-left
            iPts(9).X = iUCWidth * 0.1
            iPts(9).Y = iUCHeight * 0.57
            iPts(10).X = iUCWidth * 0.27
            iPts(10).Y = iUCHeight * 0.83
            ' point bottom
            iPts(11).X = iUCWidth * 0.465
            iPts(11).Y = iUCHeight * 0.973
            iPts(12).X = iUCWidth * 0.5
            iPts(12).Y = iUCHeight * 0.995
            iPts(13).X = iUCWidth * 0.535
            iPts(13).Y = iUCHeight * 0.973
            ' side bottom right
            iPts(14).X = iUCWidth * 0.73
            iPts(14).Y = iUCHeight * 0.83
            iPts(15).X = iUCWidth * 0.9
            iPts(15).Y = iUCHeight * 0.57
            ' point right
            iPts(16).X = iUCWidth * 0.99
            iPts(16).Y = iUCHeight * 0.28
            iPts(17).X = iUCWidth * 0.993
            iPts(17).Y = iUCHeight * 0.262
            iPts(18).X = iUCWidth * 0.995
            iPts(18).Y = iUCHeight * 0.252
            ' side top right
            iPts(19).X = iUCWidth * 0.953
            iPts(19).Y = iUCHeight * 0.237
            iPts(20).X = iUCWidth * 0.86
            iPts(20).Y = iUCHeight * 0.202
            iPts(21).X = iUCWidth * 0.725
            iPts(21).Y = iUCHeight * 0.14 '
            iPts(22).X = iUCWidth * 0.6
            iPts(22).Y = iUCHeight * 0.07 '
            
            If mBorderStyle = vbBSInsideSolid Then
                iHalfBorderWidth = mBorderWidth / 2
                For c = 0 To UBound(iPts)
                    iPts(c).X = iPts(c).X + iHalfBorderWidth
                    iPts(c).Y = iPts(c).Y + iHalfBorderWidth
                Next
            End If
            
            If iFilled Then
                FillClosedCurve iGraphics, iFillColor, iPts, 0.5
            End If
            If mBorderStyle <> vbTransparent Then
                DrawClosedCurve iGraphics, mBorderColor, mBorderWidth, iPts, 0.5
            End If
        ElseIf mShape = seShapePie Then
            If iFilled Then
                FillPie iGraphics, iFillColor, 0, 0, iUCWidth, iUCHeight, mShift + 60
            End If
            If mBorderStyle <> vbTransparent Then
                DrawPie iGraphics, mBorderColor, mBorderWidth, 0, 0, iUCWidth, iUCHeight, mShift + 60
            End If
        Else ' mShape = seShapeRectangle
            ReDim iPts(3)
            
            iPts(0).X = 0
            iPts(0).Y = 0
            iPts(1).X = 0
            iPts(1).Y = iUCHeight
            iPts(2).X = iUCWidth
            iPts(2).Y = iUCHeight
            iPts(3).X = iUCWidth
            iPts(3).Y = 0
            
            If mBorderStyle = vbBSInsideSolid Then
                iHalfBorderWidth = mBorderWidth / 2
                iPts(0).X = iPts(0).X + iHalfBorderWidth
                iPts(0).Y = iPts(0).Y + iHalfBorderWidth
                iPts(1).X = iPts(1).X + iHalfBorderWidth
                iPts(1).Y = iPts(1).Y - iHalfBorderWidth
                iPts(2).X = iPts(2).X - iHalfBorderWidth
                iPts(2).Y = iPts(2).Y - iHalfBorderWidth
                iPts(3).X = iPts(3).X - iHalfBorderWidth
                iPts(3).Y = iPts(3).Y + iHalfBorderWidth
            End If
            
            If iFilled Then
                FillPolygon iGraphics, iFillColor, iPts
            End If
            If mBorderStyle <> vbTransparent Then
                DrawPolygon iGraphics, mBorderColor, mBorderWidth, iPts
            End If
        End If
        
        Call GdipDeleteGraphics(iGraphics)
    End If
End Sub

Private Sub FillPolygon(ByVal nGraphics As Long, ByVal nColor As Long, Points() As POINTL, Optional nFillMode As FillModeConstants = FillModeAlternate)
    Dim hBrush As Long
    Dim iRet As Long
    Dim iStyle3DEffect As Long
    Dim iPath As Long
    Dim iRect As RECT
    Dim iPoints() As POINTL
    Dim c As Long
    
    If mTextureBrush = 0 Then
        If mStyle3D <> 0 Then
            If mStyle3DEffect = seStyle3EffectAuto Then
                If (mShape = seShapeParalellogram) Or (mShape = seShapeRectangle) Or (mShape = seShapeSquare) Or (mShape = seShapeTrapezoid) Or (mShape = seShapeTriangleScalene) Or (mShape = seShapeTriangleRight) Or (mShape = seShapeTriangleIsosceles) Or (mShape = seShapeTriangleEquilateral) Then
                    iStyle3DEffect = seStyle3EffectDiffuse
                Else
                    iStyle3DEffect = seStyle3EffectGem
                End If
            Else
                iStyle3DEffect = mStyle3DEffect
            End If
            
            If iStyle3DEffect = seStyle3EffectDiffuse Then
                GdipCreatePath 0&, iPath
                iRect = ScaleRect(GetPointsLRect(Points), Sqr(2) * (1 + Abs(mCurvingFactor) / 400))
                GdipAddPathEllipseI iPath, iRect.Left, iRect.Top, iRect.Right - iRect.Left, iRect.Bottom - iRect.Top
                iRet = GdipCreatePathGradientFromPath(iPath, hBrush)
            Else
                ReDim iPoints(UBound(Points))
                For c = 0 To UBound(Points)
                    iPoints(c) = Points(c)
                Next
                If mCurvingFactor <> 0 Then
                    If (mShape = seShapeTriangleScalene) Or (mShape = seShapeTriangleRight) Then
                        ' add a point
                        ReDim Preserve iPoints(3)
                        iPoints(3).X = (iPoints(0).X ^ 2 + iPoints(2).X ^ 2) ^ 0.5 + UserControl.ScaleWidth / 300 * Abs(mCurvingFactor)
                        iPoints(3).Y = (iPoints(0).Y ^ 2 + iPoints(2).Y ^ 2) ^ 0.5 - UserControl.ScaleHeight / 300 * Abs(mCurvingFactor)
                    End If
                    iPoints = ExpandPointsL(iPoints, Abs(mCurvingFactor) / 80)
                End If
                iRet = GdipCreatePathGradientI(iPoints(0), UBound(iPoints) + 1, 0&, hBrush)
            End If
            If iRet = 0 Then
                GdipSetPathGradientCenterColor hBrush, ConvertColor(ShiftColor(nColor, vbWhite, IIf(mStyle3D And seStyle3DLight, 200, 255)), mOpacity)
                GdipSetPathGradientSurroundColorsWithCount hBrush, ConvertColor(ShiftColor(nColor, vbBlack, IIf(mStyle3D And seStyle3DShadow, 200, 255)), mOpacity), 1
            End If
        Else
            iRet = GdipCreateSolidFill(ConvertColor(nColor, mOpacity), hBrush)
        End If
    End If
    
    If iRet = 0 Then
        Call GdipSetSmoothingMode(nGraphics, SmoothingMode)
        If mCurvingFactor = 0 Then
            GdipFillPolygonI nGraphics, IIf(mTextureBrush <> 0, mTextureBrush, hBrush), Points(0), UBound(Points) + 1, nFillMode
        Else
            GdipFillClosedCurve2I nGraphics, hBrush, Points(0), UBound(Points) + 1, mCurvingFactor2, nFillMode
        End If
        Call GdipDeleteBrush(hBrush)
        If iPath <> 0 Then
            GdipDeletePath iPath
        End If
    End If
    
End Sub

Private Sub DrawPolygon(ByVal nGraphics As Long, ByVal nColor As Long, ByVal nDrawWidth As Long, Points() As POINTL)
    Dim hPen As Long
    
    If GdipCreatePen1(ConvertColor(nColor, mOpacity), nDrawWidth, UnitPixel, hPen) = 0 Then
        If ((mBorderStyle = vbBSSolid) Or (mBorderStyle = vbBSInsideSolid)) Or (mBorderWidth > 1) Then
            Call GdipSetSmoothingMode(nGraphics, SmoothingMode)
        Else
            Call GdipSetSmoothingMode(nGraphics, QualityModeLow)
        End If
        If ((mBorderStyle > vbBSSolid) And (mBorderStyle < vbBSInsideSolid)) Then
            Call GdipSetPenDashStyle(hPen, mBorderStyle - 1)
        End If
        If mCurvingFactor = 0 Then
            GdipDrawPolygonI nGraphics, hPen, Points(0), UBound(Points) + 1
        Else
            GdipDrawClosedCurve2I nGraphics, hPen, Points(0), UBound(Points) + 1, mCurvingFactor2
        End If
        Call GdipDeletePen(hPen)
    End If
    
End Sub

Private Sub FillClosedCurve(ByVal nGraphics As Long, ByVal nColor As Long, Points() As POINTL, ByVal nTension As Single, Optional nFillMode As FillModeConstants = FillModeAlternate)
    Dim hBrush As Long
    Dim iRet As Long
    Dim iPoints() As POINTL
    Dim iPath As Long
    Dim iStyle3DEffect As Long
    Dim iRect As RECT
    
    If mTextureBrush = 0 Then
        If mStyle3D <> 0 Then
            If mStyle3DEffect = seStyle3EffectAuto Then
                If (mShape = seShapeCrescent) Or (mShape = seShapeLocation) Or (mShape = seShapeCloud) Then
                    iStyle3DEffect = seStyle3EffectDiffuse
                Else
                    iStyle3DEffect = seStyle3EffectGem
                End If
            Else
                iStyle3DEffect = mStyle3DEffect
            End If
            
            If iStyle3DEffect = seStyle3EffectDiffuse Then
                GdipCreatePath 0&, iPath
                iRect = ScaleRect(GetPointsLRect(Points), Sqr(2) * (1 + Abs(mCurvingFactor) / 400))
                GdipAddPathEllipseI iPath, iRect.Left, iRect.Top, iRect.Right - iRect.Left, iRect.Bottom - iRect.Top
                iRet = GdipCreatePathGradientFromPath(iPath, hBrush)
            Else
                iPoints = ExpandPointsL(Points, 0.05)
                If mCurvingFactor <> 0 Then
                    iPoints = ExpandPointsL(iPoints, Abs(mCurvingFactor) / 80)
                End If
                iRet = GdipCreatePathGradientI(iPoints(0), UBound(Points) + 1, 0&, hBrush)
            End If
            If iRet = 0 Then
                GdipSetPathGradientCenterColor hBrush, ConvertColor(ShiftColor(nColor, vbWhite, IIf(mStyle3D And seStyle3DLight, 200, 255)), mOpacity)
                GdipSetPathGradientSurroundColorsWithCount hBrush, ConvertColor(ShiftColor(nColor, vbBlack, IIf(mStyle3D And seStyle3DShadow, 200, 255)), mOpacity), 1
            End If
        Else
            iRet = GdipCreateSolidFill(ConvertColor(nColor, mOpacity), hBrush)
        End If
    End If
    
    If iRet = 0 Then
        Call GdipSetSmoothingMode(nGraphics, SmoothingMode)
        GdipFillClosedCurve2I nGraphics, IIf(mTextureBrush <> 0, mTextureBrush, hBrush), Points(0), UBound(Points) + 1, nTension, nFillMode
        Call GdipDeleteBrush(hBrush)
        If iPath <> 0 Then
            GdipDeletePath iPath
        End If
    End If
    
End Sub

Private Sub DrawClosedCurve(ByVal nGraphics As Long, ByVal nColor As Long, ByVal nDrawWidth As Long, Points() As POINTL, ByVal nTension As Single)
    Dim hPen As Long
    
    If GdipCreatePen1(ConvertColor(nColor, mOpacity), nDrawWidth, UnitPixel, hPen) = 0 Then
        Call GdipSetSmoothingMode(nGraphics, SmoothingMode)
        If ((mBorderStyle > vbBSSolid) And (mBorderStyle < vbBSInsideSolid)) Then
            Call GdipSetPenDashStyle(hPen, mBorderStyle - 1)
        End If
        GdipDrawClosedCurve2I nGraphics, hPen, Points(0), UBound(Points) + 1, nTension
        Call GdipDeletePen(hPen)
    End If
    
End Sub


Private Sub FillEllipse(ByVal nGraphics As Long, ByVal nColor As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long)
    Dim hBrush As Long
    Dim iRet As Long
    Dim iStyle3DEffect As Long
    Dim iPath As Long
    Dim iRect As RECT
    Dim iPoints(3) As POINTL
    
    If mTextureBrush = 0 Then
        If mStyle3D <> 0 Then
            If mStyle3DEffect = seStyle3EffectAuto Then
                iStyle3DEffect = seStyle3EffectDiffuse
            Else
                iStyle3DEffect = mStyle3DEffect
            End If
            
            If iStyle3DEffect = seStyle3EffectDiffuse Then
                GdipCreatePath 0&, iPath
                GdipAddPathEllipseI iPath, X, Y, nWidth, nHeight
                iRet = GdipCreatePathGradientFromPath(iPath, hBrush)
            Else
                iPoints(0).X = X
                iPoints(0).Y = Y
                iPoints(1).X = X + nWidth
                iPoints(1).Y = Y
                iPoints(2).X = X + nWidth
                iPoints(2).Y = Y + nHeight
                iPoints(3).X = X
                iPoints(3).Y = Y + nHeight
                iRet = GdipCreatePathGradientI(iPoints(0), UBound(iPoints) + 1, 0&, hBrush)
            End If
            If iRet = 0 Then
                GdipSetPathGradientCenterColor hBrush, ConvertColor(ShiftColor(nColor, vbWhite, IIf(mStyle3D And seStyle3DLight, 200, 255)), mOpacity)
                GdipSetPathGradientSurroundColorsWithCount hBrush, ConvertColor(ShiftColor(nColor, vbBlack, IIf(mStyle3D And seStyle3DShadow, 200, 255)), mOpacity), 1
            End If
        Else
            iRet = GdipCreateSolidFill(ConvertColor(nColor, mOpacity), hBrush)
        End If
    End If
    
    If iRet = 0 Then
        Call GdipSetSmoothingMode(nGraphics, SmoothingMode)
        GdipFillEllipseI nGraphics, IIf(mTextureBrush <> 0, mTextureBrush, hBrush), X, Y, nWidth, nHeight
        Call GdipDeleteBrush(hBrush)
        If iPath <> 0 Then
            GdipDeletePath iPath
        End If
    End If
End Sub

Private Sub DrawEllipse(ByVal nGraphics As Long, ByVal nColor As Long, ByVal nDrawWidth As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long)
    Dim hPen As Long
    
    If GdipCreatePen1(ConvertColor(nColor, mOpacity), nDrawWidth, UnitPixel, hPen) = 0 Then
        Call GdipSetSmoothingMode(nGraphics, SmoothingMode)
        If ((mBorderStyle > vbBSSolid) And (mBorderStyle < vbBSInsideSolid)) Then
            Call GdipSetPenDashStyle(hPen, mBorderStyle - 1)
        End If
        If mBorderStyle = vbBSInsideSolid Then
            X = X + nDrawWidth / 2
            Y = Y + nDrawWidth / 2
            nWidth = nWidth - nDrawWidth
            nHeight = nHeight - nDrawWidth
        End If
        GdipDrawEllipseI nGraphics, hPen, X, Y, nWidth, nHeight
        Call GdipDeletePen(hPen)
    End If
    
End Sub

Private Sub FillRoundRect(ByVal nGraphics As Long, ByVal nColor As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, Optional ByVal nRoundSize As Long = 10)
    Dim hBrush As Long
    Dim iRet As Long
    Dim iStyle3DEffect As Long
    Dim iPath As Long
    Dim iRect As RECT
    Dim iPoints() As POINTL
    
    nRoundSize = nRoundSize * 2
    If mTextureBrush = 0 Then
        If mStyle3D <> 0 Then
            If mStyle3DEffect = seStyle3EffectAuto Then
                iStyle3DEffect = seStyle3EffectDiffuse
            Else
                iStyle3DEffect = mStyle3DEffect
            End If
    
            ReDim iPoints(3)
            iPoints(0).X = X
            iPoints(0).Y = Y
            iPoints(1).X = X + nWidth
            iPoints(1).Y = Y
            iPoints(2).X = X + nWidth
            iPoints(2).Y = Y + nHeight
            iPoints(3).X = X
            iPoints(3).Y = Y + nHeight
            If iStyle3DEffect = seStyle3EffectDiffuse Then
                GdipCreatePath 0&, iPath
                iRect = ScaleRect(GetPointsLRect(iPoints), Sqr(2) * (1 + Abs(mCurvingFactor) / 400))
                GdipAddPathEllipseI iPath, iRect.Left, iRect.Top, iRect.Right - iRect.Left, iRect.Bottom - iRect.Top
                iRet = GdipCreatePathGradientFromPath(iPath, hBrush)
            Else
                If mCurvingFactor <> 0 Then
                    iPoints = ExpandPointsL(iPoints, Abs(mCurvingFactor) / 80)
                End If
                iRet = GdipCreatePathGradientI(iPoints(0), UBound(iPoints) + 1, 0&, hBrush)
            End If
            If iRet = 0 Then
                GdipSetPathGradientCenterColor hBrush, ConvertColor(ShiftColor(nColor, vbWhite, IIf(mStyle3D And seStyle3DLight, 200, 255)), mOpacity)
                GdipSetPathGradientSurroundColorsWithCount hBrush, ConvertColor(ShiftColor(nColor, vbBlack, IIf(mStyle3D And seStyle3DShadow, 200, 255)), mOpacity), 1
            End If
        Else
            iRet = GdipCreateSolidFill(ConvertColor(nColor, mOpacity), hBrush)
        End If
    End If
    
    If iRet = 0 Then
        GdipCreatePath &H0, iPath
        GdipAddPathArcI iPath, X, Y, nRoundSize, nRoundSize, 180, 90
        GdipAddPathArcI iPath, X + nWidth - nRoundSize, Y, nRoundSize, nRoundSize, 270, 90
        GdipAddPathArcI iPath, X + nWidth - nRoundSize, Y + nHeight - nRoundSize, nRoundSize, nRoundSize, 0, 90
        GdipAddPathArcI iPath, X, Y + nHeight - nRoundSize, nRoundSize, nRoundSize, 90, 90
        GdipClosePathFigure iPath
        GdipFillPath nGraphics, IIf(mTextureBrush <> 0, mTextureBrush, hBrush), iPath
        
        Call GdipDeletePath(iPath)
        Call GdipDeleteBrush(hBrush)
    End If
End Sub

Private Sub DrawRoundRect(ByVal nGraphics As Long, ByVal nColor As Long, ByVal nDrawWidth As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, Optional ByVal nRoundSize As Long = 10)
    Dim iPath As Long
    Dim hPen As Long
    
    nRoundSize = nRoundSize * 2
    If GdipCreatePen1(ConvertColor(nColor, mOpacity), nDrawWidth, UnitPixel, hPen) = 0 Then
        If ((mBorderStyle = vbBSSolid) Or (mBorderStyle = vbBSInsideSolid)) Or (mBorderWidth > 1) Then
            Call GdipSetSmoothingMode(nGraphics, SmoothingMode)
        Else
            Call GdipSetSmoothingMode(nGraphics, QualityModeLow)
        End If
        If ((mBorderStyle > vbBSSolid) And (mBorderStyle < vbBSInsideSolid)) Then
            Call GdipSetPenDashStyle(hPen, mBorderStyle - 1)
        End If
        If mBorderStyle = vbBSInsideSolid Then
            X = X + nDrawWidth / 2
            Y = Y + nDrawWidth / 2
            nWidth = nWidth - nDrawWidth
            nHeight = nHeight - nDrawWidth
        End If

        GdipCreatePath &H0, iPath
        GdipAddPathArcI iPath, X, Y, nRoundSize, nRoundSize, 180, 90
        GdipAddPathArcI iPath, X + nWidth - nRoundSize, Y, nRoundSize, nRoundSize, 270, 90
        GdipAddPathArcI iPath, X + nWidth - nRoundSize, Y + nHeight - nRoundSize, nRoundSize, nRoundSize, 0, 90
        GdipAddPathArcI iPath, X, Y + nHeight - nRoundSize, nRoundSize, nRoundSize, 90, 90
        GdipClosePathFigure iPath
        GdipDrawPath nGraphics, hPen, iPath
        
        Call GdipDeletePath(iPath)
        Call GdipDeletePen(hPen)
    End If
End Sub

Private Sub FillSemicircle(ByVal nGraphics As Long, ByVal nColor As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long)
    Dim hBrush As Long
    Dim iRet As Long
    Dim iStyle3DEffect As Long
    Dim iPath As Long
    Dim iRect As RECT
    Dim iPoints() As POINTL
    
    If mTextureBrush = 0 Then
        If mStyle3D <> 0 Then
            If mStyle3DEffect = seStyle3EffectAuto Then
                iStyle3DEffect = seStyle3EffectDiffuse
            Else
                iStyle3DEffect = mStyle3DEffect
            End If
            
            ReDim iPoints(3)
            iPoints(0).X = X
            iPoints(0).Y = Y
            iPoints(1).X = X + nWidth
            iPoints(1).Y = Y
            iPoints(2).X = X + nWidth
            iPoints(2).Y = Y + nHeight
            iPoints(3).X = X
            iPoints(3).Y = Y + nHeight
            If iStyle3DEffect = seStyle3EffectDiffuse Then
                GdipCreatePath 0&, iPath
                iRect = ScaleRect(GetPointsLRect(iPoints), Sqr(2) * (1 + Abs(mCurvingFactor) / 400))
                GdipAddPathEllipseI iPath, iRect.Left, iRect.Top, iRect.Right - iRect.Left, iRect.Bottom - iRect.Top
                iRet = GdipCreatePathGradientFromPath(iPath, hBrush)
            Else
                If mCurvingFactor <> 0 Then
                    iPoints = ExpandPointsL(iPoints, Abs(mCurvingFactor) / 80)
                End If
                iRet = GdipCreatePathGradientI(iPoints(0), UBound(iPoints) + 1, 0&, hBrush)
            End If
            If iRet = 0 Then
                GdipSetPathGradientCenterColor hBrush, ConvertColor(ShiftColor(nColor, vbWhite, IIf(mStyle3D And seStyle3DLight, 200, 255)), mOpacity)
                GdipSetPathGradientSurroundColorsWithCount hBrush, ConvertColor(ShiftColor(nColor, vbBlack, IIf(mStyle3D And seStyle3DShadow, 200, 255)), mOpacity), 1
            End If
        Else
            iRet = GdipCreateSolidFill(ConvertColor(nColor, mOpacity), hBrush)
        End If
    End If
    
    If iRet = 0 Then
        Call GdipSetSmoothingMode(nGraphics, SmoothingMode)
        GdipFillPieI nGraphics, IIf(mTextureBrush <> 0, mTextureBrush, hBrush), X, Y, nWidth, nHeight * 2, 180, 180
        Call GdipDeleteBrush(hBrush)
        If iPath <> 0 Then
            GdipDeletePath iPath
        End If
    End If
End Sub

Private Sub DrawSemicircle(ByVal nGraphics As Long, ByVal nColor As Long, ByVal nDrawWidth As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long)
    Dim hPen As Long
    
    If GdipCreatePen1(ConvertColor(nColor, mOpacity), nDrawWidth, UnitPixel, hPen) = 0 Then
        If ((mBorderStyle = vbBSSolid) Or (mBorderStyle = vbBSInsideSolid)) Or (mBorderWidth > 1) Then
            Call GdipSetSmoothingMode(nGraphics, SmoothingMode)
        Else
            Call GdipSetSmoothingMode(nGraphics, QualityModeLow)
        End If
        If ((mBorderStyle > vbBSSolid) And (mBorderStyle < vbBSInsideSolid)) Then
            Call GdipSetPenDashStyle(hPen, mBorderStyle - 1)
        End If
        If mBorderStyle = vbBSInsideSolid Then
            X = X + nDrawWidth / 2
            Y = Y + nDrawWidth / 2
            nWidth = nWidth - nDrawWidth
            nHeight = nHeight - nDrawWidth
        End If
        GdipDrawArcI nGraphics, hPen, X, Y, nWidth, nHeight * 2 + nDrawWidth, 180, 180
        GdipDrawLineI nGraphics, hPen, X + nDrawWidth / 2 - 1, Y + nHeight, X + nWidth - nDrawWidth / 2 + 1, Y + nHeight
        Call GdipDeletePen(hPen)
    End If
    
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

Private Sub InitGDI()
    Dim GdipStartupInput As GDIPlusStartupInput
    GdipStartupInput.GdiPlusVersion = 1&
    Call GdiplusStartup(mGdipToken, GdipStartupInput, ByVal 0)
End Sub

Private Sub TerminateGDI()
    If mTextureBrush <> 0 Then DestroyTextureBrush
    Call GdiplusShutdown(mGdipToken)
    mGdipToken = 0
End Sub

Private Property Get SmoothingMode() As Long
    If mQuality = seQualityHigh Then
        SmoothingMode = SmoothingModeAntiAlias
    Else
        SmoothingMode = QualityModeLow
    End If
End Property


Private Sub Subclass()
    Dim iDo As Boolean
    
    If mContainerHwnd <> 0 Then
        If mUseSubclassing = seSCYes Then
           iDo = True
        ElseIf mUseSubclassing = seSCNotInIDE Then
            iDo = Not InIDE
        ElseIf mUseSubclassing = seSCNotInIDEDesignTime Then
            If mUserMode Then
                iDo = True
            Else
                iDo = Not InIDE
            End If
        End If
        If iDo Then
            AttachMessage Me, mContainerHwnd, WM_INVALIDATE
            mSubclassed = True
        End If
    End If
End Sub

Private Function InIDE() As Boolean
    Static sValue As Long
    
    If sValue = 0 Then
        Err.Clear
        On Error Resume Next
        Debug.Print 1 / 0
        If Err.Number Then
            sValue = 1
        Else
            sValue = 2
        End If
        Err.Clear
    End If
    InIDE = (sValue = 1)
End Function

Private Sub Unsubclass()
    If mSubclassed Then
        DetachMessage Me, mContainerHwnd, WM_INVALIDATE
        mSubclassed = False
    End If
End Sub

Private Sub SetCurvingFactor2()
    If mCurvingFactor < 0 Then
        mCurvingFactor2 = mCurvingFactor / 100 * 0.5
    Else
        mCurvingFactor2 = mCurvingFactor / 100 * 1
    End If
End Sub

' From Leandro Ascierto
Private Function ShiftColor(ByVal clrFirst As Long, ByVal clrSecond As Long, ByVal lAlpha As Long) As Long
    Dim clrFore(3)         As Byte
    Dim clrBack(3)         As Byte
 
    OleTranslateColor clrFirst, 0, VarPtr(clrFore(0))
    OleTranslateColor clrSecond, 0, VarPtr(clrBack(0))
    
    clrFore(0) = (clrFore(0) * lAlpha + clrBack(0) * (255 - lAlpha)) / 255
    clrFore(1) = (clrFore(1) * lAlpha + clrBack(1) * (255 - lAlpha)) / 255
    clrFore(2) = (clrFore(2) * lAlpha + clrBack(2) * (255 - lAlpha)) / 255
     
    CopyMemory ShiftColor, clrFore(0), 4
End Function

Private Function ExpandPointsL(nPoints() As POINTL, nExpand As Single) As POINTL()
    Dim iCount As Long
    Dim c As Long
    Dim iRect As RECT
    Dim iCenterX As Long
    Dim iCenterY As Long
    Dim iRet() As POINTL
    
    iRect.Top = 0
    iRect.Bottom = 0
    iRect.Left = UserControl.ScaleWidth * 2
    iRect.Top = UserControl.ScaleHeight * 2
    iCount = UBound(nPoints) + 1
    For c = 0 To iCount - 1
        If nPoints(c).X < iRect.Left Then iRect.Left = nPoints(c).X
        If nPoints(c).X > iRect.Right Then iRect.Right = nPoints(c).X
        If nPoints(c).Y < iRect.Top Then iRect.Top = nPoints(c).Y
        If nPoints(c).Y > iRect.Bottom Then iRect.Bottom = nPoints(c).Y
    Next
    iCenterX = (iRect.Left + iRect.Right) / 2
    iCenterY = (iRect.Top + iRect.Bottom) / 2
    ReDim iRet(iCount - 1)
    For c = 0 To iCount - 1
        If nPoints(c).X > iCenterX Then
            iRet(c).X = nPoints(c).X + Abs(iCenterX - nPoints(c).X) * nExpand
        ElseIf nPoints(c).X < iCenterX Then
            iRet(c).X = nPoints(c).X - Abs(iCenterX - nPoints(c).X) * nExpand
        Else
            iRet(c).X = nPoints(c).X
        End If
        If nPoints(c).Y > iCenterY Then
            iRet(c).Y = nPoints(c).Y + Abs(iCenterY - nPoints(c).Y) * nExpand
        ElseIf nPoints(c).Y < iCenterY Then
            iRet(c).Y = nPoints(c).Y - Abs(iCenterY - nPoints(c).Y) * nExpand
        Else
            iRet(c).Y = nPoints(c).Y
        End If
    Next
    ExpandPointsL = iRet
End Function

Private Function GetPointsLRect(nPoints() As POINTL) As RECT
    Dim iCount As Long
    Dim c As Long
    Dim iRect As RECT
    
    iRect.Top = 0
    iRect.Bottom = 0
    iRect.Left = UserControl.ScaleWidth * 2
    iRect.Top = UserControl.ScaleHeight * 2
    iCount = UBound(nPoints) + 1
    For c = 0 To iCount - 1
        If nPoints(c).X < iRect.Left Then iRect.Left = nPoints(c).X
        If nPoints(c).X > iRect.Right Then iRect.Right = nPoints(c).X
        If nPoints(c).Y < iRect.Top Then iRect.Top = nPoints(c).Y
        If nPoints(c).Y > iRect.Bottom Then iRect.Bottom = nPoints(c).Y
    Next
    GetPointsLRect = iRect
End Function

Private Function ScaleRect(nRect As RECT, ByVal nScale As Single) As RECT
    Dim iRect As RECT
    
    nScale = (nScale - 1) / 2
    
    iRect = nRect
    InflateRect iRect, (nRect.Right - nRect.Left) * nScale, (nRect.Bottom - nRect.Top) * nScale
    ScaleRect = iRect
End Function

Private Sub DrawPie(ByVal nGraphics As Long, ByVal nColor As Long, ByVal nDrawWidth As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal nAngleMissingPart As Single)
    Dim hPen As Long
    
    If GdipCreatePen1(ConvertColor(nColor, mOpacity), nDrawWidth, UnitPixel, hPen) = 0 Then
        If ((mBorderStyle = vbBSSolid) Or (mBorderStyle = vbBSInsideSolid)) Or (mBorderWidth > 1) Then
            Call GdipSetSmoothingMode(nGraphics, SmoothingMode)
        Else
            Call GdipSetSmoothingMode(nGraphics, QualityModeLow)
        End If
        If ((mBorderStyle > vbBSSolid) And (mBorderStyle < vbBSInsideSolid)) Then
            Call GdipSetPenDashStyle(hPen, mBorderStyle - 1)
        End If
        If mBorderStyle = vbBSInsideSolid Then
            X = X + nDrawWidth / 2
            Y = Y + nDrawWidth / 2
            nWidth = nWidth - nDrawWidth
            nHeight = nHeight - nDrawWidth
        End If
        
        nAngleMissingPart = Abs(nAngleMissingPart)
        If nAngleMissingPart > 360 Then nAngleMissingPart = 360
        GdipDrawPieI nGraphics, hPen, X, Y, nWidth, nHeight, 0, 360 - nAngleMissingPart
        
        Call GdipDeletePen(hPen)
    End If
End Sub

Private Sub FillPie(ByVal nGraphics As Long, ByVal nColor As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal nAngleMissingPart As Single)
    Dim hBrush As Long
    Dim iRet As Long
    Dim iStyle3DEffect As Long
    Dim iPath As Long
    Dim iRect As RECT
    Dim iPoints() As POINTL
    
    If mTextureBrush = 0 Then
        If mStyle3D <> 0 Then
            If mStyle3DEffect = seStyle3EffectAuto Then
                iStyle3DEffect = seStyle3EffectDiffuse
            Else
                iStyle3DEffect = mStyle3DEffect
            End If
            
            ReDim iPoints(3)
            iPoints(0).X = X
            iPoints(0).Y = Y
            iPoints(1).X = X + nWidth
            iPoints(1).Y = Y
            iPoints(2).X = X + nWidth
            iPoints(2).Y = Y + nHeight
            iPoints(3).X = X
            iPoints(3).Y = Y + nHeight
            If iStyle3DEffect = seStyle3EffectDiffuse Then
                GdipCreatePath 0&, iPath
                iRect = ScaleRect(GetPointsLRect(iPoints), Sqr(2) * (1 + Abs(mCurvingFactor) / 400))
                GdipAddPathEllipseI iPath, iRect.Left, iRect.Top, iRect.Right - iRect.Left, iRect.Bottom - iRect.Top
                iRet = GdipCreatePathGradientFromPath(iPath, hBrush)
            Else
                If mCurvingFactor <> 0 Then
                    iPoints = ExpandPointsL(iPoints, Abs(mCurvingFactor) / 80)
                End If
                iRet = GdipCreatePathGradientI(iPoints(0), UBound(iPoints) + 1, 0&, hBrush)
            End If
            If iRet = 0 Then
                GdipSetPathGradientCenterColor hBrush, ConvertColor(ShiftColor(nColor, vbWhite, IIf(mStyle3D And seStyle3DLight, 200, 255)), mOpacity)
                GdipSetPathGradientSurroundColorsWithCount hBrush, ConvertColor(ShiftColor(nColor, vbBlack, IIf(mStyle3D And seStyle3DShadow, 200, 255)), mOpacity), 1
            End If
        Else
            iRet = GdipCreateSolidFill(ConvertColor(nColor, mOpacity), hBrush)
        End If
    End If
    
    If iRet = 0 Then
        Call GdipSetSmoothingMode(nGraphics, SmoothingMode)
        nAngleMissingPart = Abs(nAngleMissingPart)
        If nAngleMissingPart > 360 Then nAngleMissingPart = 360
        GdipFillPieI nGraphics, IIf(mTextureBrush <> 0, mTextureBrush, hBrush), X, Y, nWidth, nHeight, 0, 360 - nAngleMissingPart
        Call GdipDeleteBrush(hBrush)
        If iPath <> 0 Then
            GdipDeletePath iPath
        End If
    End If
End Sub

Private Function IsValidOLE_COLOR(ByVal nColor As Long) As Boolean
    Const S_OK As Long = 0
    IsValidOLE_COLOR = (TranslateColor(nColor, 0, nColor) = S_OK)
End Function

Private Sub CreateTextureBrush()
    Dim hImg As Long
    Dim hBrush As Long
    
    If mFillTexture Is Nothing Then Exit Sub
    If mFillTexture.Handle = 0 Then Exit Sub
    If mTextureBrush <> 0 Then DestroyTextureBrush
    If GdipCreateBitmapFromHBITMAP(mFillTexture.Handle, 0&, hImg) = 0 Then
        GdipCreateTexture hImg, WrapModeTile, mTextureBrush
        GdipDisposeImage hImg
    End If
End Sub

Private Sub DestroyTextureBrush()
    If mTextureBrush <> 0 Then
        GdipDeleteBrush (mTextureBrush)
        mTextureBrush = 0
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

Public Sub SetFocus()
    Extender.SetFocus
End Sub

Public Sub ZOrder(Optional ByRef Position As Variant)
    If IsMissing(Position) Then Extender.ZOrder Else Extender.ZOrder Position
End Sub

