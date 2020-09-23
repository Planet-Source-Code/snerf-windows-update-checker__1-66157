VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SplitBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ======================================================================
' Class    : cSplitDDC
' Filename : cSplitDC.cls
' Author   : SP McMahon
' Date     : 07 July 1998
'
' A splitter class using the Desktop window to draw a
' splitter bar, therefore allowing splitting of MDI forms
' as well as standard forms.
' ======================================================================

' REMARKS
' This code was extract from vbaccelarator web site and modified
' to acomplish the requirements of this project
'
' Thanks! to SP McMahon
'
' Modifications
' -------------
'
' -  The splitter object (and further calls to it) was replaced
'    by Top and Left properties once the Userobject TabDockHost
'    is already the bounds for splitting actions
'
' -  The events FormMouseMove and FormMouseUp was put together
'    in the FormMouseDown Event to supply a immediate split
'    drawing for TabDock UserControl
'
' -  The Clipping area (MouseDown() event) is now handled by
'    the TabDockHost MouseDown() event once this event calculates
'    the rectangle of the docked window
'
' -  Removed Cursor clipping (1.6)
'
' -  Removed Border references. This project does not need (1.6)
'
' -  Window Rect was replace for a Rect passed as reference
'    which is the exact area we want to split (1.6)
'
' -  Created an offset based on start and end cursor position
' removed declarations to the modAPi32 for reducing project size
'*******************
Option Explicit

Private Const SPLITBAR_SIZE As Long = 2


Private Type POINTAPI
    x       As Long
    y       As Long
End Type

'// some module declarations
Private bDraw As Boolean
Private rcCurrent As RECT
Private rcNew As RECT
Private rcWindow As RECT
Private m_Offset As Long

Private Const VK_LBUTTON = &H1
'Private Const R2_XORPEN As Long = 7

'Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
'Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
'Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long

Private m_hWnd As Long
Private m_eOrientation As eOrientationConstants

Public Event AfterResize(ByVal newSize As Long)

Public Property Get Offset() As Long
    Offset = m_Offset
End Property

Public Property Get Orientation() As eOrientationConstants
    Orientation = m_eOrientation
End Property

Public Property Let Orientation(ByVal eOrientation As eOrientationConstants)
    m_eOrientation = eOrientation
End Property

Friend Sub SplitterMouseDown(ByVal hwnd As Long, rc As RECT, ByVal x As Long, ByVal y As Long)
    Dim tP As POINTAPI
    Dim tpPrev As POINTAPI
    Dim hDC As Long
    Dim hPen As Long
    
    m_hWnd = hwnd
    ' Send subsequent mouse messages to the owner window
    SetCapture m_hWnd
    ' get window rect
    rcWindow = rc

    GetCursorPos tP
    ' Store the initial cursor position
    tpPrev.x = tP.x
    tpPrev.y = tP.y
    
    If (m_eOrientation = espHorizontal) Then
        m_Offset = tP.y
    Else
        m_Offset = tP.x
    End If
    
    bDraw = True  ' start actual drawing from next move message
    rcCurrent.Left = 0: rcCurrent.Top = 0: rcCurrent.Right = 0: rcCurrent.Bottom = 0
    
    SplitterFormMouseMove tP.x, tP.y
    
    Do While GetKeyState(VK_LBUTTON) < 0
        GetCursorPos tP
        If tP.x <> tpPrev.x Or tP.y <> tpPrev.y Then
            tpPrev.x = tP.x
            tpPrev.y = tP.y
            SplitterFormMouseMove tP.x, tP.y
        End If
        DoEvents
    Loop
    
    SplitterFormMouseUp tP.x, tP.y

End Sub

' Changed to private
Private Sub SplitterFormMouseMove(ByVal x As Long, ByVal y As Long)
    Dim pt As POINTAPI
        
    If (bDraw) Then
        'Debug.Print "splitter move"
        DrawDragRect rcCurrent, 1
        ' It is simpler to use the mouse cursor position than try to translate
        ' X,Y to screen coordinates!
        GetCursorPos pt
        ' Determine where to draw the splitter:
        If (m_eOrientation = espHorizontal) Then
            rcNew.Left = rcWindow.Left
            rcNew.Right = rcWindow.Right
            If (pt.y >= rcWindow.Top) And (pt.y < rcWindow.Bottom) Then
                rcNew.Top = pt.y - SPLITBAR_SIZE
                rcNew.Bottom = pt.y + SPLITBAR_SIZE
            Else
                If (pt.y < rcWindow.Top) Then
                    rcNew.Top = rcWindow.Top - SPLITBAR_SIZE
                    rcNew.Bottom = rcNew.Top + SPLITBAR_SIZE
                Else
                    rcNew.Top = rcWindow.Bottom - SPLITBAR_SIZE
                    rcNew.Bottom = rcNew.Top + SPLITBAR_SIZE
                End If
            End If
        Else
            rcNew.Top = rcWindow.Top
            rcNew.Bottom = rcWindow.Bottom
            If (pt.x >= rcWindow.Left) And (pt.x <= rcWindow.Right) Then
                rcNew.Left = pt.x - SPLITBAR_SIZE
                rcNew.Right = pt.x + SPLITBAR_SIZE
            Else
                If (pt.x < rcWindow.Left) Then
                    rcNew.Left = rcWindow.Left - SPLITBAR_SIZE
                    rcNew.Right = rcNew.Left + SPLITBAR_SIZE
                End If
            End If
        End If
        
        ' Draw the new rectangle
        DrawDragRect rcNew, 1
        ' Store this position so we can erase it next time:
        LSet rcCurrent = rcNew
        ' Free the reference to the Desktop DC we got (make sure you do this!)
        'DeleteDC hDc
        'DeleteObject hPen
    End If
End Sub

' changed to private
Private Function SplitterFormMouseUp(ByVal x As Long, ByVal y As Long) As Boolean
    'Dim hDC As Long
    Dim tP As POINTAPI
    'Dim hWndClient As Long

    ' Release mouse capture:
    ReleaseCapture
    ' Don't leave orphaned rectangle on desktop; erase last rectangle.
    If (bDraw) Then
        bDraw = False
        DrawDragRect rcCurrent, 1
        
        GetCursorPos tP
        If (tP.x < rcWindow.Left) Then
            tP.x = rcWindow.Left
        End If
        If (tP.x > rcWindow.Right) Then
            tP.x = rcWindow.Right
        End If
        If (tP.y < rcWindow.Top) Then
            tP.y = rcWindow.Top
        End If
        If (tP.y > rcWindow.Bottom) Then
            tP.y = rcWindow.Bottom
        End If
        'Debug.Print "up:" & tP.y & ":" & tP.x
        
        ' Move the splitter to the validated final position:
        If (m_eOrientation = espHorizontal) Then
            m_Offset = (tP.y - m_Offset) * Screen.TwipsPerPixelY
        Else
            m_Offset = (tP.x - m_Offset) * Screen.TwipsPerPixelX
        End If
        ' Return true to tell the owner we have completed splitting:
        SplitterFormMouseUp = True
    End If
    
    If Orientation = espVertical Then
       'Debug.Print "left:" & rcNew.Left & "::" & rcNew.Right & "::" & m_Offset
       RaiseEvent AfterResize(m_Offset)
    Else
       'RaiseEvent AfterResize(rcNew.Top)
       RaiseEvent AfterResize(m_Offset)
    End If
    
End Function

Private Sub Class_Initialize()
   m_eOrientation = espVertical
End Sub

Private Sub DrawDragRect(rc As RECT, Optional ByVal Size As Long = 2)
        
   Dim DrawRect As RECT
   Dim hDC As Long
   Dim i As Long
        
   hDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)

   For i = 0 To Size
           
      With DrawRect
         .Top = rc.Top + i
         .Bottom = rc.Bottom - i
         .Left = rc.Left + i
         .Right = rc.Right - i
      End With
           
      DrawFocusRect hDC, DrawRect
           
   Next i
        
   DeleteDC hDC
        
End Sub

