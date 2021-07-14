Attribute VB_Name = "mdlShaded"
Option Explicit

'=========Gdi32 Api========
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GdiAlphaBlend Lib "gdi32" (ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal BLENDFUNCT As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'=========user32 Api========
Private Declare Function GetWindowRect Lib "User32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetDC Lib "User32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "User32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function SetParent Lib "User32" (ByVal hWndChild As Long, ByVal hWndNewparent As Long) As Long
Private Declare Function DrawText Lib "user32.dll" (ByVal hdc As Long, lpStr As Long, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long

'=========Oleaut32 Api========
Private Declare Function OleTranslateColor Lib "Olepro32" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, ByVal pccolorref As Long) As Long
  
'=========Kernel32 Api========
Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
  
' USER 32 WindowsHook (LeandroA)
Private Declare Function GetPropA Lib "User32" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SetPropA Lib "User32" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
 
Private Type UcsRgbQuad
    R                       As Byte
    G                       As Byte
    B                       As Byte
    a                       As Byte
End Type
 
Private Type BLENDFUNCTION
    BlendOp                 As Byte
    BlendFlags              As Byte
    SourceConstantAlpha     As Byte
    AlphaFormat             As Byte
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


Public Function DrawShaded(bDraw As Boolean, ByRef uControl As Object, ByRef fParent As Object, Optional ByVal lBlendColor As OLE_COLOR) As Boolean
    Dim zpBox  As PictureBox
    Dim lWidth   As Long
    Dim lHeight  As Long
    Dim Area As RECT

If bDraw = True Then
    Set zpBox = fParent.Controls.Add("VB.PictureBox", "zpBox", fParent)
    Call SetPropA(fParent.hwnd, "ModalPopPicture", zpBox.hwnd)
    Call GetWindowRect(fParent.hwnd, Area)
    
    With zpBox
        .Left = 0
        .Top = 0
        .AutoRedraw = True
        .BorderStyle = 0
        '// resize picture
        lWidth = (Area.Right - Area.Left)
        lHeight = (Area.Bottom - Area.Top)
        .Width = lWidth * Screen.TwipsPerPixelX
        .Height = lHeight * Screen.TwipsPerPixelY
        .ZOrder 0
        '// Grab a copy of the parent contents
        Call BitBlt(.hdc, 0, 0, lWidth, lHeight, fParent.hdc, 0, 0, vbSrcCopy)
        '// alphablend the image with selected color
        Call DrawAlphaSelection(.hdc, 0, 0, lWidth, lHeight, lBlendColor)
        .Visible = True
    End With
    
    SetParent uControl.hwnd, zpBox.hwnd
Else
    SetParent uControl.hwnd, fParent.hwnd

    If ObjectFromhWnd(GetPropA(fParent.hwnd, "ModalPopPicture"), zpBox, fParent) Then
      '// finally, remove it
       fParent.Controls.Remove zpBox
    End If

End If
End Function

Private Sub DrawAlphaSelection(ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, _
                               ByVal Height As Long, ByVal Color As OLE_COLOR)
    Dim BF                  As BLENDFUNCTION
    Dim hDCMemory           As Long
    Dim hBmp                As Long
    Dim hOldBmp             As Long
    Dim DC                  As Long
    Dim lColor              As Long
    Dim hPen                As Long
    Dim hBrush              As Long
    Dim lBF                 As Long
  
    BF.SourceConstantAlpha = 128
  
    DC = GetDC(0)
    hDCMemory = CreateCompatibleDC(0)
    hBmp = CreateCompatibleBitmap(DC, Width, Height)
    hOldBmp = SelectObject(hDCMemory, hBmp)
  
    hPen = CreatePen(0, 1, Color)
    hBrush = CreateSolidBrush(pvAlphaBlend(vbWhite, Color, 120))
    DeleteObject SelectObject(hDCMemory, hBrush)
    DeleteObject SelectObject(hDCMemory, hPen)
  
    CopyMemory VarPtr(lBF), VarPtr(BF), 4
    GdiAlphaBlend hdc, X, Y, Width, Height, hDCMemory, 0, 0, Width, Height, lBF
  
    SelectObject hDCMemory, hOldBmp
    DeleteObject hBmp
    ReleaseDC 0&, DC
    DeleteDC hDCMemory
    DeleteObject hPen
    DeleteObject hBrush
End Sub

Private Function ObjectFromhWnd(ByVal lhWnd As Long, ByRef oObject As Object, ByRef oContainer As Object) As Boolean
    For Each oObject In oContainer.Controls
        On Local Error Resume Next
        If oObject.hwnd Then
            If Err.Number = 0 Then
                If oObject.hwnd = lhWnd Then
                    ObjectFromhWnd = True
                    Exit Function
                End If
            Call Err.Clear
            End If
        End If
    Next
End Function

Private Function pvAlphaBlend(ByVal clrFirst As Long, ByVal clrSecond As Long, ByVal lAlpha As Long) As Long
    Dim clrFore             As UcsRgbQuad
    Dim clrBack             As UcsRgbQuad
  
    OleTranslateColor clrFirst, 0, VarPtr(clrFore)
    OleTranslateColor clrSecond, 0, VarPtr(clrBack)
    With clrFore
        .R = (.R * lAlpha + clrBack.R * (255 - lAlpha)) / 255
        .G = (.G * lAlpha + clrBack.G * (255 - lAlpha)) / 255
        .B = (.B * lAlpha + clrBack.B * (255 - lAlpha)) / 255
    End With
    CopyMemory VarPtr(pvAlphaBlend), VarPtr(clrFore), 4
End Function

'Public Function DrawShaded2(bDraw As Boolean, ByRef uControl As Object, ByRef fParent As Object, Optional ByVal lBlendColor As OLE_COLOR, Optional sMsg As String = "") As Boolean
'    Dim zpBox  As PictureBox
'    Dim zShp As Shape
'    Dim lWidth   As Long
'    Dim lHeight  As Long
'    Dim Area As RECT
'    Dim aCtrl As RECT
'
'If bDraw = True Then
'    Set zpBox = fParent.Controls.Add("VB.PictureBox", "zpBox", fParent)
'    Call SetPropA(fParent.hwnd, "ModalPopPicture", zpBox.hwnd)
'    Call GetWindowRect(fParent.hwnd, Area)
'
'    Set zShp = fParent.Controls.Add("VB.Shape", "zShp", zpBox)
'
'    With zpBox
'        .Left = 0
'        .Top = 0
'        .AutoRedraw = True
'        .BorderStyle = 0
'        '// resize picture
'        lWidth = (Area.Right - Area.Left)
'        lHeight = (Area.Bottom - Area.Top)
'        .Width = lWidth * Screen.TwipsPerPixelX
'        .Height = lHeight * Screen.TwipsPerPixelY
'        .ZOrder 0
'        '// Grab a copy of the parent contents
'        Call BitBlt(.hdc, 0, 0, lWidth, lHeight, fParent.hdc, 0, 0, vbSrcCopy)
'        '// alphablend the image with selected color
'        Call DrawAlphaSelection(.hdc, 0, 0, lWidth, lHeight, lBlendColor)
'        .Visible = True
'    End With
'
'    With zShp
'        .Left = uControl.Left - 5
'        .Top = uControl.Top - 5
'        .Width = uControl.Width + 10
'        .Height = uControl.Height + 10
'        .BorderWidth = 2
'        .BorderColor = vbRed
'        .Visible = True
'    End With
'
'    SetParent uControl.hwnd, zpBox.hwnd
'Else
'    SetParent uControl.hwnd, fParent.hwnd
'
'    If ObjectFromhWnd(GetPropA(fParent.hwnd, "ModalPopPicture"), zpBox, fParent) Then
'      '// finally, remove it
'       fParent.Controls.Remove zpBox
'       Set zShp = Nothing
'    End If
'
'End If
'End Function
