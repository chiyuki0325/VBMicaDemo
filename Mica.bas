Attribute VB_Name = "Mica"
Option Explicit

Public Function GetWallpaperPath() As String
    With New WshShell
        GetWallpaperPath = .RegRead("HKCU\Control Panel\Desktop\Wallpaper")
    End With
End Function

Public Sub CalculateMica(ByVal PixelStep As Integer, ByRef Pixels As Integer, ByRef OutputR As Integer, OutputG As Integer, OutputB As Integer)
    On Error Resume Next
    Dim img As New WIA.ImageFile
    img.LoadFile GetWallpaperPath
    
    Dim PixelDelta As Integer
    PixelDelta = PixelStep * 2 - 1
    
    Dim Width As Integer, Height As Integer
    Width = img.Width
    Height = img.Height
    
    Dim R() As Integer, G() As Integer, B() As Integer
    ReDim R(0 To (Width / PixelStep * Height / PixelStep))
    ReDim G(0 To (Width / PixelStep * Height / PixelStep))
    ReDim B(0 To (Width / PixelStep * Height / PixelStep))
    
    Dim X As Long, Y As Long, ARGB As Long, Count As Long, Color As Long
    Dim R2 As Integer, G2 As Integer, B2 As Integer, H2 As Integer
    Dim R3 As Byte, G3 As Byte, B3 As Byte
    Count = 0
    For X = 1 To Width / PixelStep
        For Y = 1 To Height / PixelStep
            Color = img.ARGBData((X * PixelStep) + (Y * PixelStep) * Width - PixelDelta)
            R(Count) = ((Color And &HFF0000) \ &H10000)
            G(Count) = ((Color And &HFF00&) \ &H100)
            B(Count) = (Color And &HFF)
            'Debug.Print R(Count) & " " & G(Count) & " " & B(Count)
            Count = Count + 1
            'DoEvents
        Next
    Next
    
    R2 = AverageInt(R)
    G2 = AverageInt(G)
    B2 = AverageInt(B)
    H2 = GetHueFromRGB(R2, G2, B2)
    'MsgBox "取到的颜色 RGB：" & R2 & " " & G2 & " " & B2
    HSVtoRGB H2, CByte(8), CByte(255), R3, G3, B3
    OutputR = R3
    OutputG = G3
    OutputB = B3
    Pixels = UBound(R) + 1
End Sub

Private Function AverageInt(InputArray() As Integer) As Integer
    Dim TempAvg As Integer, Count As Long
    TempAvg = 0
    AverageInt = 0
    For Count = 0 To UBound(InputArray)
        TempAvg = TempAvg + InputArray(Count)
        DoEvents
    Next
    AverageInt = Int(TempAvg / (UBound(InputArray)))
End Function

Private Function GetHueFromRGB(ByVal R As Byte, ByVal G As Byte, ByVal B As Byte) As Integer
Dim H As Byte
Dim MinVal As Byte
Dim MaxVal As Byte
Dim Chroma As Byte
Dim TempH As Single
If R > G Then MaxVal = R Else MaxVal = G
If B > MaxVal Then MaxVal = B
If R < G Then MinVal = R Else MinVal = G
If B < MinVal Then MinVal = B
Chroma = MaxVal - MinVal

If Chroma = 0 Then
    H = 0
Else
    Select Case MaxVal
        Case R
            TempH = (1& * G - B) / Chroma
            If TempH < 0 Then TempH = TempH + 6
            H = TempH / 6 * 255
        Case G
            H = (((1& * B - R) / Chroma) + 2) / 6 * 255
        Case B
            H = (((1& * R - G) / Chroma) + 4) / 6 * 255
    End Select
End If
GetHueFromRGB = H
End Function
                    
Private Sub HSVtoRGB(ByVal H As Byte, ByVal S As Byte, ByVal V As Byte, _
                    ByRef R As Byte, ByRef G As Byte, ByRef B As Byte)
Dim MinVal As Byte
Dim MaxVal As Byte
Dim Chroma As Byte
Dim TempH As Single

If V = 0 Then
    R = 0
    G = 0
    B = 0
Else
    If S = 0 Then
        R = V
        G = V
        B = V
    Else
        MaxVal = V
        Chroma = S / 255 * MaxVal
        MinVal = MaxVal - Chroma
        Select Case H
            Case Is >= 170
                TempH = (H - 170) / 43
                If TempH < 1 Then
                    B = MaxVal
                    R = MaxVal * TempH
                Else
                    R = MaxVal
                    B = MaxVal * (2 - TempH)
                End If
                G = 0
            Case Is >= 85
                TempH = (H - 85) / 43
                If TempH < 1 Then
                    G = MaxVal
                    B = MaxVal * TempH
                Else
                    B = MaxVal
                    G = MaxVal * (2 - TempH)
                End If
                R = 0
            Case Else
                TempH = H / 43
                If TempH < 1 Then
                    R = MaxVal
                    G = MaxVal * TempH
                Else
                    G = MaxVal
                    R = MaxVal * (2 - TempH)
                End If
                B = 0
        End Select
        R = R / MaxVal * (MaxVal - MinVal) + MinVal
        G = G / MaxVal * (MaxVal - MinVal) + MinVal
        B = B / MaxVal * (MaxVal - MinVal) + MinVal
    End If
End If
End Sub



