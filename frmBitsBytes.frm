VERSION 5.00
Begin VB.Form frmBitsBytes 
   Caption         =   "Bits & Bytes, Oh My"
   ClientHeight    =   3705
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   ScaleHeight     =   3705
   ScaleWidth      =   9765
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   2415
      Left            =   4965
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "frmBitsBytes.frx":0000
      Top             =   1215
      Width           =   4695
   End
   Begin VB.ListBox List1 
      Height          =   2400
      ItemData        =   "frmBitsBytes.frx":0006
      Left            =   180
      List            =   "frmBitsBytes.frx":002E
      TabIndex        =   1
      Top             =   1200
      Width           =   4710
   End
   Begin VB.PictureBox Picture1 
      Height          =   765
      Left            =   180
      ScaleHeight     =   705
      ScaleWidth      =   9420
      TabIndex        =   0
      Top             =   150
      Width           =   9480
      Begin VB.Image Image1 
         Height          =   480
         Left            =   8355
         Picture         =   "frmBitsBytes.frx":01B7
         Top             =   90
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   270
      Index           =   1
      Left            =   210
      TabIndex        =   4
      Top             =   990
      Width           =   4620
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   270
      Index           =   0
      Left            =   4995
      TabIndex        =   3
      Top             =   990
      Width           =   4620
   End
End
Attribute VB_Name = "frmBitsBytes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (ByRef Destination As Any, ByVal Length As Long, ByVal Fill As Byte)

' all apis are used to help animate demonstrations or to asccess an icon's mask
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Private Declare Function GetIconInfo Lib "user32.dll" (ByVal hIcon As Long, ByRef piconinfo As ICONINFO) As Long
Private Type ICONINFO
    fIcon As Long
    xHotspot As Long
    yHotspot As Long
    hbmMask As Long
    hbmColor As Long
End Type
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetDIBits Lib "gdi32.dll" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, ByRef lpBits As Any, ByRef lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As Long
End Type


Private cBitsClass As cBits

Private Sub Form_Load()
    Picture1.ScaleMode = vbPixels
    Picture1.AutoRedraw = True
    
    Label1(0).Caption = "Other Functions included in the class:"
    Label1(1).Caption = "Examples"
    Text1.Text = "GetBit, SetBit" & vbCrLf & _
        "GetByte, SetByte" & vbCrLf & _
        "GetMidBits, SetMidBits" & vbCrLf & _
        "HighWord, LowWord" & vbCrLf & _
        "HighByte, LowByte" & vbCrLf & _
        "BytesToBitString, ByteArrayToBitString" & vbCrLf & _
        "StringToBytes" & vbCrLf & _
        "MakeWord, MakeDWord" & vbCrLf & _
        "CreateBitMask" & vbCrLf & _
        "CRC32 Byte,Integer,Long Arrays" & vbCrLf & _
        "And also all those used in the examples"
    
    Set cBitsClass = New cBits
    Picture1.CurrentX = 10: Picture1.CurrentY = 5
    Picture1.Print "Click on the listbox to display the example"
End Sub

Private Sub List1_Click()
    Dim l As Long, i As Integer, L2 As Long
    Dim bArray(0 To 11) As Byte ' 96 bits
    Dim bArrayLog(0 To 11) As Byte ' 96 bits, Logical fill
    
    Select Case List1.ListIndex
    Case -1
    Case 0, 1: 'shift Left
        If List1.ListIndex = 1 Then l = vbYellow Else l = -2078211838
        Picture1.Cls
        Picture1.CurrentX = 10: Picture1.CurrentY = 5
        Picture1.Print cBitsClass.BytesToBitString(l), "&H" & Hex(l)
        Picture1.Refresh
        Sleep 500
        For i = 0 To 31 * List1.ListIndex + 32
            cBitsClass.ShiftBitsLeft_Long l, 1, List1.ListIndex
            Picture1.Cls
            Picture1.CurrentX = 10: Picture1.CurrentY = 5
            Picture1.Print cBitsClass.BytesToBitString(l), "&H" & Hex(l)
            Picture1.Refresh
            Sleep 120
        Next
    Case 2: 'shift Right, no wrapping
        l = 2078211838
        L2 = -l
        Picture1.Cls
        Picture1.CurrentX = 5: Picture1.CurrentY = 5
        Picture1.Print "+  ";
        Picture1.Print cBitsClass.BytesToBitString(l); "     "; l
        Picture1.CurrentX = 5
        Picture1.Print "-   ";
        Picture1.Print cBitsClass.BytesToBitString(L2); "    "; L2
        Picture1.Refresh
        Sleep 500
        For i = 1 To 31
            cBitsClass.ShiftBitsRight_Long l, 1
            cBitsClass.ShiftBitsRight_Long L2, 1
            Picture1.Cls
            Picture1.CurrentX = 5: Picture1.CurrentY = 5
            Picture1.Print "+  ";
            Picture1.Print cBitsClass.BytesToBitString(l); "    "; l
            Picture1.CurrentX = 5
            Picture1.Print "-   ";
            Picture1.Print cBitsClass.BytesToBitString(L2); "    "; L2
            Picture1.Refresh
            Sleep 120
        Next
    
    Case 3: 'shift Right, wrapping
        l = &HFFFF0000
        Picture1.Cls
        Picture1.CurrentX = 10: Picture1.CurrentY = 5
        Picture1.Print cBitsClass.BytesToBitString(l), "&H" & Hex(l)
        Picture1.Refresh
        Sleep 500
        For i = 0 To 63
            cBitsClass.ShiftBitsRight_Long l, 1, bitShift_Wrap
            Picture1.Cls
            Picture1.CurrentX = 10: Picture1.CurrentY = 5
            Picture1.Print cBitsClass.BytesToBitString(l), "&H" & Hex(l)
            Picture1.Refresh
            Sleep 120
        Next
    Case 4 ' swap endian << what is endian?  http://en.wikipedia.org/wiki/Endianness
        l = &H12345678
        Picture1.Cls
        Picture1.CurrentX = 10: Picture1.CurrentY = 5
        Picture1.Print cBitsClass.BytesToBitString(l), "&H" & Hex(l)
        l = cBitsClass.SwapEndian_Long(l)
        Picture1.CurrentX = 10
        Picture1.Print cBitsClass.BytesToBitString(l), "&H" & Hex(l)
        Picture1.Refresh
    Case 5 ' unsigned to signed
        Picture1.Cls
        Picture1.CurrentX = 4: Picture1.CurrentY = 5
        Picture1.Print "Signed Integer of -55 Equates to Unsigned " & cBitsClass.SignedIntegerToUnsigned(-55)
        Picture1.CurrentX = 4
        Picture1.Print "Unsigned Integer of 41234 Equates to Signed " & cBitsClass.UnsignedIntegerToSigned(41234)
        Picture1.Refresh
    Case 6 ' parse icon mask
        Call ExtractIconMaskExample
    Case 7 ' crc32 example, supports ANSI & Unicode strings, Long/Byte arrays, 1 or 2 Dimensional
        Picture1.Cls
        Picture1.CurrentX = 10: Picture1.CurrentY = 5
        Picture1.Print "Test: The quick brown fox jumped over the lazy dogs."
        Picture1.CurrentX = 10
        Picture1.Print "ANSI CRC is "; cBitsClass.CRC32_String("The quick brown fox jumped over the lazy dogs.");
        Picture1.Print "     Unicode CRC is "; cBitsClass.CRC32_String("The quick brown fox jumped over the lazy dogs.", , crcUnicode)
        Picture1.Refresh
        ' call optional routine to release memory if desired
        cBitsClass.DestroyCRC32LookupTable
        
    Case 8 ' array left shift array
        FillMemory bArray(0), 12, 255 ' fill array with all 1's
        Picture1.Cls
        Picture1.CurrentY = 5: Picture1.CurrentX = 3: Picture1.Print "(0)";
        Picture1.CurrentX = 20: Picture1.Print cBitsClass.ByteArrayToBitString(bArray())
        Picture1.CurrentX = 3: Picture1.Print "(1)";
        Picture1.CurrentX = 20: Picture1.Print cBitsClass.ByteArrayToBitString(bArrayLog())
        Picture1.CurrentX = 20: Picture1.Print "Shifting 96 bits in this example..."
        Picture1.Refresh
        Sleep 500
        For i = 1 To 96
            Call cBitsClass.ShiftBitsLeft_ByteArray(bArrayLog(), 1, , fillOnes_Logical)
            Call cBitsClass.ShiftBitsLeft_ByteArray(bArray(), 1)
            Picture1.Cls
            Picture1.CurrentY = 5: Picture1.CurrentX = 3: Picture1.Print "(0)";
            Picture1.CurrentX = 20: Picture1.Print cBitsClass.ByteArrayToBitString(bArray())
            Picture1.CurrentX = 3: Picture1.Print "(1)";
            Picture1.CurrentX = 20: Picture1.Print cBitsClass.ByteArrayToBitString(bArrayLog())
            Picture1.CurrentX = 20: Picture1.Print "Shifting 96 bits in this example..."
            Picture1.Refresh
            Sleep 50 ' fast routine; gotta slow it down so we can see
        Next
    Case 9 ' array shift left with wrapping, both Arithmetic and Logical filling
        CopyMemory bArray(0), vbGreen, 4&
        CopyMemory bArray(4), vbMagenta, 4&
        CopyMemory bArray(8), vbGreen, 4&
        bArray(11) = 255
        Picture1.Cls
        Picture1.CurrentX = 10: Picture1.CurrentY = 5
        Picture1.Print cBitsClass.ByteArrayToBitString(bArray())
        Picture1.CurrentX = 10: Picture1.Print "Shifting 96 bits in this example..."
        Picture1.Refresh
        Sleep 500
        For i = 1 To 96
            Call cBitsClass.ShiftBitsLeft_ByteArray(bArray(), 1, bitShift_Wrap)
            Picture1.Cls
            Picture1.CurrentX = 10: Picture1.CurrentY = 5
            Picture1.Print cBitsClass.ByteArrayToBitString(bArray())
            Picture1.CurrentX = 10: Picture1.Print "Shifting 96 bits in this example..."
            Picture1.Refresh
            Sleep 50 ' fast routine; gotta slow it down so we can see
        Next
    Case 10 ' array right shift array, no wrapping
        FillMemory bArray(0), 12, 255 ' fill array with all 1's
        Picture1.Cls
        Picture1.CurrentY = 5: Picture1.CurrentX = 3: Picture1.Print "(0)";
        Picture1.CurrentX = 20: Picture1.Print cBitsClass.ByteArrayToBitString(bArray())
        Picture1.CurrentX = 3: Picture1.Print "(1)";
        Picture1.CurrentX = 20: Picture1.Print cBitsClass.ByteArrayToBitString(bArrayLog())
        Picture1.CurrentX = 20: Picture1.Print "Shifting 96 bits in this example..."
        Picture1.Refresh
        Sleep 500
        For i = 1 To 96
            Call cBitsClass.ShiftBitsRight_ByteArray(bArrayLog(), 1, , fillOnes_Logical)
            Call cBitsClass.ShiftBitsRight_ByteArray(bArray(), 1, , fillZeros_Logical)
            Picture1.Cls
            Picture1.CurrentY = 5: Picture1.CurrentX = 3: Picture1.Print "(0)";
            Picture1.CurrentX = 20: Picture1.Print cBitsClass.ByteArrayToBitString(bArray())
            Picture1.CurrentX = 3: Picture1.Print "(1)";
            Picture1.CurrentX = 20: Picture1.Print cBitsClass.ByteArrayToBitString(bArrayLog())
            Picture1.CurrentX = 20: Picture1.Print "Shifting 96 bits in this example..."
            Picture1.Refresh
            Sleep 50 ' fast routine; gotta slow it down so we can see
        Next
    Case 11 ' array shift right with wrapping
        CopyMemory bArray(0), vbGreen, 4&
        CopyMemory bArray(4), vbMagenta, 4&
        CopyMemory bArray(8), vbGreen, 4&
        bArray(11) = 255
        Picture1.Cls
        Picture1.CurrentX = 10: Picture1.CurrentY = 5
        Picture1.Print cBitsClass.ByteArrayToBitString(bArray())
        Picture1.CurrentX = 10: Picture1.Print "Shifting 96 bits in this example..."
        Picture1.Refresh
        Sleep 500
        For i = 1 To 96
            Call cBitsClass.ShiftBitsRight_ByteArray(bArray(), 1, bitShift_Wrap)
            Picture1.Cls
            Picture1.CurrentX = 10: Picture1.CurrentY = 5
            Picture1.Print cBitsClass.ByteArrayToBitString(bArray())
            Picture1.CurrentX = 10: Picture1.Print "Shifting 96 bits in this example..."
            Picture1.Refresh
            Sleep 50 ' fast routine; gotta slow it down so we can see
        Next
    End Select
End Sub

Private Sub ExtractIconMaskExample()
    
    Dim iconMask() As Byte
    Dim ici As ICONINFO
    Dim bmpi As BITMAPINFO
    Dim x As Long, Y As Long, Bit As Long, scanWidth As Long
    
    ' get icon details.
    GetIconInfo Image1.Picture.Handle, ici
    If ici.hbmColor Then DeleteObject ici.hbmColor ' not needed for this example
    
    ' get dib structure
    bmpi.bmiHeader.biSize = 40
    GetDIBits Me.hDC, ici.hbmMask, 0, 0, ByVal 0&, bmpi, 0&
    
    If bmpi.bmiHeader.biBitCount <> 1 Then
        MsgBox "This source icon for this example has been corrupted", vbExclamation + vbOKOnly
        If ici.hbmMask Then DeleteObject ici.hbmMask
        Exit Sub
    End If

    With bmpi.bmiHeader
        scanWidth = ((((.biWidth * .biBitCount) + &H1F) And Not &H1F&) \ &H8)
        ReDim iconMask(0 To scanWidth - 1, 0 To .biHeight - 1)
        .biHeight = -.biHeight ' flip it
    End With
    ' now get the bytes into our array
    GetDIBits Me.hDC, ici.hbmMask, 0, -bmpi.bmiHeader.biHeight, iconMask(0, 0), bmpi, 0&
    DeleteObject ici.hbmMask ' done with this now
cBitsClass.CRC32_ByteArray iconMask

    Picture1.Cls
    Picture1.PaintPicture Image1, 10, 3
    ' render the mask
    With bmpi.bmiHeader
        For Y = 0 To -.biHeight - 1
            For x = 0 To scanWidth - 1
                For Bit = 7 To 0 Step -1
                    ' note: if parsing something like a 2bit (4 color) paletted image,
                    ' the GetMidBits function would be better suited:
                    '   For Bit = 7 to 0 step -2
                    '       SomePaletteIndex = cBitsClass.GetMidBits_Byte(iconMask(X, Y), Bit, 2)
                    '   Next
                    ' and to set the bits of a 2bit paletted image...
                    '   For Bit = 6 to 0 step -2
                    '       cBitsClass.SetMidBits_Byte iconMask(X, Y), Bit, 2, SomePaletteIndex)
                    '   Next
                    
                    If cBitsClass.GetBit_FromByte(iconMask(x, Y), Bit) = 0 Then
                        Picture1.PSet (x * 8 + 7 - Bit + 52, Y + 3), vbBlack
                        ' invert the mask to test SetBit
                        cBitsClass.SetBit_Byte iconMask(x, Y), Bit, True
                    Else
                        Picture1.PSet (x * 8 + 7 - Bit + 52, Y + 3), vbWhite
                        ' invert the mask to test SetBit
                        cBitsClass.SetBit_Byte iconMask(x, Y), Bit, False
                    End If
                Next
            Next
        Next
        ' show the mask inversion
        Picture1.CurrentX = 10: Picture1.Print "Source";
        Picture1.CurrentX = 52: Picture1.Print "Mask"
        For Y = 0 To -.biHeight - 1
            For x = 0 To scanWidth - 1
                For Bit = 7 To 0 Step -1
                    If cBitsClass.GetBit_FromByte(iconMask(x, Y), Bit) = 0 Then
                        Picture1.PSet (x * 8 + 7 - Bit + 100, Y + 3), vbBlack
                    Else
                        Picture1.PSet (x * 8 + 7 - Bit + 100, Y + 3), vbWhite
                    End If
                Next
            Next
        Next
        Picture1.CurrentX = 100
        Picture1.Print "Inverted Mask (see code comments)" ' you know, that green text above ^^
    End With
    
    Picture1.Refresh

End Sub
