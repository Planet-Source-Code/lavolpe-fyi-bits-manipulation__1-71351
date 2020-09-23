VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Though most functions are replicated 3 times (1 for Longs, 1 for Integers, 1 for Bytes), it could
' be done in single functions but would require passing Variants, testing those Variants for the
' variable they contain and then performing the function, handling special overflow cases.
' So... to improve speed, each variable type has its own function

' Whenever a variable is passed to a function that modifies the variable, any bit modifications
' performed are done in the context of the variable type passed. With Integers and Longs, any
' modifications to the high bit can change the value from negative to positive and vice versa.
' Those functions are: SetBit_xxx, SetMidBits_xxx, SetByte_xxx, ShiftBitsLeft_xxx & ShiftBitsRight_xxx

' General Info
' Byte = 8 bits.  Bits are either 0 or 1.
' Integer = 2 bytes (16 bits), Long = 4 bytes (32 bits)

' Bits are generally read from right to left and each bit has the value shown below, if it is ON (has value of 1 vs 0)
'   low bit: 1, 2nd bit: 2, 3rd bit: 4, etc. The 1st 8 bits are: 1, 4, 8, 16, 32, 64, 128
'   For Integers, the additional 8 bits are: 256, 512, 1024, 2048, 4096, 8192, 16384, 32768
'   For Longs, the additional 16 bits are: 65536,131072,262144,524288,1048576,2097152,4194304,8388608,16777216,33554432,67108864,134217728,268435456,536870912,1073741824,2147483648
' Bits are referenced from zero, not 1.  So the 1st bit is bit zero.
' To calculate the bit position values, powers of two are used: Therefore for the 1st bit position, bit 0, value is 1,  2^0 = 1
'   :: For p=0 To lastBitPosition: Debug.Print p, 2^p : Next

' Integer & Long variables' high bits indicate whether or not the value is negative, not the value associated with that bit position.
'   Because of this VB Integers/Longs are signed vs unsigned
'       A signed Integer's min/max values range from -2^15 to 2^15-1
'           An unsigned integer (C++ terms) ranges from 0 to 2^16-1. Unsigned Integers should be Longs in VB
'       A signed Long's min/max values range from -2^31 to 2^31-1
'           An unsigned long (C++ terms) ranges from 0 to 2^32-1. Unsigned Longs should be Doubles in VB
'   A signed byte (C++ terms) values range from -2^7 to 2^7-1. Signed Bytes should be Integers/Longs in VB
'       An unsigned Byte ranges from 0 to 2^8-1. VB Byte variables are always unsigned
' ... Determine min/max values
' :: Multiply bytes used by the variable type by 8 to get number of bits
' :: Is the type signed or unsigned, i.e., can variable type have negative values
' Then the calculation is:
'   if Signed:      min= -2^(nrBits-1), max= (2^(nrBits-1))-1
'   if Unsigned:    min= 0              max= (2^nrBits)-1

' Note regarding Right Shift. Some shifters out there fill negative values with zeroes, but should not.
' Shifting right is done via integer division. (3) 0011 >>1 = 0001 (1) (decimal: 3 \ 2^1 = 1)
' When a negative number is shifted: (-32768) 10000000 00000000 >>1 = 11000000 00000000 (-16384)
'                                    decimal: -32768 \ 2^1 = -16384
' If the far left were filled with zeros, then the result 01000000 00000000 (16384) is wrong.
' Therefore, all the shifting functions allow you to choose how you want vacated bits to be filled
' there are 3 options:
' :: Fill with zeros, Fill with onex, or use Arithmetic fills (Default)
' What is Arithmetic? Fill with the high bit. For positive values, zeros, for negative, ones
' When forcing non-Arithmetic, then the term is a Logical fill

' following APIs & UDT are used for CRC functions
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ByRef Ptr() As Any) As Long
Private Type SafeArrayBound
    cElements As Long
    lLbound As Long
End Type
Private Type SafeArray        ' used as DMA overlay on a string
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    rgSABound(0 To 1) As SafeArrayBound ' allows using up to 2 dimensional arrays
End Type
' following API used for ShiftBitsLeft_Array & ShiftBitsRight_Array
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (ByRef Destination As Any, ByVal Length As Long, ByVal Fill As Byte)

Public Enum StringTypeConstants
    crcANSI = 0
    crcUnicode = 1
End Enum
Public Enum BitShiftWrapEnum
    bitShift_Truncate = 0
    bitShift_Wrap = 1
End Enum
Public Enum FillTypeConstants
    fillZeros_Logical = 0
    fillOnes_Logical = 1
    fillArithmetic_HighBit = 2
End Enum

Private CRC32LUT() As Long
Private CRC32LUTbuilt As Boolean
Private Const MaxUnsignedLongAnd1 As Double = 4294967296#
Private Const MaxSignedLongAnd1 As Double = 2147483648#
Private Const MaskHighBit_Long As Long = &H80000000
Private Const MaskHighBit_Int As Integer = &H8000
Private Const MaskHighBit_Byte As Byte = &H80
Private Const BitCount_Long As Byte = 32
Private Const BitCount_Int As Byte = 16
Private Const BitCount_Byte As Byte = 8
Private mPO2lut(0 To 31) As Double ' 2^31 is a double not Long & used when wrapping shifts

Public Function SignedIntegerToUnsigned(ByVal theValue As Integer) As Long
    ' function converts a signed integer to an unsigned Long
    SignedIntegerToUnsigned = (theValue And &HFFFF&)
End Function
Public Function UnsignedIntegerToSigned(ByVal theValue As Long) As Integer
    ' function converts a Long to a signed integer
    ' if the passed value contains more than 16 usable bits, we can't scrunch 16+ into a 16bit variable
    If theValue < 0& Or theValue > mPO2lut(BitCount_Int) - 1& Then Exit Function ' overflow
    If theValue < mPO2lut(15) Then
        UnsignedIntegerToSigned = theValue
    Else
        UnsignedIntegerToSigned = theValue - mPO2lut(BitCount_Int)
    End If
End Function
Public Function SignedLongToUnsigned(ByVal theValue As Long) As Double
    ' function converts a signed long to an unsigned Double
    If (theValue And MaskHighBit_Long) Then
        SignedLongToUnsigned = theValue + MaxUnsignedLongAnd1
    Else
        SignedLongToUnsigned = theValue
    End If
End Function
Public Function UnsignedLongToSigned(ByVal theValue As Double) As Long
    ' function converts a Double to a signed Long
    
    ' if the passed value contains more than 32 usable bits, we can't scrunch 32+ into a 32bit variable
    If theValue < 0# Or theValue > MaxUnsignedLongAnd1 Then Exit Function  ' overflow
    If theValue < MaxSignedLongAnd1 Then
        UnsignedLongToSigned = theValue
    Else
        UnsignedLongToSigned = theValue - MaxUnsignedLongAnd1
    End If
End Function

Public Function MakeDWord2(ByVal theHiWord As Integer, ByVal theLoWord As Integer) As Long
    ' create a Long (4 bytes) from the passed 2 Integers (2 bytes each)
    MakeDWord2 = (CLng(theHiWord) * mPO2lut(BitCount_Int)) Or (theLoWord And &HFFFF&)
End Function
Public Function MakeDWord4(ByVal theHiByte As Byte, ByVal theByte2 As Byte, ByVal theByte1 As Byte, ByVal theLoByte As Byte) As Long
    ' create a Long (4 bytes) from the passed 4 bytes
    If (theHiByte And MaskHighBit_Byte) Then
        MakeDWord4 = ((theHiByte And &H7F) * &H1000000) Or theByte2 * &H10000 Or theByte1 * &H100& Or theLoByte Or MaskHighBit_Long
    Else
        MakeDWord4 = theHiByte * &H1000000 Or theByte2 * &H10000 Or theByte1 * &H100& Or theLoByte
    End If
End Function
Public Function MakeWord(ByVal theHiByte As Byte, ByVal theLoByte As Byte) As Integer
    ' create an Integer (2 bytes) from the passed 2 bytes
    If theHiByte And MaskHighBit_Byte Then
        MakeWord = ((theHiByte And &H7F) * &H100&) Or theLoByte Or MaskHighBit_Int
    Else
        MakeWord = (theHiByte * &H100&) Or theLoByte
    End If
End Function
Public Function GetBit_FromLong(ByVal theValue As Long, ByVal Position0to31 As Byte) As Byte
    ' function returns the bit on/off from the passed value.
    ' Position must be 0 to max bits - 1 for the variable passed
    If Position0to31 = BitCount_Long - 1& Then ' high bit requested
        If (theValue And MaskHighBit_Long) Then GetBit_FromLong = 1
    ElseIf Position0to31 < BitCount_Long Then
        GetBit_FromLong = ((theValue And &H7FFFFFFF) \ mPO2lut(Position0to31)) And 1
    End If
End Function
Public Function GetBit_FromInteger(ByVal theValue As Integer, ByVal Position0to15 As Byte) As Byte
    ' function returns the bit on/off from the passed value.
    ' Position must be 0 to max bits - 1 for the variable passed
    If Position0to15 < BitCount_Int Then GetBit_FromInteger = ((theValue And &HFFFF&) \ mPO2lut(Position0to15)) And 1
End Function
Public Function GetBit_FromByte(ByVal theValue As Byte, ByVal Position0to7 As Byte) As Byte
    ' function returns the bit on/off from the passed value.
    ' Position must be 0 to max bits - 1 for the variable passed
    If Position0to7 < BitCount_Byte Then GetBit_FromByte = (theValue \ mPO2lut(Position0to7)) And 1
End Function
Public Function SetBit_Long(ByRef theValue As Long, ByVal Position0to31 As Byte, ByVal TurnOn As Boolean) As Boolean
    ' Sets a single bit either on or off within passed Long
    ' Positions are from right to left
    If Position0to31 = BitCount_Long - 1& Then ' playing with the high bit
        If TurnOn Then
            theValue = theValue Or MaskHighBit_Long
        Else
            theValue = theValue Xor MaskHighBit_Long
        End If
        SetBit_Long = True
    ElseIf Position0to31 < BitCount_Long Then
        If TurnOn Then
            theValue = theValue Or (mPO2lut(Position0to31))
        Else
            theValue = theValue Xor (mPO2lut(Position0to31))
        End If
        SetBit_Long = True
    End If
End Function
Public Function SetBit_Integer(ByRef theValue As Integer, ByVal Position0to15 As Byte, ByVal TurnOn As Boolean) As Boolean
    ' Sets a single bit either on or off within passed Integer
    ' Positions are from right to left
    If Position0to15 = BitCount_Int - 1& Then ' playing with the high bit
        If TurnOn Then
            theValue = theValue Or MaskHighBit_Int
        Else
            theValue = theValue Xor MaskHighBit_Int
        End If
        SetBit_Integer = True
    ElseIf Position0to15 < BitCount_Int Then
        If TurnOn Then
            theValue = theValue Or (mPO2lut(Position0to15))
        Else
            theValue = theValue Xor (mPO2lut(Position0to15))
        End If
        SetBit_Integer = True
    End If
End Function
Public Function SetBit_Byte(ByRef theValue As Byte, ByVal Position0to7 As Byte, ByVal TurnOn As Boolean) As Boolean
    ' Sets a single bit either on or off within passed Byte
    ' Positions are from right to left
    If Position0to7 < BitCount_Byte Then
        If TurnOn Then
            theValue = theValue Or (mPO2lut(Position0to7))
        Else
            theValue = theValue Xor (mPO2lut(Position0to7))
        End If
        SetBit_Byte = True
    End If
End Function

Public Function SetMidBits_FromByte(ByRef destValue As Byte, ByVal fromBit0to7 As Byte, ByVal BitsLength1to8 As Byte, ByVal setValue As Byte) As Boolean
    ' Sets a range of consecutive bits within passed Byte, to the pased value
    ' Positions are from right to left
    ' Note: provide the unshifted values for setValue.
    ' Example: if you want to set the far left 2 bits on, simply provide 3 which is binary 11
    Dim modMask As Long
    ' validate passed parameters
    If fromBit0to7 < BitCount_Byte And BitsLength1to8 > 0 Then
        ' validate number of bits to be modified
        If fromBit0to7 + BitsLength1to8 > BitCount_Byte Then BitsLength1to8 = BitCount_Byte - fromBit0to7
        ' validate setValue fits within the number of bits being modified (i.e, if modifying 2 bits, max value can be 3)
        If setValue < mPO2lut(BitsLength1to8) Then
            modMask = CreateBitMask(fromBit0to7, BitsLength1to8)
            destValue = (destValue Xor (destValue And modMask)) Or (setValue * mPO2lut(fromBit0to7))
            SetMidBits_FromByte = True
        End If
    End If

End Function

Public Function SetMidBits_FromLong(ByRef destValue As Long, ByVal fromBit0to31 As Byte, ByVal BitsLength1to32 As Byte, ByVal setValue As Long) As Boolean
    ' Sets a range of consecutive bits within passed Long, to the pased value
    ' Positions are from right to left
    ' Note: provide the unshifted values for setValue.
    ' Example: if you want to set the far left 2 bits on, simply provide 3 which is binary 11
    Dim modMask As Long
    ' validate passed parameters
    If fromBit0to31 < BitCount_Long And BitsLength1to32 > 0 Then
        ' validate number of bits to be modified
        If fromBit0to31 + BitsLength1to32 > BitCount_Long Then BitsLength1to32 = BitCount_Long - fromBit0to31
        ' validate setValue fits within the number of bits being modified (i.e, if modifying 2 bits, max value can be 3)
        If setValue < mPO2lut(BitsLength1to32) Or BitsLength1to32 = BitCount_Long Then
            modMask = CreateBitMask(fromBit0to31, BitsLength1to32)
            ' remove the previous bits values
            destValue = destValue Xor (destValue And modMask)
            ' now add the new bits values, shifting bits into place first
            If modMask And MaskHighBit_Long Then
                If (setValue \ mPO2lut(BitsLength1to32 - 1&)) Then ' the new bits value will set the high bit
                    destValue = destValue Or ((setValue And mPO2lut(BitsLength1to32 - 1&) - 1&) * mPO2lut(fromBit0to31)) Or MaskHighBit_Long
                Else    ' else the new bits value does not contain high bit
                    destValue = destValue Or ((setValue And mPO2lut(BitsLength1to32 - 1&) - 1&) * mPO2lut(fromBit0to31))
                End If
            Else    ' shift & set
                destValue = destValue Or (setValue * mPO2lut(fromBit0to31))
            End If
            
            SetMidBits_FromLong = True
        End If
    End If

End Function

Public Function SetMidBits_FromInteger(ByRef destValue As Integer, ByVal fromBit0to15 As Byte, ByVal BitsLength1to16 As Byte, ByVal setValue As Integer) As Boolean
    ' Sets a range of consecutive bits within passed Integer, to the pased value
    ' Positions are from right to left
    ' Note: provide the unshifted values for setValue.
    ' Example: if you want to set the far left 2 bits on, simply provide 3 which is binary 11
    Dim modMask As Long
    ' validate passed parameters
    If fromBit0to15 < BitCount_Int And BitsLength1to16 > 0 Then
        ' validate number of bits to be modified
        If fromBit0to15 + BitsLength1to16 > BitCount_Int Then BitsLength1to16 = BitCount_Int - fromBit0to15
        ' validate setValue fits within the number of bits being modified (i.e, if modifying 2 bits, max value can be 3)
        If setValue < mPO2lut(BitsLength1to16) Or BitsLength1to16 = BitCount_Int Then
            modMask = CreateBitMask(fromBit0to15, BitsLength1to16)
            ' remove the previous bits values
            destValue = destValue Xor (destValue And modMask)
            ' now add the new bits values, shifting bits into place first
            If modMask And MaskHighBit_Int Then
                ' overflow checks
                If (setValue \ mPO2lut(BitsLength1to16 - 1&)) Then ' the new bits value will set the high bit
                    destValue = destValue Or ((setValue And mPO2lut(BitsLength1to16 - 1&) - 1&) * mPO2lut(fromBit0to15)) Or MaskHighBit_Int
                Else    ' else the new bits value does not contain high bit
                    destValue = destValue Or ((setValue And mPO2lut(BitsLength1to16 - 1&) - 1&) * mPO2lut(fromBit0to15))
                End If
            Else    ' shift & set
                destValue = destValue Or (setValue * mPO2lut(fromBit0to15))
            End If
            SetMidBits_FromInteger = True
        End If
    End If

End Function

Public Function GetMidBits_Long(ByVal theValue As Long, ByVal fromBit0to31 As Byte, ByVal Length As Byte) As Long
    ' Returns the value of a range of consecutive bits within passed Long
    ' Positions are from right to left
    ' Note: the return value is unshifted bit value
    ' Example: if returning the far 2 left bits and they were binary 11, the result is 3 not -1073741824 (shifted)
    If fromBit0to31 < BitCount_Long Then
        If Length + fromBit0to31 > BitCount_Long Then Length = BitCount_Long - fromBit0to31
        If fromBit0to31 + Length = BitCount_Long Then  ' high bit is part of what is returned
            ' overflow checks
            If fromBit0to31 = 31& Then ' return high bit on/off
                If (theValue And MaskHighBit_Long) Then GetMidBits_Long = 1&
            ElseIf fromBit0to31 Then ' range starting >0 and ending <31
                If (theValue And MaskHighBit_Long) Then  ' high bit set?
                    GetMidBits_Long = ((theValue And &H7FFFFFFF) \ mPO2lut(fromBit0to31)) Or (mPO2lut(Length - 1&)) And (mPO2lut(Length) - 1&)
                Else
                    GetMidBits_Long = ((theValue And &H7FFFFFFF) \ mPO2lut(fromBit0to31)) And (mPO2lut(Length) - 1&)
                End If
            Else    ' entire 32bit range
                GetMidBits_Long = theValue
            End If
        ElseIf Length Then
            GetMidBits_Long = ((theValue And &H7FFFFFFF) \ mPO2lut(fromBit0to31)) And (mPO2lut(Length) - 1&)
        End If
    End If
End Function
Public Function GetMidBits_Integer(ByVal theValue As Integer, ByVal fromBit0to15 As Byte, ByVal Length As Byte) As Long
    ' Returns the value of a range of consecutive bits within passed Integer
    ' Positions are from right to left
    ' Note: the return value is unshifted bit value
    ' Example: if returning the far 2 left bits and they were binary 11, the result is 3 not -16384 (shifted)
    If fromBit0to15 < BitCount_Int Then
        If Length + fromBit0to15 > BitCount_Int Then Length = BitCount_Int - fromBit0to15
        GetMidBits_Integer = ((theValue And &HFFFF&) \ mPO2lut(fromBit0to15)) And (mPO2lut(Length) - 1&)
    End If
End Function
Public Function GetMidBits_Byte(ByVal theValue As Byte, ByVal fromBit0to7 As Byte, ByVal Length As Byte) As Long
    ' Returns the value of a range of consecutive bits within passed Byte
    ' Positions are from right to left
    ' Note: the return value is unshifted bit value
    ' Example: if returning the far 2 left bits and they were binary 11, the result is 3 not 192 (shifted)
    If fromBit0to7 < BitCount_Byte Then
        If Length + fromBit0to7 > BitCount_Byte Then Length = BitCount_Byte - fromBit0to7
        GetMidBits_Byte = ((theValue And &HFF) \ mPO2lut(fromBit0to7)) And (mPO2lut(Length) - 1&)
    End If
End Function
Public Function HighByte(theValue As Variant) As Byte
    ' return the last 8 bits (1 byte) of the passed value
    Select Case (VarType(theValue) And Not vbArray)
    Case vbInteger
        ' note. can also call: GetByte_FromInteger(theValue, 1)
        HighByte = (theValue And &HFF00&) \ &H100
    Case vbLong
        ' note. can also call: GetByte_FromLong(theValue, 3)
        If (theValue And MaskHighBit_Long) Then
            HighByte = (theValue And &H7FFFFFFF) \ &H1000000 Or MaskHighBit_Byte
        Else
            HighByte = theValue \ &H1000000
        End If
    Case vbByte
        HighByte = theValue
    End Select
End Function
Public Function LowByte(theValue As Variant) As Byte
    ' returns the 1st 8 bits (1byte) of the passed value
    Select Case (VarType(theValue) And Not vbArray)
    Case vbInteger
        ' note. can also call: GetByte_FromInteger(theValue, 0)
        LowByte = (theValue And &HFF&)
    Case vbLong
        ' note. can also call: GetByte_FromLong(theValue, 0)
        LowByte = (theValue And &HFF&)
    Case vbByte
        LowByte = theValue
    End Select
End Function
Public Function HighWord(ByVal theValue As Long) As Integer
    ' returns the last 16 bits (2 bytes) of the passed value
    HighWord = (theValue And &HFFFF0000) \ &H10000
End Function
Public Function LowWord(ByVal theValue As Long) As Integer
    ' returns the 1st 16 bits (2 bytes) of the passed value
    If theValue And &H8000& Then
      LowWord = theValue Or &HFFFF0000
    Else
      LowWord = theValue And &HFFFF&
    End If
End Function

Public Sub GetRGBA(ByVal fromValue As Long, ByRef Red As Byte, ByRef Green As Byte, ByRef Blue As Byte, ByRef Alpha As Byte)
    If (fromValue And MaskHighBit_Long) Then
        Alpha = (fromValue And Not MaskHighBit_Long) \ &H1000000 Or &H80
    Else
        Alpha = (fromValue And &HFF000000) \ &H1000000
    End If
    Blue = (fromValue And &HFF0000) \ &H10000
    Green = (fromValue And &HFF00&) \ &H100
    Red = (fromValue And &HFF)
    
End Sub
Public Function GetByte_FromLong(ByVal theValue As Long, ByVal Position0to3 As Byte) As Byte
    ' Function returns a byte from a long, from its zero-based Index
    ' Passing a Byte returns the entire byte
    ' Valid Positions: Integer=0,1: Long=0,1,2,3
    If Position0to3 = 3 Then
        If (theValue And MaskHighBit_Long) Then
            GetByte_FromLong = (theValue And &H7FFFFFFF) \ &H1000000 Or MaskHighBit_Byte
        Else
            GetByte_FromLong = theValue \ &H1000000
        End If
    ElseIf Position0to3 < 3 Then
        GetByte_FromLong = ((theValue And &HFFFFFF) \ mPO2lut(Position0to3 * BitCount_Byte)) And &HFF
    End If
End Function
Public Function GetByte_FromInteger(ByVal theValue As Integer, ByVal Position0to1 As Byte) As Byte
    ' Function returns a byte from an integer, from its zero-based Index
    ' Passing a Byte returns the entire byte
    ' Valid Positions: Integer=0,1: Long=0,1,2,3
    If Position0to1 = 0 Then
        GetByte_FromInteger = (theValue And &HFF&)
    Else
        GetByte_FromInteger = (theValue And &HFF00&) \ &H100
    End If
End Function

Public Function SetByte_FromLong(ByRef destValue As Long, ByVal ByteValue As Byte, ByVal Position0to3 As Byte) As Boolean
    ' function sets one of the 4 bytes in the passed Long variable
    ' Position0to3 is from right to left
    If Position0to3 = 3& Then
        If (ByteValue And MaskHighBit_Byte) Then
            destValue = (destValue And &HFFFFFF) Or (ByteValue And &H7F) * mPO2lut(24) Or MaskHighBit_Long
        Else
            destValue = (destValue And &HFFFFFF) Or ByteValue * mPO2lut(24)
        End If
        SetByte_FromLong = True
    ElseIf Position0to3 < 3& Then
        destValue = (destValue Xor (&HFF& * (mPO2lut(Position0to3 * BitCount_Byte)))) Or ByteValue * mPO2lut(Position0to3 * BitCount_Byte)
        SetByte_FromLong = True
    End If
End Function
Public Function SetByte_FromInteger(ByRef destValue As Integer, ByVal ByteValue As Byte, ByVal Position0to1 As Byte) As Boolean
    ' function sets one of the 2 bytes in the passed Integer variable
    ' Position0to3 is from right to left
    If Position0to1 = 1& Then
        If (ByteValue And MaskHighBit_Byte) Then
            destValue = (destValue And &HFF&) Or (ByteValue And &H7F) * mPO2lut(BitCount_Byte) Or MaskHighBit_Int
        Else
            destValue = (destValue And &HFF&) Or ByteValue * mPO2lut(BitCount_Byte)
        End If
        SetByte_FromInteger = True
    ElseIf Position0to1 = 0& Then
        If (destValue And MaskHighBit_Int) Then
            destValue = (destValue And &H7F00&) Or ByteValue Or MaskHighBit_Int
        Else
            destValue = (destValue And &HFF00&) Or ByteValue
        End If
        SetByte_FromInteger = True
    End If
End Function

Public Function ShiftBitsLeft_Long(ByRef theValue As Long, Optional ByVal Iterations = 1&, _
                                Optional ByVal WrapOption As BitShiftWrapEnum = bitShift_Truncate, _
                                Optional ByVal FillMethod As FillTypeConstants = fillArithmetic_HighBit) As Boolean
                                
    ' function modifies passed Long by shifting the bits n times to the left
    ' Equivalent to C++:  <<n
    ' Wrapping: when 11001001 is shifted left 1 and wrapped: 10010011
    Dim tValue As Long, tWrap As Long
    If Iterations < 1& Then
        ShiftBitsLeft_Long = (Iterations = 1&)
        Exit Function
    End If
    If WrapOption = bitShift_Wrap Then
        Iterations = Iterations Mod BitCount_Long
        If (Iterations And 31) Then ' else nothing to do
            ' we will cache the bits from the bit just left of the bit being shifted to the high bit
            ' this cache will then wrap to the far right
            If Iterations = 1& Then ' prevent overflow
                If (theValue And MaskHighBit_Long) Then tWrap = Iterations
            ElseIf (theValue And MaskHighBit_Long) Then
                tWrap = ((theValue And &H7FFFFFFF) \ mPO2lut(BitCount_Long - Iterations)) Or mPO2lut(Iterations - 1&)
            Else
                tWrap = (theValue \ mPO2lut(BitCount_Long - Iterations))
            End If
            ' now we need to shift right bits the left & append the wrap over
            tValue = (theValue And (mPO2lut(BitCount_Long - Iterations - 1&) - 1&)) * mPO2lut(Iterations) Or tWrap
            ' if the bit being shifted is now the high bit, set that here
            If (theValue And mPO2lut(BitCount_Long - Iterations - 1&)) Then theValue = tValue Or MaskHighBit_Long Else theValue = tValue
        End If
    Else
        If FillMethod = fillArithmetic_HighBit Then
            FillMethod = fillZeros_Logical
        Else
            FillMethod = &HFFFFFFFF
        End If
        If Iterations \ BitCount_Long Then  ' all bits shifted out
            theValue = FillMethod
        Else
            tWrap = BitCount_Long - Iterations - 1&
            tValue = (theValue And (mPO2lut(tWrap) - 1&)) * mPO2lut(Iterations) Or (FillMethod And (mPO2lut(Iterations) - 1))
            If (theValue And mPO2lut(tWrap)) Then theValue = tValue Or MaskHighBit_Long Else theValue = tValue
        End If
    End If
    ShiftBitsLeft_Long = True
End Function

Public Function ShiftBitsLeft_Integer(ByRef theValue As Integer, Optional ByVal Iterations = 1&, _
                                Optional ByVal WrapOption As BitShiftWrapEnum = bitShift_Truncate, _
                                Optional ByVal FillMethod As FillTypeConstants = fillArithmetic_HighBit) As Boolean
                                
    ' function modifies passed Integer by shifting the bits n times to the left
    ' Equivalent to C++:  <<n
    ' Wrapping: when 11001001 is shifted left 1 and wrapped: 10010011
    Dim tValue As Integer, tWrap As Integer
    If Iterations < 1& Then
        ShiftBitsLeft_Integer = (Iterations = 1&)
        Exit Function
    End If
    If WrapOption = bitShift_Wrap Then
        Iterations = Iterations Mod BitCount_Int
        If (Iterations And 15) Then ' else nothing to do
            ' we will cache the bits from the bit just left of the bit being shifted to the high bit
            ' this cache will then wrap to the far right
            If (theValue And MaskHighBit_Int) Then
                tWrap = ((theValue And &H7FFF) \ mPO2lut(BitCount_Int - Iterations)) Or mPO2lut(Iterations - 1&)
            Else
                tWrap = (theValue \ mPO2lut(BitCount_Int - Iterations))
            End If
            ' now we need to shift right bits the left & append the wrap over
            tValue = (theValue And (mPO2lut(BitCount_Int - Iterations - 1&) - 1&)) * mPO2lut(Iterations) Or tWrap
            ' if the bit being shifted is now the high bit, set that here
            If (theValue And mPO2lut(BitCount_Int - Iterations - 1&)) Then theValue = tValue Or MaskHighBit_Int Else theValue = tValue
        End If
    Else
        If FillMethod = fillArithmetic_HighBit Then
            FillMethod = fillZeros_Logical
        Else
            If FillMethod = fillOnes_Logical Then FillMethod = &HFFFF
        End If
        If Iterations \ BitCount_Int Then  ' all bits shifted out
            theValue = FillMethod
        Else
            tWrap = BitCount_Int - Iterations - 1
            tValue = (theValue And (mPO2lut(tWrap) - 1&)) * mPO2lut(Iterations) Or (FillMethod And (mPO2lut(Iterations) - 1))
            If (theValue And mPO2lut(tWrap)) Then theValue = tValue Or MaskHighBit_Int Else theValue = tValue
        End If
    End If
    ShiftBitsLeft_Integer = True
End Function

Public Function ShiftBitsLeft_Byte(ByRef theValue As Byte, Optional ByVal Iterations = 1&, _
                                Optional ByVal WrapOption As BitShiftWrapEnum = bitShift_Truncate, _
                                Optional ByVal FillMethod As FillTypeConstants = fillArithmetic_HighBit) As Boolean
                                
    ' function modifies passed Byte by shifting the bits n times to the left
    ' Equivalent to C++:  <<n
    ' Wrapping: when 11001001 is shifted left 1 and wrapped: 10010011
    Dim tValue As Byte, tWrap As Byte
    If Iterations < 1& Then
        ShiftBitsLeft_Byte = (Iterations = 1&)
        Exit Function
    End If
    If WrapOption = bitShift_Wrap Then
        Iterations = Iterations Mod BitCount_Byte
        If (Iterations And 7) Then ' else nothing to do
            ' we will cache the bits from the bit just left of the bit being shifted to the high bit
            ' this cache will then wrap to the far right
            If (theValue And MaskHighBit_Byte) Then
                tWrap = (theValue \ mPO2lut(BitCount_Byte - Iterations)) Or mPO2lut(Iterations - 1&)
            Else
                tWrap = (theValue \ mPO2lut(BitCount_Byte - Iterations))
            End If
            ' now we need to shift right bits the left & append the wrap over
            tValue = (theValue And (mPO2lut(BitCount_Byte - Iterations - 1&) - 1&)) * mPO2lut(Iterations) Or tWrap
            ' if the bit being shifted is now the high bit, set that here
            If (theValue And mPO2lut(BitCount_Byte - Iterations - 1&)) Then theValue = tValue Or MaskHighBit_Byte Else theValue = tValue
        End If
    Else
        If FillMethod = fillArithmetic_HighBit Then
            FillMethod = fillZeros_Logical
        Else
            If FillMethod = fillOnes_Logical Then FillMethod = &HFF
        End If
        If Iterations \ BitCount_Byte Then  ' all bits shifted out
            theValue = FillMethod
        Else
            tWrap = BitCount_Byte - Iterations - 1
            tValue = (theValue And (mPO2lut(tWrap) - 1&)) * mPO2lut(Iterations) Or (FillMethod And (mPO2lut(Iterations) - 1))
            If (theValue And mPO2lut(tWrap)) Then theValue = tValue Or MaskHighBit_Byte Else theValue = tValue
        End If
    End If
    ShiftBitsLeft_Byte = True

End Function

Public Function ShiftBitsRight_Byte(theValue As Byte, Optional ByVal Iterations As Long = 1&, _
                                Optional ByVal WrapOption As BitShiftWrapEnum = bitShift_Truncate, _
                                Optional ByVal FillMethod As FillTypeConstants = fillArithmetic_HighBit) As Boolean
                                
    ' function modifies passed Byte by shifting the bits n times to the right
    ' Equivalent to C++:  >>n
    Dim tValue As Byte, tWrap As Byte
    
    If Iterations < 1& Then
        ShiftBitsRight_Byte = (Iterations = 1&)
        Exit Function
    End If
    If WrapOption = bitShift_Wrap Then
        Iterations = Iterations Mod BitCount_Byte
        If (Iterations And 7) Then
            ' we will cache the bits from the bit just right of the bit being shifted to the lowest bit
            ' this cache will then wrap to the far left
            If ((theValue And mPO2lut(Iterations - 1&))) Then
                tWrap = (theValue And (mPO2lut(Iterations - 1&) - 1&)) * mPO2lut(BitCount_Byte - Iterations) Or MaskHighBit_Byte
            Else
                tWrap = (theValue And (mPO2lut(Iterations) - 1&)) * mPO2lut(BitCount_Byte - Iterations)
            End If
            ' we then shift the non-wrap bytes to the right
            tValue = (theValue \ mPO2lut(Iterations)) Or tWrap
            If (theValue And MaskHighBit_Byte) Then theValue = tValue Or mPO2lut(BitCount_Byte - Iterations - 1&) Else theValue = tValue
        End If
    Else
        If Iterations \ BitCount_Byte Then ' all bits were shifted out of the value
            If FillMethod = fillOnes_Logical Then theValue = &HFF Else theValue = 0
        Else
            If FillMethod = fillOnes_Logical Then
                tValue = theValue \ mPO2lut(Iterations) Or (&HFF Xor (mPO2lut(BitCount_Byte - Iterations) - 1&))
            Else
                tValue = theValue \ mPO2lut(Iterations)
            End If
            If (theValue And MaskHighBit_Byte) Then theValue = tValue Or mPO2lut(BitCount_Byte - Iterations - 1&) Else theValue = tValue
        End If
    End If
    ShiftBitsRight_Byte = True
End Function

Public Function ShiftBitsRight_Integer(theValue As Integer, Optional ByVal Iterations As Long = 1&, _
                                Optional ByVal WrapOption As BitShiftWrapEnum = bitShift_Truncate, _
                                Optional ByVal FillMethod As FillTypeConstants = fillArithmetic_HighBit) As Boolean
                                
    ' function modifies passed Integer by shifting the bits n times to the right
    ' Equivalent to C++:  >>n
    Dim tValue As Integer, tWrap As Integer
    
    ' Note per MSDN documentation: http://msdn.microsoft.com/en-us/library/b6ex274z(VS.71).aspx
    ' When shifting right, by default, any new bits coming in from the left are 0 if value is positive, else 1
    ' This makes sense: shifting right is equivalent of dividing by a positive number and
    '   when dividing a negative by positive, the result is negative (high bit is set)
    '   when dividing a positive by positive, the result is positive (high bit is not set)
    ' The FillMethod passed can override this behavior
    
    If Iterations < 1& Then
        ShiftBitsRight_Integer = (Iterations = 0&)
        Exit Function
    End If
    If WrapOption = bitShift_Wrap Then
        Iterations = Iterations Mod BitCount_Int
        If (Iterations And 15) Then
            ' we will cache the bits from the bit just right of the bit being shifted to the lowest bit
            ' this cache will then wrap to the far left
            If (theValue And mPO2lut(Iterations - 1&)) Then
                tWrap = (theValue And (mPO2lut(Iterations - 1&) - 1&)) * mPO2lut(BitCount_Int - Iterations) Or MaskHighBit_Int
            Else
                tWrap = (theValue And (mPO2lut(Iterations) - 1&)) * mPO2lut(BitCount_Int - Iterations)
            End If
            ' we then shift the non-wrap bytes to the right
            tValue = ((theValue And &HFFFF&) \ mPO2lut(Iterations)) Or tWrap
            If (theValue And MaskHighBit_Int) Then theValue = tValue Or mPO2lut(BitCount_Int - Iterations - 1&) Else theValue = tValue
        End If
    Else
        If Iterations \ BitCount_Int Then ' all bits shifted out
            If FillMethod = fillOnes_Logical Then theValue = &HFFFF Else theValue = &H0
        Else
            If FillMethod = fillArithmetic_HighBit Then
                theValue = theValue \ mPO2lut(Iterations)
            Else
                If FillMethod = fillOnes_Logical Then
                    tValue = ((theValue And Not MaskHighBit_Int) \ mPO2lut(Iterations)) Or (&HFFFF Xor (mPO2lut(BitCount_Int - Iterations) - 1))
                Else
                    tValue = ((theValue And Not MaskHighBit_Int) \ mPO2lut(Iterations))
                End If
                If (theValue And MaskHighBit_Int) Then theValue = tValue Or mPO2lut(BitCount_Int - Iterations - 1) Else theValue = tValue
            End If
        End If
    End If
    ShiftBitsRight_Integer = True
End Function

Public Function ShiftBitsRight_Long(theValue As Long, Optional ByVal Iterations As Long = 1&, _
                                Optional ByVal WrapOption As BitShiftWrapEnum = bitShift_Truncate, _
                                Optional ByVal FillMethod As FillTypeConstants = fillArithmetic_HighBit) As Boolean
                                
    ' function modifies passed Long by shifting the bits n times to the right
    ' Equivalent to C++:  >>n
    
    ' Note per MSDN documentation: http://msdn.microsoft.com/en-us/library/b6ex274z(VS.71).aspx
    ' When shifting right, by default, any new bits coming in from the left are 0 if value is positive, else 1
    ' This makes sense: shifting right is equivalent of dividing by a positive number and
    '   when dividing a negative by positive, the result is negative (high bit is set)
    '   when dividing a positive by positive, the result is positive (high bit is not set)
    ' The FillMethod passed can override this behavior
    
    Dim tValue As Long, tWrap As Long
    If Iterations < 1& Then
        ShiftBitsRight_Long = (Iterations = 0&)
        Exit Function
    End If
    If WrapOption = bitShift_Wrap Then
        Iterations = Iterations Mod BitCount_Long
        If (Iterations And 31&) Then
            ' we will cache the bits from the bit just right of the bit being shifted to the lowest bit
            ' this cache will then wrap to the far left
            If (theValue And mPO2lut(Iterations - 1&)) Then
                tWrap = (theValue And (mPO2lut(Iterations - 1&) - 1&)) * mPO2lut(BitCount_Long - Iterations) Or MaskHighBit_Long
            Else
                tWrap = (theValue And (mPO2lut(Iterations) - 1&)) * mPO2lut(BitCount_Long - Iterations)
            End If
            ' we then shift the non-wrap bytes to the right
            If Iterations = 31& Then
                tValue = tWrap
            Else
                tValue = ((theValue And &H7FFFFFFF) \ mPO2lut(Iterations)) Or tWrap
            End If
            If (theValue And MaskHighBit_Long) Then theValue = tValue Or mPO2lut(BitCount_Long - Iterations - 1&) Else theValue = tValue
        End If
    Else
        If Iterations \ BitCount_Long Then ' all bits are shifted out
            If FillMethod = fillOnes_Logical Then theValue = &HFFFFFFFF Else theValue = &H0&
        Else
            If FillMethod = fillArithmetic_HighBit Then
                theValue = theValue \ mPO2lut(Iterations)
            Else
                If FillMethod = fillOnes_Logical Then
                    tValue = ((theValue And Not MaskHighBit_Long) \ mPO2lut(Iterations)) Or (&HFFFFFFFF Xor (mPO2lut(BitCount_Long - Iterations) - 1))
                Else
                    tValue = ((theValue And Not MaskHighBit_Long) \ mPO2lut(Iterations))
                End If
                If (theValue And MaskHighBit_Long) Then theValue = tValue Or mPO2lut(BitCount_Long - Iterations - 1) Else theValue = tValue
            End If
        End If
    End If
    ShiftBitsRight_Long = True
End Function

Public Function BytesToBitString(theValue As Variant) As String
    ' Function converts byte, integer, long to a bit string
    ' String length is equal to bit count of the passed variable
    Dim b As Integer, bitLen As Long, tValue As Long, HighBitMask As Long
    Const bitOn As String = "1"
    Const bitOff As String = "0"
    
    If GetVarTypeProperties(theValue, bitLen, HighBitMask) = 0& Then Exit Function
    
    tValue = (theValue And Not HighBitMask)
    BytesToBitString = String$(bitLen, bitOff)
    For b = 0 To bitLen - 2
        If ((tValue \ mPO2lut(b)) And 1) Then Mid$(BytesToBitString, bitLen - b, 1) = bitOn
    Next
    If (theValue And HighBitMask) Then Mid$(BytesToBitString, 1, 1) = bitOn
End Function

Public Function ByteArrayToBitString(theArray() As Byte) As String
    ' String length is equal to array size
    Dim b As Integer, c As Long, cPos As Long
    Const bitOn As String = "1"
    Const bitOff As String = "0"
    
    b = Abs(UBound(theArray) - LBound(theArray)) + 1
    ByteArrayToBitString = String$(b * BitCount_Byte, bitOff)
    
    For b = UBound(theArray) To LBound(theArray) Step -1
        For c = 0 To BitCount_Byte - 1
            If (theArray(b) \ mPO2lut(c) And 1) Then Mid$(ByteArrayToBitString, cPos + BitCount_Byte - c, 1) = bitOn
        Next
        cPos = cPos + BitCount_Byte
    Next

End Function

Public Function StringToBytes(ByVal theBitString As String) As Variant
    ' Function converts bit string to a byte, integer or long
    ' The size of the string determines the variable type returned
    
    Dim rtnByte As Byte, rtnInteger As Integer, rtnLong As Long
    Dim b As Long, bitLen As Long, strLen As Long, strStart As Long
    Const bitOn As String = "1"
    
    strLen = Len(theBitString)
    Select Case strLen
        Case 0
        Case Is < 9 ' allow some flexibility
            bitLen = BitCount_Byte: strStart = 1&
            If bitLen > strLen Then b = strLen + 1& Else b = bitLen
        Case Is < 17 ' allow some flexibility
            bitLen = BitCount_Int: strStart = 1&
            If bitLen > strLen Then b = strLen + 1& Else b = bitLen
        Case Is < 33 ' allow some flexibility
            bitLen = BitCount_Long: strStart = 1&
            If bitLen > strLen Then b = strLen + 1 Else b = bitLen
        Case Else ' > 33, only use the first 32 characters
            bitLen = BitCount_Long: strStart = strLen - 31&: b = bitLen
    End Select
    
    If strLen Then
        For b = 0 To b - 2& ' extract bits from string, right to left
            If Mid$(theBitString, strLen - b, 1) = bitOn Then rtnLong = rtnLong Or mPO2lut(b)
        Next
        ' now set the high bit and the return variant type
        Select Case bitLen
        Case BitCount_Long ' long
            If strLen >= bitLen Then
                If Mid$(theBitString, strStart, 1) = bitOn Then rtnLong = rtnLong Or MaskHighBit_Long
            End If
            StringToBytes = rtnLong
        Case BitCount_Int ' integer
            rtnInteger = rtnLong
            If strLen >= bitLen Then
                If Mid$(theBitString, strStart, 1) = bitOn Then rtnInteger = rtnLong Or MaskHighBit_Int
            End If
            StringToBytes = rtnInteger
        Case Else ' byte
            rtnByte = rtnLong
            If strLen >= bitLen Then
                If Mid$(theBitString, strStart, 1) = bitOn Then rtnByte = rtnLong Or MaskHighBit_Byte
            End If
            StringToBytes = rtnByte
        End Select
    Else
        theBitString = rtnLong
    End If
End Function

Public Function CreateBitMask(ByVal fromBit0to31 As Byte, ByVal MaskLength0to31 As Byte) As Long
    ' returns a bit mask
    Dim theMask As Long, bBit As Byte
    If fromBit0to31 < BitCount_Long Then
        If fromBit0to31 + MaskLength0to31 > BitCount_Long Then MaskLength0to31 = BitCount_Long - fromBit0to31
        ' overflow check
        If fromBit0to31 + MaskLength0to31 = BitCount_Long Then
            If fromBit0to31 <> 31& Then                         ' else only setting the high bit
                For bBit = fromBit0to31 To BitCount_Long - 2&   ' create mask, append high bit mask afterwards
                    theMask = theMask Or mPO2lut(bBit)
                Next
            End If
            theMask = theMask Or MaskHighBit_Long
        Else ' no overflow potential, create mask
            For bBit = fromBit0to31 To fromBit0to31 + MaskLength0to31 - 1&
                theMask = theMask Or mPO2lut(bBit)
            Next
        End If
        CreateBitMask = theMask
    End If

End Function

Public Function SwapEndian_Long(ByVal theValue As Long) As Long
    ' funtion swaps the byte order for longs
    ' i.e., &H12345678 becomes &H78563412
    ' what is endian?  http://en.wikipedia.org/wiki/Endianness
    SwapEndian_Long = _
      (((theValue And &HFF000000) \ &H1000000) And &HFF&) Or _
      ((theValue And &HFF0000) \ &H100&) Or _
      ((theValue And &HFF00&) * &H100&) Or _
      ((theValue And &H7F&) * &H1000000)
    If (theValue And MaskHighBit_Byte) Then SwapEndian_Long = SwapEndian_Long Or MaskHighBit_Long
End Function
Public Function SwapEndian_Integer(ByVal theValue As Integer) As Integer
    ' funtion swaps the byte order for integers
    ' i.e., &H1234 becomes &H3412
    ' what is endian?  http://en.wikipedia.org/wiki/Endianness
    SwapEndian_Integer = _
      (((theValue And &HFF00) \ &H100) And &HFF&) Or _
      ((theValue And &H7F&) * &H100&)
    If (theValue And MaskHighBit_Byte) Then SwapEndian_Integer = SwapEndian_Integer Or MaskHighBit_Int
End Function

Public Function GetVarTypeProperties(theValue As Variant, Optional ByRef BitLength As Long, Optional ByRef HighBitMask As Long) As Long
    ' helper function to determine passed variable type and the high bit associated with it
    ' the return value will be the byte length of the variable: 1, 2, 4.
    ' If return value is zero, the variable type is not handled by this class
    Select Case (VarType(theValue) And Not vbArray)
        Case vbLong:
            BitLength = BitCount_Long
            HighBitMask = MaskHighBit_Long
        Case vbInteger
            BitLength = BitCount_Int
            HighBitMask = MaskHighBit_Int
        Case vbByte
            BitLength = BitCount_Byte
            HighBitMask = MaskHighBit_Byte
        Case Else
            BitLength = 0&
            HighBitMask = 0&
    End Select
    GetVarTypeProperties = BitLength \ BitCount_Byte
End Function

Public Function CRC32_String(theString As String, Optional ByVal CurrentCRC32Value As Long = &HFFFFFFFF, _
                                Optional StringType As StringTypeConstants = crcANSI) As Long
    
    ' CurrentCRC32Value can be passed if using CRC32 on multiple strings or
    ' portions of strings (i.e., loading a file)
    '  Open myFile For Input As #1
    '  Do Until EOF(1)
    '     Line Input #1, someText
    '     crc32 = classname.CRC32_String(someText, crc32, crcANSI)
    '  Loop
    '  Close #1
    
    If CRC32LUTbuilt = False Then CreateCRC32LookupTable
    
    Dim i As Long, iStep As Long
    Dim iLookup As Long
    Dim tSA As SafeArray, Buffer() As Byte

    If Not theString = vbNullString Then
        ' use an array overlay for faster results
        ' by doing so, don't have read the string into an array :: StrConv()
        With tSA
            .cbElements = 1
            .cDims = 1
            .pvData = StrPtr(theString)
            .rgSABound(0).cElements = LenB(theString) ' can use Len(theString)*2
            .rgSABound(0).lLbound = 1
        End With
        CopyMemory ByVal VarPtrArray(Buffer), VarPtr(tSA), 4&
        If StringType = crcUnicode Then iStep = 1 Else iStep = 2
        
        For i = 1 To tSA.rgSABound(0).cElements Step iStep
            iLookup = (CurrentCRC32Value And &HFF) Xor Buffer(i)
            CurrentCRC32Value = ((CurrentCRC32Value And &HFFFFFF00) \ &H100) And &HFFFFFF
            CurrentCRC32Value = CurrentCRC32Value Xor CRC32LUT(iLookup)
        Next i
        CopyMemory ByVal VarPtrArray(Buffer), 0&, 4&
        
        CRC32_String = Not (CurrentCRC32Value)
    End If
    
    
End Function

Public Function CRC32_LongArray(theArray() As Long, Optional ByVal CurrentCRC32Value As Long = &HFFFFFFFF) As Long
    
    ' CurrentCRC32Value can be passed if using CRC32 on multiple arrays
    ' 1 or 2 dimensional arrays only, are supported
    ' Max supported single array elements are (2^31-1)\4 (536870911)
    
    Dim tSA As SafeArray, Buffer() As Byte
    Dim Looper As Long
    
    ' we will overlay the array with another array to guarantee LBound=0 & byte vs Integer representation
    On Error GoTo ExitRoutine
    tSA.cDims = ArrayProps(VarPtrArray(theArray))
    If (tSA.cDims = 1 Or tSA.cDims = 2) Then ' 1 or 2 dimensional array
        With tSA
            .cbElements = 1
            If tSA.cDims = 2 Then
                .rgSABound(0).cElements = (Abs(UBound(theArray, 2) - LBound(theArray, 2)) + 1) * 4&
                .rgSABound(1).cElements = Abs(UBound(theArray, 1) - LBound(theArray, 1)) + 1
                .pvData = VarPtr(theArray(LBound(theArray, 1), LBound(theArray, 2)))
            Else
                .rgSABound(0).cElements = (Abs(UBound(theArray) - LBound(theArray)) + 1) * 4&
                .pvData = VarPtr(theArray(LBound(theArray)))
            End If
        End With
        CopyMemory ByVal VarPtrArray(Buffer), VarPtr(tSA), 4&
        CRC32_LongArray = CRCArray(Buffer, CurrentCRC32Value, (tSA.cDims = 2))
        CopyMemory ByVal VarPtrArray(Buffer), 0&, 4&
        tSA.pvData = 0&
    End If
ExitRoutine:
    If Err Then Err.Clear
    ' only potential error is > 536870911 elements in one of the dimensions
    If tSA.pvData Then CopyMemory ByVal VarPtrArray(Buffer), 0&, 4&

End Function

Public Function CRC32_IntegerArray(theArray() As Integer, Optional ByVal CurrentCRC32Value As Long = &HFFFFFFFF) As Long
    
    ' CurrentCRC32Value can be passed if using CRC32 on multiple arrays
    ' 1 or 2 dimensional arrays only, are supported
    ' Max supported single array elements are (2^31-1)\2 (1073741823)
    
    Dim tSA As SafeArray, Buffer() As Byte
    Dim Looper As Long
    
    ' we will overlay the array with another array to guarantee LBound=0 & byte vs Integer representation
    On Error GoTo ExitRoutine
    tSA.cDims = ArrayProps(VarPtrArray(theArray))
    If (tSA.cDims = 1 Or tSA.cDims = 2) Then ' 1 or 2 dimensional array
        With tSA
            .cbElements = 1
            If tSA.cDims = 2 Then
                .rgSABound(0).cElements = (Abs(UBound(theArray, 2) - LBound(theArray, 2)) + 1) * 2&
                .rgSABound(1).cElements = Abs(UBound(theArray, 1) - LBound(theArray, 1)) + 1
                .pvData = VarPtr(theArray(LBound(theArray, 1), LBound(theArray, 2)))
            Else
                .rgSABound(0).cElements = (Abs(UBound(theArray) - LBound(theArray)) + 1) * 2&
                .pvData = VarPtr(theArray(LBound(theArray)))
            End If
        End With
        CopyMemory ByVal VarPtrArray(Buffer), VarPtr(tSA), 4&
        CRC32_IntegerArray = CRCArray(Buffer, CurrentCRC32Value, (tSA.cDims = 2))
        CopyMemory ByVal VarPtrArray(Buffer), 0&, 4&
        tSA.pvData = 0&
    End If
ExitRoutine:
    If Err Then Err.Clear
    ' only potential error is > 1073741823 elements in one of the dimensions
    If tSA.pvData Then CopyMemory ByVal VarPtrArray(Buffer), 0&, 4&
End Function

Public Function CRC32_ByteArray(theArray() As Byte, Optional ByVal CurrentCRC32Value As Long = &HFFFFFFFF) As Long
    
    ' CurrentCRC32Value can be passed if using CRC32 on multiple arrays
    ' 1 or 2 dimensional arrays only, are supported
    ' Binary arrays are best used when an entire file is loaded into memory
    Dim tSA As SafeArray, Buffer() As Byte
    
    ' we will overlay the array with another array to guarantee LBound=0 if needed
    tSA.cDims = ArrayProps(VarPtrArray(theArray))
    If tSA.cDims = 1 Then
        If LBound(theArray) = 0 Then
            CRC32_ByteArray = CRCArray(theArray, CurrentCRC32Value, False)
        Else
            With tSA
                .cbElements = 1
                .pvData = VarPtr(theArray(LBound(theArray)))
                .rgSABound(0).cElements = Abs(UBound(theArray) - LBound(theArray)) + 1
            End With
            CopyMemory ByVal VarPtrArray(Buffer), VarPtr(tSA), 4&
            CRC32_ByteArray = CRCArray(Buffer, CurrentCRC32Value, False)
            CopyMemory ByVal VarPtrArray(Buffer), 0&, 4&
        End If
    ElseIf tSA.cDims = 2 Then
        If LBound(theArray, 2) = 0 And LBound(theArray, 1) = 0 Then
            CRC32_ByteArray = CRCArray(theArray, CurrentCRC32Value, True)
        Else
            With tSA
                .cbElements = 1
                .pvData = VarPtr(theArray(LBound(theArray, 1), LBound(theArray, 2)))
                .rgSABound(0).cElements = Abs(UBound(theArray, 2) - LBound(theArray, 2)) + 1
                .rgSABound(1).cElements = Abs(UBound(theArray, 1) - LBound(theArray, 1)) + 1
            End With
            CopyMemory ByVal VarPtrArray(Buffer), VarPtr(tSA), 4&
            CRC32_ByteArray = CRCArray(Buffer, CurrentCRC32Value, True)
            CopyMemory ByVal VarPtrArray(Buffer), 0&, 4&
        End If
    End If
    
End Function

Private Function CRCArray(theArray() As Byte, CurrentCRC32Value As Long, b2DArray As Boolean) As Long

    ' helper function which process byte, integer, long arrays
    ' Note: The function cannot be passed a null array. This is
    '   checked before this function is called
    
    If CRC32LUTbuilt = False Then CreateCRC32LookupTable
    
    Dim i As Long, j As Long, lb As Long, ub As Long
    Dim iLookup As Long
    
    If b2DArray Then
        lb = LBound(theArray, 1): ub = UBound(theArray, 1)
        For i = LBound(theArray, 2) To UBound(theArray, 2)
            For j = lb To ub
                iLookup = (CurrentCRC32Value And &HFF) Xor theArray(j, i)
                CurrentCRC32Value = ((CurrentCRC32Value And &HFFFFFF00) \ &H100) And &HFFFFFF
                CurrentCRC32Value = CurrentCRC32Value Xor CRC32LUT(iLookup)
            Next j
        Next i
    Else    ' single dimensional array
        For i = LBound(theArray) To UBound(theArray)
            iLookup = (CurrentCRC32Value And &HFF) Xor theArray(i)
            CurrentCRC32Value = ((CurrentCRC32Value And &HFFFFFF00) \ &H100) And &HFFFFFF
            CurrentCRC32Value = CurrentCRC32Value Xor CRC32LUT(iLookup)
        Next i
    End If
    CRCArray = Not (CurrentCRC32Value)

End Function

Public Sub CreateCRC32LookupTable(Optional ByVal CRCpolynomial As Long = &HEDB88320)
    ' Optional function call. If you want to use your own Polynomial, then
    ' call this before calling any of the CRC32_xxxx functions
    
    ' &HEDB88320 is the official polynomial used by CRC32 in PKZip.
    ' Often the polynomial is shown reversed (04C11DB7).
    Dim i As Long, j As Long

    ReDim CRC32LUT(0 To 255)
    Dim dwCrc As Long

    For i = 0 To 255
        dwCrc = i
        For j = 8 To 1 Step -1
            If (dwCrc And 1) > 0 Then
                dwCrc = ((dwCrc And &HFFFFFFFE) \ 2) And &H7FFFFFFF
                dwCrc = dwCrc Xor CRCpolynomial
            Else
                dwCrc = ((dwCrc And &HFFFFFFFE) \ 2) And &H7FFFFFFF
            End If
        Next j
        CRC32LUT(i) = dwCrc
    Next i
    CRC32LUTbuilt = True
End Sub

Public Sub DestroyCRC32LookupTable()
    ' can be called if you want to free up memory else the
    ' 256 element array is only destroyed when the class is destroyed
    Erase CRC32LUT
    CRC32LUTbuilt = False
End Sub

Public Function ShiftBitsLeft_ByteArray(theArray() As Byte, ByVal Iterations As Long, _
                                Optional ByVal WrapOption As BitShiftWrapEnum = bitShift_Truncate, _
                                Optional ByVal FillMethod As FillTypeConstants = fillArithmetic_HighBit) As Boolean

    ' Function shift an entire array of bytes and is very quick. Only supports single dimensional arrays
    ' If you need to do shifting for crypto reasons or shifting Doubles, Singles, etc, this can be for you.
    ' Simply put whatever you want into a byte array and pass it to this function (CopyMemory)
    
    ' by using a CarryOver value, we can literally shift thousands of bits and not be
    ' limited to just 32.  Additionally, there is no chance of overflows
    ' What is a carryover value?
    '   The carryover is the value that will fall off when shifting
    '   Then when we shift into the next byte, we cache that's byte carryover
    '       then add the previous carryover to the new byte after done shifting
    ' Logic:
    ' Starting: carryover1 set depending on wrap option, shift first byte
    ' Get next byte's carryover (high bits) & place in carryover2
    ' Shift next byte, and then add carryover1
    ' Place carryover2 into the carryover1 & continue to next byte
    ' As you can see below, we can shift n bits in n\8 loops
    
    ' Looks a bit complicated and it is. Wanted to handle iterations and wrapping
    ' as fast as I could, therefore some complicated memory manipulation
    
    Dim b As Long, carryFlag As Byte, carryFlag2 As Long
    Dim carryShift As Long, preserveMask As Long, preserveShift As Long
    Dim ub As Long, Offset8 As Long
    Dim tSA As SafeArray, theBytes() As Byte, Buffer() As Byte
    
    If Iterations < 1& Then 'check for invalid parameter
        ShiftBitsLeft_ByteArray = (Iterations = 0&)
        Exit Function
    End If
    If ArrayProps(VarPtrArray(theArray)) <> 1& Then Exit Function ' supports 1D arrays only
    
    ' set up masks and offset
    On Error GoTo ExitRoutine   ' error? If bit count calc throws an overflow error
    ub = Abs(UBound(theArray) - LBound(theArray)) ' calc UBound based off a zero-bound array
    b = (ub + 1) * BitCount_Byte                  ' calc total bits in the array
    On Error GoTo 0             ' stop error checking
    If Iterations >= b Then
        If WrapOption = bitShift_Truncate Then
            ' if shifting completely out of the array, then return a zeroized array
            If FillMethod = fillArithmetic_HighBit Then FillMethod = fillZeros_Logical
            FillMemory theArray(LBound(theArray)), ub + 1&, FillMethod * &HFF
            ShiftBitsLeft_ByteArray = True
            Exit Function
        End If
    End If
    Iterations = Iterations Mod b       ' get actual iterations
    If Iterations = 0& Then             ' can happen if user passed TotalBits*WholeNumber
        ShiftBitsLeft_ByteArray = True
        Exit Function
    End If
    Offset8 = Iterations \ BitCount_Byte  ' which byte will contain the 1st shifted bit?
    Iterations = Iterations Mod BitCount_Byte ' how many iterations within the byte are we going to do?
    ' calculate which bits will be shifted into the next byte & also the shift factor
    ' toBeShifted = (currentByte \ carryShift)
    carryShift = mPO2lut(BitCount_Byte - Iterations)
    ' calculate which far right bits will be kept and then shifted to the far left
    ' toBeKept = (currentByte And preserveMask) * preserveShift
    preserveMask = mPO2lut(BitCount_Byte - Iterations) - 1
    preserveShift = mPO2lut(Iterations)
    
    ' normalize array to single dimension, zero-lbound
    With tSA
        .cbElements = 1
        .cDims = 1
        .pvData = VarPtr(theArray(LBound(theArray)))
        .rgSABound(0).cElements = ub + 1&
    End With
    CopyMemory ByVal VarPtrArray(theBytes), VarPtr(tSA), 4&
    
    If WrapOption = bitShift_Wrap Then
        If Offset8 Then ' we will be BYTE wrapping from UBound() to LBound()
            ' use about 1/2 the array vs copying entire array (potential that array can be very large)
            If Offset8 < ub \ 2& + 1& Then  ' 3 steps to shift bytes
                carryFlag = Offset8 - 1: ReDim Buffer(0 To carryFlag) ' determine how many bytes to cache
                CopyMemory Buffer(0), theBytes(ub - carryFlag), carryFlag + 1& ' cache 0 to offset8
                CopyMemory theBytes(Offset8), theBytes(0), ub - Offset8 + 1& ' move 0 to offset8
                CopyMemory theBytes(0), Buffer(0), carryFlag + 1 ' move cache to 0
            Else                    ' 4 steps to shift bytes
                carryFlag = ub - Offset8: ReDim Buffer(0 To carryFlag) ' determine how many bytes to cache
                CopyMemory Buffer(0), theBytes(Offset8), carryFlag + 1& ' cache from offset8 to UBound
                CopyMemory theBytes(Offset8), theBytes(0), carryFlag + 1& ' slide right side to bytes vacated by caching
                CopyMemory theBytes(0), theBytes(carryFlag + 1&), ub - (carryFlag * 2 + 1&)
                CopyMemory theBytes(ub - (carryFlag * 2& + 1&)), Buffer(0), carryFlag + 1& ' now place the cached bytes in the middle
            End If
            Offset8 = 0& ' loop will now start from LBound
        End If
        carryFlag = theBytes(ub) \ carryShift 'get carryover for last byte in array
    
    Else
        If FillMethod = fillArithmetic_HighBit Then FillMethod = fillZeros_Logical
        FillMethod = FillMethod * &HFF
        If Offset8 Then   ' do we need to shift any byes before we start?
            CopyMemory theBytes(Offset8), theBytes(0), (ub - Offset8 + 1&) ' shift 0 to the offset8 position
            FillMemory theBytes(0), Offset8, FillMethod ' zero out the far right bytes
        End If
        carryFlag = FillMethod \ carryShift
    End If
    If Iterations Then ' here is where we do the shifting from bit to bit, byte to byte
        For b = Offset8 To ub
            ' calculate the bits to be carried over to next byte
            carryFlag2 = theBytes(b) \ carryShift
            ' shift the preserve bits to the far left & append the previous byte's carryover
            theBytes(b) = (theBytes(b) And preserveMask) * preserveShift Or carryFlag
            ' set new carryover
            carryFlag = carryFlag2
        Next
    End If
    
    CopyMemory ByVal VarPtrArray(theBytes), 0&, 4&
    ShiftBitsLeft_ByteArray = True
ExitRoutine:
End Function


Public Function ShiftBitsRight_ByteArray(theArray() As Byte, ByVal Iterations As Long, _
                                Optional ByVal WrapOption As BitShiftWrapEnum = bitShift_Truncate, _
                                Optional ByVal FillMethod As FillTypeConstants = fillArithmetic_HighBit) As Boolean

    ' Function shift an entire array of bytes and is very quick. Only supports single dimensional arrays
    ' If you need to do shifting for crypto reasons or shifting Doubles, Singles, etc, this can be for you.
    ' Simply put whatever you want shifted into a byte array and pass it to this function

    Dim b As Long, carryFlag As Byte, carryFlag2 As Long
    Dim carryMask As Long, carryShift As Long
    Dim preserveMask As Long, preserveShift As Long
    Dim ub As Long, Offset8 As Long
    
    Dim tSA As SafeArray, theBytes() As Byte, Buffer() As Byte
    
    If Iterations < 1& Then 'check for invalid parameter
        ShiftBitsRight_ByteArray = (Iterations = 0&)
        Exit Function
    End If
    If ArrayProps(VarPtrArray(theArray)) <> 1& Then Exit Function ' supports 1D arrays only
    
    ' set up masks and offset
    On Error GoTo ExitRoutine   ' error? If bit count calc throws an overflow error
    ub = Abs(UBound(theArray) - LBound(theArray)) ' calc UBound based off a zero-bound array
    b = (ub + 1) * BitCount_Byte                  ' calc total bits in the array
    On Error GoTo 0             ' stop error checking
    If Iterations >= b Then
        If WrapOption = bitShift_Truncate Then
            ' if shifting completely out of the array, then return a zeroized array
            If FillMethod = fillArithmetic_HighBit Then FillMethod = fillZeros_Logical
            FillMemory theArray(LBound(theArray)), ub + 1&, FillMethod * &HFF
            ShiftBitsRight_ByteArray = True
            Exit Function
        End If
    End If
    Iterations = Iterations Mod b       ' get actual iterations
    If Iterations = 0& Then             ' can happen if user passed TotalBits*WholeNumber
        ShiftBitsRight_ByteArray = True
        Exit Function
    End If
    Offset8 = Iterations \ BitCount_Byte  ' which byte will contain the 1st shifted bit?
    Iterations = Iterations Mod BitCount_Byte ' how many iterations within the byte are we going to do?
    ' calculate which bits will be shifted into the next byte & also the shift factor
    ' toBeShifted = (currentByte And carryMask) \ carryShift
    carryMask = mPO2lut(Iterations) - 1
    carryShift = mPO2lut(BitCount_Byte - Iterations)

    ' calculate which far right bits will be kept and then shifted to the far left
    ' toBeKept = (currentByte And preserveMask) * preserveShift
    preserveShift = mPO2lut(Iterations)
    
    ' normalize array to single dimension, zero-lbound
    With tSA
        .cbElements = 1
        .cDims = 1
        .pvData = VarPtr(theArray(LBound(theArray)))
        .rgSABound(0).cElements = ub + 1&
    End With
    CopyMemory ByVal VarPtrArray(theBytes), VarPtr(tSA), 4&
    
    If WrapOption = bitShift_Wrap Then
        If Offset8 Then ' we will be BYTE wrapping from UBound() to LBound()
            ' use about 1/2 the array vs copying entire array (potential that array can be very large)
            If Offset8 < ub \ 2& + 1& Then ' 3 steps to shift bytes
                carryFlag = Offset8 - 1: ReDim Buffer(0 To carryFlag) ' determine how many bytes to cache
                CopyMemory Buffer(0), theBytes(ub - carryFlag), carryFlag + 1& ' cache 0 to offset8
                CopyMemory theBytes(ub - carryFlag), theBytes(0), carryFlag + 1& ' move 0 to offset8
                CopyMemory theBytes(0), theBytes(carryFlag + 1), ub - carryFlag * 2 - 1
                CopyMemory theBytes(ub - carryFlag * 2 - 1), Buffer(0), carryFlag + 1  ' move cache to 0
            Else                    ' 4 steps to shift bytes
                carryFlag = ub - Offset8: ReDim Buffer(0 To carryFlag) ' determine how many bytes to cache
                CopyMemory Buffer(0), theBytes(ub - carryFlag), carryFlag + 1& ' cache from offset8 to UBound
                CopyMemory theBytes(carryFlag + 1), theBytes(0), ub - carryFlag  ' slide right side to bytes vacated by caching
                CopyMemory theBytes(0), Buffer(0), carryFlag + 1& ' now place the cached bytes in the middle
            End If
            Offset8 = 0& ' loop will now start from LBound
        End If
        carryFlag = (theBytes(0) And carryMask) * carryShift 'get carryover for last byte in array
    
    Else
        If FillMethod = fillArithmetic_HighBit Then
            If theBytes(ub) And MaskHighBit_Byte Then FillMethod = fillOnes_Logical Else FillMethod = fillZeros_Logical
        End If
        FillMethod = FillMethod * &HFF
        Offset8 = Iterations \ BitCount_Byte        ' which byte will the 1st bit be shifted to?
        If Offset8 Then                             ' do we need to shift any byes before we start?
            CopyMemory theBytes(0), theBytes(Offset8), (ub - Offset8 + 1) ' shift 0 to the offset8 position
            FillMemory theBytes(ub - Offset8 + 1), Offset8, FillMethod  ' zero out the far right bytes
        End If
        carryFlag = (FillMethod And carryMask) * carryShift
    End If
    If Iterations Then                      ' if the iterations were already completed, bug out
        For b = ub - Offset8 To 0 Step -1
            ' calculate the bits to be carried over to next byte
            carryFlag2 = (theBytes(b) And carryMask) * carryShift
            ' shift the preserve bits to the far left & append the previous byte's carryover
            theBytes(b) = (theBytes(b) \ preserveShift) Or carryFlag
            ' set new carryover
            carryFlag = carryFlag2
        Next
    End If
    
    CopyMemory ByVal VarPtrArray(theBytes), 0&, 4&
    ShiftBitsRight_ByteArray = True
ExitRoutine:
End Function


Private Function ArrayProps(ByVal arrayPtr As Long) As Long

    Dim tSA As SafeArray
    ' help function. Validates passed array is initialized and has 1 or 2 dimensions
    If arrayPtr Then
        CopyMemory arrayPtr, ByVal arrayPtr, 4&
        If arrayPtr Then    ' else uninitialized array
            CopyMemory ByVal VarPtr(tSA), ByVal arrayPtr, 16&     ' safe array structure minus bounds info
            If tSA.cDims Then
                CopyMemory tSA.rgSABound(0).cElements, ByVal arrayPtr + 16&, 8&
                If tSA.rgSABound(0).cElements Then  ' else VB null array Ubound=-1, Lbound=0
                    ArrayProps = tSA.cDims
                End If
            End If
        End If
    End If
    
End Function

Private Sub Class_Initialize()
    Dim i As Integer
    mPO2lut(0) = 1#
    For i = 1 To 31
        mPO2lut(i) = mPO2lut(i - 1) * 2&
    Next
End Sub


