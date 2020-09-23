Attribute VB_Name = "Module1"
Function CodeAToByte(Number As String) As String
'Decodes the number to a Binary A-Code
'A 1 is a Black Line
'A 0 is a White Line
Select Case Number
Case 0
    CodeAToByte = "0001101"
Case 1
    CodeAToByte = "0011001"
Case 2
    CodeAToByte = "0010011"
Case 3
    CodeAToByte = "0111101"
Case 4
    CodeAToByte = "0100011"
Case 5
    CodeAToByte = "0110001"
Case 6
    CodeAToByte = "0101111"
Case 7
    CodeAToByte = "0111011"
Case 8
    CodeAToByte = "0110111"
Case 9
    CodeAToByte = "0001011"
End Select
End Function
Function CodeBToByte(Number As String) As String
'Decodes the number to a Binary B-Code
'A 1 is a Black Line
'A 0 is a White Line
Select Case Number
Case 0
    CodeBToByte = "0100111"
Case 1
    CodeBToByte = "0110011"
Case 2
    CodeBToByte = "0011011"
Case 3
    CodeBToByte = "0100001"
Case 4
    CodeBToByte = "0011101"
Case 5
    CodeBToByte = "0111001"
Case 6
    CodeBToByte = "0000101"
Case 7
    CodeBToByte = "0010001"
Case 8
    CodeBToByte = "0001001"
Case 9
    CodeBToByte = "0010111"
End Select
End Function
Function CodeCToByte(Number As String) As String
'Decodes the number to a Binary C-Code
'A 1 is a Black Line
'A 0 is a White Line
Select Case Number
Case 0
    CodeCToByte = "1110010"
Case 1
    CodeCToByte = "1100110"
Case 2
    CodeCToByte = "1101100"
Case 3
    CodeCToByte = "1000010"
Case 4
    CodeCToByte = "1011100"
Case 5
    CodeCToByte = "1001110"
Case 6
    CodeCToByte = "1010000"
Case 7
    CodeCToByte = "1000100"
Case 8
    CodeCToByte = "1001000"
Case 9
    CodeCToByte = "1110100"
End Select
End Function
Function code(ByVal Number As String) As String
'Generates a sequence for the decoding of the next 6 numbers
Select Case Number
Case 0
    code = "AAAAAA"
Case 1
    code = "AABBAB"
Case 2
    code = "AABBAB"
Case 3
    code = "AABBBA"
Case 4
    code = "ABAABB"
Case 5
    code = "ABBAAB"
Case 6
    code = "ABBBBA"
Case 7
    code = "ABABAB"
Case 8
    code = "ABABBA"
Case 9
    code = "ABBABA"
End Select
End Function
Function PaintCode(frm As Form, fi, se, th)
Dim reihe
Dim z
Dim b
Dim d
frm.Line (1 + 10, 0)-(1 + 10, 25) 'Paint the First two lines on the begin of the Code
frm.Line (3 + 10, 0)-(3 + 10, 25)
reihe = code(fi)
For z = 1 To 6 'Use A and B code to Decode the Barcode 'For each 6 numbers use 7 Lines 6 * 7 = 47 Lines
    If Mid(reihe, z, 1) = "A" Then 'Code A
        b = CodeAToByte(Mid(se, z, 1))
        For d = 1 To 7 'Paint the 7 Lines (A Code)
            If Mid(b, d, 1) = 1 Then 'On all 7 numbers Check if it is a 1 or a 0 and Paint a Black or a White Line
                frm.Line ((z - 1) * 7 + d + 3 + 10, 0)-((z - 1) * 7 + d + 3 + 10, 20), &H0 'Black Line
            Else
                frm.Line ((z - 1) * 7 + d + 3 + 10, 0)-((z - 1) * 7 + d + 3 + 10, 20), &HFFFFFF 'White Line
            End If
        Next
    ElseIf Mid(reihe, z, 1) = "B" Then 'Code B
        b = CodeBToByte(Mid(se, z, 1))
        For d = 1 To 7 'Paint the 7 Lines (B Code)
            If Mid(b, d, 1) = 1 Then 'On all 7 numbers Check if it is a 1 or a 0 and Paint a Black or a White Line
                frm.Line ((z - 1) * 7 + d + 3 + 10, 0)-((z - 1) * 7 + d + 3 + 10, 20), &H0 'Black Line
            Else
                frm.Line ((z - 1) * 7 + d + 3 + 10, 0)-((z - 1) * 7 + d + 3 + 10, 20), &HFFFFFF 'White Line
            End If
        Next
    End If
Next
frm.Line (6 * 7 + 5 + 10, 0)-(6 * 7 + 5 + 10, 25) 'Paint the middle two lines of the Code
frm.Line (6 * 7 + 7 + 10, 0)-(6 * 7 + 7 + 10, 25)
    For z = 1 To 6 'Use C code to Decode the Barcode 'For each 6 numbers use 7 Lines 6 * 7 = 47 Lines
        b = CodeCToByte(Mid(th, z, 1)) ' Code C
        For d = 1 To 7 'Paint the 7 Lines (C Code)
            If Mid(b, d, 1) = 1 Then 'On all 7 numbers Check if it is a 1 or a 0 and Paint a Black or a White Line
                frm.Line ((z - 1) * 7 + d + 50 + 10, 0)-((z - 1) * 7 + d + 50 + 10, 20), &H0 'Black Line
            Else
                frm.Line ((z - 1) * 7 + d + 50 + 10, 0)-((z - 1) * 7 + d + 50 + 10, 20), &HFFFFFF 'White Line
            End If
        Next
    Next
frm.Line (94 + 9, 0)-(94 + 9, 25) 'The Last two lines
frm.Line (96 + 9, 0)-(96 + 9, 25)
End Function
Function CheckCode(FullCode As String) As Boolean 'Test the Code
Dim a
Dim b
Dim c
b = 1
If Len(FullCode) = 13 Then
For a = 1 To 12
    If b = 1 Then
        c = c + Mid(FullCode, a, 1)
        b = 0
    Else
        c = c + (Mid(FullCode, a, 1) * 3)
        b = 1
    End If
Next
If (c + Mid(FullCode, 13, 1)) Mod 10 = 0 Then
    CheckCode = True
Else
    CheckCode = False
End If
Else
    CheckCode = False
End If
'e.g:
'Code:   4  0  1  2  3  4  5  0  6  7  8  9  7
'        *1|*3|*1|*3|*1|*3|*1|*3|*1|*3|*1|*3|*1
'Result: 4+ 0+ 1+ 6+ 3+ 12+5+ 0+ 6+ 21+8+ 27 +7 = 100  || 100 Mod 10 = 0 Code is Correct
End Function


'Bar Code
'           ||||||||||||||||||
'           ||||||||||||||||||
'           ||||||||||||||||||
'          4||012345||067897||
'         1.    2.      3.
'1. First Number: Is used to get how 2. is Decoded
'2. 6 Numbers:Are Decoded in A and B Code
'3. Last 6 Numbers: Are always Decoded in C Code
'The First 2 Numbers are the Country code.
'The Next 5 the Manufacteur
'The Next 5 the Product
'And the Last is a Check Number to Check the Code
'e.g:   7610800002482
'       76    = Switzerland
'       10800 = Inter-Milk
'       00248 = Pastmilk
'       2     = Checknumber

'  Number 6 in code A
'  | | |   |
' 00110011111111
' 00110011111111
' 00110011111111
' 00110011111111
' 00110011111111
' 00110011111111
' ^ ^ ^ ^ ^ ^ ^
' 0 1 0 1 1 1 1
'In the created Code for a Number must be 2 Black and 2 White lines
'e.g: 2(or more) small black(or White Lines) next to each other = 1 Line
'e.g: 1 small line = 1 Line
