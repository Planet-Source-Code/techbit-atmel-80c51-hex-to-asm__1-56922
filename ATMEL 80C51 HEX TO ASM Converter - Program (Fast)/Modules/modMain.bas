Attribute VB_Name = "modMain"
Option Explicit

' ----------------------------------------------------------------------------------
'               CONSTANTS, ENUMS, TYPES AND MAIN VARIABLES DEFINITIONS
' ----------------------------------------------------------------------------------

' These API Functions are for Get Path Folder
Public Declare Function SHBrowseForFolder Lib "shell32.dll" (bBrowse As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" (ByVal lItem As Long, ByVal sDir As String) As Long
' This API Function is for generate a Delay
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' This Type is for Open Some Folder Specificated
Public Type BrowseInfo
    hWndOwner As Long
    pidlRoot As Long
    sDisplayName As String
    sTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

' This type will contain the ASM Line of the OPCode File
Public Type tASMLine
    HexOPCode As String
    NumberOfBytes As Integer
    MNemonic As String
    Operands  As String
    Comments As String
    OscillatorPeriod As String
End Type

' This type contain the Op Code List Params
Public Type tOPCList
    HexOPCode As String
    NumberOfBytes As String
    MNemonic As String
    Operands As String
    Comments As String
    OscillatorPeriod As String
End Type

' This type contain the NOB List Params
Public Type tNOBList
    Index As Integer
    Operands As String
    sReplace As String
End Type

' This type will save the Operands List, NOB and Comments
Public Type tISet
    OPCodeList(0 To 255) As tOPCList
    NOBList(0 To 63) As tNOBList
End Type

' This variable will contain the information neceessary to generate an ASM File
Public HexFileName As String
Public ISet As tISet


' ----------------------------------------------------------------------------------
'                                       SUB MAIN
' ----------------------------------------------------------------------------------

Sub main()
    Dim sBuffer As String, xLine() As String, i As Integer
    
    ' Check if the Binary File has been created
    If Dir(App.Path & "\" & "ISet.bin") = "" Then
        ' Load Instruction Operand List Code
        Open App.Path & "\" & "Instruction OPC.hta" For Input As #1
            ' Run through all file
            While Not (EOF(1))
                ' Get one line
                Line Input #1, sBuffer
                ' Split line by comma
                xLine = Split(sBuffer, ";")
                ' Save each line
                With ISet.OPCodeList(i)
                    .HexOPCode = xLine(0)
                    .NumberOfBytes = xLine(1)
                    .MNemonic = xLine(2)
                    .Operands = xLine(3)
                    .Comments = xLine(4)
                    .OscillatorPeriod = xLine(6)
                End With
                i = i + 1
            Wend
        Close #1: i = 0
        ' Load NOB Operand List
        Open App.Path & "\" & "NOB OL.hta" For Input As #1
            ' Run through all file
            While Not (EOF(1))
                ' Get one line
                Line Input #1, sBuffer
                ' Split line by comma
                xLine = Split(sBuffer, "-")
                ' Save each line
                With ISet.NOBList(i)
                    .Index = CInt(xLine(0))
                    .Operands = xLine(1)
                    .sReplace = xLine(2)
                End With
                i = i + 1
            Wend
        Close #1
        ' Crate the Binary File
        Open App.Path & "\" & "ISet.bin" For Binary Access Write As #1
            Put #1, , ISet
        Close #1
    Else
        ' Read Data
        Open App.Path & "\" & "ISet.bin" For Binary Access Read As #1
            Get #1, , ISet
        Close #1
    End If
    ' Load Form
    Call frmMain.Show
End Sub


' ----------------------------------------------------------------------------------
'                                   LOAD DATA FUNCTIONS
' ----------------------------------------------------------------------------------

' This function will obtain the data to convert the Hex Code to Asm Code
Public Function ConvertHexToAsm(lstHEXCode As ListView, lstASMCode As ListView, sAsmFileName As String) As Boolean
    Dim xBuffer() As String
    Dim xOffSet As Long
    Dim xData() As String
    Dim ASMLine As tASMLine
    Dim xSpace As String
    Dim xTabComment As String
                         
    ' Convert ListView data in matrix format to linear format
    Call LoadLinearHex(lstHEXCode, xBuffer)
    ' Set the Progress bar length
    frmMain.pgbProgress.Max = UBound(xBuffer)
    
    ' Open the New ASM File
    Open sAsmFileName & "\" & Left(GetCleanFileNameWOP(HexFileName), Len(GetCleanFileNameWOP(HexFileName)) - 3) & "asm" For Output As #10
    
    While Not ((xOffSet = UBound(xBuffer)) Or (xOffSet = UBound(xBuffer) - 1))
        ' Get the OPCode and the number of bytes that should be taken
        Call GetOPCode(xBuffer(xOffSet), ASMLine)
        
        ' Check for the NOB
        Select Case ASMLine.NumberOfBytes
        Case 1: xOffSet = xOffSet + 1
        Case 2: ReDim Preserve xData(ASMLine.NumberOfBytes - 2)
                xData(0) = xBuffer(xOffSet + 1):  xOffSet = xOffSet + 2
        Case 3: ReDim Preserve xData(ASMLine.NumberOfBytes - 2)
                xData(0) = xBuffer(xOffSet + 1): xData(1) = xBuffer(xOffSet + 2): xOffSet = xOffSet + 3
        End Select
                                
        With ASMLine
            ' Check for the Number of Bytes
            Select Case ASMLine.NumberOfBytes
            ' Save the corresponding data
            Case 1: .Operands = ASMLine.Operands
            Case 2: .Operands = GetOperand(ASMLine, xData)
            Case 3: .Operands = GetOperand(ASMLine, xData)
            End Select
    
            ' Check for the lenght of the mnemonic
            xSpace = Space$(-Len(.MNemonic) + 9)
            xTabComment = Space$(38 - Len(Format(lstASMCode.ListItems.Count + 1, "000000") & "      " & .MNemonic & xSpace & .Operands))
            ' Save data
            Print #10, Format(lstASMCode.ListItems.Count + 1, "000000") & "      " & .MNemonic & xSpace & .Operands & xTabComment & "; " & .Comments
            
            ' ---------------- Load ASM File into the List View ------------------
            Call lstASMCode.ListItems.Add(, , Format$(lstASMCode.ListItems.Count + 1, "000000"))
            lstASMCode.ListItems(lstASMCode.ListItems.Count).SubItems(1) = .MNemonic
            lstASMCode.ListItems(lstASMCode.ListItems.Count).SubItems(2) = .Operands
            lstASMCode.ListItems(lstASMCode.ListItems.Count).SubItems(3) = .Comments
            lstASMCode.ListItems(lstASMCode.ListItems.Count).SubItems(4) = .HexOPCode
            lstASMCode.ListItems(lstASMCode.ListItems.Count).SubItems(5) = .NumberOfBytes
            lstASMCode.ListItems(lstASMCode.ListItems.Count).SubItems(6) = .OscillatorPeriod
        End With
   
        ' Set the current percent progress
        If xOffSet <= frmMain.pgbProgress.Max Then frmMain.pgbProgress.Value = xOffSet
        ' Show the percent in the Status Bar
        frmMain.StatusBar.SimpleText = "Converting ...  " & CInt(frmMain.pgbProgress.Value * 100 / frmMain.pgbProgress.Max) & "% completed."
    Wend
    ' Close File
    Close #10
    ' Restore File Name and Progress Bar
    Call Sleep(1200): frmMain.StatusBar.SimpleText = HexFileName
    frmMain.pgbProgress.Value = 0
End Function


' This Sub will convert the Hex File Format to a format that the program can read
Public Sub ConvertHexFormat(sFileName As String)
    Dim sBuffer As String
    Dim xFile As String
        
    ' :1006CB0073656E736F724925642073656E7369531E
    ' 06CB |7365 6E73 6F72 4925 6420 7365 6E73 6953| 1E
        
    ' Open Source Hex File
    Open sFileName For Input As #5
        ' Read all file
        While Not (EOF(5))
            Line Input #5, sBuffer
            xFile = xFile & Mid(sBuffer, 10, Len(sBuffer) - 11)
        Wend
    Close #5
    ' Open the Destination File
    Open GetCleanPath(sFileName) & "tempHEX.hex" For Output As #5
        Print #5, xFile
    Close #5
End Sub


' This function will load the HEX File into the ListView
Public Function LoadHEX(lstHEXCode As ListView) As Boolean
    Dim xBuffer(0 To 22) As Byte
    Dim Offset As Long
    Dim i As Integer
    
    ' Open File to read byte per byte
    Open GetCleanPath(HexFileName) & "tempHEX.hex" For Binary Access Read As #40
        ' Set the Progress bar length
        frmMain.pgbProgress.Max = LOF(40)
        
        ' Run through all file
        
        While Not (EOF(40))
            ' Get one byte
            Get #40, Offset + 1, xBuffer
            
            ' Adds Line
            Call lstHEXCode.ListItems.Add(, , Format$(lstHEXCode.ListItems.Count * 11 + 1, "0000000"))
            
            ' Put data into the ListView
            For i = 0 To 10
                ' Adds each byte
                lstHEXCode.ListItems(lstHEXCode.ListItems.Count).SubItems(i + 1) = Chr(xBuffer(i * 2)) & Chr(xBuffer(i * 2 + 1))
            Next i
            ' Increment Offset
            Offset = Offset + 22
            
            ' Set the current percent progress
            If Offset <= frmMain.pgbProgress.Max Then frmMain.pgbProgress.Value = Offset
            ' Show the percent in the Status Bar
            frmMain.StatusBar.SimpleText = "Opening Hex File ...  " & CInt(frmMain.pgbProgress.Value * 100 / frmMain.pgbProgress.Max) & "% completed."
        Wend
    ' Close File
    Close #40
    ' Restore File Name
    Call Sleep(1200): frmMain.StatusBar.SimpleText = HexFileName: frmMain.pgbProgress.Value = 0
End Function


' This function will fill the lstValues into a linear vector
Private Function LoadLinearHex(lstHEXCode As ListView, Buffer() As String)
    Dim i As Long
    Dim j As Integer
    
    ' Redim Vector
    ReDim Buffer(0) As String
    
    ' Run through the whole file
    For i = 1 To lstHEXCode.ListItems.Count
        For j = 1 To 11
            ' Check if the lstviwe cell isn't empty
            If lstHEXCode.ListItems(i).SubItems(j) <> "" Then
                ' Save value
                Buffer(UBound(Buffer)) = lstHEXCode.ListItems(i).SubItems(j)
                ' Redim and convert it
                ReDim Preserve Buffer(UBound(Buffer) + 1) As String
            End If
        Next j
    Next i
End Function


' ----------------------------------------------------------------------------------
'                            FUNCTIONS TO CREATE THE ASM FILE
' ----------------------------------------------------------------------------------

' This function will compare the read Opcode with the OpCode into the OpCode File and return by reference the ASMLine
Public Function GetOPCode(OpCode As String, ASMLine As tASMLine) As Boolean
On Error GoTo Sol

    ' Reset previous obtained params
    ASMLine.HexOPCode = 0: ASMLine.MNemonic = "": ASMLine.NumberOfBytes = 0: ASMLine.Operands = ""
    ' Return ASM Line
    With ISet.OPCodeList(CInt(CHexToDec(OpCode)))
        ASMLine.HexOPCode = .HexOPCode
        ASMLine.NumberOfBytes = .NumberOfBytes
        ASMLine.MNemonic = .MNemonic
        ASMLine.Operands = .Operands
        ASMLine.Comments = .Comments
        ASMLine.OscillatorPeriod = .OscillatorPeriod
    End With
    ' Exit here
    GetOPCode = True: Exit Function
Sol:
    GetOPCode = False
End Function


' This function Get the Operand as is specificated in the OpCode File
Public Function GetOperand(ASMLine As tASMLine, OPCodeParams() As String) As String
    Dim i As Integer
    
    ' Search Operand
    For i = 0 To 63
        If ASMLine.Operands = ISet.NOBList(i).Operands Then Exit For
    Next i
    
    With ISet.NOBList(i)
        ' Check for NOB
        If UBound(OPCodeParams) = 0 Then
            ' Return the Correct Operads
            GetOperand = Replace(.sReplace, "DATA", CStr(OPCodeParams(0) & "H"))
        Else
            ' Special Cases
            If Trim(ASMLine.Operands) = "code addr" Then
                GetOperand = Replace(.sReplace, "DATA", CStr(OPCodeParams(0) & OPCodeParams(1) & "H"))
            ElseIf Trim(ASMLine.Operands) = "DPTR,#data" Then
                GetOperand = Replace(.sReplace, "DATA1DATA2", CStr(OPCodeParams(0) & OPCodeParams(1) & "H"))
            Else
                ' Return the Correct Operads
                GetOperand = Replace(.sReplace, "DATA1", CStr(OPCodeParams(0) & "H"))
                GetOperand = Replace(.sReplace, "DATA2", CStr(OPCodeParams(1) & "H"))
            End If
        End If
    End With
End Function


' ----------------------------------------------------------------------------------
'                                   COMMON FUNCTIONS
' ----------------------------------------------------------------------------------

' This function will convert the HexaDecimal to Decimal value
Public Function CHexToDec(HexNumber As String) As String
    ' Declare some constants
    Const Hx = "&H"
    Const BigShift = 65536
    Const LilShift = 256, Two = 2
    ' Declare the needed variables
    Dim Tmp$
    Dim LO1 As Integer, LO2 As Integer
    Dim HI1 As Long, HI2 As Long
    
    ' Asign the number to Tmp
    Tmp = HexNumber
    If UCase(Left$(HexNumber, 2)) = "&H" Then Tmp = Mid$(HexNumber, 3)
    
    Tmp = Right$("0000000" & Tmp, 8)
    
    ' Verify if it's numeric
    If IsNumeric(Hx & Tmp) Then
        
        LO1 = CInt(Hx & Right$(Tmp, Two))
        HI1 = CLng(Hx & Mid$(Tmp, 5, Two))
        LO2 = CInt(Hx & Mid$(Tmp, 3, Two))
        HI2 = CLng(Hx & Left$(Tmp, Two))
        
        ' Return the converted value
        CHexToDec = CCur(HI2 * LilShift + LO2) * BigShift + (HI1 * LilShift) + LO1
    End If
End Function

' This function will convert from HexValues into a Vector to String
Public Function CHexToStr(HexNumber() As Byte) As String
    Dim i As Integer
    
    ' Run throught each item of the vector
    For i = 0 To UBound(HexNumber)
        CHexToStr = CHexToStr & CStr(Hex(HexNumber(i)))
    Next i
End Function

' This function will show the Hex Value splited
Public Function SplitHexNumber(HexNumber() As Byte) As String
    Dim i As Integer
On Error Resume Next
    ' Run throught each item of the vector
    For i = 0 To UBound(HexNumber)
        SplitHexNumber = SplitHexNumber & Format(CStr(Hex(HexNumber(i))), "00") & " "
    Next i
    ' Analyze the string to determine where are the letters
    SplitHexNumber = Trim(FillNumberXX(SplitHexNumber))
End Function

' This function will return the string with the letters as "0A,0B, ..."
Public Function FillNumberXX(Str As String) As String
    ' Check for a letter
    If IsNumber(Str) = False And Len(Str) < 2 Then
        FillNumberXX = "0" & Str
    Else
        FillNumberXX = Format$(Str, "00")
    End If
End Function

' This function will check if the string is a number
Public Function IsNumber(Str As String) As Boolean
    Dim x As Integer
    
    On Error GoTo Sol
    
    x = CInt(Str)
    
    IsNumber = True
Sol:

End Function

' This function will check if the value is valid
Public Function CheckForValidValue(ItemValue As String) As Boolean
    Dim xStr() As String
    Dim i As Integer
    
    ' Split the String to check if this has seven items
    xStr = Split(ItemValue, " ")
    
    ' Fisrt check for 7 Items
    If UBound(xStr) = 7 Then
        ' Check if each item has 2 Chars
        For i = 0 To UBound(xStr)
            If Len(xStr(i)) <> 2 Then Exit Function
        Next i
        ' Return True
        CheckForValidValue = True
    End If
End Function

' This function will create a 8 bytes number from string
Public Function CreateByteValueFromStr(Str As String, xValue() As Byte)
    Dim i As Integer
    Dim xStr() As String
      
    ' Initialize Vector
    ReDim xValue(0) As Byte
    
    ' Get each byte separately
    xStr = Split(Str, " ")
    
    ' Fisrt create the number
    For i = 0 To UBound(xStr)
        ' Redim vector
        ReDim Preserve xValue(i) As Byte
        ' Save byte
        xValue(i) = CByte(CHexToDec(xStr(i)))
    Next i
End Function

' This function will Fill the Vector with zeros
Public Function FillVector(xValue() As Byte)
    Dim i As Integer
    
    ' Initialize Vector
    For i = 0 To 7
        ReDim xValue(i) As Byte
        ' Fill zero
        xValue(i) = 0
    Next i
End Function

' This function will return the filename without its path
Public Function GetCleanFileNameWOP(sFileName As String) As String
    GetCleanFileNameWOP = Right(sFileName, Len(sFileName) - InStrRev(sFileName, "\"))
End Function

' This function will return the filename without its extension
Public Function GetCleanFileName(sFileName As String) As String
    GetCleanFileName = Left(sFileName, InStrRev(sFileName, ".") - 1)
End Function

' This function will return the clean file path
Public Function GetCleanPath(sFullPath As String) As String
    GetCleanPath = Left(sFullPath, InStrRev(sFullPath, "\"))
End Function

' This is used for Explore Folders
Public Function Examined(F As Form, Title As String) As String
    Dim BI As BrowseInfo
    Dim Item As Long
    Dim Folder As String
       
    ' hWnd of Active Form
    BI.hWndOwner = F.hwnd
    ' Start to Desktop
    BI.pidlRoot = 0
    ' This is a Buffer
    BI.sDisplayName = Space(260)
    ' Windows Title
    BI.sTitle = Title
    ' Search Folders
    BI.ulFlags = 1
    BI.lpfn = 0
    BI.lParam = 0
    BI.iImage = 0
   
    Item = SHBrowseForFolder(BI)
   
    If Item Then
        ' This is a Buffer
        Folder = Space(260)
       
        ' Obtein the Truth Path Folder, from of the ID (Item) selected with SHBrowseForFolder
        If SHGetPathFromIDList(Item, Folder) Then
            Examined = Left(Folder, InStr(Folder, Chr(0)) - 1)
        Else
            Examined = ""
        End If
    End If
End Function


Private Function RestoreFile()

' ------------- Add Comments to the main file -----------------
    
    ' Open the New File
    With frmMain.cd1
        .CancelError = True
        .Filter = "All Files (*.*)|*.*"
        .Flags = cdlOFNFileMustExist
        
        ' Show Common Open
        .ShowOpen
        
        Dim sBuffer As String, xLine() As String, xSpace As String, xTabComment As String
        Open App.Path & "\" & "HF.txt" For Output As #41
        Open .FileName For Input As #40
            ' Read all file
            While Not (EOF(40))
                Line Input #40, sBuffer
                xLine = Split(sBuffer, ";")
                
                ' Check if isn't empty
                If sBuffer <> "" Then
                    ' Check if isn't an empty line
                    If InStr(1, xLine(1), ":") = 0 Then
                        
                        ' 000000          ORG    0000EH                  ' 10

                        ' Tab String
                        xSpace = Space$(-Len(xLine(1)) + 8)
                        xTabComment = Space$(45 - Len(Format(xLine(0), "000000") & "          " & xLine(1) & xSpace & xLine(2)))
                        ' Print Line
                        Print #41, Format$(xLine(0), "000000") & "          " & xLine(1) & xSpace & xLine(2) & xTabComment & "; " & xLine(3)
                    Else
                        ' 000001     L000E:  CLR    A                       ' 2 y 2
                        
                        ' Tab String
                        xSpace = Space$(-Len(xLine(2)) + 8)
                        xTabComment = Space$(45 - Len(Format(xLine(0), "000000") & "  " & xLine(1) & "  " & xLine(2) & xSpace & xLine(3)))
                        ' Print Line
                        Print #41, Format$(xLine(0), "000000") & "  " & xLine(1) & "  " & xLine(2) & xSpace & xLine(3) & xTabComment & "; " & xLine(4)
                    End If
                Else
                    Print #41, ""
                End If
            Wend
        Close #40: Close #41
    End With
End Function
    
