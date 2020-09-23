Attribute VB_Name = "basScan"
'vnAntiVirus 1.0

'Author : Dung Le Nguyen
'Email : dungcoivb@gmail.com
'This is a software open source

Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public ph As Long
Public Sub GetData(FileDat As String, lstData As ListBox, lstNameV As ListBox, XoaSach As Boolean)
If XoaSach = True Then
    lstData.Clear
    lstNameV.Clear
End If
    Open FileDat For Input As #1
        Do While Not EOF(1)
            Line Input #1, InputData
            lstData.AddItem Split(InputData, "|", , vbBinaryCompare)(0)
            lstNameV.AddItem Split(InputData, "|", , vbBinaryCompare)(1)
        Loop
    Close #1
End Sub
Public Function ScanFile(FilePath As String, ScanIcon As Boolean, ScanString As Boolean, ScanVir As Boolean, lstIma As ImageList, Picture1 As PictureBox, Picture2 As PictureBox) As String
'Thu tuc nay chi duoc su dung de quet process va cac file startup
'neu su dung thu tuc nay de quet cac thanh phan khac se cham hon binh thuong
'tai phien ban sap toi, Dung Coi se su dung thu tuc trong gan nhu tat ca cac thu tuc quet virus (Luc do se tem mot so phan tuy chon)

    On Error Resume Next
    Dim InputData As String
'If FileExists(FilePath) = True Then
'Neu su dung dong lenh tren, thi khi quet USB se xuat hien loi khong the thoat dia USB nay khoi PC

Dim i As Integer
'Tien hanh kiem tra file theo 3 thong so (Icon,MD5,String)
    If ScanMD5Main(GetMD5(FilePath), FilePath) = True Then GoTo KetThuc

    If ScanIcon = True Then
        Dim strTMP As String
        If UCase(Right(FilePath, 3)) = "EXE" Then
            strTMP = SoSanhImage(FilePath, Picture1, Picture2, lstIma)
            'Kiem tra file thong qua Icon (Chi kiem tra file exe)
            If strTMP <> "0" Then Detect GetStr("DecVirus"), "Virus : " & strTMP, FilePath: GoTo KetThuc
        End If
    End If

    If (ScanString = True) Or (ScanVir = True) Then
        Dim BoDem As String
        Open FilePath For Binary As #1
            BoDem = Space(LOF(1))
            Get #1, , BoDem
            Close #1
    Else
        GoTo KetThuc
    End If
    If ScanString = True Then
        'Kiem tra file thong qua cac chuoi String
        Dim strCodeS As String
        Dim strDecS As String
        
        Open App.Path & "\Dat\String.vnd" For Input As #1
        Do While Not EOF(1)
            Line Input #1, InputData
            
            strCodeS = Split(InputData, "|", , vbBinaryCompare)(0)
            If InStr(1, BoDem, strCodeS, vbBinaryCompare) <> 0 Then
                strDecS = Split(InputData, "|", , vbBinaryCompare)(1)
                Detect GetStr("DecFile"), strDecS, FilePath
                ph = ph + 1
                GoTo KetThuc
            End If
        Loop
            Close #1
    End If

    If ScanVir = True Then
        'Kiem tra file co phai la virus hay khong
        'quy trinh kiem tra giong kiem tra worm qua chuoi String, tuy nhien cong viec lai khac nhau
        'nen chung ta ne tach ra thanh 2 phan rieng biet
        For i = 0 To frmMnu.lstSVir.ListCount - 1
            If InStr(1, BoDem, frmMnu.lstSVir.List(i), vbBinaryCompare) <> 0 Then
                Detect GetStr("DecVir"), frmMnu.lstVirNa.List(i), FilePath, frmMnu.lstVirDat.List(i)
                GoTo KetThuc
            End If
        Next
    End If
        BoDem = ""
'Else

'End If
KetThuc:
End Function
Public Function ScanMD5Main(strCode As String, FilePath As String) As Boolean
'On Error Resume Next
    ScanMD5Main = False
If strCode = "0" Then GoTo KetThuc
    Dim InputData As String
    Dim PathTmp As String
    PathTmp = PathApp & "\Dat\Sign\" & GetExt(FilePath) & "\" & Left(strCode, 2) & ".vnd"
    If FileExists(PathTmp) = False Then Exit Function
    'Truong hop khong ton tai file (Truong hop nay tuc la ma MD5 cua file nay chua duoc cap nhat)
    'noi cach khac vnAV se khong canh bao ve file nay voi ma MD5
    
    Open PathTmp For Input As #1
        Do While Not EOF(1)
            Line Input #1, InputData
            If strCode = Split(InputData, "|", , vbBinaryCompare)(0) Then
                Detect GetStr("DecVirus"), "Virus: " & Split(InputData, "|", , vbBinaryCompare)(1), FilePath
                ph = ph + 1
                frmScan.lblPH.Caption = ph & " file"
                ScanMD5Main = True
                GoTo KetThuc
            End If
        Loop
            Close #1
KetThuc:
    Close #1

End Function
