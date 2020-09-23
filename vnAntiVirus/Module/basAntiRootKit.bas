Attribute VB_Name = "basAntiRootKit"
'vnAntiVirus 1.0

'Author : Dung Le Nguyen
'Email : dungcoivb@gmail.com
'This is a software open source

'I special thank PhamTienSinh this method

'This module use to Rootkit with method get handle :)
'The method simple but effect
'You can read introduce to see it test with some samples Rootkit

'Module nay, toi chan thanh cam on PhamTienSinh (Pham Trung Hai)
'Ok, mot phuong phap nhan dang RootKit mot cach that don gian va tuyet voi
'Cam on PTS rat nhieu voi ky thuat nay

Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Boolean
Public Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Boolean
    Dim ID As Long
    Dim CoChua As Boolean
    Dim i As Integer
    CoChua = False
    GetWindowThreadProcessId hwnd, ID
    
    With frmMnu.lstPro
        For i = 0 To .ListCount - 1
            If ID = Val(.List(i)) Then CoChua = True
        Next
        If CoChua = False Then .AddItem ID
    End With
    
If CoChua = False Then
    If CheckID(ID) <> ID Then
        Dim tmp As String
        tmp = ProcessPathByPID(ID)
    
        CoChua = False
        With frmPro

            Set lsv = .LV.ListItems.Add()
            lsv.Text = GetFileName(tmp)
            lsv.SubItems(1) = tmp
            lsv.SubItems(2) = ID
            lsv.ForeColor = vbRed
            
        End With
        
    End If
End If

    EnumWindowsProc = True
End Function
