VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "������������ ����"
   ClientHeight    =   2070
   ClientLeft      =   7965
   ClientTop       =   6300
   ClientWidth     =   3420
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   3420
   Begin VB.CommandButton Command4 
      Height          =   615
      Left            =   1560
      Picture         =   "frmMain.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "���������"
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Height          =   615
      Left            =   2640
      Picture         =   "frmMain.frx":170C
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "�����"
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Height          =   615
      Left            =   840
      Picture         =   "frmMain.frx":254E
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "� ���������"
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Default         =   -1  'True
      DisabledPicture =   "frmMain.frx":3218
      Height          =   615
      Left            =   120
      Picture         =   "frmMain.frx":508A
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "���������"
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox txtKar 
      Height          =   285
      Left            =   2640
      TabIndex        =   8
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox txtY 
      Height          =   285
      Left            =   1560
      TabIndex        =   7
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox txtM 
      Height          =   285
      Left            =   840
      TabIndex        =   6
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox txtD 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "�����"
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "���"
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "�����"
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "�����"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "������� ���� � �����"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dataTek, dataInp, dataR, ms, msg, ttl, ymx, kdy, dtIn, dtInp, msm, dm As String
Dim datTek, datInp, datR As Date
Dim yx, yxi, yt, ty, zy, mx, mt, dx, dt, ky, kd, kr, kar, kkr, d, m, y, SavData As Integer
Dim conf
Dim kVar, kName

Private Sub Command1_Click()
datTek = Now
kar = txtKar.Text
d = txtD.Text
m = txtM.Text
y = txtY.Text

'�������� ������� �� ����
If d = 0 Or m = 0 Or y = 0 Then
               Call ShowError1
               Exit Sub
End If

If y <> "" Then
    yxi = y Mod 4
Else
    Call ShowError1
    Exit Sub
End If

'�������� ������������ ����
If (m = 4 Or m = 6 Or m = 9 Or m = 11) And (d >= 31) Then
               Call ShowError3
               Exit Sub
ElseIf (m = 1 Or m = 3 Or m = 5 Or m = 7 Or m = 8 Or m = 10 Or m = 12) And (d >= 32) Then
               Call ShowError2
               Exit Sub
ElseIf m = 2 And yxi <> 0 And (d >= 29) Then
               Call ShowError5
                Exit Sub
ElseIf m = 2 And yxi = 0 And (d >= 30) Then
               Call ShowError4
               Exit Sub
ElseIf m <= 0 Or m >= 13 Or d <= 0 Or y = 0 Then
               Call ShowError1
               Exit Sub
End If


dataInp = d & "/" & m & "/" & y
datInp = CDate(dataInp)
If m = 1 Then
   msm = "������"
ElseIf m = 2 Then
   msm = "�������"
ElseIf m = 3 Then
   msm = "����"
ElseIf m = 4 Then
   msm = "������"
ElseIf m = 5 Then
   msm = "���"
ElseIf m = 6 Then
   msm = "����"
ElseIf m = 7 Then
   msm = "����"
ElseIf m = 8 Then
   msm = "������"
ElseIf m = 9 Then
   msm = "��������"
ElseIf m = 10 Then
   msm = "�������"
ElseIf m = 11 Then
   msm = "������"
Else: msm = "�������"
End If

dtIn = Weekday(datInp, vbMonday)
dm = d & "/" & m
yx = Year(datInp)
ky = yx Mod 16
kyf = yx Mod 4

'�������� ������������ �������� ������ �������
'� ���������� ������� ������������
If kar < 1 Or kar > 4 Then
    Call ShowErrorKar
    Exit Sub
End If

Select Case kVar
    Case 1
        Select Case kar
            Case 1: kkr = 0
            Case 2: kkr = -1
            Case 3: kkr = 1
            Case 4: kkr = 2
        End Select
        
    Case 2
        Select Case kar
            Case 1: kkr = -1
            Case 2: kkr = 1
            Case 3: kkr = 2
            Case 4: kkr = 0
        End Select
        
    Case 3
        Select Case kar
            Case 1: kkr = 1
            Case 2: kkr = 2
            Case 3: kkr = 0
            Case 4: kkr = -1
        End Select
        
    Case 4
        Select Case kar
            Case 1: kkr = 2
            Case 2: kkr = 0
            Case 3: kkr = -1
            Case 4: kkr = 1
        End Select
End Select

'If kVar = 1 Then
'    Select Case kar
'        Case 1: kkr = 0
'        Case 2: kkr = -1
'        Case 3: kkr = 1
'        Case 4: kkr = 2
'    End Select
'End If
'
'If kVar = 2 Then
'    Select Case kar
'        Case 1: kkr = -1
'        Case 2: kkr = 1
'        Case 3: kkr = 2
'        Case 4: kkr = 0
'    End Select
'End If
'
'If kVar = 3 Then
'    Select Case kar
'        Case 1: kkr = 1
'        Case 2: kkr = 2
'        Case 3: kkr = 0
'        Case 4: kkr = -1
'    End Select
'End If
'
'If kVar = 4 Then
'    Select Case kar
'        Case 1: kkr = 2
'        Case 2: kkr = 0
'        Case 3: kkr = -1
'        Case 4: kkr = 1
'    End Select
'End If

'�������� ������� ���� � ���
ymx = CStr(yx)
dataR = "1" + "/" + "1" + "/" + ymx
datR = CDate(dataR)
kdy = DateDiff("d", datR, datInp)
kd = kdy Mod 4

'��������� ku �� 29 �������
If (dm = "29/02" Or dm = "29/2") And ky = 4 Then ku = 0
If (dm = "29/02" Or dm = "29/2") And ky = 8 Then ku = 1
If (dm = "29/02" Or dm = "29/2") And ky = 12 Then ku = 2
If (dm = "29/02" Or dm = "29/2") And ky = 0 Then ku = 3

ky1 = (ky = 1 Or ky = 8 Or ky = 11 Or ky = 14)
ky2 = (ky = 2 Or ky = 5 Or ky = 12 Or ky = 15)
ky3 = (ky = 0 Or ky = 3 Or ky = 6 Or ky = 9)
ky4 = (ky = 4 Or ky = 7 Or ky = 10 Or ky = 13)
'��������� ku �� ��� ������ � ������� ����� 29 �������
If (dm <> "29/02" Or dm <> "29/2") And (m = 1 Or m = 2) And ky1 And kd = 3 Then ku = 0
If (dm <> "29/02" Or dm <> "29/2") And (m = 1 Or m = 2) And ky1 And kd = 0 Then ku = 1
If (dm <> "29/02" Or dm <> "29/2") And (m = 1 Or m = 2) And ky1 And kd = 1 Then ku = 2
If (dm <> "29/02" Or dm <> "29/2") And (m = 1 Or m = 2) And ky1 And kd = 2 Then ku = 3

If (dm <> "29/02" Or dm <> "29/2") And (m = 1 Or m = 2) And ky2 And kd = 2 Then ku = 0
If (dm <> "29/02" Or dm <> "29/2") And (m = 1 Or m = 2) And ky2 And kd = 3 Then ku = 1
If (dm <> "29/02" Or dm <> "29/2") And (m = 1 Or m = 2) And ky2 And kd = 0 Then ku = 2
If (dm <> "29/02" Or dm <> "29/2") And (m = 1 Or m = 2) And ky2 And kd = 1 Then ku = 3

If (dm <> "29/02" Or dm <> "29/2") And (m = 1 Or m = 2) And ky3 And kd = 1 Then ku = 0
If (dm <> "29/02" Or dm <> "29/2") And (m = 1 Or m = 2) And ky3 And kd = 2 Then ku = 1
If (dm <> "29/02" Or dm <> "29/2") And (m = 1 Or m = 2) And ky3 And kd = 3 Then ku = 2
If (dm <> "29/02" Or dm <> "29/2") And (m = 1 Or m = 2) And ky3 And kd = 0 Then ku = 3

If (dm <> "29/02" Or dm <> "29/2") And (m = 1 Or m = 2) And ky4 And kd = 0 Then ku = 0
If (dm <> "29/02" Or dm <> "29/2") And (m = 1 Or m = 2) And ky4 And kd = 1 Then ku = 1
If (dm <> "29/02" Or dm <> "29/2") And (m = 1 Or m = 2) And ky4 And kd = 2 Then ku = 2
If (dm <> "29/02" Or dm <> "29/2") And (m = 1 Or m = 2) And ky4 And kd = 3 Then ku = 3

'��������� ku �� ��� ������� ����� ������ � ������� �� ������� ���
kyO1 = (ky = 7 Or ky = 10 Or ky = 13)
kyO2 = (ky = 1 Or ky = 11 Or ky = 14)
kyO3 = (ky = 2 Or ky = 5 Or ky = 15)
kyO4 = (ky = 3 Or ky = 6 Or ky = 9)

If (m >= 3) And kyO1 And kd = 0 Then ku = 0
If (m >= 3) And kyO1 And kd = 1 Then ku = 1
If (m >= 3) And kyO1 And kd = 2 Then ku = 2
If (m >= 3) And kyO1 And kd = 3 Then ku = 3

If (m >= 3) And kyO2 And kd = 3 Then ku = 0
If (m >= 3) And kyO2 And kd = 0 Then ku = 1
If (m >= 3) And kyO2 And kd = 1 Then ku = 2
If (m >= 3) And kyO2 And kd = 2 Then ku = 3

If (m >= 3) And kyO3 And kd = 2 Then ku = 0
If (m >= 3) And kyO3 And kd = 3 Then ku = 1
If (m >= 3) And kyO3 And kd = 0 Then ku = 2
If (m >= 3) And kyO3 And kd = 1 Then ku = 3

If (m >= 3) And kyO4 And kd = 1 Then ku = 0
If (m >= 3) And kyO4 And kd = 2 Then ku = 1
If (m >= 3) And kyO4 And kd = 3 Then ku = 2
If (m >= 3) And kyO4 And kd = 0 Then ku = 3

'��������� ku �� ��� ������� ����� ������ � ������� �� ���������� ��� ������� ky=4
If (m >= 3) And (ky = 4) And kd = 0 Then ku = 0
If (m >= 3) And (ky = 4) And kd = 1 Then ku = 1
If (m >= 3) And (ky = 4) And kd = 2 Then ku = 2
If (m >= 3) And (ky = 4) And kd = 3 Then ku = 3

'��������� ku �� ��� ������� ����� ������ � ������� �� ���������� ��� ������� ky=8
If (m >= 3) And (ky = 8) And kd = 3 Then ku = 0
If (m >= 3) And (ky = 8) And kd = 0 Then ku = 1
If (m >= 3) And (ky = 8) And kd = 1 Then ku = 2
If (m >= 3) And (ky = 8) And kd = 2 Then ku = 3

'��������� ku �� ��� ������� ����� ������ � ������� �� ���������� ��� ������� ky=12
If (m >= 3) And (ky = 12) And kd = 2 Then ku = 0
If (m >= 3) And (ky = 12) And kd = 3 Then ku = 1
If (m >= 3) And (ky = 12) And kd = 0 Then ku = 2
If (m >= 3) And (ky = 12) And kd = 1 Then ku = 3

'��������� ku �� ��� ������� ����� ������ � ������� �� ���������� ��� ������� ky=0
If (m >= 3) And (ky = 0) And kd = 1 Then ku = 0
If (m >= 3) And (ky = 0) And kd = 2 Then ku = 1
If (m >= 3) And (ky = 0) And kd = 3 Then ku = 2
If (m >= 3) And (ky = 0) And kd = 0 Then ku = 3

'������ ������������ ����
kr = ku + kkr
If kr = -2 Or kr = 2 Then
                     ms = " ������ �����"
   ElseIf kr = 1 Or kr = 5 Then
                            ms = " ������� �����"
         ElseIf kr = 0 Or kr = 4 Then
                                 ms = " �������� ����"
               Else: ms = " �������� ����"
        End If

'���� ����������� ��� ������
If dtIn = "1" Then
               dtInp = "�����������"
    ElseIf dtIn = "2" Then
                     dtInp = "�������"
           ElseIf dtIn = "3" Then
                             dtInp = "�����"
                  ElseIf dtIn = "4" Then
                                    dtInp = "�������"
                         ElseIf dtIn = "5" Then
                                           dtInp = "�������"
                                ElseIf dtIn = "6" Then
                                                  dtInp = "�������"
                                       ElseIf dtIn = "7" Then
                                                         dtInp = "�����������"
End If

    '����� ��������� � �����
    dataInp = Format(datInp, "dd mmmm yyyy")
    ms = "����� " & dataInp & " (" & dtInp & ")" & vbCrLf & kar & " " & kName & " �������� �" & vbCrLf & ms
    
    'ms = msg
    ttl = "����������� �����"
    ms = MsgBox(ms, vbInformation, ttl)
    
    conf = GetConfig("FormSet", "DatVal", IniFile)
    If conf = 1 Then
        conf = SaveConfig("FormSet", "Day", txtD.Text, IniFile)
        conf = SaveConfig("FormSet", "Month", txtM.Text, IniFile)
        conf = SaveConfig("FormSet", "Year", txtY.Text, IniFile)
    End If

End Sub
Private Sub ShowError1()
'������ �������� ����
ms = "������� ���������� ����"
ttl = "������"
ms = MsgBox(ms, vbCritical, ttl)
End Sub

Private Sub ShowError2()
'������ �������� ����
ms = "������-�� " & msm & " ������������� 31-�� �����!"
ttl = "������"
ms = MsgBox(ms, vbCritical, ttl)
End Sub

Private Sub ShowError3()
'������ �������� ����
ms = "������-�� " & msm & " ������������� 30-�� �����!"
ttl = "������"
ms = MsgBox(ms, vbCritical, ttl)
End Sub

Private Sub ShowError4()
'������ �������� ����
ms = "������ ��� ���������� � ������� ������������� 29-�� �����"
ttl = "������"
ms = MsgBox(ms, vbCritical, ttl)
End Sub

Private Sub ShowError5()
'������ �������� ����
ms = "������ ��� ������� � ������� ������������� 28-�� �����"
ttl = "������"
ms = MsgBox(ms, vbCritical, ttl)
End Sub

Private Sub ShowErrorKar()
'������ ������������ �������
ms = "������� ����� ������� � ��������� �� 1 �� 4"
ttl = "������"
ms = MsgBox(ms, vbCritical, ttl)
End Sub

Private Sub Command2_Click()
frmAbout.Show
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Command4_Click()
    frmSet.Show
End Sub

Private Sub Form_Load()
    txtKar.Text = GetConfig("FormSet", "Grup", IniFile)
    txtD.Text = GetConfig("FormSet", "Day", IniFile)
    txtM.Text = GetConfig("FormSet", "Month", IniFile)
    txtY.Text = GetConfig("FormSet", "Year", IniFile)
    Me.Top = GetConfig("StartUp", "Top", IniFile)
    Me.Left = GetConfig("StartUp", "Left", IniFile)
    kName = GetConfig("FormSet", "KarName", IniFile)
    Label5.Caption = kName
    kVar = GetConfig("FormSet", "KarVar", IniFile)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    conf = SaveConfig("StartUp", "Top", Me.Top, IniFile)
    conf = SaveConfig("StartUp", "Left", Me.Left, IniFile)
    conf = GetConfig("FormSet", "DatVal", IniFile)
    If conf = 1 Then
        conf = SaveConfig("FormSet", "Day", txtD.Text, IniFile)
        conf = SaveConfig("FormSet", "Month", txtM.Text, IniFile)
        conf = SaveConfig("FormSet", "Year", txtY.Text, IniFile)
    End If
End Sub

Private Sub txtD_GotFocus()
    txtD.SelLength = 2
End Sub

Private Sub txtKar_GotFocus()
    txtKar.SelLength = 1
End Sub

Private Sub txtM_GotFocus()
    txtM.SelLength = 2
End Sub

Private Sub txtY_GotFocus()
    txtY.SelLength = 4
End Sub

