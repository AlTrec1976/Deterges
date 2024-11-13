VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Определитель смен"
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
      ToolTipText     =   "Установки"
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
      ToolTipText     =   "Выход"
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Height          =   615
      Left            =   840
      Picture         =   "frmMain.frx":254E
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "О программе"
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
      ToolTipText     =   "Расчитать"
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
      Caption         =   "Смена"
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Год"
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Месяц"
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Число"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Введите дату и смену"
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

'проверим введена ли дата
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

'проверим корректность даты
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
   msm = "январь"
ElseIf m = 2 Then
   msm = "февраль"
ElseIf m = 3 Then
   msm = "март"
ElseIf m = 4 Then
   msm = "апрель"
ElseIf m = 5 Then
   msm = "май"
ElseIf m = 6 Then
   msm = "июнь"
ElseIf m = 7 Then
   msm = "июль"
ElseIf m = 8 Then
   msm = "август"
ElseIf m = 9 Then
   msm = "сентябрь"
ElseIf m = 10 Then
   msm = "октябрь"
ElseIf m = 11 Then
   msm = "ноябрь"
Else: msm = "декабрь"
End If

dtIn = Weekday(datInp, vbMonday)
dm = d & "/" & m
yx = Year(datInp)
ky = yx Mod 16
kyf = yx Mod 4

'проверка правильности введения номера караула
'и присвоение караулу коэффициента
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

'алгоритм расчета дней и лет
ymx = CStr(yx)
dataR = "1" + "/" + "1" + "/" + ymx
datR = CDate(dataR)
kdy = DateDiff("d", datR, datInp)
kd = kdy Mod 4

'Определим ku на 29 февраля
If (dm = "29/02" Or dm = "29/2") And ky = 4 Then ku = 0
If (dm = "29/02" Or dm = "29/2") And ky = 8 Then ku = 1
If (dm = "29/02" Or dm = "29/2") And ky = 12 Then ku = 2
If (dm = "29/02" Or dm = "29/2") And ky = 0 Then ku = 3

ky1 = (ky = 1 Or ky = 8 Or ky = 11 Or ky = 14)
ky2 = (ky = 2 Or ky = 5 Or ky = 12 Or ky = 15)
ky3 = (ky = 0 Or ky = 3 Or ky = 6 Or ky = 9)
ky4 = (ky = 4 Or ky = 7 Or ky = 10 Or ky = 13)
'Определим ku на дни января и февраля кроме 29 февраля
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

'Определим ku на дни месяцев кроме января и февраля на обычный год
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

'Определим ku на дни месяцев кроме января и февраля на высокосный год имеющий ky=4
If (m >= 3) And (ky = 4) And kd = 0 Then ku = 0
If (m >= 3) And (ky = 4) And kd = 1 Then ku = 1
If (m >= 3) And (ky = 4) And kd = 2 Then ku = 2
If (m >= 3) And (ky = 4) And kd = 3 Then ku = 3

'Определим ku на дни месяцев кроме января и февраля на высокосный год имеющий ky=8
If (m >= 3) And (ky = 8) And kd = 3 Then ku = 0
If (m >= 3) And (ky = 8) And kd = 0 Then ku = 1
If (m >= 3) And (ky = 8) And kd = 1 Then ku = 2
If (m >= 3) And (ky = 8) And kd = 2 Then ku = 3

'Определим ku на дни месяцев кроме января и февраля на высокосный год имеющий ky=12
If (m >= 3) And (ky = 12) And kd = 2 Then ku = 0
If (m >= 3) And (ky = 12) And kd = 3 Then ku = 1
If (m >= 3) And (ky = 12) And kd = 0 Then ku = 2
If (m >= 3) And (ky = 12) And kd = 1 Then ku = 3

'Определим ku на дни месяцев кроме января и февраля на высокосный год имеющий ky=0
If (m >= 3) And (ky = 0) And kd = 1 Then ku = 0
If (m >= 3) And (ky = 0) And kd = 2 Then ku = 1
If (m >= 3) And (ky = 0) And kd = 3 Then ku = 2
If (m >= 3) And (ky = 0) And kd = 0 Then ku = 3

'расчет соответствия смен
kr = ku + kkr
If kr = -2 Or kr = 2 Then
                     ms = " ночную смену"
   ElseIf kr = 1 Or kr = 5 Then
                            ms = " дневную смену"
         ElseIf kr = 0 Or kr = 4 Then
                                 ms = " выходной день"
               Else: ms = " отсыпной день"
        End If

'блок определения дня недели
If dtIn = "1" Then
               dtInp = "понедельник"
    ElseIf dtIn = "2" Then
                     dtInp = "вторник"
           ElseIf dtIn = "3" Then
                             dtInp = "среда"
                  ElseIf dtIn = "4" Then
                                    dtInp = "четверг"
                         ElseIf dtIn = "5" Then
                                           dtInp = "пятница"
                                ElseIf dtIn = "6" Then
                                                  dtInp = "суббота"
                                       ElseIf dtIn = "7" Then
                                                         dtInp = "воскресенье"
End If

    'вывод сообщения о смене
    dataInp = Format(datInp, "dd mmmm yyyy")
    ms = "Число " & dataInp & " (" & dtInp & ")" & vbCrLf & kar & " " & kName & " вступает в" & vbCrLf & ms
    
    'ms = msg
    ttl = "Определение смены"
    ms = MsgBox(ms, vbInformation, ttl)
    
    conf = GetConfig("FormSet", "DatVal", IniFile)
    If conf = 1 Then
        conf = SaveConfig("FormSet", "Day", txtD.Text, IniFile)
        conf = SaveConfig("FormSet", "Month", txtM.Text, IniFile)
        conf = SaveConfig("FormSet", "Year", txtY.Text, IniFile)
    End If

End Sub
Private Sub ShowError1()
'ошибка введения даты
ms = "Введите корректную дату"
ttl = "Ошибка"
ms = MsgBox(ms, vbCritical, ttl)
End Sub

Private Sub ShowError2()
'ошибка введения даты
ms = "Вообще-то " & msm & " заканчивается 31-го числа!"
ttl = "Ошибка"
ms = MsgBox(ms, vbCritical, ttl)
End Sub

Private Sub ShowError3()
'ошибка введения даты
ms = "Вообще-то " & msm & " заканчивается 30-го числа!"
ttl = "Ошибка"
ms = MsgBox(ms, vbCritical, ttl)
End Sub

Private Sub ShowError4()
'ошибка введения даты
ms = "Данный год високосный и февраль заканчивается 29-го числа"
ttl = "Ошибка"
ms = MsgBox(ms, vbCritical, ttl)
End Sub

Private Sub ShowError5()
'ошибка введения даты
ms = "Данный год обычный и февраль заканчивается 28-го числа"
ttl = "Ошибка"
ms = MsgBox(ms, vbCritical, ttl)
End Sub

Private Sub ShowErrorKar()
'ошибка соответствия караула
ms = "Введите номер караула в диапазоне от 1 до 4"
ttl = "Ошибка"
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

