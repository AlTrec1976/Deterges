VERSION 5.00
Begin VB.Form frmSet 
   Caption         =   "Параметры"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4755
   Icon            =   "frmSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   4755
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Отмена"
      Height          =   375
      Left            =   3360
      TabIndex        =   14
      Top             =   3000
      Width           =   1215
   End
   Begin VB.OptionButton optVar 
      Caption         =   "Вариант 2"
      Height          =   255
      Index           =   3
      Left            =   2520
      TabIndex        =   10
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Выберите смену, соответствующую вам на данный день"
      Height          =   1095
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   4575
      Begin VB.OptionButton optVar 
         Caption         =   "Вариант 2"
         Height          =   255
         Index           =   2
         Left            =   2400
         TabIndex        =   13
         Top             =   360
         Width           =   2055
      End
      Begin VB.OptionButton optVar 
         Caption         =   "Вариант 2"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   2055
      End
      Begin VB.OptionButton optVar 
         Caption         =   "Вариант 2"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   2055
      End
   End
   Begin VB.TextBox txtKarName 
      Height          =   285
      Left            =   1920
      TabIndex        =   8
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton btnCan 
      Caption         =   "Отмена"
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "ОК"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txtSetKar 
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Top             =   1200
      Width           =   375
   End
   Begin VB.CheckBox chcData 
      Caption         =   "Запоминать последнюю дату"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "Смена"
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Обозначение смены"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Всегда отображается"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Отрегулируйте установки по своему усмотрению"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "frmSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conf
Dim conf2 As String
Dim SavData As Integer

Private Sub btnCan_Click()
    Unload Me
End Sub

Private Sub btnOk_Click()

    conf = SaveConfig("FormSet", "Grup", txtSetKar.Text, IniFile)
    conf = SaveConfig("FormSet", "DatVal", chcData.Value, IniFile)
    conf = SaveConfig("FormSet", "KarName", txtKarName.Text, IniFile)
    If optVar(0).Value = True Then conf2 = "1"
    If optVar(1).Value = True Then conf2 = "2"
    If optVar(2).Value = True Then conf2 = "3"
    If optVar(3).Value = True Then conf2 = "4"
    conf = SaveConfig("FormSet", "KarVar", conf2, IniFile)
    
    SavData = GetConfig("FormSet", "DatVal", IniFile)
    If SavData = 0 Then
        conf = SaveConfig("FormSet", "Day", "", IniFile)
        conf = SaveConfig("FormSet", "Month", "", IniFile)
        conf = SaveConfig("FormSet", "Year", "", IniFile)
    End If
    
    Unload Me
End Sub

Private Sub chcData_Click()
    btnOk.Enabled = True
    btnOk.Default = True
    btnCan.Default = False
End Sub

Private Sub Form_Load()
    conf2 = "Выберите смену, соответствующую вам на сегодняшний день"
    
    conf = GetConfig("FormSet", "KarVar", IniFile)
    If conf = 1 Then optVar(0).Value = True
    If conf = 2 Then optVar(1).Value = True
    If conf = 3 Then optVar(2).Value = True
    If conf = 4 Then optVar(3).Value = True
    
    txtKarName.Text = GetConfig("FormSet", "KarName", IniFile)
    Label4.Caption = GetConfig("FormSet", "KarName", IniFile)
    txtSetKar.Text = GetConfig("FormSet", "Grup", IniFile)
    chcData.Value = GetConfig("FormSet", "DatVal", IniFile)
    Me.Top = GetConfig("FormSet", "Top", IniFile)
    Me.Left = GetConfig("FormSet", "Left", IniFile)
    optVar(0).ToolTipText = conf2
    optVar(1).ToolTipText = conf2
    optVar(2).ToolTipText = conf2
    optVar(3).ToolTipText = conf2
    
    kar = txtSetKar.Text
    optVar(0).Caption = Deterges(1, kar)
    optVar(1).Caption = Deterges(2, kar)
    optVar(2).Caption = Deterges(3, kar)
    optVar(3).Caption = Deterges(4, kar)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    conf = SaveConfig("FormSet", "Top", Me.Top, IniFile)
    conf = SaveConfig("FormSet", "Left", Me.Left, IniFile)
End Sub

Private Sub optVar_Click(Index As Integer)
    Select Case Index
        Case 0
            kar = txtSetKar.Text
            optVar(0).Value = True
            optVar(1).Value = False
            optVar(2).Value = False
            optVar(3).Value = False
        Case 1
            kar = txtSetKar.Text
            optVar(0).Value = False
            optVar(1).Value = True
            optVar(2).Value = False
            optVar(3).Value = False
        Case 2
            kar = txtSetKar.Text
            optVar(0).Value = False
            optVar(1).Value = False
            optVar(2).Value = True
            optVar(3).Value = False
        Case 3
            kar = txtSetKar.Text
            optVar(0).Value = False
            optVar(1).Value = False
            optVar(2).Value = False
            optVar(3).Value = True
    End Select
End Sub

Private Sub txtKarName_LostFocus()
    Label4.Caption = txtKarName.Text
End Sub

Private Sub txtSetKar_Change()
    btnOk.Enabled = True
    btnOk.Default = True
    btnCan.Default = False
    txtSetKar.SelLength = 1
    btnOk.Default = True
    btnCan.Default = False
End Sub

Function Deterges(kVar As Variant, kar As Variant) As String

datTek = Now
d = CInt(Day(Now))
m = CInt(Month(Now))
y = CInt(Year(Now))

    yxi = y Mod 4

dataInp = d & "/" & m & "/" & y
datInp = CDate(dataInp)
dtIn = Weekday(datInp, vbMonday)
dm = d & "/" & m
yx = Year(datInp)
ky = yx Mod 16
kyf = yx Mod 4

'проверка правильности введения номера караула
'и присвоение караулу коэффициента
If kVar = 1 Then
    If kar = 1 Then
           kkr = 0
    ElseIf kar = 2 Then
           kkr = -1
    ElseIf kar = 3 Then
           kkr = 1
    ElseIf kar = 4 Then
           kkr = 2
    End If
End If
If kVar = 2 Then
    If kar = 1 Then
           kkr = -1
    ElseIf kar = 2 Then
           kkr = 1
    ElseIf kar = 3 Then
           kkr = 2
    ElseIf kar = 4 Then
           kkr = 0
    End If
End If
If kVar = 3 Then
    If kar = 1 Then
           kkr = 1
    ElseIf kar = 2 Then
           kkr = 2
    ElseIf kar = 3 Then
           kkr = 0
    ElseIf kar = 4 Then
           kkr = -1
    End If
End If
If kVar = 4 Then
    If kar = 1 Then
           kkr = 2
    ElseIf kar = 2 Then
           kkr = 0
    ElseIf kar = 3 Then
           kkr = -1
    ElseIf kar = 4 Then
           kkr = 1
    End If
End If

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
                     ms = "ночная смена"
   ElseIf kr = 1 Or kr = 5 Then
                            ms = "дневная смена"
         ElseIf kr = 0 Or kr = 4 Then
                                 ms = "выходной день"
               Else: ms = "отсыпной день"
        End If
Deterges = ms
End Function

Private Sub txtSetKar_LostFocus()
    kar = txtSetKar.Text
    optVar(0).Caption = Deterges(1, kar)
    optVar(1).Caption = Deterges(2, kar)
    optVar(2).Caption = Deterges(3, kar)
    optVar(3).Caption = Deterges(4, kar)
End Sub
