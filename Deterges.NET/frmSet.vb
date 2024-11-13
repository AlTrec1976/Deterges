Option Strict Off
Option Explicit On
Friend Class frmSet
	Inherits System.Windows.Forms.Form
#Region "Код, автоматически созданный конструктором форм Windows "
	Public Sub New()
		MyBase.New()
		If m_vb6FormDefInstance Is Nothing Then
			If m_InitializingDefInstance Then
				m_vb6FormDefInstance = Me
			Else
				Try 
					'Для начальной формы первый созданный экземпляр является экземпляром по умолчанию.
					If System.Reflection.Assembly.GetExecutingAssembly.EntryPoint.DeclaringType Is Me.GetType Then
						m_vb6FormDefInstance = Me
					End If
				Catch
				End Try
			End If
		End If
		'Этот вызов установлен конструктором форм Windows.
		InitializeComponent()
	End Sub
	'Форма переопределяет метод Dispose для очистки списка компонентов.
	Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Требуется конструктором форм Windows
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents btnCan As System.Windows.Forms.Button
	Public WithEvents btnOk As System.Windows.Forms.Button
	Public WithEvents txtSetKar As System.Windows.Forms.TextBox
	Public WithEvents chcData As System.Windows.Forms.CheckBox
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	'Примечание. Этот вызов установлен конструктором форм Windows
	'Можно изменить с помощью конструктора форм Windows.
	'Не изменять с помощью редактора исходного текста.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmSet))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ToolTip1.Active = True
		Me.btnCan = New System.Windows.Forms.Button
		Me.btnOk = New System.Windows.Forms.Button
		Me.txtSetKar = New System.Windows.Forms.TextBox
		Me.chcData = New System.Windows.Forms.CheckBox
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.Text = "Параметры"
		Me.ClientSize = New System.Drawing.Size(312, 206)
		Me.Location = New System.Drawing.Point(4, 30)
		Me.Icon = CType(resources.GetObject("frmSet.Icon"), System.Drawing.Icon)
		Me.MaximizeBox = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.MinimizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmSet"
		Me.btnCan.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.btnCan
		Me.btnCan.Text = "Отмена"
		Me.btnCan.Size = New System.Drawing.Size(81, 25)
		Me.btnCan.Location = New System.Drawing.Point(176, 168)
		Me.btnCan.TabIndex = 5
		Me.btnCan.BackColor = System.Drawing.SystemColors.Control
		Me.btnCan.CausesValidation = True
		Me.btnCan.Enabled = True
		Me.btnCan.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnCan.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnCan.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnCan.TabStop = True
		Me.btnCan.Name = "btnCan"
		Me.btnOk.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.btnOk.Text = "ОК"
		Me.AcceptButton = Me.btnOk
		Me.btnOk.Size = New System.Drawing.Size(81, 25)
		Me.btnOk.Location = New System.Drawing.Point(48, 168)
		Me.btnOk.TabIndex = 4
		Me.btnOk.BackColor = System.Drawing.SystemColors.Control
		Me.btnOk.CausesValidation = True
		Me.btnOk.Enabled = True
		Me.btnOk.ForeColor = System.Drawing.SystemColors.ControlText
		Me.btnOk.Cursor = System.Windows.Forms.Cursors.Default
		Me.btnOk.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.btnOk.TabStop = True
		Me.btnOk.Name = "btnOk"
		Me.txtSetKar.AutoSize = False
		Me.txtSetKar.Size = New System.Drawing.Size(33, 19)
		Me.txtSetKar.Location = New System.Drawing.Point(56, 112)
		Me.txtSetKar.TabIndex = 3
		Me.txtSetKar.AcceptsReturn = True
		Me.txtSetKar.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtSetKar.BackColor = System.Drawing.SystemColors.Window
		Me.txtSetKar.CausesValidation = True
		Me.txtSetKar.Enabled = True
		Me.txtSetKar.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtSetKar.HideSelection = True
		Me.txtSetKar.ReadOnly = False
		Me.txtSetKar.Maxlength = 0
		Me.txtSetKar.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtSetKar.MultiLine = False
		Me.txtSetKar.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtSetKar.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtSetKar.TabStop = True
		Me.txtSetKar.Visible = True
		Me.txtSetKar.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtSetKar.Name = "txtSetKar"
		Me.chcData.Text = "Запоминать последнюю дату"
		Me.chcData.Size = New System.Drawing.Size(185, 25)
		Me.chcData.Location = New System.Drawing.Point(8, 56)
		Me.chcData.TabIndex = 0
		Me.chcData.CheckAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.chcData.BackColor = System.Drawing.SystemColors.Control
		Me.chcData.CausesValidation = True
		Me.chcData.Enabled = True
		Me.chcData.ForeColor = System.Drawing.SystemColors.ControlText
		Me.chcData.Cursor = System.Windows.Forms.Cursors.Default
		Me.chcData.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.chcData.Appearance = System.Windows.Forms.Appearance.Normal
		Me.chcData.TabStop = True
		Me.chcData.CheckState = System.Windows.Forms.CheckState.Unchecked
		Me.chcData.Visible = True
		Me.chcData.Name = "chcData"
		Me.Label2.Text = "Смена"
		Me.Label2.Size = New System.Drawing.Size(41, 17)
		Me.Label2.Location = New System.Drawing.Point(8, 112)
		Me.Label2.TabIndex = 2
		Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label2.BackColor = System.Drawing.SystemColors.Control
		Me.Label2.Enabled = True
		Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label2.UseMnemonic = True
		Me.Label2.Visible = True
		Me.Label2.AutoSize = False
		Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label2.Name = "Label2"
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.Label1.Text = "Отрегулируйте установки по своему усмотрению"
		Me.Label1.Size = New System.Drawing.Size(313, 33)
		Me.Label1.Location = New System.Drawing.Point(0, 0)
		Me.Label1.TabIndex = 1
		Me.Label1.BackColor = System.Drawing.SystemColors.Control
		Me.Label1.Enabled = True
		Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = False
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.Controls.Add(btnCan)
		Me.Controls.Add(btnOk)
		Me.Controls.Add(txtSetKar)
		Me.Controls.Add(chcData)
		Me.Controls.Add(Label2)
		Me.Controls.Add(Label1)
	End Sub
#End Region 
#Region "Поддержка обновления "
	Private Shared m_vb6FormDefInstance As frmSet
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As frmSet
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New frmSet()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region 
	Dim conf As Object
	
	Private Sub btnCan_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnCan.Click
		Me.Close()
	End Sub
	
	Private Sub btnOk_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles btnOk.Click
		Dim SavData As Short
		
		'UPGRADE_WARNING: Не удается разрешить свойство по умолчанию объекта conf. Дополнительно: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		conf = SaveConfig("FormSet", "Grup", (txtSetKar.Text), IniFile)
		'UPGRADE_WARNING: Не удается разрешить свойство по умолчанию объекта conf. Дополнительно: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		conf = SaveConfig("FormSet", "DatVal", CStr(chcData.CheckState), IniFile)
		
		SavData = CShort(GetConfig("FormSet", "DatVal", IniFile))
		If SavData = 0 Then
			'UPGRADE_WARNING: Не удается разрешить свойство по умолчанию объекта conf. Дополнительно: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			conf = SaveConfig("FormSet", "Day", "", IniFile)
			'UPGRADE_WARNING: Не удается разрешить свойство по умолчанию объекта conf. Дополнительно: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			conf = SaveConfig("FormSet", "Month", "", IniFile)
			'UPGRADE_WARNING: Не удается разрешить свойство по умолчанию объекта conf. Дополнительно: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			conf = SaveConfig("FormSet", "Year", "", IniFile)
		End If
		
		Me.Close()
	End Sub
	
	'UPGRADE_WARNING: Событие chcData.CheckStateChanged может появиться при инициализации формы. Дополнительно: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2075"'
	Private Sub chcData_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chcData.CheckStateChanged
		btnOk.Enabled = True
		VB6.SetDefault(btnOk, True)
		VB6.SetDefault(btnCan, False)
	End Sub
	
	Private Sub frmSet_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		txtSetKar.Text = GetConfig("FormSet", "Grup", IniFile)
		chcData.CheckState = CShort(GetConfig("FormSet", "DatVal", IniFile))
		Me.Top = VB6.TwipsToPixelsY(CSng(GetConfig("FormSet", "Top", IniFile)))
		Me.Left = VB6.TwipsToPixelsX(CSng(GetConfig("FormSet", "Left", IniFile)))
	End Sub
	
	'UPGRADE_WARNING: Form событие frmSet.Unload имеет новое поведение. Дополнительно: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2065"'
	Private Sub frmSet_Closed(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Closed
		'UPGRADE_WARNING: Не удается разрешить свойство по умолчанию объекта conf. Дополнительно: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		conf = SaveConfig("FormSet", "Top", CStr(VB6.PixelsToTwipsY(Me.Top)), IniFile)
		'UPGRADE_WARNING: Не удается разрешить свойство по умолчанию объекта conf. Дополнительно: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		conf = SaveConfig("FormSet", "Left", CStr(VB6.PixelsToTwipsX(Me.Left)), IniFile)
	End Sub
	
	'UPGRADE_WARNING: Событие txtSetKar.TextChanged может появиться при инициализации формы. Дополнительно: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2075"'
	Private Sub txtSetKar_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSetKar.TextChanged
		btnOk.Enabled = True
		VB6.SetDefault(btnOk, True)
		VB6.SetDefault(btnCan, False)
		txtSetKar.SelectionLength = 1
		VB6.SetDefault(btnOk, True)
		VB6.SetDefault(btnCan, False)
	End Sub
End Class