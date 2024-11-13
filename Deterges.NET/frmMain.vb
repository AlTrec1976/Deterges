Option Strict Off
Option Explicit On
Friend Class frmMain
	Inherits System.Windows.Forms.Form
#Region "���, ������������� ��������� ������������� ���� Windows "
	Public Sub New()
		MyBase.New()
		If m_vb6FormDefInstance Is Nothing Then
			If m_InitializingDefInstance Then
				m_vb6FormDefInstance = Me
			Else
				Try 
					'��� ��������� ����� ������ ��������� ��������� �������� ����������� �� ���������.
					If System.Reflection.Assembly.GetExecutingAssembly.EntryPoint.DeclaringType Is Me.GetType Then
						m_vb6FormDefInstance = Me
					End If
				Catch
				End Try
			End If
		End If
		'���� ����� ���������� ������������� ���� Windows.
		InitializeComponent()
	End Sub
	'����� �������������� ����� Dispose ��� ������� ������ �����������.
	Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'��������� ������������� ���� Windows
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents Command4 As System.Windows.Forms.Button
	Public WithEvents Command3 As System.Windows.Forms.Button
	Public WithEvents Command2 As System.Windows.Forms.Button
	Public WithEvents Command1 As System.Windows.Forms.Button
	Public WithEvents txtKar As System.Windows.Forms.TextBox
	Public WithEvents txtY As System.Windows.Forms.TextBox
	Public WithEvents txtM As System.Windows.Forms.TextBox
	Public WithEvents txtD As System.Windows.Forms.TextBox
	Public WithEvents Label5 As System.Windows.Forms.Label
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	'����������. ���� ����� ���������� ������������� ���� Windows
	'����� �������� � ������� ������������ ���� Windows.
	'�� �������� � ������� ��������� ��������� ������.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmMain))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ToolTip1.Active = True
		Me.Command4 = New System.Windows.Forms.Button
		Me.Command3 = New System.Windows.Forms.Button
		Me.Command2 = New System.Windows.Forms.Button
		Me.Command1 = New System.Windows.Forms.Button
		Me.txtKar = New System.Windows.Forms.TextBox
		Me.txtY = New System.Windows.Forms.TextBox
		Me.txtM = New System.Windows.Forms.TextBox
		Me.txtD = New System.Windows.Forms.TextBox
		Me.Label5 = New System.Windows.Forms.Label
		Me.Label4 = New System.Windows.Forms.Label
		Me.Label3 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
		Me.Text = "������������ ����"
		Me.ClientSize = New System.Drawing.Size(228, 138)
		Me.Location = New System.Drawing.Point(531, 420)
		Me.Icon = CType(resources.GetObject("frmMain.Icon"), System.Drawing.Icon)
		Me.MaximizeBox = False
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
		Me.Name = "frmMain"
		Me.Command4.TextAlign = System.Drawing.ContentAlignment.BottomCenter
		Me.Command4.Size = New System.Drawing.Size(41, 41)
		Me.Command4.Location = New System.Drawing.Point(104, 88)
		Me.Command4.Image = CType(resources.GetObject("Command4.Image"), System.Drawing.Image)
		Me.Command4.TabIndex = 12
		Me.ToolTip1.SetToolTip(Me.Command4, "���������")
		Me.Command4.BackColor = System.Drawing.SystemColors.Control
		Me.Command4.CausesValidation = True
		Me.Command4.Enabled = True
		Me.Command4.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Command4.Cursor = System.Windows.Forms.Cursors.Default
		Me.Command4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Command4.TabStop = True
		Me.Command4.Name = "Command4"
		Me.Command3.TextAlign = System.Drawing.ContentAlignment.BottomCenter
		Me.CancelButton = Me.Command3
		Me.Command3.Size = New System.Drawing.Size(41, 41)
		Me.Command3.Location = New System.Drawing.Point(176, 88)
		Me.Command3.Image = CType(resources.GetObject("Command3.Image"), System.Drawing.Image)
		Me.Command3.TabIndex = 11
		Me.ToolTip1.SetToolTip(Me.Command3, "�����")
		Me.Command3.BackColor = System.Drawing.SystemColors.Control
		Me.Command3.CausesValidation = True
		Me.Command3.Enabled = True
		Me.Command3.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Command3.Cursor = System.Windows.Forms.Cursors.Default
		Me.Command3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Command3.TabStop = True
		Me.Command3.Name = "Command3"
		Me.Command2.TextAlign = System.Drawing.ContentAlignment.BottomCenter
		Me.Command2.Size = New System.Drawing.Size(41, 41)
		Me.Command2.Location = New System.Drawing.Point(56, 88)
		Me.Command2.Image = CType(resources.GetObject("Command2.Image"), System.Drawing.Image)
		Me.Command2.TabIndex = 10
		Me.ToolTip1.SetToolTip(Me.Command2, "� ���������")
		Me.Command2.BackColor = System.Drawing.SystemColors.Control
		Me.Command2.CausesValidation = True
		Me.Command2.Enabled = True
		Me.Command2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Command2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Command2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Command2.TabStop = True
		Me.Command2.Name = "Command2"
		Me.Command1.TextAlign = System.Drawing.ContentAlignment.BottomCenter
		Me.AcceptButton = Me.Command1
		Me.Command1.Size = New System.Drawing.Size(41, 41)
		Me.Command1.Location = New System.Drawing.Point(8, 88)
		Me.Command1.Image = CType(resources.GetObject("Command1.Image"), System.Drawing.Image)
		Me.Command1.TabIndex = 9
		Me.ToolTip1.SetToolTip(Me.Command1, "���������")
		Me.Command1.BackColor = System.Drawing.SystemColors.Control
		Me.Command1.CausesValidation = True
		Me.Command1.Enabled = True
		Me.Command1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Command1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Command1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Command1.TabStop = True
		Me.Command1.Name = "Command1"
		Me.txtKar.AutoSize = False
		Me.txtKar.Size = New System.Drawing.Size(33, 19)
		Me.txtKar.Location = New System.Drawing.Point(184, 56)
		Me.txtKar.TabIndex = 8
		Me.txtKar.AcceptsReturn = True
		Me.txtKar.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtKar.BackColor = System.Drawing.SystemColors.Window
		Me.txtKar.CausesValidation = True
		Me.txtKar.Enabled = True
		Me.txtKar.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtKar.HideSelection = True
		Me.txtKar.ReadOnly = False
		Me.txtKar.Maxlength = 0
		Me.txtKar.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtKar.MultiLine = False
		Me.txtKar.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtKar.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtKar.TabStop = True
		Me.txtKar.Visible = True
		Me.txtKar.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtKar.Name = "txtKar"
		Me.txtY.AutoSize = False
		Me.txtY.Size = New System.Drawing.Size(33, 19)
		Me.txtY.Location = New System.Drawing.Point(104, 56)
		Me.txtY.TabIndex = 7
		Me.txtY.AcceptsReturn = True
		Me.txtY.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtY.BackColor = System.Drawing.SystemColors.Window
		Me.txtY.CausesValidation = True
		Me.txtY.Enabled = True
		Me.txtY.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtY.HideSelection = True
		Me.txtY.ReadOnly = False
		Me.txtY.Maxlength = 0
		Me.txtY.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtY.MultiLine = False
		Me.txtY.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtY.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtY.TabStop = True
		Me.txtY.Visible = True
		Me.txtY.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtY.Name = "txtY"
		Me.txtM.AutoSize = False
		Me.txtM.Size = New System.Drawing.Size(33, 19)
		Me.txtM.Location = New System.Drawing.Point(56, 56)
		Me.txtM.TabIndex = 6
		Me.txtM.AcceptsReturn = True
		Me.txtM.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtM.BackColor = System.Drawing.SystemColors.Window
		Me.txtM.CausesValidation = True
		Me.txtM.Enabled = True
		Me.txtM.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtM.HideSelection = True
		Me.txtM.ReadOnly = False
		Me.txtM.Maxlength = 0
		Me.txtM.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtM.MultiLine = False
		Me.txtM.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtM.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtM.TabStop = True
		Me.txtM.Visible = True
		Me.txtM.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtM.Name = "txtM"
		Me.txtD.AutoSize = False
		Me.txtD.Size = New System.Drawing.Size(33, 19)
		Me.txtD.Location = New System.Drawing.Point(8, 56)
		Me.txtD.TabIndex = 5
		Me.txtD.AcceptsReturn = True
		Me.txtD.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtD.BackColor = System.Drawing.SystemColors.Window
		Me.txtD.CausesValidation = True
		Me.txtD.Enabled = True
		Me.txtD.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtD.HideSelection = True
		Me.txtD.ReadOnly = False
		Me.txtD.Maxlength = 0
		Me.txtD.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtD.MultiLine = False
		Me.txtD.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtD.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtD.TabStop = True
		Me.txtD.Visible = True
		Me.txtD.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtD.Name = "txtD"
		Me.Label5.Text = "�����"
		Me.Label5.Size = New System.Drawing.Size(33, 17)
		Me.Label5.Location = New System.Drawing.Point(184, 32)
		Me.Label5.TabIndex = 4
		Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label5.BackColor = System.Drawing.SystemColors.Control
		Me.Label5.Enabled = True
		Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label5.UseMnemonic = True
		Me.Label5.Visible = True
		Me.Label5.AutoSize = False
		Me.Label5.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label5.Name = "Label5"
		Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.Label4.Text = "���"
		Me.Label4.Size = New System.Drawing.Size(33, 17)
		Me.Label4.Location = New System.Drawing.Point(104, 32)
		Me.Label4.TabIndex = 3
		Me.Label4.BackColor = System.Drawing.SystemColors.Control
		Me.Label4.Enabled = True
		Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label4.UseMnemonic = True
		Me.Label4.Visible = True
		Me.Label4.AutoSize = False
		Me.Label4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label4.Name = "Label4"
		Me.Label3.Text = "�����"
		Me.Label3.Size = New System.Drawing.Size(33, 17)
		Me.Label3.Location = New System.Drawing.Point(56, 32)
		Me.Label3.TabIndex = 2
		Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label3.BackColor = System.Drawing.SystemColors.Control
		Me.Label3.Enabled = True
		Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label3.UseMnemonic = True
		Me.Label3.Visible = True
		Me.Label3.AutoSize = False
		Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label3.Name = "Label3"
		Me.Label2.Text = "�����"
		Me.Label2.Size = New System.Drawing.Size(33, 17)
		Me.Label2.Location = New System.Drawing.Point(8, 32)
		Me.Label2.TabIndex = 1
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
		Me.Label1.Text = "������� ���� � �����"
		Me.Label1.Size = New System.Drawing.Size(217, 17)
		Me.Label1.Location = New System.Drawing.Point(0, 0)
		Me.Label1.TabIndex = 0
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
		Me.Controls.Add(Command4)
		Me.Controls.Add(Command3)
		Me.Controls.Add(Command2)
		Me.Controls.Add(Command1)
		Me.Controls.Add(txtKar)
		Me.Controls.Add(txtY)
		Me.Controls.Add(txtM)
		Me.Controls.Add(txtD)
		Me.Controls.Add(Label5)
		Me.Controls.Add(Label4)
		Me.Controls.Add(Label3)
		Me.Controls.Add(Label2)
		Me.Controls.Add(Label1)
	End Sub
#End Region 
#Region "��������� ���������� "
	Private Shared m_vb6FormDefInstance As frmMain
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As frmMain
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New frmMain()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region 
	Dim dtInp, kdy, ttl, ms, dataInp, dataTek, dataR, msg, ymx, dtIn, msm As Object
	Dim dm As String
	Dim datTek, datInp As Object
	Dim datR As Date
	Dim m, kkr, kr, ky, dx, mx, ty, yxi, yx, yt, zy, mt, dt, kd, kar, d, y As Object
	Dim SavData As Short
	Dim conf As Object
	
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		Dim kyO4 As Object
		Dim kyO3 As Object
		Dim kyO2 As Object
		Dim kyO1 As Object
		Dim ky4 As Object
		Dim ky3 As Object
		Dim ky2 As Object
		Dim ky1 As Object
		Dim ku As Object
		Dim kyf As Object
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� datTek. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		datTek = Now
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kar. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		kar = txtKar.Text
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� d. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		d = txtD.Text
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		m = txtM.Text
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� y. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		y = txtY.Text
		
		'�������� ������� �� ����
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� y. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� d. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If d = 0 Or m = 0 Or y = 0 Then
			Call ShowError1()
			Exit Sub
		End If
		
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� y. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If y <> "" Then
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� y. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			'UPGRADE_WARNING: "Mod" ����� ����� ���������. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1041"'
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� yxi. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			yxi = y Mod 4
		Else
			Call ShowError1()
			Exit Sub
		End If
		
		'�������� ������������ ����
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� d. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (m = 4 Or m = 6 Or m = 9 Or m = 11) And (d >= 31) Then
			Call ShowError3()
			Exit Sub
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� d. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ElseIf (m = 1 Or m = 3 Or m = 5 Or m = 7 Or m = 8 Or m = 10 Or m = 12) And (d >= 32) Then 
			Call ShowError2()
			Exit Sub
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� d. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� yxi. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ElseIf m = 2 And yxi <> 0 And (d >= 29) Then 
			Call ShowError5()
			Exit Sub
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� d. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� yxi. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ElseIf m = 2 And yxi = 0 And (d >= 30) Then 
			Call ShowError4()
			Exit Sub
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� y. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� d. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ElseIf m <= 0 Or m >= 13 Or d <= 0 Or y = 0 Then 
			Call ShowError1()
			Exit Sub
		End If
		
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� y. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� d. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� dataInp. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		dataInp = d & "/" & m & "/" & y
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� dataInp. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� datInp. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		datInp = CDate(dataInp)
		
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If m = 1 Then
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� msm. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			msm = "������"
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ElseIf m = 2 Then 
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� msm. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			msm = "�������"
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ElseIf m = 3 Then 
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� msm. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			msm = "����"
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ElseIf m = 4 Then 
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� msm. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			msm = "������"
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ElseIf m = 5 Then 
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� msm. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			msm = "���"
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ElseIf m = 6 Then 
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� msm. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			msm = "����"
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ElseIf m = 7 Then 
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� msm. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			msm = "����"
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ElseIf m = 8 Then 
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� msm. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			msm = "������"
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ElseIf m = 9 Then 
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� msm. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			msm = "��������"
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ElseIf m = 10 Then 
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� msm. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			msm = "�������"
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ElseIf m = 11 Then 
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� msm. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			msm = "������"
		Else
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� msm. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			msm = "�������"
		End If
		
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� datInp. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� dtIn. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		dtIn = WeekDay(datInp, FirstDayOfWeek.Monday)
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� d. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		dm = d & "/" & m
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� datInp. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� yx. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		yx = Year(datInp)
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� yx. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: "Mod" ����� ����� ���������. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1041"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ky. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ky = yx Mod 16
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� yx. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: "Mod" ����� ����� ���������. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1041"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kyf. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		kyf = yx Mod 4
		
		'�������� ������������ �������� ������ �������
		'� ���������� ������� ������������
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kar. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If kar = 1 Then
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kkr. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			kkr = 0
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kar. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ElseIf kar = 2 Then 
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kkr. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			kkr = -1
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kar. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ElseIf kar = 3 Then 
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kkr. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			kkr = 1
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kar. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ElseIf kar = 4 Then 
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kkr. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			kkr = 2
		Else : Call ShowErrorKar()
			Exit Sub
		End If
		
		'�������� ������� ���� � ���
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� yx. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ymx. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ymx = CStr(yx)
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ymx. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� dataR. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		dataR = "1" & "/" & "1" & "/" + ymx
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� dataR. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		datR = CDate(dataR)
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� datInp. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kdy. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		kdy = DateDiff(Microsoft.VisualBasic.DateInterval.Day, datR, datInp)
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kdy. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: "Mod" ����� ����� ���������. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1041"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kd. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		kd = kdy Mod 4
		
		'��������� ku �� 29 �������
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ky. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (dm = "29/02" Or dm = "29/2") And ky = 4 Then ku = 0
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ky. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (dm = "29/02" Or dm = "29/2") And ky = 8 Then ku = 1
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ky. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (dm = "29/02" Or dm = "29/2") And ky = 12 Then ku = 2
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ky. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (dm = "29/02" Or dm = "29/2") And ky = 0 Then ku = 3
		
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ky. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ky1. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ky1 = (ky = 1 Or ky = 8 Or ky = 11 Or ky = 14)
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ky. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ky2. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ky2 = (ky = 2 Or ky = 5 Or ky = 12 Or ky = 15)
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ky. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ky3. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ky3 = (ky = 0 Or ky = 3 Or ky = 6 Or ky = 9)
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ky. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ky4. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ky4 = (ky = 4 Or ky = 7 Or ky = 10 Or ky = 13)
		'��������� ku �� ��� ������ � ������� ����� 29 �������
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kd. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (dm <> "29/02" Or dm <> "29/2") And (m = 1 Or m = 2) And ky1 And kd = 3 Then ku = 0
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kd. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (dm <> "29/02" Or dm <> "29/2") And (m = 1 Or m = 2) And ky1 And kd = 0 Then ku = 1
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kd. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (dm <> "29/02" Or dm <> "29/2") And (m = 1 Or m = 2) And ky1 And kd = 1 Then ku = 2
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kd. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (dm <> "29/02" Or dm <> "29/2") And (m = 1 Or m = 2) And ky1 And kd = 2 Then ku = 3
		
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kd. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (dm <> "29/02" Or dm <> "29/2") And (m = 1 Or m = 2) And ky2 And kd = 2 Then ku = 0
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kd. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (dm <> "29/02" Or dm <> "29/2") And (m = 1 Or m = 2) And ky2 And kd = 3 Then ku = 1
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kd. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (dm <> "29/02" Or dm <> "29/2") And (m = 1 Or m = 2) And ky2 And kd = 0 Then ku = 2
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kd. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (dm <> "29/02" Or dm <> "29/2") And (m = 1 Or m = 2) And ky2 And kd = 1 Then ku = 3
		
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kd. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (dm <> "29/02" Or dm <> "29/2") And (m = 1 Or m = 2) And ky3 And kd = 1 Then ku = 0
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kd. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (dm <> "29/02" Or dm <> "29/2") And (m = 1 Or m = 2) And ky3 And kd = 2 Then ku = 1
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kd. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (dm <> "29/02" Or dm <> "29/2") And (m = 1 Or m = 2) And ky3 And kd = 3 Then ku = 2
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kd. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (dm <> "29/02" Or dm <> "29/2") And (m = 1 Or m = 2) And ky3 And kd = 0 Then ku = 3
		
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kd. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (dm <> "29/02" Or dm <> "29/2") And (m = 1 Or m = 2) And ky4 And kd = 0 Then ku = 0
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kd. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (dm <> "29/02" Or dm <> "29/2") And (m = 1 Or m = 2) And ky4 And kd = 1 Then ku = 1
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kd. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (dm <> "29/02" Or dm <> "29/2") And (m = 1 Or m = 2) And ky4 And kd = 2 Then ku = 2
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kd. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (dm <> "29/02" Or dm <> "29/2") And (m = 1 Or m = 2) And ky4 And kd = 3 Then ku = 3
		
		'��������� ku �� ��� ������� ����� ������ � ������� �� ������� ���
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ky. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kyO1. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		kyO1 = (ky = 7 Or ky = 10 Or ky = 13)
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ky. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kyO2. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		kyO2 = (ky = 1 Or ky = 11 Or ky = 14)
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ky. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kyO3. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		kyO3 = (ky = 2 Or ky = 5 Or ky = 15)
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ky. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kyO4. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		kyO4 = (ky = 3 Or ky = 6 Or ky = 9)
		
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kd. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (m >= 3) And kyO1 And kd = 0 Then ku = 0
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kd. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (m >= 3) And kyO1 And kd = 1 Then ku = 1
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kd. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (m >= 3) And kyO1 And kd = 2 Then ku = 2
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kd. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (m >= 3) And kyO1 And kd = 3 Then ku = 3
		
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kd. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (m >= 3) And kyO2 And kd = 3 Then ku = 0
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kd. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (m >= 3) And kyO2 And kd = 0 Then ku = 1
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kd. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (m >= 3) And kyO2 And kd = 1 Then ku = 2
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kd. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (m >= 3) And kyO2 And kd = 2 Then ku = 3
		
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kd. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (m >= 3) And kyO3 And kd = 2 Then ku = 0
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kd. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (m >= 3) And kyO3 And kd = 3 Then ku = 1
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kd. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (m >= 3) And kyO3 And kd = 0 Then ku = 2
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kd. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (m >= 3) And kyO3 And kd = 1 Then ku = 3
		
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kd. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (m >= 3) And kyO4 And kd = 1 Then ku = 0
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kd. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (m >= 3) And kyO4 And kd = 2 Then ku = 1
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kd. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (m >= 3) And kyO4 And kd = 3 Then ku = 2
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kd. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (m >= 3) And kyO4 And kd = 0 Then ku = 3
		
		'��������� ku �� ��� ������� ����� ������ � ������� �� ���������� ��� ������� ky=4
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kd. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ky. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (m >= 3) And (ky = 4) And kd = 0 Then ku = 0
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kd. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ky. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (m >= 3) And (ky = 4) And kd = 1 Then ku = 1
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kd. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ky. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (m >= 3) And (ky = 4) And kd = 2 Then ku = 2
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kd. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ky. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (m >= 3) And (ky = 4) And kd = 3 Then ku = 3
		
		'��������� ku �� ��� ������� ����� ������ � ������� �� ���������� ��� ������� ky=8
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kd. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ky. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (m >= 3) And (ky = 8) And kd = 3 Then ku = 0
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kd. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ky. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (m >= 3) And (ky = 8) And kd = 0 Then ku = 1
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kd. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ky. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (m >= 3) And (ky = 8) And kd = 1 Then ku = 2
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kd. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ky. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (m >= 3) And (ky = 8) And kd = 2 Then ku = 3
		
		'��������� ku �� ��� ������� ����� ������ � ������� �� ���������� ��� ������� ky=12
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kd. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ky. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (m >= 3) And (ky = 12) And kd = 2 Then ku = 0
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kd. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ky. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (m >= 3) And (ky = 12) And kd = 3 Then ku = 1
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kd. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ky. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (m >= 3) And (ky = 12) And kd = 0 Then ku = 2
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kd. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ky. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (m >= 3) And (ky = 12) And kd = 1 Then ku = 3
		
		'��������� ku �� ��� ������� ����� ������ � ������� �� ���������� ��� ������� ky=0
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kd. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ky. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (m >= 3) And (ky = 0) And kd = 1 Then ku = 0
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kd. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ky. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (m >= 3) And (ky = 0) And kd = 2 Then ku = 1
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kd. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ky. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (m >= 3) And (ky = 0) And kd = 3 Then ku = 2
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kd. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ky. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� m. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If (m >= 3) And (ky = 0) And kd = 0 Then ku = 3
		
		'������ ������������ ����
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kkr. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ku. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kr. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		kr = ku + kkr
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kr. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If kr = -2 Or kr = 2 Then
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ms. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			ms = " ������ �����"
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kr. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ElseIf kr = 1 Or kr = 5 Then 
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ms. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			ms = " ������� �����"
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kr. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ElseIf kr = 0 Or kr = 4 Then 
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ms. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			ms = " ��������� ���"
		Else
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ms. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			ms = " ��������� ���"
		End If
		
		'���� ����������� ��� ������
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� dtIn. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If dtIn = "1" Then
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� dtInp. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			dtInp = "�����������"
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� dtIn. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ElseIf dtIn = "2" Then 
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� dtInp. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			dtInp = "�������"
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� dtIn. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ElseIf dtIn = "3" Then 
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� dtInp. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			dtInp = "�����"
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� dtIn. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ElseIf dtIn = "4" Then 
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� dtInp. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			dtInp = "�������"
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� dtIn. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ElseIf dtIn = "5" Then 
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� dtInp. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			dtInp = "�������"
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� dtIn. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ElseIf dtIn = "6" Then 
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� dtInp. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			dtInp = "�������"
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� dtIn. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ElseIf dtIn = "7" Then 
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� dtInp. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			dtInp = "�����������"
		End If
		
		'����� ��������� � �����
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� datInp. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� dataInp. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		dataInp = VB6.Format(datInp, "dd mmmm yyyy")
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ms. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� kar. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� dtInp. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� dataInp. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ms = "����� " & dataInp & " (" & dtInp & ")" & vbCrLf & " ��� " & kar & " �������" & " �������������" & vbCrLf & ms
		
		'ms = msg
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ttl. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ttl = "����������� �����"
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ttl. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ms. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ms = MsgBox(ms, MsgBoxStyle.Information, ttl)
		
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� conf. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		conf = GetConfig("FormSet", "DatVal", IniFile)
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� conf. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If conf = 1 Then
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� conf. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			conf = SaveConfig("FormSet", "Day", (txtD.Text), IniFile)
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� conf. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			conf = SaveConfig("FormSet", "Month", (txtM.Text), IniFile)
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� conf. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			conf = SaveConfig("FormSet", "Year", (txtY.Text), IniFile)
		End If
		
	End Sub
	Private Sub ShowError1()
		'������ �������� ����
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ms. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ms = "������� ���������� ����"
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ttl. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ttl = "������"
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ttl. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ms. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ms = MsgBox(ms, MsgBoxStyle.Critical, ttl)
	End Sub
	
	Private Sub ShowError2()
		'������ �������� ����
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� msm. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ms. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ms = "������-�� " & msm & " ������������� 31-�� �����!"
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ttl. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ttl = "������"
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ttl. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ms. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ms = MsgBox(ms, MsgBoxStyle.Critical, ttl)
	End Sub
	
	Private Sub ShowError3()
		'������ �������� ����
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� msm. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ms. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ms = "������-�� " & msm & " ������������� 30-�� �����!"
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ttl. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ttl = "������"
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ttl. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ms. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ms = MsgBox(ms, MsgBoxStyle.Critical, ttl)
	End Sub
	
	Private Sub ShowError4()
		'������ �������� ����
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ms. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ms = "������ ��� ���������� � ������� ������������� 29-�� �����"
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ttl. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ttl = "������"
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ttl. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ms. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ms = MsgBox(ms, MsgBoxStyle.Critical, ttl)
	End Sub
	
	Private Sub ShowError5()
		'������ �������� ����
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ms. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ms = "������ ��� ������� � ������� ������������� 28-�� �����"
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ttl. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ttl = "������"
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ttl. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ms. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ms = MsgBox(ms, MsgBoxStyle.Critical, ttl)
	End Sub
	
	Private Sub ShowErrorKar()
		'������ ������������ �������
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ms. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ms = "������� ����� ������� � ��������� �� 1 �� 4"
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ttl. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ttl = "������"
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ttl. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� ms. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		ms = MsgBox(ms, MsgBoxStyle.Critical, ttl)
	End Sub
	
	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
		frmAbout.DefInstance.Show()
	End Sub
	
	Private Sub Command3_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command3.Click
		Me.Close()
	End Sub
	
	Private Sub Command4_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command4.Click
		frmSet.DefInstance.Show()
	End Sub
	
	Private Sub frmMain_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		txtKar.Text = GetConfig("FormSet", "Grup", IniFile)
		txtD.Text = GetConfig("FormSet", "Day", IniFile)
		txtM.Text = GetConfig("FormSet", "Month", IniFile)
		txtY.Text = GetConfig("FormSet", "Year", IniFile)
        conf = GetConfig("StartUp", "Top", IniFile)
        Me.Top = CDbl(conf)
        Me.Left = GetConfig("StartUp", "Left", IniFile)
	End Sub
	
	'UPGRADE_WARNING: Form ������� frmMain.Unload ����� ����� ���������. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2065"'
	Private Sub frmMain_Closed(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Closed
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� conf. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		conf = SaveConfig("StartUp", "Top", CStr(VB6.PixelsToTwipsY(Me.Top)), IniFile)
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� conf. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		conf = SaveConfig("StartUp", "Left", CStr(VB6.PixelsToTwipsX(Me.Left)), IniFile)
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� conf. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		conf = GetConfig("FormSet", "DatVal", IniFile)
		'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� conf. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		If conf = 1 Then
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� conf. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			conf = SaveConfig("FormSet", "Day", (txtD.Text), IniFile)
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� conf. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			conf = SaveConfig("FormSet", "Month", (txtM.Text), IniFile)
			'UPGRADE_WARNING: �� ������� ��������� �������� �� ��������� ������� conf. �������������: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			conf = SaveConfig("FormSet", "Year", (txtY.Text), IniFile)
		End If
	End Sub
	
	Private Sub txtD_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtD.Enter
		txtD.SelectionLength = 2
	End Sub
	
	Private Sub txtKar_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtKar.Enter
		txtKar.SelectionLength = 1
	End Sub
	
	Private Sub txtM_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtM.Enter
		txtM.SelectionLength = 2
	End Sub
	
	Private Sub txtY_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtY.Enter
		txtY.SelectionLength = 4
	End Sub
End Class