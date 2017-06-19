<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmEquipmentPurchaseForm
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public PrintForm1 As Microsoft.VisualBasic.PowerPacks.Printing.PrintForm
	Public WithEvents grdEquipmentPurchase As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
	Public WithEvents comEquipmentPurchaseHelp As System.Windows.Forms.Button
	Public WithEvents LineBottom As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents labEquipmentPurchaseHelp As System.Windows.Forms.Label
	Public WithEvents labProjectName As System.Windows.Forms.Label
	Public WithEvents labBackToMenu As System.Windows.Forms.Label
	Public WithEvents imgBackToMenu As System.Windows.Forms.PictureBox
	Public WithEvents LineTop As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents LineRight As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents LineLeft As Microsoft.VisualBasic.PowerPacks.LineShape
	Public WithEvents labEquipmentPurchaseHeading As System.Windows.Forms.Label
	Public WithEvents ShapeContainer1 As Microsoft.VisualBasic.PowerPacks.ShapeContainer
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmEquipmentPurchaseForm))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.PrintForm1 = New Microsoft.VisualBasic.PowerPacks.Printing.PrintForm(Me)
		Me.ShapeContainer1 = New Microsoft.VisualBasic.PowerPacks.ShapeContainer
		Me.grdEquipmentPurchase = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
		Me.comEquipmentPurchaseHelp = New System.Windows.Forms.Button
		Me.LineBottom = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me.labEquipmentPurchaseHelp = New System.Windows.Forms.Label
		Me.labProjectName = New System.Windows.Forms.Label
		Me.labBackToMenu = New System.Windows.Forms.Label
		Me.imgBackToMenu = New System.Windows.Forms.PictureBox
		Me.LineTop = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me.LineRight = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me.LineLeft = New Microsoft.VisualBasic.PowerPacks.LineShape
		Me.labEquipmentPurchaseHeading = New System.Windows.Forms.Label
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.grdEquipmentPurchase, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
		Me.BackColor = System.Drawing.Color.Black
		Me.Text = "Equipment Capital Costs"
		Me.ClientSize = New System.Drawing.Size(610, 426)
		Me.Location = New System.Drawing.Point(103, 124)
		Me.Font = New System.Drawing.Font("Arial", 9!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.ForeColor = System.Drawing.Color.FromARGB(128, 128, 0)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
		Me.ControlBox = True
		Me.Enabled = True
		Me.KeyPreview = False
		Me.MaximizeBox = True
		Me.MinimizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmEquipmentPurchaseForm"
		grdEquipmentPurchase.OcxState = CType(resources.GetObject("grdEquipmentPurchase.OcxState"), System.Windows.Forms.AxHost.State)
		Me.grdEquipmentPurchase.Size = New System.Drawing.Size(585, 357)
		Me.grdEquipmentPurchase.Location = New System.Drawing.Point(12, 40)
		Me.grdEquipmentPurchase.TabIndex = 5
		Me.grdEquipmentPurchase.Name = "grdEquipmentPurchase"
		Me.comEquipmentPurchaseHelp.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.comEquipmentPurchaseHelp.Text = "P"
		Me.comEquipmentPurchaseHelp.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.comEquipmentPurchaseHelp.Size = New System.Drawing.Size(13, 13)
		Me.comEquipmentPurchaseHelp.Location = New System.Drawing.Point(592, 408)
		Me.comEquipmentPurchaseHelp.TabIndex = 2
		Me.comEquipmentPurchaseHelp.TabStop = False
		Me.comEquipmentPurchaseHelp.BackColor = System.Drawing.SystemColors.Control
		Me.comEquipmentPurchaseHelp.CausesValidation = True
		Me.comEquipmentPurchaseHelp.Enabled = True
		Me.comEquipmentPurchaseHelp.ForeColor = System.Drawing.SystemColors.ControlText
		Me.comEquipmentPurchaseHelp.Cursor = System.Windows.Forms.Cursors.Default
		Me.comEquipmentPurchaseHelp.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.comEquipmentPurchaseHelp.Name = "comEquipmentPurchaseHelp"
		Me.LineBottom.BorderColor = System.Drawing.Color.Cyan
		Me.LineBottom.X1 = 4
		Me.LineBottom.X2 = 604
		Me.LineBottom.Y1 = 400
		Me.LineBottom.Y2 = 400
		Me.LineBottom.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me.LineBottom.BorderWidth = 1
		Me.LineBottom.Visible = True
		Me.LineBottom.Name = "LineBottom"
		Me.labEquipmentPurchaseHelp.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.labEquipmentPurchaseHelp.BackColor = System.Drawing.Color.Black
		Me.labEquipmentPurchaseHelp.Text = "Help"
		Me.labEquipmentPurchaseHelp.Enabled = False
		Me.labEquipmentPurchaseHelp.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline Or System.Drawing.FontStyle.Italic Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.labEquipmentPurchaseHelp.ForeColor = System.Drawing.Color.White
		Me.labEquipmentPurchaseHelp.Size = New System.Drawing.Size(33, 19)
		Me.labEquipmentPurchaseHelp.Location = New System.Drawing.Point(556, 404)
		Me.labEquipmentPurchaseHelp.TabIndex = 4
		Me.labEquipmentPurchaseHelp.Cursor = System.Windows.Forms.Cursors.Default
		Me.labEquipmentPurchaseHelp.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.labEquipmentPurchaseHelp.UseMnemonic = True
		Me.labEquipmentPurchaseHelp.Visible = True
		Me.labEquipmentPurchaseHelp.AutoSize = False
		Me.labEquipmentPurchaseHelp.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.labEquipmentPurchaseHelp.Name = "labEquipmentPurchaseHelp"
		Me.labProjectName.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.labProjectName.BackColor = System.Drawing.Color.Black
		Me.labProjectName.Text = "Project Title"
		Me.labProjectName.Font = New System.Drawing.Font("Arial", 12!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline Or System.Drawing.FontStyle.Italic Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.labProjectName.ForeColor = System.Drawing.Color.Red
		Me.labProjectName.Size = New System.Drawing.Size(181, 21)
		Me.labProjectName.Location = New System.Drawing.Point(372, 8)
		Me.labProjectName.TabIndex = 3
		Me.labProjectName.Enabled = True
		Me.labProjectName.Cursor = System.Windows.Forms.Cursors.Default
		Me.labProjectName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.labProjectName.UseMnemonic = True
		Me.labProjectName.Visible = True
		Me.labProjectName.AutoSize = False
		Me.labProjectName.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.labProjectName.Name = "labProjectName"
		Me.labBackToMenu.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.labBackToMenu.BackColor = System.Drawing.Color.Black
		Me.labBackToMenu.Text = "Menu"
		Me.labBackToMenu.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline Or System.Drawing.FontStyle.Italic Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.labBackToMenu.ForeColor = System.Drawing.Color.White
		Me.labBackToMenu.Size = New System.Drawing.Size(45, 17)
		Me.labBackToMenu.Location = New System.Drawing.Point(40, 404)
		Me.labBackToMenu.TabIndex = 1
		Me.labBackToMenu.Enabled = True
		Me.labBackToMenu.Cursor = System.Windows.Forms.Cursors.Default
		Me.labBackToMenu.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.labBackToMenu.UseMnemonic = True
		Me.labBackToMenu.Visible = True
		Me.labBackToMenu.AutoSize = False
		Me.labBackToMenu.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.labBackToMenu.Name = "labBackToMenu"
		Me.imgBackToMenu.Size = New System.Drawing.Size(33, 13)
		Me.imgBackToMenu.Location = New System.Drawing.Point(4, 408)
		Me.imgBackToMenu.Image = CType(resources.GetObject("imgBackToMenu.Image"), System.Drawing.Image)
		Me.imgBackToMenu.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
		Me.imgBackToMenu.Enabled = True
		Me.imgBackToMenu.Cursor = System.Windows.Forms.Cursors.Default
		Me.imgBackToMenu.Visible = True
		Me.imgBackToMenu.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.imgBackToMenu.Name = "imgBackToMenu"
		Me.LineTop.BorderColor = System.Drawing.Color.Cyan
		Me.LineTop.X1 = 4
		Me.LineTop.X2 = 604
		Me.LineTop.Y1 = 36
		Me.LineTop.Y2 = 36
		Me.LineTop.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me.LineTop.BorderWidth = 1
		Me.LineTop.Visible = True
		Me.LineTop.Name = "LineTop"
		Me.LineRight.BorderColor = System.Drawing.Color.Cyan
		Me.LineRight.X1 = 600
		Me.LineRight.X2 = 600
		Me.LineRight.Y1 = 32
		Me.LineRight.Y2 = 404
		Me.LineRight.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me.LineRight.BorderWidth = 1
		Me.LineRight.Visible = True
		Me.LineRight.Name = "LineRight"
		Me.LineLeft.BorderColor = System.Drawing.Color.Cyan
		Me.LineLeft.X1 = 8
		Me.LineLeft.X2 = 8
		Me.LineLeft.Y1 = 32
		Me.LineLeft.Y2 = 404
		Me.LineLeft.BorderStyle = System.Drawing.Drawing2D.DashStyle.Solid
		Me.LineLeft.BorderWidth = 1
		Me.LineLeft.Visible = True
		Me.LineLeft.Name = "LineLeft"
		Me.labEquipmentPurchaseHeading.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.labEquipmentPurchaseHeading.BackColor = System.Drawing.Color.Blue
		Me.labEquipmentPurchaseHeading.Text = "Equipment Capital Costs"
		Me.labEquipmentPurchaseHeading.Font = New System.Drawing.Font("Arial", 15.75!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.labEquipmentPurchaseHeading.ForeColor = System.Drawing.Color.White
		Me.labEquipmentPurchaseHeading.Size = New System.Drawing.Size(307, 28)
		Me.labEquipmentPurchaseHeading.Location = New System.Drawing.Point(12, 4)
		Me.labEquipmentPurchaseHeading.TabIndex = 0
		Me.labEquipmentPurchaseHeading.Enabled = True
		Me.labEquipmentPurchaseHeading.Cursor = System.Windows.Forms.Cursors.Default
		Me.labEquipmentPurchaseHeading.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.labEquipmentPurchaseHeading.UseMnemonic = True
		Me.labEquipmentPurchaseHeading.Visible = True
		Me.labEquipmentPurchaseHeading.AutoSize = False
		Me.labEquipmentPurchaseHeading.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.labEquipmentPurchaseHeading.Name = "labEquipmentPurchaseHeading"
		Me.Controls.Add(grdEquipmentPurchase)
		Me.Controls.Add(comEquipmentPurchaseHelp)
		Me.ShapeContainer1.Shapes.Add(LineBottom)
		Me.Controls.Add(labEquipmentPurchaseHelp)
		Me.Controls.Add(labProjectName)
		Me.Controls.Add(labBackToMenu)
		Me.Controls.Add(imgBackToMenu)
		Me.ShapeContainer1.Shapes.Add(LineTop)
		Me.ShapeContainer1.Shapes.Add(LineRight)
		Me.ShapeContainer1.Shapes.Add(LineLeft)
		Me.Controls.Add(labEquipmentPurchaseHeading)
		Me.Controls.Add(ShapeContainer1)
		CType(Me.grdEquipmentPurchase, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class