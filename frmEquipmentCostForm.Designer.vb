<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmEquipmentCostForm
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
	Public WithEvents scrYear As System.Windows.Forms.HScrollBar
	Public WithEvents grdEquipmentCost As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
	Public WithEvents comEquipmentCostHelp As System.Windows.Forms.Button
    Public WithEvents _labYear_1 As System.Windows.Forms.Label
    Public WithEvents _labYear_0 As System.Windows.Forms.Label
    Public WithEvents labEquipmentCostHelp As System.Windows.Forms.Label
    Public WithEvents labProjectName As System.Windows.Forms.Label
    Public WithEvents labBackToMenu As System.Windows.Forms.Label
    Public WithEvents imgBackToMenu As System.Windows.Forms.PictureBox
    Public WithEvents labEquipmentCostHeading As System.Windows.Forms.Label
    Public WithEvents labYear As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmEquipmentCostForm))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.PrintForm1 = New Microsoft.VisualBasic.PowerPacks.Printing.PrintForm(Me.components)
        Me.ShapeContainer1 = New Microsoft.VisualBasic.PowerPacks.ShapeContainer
        Me.LineNotQuiteRight = New Microsoft.VisualBasic.PowerPacks.LineShape
        Me.LineNotQuiteLeft = New Microsoft.VisualBasic.PowerPacks.LineShape
        Me.LineBottom = New Microsoft.VisualBasic.PowerPacks.LineShape
        Me.LineBottomRight = New Microsoft.VisualBasic.PowerPacks.LineShape
        Me.LineBottomLeft = New Microsoft.VisualBasic.PowerPacks.LineShape
        Me.LineTop = New Microsoft.VisualBasic.PowerPacks.LineShape
        Me.LineRight = New Microsoft.VisualBasic.PowerPacks.LineShape
        Me.LineLeft = New Microsoft.VisualBasic.PowerPacks.LineShape
        Me.scrYear = New System.Windows.Forms.HScrollBar
        Me.grdEquipmentCost = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
        Me.comEquipmentCostHelp = New System.Windows.Forms.Button
        Me._labYear_1 = New System.Windows.Forms.Label
        Me._labYear_0 = New System.Windows.Forms.Label
        Me.labEquipmentCostHelp = New System.Windows.Forms.Label
        Me.labProjectName = New System.Windows.Forms.Label
        Me.labBackToMenu = New System.Windows.Forms.Label
        Me.imgBackToMenu = New System.Windows.Forms.PictureBox
        Me.labEquipmentCostHeading = New System.Windows.Forms.Label
        Me.labYear = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
        CType(Me.grdEquipmentCost, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.imgBackToMenu, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.labYear, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PrintForm1
        '
        Me.PrintForm1.DocumentName = "document"
        Me.PrintForm1.Form = Me
        Me.PrintForm1.PrintAction = System.Drawing.Printing.PrintAction.PrintToPrinter
        Me.PrintForm1.PrinterSettings = CType(resources.GetObject("PrintForm1.PrinterSettings"), System.Drawing.Printing.PrinterSettings)
        Me.PrintForm1.PrintFileName = Nothing
        '
        'ShapeContainer1
        '
        Me.ShapeContainer1.Location = New System.Drawing.Point(0, 0)
        Me.ShapeContainer1.Margin = New System.Windows.Forms.Padding(0)
        Me.ShapeContainer1.Name = "ShapeContainer1"
        Me.ShapeContainer1.Shapes.AddRange(New Microsoft.VisualBasic.PowerPacks.Shape() {Me.LineNotQuiteRight, Me.LineNotQuiteLeft, Me.LineBottom, Me.LineBottomRight, Me.LineBottomLeft, Me.LineTop, Me.LineRight, Me.LineLeft})
        Me.ShapeContainer1.Size = New System.Drawing.Size(610, 426)
        Me.ShapeContainer1.TabIndex = 10
        Me.ShapeContainer1.TabStop = False
        '
        'LineNotQuiteRight
        '
        Me.LineNotQuiteRight.BorderColor = System.Drawing.Color.Cyan
        Me.LineNotQuiteRight.Name = "LineNotQuiteRight"
        Me.LineNotQuiteRight.X1 = 516
        Me.LineNotQuiteRight.X2 = 516
        Me.LineNotQuiteRight.Y1 = 384
        Me.LineNotQuiteRight.Y2 = 420
        '
        'LineNotQuiteLeft
        '
        Me.LineNotQuiteLeft.BorderColor = System.Drawing.Color.Cyan
        Me.LineNotQuiteLeft.Name = "LineNotQuiteLeft"
        Me.LineNotQuiteLeft.X1 = 92
        Me.LineNotQuiteLeft.X2 = 92
        Me.LineNotQuiteLeft.Y1 = 384
        Me.LineNotQuiteLeft.Y2 = 420
        '
        'LineBottom
        '
        Me.LineBottom.BorderColor = System.Drawing.Color.Cyan
        Me.LineBottom.Name = "LineBottom"
        Me.LineBottom.X1 = 88
        Me.LineBottom.X2 = 520
        Me.LineBottom.Y1 = 416
        Me.LineBottom.Y2 = 416
        '
        'LineBottomRight
        '
        Me.LineBottomRight.BorderColor = System.Drawing.Color.Cyan
        Me.LineBottomRight.Name = "LineBottomRight"
        Me.LineBottomRight.X1 = 512
        Me.LineBottomRight.X2 = 604
        Me.LineBottomRight.Y1 = 388
        Me.LineBottomRight.Y2 = 388
        '
        'LineBottomLeft
        '
        Me.LineBottomLeft.BorderColor = System.Drawing.Color.Cyan
        Me.LineBottomLeft.Name = "LineBottomLeft"
        Me.LineBottomLeft.X1 = 4
        Me.LineBottomLeft.X2 = 96
        Me.LineBottomLeft.Y1 = 388
        Me.LineBottomLeft.Y2 = 388
        '
        'LineTop
        '
        Me.LineTop.BorderColor = System.Drawing.Color.Cyan
        Me.LineTop.Name = "LineTop"
        Me.LineTop.X1 = 4
        Me.LineTop.X2 = 604
        Me.LineTop.Y1 = 36
        Me.LineTop.Y2 = 36
        '
        'LineRight
        '
        Me.LineRight.BorderColor = System.Drawing.Color.Cyan
        Me.LineRight.Name = "LineRight"
        Me.LineRight.X1 = 600
        Me.LineRight.X2 = 600
        Me.LineRight.Y1 = 32
        Me.LineRight.Y2 = 392
        '
        'LineLeft
        '
        Me.LineLeft.BorderColor = System.Drawing.Color.Cyan
        Me.LineLeft.Name = "LineLeft"
        Me.LineLeft.X1 = 8
        Me.LineLeft.X2 = 8
        Me.LineLeft.Y1 = 32
        Me.LineLeft.Y2 = 392
        '
        'scrYear
        '
        Me.scrYear.Cursor = System.Windows.Forms.Cursors.Default
        Me.scrYear.LargeChange = 1
        Me.scrYear.Location = New System.Drawing.Point(100, 396)
        Me.scrYear.Maximum = 32767
        Me.scrYear.Name = "scrYear"
        Me.scrYear.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.scrYear.Size = New System.Drawing.Size(410, 13)
        Me.scrYear.TabIndex = 6
        Me.scrYear.TabStop = True
        '
        'grdEquipmentCost
        '
        Me.grdEquipmentCost.DataSource = Nothing
        Me.grdEquipmentCost.Location = New System.Drawing.Point(32, 44)
        Me.grdEquipmentCost.Name = "grdEquipmentCost"
        Me.grdEquipmentCost.OcxState = CType(resources.GetObject("grdEquipmentCost.OcxState"), System.Windows.Forms.AxHost.State)
        Me.grdEquipmentCost.Size = New System.Drawing.Size(549, 329)
        Me.grdEquipmentCost.TabIndex = 5
        '
        'comEquipmentCostHelp
        '
        Me.comEquipmentCostHelp.BackColor = System.Drawing.SystemColors.Control
        Me.comEquipmentCostHelp.Cursor = System.Windows.Forms.Cursors.Default
        Me.comEquipmentCostHelp.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.comEquipmentCostHelp.ForeColor = System.Drawing.SystemColors.ControlText
        Me.comEquipmentCostHelp.Location = New System.Drawing.Point(592, 408)
        Me.comEquipmentCostHelp.Name = "comEquipmentCostHelp"
        Me.comEquipmentCostHelp.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.comEquipmentCostHelp.Size = New System.Drawing.Size(13, 13)
        Me.comEquipmentCostHelp.TabIndex = 2
        Me.comEquipmentCostHelp.TabStop = False
        Me.comEquipmentCostHelp.Text = "P"
        Me.comEquipmentCostHelp.UseVisualStyleBackColor = False
        '
        '_labYear_1
        '
        Me._labYear_1.BackColor = System.Drawing.Color.Black
        Me._labYear_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._labYear_1.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._labYear_1.ForeColor = System.Drawing.Color.Green
        Me.labYear.SetIndex(Me._labYear_1, CType(1, Short))
        Me._labYear_1.Location = New System.Drawing.Point(348, 376)
        Me._labYear_1.Name = "_labYear_1"
        Me._labYear_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._labYear_1.Size = New System.Drawing.Size(17, 17)
        Me._labYear_1.TabIndex = 8
        Me._labYear_1.Text = "26"
        Me._labYear_1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        '_labYear_0
        '
        Me._labYear_0.BackColor = System.Drawing.Color.Black
        Me._labYear_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._labYear_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._labYear_0.ForeColor = System.Drawing.Color.Yellow
        Me.labYear.SetIndex(Me._labYear_0, CType(0, Short))
        Me._labYear_0.Location = New System.Drawing.Point(244, 376)
        Me._labYear_0.Name = "_labYear_0"
        Me._labYear_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._labYear_0.Size = New System.Drawing.Size(101, 17)
        Me._labYear_0.TabIndex = 7
        Me._labYear_0.Text = "Production Year ="
        '
        'labEquipmentCostHelp
        '
        Me.labEquipmentCostHelp.BackColor = System.Drawing.Color.Black
        Me.labEquipmentCostHelp.Cursor = System.Windows.Forms.Cursors.Default
        Me.labEquipmentCostHelp.Enabled = False
        Me.labEquipmentCostHelp.Font = New System.Drawing.Font("Arial", 9.75!, CType(((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic) _
                        Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labEquipmentCostHelp.ForeColor = System.Drawing.Color.White
        Me.labEquipmentCostHelp.Location = New System.Drawing.Point(556, 404)
        Me.labEquipmentCostHelp.Name = "labEquipmentCostHelp"
        Me.labEquipmentCostHelp.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.labEquipmentCostHelp.Size = New System.Drawing.Size(33, 19)
        Me.labEquipmentCostHelp.TabIndex = 4
        Me.labEquipmentCostHelp.Text = "Help"
        Me.labEquipmentCostHelp.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'labProjectName
        '
        Me.labProjectName.BackColor = System.Drawing.Color.Black
        Me.labProjectName.Cursor = System.Windows.Forms.Cursors.Default
        Me.labProjectName.Font = New System.Drawing.Font("Arial", 12.0!, CType(((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic) _
                        Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labProjectName.ForeColor = System.Drawing.Color.Cyan
        Me.labProjectName.Location = New System.Drawing.Point(352, 8)
        Me.labProjectName.Name = "labProjectName"
        Me.labProjectName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.labProjectName.Size = New System.Drawing.Size(181, 21)
        Me.labProjectName.TabIndex = 3
        Me.labProjectName.Text = "Project Title"
        Me.labProjectName.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'labBackToMenu
        '
        Me.labBackToMenu.BackColor = System.Drawing.Color.Black
        Me.labBackToMenu.Cursor = System.Windows.Forms.Cursors.Default
        Me.labBackToMenu.Font = New System.Drawing.Font("Arial", 9.75!, CType(((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic) _
                        Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labBackToMenu.ForeColor = System.Drawing.Color.White
        Me.labBackToMenu.Location = New System.Drawing.Point(40, 404)
        Me.labBackToMenu.Name = "labBackToMenu"
        Me.labBackToMenu.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.labBackToMenu.Size = New System.Drawing.Size(45, 17)
        Me.labBackToMenu.TabIndex = 1
        Me.labBackToMenu.Text = "Menu"
        Me.labBackToMenu.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'imgBackToMenu
        '
        Me.imgBackToMenu.Cursor = System.Windows.Forms.Cursors.Default
        Me.imgBackToMenu.Image = CType(resources.GetObject("imgBackToMenu.Image"), System.Drawing.Image)
        Me.imgBackToMenu.Location = New System.Drawing.Point(4, 408)
        Me.imgBackToMenu.Name = "imgBackToMenu"
        Me.imgBackToMenu.Size = New System.Drawing.Size(33, 13)
        Me.imgBackToMenu.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.imgBackToMenu.TabIndex = 9
        Me.imgBackToMenu.TabStop = False
        '
        'labEquipmentCostHeading
        '
        Me.labEquipmentCostHeading.BackColor = System.Drawing.Color.Blue
        Me.labEquipmentCostHeading.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.labEquipmentCostHeading.Cursor = System.Windows.Forms.Cursors.Default
        Me.labEquipmentCostHeading.Font = New System.Drawing.Font("Arial", 15.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labEquipmentCostHeading.ForeColor = System.Drawing.Color.White
        Me.labEquipmentCostHeading.Location = New System.Drawing.Point(12, 4)
        Me.labEquipmentCostHeading.Name = "labEquipmentCostHeading"
        Me.labEquipmentCostHeading.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.labEquipmentCostHeading.Size = New System.Drawing.Size(311, 28)
        Me.labEquipmentCostHeading.TabIndex = 0
        Me.labEquipmentCostHeading.Text = "Equipment Operating Costs"
        Me.labEquipmentCostHeading.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'frmEquipmentCostForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Black
        Me.ClientSize = New System.Drawing.Size(610, 426)
        Me.Controls.Add(Me.scrYear)
        Me.Controls.Add(Me.grdEquipmentCost)
        Me.Controls.Add(Me.comEquipmentCostHelp)
        Me.Controls.Add(Me._labYear_1)
        Me.Controls.Add(Me._labYear_0)
        Me.Controls.Add(Me.labEquipmentCostHelp)
        Me.Controls.Add(Me.labProjectName)
        Me.Controls.Add(Me.labBackToMenu)
        Me.Controls.Add(Me.imgBackToMenu)
        Me.Controls.Add(Me.labEquipmentCostHeading)
        Me.Controls.Add(Me.ShapeContainer1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ForeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Location = New System.Drawing.Point(103, 124)
        Me.Name = "frmEquipmentCostForm"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Equipment Operating Costs"
        CType(Me.grdEquipmentCost, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.imgBackToMenu, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.labYear, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Private WithEvents LineNotQuiteRight As Microsoft.VisualBasic.PowerPacks.LineShape
    Private WithEvents LineNotQuiteLeft As Microsoft.VisualBasic.PowerPacks.LineShape
    Private WithEvents LineBottom As Microsoft.VisualBasic.PowerPacks.LineShape
    Private WithEvents LineBottomRight As Microsoft.VisualBasic.PowerPacks.LineShape
    Private WithEvents LineBottomLeft As Microsoft.VisualBasic.PowerPacks.LineShape
    Private WithEvents LineTop As Microsoft.VisualBasic.PowerPacks.LineShape
    Private WithEvents LineRight As Microsoft.VisualBasic.PowerPacks.LineShape
    Private WithEvents LineLeft As Microsoft.VisualBasic.PowerPacks.LineShape
    Private WithEvents ShapeContainer1 As Microsoft.VisualBasic.PowerPacks.ShapeContainer
#End Region 
End Class