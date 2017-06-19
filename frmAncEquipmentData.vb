Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks
Friend Class frmAncEquipmentData
	Inherits System.Windows.Forms.Form
	Dim TempHigh As Single
	Dim NumberCells As Boolean
	Dim TempWide As Single
	Dim FootConversion As Single
	Dim DensityConversion As Single
	Dim PowderConversion As Single
	Private Sub comAncEquipmentPrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles comAncEquipmentPrint.Click
		PrintForm1.Print(Me, PowerPacks.Printing.PrintForm.PrintOption.CompatibleModeClientAreaOnly)
	End Sub
	'UPGRADE_WARNING: Form event frmAncEquipmentData.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmAncEquipmentData_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated

        'Dim baseunit As String
        'Dim baselength As Short
        Dim i As Short
		
		If IsHelpOn = True Then
			txtAncEquipmentValues(WhichCell).Focus()
			IsHelpOn = False
		Else
			For i = 0 To 5
				If optSegment(i).Checked = True Then WhichSegment = i
			Next i
			
			WhichScreen = EquipmentOne
			
			Call drawthevalues()
			
			If InsertFlag = True Then
				labInsert.Text = "Insert"
			Else
				labInsert.Text = "Typeover"
			End If
			
			WhichCell = 0
			
			txtAncEquipmentValues(12).Focus()
		End If
		
	End Sub
	Private Sub frmAncEquipmentData_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        'Dim x As Short

        Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - (VB6.PixelsToTwipsY(Me.Height) + 350)) / 2)
		Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2)
		
		If VB6.PixelsToTwipsY(Me.Top) < 0 Then Me.Top = 0
		If VB6.PixelsToTwipsX(Me.Left) < 0 Then Me.Left = 0
		
		TempWide = VB6.PixelsToTwipsX(Me.ClientRectangle.Width)
		TempHigh = VB6.PixelsToTwipsY(Me.ClientRectangle.Height)
		
		DoNotChange = True
		
		WhichSegment = 0
		optSegment(WhichSegment).Checked = True
		txtSegmentLabel.Text = SegNamie(WhichSegment)
		
		If PageChange(WhichScreen) = True Then
			Call drawthevalues()
		End If
		
		DoNotChange = False
		Call screenstuff()
	End Sub
	'UPGRADE_WARNING: Event frmAncEquipmentData.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub frmAncEquipmentData_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		TempWide = VB6.PixelsToTwipsX(Me.ClientRectangle.Width)
		TempHigh = VB6.PixelsToTwipsY(Me.ClientRectangle.Height)
		Call screenstuff()
	End Sub
	Private Sub frmAncEquipmentData_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		On Error Resume Next
        'Me.Close()
		Call InputMenuAccess(2)
	End Sub
	Private Sub imgBackToMenu_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles imgBackToMenu.Click
		Me.Close()
		Call InputMenuAccess(2)
	End Sub
	Private Sub labBackToMenu_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles labBackToMenu.Click
		Me.Close()
		Call InputMenuAccess(2)
	End Sub
	Private Sub labAncEquipmentHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles labAncEquipmentHelp.Click
		Dim StartHelp As Short
		Dim SendHelp As Short
		StartHelp = 140
		SendHelp = WhichCell - 12
		If NumberCells = True Then
			SendHelp = WhichCell - 4
		End If
		IsHelpOn = True
		Call frmSurfaceHelp.gethelptext(StartHelp, SendHelp)
		frmSurfaceHelp.Show()
	End Sub
	'UPGRADE_WARNING: Event lstAncEquipmentList.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub lstAncEquipmentList_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstAncEquipmentList.SelectedIndexChanged
		Dim Index As Short = lstAncEquipmentList.GetIndex(eventSender)
		txtAncEquipmentValues(WhichCell).Text = LTrim(RTrim(VB6.GetItemString(lstAncEquipmentList(Index), lstAncEquipmentList(Index).SelectedIndex)))
		CellValues(WhichScreen, WhichCell, WhichSegment).Changed = True
		Call Inputer(WhichCell)
		txtAncNumberValues(WhichCell).Focus()
	End Sub
	'UPGRADE_WARNING: Event optSegment.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optSegment_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optSegment.CheckedChanged
		If eventSender.Checked Then
			Dim Index As Short = optSegment.GetIndex(eventSender)
			Dim x As Short
			WhichSegment = Index
			Call drawthevalues()
			For x = 0 To 5
				labSegment(x).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFFFFFF)
			Next x
			labSegment(WhichSegment).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF)
			txtSegmentLabel.Text = SegNamie(WhichSegment)
			txtAncEquipmentValues(WhichCell).Focus()
		End If
	End Sub
	'UPGRADE_WARNING: Event txtAncEquipmentValues.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtAncEquipmentValues_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAncEquipmentValues.TextChanged
		Dim Index As Short = txtAncEquipmentValues.GetIndex(eventSender)
		If DoNotChange = True Then Exit Sub
	End Sub
	'UPGRADE_WARNING: Event txtAncNumberValues.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtAncNumberValues_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAncNumberValues.TextChanged
		Dim Index As Short = txtAncNumberValues.GetIndex(eventSender)
		If DoNotChange = True Then Exit Sub
	End Sub
	Private Sub txtAncEquipmentValues_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAncEquipmentValues.Enter
		Dim Index As Short = txtAncEquipmentValues.GetIndex(eventSender)
		Dim x As Short
		NumberCells = False
		WhichCell = Index
		For x = 0 To 8
			MachineLabel(x).Visible = False
			lstAncEquipmentList(x).Visible = False
			lstAncEquipmentList(x + 9).Visible = False
		Next x
		Select Case WhichCell
			Case 12 To 15
				MachineLabel(WhichCell - 12).Visible = True
				If UnitType = Metric Then
					lstAncEquipmentList(WhichCell - 3).Visible = True
				Else
					lstAncEquipmentList(WhichCell - 12).Visible = True
				End If
			Case 16
				If bltp = 1 Then
					LabAncEquipmentTitles(16).Text = "Powder Buggies"
					MachineLabel(4).Visible = True
					If UnitType = Metric Then
						lstAncEquipmentList(13).Visible = True
					Else
						lstAncEquipmentList(4).Visible = True
					End If
				Else
					LabAncEquipmentTitles(16).Text = "Bulk Trucks"
					MachineLabel(5).Visible = True
					If UnitType = Metric Then
						lstAncEquipmentList(14).Visible = True
					Else
						lstAncEquipmentList(5).Visible = True
					End If
				End If
			Case 17 To 19
				MachineLabel(WhichCell - 11).Visible = True
				If UnitType = Metric Then
					lstAncEquipmentList(WhichCell - 2).Visible = True
				Else
					lstAncEquipmentList(WhichCell - 11).Visible = True
				End If
		End Select
		System.Windows.Forms.SendKeys.Send("{HOME}+{END}")
		WhichCell = Index
		Call drawthevalues()
	End Sub
	Private Sub txtAncNumberValues_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAncNumberValues.Enter
		Dim Index As Short = txtAncNumberValues.GetIndex(eventSender)
		Dim x As Short
		NumberCells = True
		WhichCell = Index
		For x = 0 To 8
			lstAncEquipmentList(x).Visible = False
			MachineLabel(x).Visible = False
		Next x
		System.Windows.Forms.SendKeys.Send("{HOME}+{END}")
		Call drawthevalues()
	End Sub
	Public Sub screenstuff()
		
		Dim p As Decimal
		Dim q As Decimal
		Dim r As Decimal
		Dim s As Decimal
		Dim t As Decimal
		
		Dim x As Short
		
		Dim y As Decimal
		Dim z As Decimal
		
		Dim h As Short
		Dim w As Short
		
		h = 6420
		w = 9150
		
		
		p = (540 / h) * TempHigh
		q = (360 / w) * TempWide
		r = (120 / h) * TempHigh
		s = (120 / w) * TempWide
		t = (1020 / h) * TempHigh
		
		y = (420 / h) * TempHigh
		z = (300 / w) * TempWide
		
		For x = 0 To 2
			labAncEquipmentHeading(x).Top = VB6.TwipsToPixelsY((TempHigh * (120 / h)) + (x * p))
			labAncEquipmentHeading(x).Left = VB6.TwipsToPixelsX((TempWide * (120 / w)) + (x * q))
			labAncEquipmentHeading(x).Width = VB6.TwipsToPixelsX(TempWide * (2835 / w))
		Next x
		
		labAncEquipmentLabels(0).Top = VB6.TwipsToPixelsY(TempHigh * (2280 / h))
		labAncEquipmentLabels(0).Left = VB6.TwipsToPixelsX(TempWide * (360 / w))
		
		labAncEquipmentLabels(1).Top = VB6.TwipsToPixelsY(TempHigh * (900 / h))
		labAncEquipmentLabels(1).Left = VB6.TwipsToPixelsX(TempWide * (3900 / w))
		
		labAncEquipmentLabels(2).Top = VB6.TwipsToPixelsY(TempHigh * (4920 / h))
		labAncEquipmentLabels(2).Left = VB6.TwipsToPixelsX(TempWide * (4440 / w))
		labAncEquipmentLabels(2).Width = VB6.TwipsToPixelsX(TempWide * (1935 / w))
		
		labAncEquipmentLabels(3).Top = VB6.TwipsToPixelsY(TempHigh * (5040 / h))
		labAncEquipmentLabels(3).Left = VB6.TwipsToPixelsX(TempWide * (6780 / w))
		labAncEquipmentLabels(3).Width = VB6.TwipsToPixelsX(TempWide * (1935 / w))
		
		For x = 4 To 6
			labAncEquipmentLabels(x).Top = VB6.TwipsToPixelsY(TempHigh * (660 / h))
			If x = 4 Then
				labAncEquipmentLabels(x).Left = VB6.TwipsToPixelsX(TempWide * (4560 / w))
				labAncEquipmentLabels(x).Width = VB6.TwipsToPixelsX(TempWide * (1665 / w))
			ElseIf x = 5 Then 
				labAncEquipmentLabels(x).Left = VB6.TwipsToPixelsX(TempWide * (6300 / w))
				labAncEquipmentLabels(x).Width = VB6.TwipsToPixelsX(TempWide * (1815 / w))
			Else
				labAncEquipmentLabels(x).Left = VB6.TwipsToPixelsX(TempWide * (8220 / w))
				labAncEquipmentLabels(x).Width = VB6.TwipsToPixelsX(TempWide * (735 / w))
			End If
		Next x
		
		For x = 0 To 5
			optSegment(x).Top = VB6.TwipsToPixelsY(TempHigh * (5280 / h))
			optSegment(x).Left = VB6.TwipsToPixelsX((TempWide * (4560 / w)) + (x * z))
			optSegment(x).Width = VB6.TwipsToPixelsX(TempWide * (195 / w))
			labSegment(x).Top = VB6.TwipsToPixelsY(TempHigh * (5520 / h))
			labSegment(x).Left = VB6.TwipsToPixelsX((TempWide * (4560 / w)) + (x * z))
			labSegment(x).Width = VB6.TwipsToPixelsX(TempWide * (195 / w))
		Next x
		
		txtSegmentLabel.Top = VB6.TwipsToPixelsY(TempHigh * (5340 / h))
		txtSegmentLabel.Left = VB6.TwipsToPixelsX(TempWide * (6780 / w))
		txtSegmentLabel.Width = VB6.TwipsToPixelsX(TempWide * (1935 / w))
		
		For x = 0 To 7
			LabAncEquipmentTitles(x + 12).Top = VB6.TwipsToPixelsY((TempHigh * (1260 / h)) + (x * y))
			LabAncEquipmentTitles(x + 12).Left = VB6.TwipsToPixelsX(TempWide * (4440 / w))
			LabAncEquipmentTitles(x + 12).Width = VB6.TwipsToPixelsX(TempWide * (1635 / w))
			txtAncEquipmentValues(x + 12).Top = VB6.TwipsToPixelsY((TempHigh * (1230 / h)) + (x * y))
			txtAncEquipmentValues(x + 12).Left = VB6.TwipsToPixelsX(TempWide * (6240 / w))
			txtAncEquipmentValues(x + 12).Width = VB6.TwipsToPixelsX(TempWide * (1935 / w))
			txtAncNumberValues(x + 12).Top = VB6.TwipsToPixelsY((TempHigh * (1230 / h)) + (x * y))
			txtAncNumberValues(x + 12).Left = VB6.TwipsToPixelsX(TempWide * (8340 / w))
			txtAncNumberValues(x + 12).Width = VB6.TwipsToPixelsX(TempWide * (495 / w))
		Next x
		
		For x = 0 To 8
			lstAncEquipmentList(x).Top = VB6.TwipsToPixelsY(TempHigh * (2940 / h) + (x * r))
			lstAncEquipmentList(x).Left = VB6.TwipsToPixelsX(TempWide * (600 / w) + (x * s))
			lstAncEquipmentList(x).Height = VB6.TwipsToPixelsY(TempHigh * (1590 / h))
			lstAncEquipmentList(x).Width = VB6.TwipsToPixelsX(TempWide * (2235 / w))
			lstAncEquipmentList(x + 9).Top = VB6.TwipsToPixelsY(TempHigh * (2940 / h) + (x * r))
			lstAncEquipmentList(x + 9).Left = VB6.TwipsToPixelsX(TempWide * (600 / w) + (x * s))
			lstAncEquipmentList(x + 0).Height = VB6.TwipsToPixelsY(TempHigh * (1590 / h))
			lstAncEquipmentList(x + 9).Width = VB6.TwipsToPixelsX(TempWide * (2235 / w))
			MachineLabel(x).Top = VB6.TwipsToPixelsY(TempHigh * (2700 / h) + (x * r))
			MachineLabel(x).Left = VB6.TwipsToPixelsX(TempWide * (600 / w) + (x * s))
			MachineLabel(x).Width = VB6.TwipsToPixelsX(TempWide * (2235 / w))
		Next x
		
		LineHorizontal(0).X1 = VB6.TwipsToPixelsX(TempWide * (2580 / w))
		LineHorizontal(0).X2 = VB6.TwipsToPixelsX(TempWide * (3900 / w))
		LineHorizontal(0).Y1 = VB6.TwipsToPixelsY(TempHigh * (2340 / h))
		LineHorizontal(0).Y2 = VB6.TwipsToPixelsY(TempHigh * (2340 / h))
		
		LineHorizontal(1).X1 = VB6.TwipsToPixelsX(TempWide * (360 / w))
		LineHorizontal(1).X2 = VB6.TwipsToPixelsX(TempWide * (4020 / w))
		LineHorizontal(1).Y1 = VB6.TwipsToPixelsY(TempHigh * (5640 / h))
		LineHorizontal(1).Y2 = VB6.TwipsToPixelsY(TempHigh * (5640 / h))
		
		LineHorizontal(2).X1 = VB6.TwipsToPixelsX(TempWide * (3900 / w))
		LineHorizontal(2).X2 = VB6.TwipsToPixelsX(TempWide * (9060 / w))
		LineHorizontal(2).Y1 = VB6.TwipsToPixelsY(TempHigh * (540 / h))
		LineHorizontal(2).Y2 = VB6.TwipsToPixelsY(TempHigh * (540 / h))
		
		LineHorizontal(3).X1 = VB6.TwipsToPixelsX(TempWide * (4620 / w))
		LineHorizontal(3).X2 = VB6.TwipsToPixelsX(TempWide * (8940 / w))
		LineHorizontal(3).Y1 = VB6.TwipsToPixelsY(TempHigh * (960 / h))
		LineHorizontal(3).Y2 = VB6.TwipsToPixelsY(TempHigh * (960 / h))
		
		LineHorizontal(4).X1 = VB6.TwipsToPixelsX(TempWide * (4020 / w))
		LineHorizontal(4).X2 = VB6.TwipsToPixelsX(TempWide * (9060 / w))
		LineHorizontal(4).Y1 = VB6.TwipsToPixelsY(TempHigh * (4620 / h))
		LineHorizontal(4).Y2 = VB6.TwipsToPixelsY(TempHigh * (4620 / h))
		
		LineVertical(0).X1 = VB6.TwipsToPixelsX(TempWide * (420 / w))
		LineVertical(0).X2 = VB6.TwipsToPixelsX(TempWide * (420 / w))
		LineVertical(0).Y1 = VB6.TwipsToPixelsY(TempHigh * (2580 / h))
		LineVertical(0).Y2 = VB6.TwipsToPixelsY(TempHigh * (5700 / h))
		
		LineVertical(1).X1 = VB6.TwipsToPixelsX(TempWide * (3960 / w))
		LineVertical(1).X2 = VB6.TwipsToPixelsX(TempWide * (3960 / w))
		LineVertical(1).Y1 = VB6.TwipsToPixelsY(TempHigh * (480 / h))
		LineVertical(1).Y2 = VB6.TwipsToPixelsY(TempHigh * (840 / h))
		
		LineVertical(2).X1 = VB6.TwipsToPixelsX(TempWide * (3960 / w))
		LineVertical(2).X2 = VB6.TwipsToPixelsX(TempWide * (3960 / w))
		LineVertical(2).Y1 = VB6.TwipsToPixelsY(TempHigh * (1200 / h))
		LineVertical(2).Y2 = VB6.TwipsToPixelsY(TempHigh * (5700 / h))
		
		LineVertical(3).X1 = VB6.TwipsToPixelsX(TempWide * (9000 / w))
		LineVertical(3).X2 = VB6.TwipsToPixelsX(TempWide * (9000 / w))
		LineVertical(3).Y1 = VB6.TwipsToPixelsY(TempHigh * (480 / h))
		LineVertical(3).Y2 = VB6.TwipsToPixelsY(TempHigh * (4680 / h))
		
		labBackToMenu.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labBackToMenu.Left = VB6.TwipsToPixelsX(TempWide * (600 / w))
		
		imgBackToMenu.Top = VB6.TwipsToPixelsY(TempHigh * (6150 / h))
		imgBackToMenu.Left = VB6.TwipsToPixelsX(TempWide * (60 / w))
		imgBackToMenu.Width = VB6.TwipsToPixelsX(TempWide * (495 / w))
		
		comAncEquipmentPrint.Top = VB6.TwipsToPixelsY(TempHigh * (6150 / h))
		comAncEquipmentPrint.Left = VB6.TwipsToPixelsX(TempWide * (8880 / w))
		
		labAncEquipmentHelp.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labAncEquipmentHelp.Left = VB6.TwipsToPixelsX(TempWide * (8340 / w))
		
		labInsert.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labInsert.Left = VB6.TwipsToPixelsX(TempWide * (4080 / w))
		labInsert.Width = VB6.TwipsToPixelsX(TempWide * (975 / w))
		
	End Sub
	Public Sub drawthevalues()
		Dim i As Short
		
		DoNotChange = True
		
		Call NumberToMachine()
		
		For i = 12 To 19
			If CellValues(WhichScreen, i, WhichSegment).Changed = True Then
				txtAncEquipmentValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF0000)
			Else
				txtAncEquipmentValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
			End If
			If CellValues(EquipmentTwo, i, WhichSegment).Changed = True Then
				txtAncNumberValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF0000)
			Else
				txtAncNumberValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
			End If
			txtAncEquipmentValues(i).Text = LTrim(RTrim(CellValues(WhichScreen, i, WhichSegment).Word))
			txtAncNumberValues(i).Text = VB6.Format(LTrim(Str(CellValues(EquipmentTwo, i, WhichSegment).Value)), "##0")
		Next i
		
		If bltp = 1 Then
			LabAncEquipmentTitles(16).Text = "Powder Buggies"
		ElseIf bltp = 2 Then 
			LabAncEquipmentTitles(16).Text = "Bulk Trucks"
		End If
		
		DoNotChange = False
		
	End Sub
	Private Sub txtAncEquipmentValues_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtAncEquipmentValues.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Dim Index As Short = txtAncEquipmentValues.GetIndex(eventSender)
		
		If KeyCode = 45 Then
			If InsertFlag = True Then
				InsertFlag = False
				labInsert.Text = "Typeover"
			Else
				InsertFlag = True
				labInsert.Text = "Insert"
			End If
		End If
		
		If InsertFlag = False Then
			Select Case KeyCode
				Case 48 To 57, 190
					If KeyCode = 190 Then
						If InStr(txtAncEquipmentValues(Index).Text, ".") = 0 Then
							System.Windows.Forms.SendKeys.Send("{DELETE}")
						End If
					Else
						System.Windows.Forms.SendKeys.Send("{DELETE}")
					End If
			End Select
		End If
		
	End Sub
	Private Sub txtAncNumberValues_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtAncNumberValues.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Dim Index As Short = txtAncNumberValues.GetIndex(eventSender)
		
		If KeyCode = 45 Then
			If InsertFlag = True Then
				InsertFlag = False
				labInsert.Text = "Typeover"
			Else
				InsertFlag = True
				labInsert.Text = "Insert"
			End If
		End If
		
		If InsertFlag = False Then
			Select Case KeyCode
				Case 48 To 57, 190
					If KeyCode = 190 Then
						If InStr(txtAncNumberValues(Index).Text, ".") = 0 Then
							System.Windows.Forms.SendKeys.Send("{DELETE}")
						End If
					Else
						System.Windows.Forms.SendKeys.Send("{DELETE}")
					End If
			End Select
		End If
		
	End Sub
	Private Sub txtAncEquipmentValues_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtAncEquipmentValues.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = txtAncEquipmentValues.GetIndex(eventSender)
		Dim SetSegment As Short
		If KeyAscii > Asc("9") And KeyAscii <> Asc(",") And KeyAscii <> Asc(".") And KeyAscii <> Asc("$") Then
			Beep()
			KeyAscii = 0
		Else
			For SetSegment = 0 To MaxSegment
				CellValues(WhichScreen, WhichCell, SetSegment).Changed = True
			Next SetSegment
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	Private Sub txtAncNumberValues_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtAncNumberValues.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = txtAncNumberValues.GetIndex(eventSender)
		If KeyAscii > Asc("9") And KeyAscii <> Asc(",") And KeyAscii <> Asc(".") And KeyAscii <> Asc("$") Then
			Beep()
			KeyAscii = 0
		Else
			CellValues(EquipmentTwo, WhichCell, WhichSegment).Changed = True
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	Private Sub txtAncEquipmentValues_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAncEquipmentValues.Leave
		Dim Index As Short = txtAncEquipmentValues.GetIndex(eventSender)
		If DoNotChange = True Then Exit Sub
		WhichCell = Index
		Call Inputer(WhichCell)
	End Sub
	Private Sub txtAncNumberValues_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAncNumberValues.Leave
		Dim Index As Short = txtAncNumberValues.GetIndex(eventSender)
		If DoNotChange = True Then Exit Sub
		WhichCell = Index
		Call NumberInputer(WhichCell)
	End Sub
	Private Sub NumberInputer(ByRef Sample As Short)
		Dim x As Short
		Dim i As Short
		Dim life As Decimal
		Dim tempvalue As String
		Dim Digit As New VB6.FixedLengthString(1)
		On Error Resume Next
		If DoNotChange = True Then Exit Sub
		PageChange(WhichScreen) = True
		tempvalue = ""
		For i = 1 To Len(txtAncNumberValues(Sample).Text)
			Digit.Value = Mid(txtAncNumberValues(Sample).Text, i, 1)
			Select Case Digit.Value
				Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ".", "-"
					tempvalue = tempvalue & Digit.Value
			End Select
		Next i
		If CellValues(EquipmentTwo, Sample, WhichSegment).Changed = True Then
			CellValues(EquipmentTwo, Sample, WhichSegment).Value = Val(tempvalue)
			CellValues(EquipmentTwo, Sample, WhichSegment).Word = LTrim(RTrim(txtAncEquipmentValues(Sample).Text))
		End If
		Call drawthevalues()
	End Sub
	Private Sub Inputer(ByRef Sample As Short)
		Dim x As Short
		Dim i As Short
		Dim life As Decimal
		Dim tempvalue As String
		Dim SetSegment As Short
		Dim Digit As New VB6.FixedLengthString(1)
		On Error Resume Next
		If DoNotChange = True Then Exit Sub
		PageChange(WhichScreen) = True
		For SetSegment = 0 To MaxSegment
			If CellValues(WhichScreen, Sample, SetSegment).Changed = True Then
				CellValues(WhichScreen, Sample, SetSegment).Word = LTrim(RTrim(txtAncEquipmentValues(Sample).Text))
			End If
		Next SetSegment
		Call MachineToNumber()
		Call drawthevalues()
	End Sub
	'UPGRADE_WARNING: Event txtSegmentLabel.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtSegmentLabel_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSegmentLabel.TextChanged
		If DoNotChange = True Then Exit Sub
		SegNamie(WhichSegment) = txtSegmentLabel.Text
	End Sub
End Class