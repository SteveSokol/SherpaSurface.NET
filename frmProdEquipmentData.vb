Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks
Friend Class frmProdEquipmentData
	Inherits System.Windows.Forms.Form
	Dim TempHigh As Single
	Dim TempWide As Single
	Dim FootConversion As Single
	Dim DensityConversion As Single
	Dim PowderConversion As Single
	Private Sub comEquipmentPrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles comEquipmentPrint.Click
		PrintForm1.Print(Me, PowerPacks.Printing.PrintForm.PrintOption.CompatibleModeClientAreaOnly)
	End Sub
	'UPGRADE_WARNING: Form event frmProdEquipmentData.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmProdEquipmentData_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated

        'Dim baseunit As String
        'Dim baselength As Short
        Dim i As Short
		
		If IsHelpOn = True Then
			txtEquipmentValues(WhichCell).Focus()
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
			
			txtEquipmentValues(0).Focus()
		End If
		
	End Sub
	Private Sub frmProdEquipmentData_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
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
	'UPGRADE_WARNING: Event frmProdEquipmentData.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub frmProdEquipmentData_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		TempWide = VB6.PixelsToTwipsX(Me.ClientRectangle.Width)
		TempHigh = VB6.PixelsToTwipsY(Me.ClientRectangle.Height)
		Call screenstuff()
	End Sub
	Private Sub frmProdEquipmentData_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
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
	Private Sub labEquipmentHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles labEquipmentHelp.Click
		Dim StartHelp As Short
		Dim SendHelp As Short
		StartHelp = 127
		Select Case WhichCell
			Case 0, 1
				SendHelp = WhichCell
			Case 2, 3
				SendHelp = WhichCell - 2
			Case 10, 30
				SendHelp = 2
			Case 11, 31
				SendHelp = 3
			Case 22, 27
				SendHelp = 4
			Case 24, 29
				SendHelp = 5
		End Select
		IsHelpOn = True
		Call frmSurfaceHelp.gethelptext(StartHelp, SendHelp)
		frmSurfaceHelp.Show()
	End Sub
	'UPGRADE_WARNING: Event lstEquipmentList.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub lstEquipmentList_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstEquipmentList.SelectedIndexChanged
		Dim Index As Short = lstEquipmentList.GetIndex(eventSender)
		Dim x As Short
		Dim i As Short
		i = WhichCell
		txtEquipmentValues(WhichCell).Text = LTrim(RTrim(VB6.GetItemString(lstEquipmentList(Index), lstEquipmentList(Index).SelectedIndex)))
		Select Case i
			Case 0
				For x = 0 To 3
					If CellValues(WhichScreen, x, WhichSegment).Value <> 0 Then
						CellValues(WhichScreen, x, WhichSegment).Changed = True
					End If
				Next x
				If CellValues(WhichScreen, 20, WhichSegment).Value <> 0 Then
					CellValues(WhichScreen, 20, WhichSegment).Changed = True
				End If
			Case 1
				If CellValues(WhichScreen, 4, WhichSegment).Value <> 0 Then
					CellValues(WhichScreen, 4, WhichSegment).Changed = True
				ElseIf CellValues(WhichScreen, 21, WhichSegment).Value <> 0 Then 
					CellValues(WhichScreen, 21, WhichSegment).Changed = True
				End If
			Case 2
				For x = 5 To 8
					If CellValues(WhichScreen, x, WhichSegment).Value <> 0 Then
						CellValues(WhichScreen, x, WhichSegment).Changed = True
					End If
				Next x
				If CellValues(WhichScreen, 25, WhichSegment).Value <> 0 Then
					CellValues(WhichScreen, 25, WhichSegment).Changed = True
				End If
			Case 3
				If CellValues(WhichScreen, 9, WhichSegment).Value <> 0 Then
					CellValues(WhichScreen, 9, WhichSegment).Changed = True
				ElseIf CellValues(WhichScreen, 26, WhichSegment).Value <> 0 Then 
					CellValues(WhichScreen, 26, WhichSegment).Changed = True
				End If
			Case 10, 11
				If CellValues(WhichScreen, i, WhichSegment).Value <> 0 Then
					CellValues(WhichScreen, i, WhichSegment).Changed = True
				End If
			Case 22, 27
				If CellValues(WhichScreen, i, WhichSegment).Value <> 0 Then
					CellValues(WhichScreen, i, WhichSegment).Changed = True
				ElseIf CellValues(WhichScreen, i + 1, WhichSegment).Value <> 0 Then 
					CellValues(WhichScreen, i + 1, WhichSegment).Changed = True
				End If
			Case 24, 29
				If CellValues(WhichScreen, i, WhichSegment).Value <> 0 Then
					CellValues(WhichScreen, i, WhichSegment).Changed = True
				End If
			Case 30, 31
				If CellValues(WhichScreen, i, WhichSegment).Value <> 0 Then
					CellValues(WhichScreen, i, WhichSegment).Changed = True
				End If
		End Select
		Call Inputer(WhichCell)
		txtNumberValues(WhichCell).Focus()
	End Sub
	'UPGRADE_WARNING: Event optSegment.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optSegment_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optSegment.CheckedChanged
		If eventSender.Checked Then
			Dim Index As Short = optSegment.GetIndex(eventSender)
			Dim x As Short
			On Error Resume Next
			WhichSegment = Index
			Call drawthevalues()
			For x = 0 To 5
				labSegment(x).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFFFFFF)
			Next x
			labSegment(WhichSegment).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF)
			txtSegmentLabel.Text = SegNamie(WhichSegment)
			If txtEquipmentValues(WhichCell).Enabled = True Then
				txtEquipmentValues(WhichCell).Focus()
			End If
		End If
	End Sub
	'UPGRADE_WARNING: Event txtEquipmentValues.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtEquipmentValues_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtEquipmentValues.TextChanged
		Dim Index As Short = txtEquipmentValues.GetIndex(eventSender)
		If DoNotChange = True Then Exit Sub
	End Sub
	'UPGRADE_WARNING: Event txtNumberValues.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtNumberValues_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNumberValues.TextChanged
		Dim Index As Short = txtNumberValues.GetIndex(eventSender)
		If DoNotChange = True Then Exit Sub
	End Sub
	Private Sub txtEquipmentValues_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtEquipmentValues.Enter
		Dim Index As Short = txtEquipmentValues.GetIndex(eventSender)
		Dim x As Short
		WhichCell = Index
		For x = 0 To 11
			lstEquipmentList(x).Visible = False
			lstEquipmentList(x + 12).Visible = False
			MachineLabel(x).Visible = False
		Next x
		Select Case WhichCell
			Case 0
				For x = 0 To 3
					If CellValues(WhichScreen, x, WhichSegment).Value <> 0 Then
						If UnitType = Metric Then
							lstEquipmentList(x + 12).Visible = True
						Else
							lstEquipmentList(x).Visible = True
						End If
						MachineLabel(x).Visible = True
					End If
				Next x
				If CellValues(WhichScreen, 20, WhichSegment).Value <> 0 Then
					If UnitType = Metric Then
						lstEquipmentList(16).Visible = True
					Else
						lstEquipmentList(4).Visible = True
					End If
					MachineLabel(4).Visible = True
				End If
			Case 1
				If CellValues(WhichScreen, 4, WhichSegment).Value <> 0 Then
					If UnitType = Metric Then
						lstEquipmentList(17).Visible = True
					Else
						lstEquipmentList(5).Visible = True
					End If
					MachineLabel(5).Visible = True
				ElseIf CellValues(WhichScreen, 21, WhichSegment).Value <> 0 Then 
					If UnitType = Metric Then
						lstEquipmentList(18).Visible = True
					Else
						lstEquipmentList(6).Visible = True
					End If
					MachineLabel(6).Visible = True
				End If
			Case 2
				For x = 5 To 8
					If CellValues(WhichScreen, x, WhichSegment).Value <> 0 Then
						If UnitType = Metric Then
							lstEquipmentList(x + 7).Visible = True
						Else
							lstEquipmentList(x - 5).Visible = True
						End If
						MachineLabel(x - 5).Visible = True
					End If
				Next x
				If CellValues(WhichScreen, 25, WhichSegment).Value <> 0 Then
					If UnitType = Metric Then
						lstEquipmentList(16).Visible = True
					Else
						lstEquipmentList(4).Visible = True
					End If
					MachineLabel(4).Visible = True
				End If
			Case 3
				If CellValues(WhichScreen, 9, WhichSegment).Value <> 0 Then
					If UnitType = Metric Then
						lstEquipmentList(17).Visible = True
					Else
						lstEquipmentList(5).Visible = True
					End If
					MachineLabel(5).Visible = True
				ElseIf CellValues(WhichScreen, 26, WhichSegment).Value <> 0 Then 
					If UnitType = Metric Then
						lstEquipmentList(18).Visible = True
					Else
						lstEquipmentList(6).Visible = True
					End If
					MachineLabel(6).Visible = True
				End If
			Case 22, 27
				If CellValues(WhichScreen, WhichCell, WhichSegment).Value <> 0 Then
					If UnitType = Metric Then
						lstEquipmentList(19).Visible = True
					Else
						lstEquipmentList(7).Visible = True
					End If
					MachineLabel(7).Visible = True
				ElseIf CellValues(WhichScreen, WhichCell + 1, WhichSegment).Value <> 0 Then 
					If UnitType = Metric Then
						lstEquipmentList(20).Visible = True
					Else
						lstEquipmentList(8).Visible = True
					End If
					MachineLabel(8).Visible = True
				End If
			Case 24, 29
				If UnitType = Metric Then
					lstEquipmentList(21).Visible = True
				Else
					lstEquipmentList(9).Visible = True
				End If
				MachineLabel(9).Visible = True
			Case 11, 31
				If UnitType = Metric Then
					lstEquipmentList(22).Visible = True
				Else
					lstEquipmentList(10).Visible = True
				End If
				MachineLabel(10).Visible = True
			Case 10, 30
				If UnitType = Metric Then
					lstEquipmentList(23).Visible = True
				Else
					lstEquipmentList(11).Visible = True
				End If
				MachineLabel(11).Visible = True
		End Select
		System.Windows.Forms.SendKeys.Send("{HOME}+{END}")
		WhichCell = Index
		Call drawthevalues()
	End Sub
	Private Sub txtNumberValues_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNumberValues.Enter
		Dim Index As Short = txtNumberValues.GetIndex(eventSender)
        'Dim x As Short
        WhichCell = Index
		System.Windows.Forms.SendKeys.Send("{HOME}+{END}")
		Call drawthevalues()
	End Sub
	Public Sub screenstuff()

        'Dim p As Decimal
        Dim q As Decimal
		Dim r As Decimal
		Dim s As Decimal
		Dim t As Decimal
		Dim u As Decimal
		Dim v As Decimal
		
		Dim x As Short
		
		Dim y As Decimal
		Dim z As Decimal
		
		Dim h As Short
		Dim w As Short
		
		h = 6420
		w = 9150
		
		q = (60 / h) * TempHigh
		r = (120 / w) * TempWide
		s = (2700 / h) * TempHigh
		t = (900 / h) * TempHigh
		u = (480 / h) * TempHigh
		v = (780 / w) * TempWide
		
		y = (420 / h) * TempHigh
		z = (300 / w) * TempWide
		
		For x = 0 To 2
			labEquipmentHeading(x).Top = VB6.TwipsToPixelsY((TempHigh * (60 / h)) + (x * u))
			labEquipmentHeading(x).Left = VB6.TwipsToPixelsX((TempWide * (60 / w)) + (x * v))
			labEquipmentHeading(x).Width = VB6.TwipsToPixelsX(TempWide * (2175 / w))
		Next x
		
		For x = 0 To 1
			labEquipmentLabels(x).Top = VB6.TwipsToPixelsY((TempHigh * (1560 / h)) + (x * t))
			labEquipmentLabels(x).Left = VB6.TwipsToPixelsX(TempWide * (1020 / w))
			labEquipmentLabels(x).Width = VB6.TwipsToPixelsX(TempWide * (1935 / w))
			labEquipmentLabels(x + 3).Top = VB6.TwipsToPixelsY((TempHigh * (480 / h)) + (x * s))
			labEquipmentLabels(x + 3).Left = VB6.TwipsToPixelsX(TempWide * (4020 / w))
		Next x
		
		labEquipmentLabels(2).Top = VB6.TwipsToPixelsY(TempHigh * (3180 / h))
		labEquipmentLabels(2).Left = VB6.TwipsToPixelsX(TempWide * (160 / w))
		
		For x = 0 To 2
			labEquipmentLabels(x + 5).Top = VB6.TwipsToPixelsY(TempHigh * (240 / h))
		Next x
		
		labEquipmentLabels(5).Left = VB6.TwipsToPixelsX(TempWide * (4560 / w))
		labEquipmentLabels(5).Width = VB6.TwipsToPixelsX(TempWide * (1545 / w))
		labEquipmentLabels(6).Left = VB6.TwipsToPixelsX(TempWide * (6360 / w))
		labEquipmentLabels(6).Width = VB6.TwipsToPixelsX(TempWide * (1815 / w))
		labEquipmentLabels(7).Left = VB6.TwipsToPixelsX(TempWide * (8220 / w))
		labEquipmentLabels(7).Width = VB6.TwipsToPixelsX(TempWide * (735 / w))
		
		For x = 0 To 5
			optSegment(x).Top = VB6.TwipsToPixelsY(TempHigh * (1860 / h))
			optSegment(x).Left = VB6.TwipsToPixelsX((TempWide * (1140 / w)) + (x * z))
			optSegment(x).Width = VB6.TwipsToPixelsX(TempWide * (195 / w))
			labSegment(x).Top = VB6.TwipsToPixelsY(TempHigh * (2100 / h))
			labSegment(x).Left = VB6.TwipsToPixelsX((TempWide * (1140 / w)) + (x * z))
			labSegment(x).Width = VB6.TwipsToPixelsX(TempWide * (195 / w))
		Next x
		
		txtSegmentLabel.Top = VB6.TwipsToPixelsY(TempHigh * (2760 / h))
		txtSegmentLabel.Left = VB6.TwipsToPixelsX(TempWide * (1020 / w))
		txtSegmentLabel.Width = VB6.TwipsToPixelsX(TempWide * (1935 / w))
		
		For x = 0 To 1
			LabEquipmentTitles(x).Top = VB6.TwipsToPixelsY((TempHigh * (780 / h)) + (x * y))
			LabEquipmentTitles(x).Left = VB6.TwipsToPixelsX(TempWide * (4620 / w))
			LabEquipmentTitles(x).Width = VB6.TwipsToPixelsX(TempWide * (1575 / w))
			txtEquipmentValues(x).Top = VB6.TwipsToPixelsY((TempHigh * (750 / h)) + (x * y))
			txtEquipmentValues(x).Left = VB6.TwipsToPixelsX(TempWide * (6360 / w))
			txtEquipmentValues(x).Width = VB6.TwipsToPixelsX(TempWide * (1815 / w))
			txtNumberValues(x).Top = VB6.TwipsToPixelsY((TempHigh * (750 / h)) + (x * y))
			txtNumberValues(x).Left = VB6.TwipsToPixelsX(TempWide * (8340 / w))
			txtNumberValues(x).Width = VB6.TwipsToPixelsX(TempWide * (495 / w))
			
			LabEquipmentTitles(x + 2).Top = VB6.TwipsToPixelsY((TempHigh * (3480 / h)) + (x * y))
			LabEquipmentTitles(x + 2).Left = VB6.TwipsToPixelsX(TempWide * (4620 / w))
			LabEquipmentTitles(x + 2).Width = VB6.TwipsToPixelsX(TempWide * (1575 / w))
			txtEquipmentValues(x + 2).Top = VB6.TwipsToPixelsY((TempHigh * (3450 / h)) + (x * y))
			txtEquipmentValues(x + 2).Left = VB6.TwipsToPixelsX(TempWide * (6360 / w))
			txtEquipmentValues(x + 2).Width = VB6.TwipsToPixelsX(TempWide * (1815 / w))
			txtNumberValues(x + 2).Top = VB6.TwipsToPixelsY((TempHigh * (3450 / h)) + (x * y))
			txtNumberValues(x + 2).Left = VB6.TwipsToPixelsX(TempWide * (8340 / w))
			txtNumberValues(x + 2).Width = VB6.TwipsToPixelsX(TempWide * (495 / w))
			
			LabEquipmentTitles(x + 10).Top = VB6.TwipsToPixelsY((TempHigh * (2880 / h)) - (x * y))
			LabEquipmentTitles(x + 10).Left = VB6.TwipsToPixelsX(TempWide * (4620 / w))
			LabEquipmentTitles(x + 10).Width = VB6.TwipsToPixelsX(TempWide * (1575 / w))
			txtEquipmentValues(x + 10).Top = VB6.TwipsToPixelsY((TempHigh * (2850 / h)) - (x * y))
			txtEquipmentValues(x + 10).Left = VB6.TwipsToPixelsX(TempWide * (6360 / w))
			txtEquipmentValues(x + 10).Width = VB6.TwipsToPixelsX(TempWide * (1815 / w))
			txtNumberValues(x + 10).Top = VB6.TwipsToPixelsY((TempHigh * (2850 / h)) - (x * y))
			txtNumberValues(x + 10).Left = VB6.TwipsToPixelsX(TempWide * (8340 / w))
			txtNumberValues(x + 10).Width = VB6.TwipsToPixelsX(TempWide * (495 / w))
			
			LabEquipmentTitles(x + 30).Top = VB6.TwipsToPixelsY((TempHigh * (5580 / h)) - (x * y))
			LabEquipmentTitles(x + 30).Left = VB6.TwipsToPixelsX(TempWide * (4620 / w))
			LabEquipmentTitles(x + 30).Width = VB6.TwipsToPixelsX(TempWide * (1575 / w))
			txtEquipmentValues(x + 30).Top = VB6.TwipsToPixelsY((TempHigh * (5550 / h)) - (x * y))
			txtEquipmentValues(x + 30).Left = VB6.TwipsToPixelsX(TempWide * (6360 / w))
			txtEquipmentValues(x + 30).Width = VB6.TwipsToPixelsX(TempWide * (1815 / w))
			txtNumberValues(x + 30).Top = VB6.TwipsToPixelsY((TempHigh * (5550 / h)) - (x * y))
			txtNumberValues(x + 30).Left = VB6.TwipsToPixelsX(TempWide * (8340 / w))
			txtNumberValues(x + 30).Width = VB6.TwipsToPixelsX(TempWide * (495 / w))
			
			If x < 1 Then
				LabEquipmentTitles(x + 22).Top = VB6.TwipsToPixelsY((TempHigh * (1620 / h)) + (x * y))
				LabEquipmentTitles(x + 22).Left = VB6.TwipsToPixelsX(TempWide * (4620 / w))
				LabEquipmentTitles(x + 22).Width = VB6.TwipsToPixelsX(TempWide * (1575 / w))
				txtEquipmentValues(x + 22).Top = VB6.TwipsToPixelsY((TempHigh * (1590 / h)) + (x * y))
				txtEquipmentValues(x + 22).Left = VB6.TwipsToPixelsX(TempWide * (6360 / w))
				txtEquipmentValues(x + 22).Width = VB6.TwipsToPixelsX(TempWide * (1815 / w))
				txtNumberValues(x + 22).Top = VB6.TwipsToPixelsY((TempHigh * (1590 / h)) + (x * y))
				txtNumberValues(x + 22).Left = VB6.TwipsToPixelsX(TempWide * (8340 / w))
				txtNumberValues(x + 22).Width = VB6.TwipsToPixelsX(TempWide * (495 / w))
				
				LabEquipmentTitles(x + 24).Top = VB6.TwipsToPixelsY((TempHigh * (2040 / h)) + (x * y))
				LabEquipmentTitles(x + 24).Left = VB6.TwipsToPixelsX(TempWide * (4620 / w))
				LabEquipmentTitles(x + 24).Width = VB6.TwipsToPixelsX(TempWide * (1575 / w))
				txtEquipmentValues(x + 24).Top = VB6.TwipsToPixelsY((TempHigh * (2010 / h)) + (x * y))
				txtEquipmentValues(x + 24).Left = VB6.TwipsToPixelsX(TempWide * (6360 / w))
				txtEquipmentValues(x + 24).Width = VB6.TwipsToPixelsX(TempWide * (1815 / w))
				txtNumberValues(x + 24).Top = VB6.TwipsToPixelsY((TempHigh * (2010 / h)) + (x * y))
				txtNumberValues(x + 24).Left = VB6.TwipsToPixelsX(TempWide * (8340 / w))
				txtNumberValues(x + 24).Width = VB6.TwipsToPixelsX(TempWide * (495 / w))
				
				LabEquipmentTitles(x + 27).Top = VB6.TwipsToPixelsY((TempHigh * (4320 / h)) + (x * y))
				LabEquipmentTitles(x + 27).Left = VB6.TwipsToPixelsX(TempWide * (4620 / w))
				LabEquipmentTitles(x + 27).Width = VB6.TwipsToPixelsX(TempWide * (1575 / w))
				txtEquipmentValues(x + 27).Top = VB6.TwipsToPixelsY((TempHigh * (4290 / h)) + (x * y))
				txtEquipmentValues(x + 27).Left = VB6.TwipsToPixelsX(TempWide * (6360 / w))
				txtEquipmentValues(x + 27).Width = VB6.TwipsToPixelsX(TempWide * (1815 / w))
				txtNumberValues(x + 27).Top = VB6.TwipsToPixelsY((TempHigh * (4290 / h)) + (x * y))
				txtNumberValues(x + 27).Left = VB6.TwipsToPixelsX(TempWide * (8340 / w))
				txtNumberValues(x + 27).Width = VB6.TwipsToPixelsX(TempWide * (495 / w))
				
				LabEquipmentTitles(x + 29).Top = VB6.TwipsToPixelsY((TempHigh * (4740 / h)) + (x * y))
				LabEquipmentTitles(x + 29).Left = VB6.TwipsToPixelsX(TempWide * (4620 / w))
				LabEquipmentTitles(x + 29).Width = VB6.TwipsToPixelsX(TempWide * (1575 / w))
				txtEquipmentValues(x + 29).Top = VB6.TwipsToPixelsY((TempHigh * (4710 / h)) + (x * y))
				txtEquipmentValues(x + 29).Left = VB6.TwipsToPixelsX(TempWide * (6360 / w))
				txtEquipmentValues(x + 29).Width = VB6.TwipsToPixelsX(TempWide * (1815 / w))
				txtNumberValues(x + 29).Top = VB6.TwipsToPixelsY((TempHigh * (4710 / h)) + (x * y))
				txtNumberValues(x + 29).Left = VB6.TwipsToPixelsX(TempWide * (8340 / w))
				txtNumberValues(x + 29).Width = VB6.TwipsToPixelsX(TempWide * (495 / w))
			End If
		Next x
		
		For x = 0 To 11
			lstEquipmentList(x).Top = VB6.TwipsToPixelsY(TempHigh * (3720 / h) + (x * q))
			lstEquipmentList(x).Left = VB6.TwipsToPixelsX(TempWide * (300 / w) + (x * r))
			lstEquipmentList(x).Height = VB6.TwipsToPixelsY(TempHigh * (1635 / h))
			lstEquipmentList(x).Width = VB6.TwipsToPixelsX(TempWide * (2235 / w))
			lstEquipmentList(x + 12).Top = VB6.TwipsToPixelsY(TempHigh * (3720 / h) + (x * q))
			lstEquipmentList(x + 12).Left = VB6.TwipsToPixelsX(TempWide * (300 / w) + (x * r))
			lstEquipmentList(x + 12).Height = VB6.TwipsToPixelsY(TempHigh * (1635 / h))
			lstEquipmentList(x + 12).Width = VB6.TwipsToPixelsX(TempWide * (2235 / w))
			MachineLabel(x).Top = VB6.TwipsToPixelsY(TempHigh * (3480 / h) + (x * q))
			MachineLabel(x).Left = VB6.TwipsToPixelsX(TempWide * (300 / w) + (x * r))
			MachineLabel(x).Width = VB6.TwipsToPixelsX(TempWide * (2235 / w))
		Next x
		
		'Only for Alistair
		'  lstEquipmentList(5).Top = TempHigh * (3720 / h) + (5 * q)
		'  lstEquipmentList(5).Left = TempWide * (480 / w)
		'  lstEquipmentList(5).Height = TempHigh * (1635 / h)
		'  lstEquipmentList(5).Width = TempWide * (3195 / w)
		'  lstEquipmentList(5 + 12).Top = TempHigh * (3720 / h) + (5 * q)
		'  lstEquipmentList(5 + 12).Left = TempWide * (480 / w)
		'  lstEquipmentList(5 + 12).Height = TempHigh * (1635 / h)
		'  lstEquipmentList(5 + 12).Width = TempWide * (3195 / w)
		'  MachineLabel(5).Top = TempHigh * (3480 / h) + (5 * q)
		'  MachineLabel(5).Left = TempWide * (480 / w)
		'  MachineLabel(5).Width = TempWide * (3195 / w)
		
		LineHorizontal(0).X1 = VB6.TwipsToPixelsX(TempWide * (2340 / w))
		LineHorizontal(0).X2 = VB6.TwipsToPixelsX(TempWide * (3840 / w))
		LineHorizontal(0).Y1 = VB6.TwipsToPixelsY(TempHigh * (3240 / h))
		LineHorizontal(0).Y2 = VB6.TwipsToPixelsY(TempHigh * (3240 / h))
		
		LineHorizontal(1).X1 = VB6.TwipsToPixelsX(TempWide * (120 / w))
		LineHorizontal(1).X2 = VB6.TwipsToPixelsX(TempWide * (4020 / w))
		LineHorizontal(1).Y1 = VB6.TwipsToPixelsY(TempHigh * (6060 / h))
		LineHorizontal(1).Y2 = VB6.TwipsToPixelsY(TempHigh * (6060 / h))
		
		LineHorizontal(2).X1 = VB6.TwipsToPixelsX(TempWide * (4020 / w))
		LineHorizontal(2).X2 = VB6.TwipsToPixelsX(TempWide * (9060 / w))
		LineHorizontal(2).Y1 = VB6.TwipsToPixelsY(TempHigh * (120 / h))
		LineHorizontal(2).Y2 = VB6.TwipsToPixelsY(TempHigh * (120 / h))
		
		LineHorizontal(3).X1 = VB6.TwipsToPixelsX(TempWide * (4440 / w))
		LineHorizontal(3).X2 = VB6.TwipsToPixelsX(TempWide * (8940 / w))
		LineHorizontal(3).Y1 = VB6.TwipsToPixelsY(TempHigh * (540 / h))
		LineHorizontal(3).Y2 = VB6.TwipsToPixelsY(TempHigh * (540 / h))
		
		LineHorizontal(4).X1 = VB6.TwipsToPixelsX(TempWide * (4680 / w))
		LineHorizontal(4).X2 = VB6.TwipsToPixelsX(TempWide * (8940 / w))
		LineHorizontal(4).Y1 = VB6.TwipsToPixelsY(TempHigh * (3240 / h))
		LineHorizontal(4).Y2 = VB6.TwipsToPixelsY(TempHigh * (3240 / h))
		
		LineHorizontal(5).X1 = VB6.TwipsToPixelsX(TempWide * (4020 / w))
		LineHorizontal(5).X2 = VB6.TwipsToPixelsX(TempWide * (9060 / w))
		LineHorizontal(5).Y1 = VB6.TwipsToPixelsY(TempHigh * (5940 / h))
		LineHorizontal(5).Y2 = VB6.TwipsToPixelsY(TempHigh * (5940 / h))
		
		LineVertical(0).X1 = VB6.TwipsToPixelsX(TempWide * (180 / w))
		LineVertical(0).X2 = VB6.TwipsToPixelsX(TempWide * (180 / w))
		LineVertical(0).Y1 = VB6.TwipsToPixelsY(TempHigh * (3480 / h))
		LineVertical(0).Y2 = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		
		LineVertical(1).X1 = VB6.TwipsToPixelsX(TempWide * (3960 / w))
		LineVertical(1).X2 = VB6.TwipsToPixelsX(TempWide * (3960 / w))
		LineVertical(1).Y1 = VB6.TwipsToPixelsY(TempHigh * (60 / h))
		LineVertical(1).Y2 = VB6.TwipsToPixelsY(TempHigh * (420 / h))
		
		LineVertical(2).X1 = VB6.TwipsToPixelsX(TempWide * (3960 / w))
		LineVertical(2).X2 = VB6.TwipsToPixelsX(TempWide * (3960 / w))
		LineVertical(2).Y1 = VB6.TwipsToPixelsY(TempHigh * (780 / h))
		LineVertical(2).Y2 = VB6.TwipsToPixelsY(TempHigh * (3120 / h))
		
		LineVertical(3).X1 = VB6.TwipsToPixelsX(TempWide * (3960 / w))
		LineVertical(3).X2 = VB6.TwipsToPixelsX(TempWide * (3960 / w))
		LineVertical(3).Y1 = VB6.TwipsToPixelsY(TempHigh * (3480 / h))
		LineVertical(3).Y2 = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		
		LineVertical(4).X1 = VB6.TwipsToPixelsX(TempWide * (9000 / w))
		LineVertical(4).X2 = VB6.TwipsToPixelsX(TempWide * (9000 / w))
		LineVertical(4).Y1 = VB6.TwipsToPixelsY(TempHigh * (60 / h))
		LineVertical(4).Y2 = VB6.TwipsToPixelsY(TempHigh * (6000 / h))
		
		labBackToMenu.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labBackToMenu.Left = VB6.TwipsToPixelsX(TempWide * (600 / w))
		
		imgBackToMenu.Top = VB6.TwipsToPixelsY(TempHigh * (6150 / h))
		imgBackToMenu.Left = VB6.TwipsToPixelsX(TempWide * (60 / w))
		imgBackToMenu.Width = VB6.TwipsToPixelsX(TempWide * (495 / w))
		
		comEquipmentPrint.Top = VB6.TwipsToPixelsY(TempHigh * (6150 / h))
		comEquipmentPrint.Left = VB6.TwipsToPixelsX(TempWide * (8880 / w))
		
		labEquipmentHelp.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labEquipmentHelp.Left = VB6.TwipsToPixelsX(TempWide * (8340 / w))
		
		labInsert.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labInsert.Left = VB6.TwipsToPixelsX(TempWide * (4080 / w))
		labInsert.Width = VB6.TwipsToPixelsX(TempWide * (975 / w))
		
	End Sub
	Public Sub drawthevalues()
		
		Dim i As Short
		Dim x As Short
        'Dim TempMachine As Short
        Dim ColorChange As Boolean
		
		ColorChange = False
		DoNotChange = True
		
		Call NumberToMachine()
		
		For i = 0 To 31
			Select Case i
				Case 0 To 3, 10 To 11, 22, 24, 27, 29 To 31
					LabEquipmentTitles(i).Enabled = True
					txtEquipmentValues(i).Enabled = True
					txtNumberValues(i).Enabled = True
			End Select
		Next i
		
		For i = 0 To 31
			Select Case i
				Case 0
					ColorChange = False
					For x = 0 To 3
						If CellValues(WhichScreen, x, WhichSegment).Changed = True Then
							ColorChange = True
						End If
					Next x
					If CellValues(WhichScreen, 20, WhichSegment).Changed = True Then ColorChange = True
					If ColorChange = True Then
						txtEquipmentValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF0000)
					Else
						txtEquipmentValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
					End If
					ColorChange = False
					For x = 0 To 3
						If CellValues(EquipmentTwo, x, WhichSegment).Changed = True Then
							ColorChange = True
						End If
					Next x
					If CellValues(EquipmentTwo, 20, WhichSegment).Changed = True Then ColorChange = True
					If ColorChange = True Then
						txtNumberValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF0000)
					Else
						txtNumberValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
					End If
				Case 1
					ColorChange = False
					If CellValues(WhichScreen, 4, WhichSegment).Changed = True Then
						ColorChange = True
					End If
					If CellValues(WhichScreen, 21, WhichSegment).Changed = True Then
						ColorChange = True
					End If
					If ColorChange = True Then
						txtEquipmentValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF0000)
					Else
						txtEquipmentValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
					End If
					ColorChange = False
					If CellValues(EquipmentTwo, 4, WhichSegment).Changed = True Then
						ColorChange = True
					End If
					If CellValues(EquipmentTwo, 21, WhichSegment).Changed = True Then
						ColorChange = True
					End If
					If ColorChange = True Then
						txtNumberValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF0000)
					Else
						txtNumberValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
					End If
				Case 2
					ColorChange = False
					For x = 5 To 8
						If CellValues(WhichScreen, x, WhichSegment).Changed = True Then
							ColorChange = True
						End If
					Next x
					If CellValues(WhichScreen, 25, WhichSegment).Changed = True Then ColorChange = True
					If ColorChange = True Then
						txtEquipmentValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF0000)
					Else
						txtEquipmentValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
					End If
					ColorChange = False
					For x = 5 To 8
						If CellValues(EquipmentTwo, x, WhichSegment).Changed = True Then
							ColorChange = True
						End If
					Next x
					If CellValues(EquipmentTwo, 25, WhichSegment).Changed = True Then ColorChange = True
					If ColorChange = True Then
						txtNumberValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF0000)
					Else
						txtNumberValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
					End If
				Case 3
					ColorChange = False
					If CellValues(WhichScreen, 9, WhichSegment).Changed = True Then
						ColorChange = True
					End If
					If CellValues(WhichScreen, 26, WhichSegment).Changed = True Then
						ColorChange = True
					End If
					If ColorChange = True Then
						txtEquipmentValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF0000)
					Else
						txtEquipmentValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
					End If
					ColorChange = False
					If CellValues(EquipmentTwo, 9, WhichSegment).Changed = True Then
						ColorChange = True
					End If
					If CellValues(EquipmentTwo, 26, WhichSegment).Changed = True Then
						ColorChange = True
					End If
					If ColorChange = True Then
						txtNumberValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF0000)
					Else
						txtNumberValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
					End If
				Case 10, 11
					If CellValues(WhichScreen, i, WhichSegment).Changed = True Then
						txtEquipmentValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF0000)
					Else
						txtEquipmentValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
					End If
					If CellValues(EquipmentTwo, i, WhichSegment).Changed = True Then
						txtNumberValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF0000)
					Else
						txtNumberValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
					End If
				Case 22, 27
					ColorChange = False
					If CellValues(WhichScreen, i, WhichSegment).Changed = True Then
						ColorChange = True
					End If
					If CellValues(WhichScreen, i + 1, WhichSegment).Changed = True Then
						ColorChange = True
					End If
					If ColorChange = True Then
						txtEquipmentValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF0000)
					Else
						txtEquipmentValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
					End If
					ColorChange = False
					If CellValues(EquipmentTwo, i, WhichSegment).Changed = True Then
						ColorChange = True
					End If
					If CellValues(EquipmentTwo, i + 1, WhichSegment).Changed = True Then
						ColorChange = True
					End If
					If ColorChange = True Then
						txtNumberValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF0000)
					Else
						txtNumberValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
					End If
				Case 24, 29
					If CellValues(WhichScreen, i, WhichSegment).Changed = True Then
						txtEquipmentValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF0000)
					Else
						txtEquipmentValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
					End If
					If CellValues(EquipmentTwo, i, WhichSegment).Changed = True Then
						txtNumberValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF0000)
					Else
						txtNumberValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
					End If
				Case 30, 31
					If CellValues(WhichScreen, i, WhichSegment).Changed = True Then
						txtEquipmentValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF0000)
					Else
						txtEquipmentValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
					End If
					If CellValues(EquipmentTwo, i, WhichSegment).Changed = True Then
						txtNumberValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF0000)
					Else
						txtNumberValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
					End If
			End Select
		Next i
		
		For i = 0 To 31
			Select Case i
				Case 0
					For x = 0 To 3
						If CellValues(WhichScreen, x, WhichSegment).Value <> 0 Then
							Select Case x
								Case 0
									LabEquipmentTitles(i).Text = "Front-End Loader"
								Case 1
									LabEquipmentTitles(i).Text = "Hydraulic Shovel"
								Case 2
									LabEquipmentTitles(i).Text = "Mechanical Shovel"
								Case 3
									LabEquipmentTitles(i).Text = "Walking Dragline"
							End Select
							txtEquipmentValues(i).Text = LTrim(RTrim(CellValues(WhichScreen, x, WhichSegment).Word))
							txtNumberValues(i).Text = VB6.Format(LTrim(Str(CellValues(EquipmentTwo, x, WhichSegment).Value)), "##0")
						End If
					Next x
					If CellValues(WhichScreen, 20, WhichSegment).Value <> 0 Then
						LabEquipmentTitles(i).Text = "Scrapers"
						txtEquipmentValues(i).Text = LTrim(RTrim(CellValues(WhichScreen, 20, WhichSegment).Word))
						txtNumberValues(i).Text = VB6.Format(LTrim(Str(CellValues(EquipmentTwo, 20, WhichSegment).Value)), "##0")
						LabEquipmentTitles(i + 1).Enabled = False
						txtEquipmentValues(i + 1).Enabled = False
						txtNumberValues(i + 1).Enabled = False
					End If
				Case 1
					If CellValues(WhichScreen, 4, WhichSegment).Value <> 0 Then
						LabEquipmentTitles(i).Text = "Rear-Dump Truck"
						txtEquipmentValues(i).Text = LTrim(RTrim(CellValues(WhichScreen, 4, WhichSegment).Word))
						txtNumberValues(i).Text = VB6.Format(LTrim(Str(CellValues(EquipmentTwo, 4, WhichSegment).Value)), "##0")
					ElseIf CellValues(WhichScreen, 21, WhichSegment).Value <> 0 Then 
						LabEquipmentTitles(i).Text = "Articulated Haulers"
						txtEquipmentValues(i).Text = LTrim(RTrim(CellValues(WhichScreen, 21, WhichSegment).Word))
						txtNumberValues(i).Text = VB6.Format(LTrim(Str(CellValues(EquipmentTwo, 21, WhichSegment).Value)), "##0")
					End If
				Case 2
					For x = 5 To 8
						If CellValues(WhichScreen, x, WhichSegment).Value <> 0 Then
							Select Case x
								Case 5
									LabEquipmentTitles(i).Text = "Front-End Loader"
								Case 6
									LabEquipmentTitles(i).Text = "Hydraulic Shovel"
								Case 7
									LabEquipmentTitles(i).Text = "Mechanical Shovel"
								Case 8
									LabEquipmentTitles(i).Text = "Walking Dragline"
							End Select
							txtEquipmentValues(i).Text = LTrim(RTrim(CellValues(WhichScreen, x, WhichSegment).Word))
							txtNumberValues(i).Text = VB6.Format(LTrim(Str(CellValues(EquipmentTwo, x, WhichSegment).Value)), "##0")
						End If
					Next x
					If CellValues(WhichScreen, 25, WhichSegment).Value <> 0 Then
						LabEquipmentTitles(i).Text = "Scrapers"
						txtEquipmentValues(i).Text = LTrim(RTrim(CellValues(WhichScreen, 25, WhichSegment).Word))
						txtNumberValues(i).Text = VB6.Format(LTrim(Str(CellValues(EquipmentTwo, 25, WhichSegment).Value)), "##0")
						LabEquipmentTitles(i + 1).Enabled = False
						txtEquipmentValues(i + 1).Enabled = False
						txtNumberValues(i + 1).Enabled = False
					End If
				Case 3
					If CellValues(WhichScreen, 9, WhichSegment).Value <> 0 Then
						LabEquipmentTitles(i).Text = "Rear-Dump Truck"
						txtEquipmentValues(i).Text = LTrim(RTrim(CellValues(WhichScreen, 9, WhichSegment).Word))
						txtNumberValues(i).Text = VB6.Format(LTrim(Str(CellValues(EquipmentTwo, 9, WhichSegment).Value)), "##0")
					ElseIf CellValues(WhichScreen, 26, WhichSegment).Value <> 0 Then 
						LabEquipmentTitles(i).Text = "Articulated Haulers"
						txtEquipmentValues(i).Text = LTrim(RTrim(CellValues(WhichScreen, 26, WhichSegment).Word))
						txtNumberValues(i).Text = VB6.Format(LTrim(Str(CellValues(EquipmentTwo, 26, WhichSegment).Value)), "##0")
					End If
				Case 10, 11
					If CellValues(WhichScreen, i, WhichSegment).Value <> 0 Then
						txtEquipmentValues(i).Text = LTrim(RTrim(CellValues(WhichScreen, i, WhichSegment).Word))
						txtNumberValues(i).Text = VB6.Format(LTrim(Str(CellValues(EquipmentTwo, i, WhichSegment).Value)), "##0")
						LabEquipmentTitles(i).Enabled = True
						txtEquipmentValues(i).Enabled = True
						txtNumberValues(i).Enabled = True
					Else
						LabEquipmentTitles(i).Enabled = False
						txtEquipmentValues(i).Enabled = False
						txtNumberValues(i).Enabled = False
					End If
				Case 22, 27
					If CellValues(WhichScreen, i, WhichSegment).Value <> 0 Then
						txtEquipmentValues(i).Text = LTrim(RTrim(CellValues(WhichScreen, i, WhichSegment).Word))
						txtNumberValues(i).Text = VB6.Format(LTrim(Str(CellValues(EquipmentTwo, i, WhichSegment).Value)), "##0")
					End If
					If CellValues(WhichScreen, i, WhichSegment).Value <> 0 Then
						txtEquipmentValues(i).Text = LTrim(RTrim(CellValues(WhichScreen, i, WhichSegment).Word))
						txtNumberValues(i).Text = VB6.Format(LTrim(Str(CellValues(EquipmentTwo, i, WhichSegment).Value)), "##0")
					End If
					If CellValues(EquipmentTwo, i, WhichSegment).Value <> 0 Then
						LabEquipmentTitles(i).Text = "Jaw Crusher"
						txtEquipmentValues(i).Text = LTrim(RTrim(CellValues(WhichScreen, i, WhichSegment).Word))
						txtNumberValues(i).Text = VB6.Format(LTrim(Str(CellValues(EquipmentTwo, i, WhichSegment).Value)), "##0")
					ElseIf CellValues(EquipmentTwo, i + 1, WhichSegment).Value <> 0 Then 
						LabEquipmentTitles(i).Text = "Gyratory Crusher"
						txtEquipmentValues(i).Text = LTrim(RTrim(CellValues(WhichScreen, i + 1, WhichSegment).Word))
						txtNumberValues(i).Text = VB6.Format(LTrim(Str(CellValues(EquipmentTwo, i + 1, WhichSegment).Value)), "##0")
					Else
						LabEquipmentTitles(i).Enabled = False
						txtEquipmentValues(i).Enabled = False
						txtNumberValues(i).Enabled = False
					End If
				Case 24, 29
					If CellValues(WhichScreen, i, WhichSegment).Value <> 0 Then
						txtEquipmentValues(i).Text = LTrim(RTrim(CellValues(WhichScreen, i, WhichSegment).Word))
						txtNumberValues(i).Text = VB6.Format(LTrim(Str(CellValues(EquipmentTwo, i, WhichSegment).Value)), "##0")
						If i = 24 Then
							LabEquipmentTitles(i - 23).Enabled = False
							txtEquipmentValues(i - 23).Enabled = False
							txtNumberValues(i - 23).Enabled = False
						Else
							LabEquipmentTitles(i - 26).Enabled = False
							txtEquipmentValues(i - 26).Enabled = False
							txtNumberValues(i - 26).Enabled = False
						End If
					End If
					If CellValues(EquipmentTwo, i, WhichSegment).Value <> 0 Then
						LabEquipmentTitles(i).Enabled = True
						txtEquipmentValues(i).Text = LTrim(RTrim(CellValues(WhichScreen, i, WhichSegment).Word))
						txtNumberValues(i).Text = VB6.Format(LTrim(Str(CellValues(EquipmentTwo, i, WhichSegment).Value)), "##0")
					Else
						LabEquipmentTitles(i).Enabled = False
						txtEquipmentValues(i).Enabled = False
						txtNumberValues(i).Enabled = False
					End If
				Case 30, 31
					If CellValues(WhichScreen, i, WhichSegment).Value <> 0 Then
						txtEquipmentValues(i).Text = LTrim(RTrim(CellValues(WhichScreen, i, WhichSegment).Word))
						txtNumberValues(i).Text = VB6.Format(LTrim(Str(CellValues(EquipmentTwo, i, WhichSegment).Value)), "##0")
						LabEquipmentTitles(i).Enabled = True
						txtEquipmentValues(i).Enabled = True
						txtNumberValues(i).Enabled = True
					Else
						LabEquipmentTitles(i).Enabled = False
						txtEquipmentValues(i).Enabled = False
						txtNumberValues(i).Enabled = False
					End If
			End Select
		Next i
		
		DoNotChange = False
		
	End Sub
	Private Sub txtEquipmentValues_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtEquipmentValues.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Dim Index As Short = txtEquipmentValues.GetIndex(eventSender)
		
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
						If InStr(txtEquipmentValues(Index).Text, ".") = 0 Then
							System.Windows.Forms.SendKeys.Send("{DELETE}")
						End If
					Else
						System.Windows.Forms.SendKeys.Send("{DELETE}")
					End If
			End Select
		End If
		
	End Sub
	Private Sub txtNumberValues_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtNumberValues.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Dim Index As Short = txtNumberValues.GetIndex(eventSender)
		
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
						If InStr(txtNumberValues(Index).Text, ".") = 0 Then
							System.Windows.Forms.SendKeys.Send("{DELETE}")
						End If
					Else
						System.Windows.Forms.SendKeys.Send("{DELETE}")
					End If
			End Select
		End If
		
	End Sub
	Private Sub txtEquipmentValues_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtEquipmentValues.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = txtEquipmentValues.GetIndex(eventSender)
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
	Private Sub txtNumberValues_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtNumberValues.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = txtNumberValues.GetIndex(eventSender)
		Dim x As Short
		If KeyAscii > Asc("9") And KeyAscii <> Asc(",") And KeyAscii <> Asc(".") And KeyAscii <> Asc("$") Then
			Beep()
			KeyAscii = 0
		Else
			Select Case WhichCell
				Case 0
					For x = 0 To 3
						If CellValues(EquipmentTwo, x, WhichSegment).Value <> 0 Then
							CellValues(EquipmentTwo, x, WhichSegment).Changed = True
						End If
					Next x
					If CellValues(EquipmentTwo, 20, WhichSegment).Value <> 0 Then
						CellValues(EquipmentTwo, 20, WhichSegment).Changed = True
					End If
				Case 1
					If CellValues(EquipmentTwo, 4, WhichSegment).Value <> 0 Then
						CellValues(EquipmentTwo, 4, WhichSegment).Changed = True
					ElseIf CellValues(EquipmentTwo, 21, WhichSegment).Value <> 0 Then 
						CellValues(EquipmentTwo, 21, WhichSegment).Changed = True
					End If
				Case 2
					For x = 5 To 8
						If CellValues(EquipmentTwo, x, WhichSegment).Value <> 0 Then
							CellValues(EquipmentTwo, x, WhichSegment).Changed = True
						End If
					Next x
					If CellValues(EquipmentTwo, 25, WhichSegment).Value <> 0 Then
						CellValues(EquipmentTwo, 25, WhichSegment).Changed = True
					End If
				Case 3
					If CellValues(EquipmentTwo, 9, WhichSegment).Value <> 0 Then
						CellValues(EquipmentTwo, 9, WhichSegment).Changed = True
					ElseIf CellValues(EquipmentTwo, 26, WhichSegment).Value <> 0 Then 
						CellValues(EquipmentTwo, 26, WhichSegment).Changed = True
					End If
				Case 10, 11
					If CellValues(EquipmentTwo, WhichCell, WhichSegment).Value <> 0 Then
						CellValues(EquipmentTwo, WhichCell, WhichSegment).Changed = True
					End If
				Case 22, 27
					If CellValues(EquipmentTwo, WhichCell, WhichSegment).Value <> 0 Then
						CellValues(EquipmentTwo, WhichCell, WhichSegment).Changed = True
					ElseIf CellValues(EquipmentTwo, WhichCell + 1, WhichSegment).Value <> 0 Then 
						CellValues(EquipmentTwo, WhichCell + 1, WhichSegment).Changed = True
					End If
				Case 24, 29
					If CellValues(EquipmentTwo, WhichCell, WhichSegment).Value <> 0 Then
						CellValues(EquipmentTwo, WhichCell, WhichSegment).Changed = True
					End If
				Case 30, 31
					If CellValues(EquipmentTwo, WhichCell, WhichSegment).Value <> 0 Then
						CellValues(EquipmentTwo, WhichCell, WhichSegment).Changed = True
					End If
			End Select
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	Private Sub txtEquipmentValues_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtEquipmentValues.Leave
		Dim Index As Short = txtEquipmentValues.GetIndex(eventSender)
		If DoNotChange = True Then Exit Sub
		WhichCell = Index
		Call Inputer(WhichCell)
	End Sub
	Private Sub txtNumberValues_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNumberValues.Leave
		Dim Index As Short = txtNumberValues.GetIndex(eventSender)
		If DoNotChange = True Then Exit Sub
		WhichCell = Index
		Call NumberInputer(WhichCell)
	End Sub
	Private Sub Inputer(ByRef Sample As Short)
		Dim x As Short
		Dim i As Short
		Dim life As Decimal
		Dim TempSegment As Short
		Dim SetSegment As Short
		Dim tempvalue As String
		Dim Digit As New VB6.FixedLengthString(1)
		
		On Error Resume Next
		
		If DoNotChange = True Then Exit Sub
		PageChange(WhichScreen) = True
		i = Sample
		Select Case i
			Case 0
				For x = 0 To 3
					If CellValues(WhichScreen, x, WhichSegment).Value <> 0 Then
						For SetSegment = 0 To MaxSegment
							If CellValues(WhichScreen, x, SetSegment).Changed = True Then
								CellValues(WhichScreen, x, SetSegment).Word = LTrim(RTrim(txtEquipmentValues(Sample).Text))
							End If
						Next SetSegment
					End If
				Next x
				If CellValues(WhichScreen, 20, WhichSegment).Value <> 0 Then
					For SetSegment = 0 To MaxSegment
						If CellValues(WhichScreen, 20, SetSegment).Changed = True Then
							CellValues(WhichScreen, 20, SetSegment).Word = LTrim(RTrim(txtEquipmentValues(Sample).Text))
						End If
					Next SetSegment
				End If
			Case 1
				If CellValues(WhichScreen, 4, WhichSegment).Value <> 0 Then
					For SetSegment = 0 To MaxSegment
						If CellValues(WhichScreen, 4, SetSegment).Changed = True Then
							CellValues(WhichScreen, 4, SetSegment).Word = LTrim(RTrim(txtEquipmentValues(Sample).Text))
						End If
					Next SetSegment
				ElseIf CellValues(WhichScreen, 21, WhichSegment).Value <> 0 Then 
					For SetSegment = 0 To MaxSegment
						If CellValues(WhichScreen, 21, SetSegment).Changed = True Then
							CellValues(WhichScreen, 21, SetSegment).Word = LTrim(RTrim(txtEquipmentValues(Sample).Text))
						End If
					Next SetSegment
				End If
			Case 2
				For x = 5 To 8
					If CellValues(WhichScreen, x, WhichSegment).Value <> 0 Then
						For SetSegment = 0 To MaxSegment
							If CellValues(WhichScreen, x, SetSegment).Changed = True Then
								CellValues(WhichScreen, x, SetSegment).Word = LTrim(RTrim(txtEquipmentValues(Sample).Text))
							End If
						Next SetSegment
					End If
				Next x
				If CellValues(WhichScreen, 25, WhichSegment).Value <> 0 Then
					For SetSegment = 0 To MaxSegment
						If CellValues(WhichScreen, 25, SetSegment).Changed = True Then
							CellValues(WhichScreen, 25, SetSegment).Word = LTrim(RTrim(txtEquipmentValues(Sample).Text))
						End If
					Next SetSegment
				End If
			Case 3
				If CellValues(WhichScreen, 9, WhichSegment).Value <> 0 Then
					For SetSegment = 0 To MaxSegment
						If CellValues(WhichScreen, 9, SetSegment).Changed = True Then
							CellValues(WhichScreen, 9, SetSegment).Word = LTrim(RTrim(txtEquipmentValues(Sample).Text))
						End If
					Next SetSegment
				ElseIf CellValues(WhichScreen, 26, WhichSegment).Value <> 0 Then 
					For SetSegment = 0 To MaxSegment
						If CellValues(WhichScreen, 26, SetSegment).Changed = True Then
							CellValues(WhichScreen, 26, SetSegment).Word = LTrim(RTrim(txtEquipmentValues(Sample).Text))
						End If
					Next SetSegment
				End If
			Case 10, 11
				If CellValues(WhichScreen, i, WhichSegment).Value <> 0 Then
					For SetSegment = 0 To MaxSegment
						If CellValues(WhichScreen, i, SetSegment).Changed = True Then
							CellValues(WhichScreen, i, SetSegment).Word = LTrim(RTrim(txtEquipmentValues(Sample).Text))
						End If
					Next SetSegment
				End If
			Case 22, 27
				If CellValues(WhichScreen, i, WhichSegment).Value <> 0 Then
					For SetSegment = 0 To MaxSegment
						If CellValues(WhichScreen, i, SetSegment).Changed = True Then
							CellValues(WhichScreen, i, SetSegment).Word = LTrim(RTrim(txtEquipmentValues(Sample).Text))
						End If
					Next SetSegment
				ElseIf CellValues(WhichScreen, i + 1, WhichSegment).Value <> 0 Then 
					For SetSegment = 0 To MaxSegment
						If CellValues(WhichScreen, i + 1, SetSegment).Changed = True Then
							CellValues(WhichScreen, i + 1, SetSegment).Word = LTrim(RTrim(txtEquipmentValues(Sample).Text))
						End If
					Next SetSegment
				End If
			Case 24, 29
				If CellValues(WhichScreen, i, WhichSegment).Value <> 0 Then
					For SetSegment = 0 To MaxSegment
						If CellValues(WhichScreen, i, SetSegment).Changed = True Then
							CellValues(WhichScreen, i, SetSegment).Word = LTrim(RTrim(txtEquipmentValues(Sample).Text))
						End If
					Next SetSegment
				End If
			Case 30, 31
				If CellValues(WhichScreen, i, WhichSegment).Value <> 0 Then
					For SetSegment = 0 To MaxSegment
						If CellValues(WhichScreen, i, SetSegment).Changed = True Then
							CellValues(WhichScreen, i, SetSegment).Word = LTrim(RTrim(txtEquipmentValues(Sample).Text))
						End If
					Next SetSegment
				End If
		End Select
		
		TempSegment = WhichSegment
		For WhichSegment = 0 To MaxSegment
			Call MachineToNumber()
		Next WhichSegment
		WhichSegment = TempSegment
		Call CostItAll()
		Call drawthevalues()
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
		For i = 1 To Len(txtNumberValues(Sample).Text)
			Digit.Value = Mid(txtNumberValues(Sample).Text, i, 1)
			Select Case Digit.Value
				Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ".", "-"
					tempvalue = tempvalue & Digit.Value
			End Select
		Next i
		i = Sample
		Select Case i
			Case 0
				For x = 0 To 3
					If CellValues(EquipmentTwo, x, WhichSegment).Value <> 0 Then
						If CellValues(EquipmentTwo, x, WhichSegment).Changed = True Then
							CellValues(EquipmentTwo, x, WhichSegment).Value = CDec(tempvalue)
						End If
					End If
				Next x
				If CellValues(EquipmentTwo, 20, WhichSegment).Value <> 0 Then
					If CellValues(EquipmentTwo, 20, WhichSegment).Changed = True Then
						CellValues(EquipmentTwo, 20, WhichSegment).Value = CDec(tempvalue)
					End If
				End If
			Case 1
				If CellValues(EquipmentTwo, 4, WhichSegment).Value <> 0 Then
					If CellValues(EquipmentTwo, 4, WhichSegment).Changed = True Then
						CellValues(EquipmentTwo, 4, WhichSegment).Value = CDec(tempvalue)
					End If
				ElseIf CellValues(EquipmentTwo, 21, WhichSegment).Value <> 0 Then 
					If CellValues(EquipmentTwo, 21, WhichSegment).Changed = True Then
						CellValues(EquipmentTwo, 21, WhichSegment).Value = CDec(tempvalue)
					End If
				End If
			Case 2
				For x = 5 To 8
					If CellValues(EquipmentTwo, x, WhichSegment).Value <> 0 Then
						If CellValues(EquipmentTwo, x, WhichSegment).Changed = True Then
							CellValues(EquipmentTwo, x, WhichSegment).Value = CDec(tempvalue)
						End If
					End If
				Next x
				If CellValues(EquipmentTwo, 25, WhichSegment).Value <> 0 Then
					If CellValues(EquipmentTwo, 25, WhichSegment).Changed = True Then
						CellValues(EquipmentTwo, 25, WhichSegment).Value = CDec(tempvalue)
					End If
				End If
			Case 3
				If CellValues(EquipmentTwo, 9, WhichSegment).Value <> 0 Then
					If CellValues(EquipmentTwo, 9, WhichSegment).Changed = True Then
						CellValues(EquipmentTwo, 9, WhichSegment).Value = CDec(tempvalue)
					End If
				ElseIf CellValues(EquipmentTwo, 26, WhichSegment).Value <> 0 Then 
					If CellValues(EquipmentTwo, 26, WhichSegment).Changed = True Then
						CellValues(EquipmentTwo, 26, WhichSegment).Value = CDec(tempvalue)
					End If
				End If
			Case 10, 11
				If CellValues(EquipmentTwo, i, WhichSegment).Value <> 0 Then
					If CellValues(EquipmentTwo, i, WhichSegment).Changed = True Then
						CellValues(EquipmentTwo, i, WhichSegment).Value = CDec(tempvalue)
					End If
				End If
			Case 22, 27
				If CellValues(EquipmentTwo, i, WhichSegment).Value <> 0 Then
					If CellValues(EquipmentTwo, i, WhichSegment).Changed = True Then
						CellValues(EquipmentTwo, i, WhichSegment).Value = CDec(tempvalue)
					End If
				ElseIf CellValues(EquipmentTwo, i + 1, WhichSegment).Value <> 0 Then 
					If CellValues(EquipmentTwo, i + 1, WhichSegment).Changed = True Then
						CellValues(EquipmentTwo, i + 1, WhichSegment).Value = CDec(tempvalue)
					End If
				End If
			Case 24, 29
				If CellValues(EquipmentTwo, i, WhichSegment).Value <> 0 Then
					If CellValues(EquipmentTwo, i, WhichSegment).Changed = True Then
						CellValues(EquipmentTwo, i, WhichSegment).Value = CDec(tempvalue)
					End If
				End If
			Case 30, 31
				If CellValues(EquipmentTwo, i, WhichSegment).Value <> 0 Then
					If CellValues(EquipmentTwo, i, WhichSegment).Changed = True Then
						CellValues(EquipmentTwo, i, WhichSegment).Value = CDec(tempvalue)
					End If
				End If
		End Select
		Call CostItAll()
		Call drawthevalues()
	End Sub
	'UPGRADE_WARNING: Event txtSegmentLabel.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtSegmentLabel_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSegmentLabel.TextChanged
		If DoNotChange = True Then Exit Sub
		SegNamie(WhichSegment) = txtSegmentLabel.Text
	End Sub
End Class