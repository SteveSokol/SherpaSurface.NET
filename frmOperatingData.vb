Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks
Friend Class frmOperatingData
	Inherits System.Windows.Forms.Form
	Dim TempHigh As Single
	Dim TempWide As Single
	Dim TonConversion As Single
	Private Sub comOperatingHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles comOperatingHelp.Click
		PrintForm1.Print(Me, PowerPacks.Printing.PrintForm.PrintOption.CompatibleModeClientAreaOnly)
	End Sub
	'UPGRADE_WARNING: Form event frmOperatingData.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmOperatingData_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated

        'Dim baseunit As String
        'Dim baselength As Short
        Dim i As Short
		
		If IsHelpOn = True Then
			txtOperatingValues(WhichCell).Focus()
			IsHelpOn = False
		Else
			For i = 0 To 5
				If optSegment(i).Checked = True Then WhichSegment = i
			Next i
			
			WhichScreen = Production
			
			Call ScreenCalcs()
			
			Call drawthevalues()
			
			If InsertFlag = True Then
				labInsert.Text = "Insert"
			Else
				labInsert.Text = "Typeover"
			End If
			
			WhichCell = 0
			
			txtOperatingValues(0).Focus()
		End If
		
	End Sub
	Private Sub frmOperatingData_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        'Dim i As Short
        'Dim x As Short

        Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - (VB6.PixelsToTwipsY(Me.Height) + 350)) / 2)
		Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2)
		
		If VB6.PixelsToTwipsY(Me.Top) < 0 Then Me.Top = 0
		If VB6.PixelsToTwipsX(Me.Left) < 0 Then Me.Left = 0
		
		TempWide = VB6.PixelsToTwipsX(Me.ClientRectangle.Width)
		TempHigh = VB6.PixelsToTwipsY(Me.ClientRectangle.Height)
		
		DoNotChange = True
		
		If UnitType = Metric Then
			labOperatingUnits(0).Text = "metric tons"
			labOperatingUnits(4).Text = "metric tons"
			labOperatingUnits(5).Text = "metric tons per day"
			labOperatingUnits(10).Text = "metric tons per day"
			labOperatingUnits(15).Text = "metric tons"
			labOperatingLabels(6).Text = "Metric Tons Mined This Segment"
			TonConversion = 0.907185
		Else
			labOperatingUnits(0).Text = "tons"
			labOperatingUnits(4).Text = "tons"
			labOperatingUnits(5).Text = "tons per day"
			labOperatingUnits(10).Text = "tons per day"
			labOperatingUnits(15).Text = "tons"
			labOperatingLabels(6).Text = "Tons Mined This Segment"
			TonConversion = 1
		End If
		
		Call ScreenCalcs()
		
		WhichSegment = 0
		optSegment(WhichSegment).Checked = True
		txtSegmentLabel.Text = SegNamie(WhichSegment)
		
		If PageChange(WhichScreen) = True Then
			Call drawthevalues()
		End If
		
		DoNotChange = False
		
		Call screenstuff()
		
	End Sub
	'UPGRADE_WARNING: Event frmOperatingData.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub frmOperatingData_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		
		TempWide = VB6.PixelsToTwipsX(Me.ClientRectangle.Width)
		TempHigh = VB6.PixelsToTwipsY(Me.ClientRectangle.Height)
		
		Call screenstuff()
		
	End Sub
	Private Sub frmOperatingData_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		Call ScreenCalcs()
        'Me.Close()
		Call InputMenuAccess(1)
	End Sub
	Private Sub imgBackToMenu_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles imgBackToMenu.Click
		Call ScreenCalcs()
		Me.Close()
		Call InputMenuAccess(1)
	End Sub
	Private Sub labBackToMenu_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles labBackToMenu.Click
		Call ScreenCalcs()
		Me.Close()
		Call InputMenuAccess(1)
	End Sub
	Private Sub labOperatingHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles labOperatingHelp.Click
		Dim StartHelp As Short
		StartHelp = 0
		IsHelpOn = True
		Call frmSurfaceHelp.gethelptext(StartHelp, WhichCell)
		frmSurfaceHelp.Show()
	End Sub
	'UPGRADE_WARNING: Event lstOptions.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub lstOptions_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstOptions.SelectedIndexChanged
		Dim Index As Short = lstOptions.GetIndex(eventSender)
		Dim x As Short
		Dim ot As Decimal
		Dim wt As Decimal
		Dim TonsPerDay As Decimal
		
		Call ClearTheEquip(Index)
		
		txtOperatingValues(WhichCell).Text = LTrim(RTrim(VB6.GetItemString(lstOptions(Index), lstOptions(Index).SelectedIndex)))
		CellValues(WhichScreen, WhichCell - 2, WhichSegment).Changed = True
		
		If WhichCell = 8 And LTrim(RTrim(LCase(VB6.GetItemString(lstOptions(Index), lstOptions(Index).SelectedIndex)))) = "optimize" Then
			CellValues(Production, 40, 0).Word = "Sherpa"
		End If
		
		If LTrim(RTrim(LCase(CellValues(Production, 40, 0).Word))) = "sherpa" Then
			Call otcal(ot, wt)
			If ot < 20000 Then
				txtOperatingValues(10).Text = "Hydraulic Shovel"
				WhichCell = 10
				CellValues(WhichScreen, WhichCell - 2, WhichSegment).Changed = True
				Call Inputer(WhichCell)
				txtOperatingValues(11).Text = "Rear-Dump Truck"
				WhichCell = 11
				CellValues(WhichScreen, WhichCell - 2, WhichSegment).Changed = True
				Call Inputer(WhichCell)
				txtOperatingValues(15).Text = "Front-End Loader"
				WhichCell = 15
				CellValues(WhichScreen, WhichCell - 2, WhichSegment).Changed = True
				Call Inputer(WhichCell)
				txtOperatingValues(16).Text = "Rear-Dump Truck"
				WhichCell = 16
				CellValues(WhichScreen, WhichCell - 2, WhichSegment).Changed = True
				Call Inputer(WhichCell)
			Else
				txtOperatingValues(10).Text = "Mechanical Shovel"
				WhichCell = 10
				CellValues(WhichScreen, WhichCell - 2, WhichSegment).Changed = True
				Call Inputer(WhichCell)
				txtOperatingValues(11).Text = "Rear-Dump Truck"
				WhichCell = 11
				CellValues(WhichScreen, WhichCell - 2, WhichSegment).Changed = True
				Call Inputer(WhichCell)
				txtOperatingValues(15).Text = "Mechanical Shovel"
				WhichCell = 15
				CellValues(WhichScreen, WhichCell - 2, WhichSegment).Changed = True
				Call Inputer(WhichCell)
				txtOperatingValues(16).Text = "Rear-Dump Truck"
				WhichCell = 16
				CellValues(WhichScreen, WhichCell - 2, WhichSegment).Changed = True
				Call Inputer(WhichCell)
			End If
			Call tonsp(TonsPerDay)
			If TonsPerDay < 2501 Then
				txtOperatingValues(8).Text = "Percussion Drill"
				WhichCell = 8
				CellValues(WhichScreen, WhichCell - 2, WhichSegment).Changed = True
				Call Inputer(WhichCell)
				txtOperatingValues(9).Text = "Dynamite"
				WhichCell = 9
				CellValues(WhichScreen, WhichCell - 2, WhichSegment).Changed = True
				Call Inputer(WhichCell)
				txtOperatingValues(13).Text = "Percussion Drill"
				WhichCell = 13
				CellValues(WhichScreen, WhichCell - 2, WhichSegment).Changed = True
				Call Inputer(WhichCell)
				txtOperatingValues(14).Text = "Dynamite"
				WhichCell = 14
				CellValues(WhichScreen, WhichCell - 2, WhichSegment).Changed = True
				Call Inputer(WhichCell)
				bltp = 1
			Else
				txtOperatingValues(8).Text = "Rotary Drill"
				WhichCell = 8
				CellValues(WhichScreen, WhichCell - 2, WhichSegment).Changed = True
				Call Inputer(WhichCell)
				txtOperatingValues(9).Text = "ANFO"
				WhichCell = 9
				CellValues(WhichScreen, WhichCell - 2, WhichSegment).Changed = True
				Call Inputer(WhichCell)
				txtOperatingValues(13).Text = "Rotary Drill"
				WhichCell = 13
				CellValues(WhichScreen, WhichCell - 2, WhichSegment).Changed = True
				Call Inputer(WhichCell)
				txtOperatingValues(14).Text = "ANFO"
				WhichCell = 14
				CellValues(WhichScreen, WhichCell - 2, WhichSegment).Changed = True
				Call Inputer(WhichCell)
				bltp = 2
			End If
			
			WhichCell = 8
			For x = 9 To 12
				If x < 12 Then
					txtOperatingValues(x).Enabled = False
					LabOperatingTitles(x - 2).Enabled = False
					txtOperatingValues(x + 4).Enabled = False
					LabOperatingTitles(x + 2).Enabled = False
				Else
					txtOperatingValues(x + 4).Enabled = False
					LabOperatingTitles(x + 2).Enabled = False
				End If
			Next x
		Else
			For x = 9 To 12
				If x < 12 Then
					txtOperatingValues(x).Enabled = True
					txtOperatingValues(x + 4).Enabled = True
				Else
					txtOperatingValues(x + 4).Enabled = True
				End If
			Next x
		End If
		
		If WhichCell = 10 And LTrim(RTrim(VB6.GetItemString(lstOptions(Index), lstOptions(Index).SelectedIndex))) = "Scraper" Then
			txtOperatingValues(11).Enabled = False
			LabOperatingTitles(9).Enabled = False
		End If
		If WhichCell = 11 And LTrim(RTrim(VB6.GetItemString(lstOptions(Index), lstOptions(Index).SelectedIndex))) = "Crusher/Conveyor" Then
			CellValues(WhichScreen, WhichCell - 3, WhichSegment).Changed = False
			txtOperatingValues(10).Text = "Front-End Loader"
			Call Inputer(WhichCell - 1)
			CellValues(WhichScreen, WhichCell - 3, WhichSegment).Changed = True
			txtOperatingValues(11).Text = "Crusher/Conveyor"
		End If
		If WhichCell = 15 And LTrim(RTrim(VB6.GetItemString(lstOptions(Index), lstOptions(Index).SelectedIndex))) = "Scraper" Then
			txtOperatingValues(16).Enabled = False
			LabOperatingTitles(14).Enabled = False
		End If
		If WhichCell = 16 And LTrim(RTrim(VB6.GetItemString(lstOptions(Index), lstOptions(Index).SelectedIndex))) = "Crusher/Conveyor" Then
			CellValues(WhichScreen, WhichCell - 3, WhichSegment).Changed = False
			txtOperatingValues(15).Text = "Front-End Loader"
			Call Inputer(WhichCell - 1)
			CellValues(WhichScreen, WhichCell - 3, WhichSegment).Changed = True
			txtOperatingValues(16).Text = "Crusher/Conveyor"
		End If
		If WhichCell = 15 And LTrim(RTrim(VB6.GetItemString(lstOptions(Index), lstOptions(Index).SelectedIndex))) = "Walking Dragline" Then
			txtOperatingValues(16).Enabled = False
			LabOperatingTitles(14).Enabled = False
		End If
		
		Call Inputer(WhichCell)
		
		If WhichCell = 8 And LTrim(RTrim(LCase(CellValues(Production, 40, 0).Word))) = "sherpa" Then
			txtOperatingValues(12).Focus()
		ElseIf WhichCell = 10 And LTrim(RTrim(VB6.GetItemString(lstOptions(Index), lstOptions(Index).SelectedIndex))) = "Scraper" Then 
			txtOperatingValues(12).Focus()
		ElseIf WhichCell = 15 And LTrim(RTrim(VB6.GetItemString(lstOptions(Index), lstOptions(Index).SelectedIndex))) = "Scraper" Then 
			txtOperatingValues(0).Focus()
		ElseIf WhichCell = 15 And LTrim(RTrim(VB6.GetItemString(lstOptions(Index), lstOptions(Index).SelectedIndex))) = "Walking Dragline" Then 
			txtOperatingValues(0).Focus()
		ElseIf WhichCell = 16 Then 
			txtOperatingValues(0).Focus()
		Else
			txtOperatingValues(WhichCell + 1).Focus()
		End If
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
			Call ScreenCalcs()
			labSegment(WhichSegment).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF)
			txtSegmentLabel.Text = SegNamie(WhichSegment)
			If WhichSegment > 0 Then
				txtOperatingValues(1).Focus()
			Else
				txtOperatingValues(1).Focus()
			End If
		End If
	End Sub
	'UPGRADE_WARNING: Event txtOperatingValues.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtOperatingValues_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtOperatingValues.TextChanged
		Dim Index As Short = txtOperatingValues.GetIndex(eventSender)
		If DoNotChange = True Then Exit Sub
	End Sub
	Private Sub txtOperatingValues_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtOperatingValues.Enter
		Dim Index As Short = txtOperatingValues.GetIndex(eventSender)
		Dim x As Short
		WhichCell = Index
		For x = 0 To 3
			labOperatingLabels(x + 9).Visible = False
			lstOptions(x).Visible = False
		Next x
		
		Select Case Index
			Case 8, 13
				labOperatingLabels(9).Visible = True
				lstOptions(0).Visible = True
			Case 9, 14
				labOperatingLabels(10).Visible = True
				lstOptions(1).Visible = True
			Case 10, 15
				labOperatingLabels(11).Visible = True
				lstOptions(2).Visible = True
			Case 11, 16
				labOperatingLabels(12).Visible = True
				lstOptions(3).Visible = True
		End Select
		
		System.Windows.Forms.SendKeys.Send("{HOME}+{END}")
		
		WhichCell = Index
		Call drawthevalues()
		
	End Sub
	
	Public Sub screenstuff()
		
		Dim r As Decimal
		Dim s As Decimal
		
		Dim a As Short
		Dim b As Single
		Dim c As Single
		Dim x As Short
		Dim u As Single
		Dim v As Single
		Dim y As Decimal
		Dim z As Single
		
		Dim h As Single
		Dim w As Single
		
		h = 6420
		w = 9150
		
		a = (2040 / h) * TempHigh
		b = (750 / h) * TempHigh
		c = (870 / h) * TempHigh
		
		r = (360 / h) * TempHigh
		u = (300 / w) * TempWide
		v = (3360 / h) * TempHigh
		
		
		y = (60 / h) * TempHigh
		z = (420 / w) * TempWide
		
		labOperatingHeading.Top = VB6.TwipsToPixelsY(TempHigh * (180 / h))
		labOperatingHeading.Left = VB6.TwipsToPixelsX(TempWide * (180 / w))
		
		labOperatingLabels(0).Top = VB6.TwipsToPixelsY(TempHigh * (720 / h))
		labOperatingLabels(0).Left = VB6.TwipsToPixelsX(TempWide * (300 / w))
		
		labOperatingLabels(1).Top = VB6.TwipsToPixelsY(TempHigh * (3780 / h))
		labOperatingLabels(1).Left = VB6.TwipsToPixelsX(TempWide * (120 / w))
		
		labOperatingLabels(2).Top = VB6.TwipsToPixelsY(TempHigh * (120 / h))
		labOperatingLabels(2).Left = VB6.TwipsToPixelsX(TempWide * (3720 / w))
		
		For x = 0 To 1
			labOperatingLabels(x + 3).Top = VB6.TwipsToPixelsY((TempHigh * (2160 / h)) + (x * a))
			labOperatingLabels(x + 3).Left = VB6.TwipsToPixelsX(TempWide * (4200 / w))
		Next x
		
		For x = 0 To 1
			labOperatingLabels(x + 5).Top = VB6.TwipsToPixelsY((TempHigh * (1110 / h)) + (x * b))
			labOperatingLabels(x + 5).Left = VB6.TwipsToPixelsX(TempWide * (720 / w))
			labOperatingLabels(x + 5).Width = VB6.TwipsToPixelsX(TempWide * (2715 / w))
		Next x
		
		labOperatingLabels(7).Top = VB6.TwipsToPixelsY(TempHigh * (2760 / h))
		labOperatingLabels(7).Left = VB6.TwipsToPixelsX(TempWide * (180 / w))
		labOperatingLabels(7).Width = VB6.TwipsToPixelsX(TempWide * (1815 / w))
		
		labOperatingLabels(8).Top = VB6.TwipsToPixelsY(TempHigh * (2820 / h))
		labOperatingLabels(8).Left = VB6.TwipsToPixelsX(TempWide * (2160 / w))
		labOperatingLabels(8).Width = VB6.TwipsToPixelsX(TempWide * (1965 / w))
		
		For x = 0 To 5
			optSegment(x).Top = VB6.TwipsToPixelsY(TempHigh * (3060 / h))
			optSegment(x).Left = VB6.TwipsToPixelsX((TempWide * (240 / w)) + (x * u))
			optSegment(x).Width = VB6.TwipsToPixelsX(TempWide * (195 / w))
			labSegment(x).Top = VB6.TwipsToPixelsY(TempHigh * (3300 / h))
			labSegment(x).Left = VB6.TwipsToPixelsX((TempWide * (240 / w)) + (x * u))
			labSegment(x).Width = VB6.TwipsToPixelsX(TempWide * (195 / w))
		Next x
		
		txtSegmentLabel.Top = VB6.TwipsToPixelsY(TempHigh * (3120 / h))
		txtSegmentLabel.Left = VB6.TwipsToPixelsX(TempWide * (2160 / w))
		txtSegmentLabel.Width = VB6.TwipsToPixelsX(TempWide * (1935 / w))
		
		For x = 0 To 3
			lstOptions(x).Top = VB6.TwipsToPixelsY((TempHigh * (4320 / h)) + (x * y))
			lstOptions(x).Height = VB6.TwipsToPixelsY((TempHigh * (1410 / h)))
			lstOptions(x).Left = VB6.TwipsToPixelsX((TempWide * (540 / w)) + (x * z))
			lstOptions(x).Width = VB6.TwipsToPixelsX((TempWide * (2055 / w)))
			labOperatingLabels(x + 9).Top = VB6.TwipsToPixelsY((TempHigh * (4080 / h)) + (x * y))
			labOperatingLabels(x + 9).Left = VB6.TwipsToPixelsX((TempWide * (540 / w)) + (x * z))
			labOperatingLabels(x + 9).Width = VB6.TwipsToPixelsX(TempWide * (2055 / w))
		Next x
		
		For x = 0 To 14
			If x < 5 Then
				s = 0
			ElseIf x < 10 Then 
				s = (240 / h) * TempHigh
			Else
				s = (480 / h) * TempHigh
			End If
			If x < 5 Then
				LabOperatingTitles(x).Left = VB6.TwipsToPixelsX(TempWide * (3900 / w))
			Else
				LabOperatingTitles(x).Left = VB6.TwipsToPixelsX(TempWide * (4380 / w))
			End If
			LabOperatingTitles(x).Top = VB6.TwipsToPixelsY((TempHigh * (420 / h)) + (x * r) + s)
			txtOperatingValues(x + 2).Top = VB6.TwipsToPixelsY((TempHigh * (380 / h)) + (x * r) + s)
			labOperatingUnits(x).Top = VB6.TwipsToPixelsY((TempHigh * (420 / h)) + (x * r) + s)
			Select Case x
				Case 0 To 4
					LabOperatingTitles(x).Width = VB6.TwipsToPixelsX(TempWide * (1995 / w))
					txtOperatingValues(x + 2).Width = VB6.TwipsToPixelsX(TempWide * (1215 / w))
					txtOperatingValues(x + 2).Left = VB6.TwipsToPixelsX(TempWide * (6060 / w))
					labOperatingUnits(x).Left = VB6.TwipsToPixelsX(TempWide * (7320 / w))
				Case 5, 10
					LabOperatingTitles(x).Width = VB6.TwipsToPixelsX(TempWide * (1515 / w))
					txtOperatingValues(x + 2).Width = VB6.TwipsToPixelsX(TempWide * (1215 / w))
					txtOperatingValues(x + 2).Left = VB6.TwipsToPixelsX(TempWide * (6060 / w))
					labOperatingUnits(x).Left = VB6.TwipsToPixelsX(TempWide * (7320 / w))
				Case Else
					LabOperatingTitles(x).Width = VB6.TwipsToPixelsX(TempWide * (1155 / w))
					txtOperatingValues(x + 2).Width = VB6.TwipsToPixelsX(TempWide * (1935 / w))
					txtOperatingValues(x + 2).Left = VB6.TwipsToPixelsX(TempWide * (5700 / w))
					labOperatingUnits(x).Left = VB6.TwipsToPixelsX(TempWide * (7680 / w))
			End Select
		Next x
		
		LabOperatingTitles(15).Top = VB6.TwipsToPixelsY(TempHigh * (1440 / h))
		LabOperatingTitles(15).Left = VB6.TwipsToPixelsX(TempWide * (540 / w))
		LabOperatingTitles(15).Width = VB6.TwipsToPixelsX(TempWide * (975 / w))
		
		txtOperatingValues(17).Top = VB6.TwipsToPixelsY((TempHigh * (2190 / h)))
		txtOperatingValues(17).Left = VB6.TwipsToPixelsX(TempWide * (1200 / w))
		txtOperatingValues(17).Width = VB6.TwipsToPixelsX(TempWide * (1215 / w))
		
		txtOperatingValues(0).Top = VB6.TwipsToPixelsY((TempHigh * (1380 / h)))
		txtOperatingValues(0).Left = VB6.TwipsToPixelsX(TempWide * (1680 / w))
		txtOperatingValues(0).Width = VB6.TwipsToPixelsX(TempWide * (615 / w))
		
		txtOperatingValues(1).Top = VB6.TwipsToPixelsY((TempHigh * (1380 / h)))
		txtOperatingValues(1).Left = VB6.TwipsToPixelsX(TempWide * (2820 / w))
		txtOperatingValues(1).Width = VB6.TwipsToPixelsX(TempWide * (615 / w))
		
		labOperatingUnits(15).Top = VB6.TwipsToPixelsY((TempHigh * (2190 / h)))
		labOperatingUnits(15).Left = VB6.TwipsToPixelsX(TempWide * (2460 / w))
		
		labOperatingUnits(16).Top = VB6.TwipsToPixelsY((TempHigh * (1440 / h)))
		labOperatingUnits(16).Left = VB6.TwipsToPixelsX(TempWide * (2460 / w))
		labOperatingUnits(16).Width = VB6.TwipsToPixelsX(TempWide * (210 / w))
		
		LineHorizontal(0).X1 = VB6.TwipsToPixelsX(TempWide * (2580 / w))
		LineHorizontal(0).X2 = VB6.TwipsToPixelsX(TempWide * (3720 / w))
		LineHorizontal(0).Y1 = VB6.TwipsToPixelsY(TempHigh * (780 / h))
		LineHorizontal(0).Y2 = VB6.TwipsToPixelsY(TempHigh * (780 / h))
		
		LineHorizontal(1).X1 = VB6.TwipsToPixelsX(TempWide * (300 / w))
		LineHorizontal(1).X2 = VB6.TwipsToPixelsX(TempWide * (4320 / w))
		LineHorizontal(1).Y1 = VB6.TwipsToPixelsY(TempHigh * (2520 / h))
		LineHorizontal(1).Y2 = VB6.TwipsToPixelsY(TempHigh * (2520 / h))
		
		LineHorizontal(2).X1 = VB6.TwipsToPixelsX(TempWide * (1440 / w))
		LineHorizontal(2).X2 = VB6.TwipsToPixelsX(TempWide * (4200 / w))
		LineHorizontal(2).Y1 = VB6.TwipsToPixelsY(TempHigh * (3840 / h))
		LineHorizontal(2).Y2 = VB6.TwipsToPixelsY(TempHigh * (3840 / h))
		
		LineHorizontal(3).X1 = VB6.TwipsToPixelsX(TempWide * (120 / w))
		LineHorizontal(3).X2 = VB6.TwipsToPixelsX(TempWide * (4200 / w))
		LineHorizontal(3).Y1 = VB6.TwipsToPixelsY(TempHigh * (6000 / h))
		LineHorizontal(3).Y2 = VB6.TwipsToPixelsY(TempHigh * (6000 / h))
		
		LineHorizontal(4).X1 = VB6.TwipsToPixelsX(TempWide * (4500 / w))
		LineHorizontal(4).X2 = VB6.TwipsToPixelsX(TempWide * (9060 / w))
		LineHorizontal(4).Y1 = VB6.TwipsToPixelsY(TempHigh * (180 / h))
		LineHorizontal(4).Y2 = VB6.TwipsToPixelsY(TempHigh * (180 / h))
		
		LineHorizontal(5).X1 = VB6.TwipsToPixelsX(TempWide * (3720 / w))
		LineHorizontal(5).X2 = VB6.TwipsToPixelsX(TempWide * (4140 / w))
		LineHorizontal(5).Y1 = VB6.TwipsToPixelsY(TempHigh * (2220 / h))
		LineHorizontal(5).Y2 = VB6.TwipsToPixelsY(TempHigh * (2220 / h))
		
		LineHorizontal(6).X1 = VB6.TwipsToPixelsX(TempWide * (5640 / w))
		LineHorizontal(6).X2 = VB6.TwipsToPixelsX(TempWide * (8940 / w))
		LineHorizontal(6).Y1 = VB6.TwipsToPixelsY(TempHigh * (2220 / h))
		LineHorizontal(6).Y2 = VB6.TwipsToPixelsY(TempHigh * (2220 / h))
		
		LineHorizontal(7).X1 = VB6.TwipsToPixelsX(TempWide * (5880 / w))
		LineHorizontal(7).X2 = VB6.TwipsToPixelsX(TempWide * (8940 / w))
		LineHorizontal(7).Y1 = VB6.TwipsToPixelsY(TempHigh * (4260 / h))
		LineHorizontal(7).Y2 = VB6.TwipsToPixelsY(TempHigh * (4260 / h))
		
		LineHorizontal(8).X1 = VB6.TwipsToPixelsX(TempWide * (4200 / w))
		LineHorizontal(8).X2 = VB6.TwipsToPixelsX(TempWide * (9060 / w))
		LineHorizontal(8).Y1 = VB6.TwipsToPixelsY(TempHigh * (6300 / h))
		LineHorizontal(8).Y2 = VB6.TwipsToPixelsY(TempHigh * (6300 / h))
		
		LineVertical(0).X1 = VB6.TwipsToPixelsX(TempWide * (360 / w))
		LineVertical(0).X2 = VB6.TwipsToPixelsX(TempWide * (360 / w))
		LineVertical(0).Y1 = VB6.TwipsToPixelsY(TempHigh * (1020 / h))
		LineVertical(0).Y2 = VB6.TwipsToPixelsY(TempHigh * (2580 / h))
		
		LineVertical(1).X1 = VB6.TwipsToPixelsX(TempWide * (180 / w))
		LineVertical(1).X2 = VB6.TwipsToPixelsX(TempWide * (180 / w))
		LineVertical(1).Y1 = VB6.TwipsToPixelsY(TempHigh * (4080 / h))
		LineVertical(1).Y2 = VB6.TwipsToPixelsY(TempHigh * (6060 / h))
		
		LineVertical(2).X1 = VB6.TwipsToPixelsX(TempWide * (3780 / w))
		LineVertical(2).X2 = VB6.TwipsToPixelsX(TempWide * (3780 / w))
		LineVertical(2).Y1 = VB6.TwipsToPixelsY(TempHigh * (420 / h))
		LineVertical(2).Y2 = VB6.TwipsToPixelsY(TempHigh * (2580 / h))
		
		LineVertical(3).X1 = VB6.TwipsToPixelsX(TempWide * (4260 / w))
		LineVertical(3).X2 = VB6.TwipsToPixelsX(TempWide * (4260 / w))
		LineVertical(3).Y1 = VB6.TwipsToPixelsY(TempHigh * (2460 / h))
		LineVertical(3).Y2 = VB6.TwipsToPixelsY(TempHigh * (4140 / h))
		
		LineVertical(4).X1 = VB6.TwipsToPixelsX(TempWide * (4260 / w))
		LineVertical(4).X2 = VB6.TwipsToPixelsX(TempWide * (4260 / w))
		LineVertical(4).Y1 = VB6.TwipsToPixelsY(TempHigh * (4500 / h))
		LineVertical(4).Y2 = VB6.TwipsToPixelsY(TempHigh * (6360 / h))
		
		LineVertical(5).X1 = VB6.TwipsToPixelsX(TempWide * (9000 / w))
		LineVertical(5).X2 = VB6.TwipsToPixelsX(TempWide * (9000 / w))
		LineVertical(5).Y1 = VB6.TwipsToPixelsY(TempHigh * (120 / h))
		LineVertical(5).Y2 = VB6.TwipsToPixelsY(TempHigh * (6360 / h))
		
		labBackToMenu.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labBackToMenu.Left = VB6.TwipsToPixelsX(TempWide * (600 / w))
		
		imgBackToMenu.Top = VB6.TwipsToPixelsY(TempHigh * (6150 / h))
		imgBackToMenu.Left = VB6.TwipsToPixelsX(TempWide * (60 / w))
		imgBackToMenu.Width = VB6.TwipsToPixelsX(TempWide * (495 / w))
		
		comOperatingHelp.Top = VB6.TwipsToPixelsY(TempHigh * (6150 / h))
		comOperatingHelp.Left = VB6.TwipsToPixelsX(TempWide * (3900 / w))
		
		labOperatingHelp.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labOperatingHelp.Left = VB6.TwipsToPixelsX(TempWide * (3360 / w))
		
		labInsert.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labInsert.Left = VB6.TwipsToPixelsX(TempWide * (1800 / w))
		labInsert.Width = VB6.TwipsToPixelsX(TempWide * (975 / w))
		
	End Sub
	Public Sub drawthevalues()
		Dim i As Short
		Dim x As Short
		
		DoNotChange = True
		
		If WhichSegment > 0 Then
			txtOperatingValues(6).Enabled = False
		Else
			txtOperatingValues(6).Enabled = True
		End If
		
		If WhichSegment > 0 Then
			CellValues(Production, 6, x).Word = CellValues(Production, 6, 0).Word
			CellValues(Production, 7, x).Word = CellValues(Production, 7, 0).Word
			CellValues(Production, 8, x).Word = CellValues(Production, 8, 0).Word
			CellValues(Production, 9, x).Word = CellValues(Production, 9, 0).Word
			CellValues(Production, 11, x).Word = CellValues(Production, 11, 0).Word
			CellValues(Production, 12, x).Word = CellValues(Production, 12, 0).Word
			CellValues(Production, 13, x).Word = CellValues(Production, 13, 0).Word
			CellValues(Production, 14, x).Word = CellValues(Production, 14, 0).Word
		End If
		
		For i = 0 To 17
			Select Case i
				Case 0 To 14
					If CellValues(WhichScreen, i, WhichSegment).Changed = True Then
						txtOperatingValues(i + 2).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF0000)
					Else
						txtOperatingValues(i + 2).ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
					End If
				Case 15, 16
					If CellValues(WhichScreen, i, WhichSegment).Changed = True Then
						txtOperatingValues(i - 15).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF0000)
					Else
						txtOperatingValues(i - 15).ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
					End If
			End Select
			Select Case i
				Case 0, 4, 5, 10
					txtOperatingValues(i + 2).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value * TonConversion)), "#,###,###,##0")
				Case 1
					txtOperatingValues(i + 2).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value)), "##,###,##0.0")
				Case 2, 3
					txtOperatingValues(i + 2).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value)), "#,###,###,##0")
				Case 6 To 9, 11 To 14
					txtOperatingValues(i + 2).Text = LTrim(RTrim(CellValues(WhichScreen, i, WhichSegment).Word))
				Case 15, 16
					txtOperatingValues(i - 15).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value)), "#,###,###,##0")
				Case 17
					txtOperatingValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value * TonConversion)), "#,###,###,##0")
			End Select
		Next i
		
		DoNotChange = False
		
	End Sub
	Private Sub txtOperatingValues_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtOperatingValues.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Dim Index As Short = txtOperatingValues.GetIndex(eventSender)
		
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
						If InStr(txtOperatingValues(Index).Text, ".") = 0 Then
							System.Windows.Forms.SendKeys.Send("{DELETE}")
						End If
					Else
						System.Windows.Forms.SendKeys.Send("{DELETE}")
					End If
			End Select
		End If
		
	End Sub
	Private Sub txtOperatingValues_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtOperatingValues.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = txtOperatingValues.GetIndex(eventSender)
		If KeyAscii > Asc("9") And KeyAscii <> Asc(",") And KeyAscii <> Asc(".") And KeyAscii <> Asc("$") Then
			Beep()
			KeyAscii = 0
		Else
			If WhichCell > 1 Then
				CellValues(WhichScreen, WhichCell - 2, WhichSegment).Changed = True
			Else
				CellValues(WhichScreen, WhichCell + 15, WhichSegment).Changed = True
			End If
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	Private Sub txtOperatingValues_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtOperatingValues.Leave
		Dim Index As Short = txtOperatingValues.GetIndex(eventSender)
		If DoNotChange = True Then Exit Sub
		WhichCell = Index
		If WhichCell = 12 And LTrim(RTrim(LCase(CellValues(Production, 6, WhichSegment).Word))) = "sherpa" Then
			txtOperatingValues(0).Focus()
		End If
		Call Inputer(WhichCell)
		Call ScreenCalcs()
	End Sub
	Private Sub Inputer(ByRef Sample As Short)
		Dim x As Short
		Dim i As Short
		Dim life As Decimal
		Dim tempvalue As String
		Dim Digit As New VB6.FixedLengthString(1)
		
		On Error Resume Next
		
		If DoNotChange = True Then Exit Sub
		PageChange(WhichScreen) = True
		Select Case Sample
			Case 0 To 7, 12
				tempvalue = ""
				For i = 1 To Len(txtOperatingValues(Sample).Text)
					Digit.Value = Mid(txtOperatingValues(Sample).Text, i, 1)
					Select Case Digit.Value
						Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ".", "-"
							tempvalue = tempvalue & Digit.Value
					End Select
				Next i
				If Sample = 0 Or Sample = 1 Then
					If CellValues(WhichScreen, Sample + 15, WhichSegment).Changed = True Then
						CellValues(WhichScreen, Sample + 15, WhichSegment).Value = Val(tempvalue)
					End If
				ElseIf Sample = 2 Or Sample = 6 Or Sample = 7 Or Sample = 12 Then 
					If CellValues(WhichScreen, Sample - 2, WhichSegment).Changed = True Then
						If TonConversion <> 0 Then CellValues(WhichScreen, Sample - 2, WhichSegment).Value = Val(tempvalue) / TonConversion
					End If
				Else
					If CellValues(WhichScreen, Sample - 2, WhichSegment).Changed = True Then
						CellValues(WhichScreen, Sample - 2, WhichSegment).Value = Val(tempvalue)
					End If
				End If
			Case Else
				CellValues(WhichScreen, Sample - 2, WhichSegment).Word = txtOperatingValues(Sample).Text
		End Select
		Call ScreenCalcs()
		Call drawthevalues()
	End Sub
	Public Sub ScreenCalcs()
		Dim x As Short
		Dim z As Short
		Dim tempvalue As Decimal
		Dim TempWord As New VB6.FixedLengthString(16)
		
		On Error Resume Next
		
		DoNotChange = True
		
		Select Case WhichCell
			Case 0 To 7
				z = 0
				If CellValues(Production, 3, WhichSegment).Value = 0 Then z = 1
				If CellValues(Production, 5, WhichSegment).Value = 0 Then z = 1
				If CellValues(Production, 15, WhichSegment).Value = 0 Then z = 1
				If CellValues(Production, 16, WhichSegment).Value = 0 Then z = 1
				If z = 0 Then
					MineLife = ((CellValues(Production, 0, WhichSegment).Value) / CellValues(Production, 5, WhichSegment).Value)
					CellValues(Production, 17, WhichSegment).Value = (CellValues(Production, 3, WhichSegment).Value * CellValues(Production, 5, WhichSegment).Value) * ((CellValues(Production, 16, WhichSegment).Value - CellValues(Production, 15, WhichSegment).Value) + 1)
					If CellValues(Production, 17, WhichSegment).Value > 0 Then
						If CellValues(Production, 0, WhichSegment + 1).Changed = False Then
							CellValues(Production, 0, WhichSegment + 1).Value = CellValues(Production, 0, WhichSegment).Value - CellValues(Production, 17, WhichSegment).Value
						End If
						If CellValues(Production, 15, WhichSegment + 1).Changed = False Then
							CellValues(Production, 15, WhichSegment + 1).Value = CellValues(Production, 16, WhichSegment).Value + 1
						End If
					End If
				End If
		End Select
		
		Call TimeLineCalc()
		
		For x = 1 To MaxSegment
			CellValues(Production, 6, x).Word = CellValues(Production, 6, 0).Word
			CellValues(Production, 6, x).Changed = True
			CellValues(Production, 7, x).Word = CellValues(Production, 7, 0).Word
			CellValues(Production, 7, x).Changed = True
			CellValues(Production, 8, x).Word = CellValues(Production, 8, 0).Word
			CellValues(Production, 8, x).Changed = True
			CellValues(Production, 9, x).Word = CellValues(Production, 9, 0).Word
			CellValues(Production, 9, x).Changed = True
			CellValues(Production, 11, x).Word = CellValues(Production, 11, 0).Word
			CellValues(Production, 11, x).Changed = True
			CellValues(Production, 12, x).Word = CellValues(Production, 12, 0).Word
			CellValues(Production, 12, x).Changed = True
			CellValues(Production, 13, x).Word = CellValues(Production, 13, 0).Word
			CellValues(Production, 13, x).Changed = True
			CellValues(Production, 14, x).Word = CellValues(Production, 14, 0).Word
			CellValues(Production, 14, x).Changed = True
		Next x
		
		If WhichSegment > 0 Then
			txtOperatingValues(8).Enabled = False
			LabOperatingTitles(6).Enabled = False
			For x = 9 To 12
				If x < 12 Then
					txtOperatingValues(x).Enabled = False
					LabOperatingTitles(x - 2).Enabled = False
					txtOperatingValues(x + 4).Enabled = False
					LabOperatingTitles(x + 2).Enabled = False
				Else
					txtOperatingValues(x + 4).Enabled = False
					LabOperatingTitles(x + 2).Enabled = False
				End If
			Next x
		End If
		
		If LTrim(RTrim(CellValues(Production, 7, 0).Word)) = "Dynamite" Then bltp = 1
		If LTrim(RTrim(CellValues(Production, 12, 0).Word)) = "Dynamite" Then bltp = 1
		
		DoNotChange = False
		
	End Sub
	
	'UPGRADE_WARNING: Event txtSegmentLabel.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtSegmentLabel_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSegmentLabel.TextChanged
		If DoNotChange = True Then Exit Sub
		SegNamie(WhichSegment) = txtSegmentLabel.Text
	End Sub

End Class