Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks
Friend Class frmElectricalData
	Inherits System.Windows.Forms.Form
	Dim TempHigh As Single
	Dim TempWide As Single
	Dim FootConversion As Single
	Private Sub comElectricalPrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles comElectricalPrint.Click
		PrintForm1.Print(Me, PowerPacks.Printing.PrintForm.PrintOption.CompatibleModeClientAreaOnly)
	End Sub
	'UPGRADE_WARNING: Form event frmElectricalData.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmElectricalData_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated

        'Dim baseunit As String
        'Dim baselength As Short
        Dim i As Short
		
		If IsHelpOn = True Then
			txtElectricalValues(WhichCell).Focus()
			IsHelpOn = False
		Else
			For i = 0 To 5
				If optSegment(i).Checked = True Then WhichSegment = i
			Next i
			
			WhichScreen = Electrical
			
			Call drawthevalues()
			
			If InsertFlag = True Then
				labInsert.Text = "Insert"
			Else
				labInsert.Text = "Typeover"
			End If
			
			WhichCell = 0
			
			txtElectricalValues(0).Focus()
		End If
		
	End Sub
	Private Sub frmElectricalData_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        'Dim x As Short

        Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - (VB6.PixelsToTwipsY(Me.Height) + 350)) / 2)
		Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2)
		
		If VB6.PixelsToTwipsY(Me.Top) < 0 Then Me.Top = 0
		If VB6.PixelsToTwipsX(Me.Left) < 0 Then Me.Left = 0
		
		TempWide = VB6.PixelsToTwipsX(Me.ClientRectangle.Width)
		TempHigh = VB6.PixelsToTwipsY(Me.ClientRectangle.Height)
		
		DoNotChange = True
		
		If UnitType = Metric Then
			labElectricalUnits(9).Text = "meters"
			labElectricalUnits(10).Text = "/meter"
			FootConversion = 0.3048
		Else
			labElectricalUnits(9).Text = "feet"
			labElectricalUnits(10).Text = "/foot"
			FootConversion = 1
		End If
		
		WhichSegment = 0
		optSegment(WhichSegment).Checked = True
		txtSegmentLabel.Text = SegNamie(WhichSegment)
		
		If PageChange(WhichScreen) = True Then
			Call drawthevalues()
		End If
		
		DoNotChange = False
		Call screenstuff()
	End Sub
	'UPGRADE_WARNING: Event frmElectricalData.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub frmElectricalData_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		TempWide = VB6.PixelsToTwipsX(Me.ClientRectangle.Width)
		TempHigh = VB6.PixelsToTwipsY(Me.ClientRectangle.Height)
		Call screenstuff()
	End Sub
	Private Sub frmElectricalData_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		On Error Resume Next
        'Me.Close()
		Call InputMenuAccess(3)
	End Sub
	Private Sub imgBackToMenu_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles imgBackToMenu.Click
		Me.Close()
		Call InputMenuAccess(3)
	End Sub
	Private Sub labBackToMenu_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles labBackToMenu.Click
		Me.Close()
		Call InputMenuAccess(3)
	End Sub
	Private Sub labElectricalHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles labElectricalHelp.Click
		Dim StartHelp As Short
		StartHelp = 235
		IsHelpOn = True
		Call frmSurfaceHelp.gethelptext(StartHelp, WhichCell)
		frmSurfaceHelp.Show()
	End Sub
	'UPGRADE_WARNING: Event lstElectricalList.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub lstElectricalList_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstElectricalList.SelectedIndexChanged
		Dim Index As Short = lstElectricalList.GetIndex(eventSender)
        'Dim x As Short

        Select Case Index
			Case 2
				Select Case LTrim(RTrim(LCase(VB6.GetItemString(lstElectricalList(Index), lstElectricalList(Index).SelectedIndex))))
					Case "2 kilovolt"
						txtElectricalValues(WhichCell).Text = CStr(2)
					Case "8 kilovolt"
						txtElectricalValues(WhichCell).Text = CStr(8)
					Case "15 kilovolt"
						txtElectricalValues(WhichCell).Text = CStr(15)
				End Select
			Case 0
				Select Case LTrim(RTrim(LCase(VB6.GetItemString(lstElectricalList(Index), lstElectricalList(Index).SelectedIndex))))
					Case "150 kilovolt-ampere"
						txtElectricalValues(WhichCell).Text = CStr(150)
					Case "300 kilovolt-ampere"
						txtElectricalValues(WhichCell).Text = CStr(300)
					Case "500 kilovolt-ampere"
						txtElectricalValues(WhichCell).Text = CStr(500)
					Case "750 kilovolt-ampere"
						txtElectricalValues(WhichCell).Text = CStr(750)
					Case "1,000 kilovolt-ampere"
						txtElectricalValues(WhichCell).Text = CStr(1000)
					Case "1,500 kilovolt-ampere"
						txtElectricalValues(WhichCell).Text = CStr(1500)
					Case "5,000 kilovolt-ampere"
						txtElectricalValues(WhichCell).Text = CStr(5000)
					Case "10,000 kilovolt-ampere"
						txtElectricalValues(WhichCell).Text = CStr(10000)
				End Select
		End Select
		
		CellValues(WhichScreen, WhichCell, WhichSegment).Changed = True
		
		Call Inputer(WhichCell)
		
		txtElectricalValues(WhichCell + 1).Focus()
		
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
			txtElectricalValues(WhichCell).Focus()
		End If
	End Sub
	Private Sub txtElectricalValues_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtElectricalValues.Enter
		Dim Index As Short = txtElectricalValues.GetIndex(eventSender)
		Dim x As Short
		WhichCell = Index
		
		For x = 0 To 2
			lstElectricalList(x).Visible = False
		Next x
		Select Case WhichCell
			Case 2
				lstElectricalList(0).Visible = True
			Case 8
				lstElectricalList(2).Visible = True
			Case 5
				lstElectricalList(1).Visible = True
		End Select
		
		System.Windows.Forms.SendKeys.Send("{HOME}+{END}")
		
		WhichCell = Index
		Call drawthevalues()
		
	End Sub
	'UPGRADE_WARNING: Event txtElectricalValues.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtElectricalValues_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtElectricalValues.TextChanged
		Dim Index As Short = txtElectricalValues.GetIndex(eventSender)
		If DoNotChange = True Then Exit Sub
	End Sub
	Public Sub screenstuff()
		
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
		
		q = (1080 / h) * TempHigh
		r = (1140 / h) * TempHigh
		s = (60 / h) * TempHigh
		t = (120 / w) * TempWide
		u = (540 / h) * TempHigh
		v = (300 / w) * TempWide
		
		y = (420 / h) * TempHigh
		z = (300 / w) * TempWide
		
		For x = 0 To 1
			labElectricalHeading(x).Top = VB6.TwipsToPixelsY((TempHigh * (180 / h)) + (x * u))
			labElectricalHeading(x).Left = VB6.TwipsToPixelsX((TempWide * (180 / w)) + (x * v))
			labElectricalHeading(x).Width = VB6.TwipsToPixelsX(TempWide * (1965 / w))
			labElectricalLabels(x).Top = VB6.TwipsToPixelsY((TempHigh * (1620 / h)) + (x * r))
			labElectricalLabels(x).Left = VB6.TwipsToPixelsX(TempWide * (300 / w))
			labElectricalLabels(x).Width = VB6.TwipsToPixelsX(TempWide * (1935 / w))
		Next x
		
		labElectricalLabels(2).Left = VB6.TwipsToPixelsX(TempWide * (180 / w))
		labElectricalLabels(2).Top = VB6.TwipsToPixelsY(TempHigh * (3840 / h))
		
		For x = 3 To 7
			labElectricalLabels(x).Left = VB6.TwipsToPixelsX(TempWide * (2580 / w))
			labElectricalLabels(x).Top = VB6.TwipsToPixelsY((TempHigh * (120 / h)) + ((x - 3) * q))
		Next x
		
		For x = 0 To 5
			optSegment(x).Top = VB6.TwipsToPixelsY(TempHigh * (1920 / h))
			optSegment(x).Left = VB6.TwipsToPixelsX((TempWide * (420 / w)) + (x * z))
			optSegment(x).Width = VB6.TwipsToPixelsX(TempWide * (195 / w))
			labSegment(x).Top = VB6.TwipsToPixelsY(TempHigh * (2160 / h))
			labSegment(x).Left = VB6.TwipsToPixelsX((TempWide * (420 / w)) + (x * z))
			labSegment(x).Width = VB6.TwipsToPixelsX(TempWide * (195 / w))
		Next x
		
		txtSegmentLabel.Top = VB6.TwipsToPixelsY(TempHigh * (3060 / h))
		txtSegmentLabel.Left = VB6.TwipsToPixelsX(TempWide * (300 / w))
		txtSegmentLabel.Width = VB6.TwipsToPixelsX(TempWide * (1935 / w))
		
		For x = 0 To 14
			Select Case x
				Case 0, 1, 2, 5, 8, 12
					LabElectricalTitles(x).Left = VB6.TwipsToPixelsX(TempWide * (2700 / w))
					LabElectricalTitles(x).Width = VB6.TwipsToPixelsX(TempWide * (2715 / w))
					txtElectricalValues(x).Left = VB6.TwipsToPixelsX(TempWide * (5580 / w))
					txtElectricalValues(x).Width = VB6.TwipsToPixelsX(TempWide * (1035 / w))
					labElectricalUnits(x).Left = VB6.TwipsToPixelsX(TempWide * (6660 / w))
				Case 3, 6, 13
					LabElectricalTitles(x).Left = VB6.TwipsToPixelsX(TempWide * (2700 / w))
					LabElectricalTitles(x).Width = VB6.TwipsToPixelsX(TempWide * (1995 / w))
					txtElectricalValues(x).Left = VB6.TwipsToPixelsX(TempWide * (4860 / w))
					txtElectricalValues(x).Width = VB6.TwipsToPixelsX(TempWide * (375 / w))
					labElectricalUnits(x).Left = VB6.TwipsToPixelsX(TempWide * (5280 / w))
				Case 4, 7, 14
					LabElectricalTitles(x).Left = VB6.TwipsToPixelsX(TempWide * (5520 / w))
					LabElectricalTitles(x).Width = VB6.TwipsToPixelsX(TempWide * (1815 / w))
					txtElectricalValues(x).Left = VB6.TwipsToPixelsX(TempWide * (7500 / w))
					txtElectricalValues(x).Width = VB6.TwipsToPixelsX(TempWide * (975 / w))
					labElectricalUnits(x).Left = VB6.TwipsToPixelsX(TempWide * (8520 / w))
				Case 9
					LabElectricalTitles(x).Left = VB6.TwipsToPixelsX(TempWide * (2700 / w))
					LabElectricalTitles(x).Width = VB6.TwipsToPixelsX(TempWide * (1815 / w))
					txtElectricalValues(x).Left = VB6.TwipsToPixelsX(TempWide * (4680 / w))
					txtElectricalValues(x).Width = VB6.TwipsToPixelsX(TempWide * (735 / w))
					labElectricalUnits(x).Left = VB6.TwipsToPixelsX(TempWide * (5460 / w))
				Case 10
					LabElectricalTitles(x).Left = VB6.TwipsToPixelsX(TempWide * (6240 / w))
					LabElectricalTitles(x).Width = VB6.TwipsToPixelsX(TempWide * (1095 / w))
					txtElectricalValues(x).Left = VB6.TwipsToPixelsX(TempWide * (7500 / w))
					txtElectricalValues(x).Width = VB6.TwipsToPixelsX(TempWide * (795 / w))
					labElectricalUnits(x).Left = VB6.TwipsToPixelsX(TempWide * (8340 / w))
				Case 11
					LabElectricalTitles(x).Left = VB6.TwipsToPixelsX(TempWide * (2700 / w))
					LabElectricalTitles(x).Width = VB6.TwipsToPixelsX(TempWide * (2295 / w))
					txtElectricalValues(x).Left = VB6.TwipsToPixelsX(TempWide * (5160 / w))
					txtElectricalValues(x).Width = VB6.TwipsToPixelsX(TempWide * (1875 / w))
					labElectricalUnits(x).Left = VB6.TwipsToPixelsX(TempWide * (7080 / w))
			End Select
		Next x
		
		For x = 0 To 1
			LabElectricalTitles(x).Top = VB6.TwipsToPixelsY((TempHigh * (420 / h)) + (x * y))
			txtElectricalValues(x).Top = VB6.TwipsToPixelsY(TempHigh * (390 / h) + (x * y))
			labElectricalUnits(x).Top = VB6.TwipsToPixelsY((TempHigh * (420 / h)) + (x * y))
		Next x
		
		LabElectricalTitles(2).Top = VB6.TwipsToPixelsY(TempHigh * (1500 / h))
		txtElectricalValues(2).Top = VB6.TwipsToPixelsY(TempHigh * (1470 / h))
		labElectricalUnits(2).Top = VB6.TwipsToPixelsY(TempHigh * (1500 / h))
		
		For x = 3 To 4
			LabElectricalTitles(x).Top = VB6.TwipsToPixelsY(TempHigh * (1980 / h))
			txtElectricalValues(x).Top = VB6.TwipsToPixelsY(TempHigh * (1950 / h))
			labElectricalUnits(x).Top = VB6.TwipsToPixelsY(TempHigh * (1980 / h))
		Next x
		
		LabElectricalTitles(5).Top = VB6.TwipsToPixelsY(TempHigh * (2580 / h))
		txtElectricalValues(5).Top = VB6.TwipsToPixelsY(TempHigh * (2550 / h))
		labElectricalUnits(5).Top = VB6.TwipsToPixelsY(TempHigh * (2580 / h))
		
		For x = 6 To 7
			LabElectricalTitles(x).Top = VB6.TwipsToPixelsY(TempHigh * (3060 / h))
			txtElectricalValues(x).Top = VB6.TwipsToPixelsY(TempHigh * (3030 / h))
			labElectricalUnits(x).Top = VB6.TwipsToPixelsY(TempHigh * (3060 / h))
		Next x
		
		LabElectricalTitles(8).Top = VB6.TwipsToPixelsY(TempHigh * (3660 / h))
		txtElectricalValues(8).Top = VB6.TwipsToPixelsY(TempHigh * (3630 / h))
		labElectricalUnits(8).Top = VB6.TwipsToPixelsY(TempHigh * (3660 / h))
		
		For x = 9 To 10
			LabElectricalTitles(x).Top = VB6.TwipsToPixelsY(TempHigh * (4140 / h))
			txtElectricalValues(x).Top = VB6.TwipsToPixelsY(TempHigh * (4110 / h))
			labElectricalUnits(x).Top = VB6.TwipsToPixelsY(TempHigh * (4140 / h))
		Next x
		
		For x = 11 To 12
			LabElectricalTitles(x).Top = VB6.TwipsToPixelsY((TempHigh * (4740 / h)) + ((x - 11) * y))
			txtElectricalValues(x).Top = VB6.TwipsToPixelsY((TempHigh * (4710 / h)) + ((x - 11) * y))
			labElectricalUnits(x).Top = VB6.TwipsToPixelsY((TempHigh * (4740 / h)) + ((x - 11) * y))
		Next x
		
		For x = 13 To 14
			LabElectricalTitles(x).Top = VB6.TwipsToPixelsY(TempHigh * (5640 / h))
			txtElectricalValues(x).Top = VB6.TwipsToPixelsY(TempHigh * (5610 / h))
			labElectricalUnits(x).Top = VB6.TwipsToPixelsY(TempHigh * (5640 / h))
		Next x
		
		
		For x = 0 To 2
			lstElectricalList(x).Top = VB6.TwipsToPixelsY(TempHigh * (4140 / h) + (x * s))
			lstElectricalList(x).Height = VB6.TwipsToPixelsY(TempHigh * (1185 / h))
			lstElectricalList(x).Left = VB6.TwipsToPixelsX(TempWide * (360 / w) + (x * t))
			lstElectricalList(x).Width = VB6.TwipsToPixelsX(TempWide * (1875 / w))
		Next x
		
		LineHorizontal(0).X1 = VB6.TwipsToPixelsX(TempWide * (1440 / w))
		LineHorizontal(0).X2 = VB6.TwipsToPixelsX(TempWide * (2580 / w))
		LineHorizontal(0).Y1 = VB6.TwipsToPixelsY(TempHigh * (3900 / h))
		LineHorizontal(0).Y2 = VB6.TwipsToPixelsY(TempHigh * (3900 / h))
		
		LineHorizontal(1).X1 = VB6.TwipsToPixelsX(TempWide * (180 / w))
		LineHorizontal(1).X2 = VB6.TwipsToPixelsX(TempWide * (2580 / w))
		LineHorizontal(1).Y1 = VB6.TwipsToPixelsY(TempHigh * (5520 / h))
		LineHorizontal(1).Y2 = VB6.TwipsToPixelsY(TempHigh * (5520 / h))
		
		LineHorizontal(2).X1 = VB6.TwipsToPixelsX(TempWide * (4440 / w))
		LineHorizontal(2).X2 = VB6.TwipsToPixelsX(TempWide * (9060 / w))
		LineHorizontal(2).Y1 = VB6.TwipsToPixelsY(TempHigh * (180 / h))
		LineHorizontal(2).Y2 = VB6.TwipsToPixelsY(TempHigh * (180 / h))
		
		LineHorizontal(3).X1 = VB6.TwipsToPixelsX(TempWide * (3720 / w))
		LineHorizontal(3).X2 = VB6.TwipsToPixelsX(TempWide * (8940 / w))
		LineHorizontal(3).Y1 = VB6.TwipsToPixelsY(TempHigh * (1260 / h))
		LineHorizontal(3).Y2 = VB6.TwipsToPixelsY(TempHigh * (1260 / h))
		
		LineHorizontal(4).X1 = VB6.TwipsToPixelsX(TempWide * (5100 / w))
		LineHorizontal(4).X2 = VB6.TwipsToPixelsX(TempWide * (6420 / w))
		LineHorizontal(4).Y1 = VB6.TwipsToPixelsY(TempHigh * (1860 / h))
		LineHorizontal(4).Y2 = VB6.TwipsToPixelsY(TempHigh * (1860 / h))
		
		LineHorizontal(5).X1 = VB6.TwipsToPixelsX(TempWide * (4620 / w))
		LineHorizontal(5).X2 = VB6.TwipsToPixelsX(TempWide * (8940 / w))
		LineHorizontal(5).Y1 = VB6.TwipsToPixelsY(TempHigh * (2340 / h))
		LineHorizontal(5).Y2 = VB6.TwipsToPixelsY(TempHigh * (2340 / h))
		
		LineHorizontal(6).X1 = VB6.TwipsToPixelsX(TempWide * (5100 / w))
		LineHorizontal(6).X2 = VB6.TwipsToPixelsX(TempWide * (6420 / w))
		LineHorizontal(6).Y1 = VB6.TwipsToPixelsY(TempHigh * (2940 / h))
		LineHorizontal(6).Y2 = VB6.TwipsToPixelsY(TempHigh * (2940 / h))
		
		LineHorizontal(7).X1 = VB6.TwipsToPixelsX(TempWide * (4380 / w))
		LineHorizontal(7).X2 = VB6.TwipsToPixelsX(TempWide * (8940 / w))
		LineHorizontal(7).Y1 = VB6.TwipsToPixelsY(TempHigh * (3420 / h))
		LineHorizontal(7).Y2 = VB6.TwipsToPixelsY(TempHigh * (3420 / h))
		
		LineHorizontal(8).X1 = VB6.TwipsToPixelsX(TempWide * (5100 / w))
		LineHorizontal(8).X2 = VB6.TwipsToPixelsX(TempWide * (6420 / w))
		LineHorizontal(8).Y1 = VB6.TwipsToPixelsY(TempHigh * (4020 / h))
		LineHorizontal(8).Y2 = VB6.TwipsToPixelsY(TempHigh * (4020 / h))
		
		LineHorizontal(9).X1 = VB6.TwipsToPixelsX(TempWide * (4200 / w))
		LineHorizontal(9).X2 = VB6.TwipsToPixelsX(TempWide * (8940 / w))
		LineHorizontal(9).Y1 = VB6.TwipsToPixelsY(TempHigh * (4500 / h))
		LineHorizontal(9).Y2 = VB6.TwipsToPixelsY(TempHigh * (4500 / h))
		
		LineHorizontal(10).X1 = VB6.TwipsToPixelsX(TempWide * (5100 / w))
		LineHorizontal(10).X2 = VB6.TwipsToPixelsX(TempWide * (6420 / w))
		LineHorizontal(10).Y1 = VB6.TwipsToPixelsY(TempHigh * (5520 / h))
		LineHorizontal(10).Y2 = VB6.TwipsToPixelsY(TempHigh * (5520 / h))
		
		LineHorizontal(11).X1 = VB6.TwipsToPixelsX(TempWide * (2580 / w))
		LineHorizontal(11).X2 = VB6.TwipsToPixelsX(TempWide * (9060 / w))
		LineHorizontal(11).Y1 = VB6.TwipsToPixelsY(TempHigh * (6000 / h))
		LineHorizontal(11).Y2 = VB6.TwipsToPixelsY(TempHigh * (6000 / h))
		
		LineVertical(0).X1 = VB6.TwipsToPixelsX(TempWide * (240 / w))
		LineVertical(0).X2 = VB6.TwipsToPixelsX(TempWide * (240 / w))
		LineVertical(0).Y1 = VB6.TwipsToPixelsY(TempHigh * (4140 / h))
		LineVertical(0).Y2 = VB6.TwipsToPixelsY(TempHigh * (5580 / h))
		
		For x = 1 To 5
			LineVertical(x).X1 = VB6.TwipsToPixelsX(TempWide * (2640 / w))
			LineVertical(x).X2 = VB6.TwipsToPixelsX(TempWide * (2640 / w))
		Next x
		
		LineVertical(1).Y1 = VB6.TwipsToPixelsY(TempHigh * (420 / h))
		LineVertical(1).Y2 = VB6.TwipsToPixelsY(TempHigh * (1140 / h))
		
		LineVertical(2).Y1 = VB6.TwipsToPixelsY(TempHigh * (1500 / h))
		LineVertical(2).Y2 = VB6.TwipsToPixelsY(TempHigh * (2220 / h))
		
		LineVertical(3).Y1 = VB6.TwipsToPixelsY(TempHigh * (2580 / h))
		LineVertical(3).Y2 = VB6.TwipsToPixelsY(TempHigh * (3300 / h))
		
		LineVertical(4).Y1 = VB6.TwipsToPixelsY(TempHigh * (3660 / h))
		LineVertical(4).Y2 = VB6.TwipsToPixelsY(TempHigh * (4380 / h))
		
		LineVertical(5).Y1 = VB6.TwipsToPixelsY(TempHigh * (4740 / h))
		LineVertical(5).Y2 = VB6.TwipsToPixelsY(TempHigh * (6060 / h))
		
		For x = 6 To 9
			If x = 8 Then
				LineVertical(x).X1 = VB6.TwipsToPixelsX(TempWide * (6060 / w))
				LineVertical(x).X2 = VB6.TwipsToPixelsX(TempWide * (6060 / w))
			Else
				LineVertical(x).X1 = VB6.TwipsToPixelsX(TempWide * (5400 / w))
				LineVertical(x).X2 = VB6.TwipsToPixelsX(TempWide * (5400 / w))
			End If
		Next x
		
		LineVertical(6).Y1 = VB6.TwipsToPixelsY(TempHigh * (1920 / h))
		LineVertical(6).Y2 = VB6.TwipsToPixelsY(TempHigh * (2280 / h))
		
		LineVertical(7).Y1 = VB6.TwipsToPixelsY(TempHigh * (3000 / h))
		LineVertical(7).Y2 = VB6.TwipsToPixelsY(TempHigh * (3360 / h))
		
		LineVertical(8).Y1 = VB6.TwipsToPixelsY(TempHigh * (4080 / h))
		LineVertical(8).Y2 = VB6.TwipsToPixelsY(TempHigh * (4440 / h))
		
		LineVertical(9).Y1 = VB6.TwipsToPixelsY(TempHigh * (5580 / h))
		LineVertical(9).Y2 = VB6.TwipsToPixelsY(TempHigh * (5940 / h))
		
		LineVertical(10).X1 = VB6.TwipsToPixelsX(TempWide * (9000 / w))
		LineVertical(10).X2 = VB6.TwipsToPixelsX(TempWide * (9000 / w))
		LineVertical(10).Y1 = VB6.TwipsToPixelsY(TempHigh * (120 / h))
		LineVertical(10).Y2 = VB6.TwipsToPixelsY(TempHigh * (6060 / h))
		
		labBackToMenu.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labBackToMenu.Left = VB6.TwipsToPixelsX(TempWide * (600 / w))
		
		imgBackToMenu.Top = VB6.TwipsToPixelsY(TempHigh * (6150 / h))
		imgBackToMenu.Left = VB6.TwipsToPixelsX(TempWide * (60 / w))
		imgBackToMenu.Width = VB6.TwipsToPixelsX(TempWide * (495 / w))
		
		comElectricalPrint.Top = VB6.TwipsToPixelsY(TempHigh * (6150 / h))
		comElectricalPrint.Left = VB6.TwipsToPixelsX(TempWide * (8880 / w))
		
		labElectricalHelp.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labElectricalHelp.Left = VB6.TwipsToPixelsX(TempWide * (8340 / w))
		
		labInsert.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labInsert.Left = VB6.TwipsToPixelsX(TempWide * (4080 / w))
		labInsert.Width = VB6.TwipsToPixelsX(TempWide * (975 / w))
		
	End Sub
	
	Public Sub drawthevalues()
		
		Dim i As Short
        'Dim x As Short

        DoNotChange = True
		
		For i = 0 To 14
			If CellValues(WhichScreen, i, WhichSegment).Changed = True Then
				txtElectricalValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF0000)
			Else
				txtElectricalValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
			End If
			Select Case i
				Case 0, 1, 2, 3, 5, 6, 8, 12, 13
					txtElectricalValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value)), "##,###,##0")
				Case 9
					txtElectricalValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value * FootConversion)), "#,###,###,##0")
				Case 10
					txtElectricalValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value / FootConversion)), "$###,##0.00")
				Case 11
					txtElectricalValues(i).Text = LTrim(RTrim(CellValues(WhichScreen, i, WhichSegment).Word))
				Case Else
					txtElectricalValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value)), "$##,###,##0")
			End Select
		Next i
		
		DoNotChange = False
		
	End Sub
	Private Sub txtElectricalValues_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtElectricalValues.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Dim Index As Short = txtElectricalValues.GetIndex(eventSender)
		
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
						If InStr(txtElectricalValues(Index).Text, ".") = 0 Then
							System.Windows.Forms.SendKeys.Send("{DELETE}")
						End If
					Else
						System.Windows.Forms.SendKeys.Send("{DELETE}")
					End If
			End Select
		End If
		
	End Sub
	Private Sub txtElectricalValues_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtElectricalValues.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = txtElectricalValues.GetIndex(eventSender)
		If KeyAscii > Asc("9") And KeyAscii <> Asc(",") And KeyAscii <> Asc(".") And KeyAscii <> Asc("$") Then
			Beep()
			KeyAscii = 0
		Else
			CellValues(WhichScreen, WhichCell, WhichSegment).Changed = True
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	Private Sub txtElectricalValues_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtElectricalValues.Leave
		Dim Index As Short = txtElectricalValues.GetIndex(eventSender)
		If DoNotChange = True Then Exit Sub
		WhichCell = Index
		Call Inputer(WhichCell)
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
		tempvalue = ""
		For i = 1 To Len(txtElectricalValues(Sample).Text)
			Digit.Value = Mid(txtElectricalValues(Sample).Text, i, 1)
			Select Case Digit.Value
				Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ".", "-"
					tempvalue = tempvalue & Digit.Value
			End Select
		Next i
		Select Case Sample
			Case 9
				If CellValues(WhichScreen, Sample, WhichSegment).Changed = True Then
					If FootConversion <> 0 Then CellValues(WhichScreen, Sample, WhichSegment).Value = Val(tempvalue) / FootConversion
				End If
			Case 10
				If CellValues(WhichScreen, Sample, WhichSegment).Changed = True Then
					If FootConversion <> 0 Then CellValues(WhichScreen, Sample, WhichSegment).Value = Val(tempvalue) * FootConversion
				End If
			Case 11
				CellValues(WhichScreen, Sample, WhichSegment).Word = txtElectricalValues(Sample).Text
			Case Else
				If CellValues(WhichScreen, Sample, WhichSegment).Changed = True Then
					CellValues(WhichScreen, Sample, WhichSegment).Value = Val(tempvalue)
				End If
		End Select
		Call ElectEngr()
		Call drawthevalues()
	End Sub
	'UPGRADE_WARNING: Event txtSegmentLabel.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtSegmentLabel_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSegmentLabel.TextChanged
		If DoNotChange = True Then Exit Sub
		SegNamie(WhichSegment) = txtSegmentLabel.Text
	End Sub
End Class