Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks
Friend Class frmFuelStorageData
	Inherits System.Windows.Forms.Form
	Dim TempVolume As String
	Dim TempHigh As Single
	Dim TempWide As Single
	Dim GallonConversion As Single
	Private Sub comFuelStoragePrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles comFuelStoragePrint.Click
		PrintForm1.Print(Me, PowerPacks.Printing.PrintForm.PrintOption.CompatibleModeClientAreaOnly)
	End Sub
	'UPGRADE_WARNING: Form event frmFuelStorageData.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmFuelStorageData_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated

        'Dim baseunit As String
        'Dim baselength As Short
        Dim i As Short
		
		If IsHelpOn = True Then
			txtFuelStorageValues(WhichCell).Focus()
			IsHelpOn = False
		Else
			For i = 0 To 5
				If optSegment(i).Checked = True Then WhichSegment = i
			Next i
			
			WhichScreen = FuelStorage
			
			Call drawthevalues()
			
			If InsertFlag = True Then
				labInsert.Text = "Insert"
			Else
				labInsert.Text = "Typeover"
			End If
			
			WhichCell = 0
			
			txtFuelStorageValues(0).Focus()
		End If
		
	End Sub
	Private Sub frmFuelStorageData_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim x As Short
		
		Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - (VB6.PixelsToTwipsY(Me.Height) + 350)) / 2)
		Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2)
		
		If VB6.PixelsToTwipsY(Me.Top) < 0 Then Me.Top = 0
		If VB6.PixelsToTwipsX(Me.Left) < 0 Then Me.Left = 0
		
		TempWide = VB6.PixelsToTwipsX(Me.ClientRectangle.Width)
		TempHigh = VB6.PixelsToTwipsY(Me.ClientRectangle.Height)
		
		DoNotChange = True
		
		If UnitType = Metric Then
			For x = 0 To 7
				If x < 4 Then
					labFuelStorageUnits(x).Text = "liters/day"
				ElseIf x = 5 Then 
					labFuelStorageUnits(x).Text = "liters"
				ElseIf x = 6 Then 
					labFuelStorageUnits(x).Text = "liter"
				End If
			Next x
			GallonConversion = 3.785
		Else
			For x = 0 To 7
				If x < 4 Then
					labFuelStorageUnits(x).Text = "gallons/day"
				ElseIf x = 5 Then 
					labFuelStorageUnits(x).Text = "gallons"
				ElseIf x = 6 Then 
					labFuelStorageUnits(x).Text = "gallon"
				End If
			Next x
			GallonConversion = 1
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
	'UPGRADE_WARNING: Event frmFuelStorageData.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub frmFuelStorageData_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		TempWide = VB6.PixelsToTwipsX(Me.ClientRectangle.Width)
		TempHigh = VB6.PixelsToTwipsY(Me.ClientRectangle.Height)
		Call screenstuff()
	End Sub
	Private Sub frmFuelStorageData_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
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
	Private Sub labFuelStorageHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles labFuelStorageHelp.Click
		Dim StartHelp As Short
		StartHelp = 269
		IsHelpOn = True
		Call frmSurfaceHelp.gethelptext(StartHelp, WhichCell)
		frmSurfaceHelp.Show()
	End Sub
	'UPGRADE_WARNING: Event lstTankOptions.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub lstTankOptions_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstTankOptions.SelectedIndexChanged
		Dim Index As Short = lstTankOptions.GetIndex(eventSender)
		TempVolume = LTrim(RTrim(VB6.GetItemString(lstTankOptions(Index), lstTankOptions(Index).SelectedIndex)))
		CellValues(WhichScreen, WhichCell, WhichSegment).Changed = True
		Call NumberDecider()
		Call Inputer(WhichCell)
		WhichCell = WhichCell + 1
		txtFuelStorageValues(WhichCell).Focus()
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
			txtFuelStorageValues(WhichCell).Focus()
		End If
	End Sub
	Private Sub txtFuelStorageValues_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFuelStorageValues.Enter
		Dim Index As Short = txtFuelStorageValues.GetIndex(eventSender)
		Dim x As Short
		WhichCell = Index
		
		System.Windows.Forms.SendKeys.Send("{HOME}+{END}")
		
		WhichCell = Index
		
		For x = 0 To 1
			lstTankOptions(x).Visible = False
		Next x
		
		If WhichCell = 6 Then
			If UnitType = Metric Then
				lstTankOptions(1).Visible = True
			Else
				lstTankOptions(0).Visible = True
			End If
		End If
		
		Call drawthevalues()
	End Sub
	'UPGRADE_WARNING: Event txtFuelStorageValues.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtFuelStorageValues_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFuelStorageValues.TextChanged
		Dim Index As Short = txtFuelStorageValues.GetIndex(eventSender)
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
		
		q = (1680 / h) * TempHigh
		r = (1080 / h) * TempHigh
		s = (60 / h) * TempHigh
		t = (120 / w) * TempWide
		u = (540 / h) * TempHigh
		v = (900 / w) * TempWide
		
		y = (420 / h) * TempHigh
		z = (300 / w) * TempWide
		
		For x = 0 To 1
			labFuelStorageHeading(x).Top = VB6.TwipsToPixelsY((TempHigh * (180 / h)) + (x * u))
			labFuelStorageHeading(x).Left = VB6.TwipsToPixelsX((TempWide * (180 / w)) + (x * v))
			labFuelStorageHeading(x).Width = VB6.TwipsToPixelsX(TempWide * (1965 / w))
			labFuelStorageLabels(x).Top = VB6.TwipsToPixelsY((TempHigh * (1560 / h)) + (x * r))
			labFuelStorageLabels(x).Left = VB6.TwipsToPixelsX(TempWide * (600 / w))
			labFuelStorageLabels(x).Width = VB6.TwipsToPixelsX(TempWide * (1935 / w))
			labFuelStorageLabels(x + 3).Top = VB6.TwipsToPixelsY((TempHigh * (780 / h)) + (x * q))
			labFuelStorageLabels(x + 3).Left = VB6.TwipsToPixelsX(TempWide * (3240 / w))
			lstTankOptions(x).Top = VB6.TwipsToPixelsY((TempHigh * (4020 / h)) + (x * s))
			lstTankOptions(x).Left = VB6.TwipsToPixelsX((TempWide * (1440 / w)) + (x * t))
			lstTankOptions(x).Height = VB6.TwipsToPixelsY(TempHigh * (1635 / h))
			lstTankOptions(x).Width = VB6.TwipsToPixelsX(TempWide * (1515 / w))
		Next x
		
		labFuelStorageLabels(2).Left = VB6.TwipsToPixelsX(TempWide * (1140 / w))
		labFuelStorageLabels(2).Top = VB6.TwipsToPixelsY(TempHigh * (3660 / h))
		
		For x = 0 To 5
			optSegment(x).Top = VB6.TwipsToPixelsY(TempHigh * (1860 / h))
			optSegment(x).Left = VB6.TwipsToPixelsX((TempWide * (720 / w)) + (x * z))
			optSegment(x).Width = VB6.TwipsToPixelsX(TempWide * (195 / w))
			labSegment(x).Top = VB6.TwipsToPixelsY(TempHigh * (2100 / h))
			labSegment(x).Left = VB6.TwipsToPixelsX((TempWide * (720 / w)) + (x * z))
			labSegment(x).Width = VB6.TwipsToPixelsX(TempWide * (195 / w))
		Next x
		
		txtSegmentLabel.Top = VB6.TwipsToPixelsY(TempHigh * (2940 / h))
		txtSegmentLabel.Left = VB6.TwipsToPixelsX(TempWide * (600 / w))
		txtSegmentLabel.Width = VB6.TwipsToPixelsX(TempWide * (1935 / w))
		
		For x = 0 To 7
			LabFuelStorageTitles(x).Left = VB6.TwipsToPixelsX(TempWide * (3660 / w))
			LabFuelStorageTitles(x).Width = VB6.TwipsToPixelsX(TempWide * (2355 / w))
			txtFuelStorageValues(x).Left = VB6.TwipsToPixelsX(TempWide * (6240 / w))
			txtFuelStorageValues(x).Width = VB6.TwipsToPixelsX(TempWide * (1035 / w))
			labFuelStorageUnits(x).Left = VB6.TwipsToPixelsX(TempWide * (7320 / w))
		Next x
		
		For x = 0 To 2
			LabFuelStorageTitles(x).Top = VB6.TwipsToPixelsY((TempHigh * (1200 / h)) + (x * y))
			txtFuelStorageValues(x).Top = VB6.TwipsToPixelsY(TempHigh * (1170 / h) + (x * y))
			labFuelStorageUnits(x).Top = VB6.TwipsToPixelsY((TempHigh * (1200 / h)) + (x * y))
		Next x
		
		For x = 0 To 4
			LabFuelStorageTitles(x + 3).Top = VB6.TwipsToPixelsY((TempHigh * (2880 / h)) + (x * y))
			txtFuelStorageValues(x + 3).Top = VB6.TwipsToPixelsY((TempHigh * (2850 / h)) + (x * y))
			labFuelStorageUnits(x + 3).Top = VB6.TwipsToPixelsY((TempHigh * (2880 / h)) + (x * y))
		Next x
		
		LineHorizontal(0).X1 = VB6.TwipsToPixelsX(TempWide * (2400 / w))
		LineHorizontal(0).X2 = VB6.TwipsToPixelsX(TempWide * (3240 / w))
		LineHorizontal(0).Y1 = VB6.TwipsToPixelsY(TempHigh * (3720 / h))
		LineHorizontal(0).Y2 = VB6.TwipsToPixelsY(TempHigh * (3720 / h))
		
		LineHorizontal(1).X1 = VB6.TwipsToPixelsX(TempWide * (1140 / w))
		LineHorizontal(1).X2 = VB6.TwipsToPixelsX(TempWide * (3360 / w))
		LineHorizontal(1).Y1 = VB6.TwipsToPixelsY(TempHigh * (5820 / h))
		LineHorizontal(1).Y2 = VB6.TwipsToPixelsY(TempHigh * (5820 / h))
		
		LineHorizontal(2).X1 = VB6.TwipsToPixelsX(TempWide * (5040 / w))
		LineHorizontal(2).X2 = VB6.TwipsToPixelsX(TempWide * (8880 / w))
		LineHorizontal(2).Y1 = VB6.TwipsToPixelsY(TempHigh * (840 / h))
		LineHorizontal(2).Y2 = VB6.TwipsToPixelsY(TempHigh * (840 / h))
		
		LineHorizontal(3).X1 = VB6.TwipsToPixelsX(TempWide * (5280 / w))
		LineHorizontal(3).X2 = VB6.TwipsToPixelsX(TempWide * (8700 / w))
		LineHorizontal(3).Y1 = VB6.TwipsToPixelsY(TempHigh * (2520 / h))
		LineHorizontal(3).Y2 = VB6.TwipsToPixelsY(TempHigh * (2520 / h))
		
		LineHorizontal(4).X1 = VB6.TwipsToPixelsX(TempWide * (3360 / w))
		LineHorizontal(4).X2 = VB6.TwipsToPixelsX(TempWide * (8880 / w))
		LineHorizontal(4).Y1 = VB6.TwipsToPixelsY(TempHigh * (5040 / h))
		LineHorizontal(4).Y2 = VB6.TwipsToPixelsY(TempHigh * (5040 / h))
		
		LineVertical(0).X1 = VB6.TwipsToPixelsX(TempWide * (1200 / w))
		LineVertical(0).X2 = VB6.TwipsToPixelsX(TempWide * (1200 / w))
		LineVertical(0).Y1 = VB6.TwipsToPixelsY(TempHigh * (3960 / h))
		LineVertical(0).Y2 = VB6.TwipsToPixelsY(TempHigh * (5880 / h))
		
		LineVertical(1).X1 = VB6.TwipsToPixelsX(TempWide * (3300 / w))
		LineVertical(1).X2 = VB6.TwipsToPixelsX(TempWide * (3300 / w))
		LineVertical(1).Y1 = VB6.TwipsToPixelsY(TempHigh * (1080 / h))
		LineVertical(1).Y2 = VB6.TwipsToPixelsY(TempHigh * (2400 / h))
		
		LineVertical(2).X1 = VB6.TwipsToPixelsX(TempWide * (3300 / w))
		LineVertical(2).X2 = VB6.TwipsToPixelsX(TempWide * (3300 / w))
		LineVertical(2).Y1 = VB6.TwipsToPixelsY(TempHigh * (2760 / h))
		LineVertical(2).Y2 = VB6.TwipsToPixelsY(TempHigh * (5880 / h))
		
		LineVertical(3).X1 = VB6.TwipsToPixelsX(TempWide * (8820 / w))
		LineVertical(3).X2 = VB6.TwipsToPixelsX(TempWide * (8820 / w))
		LineVertical(3).Y1 = VB6.TwipsToPixelsY(TempHigh * (780 / h))
		LineVertical(3).Y2 = VB6.TwipsToPixelsY(TempHigh * (5100 / h))
		
		labBackToMenu.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labBackToMenu.Left = VB6.TwipsToPixelsX(TempWide * (600 / w))
		
		imgBackToMenu.Top = VB6.TwipsToPixelsY(TempHigh * (6150 / h))
		imgBackToMenu.Left = VB6.TwipsToPixelsX(TempWide * (60 / w))
		imgBackToMenu.Width = VB6.TwipsToPixelsX(TempWide * (495 / w))
		
		comFuelStoragePrint.Top = VB6.TwipsToPixelsY(TempHigh * (6150 / h))
		comFuelStoragePrint.Left = VB6.TwipsToPixelsX(TempWide * (8880 / w))
		
		labFuelStorageHelp.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labFuelStorageHelp.Left = VB6.TwipsToPixelsX(TempWide * (8340 / w))
		
		labInsert.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labInsert.Left = VB6.TwipsToPixelsX(TempWide * (4080 / w))
		labInsert.Width = VB6.TwipsToPixelsX(TempWide * (975 / w))
		
	End Sub
	Public Sub drawthevalues()
		
		Dim i As Short
        'Dim x As Short

        DoNotChange = True
		
		For i = 0 To 7
			If CellValues(WhichScreen, i, WhichSegment).Changed = True Then
				txtFuelStorageValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF0000)
			Else
				txtFuelStorageValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
			End If
			Select Case i
				Case 4, 7
					txtFuelStorageValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value)), "##,###,##0")
				Case Else
					txtFuelStorageValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value * GallonConversion)), "#,###,###,##0")
			End Select
		Next i
		
		DoNotChange = False
		
	End Sub
	Private Sub txtFuelStorageValues_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtFuelStorageValues.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Dim Index As Short = txtFuelStorageValues.GetIndex(eventSender)
		
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
						If InStr(txtFuelStorageValues(Index).Text, ".") = 0 Then
							System.Windows.Forms.SendKeys.Send("{DELETE}")
						End If
					Else
						System.Windows.Forms.SendKeys.Send("{DELETE}")
					End If
			End Select
		End If
		
	End Sub
	Private Sub txtFuelStorageValues_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtFuelStorageValues.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = txtFuelStorageValues.GetIndex(eventSender)
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
	Private Sub txtFuelStorageValues_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtFuelStorageValues.Leave
		Dim Index As Short = txtFuelStorageValues.GetIndex(eventSender)
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
		For i = 1 To Len(txtFuelStorageValues(Sample).Text)
			Digit.Value = Mid(txtFuelStorageValues(Sample).Text, i, 1)
			Select Case Digit.Value
				Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ".", "-"
					tempvalue = tempvalue & Digit.Value
			End Select
		Next i
		Select Case Sample
			Case 4, 7
				If CellValues(WhichScreen, Sample, WhichSegment).Changed = True Then
					CellValues(WhichScreen, Sample, WhichSegment).Value = Val(tempvalue)
				End If
			Case 0, 1, 2, 3, 5
				If CellValues(WhichScreen, Sample, WhichSegment).Changed = True Then
					If GallonConversion <> 0 Then CellValues(WhichScreen, Sample, WhichSegment).Value = Val(tempvalue) / GallonConversion
				End If
		End Select
		Call Recalc()
		Call drawthevalues()
	End Sub
	'UPGRADE_WARNING: Event txtSegmentLabel.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtSegmentLabel_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSegmentLabel.TextChanged
		If DoNotChange = True Then Exit Sub
		SegNamie(WhichSegment) = txtSegmentLabel.Text
	End Sub
	Private Sub Recalc()
		Dim TempFuel As Decimal
		Dim r As Decimal
		Dim t As Decimal
		Dim x As Short
		Dim y As Short
		Dim jump As Boolean
		
		On Error Resume Next
		
		Call getout(jump)
		If jump = True Then Exit Sub
		
		TempFuel = 0
		
		For x = 0 To 2
			TempFuel = TempFuel + CellValues(FuelStorage, x, WhichSegment).Value
		Next x
		
		If CellValues(FuelStorage, 3, WhichSegment).Changed = False Then
			CellValues(FuelStorage, 3, WhichSegment).Value = TempFuel
		End If
		
		If CellValues(FuelStorage, 4, WhichSegment).Changed = False Then
			CellValues(FuelStorage, 4, WhichSegment).Value = 30
		End If
		
		If CellValues(FuelStorage, 5, WhichSegment).Changed = False Then
			CellValues(FuelStorage, 5, WhichSegment).Value = CellValues(FuelStorage, 3, WhichSegment).Value * CellValues(FuelStorage, 4, WhichSegment).Value
		End If
		
		r = CellValues(FuelStorage, 5, WhichSegment).Value / 15000
		x = 0
		t = 0
		While x = 0
			If r >= 1 Then
				t = t + 1
				If t > 0 Then
					r = CellValues(FuelStorage, 5, WhichSegment).Value / (t * 15000)
				End If
			Else
				x = 1
			End If
		End While
		
		If CellValues(FuelStorage, 7, WhichSegment).Changed = False Then
			CellValues(FuelStorage, 7, WhichSegment).Value = t
		End If
		
		If CellValues(FuelStorage, 7, WhichSegment).Value <> 0 Then
			TempFuel = Int((CellValues(FuelStorage, 5, WhichSegment).Value / CellValues(FuelStorage, 7, WhichSegment).Value) + 1)
		End If
		
		Select Case TempFuel
			Case Is < 1000
				TempFuel = 1000
			Case Is < 2000
				TempFuel = 2000
			Case Is < 5000
				TempFuel = 5000
			Case Is < 10000
				TempFuel = 10000
			Case Is < 12000
				TempFuel = 12000
			Case Is < 15000
				TempFuel = 15000
		End Select
		
		If CellValues(FuelStorage, 6, WhichSegment).Changed = False Then
			CellValues(FuelStorage, 6, WhichSegment).Value = TempFuel
		ElseIf CellValues(FuelStorage, 6, WhichSegment).Changed = True Then 
			If CellValues(FuelStorage, 7, WhichSegment).Changed = False And CellValues(FuelStorage, 6, WhichSegment).Value <> 0 Then
				CellValues(FuelStorage, 7, WhichSegment).Value = Int(CellValues(FuelStorage, 5, WhichSegment).Value / CellValues(FuelStorage, 6, WhichSegment).Value) + 1
			End If
		End If
		
	End Sub
	Private Sub NumberDecider()
		
		On Error Resume Next
		
		If UnitType = Metric Then
			Select Case LTrim(RTrim(TempVolume))
				Case "3,785 liter"
					CellValues(FuelStorage, 6, WhichSegment).Value = 1000
				Case "7,575 liter"
					CellValues(FuelStorage, 6, WhichSegment).Value = 2000
				Case "18,925 liter"
					CellValues(FuelStorage, 6, WhichSegment).Value = 5000
				Case "37,850 liter"
					CellValues(FuelStorage, 6, WhichSegment).Value = 10000
				Case "45,420 liter"
					CellValues(FuelStorage, 6, WhichSegment).Value = 12000
				Case "56,775 liter"
					CellValues(FuelStorage, 6, WhichSegment).Value = 15000
			End Select
		Else
			Select Case LTrim(RTrim(TempVolume))
				Case "1,000 gallon"
					CellValues(FuelStorage, 6, WhichSegment).Value = 1000
				Case "2,000 gallon"
					CellValues(FuelStorage, 6, WhichSegment).Value = 2000
				Case "5,000 gallon"
					CellValues(FuelStorage, 6, WhichSegment).Value = 5000
				Case "10,000 gallon"
					CellValues(FuelStorage, 6, WhichSegment).Value = 10000
				Case "12,000 gallon"
					CellValues(FuelStorage, 6, WhichSegment).Value = 12000
				Case "15,000 gallon"
					CellValues(FuelStorage, 6, WhichSegment).Value = 15000
			End Select
		End If
		
		Call Recalc()
		Call drawthevalues()
		
	End Sub
End Class