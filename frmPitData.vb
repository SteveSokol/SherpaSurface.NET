Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks
Friend Class frmPitData
	Inherits System.Windows.Forms.Form
	Dim TempHigh As Single
	Dim TempWide As Single
	Dim GallonConversion As Single
	Dim FootConversion As Single
	Private Sub comPitPrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles comPitPrint.Click
		PrintForm1.Print(Me, PowerPacks.Printing.PrintForm.PrintOption.CompatibleModeClientAreaOnly)
	End Sub
	'UPGRADE_WARNING: Form event frmPitData.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmPitData_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated

        'Dim baseunit As String
        'Dim baselength As Short
        Dim i As Short
		
		If IsHelpOn = True Then
			txtPitValues(WhichCell).Focus()
			IsHelpOn = False
		Else
			For i = 0 To 5
				If optSegment(i).Checked = True Then WhichSegment = i
			Next i
			
			WhichScreen = Pit
			
			Call drawthevalues()
			
			If InsertFlag = True Then
				labInsert.Text = "Insert"
			Else
				labInsert.Text = "Typeover"
			End If
			
			WhichCell = 0
			
			txtPitValues(0).Focus()
		End If
		
	End Sub
	Private Sub frmPitData_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim x As Short
		
		Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - (VB6.PixelsToTwipsY(Me.Height) + 350)) / 2)
		Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2)
		
		If VB6.PixelsToTwipsY(Me.Top) < 0 Then Me.Top = 0
		If VB6.PixelsToTwipsX(Me.Left) < 0 Then Me.Left = 0
		
		TempWide = VB6.PixelsToTwipsX(Me.ClientRectangle.Width)
		TempHigh = VB6.PixelsToTwipsY(Me.ClientRectangle.Height)
		
		DoNotChange = True
		
		If UnitType = Metric Then
			labPitUnits(6).Text = "liters per minute"
			For x = 0 To 2
				labPitUnits(x).Text = "meters"
			Next x
			For x = 4 To 5
				labPitUnits(x).Text = "meters"
				labPitUnits(x * 2).Text = "meters"
			Next x
			GallonConversion = 3.785
			FootConversion = 0.3048
		Else
			labPitUnits(6).Text = "gallons per minute"
			For x = 0 To 2
				labPitUnits(x).Text = "feet"
			Next x
			For x = 4 To 5
				labPitUnits(x).Text = "feet"
				labPitUnits(x * 2).Text = "feet"
			Next x
			GallonConversion = 1
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
	'UPGRADE_WARNING: Event frmPitData.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub frmPitData_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		TempWide = VB6.PixelsToTwipsX(Me.ClientRectangle.Width)
		TempHigh = VB6.PixelsToTwipsY(Me.ClientRectangle.Height)
		Call screenstuff()
	End Sub
	Private Sub frmPitData_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		On Error Resume Next
        'Me.Close()
		Call InputMenuAccess(1)
	End Sub
	Private Sub imgBackToMenu_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles imgBackToMenu.Click
		Me.Close()
		Call InputMenuAccess(1)
	End Sub
	Private Sub labBackToMenu_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles labBackToMenu.Click
		Me.Close()
		Call InputMenuAccess(1)
	End Sub
	Private Sub labPitHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles labPitHelp.Click
		Dim StartHelp As Short
		IsHelpOn = True
		StartHelp = 17
		Call frmSurfaceHelp.gethelptext(StartHelp, WhichCell)
		frmSurfaceHelp.Show()
	End Sub
	'UPGRADE_WARNING: Event lstPitList.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub lstPitList_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstPitList.SelectedIndexChanged
        'Dim x As Short

        txtPitValues(WhichCell).Text = LTrim(RTrim(VB6.GetItemString(lstPitList, lstPitList.SelectedIndex)))
		CellValues(WhichScreen, WhichCell, WhichSegment).Changed = True
		
		Call Inputer(WhichCell)
		
		If WhichCell = 4 Then
			txtPitValues(WhichCell + 1).Focus()
		Else
			txtPitValues(0).Focus()
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
			labSegment(WhichSegment).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF)
			txtSegmentLabel.Text = SegNamie(WhichSegment)
			txtPitValues(WhichCell).Focus()
		End If
	End Sub
	Private Sub txtPitValues_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPitValues.Enter
		Dim Index As Short = txtPitValues.GetIndex(eventSender)
        'Dim x As Short
        WhichCell = Index
		
		lstPitList.Visible = False
		
		If WhichCell = 4 Or WhichCell = 7 Then
			lstPitList.Visible = False
		End If
		
		System.Windows.Forms.SendKeys.Send("{HOME}+{END}")
		
		WhichCell = Index
		Call drawthevalues()
	End Sub
	'UPGRADE_WARNING: Event txtPitValues.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtPitValues_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPitValues.TextChanged
		Dim Index As Short = txtPitValues.GetIndex(eventSender)
		If DoNotChange = True Then Exit Sub
	End Sub
	Public Sub screenstuff()

        'Dim q As Decimal
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
		
		r = (1080 / h) * TempHigh
		s = (60 / h) * TempHigh
		t = (120 / w) * TempWide
		u = (540 / h) * TempHigh
		v = (900 / w) * TempWide
		
		y = (420 / h) * TempHigh
		z = (300 / w) * TempWide
		
		labPitHeading.Top = VB6.TwipsToPixelsY(TempHigh * (240 / h))
		labPitHeading.Left = VB6.TwipsToPixelsX(TempWide * (240 / w))
		labPitHeading.Width = VB6.TwipsToPixelsX(TempWide * (1965 / w))
		
		For x = 0 To 1
			labPitLabels(x).Top = VB6.TwipsToPixelsY((TempHigh * (1080 / h)) + (x * r))
			labPitLabels(x).Left = VB6.TwipsToPixelsX(TempWide * (300 / w))
			labPitLabels(x).Width = VB6.TwipsToPixelsX(TempWide * (1935 / w))
		Next x
		
		labPitLabels(2).Top = VB6.TwipsToPixelsY(TempHigh * (3060 / h))
		labPitLabels(2).Left = VB6.TwipsToPixelsX(TempWide * (300 / w))
		
		For x = 3 To 4
			labPitLabels(x).Left = VB6.TwipsToPixelsX(TempWide * (2460 / w))
		Next x
		
		labPitLabels(3).Top = VB6.TwipsToPixelsY(TempHigh * (240 / h))
		labPitLabels(4).Top = VB6.TwipsToPixelsY(TempHigh * (3960 / h))
		
		For x = 0 To 5
			optSegment(x).Top = VB6.TwipsToPixelsY(TempHigh * (1380 / h))
			optSegment(x).Left = VB6.TwipsToPixelsX((TempWide * (420 / w)) + (x * z))
			optSegment(x).Width = VB6.TwipsToPixelsX(TempWide * (195 / w))
			labSegment(x).Top = VB6.TwipsToPixelsY(TempHigh * (1620 / h))
			labSegment(x).Left = VB6.TwipsToPixelsX((TempWide * (420 / w)) + (x * z))
			labSegment(x).Width = VB6.TwipsToPixelsX(TempWide * (195 / w))
		Next x
		
		txtSegmentLabel.Top = VB6.TwipsToPixelsY(TempHigh * (2460 / h))
		txtSegmentLabel.Left = VB6.TwipsToPixelsX(TempWide * (300 / w))
		txtSegmentLabel.Width = VB6.TwipsToPixelsX(TempWide * (1935 / w))
		
		lstPitList.Top = VB6.TwipsToPixelsY(TempHigh * (3420 / h))
		lstPitList.Left = VB6.TwipsToPixelsX(TempWide * (660 / w))
		lstPitList.Height = VB6.TwipsToPixelsY(TempHigh * (2085 / h))
		lstPitList.Width = VB6.TwipsToPixelsX(TempWide * (1515 / w))
		
		For x = 0 To 7
			LabPitTitles(x).Left = VB6.TwipsToPixelsX(TempWide * (2700 / w))
			LabPitTitles(x).Width = VB6.TwipsToPixelsX(TempWide * (3015 / w))
			txtPitValues(x).Left = VB6.TwipsToPixelsX(TempWide * (5940 / w))
			txtPitValues(x).Width = VB6.TwipsToPixelsX(TempWide * (795 / w))
			labPitUnits(x).Left = VB6.TwipsToPixelsX(TempWide * (6780 / w))
		Next x
		
		For x = 8 To 11
			LabPitTitles(x).Left = VB6.TwipsToPixelsX(TempWide * (3420 / w))
			LabPitTitles(x).Width = VB6.TwipsToPixelsX(TempWide * (3315 / w))
			txtPitValues(x).Left = VB6.TwipsToPixelsX(TempWide * (6960 / w))
			txtPitValues(x).Width = VB6.TwipsToPixelsX(TempWide * (795 / w))
			labPitUnits(x).Left = VB6.TwipsToPixelsX(TempWide * (7800 / w))
		Next x
		
		For x = 0 To 7
			LabPitTitles(x).Top = VB6.TwipsToPixelsY(TempHigh * (630 / h) + (x * y))
			txtPitValues(x).Top = VB6.TwipsToPixelsY(TempHigh * (600 / h) + (x * y))
			labPitUnits(x).Top = VB6.TwipsToPixelsY(TempHigh * (630 / h) + (x * y))
		Next x
		
		For x = 8 To 11
			LabPitTitles(x).Top = VB6.TwipsToPixelsY((TempHigh * (4350 / h)) + ((x - 8) * y))
			txtPitValues(x).Top = VB6.TwipsToPixelsY((TempHigh * (4320 / h)) + ((x - 8) * y))
			labPitUnits(x).Top = VB6.TwipsToPixelsY((TempHigh * (4350 / h)) + ((x - 8) * y))
		Next x
		
		LineHorizontal(0).X1 = VB6.TwipsToPixelsX(TempWide * (1620 / w))
		LineHorizontal(0).X2 = VB6.TwipsToPixelsX(TempWide * (2460 / w))
		LineHorizontal(0).Y1 = VB6.TwipsToPixelsY(TempHigh * (3120 / h))
		LineHorizontal(0).Y2 = VB6.TwipsToPixelsY(TempHigh * (3120 / h))
		
		LineHorizontal(1).X1 = VB6.TwipsToPixelsX(TempWide * (300 / w))
		LineHorizontal(1).X2 = VB6.TwipsToPixelsX(TempWide * (2460 / w))
		LineHorizontal(1).Y1 = VB6.TwipsToPixelsY(TempHigh * (5640 / h))
		LineHorizontal(1).Y2 = VB6.TwipsToPixelsY(TempHigh * (5640 / h))
		
		LineHorizontal(2).X1 = VB6.TwipsToPixelsX(TempWide * (4260 / w))
		LineHorizontal(2).X2 = VB6.TwipsToPixelsX(TempWide * (8880 / w))
		LineHorizontal(2).Y1 = VB6.TwipsToPixelsY(TempHigh * (300 / h))
		LineHorizontal(2).Y2 = VB6.TwipsToPixelsY(TempHigh * (300 / h))
		
		LineHorizontal(3).X1 = VB6.TwipsToPixelsX(TempWide * (3840 / w))
		LineHorizontal(3).X2 = VB6.TwipsToPixelsX(TempWide * (8760 / w))
		LineHorizontal(3).Y1 = VB6.TwipsToPixelsY(TempHigh * (3960 / h))
		LineHorizontal(3).Y2 = VB6.TwipsToPixelsY(TempHigh * (3960 / h))
		
		LineHorizontal(4).X1 = VB6.TwipsToPixelsX(TempWide * (2460 / w))
		LineHorizontal(4).X2 = VB6.TwipsToPixelsX(TempWide * (8880 / w))
		LineHorizontal(4).Y1 = VB6.TwipsToPixelsY(TempHigh * (6000 / h))
		LineHorizontal(4).Y2 = VB6.TwipsToPixelsY(TempHigh * (6000 / h))
		
		LineVertical(0).X1 = VB6.TwipsToPixelsX(TempWide * (360 / w))
		LineVertical(0).X2 = VB6.TwipsToPixelsX(TempWide * (360 / w))
		LineVertical(0).Y1 = VB6.TwipsToPixelsY(TempHigh * (3360 / h))
		LineVertical(0).Y2 = VB6.TwipsToPixelsY(TempHigh * (5700 / h))
		
		LineVertical(1).X1 = VB6.TwipsToPixelsX(TempWide * (2520 / w))
		LineVertical(1).X2 = VB6.TwipsToPixelsX(TempWide * (2520 / w))
		LineVertical(1).Y1 = VB6.TwipsToPixelsY(TempHigh * (540 / h))
		LineVertical(1).Y2 = VB6.TwipsToPixelsY(TempHigh * (3900 / h))
		
		LineVertical(2).X1 = VB6.TwipsToPixelsX(TempWide * (2520 / w))
		LineVertical(2).X2 = VB6.TwipsToPixelsX(TempWide * (2520 / w))
		LineVertical(2).Y1 = VB6.TwipsToPixelsY(TempHigh * (4260 / h))
		LineVertical(2).Y2 = VB6.TwipsToPixelsY(TempHigh * (6060 / h))
		
		LineVertical(3).X1 = VB6.TwipsToPixelsX(TempWide * (8820 / w))
		LineVertical(3).X2 = VB6.TwipsToPixelsX(TempWide * (8820 / w))
		LineVertical(3).Y1 = VB6.TwipsToPixelsY(TempHigh * (240 / h))
		LineVertical(3).Y2 = VB6.TwipsToPixelsY(TempHigh * (6060 / h))
		
		labBackToMenu.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labBackToMenu.Left = VB6.TwipsToPixelsX(TempWide * (600 / w))
		
		imgBackToMenu.Top = VB6.TwipsToPixelsY(TempHigh * (6150 / h))
		imgBackToMenu.Left = VB6.TwipsToPixelsX(TempWide * (60 / w))
		imgBackToMenu.Width = VB6.TwipsToPixelsX(TempWide * (495 / w))
		
		comPitPrint.Top = VB6.TwipsToPixelsY(TempHigh * (6150 / h))
		comPitPrint.Left = VB6.TwipsToPixelsX(TempWide * (8880 / w))
		
		labPitHelp.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labPitHelp.Left = VB6.TwipsToPixelsX(TempWide * (8340 / w))
		
		labInsert.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labInsert.Left = VB6.TwipsToPixelsX(TempWide * (4080 / w))
		labInsert.Width = VB6.TwipsToPixelsX(TempWide * (975 / w))
		
	End Sub
	Public Sub drawthevalues()
		
		Dim i As Short
        'Dim x As Short

        DoNotChange = True
		
		For i = 0 To 11
			If CellValues(WhichScreen, i, WhichSegment).Changed = True Then
				txtPitValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF0000)
			Else
				txtPitValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
			End If
			Select Case i
				Case 6
					txtPitValues(i).Text = VB6.Format(LTrim(Str(CellValues(Pumping, 0, WhichSegment).Value * GallonConversion)), "#,###,###,##0")
				Case 0, 1, 2, 8, 10
					txtPitValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value * FootConversion)), "#,###,###,##0")
				Case 4
					txtPitValues(i).Text = VB6.Format(LTrim(Str(CellValues(Deposit, 5, WhichSegment).Value * FootConversion)), "#,##0.0")
				Case 5
					txtPitValues(i).Text = VB6.Format(LTrim(Str(CellValues(Deposit, 12, WhichSegment).Value * FootConversion)), "#,##0.0")
				Case Else
					txtPitValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value)), "##,###,##0.0")
			End Select
		Next i
		
		DoNotChange = False
		
	End Sub
	Private Sub txtPitValues_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtPitValues.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Dim Index As Short = txtPitValues.GetIndex(eventSender)
		
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
						If InStr(txtPitValues(Index).Text, ".") = 0 Then
							System.Windows.Forms.SendKeys.Send("{DELETE}")
						End If
					Else
						System.Windows.Forms.SendKeys.Send("{DELETE}")
					End If
			End Select
		End If
		
	End Sub
	Private Sub txtPitValues_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtPitValues.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = txtPitValues.GetIndex(eventSender)
		If KeyAscii > Asc("9") And KeyAscii <> Asc(",") And KeyAscii <> Asc(".") And KeyAscii <> Asc("$") Then
			Beep()
			KeyAscii = 0
		Else
			Select Case WhichCell
				Case 6
					CellValues(Pumping, 0, WhichSegment).Changed = True
				Case 4
					CellValues(Deposit, 5, WhichSegment).Changed = True
				Case 5
					CellValues(Deposit, 12, WhichSegment).Changed = True
				Case Else
					CellValues(WhichScreen, WhichCell, WhichSegment).Changed = True
			End Select
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	Private Sub txtPitValues_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPitValues.Leave
		Dim Index As Short = txtPitValues.GetIndex(eventSender)
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
		For i = 1 To Len(txtPitValues(Sample).Text)
			Digit.Value = Mid(txtPitValues(Sample).Text, i, 1)
			Select Case Digit.Value
				Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ".", "-"
					tempvalue = tempvalue & Digit.Value
			End Select
		Next i
		Select Case Sample
			Case 6
				If CellValues(Pumping, 0, WhichSegment).Changed = True Then
					If GallonConversion <> 0 Then CellValues(Pumping, 0, WhichSegment).Value = Val(tempvalue) / GallonConversion
				End If
			Case 0, 1, 2, 8, 10
				If CellValues(WhichScreen, Sample, WhichSegment).Changed = True Then
					If FootConversion <> 0 Then CellValues(WhichScreen, Sample, WhichSegment).Value = Val(tempvalue) / FootConversion
				End If
			Case 4
				If CellValues(Deposit, 5, WhichSegment).Changed = True Then
					If FootConversion <> 0 Then CellValues(Deposit, 5, WhichSegment).Value = Val(tempvalue) / FootConversion
				End If
			Case 5
				If CellValues(Deposit, 12, WhichSegment).Changed = True Then
					If FootConversion <> 0 Then CellValues(Deposit, 12, WhichSegment).Value = Val(tempvalue) / FootConversion
				End If
			Case Else
				If CellValues(WhichScreen, Sample, WhichSegment).Changed = True Then
					CellValues(WhichScreen, Sample, WhichSegment).Value = Val(tempvalue)
				End If
		End Select
		Call drawthevalues()
	End Sub
	'UPGRADE_WARNING: Event txtSegmentLabel.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtSegmentLabel_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSegmentLabel.TextChanged
		If DoNotChange = True Then Exit Sub
		SegNamie(WhichSegment) = txtSegmentLabel.Text
	End Sub
End Class