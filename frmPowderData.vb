Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks
Friend Class frmPowderData
	Inherits System.Windows.Forms.Form
	Dim TempHigh As Single
	Dim TempWide As Single
	Dim PoundConversion As Single
	Private Sub comPowderPrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles comPowderPrint.Click
		PrintForm1.Print(Me, PowerPacks.Printing.PrintForm.PrintOption.CompatibleModeClientAreaOnly)
	End Sub
	'UPGRADE_WARNING: Form event frmPowderData.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmPowderData_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated

        'Dim baseunit As String
        'Dim baselength As Short
        Dim i As Short
		
		If IsHelpOn = True Then
			txtPowderValues(WhichCell).Focus()
			IsHelpOn = False
		Else
			For i = 0 To 5
				If optSegment(i).Checked = True Then WhichSegment = i
			Next i
			
			WhichScreen = Powder
			
			Call drawthevalues()
			
			If InsertFlag = True Then
				labInsert.Text = "Insert"
			Else
				labInsert.Text = "Typeover"
			End If
			
			WhichCell = 0
			
			txtPowderValues(0).Focus()
		End If
		
	End Sub
	Private Sub frmPowderData_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        'Dim x As Short

        Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - (VB6.PixelsToTwipsY(Me.Height) + 350)) / 2)
		Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2)
		
		If VB6.PixelsToTwipsY(Me.Top) < 0 Then Me.Top = 0
		If VB6.PixelsToTwipsX(Me.Left) < 0 Then Me.Left = 0
		
		TempWide = VB6.PixelsToTwipsX(Me.ClientRectangle.Width)
		TempHigh = VB6.PixelsToTwipsY(Me.ClientRectangle.Height)
		
		DoNotChange = True
		
		If UnitType = Metric Then
			labPowderUnits(0).Text = "kilograms/day"
			labPowderUnits(6).Text = "kilograms/day"
			PoundConversion = 0.4536
		Else
			labPowderUnits(0).Text = "pounds/day"
			labPowderUnits(6).Text = "pounds/day"
			PoundConversion = 1
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
	'UPGRADE_WARNING: Event frmPowderData.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub frmPowderData_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		TempWide = VB6.PixelsToTwipsX(Me.ClientRectangle.Width)
		TempHigh = VB6.PixelsToTwipsY(Me.ClientRectangle.Height)
		Call screenstuff()
	End Sub
	Private Sub frmPowderData_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
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
	Private Sub labPowderHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles labPowderHelp.Click
		Dim StartHelp As Short
		StartHelp = 223
		IsHelpOn = True
		Call frmSurfaceHelp.gethelptext(StartHelp, WhichCell)
		frmSurfaceHelp.Show()
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
			txtPowderValues(WhichCell).Focus()
		End If
	End Sub
	Private Sub txtPowderValues_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPowderValues.Enter
		Dim Index As Short = txtPowderValues.GetIndex(eventSender)
        'Dim x As Short
        WhichCell = Index
		
		System.Windows.Forms.SendKeys.Send("{HOME}+{END}")
		
		WhichCell = Index
		Call drawthevalues()
	End Sub
	'UPGRADE_WARNING: Event txtPowderValues.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtPowderValues_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPowderValues.TextChanged
		Dim Index As Short = txtPowderValues.GetIndex(eventSender)
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
		
		q = (2880 / h) * TempHigh
		r = (1560 / h) * TempHigh
		s = (60 / h) * TempHigh
		t = (120 / w) * TempWide
		u = (540 / h) * TempHigh
		v = (900 / w) * TempWide
		
		y = (420 / h) * TempHigh
		z = (300 / w) * TempWide
		
		For x = 0 To 1
			labPowderHeading(x).Top = VB6.TwipsToPixelsY((TempHigh * (180 / h)) + (x * u))
			labPowderHeading(x).Left = VB6.TwipsToPixelsX((TempWide * (180 / w)) + (x * v))
			labPowderHeading(x).Width = VB6.TwipsToPixelsX(TempWide * (1965 / w))
			labPowderLabels(x).Top = VB6.TwipsToPixelsY((TempHigh * (2520 / h)) + (x * r))
			labPowderLabels(x).Left = VB6.TwipsToPixelsX(TempWide * (600 / w))
			labPowderLabels(x).Width = VB6.TwipsToPixelsX(TempWide * (1935 / w))
		Next x
		
		For x = 2 To 3
			labPowderLabels(x).Left = VB6.TwipsToPixelsX(TempWide * (3240 / w))
			labPowderLabels(x).Top = VB6.TwipsToPixelsY((TempHigh * (180 / h)) + ((x - 2) * q))
		Next x
		
		For x = 0 To 5
			optSegment(x).Top = VB6.TwipsToPixelsY(TempHigh * (2820 / h))
			optSegment(x).Left = VB6.TwipsToPixelsX((TempWide * (720 / w)) + (x * z))
			optSegment(x).Width = VB6.TwipsToPixelsX(TempWide * (195 / w))
			labSegment(x).Top = VB6.TwipsToPixelsY(TempHigh * (3060 / h))
			labSegment(x).Left = VB6.TwipsToPixelsX((TempWide * (720 / w)) + (x * z))
			labSegment(x).Width = VB6.TwipsToPixelsX(TempWide * (195 / w))
		Next x
		
		txtSegmentLabel.Top = VB6.TwipsToPixelsY(TempHigh * (4380 / h))
		txtSegmentLabel.Left = VB6.TwipsToPixelsX(TempWide * (600 / w))
		txtSegmentLabel.Width = VB6.TwipsToPixelsX(TempWide * (1935 / w))
		
		For x = 0 To 11
			LabPowderTitles(x).Left = VB6.TwipsToPixelsX(TempWide * (3960 / w))
			LabPowderTitles(x).Width = VB6.TwipsToPixelsX(TempWide * (1935 / w))
			txtPowderValues(x).Left = VB6.TwipsToPixelsX(TempWide * (6060 / w))
			txtPowderValues(x).Width = VB6.TwipsToPixelsX(TempWide * (1035 / w))
			labPowderUnits(x).Left = VB6.TwipsToPixelsX(TempWide * (7140 / w))
		Next x
		
		For x = 0 To 5
			LabPowderTitles(x).Top = VB6.TwipsToPixelsY((TempHigh * (540 / h)) + (x * y))
			txtPowderValues(x).Top = VB6.TwipsToPixelsY(TempHigh * (510 / h) + (x * y))
			labPowderUnits(x).Top = VB6.TwipsToPixelsY((TempHigh * (540 / h)) + (x * y))
		Next x
		
		For x = 6 To 11
			LabPowderTitles(x).Top = VB6.TwipsToPixelsY((TempHigh * (3420 / h)) + ((x - 6) * y))
			txtPowderValues(x).Top = VB6.TwipsToPixelsY((TempHigh * (3390 / h)) + ((x - 6) * y))
			labPowderUnits(x).Top = VB6.TwipsToPixelsY((TempHigh * (3420 / h)) + ((x - 6) * y))
		Next x
		
		LineHorizontal(0).X1 = VB6.TwipsToPixelsX(TempWide * (5880 / w))
		LineHorizontal(0).X2 = VB6.TwipsToPixelsX(TempWide * (8820 / w))
		LineHorizontal(0).Y1 = VB6.TwipsToPixelsY(TempHigh * (240 / h))
		LineHorizontal(0).Y2 = VB6.TwipsToPixelsY(TempHigh * (240 / h))
		
		LineHorizontal(1).X1 = VB6.TwipsToPixelsX(TempWide * (6120 / w))
		LineHorizontal(1).X2 = VB6.TwipsToPixelsX(TempWide * (8700 / w))
		LineHorizontal(1).Y1 = VB6.TwipsToPixelsY(TempHigh * (3120 / h))
		LineHorizontal(1).Y2 = VB6.TwipsToPixelsY(TempHigh * (3120 / h))
		
		LineHorizontal(2).X1 = VB6.TwipsToPixelsX(TempWide * (3240 / w))
		LineHorizontal(2).X2 = VB6.TwipsToPixelsX(TempWide * (8820 / w))
		LineHorizontal(2).Y1 = VB6.TwipsToPixelsY(TempHigh * (6000 / h))
		LineHorizontal(2).Y2 = VB6.TwipsToPixelsY(TempHigh * (6000 / h))
		
		LineVertical(0).X1 = VB6.TwipsToPixelsX(TempWide * (3300 / w))
		LineVertical(0).X2 = VB6.TwipsToPixelsX(TempWide * (3300 / w))
		LineVertical(0).Y1 = VB6.TwipsToPixelsY(TempHigh * (480 / h))
		LineVertical(0).Y2 = VB6.TwipsToPixelsY(TempHigh * (3000 / h))
		
		LineVertical(1).X1 = VB6.TwipsToPixelsX(TempWide * (3300 / w))
		LineVertical(1).X2 = VB6.TwipsToPixelsX(TempWide * (3300 / w))
		LineVertical(1).Y1 = VB6.TwipsToPixelsY(TempHigh * (3360 / h))
		LineVertical(1).Y2 = VB6.TwipsToPixelsY(TempHigh * (6060 / h))
		
		LineVertical(2).X1 = VB6.TwipsToPixelsX(TempWide * (8760 / w))
		LineVertical(2).X2 = VB6.TwipsToPixelsX(TempWide * (8760 / w))
		LineVertical(2).Y1 = VB6.TwipsToPixelsY(TempHigh * (180 / h))
		LineVertical(2).Y2 = VB6.TwipsToPixelsY(TempHigh * (6060 / h))
		
		labBackToMenu.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labBackToMenu.Left = VB6.TwipsToPixelsX(TempWide * (600 / w))
		
		imgBackToMenu.Top = VB6.TwipsToPixelsY(TempHigh * (6150 / h))
		imgBackToMenu.Left = VB6.TwipsToPixelsX(TempWide * (60 / w))
		imgBackToMenu.Width = VB6.TwipsToPixelsX(TempWide * (495 / w))
		
		comPowderPrint.Top = VB6.TwipsToPixelsY(TempHigh * (6150 / h))
		comPowderPrint.Left = VB6.TwipsToPixelsX(TempWide * (8880 / w))
		
		labPowderHelp.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labPowderHelp.Left = VB6.TwipsToPixelsX(TempWide * (8340 / w))
		
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
				txtPowderValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF0000)
			Else
				txtPowderValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
			End If
			Select Case i
				Case 0, 6
					txtPowderValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value * PoundConversion)), "#,###,###,##0")
				Case Else
					txtPowderValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value)), "##,###,##0")
			End Select
		Next i
		
		DoNotChange = False
		
	End Sub
	Private Sub txtPowderValues_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtPowderValues.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Dim Index As Short = txtPowderValues.GetIndex(eventSender)
		
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
						If InStr(txtPowderValues(Index).Text, ".") = 0 Then
							System.Windows.Forms.SendKeys.Send("{DELETE}")
						End If
					Else
						System.Windows.Forms.SendKeys.Send("{DELETE}")
					End If
			End Select
		End If
		
	End Sub
	Private Sub txtPowderValues_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtPowderValues.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = txtPowderValues.GetIndex(eventSender)
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
	Private Sub txtPowderValues_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPowderValues.Leave
		Dim Index As Short = txtPowderValues.GetIndex(eventSender)
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
		For i = 1 To Len(txtPowderValues(Sample).Text)
			Digit.Value = Mid(txtPowderValues(Sample).Text, i, 1)
			Select Case Digit.Value
				Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ".", "-"
					tempvalue = tempvalue & Digit.Value
			End Select
		Next i
		Select Case Sample
			Case 0, 6
				If CellValues(WhichScreen, Sample, WhichSegment).Changed = True Then
					If PoundConversion <> 0 Then CellValues(WhichScreen, Sample, WhichSegment).Value = Val(tempvalue) / PoundConversion
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