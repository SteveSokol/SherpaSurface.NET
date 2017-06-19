Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks
Friend Class frmClearingData
	Inherits System.Windows.Forms.Form
	Dim TempHigh As Single
	Dim TempWide As Single
	Dim AcreConversion As Single
	Private Sub comClearingPrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles comClearingPrint.Click
		PrintForm1.Print(Me, PowerPacks.Printing.PrintForm.PrintOption.CompatibleModeClientAreaOnly)
	End Sub
	'UPGRADE_WARNING: Form event frmClearingData.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmClearingData_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated

        'Dim baseunit As String
        'Dim baselength As Short
        Dim i As Short
		
		If IsHelpOn = True Then
			txtClearingValues(WhichCell).Focus()
			IsHelpOn = False
		Else
			For i = 0 To 5
				If optSegment(i).Checked = True Then WhichSegment = i
			Next i
			
			WhichScreen = Clearing
			
			Call drawthevalues()
			
			If InsertFlag = True Then
				labInsert.Text = "Insert"
			Else
				labInsert.Text = "Typeover"
			End If
			
			WhichCell = 0
			
			txtClearingValues(0).Focus()
		End If
		
	End Sub
	Private Sub frmClearingData_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        'Dim x As Short

        Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - (VB6.PixelsToTwipsY(Me.Height) + 350)) / 2)
		Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2)
		
		If VB6.PixelsToTwipsY(Me.Top) < 0 Then Me.Top = 0
		If VB6.PixelsToTwipsX(Me.Left) < 0 Then Me.Left = 0
		
		TempWide = VB6.PixelsToTwipsX(Me.ClientRectangle.Width)
		TempHigh = VB6.PixelsToTwipsY(Me.ClientRectangle.Height)
		
		DoNotChange = True
		
		If UnitType = Metric Then
			labClearingUnits(0).Text = "hectares"
			AcreConversion = 0.4047
		Else
			labClearingUnits(0).Text = "acres"
			AcreConversion = 1
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
	'UPGRADE_WARNING: Event frmClearingData.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub frmClearingData_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		TempWide = VB6.PixelsToTwipsX(Me.ClientRectangle.Width)
		TempHigh = VB6.PixelsToTwipsY(Me.ClientRectangle.Height)
		Call screenstuff()
	End Sub
	Private Sub frmClearingData_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
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
	
	Private Sub labClearingHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles labClearingHelp.Click
		Dim StartHelp As Short
		StartHelp = 250
		IsHelpOn = True
		Call frmSurfaceHelp.gethelptext(StartHelp, WhichCell)
		frmSurfaceHelp.Show()
	End Sub
	'UPGRADE_WARNING: Event lstClearingList.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub lstClearingList_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstClearingList.SelectedIndexChanged
        'Dim x As Short

        txtClearingValues(WhichCell).Text = LTrim(RTrim(VB6.GetItemString(lstClearingList, lstClearingList.SelectedIndex)))
		CellValues(WhichScreen, WhichCell, WhichSegment).Changed = True
		
		Call Inputer(WhichCell)
		
		txtClearingValues(0).Focus()
		
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
			txtClearingValues(WhichCell).Focus()
		End If
	End Sub
	Private Sub txtClearingValues_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtClearingValues.Enter
		Dim Index As Short = txtClearingValues.GetIndex(eventSender)
        'Dim x As Short
        WhichCell = Index
		
		lstClearingList.Visible = False
		
		If WhichCell = 6 Then
			lstClearingList.Visible = True
		End If
		
		System.Windows.Forms.SendKeys.Send("{HOME}+{END}")
		
		WhichCell = Index
		Call drawthevalues()
	End Sub
	'UPGRADE_WARNING: Event txtClearingValues.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtClearingValues_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtClearingValues.TextChanged
		Dim Index As Short = txtClearingValues.GetIndex(eventSender)
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
		
		r = (1140 / h) * TempHigh
		s = (60 / h) * TempHigh
		t = (120 / w) * TempWide
		u = (540 / h) * TempHigh
		v = (900 / w) * TempWide
		
		y = (420 / h) * TempHigh
		z = (300 / w) * TempWide
		
		labClearingHeading.Top = VB6.TwipsToPixelsY(TempHigh * (240 / h))
		labClearingHeading.Left = VB6.TwipsToPixelsX(TempWide * (240 / w))
		labClearingHeading.Width = VB6.TwipsToPixelsX(TempWide * (3765 / w))
		
		For x = 0 To 1
			labClearingLabels(x).Top = VB6.TwipsToPixelsY((TempHigh * (1560 / h)) + (x * r))
			labClearingLabels(x).Left = VB6.TwipsToPixelsX(TempWide * (600 / w))
			labClearingLabels(x).Width = VB6.TwipsToPixelsX(TempWide * (1935 / w))
		Next x
		
		labClearingLabels(2).Top = VB6.TwipsToPixelsY(TempHigh * (3720 / h))
		labClearingLabels(2).Left = VB6.TwipsToPixelsX(TempWide * (600 / w))
		
		For x = 3 To 5
			labClearingLabels(x).Left = VB6.TwipsToPixelsX(TempWide * (3240 / w))
		Next x
		
		labClearingLabels(3).Top = VB6.TwipsToPixelsY(TempHigh * (1140 / h))
		labClearingLabels(4).Top = VB6.TwipsToPixelsY(TempHigh * (2040 / h))
		labClearingLabels(5).Top = VB6.TwipsToPixelsY(TempHigh * (4620 / h))
		
		For x = 0 To 5
			optSegment(x).Top = VB6.TwipsToPixelsY(TempHigh * (1860 / h))
			optSegment(x).Left = VB6.TwipsToPixelsX((TempWide * (720 / w)) + (x * z))
			optSegment(x).Width = VB6.TwipsToPixelsX(TempWide * (195 / w))
			labSegment(x).Top = VB6.TwipsToPixelsY(TempHigh * (2100 / h))
			labSegment(x).Left = VB6.TwipsToPixelsX((TempWide * (720 / w)) + (x * z))
			labSegment(x).Width = VB6.TwipsToPixelsX(TempWide * (195 / w))
		Next x
		
		txtSegmentLabel.Top = VB6.TwipsToPixelsY(TempHigh * (3000 / h))
		txtSegmentLabel.Left = VB6.TwipsToPixelsX(TempWide * (600 / w))
		txtSegmentLabel.Width = VB6.TwipsToPixelsX(TempWide * (1935 / w))
		
		lstClearingList.Top = VB6.TwipsToPixelsY(TempHigh * (4140 / h))
		lstClearingList.Left = VB6.TwipsToPixelsX(TempWide * (1080 / w))
		lstClearingList.Height = VB6.TwipsToPixelsY(TempHigh * (960 / h))
		lstClearingList.Width = VB6.TwipsToPixelsX(TempWide * (1755 / w))
		
		For x = 0 To 5
			LabClearingTitles(x).Left = VB6.TwipsToPixelsX(TempWide * (3540 / w))
			LabClearingTitles(x).Width = VB6.TwipsToPixelsX(TempWide * (2355 / w))
			txtClearingValues(x).Left = VB6.TwipsToPixelsX(TempWide * (6060 / w))
			txtClearingValues(x).Width = VB6.TwipsToPixelsX(TempWide * (1035 / w))
			labClearingUnits(x).Left = VB6.TwipsToPixelsX(TempWide * (7140 / w))
		Next x
		
		LabClearingTitles(0).Top = VB6.TwipsToPixelsY(TempHigh * (1560 / h))
		txtClearingValues(0).Top = VB6.TwipsToPixelsY(TempHigh * (1530 / h))
		labClearingUnits(0).Top = VB6.TwipsToPixelsY(TempHigh * (1560 / h))
		
		For x = 1 To 5
			LabClearingTitles(x).Top = VB6.TwipsToPixelsY((TempHigh * (2460 / h)) + ((x - 1) * y))
			txtClearingValues(x).Top = VB6.TwipsToPixelsY((TempHigh * (2430 / h)) + ((x - 1) * y))
			labClearingUnits(x).Top = VB6.TwipsToPixelsY((TempHigh * (2460 / h)) + ((x - 1) * y))
		Next x
		
		LabClearingTitles(6).Top = VB6.TwipsToPixelsY(TempHigh * (5040 / h))
		LabClearingTitles(6).Left = VB6.TwipsToPixelsX(TempWide * (3540 / w))
		LabClearingTitles(6).Width = VB6.TwipsToPixelsX(TempWide * (1875 / w))
		txtClearingValues(6).Top = VB6.TwipsToPixelsY(TempHigh * (5010 / h))
		txtClearingValues(6).Left = VB6.TwipsToPixelsX(TempWide * (5580 / w))
		txtClearingValues(6).Width = VB6.TwipsToPixelsX(TempWide * (1995 / w))
		labClearingUnits(6).Top = VB6.TwipsToPixelsY(TempHigh * (5040 / h))
		labClearingUnits(6).Left = VB6.TwipsToPixelsX(TempWide * (7620 / w))
		
		LineHorizontal(0).X1 = VB6.TwipsToPixelsX(TempWide * (1860 / w))
		LineHorizontal(0).X2 = VB6.TwipsToPixelsX(TempWide * (3240 / w))
		LineHorizontal(0).Y1 = VB6.TwipsToPixelsY(TempHigh * (3780 / h))
		LineHorizontal(0).Y2 = VB6.TwipsToPixelsY(TempHigh * (3780 / h))
		
		LineHorizontal(1).X1 = VB6.TwipsToPixelsX(TempWide * (600 / w))
		LineHorizontal(1).X2 = VB6.TwipsToPixelsX(TempWide * (3240 / w))
		LineHorizontal(1).Y1 = VB6.TwipsToPixelsY(TempHigh * (5280 / h))
		LineHorizontal(1).Y2 = VB6.TwipsToPixelsY(TempHigh * (5280 / h))
		
		LineHorizontal(2).X1 = VB6.TwipsToPixelsX(TempWide * (3660 / w))
		LineHorizontal(2).X2 = VB6.TwipsToPixelsX(TempWide * (8820 / w))
		LineHorizontal(2).Y1 = VB6.TwipsToPixelsY(TempHigh * (1200 / h))
		LineHorizontal(2).Y2 = VB6.TwipsToPixelsY(TempHigh * (1200 / h))
		
		LineHorizontal(3).X1 = VB6.TwipsToPixelsX(TempWide * (4860 / w))
		LineHorizontal(3).X2 = VB6.TwipsToPixelsX(TempWide * (8700 / w))
		LineHorizontal(3).Y1 = VB6.TwipsToPixelsY(TempHigh * (2100 / h))
		LineHorizontal(3).Y2 = VB6.TwipsToPixelsY(TempHigh * (2100 / h))
		
		LineHorizontal(4).X1 = VB6.TwipsToPixelsX(TempWide * (4800 / w))
		LineHorizontal(4).X2 = VB6.TwipsToPixelsX(TempWide * (8700 / w))
		LineHorizontal(4).Y1 = VB6.TwipsToPixelsY(TempHigh * (4680 / h))
		LineHorizontal(4).Y2 = VB6.TwipsToPixelsY(TempHigh * (4680 / h))
		
		LineHorizontal(5).X1 = VB6.TwipsToPixelsX(TempWide * (3240 / w))
		LineHorizontal(5).X2 = VB6.TwipsToPixelsX(TempWide * (8820 / w))
		LineHorizontal(5).Y1 = VB6.TwipsToPixelsY(TempHigh * (5580 / h))
		LineHorizontal(5).Y2 = VB6.TwipsToPixelsY(TempHigh * (5580 / h))
		
		LineVertical(0).X1 = VB6.TwipsToPixelsX(TempWide * (660 / w))
		LineVertical(0).X2 = VB6.TwipsToPixelsX(TempWide * (660 / w))
		LineVertical(0).Y1 = VB6.TwipsToPixelsY(TempHigh * (4020 / h))
		LineVertical(0).Y2 = VB6.TwipsToPixelsY(TempHigh * (5340 / h))
		
		LineVertical(1).X1 = VB6.TwipsToPixelsX(TempWide * (3300 / w))
		LineVertical(1).X2 = VB6.TwipsToPixelsX(TempWide * (3300 / w))
		LineVertical(1).Y1 = VB6.TwipsToPixelsY(TempHigh * (1440 / h))
		LineVertical(1).Y2 = VB6.TwipsToPixelsY(TempHigh * (1980 / h))
		
		LineVertical(2).X1 = VB6.TwipsToPixelsX(TempWide * (3300 / w))
		LineVertical(2).X2 = VB6.TwipsToPixelsX(TempWide * (3300 / w))
		LineVertical(2).Y1 = VB6.TwipsToPixelsY(TempHigh * (2340 / h))
		LineVertical(2).Y2 = VB6.TwipsToPixelsY(TempHigh * (4560 / h))
		
		LineVertical(3).X1 = VB6.TwipsToPixelsX(TempWide * (3300 / w))
		LineVertical(3).X2 = VB6.TwipsToPixelsX(TempWide * (3300 / w))
		LineVertical(3).Y1 = VB6.TwipsToPixelsY(TempHigh * (4920 / h))
		LineVertical(3).Y2 = VB6.TwipsToPixelsY(TempHigh * (5640 / h))
		
		LineVertical(4).X1 = VB6.TwipsToPixelsX(TempWide * (8760 / w))
		LineVertical(4).X2 = VB6.TwipsToPixelsX(TempWide * (8760 / w))
		LineVertical(4).Y1 = VB6.TwipsToPixelsY(TempHigh * (1140 / h))
		LineVertical(4).Y2 = VB6.TwipsToPixelsY(TempHigh * (5640 / h))
		
		labBackToMenu.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labBackToMenu.Left = VB6.TwipsToPixelsX(TempWide * (600 / w))
		
		imgBackToMenu.Top = VB6.TwipsToPixelsY(TempHigh * (6150 / h))
		imgBackToMenu.Left = VB6.TwipsToPixelsX(TempWide * (60 / w))
		imgBackToMenu.Width = VB6.TwipsToPixelsX(TempWide * (495 / w))
		
		comClearingPrint.Top = VB6.TwipsToPixelsY(TempHigh * (6150 / h))
		comClearingPrint.Left = VB6.TwipsToPixelsX(TempWide * (8880 / w))
		
		labClearingHelp.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labClearingHelp.Left = VB6.TwipsToPixelsX(TempWide * (8340 / w))
		
		labInsert.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labInsert.Left = VB6.TwipsToPixelsX(TempWide * (4080 / w))
		labInsert.Width = VB6.TwipsToPixelsX(TempWide * (975 / w))
		
	End Sub
	Public Sub drawthevalues()
		
		Dim i As Short
        'Dim x As Short

        DoNotChange = True
		
		For i = 0 To 6
			If CellValues(WhichScreen, i, WhichSegment).Changed = True Then
				txtClearingValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF0000)
			Else
				txtClearingValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
			End If
			Select Case i
				Case 0
					txtClearingValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value * AcreConversion)), "#,###,###,##0.00")
				Case 6
					txtClearingValues(i).Text = LTrim(RTrim(CellValues(WhichScreen, i, WhichSegment).Word))
				Case Else
					txtClearingValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value)), "##,###,##0.0")
			End Select
		Next i
		
		DoNotChange = False
		
	End Sub
	Private Sub txtClearingValues_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtClearingValues.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Dim Index As Short = txtClearingValues.GetIndex(eventSender)
		
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
						If InStr(txtClearingValues(Index).Text, ".") = 0 Then
							System.Windows.Forms.SendKeys.Send("{DELETE}")
						End If
					Else
						System.Windows.Forms.SendKeys.Send("{DELETE}")
					End If
			End Select
		End If
		
	End Sub
	Private Sub txtClearingValues_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtClearingValues.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = txtClearingValues.GetIndex(eventSender)
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
	Private Sub txtClearingValues_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtClearingValues.Leave
		Dim Index As Short = txtClearingValues.GetIndex(eventSender)
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
		For i = 1 To Len(txtClearingValues(Sample).Text)
			Digit.Value = Mid(txtClearingValues(Sample).Text, i, 1)
			Select Case Digit.Value
				Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ".", "-"
					tempvalue = tempvalue & Digit.Value
			End Select
		Next i
		Select Case Sample
			Case 0
				If CellValues(WhichScreen, Sample, WhichSegment).Changed = True Then
					If AcreConversion <> 0 Then CellValues(WhichScreen, Sample, WhichSegment).Value = Val(tempvalue) / AcreConversion
				End If
			Case 6
				If CellValues(WhichScreen, Sample, WhichSegment).Changed = True Then
					CellValues(WhichScreen, Sample, WhichSegment).Word = LTrim(RTrim(txtClearingValues(Sample).Text))
				End If
			Case Else
				If CellValues(WhichScreen, Sample, WhichSegment).Changed = True Then
					CellValues(WhichScreen, Sample, WhichSegment).Value = Val(tempvalue)
				End If
		End Select
		Call ClearEngr()
		Call drawthevalues()
	End Sub
	'UPGRADE_WARNING: Event txtSegmentLabel.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtSegmentLabel_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSegmentLabel.TextChanged
		If DoNotChange = True Then Exit Sub
		SegNamie(WhichSegment) = txtSegmentLabel.Text
	End Sub
End Class