Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks
Friend Class frmAncPurchaseData
	Inherits System.Windows.Forms.Form
	Dim TempHigh As Single
	Dim TempWide As Single
	Dim FootConversion As Single
	Dim DensityConversion As Single
	Dim PowderConversion As Single
	Private Sub comAncPurchasePrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles comAncPurchasePrint.Click
		PrintForm1.Print(Me, PowerPacks.Printing.PrintForm.PrintOption.CompatibleModeClientAreaOnly)
	End Sub
	'UPGRADE_WARNING: Form event frmAncPurchaseData.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmAncPurchaseData_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated

        'Dim baseunit As String
        'Dim baselength As Short
        Dim i As Short
		
		If IsHelpOn = True Then
			txtAncPurchaseValues(WhichCell).Focus()
			IsHelpOn = False
		Else
			For i = 0 To 5
				If optSegment(i).Checked = True Then WhichSegment = i
			Next i
			
			WhichScreen = Purchase
			
			Call drawthevalues()
			
			If InsertFlag = True Then
				labInsert.Text = "Insert"
			Else
				labInsert.Text = "Typeover"
			End If
			
			WhichCell = 0
			
			txtAncPurchaseValues(12).Focus()
		End If
		
	End Sub
	Private Sub frmAncPurchaseData_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        'Dim x As Short

        Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - (VB6.PixelsToTwipsY(Me.Height) + 350)) / 2)
		Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2)
		
		If VB6.PixelsToTwipsY(Me.Top) < 0 Then Me.Top = 0
		If VB6.PixelsToTwipsX(Me.Left) < 0 Then Me.Left = 0
		
		TempWide = VB6.PixelsToTwipsX(Me.ClientRectangle.Width)
		TempHigh = VB6.PixelsToTwipsY(Me.ClientRectangle.Height)
		
		DoNotChange = True
		
		If UnitType = Metric Then
		Else
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
	'UPGRADE_WARNING: Event frmAncPurchaseData.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub frmAncPurchaseData_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		TempWide = VB6.PixelsToTwipsX(Me.ClientRectangle.Width)
		TempHigh = VB6.PixelsToTwipsY(Me.ClientRectangle.Height)
		Call screenstuff()
	End Sub
	Private Sub frmAncPurchaseData_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
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
	Private Sub labAncPurchaseHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles labAncPurchaseHelp.Click
		Dim StartHelp As Short
		StartHelp = 156
		IsHelpOn = True
		Call frmSurfaceHelp.gethelptext(StartHelp, 0)
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
			txtAncPurchaseValues(WhichCell).Focus()
		End If
	End Sub
	'UPGRADE_WARNING: Event txtAncPurchaseValues.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtAncPurchaseValues_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAncPurchaseValues.TextChanged
		Dim Index As Short = txtAncPurchaseValues.GetIndex(eventSender)
		If DoNotChange = True Then Exit Sub
	End Sub
	Private Sub txtAncPurchaseValues_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAncPurchaseValues.Enter
		Dim Index As Short = txtAncPurchaseValues.GetIndex(eventSender)
		WhichCell = Index
		Call drawthevalues()
	End Sub
	Public Sub screenstuff()
		
		Dim p As Decimal
		Dim q As Decimal
		Dim r As Decimal
		Dim s As Decimal
		Dim t As Decimal
		Dim u As Decimal
		
		Dim x As Short
		
		Dim y As Decimal
		Dim z As Decimal
		
		Dim h As Short
		Dim w As Short
		
		h = 6420
		w = 9150
		
		
		p = (480 / h) * TempHigh
		q = (300 / w) * TempWide
		r = (120 / h) * TempHigh
		s = (120 / w) * TempWide
		t = (1020 / h) * TempHigh
		u = (1500 / h) * TempHigh
		
		y = (420 / h) * TempHigh
		z = (300 / w) * TempWide
		
		For x = 0 To 2
			labAncPurchaseHeading(x).Top = VB6.TwipsToPixelsY((TempHigh * (60 / h)) + (x * p))
			labAncPurchaseHeading(x).Left = VB6.TwipsToPixelsX((TempWide * (60 / w)) + (x * q))
			labAncPurchaseHeading(x).Width = VB6.TwipsToPixelsX(TempWide * (2175 / w))
			If x < 2 Then
				labAncPurchaseLabels(x).Top = VB6.TwipsToPixelsY((TempHigh * (2340 / h)) + (x * u))
				labAncPurchaseLabels(x).Left = VB6.TwipsToPixelsX(TempWide * (540 / w))
			End If
		Next x
		
		labAncPurchaseLabels(2).Top = VB6.TwipsToPixelsY(TempHigh * (1320 / h))
		labAncPurchaseLabels(2).Left = VB6.TwipsToPixelsX(TempWide * (3000 / w))
		
		For x = 3 To 5
			labAncPurchaseLabels(x).Top = VB6.TwipsToPixelsY(TempHigh * (1080 / h))
			If x = 3 Then
				labAncPurchaseLabels(x).Left = VB6.TwipsToPixelsX(TempWide * (3720 / w))
				labAncPurchaseLabels(x).Width = VB6.TwipsToPixelsX(TempWide * (1665 / w))
			ElseIf x = 4 Then 
				labAncPurchaseLabels(x).Left = VB6.TwipsToPixelsX(TempWide * (5580 / w))
				labAncPurchaseLabels(x).Width = VB6.TwipsToPixelsX(TempWide * (1335 / w))
			Else
				labAncPurchaseLabels(x).Left = VB6.TwipsToPixelsX(TempWide * (7200 / w))
				labAncPurchaseLabels(x).Width = VB6.TwipsToPixelsX(TempWide * (1575 / w))
			End If
		Next x
		
		For x = 0 To 5
			optSegment(x).Top = VB6.TwipsToPixelsY(TempHigh * (2640 / h))
			optSegment(x).Left = VB6.TwipsToPixelsX((TempWide * (660 / w)) + (x * z))
			optSegment(x).Width = VB6.TwipsToPixelsX(TempWide * (195 / w))
			labSegment(x).Top = VB6.TwipsToPixelsY(TempHigh * (2880 / h))
			labSegment(x).Left = VB6.TwipsToPixelsX((TempWide * (660 / w)) + (x * z))
			labSegment(x).Width = VB6.TwipsToPixelsX(TempWide * (195 / w))
		Next x
		
		txtSegmentLabel.Top = VB6.TwipsToPixelsY(TempHigh * (4140 / h))
		txtSegmentLabel.Left = VB6.TwipsToPixelsX(TempWide * (540 / w))
		txtSegmentLabel.Width = VB6.TwipsToPixelsX(TempWide * (1935 / w))
		
		For x = 0 To 7
			LabAncPurchaseTitles(x + 6).Top = VB6.TwipsToPixelsY((TempHigh * (1635 / h)) + (x * y))
			LabAncPurchaseTitles(x + 6).Left = VB6.TwipsToPixelsX(TempWide * (3720 / w))
			LabAncPurchaseTitles(x + 6).Width = VB6.TwipsToPixelsX(TempWide * (1635 / w))
			txtAncPurchaseValues(x + 6).Top = VB6.TwipsToPixelsY((TempHigh * (1605 / h)) + (x * y))
			txtAncPurchaseValues(x + 6).Left = VB6.TwipsToPixelsX(TempWide * (5580 / w))
			txtAncPurchaseValues(x + 6).Width = VB6.TwipsToPixelsX(TempWide * (1335 / w))
			txtAncPurchaseValues(x + 26).Top = VB6.TwipsToPixelsY((TempHigh * (1605 / h)) + (x * y))
			txtAncPurchaseValues(x + 26).Left = VB6.TwipsToPixelsX(TempWide * (7260 / w))
			txtAncPurchaseValues(x + 26).Width = VB6.TwipsToPixelsX(TempWide * (1575 / w))
		Next x
		
		LineHorizontal(0).X1 = VB6.TwipsToPixelsX(TempWide * (3000 / w))
		LineHorizontal(0).X2 = VB6.TwipsToPixelsX(TempWide * (9060 / w))
		LineHorizontal(0).Y1 = VB6.TwipsToPixelsY(TempHigh * (960 / h))
		LineHorizontal(0).Y2 = VB6.TwipsToPixelsY(TempHigh * (960 / h))
		
		LineHorizontal(1).X1 = VB6.TwipsToPixelsX(TempWide * (3780 / w))
		LineHorizontal(1).X2 = VB6.TwipsToPixelsX(TempWide * (8940 / w))
		LineHorizontal(1).Y1 = VB6.TwipsToPixelsY(TempHigh * (1380 / h))
		LineHorizontal(1).Y2 = VB6.TwipsToPixelsY(TempHigh * (1380 / h))
		
		LineHorizontal(2).X1 = VB6.TwipsToPixelsX(TempWide * (3000 / w))
		LineHorizontal(2).X2 = VB6.TwipsToPixelsX(TempWide * (9060 / w))
		LineHorizontal(2).Y1 = VB6.TwipsToPixelsY(TempHigh * (5040 / h))
		LineHorizontal(2).Y2 = VB6.TwipsToPixelsY(TempHigh * (5040 / h))
		
		LineVertical(0).X1 = VB6.TwipsToPixelsX(TempWide * (3060 / w))
		LineVertical(0).X2 = VB6.TwipsToPixelsX(TempWide * (3060 / w))
		LineVertical(0).Y1 = VB6.TwipsToPixelsY(TempHigh * (900 / h))
		LineVertical(0).Y2 = VB6.TwipsToPixelsY(TempHigh * (1260 / h))
		
		LineVertical(1).X1 = VB6.TwipsToPixelsX(TempWide * (3060 / w))
		LineVertical(1).X2 = VB6.TwipsToPixelsX(TempWide * (3060 / w))
		LineVertical(1).Y1 = VB6.TwipsToPixelsY(TempHigh * (1620 / h))
		LineVertical(1).Y2 = VB6.TwipsToPixelsY(TempHigh * (5100 / h))
		
		LineVertical(2).X1 = VB6.TwipsToPixelsX(TempWide * (7080 / w))
		LineVertical(2).X2 = VB6.TwipsToPixelsX(TempWide * (7080 / w))
		LineVertical(2).Y1 = VB6.TwipsToPixelsY(TempHigh * (1440 / h))
		LineVertical(2).Y2 = VB6.TwipsToPixelsY(TempHigh * (4980 / h))
		
		LineVertical(3).X1 = VB6.TwipsToPixelsX(TempWide * (9000 / w))
		LineVertical(3).X2 = VB6.TwipsToPixelsX(TempWide * (9000 / w))
		LineVertical(3).Y1 = VB6.TwipsToPixelsY(TempHigh * (900 / h))
		LineVertical(3).Y2 = VB6.TwipsToPixelsY(TempHigh * (5100 / h))
		
		labBackToMenu.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labBackToMenu.Left = VB6.TwipsToPixelsX(TempWide * (600 / w))
		
		imgBackToMenu.Top = VB6.TwipsToPixelsY(TempHigh * (6150 / h))
		imgBackToMenu.Left = VB6.TwipsToPixelsX(TempWide * (60 / w))
		imgBackToMenu.Width = VB6.TwipsToPixelsX(TempWide * (495 / w))
		
		comAncPurchasePrint.Top = VB6.TwipsToPixelsY(TempHigh * (6150 / h))
		comAncPurchasePrint.Left = VB6.TwipsToPixelsX(TempWide * (8880 / w))
		
		labAncPurchaseHelp.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labAncPurchaseHelp.Left = VB6.TwipsToPixelsX(TempWide * (8340 / w))
		
		labInsert.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labInsert.Left = VB6.TwipsToPixelsX(TempWide * (4080 / w))
		labInsert.Width = VB6.TwipsToPixelsX(TempWide * (975 / w))
		
	End Sub
	Public Sub drawthevalues()
		
		Dim i As Short
        'Dim x As Short

        DoNotChange = True
		Call ScreenCalc()
		For i = 6 To 13
			If CellValues(WhichScreen, i, WhichSegment).Changed = True Then
				txtAncPurchaseValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF0000)
			Else
				txtAncPurchaseValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
			End If
			If CellValues(WhichScreen, i + 20, WhichSegment).Changed = True Then
				txtAncPurchaseValues(i + 20).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF0000)
			Else
				txtAncPurchaseValues(i + 20).ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
			End If
			txtAncPurchaseValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value)), "$###,###,##0")
			txtAncPurchaseValues(i + 20).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i + 20, WhichSegment).Value)), "$###,###,##0")
		Next i
		
		If bltp = 1 Then
			LabAncPurchaseTitles(10).Text = "Powder Buggies"
		ElseIf bltp = 2 Then 
			LabAncPurchaseTitles(10).Text = "Bulk Trucks"
		End If
		
		DoNotChange = False
		
	End Sub
	
	Private Sub txtAncPurchaseValues_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtAncPurchaseValues.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Dim Index As Short = txtAncPurchaseValues.GetIndex(eventSender)
		
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
						If InStr(txtAncPurchaseValues(Index).Text, ".") = 0 Then
							System.Windows.Forms.SendKeys.Send("{DELETE}")
						End If
					Else
						System.Windows.Forms.SendKeys.Send("{DELETE}")
					End If
			End Select
		End If
		
	End Sub
	Private Sub txtAncPurchaseValues_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtAncPurchaseValues.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = txtAncPurchaseValues.GetIndex(eventSender)
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
	Private Sub txtAncPurchaseValues_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAncPurchaseValues.Leave
		Dim Index As Short = txtAncPurchaseValues.GetIndex(eventSender)
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
		For i = 1 To Len(txtAncPurchaseValues(Sample).Text)
			Digit.Value = Mid(txtAncPurchaseValues(Sample).Text, i, 1)
			Select Case Digit.Value
				Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ".", "-"
					tempvalue = tempvalue & Digit.Value
			End Select
		Next i
		If CellValues(WhichScreen, Sample, WhichSegment).Changed = True Then
			CellValues(WhichScreen, Sample, WhichSegment).Value = Val(tempvalue)
		End If
		Call drawthevalues()
	End Sub
	'UPGRADE_WARNING: Event txtSegmentLabel.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtSegmentLabel_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSegmentLabel.TextChanged
		If DoNotChange = True Then Exit Sub
		SegNamie(WhichSegment) = txtSegmentLabel.Text
	End Sub
	Public Sub ScreenCalc()
		Dim x As Short
		On Error Resume Next
		For x = 6 To 13
			If CellValues(WhichScreen, x + 20, WhichSegment).Changed = False Then
				CellValues(WhichScreen, x + 20, WhichSegment).Value = CellValues(WhichScreen, x, WhichSegment).Value * CellValues(EquipmentTwo, x + 6, WhichSegment).Value
			End If
		Next x
	End Sub
End Class