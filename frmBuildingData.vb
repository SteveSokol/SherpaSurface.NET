Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks
Friend Class frmBuildingData
	Inherits System.Windows.Forms.Form
	Dim TempHigh As Single
	Dim TempWide As Single
	Dim FootConversion As Single
	Private Sub comBuildingPrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles comBuildingPrint.Click
		PrintForm1.Print(Me, PowerPacks.Printing.PrintForm.PrintOption.CompatibleModeClientAreaOnly)
	End Sub
	'UPGRADE_WARNING: Form event frmBuildingData.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmBuildingData_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated

        'Dim baseunit As String
        'Dim baselength As Short
        Dim i As Short
		
		If IsHelpOn = True Then
			txtBuildingValues(WhichCell).Focus()
			IsHelpOn = False
		Else
			For i = 0 To 5
				If optSegment(i).Checked = True Then WhichSegment = i
			Next i
			
			WhichScreen = Building
			
			Call drawthevalues()
			
			If InsertFlag = True Then
				labInsert.Text = "Insert"
			Else
				labInsert.Text = "Typeover"
			End If
			
			WhichCell = 0
			
			txtBuildingValues(0).Focus()
		End If
		
	End Sub
	Private Sub frmBuildingData_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim x As Short
		Dim y As Short
		
		Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - (VB6.PixelsToTwipsY(Me.Height) + 350)) / 2)
		Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2)
		
		If VB6.PixelsToTwipsY(Me.Top) < 0 Then Me.Top = 0
		If VB6.PixelsToTwipsX(Me.Left) < 0 Then Me.Left = 0
		
		TempWide = VB6.PixelsToTwipsX(Me.ClientRectangle.Width)
		TempHigh = VB6.PixelsToTwipsY(Me.ClientRectangle.Height)
		
		DoNotChange = True
		
		If UnitType = Metric Then
			For y = 0 To 3
				For x = 0 To 2
					labBuildingUnits((y * 4) + x).Text = "meters"
				Next x
			Next y
			FootConversion = 0.3048
		Else
			For y = 0 To 3
				For x = 0 To 2
					labBuildingUnits((y * 4) + x).Text = "feet"
				Next x
			Next y
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
	'UPGRADE_WARNING: Event frmBuildingData.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub frmBuildingData_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		TempWide = VB6.PixelsToTwipsX(Me.ClientRectangle.Width)
		TempHigh = VB6.PixelsToTwipsY(Me.ClientRectangle.Height)
		Call screenstuff()
	End Sub
	Private Sub frmBuildingData_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
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
	Private Sub labBuildingHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles labBuildingHelp.Click
		Dim StartHelp As Short
		StartHelp = 207
		IsHelpOn = True
		Call frmSurfaceHelp.gethelptext(StartHelp, WhichCell)
		frmSurfaceHelp.Show()
	End Sub
	'UPGRADE_WARNING: Event lstBuildingList.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub lstBuildingList_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstBuildingList.SelectedIndexChanged
		txtBuildingValues(WhichCell).Text = LTrim(RTrim(VB6.GetItemString(lstBuildingList, lstBuildingList.SelectedIndex)))
		CellValues(WhichScreen, WhichCell, WhichSegment).Changed = True
		Call Inputer(WhichCell)
		If WhichCell < 15 Then
			WhichCell = WhichCell + 1
		Else
			WhichCell = 0
		End If
		txtBuildingValues(WhichCell).Focus()
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
			txtBuildingValues(WhichCell).Focus()
		End If
	End Sub
	'UPGRADE_WARNING: Event txtBuildingValues.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtBuildingValues_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBuildingValues.TextChanged
		Dim Index As Short = txtBuildingValues.GetIndex(eventSender)
		If DoNotChange = True Then Exit Sub
	End Sub
	Private Sub txtBuildingValues_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBuildingValues.Enter
		Dim Index As Short = txtBuildingValues.GetIndex(eventSender)
        'Dim x As Short
        WhichCell = Index
		Select Case WhichCell
			Case 3, 7, 11, 15
				lstBuildingList.Visible = True
			Case Else
				lstBuildingList.Visible = False
		End Select
		System.Windows.Forms.SendKeys.Send("{HOME}+{END}")
		WhichCell = Index
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
		
		q = (480 / h) * TempHigh
		r = (1920 / h) * TempHigh
		s = (2880 / h) * TempHigh
		t = (960 / h) * TempHigh
		
		y = (390 / h) * TempHigh
		z = (300 / w) * TempWide
		
		labBuildingHeading(x).Top = VB6.TwipsToPixelsY(TempHigh * (180 / h))
		labBuildingHeading(x).Left = VB6.TwipsToPixelsX(TempWide * (180 / w))
		labBuildingHeading(x).Width = VB6.TwipsToPixelsX(TempWide * (2325 / w))
		
		For x = 0 To 3
			labBuildingLabels(x).Top = VB6.TwipsToPixelsY((TempHigh * (720 / h)) + (x * t))
			labBuildingLabels(x).Left = VB6.TwipsToPixelsX(TempWide * (720 / w))
		Next x
		
		labBuildingLabels(4).Top = VB6.TwipsToPixelsY(TempHigh * (4920 / h))
		labBuildingLabels(4).Left = VB6.TwipsToPixelsX(TempWide * (660 / w))
		labBuildingLabels(4).Width = VB6.TwipsToPixelsX(TempWide * (1820 / w))
		
		labBuildingLabels(5).Top = VB6.TwipsToPixelsY(TempHigh * (4740 / h))
		labBuildingLabels(5).Left = VB6.TwipsToPixelsX(TempWide * (3180 / w))
		
		labBuildingLabels(6).Top = VB6.TwipsToPixelsY(TempHigh * (4980 / h))
		labBuildingLabels(6).Left = VB6.TwipsToPixelsX(TempWide * (6660 / w))
		labBuildingLabels(6).Width = VB6.TwipsToPixelsX(TempWide * (1785 / w))
		
		For x = 0 To 5
			optSegment(x).Top = VB6.TwipsToPixelsY(TempHigh * (5220 / h))
			optSegment(x).Left = VB6.TwipsToPixelsX((TempWide * (720 / w)) + (x * z))
			optSegment(x).Width = VB6.TwipsToPixelsX(TempWide * (195 / w))
			labSegment(x).Top = VB6.TwipsToPixelsY(TempHigh * (5460 / h))
			labSegment(x).Left = VB6.TwipsToPixelsX((TempWide * (720 / w)) + (x * z))
			labSegment(x).Width = VB6.TwipsToPixelsX(TempWide * (195 / w))
		Next x
		
		txtSegmentLabel.Top = VB6.TwipsToPixelsY(TempHigh * (5280 / h))
		txtSegmentLabel.Left = VB6.TwipsToPixelsX(TempWide * (6660 / w))
		txtSegmentLabel.Width = VB6.TwipsToPixelsX(TempWide * (1935 / w))
		
		For x = 0 To 14 Step 2
			If x <= 2 Then
				p = 0
			ElseIf x <= 6 Then 
				p = (180 / h) * TempHigh
			ElseIf x <= 10 Then 
				p = (360 / h) * TempHigh
			Else
				p = (540 / h) * TempHigh
			End If
			LabBuildingTitles(x).Top = VB6.TwipsToPixelsY((TempHigh * (960 / h)) + ((x / 2) * y) + p)
			LabBuildingTitles(x).Left = VB6.TwipsToPixelsX(TempWide * (1500 / w))
			LabBuildingTitles(x).Width = VB6.TwipsToPixelsX(TempWide * (1755 / w))
			LabBuildingTitles(x + 1).Top = VB6.TwipsToPixelsY((TempHigh * (960 / h)) + ((x / 2) * y) + p)
			LabBuildingTitles(x + 1).Left = VB6.TwipsToPixelsX(TempWide * (5040 / w))
			txtBuildingValues(x).Top = VB6.TwipsToPixelsY((TempHigh * (930 / h)) + ((x / 2) * y) + p)
			txtBuildingValues(x).Left = VB6.TwipsToPixelsX(TempWide * (3420 / w))
			txtBuildingValues(x).Width = VB6.TwipsToPixelsX(TempWide * (915 / w))
			txtBuildingValues(x + 1).Top = VB6.TwipsToPixelsY((TempHigh * (930 / h)) + ((x / 2) * y) + p)
			labBuildingUnits(x).Top = VB6.TwipsToPixelsY((TempHigh * (960 / h)) + ((x / 2) * y) + p)
			labBuildingUnits(x).Left = VB6.TwipsToPixelsX(TempWide * (4380 / w))
			labBuildingUnits(x + 1).Top = VB6.TwipsToPixelsY((TempHigh * (960 / h)) + ((x / 2) * y) + p)
			Select Case x
				Case 2, 6, 10, 14
					LabBuildingTitles(x + 1).Width = VB6.TwipsToPixelsX(TempWide * (1155 / w))
					txtBuildingValues(x + 1).Left = VB6.TwipsToPixelsX(TempWide * (6360 / w))
					txtBuildingValues(x + 1).Width = VB6.TwipsToPixelsX(TempWide * (2335 / w))
					labBuildingUnits(x + 1).Left = VB6.TwipsToPixelsX(TempWide * (8760 / w))
				Case Else
					LabBuildingTitles(x + 1).Width = VB6.TwipsToPixelsX(TempWide * (1875 / w))
					txtBuildingValues(x + 1).Left = VB6.TwipsToPixelsX(TempWide * (7080 / w))
					txtBuildingValues(x + 1).Width = VB6.TwipsToPixelsX(TempWide * (915 / w))
					labBuildingUnits(x + 1).Left = VB6.TwipsToPixelsX(TempWide * (8040 / w))
			End Select
		Next x
		
		lstBuildingList.Top = VB6.TwipsToPixelsY(TempHigh * (5100 / h))
		lstBuildingList.Left = VB6.TwipsToPixelsX(TempWide * (3420 / w))
		lstBuildingList.Height = VB6.TwipsToPixelsY(TempHigh * (735 / h))
		lstBuildingList.Width = VB6.TwipsToPixelsX(TempWide * (2295 / w))
		
		LineHorizontal(0).X1 = VB6.TwipsToPixelsX(TempWide * (1260 / w))
		LineHorizontal(0).X2 = VB6.TwipsToPixelsX(TempWide * (8940 / w))
		LineHorizontal(0).Y1 = VB6.TwipsToPixelsY(TempHigh * (780 / h))
		LineHorizontal(0).Y2 = VB6.TwipsToPixelsY(TempHigh * (780 / h))
		
		LineHorizontal(1).X1 = VB6.TwipsToPixelsX(TempWide * (1080 / w))
		LineHorizontal(1).X2 = VB6.TwipsToPixelsX(TempWide * (8820 / w))
		LineHorizontal(1).Y1 = VB6.TwipsToPixelsY(TempHigh * (1740 / h))
		LineHorizontal(1).Y2 = VB6.TwipsToPixelsY(TempHigh * (1740 / h))
		
		LineHorizontal(2).X1 = VB6.TwipsToPixelsX(TempWide * (1320 / w))
		LineHorizontal(2).X2 = VB6.TwipsToPixelsX(TempWide * (8820 / w))
		LineHorizontal(2).Y1 = VB6.TwipsToPixelsY(TempHigh * (2700 / h))
		LineHorizontal(2).Y2 = VB6.TwipsToPixelsY(TempHigh * (2700 / h))
		
		LineHorizontal(3).X1 = VB6.TwipsToPixelsX(TempWide * (1800 / w))
		LineHorizontal(3).X2 = VB6.TwipsToPixelsX(TempWide * (8820 / w))
		LineHorizontal(3).Y1 = VB6.TwipsToPixelsY(TempHigh * (3660 / h))
		LineHorizontal(3).Y2 = VB6.TwipsToPixelsY(TempHigh * (3660 / h))
		
		LineHorizontal(4).X1 = VB6.TwipsToPixelsX(TempWide * (720 / w))
		LineHorizontal(4).X2 = VB6.TwipsToPixelsX(TempWide * (8940 / w))
		LineHorizontal(4).Y1 = VB6.TwipsToPixelsY(TempHigh * (4620 / h))
		LineHorizontal(4).Y2 = VB6.TwipsToPixelsY(TempHigh * (4620 / h))
		
		LineHorizontal(5).X1 = VB6.TwipsToPixelsX(TempWide * (3000 / w))
		LineHorizontal(5).X2 = VB6.TwipsToPixelsX(TempWide * (6120 / w))
		LineHorizontal(5).Y1 = VB6.TwipsToPixelsY(TempHigh * (6000 / h))
		LineHorizontal(5).Y2 = VB6.TwipsToPixelsY(TempHigh * (6000 / h))
		
		LineVertical(0).X1 = VB6.TwipsToPixelsX(TempWide * (780 / w))
		LineVertical(0).X2 = VB6.TwipsToPixelsX(TempWide * (780 / w))
		LineVertical(0).Y1 = VB6.TwipsToPixelsY(TempHigh * (1020 / h))
		LineVertical(0).Y2 = VB6.TwipsToPixelsY(TempHigh * (1620 / h))
		
		LineVertical(1).X1 = VB6.TwipsToPixelsX(TempWide * (780 / w))
		LineVertical(1).X2 = VB6.TwipsToPixelsX(TempWide * (780 / w))
		LineVertical(1).Y1 = VB6.TwipsToPixelsY(TempHigh * (1980 / h))
		LineVertical(1).Y2 = VB6.TwipsToPixelsY(TempHigh * (2580 / h))
		
		LineVertical(2).X1 = VB6.TwipsToPixelsX(TempWide * (780 / w))
		LineVertical(2).X2 = VB6.TwipsToPixelsX(TempWide * (780 / w))
		LineVertical(2).Y1 = VB6.TwipsToPixelsY(TempHigh * (2940 / h))
		LineVertical(2).Y2 = VB6.TwipsToPixelsY(TempHigh * (3540 / h))
		
		LineVertical(3).X1 = VB6.TwipsToPixelsX(TempWide * (780 / w))
		LineVertical(3).X2 = VB6.TwipsToPixelsX(TempWide * (780 / w))
		LineVertical(3).Y1 = VB6.TwipsToPixelsY(TempHigh * (3900 / h))
		LineVertical(3).Y2 = VB6.TwipsToPixelsY(TempHigh * (4680 / h))
		
		LineVertical(4).X1 = VB6.TwipsToPixelsX(TempWide * (3060 / w))
		LineVertical(4).X2 = VB6.TwipsToPixelsX(TempWide * (3060 / w))
		LineVertical(4).Y1 = VB6.TwipsToPixelsY(TempHigh * (4560 / h))
		LineVertical(4).Y2 = VB6.TwipsToPixelsY(TempHigh * (6060 / h))
		
		LineVertical(5).X1 = VB6.TwipsToPixelsX(TempWide * (6060 / w))
		LineVertical(5).X2 = VB6.TwipsToPixelsX(TempWide * (6060 / w))
		LineVertical(5).Y1 = VB6.TwipsToPixelsY(TempHigh * (4560 / h))
		LineVertical(5).Y2 = VB6.TwipsToPixelsY(TempHigh * (6060 / h))
		
		LineVertical(6).X1 = VB6.TwipsToPixelsX(TempWide * (8880 / w))
		LineVertical(6).X2 = VB6.TwipsToPixelsX(TempWide * (8880 / w))
		LineVertical(6).Y1 = VB6.TwipsToPixelsY(TempHigh * (720 / h))
		LineVertical(6).Y2 = VB6.TwipsToPixelsY(TempHigh * (4680 / h))
		
		labBackToMenu.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labBackToMenu.Left = VB6.TwipsToPixelsX(TempWide * (600 / w))
		
		imgBackToMenu.Top = VB6.TwipsToPixelsY(TempHigh * (6150 / h))
		imgBackToMenu.Left = VB6.TwipsToPixelsX(TempWide * (60 / w))
		imgBackToMenu.Width = VB6.TwipsToPixelsX(TempWide * (495 / w))
		
		combuildingPrint.Top = VB6.TwipsToPixelsY(TempHigh * (6150 / h))
		combuildingPrint.Left = VB6.TwipsToPixelsX(TempWide * (8880 / w))
		
		labbuildingHelp.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labbuildingHelp.Left = VB6.TwipsToPixelsX(TempWide * (8340 / w))
		
		labInsert.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labInsert.Left = VB6.TwipsToPixelsX(TempWide * (4080 / w))
		labInsert.Width = VB6.TwipsToPixelsX(TempWide * (975 / w))
		
	End Sub
	Public Sub drawthevalues()
		
		Dim i As Short
        'Dim x As Short

        DoNotChange = True
		
		For i = 0 To 15
			If CellValues(WhichScreen, i, WhichSegment).Changed = True Then
				txtBuildingValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF0000)
			Else
				txtBuildingValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
			End If
			Select Case i
				Case 3, 7, 11, 15
					txtBuildingValues(i).Text = LTrim(RTrim(CellValues(WhichScreen, i, WhichSegment).Word))
				Case Else
					txtBuildingValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value * FootConv)), "###,##0")
			End Select
		Next i
		
		DoNotChange = False
		
	End Sub
	Private Sub txtBuildingValues_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtBuildingValues.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Dim Index As Short = txtBuildingValues.GetIndex(eventSender)
		
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
						If InStr(txtBuildingValues(Index).Text, ".") = 0 Then
							System.Windows.Forms.SendKeys.Send("{DELETE}")
						End If
					Else
						System.Windows.Forms.SendKeys.Send("{DELETE}")
					End If
			End Select
		End If
		
	End Sub
	Private Sub txtBuildingValues_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtBuildingValues.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = txtBuildingValues.GetIndex(eventSender)
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
	Private Sub txtBuildingValues_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBuildingValues.Leave
		Dim Index As Short = txtBuildingValues.GetIndex(eventSender)
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
		Select Case Sample
			Case 3, 7, 11, 15
				CellValues(WhichScreen, Sample, WhichSegment).Word = txtBuildingValues(Sample).Text
			Case Else
				tempvalue = ""
				For i = 1 To Len(txtBuildingValues(Sample).Text)
					Digit.Value = Mid(txtBuildingValues(Sample).Text, i, 1)
					Select Case Digit.Value
						Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ".", "-"
							tempvalue = tempvalue & Digit.Value
					End Select
				Next i
				If CellValues(WhichScreen, Sample, WhichSegment).Changed = True Then
					If FootConv <> 0 Then CellValues(WhichScreen, Sample, WhichSegment).Value = Val(tempvalue) / FootConv
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