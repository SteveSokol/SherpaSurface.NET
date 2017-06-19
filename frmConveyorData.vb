Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks
Friend Class frmConveyorData
	Inherits System.Windows.Forms.Form
	
	Dim TempHigh As Single
	Dim TempWide As Single
	Dim InchConversion As Single
	Dim FootConversion As Single
	Dim TonConversion As Single
	Dim CubicConversion As Single
	Private Sub comConveyPrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles comConveyPrint.Click
		PrintForm1.Print(Me, PowerPacks.Printing.PrintForm.PrintOption.CompatibleModeClientAreaOnly)
	End Sub
	
	'UPGRADE_WARNING: Form event frmConveyorData.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmConveyorData_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated

        'Dim baseunit As String
        'Dim baselength As Short
        Dim i As Short
		
		If IsHelpOn = True Then
			txtConveyValues(WhichCell).Focus()
			IsHelpOn = False
		Else
			For i = 0 To 5
				If optSegment(i).Checked = True Then WhichSegment = i
			Next i
			
			Call drawthevalues()
			
			If InsertFlag = True Then
				labInsert.Text = "Insert"
			Else
				labInsert.Text = "Typeover"
			End If
			
			WhichCell = 0
			
			txtConveyValues(0).Focus()
		End If
		
	End Sub
	Private Sub frmConveyorData_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        'Dim i As Short
        'Dim x As Short

        Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - (VB6.PixelsToTwipsY(Me.Height) + 350)) / 2)
		Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2)
		
		If VB6.PixelsToTwipsY(Me.Top) < 0 Then Me.Top = 0
		If VB6.PixelsToTwipsX(Me.Left) < 0 Then Me.Left = 0
		
		TempWide = VB6.PixelsToTwipsX(Me.ClientRectangle.Width)
		TempHigh = VB6.PixelsToTwipsY(Me.ClientRectangle.Height)
		
		If UnitType = Metric Then
			TonConversion = 0.9072
			InchConversion = 2.54
			FootConversion = 0.304801
		Else
			TonConversion = 1
			InchConversion = 1
			FootConversion = 1
		End If
		
		Call UnitMaker()
		
		DoNotChange = True
		
		WhichSegment = 0
		optSegment(WhichSegment).Checked = True
		
		If PageChange(WhichScreen) = True Then
			Call drawthevalues()
		End If
		
		DoNotChange = False
		
		Call screenstuff()
		
	End Sub
	'UPGRADE_WARNING: Event frmConveyorData.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub frmConveyorData_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		TempWide = VB6.PixelsToTwipsX(Me.ClientRectangle.Width)
		TempHigh = VB6.PixelsToTwipsY(Me.ClientRectangle.Height)
		Call screenstuff()
	End Sub
	Private Sub frmConveyorData_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		Call Inputer(WhichCell)
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
	Private Sub labConveyHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles labConveyHelp.Click
		Dim StartHelp As Short
		Dim SendHelp As Short
		StartHelp = 70
		If WhichCell < 11 Then
			SendHelp = WhichCell
		Else
			SendHelp = WhichCell - 11
		End If
		IsHelpOn = True
		Call frmSurfaceHelp.gethelptext(StartHelp, SendHelp)
		frmSurfaceHelp.Show()
	End Sub
	'UPGRADE_WARNING: Event lstOptions.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub lstOptions_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstOptions.SelectedIndexChanged
		Dim Index As Short = lstOptions.GetIndex(eventSender)
		Select Case Index
			Case 0, 10
				txtConveyValues(WhichCell).Text = RTrim(VB6.GetItemString(lstOptions(Index), lstOptions(Index).SelectedIndex))
			Case 10, 21
				txtConveyValues(WhichCell).Text = RTrim(VB6.GetItemString(lstOptions(Index), lstOptions(Index).SelectedIndex))
			Case Else
				txtConveyValues(WhichCell).Text = VB6.GetItemString(lstOptions(Index), lstOptions(Index).SelectedIndex)
		End Select
		CellValues(WhichScreen, WhichCell, WhichSegment).Changed = True
		Call Inputer(WhichCell)
		Call drawthevalues()
		If WhichCell = 21 Then
			WhichCell = 0
		Else
			WhichCell = WhichCell + 1
		End If
		txtConveyValues(WhichCell).Focus()
		
	End Sub
	'UPGRADE_WARNING: Event txtConveyValues.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtConveyValues_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtConveyValues.TextChanged
		Dim Index As Short = txtConveyValues.GetIndex(eventSender)
		If DoNotChange = True Then Exit Sub
		PageChange(WhichScreen) = True
		WhichCell = Index
	End Sub
	Public Sub screenstuff()
		
		Dim i As Single
		Dim j As Decimal
		Dim k As Decimal
		Dim l As Single
		Dim p As Single
		Dim r As Decimal
		Dim s As Decimal
		Dim t As Decimal
		Dim u As Decimal
		Dim v As Single
		Dim x As Short
		Dim y As Single
		Dim z As Single
		Dim w As Short
		Dim h As Short
		
		w = 9150 'Starting Form Scale Width
		h = 6420 'Starting Form Scale Height
		
		i = (480 / h) * TempHigh
		j = (720 / w) * TempWide
		k = (1260 / h) * TempHigh
		l = (3180 / h) * TempHigh
		r = (1320 / w) * TempWide
		s = (420 / h) * TempHigh
		t = (300 / h) * TempHigh
		u = (300 / w) * TempWide
		v = (660 / w) * TempWide
		y = (60 / h) * TempHigh
		z = (60 / w) * TempWide
		
		For x = 0 To 1
			labConveyHeading(x).Top = VB6.TwipsToPixelsY(((120 / h) * TempHigh) + (x * i))
			labConveyHeading(x).Left = VB6.TwipsToPixelsX(((120 / w) * TempWide) + (x * j))
			labConveyHeading(x).Width = VB6.TwipsToPixelsX((2325 / w) * TempWide)
		Next x
		
		For x = 0 To 1
			labConveyLabels(x).Top = VB6.TwipsToPixelsY(((1620 / h) * TempHigh) + (x * k))
			labConveyLabels(x).Left = VB6.TwipsToPixelsX((720 / w) * TempWide)
			labConveyLabels(x).Width = VB6.TwipsToPixelsX((1935 / w) * TempWide)
			labConveyLabels(x + 3).Top = VB6.TwipsToPixelsY(((600 / h) * TempHigh) + (x * l))
			labConveyLabels(x + 3).Left = VB6.TwipsToPixelsX((3360 / w) * TempWide)
			labConveyLabels(x + 5).Top = VB6.TwipsToPixelsY((360 / h) * TempHigh)
			labConveyLabels(x + 5).Left = VB6.TwipsToPixelsX(((5520 / w) * TempWide) + (x * r))
			labConveyLabels(x + 5).Width = VB6.TwipsToPixelsX((615 / w) * TempWide)
		Next x
		
		For x = 0 To 5
			optSegment(x).Left = VB6.TwipsToPixelsX(((840 / w) * TempWide) + (x * u))
			optSegment(x).Top = VB6.TwipsToPixelsY((1920 / h) * TempHigh)
			labOptLabel(x).Left = VB6.TwipsToPixelsX(((840 / w) * TempWide) + (x * u))
			labOptLabel(x).Top = VB6.TwipsToPixelsY((2160 / h) * TempHigh)
		Next x
		
		txtSegmentLabel.Left = VB6.TwipsToPixelsX((720 / w) * TempWide)
		txtSegmentLabel.Width = VB6.TwipsToPixelsX((1935 / w) * TempWide)
		txtSegmentLabel.Top = VB6.TwipsToPixelsY((3180 / h) * TempHigh)
		
		labConveyLabels(2).Left = VB6.TwipsToPixelsX((240 / w) * TempWide)
		labConveyLabels(2).Top = VB6.TwipsToPixelsY((4080 / h) * TempHigh)
		
		For x = 0 To 10
			If x < 7 Then
				p = 0
			Else
				p = (240 / h) * TempHigh
			End If
			If x = 0 Then
				LabConveyTitles(x).Width = VB6.TwipsToPixelsX((1575 / w) * TempWide)
				txtConveyValues(x).Left = VB6.TwipsToPixelsX((5340 / w) * TempWide)
				txtConveyValues(x).Width = VB6.TwipsToPixelsX((975 / w) * TempWide)
				txtConveyValues(x + 11).Left = VB6.TwipsToPixelsX((6660 / w) * TempWide)
				txtConveyValues(x + 11).Width = VB6.TwipsToPixelsX((975 / w) * TempWide)
				labConveyUnits(x).Left = VB6.TwipsToPixelsX((7680 / w) * TempWide)
			Else
				LabConveyTitles(x).Width = VB6.TwipsToPixelsX((1635 / w) * TempWide)
				txtConveyValues(x).Left = VB6.TwipsToPixelsX((5460 / w) * TempWide)
				txtConveyValues(x).Width = VB6.TwipsToPixelsX((735 / w) * TempWide)
				txtConveyValues(x + 11).Left = VB6.TwipsToPixelsX((6780 / w) * TempWide)
				txtConveyValues(x + 11).Width = VB6.TwipsToPixelsX((735 / w) * TempWide)
				labConveyUnits(x).Left = VB6.TwipsToPixelsX((7560 / w) * TempWide)
			End If
			LabConveyTitles(x).Left = VB6.TwipsToPixelsX((3600 / w) * TempWide)
			LabConveyTitles(x).Top = VB6.TwipsToPixelsY(((900 / h) * TempHigh) + (x * s) + p)
			txtConveyValues(x).Top = VB6.TwipsToPixelsY(((870 / h) * TempHigh) + (x * s) + p)
			labConveyUnits(x).Top = VB6.TwipsToPixelsY(((900 / h) * TempHigh) + (x * s) + p)
			txtConveyValues(x + 11).Top = VB6.TwipsToPixelsY(((870 / h) * TempHigh) + (x * s) + p)
		Next x
		
		labBackToMenu.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labBackToMenu.Left = VB6.TwipsToPixelsX(TempWide * (600 / w))
		
		imgBackToMenu.Top = VB6.TwipsToPixelsY(TempHigh * (6150 / h))
		imgBackToMenu.Left = VB6.TwipsToPixelsX(TempWide * (60 / w))
		imgBackToMenu.Width = VB6.TwipsToPixelsX(TempWide * (495 / w))
		
		comConveyPrint.Top = VB6.TwipsToPixelsY(TempHigh * (6150 / h))
		comConveyPrint.Left = VB6.TwipsToPixelsX(TempWide * (8880 / w))
		
		labConveyHelp.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labConveyHelp.Left = VB6.TwipsToPixelsX(TempWide * (8340 / w))
		
		labInsert.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labInsert.Left = VB6.TwipsToPixelsX(TempWide * (4080 / w))
		labInsert.Width = VB6.TwipsToPixelsX(TempWide * (975 / w))
		
		LineHorizontal(0).X1 = VB6.TwipsToPixelsX((1500 / w) * TempWide)
		LineHorizontal(0).X2 = VB6.TwipsToPixelsX((3480 / w) * TempWide)
		LineHorizontal(0).Y1 = VB6.TwipsToPixelsY((4140 / h) * TempHigh)
		LineHorizontal(0).Y2 = VB6.TwipsToPixelsY((4140 / h) * TempHigh)
		
		LineHorizontal(1).X1 = VB6.TwipsToPixelsX((240 / w) * TempWide)
		LineHorizontal(1).X2 = VB6.TwipsToPixelsX((3480 / w) * TempWide)
		LineHorizontal(1).Y1 = VB6.TwipsToPixelsY((5640 / h) * TempHigh)
		LineHorizontal(1).Y2 = VB6.TwipsToPixelsY((5640 / h) * TempHigh)
		
		LineHorizontal(2).X1 = VB6.TwipsToPixelsX((3360 / w) * TempWide)
		LineHorizontal(2).X2 = VB6.TwipsToPixelsX((9060 / w) * TempWide)
		LineHorizontal(2).Y1 = VB6.TwipsToPixelsY((240 / h) * TempHigh)
		LineHorizontal(2).Y2 = VB6.TwipsToPixelsY((240 / h) * TempHigh)
		
		LineHorizontal(3).X1 = VB6.TwipsToPixelsX((4200 / w) * TempWide)
		LineHorizontal(3).X2 = VB6.TwipsToPixelsX((6420 / w) * TempWide)
		LineHorizontal(3).Y1 = VB6.TwipsToPixelsY((660 / h) * TempHigh)
		LineHorizontal(3).Y2 = VB6.TwipsToPixelsY((660 / h) * TempHigh)
		
		LineHorizontal(4).X1 = VB6.TwipsToPixelsX((6540 / w) * TempWide)
		LineHorizontal(4).X2 = VB6.TwipsToPixelsX((8940 / w) * TempWide)
		LineHorizontal(4).Y1 = VB6.TwipsToPixelsY((660 / h) * TempHigh)
		LineHorizontal(4).Y2 = VB6.TwipsToPixelsY((660 / h) * TempHigh)
		
		LineHorizontal(5).X1 = VB6.TwipsToPixelsX((4320 / w) * TempWide)
		LineHorizontal(5).X2 = VB6.TwipsToPixelsX((6420 / w) * TempWide)
		LineHorizontal(5).Y1 = VB6.TwipsToPixelsY((3840 / h) * TempHigh)
		LineHorizontal(5).Y2 = VB6.TwipsToPixelsY((3840 / h) * TempHigh)
		
		LineHorizontal(6).X1 = VB6.TwipsToPixelsX((6540 / w) * TempWide)
		LineHorizontal(6).X2 = VB6.TwipsToPixelsX((8940 / w) * TempWide)
		LineHorizontal(6).Y1 = VB6.TwipsToPixelsY((3840 / h) * TempHigh)
		LineHorizontal(6).Y2 = VB6.TwipsToPixelsY((3840 / h) * TempHigh)
		
		LineHorizontal(7).X1 = VB6.TwipsToPixelsX((3360 / w) * TempWide)
		LineHorizontal(7).X2 = VB6.TwipsToPixelsX((9060 / w) * TempWide)
		LineHorizontal(7).Y1 = VB6.TwipsToPixelsY((5760 / h) * TempHigh)
		LineHorizontal(7).Y2 = VB6.TwipsToPixelsY((5760 / h) * TempHigh)
		
		LineVertical(0).X1 = VB6.TwipsToPixelsX((300 / w) * TempWide)
		LineVertical(0).X2 = VB6.TwipsToPixelsX((300 / w) * TempWide)
		LineVertical(0).Y1 = VB6.TwipsToPixelsY((5700 / h) * TempHigh)
		LineVertical(0).Y2 = VB6.TwipsToPixelsY((4380 / h) * TempHigh)
		
		LineVertical(1).X1 = VB6.TwipsToPixelsX((3420 / w) * TempWide)
		LineVertical(1).X2 = VB6.TwipsToPixelsX((3420 / w) * TempWide)
		LineVertical(1).Y1 = VB6.TwipsToPixelsY((180 / h) * TempHigh)
		LineVertical(1).Y2 = VB6.TwipsToPixelsY((540 / h) * TempHigh)
		
		LineVertical(2).X1 = VB6.TwipsToPixelsX((3420 / w) * TempWide)
		LineVertical(2).X2 = VB6.TwipsToPixelsX((3420 / w) * TempWide)
		LineVertical(2).Y1 = VB6.TwipsToPixelsY((900 / h) * TempHigh)
		LineVertical(2).Y2 = VB6.TwipsToPixelsY((3720 / h) * TempHigh)
		
		LineVertical(3).X1 = VB6.TwipsToPixelsX((3420 / w) * TempWide)
		LineVertical(3).X2 = VB6.TwipsToPixelsX((3420 / w) * TempWide)
		LineVertical(3).Y1 = VB6.TwipsToPixelsY((4080 / h) * TempHigh)
		LineVertical(3).Y2 = VB6.TwipsToPixelsY((5820 / h) * TempHigh)
		
		LineVertical(4).X1 = VB6.TwipsToPixelsX((6480 / w) * TempWide)
		LineVertical(4).X2 = VB6.TwipsToPixelsX((6480 / w) * TempWide)
		LineVertical(4).Y1 = VB6.TwipsToPixelsY((300 / h) * TempHigh)
		LineVertical(4).Y2 = VB6.TwipsToPixelsY((600 / h) * TempHigh)
		
		LineVertical(5).X1 = VB6.TwipsToPixelsX((6480 / w) * TempWide)
		LineVertical(5).X2 = VB6.TwipsToPixelsX((6480 / w) * TempWide)
		LineVertical(5).Y1 = VB6.TwipsToPixelsY((720 / h) * TempHigh)
		LineVertical(5).Y2 = VB6.TwipsToPixelsY((3780 / h) * TempHigh)
		
		LineVertical(6).X1 = VB6.TwipsToPixelsX((6480 / w) * TempWide)
		LineVertical(6).X2 = VB6.TwipsToPixelsX((6480 / w) * TempWide)
		LineVertical(6).Y1 = VB6.TwipsToPixelsY((3900 / h) * TempHigh)
		LineVertical(6).Y2 = VB6.TwipsToPixelsY((5700 / h) * TempHigh)
		
		LineVertical(7).X1 = VB6.TwipsToPixelsX((9000 / w) * TempWide)
		LineVertical(7).X2 = VB6.TwipsToPixelsX((9000 / w) * TempWide)
		LineVertical(7).Y1 = VB6.TwipsToPixelsY((180 / h) * TempHigh)
		LineVertical(7).Y2 = VB6.TwipsToPixelsY((5820 / h) * TempHigh)
		
		labConveyLabels(7).Left = VB6.TwipsToPixelsX((840 / w) * TempWide)
		labConveyLabels(7).Width = VB6.TwipsToPixelsX((2055 / w) * TempWide)
		labConveyLabels(7).Top = VB6.TwipsToPixelsY((4500 / h) * TempHigh)
		
		lstOptions(0).Left = VB6.TwipsToPixelsX((840 / w) * TempWide)
		lstOptions(0).Width = VB6.TwipsToPixelsX((2055 / w) * TempWide)
		lstOptions(0).Height = VB6.TwipsToPixelsY((735 / h) * TempHigh)
		lstOptions(0).Top = VB6.TwipsToPixelsY((4740 / h) * TempHigh)
		
	End Sub
	Public Sub drawthevalues()
		
		Dim i As Short
        'Dim x As Short
        'Dim TempWord As Short
        DoNotChange = True
		If UnitType = Metric Then
			TonConversion = 0.9072
			InchConversion = 2.54
			FootConversion = 0.304801
		Else
			TonConversion = 1
			InchConversion = 1
			FootConversion = 1
		End If
		
		For i = 0 To 21
			If CellValues(WhichScreen, i, WhichSegment).Changed = True Then
				txtConveyValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF0000)
			Else
				txtConveyValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&H0)
			End If
			Select Case i
				Case 0, 11
					txtConveyValues(i).Text = LTrim(RTrim(CellValues(WhichScreen, i, WhichSegment).Word))
				Case 1, 12
					txtConveyValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value * InchConversion)), "#,###,##0.0")
				Case 2, 13
					txtConveyValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value * InchConversion)), "#,###,##0.00")
				Case 4, 15
					If TonConversion <> 0 Then txtConveyValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value / TonConversion)), "#,###,##0.00")
				Case 5, 16
					txtConveyValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value)), "#,###,##0")
				Case 6, 17
					txtConveyValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value * TonConversion)), "#,###,##0")
				Case 7, 9, 18, 20
					txtConveyValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value * FootConversion)), "#,###,##0")
				Case 3, 8, 10, 14, 19, 21
					txtConveyValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value)), "#,###,##0.0")
			End Select
		Next i
		DoNotChange = False
	End Sub
	Private Sub txtConveyValues_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtConveyValues.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Dim Index As Short = txtConveyValues.GetIndex(eventSender)
		
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
						If InStr(txtConveyValues(Index).Text, ".") = 0 Then
							System.Windows.Forms.SendKeys.Send("{DELETE}")
						End If
					Else
						System.Windows.Forms.SendKeys.Send("{DELETE}")
					End If
			End Select
		End If
		
	End Sub
	Private Sub txtConveyValues_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtConveyValues.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = txtConveyValues.GetIndex(eventSender)
		If KeyAscii > Asc("9") And KeyAscii <> Asc(",") And KeyAscii <> Asc(".") And KeyAscii <> Asc("$") Then
			KeyAscii = 0
			Beep()
		Else
			CellValues(WhichScreen, WhichCell, WhichSegment).Changed = True
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	Private Sub txtConveyValues_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtConveyValues.Leave
		Dim Index As Short = txtConveyValues.GetIndex(eventSender)
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
		For i = 1 To Len(txtConveyValues(Sample).Text)
			Digit.Value = Mid(txtConveyValues(Sample).Text, i, 1)
			Select Case Digit.Value
				Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ".", "-"
					tempvalue = tempvalue & Digit.Value
			End Select
		Next i
		Select Case Sample
			Case 1, 2, 12, 13
				If CellValues(WhichScreen, Sample, WhichSegment).Changed = True Then
					If InchConversion <> 0 Then CellValues(WhichScreen, Sample, WhichSegment).Value = Val(tempvalue) / InchConversion
				End If
			Case 4, 15
				If CellValues(WhichScreen, Sample, WhichSegment).Changed = True Then
					CellValues(WhichScreen, Sample, WhichSegment).Value = Val(tempvalue) * TonConversion
				End If
			Case 6, 17
				If CellValues(WhichScreen, Sample, WhichSegment).Changed = True Then
					If TonConversion <> 0 Then CellValues(WhichScreen, Sample, WhichSegment).Value = Val(tempvalue) / TonConversion
				End If
			Case 7, 9, 18, 20
				If CellValues(WhichScreen, Sample, WhichSegment).Changed = True Then
					If FootConversion <> 0 Then CellValues(WhichScreen, Sample, WhichSegment).Value = Val(tempvalue) / FootConversion
				End If
			Case 0, 11
				CellValues(WhichScreen, Sample, WhichSegment).Word = txtConveyValues(Sample).Text
			Case Else
				If CellValues(WhichScreen, Sample, WhichSegment).Changed = True Then
					CellValues(WhichScreen, Sample, WhichSegment).Value = Val(tempvalue)
				End If
		End Select
		Call ccopt()
		Call drawthevalues()
	End Sub
	Private Sub txtConveyValues_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtConveyValues.Enter
		Dim Index As Short = txtConveyValues.GetIndex(eventSender)
        'Dim x As Short

        WhichCell = Index
		
		lstOptions(0).Visible = False
		labConveyLabels(7).Visible = False
		
		Select Case WhichCell
			Case 0, 11
				lstOptions(0).Visible = True
				labConveyLabels(7).Visible = True
		End Select
		
	End Sub
	'UPGRADE_WARNING: Event optSegment.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub optSegment_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optSegment.CheckedChanged
		If eventSender.Checked Then
			Dim Index As Short = optSegment.GetIndex(eventSender)
			Dim i As Short
			WhichSegment = Index
			Call drawthevalues()
			For i = 0 To 5
				labOptLabel(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFFFFFF)
			Next i
			labOptLabel(WhichSegment).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF)
			'txtConveySetLabel.Text = AditNamie(WhichSegment)
		End If
	End Sub
	Private Sub UnitMaker()
		
		If UnitType = Metric Then
			labConveyUnits(1).Text = "centimeters"
			labConveyUnits(2).Text = "centimeters"
			labConveyUnits(4).Text = "kwh/metric ton"
			labConveyUnits(6).Text = "metric tons/hour"
			labConveyUnits(7).Text = "meters"
			labConveyUnits(9).Text = "meters/minute"
		Else
			labConveyUnits(1).Text = "inches"
			labConveyUnits(2).Text = "inches"
			labConveyUnits(4).Text = "kwh/ton"
			labConveyUnits(6).Text = "tons/hour"
			labConveyUnits(7).Text = "feet"
			labConveyUnits(9).Text = "feet/minute"
		End If
		
	End Sub
End Class