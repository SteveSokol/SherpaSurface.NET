Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks
Friend Class frmDepositData
	Inherits System.Windows.Forms.Form
	Dim TempHigh As Single
	Dim TempWide As Single
	Private Sub comDepositPrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles comDepositPrint.Click
		PrintForm1.Print(Me, PowerPacks.Printing.PrintForm.PrintOption.CompatibleModeClientAreaOnly)
	End Sub
	'UPGRADE_WARNING: Form event frmDepositData.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmDepositData_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated

        'Dim baseunit As String
        'Dim baselength As Short
        Dim i As Short
		
		If IsHelpOn = True Then
			txtDepositValues(WhichCell).Focus()
			IsHelpOn = False
		Else
			For i = 0 To 5
				If optSegment(i).Checked = True Then WhichSegment = i
			Next i
			
			WhichScreen = Deposit
			
			Call drawthevalues()
			
			If InsertFlag = True Then
				labInsert.Text = "Insert"
			Else
				labInsert.Text = "Typeover"
			End If
			
			WhichCell = 0
			
			txtDepositValues(0).Focus()
		End If
		
	End Sub
	Private Sub frmDepositData_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        'Dim x As Short

        Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - (VB6.PixelsToTwipsY(Me.Height) + 350)) / 2)
		Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2)
		
		If VB6.PixelsToTwipsY(Me.Top) < 0 Then Me.Top = 0
		If VB6.PixelsToTwipsX(Me.Left) < 0 Then Me.Left = 0
		
		TempWide = VB6.PixelsToTwipsX(Me.ClientRectangle.Width)
		TempHigh = VB6.PixelsToTwipsY(Me.ClientRectangle.Height)
		
		DoNotChange = True
		
		If UnitType = Metric Then
			labDepositUnits(0).Text = "metric tons/bank cubic meter"
			labDepositUnits(7).Text = "metric tons/bank cubic meter"
			labDepositUnits(2).Text = "kilograms/metric ton"
			labDepositUnits(9).Text = "kilograms/metric ton"
			labDepositUnits(4).Text = "meters/hour"
			labDepositUnits(11).Text = "meters/hour"
		Else
			labDepositUnits(0).Text = "tons/bank cubic yard"
			labDepositUnits(7).Text = "tons/bank cubic yard"
			labDepositUnits(2).Text = "pounds/ton"
			labDepositUnits(9).Text = "pounds/ton"
			labDepositUnits(4).Text = "feet/hour"
			labDepositUnits(11).Text = "feet/hour"
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
	'UPGRADE_WARNING: Event frmDepositData.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub frmDepositData_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		TempWide = VB6.PixelsToTwipsX(Me.ClientRectangle.Width)
		TempHigh = VB6.PixelsToTwipsY(Me.ClientRectangle.Height)
		Call screenstuff()
	End Sub
	Private Sub frmDepositData_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
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
	Private Sub labDepositHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles labDepositHelp.Click
		Dim StartHelp As Short
		Dim SendHelp As Short
		IsHelpOn = True
		StartHelp = 29
		Select Case WhichCell
			Case 6 To 11
				SendHelp = WhichCell - 1
			Case 13
				SendHelp = WhichCell - 2
			Case Else
				SendHelp = WhichCell
		End Select
		Call frmSurfaceHelp.gethelptext(StartHelp, SendHelp)
		frmSurfaceHelp.Show()
	End Sub
	'UPGRADE_WARNING: Event lstDepositList.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub lstDepositList_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstDepositList.SelectedIndexChanged
		txtDepositValues(WhichCell).Text = LTrim(RTrim(VB6.GetItemString(lstDepositList, lstDepositList.SelectedIndex)))
		CellValues(WhichScreen, WhichCell, WhichSegment).Changed = True
		Call Inputer(WhichCell)
		If WhichCell < 15 Then
			WhichCell = 4
		Else
			WhichCell = 11
		End If
		txtDepositValues(WhichCell).Focus()
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
			txtDepositValues(WhichCell).Focus()
		End If
	End Sub
	'UPGRADE_WARNING: Event txtDepositValues.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtDepositValues_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDepositValues.TextChanged
		Dim Index As Short = txtDepositValues.GetIndex(eventSender)
		If DoNotChange = True Then Exit Sub
	End Sub
	Private Sub txtDepositValues_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDepositValues.Enter
		Dim Index As Short = txtDepositValues.GetIndex(eventSender)
        'Dim x As Short
        WhichCell = Index
		Select Case WhichCell
			Case 14, 15
				lstDepositList.Visible = True
			Case Else
				lstDepositList.Visible = False
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
		
		p = (3060 / h) * TempHigh
		q = (480 / h) * TempHigh
		r = (1920 / h) * TempHigh
		s = (3060 / h) * TempHigh
		t = (1260 / h) * TempHigh
		
		y = (420 / h) * TempHigh
		z = (300 / w) * TempWide
		
		For x = 0 To 2
			labDepositHeading(x).Top = VB6.TwipsToPixelsY((TempHigh * (180 / h)) + (x * q))
			labDepositHeading(x).Left = VB6.TwipsToPixelsX((TempWide * (180 / w)) + (x * z))
			labDepositHeading(x).Width = VB6.TwipsToPixelsX(TempWide * (2325 / w))
		Next x
		
		For x = 0 To 1
			labDepositLabels(x).Top = VB6.TwipsToPixelsY((TempHigh * (1920 / h)) + (x * t))
			labDepositLabels(x).Left = VB6.TwipsToPixelsX(TempWide * (360 / w))
			labDepositLabels(x).Width = VB6.TwipsToPixelsX(TempWide * (2655 / w))
		Next x
		
		labDepositLabels(2).Top = VB6.TwipsToPixelsY(TempHigh * (4140 / h))
		labDepositLabels(2).Left = VB6.TwipsToPixelsX(TempWide * (120 / w))
		
		For x = 0 To 1
			labDepositLabels(x + 3).Top = VB6.TwipsToPixelsY((TempHigh * (120 / h)) + (x * s))
			labDepositLabels(x + 3).Left = VB6.TwipsToPixelsX(TempWide * (3300 / w))
		Next x
		
		For x = 0 To 5
			optSegment(x).Top = VB6.TwipsToPixelsY(TempHigh * (2220 / h))
			optSegment(x).Left = VB6.TwipsToPixelsX((TempWide * (840 / w)) + (x * z))
			optSegment(x).Width = VB6.TwipsToPixelsX(TempWide * (195 / w))
			labSegment(x).Top = VB6.TwipsToPixelsY(TempHigh * (2460 / h))
			labSegment(x).Left = VB6.TwipsToPixelsX((TempWide * (840 / w)) + (x * z))
			labSegment(x).Width = VB6.TwipsToPixelsX(TempWide * (195 / w))
		Next x
		
		txtSegmentLabel.Top = VB6.TwipsToPixelsY(TempHigh * (3480 / h))
		txtSegmentLabel.Left = VB6.TwipsToPixelsX(TempWide * (720 / w))
		txtSegmentLabel.Width = VB6.TwipsToPixelsX(TempWide * (1935 / w))
		
		For x = 0 To 3
			LabDepositTitles(x).Top = VB6.TwipsToPixelsY((TempHigh * (360 / h)) + (x * y))
			If x = 0 Then
				LabDepositTitles(x).Left = VB6.TwipsToPixelsX(TempWide * (3840 / w))
				LabDepositTitles(x).Width = VB6.TwipsToPixelsX(TempWide * (1455 / w))
			Else
				LabDepositTitles(x).Left = VB6.TwipsToPixelsX(TempWide * (3480 / w))
				LabDepositTitles(x).Width = VB6.TwipsToPixelsX(TempWide * (1815 / w))
			End If
			txtDepositValues(x).Top = VB6.TwipsToPixelsY((TempHigh * (330 / h)) + (x * y))
			txtDepositValues(x).Left = VB6.TwipsToPixelsX(TempWide * (5460 / w))
			txtDepositValues(x).Width = VB6.TwipsToPixelsX(TempWide * (855 / w))
			labDepositUnits(x).Top = VB6.TwipsToPixelsY((TempHigh * (360 / h)) + (x * y))
			labDepositUnits(x).Left = VB6.TwipsToPixelsX(TempWide * (6360 / w))
		Next x
		
		For x = 0 To 2 Step 2
			LabDepositTitles(x + 4).Top = VB6.TwipsToPixelsY((TempHigh * (2460 / h)) + ((x / 2) * y))
			LabDepositTitles(x + 4).Left = VB6.TwipsToPixelsX(TempWide * (3480 / w))
			LabDepositTitles(x + 4).Width = VB6.TwipsToPixelsX(TempWide * (1815 / w))
			txtDepositValues(x + 4).Top = VB6.TwipsToPixelsY((TempHigh * (2430 / h)) + ((x / 2) * y))
			txtDepositValues(x + 4).Left = VB6.TwipsToPixelsX(TempWide * (5460 / w))
			txtDepositValues(x + 4).Width = VB6.TwipsToPixelsX(TempWide * (855 / w))
			labDepositUnits(x + 4).Top = VB6.TwipsToPixelsY((TempHigh * (2460 / h)) + ((x / 2) * y))
			labDepositUnits(x + 4).Left = VB6.TwipsToPixelsX(TempWide * (6360 / w))
		Next x
		
		For x = 0 To 3
			LabDepositTitles(x + 7).Top = VB6.TwipsToPixelsY((TempHigh * (3420 / h)) + (x * y))
			If x = 0 Then
				LabDepositTitles(x + 7).Left = VB6.TwipsToPixelsX(TempWide * (3840 / w))
				LabDepositTitles(x + 7).Width = VB6.TwipsToPixelsX(TempWide * (1455 / w))
			Else
				LabDepositTitles(x + 7).Left = VB6.TwipsToPixelsX(TempWide * (3480 / w))
				LabDepositTitles(x + 7).Width = VB6.TwipsToPixelsX(TempWide * (1815 / w))
			End If
			txtDepositValues(x + 7).Top = VB6.TwipsToPixelsY((TempHigh * (3390 / h)) + (x * y))
			txtDepositValues(x + 7).Left = VB6.TwipsToPixelsX(TempWide * (5460 / w))
			txtDepositValues(x + 7).Width = VB6.TwipsToPixelsX(TempWide * (855 / w))
			labDepositUnits(x + 7).Top = VB6.TwipsToPixelsY((TempHigh * (3420 / h)) + (x * y))
			labDepositUnits(x + 7).Left = VB6.TwipsToPixelsX(TempWide * (6360 / w))
		Next x
		
		For x = 0 To 2 Step 2
			LabDepositTitles(x + 11).Top = VB6.TwipsToPixelsY((TempHigh * (5520 / h)) + ((x / 2) * y))
			LabDepositTitles(x + 11).Left = VB6.TwipsToPixelsX(TempWide * (3480 / w))
			LabDepositTitles(x + 11).Width = VB6.TwipsToPixelsX(TempWide * (1815 / w))
			txtDepositValues(x + 11).Top = VB6.TwipsToPixelsY((TempHigh * (5490 / h)) + ((x / 2) * y))
			txtDepositValues(x + 11).Left = VB6.TwipsToPixelsX(TempWide * (5460 / w))
			txtDepositValues(x + 11).Width = VB6.TwipsToPixelsX(TempWide * (855 / w))
			labDepositUnits(x + 11).Top = VB6.TwipsToPixelsY((TempHigh * (5520 / h)) + ((x / 2) * y))
			labDepositUnits(x + 11).Left = VB6.TwipsToPixelsX(TempWide * (6360 / w))
		Next x
		
		For x = 0 To 1
			LabDepositTitles(x + 14).Top = VB6.TwipsToPixelsY((TempHigh * (2040 / h)) + (x * p))
			LabDepositTitles(x + 14).Left = VB6.TwipsToPixelsX(TempWide * (3480 / w))
			LabDepositTitles(x + 14).Width = VB6.TwipsToPixelsX(TempWide * (1635 / w))
			txtDepositValues(x + 14).Top = VB6.TwipsToPixelsY(TempHigh * (2010 / h) + (x * p))
			txtDepositValues(x + 14).Left = VB6.TwipsToPixelsX(TempWide * (5280 / w))
			txtDepositValues(x + 14).Width = VB6.TwipsToPixelsX(TempWide * (1215 / w))
			labDepositUnits(x + 14).Top = VB6.TwipsToPixelsY((TempHigh * (2040 / h)) + (x * p))
			labDepositUnits(x + 14).Left = VB6.TwipsToPixelsX(TempWide * (6540 / w))
		Next x
		
		lstDepositList.Top = VB6.TwipsToPixelsY(TempHigh * (4740 / h))
		lstDepositList.Left = VB6.TwipsToPixelsX(TempWide * (600 / w))
		lstDepositList.Height = VB6.TwipsToPixelsY(TempHigh * (510 / h))
		lstDepositList.Width = VB6.TwipsToPixelsX(TempWide * (2295 / w))
		
		LineHorizontal(0).X1 = VB6.TwipsToPixelsX(TempWide * (1980 / w))
		LineHorizontal(0).X2 = VB6.TwipsToPixelsX(TempWide * (3300 / w))
		LineHorizontal(0).Y1 = VB6.TwipsToPixelsY(TempHigh * (4200 / h))
		LineHorizontal(0).Y2 = VB6.TwipsToPixelsY(TempHigh * (4200 / h))
		
		LineHorizontal(1).X1 = VB6.TwipsToPixelsX(TempWide * (120 / w))
		LineHorizontal(1).X2 = VB6.TwipsToPixelsX(TempWide * (3300 / w))
		LineHorizontal(1).Y1 = VB6.TwipsToPixelsY(TempHigh * (5640 / h))
		LineHorizontal(1).Y2 = VB6.TwipsToPixelsY(TempHigh * (5640 / h))
		
		LineHorizontal(2).X1 = VB6.TwipsToPixelsX(TempWide * (3720 / w))
		LineHorizontal(2).X2 = VB6.TwipsToPixelsX(TempWide * (9060 / w))
		LineHorizontal(2).Y1 = VB6.TwipsToPixelsY(TempHigh * (180 / h))
		LineHorizontal(2).Y2 = VB6.TwipsToPixelsY(TempHigh * (180 / h))
		
		LineHorizontal(3).X1 = VB6.TwipsToPixelsX(TempWide * (3960 / w))
		LineHorizontal(3).X2 = VB6.TwipsToPixelsX(TempWide * (8940 / w))
		LineHorizontal(3).Y1 = VB6.TwipsToPixelsY(TempHigh * (3240 / h))
		LineHorizontal(3).Y2 = VB6.TwipsToPixelsY(TempHigh * (3240 / h))
		
		LineHorizontal(4).X1 = VB6.TwipsToPixelsX(TempWide * (3300 / w))
		LineHorizontal(4).X2 = VB6.TwipsToPixelsX(TempWide * (9060 / w))
		LineHorizontal(4).Y1 = VB6.TwipsToPixelsY(TempHigh * (6300 / h))
		LineHorizontal(4).Y2 = VB6.TwipsToPixelsY(TempHigh * (6300 / h))
		
		LineVertical(0).X1 = VB6.TwipsToPixelsX(TempWide * (180 / w))
		LineVertical(0).X2 = VB6.TwipsToPixelsX(TempWide * (180 / w))
		LineVertical(0).Y1 = VB6.TwipsToPixelsY(TempHigh * (4440 / h))
		LineVertical(0).Y2 = VB6.TwipsToPixelsY(TempHigh * (5700 / h))
		
		LineVertical(1).X1 = VB6.TwipsToPixelsX(TempWide * (3360 / w))
		LineVertical(1).X2 = VB6.TwipsToPixelsX(TempWide * (3360 / w))
		LineVertical(1).Y1 = VB6.TwipsToPixelsY(TempHigh * (420 / h))
		LineVertical(1).Y2 = VB6.TwipsToPixelsY(TempHigh * (3120 / h))
		
		LineVertical(2).X1 = VB6.TwipsToPixelsX(TempWide * (3360 / w))
		LineVertical(2).X2 = VB6.TwipsToPixelsX(TempWide * (3360 / w))
		LineVertical(2).Y1 = VB6.TwipsToPixelsY(TempHigh * (3360 / h))
		LineVertical(2).Y2 = VB6.TwipsToPixelsY(TempHigh * (6360 / h))
		
		LineVertical(3).X1 = VB6.TwipsToPixelsX(TempWide * (9000 / w))
		LineVertical(3).X2 = VB6.TwipsToPixelsX(TempWide * (9000 / w))
		LineVertical(3).Y1 = VB6.TwipsToPixelsY(TempHigh * (120 / h))
		LineVertical(3).Y2 = VB6.TwipsToPixelsY(TempHigh * (6360 / h))
		
		labBackToMenu.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labBackToMenu.Left = VB6.TwipsToPixelsX(TempWide * (600 / w))
		
		imgBackToMenu.Top = VB6.TwipsToPixelsY(TempHigh * (6150 / h))
		imgBackToMenu.Left = VB6.TwipsToPixelsX(TempWide * (60 / w))
		imgBackToMenu.Width = VB6.TwipsToPixelsX(TempWide * (495 / w))
		
		comDepositPrint.Top = VB6.TwipsToPixelsY(TempHigh * (6150 / h))
		comDepositPrint.Left = VB6.TwipsToPixelsX(TempWide * (3060 / w))
		
		labDepositHelp.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labDepositHelp.Left = VB6.TwipsToPixelsX(TempWide * (2580 / w))
		
		labInsert.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labInsert.Left = VB6.TwipsToPixelsX(TempWide * (1380 / w))
		labInsert.Width = VB6.TwipsToPixelsX(TempWide * (975 / w))
		
	End Sub
	
	Public Sub drawthevalues()
		
		Dim i As Short
        'Dim x As Short

        DoNotChange = True
		
		For i = 0 To 15
			Select Case i
				Case 0 To 4, 6 To 11, 13 To 15
					If CellValues(WhichScreen, i, WhichSegment).Changed = True Then
						txtDepositValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF0000)
					Else
						txtDepositValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
					End If
					Select Case i
						Case 0, 7
							txtDepositValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value * DensConv)), "##0.00")
						Case 2, 9
							txtDepositValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value * PowderConv)), "##0.00")
						Case 4, 11
							txtDepositValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value * FootConv)), "###,##0.0")
						Case 6, 13
							txtDepositValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value)), "##0.00")
						Case 14, 15
							txtDepositValues(i).Text = LTrim(RTrim(CellValues(WhichScreen, i, WhichSegment).Word))
						Case Else
							txtDepositValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value)), "##,###,##0")
					End Select
			End Select
		Next i
		
		DoNotChange = False
		
	End Sub
	
	Private Sub txtDepositValues_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtDepositValues.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Dim Index As Short = txtDepositValues.GetIndex(eventSender)
		
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
						If InStr(txtDepositValues(Index).Text, ".") = 0 Then
							System.Windows.Forms.SendKeys.Send("{DELETE}")
						End If
					Else
						System.Windows.Forms.SendKeys.Send("{DELETE}")
					End If
			End Select
		End If
		
	End Sub
	Private Sub txtDepositValues_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtDepositValues.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = txtDepositValues.GetIndex(eventSender)
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
	Private Sub txtDepositValues_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDepositValues.Leave
		Dim Index As Short = txtDepositValues.GetIndex(eventSender)
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
			Case 0 To 13
				tempvalue = ""
				For i = 1 To Len(txtDepositValues(Sample).Text)
					Digit.Value = Mid(txtDepositValues(Sample).Text, i, 1)
					Select Case Digit.Value
						Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ".", "-"
							tempvalue = tempvalue & Digit.Value
					End Select
				Next i
				If Sample = 0 Or Sample = 7 Then
					If CellValues(WhichScreen, Sample, WhichSegment).Changed = True Then
						If DensConv <> 0 Then CellValues(WhichScreen, Sample, WhichSegment).Value = Val(tempvalue) / DensConv
					End If
				ElseIf Sample = 2 Or Sample = 9 Then 
					If CellValues(WhichScreen, Sample, WhichSegment).Changed = True Then
						If PowderConv <> 0 Then CellValues(WhichScreen, Sample, WhichSegment).Value = Val(tempvalue) / PowderConv
					End If
				ElseIf Sample = 4 Or Sample = 5 Or Sample = 11 Or Sample = 12 Then 
					If CellValues(WhichScreen, Sample, WhichSegment).Changed = True Then
						If FootConv <> 0 Then CellValues(WhichScreen, Sample, WhichSegment).Value = Val(tempvalue) / FootConv
					End If
				Else
					If CellValues(WhichScreen, Sample, WhichSegment).Changed = True Then
						CellValues(WhichScreen, Sample, WhichSegment).Value = Val(tempvalue)
					End If
				End If
			Case Else
				CellValues(WhichScreen, Sample, WhichSegment).Word = txtDepositValues(Sample).Text
		End Select
		Call drawthevalues()
	End Sub
	'UPGRADE_WARNING: Event txtSegmentLabel.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtSegmentLabel_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSegmentLabel.TextChanged
		If DoNotChange = True Then Exit Sub
		SegNamie(WhichSegment) = txtSegmentLabel.Text
	End Sub
End Class