Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks
Friend Class frmDevelopmentData
	Inherits System.Windows.Forms.Form
	Dim TempHigh As Single
	Dim TempWide As Single
	Private Sub comDevelopmentPrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles comDevelopmentPrint.Click
		PrintForm1.Print(Me, PowerPacks.Printing.PrintForm.PrintOption.CompatibleModeClientAreaOnly)
	End Sub
	'UPGRADE_WARNING: Form event frmDevelopmentData.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmDevelopmentData_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated

        'Dim baseunit As String
        'Dim baselength As Short
        Dim i As Short
		
		If IsHelpOn = True Then
			txtDevelopmentValues(WhichCell).Focus()
			IsHelpOn = False
		Else
			For i = 0 To 5
				If optSegment(i).Checked = True Then WhichSegment = i
			Next i
			
			WhichScreen = Development
			
			Call drawthevalues()
			
			If InsertFlag = True Then
				labInsert.Text = "Insert"
			Else
				labInsert.Text = "Typeover"
			End If
			
			WhichCell = 0
			
			txtDevelopmentValues(0).Focus()
		End If
		
	End Sub
	Private Sub frmDevelopmentData_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        'Dim x As Short

        Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - (VB6.PixelsToTwipsY(Me.Height) + 350)) / 2)
		Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2)
		
		If VB6.PixelsToTwipsY(Me.Top) < 0 Then Me.Top = 0
		If VB6.PixelsToTwipsX(Me.Left) < 0 Then Me.Left = 0
		
		TempWide = VB6.PixelsToTwipsX(Me.ClientRectangle.Width)
		TempHigh = VB6.PixelsToTwipsY(Me.ClientRectangle.Height)
		
		DoNotChange = True
		
		If UnitType = Metric Then
			labDevelopmentUnits(0).Text = "meters"
			labDevelopmentUnits(1).Text = "meters"
			labDevelopmentUnits(2).Text = "square meters"
			labDevelopmentUnits(3).Text = "square meters"
			labDevelopmentUnits(4).Text = "square meters"
			labDevelopmentUnits(5).Text = "square meters"
			labDevelopmentUnits(6).Text = "cubic meters"
			labDevelopmentUnits(7).Text = "cubic meters"
			labDevelopmentUnits(9).Text = "hectares"
			labDevelopmentUnits(10).Text = "square meters"
			labDevelopmentUnits(22).Text = "liters/day"
			labDevelopmentUnits(24).Text = "meters"
			labDevelopmentUnits(26).Text = "liters"
			labDevelopmentUnits(11).Text = "/meter"
			labDevelopmentUnits(12).Text = "/meter"
			labDevelopmentUnits(13).Text = "/square meter"
			labDevelopmentUnits(14).Text = "/square meter"
			labDevelopmentUnits(15).Text = "/square meter"
			labDevelopmentUnits(16).Text = "/square meter"
			labDevelopmentUnits(20).Text = "/hectare"
			labDevelopmentUnits(21).Text = "/square meter"
			labDevelopmentUnits(25).Text = "/meter"
		Else
			labDevelopmentUnits(0).Text = "feet"
			labDevelopmentUnits(1).Text = "feet"
			labDevelopmentUnits(2).Text = "square feet"
			labDevelopmentUnits(3).Text = "square feet"
			labDevelopmentUnits(4).Text = "square feet"
			labDevelopmentUnits(5).Text = "square feet"
			labDevelopmentUnits(6).Text = "cubic feet"
			labDevelopmentUnits(7).Text = "cubic feet"
			labDevelopmentUnits(9).Text = "acres"
			labDevelopmentUnits(10).Text = "square feet"
			labDevelopmentUnits(22).Text = "gallons/day"
			labDevelopmentUnits(24).Text = "feet"
			labDevelopmentUnits(26).Text = "gallons"
			labDevelopmentUnits(11).Text = "/foot"
			labDevelopmentUnits(12).Text = "/foot"
			labDevelopmentUnits(13).Text = "/square foot"
			labDevelopmentUnits(14).Text = "/square foot"
			labDevelopmentUnits(15).Text = "/square foot"
			labDevelopmentUnits(16).Text = "/square foot"
			labDevelopmentUnits(20).Text = "/acre"
			labDevelopmentUnits(21).Text = "/square foot"
			labDevelopmentUnits(25).Text = "/foot"
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
	'UPGRADE_WARNING: Event frmDevelopmentData.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub frmDevelopmentData_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		TempWide = VB6.PixelsToTwipsX(Me.ClientRectangle.Width)
		TempHigh = VB6.PixelsToTwipsY(Me.ClientRectangle.Height)
		Call screenstuff()
	End Sub
	Private Sub frmDevelopmentData_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
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
	Private Sub labDevelopmentHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles labDevelopmentHelp.Click
		Dim StartHelp As Short
		Dim SendHelp As Short
		StartHelp = 277
		IsHelpOn = True
		Select Case WhichCell
			Case 0 To 10
				SendHelp = WhichCell
			Case 11 To 21
				SendHelp = WhichCell + 3
			Case 22, 24, 26
				SendHelp = WhichCell / 2
			Case 23, 25, 27
				SendHelp = ((WhichCell - 1) / 2) + 14
		End Select
		Call frmSurfaceHelp.gethelptext(StartHelp, SendHelp)
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
			txtDevelopmentValues(WhichCell).Focus()
		End If
	End Sub
	'UPGRADE_WARNING: Event txtDevelopmentValues.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtDevelopmentValues_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDevelopmentValues.TextChanged
		Dim Index As Short = txtDevelopmentValues.GetIndex(eventSender)
		If DoNotChange = True Then Exit Sub
	End Sub
	Private Sub txtDevelopmentValues_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDevelopmentValues.Enter
		Dim Index As Short = txtDevelopmentValues.GetIndex(eventSender)
        'Dim x As Short
        WhichCell = Index
		System.Windows.Forms.SendKeys.Send("{HOME}+{END}")
		Call drawthevalues()
	End Sub
	Public Sub screenstuff()

        'Dim p As Decimal
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
		
		q = (1440 / h) * TempHigh
		r = (540 / w) * TempWide
		s = (1620 / h) * TempHigh
		t = (840 / h) * TempHigh
		
		y = (360 / h) * TempHigh
		z = (300 / w) * TempWide
		
		labDevelopmentHeading.Top = VB6.TwipsToPixelsY(TempHigh * (120 / h))
		labDevelopmentHeading.Left = VB6.TwipsToPixelsX(TempWide * (120 / w))
		labDevelopmentHeading.Width = VB6.TwipsToPixelsX(TempWide * (2475 / w))
		
		For x = 0 To 1
			labDevelopmentLabels(x).Top = VB6.TwipsToPixelsY(TempHigh * (420 / h))
			labDevelopmentLabels(x + 2).Top = VB6.TwipsToPixelsY(TempHigh * ((2040 / h)) + (x * q))
			labDevelopmentLabels(x + 2).Left = VB6.TwipsToPixelsX(TempWide * (120 / w))
			labDevelopmentLabels(x + 2).Width = VB6.TwipsToPixelsX(TempWide * (1935 / w))
			If x = 0 Then
				labDevelopmentLabels(x).Left = VB6.TwipsToPixelsX(TempWide * (4140 / w))
				labDevelopmentLabels(x).Width = VB6.TwipsToPixelsX(TempWide * (2115 / w))
			ElseIf x = 1 Then 
				labDevelopmentLabels(x).Left = VB6.TwipsToPixelsX(TempWide * (6600 / w))
				labDevelopmentLabels(x).Width = VB6.TwipsToPixelsX(TempWide * (2265 / w))
			End If
		Next x
		
		For x = 0 To 5
			optSegment(x).Top = VB6.TwipsToPixelsY(TempHigh * (2340 / h))
			optSegment(x).Left = VB6.TwipsToPixelsX((TempWide * (240 / w)) + (x * z))
			optSegment(x).Width = VB6.TwipsToPixelsX(TempWide * (195 / w))
			labSegment(x).Top = VB6.TwipsToPixelsY(TempHigh * (2580 / h))
			labSegment(x).Left = VB6.TwipsToPixelsX((TempWide * (240 / w)) + (x * z))
			labSegment(x).Width = VB6.TwipsToPixelsX(TempWide * (195 / w))
		Next x
		
		txtSegmentLabel.Top = VB6.TwipsToPixelsY(TempHigh * (3780 / h))
		txtSegmentLabel.Left = VB6.TwipsToPixelsX(TempWide * (120 / w))
		txtSegmentLabel.Width = VB6.TwipsToPixelsX(TempWide * (1935 / w))
		
		For x = 0 To 13
			LabDevelopmentTitles(x).Top = VB6.TwipsToPixelsY((TempHigh * (900 / h)) + (x * y))
			LabDevelopmentTitles(x).Left = VB6.TwipsToPixelsX(TempWide * (2280 / w))
			LabDevelopmentTitles(x).Width = VB6.TwipsToPixelsX(TempWide * (1575 / w))
			If x < 11 Then
				txtDevelopmentValues(x).Top = VB6.TwipsToPixelsY((TempHigh * (870 / h)) + (x * y))
				txtDevelopmentValues(x).Left = VB6.TwipsToPixelsX(TempWide * (4140 / w))
				txtDevelopmentValues(x).Width = VB6.TwipsToPixelsX(TempWide * (915 / w))
				labDevelopmentUnits(x).Top = VB6.TwipsToPixelsY((TempHigh * (900 / h)) + (x * y))
				labDevelopmentUnits(x).Left = VB6.TwipsToPixelsX(TempWide * (5100 / w))
				
				txtDevelopmentValues(x + 11).Top = VB6.TwipsToPixelsY((TempHigh * (870 / h)) + (x * y))
				txtDevelopmentValues(x + 11).Left = VB6.TwipsToPixelsX(TempWide * (6600 / w))
				txtDevelopmentValues(x + 11).Width = VB6.TwipsToPixelsX(TempWide * (1095 / w))
				labDevelopmentUnits(x + 11).Top = VB6.TwipsToPixelsY((TempHigh * (900 / h)) + (x * y))
				labDevelopmentUnits(x + 11).Left = VB6.TwipsToPixelsX(TempWide * (7740 / w))
			End If
		Next x
		
		For x = 0 To 4 Step 2
			txtDevelopmentValues(x + 22).Top = VB6.TwipsToPixelsY((TempHigh * (4830 / h)) + ((x / 2) * y))
			txtDevelopmentValues(x + 22).Left = VB6.TwipsToPixelsX(TempWide * (4140 / w))
			txtDevelopmentValues(x + 22).Width = VB6.TwipsToPixelsX(TempWide * (915 / w))
			labDevelopmentUnits(x + 22).Top = VB6.TwipsToPixelsY((TempHigh * (4860 / h)) + ((x / 2) * y))
			labDevelopmentUnits(x + 22).Left = VB6.TwipsToPixelsX(TempWide * (5100 / w))
			txtDevelopmentValues(x + 23).Top = VB6.TwipsToPixelsY((TempHigh * (4830 / h)) + ((x / 2) * y))
			txtDevelopmentValues(x + 23).Left = VB6.TwipsToPixelsX(TempWide * (6600 / w))
			txtDevelopmentValues(x + 23).Width = VB6.TwipsToPixelsX(TempWide * (1095 / w))
			labDevelopmentUnits(x + 23).Top = VB6.TwipsToPixelsY((TempHigh * (4860 / h)) + ((x / 2) * y))
			labDevelopmentUnits(x + 23).Left = VB6.TwipsToPixelsX(TempWide * (7740 / w))
		Next x
		
		LineHorizontal(0).X1 = VB6.TwipsToPixelsX(TempWide * (3900 / w))
		LineHorizontal(0).X2 = VB6.TwipsToPixelsX(TempWide * (9060 / w))
		LineHorizontal(0).Y1 = VB6.TwipsToPixelsY(TempHigh * (360 / h))
		LineHorizontal(0).Y2 = VB6.TwipsToPixelsY(TempHigh * (360 / h))
		
		LineHorizontal(1).X1 = VB6.TwipsToPixelsX(TempWide * (2100 / w))
		LineHorizontal(1).X2 = VB6.TwipsToPixelsX(TempWide * (3900 / w))
		LineHorizontal(1).Y1 = VB6.TwipsToPixelsY(TempHigh * (720 / h))
		LineHorizontal(1).Y2 = VB6.TwipsToPixelsY(TempHigh * (720 / h))
		
		LineHorizontal(2).X1 = VB6.TwipsToPixelsX(TempWide * (4020 / w))
		LineHorizontal(2).X2 = VB6.TwipsToPixelsX(TempWide * (6360 / w))
		LineHorizontal(2).Y1 = VB6.TwipsToPixelsY(TempHigh * (720 / h))
		LineHorizontal(2).Y2 = VB6.TwipsToPixelsY(TempHigh * (720 / h))
		
		LineHorizontal(3).X1 = VB6.TwipsToPixelsX(TempWide * (6480 / w))
		LineHorizontal(3).X2 = VB6.TwipsToPixelsX(TempWide * (8940 / w))
		LineHorizontal(3).Y1 = VB6.TwipsToPixelsY(TempHigh * (720 / h))
		LineHorizontal(3).Y2 = VB6.TwipsToPixelsY(TempHigh * (720 / h))
		
		LineHorizontal(4).X1 = VB6.TwipsToPixelsX(TempWide * (2100 / w))
		LineHorizontal(4).X2 = VB6.TwipsToPixelsX(TempWide * (9060 / w))
		LineHorizontal(4).Y1 = VB6.TwipsToPixelsY(TempHigh * (6000 / h))
		LineHorizontal(4).Y2 = VB6.TwipsToPixelsY(TempHigh * (6000 / h))
		
		LineVertical(0).X1 = VB6.TwipsToPixelsX(TempWide * (2160 / w))
		LineVertical(0).X2 = VB6.TwipsToPixelsX(TempWide * (2160 / w))
		LineVertical(0).Y1 = VB6.TwipsToPixelsY(TempHigh * (660 / h))
		LineVertical(0).Y2 = VB6.TwipsToPixelsY(TempHigh * (6060 / h))
		
		LineVertical(1).X1 = VB6.TwipsToPixelsX(TempWide * (3960 / w))
		LineVertical(1).X2 = VB6.TwipsToPixelsX(TempWide * (3960 / w))
		LineVertical(1).Y1 = VB6.TwipsToPixelsY(TempHigh * (300 / h))
		LineVertical(1).Y2 = VB6.TwipsToPixelsY(TempHigh * (660 / h))
		
		LineVertical(2).X1 = VB6.TwipsToPixelsX(TempWide * (3960 / w))
		LineVertical(2).X2 = VB6.TwipsToPixelsX(TempWide * (3960 / w))
		LineVertical(2).Y1 = VB6.TwipsToPixelsY(TempHigh * (780 / h))
		LineVertical(2).Y2 = VB6.TwipsToPixelsY(TempHigh * (5940 / h))
		
		LineVertical(3).X1 = VB6.TwipsToPixelsX(TempWide * (6420 / w))
		LineVertical(3).X2 = VB6.TwipsToPixelsX(TempWide * (6420 / w))
		LineVertical(3).Y1 = VB6.TwipsToPixelsY(TempHigh * (420 / h))
		LineVertical(3).Y2 = VB6.TwipsToPixelsY(TempHigh * (660 / h))
		
		LineVertical(4).X1 = VB6.TwipsToPixelsX(TempWide * (6420 / w))
		LineVertical(4).X2 = VB6.TwipsToPixelsX(TempWide * (6420 / w))
		LineVertical(4).Y1 = VB6.TwipsToPixelsY(TempHigh * (780 / h))
		LineVertical(4).Y2 = VB6.TwipsToPixelsY(TempHigh * (5940 / h))
		
		LineVertical(5).X1 = VB6.TwipsToPixelsX(TempWide * (9000 / w))
		LineVertical(5).X2 = VB6.TwipsToPixelsX(TempWide * (9000 / w))
		LineVertical(5).Y1 = VB6.TwipsToPixelsY(TempHigh * (300 / h))
		LineVertical(5).Y2 = VB6.TwipsToPixelsY(TempHigh * (6060 / h))
		
		labBackToMenu.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labBackToMenu.Left = VB6.TwipsToPixelsX(TempWide * (600 / w))
		
		imgBackToMenu.Top = VB6.TwipsToPixelsY(TempHigh * (6150 / h))
		imgBackToMenu.Left = VB6.TwipsToPixelsX(TempWide * (60 / w))
		imgBackToMenu.Width = VB6.TwipsToPixelsX(TempWide * (495 / w))
		
		comDevelopmentPrint.Top = VB6.TwipsToPixelsY(TempHigh * (6150 / h))
		comDevelopmentPrint.Left = VB6.TwipsToPixelsX(TempWide * (8880 / w))
		
		labDevelopmentHelp.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labDevelopmentHelp.Left = VB6.TwipsToPixelsX(TempWide * (8340 / w))
		
		labInsert.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labInsert.Left = VB6.TwipsToPixelsX(TempWide * (4080 / w))
		labInsert.Width = VB6.TwipsToPixelsX(TempWide * (975 / w))
		
	End Sub
	Public Sub drawthevalues()
		
		Dim i As Short
        'Dim x As Short

        DoNotChange = True
		
		For i = 0 To 27
			If CellValues(WhichScreen, i, WhichSegment).Changed = True Then
				txtDevelopmentValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF0000)
			Else
				txtDevelopmentValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
			End If
			Select Case i
				Case 0, 1, 24
					txtDevelopmentValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value * FootConv)), "###,###,##0")
				Case 2, 3, 4, 5, 10
					txtDevelopmentValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value * SquareFootConv)), "###,###,##0")
				Case 6, 7
					txtDevelopmentValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value * CubicFootConv)), "###,###,##0")
				Case 8
					txtDevelopmentValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value)), "###,###,##0")
				Case 9
					txtDevelopmentValues(i).Text = VB6.Format(CDbl(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value))) * AcreConv, "##,###,##0.00")
				Case 11, 12, 25
					txtDevelopmentValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value / FootConv)), "$##,###,##0.00")
				Case 13, 14, 15, 16, 21
					txtDevelopmentValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value / SquareFootConv)), "$##,###,##0.00")
				Case 20
					txtDevelopmentValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value / AcreConv)), "$##,###,##0.00")
				Case 22, 26
					txtDevelopmentValues(i).Text = VB6.Format(CDbl(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value))) * GallonConv, "##,###,##0")
				Case Else
					txtDevelopmentValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value)), "$#,###,###,##0")
			End Select
		Next i
		
		DoNotChange = False
		
	End Sub
	
	Private Sub txtDevelopmentValues_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtDevelopmentValues.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Dim Index As Short = txtDevelopmentValues.GetIndex(eventSender)
		
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
						If InStr(txtDevelopmentValues(Index).Text, ".") = 0 Then
							System.Windows.Forms.SendKeys.Send("{DELETE}")
						End If
					Else
						System.Windows.Forms.SendKeys.Send("{DELETE}")
					End If
			End Select
		End If
		
	End Sub
	Private Sub txtDevelopmentValues_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtDevelopmentValues.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = txtDevelopmentValues.GetIndex(eventSender)
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
	Private Sub txtDevelopmentValues_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDevelopmentValues.Leave
		Dim Index As Short = txtDevelopmentValues.GetIndex(eventSender)
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
		For i = 1 To Len(txtDevelopmentValues(Sample).Text)
			Digit.Value = Mid(txtDevelopmentValues(Sample).Text, i, 1)
			Select Case Digit.Value
				Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ".", "-"
					tempvalue = tempvalue & Digit.Value
			End Select
		Next i
		If CellValues(WhichScreen, Sample, WhichSegment).Changed = True Then
			Select Case Sample
				Case 0, 1, 24
					CellValues(WhichScreen, Sample, WhichSegment).Value = Val(CStr(CDbl(tempvalue) / FootConv))
				Case 2, 3, 4, 5, 10
					CellValues(WhichScreen, Sample, WhichSegment).Value = Val(CStr(CDbl(tempvalue) / SquareFootConv))
				Case 6, 7
					CellValues(WhichScreen, Sample, WhichSegment).Value = Val(CStr(CDbl(tempvalue) / CubicFootConv))
				Case 9
					CellValues(WhichScreen, Sample, WhichSegment).Value = Val(CStr(CDbl(tempvalue) / AcreConv))
				Case 11, 12, 25
					CellValues(WhichScreen, Sample, WhichSegment).Value = Val(CStr(CDbl(tempvalue) * FootConv))
				Case 13, 14, 15, 16, 21
					CellValues(WhichScreen, Sample, WhichSegment).Value = Val(CStr(CDbl(tempvalue) * SquareFootConv))
				Case 20
					CellValues(WhichScreen, Sample, WhichSegment).Value = Val(CStr(CDbl(tempvalue) * AcreConv))
				Case 22, 26
					CellValues(WhichScreen, Sample, WhichSegment).Value = Val(CStr(CDbl(tempvalue) / GallonConv))
				Case Else
					CellValues(WhichScreen, Sample, WhichSegment).Value = Val(tempvalue)
			End Select
		End If
		Call drawthevalues()
	End Sub
	
	'UPGRADE_WARNING: Event txtSegmentLabel.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtSegmentLabel_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSegmentLabel.TextChanged
		If DoNotChange = True Then Exit Sub
		SegNamie(WhichSegment) = txtSegmentLabel.Text
	End Sub
End Class