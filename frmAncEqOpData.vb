Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks
Friend Class frmAncEqOpData
	Inherits System.Windows.Forms.Form
	Dim TempHigh As Single
	Dim TempWide As Single
	Private Sub comAncEqOpPrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles comAncEqOpPrint.Click
		PrintForm1.Print(Me, PowerPacks.Printing.PrintForm.PrintOption.CompatibleModeClientAreaOnly)
	End Sub
	'UPGRADE_WARNING: Event chkAncEqOpItem.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkAncEqOpItem_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkAncEqOpItem.CheckStateChanged
		Dim Index As Short = chkAncEqOpItem.GetIndex(eventSender)
        'Dim x As Short
        If DoNotChange = False Then
			WhichScreen = Diesel + Index
			Call WhichBoxIsChecked()
			Call ScreenAdjust()
			Call drawthevalues()
			txtAncEqOpValues(6).Focus()
		End If
	End Sub
	'UPGRADE_WARNING: Form event frmAncEqOpData.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmAncEqOpData_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated

        'Dim baseunit As String
        'Dim baselength As Short
        Dim i As Short
		
		If IsHelpOn = True Then
			txtAncEqOpValues(WhichCell).Focus()
			IsHelpOn = False
		Else
			For i = 0 To 5
				If optSegment(i).Checked = True Then WhichSegment = i
			Next i
			
			Call drawthevalues()
			Call ScreenAdjust()
			
			If InsertFlag = True Then
				labInsert.Text = "Insert"
			Else
				labInsert.Text = "Typeover"
			End If
			
			Call WhichBoxIsChecked()
			
			WhichCell = 0
			
			txtAncEqOpValues(6).Focus()
		End If
		
	End Sub
	Private Sub frmAncEqOpData_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        'Dim x As Short

        Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - (VB6.PixelsToTwipsY(Me.Height) + 350)) / 2)
		Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2)
		
		If VB6.PixelsToTwipsY(Me.Top) < 0 Then Me.Top = 0
		If VB6.PixelsToTwipsX(Me.Left) < 0 Then Me.Left = 0
		
		TempWide = VB6.PixelsToTwipsX(Me.ClientRectangle.Width)
		TempHigh = VB6.PixelsToTwipsY(Me.ClientRectangle.Height)
		
		DoNotChange = True
		
		WhichSegment = 0
		optSegment(WhichSegment).Checked = True
		txtSegmentLabel.Text = SegNamie(WhichSegment)
		
		If PageChange(WhichScreen) = True Then
			Call drawthevalues()
		End If
		
		DoNotChange = False
		
		Call screenstuff()
		Call ScreenAdjust()
		
	End Sub
	'UPGRADE_WARNING: Event frmAncEqOpData.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub frmAncEqOpData_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		TempWide = VB6.PixelsToTwipsX(Me.ClientRectangle.Width)
		TempHigh = VB6.PixelsToTwipsY(Me.ClientRectangle.Height)
		Call screenstuff()
	End Sub
	Private Sub frmAncEqOpData_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
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
	Private Sub labAncEqOpHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles labAncEqOpHelp.Click
		Dim StartHelp As Short
		Dim SendHelp As Short
		StartHelp = 170
		IsHelpOn = True
		If WhichCell < 14 Then
			SendHelp = WhichCell - 6
		Else
			SendHelp = WhichScreen - 5
		End If
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
			txtAncEqOpValues(WhichCell).Focus()
		End If
	End Sub
	'UPGRADE_WARNING: Event txtAncEqOpValues.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtAncEqOpValues_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAncEqOpValues.TextChanged
		Dim Index As Short = txtAncEqOpValues.GetIndex(eventSender)
		If DoNotChange = True Then Exit Sub
	End Sub
	Private Sub txtAncEqOpValues_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAncEqOpValues.Enter
		Dim Index As Short = txtAncEqOpValues.GetIndex(eventSender)
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
		Dim u As Decimal
		
		Dim x As Short
		
		Dim y As Decimal
		Dim z As Decimal
		
		Dim h As Short
		Dim w As Short
		
		h = 6420
		w = 9150
		
		q = (480 / h) * TempHigh
		r = (420 / w) * TempWide
		s = (960 / h) * TempHigh
		t = (840 / h) * TempHigh
		u = (300 / h) * TempHigh
		y = (420 / h) * TempHigh
		z = (300 / w) * TempWide
		
		For x = 0 To 2
			labAncEqOpHeading(x).Top = VB6.TwipsToPixelsY(TempHigh * (120 / h) + (x * q))
			labAncEqOpHeading(x).Left = VB6.TwipsToPixelsX(TempWide * (120 / w) + (x * r))
			labAncEqOpHeading(x).Width = VB6.TwipsToPixelsX(TempWide * (1995 / w))
			If x < 2 Then
				labAncEqOpLabels(x).Top = VB6.TwipsToPixelsY((TempHigh * (1740 / h)) + (x * s))
				labAncEqOpLabels(x).Left = VB6.TwipsToPixelsX(TempWide * (600 / w))
				labAncEqOpLabels(x).Width = VB6.TwipsToPixelsX(TempWide * (1935 / w))
			End If
		Next x
		
		labAncEqOpLabels(2).Top = VB6.TwipsToPixelsY(TempHigh * (3540 / h))
		labAncEqOpLabels(2).Left = VB6.TwipsToPixelsX(TempWide * (60 / w))
		
		labAncEqOpLabels(3).Top = VB6.TwipsToPixelsY(TempHigh * (1200 / h))
		labAncEqOpLabels(3).Left = VB6.TwipsToPixelsX(TempWide * (3120 / w))
		
		For x = 5 To 7
			labAncEqOpLabels(x).Top = VB6.TwipsToPixelsY(TempHigh * (960 / h))
			If x = 5 Then
				labAncEqOpLabels(x).Left = VB6.TwipsToPixelsX(TempWide * (3300 / w))
				labAncEqOpLabels(x).Width = VB6.TwipsToPixelsX(TempWide * (1545 / w))
			ElseIf x = 6 Then 
				labAncEqOpLabels(x).Left = VB6.TwipsToPixelsX(TempWide * (5220 / w))
				labAncEqOpLabels(x).Width = VB6.TwipsToPixelsX(TempWide * (1425 / w))
			Else
				labAncEqOpLabels(x).Left = VB6.TwipsToPixelsX(TempWide * (6960 / w))
				labAncEqOpLabels(x).Width = VB6.TwipsToPixelsX(TempWide * (2055 / w))
			End If
		Next x
		
		For x = 0 To 5
			optSegment(x).Top = VB6.TwipsToPixelsY(TempHigh * (2040 / h))
			optSegment(x).Left = VB6.TwipsToPixelsX((TempWide * (720 / w)) + (x * z))
			optSegment(x).Width = VB6.TwipsToPixelsX(TempWide * (195 / w))
			labSegment(x).Top = VB6.TwipsToPixelsY(TempHigh * (2280 / h))
			labSegment(x).Left = VB6.TwipsToPixelsX((TempWide * (720 / w)) + (x * z))
			labSegment(x).Width = VB6.TwipsToPixelsX(TempWide * (195 / w))
		Next x
		
		txtSegmentLabel.Top = VB6.TwipsToPixelsY(TempHigh * (3000 / h))
		txtSegmentLabel.Left = VB6.TwipsToPixelsX(TempWide * (600 / w))
		txtSegmentLabel.Width = VB6.TwipsToPixelsX(TempWide * (1935 / w))
		
		For x = 0 To 7
			LabAncEqOpTitles(x + 6).Top = VB6.TwipsToPixelsY((TempHigh * (1500 / h)) + (x * y))
			LabAncEqOpTitles(x + 6).Left = VB6.TwipsToPixelsX(TempWide * (3240 / w))
			LabAncEqOpTitles(x + 6).Width = VB6.TwipsToPixelsX(TempWide * (1575 / w))
			txtAncEqOpValues(x + 6).Top = VB6.TwipsToPixelsY((TempHigh * (1470 / h)) + (x * y))
			txtAncEqOpValues(x + 6).Left = VB6.TwipsToPixelsX(TempWide * (5040 / w))
			txtAncEqOpValues(x + 6).Width = VB6.TwipsToPixelsX(TempWide * (855 / w))
			txtAncEqOpValues(x + 14).Top = VB6.TwipsToPixelsY((TempHigh * (1470 / h)) + (x * y))
			txtAncEqOpValues(x + 14).Left = VB6.TwipsToPixelsX(TempWide * (7020 / w))
			txtAncEqOpValues(x + 14).Width = VB6.TwipsToPixelsX(TempWide * (855 / w))
			labAncEqOpUnits(x + 6).Top = VB6.TwipsToPixelsY((TempHigh * (1500 / h)) + (x * y))
			labAncEqOpUnits(x + 6).Left = VB6.TwipsToPixelsX(TempWide * (5940 / w))
			labAncEqOpUnits(x + 14).Top = VB6.TwipsToPixelsY((TempHigh * (1500 / h)) + (x * y))
			labAncEqOpUnits(x + 14).Left = VB6.TwipsToPixelsX(TempWide * (7920 / w))
		Next x
		
		For x = 0 To 6
			chkAncEqOpItem(x).Top = VB6.TwipsToPixelsY((TempHigh * (3840 / h)) + (x * u))
			chkAncEqOpItem(x).Left = VB6.TwipsToPixelsX(TempWide * (240 / w))
			chkAncEqOpItem(x).Width = VB6.TwipsToPixelsX(TempWide * (2715 / w))
		Next x
		
		LineHorizontal(0).X1 = VB6.TwipsToPixelsX(TempWide * (1500 / w))
		LineHorizontal(0).X2 = VB6.TwipsToPixelsX(TempWide * (3060 / w))
		LineHorizontal(0).Y1 = VB6.TwipsToPixelsY(TempHigh * (3600 / h))
		LineHorizontal(0).Y2 = VB6.TwipsToPixelsY(TempHigh * (3600 / h))
		
		LineHorizontal(1).X1 = VB6.TwipsToPixelsX(TempWide * (60 / w))
		LineHorizontal(1).X2 = VB6.TwipsToPixelsX(TempWide * (3060 / w))
		LineHorizontal(1).Y1 = VB6.TwipsToPixelsY(TempHigh * (6000 / h))
		LineHorizontal(1).Y2 = VB6.TwipsToPixelsY(TempHigh * (6000 / h))
		
		LineHorizontal(2).X1 = VB6.TwipsToPixelsX(TempWide * (3120 / w))
		LineHorizontal(2).X2 = VB6.TwipsToPixelsX(TempWide * (9120 / w))
		LineHorizontal(2).Y1 = VB6.TwipsToPixelsY(TempHigh * (840 / h))
		LineHorizontal(2).Y2 = VB6.TwipsToPixelsY(TempHigh * (840 / h))
		
		LineHorizontal(3).X1 = VB6.TwipsToPixelsX(TempWide * (3840 / w))
		LineHorizontal(3).X2 = VB6.TwipsToPixelsX(TempWide * (9000 / w))
		LineHorizontal(3).Y1 = VB6.TwipsToPixelsY(TempHigh * (1260 / h))
		LineHorizontal(3).Y2 = VB6.TwipsToPixelsY(TempHigh * (1260 / h))
		
		LineHorizontal(4).X1 = VB6.TwipsToPixelsX(TempWide * (3120 / w))
		LineHorizontal(4).X2 = VB6.TwipsToPixelsX(TempWide * (9120 / w))
		LineHorizontal(4).Y1 = VB6.TwipsToPixelsY(TempHigh * (4860 / h))
		LineHorizontal(4).Y2 = VB6.TwipsToPixelsY(TempHigh * (4860 / h))
		
		LineVertical(0).X1 = VB6.TwipsToPixelsX(TempWide * (120 / w))
		LineVertical(0).X2 = VB6.TwipsToPixelsX(TempWide * (120 / w))
		LineVertical(0).Y1 = VB6.TwipsToPixelsY(TempHigh * (3840 / h))
		LineVertical(0).Y2 = VB6.TwipsToPixelsY(TempHigh * (6060 / h))
		
		LineVertical(1).X1 = VB6.TwipsToPixelsX(TempWide * (3000 / w))
		LineVertical(1).X2 = VB6.TwipsToPixelsX(TempWide * (3000 / w))
		LineVertical(1).Y1 = VB6.TwipsToPixelsY(TempHigh * (3540 / h))
		LineVertical(1).Y2 = VB6.TwipsToPixelsY(TempHigh * (6060 / h))
		
		LineVertical(2).X1 = VB6.TwipsToPixelsX(TempWide * (3180 / w))
		LineVertical(2).X2 = VB6.TwipsToPixelsX(TempWide * (3180 / w))
		LineVertical(2).Y1 = VB6.TwipsToPixelsY(TempHigh * (780 / h))
		LineVertical(2).Y2 = VB6.TwipsToPixelsY(TempHigh * (1140 / h))
		
		LineVertical(3).X1 = VB6.TwipsToPixelsX(TempWide * (3180 / w))
		LineVertical(3).X2 = VB6.TwipsToPixelsX(TempWide * (3180 / w))
		LineVertical(3).Y1 = VB6.TwipsToPixelsY(TempHigh * (1500 / h))
		LineVertical(3).Y2 = VB6.TwipsToPixelsY(TempHigh * (4920 / h))
		
		LineVertical(4).X1 = VB6.TwipsToPixelsX(TempWide * (4920 / w))
		LineVertical(4).X2 = VB6.TwipsToPixelsX(TempWide * (4920 / w))
		LineVertical(4).Y1 = VB6.TwipsToPixelsY(TempHigh * (900 / h))
		LineVertical(4).Y2 = VB6.TwipsToPixelsY(TempHigh * (1200 / h))
		
		LineVertical(5).X1 = VB6.TwipsToPixelsX(TempWide * (4920 / w))
		LineVertical(5).X2 = VB6.TwipsToPixelsX(TempWide * (4920 / w))
		LineVertical(5).Y1 = VB6.TwipsToPixelsY(TempHigh * (1320 / h))
		LineVertical(5).Y2 = VB6.TwipsToPixelsY(TempHigh * (4800 / h))
		
		LineVertical(6).X1 = VB6.TwipsToPixelsX(TempWide * (6900 / w))
		LineVertical(6).X2 = VB6.TwipsToPixelsX(TempWide * (6900 / w))
		LineVertical(6).Y1 = VB6.TwipsToPixelsY(TempHigh * (900 / h))
		LineVertical(6).Y2 = VB6.TwipsToPixelsY(TempHigh * (1200 / h))
		
		LineVertical(7).X1 = VB6.TwipsToPixelsX(TempWide * (6900 / w))
		LineVertical(7).X2 = VB6.TwipsToPixelsX(TempWide * (6900 / w))
		LineVertical(7).Y1 = VB6.TwipsToPixelsY(TempHigh * (1320 / h))
		LineVertical(7).Y2 = VB6.TwipsToPixelsY(TempHigh * (4800 / h))
		
		LineVertical(8).X1 = VB6.TwipsToPixelsX(TempWide * (9060 / w))
		LineVertical(8).X2 = VB6.TwipsToPixelsX(TempWide * (9060 / w))
		LineVertical(8).Y1 = VB6.TwipsToPixelsY(TempHigh * (780 / h))
		LineVertical(8).Y2 = VB6.TwipsToPixelsY(TempHigh * (4920 / h))
		
		labBackToMenu.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labBackToMenu.Left = VB6.TwipsToPixelsX(TempWide * (600 / w))
		
		imgBackToMenu.Top = VB6.TwipsToPixelsY(TempHigh * (6150 / h))
		imgBackToMenu.Left = VB6.TwipsToPixelsX(TempWide * (60 / w))
		imgBackToMenu.Width = VB6.TwipsToPixelsX(TempWide * (495 / w))
		
		comAncEqOpPrint.Top = VB6.TwipsToPixelsY(TempHigh * (6150 / h))
		comAncEqOpPrint.Left = VB6.TwipsToPixelsX(TempWide * (8880 / w))
		
		labAncEqOpHelp.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labAncEqOpHelp.Left = VB6.TwipsToPixelsX(TempWide * (8340 / w))
		
		labInsert.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labInsert.Left = VB6.TwipsToPixelsX(TempWide * (4080 / w))
		labInsert.Width = VB6.TwipsToPixelsX(TempWide * (975 / w))
		
	End Sub
	Private Sub ScreenAdjust()
		
		Dim p As Decimal
		Dim y As Decimal
		
		Dim x As Short
		Dim h As Short
		Dim w As Short
		
		On Error Resume Next
		
		h = 6420
		w = 9150
		y = (420 / h) * TempHigh
		
		For x = 0 To 7
			txtAncEqOpValues(x + 14).Top = VB6.TwipsToPixelsY((TempHigh * (1470 / h)) + (x * y))
			txtAncEqOpValues(x + 14).Width = VB6.TwipsToPixelsX(TempWide * (855 / w))
			labAncEqOpUnits(x + 14).Top = VB6.TwipsToPixelsY((TempHigh * (1500 / h)) + (x * y))
			Select Case WhichScreen
				Case Lubricants, RepairParts, Undercarriage
					txtAncEqOpValues(x + 14).Left = VB6.TwipsToPixelsX(TempWide * (7380 / w))
					labAncEqOpUnits(x + 14).Left = VB6.TwipsToPixelsX(TempWide * (8280 / w))
				Case Diesel
					If UnitType = English Then
						txtAncEqOpValues(x + 14).Left = VB6.TwipsToPixelsX(TempWide * (7020 / w))
						labAncEqOpUnits(x + 14).Left = VB6.TwipsToPixelsX(TempWide * (7920 / w))
					Else
						txtAncEqOpValues(x + 14).Left = VB6.TwipsToPixelsX(TempWide * (7140 / w))
						labAncEqOpUnits(x + 14).Left = VB6.TwipsToPixelsX(TempWide * (8040 / w))
					End If
				Case Electricity
					txtAncEqOpValues(x + 14).Left = VB6.TwipsToPixelsX(TempWide * (7140 / w))
					labAncEqOpUnits(x + 14).Left = VB6.TwipsToPixelsX(TempWide * (8040 / w))
				Case RepairLabor
					txtAncEqOpValues(x + 14).Left = VB6.TwipsToPixelsX(TempWide * (7080 / w))
					labAncEqOpUnits(x + 14).Left = VB6.TwipsToPixelsX(TempWide * (7980 / w))
				Case Else
					txtAncEqOpValues(x + 14).Left = VB6.TwipsToPixelsX(TempWide * (7440 / w))
					labAncEqOpUnits(x + 14).Left = VB6.TwipsToPixelsX(TempWide * (8340 / w))
			End Select
		Next x
		
	End Sub
	Public Sub drawthevalues()
		
		Dim i As Short
        'Dim x As Short
        'Dim p As Short
        'Dim r As Short
        'Dim tempvalue As Decimal

        DoNotChange = True
		
		For i = 6 To 13
			If CellValues(WhichScreen, i, WhichSegment).Changed = True Then
				txtAncEqOpValues(i + 8).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF0000)
			Else
				txtAncEqOpValues(i + 8).ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
			End If
			If CellValues(EquipmentHours, i, WhichSegment).Changed = True Then
				txtAncEqOpValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF0000)
			Else
				txtAncEqOpValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
			End If
			txtAncEqOpValues(i).Text = VB6.Format(LTrim(Str(CellValues(EquipmentHours, i, WhichSegment).Value)), "###,##0.00")
		Next i
		
		If UnitType = Metric Then
			Select Case WhichScreen
				Case Diesel
					For i = 0 To 7
						labAncEqOpUnits(i + 14).Text = "liters/hour"
						txtAncEqOpValues(i + 14).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i + 6, WhichSegment).Value * GallonConv)), "###,##0.00")
					Next i
				Case Electricity
					For i = 0 To 7
						labAncEqOpUnits(i + 14).Text = "kWh/hour"
						txtAncEqOpValues(i + 14).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i + 6, WhichSegment).Value)), "###,###,##0")
					Next i
				Case Lubricants, RepairParts, Undercarriage
					For i = 0 To 7
						labAncEqOpUnits(i + 14).Text = "/hour"
						txtAncEqOpValues(i + 14).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i + 6, WhichSegment).Value)), "$###,##0.00")
					Next i
				Case RepairLabor
					For i = 0 To 7
						labAncEqOpUnits(i + 14).Text = "hours/hour"
						txtAncEqOpValues(i + 14).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i + 6, WhichSegment).Value)), "#,##0.000")
					Next i
				Case Tires
					For i = 0 To 7
						labAncEqOpUnits(i + 14).Text = "/set"
						txtAncEqOpValues(i + 14).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i + 6, WhichSegment).Value)), "$#,###,##0")
					Next i
			End Select
		Else
			Select Case WhichScreen
				Case Diesel
					For i = 0 To 7
						labAncEqOpUnits(i + 14).Text = "gallons/hour"
						txtAncEqOpValues(i + 14).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i + 6, WhichSegment).Value)), "###,##0.00")
					Next i
				Case Electricity
					For i = 0 To 7
						labAncEqOpUnits(i + 14).Text = "kWh/hour"
						txtAncEqOpValues(i + 14).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i + 6, WhichSegment).Value)), "###,###,##0")
					Next i
				Case Lubricants, RepairParts, Undercarriage
					For i = 0 To 7
						labAncEqOpUnits(i + 14).Text = "/hour"
						txtAncEqOpValues(i + 14).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i + 6, WhichSegment).Value)), "$###,##0.00")
					Next i
				Case RepairLabor
					For i = 0 To 7
						labAncEqOpUnits(i + 14).Text = "hours/hour"
						txtAncEqOpValues(i + 14).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i + 6, WhichSegment).Value)), "#,##0.000")
					Next i
				Case Tires
					For i = 0 To 7
						labAncEqOpUnits(i + 14).Text = "/set"
						txtAncEqOpValues(i + 14).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i + 6, WhichSegment).Value)), "$#,###,##0")
					Next i
			End Select
		End If
		
		Select Case WhichScreen
			Case Diesel
				labAncEqOpLabels(7).Text = "Diesel Consumption"
			Case Electricity
				labAncEqOpLabels(7).Text = "Electrical Consumption"
			Case Lubricants
				labAncEqOpLabels(7).Text = "Lubricant Costs"
			Case RepairParts
				labAncEqOpLabels(7).Text = "Repair Parts Costs"
			Case Undercarriage
				labAncEqOpLabels(7).Text = "Undercarriage Costs"
			Case RepairLabor
				labAncEqOpLabels(7).Text = "Repair Labor"
			Case Tires
				labAncEqOpLabels(7).Text = "Tire Prices"
		End Select
		
		For i = 0 To 7
			labAncEqOpUnits(i + 6).Text = "hours/day"
		Next i
		
		If bltp = 1 Then
			LabAncEqOpTitles(10).Text = "Powder Buggies"
		ElseIf bltp = 2 Then 
			LabAncEqOpTitles(10).Text = "Bulk Trucks"
		End If
		
		
		DoNotChange = False
		
	End Sub
	Private Sub txtAncEqOpValues_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtAncEqOpValues.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Dim Index As Short = txtAncEqOpValues.GetIndex(eventSender)
		
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
						If InStr(txtAncEqOpValues(Index).Text, ".") = 0 Then
							System.Windows.Forms.SendKeys.Send("{DELETE}")
						End If
					Else
						System.Windows.Forms.SendKeys.Send("{DELETE}")
					End If
			End Select
		End If
		
	End Sub
	Private Sub txtAncEqOpValues_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtAncEqOpValues.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = txtAncEqOpValues.GetIndex(eventSender)
		If KeyAscii > Asc("9") And KeyAscii <> Asc(",") And KeyAscii <> Asc(".") And KeyAscii <> Asc("$") Then
			Beep()
			KeyAscii = 0
		Else
			Select Case WhichCell
				Case 6 To 13
					CellValues(EquipmentHours, WhichCell, WhichSegment).Changed = True
				Case 14 To 21
					CellValues(WhichScreen, WhichCell - 8, WhichSegment).Changed = True
			End Select
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	Private Sub txtAncEqOpValues_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAncEqOpValues.Leave
		Dim Index As Short = txtAncEqOpValues.GetIndex(eventSender)
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
		For i = 1 To Len(txtAncEqOpValues(Sample).Text)
			Digit.Value = Mid(txtAncEqOpValues(Sample).Text, i, 1)
			Select Case Digit.Value
				Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ".", "-"
					tempvalue = tempvalue & Digit.Value
			End Select
		Next i
		Select Case Sample
			Case 6 To 13
				If CellValues(EquipmentHours, Sample, WhichSegment).Changed = True Then
					CellValues(EquipmentHours, Sample, WhichSegment).Value = Val(tempvalue)
				End If
			Case Else
				Select Case WhichScreen
					Case Diesel
						If CellValues(WhichScreen, Sample - 8, WhichSegment).Changed = True Then
							CellValues(WhichScreen, Sample - 8, WhichSegment).Value = Val(tempvalue) / GallonConv
						End If
					Case Else
						If CellValues(WhichScreen, Sample - 8, WhichSegment).Changed = True Then
							CellValues(WhichScreen, Sample - 8, WhichSegment).Value = Val(tempvalue)
						End If
				End Select
		End Select
		Call drawthevalues()
	End Sub
	'UPGRADE_WARNING: Event txtSegmentLabel.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtSegmentLabel_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSegmentLabel.TextChanged
		If DoNotChange = True Then Exit Sub
		SegNamie(WhichSegment) = txtSegmentLabel.Text
	End Sub
	Private Sub WhichBoxIsChecked()
		Dim TheScreen As Short
		
		On Error Resume Next
		
		DoNotChange = True
		For TheScreen = Diesel To Tires
			chkAncEqOpItem(TheScreen - Diesel).CheckState = System.Windows.Forms.CheckState.Unchecked
			chkAncEqOpItem(TheScreen - Diesel).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFFFF)
		Next TheScreen
		
		chkAncEqOpItem(WhichScreen - Diesel).CheckState = System.Windows.Forms.CheckState.Checked
		chkAncEqOpItem(WhichScreen - Diesel).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF00)
		
		DoNotChange = False
	End Sub
End Class