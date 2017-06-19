Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks
Friend Class frmProdEqOpData
	Inherits System.Windows.Forms.Form
	Dim TempHigh As Single
	Dim TempWide As Single
	Private Sub comProdEqOpPrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles comProdEqOpPrint.Click
		PrintForm1.Print(Me, PowerPacks.Printing.PrintForm.PrintOption.CompatibleModeClientAreaOnly)
	End Sub
	'UPGRADE_WARNING: Event chkProdEqOpItem.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkProdEqOpItem_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkProdEqOpItem.CheckStateChanged
		Dim Index As Short = chkProdEqOpItem.GetIndex(eventSender)
        'Dim x As Short
        If DoNotChange = False Then
			WhichScreen = Diesel + Index
			Call WhichBoxIsChecked()
			Call ScreenAdjust()
			Call drawthevalues()
			txtProdEqOpValues(0).Focus()
		End If
	End Sub
	'UPGRADE_WARNING: Form event frmProdEqOpData.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmProdEqOpData_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated

        'Dim baseunit As String
        'Dim baselength As Short
        Dim i As Short
		
		If IsHelpOn = True Then
			txtProdEqOpValues(WhichCell).Focus()
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
			
			txtProdEqOpValues(0).Focus()
		End If
		
	End Sub
	Private Sub frmProdEqOpData_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
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
	'UPGRADE_WARNING: Event frmProdEqOpData.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub frmProdEqOpData_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		TempWide = VB6.PixelsToTwipsX(Me.ClientRectangle.Width)
		TempHigh = VB6.PixelsToTwipsY(Me.ClientRectangle.Height)
		Call screenstuff()
	End Sub
	Private Sub frmProdEqOpData_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
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
	Private Sub labProdEqOpHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles labProdEqOpHelp.Click
		Dim StartHelp As Short
		Dim SendHelp As Short
		IsHelpOn = True
		StartHelp = 158
		Select Case WhichCell
			Case 0, 1
				SendHelp = WhichCell
			Case 2, 3
				SendHelp = WhichCell + 4
			Case 4
				SendHelp = 5
			Case 5
				SendHelp = 4
			Case 14, 15
				SendHelp = WhichCell - 12
			Case 16, 17
				SendHelp = WhichCell - 8
			Case 18
				SendHelp = 11
			Case 19
				SendHelp = 10
			Case Else
				SendHelp = WhichScreen + 7
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
			txtProdEqOpValues(WhichCell).Focus()
		End If
	End Sub
	'UPGRADE_WARNING: Event txtProdEqOpValues.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtProdEqOpValues_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtProdEqOpValues.TextChanged
		Dim Index As Short = txtProdEqOpValues.GetIndex(eventSender)
		If DoNotChange = True Then Exit Sub
	End Sub
	Private Sub txtProdEqOpValues_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtProdEqOpValues.Enter
		Dim Index As Short = txtProdEqOpValues.GetIndex(eventSender)
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
		t = (2700 / h) * TempHigh
		u = (300 / h) * TempHigh
		y = (420 / h) * TempHigh
		z = (300 / w) * TempWide
		
		For x = 0 To 2
			labProdEqOpHeading(x).Top = VB6.TwipsToPixelsY(TempHigh * (120 / h) + (x * q))
			labProdEqOpHeading(x).Left = VB6.TwipsToPixelsX(TempWide * (120 / w) + (x * r))
			labProdEqOpHeading(x).Width = VB6.TwipsToPixelsX(TempWide * (1995 / w))
			If x < 2 Then
				labProdEqOpLabels(x).Top = VB6.TwipsToPixelsY((TempHigh * (1740 / h)) + (x * s))
				labProdEqOpLabels(x).Left = VB6.TwipsToPixelsX(TempWide * (600 / w))
				labProdEqOpLabels(x).Width = VB6.TwipsToPixelsX(TempWide * (1935 / w))
			End If
		Next x
		
		labProdEqOpLabels(x + 2).Top = VB6.TwipsToPixelsY(TempHigh * (3540 / h))
		labProdEqOpLabels(x + 2).Left = VB6.TwipsToPixelsX(TempWide * (60 / w))
		
		For x = 0 To 1
			labProdEqOpLabels(x + 3).Top = VB6.TwipsToPixelsY((TempHigh * (480 / h)) + (x * t))
			labProdEqOpLabels(x + 3).Left = VB6.TwipsToPixelsX(TempWide * (3120 / w))
		Next x
		
		For x = 5 To 7
			labProdEqOpLabels(x).Top = VB6.TwipsToPixelsY(TempHigh * (240 / h))
			If x = 5 Then
				labProdEqOpLabels(x).Left = VB6.TwipsToPixelsX(TempWide * (3240 / w))
				labProdEqOpLabels(x).Width = VB6.TwipsToPixelsX(TempWide * (1545 / w))
			ElseIf x = 6 Then 
				labProdEqOpLabels(x).Left = VB6.TwipsToPixelsX(TempWide * (5220 / w))
				labProdEqOpLabels(x).Width = VB6.TwipsToPixelsX(TempWide * (1425 / w))
			Else
				labProdEqOpLabels(x).Left = VB6.TwipsToPixelsX(TempWide * (6960 / w))
				labProdEqOpLabels(x).Width = VB6.TwipsToPixelsX(TempWide * (2055 / w))
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
		
		For x = 0 To 1
			LabProdEqOpTitles(x).Top = VB6.TwipsToPixelsY((TempHigh * (780 / h)) + (x * y))
			LabProdEqOpTitles(x).Left = VB6.TwipsToPixelsX(TempWide * (3240 / w))
			LabProdEqOpTitles(x).Width = VB6.TwipsToPixelsX(TempWide * (1575 / w))
			txtProdEqOpValues(x).Top = VB6.TwipsToPixelsY((TempHigh * (750 / h)) + (x * y))
			txtProdEqOpValues(x).Left = VB6.TwipsToPixelsX(TempWide * (5040 / w))
			txtProdEqOpValues(x).Width = VB6.TwipsToPixelsX(TempWide * (855 / w))
			labProdEqOpUnits(x).Top = VB6.TwipsToPixelsY((TempHigh * (780 / h)) + (x * y))
			labProdEqOpUnits(x).Left = VB6.TwipsToPixelsX(TempWide * (5940 / w))
			txtProdEqOpValues(x + 8).Top = VB6.TwipsToPixelsY((TempHigh * (750 / h)) + (x * y))
			txtProdEqOpValues(x + 8).Left = VB6.TwipsToPixelsX(TempWide * (7020 / w))
			txtProdEqOpValues(x + 8).Width = VB6.TwipsToPixelsX(TempWide * (855 / w))
			labProdEqOpUnits(x + 8).Top = VB6.TwipsToPixelsY((TempHigh * (780 / h)) + (x * y))
			labProdEqOpUnits(x + 8).Left = VB6.TwipsToPixelsX(TempWide * (7920 / w))
			
			LabProdEqOpTitles(x + 2).Top = VB6.TwipsToPixelsY((TempHigh * (3480 / h)) + (x * y))
			LabProdEqOpTitles(x + 2).Left = VB6.TwipsToPixelsX(TempWide * (3240 / w))
			LabProdEqOpTitles(x + 2).Width = VB6.TwipsToPixelsX(TempWide * (1575 / w))
			txtProdEqOpValues(x + 2).Top = VB6.TwipsToPixelsY((TempHigh * (3450 / h)) + (x * y))
			txtProdEqOpValues(x + 2).Left = VB6.TwipsToPixelsX(TempWide * (5040 / w))
			txtProdEqOpValues(x + 2).Width = VB6.TwipsToPixelsX(TempWide * (855 / w))
			labProdEqOpUnits(x + 2).Top = VB6.TwipsToPixelsY((TempHigh * (3480 / h)) + (x * y))
			labProdEqOpUnits(x + 2).Left = VB6.TwipsToPixelsX(TempWide * (5940 / w))
			txtProdEqOpValues(x + 10).Top = VB6.TwipsToPixelsY((TempHigh * (3450 / h)) + (x * y))
			txtProdEqOpValues(x + 10).Left = VB6.TwipsToPixelsX(TempWide * (7020 / w))
			txtProdEqOpValues(x + 10).Width = VB6.TwipsToPixelsX(TempWide * (855 / w))
			labProdEqOpUnits(x + 10).Top = VB6.TwipsToPixelsY((TempHigh * (3480 / h)) + (x * y))
			labProdEqOpUnits(x + 10).Left = VB6.TwipsToPixelsX(TempWide * (7920 / w))
			
			LabProdEqOpTitles(x + 4).Top = VB6.TwipsToPixelsY((TempHigh * (2880 / h)) - (x * y))
			LabProdEqOpTitles(x + 4).Left = VB6.TwipsToPixelsX(TempWide * (3240 / w))
			LabProdEqOpTitles(x + 4).Width = VB6.TwipsToPixelsX(TempWide * (1575 / w))
			txtProdEqOpValues(x + 4).Top = VB6.TwipsToPixelsY((TempHigh * (2850 / h)) - (x * y))
			txtProdEqOpValues(x + 4).Left = VB6.TwipsToPixelsX(TempWide * (5040 / w))
			txtProdEqOpValues(x + 4).Width = VB6.TwipsToPixelsX(TempWide * (855 / w))
			labProdEqOpUnits(x + 4).Top = VB6.TwipsToPixelsY((TempHigh * (2880 / h)) - (x * y))
			labProdEqOpUnits(x + 4).Left = VB6.TwipsToPixelsX(TempWide * (5940 / w))
			txtProdEqOpValues(x + 12).Top = VB6.TwipsToPixelsY((TempHigh * (2850 / h)) - (x * y))
			txtProdEqOpValues(x + 12).Left = VB6.TwipsToPixelsX(TempWide * (7020 / w))
			txtProdEqOpValues(x + 12).Width = VB6.TwipsToPixelsX(TempWide * (855 / w))
			labProdEqOpUnits(x + 12).Top = VB6.TwipsToPixelsY((TempHigh * (2880 / h)) - (x * y))
			labProdEqOpUnits(x + 12).Left = VB6.TwipsToPixelsX(TempWide * (7920 / w))
			
			LabProdEqOpTitles(x + 14).Top = VB6.TwipsToPixelsY((TempHigh * (1620 / h)) + (x * y))
			LabProdEqOpTitles(x + 14).Left = VB6.TwipsToPixelsX(TempWide * (3240 / w))
			LabProdEqOpTitles(x + 14).Width = VB6.TwipsToPixelsX(TempWide * (1575 / w))
			txtProdEqOpValues(x + 14).Top = VB6.TwipsToPixelsY((TempHigh * (1590 / h)) + (x * y))
			txtProdEqOpValues(x + 14).Left = VB6.TwipsToPixelsX(TempWide * (5040 / w))
			txtProdEqOpValues(x + 14).Width = VB6.TwipsToPixelsX(TempWide * (855 / w))
			labProdEqOpUnits(x + 14).Top = VB6.TwipsToPixelsY((TempHigh * (1620 / h)) + (x * y))
			labProdEqOpUnits(x + 14).Left = VB6.TwipsToPixelsX(TempWide * (5940 / w))
			txtProdEqOpValues(x + 22).Top = VB6.TwipsToPixelsY((TempHigh * (1590 / h)) + (x * y))
			txtProdEqOpValues(x + 22).Left = VB6.TwipsToPixelsX(TempWide * (7020 / w))
			txtProdEqOpValues(x + 22).Width = VB6.TwipsToPixelsX(TempWide * (855 / w))
			labProdEqOpUnits(x + 22).Top = VB6.TwipsToPixelsY((TempHigh * (1620 / h)) + (x * y))
			labProdEqOpUnits(x + 22).Left = VB6.TwipsToPixelsX(TempWide * (7920 / w))
			
			LabProdEqOpTitles(x + 16).Top = VB6.TwipsToPixelsY((TempHigh * (4320 / h)) + (x * y))
			LabProdEqOpTitles(x + 16).Left = VB6.TwipsToPixelsX(TempWide * (3240 / w))
			LabProdEqOpTitles(x + 16).Width = VB6.TwipsToPixelsX(TempWide * (1575 / w))
			txtProdEqOpValues(x + 16).Top = VB6.TwipsToPixelsY((TempHigh * (4290 / h)) + (x * y))
			txtProdEqOpValues(x + 16).Left = VB6.TwipsToPixelsX(TempWide * (5040 / w))
			txtProdEqOpValues(x + 16).Width = VB6.TwipsToPixelsX(TempWide * (855 / w))
			txtProdEqOpValues(x + 24).Top = VB6.TwipsToPixelsY((TempHigh * (4290 / h)) + (x * y))
			txtProdEqOpValues(x + 24).Left = VB6.TwipsToPixelsX(TempWide * (7020 / w))
			txtProdEqOpValues(x + 24).Width = VB6.TwipsToPixelsX(TempWide * (855 / w))
			labProdEqOpUnits(x + 16).Top = VB6.TwipsToPixelsY((TempHigh * (4320 / h)) + (x * y))
			labProdEqOpUnits(x + 16).Left = VB6.TwipsToPixelsX(TempWide * (5940 / w))
			labProdEqOpUnits(x + 24).Top = VB6.TwipsToPixelsY((TempHigh * (4320 / h)) + (x * y))
			labProdEqOpUnits(x + 24).Left = VB6.TwipsToPixelsX(TempWide * (7920 / w))
			
			LabProdEqOpTitles(x + 18).Top = VB6.TwipsToPixelsY((TempHigh * (5580 / h)) - (x * y))
			LabProdEqOpTitles(x + 18).Left = VB6.TwipsToPixelsX(TempWide * (3240 / w))
			LabProdEqOpTitles(x + 18).Width = VB6.TwipsToPixelsX(TempWide * (1575 / w))
			txtProdEqOpValues(x + 18).Top = VB6.TwipsToPixelsY((TempHigh * (5550 / h)) - (x * y))
			txtProdEqOpValues(x + 18).Left = VB6.TwipsToPixelsX(TempWide * (5040 / w))
			txtProdEqOpValues(x + 18).Width = VB6.TwipsToPixelsX(TempWide * (855 / w))
			txtProdEqOpValues(x + 26).Top = VB6.TwipsToPixelsY((TempHigh * (5550 / h)) - (x * y))
			txtProdEqOpValues(x + 26).Left = VB6.TwipsToPixelsX(TempWide * (7020 / w))
			txtProdEqOpValues(x + 26).Width = VB6.TwipsToPixelsX(TempWide * (855 / w))
			labProdEqOpUnits(x + 18).Top = VB6.TwipsToPixelsY((TempHigh * (5580 / h)) - (x * y))
			labProdEqOpUnits(x + 18).Left = VB6.TwipsToPixelsX(TempWide * (5940 / w))
			labProdEqOpUnits(x + 26).Top = VB6.TwipsToPixelsY((TempHigh * (5580 / h)) - (x * y))
			labProdEqOpUnits(x + 26).Left = VB6.TwipsToPixelsX(TempWide * (7920 / w))
		Next x
		
		For x = 0 To 6
			chkProdEqOpItem(x).Top = VB6.TwipsToPixelsY((TempHigh * (3840 / h)) + (x * u))
			chkProdEqOpItem(x).Left = VB6.TwipsToPixelsX(TempWide * (240 / w))
			chkProdEqOpItem(x).Width = VB6.TwipsToPixelsX(TempWide * (2715 / w))
		Next x
		
		LineHorizontal(0).X1 = VB6.TwipsToPixelsX(TempWide * (1500 / w))
		LineHorizontal(0).X2 = VB6.TwipsToPixelsX(TempWide * (3070 / w))
		LineHorizontal(0).Y1 = VB6.TwipsToPixelsY(TempHigh * (3600 / h))
		LineHorizontal(0).Y2 = VB6.TwipsToPixelsY(TempHigh * (3600 / h))
		
		LineHorizontal(1).X1 = VB6.TwipsToPixelsX(TempWide * (60 / w))
		LineHorizontal(1).X2 = VB6.TwipsToPixelsX(TempWide * (3060 / w))
		LineHorizontal(1).Y1 = VB6.TwipsToPixelsY(TempHigh * (6000 / h))
		LineHorizontal(1).Y2 = VB6.TwipsToPixelsY(TempHigh * (6000 / h))
		
		LineHorizontal(2).X1 = VB6.TwipsToPixelsX(TempWide * (3120 / w))
		LineHorizontal(2).X2 = VB6.TwipsToPixelsX(TempWide * (9120 / w))
		LineHorizontal(2).Y1 = VB6.TwipsToPixelsY(TempHigh * (120 / h))
		LineHorizontal(2).Y2 = VB6.TwipsToPixelsY(TempHigh * (120 / h))
		
		LineHorizontal(3).X1 = VB6.TwipsToPixelsX(TempWide * (3540 / w))
		LineHorizontal(3).X2 = VB6.TwipsToPixelsX(TempWide * (9000 / w))
		LineHorizontal(3).Y1 = VB6.TwipsToPixelsY(TempHigh * (540 / h))
		LineHorizontal(3).Y2 = VB6.TwipsToPixelsY(TempHigh * (540 / h))
		
		LineHorizontal(4).X1 = VB6.TwipsToPixelsX(TempWide * (3780 / w))
		LineHorizontal(4).X2 = VB6.TwipsToPixelsX(TempWide * (4860 / w))
		LineHorizontal(4).Y1 = VB6.TwipsToPixelsY(TempHigh * (3240 / h))
		LineHorizontal(4).Y2 = VB6.TwipsToPixelsY(TempHigh * (3240 / h))
		
		LineHorizontal(5).X1 = VB6.TwipsToPixelsX(TempWide * (4980 / w))
		LineHorizontal(5).X2 = VB6.TwipsToPixelsX(TempWide * (6840 / w))
		LineHorizontal(5).Y1 = VB6.TwipsToPixelsY(TempHigh * (3240 / h))
		LineHorizontal(5).Y2 = VB6.TwipsToPixelsY(TempHigh * (3240 / h))
		
		LineHorizontal(6).X1 = VB6.TwipsToPixelsX(TempWide * (6960 / w))
		LineHorizontal(6).X2 = VB6.TwipsToPixelsX(TempWide * (9000 / w))
		LineHorizontal(6).Y1 = VB6.TwipsToPixelsY(TempHigh * (3240 / h))
		LineHorizontal(6).Y2 = VB6.TwipsToPixelsY(TempHigh * (3240 / h))
		
		LineHorizontal(7).X1 = VB6.TwipsToPixelsX(TempWide * (3120 / w))
		LineHorizontal(7).X2 = VB6.TwipsToPixelsX(TempWide * (9120 / w))
		LineHorizontal(7).Y1 = VB6.TwipsToPixelsY(TempHigh * (5940 / h))
		LineHorizontal(7).Y2 = VB6.TwipsToPixelsY(TempHigh * (5940 / h))
		
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
		LineVertical(2).Y1 = VB6.TwipsToPixelsY(TempHigh * (60 / h))
		LineVertical(2).Y2 = VB6.TwipsToPixelsY(TempHigh * (420 / h))
		
		LineVertical(3).X1 = VB6.TwipsToPixelsX(TempWide * (3180 / w))
		LineVertical(3).X2 = VB6.TwipsToPixelsX(TempWide * (3180 / w))
		LineVertical(3).Y1 = VB6.TwipsToPixelsY(TempHigh * (780 / h))
		LineVertical(3).Y2 = VB6.TwipsToPixelsY(TempHigh * (3120 / h))
		
		LineVertical(4).X1 = VB6.TwipsToPixelsX(TempWide * (3180 / w))
		LineVertical(4).X2 = VB6.TwipsToPixelsX(TempWide * (3180 / w))
		LineVertical(4).Y1 = VB6.TwipsToPixelsY(TempHigh * (3480 / h))
		LineVertical(4).Y2 = VB6.TwipsToPixelsY(TempHigh * (6000 / h))
		
		LineVertical(5).X1 = VB6.TwipsToPixelsX(TempWide * (4920 / w))
		LineVertical(5).X2 = VB6.TwipsToPixelsX(TempWide * (4920 / w))
		LineVertical(5).Y1 = VB6.TwipsToPixelsY(TempHigh * (240 / h))
		LineVertical(5).Y2 = VB6.TwipsToPixelsY(TempHigh * (540 / h))
		
		LineVertical(6).X1 = VB6.TwipsToPixelsX(TempWide * (4920 / w))
		LineVertical(6).X2 = VB6.TwipsToPixelsX(TempWide * (4920 / w))
		LineVertical(6).Y1 = VB6.TwipsToPixelsY(TempHigh * (600 / h))
		LineVertical(6).Y2 = VB6.TwipsToPixelsY(TempHigh * (3180 / h))
		
		LineVertical(7).X1 = VB6.TwipsToPixelsX(TempWide * (4920 / w))
		LineVertical(7).X2 = VB6.TwipsToPixelsX(TempWide * (4920 / w))
		LineVertical(7).Y1 = VB6.TwipsToPixelsY(TempHigh * (3300 / h))
		LineVertical(7).Y2 = VB6.TwipsToPixelsY(TempHigh * (5880 / h))
		
		LineVertical(8).X1 = VB6.TwipsToPixelsX(TempWide * (6900 / w))
		LineVertical(8).X2 = VB6.TwipsToPixelsX(TempWide * (6900 / w))
		LineVertical(8).Y1 = VB6.TwipsToPixelsY(TempHigh * (180 / h))
		LineVertical(8).Y2 = VB6.TwipsToPixelsY(TempHigh * (480 / h))
		
		LineVertical(9).X1 = VB6.TwipsToPixelsX(TempWide * (6900 / w))
		LineVertical(9).X2 = VB6.TwipsToPixelsX(TempWide * (6900 / w))
		LineVertical(9).Y1 = VB6.TwipsToPixelsY(TempHigh * (600 / h))
		LineVertical(9).Y2 = VB6.TwipsToPixelsY(TempHigh * (3180 / h))
		
		LineVertical(10).X1 = VB6.TwipsToPixelsX(TempWide * (6900 / w))
		LineVertical(10).X2 = VB6.TwipsToPixelsX(TempWide * (6900 / w))
		LineVertical(10).Y1 = VB6.TwipsToPixelsY(TempHigh * (3300 / h))
		LineVertical(10).Y2 = VB6.TwipsToPixelsY(TempHigh * (5880 / h))
		
		LineVertical(11).X1 = VB6.TwipsToPixelsX(TempWide * (9060 / w))
		LineVertical(11).X2 = VB6.TwipsToPixelsX(TempWide * (9060 / w))
		LineVertical(11).Y1 = VB6.TwipsToPixelsY(TempHigh * (60 / h))
		LineVertical(11).Y2 = VB6.TwipsToPixelsY(TempHigh * (6000 / h))
		
		labBackToMenu.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labBackToMenu.Left = VB6.TwipsToPixelsX(TempWide * (600 / w))
		
		imgBackToMenu.Top = VB6.TwipsToPixelsY(TempHigh * (6150 / h))
		imgBackToMenu.Left = VB6.TwipsToPixelsX(TempWide * (60 / w))
		imgBackToMenu.Width = VB6.TwipsToPixelsX(TempWide * (495 / w))
		
		comProdEqOpPrint.Top = VB6.TwipsToPixelsY(TempHigh * (6150 / h))
		comProdEqOpPrint.Left = VB6.TwipsToPixelsX(TempWide * (8880 / w))
		
		labProdEqOpHelp.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labProdEqOpHelp.Left = VB6.TwipsToPixelsX(TempWide * (8340 / w))
		
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
		
		For x = 0 To 1
			txtProdEqOpValues(x + 8).Top = VB6.TwipsToPixelsY((TempHigh * (750 / h)) + (x * y))
			txtProdEqOpValues(x + 8).Width = VB6.TwipsToPixelsX(TempWide * (855 / w))
			labProdEqOpUnits(x + 8).Top = VB6.TwipsToPixelsY((TempHigh * (780 / h)) + (x * y))
			txtProdEqOpValues(x + 10).Top = VB6.TwipsToPixelsY((TempHigh * (3450 / h)) + (x * y))
			txtProdEqOpValues(x + 10).Width = VB6.TwipsToPixelsX(TempWide * (855 / w))
			labProdEqOpUnits(x + 10).Top = VB6.TwipsToPixelsY((TempHigh * (3480 / h)) + (x * y))
			txtProdEqOpValues(x + 12).Top = VB6.TwipsToPixelsY((TempHigh * (2850 / h)) - (x * y))
			txtProdEqOpValues(x + 12).Width = VB6.TwipsToPixelsX(TempWide * (855 / w))
			labProdEqOpUnits(x + 12).Top = VB6.TwipsToPixelsY((TempHigh * (2880 / h)) - (x * y))
			txtProdEqOpValues(x + 22).Top = VB6.TwipsToPixelsY((TempHigh * (1590 / h)) + (x * y))
			txtProdEqOpValues(x + 22).Width = VB6.TwipsToPixelsX(TempWide * (855 / w))
			labProdEqOpUnits(x + 22).Top = VB6.TwipsToPixelsY((TempHigh * (1620 / h)) + (x * y))
			txtProdEqOpValues(x + 24).Top = VB6.TwipsToPixelsY((TempHigh * (4290 / h)) + (x * y))
			txtProdEqOpValues(x + 24).Width = VB6.TwipsToPixelsX(TempWide * (855 / w))
			labProdEqOpUnits(x + 24).Top = VB6.TwipsToPixelsY((TempHigh * (4320 / h)) + (x * y))
			txtProdEqOpValues(x + 26).Top = VB6.TwipsToPixelsY((TempHigh * (5550 / h)) - (x * y))
			txtProdEqOpValues(x + 26).Width = VB6.TwipsToPixelsX(TempWide * (855 / w))
			labProdEqOpUnits(x + 26).Top = VB6.TwipsToPixelsY((TempHigh * (5580 / h)) - (x * y))
		Next x
		
		For x = 0 To 5
			Select Case WhichScreen
				Case Lubricants, RepairParts, Undercarriage
					txtProdEqOpValues(x + 8).Left = VB6.TwipsToPixelsX(TempWide * (7380 / w))
					labProdEqOpUnits(x + 8).Left = VB6.TwipsToPixelsX(TempWide * (8280 / w))
					txtProdEqOpValues(x + 22).Left = VB6.TwipsToPixelsX(TempWide * (7380 / w))
					labProdEqOpUnits(x + 22).Left = VB6.TwipsToPixelsX(TempWide * (8280 / w))
				Case Diesel
					If UnitType = English Then
						txtProdEqOpValues(x + 8).Left = VB6.TwipsToPixelsX(TempWide * (7020 / w))
						labProdEqOpUnits(x + 8).Left = VB6.TwipsToPixelsX(TempWide * (7920 / w))
						txtProdEqOpValues(x + 22).Left = VB6.TwipsToPixelsX(TempWide * (7020 / w))
						labProdEqOpUnits(x + 22).Left = VB6.TwipsToPixelsX(TempWide * (7920 / w))
					Else
						txtProdEqOpValues(x + 8).Left = VB6.TwipsToPixelsX(TempWide * (7140 / w))
						labProdEqOpUnits(x + 8).Left = VB6.TwipsToPixelsX(TempWide * (8040 / w))
						txtProdEqOpValues(x + 22).Left = VB6.TwipsToPixelsX(TempWide * (7140 / w))
						labProdEqOpUnits(x + 22).Left = VB6.TwipsToPixelsX(TempWide * (8040 / w))
					End If
				Case Electricity
					txtProdEqOpValues(x + 8).Left = VB6.TwipsToPixelsX(TempWide * (7140 / w))
					labProdEqOpUnits(x + 8).Left = VB6.TwipsToPixelsX(TempWide * (8040 / w))
					txtProdEqOpValues(x + 22).Left = VB6.TwipsToPixelsX(TempWide * (7140 / w))
					labProdEqOpUnits(x + 22).Left = VB6.TwipsToPixelsX(TempWide * (8040 / w))
				Case RepairLabor
					txtProdEqOpValues(x + 8).Left = VB6.TwipsToPixelsX(TempWide * (7080 / w))
					labProdEqOpUnits(x + 8).Left = VB6.TwipsToPixelsX(TempWide * (7980 / w))
					txtProdEqOpValues(x + 22).Left = VB6.TwipsToPixelsX(TempWide * (7080 / w))
					labProdEqOpUnits(x + 22).Left = VB6.TwipsToPixelsX(TempWide * (7980 / w))
				Case Else
					txtProdEqOpValues(x + 8).Left = VB6.TwipsToPixelsX(TempWide * (7440 / w))
					labProdEqOpUnits(x + 8).Left = VB6.TwipsToPixelsX(TempWide * (8340 / w))
					txtProdEqOpValues(x + 22).Left = VB6.TwipsToPixelsX(TempWide * (7440 / w))
					labProdEqOpUnits(x + 22).Left = VB6.TwipsToPixelsX(TempWide * (8340 / w))
			End Select
		Next x
		
	End Sub
	Public Sub drawthevalues()
		
		Dim i As Short
		Dim x As Short
        'Dim p As Short
        'Dim r As Short
        'Dim tempvalue As Decimal

        DoNotChange = True
		
		For i = 0 To 19
			Select Case i
				Case 0 To 5, 14 To 19
					LabProdEqOpTitles(i).Enabled = True
					txtProdEqOpValues(i).Enabled = True
					txtProdEqOpValues(i + 8).Enabled = True
					txtProdEqOpValues(i + 8).Text = " "
					labProdEqOpUnits(i).Enabled = True
					labProdEqOpUnits(i + 8).Enabled = True
					If CellValues(EquipmentHours, i, WhichSegment).Changed = True Then
						txtProdEqOpValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF0000)
					Else
						txtProdEqOpValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
					End If
					If CellValues(WhichScreen, i, WhichSegment).Changed = True Then
						txtProdEqOpValues(i + 8).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF0000)
					Else
						txtProdEqOpValues(i + 8).ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
					End If
			End Select
			Select Case i
				Case 0
					For x = 0 To 3
						If CellValues(EquipmentOne, x, WhichSegment).Value <> 0 Then
							Select Case x
								Case 0
									LabProdEqOpTitles(i).Text = "Front-End Loader"
								Case 1
									LabProdEqOpTitles(i).Text = "Hydraulic Shovel"
								Case 2
									LabProdEqOpTitles(i).Text = "Mechanical Shovel"
								Case 3
									LabProdEqOpTitles(i).Text = "Walking Dragline"
							End Select
						End If
					Next x
					If CellValues(EquipmentOne, 20, WhichSegment).Value <> 0 Then
						LabProdEqOpTitles(i).Text = "Scraper"
						LabProdEqOpTitles(i + 1).Enabled = False
						txtProdEqOpValues(i + 1).Enabled = False
						txtProdEqOpValues(i + 9).Enabled = False
						labProdEqOpUnits(i + 1).Enabled = False
						labProdEqOpUnits(i + 9).Enabled = False
					End If
				Case 1
					If CellValues(EquipmentOne, 4, WhichSegment).Value <> 0 Then
						LabProdEqOpTitles(i).Text = "Rear-Dump Truck"
					ElseIf CellValues(EquipmentOne, 21, WhichSegment).Value <> 0 Then 
						LabProdEqOpTitles(i).Text = "Articulated Hauler"
					End If
				Case 2
					For x = 5 To 8
						If CellValues(EquipmentOne, x, WhichSegment).Value <> 0 Then
							Select Case x
								Case 5
									LabProdEqOpTitles(i).Text = "Front-End Loader"
								Case 6
									LabProdEqOpTitles(i).Text = "Hydraulic Shovel"
								Case 7
									LabProdEqOpTitles(i).Text = "Mechanical Shovel"
								Case 8
									LabProdEqOpTitles(i).Text = "Walking Dragline"
							End Select
						End If
					Next x
					If CellValues(EquipmentOne, 25, WhichSegment).Value <> 0 Then
						LabProdEqOpTitles(i).Text = "Scraper"
						LabProdEqOpTitles(i + 1).Enabled = False
						txtProdEqOpValues(i + 1).Enabled = False
						txtProdEqOpValues(i + 9).Enabled = False
						labProdEqOpUnits(i + 1).Enabled = False
						labProdEqOpUnits(i + 9).Enabled = False
					End If
				Case 3
					If CellValues(EquipmentOne, 9, WhichSegment).Value <> 0 Then
						LabProdEqOpTitles(i).Text = "Rear-Dump Truck"
					ElseIf CellValues(EquipmentOne, 26, WhichSegment).Value <> 0 Then 
						LabProdEqOpTitles(i).Text = "Articulated Hauler"
					End If
				Case 4
					If CellValues(EquipmentOne, 10, WhichSegment).Value = 0 Then
						LabProdEqOpTitles(i).Enabled = False
						txtProdEqOpValues(i).Enabled = False
						txtProdEqOpValues(i + 8).Enabled = False
						labProdEqOpUnits(i).Enabled = False
						labProdEqOpUnits(i + 8).Enabled = False
					End If
				Case 5
					If CellValues(EquipmentOne, 11, WhichSegment).Value = 0 Then
						LabProdEqOpTitles(i).Enabled = False
						txtProdEqOpValues(i).Enabled = False
						txtProdEqOpValues(i + 8).Enabled = False
						labProdEqOpUnits(i).Enabled = False
						labProdEqOpUnits(i + 8).Enabled = False
					End If
				Case 14
					If CellValues(EquipmentOne, 22, WhichSegment).Value <> 0 Then
						LabProdEqOpTitles(i).Text = "Jaw Crusher"
						LabProdEqOpTitles(i - 13).Enabled = False
						txtProdEqOpValues(i - 13).Enabled = False
						txtProdEqOpValues(i - 5).Enabled = False
						labProdEqOpUnits(i - 13).Enabled = False
						labProdEqOpUnits(i + 5).Enabled = False
					ElseIf CellValues(EquipmentOne, 23, WhichSegment).Value <> 0 Then 
						LabProdEqOpTitles(i).Text = "Gyratory Crusher"
						LabProdEqOpTitles(i).Text = "Jaw Crusher"
						LabProdEqOpTitles(i - 13).Enabled = False
						txtProdEqOpValues(i - 13).Enabled = False
						txtProdEqOpValues(i - 5).Enabled = False
						labProdEqOpUnits(i - 13).Enabled = False
						labProdEqOpUnits(i + 5).Enabled = False
					Else
						LabProdEqOpTitles(i).Enabled = False
						txtProdEqOpValues(i).Enabled = False
						txtProdEqOpValues(i + 8).Enabled = False
						labProdEqOpUnits(i).Enabled = False
						labProdEqOpUnits(i + 8).Enabled = False
					End If
				Case 15
					If CellValues(EquipmentOne, 24, WhichSegment).Value = 0 Then
						LabProdEqOpTitles(i).Enabled = False
						txtProdEqOpValues(i).Enabled = False
						txtProdEqOpValues(i + 8).Enabled = False
						labProdEqOpUnits(i).Enabled = False
						labProdEqOpUnits(i + 8).Enabled = False
					End If
				Case 16
					If CellValues(EquipmentOne, 27, WhichSegment).Value <> 0 Then
						LabProdEqOpTitles(i).Text = "Jaw Crusher"
						LabProdEqOpTitles(i).Text = "Jaw Crusher"
						LabProdEqOpTitles(i - 13).Enabled = False
						txtProdEqOpValues(i - 13).Enabled = False
						txtProdEqOpValues(i - 5).Enabled = False
						labProdEqOpUnits(i - 13).Enabled = False
						labProdEqOpUnits(i - 5).Enabled = False
					ElseIf CellValues(EquipmentOne, 28, WhichSegment).Value <> 0 Then 
						LabProdEqOpTitles(i).Text = "Gyratory Crusher"
						LabProdEqOpTitles(i).Text = "Jaw Crusher"
						LabProdEqOpTitles(i - 13).Enabled = False
						txtProdEqOpValues(i - 13).Enabled = False
						txtProdEqOpValues(i - 5).Enabled = False
						labProdEqOpUnits(i - 13).Enabled = False
						labProdEqOpUnits(i - 5).Enabled = False
					Else
						LabProdEqOpTitles(i).Enabled = False
						txtProdEqOpValues(i).Enabled = False
						txtProdEqOpValues(i + 8).Enabled = False
						labProdEqOpUnits(i).Enabled = False
						labProdEqOpUnits(i + 8).Enabled = False
					End If
				Case 17
					If CellValues(EquipmentOne, 29, WhichSegment).Value = 0 Then
						LabProdEqOpTitles(i).Enabled = False
						txtProdEqOpValues(i).Enabled = False
						txtProdEqOpValues(i + 8).Enabled = False
						labProdEqOpUnits(i).Enabled = False
						labProdEqOpUnits(i + 8).Enabled = False
					End If
				Case 18
					If CellValues(EquipmentOne, 30, WhichSegment).Value = 0 Then
						LabProdEqOpTitles(i).Enabled = False
						txtProdEqOpValues(i).Enabled = False
						txtProdEqOpValues(i + 8).Enabled = False
						labProdEqOpUnits(i).Enabled = False
						labProdEqOpUnits(i + 8).Enabled = False
					End If
				Case 19
					If CellValues(EquipmentOne, 31, WhichSegment).Value = 0 Then
						LabProdEqOpTitles(i).Enabled = False
						txtProdEqOpValues(i).Enabled = False
						txtProdEqOpValues(i + 8).Enabled = False
						labProdEqOpUnits(i).Enabled = False
						labProdEqOpUnits(i + 8).Enabled = False
					End If
			End Select
		Next i
		
		For i = 0 To 5
			If CellValues(EquipmentHours, i, WhichSegment).Value <> 0 Then
				txtProdEqOpValues(i).Text = VB6.Format(LTrim(Str(CellValues(EquipmentHours, i, WhichSegment).Value)), "###,##0.00")
			End If
		Next i
		
		For i = 14 To 19
			If CellValues(EquipmentHours, i, WhichSegment).Value <> 0 Then
				txtProdEqOpValues(i).Text = VB6.Format(LTrim(Str(CellValues(EquipmentHours, i, WhichSegment).Value)), "###,##0.00")
			End If
		Next i
		
		If UnitType = Metric Then
			Select Case WhichScreen
				Case Diesel
					For i = 0 To 5
						labProdEqOpUnits(i + 8).Text = "liters/hour"
						labProdEqOpUnits(i + 22).Text = "liters/hour"
						If CellValues(WhichScreen, i, WhichSegment).Value <> 0 Then
							txtProdEqOpValues(i + 8).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value * GallonConv)), "###,##0.00")
						End If
						If CellValues(WhichScreen, i + 14, WhichSegment).Value <> 0 Then
							txtProdEqOpValues(i + 22).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i + 14, WhichSegment).Value * GallonConv)), "###,##0.00")
						End If
					Next i
				Case Electricity
					For i = 0 To 5
						labProdEqOpUnits(i + 8).Text = "kWh/hour"
						labProdEqOpUnits(i + 22).Text = "kWh/hour"
						If CellValues(WhichScreen, i, WhichSegment).Value <> 0 Then
							txtProdEqOpValues(i + 8).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value)), "###,###,##0")
						End If
						If CellValues(WhichScreen, i + 14, WhichSegment).Value <> 0 Then
							txtProdEqOpValues(i + 22).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i + 14, WhichSegment).Value)), "###,###,##0")
						End If
					Next i
				Case Lubricants, RepairParts, Undercarriage
					For i = 0 To 5
						labProdEqOpUnits(i + 8).Text = "/hour"
						labProdEqOpUnits(i + 22).Text = "/hour"
						If CellValues(WhichScreen, i, WhichSegment).Value <> 0 Then
							txtProdEqOpValues(i + 8).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value)), "$###,##0.00")
						End If
						If CellValues(WhichScreen, i + 14, WhichSegment).Value <> 0 Then
							txtProdEqOpValues(i + 22).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i + 14, WhichSegment).Value)), "$###,##0.00")
						End If
					Next i
				Case RepairLabor
					For i = 0 To 5
						labProdEqOpUnits(i + 8).Text = "hours/hour"
						labProdEqOpUnits(i + 22).Text = "hours/hour"
						If CellValues(WhichScreen, i, WhichSegment).Value <> 0 Then
							txtProdEqOpValues(i + 8).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value)), "#,##0.000")
						End If
						If CellValues(WhichScreen, i + 14, WhichSegment).Value <> 0 Then
							txtProdEqOpValues(i + 22).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i + 14, WhichSegment).Value)), "#,##0.000")
						End If
					Next i
				Case Tires
					For i = 0 To 5
						labProdEqOpUnits(i + 8).Text = "/set"
						labProdEqOpUnits(i + 22).Text = "/set"
						If CellValues(WhichScreen, i, WhichSegment).Value <> 0 Then
							txtProdEqOpValues(i + 8).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value)), "$#,###,##0")
						End If
						If CellValues(WhichScreen, i + 14, WhichSegment).Value <> 0 Then
							txtProdEqOpValues(i + 22).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i + 14, WhichSegment).Value)), "$###,##0.00")
						End If
					Next i
			End Select
		Else
			Select Case WhichScreen
				Case Diesel
					For i = 0 To 5
						labProdEqOpUnits(i + 8).Text = "gallons/hour"
						labProdEqOpUnits(i + 22).Text = "gallons/hour"
						If CellValues(WhichScreen, i, WhichSegment).Value <> 0 Then
							txtProdEqOpValues(i + 8).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value * GallonConv)), "###,##0.00")
						End If
						If CellValues(WhichScreen, i + 14, WhichSegment).Value <> 0 Then
							txtProdEqOpValues(i + 22).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i + 14, WhichSegment).Value * GallonConv)), "###,##0.00")
						End If
					Next i
				Case Electricity
					For i = 0 To 5
						labProdEqOpUnits(i + 8).Text = "kWh/hour"
						labProdEqOpUnits(i + 22).Text = "kWh/hour"
						If CellValues(WhichScreen, i, WhichSegment).Value <> 0 Then
							txtProdEqOpValues(i + 8).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value)), "###,###,##0")
						End If
						If CellValues(WhichScreen, i + 14, WhichSegment).Value <> 0 Then
							txtProdEqOpValues(i + 22).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i + 14, WhichSegment).Value)), "###,###,##0")
						End If
					Next i
				Case Lubricants, RepairParts, Undercarriage
					For i = 0 To 5
						labProdEqOpUnits(i + 8).Text = "/hour"
						labProdEqOpUnits(i + 22).Text = "/hour"
						If CellValues(WhichScreen, i, WhichSegment).Value <> 0 Then
							txtProdEqOpValues(i + 8).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value)), "$###,##0.00")
						End If
						If CellValues(WhichScreen, i + 14, WhichSegment).Value <> 0 Then
							txtProdEqOpValues(i + 22).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i + 14, WhichSegment).Value)), "$###,##0.00")
						End If
					Next i
				Case RepairLabor
					For i = 0 To 5
						labProdEqOpUnits(i + 8).Text = "hours/hour"
						labProdEqOpUnits(i + 22).Text = "hours/hour"
						If CellValues(WhichScreen, i, WhichSegment).Value <> 0 Then
							txtProdEqOpValues(i + 8).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value)), "#,##0.000")
						End If
						If CellValues(WhichScreen, i + 14, WhichSegment).Value <> 0 Then
							txtProdEqOpValues(i + 22).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i + 14, WhichSegment).Value)), "#,##0.000")
						End If
					Next i
				Case Tires
					For i = 0 To 5
						labProdEqOpUnits(i + 8).Text = "/set"
						labProdEqOpUnits(i + 22).Text = "/set"
						If CellValues(WhichScreen, i, WhichSegment).Value <> 0 Then
							txtProdEqOpValues(i + 8).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value)), "$#,###,##0")
						End If
						If CellValues(WhichScreen, i + 14, WhichSegment).Value <> 0 Then
							txtProdEqOpValues(i + 22).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i + 14, WhichSegment).Value)), "$###,##0.00")
						End If
					Next i
			End Select
		End If
		
		Select Case WhichScreen
			Case Diesel
				labProdEqOpLabels(7).Text = "Diesel Consumption"
			Case Electricity
				labProdEqOpLabels(7).Text = "Electrical Consumption"
			Case Lubricants
				labProdEqOpLabels(7).Text = "Lubricant Costs"
			Case RepairParts
				labProdEqOpLabels(7).Text = "Repair Parts Costs"
			Case Undercarriage
				labProdEqOpLabels(7).Text = "Undercarriage Costs"
			Case RepairLabor
				labProdEqOpLabels(7).Text = "Repair Labor"
			Case Tires
				labProdEqOpLabels(7).Text = "Tire Prices"
		End Select
		
		For i = 0 To 1
			labProdEqOpUnits(i).Text = "hours/day"
		Next i
		
		DoNotChange = False
		
	End Sub
	Private Sub txtProdEqOpValues_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtProdEqOpValues.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Dim Index As Short = txtProdEqOpValues.GetIndex(eventSender)
		
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
						If InStr(txtProdEqOpValues(Index).Text, ".") = 0 Then
							System.Windows.Forms.SendKeys.Send("{DELETE}")
						End If
					Else
						System.Windows.Forms.SendKeys.Send("{DELETE}")
					End If
			End Select
		End If
		
	End Sub
	Private Sub txtProdEqOpValues_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtProdEqOpValues.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = txtProdEqOpValues.GetIndex(eventSender)
		If KeyAscii > Asc("9") And KeyAscii <> Asc(",") And KeyAscii <> Asc(".") And KeyAscii <> Asc("$") Then
			Beep()
			KeyAscii = 0
		Else
			Select Case WhichCell
				Case 0 To 5, 14 To 19
					CellValues(EquipmentHours, WhichCell, WhichSegment).Changed = True
				Case 8 To 13, 22 To 27
					CellValues(WhichScreen, WhichCell - 8, WhichSegment).Changed = True
			End Select
		End If
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	Private Sub txtProdEqOpValues_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtProdEqOpValues.Leave
		Dim Index As Short = txtProdEqOpValues.GetIndex(eventSender)
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
		For i = 1 To Len(txtProdEqOpValues(Sample).Text)
			Digit.Value = Mid(txtProdEqOpValues(Sample).Text, i, 1)
			Select Case Digit.Value
				Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ".", "-"
					tempvalue = tempvalue & Digit.Value
			End Select
		Next i
		Select Case Sample
			Case 0 To 5, 14 To 19
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
			chkProdEqOpItem(TheScreen - Diesel).CheckState = System.Windows.Forms.CheckState.Unchecked
			chkProdEqOpItem(TheScreen - Diesel).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFFFF)
		Next TheScreen
		
		chkProdEqOpItem(WhichScreen - Diesel).CheckState = System.Windows.Forms.CheckState.Checked
		chkProdEqOpItem(WhichScreen - Diesel).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF00)
		
		DoNotChange = False
	End Sub
End Class