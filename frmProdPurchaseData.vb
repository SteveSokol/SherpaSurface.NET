Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks
Friend Class frmProdPurchaseData
	Inherits System.Windows.Forms.Form
	Dim TempHigh As Single
	Dim TempWide As Single
	Dim FootConversion As Single
	Dim DensityConversion As Single
	Dim PowderConversion As Single
	Private Sub comProdPurchasePrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles comProdPurchasePrint.Click
		PrintForm1.Print(Me, PowerPacks.Printing.PrintForm.PrintOption.CompatibleModeClientAreaOnly)
	End Sub
	'UPGRADE_WARNING: Form event frmProdPurchaseData.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmProdPurchaseData_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated

        'Dim baseunit As String
        'Dim baselength As Short
        Dim i As Short
		
		If IsHelpOn = True Then
			txtProdPurchaseValues(WhichCell).Focus()
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
			
			txtProdPurchaseValues(0).Focus()
		End If
		
	End Sub
	Private Sub frmProdPurchaseData_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
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
	'UPGRADE_WARNING: Event frmProdPurchaseData.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub frmProdPurchaseData_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		TempWide = VB6.PixelsToTwipsX(Me.ClientRectangle.Width)
		TempHigh = VB6.PixelsToTwipsY(Me.ClientRectangle.Height)
		Call screenstuff()
	End Sub
	Private Sub frmProdPurchaseData_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
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
	Private Sub labProdPurchaseHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles labProdPurchaseHelp.Click
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
			txtProdPurchaseValues(WhichCell).Focus()
		End If
	End Sub
	'UPGRADE_WARNING: Event txtProdPurchaseValues.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtProdPurchaseValues_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtProdPurchaseValues.TextChanged
		Dim Index As Short = txtProdPurchaseValues.GetIndex(eventSender)
		If DoNotChange = True Then Exit Sub
	End Sub
	Private Sub txtProdPurchaseValues_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtProdPurchaseValues.Enter
		Dim Index As Short = txtProdPurchaseValues.GetIndex(eventSender)
        'Dim x As Short
        WhichCell = Index
		Call drawthevalues()
	End Sub
	Public Sub screenstuff()

        'Dim p As Decimal
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
		
		q = (60 / h) * TempHigh
		r = (120 / w) * TempWide
		s = (2700 / h) * TempHigh
		t = (1560 / h) * TempHigh
		u = (480 / h) * TempHigh
		v = (300 / w) * TempWide
		
		y = (420 / h) * TempHigh
		z = (300 / w) * TempWide
		
		For x = 0 To 2
			labProdPurchaseHeading(x).Top = VB6.TwipsToPixelsY((TempHigh * (60 / h)) + (x * u))
			labProdPurchaseHeading(x).Left = VB6.TwipsToPixelsX((TempWide * (60 / w)) + (x * v))
			labProdPurchaseHeading(x).Width = VB6.TwipsToPixelsX(TempWide * (2175 / w))
		Next x
		
		For x = 0 To 1
			labProdPurchaseLabels(x).Top = VB6.TwipsToPixelsY((TempHigh * (2280 / h)) + (x * t))
			labProdPurchaseLabels(x).Left = VB6.TwipsToPixelsX(TempWide * (540 / w))
			labProdPurchaseLabels(x).Width = VB6.TwipsToPixelsX(TempWide * (1935 / w))
			labProdPurchaseLabels(x + 2).Top = VB6.TwipsToPixelsY((TempHigh * (480 / h)) + (x * s))
			labProdPurchaseLabels(x + 2).Left = VB6.TwipsToPixelsX(TempWide * (3000 / w))
		Next x
		
		For x = 0 To 2
			labProdPurchaseLabels(x + 4).Top = VB6.TwipsToPixelsY(TempHigh * (240 / h))
		Next x
		
		labProdPurchaseLabels(4).Left = VB6.TwipsToPixelsX(TempWide * (3660 / w))
		labProdPurchaseLabels(4).Width = VB6.TwipsToPixelsX(TempWide * (1725 / w))
		labProdPurchaseLabels(5).Left = VB6.TwipsToPixelsX(TempWide * (5580 / w))
		labProdPurchaseLabels(5).Width = VB6.TwipsToPixelsX(TempWide * (1335 / w))
		labProdPurchaseLabels(6).Left = VB6.TwipsToPixelsX(TempWide * (7260 / w))
		labProdPurchaseLabels(6).Width = VB6.TwipsToPixelsX(TempWide * (1575 / w))
		
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
		
		For x = 0 To 1
			LabProdPurchaseTitles(x).Top = VB6.TwipsToPixelsY((TempHigh * (780 / h)) + (x * y))
			LabProdPurchaseTitles(x).Left = VB6.TwipsToPixelsX(TempWide * (3720 / w))
			LabProdPurchaseTitles(x).Width = VB6.TwipsToPixelsX(TempWide * (1635 / w))
			txtProdPurchaseValues(x).Top = VB6.TwipsToPixelsY((TempHigh * (750 / h)) + (x * y))
			txtProdPurchaseValues(x).Left = VB6.TwipsToPixelsX(TempWide * (5580 / w))
			txtProdPurchaseValues(x).Width = VB6.TwipsToPixelsX(TempWide * (1335 / w))
			txtProdPurchaseValues(x + 20).Top = VB6.TwipsToPixelsY((TempHigh * (750 / h)) + (x * y))
			txtProdPurchaseValues(x + 20).Left = VB6.TwipsToPixelsX(TempWide * (7260 / w))
			txtProdPurchaseValues(x + 20).Width = VB6.TwipsToPixelsX(TempWide * (1575 / w))
			
			LabProdPurchaseTitles(x + 2).Top = VB6.TwipsToPixelsY((TempHigh * (3480 / h)) + (x * y))
			LabProdPurchaseTitles(x + 2).Left = VB6.TwipsToPixelsX(TempWide * (3720 / w))
			LabProdPurchaseTitles(x + 2).Width = VB6.TwipsToPixelsX(TempWide * (1635 / w))
			txtProdPurchaseValues(x + 2).Top = VB6.TwipsToPixelsY((TempHigh * (3450 / h)) + (x * y))
			txtProdPurchaseValues(x + 2).Left = VB6.TwipsToPixelsX(TempWide * (5580 / w))
			txtProdPurchaseValues(x + 2).Width = VB6.TwipsToPixelsX(TempWide * (1335 / w))
			txtProdPurchaseValues(x + 22).Top = VB6.TwipsToPixelsY((TempHigh * (3450 / h)) + (x * y))
			txtProdPurchaseValues(x + 22).Left = VB6.TwipsToPixelsX(TempWide * (7260 / w))
			txtProdPurchaseValues(x + 22).Width = VB6.TwipsToPixelsX(TempWide * (1575 / w))
			
			LabProdPurchaseTitles(x + 4).Top = VB6.TwipsToPixelsY((TempHigh * (2880 / h)) - (x * y))
			LabProdPurchaseTitles(x + 4).Left = VB6.TwipsToPixelsX(TempWide * (3720 / w))
			LabProdPurchaseTitles(x + 4).Width = VB6.TwipsToPixelsX(TempWide * (1635 / w))
			txtProdPurchaseValues(x + 4).Top = VB6.TwipsToPixelsY((TempHigh * (2850 / h)) - (x * y))
			txtProdPurchaseValues(x + 4).Left = VB6.TwipsToPixelsX(TempWide * (5580 / w))
			txtProdPurchaseValues(x + 4).Width = VB6.TwipsToPixelsX(TempWide * (1335 / w))
			txtProdPurchaseValues(x + 24).Top = VB6.TwipsToPixelsY((TempHigh * (2850 / h)) - (x * y))
			txtProdPurchaseValues(x + 24).Left = VB6.TwipsToPixelsX(TempWide * (7260 / w))
			txtProdPurchaseValues(x + 24).Width = VB6.TwipsToPixelsX(TempWide * (1575 / w))
			
			LabProdPurchaseTitles(x + 14).Top = VB6.TwipsToPixelsY((TempHigh * (1620 / h)) + (x * y))
			LabProdPurchaseTitles(x + 14).Left = VB6.TwipsToPixelsX(TempWide * (3720 / w))
			LabProdPurchaseTitles(x + 14).Width = VB6.TwipsToPixelsX(TempWide * (1635 / w))
			txtProdPurchaseValues(x + 14).Top = VB6.TwipsToPixelsY((TempHigh * (1590 / h)) + (x * y))
			txtProdPurchaseValues(x + 14).Left = VB6.TwipsToPixelsX(TempWide * (5580 / w))
			txtProdPurchaseValues(x + 14).Width = VB6.TwipsToPixelsX(TempWide * (1335 / w))
			txtProdPurchaseValues(x + 34).Top = VB6.TwipsToPixelsY((TempHigh * (1590 / h)) + (x * y))
			txtProdPurchaseValues(x + 34).Left = VB6.TwipsToPixelsX(TempWide * (7260 / w))
			txtProdPurchaseValues(x + 34).Width = VB6.TwipsToPixelsX(TempWide * (1575 / w))
			
			LabProdPurchaseTitles(x + 16).Top = VB6.TwipsToPixelsY((TempHigh * (4320 / h)) + (x * y))
			LabProdPurchaseTitles(x + 16).Left = VB6.TwipsToPixelsX(TempWide * (3720 / w))
			LabProdPurchaseTitles(x + 16).Width = VB6.TwipsToPixelsX(TempWide * (1635 / w))
			txtProdPurchaseValues(x + 16).Top = VB6.TwipsToPixelsY((TempHigh * (4290 / h)) + (x * y))
			txtProdPurchaseValues(x + 16).Left = VB6.TwipsToPixelsX(TempWide * (5580 / w))
			txtProdPurchaseValues(x + 16).Width = VB6.TwipsToPixelsX(TempWide * (1335 / w))
			txtProdPurchaseValues(x + 36).Top = VB6.TwipsToPixelsY((TempHigh * (4290 / h)) + (x * y))
			txtProdPurchaseValues(x + 36).Left = VB6.TwipsToPixelsX(TempWide * (7260 / w))
			txtProdPurchaseValues(x + 36).Width = VB6.TwipsToPixelsX(TempWide * (1575 / w))
			
			LabProdPurchaseTitles(x + 18).Top = VB6.TwipsToPixelsY((TempHigh * (5580 / h)) - (x * y))
			LabProdPurchaseTitles(x + 18).Left = VB6.TwipsToPixelsX(TempWide * (3720 / w))
			LabProdPurchaseTitles(x + 18).Width = VB6.TwipsToPixelsX(TempWide * (1635 / w))
			txtProdPurchaseValues(x + 18).Top = VB6.TwipsToPixelsY((TempHigh * (5550 / h)) - (x * y))
			txtProdPurchaseValues(x + 18).Left = VB6.TwipsToPixelsX(TempWide * (5580 / w))
			txtProdPurchaseValues(x + 18).Width = VB6.TwipsToPixelsX(TempWide * (1335 / w))
			txtProdPurchaseValues(x + 38).Top = VB6.TwipsToPixelsY((TempHigh * (5550 / h)) - (x * y))
			txtProdPurchaseValues(x + 38).Left = VB6.TwipsToPixelsX(TempWide * (7260 / w))
			txtProdPurchaseValues(x + 38).Width = VB6.TwipsToPixelsX(TempWide * (1575 / w))
		Next x
		
		LineHorizontal(0).X1 = VB6.TwipsToPixelsX(TempWide * (3000 / w))
		LineHorizontal(0).X2 = VB6.TwipsToPixelsX(TempWide * (9060 / w))
		LineHorizontal(0).Y1 = VB6.TwipsToPixelsY(TempHigh * (120 / h))
		LineHorizontal(0).Y2 = VB6.TwipsToPixelsY(TempHigh * (120 / h))
		
		LineHorizontal(1).X1 = VB6.TwipsToPixelsX(TempWide * (3420 / w))
		LineHorizontal(1).X2 = VB6.TwipsToPixelsX(TempWide * (8940 / w))
		LineHorizontal(1).Y1 = VB6.TwipsToPixelsY(TempHigh * (540 / h))
		LineHorizontal(1).Y2 = VB6.TwipsToPixelsY(TempHigh * (540 / h))
		
		LineHorizontal(2).X1 = VB6.TwipsToPixelsX(TempWide * (3720 / w))
		LineHorizontal(2).X2 = VB6.TwipsToPixelsX(TempWide * (7020 / w))
		LineHorizontal(2).Y1 = VB6.TwipsToPixelsY(TempHigh * (3240 / h))
		LineHorizontal(2).Y2 = VB6.TwipsToPixelsY(TempHigh * (3240 / h))
		
		LineHorizontal(3).X1 = VB6.TwipsToPixelsX(TempWide * (7140 / w))
		LineHorizontal(3).X2 = VB6.TwipsToPixelsX(TempWide * (8940 / w))
		LineHorizontal(3).Y1 = VB6.TwipsToPixelsY(TempHigh * (3240 / h))
		LineHorizontal(3).Y2 = VB6.TwipsToPixelsY(TempHigh * (3240 / h))
		
		LineHorizontal(4).X1 = VB6.TwipsToPixelsX(TempWide * (3000 / w))
		LineHorizontal(4).X2 = VB6.TwipsToPixelsX(TempWide * (9060 / w))
		LineHorizontal(4).Y1 = VB6.TwipsToPixelsY(TempHigh * (5940 / h))
		LineHorizontal(4).Y2 = VB6.TwipsToPixelsY(TempHigh * (5940 / h))
		
		LineVertical(0).X1 = VB6.TwipsToPixelsX(TempWide * (3060 / w))
		LineVertical(0).X2 = VB6.TwipsToPixelsX(TempWide * (3060 / w))
		LineVertical(0).Y1 = VB6.TwipsToPixelsY(TempHigh * (60 / h))
		LineVertical(0).Y2 = VB6.TwipsToPixelsY(TempHigh * (420 / h))
		
		LineVertical(1).X1 = VB6.TwipsToPixelsX(TempWide * (3060 / w))
		LineVertical(1).X2 = VB6.TwipsToPixelsX(TempWide * (3060 / w))
		LineVertical(1).Y1 = VB6.TwipsToPixelsY(TempHigh * (780 / h))
		LineVertical(1).Y2 = VB6.TwipsToPixelsY(TempHigh * (3120 / h))
		
		LineVertical(2).X1 = VB6.TwipsToPixelsX(TempWide * (3060 / w))
		LineVertical(2).X2 = VB6.TwipsToPixelsX(TempWide * (3060 / w))
		LineVertical(2).Y1 = VB6.TwipsToPixelsY(TempHigh * (3480 / h))
		LineVertical(2).Y2 = VB6.TwipsToPixelsY(TempHigh * (6000 / h))
		
		LineVertical(3).X1 = VB6.TwipsToPixelsX(TempWide * (7080 / w))
		LineVertical(3).X2 = VB6.TwipsToPixelsX(TempWide * (7080 / w))
		LineVertical(3).Y1 = VB6.TwipsToPixelsY(TempHigh * (600 / h))
		LineVertical(3).Y2 = VB6.TwipsToPixelsY(TempHigh * (3180 / h))
		
		LineVertical(4).X1 = VB6.TwipsToPixelsX(TempWide * (7080 / w))
		LineVertical(4).X2 = VB6.TwipsToPixelsX(TempWide * (7080 / w))
		LineVertical(4).Y1 = VB6.TwipsToPixelsY(TempHigh * (3300 / h))
		LineVertical(4).Y2 = VB6.TwipsToPixelsY(TempHigh * (5880 / h))
		
		LineVertical(5).X1 = VB6.TwipsToPixelsX(TempWide * (9000 / w))
		LineVertical(5).X2 = VB6.TwipsToPixelsX(TempWide * (9000 / w))
		LineVertical(5).Y1 = VB6.TwipsToPixelsY(TempHigh * (60 / h))
		LineVertical(5).Y2 = VB6.TwipsToPixelsY(TempHigh * (6000 / h))
		
		labBackToMenu.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labBackToMenu.Left = VB6.TwipsToPixelsX(TempWide * (600 / w))
		
		imgBackToMenu.Top = VB6.TwipsToPixelsY(TempHigh * (6150 / h))
		imgBackToMenu.Left = VB6.TwipsToPixelsX(TempWide * (60 / w))
		imgBackToMenu.Width = VB6.TwipsToPixelsX(TempWide * (495 / w))
		
		comProdPurchasePrint.Top = VB6.TwipsToPixelsY(TempHigh * (6150 / h))
		comProdPurchasePrint.Left = VB6.TwipsToPixelsX(TempWide * (8880 / w))
		
		labProdPurchaseHelp.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labProdPurchaseHelp.Left = VB6.TwipsToPixelsX(TempWide * (8340 / w))
		
		labInsert.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labInsert.Left = VB6.TwipsToPixelsX(TempWide * (4080 / w))
		labInsert.Width = VB6.TwipsToPixelsX(TempWide * (975 / w))
		
	End Sub
	Public Sub drawthevalues()
		
		Dim i As Short
		Dim x As Short
		
		DoNotChange = True
		Call ScreenCalc()
		For i = 0 To 19
			Select Case i
				Case 0 To 5, 14 To 19
					LabProdPurchaseTitles(i).Enabled = True
					txtProdPurchaseValues(i).Enabled = True
					txtProdPurchaseValues(i + 20).Enabled = True
					If CellValues(WhichScreen, i, WhichSegment).Changed = True Then
						txtProdPurchaseValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF0000)
					Else
						txtProdPurchaseValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
					End If
					If CellValues(WhichScreen, i + 20, WhichSegment).Changed = True Then
						txtProdPurchaseValues(i + 20).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF0000)
					Else
						txtProdPurchaseValues(i + 20).ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
					End If
			End Select
			Select Case i
				Case 0
					For x = 0 To 3
						If CellValues(EquipmentOne, x, WhichSegment).Value <> 0 Then
							Select Case x
								Case 0
									LabProdPurchaseTitles(i).Text = "Front-End Loader"
								Case 1
									LabProdPurchaseTitles(i).Text = "Hydraulic Shovel"
								Case 2
									LabProdPurchaseTitles(i).Text = "Mechanical Shovel"
								Case 3
									LabProdPurchaseTitles(i).Text = "Walking Dragline"
							End Select
							If CellValues(WhichScreen, i, WhichSegment).Value <> 0 Then
								txtProdPurchaseValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value)), "$###,###,##0")
							End If
							If CellValues(WhichScreen, i + 20, WhichSegment).Value <> 0 Then
								txtProdPurchaseValues(i + 20).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i + 20, WhichSegment).Value)), "$###,###,##0")
							End If
						End If
					Next x
					If CellValues(EquipmentOne, 20, WhichSegment).Value <> 0 Then
						LabProdPurchaseTitles(i).Text = "Scraper"
						If CellValues(WhichScreen, i, WhichSegment).Value <> 0 Then
							txtProdPurchaseValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value)), "$###,###,##0")
						End If
						If CellValues(WhichScreen, i + 20, WhichSegment).Value <> 0 Then
							txtProdPurchaseValues(i + 20).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i + 20, WhichSegment).Value)), "$###,###,##0")
						End If
						LabProdPurchaseTitles(i + 1).Enabled = False
						txtProdPurchaseValues(i + 1).Enabled = False
						txtProdPurchaseValues(i + 21).Enabled = False
					End If
				Case 1
					If CellValues(EquipmentOne, 4, WhichSegment).Value <> 0 Then
						LabProdPurchaseTitles(i).Text = "Rear-Dump Truck"
					ElseIf CellValues(EquipmentOne, 21, WhichSegment).Value <> 0 Then 
						LabProdPurchaseTitles(i).Text = "Acticulated Hauler"
					End If
					If CellValues(WhichScreen, i, WhichSegment).Value <> 0 Then
						txtProdPurchaseValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value)), "$###,###,##0")
					End If
					If CellValues(WhichScreen, i + 20, WhichSegment).Value <> 0 Then
						txtProdPurchaseValues(i + 20).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i + 20, WhichSegment).Value)), "$###,###,##0")
					End If
				Case 2
					For x = 5 To 8
						If CellValues(EquipmentOne, x, WhichSegment).Value <> 0 Then
							Select Case x
								Case 5
									LabProdPurchaseTitles(i).Text = "Front-End Loader"
								Case 6
									LabProdPurchaseTitles(i).Text = "Hydraulic Shovel"
								Case 7
									LabProdPurchaseTitles(i).Text = "Mechanical Shovel"
								Case 8
									LabProdPurchaseTitles(i).Text = "Walking Dragline"
							End Select
							If CellValues(WhichScreen, i, WhichSegment).Value <> 0 Then
								txtProdPurchaseValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value)), "$###,###,##0")
							End If
							If CellValues(WhichScreen, i + 20, WhichSegment).Value <> 0 Then
								txtProdPurchaseValues(i + 20).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i + 20, WhichSegment).Value)), "$###,###,##0")
							End If
						End If
					Next x
					If CellValues(EquipmentOne, 25, WhichSegment).Value <> 0 Then
						LabProdPurchaseTitles(i).Text = "Scraper"
						If CellValues(WhichScreen, i, WhichSegment).Value <> 0 Then
							txtProdPurchaseValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value)), "$###,###,##0")
						End If
						If CellValues(WhichScreen, i + 20, WhichSegment).Value <> 0 Then
							txtProdPurchaseValues(i + 20).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i + 20, WhichSegment).Value)), "$###,###,##0")
						End If
						LabProdPurchaseTitles(i).Text = "Scraper"
						LabProdPurchaseTitles(i + 1).Enabled = False
						txtProdPurchaseValues(i + 1).Enabled = False
						txtProdPurchaseValues(i + 21).Enabled = False
					End If
				Case 3
					If CellValues(EquipmentOne, 9, WhichSegment).Value <> 0 Then
						LabProdPurchaseTitles(i).Text = "Rear-Dump Truck"
					ElseIf CellValues(EquipmentOne, 26, WhichSegment).Value <> 0 Then 
						LabProdPurchaseTitles(i).Text = "Acticulated Hauler"
					End If
					If CellValues(WhichScreen, i, WhichSegment).Value <> 0 Then
						txtProdPurchaseValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value)), "$###,###,##0")
					End If
					If CellValues(WhichScreen, i + 20, WhichSegment).Value <> 0 Then
						txtProdPurchaseValues(i + 20).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i + 20, WhichSegment).Value)), "$###,###,##0")
					End If
				Case 4
					If CellValues(WhichScreen, i, WhichSegment).Value <> 0 Then
						LabProdPurchaseTitles(i).Enabled = True
						txtProdPurchaseValues(i).Enabled = True
						txtProdPurchaseValues(i + 20).Enabled = True
						txtProdPurchaseValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value)), "$###,###,##0")
					Else
						LabProdPurchaseTitles(i).Enabled = False
						txtProdPurchaseValues(i).Enabled = False
						txtProdPurchaseValues(i + 20).Enabled = False
					End If
					If CellValues(WhichScreen, i + 20, WhichSegment).Value <> 0 Then
						txtProdPurchaseValues(i + 20).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i + 20, WhichSegment).Value)), "$###,###,##0")
					End If
				Case 5
					If CellValues(WhichScreen, i, WhichSegment).Value <> 0 Then
						LabProdPurchaseTitles(i).Enabled = True
						txtProdPurchaseValues(i).Enabled = True
						txtProdPurchaseValues(i + 20).Enabled = True
						txtProdPurchaseValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value)), "$###,###,##0")
					Else
						LabProdPurchaseTitles(i).Enabled = False
						txtProdPurchaseValues(i).Enabled = False
						txtProdPurchaseValues(i + 20).Enabled = False
					End If
					If CellValues(WhichScreen, i + 20, WhichSegment).Value <> 0 Then
						txtProdPurchaseValues(i + 20).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i + 20, WhichSegment).Value)), "$###,###,##0")
					End If
				Case 14
					If CellValues(EquipmentOne, 22, WhichSegment).Value <> 0 Then
						LabProdPurchaseTitles(i).Text = "Jaw Crusher"
					ElseIf CellValues(EquipmentOne, 23, WhichSegment).Value <> 0 Then 
						LabProdPurchaseTitles(i).Text = "Gyratory Crusher"
					Else
						LabProdPurchaseTitles(i).Enabled = False
						txtProdPurchaseValues(i).Enabled = False
						txtProdPurchaseValues(i + 20).Enabled = False
					End If
					If CellValues(WhichScreen, i, WhichSegment).Value <> 0 Then
						txtProdPurchaseValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value)), "$###,###,##0")
					End If
					If CellValues(WhichScreen, i + 20, WhichSegment).Value <> 0 Then
						txtProdPurchaseValues(i + 20).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i + 20, WhichSegment).Value)), "$###,###,##0")
					End If
				Case 15
					If CellValues(WhichScreen, i, WhichSegment).Value <> 0 Then
						txtProdPurchaseValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value)), "$###,###,##0")
					End If
					If CellValues(WhichScreen, i + 20, WhichSegment).Value <> 0 Then
						txtProdPurchaseValues(i + 20).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i + 20, WhichSegment).Value)), "$###,###,##0")
					End If
					If CellValues(EquipmentOne, 24, WhichSegment).Value = 0 Then
						LabProdPurchaseTitles(i).Enabled = False
						txtProdPurchaseValues(i).Enabled = False
						txtProdPurchaseValues(i + 20).Enabled = False
					Else
						LabProdPurchaseTitles(1).Enabled = False
						txtProdPurchaseValues(1).Enabled = False
						txtProdPurchaseValues(21).Enabled = False
					End If
				Case 16
					If CellValues(EquipmentOne, 27, WhichSegment).Value <> 0 Then
						LabProdPurchaseTitles(i).Text = "Jaw Crusher"
					ElseIf CellValues(EquipmentOne, 28, WhichSegment).Value <> 0 Then 
						LabProdPurchaseTitles(i).Text = "Gyratory Crusher"
					Else
						LabProdPurchaseTitles(i).Enabled = False
						txtProdPurchaseValues(i).Enabled = False
						txtProdPurchaseValues(i + 20).Enabled = False
					End If
					If CellValues(WhichScreen, i, WhichSegment).Value <> 0 Then
						txtProdPurchaseValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value)), "$###,###,##0")
					End If
					If CellValues(WhichScreen, i + 20, WhichSegment).Value <> 0 Then
						txtProdPurchaseValues(i + 20).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i + 20, WhichSegment).Value)), "$###,###,##0")
					End If
				Case 17
					If CellValues(WhichScreen, i, WhichSegment).Value <> 0 Then
						txtProdPurchaseValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value)), "$###,###,##0")
					End If
					If CellValues(WhichScreen, i + 20, WhichSegment).Value <> 0 Then
						txtProdPurchaseValues(i + 20).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i + 20, WhichSegment).Value)), "$###,###,##0")
					End If
					If CellValues(EquipmentOne, 29, WhichSegment).Value = 0 Then
						LabProdPurchaseTitles(i).Enabled = False
						txtProdPurchaseValues(i).Enabled = False
						txtProdPurchaseValues(i + 20).Enabled = False
					Else
						LabProdPurchaseTitles(3).Enabled = False
						txtProdPurchaseValues(3).Enabled = False
						txtProdPurchaseValues(23).Enabled = False
					End If
				Case 18
					If CellValues(WhichScreen, i, WhichSegment).Value <> 0 Then
						LabProdPurchaseTitles(i).Enabled = True
						txtProdPurchaseValues(i).Enabled = True
						txtProdPurchaseValues(i + 20).Enabled = True
						txtProdPurchaseValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value)), "$###,###,##0")
					Else
						LabProdPurchaseTitles(i).Enabled = False
						txtProdPurchaseValues(i).Enabled = False
						txtProdPurchaseValues(i + 20).Enabled = False
					End If
					If CellValues(WhichScreen, i + 20, WhichSegment).Value <> 0 Then
						txtProdPurchaseValues(i + 20).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i + 20, WhichSegment).Value)), "$###,###,##0")
					End If
				Case 19
					If CellValues(WhichScreen, i, WhichSegment).Value <> 0 Then
						LabProdPurchaseTitles(i).Enabled = True
						txtProdPurchaseValues(i).Enabled = True
						txtProdPurchaseValues(i + 20).Enabled = True
						txtProdPurchaseValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value)), "$###,###,##0")
					Else
						LabProdPurchaseTitles(i).Enabled = False
						txtProdPurchaseValues(i).Enabled = False
						txtProdPurchaseValues(i + 20).Enabled = False
					End If
					If CellValues(WhichScreen, i + 20, WhichSegment).Value <> 0 Then
						txtProdPurchaseValues(i + 20).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i + 20, WhichSegment).Value)), "$###,###,##0")
					End If
			End Select
		Next i
		
		DoNotChange = False
		
	End Sub
	Private Sub txtProdPurchaseValues_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtProdPurchaseValues.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Dim Index As Short = txtProdPurchaseValues.GetIndex(eventSender)
		
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
						If InStr(txtProdPurchaseValues(Index).Text, ".") = 0 Then
							System.Windows.Forms.SendKeys.Send("{DELETE}")
						End If
					Else
						System.Windows.Forms.SendKeys.Send("{DELETE}")
					End If
			End Select
		End If
		
	End Sub
	Private Sub txtProdPurchaseValues_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtProdPurchaseValues.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = txtProdPurchaseValues.GetIndex(eventSender)
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
	Private Sub txtProdPurchaseValues_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtProdPurchaseValues.Leave
		Dim Index As Short = txtProdPurchaseValues.GetIndex(eventSender)
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
		For i = 1 To Len(txtProdPurchaseValues(Sample).Text)
			Digit.Value = Mid(txtProdPurchaseValues(Sample).Text, i, 1)
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
		
		For x = 0 To 19
			Select Case x
				Case 0
					If CellValues(WhichScreen, x + 20, WhichSegment).Changed = False Then
						CellValues(WhichScreen, x + 20, WhichSegment).Value = CellValues(WhichScreen, x, WhichSegment).Value * (CellValues(EquipmentTwo, 0, WhichSegment).Value + CellValues(EquipmentTwo, 1, WhichSegment).Value + CellValues(EquipmentTwo, 2, WhichSegment).Value + CellValues(EquipmentTwo, 3, WhichSegment).Value + CellValues(EquipmentTwo, 20, WhichSegment).Value)
					End If
				Case 1
					If CellValues(WhichScreen, x + 20, WhichSegment).Changed = False Then
						If CellValues(EquipmentOne, 4, WhichSegment).Value <> 0 Then
							CellValues(WhichScreen, x + 20, WhichSegment).Value = CellValues(WhichScreen, x, WhichSegment).Value * CellValues(EquipmentTwo, 4, WhichSegment).Value
						ElseIf CellValues(EquipmentOne, 21, WhichSegment).Value <> 0 Then 
							CellValues(WhichScreen, x + 20, WhichSegment).Value = CellValues(WhichScreen, x, WhichSegment).Value * CellValues(EquipmentTwo, 21, WhichSegment).Value
						End If
					End If
				Case 2
					If CellValues(WhichScreen, x + 20, WhichSegment).Changed = False Then
						CellValues(WhichScreen, x + 20, WhichSegment).Value = CellValues(WhichScreen, x, WhichSegment).Value * (CellValues(EquipmentTwo, 5, WhichSegment).Value + CellValues(EquipmentTwo, 6, WhichSegment).Value + CellValues(EquipmentTwo, 7, WhichSegment).Value + CellValues(EquipmentTwo, 8, WhichSegment).Value + CellValues(EquipmentTwo, 25, WhichSegment).Value)
					End If
				Case 3
					If CellValues(WhichScreen, x + 20, WhichSegment).Changed = False Then
						If CellValues(EquipmentOne, 9, WhichSegment).Value <> 0 Then
							CellValues(WhichScreen, x + 20, WhichSegment).Value = CellValues(WhichScreen, x, WhichSegment).Value * CellValues(EquipmentTwo, 9, WhichSegment).Value
						ElseIf CellValues(EquipmentOne, 26, WhichSegment).Value <> 0 Then 
							CellValues(WhichScreen, x + 20, WhichSegment).Value = CellValues(WhichScreen, x, WhichSegment).Value * CellValues(EquipmentTwo, 26, WhichSegment).Value
						End If
					End If
				Case 4 To 5
					If CellValues(WhichScreen, x + 20, WhichSegment).Changed = False Then
						CellValues(WhichScreen, x + 20, WhichSegment).Value = CellValues(WhichScreen, x, WhichSegment).Value * CellValues(EquipmentTwo, x + 6, WhichSegment).Value
					End If
				Case 6 To 13
					If CellValues(WhichScreen, x + 20, WhichSegment).Changed = False Then
						CellValues(WhichScreen, x + 20, WhichSegment).Value = CellValues(WhichScreen, x, WhichSegment).Value * CellValues(EquipmentTwo, x + 6, WhichSegment).Value
					End If
				Case 14
					If CellValues(WhichScreen, x + 20, WhichSegment).Changed = False Then
						If CellValues(EquipmentTwo, 22, WhichSegment).Value <> 0 Then
							CellValues(WhichScreen, x + 20, WhichSegment).Value = CellValues(WhichScreen, x, WhichSegment).Value * CellValues(EquipmentTwo, 22, WhichSegment).Value
						ElseIf CellValues(EquipmentTwo, 23, WhichSegment).Value <> 0 Then 
							CellValues(WhichScreen, x + 20, WhichSegment).Value = CellValues(WhichScreen, x, WhichSegment).Value * CellValues(EquipmentTwo, 23, WhichSegment).Value
						End If
					End If
				Case 15
					If CellValues(WhichScreen, x + 20, WhichSegment).Changed = False Then
						CellValues(WhichScreen, x + 20, WhichSegment).Value = CellValues(WhichScreen, x, WhichSegment).Value * CellValues(EquipmentTwo, 24, WhichSegment).Value
					End If
				Case 16
					If CellValues(WhichScreen, x + 20, WhichSegment).Changed = False Then
						If CellValues(EquipmentTwo, 27, WhichSegment).Value <> 0 Then
							CellValues(WhichScreen, x + 20, WhichSegment).Value = CellValues(WhichScreen, x, WhichSegment).Value * CellValues(EquipmentTwo, 27, WhichSegment).Value
						ElseIf CellValues(EquipmentTwo, 28, WhichSegment).Value <> 0 Then 
							CellValues(WhichScreen, x + 20, WhichSegment).Value = CellValues(WhichScreen, x, WhichSegment).Value * CellValues(EquipmentTwo, 28, WhichSegment).Value
						End If
					End If
				Case 17
					If CellValues(WhichScreen, x + 20, WhichSegment).Changed = False Then
						CellValues(WhichScreen, x + 20, WhichSegment).Value = CellValues(WhichScreen, x, WhichSegment).Value * CellValues(EquipmentTwo, 29, WhichSegment).Value
					End If
				Case Else
					If CellValues(WhichScreen, x + 20, WhichSegment).Changed = False Then
						CellValues(WhichScreen, x + 20, WhichSegment).Value = CellValues(WhichScreen, x, WhichSegment).Value * CellValues(EquipmentTwo, x + 12, WhichSegment).Value
					End If
			End Select
		Next x
	End Sub
End Class