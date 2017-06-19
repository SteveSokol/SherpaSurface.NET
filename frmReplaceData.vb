Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks
Friend Class frmReplaceData
	Inherits System.Windows.Forms.Form
	Dim TempHigh As Single
	Dim TempWide As Single
	'UPGRADE_WARNING: Event chkReplaceCheck.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkReplaceCheck_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkReplaceCheck.CheckStateChanged
		If DoNotChange = True Then Exit Sub
		If chkReplaceCheck.CheckState = 1 Then
			CellValues(WhichScreen, 20, WhichSegment).Value = 1
		Else
			CellValues(WhichScreen, 20, WhichSegment).Value = 0
		End If
	End Sub
	Private Sub comReplacePrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles comReplacePrint.Click
		PrintForm1.Print(Me, PowerPacks.Printing.PrintForm.PrintOption.CompatibleModeClientAreaOnly)
	End Sub
	'UPGRADE_WARNING: Form event frmReplaceData.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmReplaceData_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated

        'Dim baseunit As String
        'Dim baselength As Short
        Dim i As Short
		
		If IsHelpOn = True Then
			txtReplaceValues(WhichCell).Focus()
			IsHelpOn = False
		Else
			For i = 0 To 5
				If optSegment(i).Checked = True Then WhichSegment = i
			Next i
			
			WhichScreen = Replace_Renamed
			
			Call drawthevalues()
			
			DoNotChange = True
			If CellValues(WhichScreen, 20, WhichSegment).Value = 1 Then
				chkReplaceCheck.CheckState = System.Windows.Forms.CheckState.Checked
			Else
				chkReplaceCheck.CheckState = System.Windows.Forms.CheckState.Unchecked
			End If
			DoNotChange = False
			
			If InsertFlag = True Then
				labInsert.Text = "Insert"
			Else
				labInsert.Text = "Typeover"
			End If
			
			WhichCell = 0
			
			txtReplaceValues(0).Focus()
		End If
		
	End Sub
	Private Sub frmReplaceData_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
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
	'UPGRADE_WARNING: Event frmReplaceData.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub frmReplaceData_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		TempWide = VB6.PixelsToTwipsX(Me.ClientRectangle.Width)
		TempHigh = VB6.PixelsToTwipsY(Me.ClientRectangle.Height)
		Call screenstuff()
	End Sub
	Private Sub frmReplaceData_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
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
	Private Sub labReplaceHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles labReplaceHelp.Click
		Dim StartHelp As Short
		IsHelpOn = True
		StartHelp = 157
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
			txtReplaceValues(WhichCell).Focus()
		End If
	End Sub
	'UPGRADE_WARNING: Event txtReplaceValues.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtReplaceValues_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtReplaceValues.TextChanged
		Dim Index As Short = txtReplaceValues.GetIndex(eventSender)
		If DoNotChange = True Then Exit Sub
	End Sub
	Private Sub txtReplaceValues_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtReplaceValues.Enter
		Dim Index As Short = txtReplaceValues.GetIndex(eventSender)
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
		
		q = (480 / h) * TempHigh
		r = (540 / w) * TempWide
		s = (1620 / h) * TempHigh
		t = (2580 / h) * TempHigh
		
		y = (420 / h) * TempHigh
		z = (300 / w) * TempWide
		
		labReplaceHeading(0).Top = VB6.TwipsToPixelsY(TempHigh * (180 / h))
		labReplaceHeading(0).Left = VB6.TwipsToPixelsX(TempWide * (180 / w))
		labReplaceHeading(0).Width = VB6.TwipsToPixelsX(TempWide * (4155 / w))
		
		labReplaceLabels(0).Top = VB6.TwipsToPixelsY(TempHigh * (720 / h))
		labReplaceLabels(0).Left = VB6.TwipsToPixelsX(TempWide * (180 / w))
		labReplaceLabels(0).Width = VB6.TwipsToPixelsX(TempWide * (1935 / w))
		labReplaceLabels(1).Top = VB6.TwipsToPixelsY(TempHigh * (840 / h))
		labReplaceLabels(1).Left = VB6.TwipsToPixelsX(TempWide * (2460 / w))
		labReplaceLabels(1).Width = VB6.TwipsToPixelsX(TempWide * (1935 / w))
		labReplaceLabels(2).Top = VB6.TwipsToPixelsY(TempHigh * (2520 / h))
		labReplaceLabels(2).Left = VB6.TwipsToPixelsX(TempWide * (120 / w))
		labReplaceLabels(9).Top = VB6.TwipsToPixelsY(TempHigh * (1800 / h))
		labReplaceLabels(9).Left = VB6.TwipsToPixelsX(TempWide * (1260 / w))
		
		chkReplaceCheck.Top = VB6.TwipsToPixelsY(TempHigh * (5880 / h))
		chkReplaceCheck.Left = VB6.TwipsToPixelsX(TempWide * (5880 / w))
		chkReplaceCheck.Width = VB6.TwipsToPixelsX(TempWide * (1935 / w))
		
		For x = 0 To 1
			labReplaceLabels(x + 3).Top = VB6.TwipsToPixelsY((TempHigh * (540 / h)) + (x * t))
			labReplaceLabels(x + 3).Left = VB6.TwipsToPixelsX(TempWide * (4800 / w))
		Next x
		
		For x = 5 To 6
			labReplaceLabels(x).Top = VB6.TwipsToPixelsY(TempHigh * (2280 / h))
			labReplaceLabels(x + 2).Top = VB6.TwipsToPixelsY(TempHigh * (300 / h))
			If x = 5 Then
				labReplaceLabels(x).Left = VB6.TwipsToPixelsX(TempWide * (1020 / w))
				labReplaceLabels(x).Width = VB6.TwipsToPixelsX(TempWide * (1665 / w))
				labReplaceLabels(x + 2).Left = VB6.TwipsToPixelsX(TempWide * (5880 / w))
				labReplaceLabels(x + 2).Width = VB6.TwipsToPixelsX(TempWide * (1665 / w))
			Else
				labReplaceLabels(x).Left = VB6.TwipsToPixelsX(TempWide * (2880 / w))
				labReplaceLabels(x).Width = VB6.TwipsToPixelsX(TempWide * (1545 / w))
				labReplaceLabels(x + 2).Left = VB6.TwipsToPixelsX(TempWide * (7380 / w))
				labReplaceLabels(x + 2).Width = VB6.TwipsToPixelsX(TempWide * (1545 / w))
			End If
		Next x
		
		For x = 0 To 5
			optSegment(x).Top = VB6.TwipsToPixelsY(TempHigh * (1020 / h))
			optSegment(x).Left = VB6.TwipsToPixelsX((TempWide * (300 / w)) + (x * z))
			optSegment(x).Width = VB6.TwipsToPixelsX(TempWide * (195 / w))
			labSegment(x).Top = VB6.TwipsToPixelsY(TempHigh * (1260 / h))
			labSegment(x).Left = VB6.TwipsToPixelsX((TempWide * (300 / w)) + (x * z))
			labSegment(x).Width = VB6.TwipsToPixelsX(TempWide * (195 / w))
		Next x
		
		txtSegmentLabel.Top = VB6.TwipsToPixelsY(TempHigh * (1140 / h))
		txtSegmentLabel.Left = VB6.TwipsToPixelsX(TempWide * (2460 / w))
		txtSegmentLabel.Width = VB6.TwipsToPixelsX(TempWide * (1935 / w))
		
		For x = 0 To 1
			LabReplaceTitles(x).Top = VB6.TwipsToPixelsY((TempHigh * (750 / h)) + (x * y))
			LabReplaceTitles(x).Left = VB6.TwipsToPixelsX(TempWide * (5520 / w))
			LabReplaceTitles(x).Width = VB6.TwipsToPixelsX(TempWide * (1635 / w))
			txtReplaceValues(x).Top = VB6.TwipsToPixelsY((TempHigh * (690 / h)) + (x * y))
			txtReplaceValues(x).Left = VB6.TwipsToPixelsX(TempWide * (7380 / w))
			txtReplaceValues(x).Width = VB6.TwipsToPixelsX(TempWide * (855 / w))
			LabReplaceUnits(x).Top = VB6.TwipsToPixelsY((TempHigh * (750 / h)) + (x * y))
			LabReplaceUnits(x).Left = VB6.TwipsToPixelsX(TempWide * (8280 / w))
			LabReplaceTitles(x + 2).Top = VB6.TwipsToPixelsY((TempHigh * (3330 / h)) + (x * y))
			LabReplaceTitles(x + 2).Left = VB6.TwipsToPixelsX(TempWide * (5520 / w))
			LabReplaceTitles(x + 2).Width = VB6.TwipsToPixelsX(TempWide * (1635 / w))
			txtReplaceValues(x + 2).Top = VB6.TwipsToPixelsY((TempHigh * (3270 / h)) + (x * y))
			txtReplaceValues(x + 2).Left = VB6.TwipsToPixelsX(TempWide * (7380 / w))
			txtReplaceValues(x + 2).Width = VB6.TwipsToPixelsX(TempWide * (855 / w))
			LabReplaceUnits(x + 2).Top = VB6.TwipsToPixelsY((TempHigh * (3330 / h)) + (x * y))
			LabReplaceUnits(x + 2).Left = VB6.TwipsToPixelsX(TempWide * (8280 / w))
			LabReplaceTitles(x + 4).Top = VB6.TwipsToPixelsY((TempHigh * (2850 / h)) - (x * y))
			LabReplaceTitles(x + 4).Left = VB6.TwipsToPixelsX(TempWide * (5520 / w))
			LabReplaceTitles(x + 4).Width = VB6.TwipsToPixelsX(TempWide * (1635 / w))
			txtReplaceValues(x + 4).Top = VB6.TwipsToPixelsY((TempHigh * (2790 / h)) - (x * y))
			txtReplaceValues(x + 4).Left = VB6.TwipsToPixelsX(TempWide * (7380 / w))
			txtReplaceValues(x + 4).Width = VB6.TwipsToPixelsX(TempWide * (855 / w))
			LabReplaceUnits(x + 4).Top = VB6.TwipsToPixelsY((TempHigh * (2850 / h)) - (x * y))
			LabReplaceUnits(x + 4).Left = VB6.TwipsToPixelsX(TempWide * (8280 / w))
			LabReplaceTitles(x + 14).Top = VB6.TwipsToPixelsY((TempHigh * (1590 / h)) + (x * y))
			LabReplaceTitles(x + 14).Left = VB6.TwipsToPixelsX(TempWide * (5520 / w))
			LabReplaceTitles(x + 14).Width = VB6.TwipsToPixelsX(TempWide * (1635 / w))
			txtReplaceValues(x + 14).Top = VB6.TwipsToPixelsY((TempHigh * (1530 / h)) + (x * y))
			txtReplaceValues(x + 14).Left = VB6.TwipsToPixelsX(TempWide * (7380 / w))
			txtReplaceValues(x + 14).Width = VB6.TwipsToPixelsX(TempWide * (855 / w))
			LabReplaceUnits(x + 14).Top = VB6.TwipsToPixelsY((TempHigh * (1590 / h)) + (x * y))
			LabReplaceUnits(x + 14).Left = VB6.TwipsToPixelsX(TempWide * (8280 / w))
			LabReplaceTitles(x + 16).Top = VB6.TwipsToPixelsY((TempHigh * (4170 / h)) + (x * y))
			LabReplaceTitles(x + 16).Left = VB6.TwipsToPixelsX(TempWide * (5520 / w))
			LabReplaceTitles(x + 16).Width = VB6.TwipsToPixelsX(TempWide * (1635 / w))
			txtReplaceValues(x + 16).Top = VB6.TwipsToPixelsY((TempHigh * (4110 / h)) + (x * y))
			txtReplaceValues(x + 16).Left = VB6.TwipsToPixelsX(TempWide * (7380 / w))
			txtReplaceValues(x + 16).Width = VB6.TwipsToPixelsX(TempWide * (855 / w))
			LabReplaceUnits(x + 16).Top = VB6.TwipsToPixelsY((TempHigh * (4170 / h)) + (x * y))
			LabReplaceUnits(x + 16).Left = VB6.TwipsToPixelsX(TempWide * (8280 / w))
			LabReplaceTitles(x + 18).Top = VB6.TwipsToPixelsY((TempHigh * (5430 / h)) - (x * y))
			LabReplaceTitles(x + 18).Left = VB6.TwipsToPixelsX(TempWide * (5520 / w))
			LabReplaceTitles(x + 18).Width = VB6.TwipsToPixelsX(TempWide * (1635 / w))
			txtReplaceValues(x + 18).Top = VB6.TwipsToPixelsY((TempHigh * (5370 / h)) - (x * y))
			txtReplaceValues(x + 18).Left = VB6.TwipsToPixelsX(TempWide * (7380 / w))
			txtReplaceValues(x + 18).Width = VB6.TwipsToPixelsX(TempWide * (855 / w))
			LabReplaceUnits(x + 18).Top = VB6.TwipsToPixelsY((TempHigh * (5430 / h)) - (x * y))
			LabReplaceUnits(x + 18).Left = VB6.TwipsToPixelsX(TempWide * (8280 / w))
		Next x
		
		For x = 0 To 7
			LabReplaceTitles(x + 6).Top = VB6.TwipsToPixelsY((TempHigh * (2730 / h)) + (x * y))
			LabReplaceTitles(x + 6).Left = VB6.TwipsToPixelsX(TempWide * (1020 / w))
			LabReplaceTitles(x + 6).Width = VB6.TwipsToPixelsX(TempWide * (1635 / w))
			txtReplaceValues(x + 6).Top = VB6.TwipsToPixelsY((TempHigh * (2670 / h)) + (x * y))
			txtReplaceValues(x + 6).Left = VB6.TwipsToPixelsX(TempWide * (2880 / w))
			txtReplaceValues(x + 6).Width = VB6.TwipsToPixelsX(TempWide * (855 / w))
			LabReplaceUnits(x + 6).Top = VB6.TwipsToPixelsY((TempHigh * (2730 / h)) + (x * y))
			LabReplaceUnits(x + 6).Left = VB6.TwipsToPixelsX(TempWide * (3780 / w))
		Next x
		
		txtReplaceValues(21).Top = VB6.TwipsToPixelsY(TempHigh * (1770 / h))
		txtReplaceValues(21).Left = VB6.TwipsToPixelsX(TempWide * (780 / w))
		txtReplaceValues(21).Width = VB6.TwipsToPixelsX(TempWide * (315 / w))
		
		LineHorizontal(0).X1 = VB6.TwipsToPixelsX(TempWide * (120 / w))
		LineHorizontal(0).X2 = VB6.TwipsToPixelsX(TempWide * (4500 / w))
		LineHorizontal(0).Y1 = VB6.TwipsToPixelsY(TempHigh * (2160 / h))
		LineHorizontal(0).Y2 = VB6.TwipsToPixelsY(TempHigh * (2160 / h))
		
		LineHorizontal(1).X1 = VB6.TwipsToPixelsX(TempWide * (960 / w))
		LineHorizontal(1).X2 = VB6.TwipsToPixelsX(TempWide * (4440 / w))
		LineHorizontal(1).Y1 = VB6.TwipsToPixelsY(TempHigh * (2580 / h))
		LineHorizontal(1).Y2 = VB6.TwipsToPixelsY(TempHigh * (2580 / h))
		
		LineHorizontal(2).X1 = VB6.TwipsToPixelsX(TempWide * (120 / w))
		LineHorizontal(2).X2 = VB6.TwipsToPixelsX(TempWide * (4560 / w))
		LineHorizontal(2).Y1 = VB6.TwipsToPixelsY(TempHigh * (6000 / h))
		LineHorizontal(2).Y2 = VB6.TwipsToPixelsY(TempHigh * (6000 / h))
		
		LineHorizontal(3).X1 = VB6.TwipsToPixelsX(TempWide * (4800 / w))
		LineHorizontal(3).X2 = VB6.TwipsToPixelsX(TempWide * (9060 / w))
		LineHorizontal(3).Y1 = VB6.TwipsToPixelsY(TempHigh * (180 / h))
		LineHorizontal(3).Y2 = VB6.TwipsToPixelsY(TempHigh * (180 / h))
		
		LineHorizontal(4).X1 = VB6.TwipsToPixelsX(TempWide * (5220 / w))
		LineHorizontal(4).X2 = VB6.TwipsToPixelsX(TempWide * (8940 / w))
		LineHorizontal(4).Y1 = VB6.TwipsToPixelsY(TempHigh * (600 / h))
		LineHorizontal(4).Y2 = VB6.TwipsToPixelsY(TempHigh * (600 / h))
		
		LineHorizontal(5).X1 = VB6.TwipsToPixelsX(TempWide * (5460 / w))
		LineHorizontal(5).X2 = VB6.TwipsToPixelsX(TempWide * (8940 / w))
		LineHorizontal(5).Y1 = VB6.TwipsToPixelsY(TempHigh * (3180 / h))
		LineHorizontal(5).Y2 = VB6.TwipsToPixelsY(TempHigh * (3180 / h))
		
		LineHorizontal(6).X1 = VB6.TwipsToPixelsX(TempWide * (4800 / w))
		LineHorizontal(6).X2 = VB6.TwipsToPixelsX(TempWide * (9060 / w))
		LineHorizontal(6).Y1 = VB6.TwipsToPixelsY(TempHigh * (5760 / h))
		LineHorizontal(6).Y2 = VB6.TwipsToPixelsY(TempHigh * (5760 / h))
		
		LineHorizontal(7).X1 = VB6.TwipsToPixelsX(TempWide * (540 / w))
		LineHorizontal(7).X2 = VB6.TwipsToPixelsX(TempWide * (4080 / w))
		LineHorizontal(7).Y1 = VB6.TwipsToPixelsY(TempHigh * (1680 / h))
		LineHorizontal(7).Y2 = VB6.TwipsToPixelsY(TempHigh * (1680 / h))
		
		LineHorizontal(8).X1 = VB6.TwipsToPixelsX(TempWide * (5580 / w))
		LineHorizontal(8).X2 = VB6.TwipsToPixelsX(TempWide * (8100 / w))
		LineHorizontal(8).Y1 = VB6.TwipsToPixelsY(TempHigh * (6240 / h))
		LineHorizontal(8).Y2 = VB6.TwipsToPixelsY(TempHigh * (6240 / h))
		
		LineVertical(0).X1 = VB6.TwipsToPixelsX(TempWide * (180 / w))
		LineVertical(0).X2 = VB6.TwipsToPixelsX(TempWide * (180 / w))
		LineVertical(0).Y1 = VB6.TwipsToPixelsY(TempHigh * (2100 / h))
		LineVertical(0).Y2 = VB6.TwipsToPixelsY(TempHigh * (2520 / h))
		
		LineVertical(1).X1 = VB6.TwipsToPixelsX(TempWide * (180 / w))
		LineVertical(1).X2 = VB6.TwipsToPixelsX(TempWide * (180 / w))
		LineVertical(1).Y1 = VB6.TwipsToPixelsY(TempHigh * (2760 / h))
		LineVertical(1).Y2 = VB6.TwipsToPixelsY(TempHigh * (6060 / h))
		
		LineVertical(2).X1 = VB6.TwipsToPixelsX(TempWide * (4500 / w))
		LineVertical(2).X2 = VB6.TwipsToPixelsX(TempWide * (4500 / w))
		LineVertical(2).Y1 = VB6.TwipsToPixelsY(TempHigh * (2100 / h))
		LineVertical(2).Y2 = VB6.TwipsToPixelsY(TempHigh * (6060 / h))
		
		LineVertical(3).X1 = VB6.TwipsToPixelsX(TempWide * (4860 / w))
		LineVertical(3).X2 = VB6.TwipsToPixelsX(TempWide * (4860 / w))
		LineVertical(3).Y1 = VB6.TwipsToPixelsY(TempHigh * (120 / h))
		LineVertical(3).Y2 = VB6.TwipsToPixelsY(TempHigh * (540 / h))
		
		LineVertical(4).X1 = VB6.TwipsToPixelsX(TempWide * (4860 / w))
		LineVertical(4).X2 = VB6.TwipsToPixelsX(TempWide * (4860 / w))
		LineVertical(4).Y1 = VB6.TwipsToPixelsY(TempHigh * (840 / h))
		LineVertical(4).Y2 = VB6.TwipsToPixelsY(TempHigh * (3120 / h))
		
		LineVertical(5).X1 = VB6.TwipsToPixelsX(TempWide * (4860 / w))
		LineVertical(5).X2 = VB6.TwipsToPixelsX(TempWide * (4860 / w))
		LineVertical(5).Y1 = VB6.TwipsToPixelsY(TempHigh * (3420 / h))
		LineVertical(5).Y2 = VB6.TwipsToPixelsY(TempHigh * (5820 / h))
		
		LineVertical(6).X1 = VB6.TwipsToPixelsX(TempWide * (9000 / w))
		LineVertical(6).X2 = VB6.TwipsToPixelsX(TempWide * (9000 / w))
		LineVertical(6).Y1 = VB6.TwipsToPixelsY(TempHigh * (120 / h))
		LineVertical(6).Y2 = VB6.TwipsToPixelsY(TempHigh * (5820 / h))
		
		LineVertical(7).X1 = VB6.TwipsToPixelsX(TempWide * (600 / w))
		LineVertical(7).X2 = VB6.TwipsToPixelsX(TempWide * (600 / w))
		LineVertical(7).Y1 = VB6.TwipsToPixelsY(TempHigh * (1620 / h))
		LineVertical(7).Y2 = VB6.TwipsToPixelsY(TempHigh * (2100 / h))
		
		LineVertical(8).X1 = VB6.TwipsToPixelsX(TempWide * (4020 / w))
		LineVertical(8).X2 = VB6.TwipsToPixelsX(TempWide * (4020 / w))
		LineVertical(8).Y1 = VB6.TwipsToPixelsY(TempHigh * (1620 / h))
		LineVertical(8).Y2 = VB6.TwipsToPixelsY(TempHigh * (2100 / h))
		
		LineVertical(9).X1 = VB6.TwipsToPixelsX(TempWide * (5640 / w))
		LineVertical(9).X2 = VB6.TwipsToPixelsX(TempWide * (5640 / w))
		LineVertical(9).Y1 = VB6.TwipsToPixelsY(TempHigh * (5820 / h))
		LineVertical(9).Y2 = VB6.TwipsToPixelsY(TempHigh * (6300 / h))
		
		LineVertical(10).X1 = VB6.TwipsToPixelsX(TempWide * (8040 / w))
		LineVertical(10).X2 = VB6.TwipsToPixelsX(TempWide * (8040 / w))
		LineVertical(10).Y1 = VB6.TwipsToPixelsY(TempHigh * (5820 / h))
		LineVertical(10).Y2 = VB6.TwipsToPixelsY(TempHigh * (6300 / h))
		
		labBackToMenu.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labBackToMenu.Left = VB6.TwipsToPixelsX(TempWide * (600 / w))
		
		imgBackToMenu.Top = VB6.TwipsToPixelsY(TempHigh * (6150 / h))
		imgBackToMenu.Left = VB6.TwipsToPixelsX(TempWide * (60 / w))
		imgBackToMenu.Width = VB6.TwipsToPixelsX(TempWide * (495 / w))
		
		comReplacePrint.Top = VB6.TwipsToPixelsY(TempHigh * (6150 / h))
		comReplacePrint.Left = VB6.TwipsToPixelsX(TempWide * (8880 / w))
		
		labReplaceHelp.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labReplaceHelp.Left = VB6.TwipsToPixelsX(TempWide * (8340 / w))
		
		labInsert.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labInsert.Left = VB6.TwipsToPixelsX(TempWide * (4080 / w))
		labInsert.Width = VB6.TwipsToPixelsX(TempWide * (975 / w))
		
	End Sub
	Public Sub drawthevalues()
		
		Dim i As Short
		Dim x As Short
		
		DoNotChange = True
		
		For i = 0 To 19
			If CellValues(WhichScreen, i, WhichSegment).Changed = True Then
				txtReplaceValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF0000)
			Else
				txtReplaceValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
			End If
			Select Case i
				Case 0
					For x = 0 To 3
						If CellValues(EquipmentOne, x, WhichSegment).Value <> 0 Then
							Select Case x
								Case 0
									LabReplaceTitles(i).Text = "Front-End Loader"
								Case 1
									LabReplaceTitles(i).Text = "Hydraulic Shovel"
								Case 2
									LabReplaceTitles(i).Text = "Mechanical Shovel"
								Case 3
									LabReplaceTitles(i).Text = "Walking Dragline"
							End Select
						End If
					Next x
				Case 1
					If CellValues(EquipmentOne, 4, WhichSegment).Value <> 0 Then
						LabReplaceTitles(i).Text = "Rear-Dump Truck"
					ElseIf CellValues(EquipmentOne, 21, WhichSegment).Value <> 0 Then 
						LabReplaceTitles(i).Text = "Acticulated Hauler"
					End If
				Case 2
					For x = 5 To 8
						If CellValues(EquipmentOne, x, WhichSegment).Value <> 0 Then
							Select Case x
								Case 5
									LabReplaceTitles(i).Text = "Front-End Loader"
								Case 6
									LabReplaceTitles(i).Text = "Hydraulic Shovel"
								Case 7
									LabReplaceTitles(i).Text = "Mechanical Shovel"
								Case 8
									LabReplaceTitles(i).Text = "Walking Dragline"
							End Select
						End If
					Next x
				Case 3
					If CellValues(EquipmentOne, 9, WhichSegment).Value <> 0 Then
						LabReplaceTitles(i).Text = "Rear-Dump Truck"
					ElseIf CellValues(EquipmentOne, 26, WhichSegment).Value <> 0 Then 
						LabReplaceTitles(i).Text = "Acticulated Hauler"
					End If
			End Select
			If CellValues(WhichScreen, i, WhichSegment).Value <> 0 Then
				txtReplaceValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value * (CellValues(WhichScreen, 21, WhichSegment).Value + 1))), "###,###,##0.0")
			End If
		Next i
		
		If bltp = 1 Then
			LabReplaceTitles(10).Text = "Powder Buggies"
		ElseIf bltp = 2 Then 
			LabReplaceTitles(10).Text = "Bulk Trucks"
		End If
		
		txtReplaceValues(21).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, 21, WhichSegment).Value)), "##0")
		
		DoNotChange = False
	End Sub
	Private Sub txtReplaceValues_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtReplaceValues.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Dim Index As Short = txtReplaceValues.GetIndex(eventSender)
		
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
						If InStr(txtReplaceValues(Index).Text, ".") = 0 Then
							System.Windows.Forms.SendKeys.Send("{DELETE}")
						End If
					Else
						System.Windows.Forms.SendKeys.Send("{DELETE}")
					End If
			End Select
		End If
		
	End Sub
	Private Sub txtReplaceValues_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtReplaceValues.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = txtReplaceValues.GetIndex(eventSender)
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
	Private Sub txtReplaceValues_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtReplaceValues.Leave
		Dim Index As Short = txtReplaceValues.GetIndex(eventSender)
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
			Case 0 To 19, 21
				tempvalue = ""
				For i = 1 To Len(txtReplaceValues(Sample).Text)
					Digit.Value = Mid(txtReplaceValues(Sample).Text, i, 1)
					Select Case Digit.Value
						Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ".", "-"
							tempvalue = tempvalue & Digit.Value
					End Select
				Next i
				If Sample < 20 Then
					If CellValues(WhichScreen, Sample, WhichSegment).Changed = True Then
						CellValues(WhichScreen, Sample, WhichSegment).Value = (Val(tempvalue) / (CellValues(WhichScreen, 21, WhichSegment).Value + 1))
					End If
				Else
					If CellValues(WhichScreen, Sample, WhichSegment).Changed = True Then
						CellValues(WhichScreen, Sample, WhichSegment).Value = Val(tempvalue)
					End If
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