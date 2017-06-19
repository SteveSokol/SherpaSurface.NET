Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks
Friend Class frmHaulData
	Inherits System.Windows.Forms.Form
	Dim TempHigh As Single
	Dim TempWide As Single
	Dim FootConversion As Single
	Private Sub comHaulPrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles comHaulPrint.Click
		PrintForm1.Print(Me, PowerPacks.Printing.PrintForm.PrintOption.CompatibleModeClientAreaOnly)
	End Sub
	'UPGRADE_WARNING: Form event frmHaulData.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmHaulData_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated

        'Dim baseunit As String
        'Dim baselength As Short
        Dim i As Short
		
		If IsHelpOn = True Then
			txtHaulValues(WhichCell).Focus()
			IsHelpOn = False
		Else
			For i = 0 To 5
				If optSegment(i).Checked = True Then WhichSegment = i
			Next i
			
			WhichScreen = Haul
			
			Call drawthevalues()
			
			If InsertFlag = True Then
				labInsert.Text = "Insert"
			Else
				labInsert.Text = "Typeover"
			End If
			
			WhichCell = 0
			
			txtHaulValues(0).Focus()
		End If
		
	End Sub
	Private Sub frmHaulData_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim x As Short
		
		Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - (VB6.PixelsToTwipsY(Me.Height) + 350)) / 2)
		Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2)
		
		If VB6.PixelsToTwipsY(Me.Top) < 0 Then Me.Top = 0
		If VB6.PixelsToTwipsX(Me.Left) < 0 Then Me.Left = 0
		
		TempWide = VB6.PixelsToTwipsX(Me.ClientRectangle.Width)
		TempHigh = VB6.PixelsToTwipsY(Me.ClientRectangle.Height)
		
		DoNotChange = True
		
		If UnitType = Metric Then
			For x = 0 To 22 Step 2
				labHaulUnits(x).Text = "meters"
			Next x
			FootConversion = 0.3048
		Else
			For x = 0 To 22 Step 2
				labHaulUnits(x).Text = "feet"
			Next x
			FootConversion = 1
		End If
		
		For x = 12 To 19
			LabHaulTitles(x).Enabled = True
			labHaulUnits(x + 12).Enabled = True
		Next x
		
		For x = 16 To 19
			txtHaulValues(x).Enabled = True
			txtHaulValues(x + 12).Enabled = True
		Next x
		
		LabHaulTitles(0).Text = "Haul Segment #1"
		LabHaulTitles(1).Text = "Haul Segment #2"
		LabHaulTitles(2).Text = "Haul Segment #3"
		LabHaulTitles(3).Text = "Haul Segment #4"
		LabHaulTitles(4).Text = "Haul Segment #5"
		LabHaulTitles(5).Text = "Haul Segment #6"
		LabHaulTitles(6).Text = "Haul Segment #1"
		LabHaulTitles(7).Text = "Haul Segment #2"
		LabHaulTitles(8).Text = "Haul Segment #3"
		LabHaulTitles(9).Text = "Haul Segment #4"
		LabHaulTitles(10).Text = "Haul Segment #5"
		LabHaulTitles(11).Text = "Haul Segment #6"
		
		If LTrim(RTrim(LCase(CellValues(Production, 9, WhichSegment).Word))) = "crusher/conveyor" Then
			LabHaulTitles(0).Text = "Access Segment #1"
			LabHaulTitles(1).Text = "Access Segment #2"
			LabHaulTitles(2).Text = "AccessSegment #3"
			LabHaulTitles(3).Text = "Access Segment #4"
			LabHaulTitles(4).Text = "Access Segment #5"
			LabHaulTitles(5).Text = "Access Segment #6"
			LabHaulTitles(14).Enabled = False
			LabHaulTitles(15).Enabled = False
			txtHaulValues(18).Enabled = False
			txtHaulValues(19).Enabled = False
			labHaulUnits(26).Enabled = False
			labHaulUnits(27).Enabled = False
		End If
		
		If LTrim(RTrim(LCase(CellValues(Production, 14, WhichSegment).Word))) = "crusher/conveyor" Or LTrim(RTrim(LCase(CellValues(Production, 13, WhichSegment).Word))) = "walking dragline" Then
			LabHaulTitles(6).Text = "Access Segment #1"
			LabHaulTitles(7).Text = "Access Segment #2"
			LabHaulTitles(8).Text = "Access Segment #3"
			LabHaulTitles(9).Text = "Access Segment #4"
			LabHaulTitles(10).Text = "Access Segment #5"
			LabHaulTitles(11).Text = "Access Segment #6"
			LabHaulTitles(18).Enabled = False
			LabHaulTitles(19).Enabled = False
			txtHaulValues(30).Enabled = False
			txtHaulValues(31).Enabled = False
			labHaulUnits(30).Enabled = False
			labHaulUnits(31).Enabled = False
		End If
		
		If LTrim(RTrim(LCase(CellValues(Production, 8, WhichSegment).Word))) = "scraper" Then
			LabHaulTitles(12).Enabled = False
			LabHaulTitles(13).Enabled = False
			txtHaulValues(16).Enabled = False
			txtHaulValues(17).Enabled = False
			labHaulUnits(24).Enabled = False
			labHaulUnits(25).Enabled = False
		End If
		
		If LTrim(RTrim(LCase(CellValues(Production, 13, WhichSegment).Word))) = "scraper" Then
			LabHaulTitles(16).Enabled = False
			LabHaulTitles(17).Enabled = False
			txtHaulValues(28).Enabled = False
			txtHaulValues(29).Enabled = False
			labHaulUnits(28).Enabled = False
			labHaulUnits(29).Enabled = False
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
	'UPGRADE_WARNING: Event frmHaulData.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub frmHaulData_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		TempWide = VB6.PixelsToTwipsX(Me.ClientRectangle.Width)
		TempHigh = VB6.PixelsToTwipsY(Me.ClientRectangle.Height)
		Call screenstuff()
	End Sub
	Private Sub frmHaulData_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
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
	Private Sub labHaulHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles labHaulHelp.Click
		Dim StartHelp As Short
		Dim CellHelp As Short
		StartHelp = 42
		Select Case WhichCell
			Case 0 To 3
				CellHelp = WhichCell
			Case 4 To 7, 16 To 19
				CellHelp = WhichCell + 8
			Case 8 To 15, 20 To 27
				CellHelp = WhichCell - 4
			Case 28 To 31
				CellHelp = WhichCell - 4
		End Select
		IsHelpOn = True
		Call frmSurfaceHelp.gethelptext(StartHelp, CellHelp)
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
			txtHaulValues(WhichCell).Focus()
		End If
	End Sub
	'UPGRADE_WARNING: Event txtHaulValues.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtHaulValues_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtHaulValues.TextChanged
		Dim Index As Short = txtHaulValues.GetIndex(eventSender)
		If DoNotChange = True Then Exit Sub
	End Sub
	Private Sub txtHaulValues_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtHaulValues.Enter
		Dim Index As Short = txtHaulValues.GetIndex(eventSender)
        'Dim x As Short
        WhichCell = Index
		System.Windows.Forms.SendKeys.Send("{HOME}+{END}")
		WhichCell = Index
		Call drawthevalues()
	End Sub
	
	Public Sub screenstuff()
		
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
		
		r = (1920 / h) * TempHigh
		s = (2700 / h) * TempHigh
		t = (1020 / h) * TempHigh
		
		y = (420 / h) * TempHigh
		z = (300 / w) * TempWide
		
		labHaulHeading.Top = VB6.TwipsToPixelsY(TempHigh * (180 / h))
		labHaulHeading.Left = VB6.TwipsToPixelsX(TempWide * (180 / w))
		
		labHaulLabels(0).Top = VB6.TwipsToPixelsY(TempHigh * (900 / h))
		labHaulLabels(0).Left = VB6.TwipsToPixelsX(TempWide * (600 / w))
		labHaulLabels(0).Width = VB6.TwipsToPixelsX(TempWide * (2655 / w))
		
		labHaulLabels(1).Top = VB6.TwipsToPixelsY(TempHigh * (5720 / h))
		labHaulLabels(1).Left = VB6.TwipsToPixelsX(TempWide * (5340 / w))
		labHaulLabels(1).Width = VB6.TwipsToPixelsX(TempWide * (2655 / w))
		
		For x = 0 To 1
			labHaulLabels(x + 2).Top = VB6.TwipsToPixelsY((TempHigh * (1920 / h)) + (x * r))
			labHaulLabels(x + 2).Left = VB6.TwipsToPixelsX(TempWide * (120 / w))
			labHaulLabels(x + 4).Top = VB6.TwipsToPixelsY((TempHigh * (60 / h)) + (x * s))
			labHaulLabels(x + 4).Left = VB6.TwipsToPixelsX(TempWide * (3780 / w))
		Next x
		
		For x = 0 To 5
			optSegment(x).Top = VB6.TwipsToPixelsY(TempHigh * (1200 / h))
			optSegment(x).Left = VB6.TwipsToPixelsX((TempWide * (1080 / w)) + (x * z))
			optSegment(x).Width = VB6.TwipsToPixelsX(TempWide * (195 / w))
			labSegment(x).Top = VB6.TwipsToPixelsY(TempHigh * (1440 / h))
			labSegment(x).Left = VB6.TwipsToPixelsX((TempWide * (1080 / w)) + (x * z))
			labSegment(x).Width = VB6.TwipsToPixelsX(TempWide * (195 / w))
		Next x
		
		txtSegmentLabel.Top = VB6.TwipsToPixelsY(TempHigh * (6020 / h))
		txtSegmentLabel.Left = VB6.TwipsToPixelsX(TempWide * (5700 / w))
		txtSegmentLabel.Width = VB6.TwipsToPixelsX(TempWide * (1935 / w))
		
		For x = 0 To 5
			LabHaulTitles(x).Left = VB6.TwipsToPixelsX(TempWide * (3900 / w))
			LabHaulTitles(x).Top = VB6.TwipsToPixelsY((TempHigh * (400 / h)) + (x * y))
			LabHaulTitles(x).Width = VB6.TwipsToPixelsX(TempWide * (1695 / w))
			LabHaulTitles(x + 6).Left = VB6.TwipsToPixelsX(TempWide * (3900 / w))
			LabHaulTitles(x + 6).Top = VB6.TwipsToPixelsY((TempHigh * (3160 / h)) + (x * y))
			LabHaulTitles(x + 6).Width = VB6.TwipsToPixelsX(TempWide * (1695 / w))
			labHaulUnits(x * 2).Top = VB6.TwipsToPixelsY((TempHigh * (390 / h)) + (x * y))
			labHaulUnits((x * 2) + 1).Top = VB6.TwipsToPixelsY((TempHigh * (390 / h)) + (x * y))
			labHaulUnits(x * 2).Left = VB6.TwipsToPixelsX(TempWide * (6660 / w))
			labHaulUnits((x * 2) + 1).Left = VB6.TwipsToPixelsX(TempWide * (8280 / w))
			labHaulUnits((x * 2) + 12).Top = VB6.TwipsToPixelsY((TempHigh * (3150 / h)) + (x * y))
			labHaulUnits((x * 2) + 13).Top = VB6.TwipsToPixelsY((TempHigh * (3150 / h)) + (x * y))
			labHaulUnits((x * 2) + 12).Left = VB6.TwipsToPixelsX(TempWide * (6660 / w))
			labHaulUnits((x * 2) + 13).Left = VB6.TwipsToPixelsX(TempWide * (8280 / w))
		Next x
		
		For x = 0 To 3
			LabHaulTitles(x + 12).Left = VB6.TwipsToPixelsX(TempWide * (300 / w))
			LabHaulTitles(x + 12).Top = VB6.TwipsToPixelsY((TempHigh * (2260 / h)) + (x * y))
			LabHaulTitles(x + 16).Left = VB6.TwipsToPixelsX(TempWide * (300 / w))
			LabHaulTitles(x + 16).Top = VB6.TwipsToPixelsY((TempHigh * (4180 / h)) + (x * y))
			labHaulUnits(x + 24).Top = VB6.TwipsToPixelsY((TempHigh * (2250 / h)) + (x * y))
			labHaulUnits(x + 24).Left = VB6.TwipsToPixelsX(TempWide * (3000 / w))
			labHaulUnits(x + 28).Top = VB6.TwipsToPixelsY((TempHigh * (4170 / h)) + (x * y))
			labHaulUnits(x + 28).Left = VB6.TwipsToPixelsX(TempWide * (3000 / w))
			txtHaulValues(x + 16).Top = VB6.TwipsToPixelsY(TempHigh * (2220 / h) + (x * y))
			txtHaulValues(x + 16).Left = VB6.TwipsToPixelsX(TempWide * (2220 / w))
			txtHaulValues(x + 16).Width = VB6.TwipsToPixelsX(TempWide * (735 / w))
			txtHaulValues(x + 28).Top = VB6.TwipsToPixelsY(TempHigh * (4140 / h) + (x * y))
			txtHaulValues(x + 28).Left = VB6.TwipsToPixelsX(TempWide * (2220 / w))
			txtHaulValues(x + 28).Width = VB6.TwipsToPixelsX(TempWide * (735 / w))
		Next x
		
		For x = 0 To 2 Step 2
			txtHaulValues(x).Top = VB6.TwipsToPixelsY((TempHigh * (360 / h)) + ((x / 2) * y))
			txtHaulValues(x).Left = VB6.TwipsToPixelsX(TempWide * (5760 / w))
			txtHaulValues(x).Width = VB6.TwipsToPixelsX(TempWide * (855 / w))
			txtHaulValues(x + 1).Top = VB6.TwipsToPixelsY((TempHigh * (360 / h)) + ((x / 2) * y))
			txtHaulValues(x + 1).Left = VB6.TwipsToPixelsX(TempWide * (7500 / w))
			txtHaulValues(x + 1).Width = VB6.TwipsToPixelsX(TempWide * (735 / w))
			txtHaulValues(x + 4).Top = VB6.TwipsToPixelsY((TempHigh * (3120 / h)) + ((x / 2) * y))
			txtHaulValues(x + 4).Left = VB6.TwipsToPixelsX(TempWide * (5760 / w))
			txtHaulValues(x + 4).Width = VB6.TwipsToPixelsX(TempWide * (855 / w))
			txtHaulValues(x + 5).Top = VB6.TwipsToPixelsY((TempHigh * (3120 / h)) + ((x / 2) * y))
			txtHaulValues(x + 5).Left = VB6.TwipsToPixelsX(TempWide * (7500 / w))
			txtHaulValues(x + 5).Width = VB6.TwipsToPixelsX(TempWide * (735 / w))
		Next x
		
		For x = 0 To 6 Step 2
			txtHaulValues(x + 8).Top = VB6.TwipsToPixelsY((TempHigh * (1200 / h)) + ((x / 2) * y))
			txtHaulValues(x + 8).Left = VB6.TwipsToPixelsX(TempWide * (5760 / w))
			txtHaulValues(x + 8).Width = VB6.TwipsToPixelsX(TempWide * (855 / w))
			txtHaulValues(x + 9).Top = VB6.TwipsToPixelsY((TempHigh * (1200 / h)) + ((x / 2) * y))
			txtHaulValues(x + 9).Left = VB6.TwipsToPixelsX(TempWide * (7500 / w))
			txtHaulValues(x + 9).Width = VB6.TwipsToPixelsX(TempWide * (735 / w))
			txtHaulValues(x + 20).Top = VB6.TwipsToPixelsY((TempHigh * (3960 / h)) + ((x / 2) * y))
			txtHaulValues(x + 20).Left = VB6.TwipsToPixelsX(TempWide * (5760 / w))
			txtHaulValues(x + 20).Width = VB6.TwipsToPixelsX(TempWide * (855 / w))
			txtHaulValues(x + 21).Top = VB6.TwipsToPixelsY((TempHigh * (3960 / h)) + ((x / 2) * y))
			txtHaulValues(x + 21).Left = VB6.TwipsToPixelsX(TempWide * (7500 / w))
			txtHaulValues(x + 21).Width = VB6.TwipsToPixelsX(TempWide * (735 / w))
		Next x
		
		LineHorizontal(0).X1 = VB6.TwipsToPixelsX(TempWide * (540 / w))
		LineHorizontal(0).X2 = VB6.TwipsToPixelsX(TempWide * (3780 / w))
		LineHorizontal(0).Y1 = VB6.TwipsToPixelsY(TempHigh * (1980 / h))
		LineHorizontal(0).Y2 = VB6.TwipsToPixelsY(TempHigh * (1980 / h))
		
		LineHorizontal(1).X1 = VB6.TwipsToPixelsX(TempWide * (780 / w))
		LineHorizontal(1).X2 = VB6.TwipsToPixelsX(TempWide * (3780 / w))
		LineHorizontal(1).Y1 = VB6.TwipsToPixelsY(TempHigh * (3900 / h))
		LineHorizontal(1).Y2 = VB6.TwipsToPixelsY(TempHigh * (3900 / h))
		
		LineHorizontal(2).X1 = VB6.TwipsToPixelsX(TempWide * (120 / w))
		LineHorizontal(2).X2 = VB6.TwipsToPixelsX(TempWide * (3900 / w))
		LineHorizontal(2).Y1 = VB6.TwipsToPixelsY(TempHigh * (5820 / h))
		LineHorizontal(2).Y2 = VB6.TwipsToPixelsY(TempHigh * (5820 / h))
		
		LineHorizontal(3).X1 = VB6.TwipsToPixelsX(TempWide * (4200 / w))
		LineHorizontal(3).X2 = VB6.TwipsToPixelsX(TempWide * (9060 / w))
		LineHorizontal(3).Y1 = VB6.TwipsToPixelsY(TempHigh * (120 / h))
		LineHorizontal(3).Y2 = VB6.TwipsToPixelsY(TempHigh * (120 / h))
		
		LineHorizontal(4).X1 = VB6.TwipsToPixelsX(TempWide * (4440 / w))
		LineHorizontal(4).X2 = VB6.TwipsToPixelsX(TempWide * (8940 / w))
		LineHorizontal(4).Y1 = VB6.TwipsToPixelsY(TempHigh * (2880 / h))
		LineHorizontal(4).Y2 = VB6.TwipsToPixelsY(TempHigh * (2880 / h))
		
		LineHorizontal(5).X1 = VB6.TwipsToPixelsX(TempWide * (3900 / w))
		LineHorizontal(5).X2 = VB6.TwipsToPixelsX(TempWide * (9060 / w))
		LineHorizontal(5).Y1 = VB6.TwipsToPixelsY(TempHigh * (5640 / h))
		LineHorizontal(5).Y2 = VB6.TwipsToPixelsY(TempHigh * (5640 / h))
		
		LineVertical(0).X1 = VB6.TwipsToPixelsX(TempWide * (180 / w))
		LineVertical(0).X2 = VB6.TwipsToPixelsX(TempWide * (180 / w))
		LineVertical(0).Y1 = VB6.TwipsToPixelsY(TempHigh * (2220 / h))
		LineVertical(0).Y2 = VB6.TwipsToPixelsY(TempHigh * (3780 / h))
		
		LineVertical(1).X1 = VB6.TwipsToPixelsX(TempWide * (180 / w))
		LineVertical(1).X2 = VB6.TwipsToPixelsX(TempWide * (180 / w))
		LineVertical(1).Y1 = VB6.TwipsToPixelsY(TempHigh * (4140 / h))
		LineVertical(1).Y2 = VB6.TwipsToPixelsY(TempHigh * (5880 / h))
		
		LineVertical(2).X1 = VB6.TwipsToPixelsX(TempWide * (3840 / w))
		LineVertical(2).X2 = VB6.TwipsToPixelsX(TempWide * (3840 / w))
		LineVertical(2).Y1 = VB6.TwipsToPixelsY(TempHigh * (360 / h))
		LineVertical(2).Y2 = VB6.TwipsToPixelsY(TempHigh * (2700 / h))
		
		LineVertical(3).X1 = VB6.TwipsToPixelsX(TempWide * (3840 / w))
		LineVertical(3).X2 = VB6.TwipsToPixelsX(TempWide * (3840 / w))
		LineVertical(3).Y1 = VB6.TwipsToPixelsY(TempHigh * (3060 / h))
		LineVertical(3).Y2 = VB6.TwipsToPixelsY(TempHigh * (5880 / h))
		
		LineVertical(4).X1 = VB6.TwipsToPixelsX(TempWide * (9000 / w))
		LineVertical(4).X2 = VB6.TwipsToPixelsX(TempWide * (9000 / w))
		LineVertical(4).Y1 = VB6.TwipsToPixelsY(TempHigh * (60 / h))
		LineVertical(4).Y2 = VB6.TwipsToPixelsY(TempHigh * (5700 / h))
		
		labBackToMenu.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labBackToMenu.Left = VB6.TwipsToPixelsX(TempWide * (600 / w))
		
		imgBackToMenu.Top = VB6.TwipsToPixelsY(TempHigh * (6150 / h))
		imgBackToMenu.Left = VB6.TwipsToPixelsX(TempWide * (60 / w))
		imgBackToMenu.Width = VB6.TwipsToPixelsX(TempWide * (495 / w))
		
		comHaulPrint.Top = VB6.TwipsToPixelsY(TempHigh * (6150 / h))
		comHaulPrint.Left = VB6.TwipsToPixelsX(TempWide * (8880 / w))
		
		labHaulHelp.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labHaulHelp.Left = VB6.TwipsToPixelsX(TempWide * (8340 / w))
		
		labInsert.Top = VB6.TwipsToPixelsY(TempHigh * (6120 / h))
		labInsert.Left = VB6.TwipsToPixelsX(TempWide * (4080 / w))
		labInsert.Width = VB6.TwipsToPixelsX(TempWide * (975 / w))
		
	End Sub
	Public Sub drawthevalues()
		
		Dim i As Short
        'Dim x As Short

        DoNotChange = True
		
		For i = 0 To 31
			If CellValues(WhichScreen, i, WhichSegment).Changed = True Then
				txtHaulValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF0000)
			Else
				txtHaulValues(i).ForeColor = System.Drawing.ColorTranslator.FromOle(&H80000008)
			End If
			Select Case i
				Case 0, 2, 4, 6, 8, 10, 12, 14, 20, 22, 24, 26
					txtHaulValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value * FootConversion)), "#,###,###,##0")
				Case Else
					txtHaulValues(i).Text = VB6.Format(LTrim(Str(CellValues(WhichScreen, i, WhichSegment).Value)), "##,###,##0.0")
			End Select
		Next i
		
		DoNotChange = False
		
	End Sub
	Private Sub txtHaulValues_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtHaulValues.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		Dim Index As Short = txtHaulValues.GetIndex(eventSender)
		
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
						If InStr(txtHaulValues(Index).Text, ".") = 0 Then
							System.Windows.Forms.SendKeys.Send("{DELETE}")
						End If
					Else
						System.Windows.Forms.SendKeys.Send("{DELETE}")
					End If
			End Select
		End If
		
	End Sub
	Private Sub txtHaulValues_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtHaulValues.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		Dim Index As Short = txtHaulValues.GetIndex(eventSender)
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
	Private Sub txtHaulValues_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtHaulValues.Leave
		Dim Index As Short = txtHaulValues.GetIndex(eventSender)
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
		For i = 1 To Len(txtHaulValues(Sample).Text)
			Digit.Value = Mid(txtHaulValues(Sample).Text, i, 1)
			Select Case Digit.Value
				Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ".", "-"
					tempvalue = tempvalue & Digit.Value
			End Select
		Next i
		Select Case Sample
			Case 0, 2, 4, 6, 8, 10, 12, 14, 20, 22, 24, 26
				If CellValues(WhichScreen, Sample, WhichSegment).Changed = True Then
					If FootConversion <> 0 Then CellValues(WhichScreen, Sample, WhichSegment).Value = Val(tempvalue) / FootConversion
				End If
			Case Else
				If CellValues(WhichScreen, Sample, WhichSegment).Changed = True Then
					CellValues(WhichScreen, Sample, WhichSegment).Value = Val(tempvalue)
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