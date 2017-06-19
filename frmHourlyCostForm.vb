Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports Microsoft.VisualBasic.PowerPacks
Friend Class frmHourlyCostForm
	Inherits System.Windows.Forms.Form
	Dim TempHigh As Single
	Dim TempWide As Single
	Dim TopItOff As Short
	Dim StartYear As Short
	Private Sub comHourlyCostHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles comHourlyCostHelp.Click
		Me.PrintForm1.Print(Me, PowerPacks.Printing.PrintForm.PrintOption.CompatibleModeClientAreaOnly)
	End Sub
	'UPGRADE_WARNING: Form event frmHourlyCostForm.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmHourlyCostForm_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        'Dim baseunit As String
        'Dim baselength As Short
        'Dim i As Short
        If IsHelpOn = True Then
			IsHelpOn = False
		Else
			DoNotChange = True
			WhichScreen = LaborResult
			labProjectName.Text = ProjectTitle(0)
			scrYear.Minimum = MinTime
			scrYear.Maximum = (MaxTime + scrYear.LargeChange - 1)
			scrYear.Value = MinTime
			StartYear = MinTime
			labYear(1).Text = Str(StartYear)
			DoNotChange = False
			WhichCell = 0
		End If
	End Sub
	Private Sub frmHourlyCostForm_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        'Dim i As Short
        'Dim x As Short
        'Dim TimeLine As Short
        Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - (VB6.PixelsToTwipsY(Me.Height) + 350)) / 2)
		Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2)
		If VB6.PixelsToTwipsY(Me.Top) < 0 Then Me.Top = 0
		If VB6.PixelsToTwipsX(Me.Left) < 0 Then Me.Left = 0
		TempWide = VB6.PixelsToTwipsX(Me.ClientRectangle.Width)
		TempHigh = VB6.PixelsToTwipsY(Me.ClientRectangle.Height)
		DoNotChange = True
		scrYear.Minimum = MinTime
		scrYear.Maximum = (MaxTime + scrYear.LargeChange - 1)
		scrYear.Value = MinTime
		StartYear = MinTime
		labYear(1).Text = Str(StartYear)
		Call LoadHourly()
		DoNotChange = False
		Call screenstuff()
	End Sub
	'UPGRADE_WARNING: Event frmHourlyCostForm.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub frmHourlyCostForm_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		TempWide = VB6.PixelsToTwipsX(Me.ClientRectangle.Width)
		TempHigh = VB6.PixelsToTwipsY(Me.ClientRectangle.Height)
		Call screenstuff()
	End Sub
	Private Sub imgBackToMenu_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles imgBackToMenu.Click
		Me.Close()
		Call InputMenuAccess(4)
	End Sub
	Private Sub labBackToMenu_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles labBackToMenu.Click
		Me.Close()
		Call InputMenuAccess(4)
	End Sub
	Public Sub screenstuff()
		Dim x As Short
		Dim y As Decimal
		Dim w As Short
		Dim h As Short
		Dim u As Single
		Dim z As Short
		
		w = 9150 'Starting Form Scale Width
		h = 6420 'Starting Form Scale Height
		u = (360 / w) * TempWide
		y = (2040 / h) * TempHigh
		z = (240 / h) * TempHigh
		
		labHourlyCostHeading.Top = VB6.TwipsToPixelsY(TempHigh * 0.0093)
		labHourlyCostHeading.Left = VB6.TwipsToPixelsX(TempWide * 0.0198)
		labProjectName.Top = VB6.TwipsToPixelsY(TempHigh * 0.0187)
		labProjectName.Left = VB6.TwipsToPixelsX((5280 / w) * TempWide)
		labProjectName.Width = VB6.TwipsToPixelsX(TempWide * 0.2967)
		LineLeft.X1 = VB6.TwipsToPixelsX(TempWide * 0.0131)
		LineLeft.X2 = VB6.TwipsToPixelsX(TempWide * 0.0131)
		LineLeft.Y1 = VB6.TwipsToPixelsY(TempHigh * 0.0748)
		LineLeft.Y2 = VB6.TwipsToPixelsY(TempHigh * 0.9439)
		LineTop.X1 = VB6.TwipsToPixelsX(TempWide * 0.0066)
		LineTop.X2 = VB6.TwipsToPixelsX(TempWide * 0.9902)
		LineTop.Y1 = VB6.TwipsToPixelsY(TempHigh * 0.0841)
		LineTop.Y2 = VB6.TwipsToPixelsY(TempHigh * 0.0841)
		LineBottom.X1 = VB6.TwipsToPixelsX(TempWide * 0.0066)
		LineBottom.X2 = VB6.TwipsToPixelsX(TempWide * 0.9902)
		LineBottom.Y1 = VB6.TwipsToPixelsY(TempHigh * 0.9346)
		LineBottom.Y2 = VB6.TwipsToPixelsY(TempHigh * 0.9346)
		LineRight.X1 = VB6.TwipsToPixelsX(TempWide * 0.9836)
		LineRight.X2 = VB6.TwipsToPixelsX(TempWide * 0.9836)
		LineRight.Y1 = VB6.TwipsToPixelsY(TempHigh * 0.0748)
		LineRight.Y2 = VB6.TwipsToPixelsY(TempHigh * 0.9439)
		grdHourlyCost.Top = VB6.TwipsToPixelsY(TempHigh * (720 / h))
		grdHourlyCost.Left = VB6.TwipsToPixelsX(TempWide * (480 / w))
		grdHourlyCost.Height = VB6.TwipsToPixelsY(TempHigh * (4395 / h))
		grdHourlyCost.Width = VB6.TwipsToPixelsX(TempWide * (8235 / w))
		labBackToMenu.Top = VB6.TwipsToPixelsY(TempHigh * 0.9532)
		labBackToMenu.Left = VB6.TwipsToPixelsX(TempWide * 0.0656)
		imgBackToMenu.Top = VB6.TwipsToPixelsY(TempHigh * 0.9609)
		imgBackToMenu.Left = VB6.TwipsToPixelsX(TempWide * 0.0066)
		imgBackToMenu.Width = VB6.TwipsToPixelsX(TempWide * 0.0541)
		labHourlyCostHelp.Top = VB6.TwipsToPixelsY(TempHigh * 0.9532)
		labHourlyCostHelp.Left = VB6.TwipsToPixelsX(TempWide * 0.9115)
		labHourlyCostHelp.Width = VB6.TwipsToPixelsX(TempWide * 0.0541)
		comHourlyCostHelp.Top = VB6.TwipsToPixelsY(TempHigh * 0.9592)
		comHourlyCostHelp.Left = VB6.TwipsToPixelsX(TempWide * 0.9705)
		scrYear.Top = VB6.TwipsToPixelsY(TempHigh * (5760 / h))
		scrYear.Height = VB6.TwipsToPixelsY(TempHigh * (195 / h))
		scrYear.Left = VB6.TwipsToPixelsX(TempWide * (1500 / w))
		scrYear.Width = VB6.TwipsToPixelsX(TempWide * (6150 / w))
		For x = 0 To 1
			labYear(x).Top = VB6.TwipsToPixelsY(TempHigh * (5280 / h) + (x * z))
			labYear(x).Left = VB6.TwipsToPixelsX(TempWide * (3840 / w))
			labYear(x).Width = VB6.TwipsToPixelsX(TempWide * (1455 / w))
		Next x
	End Sub
	Private Sub LoadHourly()
		Dim x As Short
		Dim y As Short
		Dim r As Short
		Dim TimeLine As Short
		Dim DaysPerYear As Short
		Dim TonsPerDay As Decimal
		Dim DisplayValue(26, 26, 4) As String
		Dim TempTotal(4) As Decimal
		
		On Error Resume Next
		
        grdHourlyCost.set_Cols(0, 4)
		
		For y = 0 To 3
			For x = 0 To 14
				grdHourlyCost.Row = x
				grdHourlyCost.Col = y
				grdHourlyCost.Text = ""
			Next x
		Next y
		
		For x = 0 To 14
			Select Case LCase(VB.Left(CellValues(LaborResult, x, MinTime).Word, 1))
				Case "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z"
					grdHourlyCost.Row = x + 1
					grdHourlyCost.Col = 0
                    grdHourlyCost.set_ColWidth(0, 0, 2722)
					grdHourlyCost.Text = CellValues(LaborResult, x, MinTime).Word
					TopItOff = x + 2
			End Select
		Next x
		
		grdHourlyCost.Row = TopItOff + 1
		grdHourlyCost.Col = 0
        grdHourlyCost.set_ColWidth(0, 0, 2722)
		grdHourlyCost.Text = "Total"
		
		For TimeLine = 1 To MaxTime + 1
			For y = 0 To 6
				If (TimeLine >= CellValues(Production, 15, y).Value And TimeLine <= CellValues(Production, 16, y).Value) Then
					DaysPerYear = CellValues(Production, 3, y).Value
					TonsPerDay = CellValues(Production, 5, y).Value
				End If
			Next y
			For r = 0 To 4
				TempTotal(r) = 0
			Next r
			For x = 1 To TopItOff
				If UnitType = Metric Then
					If CellValues(Production, 20, TimeLine).Value > 0 Then
						If TonsPerDay <> 0 Then
							DisplayValue(TimeLine, x, 0) = VB6.Format(Str(((CellValues(LaborResult, x - 1, TimeLine).Value) / TonsPerDay) / 0.9072), "$##,##0.00")
						End If
					Else
						DisplayValue(TimeLine, x, 0) = VB6.Format(Str(0), "$##,##0.00")
					End If
				Else
					If CellValues(Production, 20, TimeLine).Value > 0 Then
						If TonsPerDay <> 0 Then
							DisplayValue(TimeLine, x, 0) = VB6.Format(Str((CellValues(LaborResult, x - 1, TimeLine).Value) / TonsPerDay), "$##,##0.00")
						End If
					Else
						DisplayValue(TimeLine, x, 0) = VB6.Format(Str(0), "$##,##0.00")
					End If
				End If
				TempTotal(0) = (TempTotal(0) + Val(VB6.Format(DisplayValue(TimeLine, x, 0), "##############.##")))
				DisplayValue(TimeLine, x, 1) = VB6.Format(Str(CellValues(LaborResult, x - 1, TimeLine).Value), "$##,###.#0")
				TempTotal(1) = (TempTotal(1) + Val(VB6.Format(DisplayValue(TimeLine, x, 1), "##############.##")))
				DisplayValue(TimeLine, x, 2) = VB6.Format(Str(CellValues(LaborResult, x - 1, TimeLine).Value * DaysPerYear), "$##,###,###,###")
				TempTotal(2) = (TempTotal(2) + Val(VB6.Format(DisplayValue(TimeLine, x, 2), "##############.##")))
			Next x
			DisplayValue(TimeLine, TopItOff, 0) = VB6.Format(Str(TempTotal(0)), "$###,##0.00")
			DisplayValue(TimeLine, TopItOff, 1) = VB6.Format(Str(TempTotal(1)), "$###,##0.00")
			DisplayValue(TimeLine, TopItOff, 2) = VB6.Format(Str(TempTotal(2)), "$##,###,###,###")
		Next TimeLine
		
		r = 0
		For TimeLine = 0 To 2
			r = r + 1
			For x = 0 To TopItOff
				If x = TopItOff Then
					grdHourlyCost.Row = x + 1
				Else
					grdHourlyCost.Row = x
				End If
				grdHourlyCost.Col = r
                grdHourlyCost.set_ColWidth(r, 0, 1820)
				grdHourlyCost.CellAlignment = 4
				If x = 0 Then
					If TimeLine = 0 Then
						If UnitType = Metric Then
							grdHourlyCost.Text = "Dollars/Metric Ton Ore"
						Else
							grdHourlyCost.Text = "Dollars/Ton Ore"
						End If
					ElseIf TimeLine = 1 Then 
						grdHourlyCost.Text = "Dollars/Day"
					Else
						grdHourlyCost.Text = "Dollars/Year"
					End If
				Else
					grdHourlyCost.CellAlignment = 7
					grdHourlyCost.Text = DisplayValue(StartYear, x, TimeLine)
				End If
			Next x
		Next TimeLine
		
	End Sub
	'UPGRADE_NOTE: scrYear.Change was changed from an event to a procedure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="4E2DC008-5EDA-4547-8317-C9316952674F"'
	'UPGRADE_WARNING: HScrollBar event scrYear.Change has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub scrYear_Change(ByVal newScrollValue As Integer)
		DoNotChange = True
		StartYear = newScrollValue
		labYear(1).Text = Str(StartYear)
		grdHourlyCost.Visible = False
		Call LoadHourly()
		grdHourlyCost.Visible = True
		DoNotChange = False
	End Sub
	Private Sub scrYear_Scroll(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.ScrollEventArgs) Handles scrYear.Scroll
		Select Case eventArgs.type
			Case System.Windows.Forms.ScrollEventType.EndScroll
				scrYear_Change(eventArgs.newValue)
		End Select
	End Sub
End Class