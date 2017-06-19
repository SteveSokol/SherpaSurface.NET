Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports Microsoft.VisualBasic.PowerPacks
Friend Class frmDevelopmentCostForm
	Inherits System.Windows.Forms.Form
	Dim TempHigh As Single
	Dim TempWide As Single
	Dim TopItOff As Short
	Dim StartYear As Short
	Private Sub comDevelopmentCostHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles comDevelopmentCostHelp.Click
		Me.PrintForm1.Print(Me, PowerPacks.Printing.PrintForm.PrintOption.CompatibleModeClientAreaOnly)
	End Sub
	'UPGRADE_WARNING: Form event frmDevelopmentCostForm.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmDevelopmentCostForm_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        'Dim baseunit As String
        'Dim baselength As Short
        'Dim i As Short
        If IsHelpOn = True Then
			IsHelpOn = False
		Else
			DoNotChange = True
			WhichScreen = DevelopmentResult
			labProjectName.Text = ProjectTitle(0)
			scrYear.Minimum = 1
			scrYear.Maximum = (MaxTime + scrYear.LargeChange - 1)
			scrYear.Value = 1
			StartYear = 1
			labYear(1).Text = Str(StartYear)
			DoNotChange = False
			WhichCell = 0
		End If
	End Sub
	Private Sub frmDevelopmentCostForm_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
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
		scrYear.Minimum = 1
		scrYear.Maximum = (MaxTime + scrYear.LargeChange - 1)
		scrYear.Value = 1
		StartYear = 1
		labYear(1).Text = Str(StartYear)
		Call LoadDevelopment()
		DoNotChange = False
		Call screenstuff()
	End Sub
	'UPGRADE_WARNING: Event frmDevelopmentCostForm.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub frmDevelopmentCostForm_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
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
		z = (1560 / w) * TempWide
		
		labDevelopmentCostHeading.Top = VB6.TwipsToPixelsY(TempHigh * 0.0093)
		labDevelopmentCostHeading.Left = VB6.TwipsToPixelsX(TempWide * 0.0198)
		labProjectName.Top = VB6.TwipsToPixelsY(TempHigh * 0.0187)
		labProjectName.Left = VB6.TwipsToPixelsX((5280 / w) * TempWide)
		labProjectName.Width = VB6.TwipsToPixelsX(TempWide * 0.2967)
		LineLeft.X1 = VB6.TwipsToPixelsX(TempWide * (120 / w))
		LineLeft.X2 = VB6.TwipsToPixelsX(TempWide * (120 / w))
		LineLeft.Y1 = VB6.TwipsToPixelsY(TempHigh * (480 / h))
		LineLeft.Y2 = VB6.TwipsToPixelsY(TempHigh * (5880 / h))
		LineNotQuiteLeft.X1 = VB6.TwipsToPixelsX(TempWide * (1380 / w))
		LineNotQuiteLeft.X2 = VB6.TwipsToPixelsX(TempWide * (1380 / w))
		LineNotQuiteLeft.Y1 = VB6.TwipsToPixelsY(TempHigh * (5760 / h))
		LineNotQuiteLeft.Y2 = VB6.TwipsToPixelsY(TempHigh * (6300 / h))
		LineTop.X1 = VB6.TwipsToPixelsX(TempWide * 0.0066)
		LineTop.X2 = VB6.TwipsToPixelsX(TempWide * 0.9902)
		LineTop.Y1 = VB6.TwipsToPixelsY(TempHigh * 0.0841)
		LineTop.Y2 = VB6.TwipsToPixelsY(TempHigh * 0.0841)
		LineBottomLeft.X1 = VB6.TwipsToPixelsX(TempWide * (60 / w))
		LineBottomLeft.X2 = VB6.TwipsToPixelsX(TempWide * (1440 / w))
		LineBottomLeft.Y1 = VB6.TwipsToPixelsY(TempHigh * (5820 / h))
		LineBottomLeft.Y2 = VB6.TwipsToPixelsY(TempHigh * (5820 / h))
		LineBottom.X1 = VB6.TwipsToPixelsX(TempWide * (1320 / w))
		LineBottom.X2 = VB6.TwipsToPixelsX(TempWide * (7800 / w))
		LineBottom.Y1 = VB6.TwipsToPixelsY(TempHigh * (6240 / h))
		LineBottom.Y2 = VB6.TwipsToPixelsY(TempHigh * (6240 / h))
		LineBottomRight.X1 = VB6.TwipsToPixelsX(TempWide * (7680 / w))
		LineBottomRight.X2 = VB6.TwipsToPixelsX(TempWide * (9060 / w))
		LineBottomRight.Y1 = VB6.TwipsToPixelsY(TempHigh * (5820 / h))
		LineBottomRight.Y2 = VB6.TwipsToPixelsY(TempHigh * (5820 / h))
		LineNotQuiteRight.X1 = VB6.TwipsToPixelsX(TempWide * (7740 / w))
		LineNotQuiteRight.X2 = VB6.TwipsToPixelsX(TempWide * (7740 / w))
		LineNotQuiteRight.Y1 = VB6.TwipsToPixelsY(TempHigh * (5760 / h))
		LineNotQuiteRight.Y2 = VB6.TwipsToPixelsY(TempHigh * (6300 / h))
		LineRight.X1 = VB6.TwipsToPixelsX(TempWide * (9000 / w))
		LineRight.X2 = VB6.TwipsToPixelsX(TempWide * (9000 / w))
		LineRight.Y1 = VB6.TwipsToPixelsY(TempHigh * (480 / h))
		LineRight.Y2 = VB6.TwipsToPixelsY(TempHigh * (5880 / h))
		grdDevelopmentCost.Top = VB6.TwipsToPixelsY(TempHigh * (660 / h))
		grdDevelopmentCost.Left = VB6.TwipsToPixelsX(TempWide * (480 / w))
		grdDevelopmentCost.Height = VB6.TwipsToPixelsY(TempHigh * (4935 / h))
		grdDevelopmentCost.Width = VB6.TwipsToPixelsX(TempWide * (8235 / w))
		labBackToMenu.Top = VB6.TwipsToPixelsY(TempHigh * 0.9532)
		labBackToMenu.Left = VB6.TwipsToPixelsX(TempWide * 0.0656)
		imgBackToMenu.Top = VB6.TwipsToPixelsY(TempHigh * 0.9609)
		imgBackToMenu.Left = VB6.TwipsToPixelsX(TempWide * 0.0066)
		imgBackToMenu.Width = VB6.TwipsToPixelsX(TempWide * 0.0541)
		labDevelopmentCostHelp.Top = VB6.TwipsToPixelsY(TempHigh * 0.9532)
		labDevelopmentCostHelp.Left = VB6.TwipsToPixelsX(TempWide * 0.9115)
		labDevelopmentCostHelp.Width = VB6.TwipsToPixelsX(TempWide * 0.0541)
		comDevelopmentCostHelp.Top = VB6.TwipsToPixelsY(TempHigh * 0.9592)
		comDevelopmentCostHelp.Left = VB6.TwipsToPixelsX(TempWide * 0.9705)
		scrYear.Top = VB6.TwipsToPixelsY(TempHigh * (5940 / h))
		scrYear.Height = VB6.TwipsToPixelsY(TempHigh * (195 / h))
		scrYear.Left = VB6.TwipsToPixelsX(TempWide * (1500 / w))
		scrYear.Width = VB6.TwipsToPixelsX(TempWide * (6150 / w))
		For x = 0 To 1
			labYear(x).Top = VB6.TwipsToPixelsY(TempHigh * (5640 / h))
			labYear(x).Left = VB6.TwipsToPixelsX(TempWide * (3660 / w) + (x * z))
			If x = 0 Then
				labYear(x).Width = VB6.TwipsToPixelsX(TempWide * (1515 / w))
			Else
				labYear(x).Width = VB6.TwipsToPixelsX(TempWide * (255 / w))
			End If
		Next x
	End Sub
	Private Sub LoadDevelopment()
		Dim x As Short
		Dim y As Short
		Dim r As Short
		Dim TimeLine As Short
		Dim DaysPerYear As Short
		Dim TonsPerDay As Decimal
		Dim DisplayValue(26, 26, 4) As String
		Dim TempTotal As Decimal
		
		On Error Resume Next
		
        grdDevelopmentCost.set_Cols(0, 2)
		
		For y = 0 To 1
			For x = 0 To 14
				grdDevelopmentCost.Row = x
				grdDevelopmentCost.Col = y
				grdDevelopmentCost.Text = ""
			Next x
		Next y
		
		For x = 0 To 14
			Select Case LCase(VB.Left(CellValues(DevelopmentResult, x, 1).Word, 1))
				Case "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z"
					grdDevelopmentCost.Row = x + 1
					grdDevelopmentCost.Col = 0
                    grdDevelopmentCost.set_ColWidth(0, 0, 2722)
					grdDevelopmentCost.Text = CellValues(DevelopmentResult, x, 1).Word
					TopItOff = x + 2
			End Select
		Next x
		
		grdDevelopmentCost.Row = TopItOff + 1
		grdDevelopmentCost.Col = 0
        grdDevelopmentCost.set_ColWidth(0, 0, 2722)
		grdDevelopmentCost.Text = "Total"
		
		For TimeLine = 1 To MaxTime
			TempTotal = 0
			For x = 1 To TopItOff
				DisplayValue(TimeLine, x, 0) = VB6.Format(Str(CellValues(DevelopmentResult, x - 1, TimeLine).Value), "$###,###,##0")
				TempTotal = (TempTotal + Val(VB6.Format(DisplayValue(TimeLine, x, 0), "##############")))
			Next x
			DisplayValue(TimeLine, TopItOff, 0) = VB6.Format(Str(TempTotal), "$###,###,##0")
		Next TimeLine
		
		r = 1
		For x = 0 To TopItOff
			If x = TopItOff Then
				grdDevelopmentCost.Row = x + 1
			Else
				grdDevelopmentCost.Row = x
			End If
			grdDevelopmentCost.Col = r
            grdDevelopmentCost.set_ColWidth(r, 0, 5460)
			grdDevelopmentCost.CellAlignment = 4
			If x = 0 Then
				grdDevelopmentCost.Text = "Dollars/Year"
			Else
				grdDevelopmentCost.CellAlignment = 7
				grdDevelopmentCost.Text = DisplayValue(StartYear, x, 0)
			End If
		Next x
	End Sub
	'UPGRADE_NOTE: scrYear.Change was changed from an event to a procedure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="4E2DC008-5EDA-4547-8317-C9316952674F"'
	'UPGRADE_WARNING: HScrollBar event scrYear.Change has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub scrYear_Change(ByVal newScrollValue As Integer)
		DoNotChange = True
		StartYear = newScrollValue
		labYear(1).Text = Str(StartYear)
		grdDevelopmentCost.Visible = False
		Call LoadDevelopment()
		grdDevelopmentCost.Visible = True
		DoNotChange = False
	End Sub
	Private Sub scrYear_Scroll(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.ScrollEventArgs) Handles scrYear.Scroll
		Select Case eventArgs.type
			Case System.Windows.Forms.ScrollEventType.EndScroll
				scrYear_Change(eventArgs.newValue)
		End Select
	End Sub

    Private Sub AxMSHFlexGrid1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
End Class