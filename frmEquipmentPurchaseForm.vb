Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports Microsoft.VisualBasic.PowerPacks
Friend Class frmEquipmentPurchaseForm
	Inherits System.Windows.Forms.Form
	Dim TempHigh As Single
	Dim TempWide As Single
	Dim TonneTon As Decimal
	Private Sub comEquipmentPurchaseHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles comEquipmentPurchaseHelp.Click
		Me.PrintForm1.Print(Me, PowerPacks.Printing.PrintForm.PrintOption.CompatibleModeClientAreaOnly)
	End Sub
	'UPGRADE_WARNING: Form event frmEquipmentPurchaseForm.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmEquipmentPurchaseForm_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        'Dim baseunit As String
        'Dim baselength As Short
        'Dim i As Short
        If IsHelpOn = True Then
			IsHelpOn = False
		Else
			DoNotChange = True
			WhichScreen = EquipmentPurchaseResult
			labProjectName.Text = ProjectTitle(0)
			DoNotChange = False
			WhichCell = 0
		End If
	End Sub
	Private Sub frmEquipmentPurchaseForm_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        'Dim i As Short
        Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - (VB6.PixelsToTwipsY(Me.Height) + 350)) / 2)
		Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2)
		If VB6.PixelsToTwipsY(Me.Top) < 0 Then Me.Top = 0
		If VB6.PixelsToTwipsX(Me.Left) < 0 Then Me.Left = 0
		TempWide = VB6.PixelsToTwipsX(Me.ClientRectangle.Width)
		TempHigh = VB6.PixelsToTwipsY(Me.ClientRectangle.Height)
		DoNotChange = True
		If UnitType = Metric Then
			TonneTon = 1.102
		Else
			TonneTon = 1
		End If
		Call LoadPurchase()
		Call screenstuff()
		DoNotChange = False
	End Sub
	'UPGRADE_WARNING: Event frmEquipmentPurchaseForm.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub frmEquipmentPurchaseForm_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
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
        'Dim x As Short
        Dim y As Decimal
		Dim w As Short
		Dim h As Short
		Dim u As Single
		
		w = 9150 'Starting Form Scale Width
		h = 6420 'Starting Form Scale Height
		u = (360 / w) * TempWide
		y = (2040 / h) * TempHigh
		
		labEquipmentPurchaseHeading.Top = VB6.TwipsToPixelsY(TempHigh * 0.0093)
		labEquipmentPurchaseHeading.Left = VB6.TwipsToPixelsX(TempWide * 0.0197)
		labEquipmentPurchaseHeading.Width = VB6.TwipsToPixelsX((4605 / w) * TempWide)
		labProjectName.Top = VB6.TwipsToPixelsY(TempHigh * 0.0187)
		labProjectName.Left = VB6.TwipsToPixelsX((5580 / w) * TempWide)
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
		grdEquipmentPurchase.Top = VB6.TwipsToPixelsY(TempHigh * 0.1028)
		grdEquipmentPurchase.Left = VB6.TwipsToPixelsX(TempWide * 0.0197)
		grdEquipmentPurchase.Height = VB6.TwipsToPixelsY(TempHigh * 0.8154)
		grdEquipmentPurchase.Width = VB6.TwipsToPixelsX(TempWide * 0.959)
		labBackToMenu.Top = VB6.TwipsToPixelsY(TempHigh * 0.9532)
		labBackToMenu.Left = VB6.TwipsToPixelsX(TempWide * 0.0656)
		imgBackToMenu.Top = VB6.TwipsToPixelsY(TempHigh * 0.9609)
		imgBackToMenu.Left = VB6.TwipsToPixelsX(TempWide * 0.0066)
		imgBackToMenu.Width = VB6.TwipsToPixelsX(TempWide * 0.0541)
		labEquipmentPurchaseHelp.Top = VB6.TwipsToPixelsY(TempHigh * 0.9532)
		labEquipmentPurchaseHelp.Left = VB6.TwipsToPixelsX(TempWide * 0.9115)
		labEquipmentPurchaseHelp.Width = VB6.TwipsToPixelsX(TempWide * 0.0541)
		comEquipmentPurchaseHelp.Top = VB6.TwipsToPixelsY(TempHigh * 0.9592)
		comEquipmentPurchaseHelp.Left = VB6.TwipsToPixelsX(TempWide * 0.9705)
	End Sub
	Private Sub LoadPurchase()
		Dim x As Short
		Dim y As Short
		Dim r As Short
		Dim s As Short
		Dim TopItOff As Short
		Dim TimeLine As Short
		Dim TempWord As String
		Dim TempTotal(26) As Decimal
		
		On Error Resume Next
		
        grdEquipmentPurchase.set_Cols(0, MaxTime + 2)
		
		For x = 0 To 14
			Select Case LCase(VB.Left(CellValues(EquipmentHourlyResult, x, MinTime).Word, 1))
				Case "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z"
					grdEquipmentPurchase.Row = x + 1
					grdEquipmentPurchase.Col = 0
                    grdEquipmentPurchase.set_ColWidth(0, 0, 2400)
					grdEquipmentPurchase.Text = CellValues(EquipmentHourlyResult, x, MinTime).Word
					TopItOff = x + 2
			End Select
		Next x
		
		grdEquipmentPurchase.Row = TopItOff + 1
		grdEquipmentPurchase.Col = 0
        grdEquipmentPurchase.set_ColWidth(0, 0, 2400)
		grdEquipmentPurchase.Text = "Total"
		
		For TimeLine = 1 To MaxTime + 1
			TempTotal(TimeLine) = 0
			s = -1
			For y = 0 To 6
				If y > 0 Then
					If (TimeLine >= CellValues(Production, 15, y).Value And TimeLine <= CellValues(Production, 16, y).Value) Then
						s = y
					End If
				ElseIf y = 0 Then 
					s = 0
				End If
			Next y
			If s <> -1 Then
				r = 0
				For x = 0 To TopItOff
					grdEquipmentPurchase.Row = r
					grdEquipmentPurchase.Col = TimeLine
                    grdEquipmentPurchase.set_ColWidth(TimeLine, 0, 1200)
					If x = 0 Then
						grdEquipmentPurchase.CellAlignment = 4
						grdEquipmentPurchase.Text = Str(TimeLine)
						r = r + 1
					ElseIf x = TopItOff Then 
						If TimeLine = 1 Then
							If TempTotal(TimeLine) > 0 Then
								grdEquipmentPurchase.Row = r + 1
								grdEquipmentPurchase.CellAlignment = 7
								grdEquipmentPurchase.Text = VB6.Format(Str(TempTotal(TimeLine)), "$##,###,###,##0")
							End If
						ElseIf s > 0 And TimeLine = CellValues(Production, 15, s).Value Then 
							If TempTotal(TimeLine) > 0 Then
								grdEquipmentPurchase.Row = r + 1
								grdEquipmentPurchase.CellAlignment = 7
								grdEquipmentPurchase.Text = VB6.Format(Str(TempTotal(TimeLine)), "$##,###,###,##0")
							End If
						End If
					Else
						If TimeLine = 1 Then
							grdEquipmentPurchase.CellAlignment = 7
							If CellValues(EquipmentPurchaseResult, x - 1, s).Value > 0 Then
								grdEquipmentPurchase.Text = VB6.Format(Str(CellValues(EquipmentPurchaseResult, x - 1, s).Value), "$##,###,###,##0")
								TempTotal(TimeLine) = TempTotal(TimeLine) + CellValues(EquipmentPurchaseResult, x - 1, s).Value
							End If
							r = r + 1
						ElseIf s > 0 And TimeLine = CellValues(Production, 15, s).Value Then 
							grdEquipmentPurchase.CellAlignment = 7
							If CellValues(EquipmentPurchaseResult, x - 1, s).Value > 0 Then
								grdEquipmentPurchase.Text = VB6.Format(Str(CellValues(EquipmentPurchaseResult, x - 1, s).Value), "$##,###,###,##0")
								TempTotal(TimeLine) = TempTotal(TimeLine) + CellValues(EquipmentPurchaseResult, x - 1, s).Value
							End If
							r = r + 1
						End If
					End If
				Next x
			End If
		Next TimeLine
		
	End Sub
	Private Sub labEquipmentPurchaseHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles labEquipmentPurchaseHelp.Click
		IsHelpOn = True
		Call frmSurfaceHelp.gethelptext(WhichScreen, WhichCell)
		frmSurfaceHelp.Show()
	End Sub
End Class