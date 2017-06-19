Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports Microsoft.VisualBasic.PowerPacks
Friend Class frmEquipmentNumberForm
	Inherits System.Windows.Forms.Form
	Dim TempHigh As Single
	Dim TempWide As Single
	Dim TonneTon As Decimal
	Private Sub comEquipmentNumberHelp_Click()
		Me.PrintForm1.Print(Me, PowerPacks.Printing.PrintForm.PrintOption.CompatibleModeClientAreaOnly)
	End Sub
	'UPGRADE_WARNING: Form event frmEquipmentNumberForm.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmEquipmentNumberForm_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        'Dim baseunit As String
        'Dim baselength As Short
        'Dim i As Short
        If IsHelpOn = True Then
			IsHelpOn = False
		Else
			DoNotChange = True
			WhichScreen = EquipmentNumberResult
			labProjectName.Text = ProjectTitle(0)
			DoNotChange = False
			WhichCell = 0
		End If
	End Sub
	Private Sub frmEquipmentNumberForm_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
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
		Call LoadNumber()
		Call screenstuff()
		DoNotChange = False
	End Sub
	'UPGRADE_WARNING: Event frmEquipmentNumberForm.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub frmEquipmentNumberForm_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
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
		
		labEquipmentNumberHeading.Top = VB6.TwipsToPixelsY(TempHigh * 0.0093)
		labEquipmentNumberHeading.Left = VB6.TwipsToPixelsX(TempWide * 0.0197)
		labEquipmentNumberHeading.Width = VB6.TwipsToPixelsX((4605 / w) * TempWide)
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
		grdEquipmentNumber.Top = VB6.TwipsToPixelsY(TempHigh * 0.1028)
		grdEquipmentNumber.Left = VB6.TwipsToPixelsX(TempWide * 0.0197)
		grdEquipmentNumber.Height = VB6.TwipsToPixelsY(TempHigh * 0.8154)
		grdEquipmentNumber.Width = VB6.TwipsToPixelsX(TempWide * 0.959)
		labBackToMenu.Top = VB6.TwipsToPixelsY(TempHigh * 0.9532)
		labBackToMenu.Left = VB6.TwipsToPixelsX(TempWide * 0.0656)
		imgBackToMenu.Top = VB6.TwipsToPixelsY(TempHigh * 0.9609)
		imgBackToMenu.Left = VB6.TwipsToPixelsX(TempWide * 0.0066)
		imgBackToMenu.Width = VB6.TwipsToPixelsX(TempWide * 0.0541)
		labEquipmentNumberHelp.Top = VB6.TwipsToPixelsY(TempHigh * 0.9532)
		labEquipmentNumberHelp.Left = VB6.TwipsToPixelsX(TempWide * 0.9115)
		labEquipmentNumberHelp.Width = VB6.TwipsToPixelsX(TempWide * 0.0541)
		comEquipmentNumberPrint.Top = VB6.TwipsToPixelsY(TempHigh * 0.9592)
		comEquipmentNumberPrint.Left = VB6.TwipsToPixelsX(TempWide * 0.9705)
	End Sub
	Private Sub LoadNumber()
		Dim x As Short
		Dim y As Short
		Dim s As Short
		Dim r As Short
		Dim TopItOff As Short
		Dim TimeLine As Short
		Dim TempWord As String
		Dim tempnumber As Decimal
		
		On Error Resume Next
		
        grdEquipmentNumber.set_Cols(0, MaxTime + 2)
		
		r = 0
		
		For x = 0 To 14
			Select Case LCase(VB.Left(CellValues(EquipmentHourlyResult, x, MinTime).Word, 1))
				Case "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z"
					r = r + 1
					grdEquipmentNumber.Row = r
					grdEquipmentNumber.Col = 0
                    grdEquipmentNumber.set_ColWidth(0, 0, 2400)
					grdEquipmentNumber.Text = CellValues(EquipmentHourlyResult, x, MinTime).Word
					TopItOff = r + 2
			End Select
		Next x
		
		r = 0
		
		For TimeLine = 1 To MaxTime + 1
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
				For x = 0 To 14
					grdEquipmentNumber.Row = r
					grdEquipmentNumber.Col = TimeLine
                    grdEquipmentNumber.set_ColWidth(TimeLine, 0, 1200)
					If x = 0 Then
						grdEquipmentNumber.CellAlignment = 4
						grdEquipmentNumber.Text = Str(TimeLine)
						r = r + 1
					Else
						If TimeLine = 1 Then
							If CellValues(EquipmentNumberResult, x - 1, s).Value > 0 Then
								grdEquipmentNumber.CellAlignment = 7
								grdEquipmentNumber.Text = VB6.Format(Str(CellValues(EquipmentNumberResult, x - 1, s).Value), "#,##0")
								r = r + 1
							End If
						ElseIf s > 0 And TimeLine = CellValues(Production, 15, s).Value Then 
							If CellValues(EquipmentNumberResult, x - 1, s).Value > 0 Then
								grdEquipmentNumber.CellAlignment = 7
								grdEquipmentNumber.Text = VB6.Format(Str(CellValues(EquipmentNumberResult, x - 1, s).Value), "#,##0")
								r = r + 1
							End If
						End If
					End If
				Next x
			End If
		Next TimeLine
		
	End Sub
	Private Sub labEquipmentNumberHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles labEquipmentNumberHelp.Click
		IsHelpOn = True
		Call frmSurfaceHelp.gethelptext(WhichScreen, WhichCell)
		frmSurfaceHelp.Show()
	End Sub
End Class