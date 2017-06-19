Option Strict Off
Option Explicit On
Friend Class frmResultsMenu
	Inherits System.Windows.Forms.Form
	Private Sub frmResultsMenu_GotFocus(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.GotFocus
		If Len(PermFileName) > 0 Then
			Me.Text = "Results Menu - " & PermFileName
		Else
			Me.Text = "Results Menu"
		End If
		If ArrowOn <> 0 Then
			imgBackToEngineering.Visible = True
			ArrowOn = 0
		End If
	End Sub
	Private Sub frmResultsMenu_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		Dim x As Short
		Dim TempHigh As Short
		Dim TempWide As Short
		
		Me.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) * 0.6)
		Me.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) * 0.3)
		Me.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) * 0.32)
		Me.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) * 0.65)
		
		TempHigh = VB6.PixelsToTwipsY(Me.ClientRectangle.Height)
		TempWide = VB6.PixelsToTwipsX(Me.ClientRectangle.Width)
		
		For x = 0 To 9
			labResultsTitle(x).Top = VB6.TwipsToPixelsY(TempHigh * (0.1413 + (x * 0.0684)))
			labResultsTitle(x).Left = VB6.TwipsToPixelsX(TempWide * 0.1488)
		Next x
		
		labResultsHeading.Top = VB6.TwipsToPixelsY(TempHigh * 0.0567)
		labResultsHeading.Left = VB6.TwipsToPixelsX(TempWide * 0.0661)
		
		imgBackToEngineering.Top = VB6.TwipsToPixelsY(TempHigh * 0.9518)
		imgBackToEngineering.Left = VB6.TwipsToPixelsX(TempWide * 0.0113)
		imgBackToEngineering.Width = VB6.TwipsToPixelsX(TempWide * 0.141)
		
	End Sub
	Private Sub imgBackToEngineering_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles imgBackToEngineering.Click
		MenuNumber = 2
		ArrowOn = 2
		imgBackToEngineering.Visible = False
		Me.Enabled = False
        frmEngineeringMenu.Show()
        frmEngineeringMenu.Activate()
        frmEngineeringMenu.Enabled = True
        frmEngineeringMenu.imgBackToEquipment.Visible = True
        frmEngineeringMenu.imgOnToResults.Visible = True
	End Sub
	Private Sub labResultsTitle_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles labResultsTitle.Click
		Dim Index As Short = labResultsTitle.GetIndex(eventSender)
		Select Case Index
			Case 0
				Call SupplyCost()
				WhichScreen = SupplyResult
				frmSupplyCostForm.Show()
			Case 1
				Call LaborCost()
				Call SalaryCost()
				WhichScreen = LaborResult
				frmHourlyCostForm.Show()
			Case 2
				Call LaborCost()
				Call SalaryCost()
				WhichScreen = SalaryResult
				frmSalaryCostForm.Show()
			Case 3
				Call EquipmentCost()
				WhichScreen = EquipmentSupplyResult
				frmEquipmentCostForm.Show()
			Case 4
				WhichScreen = EquipmentNumberResult
				frmEquipmentNumberForm.Show()
			Case 5
				WhichScreen = EquipmentPurchaseResult
				frmEquipmentPurchaseForm.Show()
			Case 6
				Call DevelopmentCost()
				WhichScreen = DevelopmentResult
				frmDevelopmentCostForm.Show()
			Case 7
				Call StripCostCalc()
				WhichScreen = StrippingResult
				frmStrippingCostForm.Show()
			Case 8
				WhichScreen = Summary
				frmTotalCostForm.Show()
            Case 9
                WhichScreen = UtilScreen
                frmUtility.Show()

        End Select
    End Sub
End Class