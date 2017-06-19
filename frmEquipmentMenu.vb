Option Strict Off
Option Explicit On
Friend Class frmEquipmentMenu
	Inherits System.Windows.Forms.Form
	Private Sub frmEquipmentMenu_GotFocus(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.GotFocus
		If Len(PermFileName) > 0 Then
			Me.Text = "Equipment Menu - " & PermFileName
		Else
			Me.Text = "Equipment Menu"
		End If
		If ArrowOn <> 0 Then
			MenuNumber = 1
			imgBackToInput.Visible = True
			imgOnToEngineering.Visible = True
			ArrowOn = 0
			If LTrim(RTrim(LCase(CellValues(Production, 40, 0).Word))) = "sherpa" Then
				labEquipmentTitle(0).Enabled = True
				labEquipmentTitle(1).Enabled = True
			Else
				labEquipmentTitle(0).Enabled = False
				labEquipmentTitle(1).Enabled = False
			End If
		End If
	End Sub
	Private Sub frmEquipmentMenu_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		Dim x As Short
		Dim TempHigh As Short
		Dim TempWide As Short
		
		Me.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) * 0.6)
		Me.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) * 0.3)
		Me.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) * 0.14)
		Me.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) * 0.25)
		
		TempWide = VB6.PixelsToTwipsX(Me.ClientRectangle.Width)
		TempHigh = VB6.PixelsToTwipsY(Me.ClientRectangle.Height)
		
		For x = 0 To 8
			labEquipmentTitle(x).Top = VB6.TwipsToPixelsY(TempHigh * (0.1413 + (x * 0.0684)))
			labEquipmentTitle(x).Left = VB6.TwipsToPixelsX(TempWide * 0.1488)
		Next x
		
		labEquipmentHeading.Top = VB6.TwipsToPixelsY(TempHigh * 0.0567)
		labEquipmentHeading.Left = VB6.TwipsToPixelsX(TempWide * 0.0661)
		
		imgBackToInput.Top = VB6.TwipsToPixelsY(TempHigh * 0.9518)
		imgBackToInput.Left = VB6.TwipsToPixelsX(TempWide * 0.0113)
		imgBackToInput.Width = VB6.TwipsToPixelsX(TempWide * 0.141)
		
		imgOnToEngineering.Top = VB6.TwipsToPixelsY(TempHigh * 0.9518)
		imgOnToEngineering.Left = VB6.TwipsToPixelsX(TempWide * 0.843)
		imgOnToEngineering.Width = VB6.TwipsToPixelsX(TempWide * 0.141)
		
	End Sub
	Private Sub imgBackToInput_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles imgBackToInput.Click
		MenuNumber = 0
		ArrowOn = 2
		imgBackToInput.Visible = False
		imgOnToEngineering.Visible = False
		Me.Enabled = False
        frmInputMenu.Show()
        frmInputMenu.Activate()
		frmInputMenu.Enabled = True
		frmInputMenu.imgOnToEquipment.Visible = True
	End Sub
	Private Sub imgOnToEngineering_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles imgOnToEngineering.Click
		MenuNumber = 2
		ArrowOn = 2
		Call TimeLineCalc()
		ResetMachines = False
		Call enteq()
		Call CostItAll()
		OpBin(1) = 0
		For WhichSegment = 0 To MaxSegment
			Call BlastEngr()
			Call ElectEngr()
			Call FuelEngr()
		Next WhichSegment
		WhichSegment = 0
		Call CostItAll()
		Call bcalc()
		Call StripEngr()
		Call RoadEngr()
		Call BuildEngr()
		Call ClearEngr()
		Call bcalc()
		Call SiteEngr()
		Call RoadCost()
		Call StripCostCalc()
		imgBackToInput.Visible = False
		imgOnToEngineering.Visible = False
		Me.Enabled = False
        frmEngineeringMenu.Show()
        frmEngineeringMenu.Activate()
        frmEngineeringMenu.Enabled = True
        frmEngineeringMenu.imgBackToEquipment.Visible = True
        frmEngineeringMenu.imgOnToResults.Visible = True
    End Sub
	
	Private Sub labEquipmentTitle_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles labEquipmentTitle.Click
		Dim Index As Short = labEquipmentTitle.GetIndex(eventSender)
		Select Case Index
			Case 0
				Call TimeLineCalc()
				ResetMachines = True
				Call enteq()
				ResetMachines = False
				AutoSave = True
				Call FileStuff(2)
				WhichSegment = 0
				WhichScreen = EquipmentOne
				frmProdEquipmentData.Show()
			Case 1
				Call TimeLineCalc()
				ResetMachines = True
				Call enteq()
				ResetMachines = False
				AutoSave = True
				Call FileStuff(2)
				WhichSegment = 0
				WhichScreen = EquipmentOne
				frmAncEquipmentData.Show()
			Case 2
				WhichScreen = EquipmentOne
				Call TimeLineCalc()
				ResetMachines = False
				Call enteq()
				AutoSave = True
				Call FileStuff(2)
				WhichSegment = 0
				frmProdEquipmentData.Show()
			Case 3
				WhichScreen = EquipmentOne
				Call TimeLineCalc()
				ResetMachines = False
				Call enteq()
				AutoSave = True
				Call FileStuff(2)
				WhichSegment = 0
				frmAncEquipmentData.Show()
			Case 4
				Call CostItAll()
				AutoSave = True
				Call FileStuff(2)
				WhichScreen = Purchase
				frmProdPurchaseData.Show()
			Case 5
				Call CostItAll()
				AutoSave = True
				Call FileStuff(2)
				WhichScreen = Purchase
				frmAncPurchaseData.Show()
			Case 6
				Call CostItAll()
				AutoSave = True
				Call FileStuff(2)
				WhichScreen = Replace_Renamed
				frmReplaceData.Show()
			Case 7
				Call CostItAll()
				AutoSave = True
				Call FileStuff(2)
				WhichScreen = Diesel
				frmProdEqOpData.Show()
			Case 8
				Call CostItAll()
				AutoSave = True
				Call FileStuff(2)
				WhichScreen = Diesel
				frmAncEqOpData.Show()
		End Select
	End Sub
End Class