Option Strict Off
Option Explicit On
Friend Class frmEngineeringMenu
	Inherits System.Windows.Forms.Form
	Private Sub frmEngineeringMenu_GotFocus(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.GotFocus
		If Len(PermFileName) > 0 Then
			Me.Text = "Development Menu - " & PermFileName
		Else
			Me.Text = "Development Menu"
		End If
		If ArrowOn <> 0 Then
			MenuNumber = 2
			imgBackToEquipment.Visible = True
			imgOnToResults.Visible = True
			ArrowOn = 0
		End If
	End Sub
	Private Sub frmEngineeringMenu_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		Dim x As Short
		Dim TempHigh As Short
		Dim TempWide As Short
		
		Me.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) * 0.6)
		Me.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) * 0.3)
		Me.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) * 0.23)
		Me.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) * 0.45)
		
		TempHigh = VB6.PixelsToTwipsY(Me.ClientRectangle.Height)
		TempWide = VB6.PixelsToTwipsX(Me.ClientRectangle.Width)
		
		For x = 0 To 10
			labEngineeringTitle(x).Top = VB6.TwipsToPixelsY(TempHigh * (0.1413 + (x * 0.0684)))
			labEngineeringTitle(x).Left = VB6.TwipsToPixelsX(TempWide * 0.1488)
		Next x
		
		labEngineeringHeading.Top = VB6.TwipsToPixelsY(TempHigh * 0.0567)
		labEngineeringHeading.Left = VB6.TwipsToPixelsX(TempWide * 0.0661)
		
		imgBackToEquipment.Top = VB6.TwipsToPixelsY(TempHigh * 0.9518)
		imgBackToEquipment.Left = VB6.TwipsToPixelsX(TempWide * 0.0113)
		imgBackToEquipment.Width = VB6.TwipsToPixelsX(TempWide * 0.141)
		
		imgOnToResults.Top = VB6.TwipsToPixelsY(TempHigh * 0.9518)
		imgOnToResults.Left = VB6.TwipsToPixelsX(TempWide * 0.843)
		imgOnToResults.Width = VB6.TwipsToPixelsX(TempWide * 0.141)
	End Sub
	Private Sub imgBackToEquipment_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles imgBackToEquipment.Click
		MenuNumber = 1
		ArrowOn = 2
		imgBackToEquipment.Visible = False
		imgOnToResults.Visible = False
		Me.Enabled = False
        frmEquipmentMenu.Show()
        frmEquipmentMenu.Activate()
        frmEquipmentMenu.Enabled = True
        frmEquipmentMenu.imgBackToInput.Visible = True
        frmEquipmentMenu.imgOnToEngineering.Visible = True
	End Sub
	Private Sub imgOnToResults_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles imgOnToResults.Click
		ArrowOn = 2
		Call TimeLineCalc()
		OpBin(1) = 0
		For WhichSegment = 0 To MaxSegment
			Call BlastEngr()
			Call ElectEngr()
			Call FuelEngr()
		Next WhichSegment
		WhichSegment = 0
		Call bcalc()
		Call RoadCost()
		Call ProductionCalc()
		Call LaborCost()
		Call SalaryCost()
		Call SupplyCost()
		Call EquipmentCost()
		Call NewEquipEngr()
		Call EquipmentCost()
		Call StripCostCalc()
		Call DevelopmentCost()
		Call SummaryCost()
		WhichSegment = 0
		imgBackToEquipment.Visible = False
		imgOnToResults.Visible = False
		Me.Enabled = False
        frmResultsMenu.Show()
        frmResultsMenu.Activate()
        frmResultsMenu.Enabled = True
        frmResultsMenu.imgBackToEngineering.Visible = True
	End Sub
	
	Private Sub labEngineeringTitle_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles labEngineeringTitle.Click
		Dim Index As Short = labEngineeringTitle.GetIndex(eventSender)
		Select Case Index
			Case 0
				Call StripEngr()
				Call dcalc()
				Call StripSummary()
				AutoSave = True
				Call FileStuff(2)
				WhichScreen = Strip
				frmStripData.Show()
			Case 1
				Call RoadEngr()
				AutoSave = True
				Call FileStuff(2)
				WhichScreen = Road
				frmRoadData.Show()
			Case 2
				Call BuildEngr()
				AutoSave = True
				Call FileStuff(2)
				WhichScreen = Building
				frmBuildingData.Show()
			Case 3
				Call TimeLineCalc()
				For WhichSegment = 0 To MaxSegment
					Call BlastEngr()
				Next WhichSegment
				AutoSave = True
				Call FileStuff(2)
				WhichSegment = 0
				WhichScreen = Powder
				frmPowderData.Show()
			Case 4
				Call TimeLineCalc()
				For WhichSegment = 0 To MaxSegment
					Call ElectEngr()
				Next WhichSegment
				AutoSave = True
				Call FileStuff(2)
				WhichSegment = 0
				WhichScreen = Electrical
				frmElectricalData.Show()
			Case 5
				Call ClearEngr()
				AutoSave = True
				Call FileStuff(2)
				WhichScreen = Clearing
				frmClearingData.Show()
			Case 6
				Call SiteEngr()
				AutoSave = True
				Call FileStuff(2)
                WhichScreen = Convert.ToInt16(Site)
				frmSiteData.Show()
			Case 7
				Call TimeLineCalc()
				OpBin(1) = 0
				For WhichSegment = 0 To MaxSegment
					Call FuelEngr()
				Next WhichSegment
				WhichSegment = 0
				Call bcalc()
				AutoSave = True
				Call FileStuff(2)
				WhichSegment = 0
				WhichScreen = FuelStorage
				frmFuelStorageData.Show()
			Case 8
				WhichSegment = 0
				Call RoadCost()
				Call ClearEngr()
				Call SiteEngr()
				Call bcalc()
				AutoSave = True
				Call FileStuff(2)
				WhichSegment = 0
				WhichScreen = Development
				frmDevelopmentData.Show()
			Case 9
				AutoSave = True
				Call FileStuff(2)
				WhichScreen = WorkForce
				frmWorkForceData.Show()
			Case 10
				AutoSave = True
				Call FileStuff(2)
				WhichScreen = Staff
				frmStaffData.Show()
		End Select
		
	End Sub
End Class