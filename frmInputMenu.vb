Option Strict Off
Option Explicit On
Friend Class frmInputMenu
	Inherits System.Windows.Forms.Form
	Private Sub frmInputMenu_GotFocus(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.GotFocus
		If Len(PermFileName) > 0 Then
			Me.Text = "Project Design Menu - " & PermFileName
		Else
			Me.Text = "Project Design Menu"
		End If
		If ArrowOn = 2 Then
			MenuNumber = 0
			imgOnToEquipment.Visible = True
			ArrowOn = 0
		End If
	End Sub
	Private Sub frmInputMenu_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim x As Short
		Dim TempHigh As Short
		Dim TempWide As Short
        'Dim findintro As String
        frmResultsMenu.Show()
		frmEngineeringMenu.Show()
		frmEquipmentMenu.Show()
		frmResultsMenu.Enabled = False
		frmEngineeringMenu.Enabled = False
		frmEquipmentMenu.Enabled = False
		Me.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) * 0.6)
		Me.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) * 0.3)
		Me.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) * 0.05)
		Me.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) * 0.05)
		
		TempWide = VB6.PixelsToTwipsX(Me.ClientRectangle.Width)
		TempHigh = VB6.PixelsToTwipsY(Me.ClientRectangle.Height)
		
		For x = 0 To 9
			labInputTitle(x).Top = VB6.TwipsToPixelsY(TempHigh * (0.1413 + (x * 0.0684)))
			labInputTitle(x).Left = VB6.TwipsToPixelsX(TempWide * 0.1488)
		Next x
		
		labInputHeading.Top = VB6.TwipsToPixelsY(TempHigh * 0.0567)
		labInputHeading.Left = VB6.TwipsToPixelsX(TempWide * 0.0661)
		'labInputHeading.Width = TempWide * 0.6736
		
		imgOnToEquipment.Top = VB6.TwipsToPixelsY(TempHigh * 0.9518)
		imgOnToEquipment.Left = VB6.TwipsToPixelsX(TempWide * 0.8345)
		imgOnToEquipment.Width = VB6.TwipsToPixelsX(TempWide * 0.141)
		Me.Show()
	End Sub
	Private Sub imgOnToEquipment_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles imgOnToEquipment.Click
		MenuNumber = 1
		ArrowOn = 2
		Call TimeLineCalc()
		For WhichSegment = 0 To MaxSegment
			Call HaulDesign()
			Call dlopt()
			Call slopt()
		Next WhichSegment
		WhichSegment = 0
		imgOnToEquipment.Visible = False
		Me.Enabled = False
        frmEquipmentMenu.Show()
        frmEquipmentMenu.Activate()
        frmEquipmentMenu.Enabled = True
        frmEquipmentMenu.imgOnToEngineering.Visible = True
        frmEquipmentMenu.imgBackToInput.Visible = True

	End Sub
	Private Sub labInputTitle_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles labInputTitle.Click
		Dim Index As Short = labInputTitle.GetIndex(eventSender)
		
		Select Case Index
			Case 0
				frmProjectTitle.Show()
			Case 1
                WhichScreen = Production
                frmOperatingData.Show()
            Case 2
                Call TimeLineCalc()
                For WhichSegment = 0 To MaxSegment
                    Call HaulDesign()
                Next WhichSegment
                AutoSave = True
                Call FileStuff(2)
                WhichSegment = 0
                WhichScreen = Pit
                frmPitData.Show()
			Case 3
				Call TimeLineCalc()
				For WhichSegment = 0 To MaxSegment
					Call dlopt()
					Call HaulDesign()
				Next WhichSegment
				AutoSave = True
				Call FileStuff(2)
				WhichSegment = 0
				WhichScreen = Deposit
				frmDepositData.Show()
			Case 4
				Call TimeLineCalc()
				For WhichSegment = 0 To MaxSegment
					Call HaulDesign()
				Next WhichSegment
				AutoSave = True
				Call FileStuff(2)
				WhichSegment = 0
				WhichScreen = Haul
				frmHaulData.Show()
			Case 5
				Call TimeLineCalc()
				For WhichSegment = 0 To MaxSegment
					Call HaulDesign()
					Call ccopt()
				Next WhichSegment
				AutoSave = True
				Call FileStuff(2)
				WhichSegment = 0
				WhichScreen = Convey
				frmConveyorData.Show()
			Case 6
				Call TimeLineCalc()
				For WhichSegment = 0 To MaxSegment
					Call HaulDesign()
					Call PumpEngr()
				Next WhichSegment
				AutoSave = True
				Call FileStuff(2)
				WhichSegment = 0
				WhichScreen = Pumping
				frmPumpingData.Show()
			Case 7
				Call TimeLineCalc()
				For WhichSegment = 0 To MaxSegment
					Call HaulDesign()
					Call dlopt()
				Next WhichSegment
				AutoSave = True
				Call FileStuff(2)
				WhichSegment = 0
				WhichScreen = Supply
				frmSupplyData.Show()
			Case 8
				Call TimeLineCalc()
				For WhichSegment = 0 To MaxSegment
					Call HaulDesign()
				Next WhichSegment
				AutoSave = True
				Call FileStuff(2)
				WhichSegment = 0
				WhichScreen = Wage
				frmWageData.Show()
			Case 9
				Call TimeLineCalc()
				For WhichSegment = 0 To MaxSegment
					Call HaulDesign()
					Call slopt()
				Next WhichSegment
				AutoSave = True
				Call FileStuff(2)
				WhichSegment = 0
				WhichScreen = Salary
				frmSalaryData.Show()
		End Select
		
	End Sub
End Class