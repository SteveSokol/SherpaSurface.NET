Option Strict Off
Option Explicit On
Friend Class frmIntro
	Inherits System.Windows.Forms.Form
	Private Sub comBeginProgram_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles comBeginProgram.Click
		Dim Index As Short = comBeginProgram.GetIndex(eventSender)
		
		'If compiling for cd, get rid of the following
		'UserYes = 2
		'If compiling for cd, get rid of the above
		
		If UserYes = 0 Then End
		
		If Index = 0 Then
			UnitType = English
			FootConv = 1
			AcreConv = 1
			SquareFootConv = 1
			CubicFootConv = 1
			GallonConv = 1
			TonConv = 1
			PoundConv = 1
			MileConv = 1
			DensConv = 1
			PowderConv = 1
			InchConv = 1
		Else
			UnitType = Metric
			FootConv = 0.3048
			AcreConv = 0.4047
			SquareFootConv = 0.0929
			CubicFootConv = 0.0283
			GallonConv = 3.785
			TonConv = 0.9071
			PoundConv = 0.4536
			MileConv = 1.609
			DensConv = 1.1865
			PowderConv = 0.4999
			InchConv = 2.54
		End If
		
		Call InDat()
		
		Me.Hide()
		
		'If compiling for cd, get rid of the following:
		'  Call openall
		'  frmSetItUp.Hide
		'  frmInputMenu.Show
		'If compiling for cd, get rid of the above
		
	End Sub
	Private Sub frmIntro_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		Me.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) * 0.1)
		Me.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) * 0.25)
		Me.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) * 0.8)
		Me.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) * 0.5)
		
		labIntroTitle.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Height) * 0.0004)
		labIntroTitle.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Width) * 0.008)
		
		imgIntro.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.ClientRectangle.Height) * 0.9782)
		imgIntro.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.ClientRectangle.Width) * 0.9835)
		
		labTitleOne.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.ClientRectangle.Width) - 4560)
		labTitleTwo.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.ClientRectangle.Width) - 4560)
		labTitleThree.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.ClientRectangle.Width) - 5640)
		
		labTitleOne.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.ClientRectangle.Height) - 780)
		labTitleTwo.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.ClientRectangle.Height) - 540)
		labTitleThree.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.ClientRectangle.Height) - 300)
		
		comBeginProgram(0).Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.ClientRectangle.Height) - 740)
		comBeginProgram(0).Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.ClientRectangle.Width) * 0.0188)
		
		comBeginProgram(1).Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.ClientRectangle.Height) - 420)
		comBeginProgram(1).Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.ClientRectangle.Width) * 0.0188)
		
	End Sub
	Public Sub UpInDat()
        Dim TimeLine As Object = Nothing
		
		On Error Resume Next
		
        If (Val(CStr(CellValues(Production, 5, TimeLine).Value)) + Val(CStr(CellValues(Production, 10, TimeLine).Value))) < 25000 Then
            If CellValues(Supply, 2, TimeLine).Changed = False Then CellValues(Supply, 2, TimeLine).Value = 2.6238
            If CellValues(Supply, 14, TimeLine).Changed = False Then CellValues(Supply, 14, TimeLine).Value = 2.6238
            If CellValues(Supply, 7, TimeLine).Changed = False Then CellValues(Supply, 7, TimeLine).Word = "Cartridge"
            If CellValues(Supply, 9, TimeLine).Changed = False Then CellValues(Supply, 9, TimeLine).Value = 900
            If CellValues(Supply, 19, TimeLine).Changed = False Then CellValues(Supply, 19, TimeLine).Value = 900
            If CellValues(Supply, 10, TimeLine).Changed = False Then CellValues(Supply, 10, TimeLine).Value = 6500
            If CellValues(Supply, 20, TimeLine).Changed = False Then CellValues(Supply, 20, TimeLine).Value = 6500
        End If
		
	End Sub
End Class