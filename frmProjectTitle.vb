Option Strict Off
Option Explicit On
Friend Class frmProjectTitle
	Inherits System.Windows.Forms.Form
	Private Sub comTitleHelp_Click()
		IsHelpOn = True
		Call frmSurfaceHelp.gethelptext(WhichScreen, WhichCell)
		frmSurfaceHelp.Show()
	End Sub
	'UPGRADE_WARNING: Form event frmProjectTitle.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmProjectTitle_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		Dim i As Short
		If IsHelpOn = True Then
			txtProjectTitles(WhichCell).Focus()
			IsHelpOn = False
		Else
			If txtProjectTitles(4).Text = "" Then
				txtProjectTitles(4).Text = VB6.Format(Today, "Long Date")
				CellValues(Project, 4, 0).Word = txtProjectTitles(4).Text
			End If
			If PageChange(0) = True Then
				For i = 0 To 4
					txtProjectTitles(i).Text = CellValues(Project, i, 0).Word
				Next i
			End If
			WhichCell = 0
			txtProjectTitles(0).Focus()
		End If
	End Sub
	Private Sub frmProjectTitle_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Dim x As Short
		Dim TempHigh As Decimal
		Dim TempWide As Decimal
		
		Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - (VB6.PixelsToTwipsY(Me.Height) + 350)) / 2)
		Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2)
		
		If VB6.PixelsToTwipsY(Me.Top) < 0 Then Me.Top = 0
		If VB6.PixelsToTwipsX(Me.Left) < 0 Then Me.Left = 0
		
		TempHigh = VB6.PixelsToTwipsY(Me.ClientRectangle.Height)
		TempWide = VB6.PixelsToTwipsX(Me.ClientRectangle.Width)
		
		For x = 0 To 4
			labProjectTitles(x).Top = VB6.TwipsToPixelsY((TempHigh * (0.105)) + (TempHigh * (0.15 + (x / 10))))
			labProjectTitles(x).Left = VB6.TwipsToPixelsX(TempWide * 0.0748)
			labProjectTitles(x).Width = VB6.TwipsToPixelsX(TempWide * 0.2318)
			txtProjectTitles(x).Top = VB6.TwipsToPixelsY((TempHigh * (0.1)) + (TempHigh * (0.15 + (x / 10))))
			txtProjectTitles(x).Left = VB6.TwipsToPixelsX(TempWide * 0.32)
			If x = 4 Then
				txtProjectTitles(x).Width = VB6.TwipsToPixelsX(TempWide * 0.4619)
			Else
				txtProjectTitles(x).Width = VB6.TwipsToPixelsX(TempWide * 0.6176)
			End If
		Next x
		
		labProjectHeading.Left = VB6.TwipsToPixelsX(TempWide * 0.0194)
		labProjectHeading.Top = VB6.TwipsToPixelsY(TempHigh * 0.0334)
		
		labBackToMenu.Left = VB6.TwipsToPixelsX(TempWide * 0.08)
		labBackToMenu.Top = VB6.TwipsToPixelsY(TempHigh * 0.9425)
		
		imgBackToMenu.Left = VB6.TwipsToPixelsX(TempWide * 0.012)
		imgBackToMenu.Top = VB6.TwipsToPixelsY(TempHigh * 0.9496)
		
	End Sub
	Private Sub frmProjectTitle_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		
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
	'UPGRADE_WARNING: Event txtProjectTitles.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtProjectTitles_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtProjectTitles.TextChanged
		Dim Index As Short = txtProjectTitles.GetIndex(eventSender)
		
		PageChange(0) = True
		CellValues(Project, Index, 0).Word = txtProjectTitles(Index).Text
		
	End Sub
	Private Sub txtProjectTitles_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtProjectTitles.Enter
		Dim Index As Short = txtProjectTitles.GetIndex(eventSender)
		WhichCell = Index
	End Sub
End Class