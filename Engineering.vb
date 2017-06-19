Option Strict Off
Option Explicit On
Module Engineering
	Public Sub StripEngr()
		Dim x As Object
		
		On Error Resume Next
		
		WhichSegment = 0
		
		If CellValues(Strip, 0, WhichSegment).Changed = False Then
			CellValues(Strip, 0, WhichSegment).Value = CellValues(Production, 4, WhichSegment).Value
		End If
		
		If CellValues(Strip, 1, WhichSegment).Changed = False Then
			CellValues(Strip, 1, WhichSegment).Value = CellValues(Production, 10, WhichSegment).Value
		End If
		
		For x = 2 To 5
			'UPGRADE_WARNING: Couldn't resolve default property of object x. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If CellValues(Strip, x, WhichSegment).Changed = False Then
				'UPGRADE_WARNING: Couldn't resolve default property of object x. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				CellValues(Strip, x, WhichSegment).Value = CellValues(Haul, x + 2, WhichSegment).Value
			End If
		Next x
		
		For x = 6 To 17
			'UPGRADE_WARNING: Couldn't resolve default property of object x. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If CellValues(Strip, x, WhichSegment).Changed = False Then
				'UPGRADE_WARNING: Couldn't resolve default property of object x. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				CellValues(Strip, x, WhichSegment).Value = CellValues(Haul, x + 14, WhichSegment).Value
			End If
		Next x
		
	End Sub
	Public Sub BuildEngr()
		Dim OfficeSize As Object
		Dim ShopSize As Decimal
		Dim DrySize As Decimal
		Dim TempAdmin As Decimal
		Dim TotalAdmin As Decimal
		Dim TestSegment As Decimal
		Dim WarehouseSize As Decimal
		Dim hrsh As Decimal
		Dim shdy As Decimal
		
		Dim TotalMan As Short
		Dim x As Short
		Dim ao As Short
		Dim aw As Short
		Dim bo As Short
		Dim bw As Short
		Dim dore As Short
		Dim dw As Short
		Dim eo As Short
		Dim ew As Short
		Dim co As Short
		Dim cw As Short
		Dim ot As Decimal
		Dim wt As Decimal
		Dim TotalArea As Decimal
		Dim CheckOreHeight As Decimal
		Dim CheckWasteHeight As Decimal
		
		On Error Resume Next
		
		Call otcal(ot, wt)
		
		ao = Int(CellValues(EquipmentOne, 0, WhichSegment).Value)
		aw = Int(CellValues(EquipmentOne, 5, WhichSegment).Value)
		bo = Int(CellValues(EquipmentOne, 1, WhichSegment).Value)
		bw = Int(CellValues(EquipmentOne, 6, WhichSegment).Value)
		dore = Int(CellValues(EquipmentOne, 2, WhichSegment).Value)
		dw = Int(CellValues(EquipmentOne, 7, WhichSegment).Value)
		eo = Int(CellValues(EquipmentOne, 3, WhichSegment).Value)
		ew = Int(CellValues(EquipmentOne, 8, WhichSegment).Value)
		co = Int(CellValues(EquipmentOne, 4, WhichSegment).Value)
		cw = Int(CellValues(EquipmentOne, 9, WhichSegment).Value)
		
		'=============================== Shop Size =================================='
		
		TotalArea = 0
		TotalArea = (EqCost(Ore, Truck, OutNumber) * EqDefault(Truck, co, Area)) + (EqCost(Waste, Truck, OutNumber) * EqDefault(Truck, cw, Area))
		TotalArea = TotalArea + ((EqCost(Ore, Articulated, OutNumber) * EqDefault(Articulated, co, Area)) + (EqCost(Waste, Articulated, OutNumber) * EqDefault(Articulated, cw, Area)))
		TotalArea = TotalArea + ((EqCost(Ore, Scraper, OutNumber) * EqDefault(Scraper, co, Area)) + (EqCost(Waste, Scraper, OutNumber) * EqDefault(Scraper, cw, Area)))
		TotalArea = TotalArea + ((CellValues(Convey, 7, WhichSegment).Value + CellValues(Convey, 18, WhichSegment).Value) / 40)
		
		If (EqCost(Ore, Truck, OutNumber) + EqCost(Waste, Truck, OutNumber) + EqCost(Waste, Articulated, OutNumber) + EqCost(Waste, Articulated, OutNumber) + EqCost(Waste, Scraper, OutNumber) + EqCost(Waste, Scraper, OutNumber)) > 8 Then
			ShopSize = 2.25 * TotalArea
		Else
			ShopSize = 2.65 * TotalArea
		End If
		
		If CellValues(Building, 0, WhichSegment).Changed = False Then
			CellValues(Building, 0, WhichSegment).Value = System.Math.Round(((ShopSize) ^ 0.5) * 1.4, 1)
		End If
		
		If CellValues(Building, 1, WhichSegment).Changed = False And CellValues(Building, 0, WhichSegment).Value <> 0 Then
			CellValues(Building, 1, WhichSegment).Value = System.Math.Round(ShopSize / CellValues(Building, 0, WhichSegment).Value, 1)
		End If
		
		CheckOreHeight = (EqDefault(Truck, co, TruckHeight) + EqDefault(Articulated, co, TruckHeight) + EqDefault(Scraper, co, TruckHeight))
		CheckWasteHeight = (EqDefault(Truck, cw, TruckHeight) + EqDefault(Articulated, cw, TruckHeight) + EqDefault(Scraper, cw, TruckHeight))
		
		
		If CellValues(Building, 2, WhichSegment).Changed = False Then
			If (CheckOreHeight < 11 And CheckWasteHeight < 11) Then
				CellValues(Building, 2, WhichSegment).Value = 12
			ElseIf (CheckOreHeight >= 11 Or CheckWasteHeight >= 11) And (CheckOreHeight < 18 And CheckWasteHeight < 18) Then 
				CellValues(Building, 2, WhichSegment).Value = 20
			ElseIf (CheckOreHeight >= 18 Or CheckWasteHeight >= 18) Then 
				CellValues(Building, 2, WhichSegment).Value = 28
			End If
		End If
		
		Select Case ot + wt
			Case Is <= 5000
				If CellValues(Building, 3, WhichSegment).Changed = False Then
					CellValues(Building, 3, WhichSegment).Word = "Mobile"
				End If
			Case Is <= 20000
				If CellValues(Building, 3, WhichSegment).Changed = False Then
					CellValues(Building, 3, WhichSegment).Word = "Wood Frame/Steel Siding"
				End If
			Case Else
				If CellValues(Building, 3, WhichSegment).Changed = False Then
					CellValues(Building, 3, WhichSegment).Word = "Permanent Steel"
				End If
		End Select
		
		'================================ Dry Size =================================='
		
		Call hrlab()
		Call hrcal(hrsh)
		Call shcal(shdy)
		
		TotalMan = 12
		
		Call TimeLineCalc()
		For WhichSegment = 0 To MaxSegment
			Call hrlab()
		Next WhichSegment
		WhichSegment = 0
		Hour_Renamed(TotalMan) = 0
		
		For x = 0 To 8
			Hour_Renamed(TotalMan) = Hour_Renamed(TotalMan) + Hour_Renamed(x)
		Next x
		
		Hour_Renamed(TotalMan) = Hour_Renamed(TotalMan) + Hour_Renamed(11)
		
		If shdy > 0 Then DrySize = ((Hour_Renamed(TotalMan) / hrsh) / shdy) * 125
		
		If CellValues(Building, 4, WhichSegment).Changed = False Then
			CellValues(Building, 4, WhichSegment).Value = System.Math.Round(((DrySize) ^ 0.5) * 1.4, 1)
		End If
		
		If CellValues(Building, 5, WhichSegment).Changed = False Then
			If CellValues(Building, 4, WhichSegment).Value > 0 Then CellValues(Building, 5, WhichSegment).Value = System.Math.Round(DrySize / CellValues(Building, 4, WhichSegment).Value, 1)
		End If
		
		If CellValues(Building, 6, WhichSegment).Changed = False Then
			CellValues(Building, 6, WhichSegment).Value = 12
		End If
		
		Select Case ot + wt
			Case Is <= 5000
				If CellValues(Building, 7, WhichSegment).Changed = False Then
					CellValues(Building, 7, WhichSegment).Word = "Mobile"
				End If
			Case Is <= 20000
				If CellValues(Building, 7, WhichSegment).Changed = False Then
					CellValues(Building, 7, WhichSegment).Word = "Wood Frame/Steel Siding"
				End If
			Case Else
				If CellValues(Building, 7, WhichSegment).Changed = False Then
					CellValues(Building, 7, WhichSegment).Word = "Permanent Steel"
				End If
		End Select
		
		'=============================== Office Cost ================================'
		
		TotalAdmin = 0
		TempAdmin = 0
		
		For TestSegment = 0 To MaxSegment
			TempAdmin = 0
			For x = 0 To 11
				TempAdmin = TempAdmin + CellValues(Staff, x, TestSegment).Value
			Next x
			If TempAdmin > TotalAdmin Then TotalAdmin = TempAdmin
		Next TestSegment
		
		'UPGRADE_WARNING: Couldn't resolve default property of object OfficeSize. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		OfficeSize = TotalAdmin * 275
		
		If CellValues(Building, 8, WhichSegment).Changed = False Then
			CellValues(Building, 8, WhichSegment).Value = System.Math.Round(((OfficeSize) ^ 0.5) * 1.4, 1)
		End If
		
		If CellValues(Building, 9, WhichSegment).Changed = False And CellValues(Building, 8, WhichSegment).Value <> 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object OfficeSize. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CellValues(Building, 9, WhichSegment).Value = System.Math.Round(OfficeSize / CellValues(Building, 8, WhichSegment).Value, 1)
		End If
		
		If CellValues(Building, 10, WhichSegment).Changed = False Then
			CellValues(Building, 10, WhichSegment).Value = 12
			
		End If
		Select Case ot + wt
			Case Is <= 5000
				If CellValues(Building, 11, WhichSegment).Changed = False Then
					CellValues(Building, 11, WhichSegment).Word = "Mobile"
				End If
			Case Is <= 20000
				If CellValues(Building, 11, WhichSegment).Changed = False Then
					CellValues(Building, 11, WhichSegment).Word = "Wood Frame/Steel Siding"
				End If
			Case Else
				If CellValues(Building, 11, WhichSegment).Changed = False Then
					CellValues(Building, 11, WhichSegment).Word = "Permanent Steel"
				End If
		End Select
		
		'============================= Warehouse Cost ==============================='
		
		WarehouseSize = 0
		
		WarehouseSize = (EqCost(Waste, Loader, OutNumber) * EqDefault(Loader, aw, Area) * 3.4) + (EqCost(Ore, Loader, OutNumber) * EqDefault(Loader, ao, Area) * 3.4)
		WarehouseSize = WarehouseSize + ((EqCost(Waste, Shovel, OutNumber) * EqDefault(Shovel, bw, TrackArea) * 4.9) + (EqCost(Ore, Shovel, OutNumber) * EqDefault(Shovel, bo, TrackArea) * 4.9))
		WarehouseSize = WarehouseSize + ((EqCost(Waste, CableShovel, OutNumber) * EqDefault(CableShovel, dw, TrackArea) * 4.4) + (EqCost(Ore, CableShovel, OutNumber) * EqDefault(CableShovel, dore, TrackArea) * 4.4))
		WarehouseSize = WarehouseSize + ((EqCost(Waste, Dragline, OutNumber) + EqDefault(Dragline, ew, TrackArea) * 0.8) + (EqCost(Ore, Dragline, OutNumber) * EqDefault(Dragline, eo, TrackArea) * 0.8))
		WarehouseSize = WarehouseSize + ((CellValues(Convey, 7, WhichSegment).Value + CellValues(Convey, 18, WhichSegment).Value) * 0.8)
		
		If WarehouseSize > 0 Then
			If CellValues(Building, 12, WhichSegment).Changed = False Then
				CellValues(Building, 12, WhichSegment).Value = System.Math.Round(((WarehouseSize) ^ 0.5) * 1.4, 1)
			End If
		End If
		
		If CellValues(Building, 12, WhichSegment).Value <> 0 Then
			If CellValues(Building, 13, WhichSegment).Changed = False Then
				CellValues(Building, 13, WhichSegment).Value = System.Math.Round(WarehouseSize / CellValues(Building, 12, WhichSegment).Value, 1)
			End If
		End If
		
		If CellValues(Building, 14, WhichSegment).Changed = False Then
			CellValues(Building, 14, WhichSegment).Value = 18
		End If
		Select Case ot + wt
			Case Is <= 5000
				If CellValues(Building, 15, WhichSegment).Changed = False Then
					CellValues(Building, 15, WhichSegment).Word = "Mobile"
				End If
			Case Is <= 20000
				If CellValues(Building, 15, WhichSegment).Changed = False Then
					CellValues(Building, 15, WhichSegment).Word = "Wood Frame/Steel Siding"
				End If
			Case Else
				If CellValues(Building, 15, WhichSegment).Changed = False Then
					CellValues(Building, 15, WhichSegment).Word = "Permanent Steel"
				End If
		End Select
		
	End Sub
	Public Sub PumpEngr()
		Dim x As Short
		Dim JumpOutOfThis As Short
		Dim TempHeadLoss As Decimal
		Dim TempHorizontal As Decimal
		Dim TempSurfaceLoss As Decimal
		Dim TempSurfaceHor As Decimal
		Dim TempDiameter As Decimal
		Dim Pipeit As Boolean
		
		On Error Resume Next
		
		TempHeadLoss = 0
		TempHorizontal = 0
		TempSurfaceLoss = 0
		TempSurfaceHor = 0
		JumpOutOfThis = 0
		Pipeit = False
		
		For x = 0 To 2 Step 2
			TempHeadLoss = TempHeadLoss + (CellValues(Haul, x, WhichSegment).Value * (CellValues(Haul, x + 1, WhichSegment).Value / 100))
			TempSurfaceHor = TempSurfaceHor + CellValues(Haul, x, WhichSegment).Value
		Next x
		
		For x = 8 To 14 Step 2
			TempHeadLoss = TempHeadLoss + (CellValues(Haul, x, WhichSegment).Value * (CellValues(Haul, x + 1, WhichSegment).Value / 100))
			TempSurfaceHor = TempSurfaceHor + CellValues(Haul, x, WhichSegment).Value
		Next x
		
		TempHorizontal = ((CellValues(Pit, 1, WhichSegment).Value + CellValues(Pit, 2, WhichSegment).Value) / 2)
		TempHorizontal = TempHorizontal + (CellValues(Pit, 0, WhichSegment).Value / (System.Math.Tan(CellValues(Pit, 3, WhichSegment).Value * (pi / 180))))
		
		TempSurfaceLoss = CellValues(Pit, 0, WhichSegment).Value - TempHeadLoss
		
		If TempSurfaceLoss > 0 Then Pipeit = True
		
		TempSurfaceHor = TempSurfaceHor - TempHorizontal
		
		If CellValues(Pumping, 1, WhichSegment).Changed = False Then
			CellValues(Pumping, 1, WhichSegment).Value = CellValues(Pit, 0, WhichSegment).Value
		End If
		
		If CellValues(Pumping, 2, WhichSegment).Changed = False Then
			CellValues(Pumping, 2, WhichSegment).Value = TempHorizontal
		End If
		
		If CellValues(Pumping, 3, WhichSegment).Changed = False Then
			CellValues(Pumping, 3, WhichSegment).Value = Int(CellValues(Pumping, 1, WhichSegment).Value / 250) + 1
		End If
		
		If CellValues(Pumping, 5, WhichSegment).Changed = False Then
			If TempSurfaceLoss > 0 Then
				CellValues(Pumping, 5, WhichSegment).Value = TempSurfaceLoss
			Else
				CellValues(Pumping, 5, WhichSegment).Value = 0
			End If
		End If
		
		If CellValues(Pumping, 6, WhichSegment).Changed = False Then
			If TempSurfaceHor > 0 Then
				CellValues(Pumping, 6, WhichSegment).Value = TempSurfaceHor
			End If
		End If
		
		TempDiameter = (((CellValues(Pumping, 0, WhichSegment).Value / 60) * 0.1337) / 4) * 144
		TempDiameter = ((4 * TempDiameter) / pi) ^ 0.5
		
		If CellValues(Pumping, 4, WhichSegment).Changed = False Then
			If UnitType = Metric Then
				Select Case TempDiameter
					Case Is <= 1
						CellValues(Pumping, 4, WhichSegment).Word = "2.54 centimeter"
					Case Is <= 2
						CellValues(Pumping, 4, WhichSegment).Word = "5.08 centimeter"
					Case Is <= 3
						CellValues(Pumping, 4, WhichSegment).Word = "7.62 centimeter"
					Case Is <= 4
						CellValues(Pumping, 4, WhichSegment).Word = "10.16 centimeter"
					Case Is <= 6
						CellValues(Pumping, 4, WhichSegment).Word = "15.24 centimeter"
					Case Is <= 8
						CellValues(Pumping, 4, WhichSegment).Word = "20.32 centimeter"
					Case Is <= 10
						CellValues(Pumping, 4, WhichSegment).Word = "25.40 centimeter"
					Case Is <= 12
						CellValues(Pumping, 4, WhichSegment).Word = "30.48 centimeter"
				End Select
			Else
				Select Case TempDiameter
					Case Is <= 1
						CellValues(Pumping, 4, WhichSegment).Word = "1 inch"
					Case Is <= 2
						CellValues(Pumping, 4, WhichSegment).Word = "2 inch"
					Case Is <= 3
						CellValues(Pumping, 4, WhichSegment).Word = "3 inch"
					Case Is <= 4
						CellValues(Pumping, 4, WhichSegment).Word = "4 inch"
					Case Is <= 6
						CellValues(Pumping, 4, WhichSegment).Word = "6 inch"
					Case Is <= 8
						CellValues(Pumping, 4, WhichSegment).Word = "8 inch"
					Case Is <= 10
						CellValues(Pumping, 4, WhichSegment).Word = "10 inch"
					Case Is <= 12
						CellValues(Pumping, 4, WhichSegment).Word = "12 inch"
				End Select
			End If
		End If
		
		If CellValues(Pumping, 7, WhichSegment).Changed = False And Pipeit = True Then
			CellValues(Pumping, 7, WhichSegment).Word = CellValues(Pumping, 4, WhichSegment).Word
		ElseIf Pipeit = False Then 
			CellValues(Pumping, 7, WhichSegment).Word = "None Required"
		End If
		
	End Sub
	Public Sub RoadEngr()
		Dim x As Short
		Dim y As Short
		Dim co As Short
		Dim cw As Short
		Dim WhichOreMachine As Short
		Dim WhichWasteMachine As Short
		
		On Error Resume Next
		
		co = Int(CellValues(EquipmentOne, 4, WhichSegment).Value)
		WhichOreMachine = Truck
		If co = 0 Then
			co = Int(CellValues(EquipmentOne, 20, WhichSegment).Value)
			WhichOreMachine = Scraper
		End If
		If co = 0 Then
			co = Int(CellValues(EquipmentOne, 21, WhichSegment).Value)
			WhichOreMachine = Articulated
		End If
		If co = 0 Then
			co = Int(CellValues(EquipmentOne, 24, WhichSegment).Value)
			WhichOreMachine = Conveyor
		End If
		
		cw = Int(CellValues(EquipmentOne, 9, WhichSegment).Value)
		WhichWasteMachine = Truck
		If cw = 0 Then
			cw = Int(CellValues(EquipmentOne, 25, WhichSegment).Value)
			WhichWasteMachine = Scraper
		End If
		If cw = 0 Then
			cw = Int(CellValues(EquipmentOne, 26, WhichSegment).Value)
			WhichWasteMachine = Articulated
		End If
		If cw = 0 Then
			cw = Int(CellValues(EquipmentOne, 29, WhichSegment).Value)
			WhichWasteMachine = Conveyor
		End If
		If cw = 0 And Int(CellValues(EquipmentOne, 8, WhichSegment).Value) <> 0 Then
			WhichWasteMachine = Dragline
		End If
		
		For x = 0 To 10 Step 2
			If x <= 2 Then
				y = x
			Else
				y = x + 4
			End If
			If CellValues(Haul, y, WhichSegment).Value > 0 Then
				If CellValues(Road, x, WhichSegment).Changed = False Then
					If WhichOreMachine = Conveyor Then
						CellValues(Road, x, WhichSegment).Value = 18
					Else
						CellValues(Road, x, WhichSegment).Value = (EqDefault(WhichOreMachine, co, TruckWidth) * 3.2)
						If CellValues(Road, x, WhichSegment).Value < 24 Then CellValues(Road, x, WhichSegment).Value = 24
					End If
				End If
				If CellValues(Road, x + 1, WhichSegment).Changed = False Then
					If WhichOreMachine = Conveyor Then
						CellValues(Road, x + 1, WhichSegment).Value = 3
					Else
						CellValues(Road, x + 1, WhichSegment).Value = (0.748821 * (EqDefault(WhichOreMachine, co, WeightCap) ^ 0.580262))
					End If
				End If
			Else
				CellValues(Road, x, WhichSegment).Value = 0
			End If
			
			If x <= 2 Then
				y = x + 4
			Else
				y = x + 16
			End If
			If CellValues(Haul, y, WhichSegment).Value > 0 Then
				If CellValues(Road, x + 12, WhichSegment).Changed = False Then
					If WhichWasteMachine = Conveyor Or WhichWasteMachine = Dragline Then
						CellValues(Road, x + 12, WhichSegment).Value = 18
					Else
						CellValues(Road, x + 12, WhichSegment).Value = (EqDefault(WhichWasteMachine, cw, TruckWidth) * 3.2)
						If CellValues(Road, x + 12, WhichSegment).Value < 24 Then CellValues(Road, x + 12, WhichSegment).Value = 24
					End If
				End If
				If CellValues(Road, x + 13, WhichSegment).Changed = False Then
					If WhichWasteMachine = Conveyor Or WhichWasteMachine = Dragline Then
						CellValues(Road, x + 13, WhichSegment).Value = 3
					Else
						CellValues(Road, x + 13, WhichSegment).Value = (0.748821 * (EqDefault(WhichWasteMachine, cw, WeightCap) ^ 0.580262))
					End If
				End If
			Else
				CellValues(Road, x + 12, WhichSegment).Value = 0
			End If
		Next x
		
	End Sub
	Public Sub BlastEngr()
		Dim x As Object
		Dim Volume As Decimal
		Dim obh As Decimal
		Dim wbh As Decimal
		Dim ot As Decimal
		Dim wt As Decimal
		Dim opf As Decimal
		Dim wpf As Decimal
		
		On Error Resume Next
		
		Call bhcal(obh, wbh)
		Call otcal(ot, wt)
		Call pfcal(opf, wpf)
		
		If CellValues(Powder, 0, WhichSegment).Changed = False Then
			CellValues(Powder, 0, WhichSegment).Value = (ot * opf)
		End If
		
		If CellValues(Supply, 8, WhichSegment).Value <> 0 Then
			Volume = Int(CellValues(Powder, 0, WhichSegment).Value / (62.4 * CellValues(Supply, 8, WhichSegment).Value))
			Volume = Int(Volume + (CellValues(Powder, 3, WhichSegment).Value / 12) + (CellValues(Powder, 5, WhichSegment).Value / 12))
		End If
		
		CellValues(Powder, 12, WhichSegment).Value = Volume
		
		If CellValues(Powder, 1, WhichSegment).Changed = False Then
			CellValues(Powder, 1, WhichSegment).Value = 7
		End If
		
		If CellValues(Powder, 2, WhichSegment).Changed = False Then
			If CellValues(Powder, 0, WhichSegment).Value <> 0 Then
				CellValues(Powder, 2, WhichSegment).Value = (Int((EqCost(Ore, Percussion, OutFeet) + EqCost(Ore, Rotary, OutFeet)) / (obh * 1.15)) + 2)
			Else
				CellValues(Powder, 2, WhichSegment).Value = 0
			End If
		End If
		
		If CellValues(Powder, 4, WhichSegment).Changed = False Then
			If CellValues(Supply, 4, WhichSegment).Value <> 0 Then
				CellValues(Powder, 4, WhichSegment).Value = Int((EqCost(Ore, Percussion, OutFeet) + EqCost(Ore, Rotary, OutFeet)) / (obh * 1.15))
			Else
				CellValues(Powder, 4, WhichSegment).Value = 0
			End If
		End If
		
		If CellValues(Powder, 6, WhichSegment).Changed = False Then
			CellValues(Powder, 6, WhichSegment).Value = (wt * wpf)
		End If
		
		If CellValues(Supply, 18, WhichSegment).Value <> 0 Then
			Volume = (CellValues(Powder, 6, WhichSegment).Value / (62.4 * CellValues(Supply, 18, WhichSegment).Value))
			Volume = (Volume + (CellValues(Powder, 8, WhichSegment).Value / 12) + (CellValues(Powder, 10, WhichSegment).Value / 12))
		End If
		
		'If Volume < 11000 Then
		'  Volume = 30
		'Else
		'  Volume = Int((11000 / Volume) * 30)
		'End If
		
		CellValues(Powder, 13, WhichSegment).Value = Volume
		
		If CellValues(Powder, 7, WhichSegment).Changed = False Then
			CellValues(Powder, 7, WhichSegment).Value = 7
		End If
		
		If CellValues(Powder, 8, WhichSegment).Changed = False Then
			If CellValues(Powder, 6, WhichSegment).Value <> 0 And wbh <> 0 Then
				CellValues(Powder, 8, WhichSegment).Value = (Int((EqCost(Waste, Percussion, OutFeet) + EqCost(Waste, Rotary, OutFeet)) / (wbh * 1.15)) + 2)
			Else
				CellValues(Powder, 8, WhichSegment).Value = 0
			End If
		End If
		
		If CellValues(Powder, 10, WhichSegment).Changed = False Then
			If CellValues(Supply, 16, WhichSegment).Value <> 0 And wbh <> 0 Then
				CellValues(Powder, 10, WhichSegment).Value = Int((EqCost(Waste, Percussion, OutFeet) + EqCost(Waste, Rotary, OutFeet)) / (wbh * 1.15))
			Else
				CellValues(Powder, 10, WhichSegment).Value = 0
			End If
		End If
		
		For x = 3 To 11 Step 2
			'UPGRADE_WARNING: Couldn't resolve default property of object x. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If x <> 7 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object x. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If CellValues(Powder, x, WhichSegment).Changed = False Then
					'UPGRADE_WARNING: Couldn't resolve default property of object x. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					CellValues(Powder, x, WhichSegment).Value = 7
				End If
			End If
		Next x
	End Sub
	Public Sub ElectEngr()
		Dim DoubleSwitch As Object
        Dim a As Object = Nothing
        Dim cw As Object
		Dim co As Object
		Dim KvaStuff As Decimal
		Dim tempvalue As Decimal
		Dim NumStations As Short
		Dim SubStationPrice As Decimal
		Dim SingleSwitch As Decimal
		Dim dore As Short
		Dim dw As Short
		Dim eo As Short
		Dim ew As Short
		Dim WireSize As Decimal
		Dim WirePrice As Decimal
		Dim SumItUp As Decimal
		Dim x As Short
		Dim TestSegment As Decimal
		
		On Error Resume Next
		
		dore = Int(CellValues(EquipmentOne, 2, WhichSegment).Value)
		dw = Int(CellValues(EquipmentOne, 7, WhichSegment).Value)
		eo = Int(CellValues(EquipmentOne, 3, WhichSegment).Value)
		ew = Int(CellValues(EquipmentOne, 8, WhichSegment).Value)
		'UPGRADE_WARNING: Couldn't resolve default property of object co. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		co = Int(CellValues(EquipmentOne, 4, WhichSegment).Value)
		'UPGRADE_WARNING: Couldn't resolve default property of object cw. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		cw = Int(CellValues(EquipmentOne, 9, WhichSegment).Value)
		
		KvaStuff = ((dore * 500) + 500) * EqCost(Ore, CableShovel, OutNumber)
		KvaStuff = KvaStuff + (((dw * 500) + 500) * EqCost(Waste, CableShovel, OutNumber))
		KvaStuff = KvaStuff + (eo * 2500 * EqCost(Ore, Dragline, OutNumber))
		KvaStuff = KvaStuff + (ew * 2500 * EqCost(Waste, Dragline, OutNumber))
		KvaStuff = KvaStuff + ((CellValues(Convey, 5, WhichSegment).Value + CellValues(Convey, 5, WhichSegment).Value) * 0.746)
		KvaStuff = KvaStuff + (CellValues(Electricity, 15, WhichSegment).Value + CellValues(Electricity, 17, WhichSegment).Value)
		If (EqCost(Ore, CableShovel, OutNumber) + EqCost(Waste, CableShovel, OutNumber) + EqCost(Ore, Dragline, OutNumber) + EqCost(Waste, Dragline, OutNumber)) > 0 Then
			KvaStuff = KvaStuff * ((EqCost(Ore, CableShovel, OutNumber) + EqCost(Waste, CableShovel, OutNumber) + EqCost(Ore, Dragline, OutNumber) + EqCost(Waste, Dragline, OutNumber)) ^ -0.5)
		End If
		
		Select Case KvaStuff
			Case Is > 99999
				KvaStuff = (System.Math.Round(KvaStuff / 10000)) * 10000
			Case Is > 9999
				KvaStuff = (System.Math.Round(KvaStuff / 1000)) * 1000
			Case Is > 999
				KvaStuff = (System.Math.Round(KvaStuff / 100)) * 100
			Case Is > 99
				KvaStuff = (System.Math.Round(KvaStuff / 10)) * 10
		End Select
		
		
		'Work on This
		Select Case KvaStuff
			Case Is <= 1000
				WireSize = 2
			Case Is <= 3500
				WireSize = 8
			Case Else
				WireSize = 15
		End Select
		
		If CellValues(Electrical, 8, WhichSegment).Changed = True Then
			WireSize = CellValues(Electrical, 8, WhichSegment).Value
		End If
		
		'Update These Costs - Averages from MCS 2014
		Select Case WireSize
			Case 2
				WirePrice = 31.36
			Case 8
				WirePrice = 33.83
			Case 15
				WirePrice = 61.44
		End Select
		
		If CellValues(Electrical, 0, WhichSegment).Changed = False Then
			CellValues(Electrical, 0, WhichSegment).Value = KvaStuff
		End If
		
		KvaStuff = 0
		
		Select Case CellValues(Building, 0, WhichSegment).Value * CellValues(Building, 1, WhichSegment).Value
			Case 0 To 1999
				KvaStuff = 0
			Case 2000 To 7999
				KvaStuff = 300
			Case 8000 To 13799
				KvaStuff = 750
			Case Else
				KvaStuff = 1500
		End Select
		
		If CellValues(Electrical, 1, WhichSegment).Changed = False Then
			CellValues(Electrical, 1, WhichSegment).Value = KvaStuff
		End If
		
		If CellValues(Development, 8, WhichSegment).Changed = False Then
			CellValues(Development, 8, WhichSegment).Value = (CellValues(Electrical, 0, WhichSegment).Value + CellValues(Electrical, 1, WhichSegment).Value)
		End If
		
		KvaStuff = 0
		
		KvaStuff = (CellValues(Electrical, 0, WhichSegment).Value + CellValues(Electrical, 1, WhichSegment).Value) * 1.25
		
		If CellValues(Development, 8, WhichSegment).Changed = False Then
			CellValues(Development, 8, WhichSegment).Value = (CellValues(Electrical, 0, WhichSegment).Value + CellValues(Electrical, 1, WhichSegment).Value) * 1.25
		End If
		
		NumStations = 1
		
		'UPGRADE_WARNING: Couldn't resolve default property of object a. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		While a = 0
			KvaStuff = KvaStuff / NumStations
			If KvaStuff > 10000 Then
				NumStations = NumStations + 1
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object a. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				a = 1
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object a. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If NumStations > 1000 Then a = 1
		End While
		
		Select Case KvaStuff
			Case Is <= 150
				tempvalue = 150
			Case Is <= 300
				tempvalue = 300
			Case Is <= 500
				tempvalue = 500
			Case Is <= 750
				tempvalue = 750
			Case Is <= 1000
				tempvalue = 1000
			Case Is <= 1500
				tempvalue = 1500
			Case Is <= 5000
				tempvalue = 5000
			Case Is <= 10000
				tempvalue = 10000
		End Select
		
		If CellValues(Electrical, 2, WhichSegment).Changed = False Then
			CellValues(Electrical, 2, WhichSegment).Value = tempvalue
		End If
		
		If CellValues(Electrical, 3, WhichSegment).Changed = False Then
			CellValues(Electrical, 3, WhichSegment).Value = CDec(NumStations)
		End If
		
		'Update these costs 2014
		
		Select Case CellValues(Electrical, 2, WhichSegment).Value
			Case 0 To 24
				SubStationPrice = 0
			Case 25 To 150
				SubStationPrice = 25300 * 1.8
			Case 151 To 300
				SubStationPrice = 32150 * 1.7
			Case 301 To 500
				SubStationPrice = 44800 * 1.6
			Case 501 To 750
				SubStationPrice = 70250 * 1.5
			Case 751 To 1000
				SubStationPrice = 90000 * 1.45
			Case 1001 To 1500
				SubStationPrice = 127200 * 1.4
			Case 1501 To 5000
				SubStationPrice = 279800 * 1.35
			Case Else
				SubStationPrice = 475900 * 1.3
		End Select
		
		If CellValues(Electrical, 4, WhichSegment).Changed = False Then
			CellValues(Electrical, 4, WhichSegment).Value = SubStationPrice
		End If
		
		SingleSwitch = 0
		
		SingleSwitch = EqCost(Ore, CableShovel, OutNumber) + EqCost(Waste, CableShovel, OutNumber)
		SingleSwitch = SingleSwitch + (EqCost(Ore, Dragline, OutNumber) + EqCost(Waste, Dragline, OutNumber))
		
		If CellValues(Electrical, 6, WhichSegment).Changed = False Then
			If SingleSwitch > 0 Then
				CellValues(Electrical, 6, WhichSegment).Value = SingleSwitch + 1
			Else
				CellValues(Electrical, 6, WhichSegment).Value = 0
			End If
		End If
		
		tempvalue = 0
		
		If CellValues(Electrical, 6, WhichSegment).Value <> 0 Then
			tempvalue = (CellValues(Electrical, 2, WhichSegment).Value * CellValues(Electrical, 3, WhichSegment).Value) / CellValues(Electrical, 6, WhichSegment).Value
		End If
		
		Select Case tempvalue
			Case Is > 99999
				tempvalue = (System.Math.Round(tempvalue / 10000)) * 10000
			Case Is > 9999
				tempvalue = (System.Math.Round(tempvalue / 1000)) * 1000
			Case Is > 999
				tempvalue = (System.Math.Round(tempvalue / 100)) * 100
			Case Is > 99
				tempvalue = (System.Math.Round(tempvalue / 10)) * 10
			Case Else
				tempvalue = 0
		End Select
		
		If CellValues(Electrical, 5, WhichSegment).Changed = False Then
			CellValues(Electrical, 5, WhichSegment).Value = tempvalue
		End If
		
		SingleSwitch = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object DoubleSwitch. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		DoubleSwitch = 0
		
		'Update These Costs 2014
		If (EqCost(Ore, CableShovel, OutNumber) + EqCost(Waste, CableShovel, OutNumber)) > 0 Then
			SingleSwitch = 31000
		End If
		
		If (EqCost(Ore, Dragline, OutNumber) + EqCost(Waste, Dragline, OutNumber)) > 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object DoubleSwitch. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			DoubleSwitch = 48650
		End If
		
		If CellValues(Electrical, 7, WhichSegment).Changed = False Then
			'UPGRADE_WARNING: Couldn't resolve default property of object DoubleSwitch. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If DoubleSwitch > 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object DoubleSwitch. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				CellValues(Electrical, 7, WhichSegment).Value = DoubleSwitch
			Else
				CellValues(Electrical, 7, WhichSegment).Value = SingleSwitch
			End If
		End If
		
		If CellValues(Electrical, 8, WhichSegment).Changed = False Then
			CellValues(Electrical, 8, WhichSegment).Value = WireSize
		End If
		
		tempvalue = 0
		
		tempvalue = (EqCost(Ore, CableShovel, OutNumber) + EqCost(Ore, Dragline, OutNumber)) * CellValues(Haul, 0, WhichSegment).Value
		tempvalue = tempvalue + (EqCost(Waste, CableShovel, OutNumber) + EqCost(Waste, Dragline, OutNumber)) * CellValues(Haul, 4, WhichSegment).Value
		If tempvalue > 0 Then
			tempvalue = tempvalue + CellValues(Haul, 2, WhichSegment).Value + CellValues(Haul, 6, WhichSegment).Value
		End If
		tempvalue = tempvalue + CellValues(Haul, 8, WhichSegment).Value
		
		If CellValues(Electrical, 9, WhichSegment).Changed = False Then
			CellValues(Electrical, 9, WhichSegment).Value = tempvalue
		End If
		
		If CellValues(Electrical, 10, WhichSegment).Changed = False Then
			CellValues(Electrical, 10, WhichSegment).Value = WirePrice
		End If
		
		SumItUp = 0
		For x = 3 To 9 Step 3
			SumItUp = SumItUp + (CellValues(Electrical, x, WhichSegment).Value * CellValues(Electrical, x + 1, WhichSegment).Value)
		Next x
		
		SumItUp = SumItUp + (CellValues(Electrical, 13, WhichSegment).Value * CellValues(Electrical, 14, WhichSegment).Value)
		
		If CellValues(Development, 19, WhichSegment).Changed = False Then
			CellValues(Development, 19, WhichSegment).Value = SumItUp
		End If
		
		If WhichSegment = 0 Then
			CellValues(DevelopmentResult, 4, 1).Value = SumItUp
		End If
		
		For TestSegment = 0 To MaxSegment
			If TestSegment = 0 Then
				EngBase(4) = CellValues(Development, 19, TestSegment).Value
			Else
				If CellValues(Development, 19, TestSegment).Value > EngBase(8) Then
					EngBase(4) = CellValues(Development, 19, TestSegment).Value
				End If
			End If
		Next TestSegment
		
	End Sub
	Public Sub ClearEngr()
		Dim TestSegment As Object
		Dim ClearVolume As Decimal
		Dim ClearDepth As Decimal
		Dim ClearArea As Decimal
		Dim TreeDiameter As Decimal
		Dim CutHours As Decimal
		Dim GrubHours As Decimal
		Dim ClearHours As Decimal
		Dim DozerHours As Decimal
		Dim TruckHours As Decimal
		Dim DozerLaborHours As Decimal
		Dim ClearLaborCost As Decimal
		Dim ClearMachineCost As Decimal
		Dim ClearTruckCost As Decimal
		Dim ClearOperatorCost As Decimal
		Dim laef As Decimal
		
		On Error Resume Next
		
		Call lbcal(laef)
		
		ClearVolume = (CellValues(Production, 4, WhichSegment).Value / CellValues(Deposit, 7, WhichSegment).Value) * 27
		ClearArea = ClearVolume ^ 0.33333
		ClearDepth = ClearArea / 3
		If ClearDepth <> 0 Then ClearArea = ClearVolume / ClearDepth
		
		If CellValues(Clearing, 0, WhichSegment).Changed = False Then
			CellValues(Clearing, 0, WhichSegment).Value = ClearArea / 43560
		End If
		
		If CellValues(Clearing, 2, WhichSegment).Changed = False Then
			CellValues(Clearing, 2, WhichSegment).Value = 20
		End If
		
		If CellValues(Clearing, 3, WhichSegment).Changed = False Then
			CellValues(Clearing, 3, WhichSegment).Value = 75
		End If
		
		If CellValues(Clearing, 4, WhichSegment).Changed = False Then
			CellValues(Clearing, 4, WhichSegment).Value = 25
		End If
		
		If CellValues(Clearing, 6, WhichSegment).Changed = False Then
			CellValues(Clearing, 6, WhichSegment).Word = "Bury On-Site"
		End If
		
		If CellValues(Development, 9, WhichSegment).Changed = False Then
			CellValues(Development, 9, WhichSegment).Value = CellValues(Clearing, 0, WhichSegment).Value
		End If
		
		TreeDiameter = 24 * (CellValues(Clearing, 1, WhichSegment).Value / 100)
		TreeDiameter = TreeDiameter + (12 * (CellValues(Clearing, 2, WhichSegment).Value / 100))
		TreeDiameter = TreeDiameter + (6 * (CellValues(Clearing, 3, WhichSegment).Value / 100))
		
		CutHours = 9.304785 * (TreeDiameter ^ 0.868483)
		GrubHours = 13.776482 * (TreeDiameter ^ 0.575121)
		ClearHours = 32 * ((CellValues(Clearing, 4, WhichSegment).Value / 100) + (CellValues(Clearing, 5, WhichSegment).Value / 100))
		
		ClearHours = ClearHours + CutHours + GrubHours
		DozerHours = 1.236893 * (TreeDiameter ^ 1.207519)
		
		Select Case LTrim(RTrim(LCase(CellValues(Clearing, 6, WhichSegment).Word)))
			Case "burn on-site"
				ClearHours = ClearHours * 1.1
			Case "bury on-site"
				ClearHours = ClearHours * 1.05
				DozerHours = DozerHours * 1.25
			Case "haul off-site"
				ClearHours = ClearHours * 1.25
				TruckHours = ClearHours * 0.25
		End Select
		
		ClearLaborCost = ClearHours * (CellValues(Wage, 7, WhichSegment).Value * (1 + (CellValues(Wage, 9, WhichSegment).Value / 100)))
		
		ClearMachineCost = (DozerHours * EqCost(Ore, Dozer, OutUnit))
		ClearMachineCost = ClearMachineCost + (TruckHours * EqCost(Ore, MainTruck, OutUnit))
		
		DozerLaborHours = DozerHours / laef
		
		ClearOperatorCost = DozerLaborHours * (CellValues(Wage, 4, WhichSegment).Value * (1 + (CellValues(Wage, 9, WhichSegment).Value / 100)))
		
		If CellValues(Development, 20, WhichSegment).Changed = False Then
			CellValues(Development, 20, WhichSegment).Value = ClearLaborCost + ClearMachineCost + ClearOperatorCost
		End If
		
		If WhichSegment = 0 Then
			CellValues(DevelopmentResult, 5, 1).Value = CellValues(Development, 9, WhichSegment).Value * CellValues(Development, 20, WhichSegment).Value
		End If
		
		For TestSegment = 0 To MaxSegment
			'UPGRADE_WARNING: Couldn't resolve default property of object TestSegment. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If TestSegment = 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object TestSegment. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				EngBase(5) = CellValues(Development, 20, TestSegment).Value
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object TestSegment. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If CellValues(Development, 20, TestSegment).Value > EngBase(8) Then
					'UPGRADE_WARNING: Couldn't resolve default property of object TestSegment. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					EngBase(5) = CellValues(Development, 20, TestSegment).Value
				End If
			End If
		Next TestSegment
		
	End Sub
	Public Sub SiteEngr()
        Dim cw As Object = Nothing
        Dim co As Object = Nothing
        Dim ot As Decimal
		Dim wt As Decimal
		Dim FenceLength As Decimal
		Dim FenceWidth As Decimal
		Dim YardSize As Decimal
		
		On Error Resume Next
		
		Call otcal(ot, wt)
		Call hrlab()
		
		If CellValues(Site, 0, WhichSegment).Changed = False Then
			CellValues(Site, 0, WhichSegment).Value = CellValues(WorkForce, 9, WhichSegment).Value
		End If
		
		If CellValues(Site, 1, WhichSegment).Changed = False Then
			CellValues(Site, 1, WhichSegment).Value = 300
		End If
		
		If CellValues(Site, 2, WhichSegment).Changed = False Then
			CellValues(Site, 2, WhichSegment).Value = CellValues(Site, 0, WhichSegment).Value * CellValues(Site, 1, WhichSegment).Value
		End If
		
		If CellValues(Development, 10, WhichSegment).Changed = False Then
			CellValues(Development, 10, WhichSegment).Value = CellValues(Site, 2, WhichSegment).Value
		End If
		
		If CellValues(Site, 4, WhichSegment).Changed = False Then
			If ot + wt < 20000 Then
				CellValues(Site, 4, WhichSegment).Word = "Gravel"
			Else
				CellValues(Site, 4, WhichSegment).Word = "Crushed Rock"
			End If
		End If
		
		YardSize = ((EqCost(Ore, Truck, OutNumber) * EqDefault(Truck, co, Area)) + (EqCost(Waste, Truck, OutNumber) * EqDefault(Truck, cw, Area)))
		YardSize = YardSize + ((EqCost(Ore, Articulated, OutNumber) * EqDefault(Articulated, co, Area)) + (EqCost(Waste, Articulated, OutNumber) * EqDefault(Articulated, cw, Area)))
		YardSize = YardSize + ((EqCost(Ore, Scraper, OutNumber) * EqDefault(Scraper, co, Area)) + (EqCost(Waste, Scraper, OutNumber) * EqDefault(Scraper, cw, Area)))
		
		If CellValues(Site, 5, WhichSegment).Changed = False Then
			CellValues(Site, 5, WhichSegment).Value = 100
		End If
		
		If CellValues(Site, 6, WhichSegment).Changed = False Then
			Select Case CellValues(Site, 0, WhichSegment).Value
				Case Is < 10
					CellValues(Site, 6, WhichSegment).Word = "Portable Self-Contained"
				Case Is < 50
					CellValues(Site, 6, WhichSegment).Word = "Septic System"
				Case Else
					CellValues(Site, 6, WhichSegment).Word = "Sewage Treatment Plant"
			End Select
		End If
		
		FenceLength = CellValues(Pit, 1, MaxSegment).Value + (2 * (CellValues(Pit, 0, MaxSegment).Value / (System.Math.Tan(CellValues(Pit, 3, MaxSegment).Value * (pi / 180)))))
		FenceWidth = CellValues(Pit, 2, MaxSegment).Value + (2 * (CellValues(Pit, 0, MaxSegment).Value / (System.Math.Tan(CellValues(Pit, 3, MaxSegment).Value * (pi / 180)))))
		FenceLength = (2 * FenceLength) + (2 * FenceWidth)
		
		If CellValues(Site, 7, WhichSegment).Changed = False Then
			CellValues(Site, 7, WhichSegment).Value = FenceLength
		End If
		
		If CellValues(Development, 24, WhichSegment).Changed = False Then
			CellValues(Development, 24, WhichSegment).Value = CellValues(Site, 7, WhichSegment).Value
		End If
		
		If CellValues(Site, 8, WhichSegment).Changed = False Then
			CellValues(Site, 8, WhichSegment).Value = 4
		End If
		
		If CellValues(Site, 9, WhichSegment).Changed = False Then
			CellValues(Site, 9, WhichSegment).Word = "Chain Link/Barbed Wire"
		End If
		
	End Sub
	Public Sub FuelEngr()
		Dim TempFuel As Decimal
		Dim r As Decimal
		Dim t As Decimal
		Dim x As Short
		Dim y As Short
		
		On Error Resume Next
		
		TempFuel = 0
		TempFuel = CellValues(EquipmentHours, 0, WhichSegment).Value * CellValues(Diesel, 0, WhichSegment).Value
		TempFuel = TempFuel + (CellValues(EquipmentHours, 1, WhichSegment).Value * CellValues(Diesel, 1, WhichSegment).Value)
		TempFuel = TempFuel + (CellValues(EquipmentHours, 4, WhichSegment).Value * CellValues(Diesel, 4, WhichSegment).Value)
		TempFuel = TempFuel + (CellValues(EquipmentHours, 5, WhichSegment).Value * CellValues(Diesel, 5, WhichSegment).Value)
		
		If CellValues(FuelStorage, 0, WhichSegment).Changed = False Then
			CellValues(FuelStorage, 0, WhichSegment).Value = TempFuel
		End If
		
		TempFuel = 0
		TempFuel = CellValues(EquipmentHours, 2, WhichSegment).Value * CellValues(Diesel, 2, WhichSegment).Value
		TempFuel = TempFuel + (CellValues(EquipmentHours, 3, WhichSegment).Value * CellValues(Diesel, 3, WhichSegment).Value)
		TempFuel = TempFuel + (CellValues(EquipmentHours, 18, WhichSegment).Value * CellValues(Diesel, 18, WhichSegment).Value)
		TempFuel = TempFuel + (CellValues(EquipmentHours, 19, WhichSegment).Value * CellValues(Diesel, 19, WhichSegment).Value)
		
		If CellValues(FuelStorage, 1, WhichSegment).Changed = False Then
			CellValues(FuelStorage, 1, WhichSegment).Value = TempFuel
		End If
		
		TempFuel = 0
		
		For x = 6 To 13
			TempFuel = TempFuel + (CellValues(EquipmentHours, x, WhichSegment).Value * CellValues(Diesel, x, WhichSegment).Value)
		Next x
		
		If CellValues(FuelStorage, 2, WhichSegment).Changed = False Then
			CellValues(FuelStorage, 2, WhichSegment).Value = TempFuel
		End If
		
		TempFuel = 0
		For x = 0 To 2
			TempFuel = TempFuel + CellValues(FuelStorage, x, WhichSegment).Value
		Next x
		
		If CellValues(FuelStorage, 3, WhichSegment).Changed = False Then
			CellValues(FuelStorage, 3, WhichSegment).Value = TempFuel
		End If
		
		OpBin(1) = OpBin(1) + (TempFuel * CellValues(Supply, 0, WhichSegment).Value)
		
		If CellValues(FuelStorage, 4, WhichSegment).Changed = False Then
			CellValues(FuelStorage, 4, WhichSegment).Value = 30
		End If
		
		If CellValues(FuelStorage, 5, WhichSegment).Changed = False Then
			CellValues(FuelStorage, 5, WhichSegment).Value = CellValues(FuelStorage, 3, WhichSegment).Value * CellValues(FuelStorage, 4, WhichSegment).Value
		End If
		
		r = CellValues(FuelStorage, 5, WhichSegment).Value / 15000
		x = 0
		t = 0
		While x = 0
			If r >= 1 Then
				t = t + 1
				If t > 0 Then
					r = CellValues(FuelStorage, 5, WhichSegment).Value / (t * 15000)
				End If
			Else
				x = 1
			End If
		End While
		
		If CellValues(FuelStorage, 7, WhichSegment).Changed = False Then
			CellValues(FuelStorage, 7, WhichSegment).Value = t
		End If
		
		If CellValues(FuelStorage, 7, WhichSegment).Value <> 0 Then
			TempFuel = Int((CellValues(FuelStorage, 5, WhichSegment).Value / CellValues(FuelStorage, 7, WhichSegment).Value) + 1)
		End If
		
		Select Case TempFuel
			Case Is < 1000
				TempFuel = 1000
			Case Is < 2000
				TempFuel = 2000
			Case Is < 5000
				TempFuel = 5000
			Case Is < 10000
				TempFuel = 10000
			Case Is < 12000
				TempFuel = 12000
			Case Is < 15000
				TempFuel = 15000
		End Select
		
		If CellValues(FuelStorage, 6, WhichSegment).Changed = False Then
			CellValues(FuelStorage, 6, WhichSegment).Value = TempFuel
		ElseIf CellValues(FuelStorage, 6, WhichSegment).Changed = True Then 
			If CellValues(FuelStorage, 7, WhichSegment).Changed = False And CellValues(FuelStorage, 6, WhichSegment).Value <> 0 Then
				CellValues(FuelStorage, 7, WhichSegment).Value = Int(CellValues(FuelStorage, 5, WhichSegment).Value / CellValues(FuelStorage, 6, WhichSegment).Value) + 1
			End If
		End If
		
	End Sub
	Public Sub NewEquipEngr()
		Dim x As Short
		Dim NewSegment As Short
		Dim OldSegment As Short
		Dim AddNumber As Short
		Dim AddPrice As Decimal
		
		On Error Resume Next
		
		For OldSegment = 0 To MaxSegment - 1
			NewSegment = OldSegment + 1
			For x = 0 To 31
				If CellValues(EquipmentTwo, x, NewSegment).Value > CellValues(EquipmentTwo, x, OldSegment).Value Then
					AddNumber = CellValues(EquipmentTwo, x, NewSegment).Value - CellValues(EquipmentTwo, x, OldSegment).Value
				ElseIf CellValues(EquipmentTwo, x, NewSegment).Value <= CellValues(EquipmentTwo, x, OldSegment).Value Then 
					AddNumber = 0
				End If
				Select Case x
					Case 0 To 3, 20
						If CellValues(EquipmentOne, x, NewSegment).Value > 0 Then
							AddPrice = AddNumber * CellValues(Purchase, 0, NewSegment).Value
							CellValues(Purchase, 20, NewSegment).Value = AddPrice
						End If
					Case 4, 21
						If CellValues(EquipmentOne, x, NewSegment).Value > 0 Then
							AddPrice = AddNumber * CellValues(Purchase, 1, NewSegment).Value
							CellValues(Purchase, 21, NewSegment).Value = AddPrice
						End If
					Case 5 To 8, 25
						If CellValues(EquipmentOne, x, NewSegment).Value > 0 Then
							AddPrice = AddNumber * CellValues(Purchase, 2, NewSegment).Value
							CellValues(Purchase, 22, NewSegment).Value = AddPrice
						End If
					Case 9, 26
						If CellValues(EquipmentOne, x, NewSegment).Value > 0 Then
							AddPrice = AddNumber * CellValues(Purchase, 3, NewSegment).Value
							CellValues(Purchase, 23, NewSegment).Value = AddPrice
						End If
					Case 10 To 19
						If CellValues(EquipmentOne, x, NewSegment).Value > 0 Then
							AddPrice = AddNumber * CellValues(Purchase, x - 6, NewSegment).Value
							CellValues(Purchase, x + 14, NewSegment).Value = AddPrice
						End If
					Case 22, 23
						If CellValues(EquipmentOne, x, NewSegment).Value > 0 Then
							AddPrice = AddNumber * CellValues(Purchase, 14, NewSegment).Value
							CellValues(Purchase, 34, NewSegment).Value = AddPrice
						End If
					Case 24
						AddPrice = AddNumber * CellValues(Purchase, 15, NewSegment).Value
						CellValues(Purchase, 35, NewSegment).Value = AddPrice
					Case 27, 28
						If CellValues(EquipmentOne, x, NewSegment).Value > 0 Then
							AddPrice = AddNumber * CellValues(Purchase, 16, NewSegment).Value
							CellValues(Purchase, 36, NewSegment).Value = AddPrice
						End If
					Case 29
						AddPrice = AddNumber * CellValues(Purchase, 17, NewSegment).Value
						CellValues(Purchase, 37, NewSegment).Value = AddPrice
					Case 30, 31
						If CellValues(EquipmentOne, x, NewSegment).Value > 0 Then
							AddPrice = AddNumber * CellValues(Purchase, x - 12, NewSegment).Value
							CellValues(Purchase, x + 8, NewSegment).Value = AddPrice
						End If
				End Select
			Next x
		Next OldSegment
	End Sub
	Public Sub ReplaceEngr()
		Dim y As Object
		Dim x As Short
		Dim NewSegment As Short
		Dim ThisSegment As Short
		Dim OldSegment As Short
		Dim AddNumber As Short
		Dim AddPrice As Decimal
		Dim YearOfReplace As Short
		Dim CostBin(25) As Decimal
		Dim StartAdd As Decimal
		Dim AddMonth As Decimal
		Dim OverHauls As Decimal
		Dim OneLife As Decimal
		Dim WhichReplace As Short
		
		On Error Resume Next
		
		Call TimeLineCalc()
		
		For x = 0 To 25
			CostBin(x) = 0
		Next x
		
		For ThisSegment = 0 To MaxSegment
			For x = 0 To 19
				For WhichReplace = 1 To MaxTime
					StartAdd = CellValues(Production, 15, ThisSegment).Value
					AddMonth = CellValues(Replace_Renamed, x, ThisSegment).Value
					OverHauls = CellValues(Replace_Renamed, 21, ThisSegment).Value
					OneLife = CellValues(Replace_Renamed, x, ThisSegment).Value
					If OneLife > 0 Then
						If ((OneLife + (AddMonth * OverHauls)) / 12) < 100 Then YearOfReplace = StartAdd + (WhichReplace * (Int((OneLife + (AddMonth * OverHauls)) / 12) + 1))
						If YearOfReplace < MaxTime Then
							CostBin(YearOfReplace) = CostBin(YearOfReplace) + CellValues(Purchase, x + 20, ThisSegment).Value
							'          CellValues(EquipmentPurchaseResult, x, YearOfReplace).Value = CellValues(EquipmentPurchaseResult, x, YearOfReplace).Value + CellValues(Purchase, x + 20, ThisSegment).Value
							'          CellValues(EquipmentNumberResult, x, YearOfReplace).Value = CellValues(EquipmentNumberResult, x, YearOfReplace).Value + (CellValues(Purchase, x + 20, ThisSegment).Value / CellValues(Purchase, x, ThisSegment).Value)
						End If
					End If
				Next WhichReplace
			Next x
		Next ThisSegment
		
		For y = CellValues(Production, 15, 0).Value To MaxTime
			If CellValues(Replace_Renamed, 20, 0).Value = 1 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object y. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				CellValues(Summary, 8, y).Value = CellValues(Summary, 8, y).Value + CostBin(y)
			End If
		Next y
		
	End Sub
	Public Sub MaxEquipEngr()
		Dim x As Short
		Dim NewSegment As Short
		Dim OldSegment As Short
		Dim AddNumber As Short
		Dim AddPrice As Decimal
		
		On Error Resume Next
		
		For OldSegment = 0 To MaxSegment
			For x = 0 To 31
				If CellValues(EquipmentOne, x, OldSegment).Value > CellValues(EquipmentOne, x, 0).Value Then
					CellValues(EquipmentOne, x, 0).Value = CellValues(EquipmentOne, x, OldSegment).Value
				End If
			Next x
		Next OldSegment
		
		For x = 0 To 31
			For NewSegment = 1 To MaxSegment
				CellValues(EquipmentOne, x, NewSegment).Changed = True
				CellValues(EquipmentOne, x, NewSegment).Value = CellValues(EquipmentOne, x, 0).Value
			Next NewSegment
		Next x
		
		Call CostItAll()
		
	End Sub
End Module