Option Strict Off
Option Explicit On
Module Calculations
	Public Sub HaulDesign()
		
		Dim x As Short
		Dim StarterDepth As Decimal
		Dim CompVolume As Decimal
		Dim Slope As Decimal
		Dim Depth As Decimal
		Dim Width As Decimal
		Dim Length As Decimal
		Dim WasteVolume As Decimal
		Dim WasteSwell As Decimal
		Dim WasteTons As Decimal
		Dim OreVolume As Decimal
		Dim OutSwitch As Short
		Dim OutDepth As Decimal
		Dim OutWidth As Decimal
		Dim LostWidth As Decimal
		Dim FloorWidth As Decimal
		Dim OutFloorWidth As Decimal
		Dim OutTopWidth As Decimal
		Dim TonsPerDay As Decimal
		Dim DumpFloorWidth As Decimal
		Dim DumpTopWidth As Decimal
		Dim DumpHeight As Decimal
		Dim xSegment As Decimal
		
		On Error Resume Next
		
		'Pit Dimensions
		
		For x = 16 To 19
			Select Case x
				Case 16
					If CellValues(Haul, x, WhichSegment).Changed = False Then CellValues(Haul, x, WhichSegment).Value = 82.5
					If CellValues(Haul, x + 12, WhichSegment).Changed = False Then CellValues(Haul, x + 12, WhichSegment).Value = 82.5
				Case 18
					If CellValues(Haul, x, WhichSegment).Changed = False Then CellValues(Haul, x, WhichSegment).Value = 86
					If CellValues(Haul, x + 12, WhichSegment).Changed = False Then CellValues(Haul, x + 12, WhichSegment).Value = 86
				Case Else
					If CellValues(Haul, x, WhichSegment).Changed = False Then CellValues(Haul, x, WhichSegment).Value = 92
					If CellValues(Haul, x + 12, WhichSegment).Changed = False Then CellValues(Haul, x + 12, WhichSegment).Value = 92
			End Select
		Next x
		
		OreVolume = 0
		WasteVolume = 0
		
		For xSegment = 0 To WhichSegment
			OreVolume = OreVolume + ((CellValues(Production, 17, xSegment).Value / CellValues(Deposit, 0, xSegment).Value) * 27)
			If CellValues(Production, 5, xSegment).Value <> 0 Then
				WasteTons = CellValues(Production, 17, xSegment).Value * (CellValues(Production, 10, xSegment).Value / CellValues(Production, 5, xSegment).Value)
			End If
			WasteVolume = WasteVolume + ((WasteTons / CellValues(Deposit, 7, xSegment).Value) * 27)
		Next xSegment
		
		WasteSwell = CellValues(Deposit, 8, WhichSegment).Value
		
		If CellValues(Supply, 11, WhichSegment).Changed = False Then
			CellValues(Supply, 11, WhichSegment).Changed = (OreVolume + WasteVolume) / 25
		End If
		
		If CellValues(Pit, 0, WhichSegment).Changed = False Then
			StarterDepth = ((OreVolume + WasteVolume) / 4) ^ 0.33333
		Else
			StarterDepth = CellValues(Pit, 0, WhichSegment).Value
		End If
		OutSwitch = 0
		Slope = 1 / (CDec(System.Math.Tan(CDbl(CellValues(Pit, 3, WhichSegment).Value * (pi / 180)))))
		Slope = (Slope * StarterDepth) * 2
		For Width = (Slope + 1) To 20000
			CompVolume = ((pi * StarterDepth * (((Width / 2) ^ 2) + (((Width - Slope) / 2) ^ 2) + ((Width / 2) * ((Width - Slope) / 2))))) / 3
			If CompVolume > (WasteVolume + OreVolume) And OutSwitch = 0 Then
				OutFloorWidth = Width - Slope
				OutTopWidth = Width
				OutSwitch = 1
			End If
		Next Width
		
		Call tonsp(TonsPerDay)
		
		If CellValues(Pit, 0, WhichSegment).Changed = False Then CellValues(Pit, 0, WhichSegment).Value = StarterDepth
		If CellValues(Pumping, 0, WhichSegment).Changed = False Then CellValues(Pumping, 0, WhichSegment).Value = (TonsPerDay / 50)
		If CellValues(Pumping, 1, WhichSegment).Changed = False Then CellValues(Pumping, 1, WhichSegment).Value = StarterDepth
		If CellValues(Pit, 1, WhichSegment).Changed = False Then CellValues(Pit, 1, WhichSegment).Value = OutFloorWidth * (4 / 3)
		If CellValues(Pit, 2, WhichSegment).Changed = False Then CellValues(Pit, 2, WhichSegment).Value = OutFloorWidth * (2 / 3)
		If CellValues(Pit, 8, WhichSegment).Changed = False Then CellValues(Pit, 8, WhichSegment).Value = OutTopWidth * (4 / 3)
		If CellValues(Pit, 9, WhichSegment).Changed = False Then CellValues(Pit, 9, WhichSegment).Value = 2
		If CellValues(Pit, 10, WhichSegment).Changed = False Then CellValues(Pit, 10, WhichSegment).Value = OutTopWidth * (2 / 3)
		If CellValues(Pit, 11, WhichSegment).Changed = False Then CellValues(Pit, 11, WhichSegment).Value = 2
		
		'Waste Pit Haul Profile Dimensions
		
		Slope = 1 / (CDec(System.Math.Tan(CDbl(CellValues(Pit, 3, WhichSegment).Value * (pi / 180)))))
		CompVolume = 0
		OutSwitch = 0
		For Depth = 1 To 20000
			LostWidth = (Depth * Slope) * 2
			CompVolume = (pi * Depth * (((OutTopWidth / 2) ^ 2) + (((OutTopWidth - LostWidth) / 2) ^ 2) + ((OutTopWidth / 2) * ((OutTopWidth - LostWidth) / 2)))) / 3
			If CompVolume > (WasteVolume / 2) And OutSwitch = 0 Then
				OutDepth = Depth
				FloorWidth = (OutTopWidth - LostWidth)
				WasteFloor(WhichSegment) = FloorWidth
				OutSwitch = 1
			End If
		Next Depth
		
		If CellValues(Haul, 4, WhichSegment).Changed = False Then CellValues(Haul, 4, WhichSegment).Value = FloorWidth / 2
		If CellValues(Haul, 5, WhichSegment).Changed = False Then CellValues(Haul, 5, WhichSegment).Value = 0
		If CellValues(Haul, 6, WhichSegment).Changed = False Then CellValues(Haul, 6, WhichSegment).Value = OutDepth / (CellValues(Pit, 7, WhichSegment).Value / 100)
		If CellValues(Haul, 7, WhichSegment).Changed = False Then CellValues(Haul, 7, WhichSegment).Value = CellValues(Pit, 7, WhichSegment).Value
		If CellValues(Haul, 20, WhichSegment).Changed = False Then CellValues(Haul, 20, WhichSegment).Value = CellValues(Pit, 10, WhichSegment).Value
		If CellValues(Haul, 21, WhichSegment).Changed = False Then CellValues(Haul, 21, WhichSegment).Value = CellValues(Pit, 11, WhichSegment).Value
		
		'Waste Dump Haul Profile Dimensions
		
		DumpFloorWidth = 0
		DumpTopWidth = 0
		Slope = 1
		DumpHeight = (((WasteVolume * (1 + (WasteSwell / 100))) / 4) ^ 0.33333) / 3
		OutSwitch = 0
		Slope = (Slope * DumpHeight) * 2
		For Width = (Slope + 1) To 20000
			CompVolume = ((pi * DumpHeight * (((Width / 2) ^ 2) + (((Width - Slope) / 2) ^ 2) + ((Width / 2) * ((Width - Slope) / 2))))) / 3
			If CompVolume > (WasteVolume * (1 + (WasteSwell / 100))) And OutSwitch = 0 Then
				DumpFloorWidth = Width - Slope
				DumpTopWidth = Width
				OutSwitch = 1
			End If
		Next Width
		
		If CellValues(Haul, 22, WhichSegment).Changed = False Then CellValues(Haul, 22, WhichSegment).Value = DumpHeight / (CellValues(Pit, 7, WhichSegment).Value / 100)
		If CellValues(Haul, 23, WhichSegment).Changed = False Then CellValues(Haul, 23, WhichSegment).Value = CellValues(Pit, 7, WhichSegment).Value
		If CellValues(Haul, 24, WhichSegment).Changed = False Then CellValues(Haul, 24, WhichSegment).Value = DumpTopWidth / 3
		If CellValues(Haul, 25, WhichSegment).Changed = False Then CellValues(Haul, 25, WhichSegment).Value = 0
		
		'Ore Haul Profile Dimensions
		
		OutDepth = 0
		FloorWidth = 0
		Slope = 1 / (CDec(System.Math.Tan(CDbl(CellValues(Pit, 3, WhichSegment).Value * (pi / 180)))))
		CompVolume = 0
		OutSwitch = 0
		For Depth = 1 To 20000
			LostWidth = (Depth * Slope) * 2
			CompVolume = (pi * Depth * (((OutTopWidth / 2) ^ 2) + (((OutTopWidth - LostWidth) / 2) ^ 2) + ((OutTopWidth / 2) * ((OutTopWidth - LostWidth) / 2)))) / 3
			If CompVolume > (WasteVolume + (OreVolume / 2)) And OutSwitch = 0 Then
				OutDepth = Depth
				FloorWidth = (OutTopWidth - LostWidth)
				OreFloor(WhichSegment) = FloorWidth
				OutSwitch = 1
			End If
		Next Depth
		
		If CellValues(Haul, 0, WhichSegment).Changed = False Then CellValues(Haul, 0, WhichSegment).Value = FloorWidth / 2
		If CellValues(Haul, 1, WhichSegment).Changed = False Then CellValues(Haul, 1, WhichSegment).Value = 0
		If CellValues(Haul, 2, WhichSegment).Changed = False Then CellValues(Haul, 2, WhichSegment).Value = OutDepth / (CellValues(Pit, 7, WhichSegment).Value / 100)
		If CellValues(Haul, 3, WhichSegment).Changed = False Then CellValues(Haul, 3, WhichSegment).Value = CellValues(Pit, 7, WhichSegment).Value
		If CellValues(Haul, 8, WhichSegment).Changed = False Then CellValues(Haul, 8, WhichSegment).Value = CellValues(Pit, 8, WhichSegment).Value
		If CellValues(Haul, 9, WhichSegment).Changed = False Then CellValues(Haul, 9, WhichSegment).Value = CellValues(Pit, 9, WhichSegment).Value
		
	End Sub
	Public Sub bhcal(ByRef OreBenchHeight As Decimal, ByRef WasteBenchHeight As Decimal)
		
		OreBenchHeight = CellValues(Deposit, 5, WhichSegment).Value
		WasteBenchHeight = CellValues(Deposit, 12, WhichSegment).Value
		
	End Sub
	Sub dcalc()
		Dim g As Object
		Dim f As Object
        Dim x As Object = Nothing
        Dim y As Object = Nothing
        Dim mnwg As Object
		Dim lawg As Object
		Dim blwg As Object
		
		Dim drwg As Decimal
		Dim hrsh As Decimal
		Dim olf As Decimal
		Dim wlf As Decimal
		Dim ElectricalRate As Decimal
		Dim opsg As Decimal
		Dim wpsg As Decimal
		Dim TaxRate As Decimal
		Dim obh As Decimal
		Dim wbh As Decimal
		Dim ot As Decimal
		Dim wt As Decimal
		Dim opf As Decimal
		Dim wpf As Decimal
		Dim prim As Decimal
		Dim btw As Decimal
		Dim rdw As Decimal
		Dim BurdenRate As Decimal
		Dim mcwg As Decimal
		Dim laef As Decimal
		Dim pow As Decimal
		Dim wpow As Decimal
		Dim cap As Decimal
		Dim det As Decimal
		Dim ppt As Decimal
		Dim TempMech As Decimal
		Dim DvFtDr As Decimal
		Dim DvPdVl As Decimal
		Dim DvPrNm As Decimal
		Dim DvCpNm As Decimal
		Dim AddToIt As Decimal
		Dim DevelopmentSupplyTemp As Decimal
		Dim DevelopmentLabor(9) As Decimal
		
		Dim jump As Boolean
		
		On Error Resume Next
		
		Call getout(jump)
		If jump = True Then Exit Sub
		
		Call hrcal(hrsh)
		Call lfcal(olf, wlf)
		Call elcal(ElectricalRate)
		Call sgcal(opsg, wpsg)
		Call txcal(TaxRate)
		Call bhcal(obh, wbh)
		Call otcal(ot, wt)
		Call pfcal(opf, wpf)
		Call prcal(prim, btw, rdw)
		Call mccal(BurdenRate, mcwg)
		Call lbcal(laef)
		Call excal(wpow, pow, cap, det)
		Call ptcal(ppt)
		
		drwg = CellValues(Wage, 0, WhichSegment).Value
		'UPGRADE_WARNING: Couldn't resolve default property of object blwg. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		blwg = CellValues(Wage, 1, WhichSegment).Value
		'UPGRADE_WARNING: Couldn't resolve default property of object lawg. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		lawg = CellValues(Wage, 7, WhichSegment).Value
		'UPGRADE_WARNING: Couldn't resolve default property of object mnwg. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mnwg = CellValues(Wage, 8, WhichSegment).Value
		
		If wt = 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object y. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			y = 15
			wt = ot
		End If
		
		DevelopmentLabor(1) = ((Int((EqCost(Waste, Percussion, OutHours) + EqCost(Waste, Rotary, OutHours)) / laef / hrsh) + 1) * hrsh * drwg * BurdenRate) * (ppt / wt)
		'UPGRADE_WARNING: Couldn't resolve default property of object blwg. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		DevelopmentLabor(2) = (EqCost(Ore, PowderBuggy, OutHours) * blwg * BurdenRate) * (ppt / wt)
		DevelopmentLabor(3) = (StripCost(Loader, OutOperator) + StripCost(Shovel, OutOperator) + StripCost(CableShovel, OutOperator) + StripCost(Dragline, OutOperator) + StripCost(Scraper, OutOperator))
		DevelopmentLabor(4) = StripCost(Truck, OutOperator) + StripCost(Articulated, OutOperator)
		DevelopmentLabor(5) = EqCost(Ore, Dozer, OutOperator) * (ppt / wt)
		DevelopmentLabor(6) = EqCost(Ore, WaterTanker, OutOperator) * (ppt / wt)
		DevelopmentLabor(7) = (StripCost(CableShovel, OutMechanicCost) + StripCost(Dragline, OutMechanicCost) + EqCost(Waste, Percussion, OutMechanicCost) + EqCost(Waste, Rotary, OutMechanicCost) + StripCost(Shovel, OutMechanicCost) + StripCost(Loader, OutMechanicCost) + StripCost(Scraper, OutMechanicCost))
		DevelopmentLabor(7) = DevelopmentLabor(7) + (StripCost(Truck, OutMechanicCost) + StripCost(Articulated, OutMechanicCost))
		DevelopmentLabor(7) = DevelopmentLabor(7) + ((EqCost(Ore, Light, OutMechanicTime) + EqCost(Ore, Pump, OutMechanicTime) + EqCost(Ore, Grader, OutMechanicTime)) * mcwg * BurdenRate * (ppt / wt))
		DevelopmentLabor(7) = DevelopmentLabor(7) + ((EqCost(Ore, Dozer, OutMechanicTime) + EqCost(Ore, WaterTanker, OutMechanicTime) + EqCost(Ore, PowderBuggy, OutMechanicTime) + EqCost(Ore, MainTruck, OutMechanicTime) + EqCost(Ore, Pickup, OutMechanicTime)) * mcwg * BurdenRate * (ppt / wt))
		If (mcwg * BurdenRate * (ppt / wt)) <> 0 Then
			TempMech = DevelopmentLabor(7) / (mcwg * BurdenRate * (ppt / wt))
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object lawg. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		DevelopmentLabor(8) = (Int((TempMech * (1.07187 * ((ot + wt) ^ -0.057219))) / hrsh) + 1) * hrsh * lawg * BurdenRate * (ppt / wt)
		'UPGRADE_WARNING: Couldn't resolve default property of object mnwg. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		DevelopmentLabor(9) = (Int((TempMech * (1.72809 * ((ot + wt) ^ -0.03769))) / hrsh) + 1) * hrsh * mnwg * BurdenRate * (ppt / wt)
		
		DevelopmentArray(1) = 0
		
		For x = 1 To 9
			'UPGRADE_WARNING: Couldn't resolve default property of object x. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			DevelopmentArray(1) = DevelopmentArray(1) + DevelopmentLabor(x)
		Next x
		
		DevelopmentArray(2) = StripCost(Loader, OutFuel) + StripCost(Shovel, OutFuel) + StripCost(Scraper, OutFuel) + StripCost(Truck, OutFuel) + StripCost(Articulated, OutFuel)
		DevelopmentArray(2) = DevelopmentArray(2) + ((EqCost(Waste, Percussion, OutFuel) + EqCost(Waste, Rotary, OutFuel)) * (ppt / wt))
		DevelopmentArray(2) = DevelopmentArray(2) + ((EqCost(Ore, Grader, OutFuel) + EqCost(Ore, Pickup, OutFuel) + EqCost(Ore, Light, OutFuel) + EqCost(Ore, Pump, OutFuel) + EqCost(Ore, MainTruck, OutFuel)) * (ppt / wt))
		DevelopmentArray(2) = DevelopmentArray(2) + ((EqCost(Ore, Dozer, OutFuel) + EqCost(Ore, WaterTanker, OutFuel) + EqCost(Ore, PowderBuggy, OutFuel)) * (ppt / wt))
		
		DevelopmentArray(3) = (StripCost(CableShovel, OutElectricity) + StripCost(Dragline, OutElectricity))
		
		DevelopmentArray(4) = StripCost(Loader, OutParts) + StripCost(Shovel, OutParts) + StripCost(CableShovel, OutParts) + StripCost(Dragline, OutParts) + StripCost(Scraper, OutParts) + StripCost(Truck, OutParts) + StripCost(Articulated, OutParts)
		DevelopmentArray(4) = DevelopmentArray(4) + ((EqCost(Waste, Percussion, OutParts) + EqCost(Waste, Rotary, OutParts)) * (ppt / wt))
		DevelopmentArray(4) = DevelopmentArray(4) + ((EqCost(Ore, Pump, OutParts) + EqCost(Ore, Dozer, OutParts) + EqCost(Ore, WaterTanker, OutParts)) * (ppt / wt))
		DevelopmentArray(4) = DevelopmentArray(4) + ((EqCost(Ore, PowderBuggy, OutParts) + EqCost(Ore, MainTruck, OutParts) + EqCost(Ore, Grader, OutParts) + EqCost(Ore, Light, OutParts) + EqCost(Ore, Pickup, OutParts)) * (ppt / wt))
		
		DevelopmentArray(5) = StripCost(Loader, OutLube) + StripCost(Shovel, OutLube) + StripCost(CableShovel, OutLube) + StripCost(Dragline, OutLube) + StripCost(Scraper, OutLube) + StripCost(Truck, OutLube) + StripCost(Articulated, OutLube)
		DevelopmentArray(5) = DevelopmentArray(5) + ((EqCost(Waste, Percussion, OutLube) + EqCost(Waste, Rotary, OutLube)) * (ppt / wt))
		DevelopmentArray(5) = DevelopmentArray(5) + ((EqCost(Ore, Pump, OutLube) + EqCost(Ore, Dozer, OutLube) + EqCost(Ore, WaterTanker, OutLube)) * (ppt / wt))
		DevelopmentArray(5) = DevelopmentArray(5) + ((EqCost(Ore, PowderBuggy, OutLube) + EqCost(Ore, MainTruck, OutLube) + EqCost(Ore, Grader, OutLube) + EqCost(Ore, Light, OutLube) + EqCost(Ore, Pickup, OutLube)) * (ppt / wt))
		
		DevelopmentArray(6) = (StripCost(Loader, OutTires) + StripCost(Scraper, OutTires) + StripCost(Truck, OutTires) + StripCost(Articulated, OutTires))
		DevelopmentArray(6) = DevelopmentArray(6) + ((EqCost(Ore, WaterTanker, OutTires) + EqCost(Ore, PowderBuggy, OutTires) + EqCost(Ore, Grader, OutTires) + EqCost(Ore, MainTruck, OutTires) + EqCost(Ore, Pickup, OutTires)) * (ppt / wt))
		
		DevelopmentArray(7) = (ppt * wpf * wpow) * TaxRate
		
		'UPGRADE_WARNING: Couldn't resolve default property of object f. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		f = CellValues(EquipmentOne, 30, WhichSegment).Value
		'UPGRADE_WARNING: Couldn't resolve default property of object g. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		g = CellValues(EquipmentOne, 31, WhichSegment).Value
		
		If wpsg <> 0 Then
			DvPdVl = (ppt * wpf) / (wpsg * 62.4)
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object f. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If f > 0 Then
			DvFtDr = (DvPdVl / (wlf / 100)) / ((3.141593 * (EqDefault(Percussion, f, HoleDiameter) / 12) ^ 2) / 4)
		Else
			If wlf > 0 Then
				DvFtDr = (DvPdVl / (wlf / 100)) / ((3.141593 * (EqDefault(Rotary, g, HoleDiameter) / 12) ^ 2) / 4)
			End If
		End If
		
		DvPrNm = DvFtDr / (wbh * 1.15)
		
		AddToIt = (Int((EqCost(Waste, Percussion, OutFeet) + EqCost(Waste, Rotary, OutFeet)) / (wbh * 1.15)) + 1)
		DvCpNm = ((DvPrNm / AddToIt) * 2) + DvPrNm
		
		DevelopmentArray(8) = DvCpNm * cap * TaxRate
		DevelopmentArray(9) = DvPrNm * prim * TaxRate
		DevelopmentArray(10) = (DvFtDr / btw) * (EqCost(Waste, Percussion, OutBit) + EqCost(Waste, Rotary, OutBit)) * TaxRate
		DevelopmentArray(11) = (DvFtDr / rdw) * (EqCost(Waste, Percussion, OutSteel) + EqCost(Waste, Rotary, OutSteel)) * TaxRate
		DevelopmentArray(12) = (DvFtDr * 2) * det * TaxRate
		
		For x = 1 To 12
			'UPGRADE_WARNING: Couldn't resolve default property of object x. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			DevelopmentSupplyTemp = DevelopmentSupplyTemp + DevelopmentArray(x)
		Next x
		
		DevelopmentArray(13) = DevelopmentSupplyTemp * (CellValues(Supply, 13, WhichSegment).Value / 100)
		
		CellValues(DevelopmentResult, 0, Int((CellValues(Production, 15, 0).Value) - 1)).Value = DevelopmentArray(13) + DevelopmentSupplyTemp
		
		'UPGRADE_WARNING: Couldn't resolve default property of object y. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If y = 15 Then
			wt = 0
			'UPGRADE_WARNING: Couldn't resolve default property of object y. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			y = 0
		End If
		
	End Sub
	Sub elcal(ByRef ElectricalRate As Object)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object ElectricalRate. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElectricalRate = CellValues(Supply, 1, WhichSegment).Value
		
	End Sub
	Sub excal(ByRef WastePowderCost As Decimal, ByRef PowderCost As Decimal, ByRef CapCost As Decimal, ByRef DetCordCost As Decimal)
		
		PowderCost = CellValues(Supply, 2, WhichSegment).Value
		WastePowderCost = CellValues(Supply, 14, WhichSegment).Value
		CapCost = CellValues(Supply, 3, WhichSegment).Value
		DetCordCost = CellValues(Supply, 5, WhichSegment).Value
		
	End Sub
	Sub prcal(ByRef prim As Decimal, ByRef btw As Decimal, ByRef rdw As Decimal)
		
		prim = CellValues(Supply, 4, WhichSegment).Value
		btw = CellValues(Supply, 9, WhichSegment).Value
		rdw = CellValues(Supply, 10, WhichSegment).Value
		
	End Sub
	Sub ptcal(ByRef ppt As Decimal)
		
		ppt = CellValues(Production, 4, WhichSegment).Value
		
	End Sub
	Sub rscal(ByRef rs As Decimal)
		
		rs = CellValues(Production, 0, WhichSegment).Value
		
	End Sub
	Sub sbcal(ByRef burden As Decimal, ByRef annualprod As Decimal)
		
		annualprod = (CellValues(Production, 5, WhichSegment).Value * CellValues(Production, 3, WhichSegment).Value)
		burden = (1 + (CellValues(Salary, 24, WhichSegment).Value / 100))
		
	End Sub
	Sub scalc()
        Dim tempsupply As Object = Nothing

        Dim x As Short
		Dim y As Short
		Dim ElectricalRate As Decimal
		Dim TaxRate As Decimal
		Dim obh As Decimal
		Dim wbh As Decimal
		Dim ot As Decimal
		Dim wt As Decimal
		Dim opf As Decimal
		Dim wpf As Decimal
		Dim prim As Decimal
		Dim btw As Decimal
		Dim rdw As Decimal
		Dim pow As Decimal
		Dim wpow As Decimal
		Dim cap As Decimal
		Dim det As Decimal
		Dim fuel As Decimal
		Dim bittemp As Decimal
		Dim cordtemp As Decimal
		
		On Error Resume Next
		
		Call txcal(TaxRate)
		Call bhcal(obh, wbh)
		Call otcal(ot, wt)
		Call pfcal(opf, wpf)
		Call prcal(prim, btw, rdw)
		Call excal(wpow, pow, cap, det)
		Call fucal(fuel)
		
		For x = 0 To 11
			For y = 0 To 6
				OutSupply(x, y) = 0
			Next y
		Next x
		
		For WhichSegment = 0 To MaxSegment
			Call txcal(TaxRate)
			Call bhcal(obh, wbh)
			Call otcal(ot, wt)
			Call pfcal(opf, wpf)
			Call prcal(prim, btw, rdw)
			Call excal(wpow, pow, cap, det)
			Call fucal(fuel)
			If opf + wpf > 0 Then
				bittemp = 0
				cordtemp = 0
				OutSupply(0, WhichSegment) = ((ot * opf * CellValues(Supply, 2, WhichSegment).Value) * TaxRate)
				OutSupply(1, WhichSegment) = (CellValues(Powder, 2, WhichSegment).Value * CellValues(Supply, 3, WhichSegment).Value) * TaxRate
				OutSupply(2, WhichSegment) = (CellValues(Powder, 4, WhichSegment).Value * CellValues(Supply, 4, WhichSegment).Value) * TaxRate
				OutSupply(3, WhichSegment) = ((CellValues(Powder, 25, WhichSegment).Value + CellValues(Powder, 18, WhichSegment).Value) * CellValues(Supply, 5, WhichSegment).Value) * 2 * TaxRate
				cordtemp = (CellValues(Powder, 25, WhichSegment).Value + CellValues(Powder, 18, WhichSegment).Value) * 2
				If CellValues(Supply, 9, WhichSegment).Value > 0 Then
					bittemp = (CellValues(Powder, 25, WhichSegment).Value / CellValues(Supply, 9, WhichSegment).Value) + (CellValues(Powder, 18, WhichSegment).Value / CellValues(Supply, 9, WhichSegment).Value)
					OutSupply(4, WhichSegment) = (((CellValues(Powder, 25, WhichSegment).Value / CellValues(Supply, 9, WhichSegment).Value) * CellValues(Powder, 13, WhichSegment).Value) + ((CellValues(Powder, 18, WhichSegment).Value / CellValues(Supply, 9, WhichSegment).Value) * CellValues(Powder, 19, WhichSegment).Value)) * TaxRate
				End If
				If CellValues(Supply, 10, WhichSegment).Value > 0 Then
					OutSupply(5, WhichSegment) = (((CellValues(Powder, 25, WhichSegment).Value / CellValues(Supply, 10, WhichSegment).Value) * CellValues(Powder, 14, WhichSegment).Value) + ((CellValues(Powder, 18, WhichSegment).Value / CellValues(Supply, 10, WhichSegment).Value) * CellValues(Powder, 20, WhichSegment).Value)) * TaxRate
				End If
				OutSupply(6, WhichSegment) = ((wt * wpf * CellValues(Supply, 14, WhichSegment).Value) * TaxRate)
				OutSupply(7, WhichSegment) = (CellValues(Powder, 8, WhichSegment).Value * CellValues(Supply, 15, WhichSegment).Value) * TaxRate
				OutSupply(8, WhichSegment) = (CellValues(Powder, 10, WhichSegment).Value * CellValues(Supply, 16, WhichSegment).Value) * TaxRate
				OutSupply(9, WhichSegment) = ((CellValues(Powder, 15, WhichSegment).Value + CellValues(Powder, 21, WhichSegment).Value) * CellValues(Supply, 17, WhichSegment).Value) * 2 * TaxRate
				cordtemp = cordtemp + (CellValues(Powder, 15, WhichSegment).Value + CellValues(Powder, 21, WhichSegment).Value) * 2
				cordtemp = cordtemp / 3.281
				If CellValues(Supply, 19, WhichSegment).Value > 0 Then
					bittemp = bittemp + (CellValues(Powder, 15, WhichSegment).Value / CellValues(Supply, 19, WhichSegment).Value) + (CellValues(Powder, 21, WhichSegment).Value / CellValues(Supply, 19, WhichSegment).Value)
					OutSupply(10, WhichSegment) = (((CellValues(Powder, 15, WhichSegment).Value / CellValues(Supply, 19, WhichSegment).Value) * CellValues(Powder, 16, WhichSegment).Value) + ((CellValues(Powder, 21, WhichSegment).Value / CellValues(Supply, 19, WhichSegment).Value) * CellValues(Powder, 22, WhichSegment).Value)) * TaxRate
				End If
				If CellValues(Supply, 20, WhichSegment).Value > 0 Then
					OutSupply(11, WhichSegment) = (((CellValues(Powder, 15, WhichSegment).Value / CellValues(Supply, 20, WhichSegment).Value) * CellValues(Powder, 17, WhichSegment).Value) + ((CellValues(Powder, 21, WhichSegment).Value / CellValues(Supply, 20, WhichSegment).Value) * CellValues(Powder, 23, WhichSegment).Value)) * TaxRate
				End If
			End If
			For x = 0 To 11
				'UPGRADE_WARNING: Couldn't resolve default property of object tempsupply. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				tempsupply = tempsupply + OutSupply(x, WhichSegment)
			Next x
			
			'UPGRADE_WARNING: Couldn't resolve default property of object tempsupply. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			OutSupply(12, WhichSegment) = tempsupply * (CellValues(Supply, 13, WhichSegment).Value / 100)
			
			'UPGRADE_WARNING: Couldn't resolve default property of object tempsupply. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			SupSum = (OutSupply(12, WhichSegment) + tempsupply) / ot
		Next WhichSegment
		WhichSegment = 0
		
	End Sub
	Public Sub hrlab()
		
		Dim x As Short
		
		Dim WageArray(12) As Decimal
		Dim hrsh As Decimal
		Dim ot As Decimal
		Dim wt As Decimal
		Dim BurdenRate As Decimal
		Dim mcwg As Decimal
		Dim laef As Decimal
		Dim opf As Decimal
		Dim wpf As Decimal
		Dim mechtime As Decimal
		
		Dim jump As Boolean
		
		On Error Resume Next
		
		Call getout(jump)
		If jump = True Then Exit Sub
		
		Call hrcal(hrsh)
		Call otcal(ot, wt)
		Call mccal(BurdenRate, mcwg)
		Call lbcal(laef)
		Call pfcal(opf, wpf)
		
		For x = 0 To 8
			WageArray(x) = CellValues(Wage, x, WhichSegment).Value
			LaborArray(x) = 0
		Next x
		
		WageArray(11) = CellValues(Wage, 11, WhichSegment).Value
		LaborArray(11) = 0
		
		If ((opf + wpf) > 0) And (laef > 0) And (hrsh > 0) Then
			Hour_Renamed(DrillMan) = (Int((CellValues(EquipmentHours, 4, WhichSegment).Value + CellValues(EquipmentHours, 18, WhichSegment).Value + CellValues(EquipmentHours, 5, WhichSegment).Value + CellValues(EquipmentHours, 19, WhichSegment).Value) / laef / hrsh) + 1) * hrsh
		Else
			Hour_Renamed(DrillMan) = 0
		End If
		Hour_Renamed(BlastMan) = blhr
		Hour_Renamed(LoadMan) = (Int((CellValues(EquipmentHours, 0, WhichSegment).Value + CellValues(EquipmentHours, 2, WhichSegment).Value) / laef / hrsh) + 1) * hrsh
		Hour_Renamed(DriverMan) = (Int((CellValues(EquipmentHours, 1, WhichSegment).Value + CellValues(EquipmentHours, 3, WhichSegment).Value) / laef / hrsh) + 1) * hrsh
		Hour_Renamed(HeavyMan) = (Int(CellValues(EquipmentHours, 6, WhichSegment).Value / laef / hrsh) + 1) * hrsh
		If WageArray(UtilMan) > 0 Then
			Hour_Renamed(UtilMan) = (Int((CellValues(EquipmentHours, 7, WhichSegment).Value + CellValues(EquipmentHours, 8, WhichSegment).Value) / laef / hrsh) + 1) * hrsh
		Else
			Hour_Renamed(UtilMan) = 0
		End If
		
		mechtime = (CellValues(EquipmentHours, 0, WhichSegment).Value * CellValues(RepairLabor, 0, WhichSegment).Value) + (CellValues(EquipmentHours, 1, WhichSegment).Value * CellValues(RepairLabor, 1, WhichSegment).Value)
		mechtime = mechtime + ((CellValues(EquipmentHours, 2, WhichSegment).Value * CellValues(RepairLabor, 2, WhichSegment).Value) + (CellValues(EquipmentHours, 3, WhichSegment).Value * CellValues(RepairLabor, 3, WhichSegment).Value))
		mechtime = mechtime + ((CellValues(EquipmentHours, 4, WhichSegment).Value * CellValues(RepairLabor, 4, WhichSegment).Value) + (CellValues(EquipmentHours, 5, WhichSegment).Value * CellValues(RepairLabor, 5, WhichSegment).Value))
		mechtime = mechtime + ((CellValues(EquipmentHours, 6, WhichSegment).Value * CellValues(RepairLabor, 6, WhichSegment).Value) + (CellValues(EquipmentHours, 7, WhichSegment).Value * CellValues(RepairLabor, 7, WhichSegment).Value))
		mechtime = mechtime + ((CellValues(EquipmentHours, 8, WhichSegment).Value * CellValues(RepairLabor, 8, WhichSegment).Value) + (CellValues(EquipmentHours, 9, WhichSegment).Value * CellValues(RepairLabor, 9, WhichSegment).Value))
		mechtime = mechtime + ((CellValues(EquipmentHours, 10, WhichSegment).Value * CellValues(RepairLabor, 10, WhichSegment).Value) + (CellValues(EquipmentHours, 11, WhichSegment).Value * CellValues(RepairLabor, 11, WhichSegment).Value))
		mechtime = mechtime + ((CellValues(EquipmentHours, 12, WhichSegment).Value * CellValues(RepairLabor, 12, WhichSegment).Value) + (CellValues(EquipmentHours, 13, WhichSegment).Value * CellValues(RepairLabor, 13, WhichSegment).Value))
		mechtime = mechtime + ((CellValues(EquipmentHours, 14, WhichSegment).Value * CellValues(RepairLabor, 14, WhichSegment).Value) + (CellValues(EquipmentHours, 15, WhichSegment).Value * CellValues(RepairLabor, 15, WhichSegment).Value))
		mechtime = mechtime + ((CellValues(EquipmentHours, 16, WhichSegment).Value * CellValues(RepairLabor, 16, WhichSegment).Value) + (CellValues(EquipmentHours, 17, WhichSegment).Value * CellValues(RepairLabor, 17, WhichSegment).Value))
		mechtime = mechtime + ((CellValues(EquipmentHours, 18, WhichSegment).Value * CellValues(RepairLabor, 18, WhichSegment).Value) + (CellValues(EquipmentHours, 19, WhichSegment).Value * CellValues(RepairLabor, 19, WhichSegment).Value))
		If hrsh <> 0 Then
			mechtime = ((Int(mechtime / hrsh)) + 1) * hrsh
		End If
		
		Hour_Renamed(MechMan) = mechtime
		
		If WageArray(LaborMan) > 0 And hrsh > 0 Then
			Hour_Renamed(LaborMan) = (Int((mechtime * (1.07187 * ((ot + wt) ^ -0.057219))) / hrsh) + 1) * hrsh
		Else
			Hour_Renamed(LaborMan) = 0
		End If
		
		If WageArray(MaintMan) > 0 And hrsh > 0 Then
			Hour_Renamed(MaintMan) = (Int((mechtime * (1.72809 * ((ot + wt) ^ -0.03769))) / hrsh) + 1) * hrsh
		Else
			Hour_Renamed(MaintMan) = 0
		End If
		
		If WageArray(ElectricMan) > 0 And hrsh > 0 Then
			Hour_Renamed(ElectricMan) = (Int(CellValues(Development, 8, WhichSegment).Value / 950) + 1) * hrsh
		Else
			Hour_Renamed(ElectricMan) = 0
		End If
		
		For x = 0 To 8
			If CellValues(WorkForce, x, WhichSegment).Changed = True Then
				Hour_Renamed(x) = CellValues(WorkForce, x, WhichSegment).Value * hrsh
			End If
		Next x
		
		If CellValues(WorkForce, 11, WhichSegment).Changed = True Then
			Hour_Renamed(11) = CellValues(WorkForce, 11, WhichSegment).Value * hrsh
		End If
		
		If laef > 0 And hrsh > 0 Then
			If opf + wpf > 0 Then
				LaborArray(DrillMan) = (Int((CellValues(EquipmentHours, 4, WhichSegment).Value + CellValues(EquipmentHours, 5, WhichSegment).Value + CellValues(EquipmentHours, 18, WhichSegment).Value + CellValues(EquipmentHours, 19, WhichSegment).Value) / laef / hrsh) + 1) * hrsh * WageArray(DrillMan) * BurdenRate
				LaborArray(BlastMan) = blnm * hrsh * WageArray(BlastMan) * BurdenRate
			Else
				LaborArray(DrillMan) = 0
				LaborArray(BlastMan) = 0
			End If
		End If
		
		LaborArray(LoadMan) = CellValues(EquipmentHours, 0, WhichSegment).Value + CellValues(EquipmentHours, 2, WhichSegment).Value
		LaborArray(DriverMan) = CellValues(EquipmentHours, 1, WhichSegment).Value + CellValues(EquipmentHours, 3, WhichSegment).Value
		
		If CellValues(WorkForce, HeavyMan, WhichSegment).Changed = True Then
			LaborArray(HeavyMan) = Hour_Renamed(HeavyMan) * WageArray(HeavyMan) * BurdenRate
		Else
			LaborArray(HeavyMan) = CellValues(EquipmentHours, 6, WhichSegment).Value
		End If
		
		LaborArray(UtilMan) = CellValues(EquipmentHours, 7, WhichSegment).Value + CellValues(EquipmentHours, 8, WhichSegment).Value
		LaborArray(MechMan) = mechtime * WageArray(MechMan) * BurdenRate
		If (ot + wt) > 0 And hrsh > 0 Then
			LaborArray(LaborMan) = (Int((mechtime * (1.07187 * ((ot + wt) ^ -0.057219))) / hrsh) + 1) * hrsh * WageArray(LaborMan) * BurdenRate
			LaborArray(MaintMan) = (Int((mechtime * (1.72809 * ((ot + wt) ^ -0.03769))) / hrsh) + 1) * hrsh * WageArray(MaintMan) * BurdenRate
		End If
		
		LaborArray(ElectricMan) = (Int(CellValues(Development, 8, WhichSegment).Value / 950) + 1) * hrsh * WageArray(ElectricMan) * BurdenRate
		
		LaborArray(TotalMan) = 0
		
		For x = 0 To 8
			If CellValues(WorkForce, x, WhichSegment).Changed = True Then
				LaborArray(x) = Hour_Renamed(x) * WageArray(x) * BurdenRate
			End If
			LaborArray(TotalMan) = LaborArray(TotalMan) + LaborArray(x)
		Next x
		
		LaborArray(ElectricMan) = Hour_Renamed(ElectricMan) * WageArray(ElectricMan) * BurdenRate
		
		LaborArray(TotalMan) = LaborArray(TotalMan) + LaborArray(ElectricMan)
		
		CellValues(WorkForce, 9, WhichSegment).Value = 0
		
		For x = 0 To 8
			If CellValues(WorkForce, x, WhichSegment).Changed <> True Then
				If hrsh > 0 Then CellValues(WorkForce, x, WhichSegment).Value = Hour_Renamed(x) / hrsh
				CellValues(WorkForce, 9, WhichSegment).Value = CellValues(WorkForce, 9, WhichSegment).Value + CellValues(WorkForce, x, WhichSegment).Value
			End If
		Next x
		
		If hrsh > 0 Then CellValues(WorkForce, 11, WhichSegment).Value = Hour_Renamed(11) / hrsh
		CellValues(WorkForce, 9, WhichSegment).Value = CellValues(WorkForce, 9, WhichSegment).Value + CellValues(WorkForce, 11, WhichSegment).Value
		
	End Sub
	Public Sub hrcal(ByRef hrsh As Decimal)
		
		hrsh = CellValues(Production, 1, WhichSegment).Value
		
	End Sub
	Public Sub sgcal(ByRef opsg As Decimal, ByRef wpsg As Decimal)
		
		opsg = CellValues(Supply, 8, WhichSegment).Value
		wpsg = CellValues(Supply, 18, WhichSegment).Value
		
	End Sub
	Public Sub shcal(ByRef shdy As Decimal)
		
		shdy = CellValues(Production, 2, WhichSegment).Value
		
	End Sub
	Public Sub txcal(ByRef TaxRate As Decimal)
		
		TaxRate = 1 + (CellValues(Supply, 6, WhichSegment).Value / 100)
		
	End Sub
	Public Sub drcal(ByRef odr As Decimal, ByRef wdr As Decimal)
		
		odr = CellValues(Deposit, 4, WhichSegment).Value
		wdr = CellValues(Deposit, 11, WhichSegment).Value
		
	End Sub
	Public Sub dycal(ByRef dyyr As Decimal)
		
		dyyr = CellValues(Production, 3, WhichSegment).Value
		
	End Sub
	Public Sub fucal(ByRef fuel As Decimal)
		
		fuel = CellValues(Supply, 0, WhichSegment).Value
		
	End Sub
	Sub lbcal(ByRef laef As Decimal)
		
		laef = CellValues(Wage, 10, WhichSegment).Value / 100
		
	End Sub
	Sub lfcal(ByRef olf As Decimal, ByRef wlf As Decimal)
		
		olf = CellValues(Deposit, 3, WhichSegment).Value
		wlf = CellValues(Deposit, 10, WhichSegment).Value
		
	End Sub
	Sub mccal(ByRef BurdenRate As Decimal, ByRef mcwg As Decimal)
		
		BurdenRate = 1 + (CellValues(Wage, 9, WhichSegment).Value / 100)
		mcwg = CellValues(Wage, 6, WhichSegment).Value
		
	End Sub
	Sub otcal(ByRef ot As Decimal, ByRef wt As Decimal)
		
		ot = CellValues(Production, 5, WhichSegment).Value
		wt = CellValues(Production, 10, WhichSegment).Value
		
	End Sub
	Sub bcalc()
		Dim NewSegment As Object
		
		Dim SepticPrice As Decimal
		Dim ParkingPrice As Decimal
		Dim FencePrice As Decimal
		Dim GatePrice As Decimal
		Dim TanCo As Decimal
		Dim ShopSize As Decimal
		Dim DrySize As Decimal
		Dim OfficeSize As Decimal
		Dim WarehouseSize As Decimal
		Dim PowderMagazineSize As Decimal
		Dim unspco As Decimal
		Dim undrco As Decimal
		Dim unofco As Decimal
		Dim unwhco As Decimal
		Dim pmgco As Decimal
		Dim adnm As Decimal
		Dim PmNumber As Decimal
		Dim KvaStuff As Decimal
		Dim OreRoadLength As Decimal
		Dim WasteRoadLength As Decimal
		Dim kva As Decimal
		Dim SubStation As Decimal
		Dim SingleSwitch As Decimal
		Dim DoubleSwitch As Decimal
		Dim WorkingBase As Decimal
		Dim BaseHead As Decimal
		Dim OreMagazineSize As Decimal
		Dim WasteMagazineSize As Decimal
		Dim OreBinSize As Decimal
		Dim WasteBinSize As Decimal
		Dim HeightAdjust As Decimal
		Dim MaterialAdjust As Decimal
		Dim SegYear As Decimal
		
		Dim hrsh As Decimal
		Dim shdy As Decimal
		Dim dyyr As Decimal
		Dim opsg As Decimal
		Dim wpsg As Decimal
		Dim ot As Decimal
		Dim wt As Decimal
		Dim opf As Decimal
		Dim wpf As Decimal
		Dim burden As Decimal
		Dim annualprod As Decimal
		Dim TempResult As Decimal
		
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
		Dim MagazineNumber As Short
		Dim BinNumber As Short
		
		Dim jump As Boolean
		
		On Error Resume Next
		
		Call getout(jump)
		If jump = True Then Exit Sub
		
		ao = Int(CellValues(EquipmentOne, 0, WhichSegment).Value)
		aw = Int(CellValues(EquipmentOne, 5, WhichSegment).Value)
		bo = Int(CellValues(EquipmentOne, 1, WhichSegment).Value)
		bw = Int(CellValues(EquipmentOne, 6, WhichSegment).Value)
		dore = Int(CellValues(EquipmentOne, 2, WhichSegment).Value)
		dw = Int(CellValues(EquipmentOne, 7, WhichSegment).Value)
		eo = Int(CellValues(EquipmentOne, 3, WhichSegment).Value)
		ew = Int(CellValues(EquipmentOne, 8, WhichSegment).Value)
		co = Int(CellValues(EquipmentOne, 4, WhichSegment).Value) + Int(CellValues(EquipmentOne, 20, WhichSegment).Value) + Int(CellValues(EquipmentOne, 21, WhichSegment).Value)
		cw = Int(CellValues(EquipmentOne, 9, WhichSegment).Value) + Int(CellValues(EquipmentOne, 25, WhichSegment).Value) + Int(CellValues(EquipmentOne, 26, WhichSegment).Value)
		
		Call hrcal(hrsh)
		Call shcal(shdy)
		Call dycal(dyyr)
		Call sgcal(opsg, wpsg)
		Call otcal(ot, wt)
		Call pfcal(opf, wpf)
		Call sbcal(burden, annualprod)
		
		'=============================== Road Cost =================================='
		
		If WhichSegment = 0 Then
			
			OreRoadLength = 0
			WasteRoadLength = 0
			
			For x = 0 To 2 Step 2
				OreRoadLength = OreRoadLength + CellValues(Haul, x, WhichSegment).Value
				WasteRoadLength = WasteRoadLength + CellValues(Haul, x + 4, WhichSegment).Value
			Next x
			For x = 8 To 14 Step 2
				OreRoadLength = OreRoadLength + CellValues(Haul, x, WhichSegment).Value
				WasteRoadLength = WasteRoadLength + CellValues(Haul, x + 12, WhichSegment).Value
			Next x
			
			EngBase(1) = (OreRoadLength * CellValues(Development, 11, WhichSegment).Value)
			EngBase(1) = EngBase(1) + (WasteRoadLength * CellValues(Development, 12, WhichSegment).Value)
			
			CellValues(DevelopmentResult, 1, 1).Value = EngBase(1)
			
		End If
		'============================= Building Costs ================================='
		
		For x = 0 To 3
			OutBuilding(x) = 0
		Next x
		
		MaterialAdjust = 1
		HeightAdjust = 1
		
		'================================= Shop ===================================='
		
		If co + cw > 0 Then
			
			If CellValues(Development, 2, WhichSegment).Changed = False Then
				CellValues(Development, 2, WhichSegment).Value = CellValues(Building, 0, WhichSegment).Value * CellValues(Building, 1, WhichSegment).Value
			End If
			
			ShopSize = CellValues(Development, 2, WhichSegment).Value
			
			Select Case LTrim(RTrim(LCase(CellValues(Building, 3, WhichSegment).Word)))
				Case "permanent wood"
					MaterialAdjust = 1.046747
				Case "permanent steel"
					MaterialAdjust = 1.042956
				Case "permanent concrete"
					MaterialAdjust = 1
				Case "wood frame/steel siding"
					MaterialAdjust = 0.966519
				Case "modular"
					MaterialAdjust = 0.420359
				Case "mobile"
					MaterialAdjust = 0.624331
				Case Else
					MaterialAdjust = 1
			End Select
			
			If CellValues(Development, 13, WhichSegment).Changed = False Then
				If CellValues(Building, 0, WhichSegment).Value < 11 Then
					Select Case ShopSize
						Case 0 To 1999
							unspco = EqDefault(Shop, 1, Price)
						Case 2000 To 7999
							unspco = EqDefault(Shop, 2, Price)
						Case 8000 To 13799
							unspco = EqDefault(Shop, 3, Price)
						Case Else
							unspco = EqDefault(Shop, 4, Price)
					End Select
				End If
				If CellValues(Building, 0, WhichSegment).Value >= 11 And CellValues(Building, 0, WhichSegment).Value < 18 Then
					Select Case ShopSize
						Case 0 To 1999
							unspco = EqDefault(Shop, 5, Price)
						Case 2000 To 7999
							unspco = EqDefault(Shop, 6, Price)
						Case 8000 To 13799
							unspco = EqDefault(Shop, 7, Price)
						Case Else
							unspco = EqDefault(Shop, 8, Price)
					End Select
				End If
				If CellValues(Building, 0, WhichSegment).Value >= 18 Then
					Select Case ShopSize
						Case 0 To 1999
							unspco = EqDefault(Shop, 9, Price)
						Case 2000 To 7999
							unspco = EqDefault(Shop, 10, Price)
						Case 8000 To 13799
							unspco = EqDefault(Shop, 11, Price)
						Case Else
							unspco = EqDefault(Shop, 12, Price)
					End Select
				End If
				CellValues(Development, 13, WhichSegment).Value = (unspco * MaterialAdjust)
			End If
		Else
			
		End If
		
		For x = 1 To (Int(CellValues(Production, 15, 0).Value) - 1)
			CellValues(Building, 20, x).Value = (CellValues(Development, 2, WhichSegment).Value * CellValues(Development, 13, WhichSegment).Value) / (Int(CellValues(Production, 15, 0).Value) - 1)
		Next x
		
		EngBase(2) = CellValues(Development, 2, WhichSegment).Value * CellValues(Development, 13, WhichSegment).Value
		
		'==================================== Dry ==================================='
		
		If CellValues(Development, 3, WhichSegment).Changed = False Then
			CellValues(Development, 3, WhichSegment).Value = CellValues(Building, 4, WhichSegment).Value * CellValues(Building, 5, WhichSegment).Value
		End If
		
		If CellValues(Building, 6, WhichSegment).Value > 0 Then
			HeightAdjust = (0.185619 * (CellValues(Building, 6, WhichSegment).Value ^ 0.754487))
		Else
			HeightAdjust = 1
		End If
		
		DrySize = CellValues(Development, 3, WhichSegment).Value
		
		Select Case LTrim(RTrim(LCase(CellValues(Building, 7, WhichSegment).Word)))
			Case "permanent wood"
				MaterialAdjust = 1.046747
			Case "permanent steel"
				MaterialAdjust = 1.042956
			Case "permanent concrete"
				MaterialAdjust = 1
			Case "wood frame/steel siding"
				MaterialAdjust = 0.966519
			Case "modular"
				MaterialAdjust = 0.420359
			Case "mobile"
				MaterialAdjust = 0.624331
			Case Else
				MaterialAdjust = 1
		End Select
		
		If CellValues(Development, 14, WhichSegment).Changed = False Then
			Select Case DrySize
				Case 0 To 1999
					undrco = EqDefault(Dry, 1, Price)
				Case 2000 To 5999
					undrco = EqDefault(Dry, 2, Price)
				Case 6000 To 11999
					undrco = EqDefault(Dry, 3, Price)
				Case Else
					undrco = EqDefault(Dry, 4, Price)
			End Select
			CellValues(Development, 14, WhichSegment).Value = (undrco * HeightAdjust * MaterialAdjust)
		End If
		
		For x = 1 To (Int(CellValues(Production, 15, 0).Value) - 1)
			CellValues(Building, 21, x).Value = (CellValues(Development, 3, WhichSegment).Value * CellValues(Development, 14, WhichSegment).Value) / (Int(CellValues(Production, 15, 0).Value) - 1)
		Next x
		
		EngBase(2) = EngBase(2) + (CellValues(Development, 3, WhichSegment).Value * CellValues(Development, 14, WhichSegment).Value)
		
		'================================== Office ==================================='
		
		For x = 0 To 11
			adnm = adnm + CellValues(Staff, x, WhichSegment).Value
		Next x
		
		If CellValues(Development, 4, WhichSegment).Changed = False Then
			CellValues(Development, 4, WhichSegment).Value = CellValues(Building, 8, WhichSegment).Value * CellValues(Building, 9, WhichSegment).Value
		End If
		
		If CellValues(Building, 10, WhichSegment).Value > 0 Then
			HeightAdjust = (0.185619 * (CellValues(Building, 10, WhichSegment).Value ^ 0.754487))
		Else
			HeightAdjust = 1
		End If
		
		OfficeSize = CellValues(Development, 4, WhichSegment).Value
		
		Select Case LTrim(RTrim(LCase(CellValues(Building, 11, WhichSegment).Word)))
			Case "permanent wood"
				MaterialAdjust = 1.046747
			Case "permanent steel"
				MaterialAdjust = 1.042956
			Case "permanent concrete"
				MaterialAdjust = 1
			Case "wood frame/steel siding"
				MaterialAdjust = 0.966519
			Case "modular"
				MaterialAdjust = 0.420359
			Case "mobile"
				MaterialAdjust = 0.624331
			Case Else
				MaterialAdjust = 1
		End Select
		
		If CellValues(Development, 15, WhichSegment).Changed = False Then
			Select Case OfficeSize
				Case 0 To 3999
					unofco = EqDefault(Office, 1, Price)
				Case 4000 To 5499
					unofco = EqDefault(Office, 2, Price)
				Case 5500 To 6999
					unofco = EqDefault(Office, 3, Price)
				Case 7000 To 8499
					unofco = EqDefault(Office, 4, Price)
				Case 8500 To 9999
					unofco = EqDefault(Office, 5, Price)
				Case 10000 To 21999
					unofco = EqDefault(Office, 6, Price)
				Case 22000 To 33999
					unofco = EqDefault(Office, 7, Price)
				Case Else
					unofco = EqDefault(Office, 8, Price)
			End Select
			CellValues(Development, 15, WhichSegment).Value = (unofco * HeightAdjust * MaterialAdjust)
		End If
		
		For x = 1 To (Int(CellValues(Production, 15, 0).Value) - 1)
			CellValues(Building, 22, x).Value = (CellValues(Development, 4, WhichSegment).Value * CellValues(Development, 15, WhichSegment).Value) / (Int(CellValues(Production, 15, 0).Value) - 1)
		Next x
		
		EngBase(2) = EngBase(2) + (CellValues(Development, 4, WhichSegment).Value * CellValues(Development, 15, WhichSegment).Value)
		
		'================================== Warehouse ==============================='
		
		WarehouseSize = 0
		
		If CellValues(Development, 5, WhichSegment).Changed = False Then
			CellValues(Development, 5, WhichSegment).Value = CellValues(Building, 12, WhichSegment).Value * CellValues(Building, 13, WhichSegment).Value
		End If
		
		If CellValues(Building, 14, WhichSegment).Value > 0 Then
			HeightAdjust = (0.185619 * (CellValues(Building, 14, WhichSegment).Value ^ 0.754487))
		Else
			HeightAdjust = 1
		End If
		
		WarehouseSize = CellValues(Development, 5, WhichSegment).Value
		
		Select Case LTrim(RTrim(LCase(CellValues(Building, 15, WhichSegment).Word)))
			Case "permanent wood"
				MaterialAdjust = 1.046747
			Case "permanent steel"
				MaterialAdjust = 1.042956
			Case "permanent concrete"
				MaterialAdjust = 1
			Case "wood frame/steel siding"
				MaterialAdjust = 0.966519
			Case "modular"
				MaterialAdjust = 0.420359
			Case "mobile"
				MaterialAdjust = 0.624331
			Case Else
				MaterialAdjust = 1
		End Select
		
		If CellValues(Development, 16, WhichSegment).Changed = False Then
			Select Case WarehouseSize
				Case 0 To 9999
					unwhco = EqDefault(Warehouse, 1, Price)
				Case 10000 To 24999
					unwhco = EqDefault(Warehouse, 2, Price)
				Case 25000 To 34499
					unwhco = EqDefault(Warehouse, 3, Price)
				Case Else
					unwhco = EqDefault(Warehouse, 4, Price)
			End Select
			CellValues(Development, 16, WhichSegment).Value = (unwhco * HeightAdjust * MaterialAdjust)
		End If
		
		For x = 0 To (Int(CellValues(Production, 15, 0).Value) - 1)
			If Int(CellValues(Production, 15, 0).Value) <> 1 Then
				CellValues(Building, 23, x).Value = (CellValues(Development, 5, WhichSegment).Value * CellValues(Development, 16, WhichSegment).Value) / (Int(CellValues(Production, 15, 0).Value) - 1)
			End If
		Next x
		
		EngBase(2) = EngBase(2) + (CellValues(Development, 5, WhichSegment).Value * CellValues(Development, 16, WhichSegment).Value)
		
		For x = 1 To (Int(CellValues(Production, 15, 0).Value) - 1)
			CellValues(DevelopmentResult, 2, x).Value = (EngBase(2) / (Int(CellValues(Production, 15, 0).Value) - 1))
		Next x
		
		'=================== Powder Magazine/ANFO Storage Bin Cost ======================'
		
		Call TimeLineCalc()
		For NewSegment = 0 To MaxSegment
			Select Case LTrim(RTrim(LCase(CellValues(Production, 7, 0).Word)))
				Case "anfo"
					'UPGRADE_WARNING: Couldn't resolve default property of object NewSegment. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If OreBinSize < (CellValues(Powder, 12, NewSegment).Value * CellValues(Powder, 1, NewSegment).Value) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object NewSegment. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						OreBinSize = (CellValues(Powder, 12, NewSegment).Value * CellValues(Powder, 1, NewSegment).Value)
					End If
				Case Else
					'UPGRADE_WARNING: Couldn't resolve default property of object NewSegment. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If OreMagazineSize < (CellValues(Powder, 12, NewSegment).Value * CellValues(Powder, 1, NewSegment).Value) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object NewSegment. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						OreMagazineSize = (CellValues(Powder, 12, NewSegment).Value * CellValues(Powder, 1, NewSegment).Value)
					End If
			End Select
			
			Select Case LTrim(RTrim(LCase(CellValues(Production, 12, 0).Word)))
				Case "anfo"
					'UPGRADE_WARNING: Couldn't resolve default property of object NewSegment. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If WasteBinSize < (CellValues(Powder, 13, NewSegment).Value * CellValues(Powder, 7, NewSegment).Value) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object NewSegment. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						WasteBinSize = (CellValues(Powder, 13, NewSegment).Value * CellValues(Powder, 7, NewSegment).Value)
					End If
				Case Else
					'UPGRADE_WARNING: Couldn't resolve default property of object NewSegment. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If WasteMagazineSize < (CellValues(Powder, 13, NewSegment).Value * CellValues(Powder, 7, NewSegment).Value) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object NewSegment. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						WasteMagazineSize = (CellValues(Powder, 13, NewSegment).Value * CellValues(Powder, 7, NewSegment).Value)
					End If
			End Select
		Next NewSegment
		
		If CellValues(Development, 7, WhichSegment).Changed = False Then
			CellValues(Development, 7, WhichSegment).Value = (OreMagazineSize + WasteMagazineSize)
		End If
		
		If CellValues(Development, 6, WhichSegment).Changed = False Then
			CellValues(Development, 6, WhichSegment).Value = (OreBinSize + WasteBinSize)
		End If
		
		'Update These Costs 2014
		If CellValues(Development, 18, WhichSegment).Changed = False And (OreMagazineSize + WasteMagazineSize) <> 0 Then
			pmgco = 0
			Select Case (OreMagazineSize + WasteMagazineSize)
				Case 1 To 16
					pmgco = 1110
				Case 17 To 34
					pmgco = 2280
				Case 35 To 43
					pmgco = 3010
				Case 44 To 96
					pmgco = 4410
				Case 97 To 138
					pmgco = 6200
				Case 139 To 212
					pmgco = 8060
				Case 213 To 270
					pmgco = 9930
				Case 271 To 392
					pmgco = 12230
				Case 393 To 539
					pmgco = 15280
				Case 540 To 1078
					pmgco = 21940
				Case 1079 To 1920
					pmgco = 32870
				Case 1921 To 2760
					pmgco = 44010
				Case 2761 To 3600
					pmgco = 53800
				Case 3601 To 4400
					pmgco = 62530
				Case 4401 To 5280
					pmgco = 70850
				Case 5281 To 6120
					pmgco = 78870
				Case 6121 To 6960
					pmgco = 86300
				Case 6961 To 7800
					pmgco = 93410
				Case 7801 To 8640
					pmgco = 100300
				Case 8641 To 9280
					pmgco = 106100
				Case 9281 To 10400
					pmgco = 112800
				Case Is > 10401
					pmgco = 122200
				Case Else
					pmgco = 0
			End Select
			CellValues(Development, 18, WhichSegment).Value = pmgco
		End If
		
		'Update These Costs 2014
		If CellValues(Development, 17, WhichSegment).Changed = False And (OreBinSize + WasteBinSize) <> 0 Then
			pmgco = 0
			Select Case (OreBinSize + WasteBinSize)
				Case 1 To 199
					pmgco = 20780
				Case 200 To 549
					pmgco = 30770
				Case 550 To 869
					pmgco = 37220
				Case 870 To 1279
					pmgco = 42110
				Case 1280 To 2349
					pmgco = 49220
				Case 2350 To 3149
					pmgco = 55700
				Case 3150 To 6134
					pmgco = 65100
				Case Is > 6135
					pmgco = 75600
				Case Else
					pmgco = 0
			End Select
			CellValues(Development, 17, WhichSegment).Value = pmgco
		End If
		
		MagazineNumber = 0
		BinNumber = 0
		
		If CellValues(Development, 18, WhichSegment).Value <> 0 Then
			If CellValues(Development, 18, WhichSegment).Value > 10400 Then
				MagazineNumber = Int(CellValues(Development, 18, WhichSegment).Value / 10400) + 1
			Else
				MagazineNumber = 1
			End If
		ElseIf CellValues(Development, 17, WhichSegment).Value <> 0 Then 
			If CellValues(Development, 17, WhichSegment).Value > 6134 Then
				BinNumber = Int(CellValues(Development, 17, WhichSegment).Value / 6134) + 1
			Else
				BinNumber = 1
			End If
		End If
		
		EngBase(3) = (CellValues(Development, 17, WhichSegment).Value * BinNumber) + (CellValues(Development, 18, WhichSegment).Value * MagazineNumber)
		
		If WhichSegment = 0 Then
			CellValues(DevelopmentResult, 3, Int(CellValues(Production, 15, WhichSegment).Value) - 1).Value = EngBase(3)
		End If
		
		'=========================== Electrical System ================================='
		
		Call ElectEngr()
		
		'============================== Clearing ====================================='
		
		Call ClearEngr()
		
		'====================== Employee Parking/Mine Yard ============================='
		
		'Update These Costs (dollars per square foot) 2014 (spreadsheet in Update 2014 folder)
		
		Select Case LTrim(RTrim(LCase(CellValues(Site, 4, WhichSegment).Word)))
			Case "crushed rock"
				ParkingPrice = 0.287
			Case "bituminous"
				ParkingPrice = 0.481
				'includes the $0.287 crsuhed rock price - 5% bitumen
			Case "gravel"
				ParkingPrice = 0.126
		End Select
		
		ParkingPrice = CellValues(Site, 2, WhichSegment).Value * ParkingPrice
		
		EngBase(6) = ParkingPrice
		
		If WhichSegment = 0 And CellValues(Development, 10, WhichSegment).Value <> 0 Then
			If CellValues(Development, 21, WhichSegment).Changed = False Then
				CellValues(Development, 21, WhichSegment).Value = ParkingPrice / CellValues(Development, 10, WhichSegment).Value
			End If
		End If
		
		If WhichSegment = 0 Then
			CellValues(DevelopmentResult, 6, 1).Value = CellValues(Development, 21, WhichSegment).Value * CellValues(Development, 10, WhichSegment).Value
		End If
		
		'======================= Sewage Treatment and Disposal ========================='
		
		'Update These Costs (dollars per daily capacity (gallons)) 2014
		
		SepticPrice = (CellValues(Site, 0, WhichSegment).Value * CellValues(Site, 5, WhichSegment).Value)
		
		If CellValues(Development, 22, WhichSegment).Changed = False Then
			CellValues(Development, 22, WhichSegment).Value = SepticPrice
		End If
		
		Select Case LTrim(RTrim(LCase(CellValues(Site, 6, WhichSegment).Word)))
			Case "sewage treatment plant"
				SepticPrice = (397.6147923 * (SepticPrice ^ 0.686141))
			Case "septic system"
				SepticPrice = (97.010388 * (SepticPrice ^ 0.570384))
			Case "portable self-contained"
				SepticPrice = 740 * ((Int(CellValues(Site, 0, WhichSegment).Value / 6)) + 1)
		End Select
		
		EngBase(7) = SepticPrice
		
		If WhichSegment = 0 Then
			If CellValues(Development, 23, WhichSegment).Changed = False Then
				CellValues(Development, 23, WhichSegment).Value = SepticPrice
			End If
		End If
		
		If WhichSegment = 0 Then
			CellValues(DevelopmentResult, 7, Int(CellValues(Production, 15, 0).Value) - 1).Value = CellValues(Development, 23, WhichSegment).Value
		End If
		
		'=============================== Fence Costs ================================='
		
		'Update These Costs 2014
		
		Select Case LTrim(RTrim(LCase(CellValues(Site, 9, WhichSegment).Word)))
			Case "chain link"
				FencePrice = 19.25
				GatePrice = 335
			Case "chain link/barbed wire"
				FencePrice = 43.5
				GatePrice = 2775
			Case "barbed wire"
				FencePrice = 8.041
				GatePrice = 273
			Case "straight wire"
				FencePrice = 7.73
				GatePrice = 245
			Case "wood rail"
				FencePrice = 18.85
				GatePrice = 279
		End Select
		
		GatePrice = GatePrice * CellValues(Site, 8, WhichSegment).Value
		
		If CellValues(Site, 7, WhichSegment).Value <> 0 Then
			GatePrice = GatePrice / CellValues(Site, 7, WhichSegment).Value
		Else
			GatePrice = 0
		End If
		
		If WhichSegment = 0 Then
			If CellValues(Development, 25, WhichSegment).Changed = False Then
				CellValues(Development, 25, WhichSegment).Value = FencePrice + GatePrice
			End If
		End If
		
		FencePrice = (CellValues(Development, 25, WhichSegment).Value * CellValues(Development, 24, WhichSegment).Value)
		
		EngBase(7) = EngBase(7) + FencePrice
		
		If WhichSegment = 0 Then
			CellValues(DevelopmentResult, 8, Int(CellValues(Production, 15, 0).Value) - 1).Value = FencePrice
		End If
		
		'=============================== Fuel Tanks =================================='
		
		CellValues(Development, 26, WhichSegment).Value = CellValues(FuelStorage, 5, WhichSegment).Value
		
		Select Case CellValues(FuelStorage, 6, WhichSegment).Value
			Case Is <= 1000
				TanCo = 2550
			Case Is <= 2000
				TanCo = 3860
			Case Is <= 5000
				TanCo = 8820
			Case Is <= 10000
				TanCo = 12460
			Case Is <= 12000
				TanCo = 13750
			Case Is <= 15000
				TanCo = 15720
		End Select
		
		CellValues(Development, 27, WhichSegment).Value = TanCo
		
		EngBase(7) = EngBase(7) + (CellValues(FuelStorage, 7, WhichSegment).Value * CellValues(Development, 27, WhichSegment).Value)
		
		If WhichSegment = 0 Then
			CellValues(DevelopmentResult, 9, Int(CellValues(Production, 15, 0).Value) - 1).Value = EngBase(7)
		End If
		
		For x = 1 To (Int(CellValues(Production, 15, 0).Value) - 1)
			CellValues(Summary, 9, x).Value = CellValues(DevelopmentResult, 0, x).Value + CellValues(DevelopmentResult, 1, x).Value + CellValues(DevelopmentResult, 5, x).Value + CellValues(DevelopmentResult, 6, x).Value
			'CellValues(Summary, 9, x).Value = (EngBase(1) + EngBase(5) + EngBase(6)) / ((Int(CellValues(Production, 15, 0).Value) - 1))
			'If x = (Int(CellValues(Production, 15, 0).Value) - 1) Then
			'  CellValues(Summary, 9, x).Value = CellValues(Summary, 9, x).Value + CellValues(DevelopmentResult, 0, x).Value
			'End If
			CellValues(Summary, 10, x).Value = CellValues(DevelopmentResult, 2, x).Value + CellValues(DevelopmentResult, 3, x).Value + CellValues(DevelopmentResult, 4, x).Value + CellValues(DevelopmentResult, 7, x).Value + CellValues(DevelopmentResult, 8, x).Value + CellValues(DevelopmentResult, 9, x).Value
			'CellValues(Summary, 10, x).Value = (EngBase(2) + EngBase(3) + EngBase(4) + EngBase(7)) / ((Int(CellValues(Production, 15, 0).Value) - 1))
		Next x
		
		'=========================== Working Capital Cost  =============================='
		Call scalc()
		
		Call dycal(dyyr)
		
		WorkingBase = 0
		For x = 0 To 11
			WorkingBase = WorkingBase + OutSupply(x, 0)
		Next x
		
		For x = 0 To 8
			WorkingBase = WorkingBase + LaborArray(x)
		Next x
		
		For x = 0 To 11
			WorkingBase = WorkingBase + (((CellValues(Staff, x, 0).Value * CellValues(Salary, x + 12, 0).Value * burden) / annualprod) * ot)
		Next x
		
		If CellValues(Summary, 11, Int(CellValues(Production, 15, 0).Value)).Changed = False Then
			If dyyr <> 0 Then
				CellValues(Summary, 11, Int(CellValues(Production, 15, 0).Value)).Value = (WorkingBase * dyyr) / 6
				CapBin(13) = CellValues(Summary, 11, Int(CellValues(Production, 15, 0).Value)).Value
			End If
		End If
		
		'============================= Engineering Cost ==============================='
		
		Call ecalc()
		EngBase(8) = EquipmentSum
		
		If WhichSegment = 0 Then
			Call dcalc()
		End If
		
		EngBase(9) = 0
		
		For x = 1 To 13
			EngBase(9) = EngBase(9) + DevelopmentArray(x)
		Next x
		
		EngBase(10) = CellValues(Development, 22, WhichSegment).Value
		
		BaseHead = 0
		
		For x = 0 To 10
			BaseHead = BaseHead + EngBase(x)
		Next x
		
		If CellValues(Summary, 12, 1).Changed = False Then
			Select Case (ot + wt)
				Case 0 To 9999
					CellValues(Summary, 12, 1).Value = 8
				Case 10000 To 99999
					CellValues(Summary, 12, 1).Value = 10
				Case Else
					CellValues(Summary, 12, 1).Value = 12
			End Select
		End If
		
		BaseHead = BaseHead * (CellValues(Summary, 12, 1).Value / 100)
		
		For x = 1 To ((Int(CellValues(Production, 15, 0).Value) - 1))
			CellValues(Summary, 12, x).Value = BaseHead / ((Int(CellValues(Production, 15, 0).Value) - 1))
		Next x
		
		'============================= Management Cost ================================'
		
		BaseHead = 0
		
		For x = 1 To 10
			BaseHead = BaseHead + EngBase(x)
		Next x
		
		If CellValues(Summary, 12, 1).Changed = False Then
			Select Case (ot + wt)
				Case 0 To 9999
					TempResult = 7
				Case 10000 To 99999
					TempResult = 8
				Case Else
					TempResult = 9
			End Select
		End If
		
		BaseHead = BaseHead * (TempResult / 100)
		
		For x = 1 To ((Int(CellValues(Production, 15, 0).Value) - 1))
			CellValues(Summary, 12, x).Value = CellValues(Summary, 12, x).Value + (BaseHead / ((Int(CellValues(Production, 15, 0).Value) - 1)))
		Next x
		
	End Sub
	Private Sub ecalc()
		
		Dim x As Short
		Dim TaxRate As Decimal
		
		On Error Resume Next
		
		Call txcal(TaxRate)
		
		EquipmentSum = 0
		
		For x = 0 To 20
			EquipmentSum = EquipmentSum + (CellValues(Purchase, x + 20, WhichSegment).Value * TaxRate)
		Next x
		
		For x = 2 To 3
			RoadSum = RoadSum + (CellValues(Production, x * 3, WhichSegment).Value * CellValues(Development, x + 8, WhichSegment).Value)
		Next x
		
		If WhichSegment = 0 Then
			Call dcalc()
		End If
		
		For x = 1 To 13
			DevelopmentSum = DevelopmentSum + DevelopmentArray(x)
		Next x
		
	End Sub
	Sub tonsp(ByRef TonsPerDay As Decimal)
		
		TonsPerDay = CellValues(Production, 5, WhichSegment).Value + CellValues(Production, 10, WhichSegment).Value
		
	End Sub
	Sub pfcal(ByRef opf As Decimal, ByRef wpf As Decimal)
		
		opf = CellValues(Deposit, 2, WhichSegment).Value
		wpf = CellValues(Deposit, 9, WhichSegment).Value
		
	End Sub
	Sub pmcal(ByRef vol As Decimal, ByRef head As Decimal)
		
		vol = CellValues(Pumping, 0, WhichSegment).Value
		head = CellValues(Pumping, 1, WhichSegment).Value + CellValues(Pumping, 5, WhichSegment).Value
		
	End Sub
	Public Sub TimeLineCalc()
		Dim x As Short
		
		MinTime = CellValues(Production, 15, 0).Value
		
		MaxSegment = 0
		
		For x = 0 To 5
			If CellValues(Production, 16, x).Value > 0 Then
				MaxTime = CellValues(Production, 16, x).Value
				MaxSegment = x
			End If
		Next x
		
	End Sub
	Public Sub ProductionCalc()
        Dim r As Object = Nothing
        Dim s As Object
		Dim x As Short
		
		For x = MinTime To MaxTime
			For s = 0 To 6
				'UPGRADE_WARNING: Couldn't resolve default property of object s. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If x > CellValues(Production, 15, s).Value And x <= CellValues(Production, 16, s).Value Then
					'UPGRADE_WARNING: Couldn't resolve default property of object s. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object r. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					r = s
				End If
			Next s
			'UPGRADE_WARNING: Couldn't resolve default property of object r. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CellValues(Production, 20, x).Value = CellValues(Production, 5, r).Value
		Next x
		
	End Sub
End Module