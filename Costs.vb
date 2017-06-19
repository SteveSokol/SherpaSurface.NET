Option Strict Off
Option Explicit On
Module Costs
	Public Sub crcst(ByRef x As Short)
		Dim crhr As Object
		Dim crwg As Object
        Dim y As Object = Nothing

        Dim MachineType As Short
		Dim PowerDraw As Decimal
		Dim CrusherArray(2, 9) As Decimal
		Dim jump As Boolean
		Dim hrsh As Decimal
		Dim shdy As Decimal
		Dim dyyr As Decimal
		Dim PowerRate As Decimal
		Dim TaxRate As Decimal
		Dim rs As Decimal
		Dim BurdenRate As Decimal
		Dim mcwg As Decimal
		Dim laef As Decimal
		Dim ot As Decimal
		Dim wt As Decimal
		
		On Error Resume Next
		
		Call getout(jump)
		
		If jump = True Then Exit Sub
		
		Call hrcal(hrsh)
		Call shcal(shdy)
		Call dycal(dyyr)
		Call txcal(TaxRate)
		Call elcal(PowerRate)
		Call mccal(BurdenRate, mcwg)
		Call lbcal(laef)
		Call rscal(rs)
		
		If CellValues(EquipmentOne, (x * 5) + 22, WhichSegment).Value <> 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object y. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			y = Int(CellValues(EquipmentOne, (x * 5) + 22, WhichSegment).Value)
			MachineType = Jaw
			CrusherArray(x, Number) = CellValues(EquipmentTwo, (x * 5) + 22, WhichSegment).Value
		ElseIf CellValues(EquipmentOne, (x * 5) + 23, WhichSegment).Value <> 0 Then 
			'UPGRADE_WARNING: Couldn't resolve default property of object y. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			y = Int(CellValues(EquipmentOne, (x * 5) + 23, WhichSegment).Value)
			MachineType = Gyratory
			CrusherArray(x, Number) = CellValues(EquipmentTwo, (x * 5) + 23, WhichSegment).Value
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object crwg. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		crwg = CellValues(Wage, 5, WhichSegment).Value
		
		If CellValues(EquipmentHours, (x * 2) + 14, WhichSegment).Changed = False Then
			'UPGRADE_WARNING: Couldn't resolve default property of object crhr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			crhr = hrsh * shdy * (CellValues(Convey, 3 + (x * 11), WhichSegment).Value / 100)
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object crhr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			crhr = CellValues(EquipmentHours, (x * 2) + 14, WhichSegment).Value
		End If
		
		PowerDraw = (CellValues(Convey, 5 + (x * 11), WhichSegment).Value * 0.746)
		
		EqCost(x, MachineType, OutNumber) = 1
		
		CrusherArray(x, Number) = EqCost(x, MachineType, OutNumber)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object crhr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If crhr <> 0 And CrusherArray(x, Number) <> 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object crhr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			EqCost(x, MachineType, OutLife) = EqDefault(MachineType, y, life) / ((crhr / CrusherArray(x, Number)) * dyyr)
		End If
		
		If CellValues(Purchase, (2 * x) + 14, WhichSegment).Changed = False Then
			'UPGRADE_WARNING: Couldn't resolve default property of object crhr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If crhr <> 0 Then
				CellValues(Purchase, (2 * x) + 14, WhichSegment).Value = EqDefault(MachineType, y, Price)
			Else
				CellValues(Purchase, (2 * x) + 14, WhichSegment).Value = 0
			End If
		End If
		
		If x = Ore Then
			'UPGRADE_WARNING: Couldn't resolve default property of object crhr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If CellValues(EquipmentHours, 14, WhichSegment).Changed = False Then CellValues(EquipmentHours, 14, WhichSegment).Value = crhr
			If CellValues(RepairParts, 14, WhichSegment).Changed = False Then CellValues(RepairParts, 14, WhichSegment).Value = EqDefault(MachineType, y, Parts)
			If CellValues(Electricity, 14, WhichSegment).Changed = False Then CellValues(Electricity, 14, WhichSegment).Value = PowerDraw
			If CellValues(Lubricants, 14, WhichSegment).Changed = False Then CellValues(Lubricants, 14, WhichSegment).Value = EqDefault(MachineType, y, LubeCost)
			If CellValues(RepairLabor, 14, WhichSegment).Changed = False Then CellValues(RepairLabor, 14, WhichSegment).Value = EqDefault(MachineType, y, Mechanic)
			If CellValues(Replace_Renamed, 14, WhichSegment).Changed = False Then CellValues(Replace_Renamed, 14, WhichSegment).Value = CDec(EqCost(x, MachineType, OutLife) * 12)
			If CellValues(Purchase, 34, WhichSegment).Changed = False Then
				CellValues(Purchase, 34, WhichSegment).Value = CellValues(Purchase, 14, WhichSegment).Value * CrusherArray(x, Number)
			End If
			
			EqCost(x, MachineType, OutParts) = CellValues(RepairParts, 14, WhichSegment).Value * CellValues(EquipmentHours, 14, WhichSegment).Value * TaxRate
			EqCost(x, MachineType, OutElectricity) = (CellValues(Electricity, 14, WhichSegment).Value * CellValues(EquipmentHours, 14, WhichSegment).Value * PowerRate)
			EqCost(x, MachineType, OutLube) = CellValues(Lubricants, 14, WhichSegment).Value * CellValues(EquipmentHours, 14, WhichSegment).Value * TaxRate
			'UPGRADE_WARNING: Couldn't resolve default property of object crwg. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			EqCost(x, MachineType, OutOperator) = (Int((CellValues(EquipmentHours, 14, WhichSegment).Value / laef) / hrsh) + 1) * hrsh * crwg * BurdenRate
			EqCost(x, MachineType, OutMechanicTime) = CellValues(RepairLabor, 14, WhichSegment).Value * CellValues(EquipmentHours, 14, WhichSegment).Value
			EqCost(x, MachineType, OutMechanicCost) = EqCost(x, MachineType, OutMechanicTime) * mcwg * BurdenRate
			
			EqCost(x, MachineType, OutOwn) = ((CrusherArray(x, Number) * EqDefault(Conveyor, y, Price)) / rs) * CrusherArray(x, Ton)
			
			OutEquipment(14, WhichSegment) = (EqCost(x, MachineType, OutParts) + EqCost(x, MachineType, OutElectricity) + EqCost(x, MachineType, OutLube))
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object crhr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If CellValues(EquipmentHours, 16, WhichSegment).Changed = False Then CellValues(EquipmentHours, 16, WhichSegment).Value = crhr
			If CellValues(RepairParts, 16, WhichSegment).Changed = False Then CellValues(RepairParts, 16, WhichSegment).Value = EqDefault(MachineType, y, Parts)
			If CellValues(Electricity, 16, WhichSegment).Changed = False Then CellValues(Electricity, 16, WhichSegment).Value = PowerDraw
			If CellValues(Lubricants, 16, WhichSegment).Changed = False Then CellValues(Lubricants, 16, WhichSegment).Value = EqDefault(MachineType, y, LubeCost)
			If CellValues(RepairLabor, 16, WhichSegment).Changed = False Then CellValues(RepairLabor, 16, WhichSegment).Value = EqDefault(MachineType, y, Mechanic)
			If CellValues(Replace_Renamed, 16, WhichSegment).Changed = False Then CellValues(Replace_Renamed, 16, WhichSegment).Value = CDec(EqCost(x, MachineType, OutLife) * 12)
			If CellValues(Purchase, 36, WhichSegment).Changed = False Then
				CellValues(Purchase, 36, WhichSegment).Value = CellValues(Purchase, 16, WhichSegment).Value * CrusherArray(x, Number)
			End If
			
			EqCost(x, MachineType, OutParts) = CellValues(RepairParts, 16, WhichSegment).Value * CellValues(EquipmentHours, 16, WhichSegment).Value * TaxRate
			EqCost(x, MachineType, OutElectricity) = CellValues(Electricity, 16, WhichSegment).Value * CellValues(EquipmentHours, 16, WhichSegment).Value * PowerRate
			EqCost(x, MachineType, OutLube) = CellValues(Lubricants, 16, WhichSegment).Value * CellValues(EquipmentHours, 16, WhichSegment).Value * TaxRate
			'UPGRADE_WARNING: Couldn't resolve default property of object crwg. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			EqCost(x, MachineType, OutOperator) = (Int((CellValues(EquipmentHours, 16, WhichSegment).Value / laef) / hrsh) + 1) * hrsh * crwg * BurdenRate
			EqCost(x, MachineType, OutMechanicTime) = CellValues(RepairLabor, 16, WhichSegment).Value * CellValues(EquipmentHours, 16, WhichSegment).Value
			
			EqCost(x, MachineType, OutMechanicCost) = EqCost(x, MachineType, OutMechanicTime) * mcwg * BurdenRate
			EqCost(x, MachineType, OutOwn) = ((CrusherArray(x, Number) * EqDefault(Conveyor, y, Price)) / rs) * CrusherArray(x, Ton)
			
			OutEquipment(16, WhichSegment) = (EqCost(x, MachineType, OutParts) + EqCost(x, MachineType, OutElectricity) + EqCost(x, MachineType, OutLube))
		End If
		
		EqCost(x, Conveyor, OutNumber) = 1
		
		CrusherArray(x, Number) = EqCost(x, Conveyor, OutNumber)
		
	End Sub
	Public Sub cocst(ByRef x As Short)
        Dim Conveyore As Object = Nothing
        Dim AddCost As Object
		
		Dim BaseCost As Decimal
		Dim bWaste As Short
		Dim y As Short
		Dim z As Short
		
		Dim jump As Boolean
		
		Dim ConveyArray(2, 9) As Decimal
		Dim TravelArray(2, 4, 6) As Decimal
		Dim TemperTime(2) As Decimal
		Dim ConveyFill(2, 6) As Decimal
		
		Dim Length As Decimal
		Dim Gradient As Decimal
		Dim lift As Decimal
		Dim ConveyWidth As Decimal
		Dim hp1 As Decimal
		Dim hp2 As Decimal
		Dim hp3 As Decimal
		
		Dim a As Decimal
		Dim b As Decimal
		Dim cohr As Decimal
		Dim cowg As Decimal
		Dim hrsh As Decimal
		Dim shdy As Decimal
		Dim dyyr As Decimal
		Dim PowerRate As Decimal
		Dim fuel As Decimal
		Dim tph As Decimal
		Dim PowerDraw As Decimal
		
		Dim TaxRate As Decimal
		Dim rs As Decimal
		Dim BurdenRate As Decimal
		Dim mcwg As Decimal
		Dim laef As Decimal
		
		On Error Resume Next
		
		Call getout(jump)
		If jump = True Then Exit Sub
		
		y = Int(CellValues(EquipmentOne, (x * 5) + 24, WhichSegment).Value)
		
		ConveyWidth = Val(Left(CellValues(EquipmentOne, (x * 5) + 24, WhichSegment).Word, 2))
		
		Call hrcal(hrsh)
		Call shcal(shdy)
		Call dycal(dyyr)
		Call fucal(fuel)
		Call txcal(TaxRate)
		Call rscal(rs)
		Call elcal(PowerRate)
		Call mccal(BurdenRate, mcwg)
		Call lbcal(laef)
		
		cowg = CellValues(Wage, 5, WhichSegment).Value
		
		Length = CellValues(Convey, 7 + (x * 11), WhichSegment).Value
		Gradient = CellValues(Convey, 8 + (x * 11), WhichSegment).Value
		lift = Length * (Gradient / 100)
		tph = CellValues(Production, 5 + (x * 5), WhichSegment).Value / (hrsh * shdy)
		hp1 = (0.072 * ConveyWidth)
		hp1 = hp1 * (Length / 1000) * 2.5
		hp2 = ((0.033 * tph) * (Length / 1000))
		hp3 = (tph * lift) / 990
		If hp3 <= hp1 + hp2 Then
			a = (1.17647 * ((hp1 + hp2) - ((2 * hp3) / 3)))
		Else
			a = (0.85 * (hp3 - ((hp1 + hp2) / 2)))
		End If
		PowerDraw = a
		b = 1.17647 * hp1
		If b > a Then PowerDraw = b
		
		If CellValues(EquipmentHours, (2 * x) + 15, WhichSegment).Changed = False Then
			cohr = hrsh * shdy * (CellValues(Convey, 10 + (x * 11), WhichSegment).Value / 100)
		Else
			cohr = CellValues(EquipmentHours, (2 * x) + 15, WhichSegment).Value
		End If
		
		BaseCost = EqDefault(Conveyor, y, Price)
		'UPGRADE_WARNING: Couldn't resolve default property of object AddCost. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		AddCost = EqDefault(Conveyor, y, ConveyAdd)
		'UPGRADE_WARNING: Couldn't resolve default property of object AddCost. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		BaseCost = BaseCost + (AddCost * (Length - 5300))
		
		If CellValues(Purchase, (2 * x) + 15, WhichSegment).Changed = False Then
			If cohr <> 0 Then
				CellValues(Purchase, (2 * x) + 15, WhichSegment).Value = BaseCost
			Else
				CellValues(Purchase, (2 * x) + 15, WhichSegment).Value = 0
			End If
		End If
		
		EqCost(x, Conveyor, OutNumber) = 1
		
		ConveyArray(x, Number) = EqCost(x, Conveyor, OutNumber)
		
		If cohr <> 0 And ConveyArray(x, Number) <> 0 Then
			EqCost(x, Conveyor, OutLife) = EqDefault(Conveyor, y, life) / ((cohr / ConveyArray(x, Number)) * dyyr)
		End If
		
		If x = Ore Then
			If CellValues(EquipmentHours, 15, WhichSegment).Changed = False Then CellValues(EquipmentHours, 15, WhichSegment).Value = cohr
			If CellValues(RepairParts, 15, WhichSegment).Changed = False Then CellValues(RepairParts, 15, WhichSegment).Value = ((BaseCost * 0.2) / 8400)
			If CellValues(Electricity, 15, WhichSegment).Changed = False Then CellValues(Electricity, 15, WhichSegment).Value = PowerDraw * 0.746
			If CellValues(Lubricants, 15, WhichSegment).Changed = False Then CellValues(Lubricants, 15, WhichSegment).Value = (BaseCost / 175000)
			If CellValues(RepairLabor, 15, WhichSegment).Changed = False Then CellValues(RepairLabor, 15, WhichSegment).Value = CellValues(RepairParts, 15, WhichSegment).Value * 0.048
			If CellValues(Replace_Renamed, 15, WhichSegment).Changed = False Then CellValues(Replace_Renamed, 15, WhichSegment).Value = CDec(EqCost(x, Conveyor, OutLife) * 12)
			If CellValues(EquipmentTwo, 24, WhichSegment).Changed = False Then CellValues(EquipmentTwo, 24, WhichSegment).Value = CDec(EqCost(x, Conveyor, OutNumber))
			If CellValues(Purchase, 35, WhichSegment).Changed = False Then
				CellValues(Purchase, 35, WhichSegment).Value = CellValues(Purchase, 15, WhichSegment).Value * CellValues(EquipmentTwo, 24, WhichSegment).Value
			End If
			EqCost(x, Conveyor, OutParts) = CellValues(RepairParts, 15, WhichSegment).Value * CellValues(EquipmentHours, 15, WhichSegment).Value * TaxRate
			EqCost(x, Conveyor, OutElectricity) = CellValues(Electricity, 15, WhichSegment).Value * CellValues(EquipmentHours, 15, WhichSegment).Value * PowerRate
			EqCost(x, Conveyor, OutLube) = CellValues(Lubricants, 15, WhichSegment).Value * CellValues(EquipmentHours, 15, WhichSegment).Value * TaxRate
			EqCost(x, Conveyor, OutOperator) = (Int((CellValues(EquipmentHours, 15, WhichSegment).Value / laef) / hrsh) + 1) * hrsh * cowg * BurdenRate
			EqCost(x, Conveyor, OutMechanicTime) = CellValues(RepairLabor, 15, WhichSegment).Value * CellValues(EquipmentHours, 15, WhichSegment).Value
			EqCost(x, Conveyor, OutMechanicCost) = EqCost(x, Conveyor, OutMechanicTime) * mcwg * BurdenRate
			EqCost(x, Conveyor, OutOwn) = ((ConveyArray(x, Number) * EqDefault(Conveyor, y, Price)) / rs) * ConveyArray(x, Ton)
			If ConveyArray(x, Ton) <> 0 Then
				'  EqCost(x, Conveyor, OutUnit) = (EqCost(x, Conveyor, OutParts) + EqCost(x, Conveyor, OutElectricity) + EqCost(x, Conveyor, OutLube) + EqCost(x, Conveyor, OutOperator) + EqCost(x, Conveyor, OutMechanicCost) + EqCost(x, Conveyor, OutOwn)) / ConveyArray(x, Ton)
			End If
			EqCost(x, Conveyor, OutUnit) = EqCost(x, Conveyor, OutUnit) + ((ConveyArray(x, Number) / 25) * EqCost(x, Conveyor, OutUnit))
			'UPGRADE_WARNING: Couldn't resolve default property of object Conveyore. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			OutEquipment(15, WhichSegment) = (EqCost(x, Conveyor, OutParts) + EqCost(x, Conveyore, OutElectricity) + EqCost(x, Conveyor, OutLube))
		Else
			If CellValues(EquipmentHours, 17, WhichSegment).Changed = False Then CellValues(EquipmentHours, 17, WhichSegment).Value = cohr
			If CellValues(RepairParts, 17, WhichSegment).Changed = False Then CellValues(RepairParts, 17, WhichSegment).Value = ((BaseCost * 0.2) / 8400)
			If CellValues(Electricity, 17, WhichSegment).Changed = False Then CellValues(Electricity, 17, WhichSegment).Value = PowerDraw * 0.746
			If CellValues(Lubricants, 17, WhichSegment).Changed = False Then CellValues(Lubricants, 17, WhichSegment).Value = (BaseCost / 175000)
			If CellValues(RepairLabor, 17, WhichSegment).Changed = False Then CellValues(RepairLabor, 17, WhichSegment).Value = CellValues(RepairParts, 17, WhichSegment).Value * 0.048
			If CellValues(Replace_Renamed, 17, WhichSegment).Changed = False Then CellValues(Replace_Renamed, 17, WhichSegment).Value = CDec(EqCost(x, Conveyor, OutLife) * 12)
			If CellValues(EquipmentTwo, 29, WhichSegment).Changed = False Then CellValues(EquipmentTwo, 29, WhichSegment).Value = CDec(EqCost(x, Conveyor, OutNumber))
			If CellValues(Purchase, 37, WhichSegment).Changed = False Then
				CellValues(Purchase, 37, WhichSegment).Value = CellValues(Purchase, 17, WhichSegment).Value * CellValues(EquipmentTwo, 29, WhichSegment).Value
			End If
			EqCost(x, Conveyor, OutParts) = CellValues(RepairParts, 17, WhichSegment).Value * CellValues(EquipmentHours, 17, WhichSegment).Value * TaxRate
			EqCost(x, Conveyor, OutElectricity) = CellValues(Electricity, 17, WhichSegment).Value * CellValues(EquipmentHours, 17, WhichSegment).Value * PowerRate
			EqCost(x, Conveyor, OutLube) = CellValues(Lubricants, 17, WhichSegment).Value * CellValues(EquipmentHours, 17, WhichSegment).Value * TaxRate
			EqCost(x, Conveyor, OutOperator) = (Int((CellValues(EquipmentHours, 17, WhichSegment).Value / laef) / hrsh) + 1) * hrsh * cowg * BurdenRate
			EqCost(x, Conveyor, OutMechanicTime) = CellValues(RepairLabor, 17, WhichSegment).Value * CellValues(EquipmentHours, 17, WhichSegment).Value
			EqCost(x, Conveyor, OutMechanicCost) = EqCost(x, Conveyor, OutMechanicTime) * mcwg * BurdenRate
			EqCost(x, Conveyor, OutOwn) = ((ConveyArray(x, Number) * EqDefault(Conveyor, y, Price)) / rs) * ConveyArray(x, Ton)
			If ConveyArray(x, Ton) <> 0 Then
				'  EqCost(x, Conveyor, OutUnit) = (EqCost(x, Conveyor, OutParts) + EqCost(x, Conveyor, OutElectricity) + EqCost(x, Conveyor, OutLube) + EqCost(x, Conveyor, OutOperator) + EqCost(x, Conveyor, OutMechanicCost) + EqCost(x, Conveyor, OutOwn)) / ConveyArray(x, Ton)
			End If
			EqCost(x, Conveyor, OutUnit) = EqCost(x, Conveyor, OutUnit) + ((ConveyArray(x, Number) / 25) * EqCost(x, Conveyor, OutUnit))
			'UPGRADE_WARNING: Couldn't resolve default property of object Conveyore. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			OutEquipment(17, WhichSegment) = (EqCost(x, Conveyor, OutParts) + EqCost(x, Conveyore, OutElectricity) + EqCost(x, Conveyor, OutLube))
		End If
		
	End Sub
	Sub prcst()
        Dim MaxNumOut As Object = Nothing
        Dim x As Object = Nothing
        Dim BurdenRate As Object = Nothing
        Dim mcwg As Object = Nothing

        Dim f As Short
		
		Dim hrsh As Decimal
		Dim shdy As Decimal
		Dim dyyr As Decimal
		Dim TaxRate As Decimal
		Dim laef As Decimal
		Dim fuel As Decimal
		Dim ot As Decimal
		Dim wt As Decimal
		Dim opf As Decimal
		Dim wpf As Decimal
		Dim opsg As Decimal
		Dim wpsg As Decimal
		Dim olf As Decimal
		Dim wlf As Decimal
		Dim odr As Decimal
		Dim wdr As Decimal
		Dim prlbpr As Decimal
		Dim prptpr As Decimal
		Dim prexporedaylb As Decimal
		Dim prexpwasdaylb As Decimal
		Dim prexporedayvol As Decimal
		Dim prexpwasdayvol As Decimal
		Dim prorevoldri As Decimal
		Dim prwasvoldri As Decimal
		Dim prohr As Decimal
		Dim prwhr As Decimal
		
		Dim jump As Boolean
		
		On Error Resume Next
		
		Call getout(jump)
		
		If jump = True Then Exit Sub
		
		Call hrcal(hrsh)
		Call shcal(shdy)
		Call dycal(dyyr)
		Call txcal(TaxRate)
		Call lbcal(laef)
		Call fucal(fuel)
		Call otcal(ot, wt)
		Call pfcal(opf, wpf)
		Call sgcal(opsg, wpsg)
		Call lfcal(olf, wlf)
		Call drcal(odr, wdr)
		
		'Ore
		
		f = CellValues(EquipmentOne, 10, WhichSegment).Value
		
		If f > 0 Then
			prlbpr = ((EqDefault(Percussion, f, Price) * 0.0000063) + (EqDefault(Percussion, f, DrillFuel) * fuel * 0.1))
			
			prptpr = ((EqDefault(Percussion, f, Price) * 0.75 * 0.55) / EqDefault(Percussion, f, life))
			
			prexporedaylb = opf * ot
			prexporedayvol = prexporedaylb / (opsg * 62.4)
			prorevoldri = prexporedayvol / (olf / 100)
			
			EqCost(Ore, Percussion, OutFeet) = prorevoldri / ((3.141593 * (EqDefault(Percussion, f, HoleDiameter) / 12) ^ 2) / 4)
			
			CellValues(Powder, 25, WhichSegment).Value = EqCost(Ore, Percussion, OutFeet)
			
			If opf > 0 Then
				prohr = EqCost(Ore, Percussion, OutFeet) / odr
			Else
				prohr = 0
			End If
			
			EqCost(Ore, Percussion, OutHours) = prohr
			
			If opf > 0 Then EqCost(Ore, Percussion, OutNumber) = Int(((EqCost(Ore, Percussion, OutHours) / laef) / 0.92) / (hrsh * shdy)) + 1
			
			If EqCost(Ore, Percussion, OutHours) > 0 And EqCost(Ore, Percussion, OutNumber) > 0 Then
				EqCost(Ore, Percussion, OutLife) = EqDefault(Percussion, f, life) / ((EqCost(Ore, Percussion, OutHours) / EqCost(Ore, Percussion, OutNumber)) * dyyr)
			End If
			
			If EqCost(Ore, Percussion, OutHours) <> 0 Then
				If CellValues(EquipmentHours, 4, WhichSegment).Changed = False Then CellValues(EquipmentHours, 4, WhichSegment).Value = EqCost(Ore, Percussion, OutHours)
				If CellValues(RepairParts, 4, WhichSegment).Changed = False Then CellValues(RepairParts, 4, WhichSegment).Value = prptpr
				If CellValues(Diesel, 4, WhichSegment).Changed = False Then CellValues(Diesel, 4, WhichSegment).Value = EqDefault(Percussion, f, DrillFuel)
				If CellValues(Lubricants, 4, WhichSegment).Changed = False Then CellValues(Lubricants, 4, WhichSegment).Value = prlbpr
				If CellValues(RepairLabor, 4, WhichSegment).Changed = False Then CellValues(RepairLabor, 4, WhichSegment).Value = EqDefault(Percussion, f, DrillMechanic)
				If CellValues(Powder, 13, WhichSegment).Changed = False Then CellValues(Powder, 13, WhichSegment).Value = EqCost(Ore, Percussion, OutBit)
				If CellValues(Powder, 14, WhichSegment).Changed = False Then CellValues(Powder, 14, WhichSegment).Value = EqCost(Ore, Percussion, OutSteel)
			End If
			
			If CellValues(Purchase, 4, WhichSegment).Changed = False Then
				CellValues(Purchase, 4, WhichSegment).Value = EqDefault(Percussion, f, Price)
			End If
			
			EqCost(Ore, Percussion, OutParts) = CellValues(RepairParts, 4, WhichSegment).Value * CellValues(EquipmentHours, 4, WhichSegment).Value * TaxRate
			OpBin(4) = OpBin(4) + EqCost(Ore, Percussion, OutParts)
			EqCost(Ore, Percussion, OutFuel) = fuel * CellValues(Diesel, 4, WhichSegment).Value * CellValues(EquipmentHours, 4, WhichSegment).Value
			EqCost(Ore, Percussion, OutLube) = CellValues(Lubricants, 4, WhichSegment).Value * CellValues(EquipmentHours, 4, WhichSegment).Value * TaxRate
			OpBin(2) = OpBin(2) + EqCost(Ore, Percussion, OutLube)
			EqCost(Ore, Percussion, OutMechanicTime) = CellValues(EquipmentHours, 4, WhichSegment).Value * CellValues(RepairLabor, 4, WhichSegment).Value
			'UPGRADE_WARNING: Couldn't resolve default property of object BurdenRate. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object mcwg. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			EqCost(Ore, Percussion, OutMechanicCost) = EqCost(Ore, Percussion, OutMechanicTime) * mcwg * BurdenRate
			EqCost(Ore, Percussion, OutBit) = EqDefault(Percussion, f, BitCost)
			EqCost(Ore, Percussion, OutSteel) = EqDefault(Percussion, f, SteelCost)
			EqCost(Ore, Percussion, OutPrice) = CellValues(Purchase, 4, WhichSegment).Value
			
			OutEquipment(4, WhichSegment) = EqCost(Ore, Percussion, OutParts) + EqCost(Ore, Percussion, OutFuel) + EqCost(Ore, Percussion, OutLube)
			
			If CellValues(Replace_Renamed, 4, WhichSegment).Changed = False Then
				CellValues(Replace_Renamed, 4, WhichSegment).Value = EqCost(Ore, Percussion, OutLife) * 12
			End If
			
			If CellValues(EquipmentTwo, 10, WhichSegment).Changed = False Then
				CellValues(EquipmentTwo, 10, WhichSegment).Value = EqCost(Ore, Percussion, OutNumber)
			End If
			
			If CellValues(Purchase, 24, WhichSegment).Changed = False Then
				CellValues(Purchase, 24, WhichSegment).Value = CellValues(Purchase, 4, WhichSegment).Value * CellValues(EquipmentTwo, 10, WhichSegment).Value
			End If
			
			'UPGRADE_WARNING: Couldn't resolve default property of object MaxNumOut. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			For x = 0 To MaxNumOut
				'UPGRADE_WARNING: Couldn't resolve default property of object x. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				EqCost(0, Rotary, x) = 0
			Next x
			CellValues(EquipmentTwo, 11, WhichSegment).Value = 0
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object MaxNumOut. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			For x = 0 To MaxNumOut
				'UPGRADE_WARNING: Couldn't resolve default property of object x. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				EqCost(0, Percussion, x) = 0
			Next x
		End If
		
		'Waste
		
		f = CellValues(EquipmentOne, 30, WhichSegment).Value
		
		If f > 0 Then
			prlbpr = ((EqDefault(Percussion, f, Price) * 0.0000063) + (EqDefault(Percussion, f, DrillFuel) * fuel * 0.1))
			
			prptpr = ((EqDefault(Percussion, f, Price) * 0.75 * 0.55) / EqDefault(Percussion, f, life))
			
			prexpwasdaylb = wpf * wt
			prexpwasdayvol = prexpwasdaylb / (wpsg * 62.4)
			If wlf <> 0 Then
				prwasvoldri = prexpwasdayvol / (wlf / 100)
			End If
			
			EqCost(Waste, Percussion, OutFeet) = prwasvoldri / ((3.141593 * (EqDefault(Percussion, f, HoleDiameter) / 12) ^ 2) / 4)
			
			CellValues(Powder, 15, WhichSegment).Value = EqCost(Waste, Percussion, OutFeet)
			
			If wpf > 0 Then
				prwhr = EqCost(Waste, Percussion, OutFeet) / wdr
			Else
				prwhr = 0
			End If
			
			EqCost(Waste, Percussion, OutHours) = prwhr
			
			If wpf > 0 Then EqCost(Waste, Percussion, OutNumber) = Int(((EqCost(Waste, Percussion, OutHours) / laef) / 0.92) / (hrsh * shdy)) + 1
			
			
			If EqCost(Waste, Percussion, OutHours) > 0 And EqCost(Waste, Percussion, OutNumber) > 0 Then
				EqCost(Waste, Percussion, OutLife) = EqDefault(Percussion, f, life) / ((EqCost(Waste, Percussion, OutHours) / EqCost(Waste, Percussion, OutNumber)) * dyyr)
			End If
			
			If EqCost(Waste, Percussion, OutHours) <> 0 Then
				If CellValues(EquipmentHours, 18, WhichSegment).Changed = False Then CellValues(EquipmentHours, 18, WhichSegment).Value = EqCost(Waste, Percussion, OutHours)
				If CellValues(RepairParts, 18, WhichSegment).Changed = False Then CellValues(RepairParts, 18, WhichSegment).Value = prptpr
				If CellValues(Diesel, 18, WhichSegment).Changed = False Then CellValues(Diesel, 18, WhichSegment).Value = EqDefault(Percussion, f, DrillFuel)
				If CellValues(Lubricants, 18, WhichSegment).Changed = False Then CellValues(Lubricants, 18, WhichSegment).Value = prlbpr
				If CellValues(RepairLabor, 18, WhichSegment).Changed = False Then CellValues(RepairLabor, 18, WhichSegment).Value = EqDefault(Percussion, f, DrillMechanic)
				If CellValues(Powder, 16, WhichSegment).Changed = False Then CellValues(Powder, 16, WhichSegment).Value = EqCost(Waste, Percussion, OutBit)
				If CellValues(Powder, 17, WhichSegment).Changed = False Then CellValues(Powder, 17, WhichSegment).Value = EqCost(Waste, Percussion, OutSteel)
			End If
			
			If CellValues(Purchase, 18, WhichSegment).Changed = False Then
				CellValues(Purchase, 18, WhichSegment).Value = EqDefault(Percussion, f, Price)
			End If
			
			EqCost(Waste, Percussion, OutParts) = CellValues(RepairParts, 18, WhichSegment).Value * CellValues(EquipmentHours, 18, WhichSegment).Value * TaxRate
			OpBin(4) = OpBin(4) + EqCost(Waste, Percussion, OutParts)
			EqCost(Waste, Percussion, OutFuel) = fuel * CellValues(Diesel, 18, WhichSegment).Value * CellValues(EquipmentHours, 18, WhichSegment).Value
			EqCost(Waste, Percussion, OutLube) = CellValues(Lubricants, 18, WhichSegment).Value * CellValues(EquipmentHours, 18, WhichSegment).Value * TaxRate
			OpBin(2) = OpBin(2) + EqCost(Waste, Percussion, OutLube)
			EqCost(Waste, Percussion, OutMechanicTime) = CellValues(EquipmentHours, 18, WhichSegment).Value * CellValues(RepairLabor, 18, WhichSegment).Value
			'UPGRADE_WARNING: Couldn't resolve default property of object BurdenRate. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object mcwg. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			EqCost(Waste, Percussion, OutMechanicCost) = EqCost(Waste, Percussion, OutMechanicTime) * mcwg * BurdenRate
			EqCost(Waste, Percussion, OutBit) = EqDefault(Percussion, f, BitCost)
			EqCost(Waste, Percussion, OutSteel) = EqDefault(Percussion, f, SteelCost)
			EqCost(Waste, Percussion, OutPrice) = CellValues(Purchase, 18, WhichSegment).Value
			
			OutEquipment(18, WhichSegment) = EqCost(Waste, Percussion, OutParts) + EqCost(Waste, Percussion, OutFuel) + EqCost(Waste, Percussion, OutLube)
			
			If CellValues(Replace_Renamed, 18, WhichSegment).Changed = False Then
				CellValues(Replace_Renamed, 18, WhichSegment).Value = EqCost(Waste, Percussion, OutLife) * 12
			End If
			
			If CellValues(EquipmentTwo, 30, WhichSegment).Changed = False Then
				CellValues(EquipmentTwo, 30, WhichSegment).Value = EqCost(Waste, Percussion, OutNumber)
			End If
			
			If CellValues(Purchase, 38, WhichSegment).Changed = False Then
				CellValues(Purchase, 38, WhichSegment).Value = CellValues(Purchase, 18, WhichSegment).Value * CellValues(EquipmentTwo, 30, WhichSegment).Value
			End If
			
			'UPGRADE_WARNING: Couldn't resolve default property of object MaxNumOut. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			For x = 0 To MaxNumOut
				'UPGRADE_WARNING: Couldn't resolve default property of object x. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				EqCost(1, Rotary, x) = 0
			Next x
			CellValues(EquipmentTwo, 31, WhichSegment).Value = 0
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object MaxNumOut. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			For x = 0 To MaxNumOut
				'UPGRADE_WARNING: Couldn't resolve default property of object x. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				EqCost(1, Percussion, x) = 0
			Next x
		End If
		
	End Sub
	Sub rtcst()
        Dim x As Object = Nothing
        Dim MaxNumOut As Object = Nothing
        Dim y As Object = Nothing
        Dim BurdenRate As Object = Nothing
        Dim mcwg As Object = Nothing

        Dim g As Short
		
		Dim hrsh As Decimal
		Dim shdy As Decimal
		Dim dyyr As Decimal
		Dim TaxRate As Decimal
		Dim laef As Decimal
		Dim fuel As Decimal
		Dim ot As Decimal
		Dim wt As Decimal
		Dim opf As Decimal
		Dim wpf As Decimal
		Dim opsg As Decimal
		Dim wpsg As Decimal
		Dim olf As Decimal
		Dim wlf As Decimal
		Dim odr As Decimal
		Dim wdr As Decimal
		Dim rtlbpr As Decimal
		Dim rtptpr As Decimal
		Dim rtexporedaylb As Decimal
		Dim rtexpwasdaylb As Decimal
		Dim rtexporedayvol As Decimal
		Dim rtexpwasdayvol As Decimal
		Dim rtorevoldri As Decimal
		Dim rtwasvoldri As Decimal
		Dim rtoreuse As Decimal
		Dim rtwasuse As Decimal
		
		Dim jump As Boolean
		
		On Error Resume Next
		
		Call getout(jump)
		
		If jump = True Then Exit Sub
		
		Call hrcal(hrsh)
		Call shcal(shdy)
		Call dycal(dyyr)
		Call txcal(TaxRate)
		Call lbcal(laef)
		Call fucal(fuel)
		Call otcal(ot, wt)
		Call pfcal(opf, wpf)
		Call sgcal(opsg, wpsg)
		Call lfcal(olf, wlf)
		Call drcal(odr, wdr)
		
		'Ore
		
		g = CellValues(EquipmentOne, 11, WhichSegment).Value
		
		If g > 0 Then
			
			rtlbpr = ((EqDefault(Rotary, g, Price) * 0.0000063) + (EqDefault(Rotary, g, DrillFuel) * fuel * 0.1))
			
			rtptpr = ((EqDefault(Rotary, g, Price) * 0.75 * 0.55) / EqDefault(Rotary, g, life))
			
			rtexporedaylb = opf * ot
			rtexporedayvol = rtexporedaylb / (opsg * 62.4)
			rtorevoldri = rtexporedayvol / (olf / 100)
			
			EqCost(Ore, Rotary, OutFeet) = rtorevoldri / ((3.141593 * (EqDefault(Rotary, g, HoleDiameter) / 12) ^ 2) / 4)
			
			CellValues(Powder, 18, WhichSegment).Value = EqCost(Ore, Rotary, OutFeet)
			
			If opf > 0 Then
				rtoreuse = EqCost(Ore, Rotary, OutFeet) / odr
			Else
				rtoreuse = 0
			End If
			
			EqCost(Ore, Rotary, OutHours) = rtoreuse
			
			If opf > 0 Then EqCost(Ore, Rotary, OutNumber) = Int(((EqCost(Ore, Rotary, OutHours) / laef) / 0.92) / (hrsh * shdy)) + 1
			
			If EqCost(Ore, Rotary, OutHours) > 0 And EqCost(Ore, Rotary, OutNumber) > 0 Then
				EqCost(Ore, Rotary, OutLife) = EqDefault(Rotary, g, life) / ((EqCost(Ore, Rotary, OutHours) / EqCost(Ore, Rotary, OutNumber)) * dyyr)
			End If
			
			If EqCost(Ore, Rotary, OutHours) > 0 Then
				If CellValues(EquipmentHours, 5, WhichSegment).Changed = False Then CellValues(EquipmentHours, 5, WhichSegment).Value = EqCost(Ore, Rotary, OutHours)
				If CellValues(RepairParts, 5, WhichSegment).Changed = False Then CellValues(RepairParts, 5, WhichSegment).Value = rtptpr
				If CellValues(Diesel, 5, WhichSegment).Changed = False Then CellValues(Diesel, 5, WhichSegment).Value = EqDefault(Rotary, g, DrillFuel)
				If CellValues(Lubricants, 5, WhichSegment).Changed = False Then CellValues(Lubricants, 5, WhichSegment).Value = rtlbpr
				If CellValues(RepairLabor, 5, WhichSegment).Changed = False Then CellValues(RepairLabor, 5, WhichSegment).Value = EqDefault(Rotary, g, DrillMechanic)
				If CellValues(Powder, 19, WhichSegment).Changed = False Then CellValues(Powder, 19, WhichSegment).Value = EqCost(Ore, Rotary, OutBit)
				If CellValues(Powder, 20, WhichSegment).Changed = False Then CellValues(Powder, 20, WhichSegment).Value = EqCost(Ore, Rotary, OutSteel)
			End If
			
			If CellValues(Purchase, 5, WhichSegment).Changed = False Then
				CellValues(Purchase, 5, WhichSegment).Value = EqDefault(Rotary, g, Price)
			End If
			
			EqCost(Ore, Rotary, OutParts) = CellValues(RepairParts, 5, WhichSegment).Value * CellValues(EquipmentHours, 5, WhichSegment).Value * TaxRate
			OpBin(4) = OpBin(4) + EqCost(Ore, Rotary, OutParts)
			EqCost(Ore, Rotary, OutFuel) = fuel * CellValues(Diesel, 5, WhichSegment).Value * CellValues(EquipmentHours, 5, WhichSegment).Value
			EqCost(Ore, Rotary, OutLube) = CellValues(Lubricants, 5, WhichSegment).Value * CellValues(EquipmentHours, 5, WhichSegment).Value * TaxRate
			OpBin(2) = OpBin(2) + EqCost(Ore, Rotary, OutLube)
			EqCost(Ore, Rotary, OutMechanicTime) = CellValues(EquipmentHours, 5, WhichSegment).Value * CellValues(RepairLabor, 5, WhichSegment).Value
			'UPGRADE_WARNING: Couldn't resolve default property of object BurdenRate. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object mcwg. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			EqCost(Ore, Rotary, OutMechanicCost) = EqCost(Ore, Rotary, OutMechanicTime) * mcwg * BurdenRate
			EqCost(Ore, Rotary, OutBit) = EqDefault(Rotary, g, BitCost)
			EqCost(Ore, Rotary, OutSteel) = EqDefault(Rotary, g, SteelCost)
			EqCost(Ore, Rotary, OutPrice) = CellValues(Purchase, 5, WhichSegment).Value
			
			OutEquipment(5, WhichSegment) = EqCost(Ore, Rotary, OutParts) + EqCost(Ore, Rotary, OutFuel) + EqCost(Ore, Rotary, OutLube)
			
			If CellValues(Replace_Renamed, 5, WhichSegment).Changed = False Then
				CellValues(Replace_Renamed, 5, WhichSegment).Value = EqCost(Ore, Rotary, OutLife) * 12
			End If
			
			If CellValues(EquipmentTwo, 11, WhichSegment).Changed = False Then
				CellValues(EquipmentTwo, 11, WhichSegment).Value = EqCost(Ore, Rotary, OutNumber)
			End If
			
			If CellValues(Purchase, 25, WhichSegment).Changed = False Then
				CellValues(Purchase, 25, WhichSegment).Value = CellValues(Purchase, 5, WhichSegment).Value * CellValues(EquipmentTwo, 11, WhichSegment).Value
			End If
			
			'UPGRADE_WARNING: Couldn't resolve default property of object MaxNumOut. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			For y = 0 To MaxNumOut
				'UPGRADE_WARNING: Couldn't resolve default property of object y. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				EqCost(Ore, Percussion, y) = 0
			Next y
			CellValues(EquipmentTwo, 10, WhichSegment).Value = 0
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object MaxNumOut. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			For y = 0 To MaxNumOut
				'UPGRADE_WARNING: Couldn't resolve default property of object y. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				EqCost(Ore, Rotary, y) = 0
			Next y
		End If
		
		'Waste
		
		g = CellValues(EquipmentOne, 31, WhichSegment).Value
		
		If g > 0 Then
			rtlbpr = ((EqDefault(Rotary, g, Price) * 0.0000063) + (EqDefault(Rotary, g, DrillFuel) * fuel * 0.1))
			
			rtptpr = ((EqDefault(Rotary, g, Price) * 0.75 * 0.55) / EqDefault(Rotary, g, life))
			
			rtexpwasdaylb = wpf * wt
			rtexpwasdayvol = rtexpwasdaylb / (wpsg * 62.4)
			If wlf > 0 Then
				rtwasvoldri = rtexpwasdayvol / (wlf / 100)
			End If
			
			EqCost(Waste, Rotary, OutFeet) = rtwasvoldri / ((3.141593 * (EqDefault(Rotary, g, HoleDiameter) / 12) ^ 2) / 4)
			
			CellValues(Powder, 21, WhichSegment).Value = EqCost(Waste, Rotary, OutFeet)
			
			If wpf > 0 Then
				If wdr > 0 Then
					rtwasuse = EqCost(Waste, Rotary, OutFeet) / wdr
				End If
			Else
				rtwasuse = 0
			End If
			
			EqCost(Waste, Rotary, OutHours) = rtwasuse
			
			If wpf > 0 Then EqCost(Waste, Rotary, OutNumber) = Int(((EqCost(Waste, Rotary, OutHours) / laef) / 0.92) / (hrsh * shdy)) + 1
			
			If EqCost(Waste, Rotary, OutHours) > 0 And EqCost(Waste, Rotary, OutNumber) > 0 Then
				EqCost(Waste, Rotary, OutLife) = EqDefault(Rotary, g, life) / ((EqCost(Waste, Rotary, OutHours) / EqCost(Waste, Rotary, OutNumber)) * dyyr)
			End If
			
			If EqCost(Waste, Rotary, OutHours) <> 0 Then
				If CellValues(EquipmentHours, 19, WhichSegment).Changed = False Then CellValues(EquipmentHours, 19, WhichSegment).Value = EqCost(Waste, Rotary, OutHours)
				If CellValues(RepairParts, 19, WhichSegment).Changed = False Then CellValues(RepairParts, 19, WhichSegment).Value = rtptpr
				If CellValues(Diesel, 19, WhichSegment).Changed = False Then CellValues(Diesel, 19, WhichSegment).Value = EqDefault(Rotary, g, DrillFuel)
				If CellValues(Lubricants, 19, WhichSegment).Changed = False Then CellValues(Lubricants, 19, WhichSegment).Value = rtlbpr
				If CellValues(RepairLabor, 19, WhichSegment).Changed = False Then CellValues(RepairLabor, 19, WhichSegment).Value = EqDefault(Rotary, g, DrillMechanic)
				If CellValues(Powder, 22, WhichSegment).Changed = False Then CellValues(Powder, 22, WhichSegment).Value = EqCost(Waste, Rotary, OutBit)
				If CellValues(Powder, 23, WhichSegment).Changed = False Then CellValues(Powder, 23, WhichSegment).Value = EqCost(Waste, Rotary, OutSteel)
			End If
			
			If CellValues(Purchase, 19, WhichSegment).Changed = False Then
				CellValues(Purchase, 19, WhichSegment).Value = EqDefault(Rotary, g, Price)
			End If
			
			EqCost(Waste, Rotary, OutParts) = CellValues(RepairParts, 19, WhichSegment).Value * CellValues(EquipmentHours, 19, WhichSegment).Value * TaxRate
			OpBin(4) = OpBin(4) + EqCost(Waste, Rotary, OutParts)
			EqCost(Waste, Rotary, OutFuel) = fuel * CellValues(Diesel, 19, WhichSegment).Value * CellValues(EquipmentHours, 19, WhichSegment).Value
			EqCost(Waste, Rotary, OutLube) = CellValues(Lubricants, 19, WhichSegment).Value * CellValues(EquipmentHours, 19, WhichSegment).Value * TaxRate
			OpBin(2) = OpBin(2) + EqCost(Ore, Rotary, OutLube)
			EqCost(Waste, Rotary, OutMechanicTime) = CellValues(EquipmentHours, 19, WhichSegment).Value * CellValues(RepairLabor, 19, WhichSegment).Value
			'UPGRADE_WARNING: Couldn't resolve default property of object BurdenRate. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object mcwg. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			EqCost(Waste, Rotary, OutMechanicCost) = EqCost(Waste, Rotary, OutMechanicTime) * mcwg * BurdenRate
			EqCost(Waste, Rotary, OutBit) = EqDefault(Rotary, g, BitCost)
			EqCost(Waste, Rotary, OutSteel) = EqDefault(Rotary, g, SteelCost)
			EqCost(Waste, Rotary, OutPrice) = CellValues(Purchase, 19, WhichSegment).Value
			
			OutEquipment(19, WhichSegment) = EqCost(Waste, Rotary, OutParts) + EqCost(Waste, Rotary, OutFuel) + EqCost(Waste, Rotary, OutLube)
			
			If CellValues(Replace_Renamed, 19, WhichSegment).Changed = False Then
				CellValues(Replace_Renamed, 19, WhichSegment).Value = EqCost(Waste, Rotary, OutLife) * 12
			End If
			
			If CellValues(EquipmentTwo, 31, WhichSegment).Changed = False Then
				CellValues(EquipmentTwo, 31, WhichSegment).Value = EqCost(Waste, Rotary, OutNumber)
			End If
			
			If CellValues(Purchase, 39, WhichSegment).Changed = False Then
				CellValues(Purchase, 39, WhichSegment).Value = CellValues(Purchase, 19, WhichSegment).Value * CellValues(EquipmentTwo, 31, WhichSegment).Value
			End If
			
			'UPGRADE_WARNING: Couldn't resolve default property of object MaxNumOut. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			For y = 0 To MaxNumOut
				'UPGRADE_WARNING: Couldn't resolve default property of object y. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				EqCost(Waste, Percussion, y) = 0
			Next y
			CellValues(EquipmentTwo, 30, WhichSegment).Value = 0
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object MaxNumOut. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			For y = 0 To MaxNumOut
				'UPGRADE_WARNING: Couldn't resolve default property of object y. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object x. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				EqCost(x, Rotary, y) = 0
			Next y
		End If
		
	End Sub
	Sub pkcst()
        Dim BurdenRate As Object = Nothing
        Dim mcwg As Object = Nothing
        Dim MaxNumOut As Object = Nothing
        Dim x As Object = Nothing

        Dim shdy As Decimal
		Dim dyyr As Decimal
		Dim fuel As Decimal
		Dim TaxRate As Decimal
		Dim pkhr As Decimal
		Dim pknm As Decimal
		
		Dim jump As Boolean
		
		On Error Resume Next
		
		Call getout(jump)
		If jump = True Then Exit Sub
		
		Call shcal(shdy)
		Call dycal(dyyr)
		Call fucal(fuel)
		Call txcal(TaxRate)
		
		If CellValues(EquipmentHours, 13, WhichSegment).Changed = False Then
			pkhr = 2 * shdy * CellValues(EquipmentTwo, 19, WhichSegment).Value
		Else
			pkhr = CellValues(EquipmentHours, 13, WhichSegment).Value
		End If
		
		EqCost(Ore, Pickup, OutHours) = pkhr
		
		pknm = CellValues(EquipmentTwo, 19, WhichSegment).Value
		
		If CellValues(EquipmentTwo, 19, WhichSegment).Value = 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object MaxNumOut. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			For x = 0 To MaxNumOut
				'UPGRADE_WARNING: Couldn't resolve default property of object x. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				EqCost(Ore, Pickup, x) = 0
			Next x
			If CellValues(Purchase, 13, WhichSegment).Changed = False Then CellValues(Purchase, 13, WhichSegment).Value = 0
			Exit Sub
		End If
		
		If CellValues(EquipmentHours, 13, WhichSegment).Changed = False Then CellValues(EquipmentHours, 13, WhichSegment).Value = pkhr
		If CellValues(RepairParts, 13, WhichSegment).Changed = False Then CellValues(RepairParts, 13, WhichSegment).Value = EqDefault(Pickup, 1, Parts)
		If CellValues(Diesel, 13, WhichSegment).Changed = False Then CellValues(Diesel, 13, WhichSegment).Value = EqDefault(Pickup, 1, FuelUse)
		If CellValues(Lubricants, 13, WhichSegment).Changed = False Then CellValues(Lubricants, 13, WhichSegment).Value = EqDefault(Pickup, 1, LubeCost)
		If CellValues(Tires, 13, WhichSegment).Changed = False Then CellValues(Tires, 13, WhichSegment).Value = EqDefault(Pickup, 1, TirePrice)
		If CellValues(RepairLabor, 13, WhichSegment).Changed = False Then CellValues(RepairLabor, 13, WhichSegment).Value = EqDefault(Pickup, 1, Mechanic)
		
		If CellValues(Purchase, 13, WhichSegment).Changed = False Then
			CellValues(Purchase, 13, WhichSegment).Value = EqDefault(Pickup, 1, Price)
		End If
		
		EqCost(Ore, Pickup, OutParts) = CellValues(RepairParts, 13, WhichSegment).Value * CellValues(EquipmentHours, 13, WhichSegment).Value * TaxRate
		OpBin(4) = OpBin(4) + EqCost(Ore, Pickup, OutParts)
		EqCost(Ore, Pickup, OutFuel) = fuel * CellValues(Diesel, 13, WhichSegment).Value * CellValues(EquipmentHours, 13, WhichSegment).Value
		EqCost(Ore, Pickup, OutLube) = CellValues(Lubricants, 13, WhichSegment).Value * CellValues(EquipmentHours, 13, WhichSegment).Value * TaxRate
		OpBin(2) = OpBin(2) + EqCost(Ore, Pickup, OutLube)
		EqCost(Ore, Pickup, OutTires) = CellValues(Tires, 13, WhichSegment).Value / EqDefault(Pickup, 1, TireLife) * CellValues(EquipmentHours, 13, WhichSegment).Value * TaxRate
		OpBin(5) = OpBin(5) + EqCost(Ore, Pickup, OutTires)
		EqCost(Ore, Pickup, OutMechanicTime) = CellValues(RepairLabor, 13, WhichSegment).Value * CellValues(EquipmentHours, 13, WhichSegment).Value
		'UPGRADE_WARNING: Couldn't resolve default property of object BurdenRate. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object mcwg. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		EqCost(Ore, Pickup, OutMechanicCost) = EqCost(Ore, Pickup, OutMechanicTime) * mcwg * BurdenRate
		
		OutEquipment(13, WhichSegment) = EqCost(Ore, Pickup, OutParts) + EqCost(Ore, Pickup, OutFuel) + EqCost(Ore, Pickup, OutLube) + EqCost(Ore, Pickup, OutTires)
		
		If CellValues(EquipmentHours, 13, WhichSegment).Value > 0 And pknm > 0 And dyyr > 0 Then
			EqCost(Ore, Pickup, OutLife) = EqDefault(Pickup, 1, life) / ((CellValues(EquipmentHours, 13, WhichSegment).Value / pknm) * dyyr)
		End If
		
		If CellValues(Purchase, 33, WhichSegment).Changed = False Then
			CellValues(Purchase, 33, WhichSegment).Value = CellValues(Purchase, 13, WhichSegment).Value * CellValues(EquipmentTwo, 19, WhichSegment).Value
		End If
		
		If CellValues(Replace_Renamed, 13, WhichSegment).Changed = False Then
			CellValues(Replace_Renamed, 13, WhichSegment).Value = EqCost(Ore, Pickup, OutLife) * 12
		End If
		
	End Sub
	Sub pocst()
        Dim MaxNumOut As Object = Nothing
        Dim x As Object = Nothing

        Dim l As Short
		
		Dim dyyr As Decimal
		Dim fuel As Decimal
		Dim TaxRate As Decimal
		Dim BurdenRate As Decimal
		Dim mcwg As Decimal
		Dim laef As Decimal
		Dim vol As Decimal
		Dim head As Decimal
		Dim pumphp As Decimal
		Dim punm As Decimal
		Dim puhr As Decimal
		
		Dim jump As Boolean
		
		On Error Resume Next
		
		Call getout(jump)
		If jump = True Then Exit Sub
		
		Call dycal(dyyr)
		Call fucal(fuel)
		Call txcal(TaxRate)
		Call mccal(BurdenRate, mcwg)
		Call lbcal(laef)
		Call pmcal(vol, head)
		
		l = CellValues(EquipmentOne, 18, WhichSegment).Value
		
		pumphp = (vol * head * 8.33) / (33000 * 0.6)
		
		punm = CellValues(Pumping, 3, WhichSegment).Value
		
		If CellValues(EquipmentTwo, 18, WhichSegment).Changed = False Then
			CellValues(EquipmentTwo, 18, WhichSegment).Value = punm
		End If
		
		If CellValues(EquipmentHours, 12, WhichSegment).Changed = False Then
			puhr = 24 * punm
		Else
			puhr = CellValues(EquipmentHours, 12, WhichSegment).Value
		End If
		
		EqCost(Ore, Pump, OutHours) = puhr
		
		If CellValues(EquipmentOne, 18, WhichSegment).Value = 0 Or vol = 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object MaxNumOut. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			For x = 0 To MaxNumOut
				'UPGRADE_WARNING: Couldn't resolve default property of object x. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				EqCost(Ore, Pump, x) = 0
			Next x
			If CellValues(EquipmentOne, 18, WhichSegment).Changed = False Then CellValues(EquipmentOne, 18, WhichSegment).Value = 0
			If CellValues(EquipmentTwo, 18, WhichSegment).Changed = False Then CellValues(EquipmentTwo, 18, WhichSegment).Value = 0
			If CellValues(Purchase, 12, WhichSegment).Changed = False Then CellValues(Purchase, 12, WhichSegment).Value = 0
			Exit Sub
		End If
		
		If CellValues(EquipmentHours, 12, WhichSegment).Changed = False Then CellValues(EquipmentHours, 12, WhichSegment).Value = puhr
		If CellValues(RepairParts, 12, WhichSegment).Changed = False Then CellValues(RepairParts, 12, WhichSegment).Value = EqDefault(Pump, l, Parts)
		If CellValues(Diesel, 12, WhichSegment).Changed = False Then CellValues(Diesel, 12, WhichSegment).Value = EqDefault(Pump, l, FuelUse)
		If CellValues(Lubricants, 12, WhichSegment).Changed = False Then CellValues(Lubricants, 12, WhichSegment).Value = EqDefault(Pump, l, LubeCost)
		If CellValues(RepairLabor, 12, WhichSegment).Changed = False Then CellValues(RepairLabor, 12, WhichSegment).Value = EqDefault(Pump, l, Mechanic)
		
		EqCost(Ore, Pump, OutParts) = CellValues(RepairParts, 12, WhichSegment).Value * CellValues(EquipmentHours, 12, WhichSegment).Value * TaxRate
		OpBin(4) = OpBin(4) + EqCost(Ore, Pump, OutParts)
		EqCost(Ore, Pump, OutFuel) = CellValues(Diesel, 12, WhichSegment).Value * CellValues(EquipmentHours, 12, WhichSegment).Value * fuel
		EqCost(Ore, Pump, OutLube) = CellValues(Lubricants, 12, WhichSegment).Value * CellValues(EquipmentHours, 12, WhichSegment).Value
		OpBin(2) = OpBin(2) + EqCost(Ore, Pump, OutLube)
		EqCost(Ore, Pump, OutMechanicTime) = CellValues(RepairLabor, 12, WhichSegment).Value * CellValues(EquipmentHours, 12, WhichSegment).Value
		EqCost(Ore, Pump, OutMechanicCost) = EqCost(Ore, Pump, OutMechanicTime) * mcwg * BurdenRate
		
		OutEquipment(12, WhichSegment) = EqCost(Ore, Pump, OutParts) + EqCost(Ore, Pump, OutFuel) + EqCost(Ore, Pump, OutLube)
		
		If dyyr > 0 Then
			EqCost(Ore, Pump, OutLife) = EqDefault(Pump, l, life) / ((CellValues(EquipmentHours, 12, WhichSegment).Value / punm) * dyyr)
		End If
		
		If CellValues(Purchase, 12, WhichSegment).Changed = False Then
			CellValues(Purchase, 12, WhichSegment).Value = EqDefault(Pump, l, Price)
		End If
		
		If CellValues(Replace_Renamed, 12, WhichSegment).Changed = False Then
			CellValues(Replace_Renamed, 12, WhichSegment).Value = EqCost(Ore, Pump, OutLife) * 12
		End If
		
		If CellValues(Purchase, 32, WhichSegment).Changed = False Then
			CellValues(Purchase, 32, WhichSegment).Value = CellValues(Purchase, 12, WhichSegment).Value * CellValues(EquipmentTwo, 18, WhichSegment).Value
		End If
		
	End Sub
	Sub dzcst()
        Dim MaxNumOut As Object = Nothing
        Dim x As Object = Nothing

        Dim h As Short
		
		Dim hrsh As Decimal
		Dim shdy As Decimal
		Dim dyyr As Decimal
		Dim TaxRate As Decimal
		Dim BurdenRate As Decimal
		Dim mcwg As Decimal
		Dim laef As Decimal
		Dim fuel As Decimal
		Dim dznm As Decimal
		Dim dzhr As Decimal
		Dim hvwg As Decimal
		
		Dim jump As Boolean
		
		On Error Resume Next
		
		Call getout(jump)
		If jump = True Then Exit Sub
		
		Call hrcal(hrsh)
		Call shcal(shdy)
		Call dycal(dyyr)
		Call txcal(TaxRate)
		Call mccal(BurdenRate, mcwg)
		Call lbcal(laef)
		Call fucal(fuel)
		
		hvwg = CellValues(Wage, 4, WhichSegment).Value
		
		h = CellValues(EquipmentOne, 12, WhichSegment).Value
		
		If CellValues(EquipmentOne, 12, WhichSegment).Value = 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object MaxNumOut. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			For x = 0 To MaxNumOut
				'UPGRADE_WARNING: Couldn't resolve default property of object x. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				EqCost(Ore, Dozer, x) = 0
			Next x
			CellValues(Purchase, 6, WhichSegment).Value = 0
			CellValues(EquipmentTwo, 12, WhichSegment).Value = 0
			Exit Sub
		End If
		
		dznm = CellValues(EquipmentTwo, 12, WhichSegment).Value
		
		If CellValues(EquipmentHours, 6, WhichSegment).Changed = False Then
			dzhr = dznm * hrsh * shdy * laef
		Else
			dzhr = CellValues(EquipmentHours, 6, WhichSegment).Value
		End If
		
		If CellValues(EquipmentHours, 6, WhichSegment).Changed = False Then CellValues(EquipmentHours, 6, WhichSegment).Value = dzhr
		If CellValues(RepairParts, 6, WhichSegment).Changed = False Then CellValues(RepairParts, 6, WhichSegment).Value = EqDefault(Dozer, h, Parts)
		If CellValues(Undercarriage, 6, WhichSegment).Changed = False Then CellValues(Undercarriage, 6, WhichSegment).Value = EqDefault(Dozer, h, uCarriage)
		If CellValues(Diesel, 6, WhichSegment).Changed = False Then CellValues(Diesel, 6, WhichSegment).Value = EqDefault(Dozer, h, FuelUse)
		If CellValues(Lubricants, 6, WhichSegment).Changed = False Then CellValues(Lubricants, 6, WhichSegment).Value = EqDefault(Dozer, h, LubeCost)
		If CellValues(RepairLabor, 6, WhichSegment).Changed = False Then CellValues(RepairLabor, 6, WhichSegment).Value = EqDefault(Dozer, h, Mechanic)
		
		EqCost(Ore, Dozer, OutParts) = (CellValues(RepairParts, 6, WhichSegment).Value + CellValues(Undercarriage, 6, WhichSegment).Value) * CellValues(EquipmentHours, 6, WhichSegment).Value * TaxRate
		OpBin(4) = OpBin(4) + EqCost(Ore, Dozer, OutParts)
		EqCost(Ore, Dozer, OutFuel) = CellValues(Diesel, 6, WhichSegment).Value * CellValues(EquipmentHours, 6, WhichSegment).Value * fuel
		EqCost(Ore, Dozer, OutLube) = CellValues(Lubricants, 6, WhichSegment).Value * CellValues(EquipmentHours, 6, WhichSegment).Value * TaxRate
		OpBin(2) = OpBin(2) + EqCost(Ore, Dozer, OutLube)
		EqCost(Ore, Dozer, OutOperator) = (Int((CellValues(EquipmentHours, 6, WhichSegment).Value / laef) / hrsh) + 1) * hvwg * hrsh * BurdenRate
		EqCost(Ore, Dozer, OutMechanicTime) = CellValues(RepairLabor, 6, WhichSegment).Value * CellValues(EquipmentHours, 6, WhichSegment).Value
		EqCost(Ore, Dozer, OutMechanicCost) = EqCost(Ore, Dozer, OutMechanicTime) * mcwg * BurdenRate
		
		OutEquipment(6, WhichSegment) = EqCost(Ore, Dozer, OutParts) + EqCost(Ore, Dozer, OutFuel) + EqCost(Ore, Dozer, OutLube)
		
		EqCost(Ore, Dozer, OutUnit) = OutEquipment(6, WhichSegment) / CellValues(EquipmentHours, 6, WhichSegment).Value
		
		If CellValues(Purchase, 6, WhichSegment).Changed = False Then
			CellValues(Purchase, 6, WhichSegment).Value = EqDefault(Dozer, h, Price)
		End If
		
		If dzhr > 0 And dznm > 0 And dyyr > 0 Then
			EqCost(Ore, Dozer, OutLife) = EqDefault(Dozer, h, life) / ((dzhr / dznm) * dyyr)
		End If
		
		If CellValues(Replace_Renamed, 6, WhichSegment).Changed = False Then
			CellValues(Replace_Renamed, 6, WhichSegment).Value = EqCost(Ore, Dozer, OutLife) * 12
		End If
		
		If CellValues(Purchase, 26, WhichSegment).Changed = False Then
			CellValues(Purchase, 26, WhichSegment).Value = CellValues(Purchase, 6, WhichSegment).Value * CellValues(EquipmentTwo, 12, WhichSegment).Value
		End If
		
	End Sub
	Sub grcst()
        Dim MaxNumOut As Object = Nothing

        Dim x As Short
		Dim co As Short
		Dim cw As Short
		Dim m As Short
		
		Dim hrsh As Decimal
		Dim shdy As Decimal
		Dim dyyr As Decimal
		Dim TaxRate As Decimal
		Dim BurdenRate As Decimal
		Dim mcwg As Decimal
		Dim laef As Decimal
		Dim fuel As Decimal
		Dim grhr As Decimal
		Dim grnm As Decimal
		Dim ohd As Decimal
		Dim whd As Decimal
		Dim utwg As Decimal
		Dim TempDist(2) As Decimal
		
		Dim jump As Boolean
		
		On Error Resume Next
		
		Call getout(jump)
		If jump = True Then Exit Sub
		
		co = Int(CellValues(EquipmentOne, 4, WhichSegment).Value)
		cw = Int(CellValues(EquipmentOne, 9, WhichSegment).Value)
		
		Call hrcal(hrsh)
		Call shcal(shdy)
		Call dycal(dyyr)
		Call txcal(TaxRate)
		Call mccal(BurdenRate, mcwg)
		Call lbcal(laef)
		Call fucal(fuel)
		
		For x = 0 To 6 Step 2
			TempDist(1) = TempDist(1) + CellValues(Haul, x + 8, WhichSegment).Value
			TempDist(2) = TempDist(2) + CellValues(Haul, x + 20, WhichSegment).Value
		Next x
		
		ohd = CellValues(Haul, 0, WhichSegment).Value + CellValues(Haul, 2, WhichSegment).Value + TempDist(1)
		whd = CellValues(Haul, 4, WhichSegment).Value + CellValues(Haul, 6, WhichSegment).Value + TempDist(2)
		utwg = CellValues(Wage, 5, WhichSegment).Value
		
		m = CellValues(EquipmentOne, 13, WhichSegment).Value
		
		If CellValues(EquipmentHours, 7, WhichSegment).Changed = False Then
			grhr = (((ohd * EqDefault(Truck, co, TruckWidth) * 4) + (whd * EqDefault(Truck, cw, TruckWidth) * 4)) * 2 * shdy) / EqDefault(Grader, m, Productivity)
		Else
			grhr = CellValues(EquipmentHours, 7, WhichSegment).Value
		End If
		
		grnm = CellValues(EquipmentTwo, 13, WhichSegment).Value
		
		If CellValues(EquipmentOne, 13, WhichSegment).Value = 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object MaxNumOut. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			For x = 0 To MaxNumOut
				EqCost(Ore, Grader, x) = 0
			Next x
			If CellValues(EquipmentTwo, 13, WhichSegment).Changed = False Then CellValues(EquipmentTwo, 13, WhichSegment).Value = 0
			If CellValues(Purchase, 7, WhichSegment).Changed = False Then CellValues(Purchase, 7, WhichSegment).Value = 0
			Call later()
		End If
		
		If CellValues(EquipmentHours, 7, WhichSegment).Changed = False Then CellValues(EquipmentHours, 7, WhichSegment).Value = grhr
		If CellValues(RepairParts, 7, WhichSegment).Changed = False Then CellValues(RepairParts, 7, WhichSegment).Value = EqDefault(Grader, m, Parts)
		If CellValues(Diesel, 7, WhichSegment).Changed = False Then CellValues(Diesel, 7, WhichSegment).Value = EqDefault(Grader, m, FuelUse)
		If CellValues(Lubricants, 7, WhichSegment).Changed = False Then CellValues(Lubricants, 7, WhichSegment).Value = EqDefault(Grader, m, LubeCost)
		If CellValues(Tires, 7, WhichSegment).Changed = False Then CellValues(Tires, 7, WhichSegment).Value = EqDefault(Grader, m, TirePrice)
		If CellValues(RepairLabor, 7, WhichSegment).Changed = False Then CellValues(RepairLabor, 7, WhichSegment).Value = EqDefault(Grader, m, Mechanic)
		
		If CellValues(Purchase, 7, WhichSegment).Changed = False Then
			CellValues(Purchase, 7, WhichSegment).Value = EqDefault(Grader, m, Price)
		End If
		
		EqCost(Ore, Grader, OutParts) = CellValues(RepairParts, 7, WhichSegment).Value * CellValues(EquipmentHours, 7, WhichSegment).Value * TaxRate
		OpBin(4) = OpBin(4) + EqCost(Ore, Grader, OutParts)
		EqCost(Ore, Grader, OutFuel) = fuel * CellValues(Diesel, 7, WhichSegment).Value * CellValues(EquipmentHours, 7, WhichSegment).Value
		EqCost(Ore, Grader, OutLube) = CellValues(Lubricants, 7, WhichSegment).Value * CellValues(EquipmentHours, 7, WhichSegment).Value * TaxRate
		OpBin(2) = OpBin(2) + EqCost(Ore, Grader, OutLube)
		EqCost(Ore, Grader, OutTires) = CellValues(Tires, 7, WhichSegment).Value / EqDefault(Grader, m, TireLife) * CellValues(EquipmentHours, 7, WhichSegment).Value * TaxRate
		OpBin(5) = OpBin(5) + EqCost(Ore, Grader, OutTires)
		EqCost(Ore, Grader, OutOperator) = (Int((CellValues(EquipmentHours, 7, WhichSegment).Value / laef) / hrsh) + 1) * utwg * hrsh * BurdenRate
		EqCost(Ore, Grader, OutMechanicTime) = CellValues(RepairLabor, 7, WhichSegment).Value * CellValues(EquipmentHours, 7, WhichSegment).Value
		EqCost(Ore, Grader, OutMechanicCost) = EqCost(Ore, Grader, OutMechanicTime) * mcwg * BurdenRate
		
		OutEquipment(7, WhichSegment) = EqCost(Ore, Grader, OutParts) + EqCost(Ore, Grader, OutFuel) + EqCost(Ore, Grader, OutLube) + EqCost(Ore, Grader, OutTires)
		
		If CellValues(EquipmentHours, 7, WhichSegment).Value > 0 And grnm > 0 And dyyr > 0 Then
			EqCost(Ore, Grader, OutLife) = EqDefault(Grader, m, life) / ((CellValues(EquipmentHours, 7, WhichSegment).Value / grnm) * dyyr)
		End If
		
		If CellValues(Replace_Renamed, 7, WhichSegment).Changed = False Then
			CellValues(Replace_Renamed, 7, WhichSegment).Value = EqCost(Ore, Grader, OutLife) * 12
		End If
		
		If CellValues(Purchase, 27, WhichSegment).Changed = False Then
			CellValues(Purchase, 27, WhichSegment).Value = CellValues(Purchase, 7, WhichSegment).Value * CellValues(EquipmentTwo, 13, WhichSegment).Value
		End If
		
	End Sub
	Sub later()
		
		On Error Resume Next
		
		If CellValues(Purchase, 0, WhichSegment).Changed = False Then
			CellValues(Purchase, 0, WhichSegment).Value = EqCost(Ore, Loader, OutPrice) + EqCost(Ore, Shovel, OutPrice) + EqCost(Ore, CableShovel, OutPrice) + EqCost(Ore, Dragline, OutPrice) + EqCost(Ore, Scraper, OutPrice)
		End If
		
		If CellValues(Purchase, 20, WhichSegment).Changed = False Then
			'CellValues(Purchase, 20, WhichSegment).Value = CellValues(Purchase, 0, WhichSegment).Value * (EqCost(Ore, Loader, OutNumber) + EqCost(Ore, Shovel, OutNumber) + EqCost(Ore, CableShovel, OutNumber) + EqCost(Ore, Dragline, OutNumber) + EqCost(Ore, Scraper, OutNumber))
			CellValues(Purchase, 20, WhichSegment).Value = CellValues(Purchase, 0, WhichSegment).Value * (CellValues(EquipmentTwo, 0, WhichSegment).Value + CellValues(EquipmentTwo, 1, WhichSegment).Value + CellValues(EquipmentTwo, 2, WhichSegment).Value + CellValues(EquipmentTwo, 3, WhichSegment).Value + CellValues(EquipmentTwo, 20, WhichSegment).Value)
		End If
		
		If CellValues(Purchase, 2, WhichSegment).Changed = False Then
			CellValues(Purchase, 2, WhichSegment).Value = EqCost(Waste, Loader, OutPrice) + EqCost(Waste, Shovel, OutPrice) + EqCost(Waste, CableShovel, OutPrice) + EqCost(Waste, Dragline, OutPrice) + EqCost(Waste, Scraper, OutPrice)
		End If
		
		If CellValues(Purchase, 22, WhichSegment).Changed = False Then
			'CellValues(Purchase, 22, WhichSegment).Value = CellValues(Purchase, 2, WhichSegment).Value * (EqCost(Waste, Loader, OutNumber) + EqCost(Waste, Shovel, OutNumber) + EqCost(Waste, CableShovel, OutNumber) + EqCost(Waste, Dragline, OutNumber) + EqCost(Waste, Scraper, OutNumber))
			CellValues(Purchase, 22, WhichSegment).Value = CellValues(Purchase, 2, WhichSegment).Value * (CellValues(EquipmentTwo, 5, WhichSegment).Value + CellValues(EquipmentTwo, 6, WhichSegment).Value + CellValues(EquipmentTwo, 7, WhichSegment).Value + CellValues(EquipmentTwo, 8, WhichSegment).Value + CellValues(EquipmentTwo, 25, WhichSegment).Value)
		End If
		
	End Sub
	Sub ltcst()
        Dim MaxNumOut As Object = Nothing
        Dim x As Object = Nothing
        Dim i As Object = Nothing

        Dim l As Short
		
		Dim dyyr As Decimal
		Dim TaxRate As Decimal
		Dim BurdenRate As Decimal
		Dim mcwg As Decimal
		Dim fuel As Decimal
		Dim ltnm As Decimal
		Dim lthrsingle As Decimal
		Dim lthr As Decimal
		
		Dim jump As Boolean
		
		On Error Resume Next
		
		Call getout(jump)
		If jump = True Then Exit Sub
		
		'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		i = CellValues(EquipmentOne, 17, WhichSegment).Value
		
		ltnm = CellValues(EquipmentTwo, 17, WhichSegment).Value
		
		Call dycal(dyyr)
		Call txcal(TaxRate)
		Call mccal(BurdenRate, mcwg)
		Call fucal(fuel)
		
		lthrsingle = (CellValues(Production, 1, WhichSegment).Value * CellValues(Production, 2, WhichSegment).Value) - 12
		
		If CellValues(EquipmentHours, 11, WhichSegment).Changed = True Then
			lthr = CellValues(EquipmentHours, 11, WhichSegment).Value
		Else
			lthr = lthrsingle * ltnm
		End If
		
		If CellValues(EquipmentOne, 17, WhichSegment).Value = 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object MaxNumOut. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			For x = 0 To MaxNumOut
				'UPGRADE_WARNING: Couldn't resolve default property of object x. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				EqCost(Ore, Light, x) = 0
			Next x
			If CellValues(EquipmentTwo, 17, WhichSegment).Changed = False Then CellValues(EquipmentTwo, 17, WhichSegment).Value = 0
			If CellValues(Purchase, 11, WhichSegment).Changed = False Then CellValues(Purchase, 11, WhichSegment).Value = 0
			Exit Sub
		End If
		
		If CellValues(EquipmentHours, 11, WhichSegment).Changed = False Then CellValues(EquipmentHours, 11, WhichSegment).Value = lthr
		If CellValues(RepairParts, 11, WhichSegment).Changed = False Then CellValues(RepairParts, 11, WhichSegment).Value = EqDefault(Light, i, Parts)
		If CellValues(Diesel, 11, WhichSegment).Changed = False Then CellValues(Diesel, 11, WhichSegment).Value = EqDefault(Light, i, FuelUse)
		If CellValues(Lubricants, 11, WhichSegment).Changed = False Then CellValues(Lubricants, 11, WhichSegment).Value = EqDefault(Light, i, LubeCost)
		If CellValues(RepairLabor, 11, WhichSegment).Changed = False Then CellValues(RepairLabor, 11, WhichSegment).Value = EqDefault(Light, i, Mechanic)
		
		If CellValues(Purchase, 11, WhichSegment).Changed = False Then
			CellValues(Purchase, 11, WhichSegment).Value = EqDefault(Light, i, Price)
		End If
		
		EqCost(Ore, Light, OutParts) = CellValues(RepairParts, 11, WhichSegment).Value * CellValues(EquipmentHours, 11, WhichSegment).Value * TaxRate
		OpBin(4) = OpBin(4) + EqCost(Ore, Light, OutParts)
		EqCost(Ore, Light, OutFuel) = CellValues(Diesel, 11, WhichSegment).Value * CellValues(EquipmentHours, 11, WhichSegment).Value * fuel
		EqCost(Ore, Light, OutLube) = CellValues(Lubricants, 11, WhichSegment).Value * CellValues(EquipmentHours, 11, WhichSegment).Value
		OpBin(2) = OpBin(2) + EqCost(Ore, Light, OutLube)
		EqCost(Ore, Light, OutMechanicTime) = CellValues(RepairLabor, 11, WhichSegment).Value * CellValues(EquipmentHours, 11, WhichSegment).Value
		EqCost(Ore, Light, OutMechanicCost) = EqCost(Ore, Light, OutMechanicTime) * mcwg * BurdenRate
		
		OutEquipment(11, WhichSegment) = EqCost(Ore, Light, OutParts) + EqCost(Ore, Light, OutFuel) + EqCost(Ore, Light, OutLube)
		
		If ltnm > 0 And CellValues(EquipmentHours, 11, WhichSegment).Value > 0 And dyyr > 0 Then
			EqCost(Ore, Light, OutLife) = EqDefault(Light, i, life) / ((CellValues(EquipmentHours, 11, WhichSegment).Value / ltnm) * dyyr)
		End If
		
		If CellValues(Replace_Renamed, 11, WhichSegment).Changed = False Then
			CellValues(Replace_Renamed, 11, WhichSegment).Value = EqCost(Ore, Light, OutLife) * 12
		End If
		
		If CellValues(Purchase, 31, WhichSegment).Changed = False Then
			CellValues(Purchase, 31, WhichSegment).Value = CellValues(Purchase, 11, WhichSegment).Value * CellValues(EquipmentTwo, 17, WhichSegment).Value
		End If
		
	End Sub
	Sub wtcst()
        Dim MaxNumOut As Object = Nothing
        Dim x As Object = Nothing

        Dim k As Short
		Dim Test As Short
		Dim wthrsh As Short
		
		Dim hrsh As Decimal
		Dim shdy As Decimal
		Dim dyyr As Decimal
		Dim TaxRate As Decimal
		Dim BurdenRate As Decimal
		Dim mcwg As Decimal
		Dim laef As Decimal
		Dim fuel As Decimal
		Dim utwg As Decimal
		Dim wthr As Decimal
		Dim wtnm As Decimal
		Dim TonsPerDay As Decimal
		
		Dim jump As Boolean
		
		On Error Resume Next
		
		Call getout(jump)
		If jump = True Then Exit Sub
		
		Call tonsp(TonsPerDay)
		
		Test = Int(TonsPerDay / 5000)
		
		Select Case Test
			Case 0, 3, 6, 9, 12
				wthrsh = 2
			Case 1, 4, 7, 10, 13
				wthrsh = 4
			Case 2, 5, 8, 11
				wthrsh = 6
			Case Is > 13
				wthrsh = 6
		End Select
		
		k = CellValues(EquipmentOne, 14, WhichSegment).Value
		
		If k = 0 Then
			If CellValues(EquipmentTwo, 14, WhichSegment).Changed = False Then CellValues(EquipmentTwo, 14, WhichSegment).Value = 0
			'UPGRADE_WARNING: Couldn't resolve default property of object MaxNumOut. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			For x = 0 To MaxNumOut
				'UPGRADE_WARNING: Couldn't resolve default property of object x. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				EqCost(Ore, WaterTanker, x) = 0
			Next x
			If CellValues(Purchase, 8, WhichSegment).Changed = False Then CellValues(Purchase, 8, WhichSegment).Value = 0
			Exit Sub
		Else
			If CellValues(EquipmentTwo, 14, WhichSegment).Changed = False Then CellValues(EquipmentTwo, 14, WhichSegment).Value = 1
		End If
		
		Call hrcal(hrsh)
		Call shcal(shdy)
		Call dycal(dyyr)
		Call txcal(TaxRate)
		Call mccal(BurdenRate, mcwg)
		Call lbcal(laef)
		Call fucal(fuel)
		
		utwg = CellValues(Wage, 5, WhichSegment).Value
		
		EqCost(Ore, WaterTanker, OutHours) = wthrsh * shdy
		
		If CellValues(EquipmentHours, 8, WhichSegment).Changed = False Then
			wthr = EqCost(Ore, WaterTanker, OutHours)
		Else
			wthr = CellValues(EquipmentHours, 8, WhichSegment).Value
		End If
		
		wtnm = CellValues(EquipmentTwo, 14, WhichSegment).Value
		
		If CellValues(EquipmentHours, 8, WhichSegment).Changed = False Then CellValues(EquipmentHours, 8, WhichSegment).Value = wthr
		If CellValues(RepairParts, 8, WhichSegment).Changed = False Then CellValues(RepairParts, 8, WhichSegment).Value = EqDefault(WaterTanker, k, Parts)
		If CellValues(Diesel, 8, WhichSegment).Changed = False Then CellValues(Diesel, 8, WhichSegment).Value = EqDefault(WaterTanker, k, FuelUse)
		If CellValues(Lubricants, 8, WhichSegment).Changed = False Then CellValues(Lubricants, 8, WhichSegment).Value = EqDefault(WaterTanker, k, LubeCost)
		If CellValues(Tires, 8, WhichSegment).Changed = False Then CellValues(Tires, 8, WhichSegment).Value = EqDefault(WaterTanker, k, TirePrice)
		If CellValues(RepairLabor, 8, WhichSegment).Changed = False Then CellValues(RepairLabor, 8, WhichSegment).Value = EqDefault(WaterTanker, k, Mechanic)
		
		If CellValues(Purchase, 8, WhichSegment).Changed = False Then
			CellValues(Purchase, 8, WhichSegment).Value = EqDefault(WaterTanker, k, Price)
		End If
		
		EqCost(Ore, WaterTanker, OutParts) = CellValues(RepairParts, 8, WhichSegment).Value * CellValues(EquipmentHours, 8, WhichSegment).Value * TaxRate
		OpBin(4) = OpBin(4) + EqCost(Ore, WaterTanker, OutParts)
		EqCost(Ore, WaterTanker, OutFuel) = fuel * CellValues(Diesel, 8, WhichSegment).Value * CellValues(EquipmentHours, 8, WhichSegment).Value
		EqCost(Ore, WaterTanker, OutLube) = CellValues(Lubricants, 8, WhichSegment).Value * CellValues(EquipmentHours, 8, WhichSegment).Value * TaxRate
		OpBin(2) = OpBin(2) + EqCost(Ore, WaterTanker, OutLube)
		EqCost(Ore, WaterTanker, OutTires) = (CellValues(Tires, 8, WhichSegment).Value / EqDefault(WaterTanker, k, TireLife)) * CellValues(EquipmentHours, 8, WhichSegment).Value * TaxRate
		OpBin(5) = OpBin(5) + EqCost(Ore, WaterTanker, OutTires)
		EqCost(Ore, WaterTanker, OutOperator) = (Int((CellValues(EquipmentHours, 8, WhichSegment).Value / laef) / hrsh) + 1) * utwg * hrsh * BurdenRate
		EqCost(Ore, WaterTanker, OutMechanicTime) = CellValues(RepairLabor, 8, WhichSegment).Value * CellValues(EquipmentHours, 8, WhichSegment).Value
		EqCost(Ore, WaterTanker, OutMechanicCost) = EqCost(Ore, WaterTanker, OutMechanicTime) * mcwg * BurdenRate
		
		OutEquipment(8, WhichSegment) = EqCost(Ore, WaterTanker, OutParts) + EqCost(Ore, WaterTanker, OutFuel) + EqCost(Ore, WaterTanker, OutLube) + EqCost(Ore, WaterTanker, OutTires)
		
		If CellValues(EquipmentHours, 8, WhichSegment).Value > 0 And wtnm > 0 And dyyr > 0 Then
			EqCost(Ore, WaterTanker, OutLife) = EqDefault(WaterTanker, k, life) / ((CellValues(EquipmentHours, 8, WhichSegment).Value / wtnm) * dyyr)
		End If
		
		If CellValues(Replace_Renamed, 8, WhichSegment).Changed = False Then
			CellValues(Replace_Renamed, 8, WhichSegment).Value = EqCost(Ore, WaterTanker, OutLife) * 12
		End If
		
		If CellValues(Purchase, 28, WhichSegment).Changed = False Then
			CellValues(Purchase, 28, WhichSegment).Value = CellValues(Purchase, 8, WhichSegment).Value * CellValues(EquipmentTwo, 14, WhichSegment).Value
		End If
		
	End Sub
	Sub mtcst()
        Dim MaxNumOut As Object = Nothing
        Dim x As Object = Nothing

        Dim n As Short
		
		Dim hrsh As Decimal
		Dim shdy As Decimal
		Dim dyyr As Decimal
		Dim TaxRate As Decimal
		Dim BurdenRate As Decimal
		Dim mcwg As Decimal
		Dim laef As Decimal
		Dim fuel As Decimal
		Dim mthr As Decimal
		Dim mtnm As Decimal
		
		Dim jump As Boolean
		
		On Error Resume Next
		
		Call getout(jump)
		If jump = True Then Exit Sub
		
		Call hrcal(hrsh)
		Call shcal(shdy)
		Call dycal(dyyr)
		Call txcal(TaxRate)
		Call mccal(BurdenRate, mcwg)
		Call lbcal(laef)
		Call fucal(fuel)
		
		n = CellValues(EquipmentOne, 15, WhichSegment).Value
		
		If CellValues(EquipmentOne, 15, WhichSegment).Value = 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object MaxNumOut. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			For x = 0 To MaxNumOut
				'UPGRADE_WARNING: Couldn't resolve default property of object x. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				EqCost(Ore, MainTruck, x) = 0
			Next x
			If CellValues(EquipmentTwo, 15, WhichSegment).Changed = False Then CellValues(EquipmentTwo, 15, WhichSegment).Value = 0
			If CellValues(Purchase, 9, WhichSegment).Changed = False Then CellValues(Purchase, 9, WhichSegment).Value = 0
			Exit Sub
		End If
		
		mthr = (EqCost(Waste, Loader, OutHours) + EqCost(Waste, Shovel, OutHours) + EqCost(Waste, CableShovel, OutHours) + EqCost(Waste, Dragline, OutHours) + EqCost(Waste, Truck, OutHours))
		mthr = mthr + (EqCost(Ore, Loader, OutHours) + EqCost(Ore, Shovel, OutHours) + EqCost(Ore, CableShovel, OutHours) + EqCost(Ore, Dragline, OutHours) + EqCost(Ore, Truck, OutHours))
		
		If CellValues(EquipmentHours, 9, WhichSegment).Changed = False Then
			mthr = mthr * 0.1875
		Else
			mthr = CellValues(EquipmentHours, 9, WhichSegment).Value
		End If
		
		mtnm = CellValues(EquipmentTwo, 15, WhichSegment).Value
		
		If CellValues(EquipmentHours, 9, WhichSegment).Changed = False Then CellValues(EquipmentHours, 9, WhichSegment).Value = mthr
		If CellValues(RepairParts, 9, WhichSegment).Changed = False Then CellValues(RepairParts, 9, WhichSegment).Value = EqDefault(MainTruck, n, Parts)
		If CellValues(Diesel, 9, WhichSegment).Changed = False Then CellValues(Diesel, 9, WhichSegment).Value = EqDefault(MainTruck, n, FuelUse)
		If CellValues(Lubricants, 9, WhichSegment).Changed = False Then CellValues(Lubricants, 9, WhichSegment).Value = EqDefault(MainTruck, n, LubeCost)
		If CellValues(Tires, 9, WhichSegment).Changed = False Then CellValues(Tires, 9, WhichSegment).Value = EqDefault(MainTruck, n, TirePrice)
		If CellValues(RepairLabor, 9, WhichSegment).Changed = False Then CellValues(RepairLabor, 9, WhichSegment).Value = EqDefault(MainTruck, n, Mechanic)
		
		If CellValues(Purchase, 9, WhichSegment).Changed = False Then
			CellValues(Purchase, 9, WhichSegment).Value = EqDefault(MainTruck, n, Price)
		End If
		
		EqCost(Ore, MainTruck, OutParts) = CellValues(RepairParts, 9, WhichSegment).Value * CellValues(EquipmentHours, 9, WhichSegment).Value * TaxRate
		OpBin(4) = OpBin(4) + EqCost(Ore, MainTruck, OutParts)
		EqCost(Ore, MainTruck, OutFuel) = fuel * CellValues(Diesel, 9, WhichSegment).Value * CellValues(EquipmentHours, 9, WhichSegment).Value
		EqCost(Ore, MainTruck, OutLube) = CellValues(Lubricants, 9, WhichSegment).Value * CellValues(EquipmentHours, 9, WhichSegment).Value * TaxRate
		OpBin(2) = OpBin(2) + EqCost(Ore, MainTruck, OutLube)
		EqCost(Ore, MainTruck, OutTires) = (CellValues(Tires, 9, WhichSegment).Value / EqDefault(MainTruck, n, TireLife)) * CellValues(EquipmentHours, 9, WhichSegment).Value * TaxRate
		OpBin(5) = OpBin(5) + EqCost(Ore, MainTruck, OutTires)
		EqCost(Ore, MainTruck, OutMechanicTime) = CellValues(RepairLabor, 9, WhichSegment).Value * CellValues(EquipmentHours, 9, WhichSegment).Value
		EqCost(Ore, MainTruck, OutMechanicCost) = EqCost(Ore, MainTruck, OutMechanicTime) * mcwg * BurdenRate
		
		OutEquipment(9, WhichSegment) = EqCost(Ore, MainTruck, OutParts) + EqCost(Ore, MainTruck, OutFuel) + EqCost(Ore, MainTruck, OutLube) + EqCost(Ore, MainTruck, OutTires)
		
		If CellValues(EquipmentHours, 9, WhichSegment).Value <> 0 Then
			EqCost(Ore, MainTruck, OutUnit) = OutEquipment(9, WhichSegment) / CellValues(EquipmentHours, 9, WhichSegment).Value
		End If
		
		If CellValues(EquipmentHours, 9, WhichSegment).Value > 0 And CellValues(EquipmentTwo, 15, WhichSegment).Value > 0 And dyyr > 0 Then
			EqCost(Ore, MainTruck, OutLife) = EqDefault(MainTruck, n, life) / ((CellValues(EquipmentHours, 9, WhichSegment).Value / CellValues(EquipmentTwo, 15, WhichSegment).Value) * dyyr)
		End If
		
		If CellValues(Replace_Renamed, 9, WhichSegment).Changed = False Then
			CellValues(Replace_Renamed, 9, WhichSegment).Value = EqCost(Ore, MainTruck, OutLife) * 12
		End If
		
		If CellValues(Purchase, 29, WhichSegment).Changed = False Then
			CellValues(Purchase, 29, WhichSegment).Value = CellValues(Purchase, 9, WhichSegment).Value * CellValues(EquipmentTwo, 15, WhichSegment).Value
		End If
		
	End Sub
	Sub pbcst()
        Dim MaxNumOut As Object = Nothing
        Dim x As Object = Nothing

        Dim j As Short
		
		Dim hrsh As Decimal
		Dim shdy As Decimal
		Dim dyyr As Decimal
		Dim TaxRate As Decimal
		Dim BurdenRate As Decimal
		Dim mcwg As Decimal
		Dim laef As Decimal
		Dim fuel As Decimal
		Dim ot As Decimal
		Dim wt As Decimal
		Dim opf As Decimal
		Dim wpf As Decimal
		Dim blhrdy As Decimal
		Dim blnm As Decimal
		Dim pbhr As Decimal
		Dim pbnm As Decimal
		
		Dim jump As Boolean
		
		On Error Resume Next
		
		Call getout(jump)
		
		If jump = True Then Exit Sub
		
		Call hrcal(hrsh)
		Call shcal(shdy)
		Call dycal(dyyr)
		Call txcal(TaxRate)
		Call mccal(BurdenRate, mcwg)
		Call lbcal(laef)
		Call fucal(fuel)
		Call otcal(ot, wt)
		Call pfcal(opf, wpf)
		
		j = CellValues(EquipmentOne, 16, WhichSegment).Value
		
		If j = 1 Then
			blhrdy = ((opf * ot) + (wpf * wt)) / (151.9 * laef * 0.75)
		ElseIf j = 2 Then 
			blhrdy = ((opf * ot) + (wpf * wt)) / (253.1 * laef * 0.75)
		ElseIf j = 3 Then 
			blhrdy = ((opf * ot) + (wpf * wt)) / (3375 * laef * 0.75)
		ElseIf j = 4 Then 
			blhrdy = ((opf * ot) + (wpf * wt)) / (5625 * laef * 0.75)
		End If
		
		If hrsh > 0 Then
			If j = 1 Or j = 2 Then
				blnm = Int(blhrdy / hrsh) + 1
			ElseIf j = 3 Or j = 4 Then 
				blnm = (Int(blhrdy / hrsh) + 1) * 2
			End If
		End If
		
		If opf + wpf > 0 Then
			blhr = blnm * hrsh
		Else
			blhr = 0
		End If
		
		If bltp = 1 Then
			pbhr = blhrdy / 3
		ElseIf bltp = 2 Then 
			pbhr = blhrdy
		End If
		
		EqCost(Ore, PowderBuggy, OutHours) = pbhr
		
		If opf + wpf > 0 Then
			pbnm = (Int((pbhr / laef) / (hrsh * shdy))) + 1
		Else
			pbnm = 0
		End If
		
		If CellValues(EquipmentTwo, 16, WhichSegment).Changed = False Then
			CellValues(EquipmentTwo, 16, WhichSegment).Value = pbnm
		End If
		
		If j = 0 Then
			If CellValues(EquipmentTwo, 16, WhichSegment).Changed = False Then CellValues(EquipmentTwo, 16, WhichSegment).Value = 0
			'UPGRADE_WARNING: Couldn't resolve default property of object MaxNumOut. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			For x = 0 To MaxNumOut
				'UPGRADE_WARNING: Couldn't resolve default property of object x. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				EqCost(Ore, PowderBuggy, x) = 0
			Next x
			If CellValues(Purchase, 10, WhichSegment).Changed = False Then CellValues(Purchase, 10, WhichSegment).Value = 0
			Exit Sub
		End If
		
		If CellValues(EquipmentHours, 10, WhichSegment).Changed = False Then CellValues(EquipmentHours, 10, WhichSegment).Value = pbhr
		If CellValues(RepairParts, 10, WhichSegment).Changed = False Then CellValues(RepairParts, 10, WhichSegment).Value = EqDefault(PowderBuggy, j, Parts)
		If CellValues(Diesel, 10, WhichSegment).Changed = False Then CellValues(Diesel, 10, WhichSegment).Value = EqDefault(PowderBuggy, j, FuelUse)
		If CellValues(Lubricants, 10, WhichSegment).Changed = False Then CellValues(Lubricants, 10, WhichSegment).Value = EqDefault(PowderBuggy, j, LubeCost)
		If CellValues(Tires, 10, WhichSegment).Changed = False Then CellValues(Tires, 10, WhichSegment).Value = EqDefault(PowderBuggy, j, TirePrice)
		If CellValues(RepairLabor, 10, WhichSegment).Changed = False Then CellValues(RepairLabor, 10, WhichSegment).Value = EqDefault(PowderBuggy, j, Mechanic)
		
		If CellValues(Purchase, 10, WhichSegment).Changed = False Then
			CellValues(Purchase, 10, WhichSegment).Value = EqDefault(PowderBuggy, j, Price)
		End If
		
		EqCost(Ore, PowderBuggy, OutParts) = CellValues(RepairParts, 10, WhichSegment).Value * CellValues(EquipmentHours, 10, WhichSegment).Value * TaxRate
		OpBin(4) = OpBin(4) + EqCost(Ore, PowderBuggy, OutParts)
		EqCost(Ore, PowderBuggy, OutFuel) = fuel * CellValues(Diesel, 10, WhichSegment).Value * CellValues(EquipmentHours, 10, WhichSegment).Value
		EqCost(Ore, PowderBuggy, OutLube) = CellValues(Lubricants, 10, WhichSegment).Value * CellValues(EquipmentHours, 10, WhichSegment).Value * TaxRate
		OpBin(2) = OpBin(2) + EqCost(Ore, PowderBuggy, OutLube)
		EqCost(Ore, PowderBuggy, OutTires) = (CellValues(Tires, 10, WhichSegment).Value / EqDefault(PowderBuggy, j, TireLife)) * CellValues(EquipmentHours, 10, WhichSegment).Value * TaxRate
		OpBin(5) = OpBin(5) + EqCost(Ore, PowderBuggy, OutTires)
		EqCost(Ore, PowderBuggy, OutMechanicTime) = CellValues(RepairLabor, 10, WhichSegment).Value * CellValues(EquipmentHours, 10, WhichSegment).Value
		EqCost(Ore, PowderBuggy, OutMechanicCost) = EqCost(Ore, PowderBuggy, OutMechanicTime) * mcwg * BurdenRate
		
		OutEquipment(10, WhichSegment) = EqCost(Ore, PowderBuggy, OutParts) + EqCost(Ore, PowderBuggy, OutFuel) + EqCost(Ore, PowderBuggy, OutLube) + EqCost(Ore, PowderBuggy, OutTires)
		
		If CellValues(EquipmentHours, 10, WhichSegment).Value > 0 And pbnm > 0 And dyyr > 0 Then
			EqCost(Ore, PowderBuggy, OutLife) = EqDefault(PowderBuggy, j, life) / ((CellValues(EquipmentHours, 10, WhichSegment).Value / pbnm) * dyyr)
		End If
		
		If CellValues(Replace_Renamed, 10, WhichSegment).Changed = False Then
			CellValues(Replace_Renamed, 10, WhichSegment).Value = EqCost(Ore, PowderBuggy, OutLife) * 12
		End If
		
		If CellValues(Purchase, 30, WhichSegment).Changed = False Then
			CellValues(Purchase, 30, WhichSegment).Value = CellValues(Purchase, 10, WhichSegment).Value * CellValues(EquipmentTwo, 16, WhichSegment).Value
		End If
		
	End Sub
	Public Sub RoadCost()
		Dim x As Short
		Dim OreRoadLength As Decimal
		Dim WasteRoadLength As Decimal
		Dim TempArea As Decimal
		Dim TempVolume As Decimal
		Dim TempOreThickness As Decimal
		Dim TempWasteThickness As Decimal
		Dim TempOreLength(5) As Decimal
		Dim TempWasteLength(5) As Decimal
		Dim DivOreVolume As Decimal
		Dim DivWasteVolume As Decimal
		Dim AverageWidthOre As Decimal
		Dim AverageWidthWaste As Decimal
		
		On Error Resume Next
		
		OreRoadLength = 0
		WasteRoadLength = 0
		
		For x = 0 To 2 Step 2
			OreRoadLength = OreRoadLength + CellValues(Haul, x, WhichSegment).Value
			TempOreLength(x / 2) = CellValues(Haul, x, WhichSegment).Value
			WasteRoadLength = WasteRoadLength + CellValues(Haul, x + 4, WhichSegment).Value
			TempWasteLength(x / 2) = CellValues(Haul, x + 4, WhichSegment).Value
		Next x
		For x = 8 To 14 Step 2
			OreRoadLength = OreRoadLength + CellValues(Haul, x, WhichSegment).Value
			TempOreLength((x - 4) / 2) = CellValues(Haul, x, WhichSegment).Value
			WasteRoadLength = WasteRoadLength + CellValues(Haul, x + 12, WhichSegment).Value
			TempWasteLength((x - 4) / 2) = CellValues(Haul, x + 12, WhichSegment).Value
		Next x
		
		If CellValues(Development, 0, WhichSegment).Changed = False Then
			CellValues(Development, 0, WhichSegment).Value = OreRoadLength
		End If
		
		If CellValues(Development, 1, WhichSegment).Changed = False Then
			CellValues(Development, 1, WhichSegment).Value = WasteRoadLength
		End If
		
		TempArea = 0
		TempVolume = 0
		
		For x = 0 To 10 Step 2
			TempArea = TempArea + (TempOreLength(x / 2) * CellValues(Road, x, WhichSegment).Value)
			TempVolume = TempVolume + (TempOreLength(x / 2) * CellValues(Road, x, WhichSegment).Value * (CellValues(Road, x + 1, WhichSegment).Value / 12))
		Next x
		
		TempOreThickness = TempVolume / TempArea
		DivOreVolume = ((TempVolume / OreRoadLength) / 27)
		AverageWidthOre = TempArea / OreRoadLength
		
		TempArea = 0
		TempVolume = 0
		
		For x = 0 To 10 Step 2
			TempArea = TempArea + (TempWasteLength(x / 2) * CellValues(Road, x + 12, WhichSegment).Value)
			TempVolume = TempVolume + (TempWasteLength(x / 2) * CellValues(Road, x + 12, WhichSegment).Value * (CellValues(Road, x + 13, WhichSegment).Value / 12))
		Next x
		
		TempWasteThickness = TempVolume / TempArea
		DivWasteVolume = ((TempVolume / WasteRoadLength) / 27)
		AverageWidthWaste = TempArea / WasteRoadLength
		
		'Update These Costs - 2014
		If CellValues(Development, 11, WhichSegment).Changed = False And TempOreThickness > 0 Then
			CellValues(Development, 11, WhichSegment).Value = (((2.182778 * ((TempOreThickness * 10) ^ 0.890617)) / 9) * AverageWidthOre)
		End If
		
		If CellValues(Development, 12, WhichSegment).Changed = False And TempWasteThickness > 0 Then
			CellValues(Development, 12, WhichSegment).Value = (((2.182778 * ((TempWasteThickness * 10) ^ 0.890617)) / 9) * AverageWidthWaste)
		End If
		
	End Sub
	Public Sub LaborCost()
		Dim ShiftDiff As Decimal
		Dim r As Short
		Dim s As Short
		Dim x As Short
		Dim y As Short
		
		On Error Resume Next
		
		OpBin(0) = 0
		
		For x = MinTime To MaxTime
			CellValues(Summary, 1, 0).Value = 0
			For s = 0 To 6
				If x >= CellValues(Production, 15, s).Value And x <= CellValues(Production, 16, s).Value Then
					r = s
				End If
			Next s
			Select Case CellValues(Production, 2, r).Value
				Case 1
					ShiftDiff = 0
				Case 2
					ShiftDiff = CellValues(Wage, 12, r).Value / 2
				Case 3
					ShiftDiff = CellValues(Wage, 12, r).Value / 3
			End Select
			CellValues(Summary, 1, x).Value = 0
			For y = 0 To 6
				CellValues(LaborResult, y, x).Value = ((CellValues(Wage, y, r).Value + ShiftDiff) * (1 + (CellValues(Wage, 9, r).Value / 100)) * CellValues(Production, 1, r).Value) * CellValues(WorkForce, y, r).Value
				CellValues(LaborResult, y, x).Word = CellValues(Wage, y, r).Word
				CellValues(Summary, 1, x).Value = CellValues(Summary, 1, x).Value + CellValues(LaborResult, y, x).Value
			Next y
			CellValues(LaborResult, 7, x).Value = ((CellValues(Wage, 11, r).Value + ShiftDiff) * (1 + (CellValues(Wage, 9, r).Value / 100)) * CellValues(Production, 1, r).Value) * CellValues(WorkForce, 11, r).Value
			CellValues(Summary, 1, x).Value = CellValues(Summary, 1, x).Value + CellValues(LaborResult, 7, x).Value
			CellValues(LaborResult, 7, x).Word = CellValues(Wage, 11, r).Word
			CellValues(LaborResult, 8, x).Value = ((CellValues(Wage, 8, r).Value + ShiftDiff) * (1 + (CellValues(Wage, 9, r).Value / 100)) * CellValues(Production, 1, r).Value) * CellValues(WorkForce, 8, r).Value
			CellValues(Summary, 1, x).Value = CellValues(Summary, 1, x).Value + CellValues(LaborResult, 8, x).Value
			CellValues(LaborResult, 8, x).Word = CellValues(Wage, 8, r).Word
			CellValues(LaborResult, 9, x).Value = ((CellValues(Wage, 7, r).Value + ShiftDiff) * (1 + (CellValues(Wage, 9, r).Value / 100)) * CellValues(Production, 1, r).Value) * CellValues(WorkForce, 7, r).Value
			CellValues(Summary, 1, x).Value = CellValues(Summary, 1, x).Value + CellValues(LaborResult, 9, x).Value
			CellValues(LaborResult, 9, x).Word = CellValues(Wage, 7, r).Word
			OpBin(0) = OpBin(0) + CellValues(Summary, 1, x).Value
		Next x
		
	End Sub
	Public Sub SalaryCost()
		Dim r As Short
		Dim s As Short
		Dim x As Short
		Dim y As Short
		Dim z As Short
		Dim MaxNum As Short
		Dim MaxSeg As Short
		Dim count As Short
		
		On Error Resume Next
		
		For z = 0 To 6
			If MaxNum < CellValues(Production, 15, z).Value - 1 Then
				MaxNum = CellValues(Production, 15, z).Value - 1
				MaxSeg = z
			End If
		Next z
		
		count = 0
		
		On Error Resume Next
		
		For x = MinTime To MaxTime
			For s = 0 To 6
				If x >= CellValues(Production, 15, s).Value And x <= CellValues(Production, 16, s).Value Then
					r = s
				End If
			Next s
			CellValues(Summary, 2, x).Value = 0
			For y = 0 To 11
				If CellValues(Production, 3, r).Value <> 0 Then
					CellValues(SalaryResult, y, x).Value = ((CellValues(Salary, y + 12, r).Value * (1 + (CellValues(Salary, 24, r).Value / 100))) * CellValues(Staff, y, r).Value) / CellValues(Production, 3, r).Value
				End If
				CellValues(Summary, 2, x).Value = CellValues(Summary, 2, x).Value + CellValues(SalaryResult, y, x).Value
				CellValues(SalaryResult, y, x).Word = CellValues(Salary, y + 12, r).Word
			Next y
			OpBin(0) = OpBin(0) + CellValues(Summary, 2, x).Value
			count = count + 1
		Next x
		
		If count > 0 Then
			OpBin(0) = (OpBin(0) / count) * MaxSeg
		End If
		
	End Sub
	Public Sub SupplyCost()
		Dim r As Short
		Dim s As Short
		Dim x As Short
		Dim y As Short
		Dim z As Short
		Dim MaxNum As Short
		Dim MaxSeg As Short
		Dim count As Short
		
		On Error Resume Next
		
		For z = 0 To 6
			If MaxNum < CellValues(Production, 15, z).Value - 1 Then
				MaxNum = CellValues(Production, 15, z).Value - 1
				MaxSeg = z
			End If
		Next z
		
		count = 0
		OpBin(6) = 0
		OpBin(7) = 0
		For x = MinTime To MaxTime
			For s = 0 To 6
				If x >= CellValues(Production, 15, s).Value And x <= CellValues(Production, 16, s).Value Then
					r = s
				End If
			Next s
			CellValues(Summary, 0, x).Value = 0
			For y = 0 To 12
				CellValues(SupplyResult, y, x).Value = OutSupply(y, r)
				CellValues(Summary, 0, x).Value = CellValues(Summary, 0, x).Value + OutSupply(y, r)
				Select Case y
					Case 0, 1, 2, 3, 6, 7, 8, 9
						OpBin(6) = OpBin(6) + OutSupply(y, r)
					Case 4, 5, 10, 11
						OpBin(7) = OpBin(7) + OutSupply(y, r)
				End Select
			Next y
			CellValues(SupplyResult, 0, x).Word = "Explosives - Ore"
			CellValues(SupplyResult, 1, x).Word = "Caps - Ore"
			CellValues(SupplyResult, 2, x).Word = "Primers - Ore"
			CellValues(SupplyResult, 3, x).Word = "Detonation Cord - Ore"
			CellValues(SupplyResult, 4, x).Word = "Drill Bits - Ore"
			CellValues(SupplyResult, 5, x).Word = "Drill Steel - Ore"
			CellValues(SupplyResult, 6, x).Word = "Explosives - Waste"
			CellValues(SupplyResult, 7, x).Word = "Caps - Waste"
			CellValues(SupplyResult, 8, x).Word = "Primers - Waste"
			CellValues(SupplyResult, 9, x).Word = "Detonation Cord - Waste"
			CellValues(SupplyResult, 10, x).Word = "Drill Bits - Waste"
			CellValues(SupplyResult, 11, x).Word = "Drill Steel - Waste"
			CellValues(SupplyResult, 12, x).Word = "Sundry Items"
			count = count + 1
		Next x
		
		If count > 0 Then
			OpBin(6) = (OpBin(6) / count) * MaxSeg
			OpBin(7) = (OpBin(7) / count) * MaxSeg
		End If
		
	End Sub
	Public Sub EquipmentCost()
		Dim r As Short
		Dim s As Short
		Dim x As Short
		Dim y As Short
		Dim i As Short
		Dim AddNumber As Short
		Dim SumCapital(6) As Decimal
		
		On Error Resume Next
		
		Call TimeLineCalc()
		
		For y = 0 To MaxTime
			CellValues(Summary, 8, y).Value = 0
		Next y
		
		For y = MinTime To MaxTime
			r = 0
			For s = 0 To 6
				If y >= CellValues(Production, 15, s).Value And y <= CellValues(Production, 16, s).Value Then
					r = s
				End If
			Next s
			CellValues(Summary, 3, y).Value = 0
			AddNumber = 0
			For i = 0 To 13
				Select Case i
					Case 0
						For x = 0 To 3
							If CellValues(EquipmentOne, x, WhichSegment).Value <> 0 Then
								Select Case x
									Case 0
										CellValues(EquipmentHourlyResult, AddNumber, y).Word = "Front-End Loader - Ore"
										CellValues(EquipmentHourlyResult, AddNumber, y).Value = OutEquipment(i, r)
										CellValues(Summary, 3, y).Value = CellValues(Summary, 3, y).Value + OutEquipment(i, r)
										CellValues(EquipmentPurchaseResult, AddNumber, r).Value = CellValues(Purchase, i + 20, r).Value
										CellValues(EquipmentNumberResult, AddNumber, r).Value = CellValues(EquipmentTwo, x, r).Value
										AddNumber = AddNumber + 1
									Case 1
										CellValues(EquipmentHourlyResult, AddNumber, y).Word = "Hydraulic Shovel - Ore"
										CellValues(EquipmentHourlyResult, AddNumber, y).Value = OutEquipment(i, r)
										CellValues(Summary, 3, y).Value = CellValues(Summary, 3, y).Value + OutEquipment(i, r)
										CellValues(EquipmentPurchaseResult, AddNumber, r).Value = CellValues(Purchase, i + 20, r).Value
										CellValues(EquipmentNumberResult, AddNumber, r).Value = CellValues(EquipmentTwo, x, r).Value
										AddNumber = AddNumber + 1
									Case 2
										CellValues(EquipmentHourlyResult, AddNumber, y).Word = "Mechanical Shovel - Ore"
										CellValues(EquipmentHourlyResult, AddNumber, y).Value = OutEquipment(i, r)
										CellValues(Summary, 3, y).Value = CellValues(Summary, 3, y).Value + OutEquipment(i, r)
										CellValues(EquipmentPurchaseResult, AddNumber, r).Value = CellValues(Purchase, i + 20, r).Value
										CellValues(EquipmentNumberResult, AddNumber, r).Value = CellValues(EquipmentTwo, x, r).Value
										AddNumber = AddNumber + 1
									Case 3
										CellValues(EquipmentHourlyResult, AddNumber, y).Word = "Walking Dragline - Ore"
										CellValues(EquipmentHourlyResult, AddNumber, y).Value = OutEquipment(i, r)
										CellValues(Summary, 3, y).Value = CellValues(Summary, 3, y).Value + OutEquipment(i, r)
										CellValues(EquipmentPurchaseResult, AddNumber, r).Value = CellValues(Purchase, i + 20, r).Value
										CellValues(EquipmentNumberResult, AddNumber, r).Value = CellValues(EquipmentTwo, x, r).Value
										AddNumber = AddNumber + 1
								End Select
							End If
						Next x
						If CellValues(EquipmentOne, 20, WhichSegment).Value <> 0 Then
							CellValues(EquipmentHourlyResult, AddNumber, y).Word = "Scraper - Ore"
							CellValues(EquipmentHourlyResult, AddNumber, y).Value = OutEquipment(i, r)
							CellValues(Summary, 3, y).Value = CellValues(Summary, 3, y).Value + OutEquipment(i, r)
							CellValues(EquipmentPurchaseResult, AddNumber, r).Value = CellValues(Purchase, i + 20, r).Value
							CellValues(EquipmentNumberResult, AddNumber, r).Value = CellValues(EquipmentTwo, 20, r).Value
							AddNumber = AddNumber + 1
						End If
					Case 1
						If CellValues(EquipmentOne, 4, WhichSegment).Value <> 0 Then
							CellValues(EquipmentHourlyResult, AddNumber, y).Word = "Rear-Dump Truck - Ore"
							CellValues(EquipmentHourlyResult, AddNumber, y).Value = OutEquipment(i, r)
							CellValues(Summary, 3, y).Value = CellValues(Summary, 3, y).Value + OutEquipment(i, r)
							CellValues(EquipmentPurchaseResult, AddNumber, r).Value = CellValues(Purchase, i + 20, r).Value
							CellValues(EquipmentNumberResult, AddNumber, r).Value = CellValues(EquipmentTwo, 4, r).Value
							AddNumber = AddNumber + 1
						ElseIf CellValues(EquipmentOne, 21, WhichSegment).Value <> 0 Then 
							CellValues(EquipmentHourlyResult, AddNumber, y).Word = "Articulated Haul Truck - Ore"
							CellValues(EquipmentHourlyResult, AddNumber, y).Value = OutEquipment(i, r)
							CellValues(Summary, 3, y).Value = CellValues(Summary, 3, y).Value + OutEquipment(i, r)
							CellValues(EquipmentPurchaseResult, AddNumber, r).Value = CellValues(Purchase, i + 20, r).Value
							CellValues(EquipmentNumberResult, AddNumber, r).Value = CellValues(EquipmentTwo, 21, r).Value
							AddNumber = AddNumber + 1
						ElseIf CellValues(EquipmentOne, 22, WhichSegment).Value <> 0 Then 
							CellValues(EquipmentHourlyResult, AddNumber, y).Word = "Jaw Crusher - Ore"
							CellValues(EquipmentHourlyResult, AddNumber, y).Value = OutEquipment(14, r)
							CellValues(Summary, 3, y).Value = CellValues(Summary, 3, y).Value + OutEquipment(14, r)
							CellValues(EquipmentPurchaseResult, AddNumber, r).Value = CellValues(Purchase, i + 33, r).Value
							CellValues(EquipmentNumberResult, AddNumber, r).Value = CellValues(EquipmentTwo, 22, r).Value
							AddNumber = AddNumber + 1
							CellValues(EquipmentHourlyResult, AddNumber, y).Word = "Conveyor - Ore"
							CellValues(EquipmentHourlyResult, AddNumber, y).Value = OutEquipment(15, r)
							CellValues(Summary, 3, y).Value = CellValues(Summary, 3, y).Value + OutEquipment(15, r)
							CellValues(EquipmentPurchaseResult, AddNumber, r).Value = CellValues(Purchase, i + 34, r).Value
							CellValues(EquipmentNumberResult, AddNumber, r).Value = CellValues(EquipmentTwo, 24, r).Value
							AddNumber = AddNumber + 1
						ElseIf CellValues(EquipmentOne, 23, WhichSegment).Value <> 0 Then 
							CellValues(EquipmentHourlyResult, AddNumber, y).Word = "Gyratory Crusher - Ore"
							CellValues(EquipmentHourlyResult, AddNumber, y).Value = OutEquipment(14, r)
							CellValues(Summary, 3, y).Value = CellValues(Summary, 3, y).Value + OutEquipment(14, r)
							CellValues(EquipmentPurchaseResult, AddNumber, r).Value = CellValues(Purchase, i + 33, r).Value
							CellValues(EquipmentNumberResult, AddNumber, r).Value = CellValues(EquipmentTwo, 23, r).Value
							AddNumber = AddNumber + 1
							CellValues(EquipmentHourlyResult, AddNumber, y).Word = "Conveyor - Ore"
							CellValues(EquipmentHourlyResult, AddNumber, y).Value = OutEquipment(15, r)
							CellValues(Summary, 3, y).Value = CellValues(Summary, 3, y).Value + OutEquipment(15, r)
							CellValues(EquipmentPurchaseResult, AddNumber, r).Value = CellValues(Purchase, i + 34, r).Value
							CellValues(EquipmentNumberResult, AddNumber, r).Value = CellValues(EquipmentTwo, 24, r).Value
							AddNumber = AddNumber + 1
						End If
					Case 2
						For x = 5 To 8
							If CellValues(EquipmentOne, x, WhichSegment).Value <> 0 Then
								Select Case x
									Case 5
										CellValues(EquipmentHourlyResult, AddNumber, y).Word = "Front-End Loader - Waste"
										CellValues(EquipmentHourlyResult, AddNumber, y).Value = OutEquipment(i, r)
										CellValues(Summary, 3, y).Value = CellValues(Summary, 3, y).Value + OutEquipment(i, r)
										CellValues(EquipmentPurchaseResult, AddNumber, r).Value = CellValues(Purchase, i + 20, r).Value
										CellValues(EquipmentNumberResult, AddNumber, r).Value = CellValues(EquipmentTwo, x, r).Value
										AddNumber = AddNumber + 1
									Case 6
										CellValues(EquipmentHourlyResult, AddNumber, y).Word = "Hydraulic Shovel - Waste"
										CellValues(EquipmentHourlyResult, AddNumber, y).Value = OutEquipment(i, r)
										CellValues(Summary, 3, y).Value = CellValues(Summary, 3, y).Value + OutEquipment(i, r)
										CellValues(EquipmentPurchaseResult, AddNumber, r).Value = CellValues(Purchase, i + 20, r).Value
										CellValues(EquipmentNumberResult, AddNumber, r).Value = CellValues(EquipmentTwo, x, r).Value
										AddNumber = AddNumber + 1
									Case 7
										CellValues(EquipmentHourlyResult, AddNumber, y).Word = "Mechanical Shovel - Waste"
										CellValues(EquipmentHourlyResult, AddNumber, y).Value = OutEquipment(i, r)
										CellValues(Summary, 3, y).Value = CellValues(Summary, 3, y).Value + OutEquipment(i, r)
										CellValues(EquipmentPurchaseResult, AddNumber, r).Value = CellValues(Purchase, i + 20, r).Value
										CellValues(EquipmentNumberResult, AddNumber, r).Value = CellValues(EquipmentTwo, x, r).Value
										AddNumber = AddNumber + 1
									Case 8
										CellValues(EquipmentHourlyResult, AddNumber, y).Word = "Walking Dragline - Waste"
										CellValues(EquipmentHourlyResult, AddNumber, y).Value = OutEquipment(i, r)
										CellValues(Summary, 3, y).Value = CellValues(Summary, 3, y).Value + OutEquipment(i, r)
										CellValues(EquipmentPurchaseResult, AddNumber, r).Value = CellValues(Purchase, i + 20, r).Value
										CellValues(EquipmentNumberResult, AddNumber, r).Value = CellValues(EquipmentTwo, x, r).Value
										AddNumber = AddNumber + 1
								End Select
							End If
						Next x
						If CellValues(EquipmentOne, 25, WhichSegment).Value <> 0 Then
							CellValues(EquipmentHourlyResult, AddNumber, y).Word = "Scraper - Waste"
							CellValues(EquipmentHourlyResult, AddNumber, y).Value = OutEquipment(i, r)
							CellValues(Summary, 3, y).Value = CellValues(Summary, 3, y).Value + OutEquipment(i, r)
							CellValues(EquipmentPurchaseResult, AddNumber, r).Value = CellValues(Purchase, i + 20, r).Value
							CellValues(EquipmentNumberResult, AddNumber, r).Value = CellValues(EquipmentTwo, 25, r).Value
							AddNumber = AddNumber + 1
						End If
					Case 3
						If CellValues(EquipmentOne, 9, WhichSegment).Value <> 0 Then
							CellValues(EquipmentHourlyResult, AddNumber, y).Word = "Rear-Dump Truck - Waste"
							CellValues(EquipmentHourlyResult, AddNumber, y).Value = OutEquipment(i, r)
							CellValues(Summary, 3, y).Value = CellValues(Summary, 3, y).Value + OutEquipment(i, r)
							CellValues(EquipmentPurchaseResult, AddNumber, r).Value = CellValues(Purchase, i + 20, r).Value
							CellValues(EquipmentNumberResult, AddNumber, r).Value = CellValues(EquipmentTwo, 9, r).Value
							AddNumber = AddNumber + 1
						ElseIf CellValues(EquipmentOne, 26, WhichSegment).Value <> 0 Then 
							CellValues(EquipmentHourlyResult, AddNumber, y).Word = "Articulated Haul Truck - Waste"
							CellValues(EquipmentHourlyResult, AddNumber, y).Value = OutEquipment(i, r)
							CellValues(Summary, 3, y).Value = CellValues(Summary, 3, y).Value + OutEquipment(i, r)
							CellValues(EquipmentPurchaseResult, AddNumber, r).Value = CellValues(Purchase, i + 20, r).Value
							CellValues(EquipmentNumberResult, AddNumber, r).Value = CellValues(EquipmentTwo, 26, r).Value
							AddNumber = AddNumber + 1
						ElseIf CellValues(EquipmentOne, 27, WhichSegment).Value <> 0 Then 
							CellValues(EquipmentHourlyResult, AddNumber, y).Word = "Jaw Crusher - Waste"
							CellValues(EquipmentHourlyResult, AddNumber, y).Value = OutEquipment(16, r)
							CellValues(Summary, 3, y).Value = CellValues(Summary, 3, y).Value + OutEquipment(16, r)
							CellValues(EquipmentPurchaseResult, AddNumber, r).Value = CellValues(Purchase, i + 33, r).Value
							CellValues(EquipmentNumberResult, AddNumber, r).Value = CellValues(EquipmentTwo, 27, r).Value
							AddNumber = AddNumber + 1
							CellValues(EquipmentHourlyResult, AddNumber, y).Word = "Conveyor - Waste"
							CellValues(EquipmentHourlyResult, AddNumber, y).Value = OutEquipment(17, r)
							CellValues(Summary, 3, y).Value = CellValues(Summary, 3, y).Value + OutEquipment(17, r)
							CellValues(EquipmentPurchaseResult, AddNumber, r).Value = CellValues(Purchase, i + 34, r).Value
							CellValues(EquipmentNumberResult, AddNumber, r).Value = CellValues(EquipmentTwo, 29, r).Value
							AddNumber = AddNumber + 1
						ElseIf CellValues(EquipmentOne, 28, WhichSegment).Value <> 0 Then 
							CellValues(EquipmentHourlyResult, AddNumber, y).Word = "Gyratory Crusher - Waste"
							CellValues(EquipmentHourlyResult, AddNumber, y).Value = OutEquipment(16, r)
							CellValues(Summary, 3, y).Value = CellValues(Summary, 3, y).Value + OutEquipment(16, r)
							CellValues(EquipmentPurchaseResult, AddNumber, r).Value = CellValues(Purchase, i + 33, r).Value
							CellValues(EquipmentNumberResult, AddNumber, r).Value = CellValues(EquipmentTwo, 28, r).Value
							AddNumber = AddNumber + 1
							CellValues(EquipmentHourlyResult, AddNumber, y).Word = "Conveyor - Waste"
							CellValues(EquipmentHourlyResult, AddNumber, y).Value = OutEquipment(17, r)
							CellValues(Summary, 3, y).Value = CellValues(Summary, 3, y).Value + OutEquipment(17, r)
							CellValues(EquipmentPurchaseResult, AddNumber, r).Value = CellValues(Purchase, i + 34, r).Value
							CellValues(EquipmentNumberResult, AddNumber, r).Value = CellValues(EquipmentTwo, 29, r).Value
							AddNumber = AddNumber + 1
						End If
					Case 4
						If CellValues(EquipmentOne, 10, WhichSegment).Value <> 0 Then
							CellValues(EquipmentHourlyResult, AddNumber, y).Word = "Percussion Drill - Ore"
							CellValues(EquipmentHourlyResult, AddNumber, y).Value = OutEquipment(i, r)
							CellValues(Summary, 3, y).Value = CellValues(Summary, 3, y).Value + OutEquipment(i, r)
							CellValues(EquipmentPurchaseResult, AddNumber, r).Value = CellValues(Purchase, i + 20, r).Value
							CellValues(EquipmentNumberResult, AddNumber, r).Value = CellValues(EquipmentTwo, 10, r).Value
							AddNumber = AddNumber + 1
						ElseIf CellValues(EquipmentOne, 11, WhichSegment).Value <> 0 Then 
							CellValues(EquipmentHourlyResult, AddNumber, y).Word = "Rotary Drill - Ore"
							CellValues(EquipmentHourlyResult, AddNumber, y).Value = OutEquipment(i + 1, r)
							CellValues(Summary, 3, y).Value = CellValues(Summary, 3, y).Value + OutEquipment(i + 1, r)
							CellValues(EquipmentPurchaseResult, AddNumber, r).Value = CellValues(Purchase, i + 21, r).Value
							CellValues(EquipmentNumberResult, AddNumber, r).Value = CellValues(EquipmentTwo, 11, r).Value
							AddNumber = AddNumber + 1
						End If
					Case 5
						If CellValues(EquipmentOne, 30, WhichSegment).Value <> 0 Then
							CellValues(EquipmentHourlyResult, AddNumber, y).Word = "Percussion Drill - Waste"
							CellValues(EquipmentHourlyResult, AddNumber, y).Value = OutEquipment(i + 13, r)
							CellValues(Summary, 3, y).Value = CellValues(Summary, 3, y).Value + OutEquipment(i + 13, r)
							CellValues(EquipmentPurchaseResult, AddNumber, r).Value = CellValues(Purchase, i + 33, r).Value
							CellValues(EquipmentNumberResult, AddNumber, r).Value = CellValues(EquipmentTwo, 30, r).Value
							AddNumber = AddNumber + 1
						ElseIf CellValues(EquipmentOne, 31, WhichSegment).Value <> 0 Then 
							CellValues(EquipmentHourlyResult, AddNumber, y).Word = "Rotary Drill - Waste"
							CellValues(EquipmentHourlyResult, AddNumber, y).Value = OutEquipment(i + 14, r)
							CellValues(Summary, 3, y).Value = CellValues(Summary, 3, y).Value + OutEquipment(i + 14, r)
							CellValues(EquipmentPurchaseResult, AddNumber, r).Value = CellValues(Purchase, i + 34, r).Value
							CellValues(EquipmentNumberResult, AddNumber, r).Value = CellValues(EquipmentTwo, 31, r).Value
							AddNumber = AddNumber + 1
						End If
					Case 6
						If CellValues(EquipmentOne, i + 6, WhichSegment).Value <> 0 Then
							CellValues(EquipmentHourlyResult, AddNumber, y).Word = "Bulldozers"
							CellValues(EquipmentHourlyResult, AddNumber, y).Value = OutEquipment(i, r)
							CellValues(Summary, 3, y).Value = CellValues(Summary, 3, y).Value + OutEquipment(i, r)
							CellValues(EquipmentPurchaseResult, AddNumber, r).Value = CellValues(Purchase, i + 20, r).Value
							CellValues(EquipmentNumberResult, AddNumber, r).Value = CellValues(EquipmentTwo, i + 6, r).Value
							AddNumber = AddNumber + 1
						End If
					Case 7
						If CellValues(EquipmentOne, i + 6, WhichSegment).Value <> 0 Then
							CellValues(EquipmentHourlyResult, AddNumber, y).Word = "Graders"
							CellValues(EquipmentHourlyResult, AddNumber, y).Value = OutEquipment(i, r)
							CellValues(Summary, 3, y).Value = CellValues(Summary, 3, y).Value + OutEquipment(i, r)
							CellValues(EquipmentPurchaseResult, AddNumber, r).Value = CellValues(Purchase, i + 20, r).Value
							CellValues(EquipmentNumberResult, AddNumber, r).Value = CellValues(EquipmentTwo, i + 6, r).Value
							AddNumber = AddNumber + 1
						End If
					Case 8
						If CellValues(EquipmentOne, i + 6, WhichSegment).Value <> 0 Then
							CellValues(EquipmentHourlyResult, AddNumber, y).Word = "Water Tankers"
							CellValues(EquipmentHourlyResult, AddNumber, y).Value = OutEquipment(i, r)
							CellValues(Summary, 3, y).Value = CellValues(Summary, 3, y).Value + OutEquipment(i, r)
							CellValues(EquipmentPurchaseResult, AddNumber, r).Value = CellValues(Purchase, i + 20, r).Value
							CellValues(EquipmentNumberResult, AddNumber, r).Value = CellValues(EquipmentTwo, i + 6, r).Value
							AddNumber = AddNumber + 1
						End If
					Case 9
						If CellValues(EquipmentOne, i + 6, WhichSegment).Value <> 0 Then
							CellValues(EquipmentHourlyResult, AddNumber, y).Word = "Tire Service Trucks"
							CellValues(EquipmentHourlyResult, AddNumber, y).Value = OutEquipment(i, r)
							CellValues(Summary, 3, y).Value = CellValues(Summary, 3, y).Value + OutEquipment(i, r)
							CellValues(EquipmentPurchaseResult, AddNumber, r).Value = CellValues(Purchase, i + 20, r).Value
							CellValues(EquipmentNumberResult, AddNumber, r).Value = CellValues(EquipmentTwo, i + 6, r).Value
							AddNumber = AddNumber + 1
						End If
					Case 10
						If CellValues(EquipmentOne, i + 6, WhichSegment).Value <> 0 Then
							If bltp = 1 Then
								CellValues(EquipmentHourlyResult, AddNumber, y).Word = "Powder Buggies"
							ElseIf bltp = 2 Then 
								CellValues(EquipmentHourlyResult, AddNumber, y).Word = "Bulk Trucks"
							End If
							CellValues(EquipmentHourlyResult, AddNumber, y).Value = OutEquipment(i, r)
							CellValues(Summary, 3, y).Value = CellValues(Summary, 3, y).Value + OutEquipment(i, r)
							CellValues(EquipmentPurchaseResult, AddNumber, r).Value = CellValues(Purchase, i + 20, r).Value
							CellValues(EquipmentNumberResult, AddNumber, r).Value = CellValues(EquipmentTwo, i + 6, r).Value
							AddNumber = AddNumber + 1
						End If
					Case 11
						If CellValues(EquipmentOne, i + 6, WhichSegment).Value <> 0 Then
							CellValues(EquipmentHourlyResult, AddNumber, y).Word = "Lighting Plants"
							CellValues(EquipmentHourlyResult, AddNumber, y).Value = OutEquipment(i, r)
							CellValues(Summary, 3, y).Value = CellValues(Summary, 3, y).Value + OutEquipment(i, r)
							CellValues(EquipmentPurchaseResult, AddNumber, r).Value = CellValues(Purchase, i + 20, r).Value
							CellValues(EquipmentNumberResult, AddNumber, r).Value = CellValues(EquipmentTwo, i + 6, r).Value
							AddNumber = AddNumber + 1
						End If
					Case 12
						If CellValues(EquipmentOne, i + 6, WhichSegment).Value <> 0 Then
							CellValues(EquipmentHourlyResult, AddNumber, y).Word = "Pumps"
							CellValues(EquipmentHourlyResult, AddNumber, y).Value = OutEquipment(i, r)
							CellValues(Summary, 3, y).Value = CellValues(Summary, 3, y).Value + OutEquipment(i, r)
							CellValues(EquipmentPurchaseResult, AddNumber, r).Value = CellValues(Purchase, i + 20, r).Value
							CellValues(EquipmentNumberResult, AddNumber, r).Value = CellValues(EquipmentTwo, i + 6, r).Value
							AddNumber = AddNumber + 1
						End If
					Case 13
						If CellValues(EquipmentOne, i + 6, WhichSegment).Value <> 0 Then
							CellValues(EquipmentHourlyResult, AddNumber, y).Word = "Pick-up Trucks"
							CellValues(EquipmentHourlyResult, AddNumber, y).Value = OutEquipment(i, r)
							CellValues(Summary, 3, y).Value = CellValues(Summary, 3, y).Value + OutEquipment(i, r)
							CellValues(EquipmentPurchaseResult, AddNumber, r).Value = CellValues(Purchase, i + 20, r).Value
							CellValues(EquipmentNumberResult, AddNumber, r).Value = CellValues(EquipmentTwo, i + 6, r).Value
						End If
				End Select
			Next i
		Next y
		
		For y = 0 To MaxSegment
			SumCapital(y) = 0
			AddNumber = 0
			For i = 0 To 13
				Select Case i
					Case 0
						For x = 0 To 3
							If CellValues(EquipmentOne, x, WhichSegment).Value <> 0 Then
								Select Case x
									Case 0
										SumCapital(y) = SumCapital(y) + CellValues(Purchase, i + 20, y).Value
										AddNumber = AddNumber + 1
									Case 1
										SumCapital(y) = SumCapital(y) + CellValues(Purchase, i + 20, y).Value
										AddNumber = AddNumber + 1
									Case 2
										SumCapital(y) = SumCapital(y) + CellValues(Purchase, i + 20, y).Value
										AddNumber = AddNumber + 1
									Case 3
										SumCapital(y) = SumCapital(y) + CellValues(Purchase, i + 20, y).Value
										AddNumber = AddNumber + 1
								End Select
							End If
						Next x
						If CellValues(EquipmentOne, 20, WhichSegment).Value <> 0 Then
							SumCapital(y) = SumCapital(y) + CellValues(Purchase, i + 20, y).Value
							AddNumber = AddNumber + 1
						End If
					Case 1
						If CellValues(EquipmentOne, 4, WhichSegment).Value <> 0 Then
							SumCapital(y) = SumCapital(y) + CellValues(Purchase, i + 20, y).Value
							AddNumber = AddNumber + 1
						ElseIf CellValues(EquipmentOne, 21, WhichSegment).Value <> 0 Then 
							SumCapital(y) = SumCapital(y) + CellValues(Purchase, i + 20, y).Value
							AddNumber = AddNumber + 1
						ElseIf CellValues(EquipmentOne, 22, WhichSegment).Value <> 0 Then 
							SumCapital(y) = SumCapital(y) + CellValues(Purchase, i + 20, y).Value
							AddNumber = AddNumber + 1
							SumCapital(y) = SumCapital(y) + CellValues(Purchase, i + 20, y).Value
							AddNumber = AddNumber + 1
						ElseIf CellValues(EquipmentOne, 23, WhichSegment).Value <> 0 Then 
							SumCapital(y) = SumCapital(y) + CellValues(Purchase, i + 20, y).Value
							AddNumber = AddNumber + 1
							SumCapital(y) = SumCapital(y) + CellValues(Purchase, i + 20, y).Value
							AddNumber = AddNumber + 1
						End If
					Case 2
						For x = 5 To 8
							If CellValues(EquipmentOne, x, WhichSegment).Value <> 0 Then
								Select Case x
									Case 5
										SumCapital(y) = SumCapital(y) + CellValues(Purchase, i + 20, y).Value
										AddNumber = AddNumber + 1
									Case 6
										SumCapital(y) = SumCapital(y) + CellValues(Purchase, i + 20, y).Value
										AddNumber = AddNumber + 1
									Case 7
										SumCapital(y) = SumCapital(y) + CellValues(Purchase, i + 20, y).Value
										AddNumber = AddNumber + 1
									Case 8
										SumCapital(y) = SumCapital(y) + CellValues(Purchase, i + 20, y).Value
										AddNumber = AddNumber + 1
								End Select
							End If
						Next x
						If CellValues(EquipmentOne, 25, WhichSegment).Value <> 0 Then
							SumCapital(y) = SumCapital(y) + CellValues(Purchase, i + 20, y).Value
							AddNumber = AddNumber + 1
						End If
					Case 3
						If CellValues(EquipmentOne, 9, WhichSegment).Value <> 0 Then
							SumCapital(y) = SumCapital(y) + CellValues(Purchase, i + 20, y).Value
							AddNumber = AddNumber + 1
						ElseIf CellValues(EquipmentOne, 26, WhichSegment).Value <> 0 Then 
							SumCapital(y) = SumCapital(y) + CellValues(Purchase, i + 20, y).Value
							AddNumber = AddNumber + 1
						ElseIf CellValues(EquipmentOne, 27, WhichSegment).Value <> 0 Then 
							SumCapital(y) = SumCapital(y) + CellValues(Purchase, i + 20, y).Value
							AddNumber = AddNumber + 1
							SumCapital(y) = SumCapital(y) + CellValues(Purchase, i + 20, y).Value
							AddNumber = AddNumber + 1
						ElseIf CellValues(EquipmentOne, 28, WhichSegment).Value <> 0 Then 
							SumCapital(y) = SumCapital(y) + CellValues(Purchase, i + 20, y).Value
							AddNumber = AddNumber + 1
							SumCapital(y) = SumCapital(y) + CellValues(Purchase, i + 20, y).Value
							AddNumber = AddNumber + 1
						End If
					Case 4
						If CellValues(EquipmentOne, 10, WhichSegment).Value <> 0 Then
							SumCapital(y) = SumCapital(y) + CellValues(Purchase, 4, y).Value
							AddNumber = AddNumber + 1
						ElseIf CellValues(EquipmentOne, 11, WhichSegment).Value <> 0 Then 
							SumCapital(y) = SumCapital(y) + CellValues(Purchase, 25, y).Value
							AddNumber = AddNumber + 1
						End If
					Case 5
						If CellValues(EquipmentOne, 30, WhichSegment).Value <> 0 Then
							SumCapital(y) = SumCapital(y) + CellValues(Purchase, 38, y).Value
							AddNumber = AddNumber + 1
						ElseIf CellValues(EquipmentOne, 31, WhichSegment).Value <> 0 Then 
							SumCapital(y) = SumCapital(y) + CellValues(Purchase, 39, y).Value
							AddNumber = AddNumber + 1
						End If
					Case 6
						If CellValues(EquipmentOne, i + 6, WhichSegment).Value <> 0 Then
							SumCapital(y) = SumCapital(y) + CellValues(Purchase, i + 20, y).Value
							AddNumber = AddNumber + 1
						End If
					Case 7
						If CellValues(EquipmentOne, i + 6, WhichSegment).Value <> 0 Then
							SumCapital(y) = SumCapital(y) + CellValues(Purchase, i + 20, y).Value
							AddNumber = AddNumber + 1
						End If
					Case 8
						If CellValues(EquipmentOne, i + 6, WhichSegment).Value <> 0 Then
							SumCapital(y) = SumCapital(y) + CellValues(Purchase, i + 20, y).Value
							AddNumber = AddNumber + 1
						End If
					Case 9
						If CellValues(EquipmentOne, i + 6, WhichSegment).Value <> 0 Then
							SumCapital(y) = SumCapital(y) + CellValues(Purchase, i + 20, y).Value
							AddNumber = AddNumber + 1
						End If
					Case 10
						If CellValues(EquipmentOne, i + 6, WhichSegment).Value <> 0 Then
							SumCapital(y) = SumCapital(y) + CellValues(Purchase, i + 20, y).Value
							AddNumber = AddNumber + 1
						End If
					Case 11
						If CellValues(EquipmentOne, i + 6, WhichSegment).Value <> 0 Then
							SumCapital(y) = SumCapital(y) + CellValues(Purchase, i + 20, y).Value
							AddNumber = AddNumber + 1
						End If
					Case 12
						If CellValues(EquipmentOne, i + 6, WhichSegment).Value <> 0 Then
							SumCapital(y) = SumCapital(y) + CellValues(Purchase, i + 20, y).Value
							AddNumber = AddNumber + 1
						End If
					Case 13
						If CellValues(EquipmentOne, i + 6, WhichSegment).Value <> 0 Then
							SumCapital(y) = SumCapital(y) + CellValues(Purchase, i + 20, y).Value
						End If
				End Select
			Next i
			If y = 0 Then
				CellValues(Summary, 8, 1).Value = SumCapital(y)
			Else
				CellValues(Summary, 8, Int(CellValues(Production, 15, y).Value)).Value = SumCapital(y)
			End If
		Next y
		
		Call ReplaceEngr()
		
	End Sub
	Public Sub SummaryCost()
		Dim x As Short
		Dim y As Short
		Dim ContBasis As Decimal
		
		On Error Resume Next
		
		For y = MinTime To MaxTime
			CellValues(Summary, 4, y).Value = 0
			CellValues(Summary, 5, y).Value = 0
			For x = 0 To 3
				CellValues(Summary, 5, y).Value = CellValues(Summary, 5, y).Value + CellValues(Summary, x, y).Value
			Next x
			CellValues(Summary, 4, y).Value = CellValues(Summary, 5, y).Value * 0.1
			CellValues(Summary, 5, y).Value = CellValues(Summary, 5, y).Value + CellValues(Summary, 4, y).Value
		Next y
		
		For y = 1 To MaxTime
			ContBasis = 0
			If y < CellValues(Production, 15, 0).Value Then
				For x = 8 To 12
					ContBasis = ContBasis + (CellValues(Summary, x, y).Value * 0.1)
				Next x
				CellValues(Summary, 13, y).Value = ContBasis
			End If
			CellValues(Summary, 14, y).Value = 0
			For x = 8 To 13
				CellValues(Summary, 14, y).Value = CellValues(Summary, 14, y).Value + CellValues(Summary, x, y).Value
			Next x
		Next y
		
	End Sub
	Public Sub DevelopmentCost()
		
		On Error Resume Next
		
		If WhichSegment = 0 Then
			Call dcalc()
		End If
		
		CellValues(DevelopmentResult, 0, 1).Word = "Preproduction Stripping"
		CellValues(DevelopmentResult, 1, 1).Word = "Haul Roads"
		CellValues(DevelopmentResult, 2, 1).Word = "Buildings"
		CellValues(DevelopmentResult, 3, 1).Word = "Explosives Storage"
		CellValues(DevelopmentResult, 4, 1).Word = "Electrical System"
		CellValues(DevelopmentResult, 5, 1).Word = "Clearing"
		CellValues(DevelopmentResult, 6, 1).Word = "Yard Preparation"
		CellValues(DevelopmentResult, 7, 1).Word = "Sewage Treatment"
		CellValues(DevelopmentResult, 8, 1).Word = "Fencing"
		CellValues(DevelopmentResult, 9, 1).Word = "Fuel Storage"
		
	End Sub
End Module