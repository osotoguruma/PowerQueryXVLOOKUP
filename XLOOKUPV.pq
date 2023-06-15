/*
AUTHOR       : OGURUMA1209@OUTLOOK.FR
DATE         : 11/06/2023
VERSION      : Initial
PURPOSE      : Excel RECHERCHEX in PowerQuery
DESCRIPTION  :

pLookupValue  : searched value
pTableName    : table lookup for - type may be text or pq table
tmpColumnKeyNumberIndex : searched column -- type may be text or number -- not necessary to be the firts column (as RechercheX)
tmpColumnReturnIndexNumber : returned column -- type me be texte or number
pSort : table lookup for need sorted or not -- eg speed
optional tmpMSG : msg if not found - default "#N/A"
*/

// Nested let started - 1
let pqXLOOKUPV1 = (pLookupValue as any, pTableName as any, tmpColumnKeyNumberIndex as any, tmpColumnReturnIndexNumber as any, pSort as logical, optional tmpMSG as any ) as any =>

	// Nested let started - 2
	let

		// Fixe parameters
		// ---------------
		pMSG = if tmpMSG = null then "#N/A" else tmpMSG,
		pColumnKeyNumberIndex = if tmpColumnKeyNumberIndex = 0 then 1 else tmpColumnKeyNumberIndex,
		pColumnReturnIndexNumber = if tmpColumnReturnIndexNumber = 0 then 1 else tmpColumnReturnIndexNumber,		
		pTableArray = if pTableName is table then pTableName else Excel.CurrentWorkbook(){[Name=pTableName]}[Content],	

		// ----------------------------------------------------------------------
		// Fixe index colunm index and return column -- either number or string
		// ----------------------------------------------------------------------
		pKeyIndex=if tmpColumnKeyNumberIndex is text then List.PositionOf(Columns,tmpColumnKeyNumberIndex) else tmpColumnKeyNumberIndex - 1,
		pKeyReturn=if tmpColumnReturnIndexNumber is text then List.PositionOf(Columns,tmpColumnReturnIndexNumber) else tmpColumnReturnIndexNumber - 1,

		// -------------------------------------
		// fixe columns in table lookup for
		// -------------------------------------
		Columns = Table.ColumnNames(pTableArray),

		// ------------------------------------
		// transform to table
		// ------------------------------------
		ColumnsTable = Table.FromList(
			Columns, Splitter.SplitByNothing(), null, null, ExtraValues.Error),

		// ---------------------------------------		
		// fixe index key in index column
		// ---------------------------------------
		ColumnNameMatch = Record.Field(
			ColumnsTable{pKeyIndex},"Column1"),

		//-----------------------------------------------------
		// fixe column return value
		//-----------------------------------------------------
		ColumnNameReturn = Record.Field(
			ColumnsTable{pKeyReturn},"Column1"),
		
		//-----------------------------------------------------
		//  sort ? yes ? no ?
		//-----------------------------------------------------
		SortTable = 
			if pSort = true 
			then Table.Sort(pTableArray,{{ColumnNameMatch, Order.Ascending}})
			else pTableArray,
		
		//----------------------------------------------------------------------
		// fixe temp column to get value with a secured name as __Lookup__
		//----------------------------------------------------------------------
		RenameTargetCol = Table.RenameColumns(
			SortTable,{{ColumnNameMatch, "__Lookup__"}}),

		//-----------------------------------------------------
		// get value from [__Lookup__] column
		//-----------------------------------------------------
		Lookup = Table.SelectRows(
			RenameTargetCol, each [__Lookup__] = pLookupValue),
		
		//-------------------------------------------------------------------------------
		// return value or not in this case msg error (customized or default #N/A)
		//-------------------------------------------------------------------------------
		ReturnValue=
			if Table.IsEmpty(Lookup)=true 
			then pMSG
			// Get first record mandatory 
			else Record.Field(Lookup{0},ColumnNameReturn)

	// end nested Let
	in ReturnValue

	// end function RechercheX as pq
in pqXLOOKUPV1