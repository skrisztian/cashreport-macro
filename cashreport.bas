REM  *****  BASIC  *****
REM
REM This set of macros prepare a Cash Book report based on a properly
REM filled and formatted table, and export it into pdf.
REM
REM Writen by Krisztian Stancz
REM Version: 2018-Mar-20-v3

Type DefaultData
    FirstYear as Long
	FirstMonth as Long
    CityTwoLetters as String
    CityName as String
    OrgName as String
    OrgAddress as String
    TaxNumber as String
End Type

Dim HufFormatId, HufFormatIdNoPoint As Long
Dim Defaults As New DefaultData

Sub CashBookSortData

	Dim Sheet As Object
	Dim ColumnLastCell, LastRow, LastColumn As Long
	Dim I, J As Long

	'Get active sheet
	Sheet = ThisComponent.GetCurrentController.ActiveSheet

	'Get the last cell's position in column "A"
	LastRow = GetCellCountOfColumn(0)-1
	
	'Get the last column
	LastColumn = GetLastColumn(Sheet)

	'Fill in the column after the last one with a header and functions
	'The functions result +1 for income and -1 for expense rows
	LastColumn = LastColumn + 1
	Sheet.GetCellByPosition(LastColumn, 0).String = "Tipus"
	For J = 1 To LastRow
		Sheet.GetCellByPosition(LastColumn, J).Formula = "=IF($C" & J+1 &">=$D" & J+1 & ";1;-1)"
	Next J

	'Select range, we skip the headers
	'From Col, From Row To Col, To Row
	Dim SortRange As Object
	SortRange = Sheet.getCellRangeByPosition(0, 1, LastColumn, LastRow)
	
	'Set up the sort parameters
	' Dátum - column 0 - ascending
	' Tipus - lastColumn - descending
	' Számla sorszám - column 1 - ascending
	Dim SortFields(2) As New com.sun.star.util.SortField
	SortFields(0).Field = 0
	SortFields(0).SortAscending = TRUE
	SortFields(1).Field = LastColumn
	SortFields(1).SortAscending = FALSE
	SortFields(2).Field = 1
	SortFields(2).SortAscending = TRUE
	
	'Set up sort descriptor
	Dim SortDesc(0) As New com.sun.star.beans.PropertyValue
	SortDesc(0).Name = "SortFields"
  	SortDesc(0).Value = SortFields()
  	
  	'Sort the range.
 	SortRange.Sort(SortDesc())

	'Delete the "Tipus column"
	Dim ColumnToDelete As Object
	ColumnToDelete = Sheet.getCellRangeByPosition(LastColumn, 0, LastColumn, 0).Columns
	ColumnToDelete.removeByIndex(0, 1)
	LastColumn = LastColumn-1
	
	'Fill in the rolling summary functions in column E
	For I = 1 To LastRow
		Sheet.GetCellByPosition(4, I).Formula = "=E" & I & "+C" & I+1 & "-D" & I+1
	Next I

End Sub

Sub CashBookCreateReport

	Dim Sheet, TitleRange, SignatureRange, oRow, oColumn, x, v, Doc As Object
	Dim LastColumn, ReportFirstClumn, ReportLastColumn As Long
	Dim Row, LastRow As Long
	Dim I, L As Long
	Dim ColumnNameIncome, ColumnNameExpense, DQuote, LBreak As String
	Dim NumberFormats As Object
	Dim HufFormatString, HufFormatStringNoPoint, DateFormatString As String
	Dim DateFormatId As Long

	DQuote = Chr$(34)&Chr$(34)
	LBreak = Chr$(10)

	' Sort data
	CashBookSortData

	'Get document and active sheet
	Doc = ThisComponent
	Sheet = ThisComponent.GetCurrentController.ActiveSheet

	'Set up currency and date formatting options
	Dim LocalSettings As New com.sun.star.lang.Locale
	
	LocalSettings.Language = "hu"
	LocalSettings.Country = "hu"
	
	NumberFormats = Doc.NumberFormats
	HufFormatString = "#.##0,00 Ft"
	HufFormatStringNoPoint = "##0,00 Ft"
	DateFormatString = "YYYY. MM. DD."
	
	HufFormatId = NumberFormats.queryKey(HufFormatString, LocalSettings, True)
	If HufFormatId = -1 Then
		HufFormatId = NumberFormats.addNew(HufFormatString, LocalSettings)
	End If
	
	HufFormatIdNoPoint = NumberFormats.queryKey(HufFormatStringNoPoint, LocalSettings, True)
	If HufFormatIdNoPoint = -1 Then
		HufFormatIdNoPoint = NumberFormats.addNew(HufFormatStringNoPoint, LocalSettings)
	End If
	
	DateFormatId = NumberFormats.queryKey(DateFormatString, LocalSettings, True)
	If DateFormatId = -1 Then
		DateFormatId = NumberFormats.addNew(DateFormatString, LocalSettings)
	End If
	
	'Get last row and column (of data)
	LastRow = GetCellCountOfColumn(0)-1
	LastColumn = GetLastColumn(Sheet)
	ReportFirstColumn = LastColumn + 2
	ReportLastColumn = ReportFirstColumn + 5

	'Create report title
	Sheet.getCellByPosition(ReportFirstColumn, 0).String = "Időszaki pénztárjelentés"
	TitleRange = Sheet.getCellRangeByPosition(ReportFirstColumn, 0, ReportLastColumn, 1)
	TitleRange.merge(TRUE)
	Titlerange.HoriJustify = com.sun.star.table.CellHoriJustify.CENTER
	TitleRange.VertJustify = com.sun.star.table.CellVertJustify.CENTER
	TitleRange.CharWeight = com.sun.star.awt.FontWeight.BOLD
	TitleRange.CellBackColor = RGB(220, 220, 220)
	SetBoxBorder(TitleRange, 8, FALSE)
	
	'Create report header
	Sheet.getCellByPosition(ReportFirstColumn, 3).String = "Sor-" &LBreak & "szám"
	Sheet.getCellByPosition(ReportFirstColumn + 1, 3).String = "Dátum"
	Sheet.getCellByPosition(ReportFirstColumn + 2, 3).String = "Bizonylatszám"
	Sheet.getCellByPosition(ReportFirstColumn + 3, 3).String = "Megnevezés"
	Sheet.getCellByPosition(ReportFirstColumn + 4, 3).String = "Bevétel"
	Sheet.getCellByPosition(ReportFirstColumn + 5, 3).String = "Kiadás"
	Sheet.Rows(3).OptimalHeight = True
	HeaderRange = Sheet.getCellRangeByPosition(ReportFirstColumn, 3, ReportLastColumn, 3)
	HeaderRange.HoriJustify = com.sun.star.table.CellHoriJustify.CENTER
	HeaderRange.VertJustify = com.sun.star.table.CellVertJustify.CENTER
	HeaderRange.CharWeight = com.sun.star.awt.FontWeight.BOLD
	
	'Print data into report
	For L = 0 To LastRow-1

		Row = L+4

		'Sorszám
		Sheet.GetCellByPosition(ReportFirstColumn, Row).Value = L+1
		Sheet.GetCellByPosition(ReportFirstColumn, Row).HoriJustify = com.sun.star.table.CellHoriJustify.CENTER

		'Dátum
		Sheet.GetCellByPosition(ReportFirstColumn+1, Row).Formula = "=A" & L+2
		Sheet.GetCellByPosition(ReportFirstColumn+1, Row).HoriJustify = com.sun.star.table.CellHoriJustify.RIGHT

		'Bizonylatszám
		Sheet.GetCellByPosition(ReportFirstColumn+2, Row).Formula = "=B" & L+2
		Sheet.GetCellByPosition(ReportFirstColumn+2, Row).NumberFormat = 100
		Sheet.GetCellByPosition(ReportFirstColumn+2, Row).HoriJustify = com.sun.star.table.CellHoriJustify.RIGHT
		
		'Megnevezés
		Sheet.GetCellByPosition(ReportFirstColumn+3, Row).Formula =  "=IF(F" & L+2 & "=" & DQuote & ";" & DQuote & ";F" & L+2 & ")"
		Sheet.GetCellByPosition(ReportFirstColumn+3, Row).NumberFormat = 100
		Sheet.GetCellByPosition(ReportFirstColumn+3, Row).HoriJustify = com.sun.star.table.CellHoriJustify.RIGHT

		'Bevétel
		Sheet.GetCellByPosition(ReportFirstColumn+4, Row).Formula = "=IF(C" & L+2 & "=" & DQuote & ";" & DQuote & ";C" & L+2 & ")"
		FormatHufValue(Sheet.GetCellByPosition(ReportFirstColumn+4, Row))

		'Kiadás
		Sheet.GetCellByPosition(ReportFirstColumn+5, Row).Formula = "=IF(D" & L+2 & "=" & DQuote & ";" & DQuote & ";D" & L+2 & ")"
		FormatHufValue(Sheet.GetCellByPosition(ReportFirstColumn+5, Row))

	Next L

	'Print summary section
	ColumnNameIncome = GetColumnName(Sheet.GetCellByPosition(ReportFirstColumn+4, 0))
	ColumnNameExpense = GetColumnName(Sheet.GetCellByPosition(ReportFirstColumn+5, 0))
	
	'Forgalom
	Sheet.GetCellByPosition(ReportFirstColumn+3, LastRow+4).String = "Forgalom"
	Sheet.GetCellByPosition(ReportFirstColumn+4, LastRow+4).Formula = "=SUM(" & ColumnNameIncome & 5 & ":" & ColumnNameIncome & LastRow+4 & ")"
	Sheet.GetCellByPosition(ReportFirstColumn+5, LastRow+4).Formula = "=SUM(" & ColumnNameExpense & 5 & ":" & ColumnNameExpense & LastRow+4 & ")"
	FormatHufValue(Sheet.GetCellByPosition(ReportFirstColumn+4, LastRow+4))
	FormatHufValue(Sheet.GetCellByPosition(ReportFirstColumn+5, LastRow+4))

	'Kezdő pénzkészlet
	Sheet.GetCellByPosition(ReportFirstColumn+3, LastRow+5).String = "Kezdő pénzkészlet"
	Sheet.GetCellByPosition(ReportFirstColumn+4, LastRow+5).Formula = "=E1"
	FormatHufValue(Sheet.GetCellByPosition(ReportFirstColumn+4, LastRow+5))
	Sheet.GetCellByPosition(ReportFirstColumn+5, LastRow+5).CellBackColor = RGB(210, 210, 210)
	
	'Záró pénzkészlet
	Sheet.GetCellByPosition(ReportFirstColumn+3, LastRow+6).String = "Záró pénzkészlet"
	Sheet.GetCellByPosition(ReportFirstColumn+5, LastRow+6).Formula = "=E" & LastRow+1
	FormatHufValue(Sheet.GetCellByPosition(ReportFirstColumn+5, LastRow+6))
	Sheet.GetCellByPosition(ReportFirstColumn+4, LastRow+6).CellBackColor = RGB(210, 210, 210)

	'Összesen
	Sheet.GetCellByPosition(ReportFirstColumn+3, LastRow+7).String = "Összesen"
	Sheet.GetCellByPosition(ReportFirstColumn+4, LastRow+7).Formula = "=SUM(" & ColumnNameIncome & LastRow+5 & ":" & ColumnNameIncome & LastRow+7 & ")"
	FormatHufValue(Sheet.GetCellByPosition(ReportFirstColumn+4, LastRow+7))
	Sheet.GetCellByPosition(ReportFirstColumn+4, LastRow+7).CharWeight = com.sun.star.awt.FontWeight.BOLD
	Sheet.GetCellByPosition(ReportFirstColumn+5, LastRow+7).Formula = "=SUM(" & ColumnNameExpense & LastRow+5 & ":" & ColumnNameExpense & LastRow+7 & ")"
	FormatHufValue(Sheet.GetCellByPosition(ReportFirstColumn+5, LastRow+7))
	Sheet.GetCellByPosition(ReportFirstColumn+5, LastRow+7).CharWeight = com.sun.star.awt.FontWeight.BOLD
	
	Sheet.GetCellByPosition(ReportFirstColumn+4, LastRow+8).String = "Bevétel"
	Sheet.GetCellByPosition(ReportFirstColumn+5, LastRow+8).String = "Kiadás"
	Sheet.GetCellByPosition(ReportFirstColumn+4, LastRow+8).HoriJustify = com.sun.star.table.CellHoriJustify.CENTER
	Sheet.GetCellByPosition(ReportFirstColumn+5, LastRow+8).HoriJustify = com.sun.star.table.CellHoriJustify.CENTER
	
	'Change date format
	oRange = Sheet.getCellRangeByPosition(ReportFirstColumn+1, 4, ReportFirstColumn+1, LastRow+3)
	oRange.NumberFormat = DateFormatId
	
	'Set report borders and grids
	For I = ReportFirstColumn To ReportLastColumn
		If I > ReportFirstColumn+2 Then
			L = 5
		Else
			L = 0
		End If
		oRange = Sheet.getCellRangeByPosition(I, 3, I, LastRow+3+L)
		oRange.ParaRightMargin = 100
		oRange.ParaLeftMargin = 100
		SetBoxBorder(oRange, 8, TRUE)
	Next I
	
	'Set report columns to auto adjust width
	For I = ReportFirstColumn To ReportLastColumn
		oColumn = Sheet.Columns(I)
  		oColumn.OptimalWidth = True
  	Next I
	
	'Create signature place
	Sheet.GetCellByPosition(ReportFirstColumn, LastRow+15).String = "pénztáros"
	SignatureRange = Sheet.getCellRangeByPosition(ReportFirstColumn, LastRow+15, ReportFirstColumn+2, LastRow+15)
	SignatureRange.merge(TRUE)
	SignatureRange.HoriJustify = com.sun.star.table.CellHoriJustify.CENTER
	SignatureRange.VertJustify = com.sun.star.table.CellVertJustify.CENTER

	v = SignatureRange.TableBorder
	x = v.TopLine : x.OuterLineWidth = 2 : v.TopLine = x
	x = v.TopLine : x.InnerLineWidth = 0 : v.TopLine = x
	x = v.TopLine : x.LineDistance = 0 : v.TopLine = x	
	SignatureRange.TableBorder = v
	
End Sub

Sub PrintReport

	Dim Sheet, Cell, ReportRange, Doc As Object
	Dim StyleFamilies, PageStyles, DefPage, HText As Object
	Dim LBreak As String

	LBreak = Chr$(10)
	
	' Read default data into global var
	GetDefaults

	'Get active sheet
	Doc = ThisComponent
	Sheet = Doc.GetCurrentController.ActiveSheet
	
	'Get report range
	RiportRange = GetReportRange(Sheet)
	
	'Setup page params
	StyleFamilies = Doc.StyleFamilies
	PageStyles = StyleFamilies.getByName("PageStyles")
	DefPage = PageStyles.getByName("Default")
	DefPage.IsLandscape = FALSE
	DefPage.Width = 21000
	DefPage.Height = 29700
	DefPage.LeftMargin = 1000
	DefPage.RightMargin = 1000
	DefPage.TopMargin = 1000
	DefPage.BottomMargin = 1000
	DefPage.CenterHorizontally = TRUE
	defPage.PrintHeaders = 0
	
	'Setup header
	DefPage.HeaderIsOn = TRUE
	HContent = DefPage.RightPageHeaderContent
	HText = HContent.LeftText
	HText.String = Defaults.OrgName & LBreak & Defaults.OrgAddress & LBreak &_
	               "Adószám: " & Defaults.TaxNumber
	DefPage.RightPageHeaderContent = HContent
	HText = HContent.CenterText
	HText.String = ""
	DefPage.RightPageHeaderContent = HContent
	HText = HContent.RightText
	HText.String = "Sorszám: " & GetReportId(Sheet.GetCellByPosition(0, 1)) & LBreak &_
	               "Időszak: " & GetReportPeriod(Sheet.GetCellByPosition(0, 1))
	DefPage.RightPageHeaderContent = HContent
	
	'Setup footer
	'Dim PageNum As New com.sun.star.text.TextField.PageNumber
	
	DefPage.FooterIsOn = TRUE
	HContent = DefPage.RightPageFooterContent
	HText = HContent.LeftText
	HText.String = "Időszaki pénztárjelentés (" & Defaults.CityName & ")"
	DefPage.RightPageFooterContent = HContent
	HText = HContent.CenterText
	HText.String = ""
	DefPage.RightPageFooterContent = HContent
	'HText = HContent.RightText
	'HText.String = CStr(Doc.CurrentController.PageCount) & " / " & PageNum
	'DefPage.RightPageFooterContent = HContent
	
	Dim dispatcher As Object
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
  
	'Get Print Range - if no print range has been defined for the Active Sheet, then it will export the entire Active Sheet
	Dim aFilterData(0) As New com.sun.star.beans.PropertyValue
	aFilterData(0).Name = "Selection"
	aFilterData(0).Value = GetReportRange 'Sheet

	'Export to PDF
    Dim args1(1) As New com.sun.star.beans.PropertyValue
    Dim document As Object
    document = ThisComponent.CurrentController.Frame
	args1(0).Name = "FilterName"
    args1(0).Value = "calc_pdf_Export"
    args1(1).Name = "FilterData"
    args1(1).Value = aFilterData() 'GetReportRange 
    dispatcher.executeDispatch(document, ".uno:ExportDirectToPDF", "", 0, args1())

	' To export with fixed automatic name:
	' Dim oDoc As Object
	' Dim PdfURL as String -> new file name with full path
	' oDoc = ThisComponent
	' oDoc.getURL() -> current file name with full path
	' oDoc.storeToURL(PdfURL,args())
    
End Sub

REM ***** Helper functions *****

Sub GetDefaults

	'First year of report
	'A 0001 számú jelentés éve, 4 számjegy
    Defaults.FirstYear = 2018

	'First month of report
	'A 0001 számú jelentés hónapja 1 vagy 2 számjegy (január = 1, december = 12)
	Defaults.FirstMonth = 2
	
	' City long name
	' A város teljes neve, idézőjelekkel
    Defaults.CityName = "Budapest"

	' City two letter code
	' A város kétbetűs rövidítése, ami a sorszámba kerül
	' Két nagy betű, idézőjelekkel
    Defaults.CityTwoLetters = "BP"
    
    ' Organization name
    ' A társaság neve
    Defaults.OrgName = "Magyarországi Taoista Tai Chi Társaság"
    
    ' Organization address
    ' A társaság központi hivatalos címe
    Defaults.OrgAddress = "1053 Budapest, Kossuth L. u. 18. II./3."
    
    ' Organizations tax number
    ' A társaság adószáma
    Defaults.TaxNumber = "18070022-1-41"

End Sub

Function GetCellCountOfColumn(Column As Long) As Long

	Dim Sheet, CellRange As Object
	Dim CellCount As Long

	'Select column as range
	Sheet = ThisComponent.GetCurrentController.ActiveSheet
	CellRange = Sheet.getCellRangeByPosition(Column, 0, Column, 10000)

	'Count the number of cells in the range
	GetCellCountOfColumn = CellRange.computeFunction(com.sun.star.sheet.GeneralFunction.COUNT)
	
End Function

Function GetLastColumn(Sheet As Object) As Long
'Returns the last used column number as a zero based position, i.e GetNameByPosition value

	Dim Cell, Cursor As Object
	
	Cell = Sheet.GetCellByPosition(0, 0)
	Cursor = Sheet.createCursorByRange(Cell)
	Cursor.GotoEndOfUsedArea(FALSE)
	GetLastColumn = Cursor.RangeAddress.EndColumn

End Function

Function GetColumnName(Cell As Object) As String

	'Absolute cell name e.g.: $Sheet1.$A$1
	aCellName = Split(Cell.AbsoluteName, "$")
	GetColumnName = aCellName(UBound(aCellName)-1)

End Function

Sub SetBoxBorder(Range As Object, LineWidth As Long, SetInnerBorders As Boolean)

	Dim x, v As Object
	
	v = Range.TableBorder
	If SetInnerBorders = TRUE Then
		x = v.HorizontalLine : x.OuterLineWidth = 2 : v.HorizontalLine = x
		x = v.HorizontalLine : x.InnerLineWidth = 0 : v.HorizontalLine = x
		x = v.HorizontalLine : x.LineDistance = 0 : v.HorizontalLine = x	
	End If
	x = v.TopLine : x.OuterLineWidth = LineWidth : v.TopLine = x
	x = v.TopLine : x.InnerLineWidth = 0 : v.TopLine = x
	x = v.TopLine : x.LineDistance = 0 : v.TopLine = x	
	x = v.LeftLine : x.OuterLineWidth = LineWidth : v.LeftLine = x
	x = v.LeftLine : x.InnerLineWidth = 0 : v.LeftLine = x
	x = v.LeftLine : x.LineDistance = 0 : v.LeftLine = x	
	x = v.RightLine : x.OuterLineWidth = LineWidth : v.RightLine = x
	x = v.RightLine : x.InnerLineWidth = 0 : v.RightLine = x
	x = v.RightLine : x.LineDistance = 0 : v.RightLine = x
	x = v.BottomLine : x.OuterLineWidth = LineWidth : v.BottomLine = x
	x = v.BottomLine : x.InnerLineWidth = 0 : v.BottomLine = x
	x = v.BottomLine : x.LineDistance = 0 : v.BottomLine = x
	Range.TableBorder = v

End Sub

Sub FormatHufValue(Cell As Object)

	If Cell.Value < 1000 Then
		Cell.NumberFormat = HufFormatIdNoPoint
	Else
		Cell.NumberFormat = HufFormatId
	End If
	
End Sub

Function GetReportRange As Object

	Dim Cell, Search, ReportRange, ColumnRange, Sheet As Object
	Dim LastCell As Long

	Sheet = ThisComponent.GetCurrentController.ActiveSheet
	Search = Sheet.createSearchDescriptor()
	Search.SearchString = "Időszaki pénztárjelentés"
	Cell = Sheet.findFirst(Search)
	ColumnRange = Sheet.getCellRangeByPosition(Cell.CellAddress.Column, 3, Cell.CellAddress.Column, 10000)
	LastCell = ColumnRange.computeFunction(com.sun.star.sheet.GeneralFunction.COUNT) + 14
	GetReportRange = Sheet.GetCellRangeByPosition(Cell.CellAddress.Column, 0, Cell.CellAddress.Column+5, LastCell)

End Function

Function GetReportPeriod(Cell As Object) As String

	Dim Period As Date

	Period = Cell.Value
	GetReportPeriod = Format(Period, "yyyy. MMMM")

End Function

Function GetReportId(Cell As Object) As String

	Dim Period As Date
	Dim Id As Long
	Dim IdString As String

	' Read default data into global var
	GetDefaults
	
	Period = Cell.Value
	IdNum = (Year(Period) - Defaults.FirstYear) * 12 + Month(Period) - (Defaults.FirstMonth - 1)

	Select Case IdNum
	  Case 0 To 9                   
	    IdString = "000" & CStr(IdNum)
	  Case 10 To 99
	    IdString = "00" & CStr(IdNum)
	  Case 100 To 999                  
	    IdString = "0" & CStr(IdNum)
	  Case Else
	    IdString = CStr(IdNum)
	End Select

	GetReportId = "HP" & Defaults.CityTwoLetters & IdString

End Function

REM ***** Debug functions *****

Sub PrintProperties

	Dim Sheet, Cell As Object

	'Get active sheet
	Sheet = ThisComponent.GetCurrentController.ActiveSheet
	Cell = Sheet.getCellByPosition(0, 0).CellAddress
	MsgBox ThisComponent.dbg_properties

End Sub

Sub PrintMethods

	Dim Sheet, Cell As Object

	'Get active sheet
	Sheet = ThisComponent.GetCurrentController.ActiveSheet
	Cell = Sheet.GetCellByPosition(0, 0)
	MsgBox Cell.dbg_methods

End Sub
