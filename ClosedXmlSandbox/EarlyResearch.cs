using System;
using ClosedXML.Excel;

namespace Mo.ClosedXmlSandbox
{
	public class EarlyResearch
	{
		public static void Pantsit()
		{
			var wb = new XLWorkbook("BasicTable.xlsx");
			var ws = wb.Worksheet(1);
			var firstCell = ws.FirstCellUsed();
			var lastCell = ws.LastCellUsed();
			var range = ws.Range(firstCell.Address, lastCell.Address);
			range.Row(1).Delete(); // Deleting the "Contacts" header (we don't need it for our purposes)

			// We want to use a theme for table, not the hard coded format of the BasicTable
			range.Clear(XLClearOptions.Formats);
			// Put back the date and number formats
			range.Column(4).Style.NumberFormat.NumberFormatId = 15;
			range.Column(5).Style.NumberFormat.Format = "$ #,##0";

			//var table = range.CreateTable();    // You can also use range.AsTable() if you want to
			// manipulate the range as a table but don't want 
			// to create the table in the worksheet.

			wb.SaveAs("UsingTables.xlsx");
		}
		public static void Closeit()
		{
			var workbook = new XLWorkbook();
			var worksheet = workbook.Worksheets.Add("Rippit");
			worksheet.Cell("B5").Value = "Snart that!";
			//worksheet.Column("A")
			workbook.SaveAs("Yosup.xlsx");
		}
		public static void FromData()
		{
			// Creating a new workbook
			var wb = new XLWorkbook();


			// Adding a worksheet
			var ws = wb.Worksheets.Add("Contacts");

			var data = new[]
				{
					new[] {"Food", "Color", "Size"},
					new[] {"Fruit"},
					new[] {"Banana", "Yellow", "Med"},
					new[] {"Apple", "Green", "Med"},
					new[] {"Raspberry", "Pink", "Small"},
					new[] {"Grain"},
					new[] {"Bread", "Brown", "Med"},
					new[] {"Dairy"},
					new[] {"Milk", "White", "Large"},
					new[] {"Cheese", "Yellow", "Med"}
				};

			ws.Cell("A1").Value = data;


			var colHeaders = ws.Range("A1:C1");

			//var subHeaders = ws.Ranges("A2:C2,A6:C6,A8:C8");
			//subHeaders.ForEach(x => x.Merge());
			//subHeaders.Style.Fill.BackgroundColor = XLColor.Amber;
			//subHeaders.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

			var poo = ws.FirstColumnUsed().ColumnNumber().ToString();

			ws.SheetView.FreezeRows(1);
			var table = ws.RangeUsed().CreateTable();
			table.Theme = XLTableTheme.TableStyleMedium2;
			table.ShowAutoFilter = false;
			var backToRange = table.AsRange();

			var subHeaders = backToRange.Ranges("A2:C2,A6:C6,A8:C8");
			//subHeaders.ForEach(x => x.Merge());
			subHeaders.Style.Font.SetBold();

			ws.Clear();
			ws.Cell("A1").Value = backToRange;
			
			wb.SaveAs("FromData.xlsx");
		}
		public static void Fupp()
		{
			// Creating a new workbook
			var wb = new XLWorkbook();


			// Adding a worksheet
			var ws = wb.Worksheets.Add("Contacts");

			// Title
			ws.Cell("B2").Value = "Contacts";

			// First Names
			ws.Cell("B3").Value = "FName";
			ws.Cell("B4").Value = "John";
			ws.Cell("B5").Value = "Hank";
			ws.Cell("B6").Value = "Dagny";

			// Last Names
			ws.Cell("C3").Value = "LName";
			ws.Cell("C4").Value = "Galt";
			ws.Cell("C5").Value = "Rearden";
			ws.Cell("C6").Value = "Taggart";


			// Adding more data types
			// Boolean
			ws.Cell("D3").Value = "Outcast";
			ws.Cell("D4").Value = true;
			ws.Cell("D5").Value = false;
			ws.Cell("D6").Value = false;

			// DateTime
			ws.Cell("E3").Value = "DOB";
			ws.Cell("E4").Value = new DateTime(1919, 1, 21);
			ws.Cell("E5").Value = new DateTime(1907, 3, 4);
			ws.Cell("E6").Value = new DateTime(1921, 12, 15);

			// Numeric
			ws.Cell("F3").Value = "Income";
			ws.Cell("F4").Value = 2000;
			ws.Cell("F5").Value = 40000;
			ws.Cell("F6").Value = 10000;

			// Defining ranges
			// From worksheet
			var rngTable = ws.Range("B2:F6");

			// From another range
			var rngDates = rngTable.Range("D3:D5"); // The address is relative to rngTable (NOT the worksheet)
			var rngNumbers = rngTable.Range("E3:E5"); // The address is relative to rngTable (NOT the worksheet)


			// Formatting dates and numbers
			// Using OpenXML's predefined formats
			rngDates.Style.NumberFormat.NumberFormatId = 15;
			IXLStyle poo = rngTable.Cells().Style;
			// Using a custom format
			rngNumbers.Style.NumberFormat.Format = "$ #,##0";
			

			// Formatting headers
			// Formatting headers
			var rngHeaders = rngTable.Range("A2:E2"); // The address is relative to rngTable (NOT the worksheet)
			rngHeaders.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
			rngHeaders.Style.Font.Bold = true;
			rngHeaders.Style.Fill.BackgroundColor = XLColor.Aqua;


			// Adding grid lines
			// Adding grid lines
			rngTable.Style.Border.BottomBorder = XLBorderStyleValues.Thin;


			// Format title cell
			// Format title cell
			rngTable.Cell(1, 1).Style.Font.Bold = true;
			rngTable.Cell(1, 1).Style.Fill.BackgroundColor = XLColor.CornflowerBlue;
			rngTable.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;


			rngTable.Row(1).Merge(); // We could've also used: rngTable.Range("A1:E1").Merge()


			// Add thick borders
			// Add a thick outside border
			rngTable.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;

			// You can also specify the border for each side with:
			// rngTable.FirstColumn().Style.Border.LeftBorder = XLBorderStyleValues.Thick;
			// rngTable.LastColumn().Style.Border.RightBorder = XLBorderStyleValues.Thick;
			// rngTable.FirstRow().Style.Border.TopBorder = XLBorderStyleValues.Thick;
			// rngTable.LastRow().Style.Border.BottomBorder = XLBorderStyleValues.Thick;


			// Adjust column widths to their content
			ws.Columns(2, 6).AdjustToContents();


			// Saving the workbook
			wb.SaveAs("BasicTable.xlsx");
		}
	}
}
