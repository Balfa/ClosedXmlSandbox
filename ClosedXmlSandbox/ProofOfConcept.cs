using System.Reflection;
using ClosedXML.Excel;

namespace Mo.ClosedXmlSandbox
{
	public class ProofOfConcept
	{
		static string[][] data = new[]
		{
			new[] {"Food", "Color", "Oz."},
			new[] {"Fruit"},
			new[] {"Banana", "Yellow", "3"},
			new[] {"Apple", "Green", "2"},
			new[] {"Raspberry", "Pink", "0.1"},
			new[] {"Grain"},
			new[] {"Bread", "Brown", "1"},
			new[] {"Dairy"},
			new[] {"Milk", "White", "133"},
			new[] {"Cheese", "Yellow", "4"}
		};
		public static void Basic()
		{
			var wb = new XLWorkbook();
			var ws = wb.AddWorksheet("Worksheet Name");
			ws.Cell("A1").Value = data;
			wb.SaveAs(MethodBase.GetCurrentMethod().Name + ".xlsx");
		}
		public static void FontFaceStyleWeightColorSize()
		{
			var wb = new XLWorkbook();
			var ws = wb.AddWorksheet("Worksheet Name");
			ws.Cell("A1").Value = data;

			var subHeaders = ws.Ranges("A2:C2,A6:C6,A8:C8");

			subHeaders.Style
				.Font.SetFontName("Century Gothic")
				.Font.SetUnderline().Font.SetStrikethrough()
				.Font.SetBold()
				.Font.SetFontColor(XLColor.Amethyst)
				.Font.SetFontSize(18);

			wb.SaveAs(MethodBase.GetCurrentMethod().Name + ".xlsx");
		}
		public static void BorderStyleColor()
		{
			var wb = new XLWorkbook();
			var ws = wb.AddWorksheet("Worksheet Name");
			ws.Cell("A1").Value = data;

			var subHeaders = ws.Ranges("A2:C2,A6:C6,A8:C8");

			subHeaders.Style
				.Border.SetBottomBorder(XLBorderStyleValues.Double)
				.Border.SetBottomBorderColor(XLColor.Aquamarine);

			wb.SaveAs(MethodBase.GetCurrentMethod().Name + ".xlsx");
		}
		public static void BackgroundColor()
		{
			var wb = new XLWorkbook();
			var ws = wb.AddWorksheet("Worksheet Name");
			ws.Cell("A1").Value = data;

			var subHeaders = ws.Ranges("A2:C2,A6:C6,A8:C8");

			subHeaders.Style
				.Fill.SetBackgroundColor(XLColor.Almond);

			wb.SaveAs(MethodBase.GetCurrentMethod().Name + ".xlsx");
		}
		public static void WordWrap()
		{
			var wb = new XLWorkbook();
			var ws = wb.AddWorksheet("Worksheet Name");
			ws.Cell("A1").Value = data;

			ws.Cell("A2").Value = "Extraordinarily long string that deserves a good wrapping!";
			ws.Cell("A2").Style.Alignment.SetWrapText();
			ws.Cell("A6").Value = "Extraordinarily long string that shan't get any wrapping today.";

			wb.SaveAs(MethodBase.GetCurrentMethod().Name + ".xlsx");
		}
		public static void RowHeight()
		{
			var wb = new XLWorkbook();
			var ws = wb.AddWorksheet("Worksheet Name");
			ws.Cell("A1").Value = data;

			var rows = ws.Rows("2, 6, 8");
			rows.Height = 35;
			
			wb.SaveAs(MethodBase.GetCurrentMethod().Name + ".xlsx");
		}
		public static void VerticalAlign()
		{
			var wb = new XLWorkbook();
			var ws = wb.AddWorksheet("Worksheet Name");
			ws.Cell("A1").Value = data;

			var rows = ws.Rows("2, 6, 8");
			rows.Height = 35;

			var subHeaders = ws.Ranges("A2:C2");
			subHeaders.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Top);
			
			wb.SaveAs(MethodBase.GetCurrentMethod().Name + ".xlsx");
		}
		public static void ColWidth()
		{
			var wb = new XLWorkbook();
			var ws = wb.AddWorksheet("Worksheet Name");
			ws.Cell("A1").Value = data;

			var cols = ws.Columns(1,3);
			cols.Width = 44;
			
			wb.SaveAs(MethodBase.GetCurrentMethod().Name + ".xlsx");
		}
		public static void HorizontalAlign()
		{
			var wb = new XLWorkbook();
			var ws = wb.AddWorksheet("Worksheet Name");
			ws.Cell("A1").Value = data;

			var cols = ws.Columns(1, 3);
			cols.Width = 44;

			ws.Cell("B4").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
			
			wb.SaveAs(MethodBase.GetCurrentMethod().Name + ".xlsx");
		}
		public static void HorizontalMerge()
		{
			var wb = new XLWorkbook();
			var ws = wb.AddWorksheet("Worksheet Name");
			ws.Cell("A1").Value = data;

			var subHeaders = ws.Ranges("A2:C2,A6:C6,A8:C8");

			subHeaders.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
			subHeaders.ForEach(x => x.Merge());

			wb.SaveAs(MethodBase.GetCurrentMethod().Name + ".xlsx");
		}
		public static void VerticalMerge()
		{
			var wb = new XLWorkbook();
			var ws = wb.AddWorksheet("Worksheet Name");
			ws.Cell("A1").Value = data;

			var range = ws.Range("B1:B4");

			range.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
			range.Merge();

			wb.SaveAs(MethodBase.GetCurrentMethod().Name + ".xlsx");
		}
		public static void Annotation()
		{
			var wb = new XLWorkbook();
			var ws = wb.AddWorksheet("Worksheet Name");
			ws.Name = "JOEY";
			ws.Cell("A1").Value = data;
			ws.Cell("B4").Comment.SetAuthor("Batman")
			  .AddSignature()
			  .AddText("Zowie! An annotation!")
			  .AddNewLine()
			  .AddNewLine()
			  .AddText("And it even supports new lines!");
			wb.SaveAs(MethodBase.GetCurrentMethod().Name + ".xlsx");
		}
		public static void DataTypes()
		{
			var wb = new XLWorkbook();
			var ws = wb.AddWorksheet("Worksheet Name");
			ws.Cell("A1").Value = data;

			var nums = ws.Range("C2:C10");
			nums.SetDataType(XLCellValues.Number);
			wb.SaveAs(MethodBase.GetCurrentMethod().Name + ".xlsx");
		}
		public static void NumberFormats()
		{
			var wb = new XLWorkbook();
			var ws = wb.AddWorksheet("Worksheet Name");
			ws.Cell("A1").Value = data;

			var nums = ws.Range("C2:C10");
			nums.SetDataType(XLCellValues.Number);
			nums.Style.NumberFormat.SetFormat("#0.0");
			wb.SaveAs(MethodBase.GetCurrentMethod().Name + ".xlsx");
		}
		public static void MultipleWorksheetsAndWorksheetName()
		{
			var wb = new XLWorkbook();
			var ws = wb.AddWorksheet("Worksheet1");
			ws.Cell("A1").Value = data;
			var ws2 = wb.AddWorksheet("Worksheet2");
			ws2.Cell("D4").Value = "Hi, there";
			wb.SaveAs(MethodBase.GetCurrentMethod().Name + ".xlsx");
		}
		public static void FreezeRow()
		{
			var wb = new XLWorkbook();
			var ws = wb.AddWorksheet("Worksheet Name");
			ws.Cell("A1").Value = data;

			ws.SheetView.FreezeRows(1);
			wb.SaveAs(MethodBase.GetCurrentMethod().Name + ".xlsx");
		}
		public static void FreezeColumn()
		{
			var wb = new XLWorkbook();
			var ws = wb.AddWorksheet("Worksheet Name");
			ws.Cell("A1").Value = data;

			ws.SheetView.FreezeColumns(1);
			wb.SaveAs(MethodBase.GetCurrentMethod().Name + ".xlsx");
		}
		public static void FreezeBoth()
		{
			var wb = new XLWorkbook();
			var ws = wb.AddWorksheet("Worksheet Name");
			ws.Cell("A1").Value = data;

			ws.SheetView.Freeze(1,1);
			wb.SaveAs(MethodBase.GetCurrentMethod().Name + ".xlsx");
		}
	}
}
