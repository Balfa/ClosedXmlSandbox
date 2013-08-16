using System.Reflection;
using ClosedXML.Excel;

namespace Mo.ClosedXmlSandbox
{
	public class CompactnessResearch
	{
		static readonly string[][] Data = new[]
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
		// Oh, that's pretty epic. UseStylePerCell() & UseStylePerRange() produce practically identical xlsx files!
		public static void UseStylePerCell()
		{
			var wb = new XLWorkbook();
			var ws = wb.AddWorksheet("Worksheet Name");
			ws.Cell("A1").Value = Data;
			// loop through cells in first col and make each one SAME colour
			for (int i = 0; i < 10; i++)
			{
				int rowNum = i + 1; // Rows are 1-based
				ws.Cell("A" + rowNum).Style.Fill.SetBackgroundColor(XLColor.FromHtml("#FF9"));
			}

			wb.SaveAs(MethodBase.GetCurrentMethod().Name + ".xlsx");
		}
		public static void UseStylePerRange()
		{
			var wb = new XLWorkbook();
			var ws = wb.AddWorksheet("Worksheet Name");
			ws.Cell("A1").Value = Data;
			// loop through cells in first col and make each one SAME colour
			var range = ws.Range("A1:A10");
			range.Style.Fill.SetBackgroundColor(XLColor.FromHtml("#FF9"));

			wb.SaveAs(MethodBase.GetCurrentMethod().Name + ".xlsx");
		}
		// As do StylePerRangeRowBold() & StylePerCellRowBold()
		public static void StylePerCellRowBold()
		{
			var wb = new XLWorkbook();
			var ws = wb.AddWorksheet("Worksheet Name");
			ws.Cell("A1").Value = Data;
			// loop through cells in first col and make each one SAME colour
			for (int i = 0; i < 10; i++)
			{
				int rowNum = i + 1; // Rows are 1-based
				ws.Cell("A" + rowNum).Style.Fill.SetBackgroundColor(XLColor.FromHtml("#FF9"));
			}
			// Bold a row
			var boldRow = ws.Range("A3:C3");
			boldRow.Style.Font.SetBold();

			wb.SaveAs(MethodBase.GetCurrentMethod().Name + ".xlsx");
		}
		public static void StylePerRangeRowBold()
		{
			var wb = new XLWorkbook();
			var ws = wb.AddWorksheet("Worksheet Name");
			ws.Cell("A1").Value = Data;
			// loop through cells in first col and make each one SAME colour
			var range = ws.Range("A1:A10");
			range.Style.Fill.SetBackgroundColor(XLColor.FromHtml("#FF9"));
			// Bold a row
			var boldRow = ws.Range("A3:C3");
			boldRow.Style.Font.SetBold();

			wb.SaveAs(MethodBase.GetCurrentMethod().Name + ".xlsx");
		}
	}
}
