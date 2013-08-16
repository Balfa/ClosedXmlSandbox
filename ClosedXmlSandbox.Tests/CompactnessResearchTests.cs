using NUnit.Framework;

namespace Mo.ClosedXmlSandbox.Tests
{
	[TestFixture]
	public class CompactnessResearchTests
	{
		[Test]
		public void UseStylePerCell()
		{
			CompactnessResearch.UseStylePerCell();
		}
		[Test]
		public void UseStylePerRange()
		{
			CompactnessResearch.UseStylePerRange();
		}
		[Test]
		public void StylePerCellRowBold()
		{
			CompactnessResearch.StylePerCellRowBold();
		}
		[Test]
		public void StylePerRangeRowBold()
		{
			CompactnessResearch.StylePerRangeRowBold();
		}
	}
}
