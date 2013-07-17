using NUnit.Framework;

namespace Mo.ClosedXmlSandbox.Tests
{
	[TestFixture]
	public class EarlyResearchTests
	{
		[Test]
		public void Closeit_HappyPath()
		{
			EarlyResearch.Fupp();
			EarlyResearch.Pantsit();
			EarlyResearch.Closeit();
			EarlyResearch.FromData();
		}
	}
}
