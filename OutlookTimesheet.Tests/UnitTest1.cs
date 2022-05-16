using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace OutlookTimesheet.Tests
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            List<string> values = new List<string>(OutlookTimesheet.Process.TimesheetProcess.GetAllCalendarItems());

            Assert.IsTrue(true);
        }
    }
}