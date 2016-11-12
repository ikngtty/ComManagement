using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using Ikngtty.ComManagement;
using Ikngtty.ComManagement.Excel;
using NUnit.Framework;

namespace Ikngtty.ComManagementTest
{
	[TestFixture]
	public class RoughTest
	{
		[Test]
		public void Scenario1()
		{
			// Get a test file path.
			const string relativeExcelFilePath = @"../../TestFile/RoughTest/ForTest.xlsx";
			string absoluteExcelFilePath = Path.GetFullPath(relativeExcelFilePath);
			
			// Get an excel process count.
			int processCount = Process.GetProcessesByName("Excel").Length;
			
			// Get a cell value.
			using (var comManager = new ComManager())
			{
				var excel = new ComManagedApplication(comManager);
				
				// The process is created.
				Assert.AreEqual(processCount + 1, Process.GetProcessesByName("Excel").Length);
				
				ComManagedWorkbook book = excel.Workbooks.Open(absoluteExcelFilePath);
				var cellValue = (string)book.Sheets[1].Cells[1, 1].Value;
				
				Assert.AreEqual("foo1", cellValue);
			}
			
			// The process is killed.
			Thread.Sleep(3000);
			Assert.AreEqual(processCount, Process.GetProcessesByName("Excel").Length);
		}
	}
}
