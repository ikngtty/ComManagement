using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using MicrosoftExcel = Microsoft.Office.Interop.Excel;

namespace Ikngtty.ComManagement
{
	/// <summary>
	/// A COM objects' life span manager.
	/// </summary>
	public class ComManager : IDisposable
	{
		private readonly Stack<object> coms;
		
		public ComManager()
		{
			// Initialize.
			this.coms = new Stack<object>();
		}
		
		public void Add(object com)
		{
			// Validate.
			if (com == null) throw new ArgumentNullException();
			if (!Marshal.IsComObject(com)) throw new ArithmeticException();
			
			// Add.
			this.coms.Push(com);
		}
		
		public void Dispose()
		{
			// Release COM objects.
			while (this.coms.Count > 0)
			{
				object com = this.coms.Pop();
				
				// Terminate in response to COM object's type.
				{
					var workbook = com as MicrosoftExcel.Workbook;
					if (workbook != null) workbook.Close(false);		// Don't save.
				}
				{
					var application = com as MicrosoftExcel.Application;
					if (application != null) application.Quit();
				}
				
				// Release.
				Marshal.ReleaseComObject(com);
			}
		}
	}
}
