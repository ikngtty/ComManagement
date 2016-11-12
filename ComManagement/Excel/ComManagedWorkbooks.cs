using System;
using Microsoft.Office.Interop.Excel;

namespace Ikngtty.ComManagement.Excel
{
	/// <summary>
	/// An excel workbooks COM object wrapper, which delegates it's life span management to the manager.
	/// </summary>
	public class ComManagedWorkbooks
	{
		private readonly ComManager comManager;
		private readonly Workbooks com;
		
		public ComManagedWorkbooks(ComManager comManager, Workbooks com)
		{
			// Validate.
			if (comManager == null) throw new ArgumentNullException();
			if (com == null) throw new ArgumentNullException();
			
			// Initialize.
			this.comManager = comManager;
			this.com = com;
			
			// Delegate COM object life span managament to the manager.
			this.comManager.Add(this.com);
		}
		
		public ComManagedWorkbook Add()
		{
			return new ComManagedWorkbook(this.comManager, this.com.Add());
		}
		
		public ComManagedWorkbook Open(string filename)
		{
			return new ComManagedWorkbook(this.comManager, this.com.Open(filename));
		}
	}
}
