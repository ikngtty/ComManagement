using System;
using Microsoft.Office.Interop.Excel;

namespace Ikngtty.ComManagement.Excel
{
	/// <summary>
	/// An excel workbook COM object wrapper, which delegates it's life span management to the manager.
	/// </summary>
	public class ComManagedWorkbook
	{
		private readonly ComManager comManager;
		private readonly Workbook com;
		
		public ComManagedWorkbook(ComManager comManager, Workbook com)
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
		
		private ComManagedSheets sheets;		// For cache.
		public ComManagedSheets Sheets
		{
			get
			{
				// Return the cache.
				if (this.sheets != null) return this.sheets;
				
				// Return a new sheets.
				var sheets = new ComManagedSheets(this.comManager, this.com.Sheets);
				this.sheets = sheets;			// Cache.
				return sheets;
			}
		}
		
		public void Close()
		{
			// Don't save.
			this.com.Close(false);
		}
		
		public void Save()
		{
			this.com.Save();
		}
		
		public void SaveAs(string filename)
		{
			this.com.SaveAs(filename);
		}
	}
}
