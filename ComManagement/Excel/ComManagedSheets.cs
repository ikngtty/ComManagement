using System;
using Microsoft.Office.Interop.Excel;

namespace Ikngtty.ComManagement.Excel
{
	/// <summary>
	/// An excel sheets COM object wrapper, which delegates it's life span management to the manager.
	/// </summary>
	public class ComManagedSheets
	{
		private readonly ComManager comManager;
		private readonly Sheets com;
		private readonly ComManagedWorksheet[] sheets;		// For cache.
		
		public ComManagedSheets(ComManager comManager, Sheets com)
		{
			// Validate.
			if (comManager == null) throw new ArgumentNullException();
			if (com == null) throw new ArgumentNullException();
			
			// Initialize.
			this.comManager = comManager;
			this.com = com;
			this.sheets = new ComManagedWorksheet[this.com.Count + 1];
			
			// Delegate COM object life span managament to the manager.
			this.comManager.Add(this.com);
		}
		
		public ComManagedWorksheet this[int index]
		{
			get
			{
				// Validate.
				if (index < 1 || this.com.Count < index) throw new ArgumentOutOfRangeException();
				
				// Return the cache.
				if (this.sheets[index] != null) return this.sheets[index];
				
				// Return a new worksheet.
				var sheet = new ComManagedWorksheet(this.comManager, (Worksheet)this.com[index]);
				this.sheets[index] = sheet;			// Cache.
				return sheet;
			}
		}
	}
}
