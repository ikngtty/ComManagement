using System;
using Microsoft.Office.Interop.Excel;

namespace Ikngtty.ComManagement.Excel
{
	/// <summary>
	/// An excel worksheet COM object wrapper, which delegates it's life span management to the manager.
	/// </summary>
	public class ComManagedWorksheet
	{
		private readonly ComManager comManager;
		private readonly Worksheet com;
		
		public ComManagedWorksheet(ComManager comManager, Worksheet com)
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
		
		private ComManagedRange cells;		// For cache.
		public ComManagedRange Cells
		{
			get
			{
				// Return the cache.
				if (this.cells != null) return this.cells;
				
				// Return a new sheets.
				var cells = new ComManagedRange(this.comManager, this.com.Cells);
				this.cells = cells;			// Cache.
				return cells;
			}
		}
	}
}
