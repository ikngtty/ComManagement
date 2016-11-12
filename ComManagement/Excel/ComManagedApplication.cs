using System;
using Microsoft.Office.Interop.Excel;

namespace Ikngtty.ComManagement.Excel
{
	/// <summary>
	/// An excel application COM object wrapper, which delegates it's life span management to the manager.
	/// </summary>
	public class ComManagedApplication
	{
		private readonly ComManager comManager;
		private readonly Application com;
		
		public ComManagedApplication(ComManager comManager)
		{
			// Validate.
			if (comManager == null) throw new ArgumentNullException();
			
			// Initialize.
			this.comManager = comManager;
			this.com = new Application();
			
			// Delegate COM object life span managament to the manager.
			this.comManager.Add(this.com);
		}
		
		private ComManagedWorkbooks workbooks;		// For cache.
		public ComManagedWorkbooks Workbooks
		{
			get
			{
				// Return the cache.
				if (this.workbooks != null) return this.workbooks;
				
				// Return a new workbooks.
				var workbooks = new ComManagedWorkbooks(this.comManager, this.com.Workbooks);
				this.workbooks = workbooks;			// Cache.
				return workbooks;
			}
		}
		
		public void Quit()
		{
			this.com.Quit();
		}
	}
}
