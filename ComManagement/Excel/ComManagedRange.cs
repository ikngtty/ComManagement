using System;
using Microsoft.Office.Interop.Excel;

namespace Ikngtty.ComManagement.Excel
{
	/// <summary>
	/// An excel range COM object wrapper, which delegates it's life span management to the manager.
	/// </summary>
	public class ComManagedRange
	{
		private readonly ComManager comManager;
		private readonly Range com;
		
		public ComManagedRange(ComManager comManager, Range com)
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
		
		public ComManagedRange this[int rowIndex, int columnIndex]
		{
			get
			{
				// Validate.
				if (rowIndex < 1) throw new ArgumentOutOfRangeException();
				if (columnIndex < 1) throw new ArgumentOutOfRangeException();
				
				// Return a new range.
				return new ComManagedRange(this.comManager, (Range)this.com[rowIndex, columnIndex]);
			}
		}
		
		public object Value
		{
			get { return this.com.Value; }
			set { this.com.Value = value; }
		}
	}
}
