# ComManagement
COM object management objects to reduce the annoying "ReleaseComObject" steps.

## Attension
* This is my idea memo, so this program is incomplete. Many necessarry features aren't implemented.
* This program needs Microsoft.Office.Interop.Excel.dll.
* This program is tested with NUnit.
* This program is created with Sharp Develop.

## Example of use
```cs
// Get an excel cell value.
using (var comManager = new ComManager())
{
	var excel = new ComManagedApplication(comManager);
	ComManagedWorkbook book = excel.Workbooks.Open(@"C:/TestFile/ForTest.xlsx");
	var cellValue = (string)book.Sheets[1].Cells[1, 1].Value;
	
	Console.WriteLine(cellValue);
}

// All COM objects are released by a ComManager.
```

## License
MIT License (see [LICENSE.txt](LICENSE.txt))

## Used Library
### NUnit 2.6.4 ([nunit.framework.dll](Library/NUnit-2.6.4/nunit.framework.dll))
* Author  
Portions Copyright © 2002-2014 Charlie Poole  
or Copyright © 2002-2004 James W. Newkirk, Michael C. Two, Alexei A. Vorontsov  
or Copyright © 2000-2002 Philip A. Craig

* License  
NUnit License (see [License/NUnit.txt](License/NUnit.txt))
