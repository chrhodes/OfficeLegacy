using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;

namespace ExcelUFD
{
    public interface IFunctions
    {
        string Func1a();

        string Func2a(Range inputRange);
    }
}
