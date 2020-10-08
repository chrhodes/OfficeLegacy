using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;

namespace SupportTools_Excel
{

  [Guid("246A6A1B-0952-413F-920C-583DFBF79DDC")] 
  [ClassInterface(ClassInterfaceType.AutoDual)]
  [ComVisible(true)]
    class UDF
    {
        public string Func1()
        {
            return "bar";
        }

        public string Func2(Range inputRange)
        {
            return "foo";
        }

        [ComRegisterFunctionAttribute]
        public static void RegisterFunction(Type type)
        { 
            Registry.ClassesRoot.CreateSubKey(GetSubKeyName(type, "Programmable"));
 
            RegistryKey key = Registry.ClassesRoot.OpenSubKey(GetSubKeyName(type, "InprocServer32"), true);
 
            key.SetValue(
                "", 
                System.Environment.SystemDirectory + @"\mscoree.dll",
                RegistryValueKind.String);
        }

 
        [ComUnregisterFunctionAttribute]
        public static void UnregisterFunction(Type type)
        {
          Registry.ClassesRoot.DeleteSubKey(GetSubKeyName(type, "Programmable"), false);
        }

 
        private static string GetSubKeyName(Type type, string subKeyName)
        {
 
          System.Text.StringBuilder s = new System.Text.StringBuilder();
 
          s.Append(@"CLSID\{");
          s.Append(type.GUID.ToString().ToUpper());
          s.Append(@"}\");
          s.Append(subKeyName);
 
          return s.ToString();
        }  

    }
}
