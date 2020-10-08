using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;

namespace ExcelExtensionFunctions
{

  [Guid("246A6A1B-0952-413F-920C-583DFBF79DDC")] 
  [ClassInterface(ClassInterfaceType.AutoDual)]
  [ComVisible(true)]
    public class UDF
    {

        internal const string cIllegalFileCharacters = "[/\\\\:\\*\\?\"<>\\|#\\{}%~&]|\\.\\.";

        internal const string cIllegalFolderCharacters = "[:\\*\\?\"<>\\|#\\{}%~&]";

        public string RegularExpression(Range inputRange, Range regularExpression)
        {
            string input = inputRange.Value;
            string expression = regularExpression.Value;

            Match foundMatch = Regex.Match(input, expression);

            return foundMatch.Value;
        }

        //public string Func2(object inputRange, object regEx, object outputRange)
        //{
        //    Range input = (Range)inputRange;
        //    Range expression = (Range)regEx;
        //    Range output = (Range)outputRange;

        //    output.Value = input.Value;
        //    output.Offset[0, 2].Value = expression.Value;

        //    return "wowwee";
        //    return inputRange.GetType().ToString();
        //    //outputRange.Value = "wee";
        //    //outputRange.Value = inputRange.Value;
        //    //outputRange.Offset[0, 2].Value = regEx.Value;
            
        //    //return "wow";
        //}

        #region Registration Functions

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

        #endregion

    }
}
