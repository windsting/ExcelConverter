using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;

namespace ExcelConverter
{
    public class HelpInfo
    {
        static string[] ShortInfo = new string[]
        {
            "No file or directory specified.",
            "Usage: ExcelConverter [-i | --indent] [-q | --quiet] [-k | --keep-extension]",
            "                      [-s | --stop-on-empty-row]",
            "                      [-p=postfixToPass | --postfix-to-pass=postfixToPass]",
            "                      [-e=postfixToErase | --erase-postfix=postfixToErase]",
            "                      [-n=rowNumberOfName | --name-row=rowNumberOfName]",
            "                      [excelFilePath | dirContainExcel]... [-o=outputDir]",
            "       ExcelConverter -h | --help",
            "",
            "       ExcelConverter data.xlsx",
            "       ExcelConverter data.xlsx -o=../config/",
            "       ExcelConverter .",
            "       ExcelConverter . -o=output_dir",
            "       ExcelConverter -i Documents/test.xlsx",
            "       ExcelConverter -k data.xlsx",
            "       ExcelConverter Documents/ -o=config_file_dir",
            "       ExcelConverter -p=ForServer -p=ForWeb -e=ForClient data.xlsx -o=cfgDir",
            "       ExcelConverter data.xlsx -n=3",
        };

        static string[] LongInfo = new string[] 
        {
            "Convert .excel file(s) to .json file(s).",
            "Usage: ExcelConverter [OPTION]... [FILE]...",
            "",
            "Options:",
            "  -h, --help           print this help information, and exit immediately",
            "  -i, --indent         Indent output json file",
            "  -q, --quiet          Don't report processing result",
            "  -o, --output         specify output directory, without this provided,",
            "                       the result .json file will be output to directory",
            "                       which contain the .xlsx file",
            "  -k, --keep-extension keep extension name, i.e. data.xlsx produce" ,
            "                       data.xlsx.json",
            "  -p, --pass-postfix   pass column(s) end with specified postfix, this",
            "                       option can be specified repeatedly",
            "  -e, --erase-postfix  erase postfix for column(s) end with specified",
            "                       postfix, this option can be specified repeatedly",
            "  -n, --name-row       specify the row number for the row contains",
            "                       property names. default value is 2, left the 1st",
            "                       row for describe name",
            "  -s, --stop-on-empty-row",
            "                       stop converting on the first empty row, the result",
            "                       only contains datas above the empty row",
        };

        static string[] ExtraInfo = new string[] 
        {
            "",
            "More detail can be found on help page: ",
            "   <https://github.com/windsting/ExcelConverter>",
        };

        static string[] OptionsForHelp = new string[] { "-h", "--help" };
        public static bool HelpPrinted(string[] args)
        {
            if(args.Length == 0)
            {
                PrintInfo(ShortInfo);
                return true;
            }

            if(args.Length == 1 && OptionsForHelp.Contains(args[0].Trim()))
            {
                // print long form of help information
                PrintInfo(LongInfo);
                PrintInfo(ExtraInfo);
                return true;
            }

            return false;
        }

        static void PrintInfo(string[] infoArray)
        {
            foreach (var line in infoArray)
                AppLog.Instance.Log(line);
        }
    }
}
