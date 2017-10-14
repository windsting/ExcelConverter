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
            "Usage: ExcelConverter [-i | --indent] [-q | --quiet] [-k | --keepExtension]",
            "                      [-p=postfixToPass | --passPostfix=postfixToPass]",
            "                      [-e=postfixToErase | --erasePostfix=postfixToErase]",
            "                      [-n=rowNumberOfName | --nameRow=rowNumberOfName]",
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
            "       ExcelConverter -p=ForServer -p=ForWeb -e=ForClient data.xlsx -o=cfg_dir",
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
            "  -o, --output         specify output directory, without this provided, the result",
            "                       .json file will be output to directory .xlsx file within",
            "  -k, --keepExtension  keep extension name, i.e. data.xlsx produce data.xlsx.json",
            "  -p, --postfixToPass  pass column(s) end with specified postfix, this option",
            "                       can be specified repeatedly",
            "  -e, --postfixToErase erase postfix for column(s) end with specified postfix,",
            "                       this option can be specified repeatedly",
            "  -n, --nameRow        specify the row number for the row contains property names,",
            "                       default value is 2, left the 1st row for describe name"
        };

        static string[] ExtraInfo = new string[] 
        {
            "",
            "More detail online help: <http://192.168.8.172/wangg/ExcelConverter/wikis/home>",
            ""
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
