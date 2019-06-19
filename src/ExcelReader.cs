using System;
using System.IO;
using OfficeOpenXml;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System.Linq;

namespace ExcelConverter
{
    public class ExcelInfo
    {
        public Config Config { get; set; }
        public JArray JArray { get; set; }
    }

    public class ConvertObj {
        public string FileName { get; set; }
        public Config Config { get; set; }
        public ExcelWorksheets Sheets { get; set; }
        public Stack<string> SheetStack { get; set; } = new Stack<string>();
        public Dictionary<string, JArray> NamedSheets { get; set; } = new Dictionary<string, JArray>();

        public ConvertObj(ExcelWorksheets sheets, string fileName, Config config)
        {
            Sheets = sheets;
            FileName = fileName;
            Config = config;
            SheetStack.Push(sheets[1].Name);
        }
    }

    public class ExcelReader : Singleton<ExcelReader>
    {
        const int MinRowCount = 2;
        public ExcelInfo ReadFile(string fileName, Config config)
        {
            FileInfo fileInfo = new FileInfo(fileName);
            if (!fileInfo.Exists)
            {
                AppLog.Instance.Log($"file not found:\n    {fileName}");
                return null;
            }

            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                var co = new ConvertObj(package.Workbook.Worksheets, fileName, config);
                var array = ConvertSheet(co);
                if (co.SheetStack.Count > 0)
                {
                    var stackContent = JsonConvert.SerializeObject(co.SheetStack);
                    throw new Exception($"ConvertSheet finished with stack content:{stackContent}");
                }

                //AppLog.Instance.Log(array.ToString());
                return new ExcelInfo { Config = config, JArray = array };
            }
        }

        static bool IsPassProperty(string propertyName)
        {
            foreach (var postfix in Config.Instance.PassPostfix)
                if (propertyName.EndsWith(postfix))
                    return true;

            return false;
        }

        static string ErasePostfix(string propertyName)
        {
            foreach (var postfix in Config.Instance.ErasePostfix)
                if (propertyName.EndsWith(postfix))
                    return propertyName.Substring(0,propertyName.IndexOf(postfix));

            return propertyName;
        }

        static JArray ConvertSheet(ConvertObj co)
        {
            var fileName = co.FileName;
            var config = co.Config;
            var stack = co.SheetStack;
            var sheetName = stack.Peek();
            var worksheet = co.Sheets[sheetName];
            if (worksheet == null)
            {
                throw new Exception($"Invalid referenced sheet name:[{sheetName}]");
            }
            int rowCount = worksheet.Dimension.End.Row;
            int colCount = worksheet.Dimension.Columns;

            if (rowCount < MinRowCount)
            {
                AppLog.Instance.Log($"invalid rowCount:{rowCount} {nameof(MinRowCount)}:{MinRowCount} in file:\n    {fileName}");
                return null;
            }

            Func<int, int, object> vCell = (r, c) => worksheet.Cells[r, c].Value;
            Func<int, int, string> sCell = (r, c) =>
            {
                var value = vCell(r, c);
                return value?.ToString();
            };

            config = TryParseConfig(sCell(1, 1)) ?? config;

            List<string> keys = FetchKeys(config, colCount, sCell);
            var keyCount = keys.Count;
            if (keyCount < 1)
            {
                AppLog.Instance.Log($"name row has no value in file:\n    {fileName}");
                return null;
            }

            JArray array = new JArray();
            for (int row = config.NameRow + 1; row <= rowCount; ++row)
            {
                Func<int, string> getKey = (c) => keys[c - 1];
                Func<int, string> getValue = (c) => sCell(row, c);
                var dataRow = worksheet.Cells[row, 1, row, keyCount];
                JObject jobj = GenerateObject(keyCount, getKey, getValue, co);
                if (jobj == null && co.Config.StopOnEmptyRow) {
                    break;
                }
                if (jobj != null)
                    array.Add(jobj);
            }

            stack.Pop();
            return array;
        }

        static JArray GetSheet(string sheetName, ConvertObj co) {
            co.NamedSheets.TryGetValue(sheetName, out var sheet);
            if (sheet != null) {
                return sheet;
            }
            if (co.SheetStack.Contains(sheetName)) {
                var strStack = JsonConvert.SerializeObject(co.SheetStack);
                throw new Exception($"sheet [{sheetName}] gonna be referenced recursively, convert stack is:{strStack}");
            }
            co.SheetStack.Push(sheetName);
            sheet = ConvertSheet(co);
            co.NamedSheets.Add(sheetName, sheet);
            return sheet;
        }

        const string ArraySplitters = "|;,";
        private static JToken ParseJArray(string value, int level = 0) {
            if (level == 0 && value.IndexOf(ArraySplitters[0]) < 0) {
                //  this is not an array notation
                return null;
            }

            if(level >= ArraySplitters.Length) {
                //  last level treat as a token
                return ParseArrayElement(value);
            }

            var splitter = ArraySplitters[level];
            var parts = value.Split(splitter);
            if (parts.Length > 1) {
                JArray arr = new JArray();
                bool empty = true;
                foreach (var part in parts) {
                    var val = ParseJArray(part, level + 1);
                    if (val != null) {
                        arr.Add(val);
                        empty = false;
                    }
                }
                if (empty)
                    return null;
                return arr;
            }

            // this is the element of a [level + 1] dimention array
            return ParseArrayElement(value);
        }

        private static JToken ParseArrayElement(string value) {
            if (string.IsNullOrEmpty(value))
                return null;

            return ParseToken(value);
        }

        private static JToken ParseToken(string value) {
            if (long.TryParse(value, out var number))
                return number;

            return value;
        }

        private static List<JToken> FetchRows(ConvertObj co, string sheetName, string refColName, string[] refColValues) {
            var jarray = GetSheet(sheetName, co);
            List<JToken> rows = new List<JToken>();
            foreach(var refValue in refColValues) {
                var lst = jarray.Children<JToken>();
                var matched = lst.Where(o => o[refColName] != null && o[refColName].ToString() == refValue);
                if (matched != null) {
                    foreach(var jobj in matched) {
                        rows.Add(jobj);
                    }
                }
            }

            return rows;
        }

        private static List<JToken> GetRefList(string[] parts, string value, ConvertObj co) {
            var sheetName = parts[1];
            var refCol = parts[2];
            var refValues = SubArray(parts, 3, parts.Length - 3);
            var lst = FetchRows(co, sheetName, refCol, refValues);
            return lst;
        }

        public static T[] SubArray<T>(T[] data, int index, int length) {
            if(length <= 0) {
                return new T[] { };
            }
            T[] result = new T[length];
            Array.Copy(data, index, result, 0, length);
            return result;
        }

        public static JArray ToJarray<T>(IEnumerable<T> lst) {
            var jarray = new JArray();
            foreach (var obj in lst)
                jarray.Add(obj);
            return jarray;
        }

        private static JToken ParseRef(string value, ConvertObj co) {
            if (!value.StartsWith("[") || !value.EndsWith("]"))
                return null;
            var strparts = value.Trim('[', ']');
            var parts = strparts.Split(":");
            if(parts.Length < 2) {
                throw new Exception($"no sheet name provided in ref cell: {value}");
            }
            switch (parts[0]) {
                case "ref": {
                        var sheetName = parts[1];
                        return GetSheet(sheetName, co);
                    }
                case "refOne": {
                        var lst = GetRefList(parts, value, co);
                        if (lst.Count == 0) {
                            throw new Exception($"no value found for reference: {value}");
                        }
                        if (lst.Count > 1) {
                            throw new Exception($"more than ONE values found for reference: {value}, they are: {JsonConvert.SerializeObject(lst, Formatting.Indented)}");
                        }
                        return lst[0];
                    }
                case "refMany": {
                        var lst = GetRefList(parts, value, co);
                        return ToJarray(lst);
                    }
                default:
                    throw new Exception($"Invalid command:[{parts[0]}] in {nameof(ParseRef)} value is:{value}");
            } // all paths returned, "return" not necessary
        }

        private static JToken ParseCell(string value, ConvertObj co) {
            if (value == null)
                return null;

            var refData = ParseRef(value, co);
            if (refData != null)
                return refData;

            var array = ParseJArray(value);
            if (array != null)
                return array;

            return ParseToken(value);
        }

        private static JObject GenerateObject(int keyCount, Func<int, string> getKey, Func<int, string> getValue, ConvertObj co)
        {
            JObject jobj = new JObject();
            bool IsAllPropertyNull = true;
            for (int col = 1; col <= keyCount; ++col)
            {
                var key = getKey(col);
                if (IsPassProperty(key))
                    continue;

                key = ErasePostfix(key);

                var stringValue = getValue(col);
                var cellValue = ParseCell(stringValue, co);
                if(cellValue != null)
                {
                    jobj.Add(key, cellValue);
                }

                if (stringValue != null)
                    IsAllPropertyNull = false;
            }

            return IsAllPropertyNull ? null : jobj;
        }

        const string ConfigMagicKey = "ecconfig";
        private static Config TryParseConfig(string firstCell)
        {
            Config config = null;
            if (firstCell != null && firstCell.StartsWith(ConfigMagicKey))
            {
                var argString = firstCell.Replace(ConfigMagicKey, string.Empty);
                var args = Regex.Split(argString, @"\s");
                config = Config.FromArgs(args);
            }

            return config;
        }

        private static List<string> FetchKeys(Config config, int colCount, Func<int, int, string> sCell)
        {
            var nameRow = config.NameRow;
            List<string> keys = new List<string>();
            for (var col = 1; col <= colCount; ++col)
            {
                var val = sCell(nameRow, col);
                if (val == null || string.IsNullOrWhiteSpace(val))
                    break;
                keys.Add(val);
            }

            return keys;
        }
    }
}
