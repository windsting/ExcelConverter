using System;
using System.IO;
using OfficeOpenXml;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;

namespace ExcelConverter
{
    public class ExcelInfo
    {
        public Config Config { get; set; }
        public JArray JArray { get; set; }
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
                var worksheet = package.Workbook.Worksheets[1];
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
                if(keyCount < 1)
                {
                    AppLog.Instance.Log($"name row has no value in file:\n    {fileName}");
                    return null;
                }

                JArray array = new JArray();
                for (int row = config.NameRow+1; row <= rowCount; ++row)
                {
                    Func<int, string> getKey = (c) => keys[c - 1];
                    Func<int, string> getValue = (c) => sCell(row, c);
                    var dataRow = worksheet.Cells[row, 1, row, keyCount];
                    JObject jobj = GenerateObject(keyCount, getKey, getValue);
                    if (jobj != null)
                        array.Add(jobj);
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

        private static JObject GenerateObject(int keyCount, Func<int, string> getKey, Func<int, string> getValue)
        {
            JObject jobj = new JObject();
            bool IsAllPropertyNull = true;
            for (int col = 1; col <= keyCount; ++col)
            {
                var key = getKey(col);
                if (IsPassProperty(key))
                    continue;

                key = ErasePostfix(key);

                var value = getValue(col);
                if (long.TryParse(value, out var number))
                    jobj.Add(key, number);
                else
                    jobj.Add(key, value);

                if (value != null)
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