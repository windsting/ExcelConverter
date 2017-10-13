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
                AppLog.Instance.Log($"file:{fileName} not exist!");
                return null;
            }

            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                var worksheet = package.Workbook.Worksheets[1];
                int rowCount = worksheet.Dimension.End.Row;
                int colCount = worksheet.Dimension.Columns;

                if (rowCount < MinRowCount)
                {
                    AppLog.Instance.Log($"invalid rowCount:{rowCount} {nameof(MinRowCount)}:{MinRowCount}");
                    return null;
                }

                Func<int, int, object> vCell = (r, c) => worksheet.Cells[r, c].Value;
                Func<int, int, string> sCell = (r, c) =>
                {
                    var value = vCell(r, c);
                    if (value != null)
                        return value.ToString();
                    return null;
                };

                config = TryParseConfig(sCell(1, 1), sCell(1, 2)) ?? config;

                List<string> keys = FetchKeys(config, colCount, sCell);
                var keyCount = keys.Count;

                JArray array = new JArray();
                for (int row = config.nameRow+1; row <= rowCount; ++row)
                {
                    Func<int, string> getKey = (c) => keys[c - 1];
                    Func<int, string> getValue = (c) => sCell(row, c);
                    var dataRow = worksheet.Cells[row, 1, row, keyCount];
                    JObject jobj = GenerateObject(keyCount, getKey, getValue);
                    array.Add(jobj);
                }

                //AppLog.Instance.Log(array.ToString());
                return new ExcelInfo { Config = config, JArray = array };
            }
        }

        static bool IsPassProperty(string propertyName)
        {
            foreach (var postfix in Config.Instance.passPostfix)
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
        private static Config TryParseConfig(string firstCell, string argString)
        {
            Config config = null;
            if (firstCell == ConfigMagicKey)
            {
                var args = Regex.Split(argString, @"\s");
                config = Config.FromArgs(args);
            }

            return config;
        }

        private static List<string> FetchKeys(Config config, int colCount, Func<int, int, string> sCell)
        {
            List<string> keys = new List<string>();
            for (var col = 1; col <= colCount; ++col)
            {
                var val = sCell((int)config.nameRow, col);
                if (val == null)
                    break;
                keys.Add(val);
            }

            return keys;
        }
    }
}