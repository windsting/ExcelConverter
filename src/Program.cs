using System;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            var cfg = Config.FromArgs(args, Config.Instance);
            var json = JsonConvert.SerializeObject(cfg);

            foreach(var fileName in Config.Instance.NoneOption)
            {
                Process(Config.Instance.NoneOption);
            }

            Func<List<string>, string> toString = (obj) => JsonConvert.SerializeObject(obj, Formatting.Indented);
            if (!Config.Instance.Quiet)
            {
                AppLog.Instance.Log($"{Succceed.Count} {nameof(Succceed)}:{toString(Succceed)}");
                AppLog.Instance.Log($"{Failed.Count} {nameof(Failed)}:{toString(Failed)}");
            }
        }

        static List<string> Succceed = new List<string>();
        static List<string> Failed = new List<string>();

        static string[] ExcelExtensions = new string[] { ".xls", ".xlsx" };
        static void Process(IEnumerable<string> paths)
        {
            foreach(var path in paths)
            {
                if (Directory.Exists(path))
                {
                    var subPaths = Directory.EnumerateFileSystemEntries(path);
                    Process(subPaths);
                    continue;
                }

                if (!File.Exists(path))
                {
                    AppLog.Instance.Log($"invalid arg:{path}");
                    continue;
                }

                var extName = Path.GetExtension(path);
                if (ExcelExtensions.Contains(extName))
                {
                    ConvertFile(path);
                }
            }
        }

        static void ConvertFile(string orgName)
        {
            var excel = ExcelReader.Instance.ReadFile(orgName, Config.Instance);
            if (excel == null)
            {
                AppLog.Instance.Log($"ExcelReader failed on {orgName}");
                Failed.Add(orgName);
                return;
            }

            var config = excel.Config;
            var outName = FileNameConverter.Instance.Convert(orgName, config.outFormat, config);
            var text = excel.JArray.ToString(config.Formatting);
            WriteTextFile(outName, text);

            Succceed.Add(orgName);
        }

        private static async void WriteTextFile(string outName, string text)
        {
            var path = Path.GetDirectoryName(outName);
            Directory.CreateDirectory(path);
            using (var writer = new StreamWriter(outName))
            {
                await writer.WriteAsync(text);
            }
        }
    }
}
