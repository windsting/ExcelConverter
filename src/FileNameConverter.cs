using System.IO;

namespace ExcelConverter
{
    public class FileNameConverter : Singleton<FileNameConverter>
    {
        public static bool IsExcelFile(string fileName)
        {
            foreach (var ext in Program.ExcelExtensions)
            {
                if (fileName.EndsWith(ext))
                    return true;
            }
            return false;
        }

        public string Convert(string orgName, string extension, Config config)
        {
            var orgFull = Path.GetFullPath(orgName);
            var orgExt = Path.GetExtension(orgName).Replace(".", string.Empty);

            var basename = Path.GetFileNameWithoutExtension(orgFull);
            var Output = Config.Instance.Output ?? config.Output;
            if (Output != null
                && Output.ToLower().EndsWith(".json")
                && Config.Instance.NoneOption.Count == 1
                && IsExcelFile(Config.Instance.NoneOption[0]))
                return Path.GetFullPath(Output);

            var dir = Output ?? Path.GetDirectoryName(orgFull);
            extension = Config.Instance.KeepExtension ? $"{orgExt}.{extension}" : extension;
            var fullname = $"{dir}/{basename}.{extension}";

            return fullname;
        }
    }
}
