using System.IO;

namespace ExcelConverter
{
    public class FileNameConverter : Singleton<FileNameConverter>
    {
        public string Convert(string orgName, string extension, Config config)
        {
            var orgFull = Path.GetFullPath(orgName);
            var orgExt = Path.GetExtension(orgName).Replace(".", string.Empty);

            var basename = Path.GetFileNameWithoutExtension(orgFull);
            var dir = Config.Instance.output ?? config.output ?? Path.GetDirectoryName(orgFull);
            extension = config.KeepExtension ? $"{orgExt}.{extension}" : extension;
            var fullname = $"{dir}/{basename}.{extension}";

            return fullname;
        }
    }
}
