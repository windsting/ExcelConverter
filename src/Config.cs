using System;
using System.Collections.Generic;
using System.Reflection;
using System.Linq;

namespace ExcelConverter
{
    public class Config : Singleton<Config>
    {
        public string OutFormat { get; set; } = "json";
        public string Output { get; set; } = null;
        public List<string> PassPostfix { get; set; } = new List<string>();
        public List<string> ErasePostfix { get; set; } = new List<string>();
        public int NameRow { get; set; } = 2;
        public bool Indent { get; set; } = DefaultIndent;
        public bool KeepExtension { get; set; } = DefaultExtensionSetting;
        public bool Quiet { get; set; } = DefaultQuiet;
        public bool StopOnEmptyRow { get; set; } = DefaultStopOnEmptyRow;

        public Newtonsoft.Json.Formatting Formatting
        {
            get
            {
                return Indent ? Newtonsoft.Json.Formatting.Indented : Newtonsoft.Json.Formatting.None;
            }
        }

        const bool DefaultIndent = false;
        const bool DefaultExtensionSetting = false;
        const bool DefaultQuiet = false;
        const bool DefaultStopOnEmptyRow = false;
        public List<string> NoneOption { get; set; } = new List<string>();

        public List<string> MandatoryOptions = new List<string>();

        const BindingFlags FieldFlags = BindingFlags.Instance | BindingFlags.Public | BindingFlags.SetProperty | BindingFlags.IgnoreCase;
        public void AddField(string strOption)
        {
            var parts = strOption.Split('=');
            strOption = parts[0];
            var value = parts.Length > 1 ? parts[1] : null;

            var name = GetFullName(strOption);
            var prop = this.GetType().GetProperty(name, FieldFlags);
            if (prop == null)
            {
                throw new ArgumentException($"Invalid FieldName:[{strOption}] for value:[{value}]");
            }

            try
            {
                SetValue(value, prop, strOption);
            }
            catch (Exception ex)
            {
                AppLog.Instance.Log($"Field:{name} set value:[{value}] got Exception:{ex.ToString()}");
            }
        }

        public bool IsOptionMandatory(string optionName)
        {
            return MandatoryOptions.Contains(optionName);
        }

        private void SetValue(string value, PropertyInfo prop, string option)
        {
            var obj = this;
            var propName = prop.PropertyType.Name;
            MandatoryOptions.Add(propName);
            switch (propName)
            {
                case nameof(Int32):
                    {
                        if (!Int32.TryParse(value, out int iValue))
                        {
                            AppLog.Instance.Log($"Invalid value:[{value}] for option:{option}");
                            return;
                        }
                        prop.SetValue(obj, Int32.Parse(value));
                    }
                    break;
                case nameof(String):
                    prop.SetValue(obj, value);
                    break;
                case "List`1":
                    {
                        var type = prop.PropertyType;
                        if (type.IsGenericType && type.GetGenericTypeDefinition() == typeof(List<>))
                        {
                            var TType = type.GetGenericArguments()[0];
                            switch (TType.Name)
                            {
                                case nameof(Int32):
                                    if (!Int32.TryParse(value, out int iValue))
                                    {
                                        AppLog.Instance.Log($"Invalid value for option:{option}");
                                        return;
                                    }
                                    type.GetMethod($"Add").Invoke(prop.GetValue(obj), new object[] { Int32.Parse(value) });
                                    break;
                                case nameof(String):
                                    type.GetMethod($"Add").Invoke(prop.GetValue(obj), new[] { value });
                                    break;
                                default:
                                    throw new Exception("Invalid Config property type:[{prop.PropertyType.Name}], only:[{nameof(Int32)},{nameof(String)}] supported");
                            }
                        }
                    }
                    break;
                case nameof(Boolean):
                    var oldVal = (bool)prop.GetValue(obj);
                    prop.SetValue(obj, !oldVal);
                    break;
                default:
                    throw new Exception($"Invalid Config property type:[{prop.PropertyType.Name}], only:[{nameof(Int32)},{nameof(String)}] supported");
            }
        }

        #region NamePair
        public class NamePair
        {
            public string Short { get; set; }
            public string Full { get; set; }

            public static NamePair Create(string s, string f) { return new NamePair { Short = s, Full = f }; }
        }

        static NamePair[] NamePairs = new NamePair[]
        {
            NamePair.Create("i",$"{nameof(Indent)}"),
            NamePair.Create("q",$"{nameof(Quiet)}"),
            NamePair.Create("o",$"{nameof(Output)}"),
            NamePair.Create("k",$"{nameof(KeepExtension)}"),
            NamePair.Create("p",$"{nameof(PassPostfix)}"),
            NamePair.Create("e",$"{nameof(ErasePostfix)}"),
            NamePair.Create("n",$"{nameof(NameRow)}"),
            NamePair.Create("s",$"{nameof(StopOnEmptyRow)}"),
        };

        static string GetFullName(string shortOption)
        {
            var shortName = shortOption.Replace(ShortOptionPrefix, string.Empty);
            var pair = NamePairs.FirstOrDefault(p => p.Short == shortName);
            return pair == null ? shortName : pair.Full;
        }
        #endregion

        const char OptionPrefix = '-';
        const string ShortOptionPrefix = "-";
        const string LongOptionPrefix = ShortOptionPrefix + ShortOptionPrefix;
        public static Config FromArgs(string[] args, Config obj = null)
        {
            Config config = obj ?? new Config();

            for (int i = 0; i < args.Length; ++i)
            {
                var current = args[i];

                if (current.StartsWith(OptionPrefix))
                {
                    config.AddField(current);
                }
                else
                {
                    config.NoneOption.Add(current);
                }
            }

            if (config == Config.Instance && config.NoneOption.Count == 0)
            {
                config.NoneOption.Add(System.IO.Directory.GetCurrentDirectory());
            }

            return config;
        }
    }
}