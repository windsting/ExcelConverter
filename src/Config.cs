using System;
using System.Collections.Generic;
using System.Reflection;
using System.Linq;

namespace ExcelConverter
{
    public class Config : Singleton<Config>
    {
        public string outFormat { get; set; } = "json";
        public string output { get; set; } = null;
        public List<string> passPostfix { get; set; } = new List<string>();
        public List<string> ErasePostfix { get; set; } = new List<string>();
        public int nameRow { get; set; } = 2;
        public bool Indent { get; set; } = DefaultIndent;
        public bool KeepExtension { get; set; } = DefaultExtensionSetting;
        public bool Quiet { get; set; } = DefaultQuiet;

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
        public List<string> NoneOption { get; set; } = new List<string>();

        const BindingFlags FieldFlags = BindingFlags.Instance | BindingFlags.Public | BindingFlags.SetProperty | BindingFlags.IgnoreCase;
        public int AddField(string strOption, string value = null)
        {
            int jumpCount = 1;  //  most 
            var name = GetFullName(strOption);
            var obj = this;
            var prop = obj.GetType().GetProperty(name, FieldFlags);
            if (prop == null)
            {
                throw new ArgumentException($"Invalid FieldName:[{name}] for value:[{value}]");
            }

            var propName = prop.PropertyType.Name;
            switch (propName)
            {
                case nameof(Int32):
                    prop.SetValue(obj, Int32.Parse(value));
                    break;
                case nameof(String):
                    prop.SetValue(obj, value);
                    break;
                case "List`1":
                    {
                        var type = prop.PropertyType;
                        if(type.IsGenericType && type.GetGenericTypeDefinition() == typeof(List<>))
                        {
                            var TType = type.GetGenericArguments()[0];
                            switch(TType.Name)
                            {
                                case nameof(Int32):
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
                    jumpCount = 0;
                    break;
                default:
                    throw new Exception($"Invalid Config property type:[{prop.PropertyType.Name}], only:[{nameof(Int32)},{nameof(String)}] supported");
            }

            return jumpCount;
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
            NamePair.Create("o",$"{nameof(output)}"),
            NamePair.Create("i",$"{nameof(Indent)}"),
            NamePair.Create("q",$"{nameof(Quiet)}"),
            NamePair.Create("p",$"{nameof(passPostfix)}"),
            NamePair.Create("e",$"{nameof(ErasePostfix)}"),
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
                    var nextIndex = i + 1;
                    var next = nextIndex < args.Length ? args[nextIndex] : null;
                    i += config.AddField(current, next);
                }
                else
                {
                    config.NoneOption.Add(current);
                    if (config == Config.Instance && config.NoneOption.Count == 0)
                    {
                        config.NoneOption.Add(System.IO.Directory.GetCurrentDirectory());
                    }
                }
            }

            return config;
        }
    }
}