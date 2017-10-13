using System;
using System.Reflection;

namespace ExcelConverter
{
    /// <summary>
    /// Template for Singleton
    /// </summary>
    /// <typeparam name="T">The class you want to use in the Singleton</typeparam>
    public class Singleton<T> : SingletonBase<T>
        where T : class
    {
        static object Lock = new object();

        public static T Instance
        {
            get
            {
                if (_instance == null)
                {
                    SetInstance(CreateInstance);
                }

                return _instance;
            }
        }

        static T CreateInstance()
        {
            return Activator.CreateInstance(typeof(T), true) as T;
        }
    }


    /// <summary>
    /// Template for Configurable Singleton
    /// </summary>
    /// <typeparam name="T">The class you want to use in the Singleton</typeparam>
    /// <typeparam name="TConfig">The parameter type your Constructor accept</typeparam>
    public class ConfigurableSingleton<T, TConfig> : SingletonBase<T>
        where T : class
    {
        static TConfig _ConfigSetting;
        static bool _ConfigCalled = false;
        public static void Init(TConfig config)
        {
            _ConfigSetting = config;
            _ConfigCalled = true;
        }

        public static T Instance
        {
            get
            {
                if (_instance == null)
                    SetInstance(CreateInstance);
                return _instance;
            }
        }

        static T CreateInstance()
        {
            if (!_ConfigCalled)
            {
                var tn = typeof(T).Name;
                var ctn = typeof(TConfig).Name;
                var msg = $"Call {tn}.{nameof(Init)}({ctn} config) before using {tn}.Instance";
                throw new InvalidOperationException(msg);
            }

            var bindingAttr = BindingFlags.NonPublic | BindingFlags.CreateInstance | BindingFlags.Instance;
            var constructor = typeof(T).GetConstructor(bindingAttr, null, new[] { typeof(TConfig) }, null);
            return constructor.Invoke(new object[] { _ConfigSetting }) as T;
        }
    }

    public abstract class SingletonBase<T>
        where T : class
    {
        protected SingletonBase() { }

        static object Lock = new object();
        protected static T _instance { get; private set; } = null;
        protected static void SetInstance(Func<T> generator)
        {
            if (_instance == null)
            {
                lock (Lock)
                {
                    if (_instance == null)
                    {
                        _instance = generator();
                    }
                }
            }
        }
    }
}