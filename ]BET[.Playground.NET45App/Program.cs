using System;
using System.Reflection;

namespace _BET_.Playground.NET45App
{
    class Program
    {
        static void Main(string[] args)
        {
#if DEBUG
            System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("en-US");
#endif
            var main = new MainWrapper();
            Console.WriteLine(main.Call());
        }
    }

    public sealed class MainWrapper : MarshalByRefObject
    {
        public static readonly string ErrorMsg = "type not found, check your manifest!";

        public string Call()
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            var result = $"This assembly is NET Dll Version {assembly.ImageRuntimeVersion}";
            System.Diagnostics.Trace.WriteLine(result);

            return result;
        }
    }
}
