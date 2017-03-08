using System;
using System.Reflection;
using System.Runtime.InteropServices;

namespace _BET_.Playground.Interop.COM.NET45Caller
{
    public class Program : MarshalByRefObject
    {
        // this will only start when the assembly of NET4Callee is in gac or in local folder
        static void Main(string[] args)
        {
            MainWrapper main = new NET45Caller.MainWrapper();

            Console.WriteLine(main.CallCOM());

            if (args.Length == 0 ||
               args[0] != "UT")
                Console.ReadLine();
        }
    }

    // when implementing MarshalByRefObject we wont require the serializable attribute
    public sealed class MainWrapper : MarshalByRefObject
    {
        public static readonly string ErrorMsg = "type not found, check your manifest!";

        public string CallCOM()
        {
#if DEBUG
            System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("en-US");
#endif
            Assembly assembly = Assembly.GetExecutingAssembly();
            System.Diagnostics.Trace.WriteLine(string.Format("This assembly is NET Dll Version {0}", assembly.ImageRuntimeVersion));

            Type type = Type.GetTypeFromProgID("NET4Callee.COMAdapter");

            if (type == null)
            {
                System.Diagnostics.Trace.WriteLine(ErrorMsg);
                return ErrorMsg;
            }

            dynamic instance = Activator.CreateInstance(type);
            var msg = instance.WhoIAm() as string;

            System.Diagnostics.Trace.WriteLine(msg);
            return msg;
        }
    }
}
