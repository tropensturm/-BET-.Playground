using System;
using System.Reflection;
using System.Runtime.InteropServices;

namespace _BET_.Playground.Interop.COM.NET2Caller
{
    // remark: Guid need to be identical with the one from the NET4Callee.ICOMAdapter!
    [ComVisible(true)]
    [Guid("af99c64e-94d4-494a-b822-660ce2ec9a56")]
    public interface ICOMAdapter
    {
        string WhoIAm();
    }

    //[Serializable]
    public class Program
    {
        // this will only start when the assembly of NET4Callee is in gac or in local folder
        static void Main(string[] args)
        {
            MainWrapper main = new NET2Caller.MainWrapper();
            main.CallCOM();
        }
    }

    public sealed class MainWrapper : MarshalByRefObject
    {
        public void CallCOM()
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            System.Diagnostics.Trace.WriteLine(string.Format("NET Dll Version {0}", assembly.ImageRuntimeVersion));

            Type type = Type.GetTypeFromProgID("NET4Callee.COMAdapter");

            if (type == null)
            {
                System.Diagnostics.Trace.WriteLine(string.Format("type not found, check your manifest!"));
                return;
            }

            object instance = Activator.CreateInstance(type);
            ICOMAdapter adapter = (ICOMAdapter)instance;

            System.Diagnostics.Trace.WriteLine(adapter.WhoIAm());
        }
    }
}
