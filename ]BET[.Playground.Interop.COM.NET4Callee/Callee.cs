using System.Reflection;

namespace _BET_.Playground.Interop.COM.NET4Callee
{
    public class Callee
    {
        public string WhoIAm()
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            return $"NET Dll Version {assembly.ImageRuntimeVersion}";
        }
    }
}