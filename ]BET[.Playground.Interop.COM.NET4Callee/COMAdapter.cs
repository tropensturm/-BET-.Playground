using System;
using System.Runtime.InteropServices;

namespace _BET_.Playground.Interop.COM.NET4Callee
{
    [ComVisible(true)]
    [Guid("e0bafa8b-125e-4ac0-ad5c-72fd850024b9")]
    public class COMAdapter : ICOMAdapter
    {
        private Callee target = new Callee();

        public string WhoIAm()
        {
            return target.WhoIAm();
        }
    }
}
