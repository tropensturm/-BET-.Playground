using System;
using System.Runtime.InteropServices;

namespace _BET_.Playground.Interop.COM.NET4Callee
{
    [ComVisible(true)]
    [Guid("af99c64e-94d4-494a-b822-660ce2ec9a56")]
    public interface ICOMAdapter
    {
        string WhoIAm();
    }
}
