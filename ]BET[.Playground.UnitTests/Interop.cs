using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using _BET_.Playground.Interop.COM.NET4Callee;
using System.Reflection;

namespace _BET_.Playground.UnitTests
{
    [TestClass]
    public class Interop : MarshalByRefObject
    {
        [TestMethod]
        [TestCategory("Interop")]
        public void TestNET4Callee()
        {
            var target4 = new Callee();
            var target4Result = target4.WhoIAm();
            var target4Adapter = new COMAdapter();
            var target4AdapterResult = target4Adapter.WhoIAm();

            Assert.AreEqual<string>(target4Result, target4AdapterResult);
            Assert.AreEqual<string>("NET Dll Version v4.0.30319", target4Result);
        }

        public void Execute()
        {
            var net2CallerType = typeof(_BET_.Playground.Interop.COM.NET2Caller.Program);
            Assembly net2Caller = Assembly.GetAssembly(net2CallerType);
            Console.WriteLine(net2Caller.EntryPoint.Invoke(null, new object[] { new string[] { } }));
        }

        [TestMethod]
        [TestCategory("Interop")]
        public void TestNET2Caller()
        {
            var net2CallerType = typeof(_BET_.Playground.Interop.COM.NET2Caller.MainWrapper);
            Assembly net2Caller = Assembly.GetAssembly(net2CallerType);
            //Console.WriteLine(net2Caller.EntryPoint.Invoke(null, new object[] { new string[] { } }));

            // setup
            AppDomainSetup setup2 = new AppDomainSetup()
            {
            };
            AppDomainSetup setup = AppDomain.CurrentDomain.SetupInformation;
            
            // create domain
            AppDomain net2CallerDomain = AppDomain.CreateDomain(
                "TestNET2Caller",
                AppDomain.CurrentDomain.Evidence,
                setup
                );

            try
            {
                // create instance of object to execute by using CreateInstanceAndUnwrap on the domain object
                // returns a remoting proxy you can use to communicate with in your main domain
                // the object behind the proxy must implement MarshalByRefObject or else it will just serialize
                // a copy back to the main domain!
                var prg = net2CallerDomain.CreateInstanceAndUnwrap(net2Caller.FullName, net2CallerType.FullName) 
                    as _BET_.Playground.Interop.COM.NET2Caller.MainWrapper;

                prg.CallCOM();
            }
            finally
            {
                AppDomain.Unload(net2CallerDomain);
            }
        }
    }
}
