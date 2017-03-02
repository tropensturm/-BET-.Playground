using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using _BET_.Playground.Interop.COM.NET4Callee;
using System.Reflection;
using _BET_.Playground.Core;

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

        [TestMethod]
        [TestCategory("Interop")]
        public void TestNET2Caller_Try1()
        {
            var net2CallerType = typeof(_BET_.Playground.Interop.COM.NET2Caller.Program);
            Assembly net2Caller = Assembly.GetAssembly(net2CallerType);

            net2Caller.EntryPoint.Invoke(null, new object[] { new string[] { "UT" } });
            var mainWrapper = net2Caller.GetType("_BET_.Playground.Interop.COM.NET2Caller.MainWrapper");
            var callCOM = mainWrapper.GetMethod("CallCOM");

            var instance = Activator.CreateInstance(mainWrapper);
            var result = callCOM.Invoke(instance, null) as string;

            Assert.AreEqual<string>(_BET_.Playground.Interop.COM.NET2Caller.MainWrapper.ErrorMsg, result); // fail
            Assert.Inconclusive("applies not the app.manifest");
        }

        [TestMethod]
        [TestCategory("Interop")]
        public void TestNET2Caller_Try4()
        {
            var path = @"..\..\..\]BET[.Playground.Interop.COM.NET2Caller\bin\Debug\Playground.Interop.COM.NET2Caller.exe";

            Assembly net2Caller = Assembly.LoadFrom(@"..\..\..\]BET[.Playground.Interop.COM.NET2Caller\bin\Debug\Playground.Interop.COM.NET2Caller.exe");

            net2Caller.EntryPoint.Invoke(null, new object[] { new string[] { "UT" } });
            var mainWrapper = net2Caller.GetType("_BET_.Playground.Interop.COM.NET2Caller.MainWrapper");

            var handler = Activator.CreateComInstanceFrom(path, "_BET_.Playground.Interop.COM.NET2Caller.MainWrapper");
            var prg = handler.Unwrap();

            var callCOM = prg.GetType().GetMethod("CallCOM");
            var result = callCOM.Invoke(prg, null) as string;

            Assert.AreEqual<string>(_BET_.Playground.Interop.COM.NET2Caller.MainWrapper.ErrorMsg, result); // fail
            Assert.Inconclusive("applies not the app.manifest");
        }

        [TestMethod]
        [TestCategory("Interop")]
        public void TestNET2Caller_Try2()
        {
            var net2CallerType = typeof(_BET_.Playground.Interop.COM.NET2Caller.MainWrapper);
            Assembly net2Caller = Assembly.GetAssembly(net2CallerType);

            // setup
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

                var result = prg.CallCOM();

                Assert.AreEqual<string>(_BET_.Playground.Interop.COM.NET2Caller.MainWrapper.ErrorMsg, result); // fail
            }
            catch(Exception ex)
            {
                Assert.Fail(ex.Message);
            }
            finally
            {
                AppDomain.Unload(net2CallerDomain);
            }
            
            Assert.Inconclusive("applies not the app.manifest");
        }

        [TestMethod]
        [TestCategory("Interop")]
        public void TestNET2Caller_Try3()
        {
            Assembly net2Caller = Assembly.LoadFrom(@"..\..\..\]BET[.Playground.Interop.COM.NET2Caller\bin\Debug\Playground.Interop.COM.NET2Caller.exe");
            
            AppDomainSetup setup = new AppDomainSetup()
            {
                PrivateBinPath = net2Caller.CodeBase,
                ApplicationBase = net2Caller.CodeBase,
                ApplicationName = "TestNET2Caller",
                ShadowCopyFiles = Boolean.TrueString
            };
            
            // create domain
            AppDomain net2CallerDomain = AppDomain.CreateDomain(
                "TestNET2Caller",
                null,
                setup
                );

            try
            {
                // create instance of object to execute by using CreateInstanceAndUnwrap on the domain object
                // returns a remoting proxy you can use to communicate with in your main domain
                // the object behind the proxy must implement MarshalByRefObject or else it will just serialize
                // a copy back to the main domain!
                var handle = net2CallerDomain.CreateComInstanceFrom(net2Caller.ManifestModule.FullyQualifiedName, net2Caller.GetType().FullName);
                var prg = handle.Unwrap() as _BET_.Playground.Interop.COM.NET2Caller.MainWrapper;
                var result = prg.CallCOM();
                
                Assert.AreEqual<string>(_BET_.Playground.Interop.COM.NET2Caller.MainWrapper.ErrorMsg, result); // fail
            }
            catch (Exception ex)
            {
                Assert.AreEqual<int>(ex.HResult, -2146233054, "failed as expected");
                Assert.Inconclusive("cannot load assembly because in .net 4.5 System.Runtime.InteropServices moved from system.core to mscorlib :( no work around known to me");
            }
            finally
            {
                AppDomain.Unload(net2CallerDomain);
            }

            Assert.Fail("should never hit here");
        }
    }
}
