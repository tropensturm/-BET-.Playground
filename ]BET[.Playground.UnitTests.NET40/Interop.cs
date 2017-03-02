using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Reflection;

namespace _BET_.Playground.UnitTests.NET40
{
    [TestClass]
    public class Interop
    {
        [TestMethod]
        [TestCategory("Interop40")]
        public void Test40NET2Caller_Try3()
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
                Assert.Fail(ex.Message);
            }
            finally
            {
                AppDomain.Unload(net2CallerDomain);
            }

            Assert.Fail("should never hit here");
        }
    }
}
