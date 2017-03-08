using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace _BET_.Playground.UnitTests
{
    [TestClass]
    public class AppDomainTests
    {
        /// <summary>
        /// ClassInitialize: executed one time, setting state that does not change during test
        /// TestInitialize: execute before each method, reset members for each test run
        /// </summary>
        [ClassInitialize]
        public static void InteropInit(TestContext context)
        {
#if DEBUG
            System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("en-US");
#endif
        }

        private static readonly string TestPath = @"..\..\..\deployment";
        private static readonly string TestObject45 = "_BET_.Playground.NET45App.MainWrapper";
        private static readonly string TestMethod45 = "Call";
        //private static readonly string TestAssembly45RelativePath = @"..\..\..\]BET[.Playground.NET45App\bin\Debug\Playground.NET45App.exe";
        //private static readonly string TestAssembly45RelativePath = @"..\..\..\deployment\Playground.NET45App.exe";
        private static readonly string TestAssembly45RelativePath = @"C:\Users\mark\OneDrive\Dokumente\Visual Studio 2015\Projects\]BET[.Playground\deployment\Playground.NET45App.exe";
        private static readonly string NET45Config = @"<?xml version='1.0' encoding='utf-8'?>
<configuration>  
    <startup>  
        <supportedRuntime version='v4.0' sku='.NETFramework,Version=v4.5'/>  
    </startup>  
</configuration>";

        [TestMethod]
        [TestCategory("AppDomain")]
        public void TestConfig()
        {
            // just verify that it is solid xml
            XmlDocument xml = new XmlDocument();
            xml.LoadXml(NET45Config);

            Assert.IsNotNull(xml.InnerXml);
        }

        [TestMethod]
        [TestCategory("AppDomain")]
        public void TestNET45App_Try2()
        {
            var result = Helper.InvokeMethodFromObjectByPath(TestAssembly45RelativePath, TestObject45, TestMethod45) as string;

            Assert.AreEqual<string>("This assembly is NET Dll Version v4.0.30319", result);
        }

        [TestMethod]
        [TestCategory("AppDomain")]
        public void TestNET45App_Try3()
        {
            var result = Helper.InvokeMethodFromCOMObjectByAssembly(TestAssembly45RelativePath, TestObject45, TestMethod45, null) as string;

            Assert.AreEqual<string>("This assembly is NET Dll Version v4.0.30319", result);
        }

        [TestMethod]
        [TestCategory("AppDomain")]
        public void TestNET45App_AppDomain2()
        {
            AppDomainSetup domain = new AppDomainSetup();
            domain.ApplicationBase = TestPath;
            var adevidence = AppDomain.CurrentDomain.Evidence;

            AppDomain net45AppDomain = AppDomain.CreateDomain("TestNET45App", adevidence, domain);
            Type proxy = typeof(Proxy);

            try
            {
                // create instance of object to execute by using CreateInstanceAndUnwrap on the domain object
                // returns a remoting proxy you can use to communicate with in your main domain
                // the object behind the proxy must implement MarshalByRefObject or else it will just serialize
                // a copy back to the main domain!
                var prg = (Proxy) net45AppDomain.CreateInstanceAndUnwrap(proxy.Assembly.FullName, proxy.FullName);

                // this will only work when it is referenced in the project?
                var assembly = prg.GetAssembly(TestAssembly45RelativePath);
                var instance = Activator.CreateInstance(assembly.GetType(TestObject45));

                var call = assembly.GetType(TestObject45).GetMethod("Call");
                var result = call.Invoke(instance, null);

                Assert.AreEqual<string>("This assembly is NET Dll Version v4.0.30319", (string)result);
            }
            catch (Exception ex)
            {
                Assert.Fail($"Unknown Fail: {ex.Message}");
            }
            finally
            {
                AppDomain.Unload(net45AppDomain);
            }
        }


        [TestMethod]
        [TestCategory("AppDomain")]
        public void TestNET45App_AppDomain2_CustomDomain2()
        {
            AppDomainSetup setup = new AppDomainSetup()
            {
                PrivateBinPath = TestPath,
                ApplicationBase = TestPath,
                ApplicationName = "TestNET45App",
                TargetFrameworkName = ".NETFramework,Version=v4.5",
            };

            setup.SetConfigurationBytes(System.Text.Encoding.UTF8.GetBytes(NET45Config));

            Type proxy = typeof(Proxy);
            System.Security.Policy.Evidence evidence = AppDomain.CurrentDomain.Evidence;
            
            // create domain
            AppDomain net45AppDomain = AppDomain.CreateDomain(
                "TestNET45App",
                evidence,
                setup
                );

            try
            {
                var prg = (Proxy) net45AppDomain.CreateInstanceAndUnwrap(proxy.Assembly.FullName, proxy.FullName);

                // this will only work when it is referenced in the project?
                var assembly = prg.GetAssembly(TestAssembly45RelativePath);
                var instance = Activator.CreateInstance(assembly.GetType(TestObject45));

                var call = assembly.GetType(TestObject45).GetMethod("Call");
                var result = call.Invoke(instance, null) as string;

                Assert.AreEqual<string>("This assembly is NET Dll Version v4.0.30319", result);
            }
            catch (Exception ex)
            {
                Assert.Fail($"Unknown Fail: {ex.Message}");
            }
            finally
            {
                AppDomain.Unload(net45AppDomain);
            }
        }
    }
}
