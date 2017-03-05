using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using _BET_.Playground.Interop.COM.NET4Callee;
using System.Reflection;
using _BET_.Playground.Core;
using System.Linq;
using System.Xml;

namespace _BET_.Playground.UnitTests
{
    [TestClass]
    public class Interop
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

        private static readonly string TestObject = "_BET_.Playground.Interop.COM.NET2Caller.MainWrapper";
        private static readonly string TestMethod = "CallCOM";
        private static readonly string TestAssemblyRelativePath = @"..\..\..\]BET[.Playground.Interop.COM.NET2Caller\bin\Debug\Playground.Interop.COM.NET2Caller.exe";
        private static readonly string InconclusiveMsg = "app.manifest from target not applied";
        private static readonly string NET2Config = @"<configuration>  
    <startup>  
        <supportedRuntime version='v2.0' />  
    </startup>  
</configuration>";

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
        public void TestLoadCorrectAssemblies()
        {
            Assembly net2Caller = Assembly.LoadFrom(TestAssemblyRelativePath);
            var assemblyNames = net2Caller.GetReferencedAssemblies();

            int errors = 0;

            foreach (var an in assemblyNames)
            {
                var rA = Assembly.Load(an);
                if (rA.FullName != an.FullName)
                    errors += 1;
            }

            Assert.AreEqual<int>(0, errors, $"{errors} of {assemblyNames.Length} are loaded in wrong version!");
        }

        [TestMethod]
        [TestCategory("Interop")]
        public void TestNET2Caller_Try1()
        {
            var net2CallerType = typeof(_BET_.Playground.Interop.COM.NET2Caller.Program);
            var result = Helper.InvokeMethodFromObjectByType(net2CallerType, TestObject, TestMethod) as string;

            Assert.AreEqual<string>(_BET_.Playground.Interop.COM.NET2Caller.MainWrapper.ErrorMsg, result); // fail
            Assert.Inconclusive(InconclusiveMsg);
        }

        [TestMethod]
        [TestCategory("Interop")]
        public void TestNET2Caller_Try2()
        {
            var result = Helper.InvokeMethodFromObjectByPath(TestAssemblyRelativePath, TestObject, TestMethod) as string;

            Assert.AreEqual<string>(_BET_.Playground.Interop.COM.NET2Caller.MainWrapper.ErrorMsg, result); // fail
            Assert.Inconclusive(InconclusiveMsg);
        }

        [TestMethod]
        [TestCategory("Interop")]
        public void TestNET2Caller_Try3()
        {
            var result = Helper.InvokeMethodFromCOMObjectByAssembly(TestAssemblyRelativePath, TestObject, TestMethod, null) as string;

            Assert.AreEqual<string>(_BET_.Playground.Interop.COM.NET2Caller.MainWrapper.ErrorMsg, result); // fail
            Assert.Inconclusive(InconclusiveMsg);
        }

        [TestMethod]
        [TestCategory("Interop")]
        public void TestNET2Caller_AppDomain1()
        {
            Assembly net2Caller = Assembly.LoadFrom(TestAssemblyRelativePath);

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
                var prg = net2CallerDomain.CreateInstanceAndUnwrap(net2Caller.FullName, net2Caller.GetType().FullName)
                    as _BET_.Playground.Interop.COM.NET2Caller.MainWrapper;

                var result = prg.CallCOM();

                Assert.AreEqual<string>(_BET_.Playground.Interop.COM.NET2Caller.MainWrapper.ErrorMsg, result); // fail
                Assert.Inconclusive(InconclusiveMsg);
            }
            catch (Exception ex)
            {
                Assert.AreEqual<int>(ex.HResult, -2146233054, "failed as expected");
                Assert.Inconclusive(InconclusiveMsg);
            }
            finally
            {
                AppDomain.Unload(net2CallerDomain);
            }
        }

        [TestMethod]
        [TestCategory("Interop")]
        public void TestNET2Caller_AppDomain2_CustomDomain1()
        {
            Assembly net2Caller = Assembly.LoadFrom(TestAssemblyRelativePath);

            AppDomainSetup setup = new AppDomainSetup()
            {
                PrivateBinPath = net2Caller.CodeBase,
                ApplicationBase = net2Caller.CodeBase,
                ApplicationName = "TestNET2Caller",
                SandboxInterop = true,
                ShadowCopyFiles = Boolean.TrueString,
                TargetFrameworkName = ".NETFramework,Version=v2.0"
            };

            System.Security.Policy.Evidence evidence = new System.Security.Policy.Evidence(AppDomain.CurrentDomain.Evidence);
            evidence.AddAssemblyEvidence(new System.Security.Policy.ApplicationDirectory(@"..\..\..\]BET[.Playground.Interop.COM.NET2Caller\bin\Debug\"));

            // create domain
            AppDomain net2CallerDomain = AppDomain.CreateDomain(
                "TestNET2Caller",
                evidence,
                setup
                );

            try
            {
                var prg = net2CallerDomain.CreateInstanceAndUnwrap(net2Caller.FullName, net2Caller.GetType().FullName) 
                    as _BET_.Playground.Interop.COM.NET2Caller.MainWrapper;

                var result = prg.CallCOM();
            }
            catch(Exception ex)
            {
                if (ex.HResult == -2146233054)
                    Assert.Fail($"Expected Fail 1: {ex.Message}");
                if (ex.HResult == -2147024894)
                    Assert.Fail($"Expected Fail 2: {ex.Message}");
                if (ex.HResult == -2147024773)
                    Assert.Fail($"Expected Fail 3: {ex.Message}");

                Assert.Fail($"Unknown Fail: {ex.Message}");
            }
            finally
            {
                AppDomain.Unload(net2CallerDomain);
            }
        }

        [TestMethod]
        [TestCategory("Interop")]
        public void TestNET2Caller_AppDomain2_CustomDomain2()
        {
            Assembly net2Caller = Assembly.LoadFrom(TestAssemblyRelativePath);

            // just verify that it is solid xml
            XmlDocument xml = new XmlDocument();
            xml.LoadXml(NET2Config);

            AppDomainSetup setup = new AppDomainSetup()
            {
                PrivateBinPath = net2Caller.CodeBase,
                ApplicationBase = net2Caller.CodeBase,
                ApplicationName = "TestNET2Caller",
                SandboxInterop = true,
                ShadowCopyFiles = Boolean.TrueString,
                TargetFrameworkName = ".NETFramework,Version=v2.0",
                ConfigurationFile = xml.InnerXml
            };

            System.Security.Policy.Evidence evidence = new System.Security.Policy.Evidence(AppDomain.CurrentDomain.Evidence);
            evidence.AddAssemblyEvidence(new System.Security.Policy.ApplicationDirectory(@"..\..\..\]BET[.Playground.Interop.COM.NET2Caller\bin\Debug\"));

            // create domain
            AppDomain net2CallerDomain = AppDomain.CreateDomain(
                "TestNET2Caller",
                evidence,
                setup
                );

            try
            {
                var prg = net2CallerDomain.CreateInstanceAndUnwrap(net2Caller.FullName, net2Caller.GetType().FullName)
                    as _BET_.Playground.Interop.COM.NET2Caller.MainWrapper;

                var result = prg.CallCOM();
            }
            catch (Exception ex)
            {
                if (ex.HResult == -2146233054)
                    Assert.Fail($"Expected Fail 1: {ex.Message}");
                if (ex.HResult == -2147024894)
                    Assert.Fail($"Expected Fail 2: {ex.Message}");
                if (ex.HResult == -2147024773)
                    Assert.Fail($"Expected Fail 3: {ex.Message}");

                Assert.Fail($"Unknown Fail: {ex.Message}");
            }
            finally
            {
                AppDomain.Unload(net2CallerDomain);
            }
        }

        protected static Assembly AssemblyResolve(object sender, ResolveEventArgs args)
        {
            Assembly net2Caller = Assembly.LoadFrom(TestAssemblyRelativePath);
            var name = net2Caller.GetReferencedAssemblies().FirstOrDefault(a => a.FullName == args.Name);

            return Assembly.Load(name);
        }

        [TestMethod]
        [TestCategory("Interop")]
        public void TestNET2Caller_AppDomain2_CustomDomain3()
        {
            Assembly net2Caller = Assembly.LoadFrom(TestAssemblyRelativePath);

            // just verify that it is solid xml
            XmlDocument xml = new XmlDocument();
            xml.LoadXml(NET2Config);

            AppDomainSetup setup = new AppDomainSetup()
            {
                ApplicationBase = AppDomain.CurrentDomain.SetupInformation.ApplicationBase,
                ApplicationName = "TestNET2Caller",
                SandboxInterop = true,
                ShadowCopyFiles = Boolean.TrueString,
                TargetFrameworkName = ".NETFramework,Version=v2.0",
                ConfigurationFile = xml.InnerXml
            };

            System.Security.Policy.Evidence evidence = new System.Security.Policy.Evidence(AppDomain.CurrentDomain.Evidence);
            evidence.AddAssemblyEvidence(new System.Security.Policy.ApplicationDirectory(@"..\..\..\]BET[.Playground.Interop.COM.NET2Caller\bin\Debug\"));

            // create domain
            AppDomain net2CallerDomain = AppDomain.CreateDomain(
                "TestNET2Caller",
                evidence,
                setup
                );

            try
            {
                var prg = net2CallerDomain.CreateInstanceAndUnwrap(net2Caller.FullName, net2Caller.GetType().FullName)
                    as _BET_.Playground.Interop.COM.NET2Caller.MainWrapper;

                var result = prg.CallCOM();
            }
            catch (Exception ex)
            {
                if (ex.HResult == -2146233054)
                    Assert.Fail($"Expected Fail 1: {ex.Message}");
                if (ex.HResult == -2147024894)
                    Assert.Fail($"Expected Fail 2: {ex.Message}");
                if (ex.HResult == -2147024773)
                    Assert.Fail($"Expected Fail 3: {ex.Message}");

                Assert.Fail($"Unknown Fail: {ex.Message}");
            }
            finally
            {
                AppDomain.Unload(net2CallerDomain);
            }
        }
        
        [TestMethod]
        [TestCategory("Interop")]
        public void TestNET2Caller_AppDomain2_CustomDomain4()
        {
            Assembly net2Caller = Assembly.LoadFrom(TestAssemblyRelativePath);

            AppDomainSetup setup = new AppDomainSetup()
            {
                ApplicationBase = AppDomain.CurrentDomain.SetupInformation.ApplicationBase,
                ApplicationName = "TestNET2Caller",
                SandboxInterop = true,
                ShadowCopyFiles = Boolean.TrueString,
                TargetFrameworkName = ".NETFramework,Version=v2.0"
            };

            System.Security.Policy.Evidence evidence = new System.Security.Policy.Evidence(AppDomain.CurrentDomain.Evidence);
            evidence.AddAssemblyEvidence(new System.Security.Policy.ApplicationDirectory(@"..\..\..\]BET[.Playground.Interop.COM.NET2Caller\bin\Debug\"));

            // create domain
            AppDomain net2CallerDomain = AppDomain.CreateDomain(
                "TestNET2Caller",
                evidence,
                setup
                );

            net2CallerDomain.AssemblyResolve += AssemblyResolve;

            try
            {
                var prg = net2CallerDomain.CreateInstanceAndUnwrap(net2Caller.FullName, net2Caller.GetType().FullName)
                    as _BET_.Playground.Interop.COM.NET2Caller.MainWrapper;

                var result = prg.CallCOM();
            }
            catch (Exception ex)
            {
                if (ex.HResult == -2146233054)
                    Assert.Fail($"Expected Fail 1: {ex.Message}");
                if (ex.HResult == -2147024894)
                    Assert.Fail($"Expected Fail 2: {ex.Message}");
                if (ex.HResult == -2147024773)
                    Assert.Fail($"Expected Fail 3: {ex.Message}");

                Assert.Fail($"Unknown Fail: {ex.Message}");
            }
            finally
            {
                AppDomain.Unload(net2CallerDomain);
            }
        }

        [TestMethod]
        [TestCategory("Interop")]
        public void TestNET2Caller_AppDomain3_CreateCom1()
        {
            Assembly net2Caller = Assembly.LoadFrom(TestAssemblyRelativePath);

            AppDomainSetup setup = new AppDomainSetup()
            {
                PrivateBinPath = net2Caller.CodeBase,
                ApplicationBase = net2Caller.CodeBase,
                ApplicationName = "TestNET2Caller",
                SandboxInterop = true,
                ShadowCopyFiles = Boolean.TrueString,
                TargetFrameworkName = ".NETFramework,Version=v2.0"
            };

            System.Security.Policy.Evidence evidence = new System.Security.Policy.Evidence(AppDomain.CurrentDomain.Evidence);
            evidence.AddAssemblyEvidence(new System.Security.Policy.ApplicationDirectory(@"..\..\..\]BET[.Playground.Interop.COM.NET2Caller\bin\Debug\"));

            // create domain
            AppDomain net2CallerDomain = AppDomain.CreateDomain(
                "TestNET4Caller",
                evidence,
                setup
                );

            try
            {
                var handle = net2CallerDomain.CreateComInstanceFrom(net2Caller.ManifestModule.FullyQualifiedName, net2Caller.GetType().FullName);
                var prg = handle.Unwrap() as _BET_.Playground.Interop.COM.NET2Caller.MainWrapper;
                var result = prg.CallCOM();

                Assert.AreEqual<string>(_BET_.Playground.Interop.COM.NET2Caller.MainWrapper.ErrorMsg, result); // fail
            }
            catch (Exception ex)
            {
                if (ex.HResult == -2146233054)
                    Assert.Fail($"Expected Fail 1: {ex.Message}");
                if (ex.HResult == -2147024894)
                    Assert.Fail($"Expected Fail 2: {ex.Message}");
                if (ex.HResult == -2147024773)
                    Assert.Fail($"Expected Fail 3: {ex.Message}");

                Assert.Fail($"Unknown Fail: {ex.Message}");
            }
            finally
            {
                AppDomain.Unload(net2CallerDomain);
            }
        }

        [TestMethod]
        [TestCategory("Interop")]
        public void TestNET2Caller_AppDomain3_CreateCom2()
        {
            Assembly net2Caller = Assembly.LoadFrom(TestAssemblyRelativePath);

            // just verify that it is solid xml
            XmlDocument xml = new XmlDocument();
            xml.LoadXml(NET2Config);

            AppDomainSetup setup = new AppDomainSetup()
            {
                PrivateBinPath = net2Caller.CodeBase,
                ApplicationBase = AppDomain.CurrentDomain.SetupInformation.ApplicationBase, //net2Caller.CodeBase,
                ApplicationName = "TestNET2Caller",
                SandboxInterop = true,
                ShadowCopyFiles = Boolean.TrueString,
                TargetFrameworkName = ".NETFramework,Version=v2.0",
                ConfigurationFile = xml.InnerXml
            };

            System.Security.Policy.Evidence evidence = new System.Security.Policy.Evidence(AppDomain.CurrentDomain.Evidence);
            evidence.AddAssemblyEvidence(new System.Security.Policy.ApplicationDirectory(@"..\..\..\]BET[.Playground.Interop.COM.NET2Caller\bin\Debug\"));

            // create domain
            AppDomain net2CallerDomain = AppDomain.CreateDomain(
                "TestNET4Caller",
                evidence,
                setup
                );

            net2CallerDomain.AssemblyResolve += AssemblyResolve;

            try
            {
                var handle = net2CallerDomain.CreateComInstanceFrom(net2Caller.ManifestModule.FullyQualifiedName, net2Caller.GetType().FullName);
                var prg = handle.Unwrap() as _BET_.Playground.Interop.COM.NET2Caller.MainWrapper;
                var result = prg.CallCOM();

                Assert.AreEqual<string>(_BET_.Playground.Interop.COM.NET2Caller.MainWrapper.ErrorMsg, result); // fail
            }
            catch (Exception ex)
            {
                if (ex.HResult == -2146233054)
                    Assert.Fail($"Expected Fail 1: {ex.Message}");
                if (ex.HResult == -2147024894)
                    Assert.Fail($"Expected Fail 2: {ex.Message}");
                if (ex.HResult == -2147024773)
                    Assert.Fail($"Expected Fail 3: {ex.Message}");

                Assert.Fail($"Unknown Fail: {ex.Message}");
            }
            finally
            {
                AppDomain.Unload(net2CallerDomain);
            }
        }
    }
}
