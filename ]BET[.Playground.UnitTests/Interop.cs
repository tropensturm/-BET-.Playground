using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using _BET_.Playground.Interop.COM.NET4Callee;
using System.Reflection;
using _BET_.Playground.NET20Core;
using System.Linq;
using System.Xml;
using System.Collections.Generic;
using System.Diagnostics;

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
        private static readonly string NET2Config = @"<?xml version='1.0' encoding='utf-8'?>
<configuration>  
    <startup>  
        <supportedRuntime version='v2.0.50727' />  
    </startup>  
</configuration>";

        [TestMethod]
        [TestCategory("Interop")]
        public void TestNET2Config()
        {
            // just verify that it is solid xml
            XmlDocument xml = new XmlDocument();
            xml.LoadXml(NET2Config);

            Assert.IsNotNull(xml.InnerXml);
        }

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
        public void TestLoadCorrectAssemblies1()
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
        public void TestLoadCorrectAssemblies2()
        {
            string net2Path = @"C:\Windows\Microsoft.NET\Framework\v2.0.50727\{0}.dll";

            Assembly net2Caller = Assembly.LoadFrom(TestAssemblyRelativePath);
            var assemblyNames = net2Caller.GetReferencedAssemblies();

            int errors = 0;

            foreach (var an in assemblyNames)
            {
                var rA = Assembly.LoadFile(string.Format(net2Path, an.Name));
                if (rA.FullName != an.FullName)
                    errors += 1;
            }

            Assert.AreEqual<int>(0, errors, $"{errors} of {assemblyNames.Length} are loaded in wrong version!");
        }

        [TestMethod]
        [TestCategory("Interop")]
        public void TestLoadCorrectAssemblies3()
        {
            string net2Path = @"C:\Windows\Microsoft.NET\Framework\v2.0.50727\{0}.dll";
            Assembly net2Caller = Assembly.LoadFrom(TestAssemblyRelativePath);
            var assemblyNames = net2Caller.GetReferencedAssemblies();

            int errors = 0;

            foreach (var an in assemblyNames)
            {
                byte[] bytes = System.IO.File.ReadAllBytes(string.Format(net2Path, an.Name));

                var rA = Assembly.Load(bytes);
                if (rA.FullName != an.FullName)
                    errors += 1;
            }

            Assert.AreEqual<int>(0, errors, $"{errors} of {assemblyNames.Length} are loaded in wrong version!");
        }

        [TestMethod]
        [TestCategory("Interop")]
        public void TestLoadCorrectAssemblies4()
        {
            string net2Path = @"..\..\..\]BET[.Playground.Interop.COM.NET2Caller\bin\Debug\{0}.dll";
            Assembly net2Caller = Assembly.LoadFrom(TestAssemblyRelativePath);
            var assemblyNames = net2Caller.GetReferencedAssemblies();

            int errors = 0;

            foreach (var an in assemblyNames)
            {
                byte[] bytes = System.IO.File.ReadAllBytes(string.Format(net2Path, an.Name));

                var rA = Assembly.Load(bytes);
                if (rA.FullName != an.FullName)
                    errors += 1;
            }

            Assert.AreEqual<int>(0, errors, $"{errors} of {assemblyNames.Length} are loaded in wrong version!");
        }

        [TestMethod]
        [TestCategory("Interop")]
        public void TestLoadCorrectAssemblies5()
        {
            string net2Path = @"C:\Windows\Microsoft.NET\Framework\v2.0.50727\{0}.dll";
            Assembly net2Caller = Assembly.LoadFrom(TestAssemblyRelativePath);
            var assemblyNames = net2Caller.GetReferencedAssemblies();

            int errors = 0;

            foreach (var an in assemblyNames)
            {
                byte[] bytes = System.IO.File.ReadAllBytes(string.Format(net2Path, an.Name));

                var rA = Assembly.ReflectionOnlyLoadFrom(string.Format(net2Path, an.Name));
                if (rA.FullName != an.FullName)
                    errors += 1;
            }

            Assert.AreEqual<int>(0, errors, $"{errors} of {assemblyNames.Length} are loaded in wrong version!");
            Assert.Inconclusive("loaded, but reflection only");
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

        [TestMethod]
        [TestCategory("Interop")]
        public void TestNET2Caller_AppDomain2_CustomDomain3()
        {
            Assembly net2Caller = Assembly.LoadFrom(TestAssemblyRelativePath);

            AppDomainSetup setup = new AppDomainSetup()
            {
                //ApplicationBase = net2Caller.CodeBase, --> error code 2
                ApplicationBase = AppDomain.CurrentDomain.SetupInformation.ApplicationBase, // --> error code 3
                ApplicationName = "TestNET2Caller",
                SandboxInterop = true,
                ShadowCopyFiles = Boolean.TrueString,
                TargetFrameworkName = ".NETFramework,Version=v2.0", // appdomain is more or less ignoring < 4.5 :(
                ConfigurationFile = NET2Config
            };

            setup.SetConfigurationBytes(System.Text.Encoding.UTF8.GetBytes(NET2Config));

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
                TargetFrameworkName = ".NETFramework,Version=v2.0" // appdomain is more or less ignoring < 4.5 :(
            };

            setup.SetConfigurationBytes(System.Text.Encoding.UTF8.GetBytes(NET2Config));

            System.Security.Policy.Evidence evidence = new System.Security.Policy.Evidence(AppDomain.CurrentDomain.Evidence);
            evidence.AddAssemblyEvidence(new System.Security.Policy.ApplicationDirectory(@"..\..\..\]BET[.Playground.Interop.COM.NET2Caller\bin\Debug\"));

            // create domain
            AppDomain net2CallerDomain = AppDomain.CreateDomain(
                "TestNET2Caller",
                evidence,
                setup
                );

            AssemblyLocator.Init(net2CallerDomain);

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
                ApplicationBase = AppDomain.CurrentDomain.SetupInformation.ApplicationBase,
                //ApplicationBase = net2Caller.CodeBase, // -> AssemblyLocator will crash
                ApplicationName = "TestNET2Caller",
                SandboxInterop = true,
                ShadowCopyFiles = Boolean.TrueString,
                TargetFrameworkName = ".NETFramework,Version=v2.0",
            };

            setup.SetConfigurationBytes(System.Text.Encoding.UTF8.GetBytes(NET2Config));

            System.Security.Policy.Evidence evidence = new System.Security.Policy.Evidence(AppDomain.CurrentDomain.Evidence);
            evidence.AddAssemblyEvidence(new System.Security.Policy.ApplicationDirectory(@"..\..\..\]BET[.Playground.Interop.COM.NET2Caller\bin\Debug\"));

            // create domain
            AppDomain net2CallerDomain = AppDomain.CreateDomain(
                "TestNET4Caller",
                evidence,
                setup
                );

            AssemblyLocator.Init(net2CallerDomain);

            try
            {
                var handle = net2CallerDomain.CreateComInstanceFrom(net2Caller.ManifestModule.FullyQualifiedName, net2Caller.GetType().FullName);
                var prg = handle.Unwrap() as _BET_.Playground.Interop.COM.NET2Caller.MainWrapper;
                var result = prg.CallCOM();

                Assert.AreEqual<string>(_BET_.Playground.Interop.COM.NET2Caller.MainWrapper.ErrorMsg, result); // fail
            }
            catch (Exception ex)
            {
                // Could not load type 'System.Reflection.RuntimeAssembly' from 
                // assembly 'Playground.Interop.COM.NET2Caller, Version=1.0.0.0, 
                // Culture =neutral, PublicKeyToken=null'
                if (ex.HResult == -2146233054)
                    Assert.Fail($"Expected Fail 1: {ex.Message}");
                // Could not load file or assembly 'Playground.Interop.COM.NET2Caller, 
                // Version =1.0.0.0, Culture=neutral, PublicKeyToken=null' or one of 
                // its dependencies. Das System kann die angegebene Datei nicht finden.
                if (ex.HResult == -2147024894)
                    Assert.Fail($"Expected Fail 2: {ex.Message}");
                // Could not load file or assembly 'Playground.Interop.COM.NET2Caller, 
                // Version =1.0.0.0, Culture=neutral, PublicKeyToken=null' or one of 
                // its dependencies. Die Syntax für den Dateinamen, Verzeichnisnamen 
                // oder die Datenträgerbezeichnung ist falsch.
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
        public void TestNET2Caller_AppDomain4_AssemblyLocator()
        {
            Assembly net2Caller = Assembly.LoadFrom(TestAssemblyRelativePath);

            AppDomainSetup setup = new AppDomainSetup()
            {
                //ApplicationBase = AppDomain.CurrentDomain.SetupInformation.ApplicationBase,
                ApplicationBase = net2Caller.CodeBase, // -> AssemblyLocator will crash
                ApplicationName = "TestNET2Caller",
                SandboxInterop = true,
                ShadowCopyFiles = Boolean.TrueString,
                TargetFrameworkName = ".NETFramework,Version=v2.0", // appdomain is more or less ignoring < 4.5 :(
            };

            setup.SetConfigurationBytes(System.Text.Encoding.UTF8.GetBytes(NET2Config));

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
                AssemblyLocator.Init(net2CallerDomain);
            }
            catch(Exception ex)
            {
                Assert.Fail(ex.Message);

                /*
Could not load file or assembly ']BET[.Playground.Core, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null' or one of its dependencies. Das System kann die angegebene Datei nicht finden.
                
                === Pre-bind state information ===
LOG: DisplayName = ]BET[.Playground.Core, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
 (Fully-specified)

LOG: Appbase = ../]BET[.Playground/]BET[.Playground.Interop.COM.NET2Caller/bin/Debug/Playground.Interop.COM.NET2Caller.exe
LOG: Initial PrivatePath = ../Projects/]BET[.Playground/]BET[.Playground.Interop.COM.NET2Caller/bin/Debug/Playground.Interop.COM.NET2Caller.exe
Calling assembly : (Unknown).
===
LOG: This bind starts in default load context.
LOG: Found application configuration file (\PROGRAM FILES (X86)\MICROSOFT VISUAL STUDIO 14.0\COMMON7\IDE\COMMONEXTENSIONS\MICROSOFT\TESTWINDOW\vstest.executionengine.x86.exe.Config).
LOG: Using application configuration file: \PROGRAM FILES (X86)\MICROSOFT VISUAL STUDIO 14.0\COMMON7\IDE\COMMONEXTENSIONS\MICROSOFT\TESTWINDOW\vstest.executionengine.x86.exe.Config
LOG: Using host configuration file: 
LOG: Using machine configuration file from \Windows\Microsoft.NET\Framework\v4.0.30319\config\machine.config.
LOG: Policy not being applied to reference at this time (private, custom, partial, or location-based assembly bind).
LOG: Attempting download of new URL ../]BET[.Playground/]BET[.Playground.Interop.COM.NET2Caller/bin/Debug/Playground.Interop.COM.NET2Caller.exe/]BET[.Playground.Core.DLL.
LOG: Attempting download of new URL ../]BET[.Playground/]BET[.Playground.Interop.COM.NET2Caller/bin/Debug/Playground.Interop.COM.NET2Caller.exe/]BET[.Playground.Core/]BET[.Playground.Core.DLL.
LOG: Attempting download of new URL ../]BET[.Playground/]BET[.Playground.Interop.COM.NET2Caller/bin/Debug/Playground.Interop.COM.NET2Caller.exe/]BET[.Playground.Core.EXE.
LOG: Attempting download of new URL ../]BET[.Playground/]BET[.Playground.Interop.COM.NET2Caller/bin/Debug/Playground.Interop.COM.NET2Caller.exe/]BET[.Playground.Core/]BET[.Playground.Core.EXE.


   at System.Reflection.RuntimeAssembly._nLoad(AssemblyName fileName, String codeBase, Evidence assemblySecurity, RuntimeAssembly locationHint, StackCrawlMark& stackMark, IntPtr pPrivHostBinder, Boolean throwOnFileNotFound, Boolean forIntrospection, Boolean suppressSecurityChecks)
   at System.Reflection.RuntimeAssembly.nLoad(AssemblyName fileName, String codeBase, Evidence assemblySecurity, RuntimeAssembly locationHint, StackCrawlMark& stackMark, IntPtr pPrivHostBinder, Boolean throwOnFileNotFound, Boolean forIntrospection, Boolean suppressSecurityChecks)
   at System.Reflection.RuntimeAssembly.InternalLoadAssemblyName(AssemblyName assemblyRef, Evidence assemblySecurity, RuntimeAssembly reqAssembly, StackCrawlMark& stackMark, IntPtr pPrivHostBinder, Boolean throwOnFileNotFound, Boolean forIntrospection, Boolean suppressSecurityChecks)
   at System.Reflection.RuntimeAssembly.InternalLoad(String assemblyString, Evidence assemblySecurity, StackCrawlMark& stackMark, IntPtr pPrivHostBinder, Boolean forIntrospection)
   at System.Reflection.RuntimeAssembly.InternalLoad(String assemblyString, Evidence assemblySecurity, StackCrawlMark& stackMark, Boolean forIntrospection)
   at System.Reflection.Assembly.Load(String assemblyString)
   at System.Runtime.Serialization.FormatterServices.LoadAssemblyFromString(String assemblyName)
   at System.Reflection.MemberInfoSerializationHolder..ctor(SerializationInfo info, StreamingContext context)
   at System.AppDomain.add_AssemblyLoad(AssemblyLoadEventHandler value)
   at _BET_.Playground.Core.AssemblyLocator.Init(AppDomain domain) in ..\]BET[.Playground\]BET[.Playground.Core\Assembly.cs:line 14
   at _BET_.Playground.UnitTests.Interop.TestNET2Caller_AppDomain3_CreateCom2() in ..\]BET[.Playground\]BET[.Playground.UnitTests\Interop.cs:line 529

                 */
            }
        }

        [TestMethod]
        [TestCategory("Interop")]
        public void TestNET2Caller_Process1()
        {
            // this will simply work

            var p = new Process();

            p.StartInfo.FileName = TestAssemblyRelativePath;
            p.StartInfo.Arguments = "-UT";
            p.StartInfo.RedirectStandardOutput = true;
            p.StartInfo.UseShellExecute = false; // implicitly required when redirecting output
            p.Start();
            while (!p.StandardOutput.EndOfStream)
            {
                string output = p.StandardOutput.ReadLine();
                Assert.AreEqual<string>("NET Dll Version v4.0.30319", output);
                break;
            }
        }
    }
}
