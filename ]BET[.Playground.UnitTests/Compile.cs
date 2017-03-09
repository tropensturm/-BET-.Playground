using System;
using System.Text;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using _BET_.Playground.Compile;
using Microsoft.CodeAnalysis.CSharp;
using System.Reflection;
using System.IO;

namespace _BET_.Playground.UnitTests
{
    /// <summary>
    /// Summary description for Compile
    /// </summary>
    [TestClass]
    public class Compile
    {
        private TestContext testContextInstance;

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }


        //
        // You can use the following additional attributes as you write your tests:
        //
        // Use ClassInitialize to run code before running the first test in the class
        [ClassInitialize()]
        public static void MyClassInitialize(TestContext testContext)
        {
#if DEBUG
            System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("en-US");
#endif
        }

        #region Additional test attributes
        //
        // Use ClassCleanup to run code after all tests in a class have run
        // [ClassCleanup()]
        // public static void MyClassCleanup() { }
        //
        // Use TestInitialize to run code before running each test 
        // [TestInitialize()]
        // public void MyTestInitialize() { }
        //
        // Use TestCleanup to run code after each test has run
        // [TestCleanup()]
        // public void MyTestCleanup() { }
        //
        #endregion

        [TestMethod]
        public void TestCompile()
        {
            var cs1 = Helper.GetRessourceAsString(1, "Program.cs");
            var man = Helper.GetRessourceAsByteArray(1, "app.manifest");

            var roslyn = new _BET_.Playground.Compile.Roslyn();
            roslyn.TargetFrameworkVersion = Microsoft.Build.Utilities.TargetDotNetFrameworkVersion.Version45;
            roslyn.References = new[] {
                "mscorlib",
                "System",
                "System.Core",
                "Microsoft.CSharp"};
            roslyn.Usings = new[] {
                "System",
                "System.Reflection" };

            bool isCompiled = false;
            Assembly assembly = null;

            string error = "";
            var ok = roslyn.Verify(out error);
            Assert.AreEqual<bool>(true, ok);

            using (var assemblyStream = roslyn.Compile(out isCompiled, "NET45Caller", new[] { cs1 }, man, CSharpParseOptions.Default.WithLanguageVersion(LanguageVersion.CSharp4)))
            {
                Assert.AreEqual<bool>(true, isCompiled);

                assembly = Assembly.Load(assemblyStream.ToArray());
            }

            using (var output = Console.Out)
            using (var writer = new StringWriter())
            {
                Console.SetOut(writer);

                assembly.EntryPoint.Invoke(null, new object[] { new string[] { "UT" } });

                Console.SetOut(output);
                var s = writer.ToString();
                s = s.Replace(System.Environment.NewLine, string.Empty);

                Assert.AreEqual<string>("type not found, check your manifest!", s); // expected fail
                Assert.Inconclusive("does not apply manifest");
            }
        }
    }
}
