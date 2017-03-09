using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.Text;
using System.IO;
using Microsoft.Build.Utilities;

namespace _BET_.Playground.Compile
{
    public class Roslyn
    {
        private string runtimePath;
        private IEnumerable<MetadataReference> referencedAssemblies;
        private CSharpCompilationOptions compilationOptions;

        public IEnumerable<string> Usings { get; set; }
        public IEnumerable<string> References { get; set; }
        public TargetDotNetFrameworkVersion TargetFrameworkVersion { get; set; }
        private OutputKind Output { get; set; } = OutputKind.ConsoleApplication;

        public bool Verify(out string error)
        {
            error = "Usings Fail";
            if (null == Usings || Usings.Count() < 1)
                return false;

            error = "References Fail";
            if (null == References || References.Count() < 1)
                return false;

            try
            {
                runtimePath = $"{ToolLocationHelper.GetPathToDotNetFramework(TargetFrameworkVersion)}\\{{0}}.dll";

                //    referencedAssemblies = new[]
                //{
                //    MetadataReference.CreateFromFile(string.Format(runtimePath, "System.Core")),
                //    MetadataReference.CreateFromFile(string.Format(runtimePath, "System"))
                //};
                referencedAssemblies = References.
                    Select(
                    x => MetadataReference.CreateFromFile(string.Format(runtimePath, x))
                    ).ToArray();

                compilationOptions = new CSharpCompilationOptions(Output).
                    WithOverflowChecks(true).
                    WithOptimizationLevel(OptimizationLevel.Release)
                    .WithUsings(Usings);
            }
            catch(Exception ex)
            {
                error = ex.Message;
                return false;
            }

            error = string.Empty;
            return true;
        }

        private SyntaxTree Parse(string text, string filename = "", CSharpParseOptions options = null)
        {
            var stringText = SourceText.From(text, Encoding.UTF8);
            return SyntaxFactory.ParseSyntaxTree(stringText, options, filename);
        }

        // CSharpParseOptions.Default.WithLanguageVersion(LanguageVersion.CSharp2)
        public MemoryStream Compile(out bool success, string name, string[] sourcefiles, byte[] manifest, CSharpParseOptions options = null)
        {
            success = false;
            string error = string.Empty;

            if (!Verify(out error))
                return null;

            SyntaxTree[] tree = sourcefiles.
                Select(
                x => Parse(x, string.Empty, options)
                ).ToArray();

            var compilation = CSharpCompilation.Create(name, tree, referencedAssemblies, compilationOptions);

            using (var output = new MemoryStream())
            using (var ms = new MemoryStream())
            {
                ms.Write(manifest, 0, manifest.Length);
                using (var win32res = compilation.CreateDefaultWin32Resources(true, false, ms, null))
                {

                    var result = compilation.Emit(output, win32Resources: win32res);
                    success = result.Success;
                    output.Position = 0;
                    return output;
                }
            }
        }

        //private static readonly IEnumerable<MetadataReference> DefaultReferences =
        //    new[]
        //    {
        //        MetadataReference.CreateFromFile(string.Format(runtimePath, "mscorlib")),
        //        MetadataReference.CreateFromFile(string.Format(runtimePath, "System"))
        //    };
    }
}
