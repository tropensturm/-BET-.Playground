using System;
using System.Collections.Generic;
using System.Reflection;

namespace _BET_.Playground.NET20Core
{
    public static class AssemblyLocator
    {
        public static Dictionary<string, Assembly> assemblies;

        public static void Init(AppDomain domain)
        {
            assemblies = new Dictionary<string, Assembly>();
            domain.AssemblyLoad += new AssemblyLoadEventHandler(AssemblyLoad);
            domain.AssemblyResolve += new ResolveEventHandler(AssemblyResolve);
        }

        static Assembly AssemblyResolve(object sender, ResolveEventArgs args)
        {
            Assembly assembly = null;
            assemblies.TryGetValue(args.Name, out assembly);
            return assembly;
        }

        static void AssemblyLoad(object sender, AssemblyLoadEventArgs args)
        {
            Assembly assembly = args.LoadedAssembly;
            if (assemblies == null)
                assemblies = new Dictionary<string, Assembly>();
            if (!assemblies.ContainsKey(assembly.FullName))
                assemblies.Add(assembly.FullName, assembly);
        }
    }
}
