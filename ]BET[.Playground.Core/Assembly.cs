﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace _BET_.Playground.Core
{
    public static class AssemblyLocator
    {
        static Dictionary<string, Assembly> assemblies;

        public static void Init()
        {
            assemblies = new Dictionary<string, Assembly>();
            AppDomain.CurrentDomain.AssemblyLoad += new AssemblyLoadEventHandler(CurrentDomain_AssemblyLoad);
            AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(CurrentDomain_AssemblyResolve);
        }

        static Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            Assembly assembly = null;
            assemblies.TryGetValue(args.Name, out assembly);
            return assembly;
        }

        static void CurrentDomain_AssemblyLoad(object sender, AssemblyLoadEventArgs args)
        {
            Assembly assembly = args.LoadedAssembly;
            assemblies[assembly.FullName] = assembly;
        }
    }
}
