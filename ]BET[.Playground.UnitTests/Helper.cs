using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace _BET_.Playground.UnitTests
{
    internal static class Helper
    {
        /// <param name="parameter">example: new string[] { "UT" }</param>
        public static void InvokeMainFromApp(Type type, string[] parameter)
        {
            Assembly assembly = Assembly.GetAssembly(type);
            assembly.EntryPoint.Invoke(null, new object[] { parameter });
        }

        public static object InvokeMethodFromCOMObjectByAssembly(string assemblyPath, string testObject, string testMethod, object[] methodParameters)
        {
            var handler = Activator.CreateComInstanceFrom(assemblyPath, testObject);
            var instance = handler.Unwrap();

            var method = instance.GetType().GetMethod(testMethod);
            return method.Invoke(instance, methodParameters);
        }

        public static object InvokeMethodFromObjectByAssembly(Assembly assembly, string testObject, string testMethod, object[] methodParameters)
        {
            var objectType = assembly.GetType(testObject);
            var method = objectType.GetMethod(testMethod);

            var instance = Activator.CreateInstance(objectType);
            return method.Invoke(instance, methodParameters);
        }

        public static object InvokeMethodFromObjectByType(Type type, string testObject, string testMethod, object[] methodParameters = null)
        {
            Assembly assembly = Assembly.GetAssembly(type);
            return InvokeMethodFromObjectByAssembly(assembly, testObject, testMethod, methodParameters);
        }

        public static object InvokeMethodFromObjectByPath(string assemblyPath, string testObject, string testMethod, object[] methodParameters = null)
        {
            Assembly assembly = Assembly.LoadFrom(assemblyPath);
            return InvokeMethodFromObjectByAssembly(assembly, testObject, testMethod, methodParameters);
        }

        private static readonly string AssemblyRessource = "_BET_.Playground.UnitTests.Ressources";

        public static string GetAssemblyRessource(int resNo, string value)
        {
            return $"{AssemblyRessource}{resNo}.{value}";
        }

        public static string GetRessourceAsString(int resNo, string ressourceName)
        {
            var assembly = Assembly.GetExecutingAssembly();

            using (Stream stream = assembly.GetManifestResourceStream(GetAssemblyRessource(resNo, ressourceName)))
            using (StreamReader reader = new StreamReader(stream, true))
            {
                return reader.ReadToEnd();
            }
        }

        public static byte[] GetRessourceAsByteArray(int resNo, string ressourceName)
        {
            var assembly = Assembly.GetExecutingAssembly();

            using (var stream = assembly.GetManifestResourceStream(GetAssemblyRessource(resNo, ressourceName)))
            using (var mstream = new MemoryStream())
            {
                stream.CopyTo(mstream);
                return mstream.ToArray();
            }
        }
    }

    /// <summary>
    /// I do not get it but it is required or the CreateInstanceAndUnwrap call will crash no matter what I try
    /// this is creating a remoting object
    /// 
    /// my best guess is that CreateInstanceAndUnwrap is not loading all referenced assemblies, with the proxy
    /// it will load it completley because it is within the same assembly -> than we can load completely with loadfrom
    /// </summary>
    internal class Proxy : MarshalByRefObject
    {
        public Assembly GetAssembly(string assemblyPath)
        {
            try
            {
                return Assembly.LoadFrom(assemblyPath);
            }
            catch (Exception)
            {
                return null;
            }
        }
    }
}
