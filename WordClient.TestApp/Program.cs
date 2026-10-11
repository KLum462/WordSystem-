using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace WordClient.TestApp
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            // Resolve office.dll assembly dynamically if requested
            AppDomain.CurrentDomain.AssemblyResolve += (sender, args) =>
            {
                var assemblyName = new AssemblyName(args.Name).Name;
                if (string.Equals(assemblyName, "office", StringComparison.OrdinalIgnoreCase))
                {
                    var localPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "office.dll");
                    if (File.Exists(localPath))
                    {
                        return Assembly.LoadFrom(localPath);
                    }

                    var gacPath = @"C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c\OFFICE.DLL";
                    if (File.Exists(gacPath))
                    {
                        return Assembly.LoadFrom(gacPath);
                    }
                }
                return null;
            };

            ApplicationConfiguration.Initialize();
            Application.Run(new Form1());
        }    
    }
}