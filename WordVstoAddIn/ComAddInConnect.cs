using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using WordClient.Core.Services;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace WordVstoAddIn
{
    [ComImport]
    [Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDTExtensibility2
    {
        void OnConnection(
            [MarshalAs(UnmanagedType.IDispatch)] object Application,
            int ConnectMode,
            [MarshalAs(UnmanagedType.IDispatch)] object AddInInst,
            [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        void OnDisconnection(
            int RemoveMode,
            [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        void OnAddInsUpdate(
            [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        void OnStartupComplete(
            [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        void OnBeginShutdown(
            [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
    }

    /// <summary>
    /// Native Microsoft Word COM Add-In entry point exposing the custom Ribbon ("Internship Word Tools").
    /// Register via Register-WordAddIn.ps1 to load directly inside Microsoft Word without VSTO click-once certificates.
    /// </summary>
    [ComVisible(true)]
    [Guid("8F5E3B9A-6D4C-4B28-9A1D-2E7B8C4F1092")]
    [ProgId("WordVstoAddIn.Connect")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class ComAddInConnect : IDTExtensibility2, Office.IRibbonExtensibility
    {
        private static readonly object ResolveLock = new object();
        private static bool _resolverRegistered;

        [ThreadStatic]
        private static HashSet<string>? _resolvingAssemblies;

        static ComAddInConnect()
        {
            EnsureAssemblyResolverRegistered();
        }

        private readonly Ribbon _ribbon = new Ribbon();
        private ThisAddIn? _addIn;

        /// <summary>
        /// Resolves NuGet dependencies (such as System.Runtime.CompilerServices.Unsafe) from the add-in
        /// directory when hosted inside WINWORD.EXE, which lacks binding redirects for them.
        /// </summary>
        private static void EnsureAssemblyResolverRegistered()
        {
            lock (ResolveLock)
            {
                if (_resolverRegistered)
                {
                    return;
                }

                AppDomain.CurrentDomain.AssemblyResolve += OnAssemblyResolve;
                _resolverRegistered = true;
            }
        }

        private static Assembly? OnAssemblyResolve(object? sender, ResolveEventArgs args)
        {
            var simpleName = new AssemblyName(args.Name).Name;
            if (string.IsNullOrEmpty(simpleName) ||
                simpleName.EndsWith(".resources", StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            _resolvingAssemblies ??= new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (!_resolvingAssemblies.Add(simpleName))
            {
                return null;
            }

            try
            {
                foreach (var loaded in AppDomain.CurrentDomain.GetAssemblies())
                {
                    if (string.Equals(loaded.GetName().Name, simpleName, StringComparison.OrdinalIgnoreCase))
                    {
                        return loaded;
                    }
                }

                var addInDir = Path.GetDirectoryName(typeof(ComAddInConnect).Assembly.Location);
                if (string.IsNullOrEmpty(addInDir))
                {
                    return null;
                }

                var candidatePath = Path.Combine(addInDir, $"{simpleName}.dll");
                return File.Exists(candidatePath) ? Assembly.LoadFrom(candidatePath) : null;
            }
            finally
            {
                _resolvingAssemblies.Remove(simpleName);
            }
        }

        public void OnConnection(object Application, int ConnectMode, object AddInInst, ref Array custom)
        {
            try
            {
                _addIn = new ThisAddIn();
                if (Application is Word.Application wordApp)
                {
                    _addIn.Application = wordApp;
                }
                Globals.ThisAddIn = _addIn;
                _addIn.ThisAddIn_Startup(this, EventArgs.Empty);
                ClientLogger.LogInfo("ComAddInConnect connected to Microsoft Word.");
            }
            catch (Exception ex)
            {
                ClientLogger.LogError("ComAddInConnect OnConnection failed.", ex);
            }
        }

        public void OnDisconnection(int RemoveMode, ref Array custom)
        {
            _addIn?.ThisAddIn_Shutdown(this, EventArgs.Empty);
            ClientLogger.LogInfo("ComAddInConnect disconnected from Microsoft Word.");
        }

        public void OnAddInsUpdate(ref Array custom) { }

        public void OnStartupComplete(ref Array custom)
        {
            ClientLogger.LogInfo("Microsoft Word startup complete with WordVstoAddIn loaded.");
        }

        public void OnBeginShutdown(ref Array custom)
        {
            _addIn?.ThisAddIn_Shutdown(this, EventArgs.Empty);
        }

        // IRibbonExtensibility forwarding
        public string GetCustomUI(string ribbonID) => _ribbon.GetCustomUI(ribbonID);

        // Ribbon callback forwarding for Word COM invocation
        public void Ribbon_Load(Office.IRibbonUI ribbonUI) => _ribbon.Ribbon_Load(ribbonUI);
        public void OnAuthenticateClick(Office.IRibbonControl control) => _ribbon.OnAuthenticateClick(control);
        public void OnLoadDocumentClick(Office.IRibbonControl control) => _ribbon.OnLoadDocumentClick(control);
        public void OnSendSelectedTextClick(Office.IRibbonControl control) => _ribbon.OnSendSelectedTextClick(control);
        public void OnSendDocumentClick(Office.IRibbonControl control) => _ribbon.OnSendDocumentClick(control);
        public void OnCheckDocumentClick(Office.IRibbonControl control) => _ribbon.OnCheckDocumentClick(control);
        public void OnFindReplaceClick(Office.IRibbonControl control) => _ribbon.OnFindReplaceClick(control);
        public void OnSaveTemplateClick(Office.IRibbonControl control) => _ribbon.OnSaveTemplateClick(control);
    }
}
