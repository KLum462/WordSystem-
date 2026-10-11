using System;
using System.Runtime.InteropServices;
using WordClient.Core.Services;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace WordVstoAddIn
{
    public class ThisAddIn
    {
        public Word.Application Application { get; set; } = null!;

        public void ThisAddIn_Startup(object sender, EventArgs e)
        {
            ClientLogger.LogInfo("Word 2016 VSTO Add-In started.");
        }

        public void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            ClientLogger.LogInfo("Word 2016 VSTO Add-In shutting down.");

            try
            {
                if (Application != null && Marshal.IsComObject(Application))
                {
                    WordDocumentService.ReleaseComObject(Application);
                }
            }
            catch
            {
                // Suppress COM release errors during shutdown
            }
            finally
            {
                // Force Garbage Collection to ensure all COM wrappers are finalized
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
        }

        protected virtual Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon();
        }
    }

    public static class Globals
    {
        private static ThisAddIn? _thisAddIn;

        public static ThisAddIn ThisAddIn
        {
            get
            {
                if (_thisAddIn == null)
                {
                    _thisAddIn = new ThisAddIn();
                    try
                    {
                        var service = new WordDocumentService();
                        _thisAddIn.Application = service.GetOrCreateApplication();
                    }
                    catch
                    {
                        // Deferred until Word connection
                    }
                }
                return _thisAddIn;
            }
            internal set => _thisAddIn = value;
        }
    }
}
