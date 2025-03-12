using Microsoft.Extensions.Logging;
using System;
using System.IO;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using SHDocVw;

namespace Listary.FileAppPlugin.TablacusExplorer
{
    public class TEPlugin : IFileAppPlugin
    {
        private IFileAppPluginHost _host;
        // 这个属性表示此应用打开的文件夹是否与其他应用共享
        bool IFileAppPlugin.IsOpenedFolderProvider => true;

        bool IFileAppPlugin.IsQuickSwitchTarget => false;

        bool IFileAppPlugin.IsSharedAcrossApplications => false;

        SearchBarType IFileAppPlugin.SearchBarType => SearchBarType.Fixed;


        IFileWindow IFileAppPlugin.BindFileWindow(IntPtr hWnd)
        {
            string className = Win32Utils.GetClassName(hWnd);
            if (className == "TablacusExplorer")
            {
                string processName = Path.GetFileName(Win32Utils.GetProcessPathFromHwnd(hWnd)).ToLowerInvariant();
                if (processName == "te32.exe" || processName == "te64.exe")
                {
                    try
                    {
                        return new TEFileWindow(_host, hWnd);
                    }
                    catch (COMException e)
                    {
                        _host.Logger.LogError($"COM exception: {e}");
                    }
                    catch (TimeoutException e)
                    {
                        _host.Logger.LogWarning($"UIA timeout: {e}");
                    }
                    catch (Exception e)
                    {
                        _host.Logger.LogError($"Failed to bind window: {e}");
                    }
                }
            }
            
            return null;
        }

        public Task<bool> Initialize(IFileAppPluginHost host)
        {
            _host = host;
            return Task.FromResult(true);
        }
    }
    public class TEFileWindow : IFileWindow
    {
        private readonly IFileAppPluginHost _host;

        public IntPtr Handle { get; }

        public TEFileWindow(IFileAppPluginHost host, IntPtr handle)
        {
            Handle = handle;
            _host = host;
        }
        public Task<IFileTab> GetCurrentTab()
        {
            var fTab = new TEFileTab(_host, Handle);
            return Task.FromResult<IFileTab>(fTab);
        }
    }
    public class TEFileTab : IFileTab, IGetFolder, IOpenFolder
    {
        private readonly IFileAppPluginHost _host;
        private readonly IntPtr _hWnd;
        private readonly InternetExplorer _ie = null;

        private const uint WM_GETOBJECT = 0x003D;
        private const int OBJID_CLIENT = 0x00000001;
        public TEFileTab(IFileAppPluginHost host, IntPtr hWnd)
        {
            _host = host;
            _hWnd = hWnd;

            ShellWindows shellWindows = new ShellWindows();
            foreach (InternetExplorer ie in shellWindows)
            {
                string processName = Path.GetFileName(ie.FullName).ToLowerInvariant();
                if (processName == "te32.exe" || processName == "te64.exe")
                {
                    _ie = ie;
                    break;
                }

            }
        }
        public Task<string> GetCurrentFolder()
        {
            if (_ie != null)
            {
                var tParentWnd = _ie.Document.GetType().GetProperty("parentWindow");
                if(tParentWnd != null)
                {
                    var parent = tParentWnd.GetValue(_ie.Document);
                    if (parent != null)
                    {
                        parent.execScript("document._tophwnd = GetTopWindow(WebBrowser.hwnd)");
                        if (new IntPtr(_ie.Document._tophwnd) == _hWnd)
                        {
                            string dir = _ie.Document.F.addressbar.value;
                            return Task.FromResult(dir);
                        }
                    }
                }
            }
            return Task.FromResult(string.Empty);
        }

        public Task<bool> OpenFolder(string path)
        {
            return Task.FromResult(false);
        }
    }
}
