// OneFinder.AddIn - OneNote COM addin (.NET Framework 4.8)
// Load chain: OneNote.exe -> mscoree.dll -> CLR v4 -> this DLL
// Registry: HKCU\SOFTWARE\Microsoft\Office\OneNote\AddIns\OneFinder.AddIn

using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using Extensibility;
using Microsoft.Office.Core;

namespace OneFinder.AddIn
{
    [ComVisible(true)]
    [Guid(AddinClsid)]
    [ProgId(AddinProgId)]
    public class AddIn : IDTExtensibility2, IRibbonExtensibility
    {
        public const string AddinClsid  = "6B29FC40-CA47-1067-B31D-00DD010662DA";
        public const string AddinProgId = "OneFinder.AddIn";

        private IRibbonUI? _ribbon;

        private static string GetInstallDir()
        {
            try
            {
                var codeBase = Assembly.GetExecutingAssembly().CodeBase;
                if (!string.IsNullOrEmpty(codeBase))
                {
                    var uri = new Uri(codeBase);
                    if (uri.IsFile)
                        return Path.GetDirectoryName(uri.LocalPath) ?? AppDomain.CurrentDomain.BaseDirectory;
                }
            }
            catch (Exception ex)
            {
                Log($"GetInstallDir (CodeBase): {ex.Message}");
            }

            try
            {
                var loc = Assembly.GetExecutingAssembly().Location;
                if (!string.IsNullOrEmpty(loc))
                    return Path.GetDirectoryName(loc) ?? AppDomain.CurrentDomain.BaseDirectory;
            }
            catch (Exception ex)
            {
                Log($"GetInstallDir (Location): {ex.Message}");
            }

            return AppDomain.CurrentDomain.BaseDirectory;
        }

        private static void Log(string msg)
        {
            try
            {
                var path = Path.Combine(Path.GetTempPath(), "OneFinder.AddIn.log");
                File.AppendAllText(path, $"[{DateTime.Now:HH:mm:ss}] {msg}{Environment.NewLine}");
            }
            catch (Exception ex)
            {
                // Silent fallback - can't log if logging fails
                System.Diagnostics.Debug.WriteLine($"Log write failed: {ex.Message}");
            }
        }

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            Log($"OnConnection ConnectMode={ConnectMode}");
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            Log($"OnDisconnection RemoveMode={RemoveMode}");
        }

        public void OnAddInsUpdate(ref Array custom) { }

        public void OnStartupComplete(ref Array custom)
        {
            Log("OnStartupComplete");
        }

        public void OnBeginShutdown(ref Array custom)
        {
            Log("OnBeginShutdown");
            try
            {
                if (EventWaitHandle.TryOpenExisting("Local\\OneFinder-OneNoteShutdown", out var handle))
                {
                    using (handle)
                        handle.Set();
                }
            }
            catch (Exception ex)
            {
                Log($"OnBeginShutdown: {ex.Message}");
            }
        }

        public string GetCustomUI(string RibbonID)
        {
            Log($"GetCustomUI RibbonID={RibbonID}");
            try
            {
                var asm = Assembly.GetExecutingAssembly();
                using (var stream = asm.GetManifestResourceStream("OneFinder.AddIn.Ribbon.xml"))
                {
                    if (stream == null)
                    {
                        Log("ERROR: embedded Ribbon.xml not found");
                        return string.Empty;
                    }
                    using (var reader = new StreamReader(stream))
                    {
                        var xml = reader.ReadToEnd();
                        Log($"GetCustomUI OK, {xml.Length} chars");
                        return xml;
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"GetCustomUI: {ex.Message}");
                return string.Empty;
            }
        }

        public void RibbonLoaded(IRibbonUI ribbon)
        {
            _ribbon = ribbon;
            Log("RibbonLoaded OK");
        }

        public void OnSearchClick(IRibbonControl control)
        {
            Log("OnSearchClick fired");
            try
            {
                var dir = GetInstallDir();
                var exe = Path.Combine(dir, "OneFinder.exe");
                Log($"exe path={exe} exists={File.Exists(exe)}");

                if (!File.Exists(exe))
                {
                    Log($"ERROR: exe not found at {exe}");
                    return;
                }

                var procName = Path.GetFileNameWithoutExtension(exe);
                var procs = Process.GetProcessesByName(procName);
                var hwnd = procs.Length > 0 ? procs[0].MainWindowHandle : IntPtr.Zero;
                
                if (hwnd != IntPtr.Zero)
                {
                    Log("already running - bring to front");
                    NativeMethods.ShowWindow(hwnd, 9);
                    NativeMethods.SetForegroundWindow(hwnd);
                }
                else
                {
                    // Kill stale processes before starting new one
                    foreach (var stale in procs)
                    {
                        try
                        {
                            Log($"killing stale pid={stale.Id}");
                            stale.Kill();
                            stale.WaitForExit(1000); // Wait up to 1s for process to exit
                        }
                        catch (Exception ex)
                        {
                            Log($"kill stale process: {ex.Message}");
                        }
                        finally
                        {
                            stale.Dispose();
                        }
                    }

                    Log($"starting {exe}");
                    using (var p = Process.Start(new ProcessStartInfo
                    {
                        FileName = exe,
                        UseShellExecute = true,
                        WorkingDirectory = dir
                    }))
                    {
                        Log($"started pid={p?.Id ?? -1}");
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"OnSearchClick: {ex.Message}\n{ex.StackTrace}");
            }
        }
    }

    internal static class NativeMethods
    {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        internal static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    }
}