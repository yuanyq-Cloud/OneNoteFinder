using System;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace OneFinder
{
    /// <summary>
    /// 在专用 STA 线程上序列化所有 OneNote COM 调用
    /// </summary>
    internal sealed class OneNoteScheduler : IDisposable
    {
        private readonly BlockingCollection<Action> _queue = new();
        private readonly Thread _thread;
        private OneNoteService? _service;
        private bool _disposed;

        private static void Log(string msg) => OneNoteService.Log($"[Scheduler] {msg}");

        public OneNoteScheduler()
        {
            _thread = new Thread(ThreadProc)
            {
                Name = "OneNote-STA",
                IsBackground = true,
            };
            _thread.SetApartmentState(ApartmentState.STA);
            _thread.Start();
            Log("STA thread started");
        }

        public Task<T> Run<T>(Func<OneNoteService, T> func)
        {
            var tcs = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
            _queue.Add(() =>
            {
                try
                {
                    EnsureService();
                    tcs.SetResult(func(_service!));
                }
                catch (OperationCanceledException ex)
                {
                    tcs.SetException(ex);
                }
                catch (Exception ex)
                {
                    Log($"Run<T> error: {ex.Message}");
                    InvalidateService();
                    tcs.SetException(ex);
                }
            });
            return tcs.Task;
        }

        public Task Run(Action<OneNoteService> action)
            => Run<bool>(svc => { action(svc); return true; });

        /// <summary>
        /// 立即在 STA 线程上释放 COM 对象（不关闭队列，不影响窗口生命周期）。
        /// 供 OneFinder-ReleaseCOM 信号触发，确保 ONENOTE.EXE 可以干净退出。
        /// </summary>
        public Task ReleaseCom()
        {
            var tcs = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            if (_disposed)
            {
                tcs.SetResult(false);
                return tcs.Task;
            }
            _queue.Add(() =>
            {
                Log("ReleaseCom: releasing COM service on STA thread");
                InvalidateService();
                Log("ReleaseCom: done");
                tcs.SetResult(true);
            });
            return tcs.Task;
        }

        private void EnsureService()
        {
            if (_service != null && !IsOneNoteRunning())
            {
                Log("EnsureService: OneNote not running, invalidating service");
                InvalidateService();
            }

            if (_service == null)
            {
                Log("EnsureService: creating new OneNoteService");
                _service = new OneNoteService();
            }
        }

        private static bool IsOneNoteRunning()
            => Process.GetProcessesByName("ONENOTE").Any()
            || Process.GetProcessesByName("OneNote").Any();

        private void InvalidateService()
        {
            if (_service != null)
            {
                Log("InvalidateService: disposing service");
                try { _service.Dispose(); } catch (Exception ex) { Log($"InvalidateService error: {ex.Message}"); }
                _service = null;
            }
        }

        private void ThreadProc()
        {
            Log("ThreadProc: entering consumer loop");
            foreach (var action in _queue.GetConsumingEnumerable())
                action();

            // Must release COM object on the same STA thread that created it,
            // so that the COM reference count drops to zero and ONENOTE.EXE can exit.
            Log("ThreadProc: queue completed, releasing service");
            InvalidateService();
            Log("ThreadProc: exiting");
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            Log("Dispose: completing queue");
            _queue.CompleteAdding();       // signal STA thread to finish and dispose _service
            _thread.Join(millisecondsTimeout: 5000); // wait long enough for COM release
            Log("Dispose: STA thread joined");
            _queue.Dispose();
        }
    }
}
