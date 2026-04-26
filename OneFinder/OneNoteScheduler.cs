using System;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace OneFinder
{
    /// <summary>
    /// 在专用 STA 线程上序列化所有 OneNote COM 调用。
    /// OneNote 是 STA COM 服务器，若从 MTA 线程池线程调用会触发跨
    /// apartment 代理，在快速连续操作时导致 OneNote 崩溃。
    /// 此类确保 COM 对象的创建与所有调用始终在同一 STA 线程上进行。
    /// </summary>
    internal sealed class OneNoteScheduler : IDisposable
    {
        private readonly BlockingCollection<Action> _queue = new();
        private readonly Thread _thread;
        private OneNoteService? _service;
        private bool _disposed;

        public OneNoteScheduler()
        {
            _thread = new Thread(ThreadProc)
            {
                Name = "OneNote-STA",
                IsBackground = true,
            };
            _thread.SetApartmentState(ApartmentState.STA);
            _thread.Start();
        }

        /// <summary>在 STA 线程上执行有返回值的 COM 操作。</summary>
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
                    // 取消不代表 COM 连接损坏，不需要丢弃 service
                    tcs.SetException(ex);
                }
                catch (Exception ex)
                {
                    InvalidateService();
                    tcs.SetException(ex);
                }
            });
            return tcs.Task;
        }

        /// <summary>在 STA 线程上执行无返回值的 COM 操作。</summary>
        public Task Run(Action<OneNoteService> action)
            => Run<bool>(svc => { action(svc); return true; });

        private void EnsureService()
        {
            // 如果 OneNote 进程已不存在，先主动丢弃旧连接，
            // 避免对死连接发出 COM 调用（会触发 COM 自动激活，
            // 在 OneNote 未就绪时调用导致崩溃）。
            if (_service != null && !IsOneNoteRunning())
                InvalidateService();

            _service ??= new OneNoteService();
        }

        private static bool IsOneNoteRunning()
            => Process.GetProcessesByName("ONENOTE").Any()
            || Process.GetProcessesByName("OneNote").Any();

        private void InvalidateService()
        {
            try { _service?.Dispose(); } catch { }
            _service = null;
        }

        private void ThreadProc()
        {
            foreach (var action in _queue.GetConsumingEnumerable())
                action();
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            _queue.CompleteAdding();
            _thread.Join(millisecondsTimeout: 2000);
            _service?.Dispose();
            _queue.Dispose();
        }
    }
}
