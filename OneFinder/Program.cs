using System;
using System.Threading;
using System.Windows.Forms;

namespace OneFinder
{
    internal static class Program
    {
        private const string MutexName = "Local\\OneFinder-SingleInstance";

        [STAThread]
        static void Main()
        {
            using var mutex = new Mutex(initiallyOwned: true, MutexName, out bool createdNew);
            if (!createdNew)
            {
                // 已有实例在运行，直接退出
                return;
            }

            ApplicationConfiguration.Initialize();
            Application.Run(new MainForm());
        }
    }
}
