using System;
using System.Windows.Forms;

namespace StyleWatcherWin
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            ApplicationConfiguration.Initialize();
            Application.Run(new TrayApp());
        }
    }
}
