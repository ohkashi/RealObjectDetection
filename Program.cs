using Microsoft.Win32;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace RealObjectDetection
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            ApplicationConfiguration.Initialize();
			//Application.Run(new Form1());
			using Form1 frm = new();
			frm.Show();
			var sw = Stopwatch.StartNew();
			while (frm.Created) {
				Application.DoEvents();
				frm.Render(sw.ElapsedMilliseconds);
			}
			sw.Stop();
		}
	}

	public static class DarkMode
	{
		public static bool IsLightTheme()
		{
			using var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize");
			var value = key?.GetValue("AppsUseLightTheme");
			return value is int i && i > 0;
		}

		[DllImport("dwmapi.dll")]
		private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);

		private const int DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1 = 19;
		private const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;

		public static bool UseImmersiveDarkMode(IntPtr handle, bool enabled)
		{
			if (OperatingSystem.IsWindowsVersionAtLeast(10, 0, 17763)) {
				var attribute = DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1;
				if (OperatingSystem.IsWindowsVersionAtLeast(10, 0, 18985)) {
					attribute = DWMWA_USE_IMMERSIVE_DARK_MODE;
				}

				int useImmersiveDarkMode = enabled ? 1 : 0;
				return DwmSetWindowAttribute(handle, attribute, ref useImmersiveDarkMode, sizeof(int)) == 0;
			}

			return false;
		}
	}
}