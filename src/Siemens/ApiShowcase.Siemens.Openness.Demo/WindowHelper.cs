using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace ApiShowcase.Siemens.Openness.Demo
{
  public static class WindowHelper
  {
    const int SW_RESTORE = 9;

    public static void BringProcessToFront(Process process)
    {
      IntPtr handle = process.MainWindowHandle;
      if (IsIconic(handle))
      {
        ShowWindow(handle, SW_RESTORE);
      }

      SetForegroundWindow(handle);
    }

    [DllImport("User32.dll")]
    private static extern bool SetForegroundWindow(IntPtr handle);

    [DllImport("User32.dll")]
    private static extern bool ShowWindow(IntPtr handle, int nCmdShow);

    [DllImport("User32.dll")]
    private static extern bool IsIconic(IntPtr handle);
  }
}