using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Addon_AutoDivSalesOrd.Addons.Tools
{
    public class MarshalGC
    {
        [DllImport("kernel32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetProcessWorkingSetSize(IntPtr process, UIntPtr minimumWorkingSetSize, UIntPtr maximumWorkingSetSize);

        public static void ReleaseComObject(object obj)
        {
            if (obj != null && Marshal.IsComObject(obj))
            {
                Marshal.ReleaseComObject(obj);
            }
        }

        public static void ReleaseComObjects(params object[] objects)
        {
            ReleaseValidComObjects(objects);
        }

        public static void ReleaseComObjectsAndMinimize(params object[] objects)
        {
            // Avoid forcing GC/Finalizers/WorkingSet trimming during UI events.
            // This often causes instability with SAP B1 COM objects (RPC failures / invalid RCWs).
            ReleaseValidComObjects(objects);
        }

        public static void MinimizeMemory()
        {
            GC.Collect(GC.MaxGeneration);
            GC.WaitForPendingFinalizers();
            SetProcessWorkingSetSize(Process.GetCurrentProcess().Handle, (UIntPtr)0xFFFFFFFF, (UIntPtr)0xFFFFFFFF);
        }

        private static bool ReleaseValidComObjects(params object[] objects)
        {
            bool release = false;
            int index = 0;
            foreach (var obj in objects)
            {
                if (obj != null && Marshal.IsComObject(obj))
                {
                    Marshal.ReleaseComObject(obj);
                    objects[index] = null;
                    release = true;
                }
            }

            return release;
        }
    }
}
