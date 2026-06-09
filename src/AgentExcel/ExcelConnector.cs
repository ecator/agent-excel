using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Excel = Microsoft.Office.Interop.Excel;

namespace AgentExcel
{
    /// <summary>
    /// Excel 实例连接器，负责通过 COM 接口探测和接管运行中的 Excel
    /// </summary>
    internal static class ExcelConnector
    {
        [DllImport("oleaut32.dll", PreserveSig = false)]
        private static extern void GetActiveObject([In] ref Guid rclsid, IntPtr pvReserved, [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

        [DllImport("ole32.dll")]
        private static extern int CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] string lpszProgID, out Guid pclsid);

        /// <summary>
        /// 确保获得一个运行中的 Excel 实例 (优先接管现有进程，没有则创建并设为 Visible)
        /// </summary>
        /// <returns>Excel Application 实例</returns>
        public static Excel.Application EnsureExcelApplication()
        {
            Excel.Application? app = GetRunningExcelApplication();
            if (app == null)
            {
                // 注意：这里新建后必须 Visible=true，以符合“共享/用户感知模式”
                app = new Excel.Application { Visible = true };
            }
            return app;
        }

        /// <summary>
        /// 尝试获取现有的 Excel 实例 (不抛出异常，找不到则返回 null)
        /// </summary>
        public static Excel.Application? GetRunningExcelApplication()
        {
            try
            {
                Guid clsid;
                CLSIDFromProgID("Excel.Application", out clsid);
                GetActiveObject(ref clsid, IntPtr.Zero, out object obj);
                if (obj != null) return obj as Excel.Application;
            }
            catch
            {
                // 忽略错误，找不到运行中的对象是常见情况
            }
            return null;
        }

        /// <summary>
        /// 安全释放 COM 对象，并置为 null
        /// </summary>
        public static void SafeReleaseComObject(ref object? obj)
        {
            if (obj != null && Marshal.IsComObject(obj))
            {
                try
                {
                    Marshal.ReleaseComObject(obj);
                }
                catch
                {
                    // 记录日志或忽略
                }
                finally
                {
                    obj = null;
                }
            }
        }

        /// <summary>
        /// 显式调用一次 GC.Collect 以清理底层残留的 COM 组件
        /// </summary>
        public static void ForceGarbageCollection()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
