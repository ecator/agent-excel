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

        [DllImport("ole32.dll")]
        private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable pprot);

        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(int reserved, out IBindCtx ppbc);

        /// <summary>
        /// 尝试获取现有的 Excel 实例
        /// </summary>
        /// <returns>Excel Application 实例，未找到则返回 null</returns>
        public static Excel.Application? GetRunningExcelApplication()
        {
            // 方法 1: 使用自定义 GetActiveObject (替代 .NET Framework 的 Marshal.GetActiveObject)
            try
            {
                Guid clsid;
                CLSIDFromProgID("Excel.Application", out clsid);
                GetActiveObject(ref clsid, IntPtr.Zero, out object obj);
                if (obj != null) return obj as Excel.Application;
            }
            catch
            {
                // 忽略错误，尝试方法 2
            }

            // 方法 2: ROT (Running Object Table) 枚举回退实现
            IRunningObjectTable? rot = null;
            IEnumMoniker? enumMoniker = null;
            try
            {
                if (GetRunningObjectTable(0, out rot) != 0 || rot == null)
                    return null;

                rot.EnumRunning(out enumMoniker);
                enumMoniker.Reset();

                IMoniker[] monikers = new IMoniker[1];
                IntPtr fetched = IntPtr.Zero;

                while (enumMoniker.Next(1, monikers, fetched) == 0)
                {
                    IBindCtx? bindCtx = null;
                    try
                    {
                        CreateBindCtx(0, out bindCtx);
                        if (bindCtx == null) continue;

                        monikers[0].GetDisplayName(bindCtx, null, out string displayName);

                        // Excel 的 ROT 名字通常包含文件名或 "!{000208D5-0000-0000-C000-000000000046}"
                        if (displayName.Contains("Excel", StringComparison.OrdinalIgnoreCase) ||
                            displayName.Contains("{000208D5-0000-0000-C000-000000000046}"))
                        {
                            rot.GetObject(monikers[0], out object comObj);
                            if (comObj is Excel.Application app) return app;

                            // 尝试从 Workbook 向上获取 Application
                            if (comObj is Excel._Workbook wb) return wb.Application;
                        }
                    }
                    catch
                    {
                        // 忽略单个条目的错误
                    }
                    finally
                    {
                        if (bindCtx != null) Marshal.ReleaseComObject(bindCtx);
                        if (monikers[0] != null) Marshal.ReleaseComObject(monikers[0]);
                    }
                }
            }
            catch
            {
                // 最终失败
            }
            finally
            {
                if (enumMoniker != null) Marshal.ReleaseComObject(enumMoniker);
                if (rot != null) Marshal.ReleaseComObject(rot);
            }

            return null;
        }
    }
}
