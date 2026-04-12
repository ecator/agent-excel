using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace AgentExcel
{
    /// <summary>
    /// Excel 处理工具类，实现 IDisposable 以确保 COM 对象被正确释放
    /// </summary>
    public class ExcelHandler : IDisposable
    {
        private Excel.Application? _application;
        private Excel.Workbooks? _workbooks;
        private Excel.Workbook? _workbook;
        private Excel.Worksheet? _worksheet;
        private bool _disposed;
        private readonly bool _wasAttached;

        /// <summary>
        /// 获取或创建一个 Excel 实例
        /// </summary>
        /// <param name="visible">是否显示 Excel 窗口</param>
        public ExcelHandler(bool visible = false)
        {
            // 尝试连接现有实例
            _application = ExcelConnector.GetRunningExcelApplication();
            if (_application != null)
            {
                _wasAttached = true;
            }
            else
            {
                _application = new Excel.Application();
                _wasAttached = false;
            }

            _application.Visible = visible;
            _workbooks = _application.Workbooks;
        }

        /// <summary>
        /// 打开现有的 Excel 文件
        /// </summary>
        public void OpenWorkbook(string filePath)
        {
            ReleaseObject(_workbook);
            // 确保使用绝对路径，避免 COM 接口在不同工作目录下找不到文件
            string fullPath = System.IO.Path.GetFullPath(filePath);
            _workbook = _workbooks?.Open(fullPath);
        }

        /// <summary>
        /// 创建一个新的工作簿
        /// </summary>
        public void CreateWorkbook()
        {
            ReleaseObject(_workbook);
            _workbook = _workbooks?.Add();
        }

        /// <summary>
        /// 选择指定名称的工作表
        /// </summary>
        /// <param name="sheetName">工作表名称</param>
        public void SelectSheet(string sheetName)
        {
            ReleaseObject(_worksheet);
            _worksheet = _workbook?.Sheets[sheetName] as Excel.Worksheet;
        }

        /// <summary>
        /// 选择指定索引的工作表（从 1 开始）
        /// </summary>
        /// <param name="index">索引</param>
        public void SelectSheet(int index)
        {
            ReleaseObject(_worksheet);
            _worksheet = _workbook?.Sheets[index] as Excel.Worksheet;
        }

        /// <summary>
        /// 读取单元格的值
        /// </summary>
        /// <param name="row">行号（从 1 开始）</param>
        /// <param name="column">列号（从 1 开始）</param>
        /// <returns>单元格内容</returns>
        public object? ReadCell(int row, int column)
        {
            Excel.Range? range = null;
            try
            {
                range = _worksheet?.Cells[row, column] as Excel.Range;
                return range?.Value2;
            }
            finally
            {
                ReleaseObject(range);
            }
        }

        /// <summary>
        /// 向单元格写入值
        /// </summary>
        /// <param name="row">行号（从 1 开始）</param>
        /// <param name="column">列号（从 1 开始）</param>
        /// <param name="value">要写入的值</param>
        public void WriteCell(int row, int column, object value)
        {
            Excel.Range? range = null;
            try
            {
                range = _worksheet?.Cells[row, column] as Excel.Range;
                if (range != null)
                {
                    range.Value2 = value;
                }
            }
            finally
            {
                ReleaseObject(range);
            }
        }

        /// <summary>
        /// 保存当前工作簿
        /// </summary>
        public void Save()
        {
            _workbook?.Save();
        }

        /// <summary>
        /// 将工作簿另存为
        /// </summary>
        /// <param name="filePath">文件路径</param>
        public void SaveAs(string filePath)
        {
            _workbook?.SaveAs(filePath);
        }

        /// <summary>
        /// 释放所有 COM 对象
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                // 按照创建相反的顺序释放 COM 对象
                ReleaseObject(_worksheet);
                _worksheet = null;

                ReleaseObject(_workbook);
                _workbook = null;

                ReleaseObject(_workbooks);
                _workbooks = null;

                if (_application != null)
                {
                    try
                    {
                        // 如果是新创建的实例，退出进程；如果是连接的实例，保持运行
                        if (!_wasAttached)
                        {
                            _application.Quit();
                        }
                    }
                    catch { }
                    ReleaseObject(_application);
                    _application = null;
                }

                _disposed = true;
            }
        }

        private void ReleaseObject(object? obj)
        {
            if (obj != null && Marshal.IsComObject(obj))
            {
                try
                {
                    Marshal.ReleaseComObject(obj);
                }
                catch
                {
                    // 忽略释放过程中的错误
                }
                finally
                {
                    obj = null;
                }
            }
        }

        ~ExcelHandler()
        {
            Dispose(false);
        }
    }
}
