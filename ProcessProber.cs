using System.Diagnostics;
using System.Text.RegularExpressions;

namespace AgentExcel
{
    /// <summary>
    /// 进程与端口探测器
    /// </summary>
    internal class ProcessProber
    {
        /// <summary>
        /// 查找正在运行的服务器端口
        /// </summary>
        /// <returns>端口号，未找到则返回 null</returns>
        public static int? FindRunningServerPort()
        {
            string currentName = Process.GetCurrentProcess().ProcessName;
            int currentPid = Process.GetCurrentProcess().Id;

            // 查找所有同名的后台进程，排除当前正在敲击命令的自身
            Process backgroundProcess = Process.GetProcessesByName(currentName)
                                               .FirstOrDefault(p => p.Id != currentPid);

            if (backgroundProcess == null) return null;

            return GetListeningPortByPid(backgroundProcess.Id);
        }

        /// <summary>
        /// 根据进程 PID 获取其监听的端口
        /// </summary>
        /// <param name="pid">进程 PID</param>
        /// <returns>监听端口，未找到则返回 null</returns>
        private static int? GetListeningPortByPid(int pid)
        {
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = "netstat",
                    Arguments = "-ano",
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using Process p = Process.Start(psi);
                string output = p.StandardOutput.ReadToEnd();
                p.WaitForExit();

                // 提取属于该 PID 并在监听状态的 TCP 本地端口
                string pattern = $@"\s+TCP\s+(?:127\.0\.0\.1|0\.0\.0\.0):(\d+)\s+.*?\s+LISTENING\s+{pid}\b";
                Regex regex = new Regex(pattern, RegexOptions.IgnoreCase);

                Match match = regex.Match(output);
                if (match.Success && int.TryParse(match.Groups[1].Value, out int port))
                {
                    return port;
                }
            }
            catch
            {
                // 忽略系统探测时的权限异常
            }
            return null;
        }
    }
}
