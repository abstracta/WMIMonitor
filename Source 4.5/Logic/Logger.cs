namespace Abstracta.WMIMonitor.Logic
{
    using System.Collections.Generic;
    using System.IO;

    internal static class Logger
    {
        internal static void Log(List<string> result)
        {
            var log = ConfigManager.GetInstance().LogResults;
            if (!log)
            {
                return;
            }

            var logFileName = ConfigManager.GetInstance().LogFileName;
            var strWriter = new StreamWriter(logFileName);

            foreach (var line in result)
            {
                strWriter.WriteLine(line);
            }

            strWriter.Close();
        }
    }
}
