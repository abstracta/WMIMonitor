namespace Abstracta.WMIMonitor.Logic
{
    using System;
    
    internal class ConfigException : Exception
    {
        public ConfigException()
        {
        }

        public ConfigException(string message) : base(message)
        {
        }
    }
}
