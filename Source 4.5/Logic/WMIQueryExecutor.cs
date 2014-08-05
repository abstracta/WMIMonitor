namespace Abstracta.WMIMonitor.Logic
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Management;

    public class WMIWrapper
    {
        private static volatile WMIWrapper _instance;

        // Property that identifies an object of the instance
        private readonly string _propId = ConfigManager.GetInstance().WMIKeyProperty;

        private const string Query = "select * from ";

        private static readonly object Lock = new object();

        public const string ErrorPrefix = "ERROR: ";

        public const string DetailSeparator = ": ";

        public static WMIWrapper GetInstance()
        {
            if (_instance == null)
            {
                lock (Lock)
                {
                    if (_instance == null)
                    {
                        _instance = new WMIWrapper();
                    }
                }
            }

            return _instance;
        }

        public List<string> GetWMIClassesInNamespace(string serverName, string wmiNamespace)
        {
            var searcher = new ManagementObjectSearcher(
                    new ManagementScope("root/" + wmiNamespace),
                    new WqlObjectQuery("select * from meta_class"),
                    null);

            var result = (from ManagementClass wmiClass in searcher.Get() select wmiClass["__CLASS"].ToString()).ToList();

            result.Sort();

            return result;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="serverName">The name of the server that exposes the WMI objects</param>
        /// <returns></returns>
        public List<string> GetWMIInstancesOfClass(string serverName)
        {
            return GetWMIInstancesOfClass(serverName, ConfigManager.GetInstance().WMIClassName);
        }

        public List<string> GetWMIInstancesOfClass(string serverName, string className)
        {
            var result = new List<string>();
            var provider = ConfigManager.GetInstance().GetProviderByName(serverName);

            if (provider == null)
            {
                throw new ConfigException("Provider (server name) not found in config file: " + serverName);
            }

            var scope = CreateNewManagementScope(provider.ComputerName, provider.Credential);
            var query = new SelectQuery(Query + className);

            try
            {
                using (var searcher = new ManagementObjectSearcher(scope, query))
                {
                    var services = searcher.Get();
                    result.AddRange(from ManagementBaseObject service in services select service[_propId].ToString());
                }
            }
            catch (Exception exception)
            {
                var errorResult = ErrorPrefix + exception.Message;
                result.Add(errorResult);
            }

            return result;
        }

        public List<string> GetWMIPropertiesOfInstance(string serverName, string className, string classInstance)
        {
            return GetWMIPropertiesOfInstance(serverName, className, classInstance, OutputFormatType.GUIDetailFormat);
        }

        public List<string> GetWMIPropertiesOfInstance(string serverName, string className, string classInstance, OutputFormatType format)
        {
            var result = new List<string>();
            var provider = ConfigManager.GetInstance().GetProviderByName(serverName);

            if (provider == null)
            {
                throw new ConfigException("Provider (server name) not found in config file: " + serverName);
            }

            var wmiInstance = GetManagementBaseObject(provider, className, classInstance);

            var props = ConfigManager.GetInstance().WMIProperties;
            switch (format)
            {
                case OutputFormatType.GUIDetailFormat:
                    result.AddRange(props.Select(prop => prop + DetailSeparator + wmiInstance[prop]));
                    break;

                case OutputFormatType.CSVFormat:
                    var propsString = String.Join(";", props.Select(prop => wmiInstance[prop]));
                    result.Add(propsString);
                    break;

                case OutputFormatType.XMLFormat:
                    var tmp = "<WMIObject " + String.Join(" ", props.Select(prop => prop + "=" + "\"" + wmiInstance[prop] + "\"")) + "/>";
                    result.Add(tmp);
                    break;
            }

            return result;
        }

        public List<MethodWapper> GetWMIMethodsOfInstance(string serverName, string className, string classInstance)
        {
            var result = new List<MethodWapper>();
            var provider = ConfigManager.GetInstance().GetProviderByName(serverName);

            if (provider == null)
            {
                throw new ConfigException("Provider (server name) not found in config file: " + serverName);
            }

            var scope = CreateNewManagementScope(provider.ComputerName, provider.Credential);
            var queryStr = "SELECT * FROM meta_class WHERE __CLASS = \"" + className + "\"";
            // var query = new SelectQuery(_query + className);
            var query = new SelectQuery(queryStr);

            var wmiInstance = GetManagementObject(provider, className, classInstance);

            using (var searcher = new ManagementObjectSearcher(scope, query))
            {
                // get the each class from the select query.
                foreach (ManagementClass c in searcher.Get())
                {
                    //// get the each propery name of WMI class
                    //foreach (PropertyData d in c.Properties)
                    //{
                    //    listBox1.Items.Add(d.Name);
                    //}

                    // get the each method name of WMI class
                    result.AddRange(from MethodData m in c.Methods select new MethodWapper(m, wmiInstance));
                }
            }

            return result;
        }

        internal string GetWMIInstanceAsXML(string serverName, string className, string classInstance)
        {
            var provider = ConfigManager.GetInstance().GetProviderByName(serverName);

            if (provider == null)
            {
                throw new ConfigException("Provider (server name) not found in config file: " + serverName);
            }

            var wmiInstance = GetManagementBaseObject(provider, className, classInstance);

            return (wmiInstance != null)
                ? wmiInstance.GetText(TextFormat.WmiDtd20)
                : null;
        }

        private ManagementBaseObject GetManagementBaseObject(Provider provider, string className, string classInstance)
        {
            // Example of Query: new ManagementObject("root\\CIMV2", "Win32_Service.Name='PlugPlay'", null);
            // return new ManagementObject(serverString, className + "." + _propId + "='" + classInstance + "'", null);

            var scope = CreateNewManagementScope(provider.ComputerName, provider.Credential);
            var query = new SelectQuery(Query + className);

            using (var searcher = new ManagementObjectSearcher(scope, query))
            {
                var services = searcher.Get();
                foreach (var service in services.Cast<ManagementBaseObject>().Where(service => classInstance == (string)service[_propId]))
                {
                    return service;
                }
            }

            return null;
        }

        private ManagementObject GetManagementObject(Provider provider, string className, string classInstance)
        {
            var serverString = GetServerString(provider.ComputerName, ConfigManager.GetInstance().WMINamespace);

            // Example of Query: new ManagementObject("root\\CIMV2", "Win32_Service.Name='PlugPlay'", null);
            return new ManagementObject(serverString, className + "." + _propId + "='" + classInstance + "'", null);
        }

        private static ManagementScope CreateNewManagementScope(string serverName, Credentials credential)
        {
            var serverString = GetServerString(serverName, ConfigManager.GetInstance().WMINamespace);

            var scope = new ManagementScope(serverString);

            var upa = credential as UserPasswAuthentication;
            if (upa != null)
            {
                scope.Options = new ConnectionOptions
                {
                    Username = upa.User,
                    Password = upa.Password,
                    Impersonation = ImpersonationLevel.Impersonate,
                    Authentication = AuthenticationLevel.PacketPrivacy
                };
            }

            return scope;
        }

        private static string GetServerString(string computerName, string wmiNamespace)
        {
            return @"\\" + computerName + @"\root\" + wmiNamespace;
        }
    }
}
