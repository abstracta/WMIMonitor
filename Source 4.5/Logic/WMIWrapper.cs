namespace Abstracta.WMIMonitor.Logic
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Management;

    using LogicInterface;

    public class WMIWrapper
    {
        private static volatile WMIWrapper _instance;

        private const string Query = "select * from ";

        private static readonly object Lock = new object();

        internal const string ErrorPrefix = "ERROR: ";

        internal const string DetailSeparator = ": ";

        internal const char CSVSeparator = '\t';

        private WMIWrapper()
        {
        }

        internal static WMIWrapper GetInstance()
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

        internal List<string> GetWMIClassesInNamespace()
        {
            var wmiNamespace = ConfigManager.GetInstance().WMINamespace;
         
            var searcher = new ManagementObjectSearcher(
                    new ManagementScope("root/" + wmiNamespace),
                    new WqlObjectQuery("select * from meta_class"),
                    null);

            var result = (from ManagementClass wmiClass in searcher.Get() select wmiClass["__CLASS"].ToString()).ToList();

            result.Sort();

            return result;
        }

        internal List<string> GetWMIInstances()
        {
            var result = new List<string>();

            var className = ConfigManager.GetInstance().WMIClassName;
            var provider = ConfigManager.GetInstance().SelectedProvider;

            var scope = CreateNewManagementScope(provider.ComputerName, provider.Credential);
            var query = new SelectQuery(Query + className);

            var propId = ConfigManager.GetInstance().WMIKeyProperty;

            try
            {
                using (var searcher = new ManagementObjectSearcher(scope, query))
                {
                    var wmiInstances = searcher.Get();

                    if (wmiInstances.Count > 0)
                    {
                        var enumerator = wmiInstances.GetEnumerator();
                        enumerator.MoveNext();
                        var firstElement = enumerator.Current;

                        try
                        {
                            if (firstElement[propId] != null)
                            {
                            }
                        }
                        catch
                        {
                            propId = GetFirstWMIPropertyOfInstance(className);
                            ConfigManager.GetInstance().WMIKeyProperty = propId;
                        }

                        result.AddRange(from ManagementBaseObject service in wmiInstances select (service[propId] != null ? service[propId].ToString() : string.Empty));
                    }
                }
            }
            catch (Exception exception)
            {
                var errorResult = ErrorPrefix + exception.Message;
                result.Add(errorResult);
            }

            return result;
        }

        internal List<string> GetWMIPropertiesOfInstance()
        {
            return GetWMIPropertiesOfInstance(OutputFormatType.GUIDetailFormat);
        }

        internal List<string> GetWMIPropertiesOfInstance(OutputFormatType format)
        {
            var result = new List<string>();

            var className = ConfigManager.GetInstance().WMIClassName;
            var classInstance = ConfigManager.GetInstance().WMIInstanceName;
            var provider = ConfigManager.GetInstance().SelectedProvider;
            
            var wmiInstance = GetManagementBaseObject(provider, className, classInstance);

            var props = ConfigManager.GetInstance().WMIFilterProperties;

            if (props.Count == 1 && props[0] == ConfigManager.AllWMIProperties)
            {
                props = GetAllWMIPropertyNamesOfInstance();
            }

            switch (format)
            {
                case OutputFormatType.GUIDetailFormat:
                    result.AddRange(props.Select(prop => prop + DetailSeparator + wmiInstance[prop]));
                    break;

                case OutputFormatType.CSVFormat:
                    var propsString = String.Join(CSVSeparator.ToString(CultureInfo.InvariantCulture), props.Select(prop => wmiInstance[prop]));
                    result.Add(propsString);
                    break;

                case OutputFormatType.XMLFormat:
                    var tmp = "<WMIObject " + String.Join(" ", props.Select(prop => prop + "=" + "\"" + wmiInstance[prop] + "\"")) + "/>";
                    result.Add(tmp);
                    break;
            }

            return result;
        }

        internal string GetFirstWMIPropertyOfInstance(string className)
        {
            var provider = ConfigManager.GetInstance().SelectedProvider;
            var props = GetAllWMIPropertyNamesOfInstance(className, provider);

            return props.First();
        }

        internal object[] GetWMIPropertiesOfInstanceAsObjectArray()
        {
            var props = ConfigManager.GetInstance().WMIFilterProperties;
            if (props.Count == 1 && props[0] == ConfigManager.AllWMIProperties)
            {
                props = GetAllWMIPropertyNamesOfInstance();
            }

            var result = new object[props.Count];

            var className = ConfigManager.GetInstance().WMIClassName;
            var classInstance = ConfigManager.GetInstance().WMIInstanceName;
            var provider = ConfigManager.GetInstance().SelectedProvider;

            var wmiInstance = GetManagementBaseObject(provider, className, classInstance);

            var i = 0;
            foreach (var prop in props)
            {
                result[i] = wmiInstance[prop];
                i++;
            }

            return result;
        }

        internal List<string> GetAllWMIPropertyNamesOfInstance()
        {
            var className = ConfigManager.GetInstance().WMIClassName;
            var provider = ConfigManager.GetInstance().SelectedProvider;

            return GetAllWMIPropertyNamesOfInstance(className, provider);
        }

        internal List<MethodWapper> GetWMIMethodsOfInstance()
        {
            var result = new List<MethodWapper>();

            var className = ConfigManager.GetInstance().WMIClassName;
            var classInstance = ConfigManager.GetInstance().WMIInstanceName;
            var provider = ConfigManager.GetInstance().SelectedProvider;

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
                    // get the each method name of WMI class
                    result.AddRange(from MethodData m in c.Methods 
                                    where IsInstanceMethod(m)
                                    select new MethodWapper(m, wmiInstance));
                }
            }

            return result;
        }

        internal string GetWMIInstanceAsXML()
        {
            var className = ConfigManager.GetInstance().WMIClassName;
            var classInstance = ConfigManager.GetInstance().WMIInstanceName;
            var provider = ConfigManager.GetInstance().SelectedProvider;

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

            var propId = ConfigManager.GetInstance().WMIKeyProperty;

            using (var searcher = new ManagementObjectSearcher(scope, query))
            {
                var services = searcher.Get();
                foreach (var service in services.Cast<ManagementBaseObject>().Where(service => classInstance == (service[propId] != null ? service[propId].ToString() : string.Empty)))
                {
                    return service;
                }
            }

            return null;
        }

        private ManagementObject GetManagementObject(Provider provider, string className, string classInstance)
        {
            var serverString = GetServerString(provider.ComputerName, ConfigManager.GetInstance().WMINamespace);

            var propId = ConfigManager.GetInstance().WMIKeyProperty;

            // Example of Query: new ManagementObject("root\\CIMV2", "Win32_Service.Name='PlugPlay'", null);
            return new ManagementObject(serverString, className + "." + propId + "='" + classInstance + "'", null);
        }
        
        private List<string> GetAllWMIPropertyNamesOfInstance(string className, Provider provider)
        {
            var result = new List<string>();

            var scope = CreateNewManagementScope(provider.ComputerName, provider.Credential);
            var queryStr = "SELECT * FROM meta_class WHERE __CLASS = \"" + className + "\"";
            var query = new SelectQuery(queryStr);

            using (var searcher = new ManagementObjectSearcher(scope, query))
            {
                // get the each class from the select query.
                foreach (ManagementClass c in searcher.Get())
                {
                    result.AddRange(from PropertyData m in c.Properties select m.Name);
                }
            }

            return result;
        }

        private static bool IsInstanceMethod(MethodData m)
        {
            foreach (var q in m.Qualifiers)
            {
                if (q.Name == "Static" && (bool)q.Value)
                {
                    return false;
                }
            }

            return true;
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
                    Authentication = AuthenticationLevel.PacketPrivacy,
                    EnablePrivileges = true
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
