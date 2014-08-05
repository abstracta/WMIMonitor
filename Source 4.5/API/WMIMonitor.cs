namespace Abstracta.WMIMonitor.API
{
    using Logic;

    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Management;
    
    public class WMIMonitor
    {
        private const string Query = "select * from ";

        private readonly string _wmiNamespace;

        private readonly string _wmiClassName;

        private readonly List<string> _wmiProperties;

        private readonly string _wmiKeyProperty;

        private Provider _selectedProvider;

        public WMIMonitor(string wmiNamespace, string wmiClassName, string wmiKeyProperty, string wmiProperties)
        {
            _wmiNamespace = wmiNamespace;
            _wmiClassName = wmiClassName;
            _wmiKeyProperty = wmiKeyProperty;
            _wmiProperties = wmiProperties.Split(',').ToList();
        }

        public Dictionary<string, Dictionary<string, string>> GetWMIValuesFromLocalMachine()
        {
            _selectedProvider = new Provider
            {
                ComputerName = Environment.MachineName,
                Credential = new CurrentUser(),
            };

            return GetWMIValues();
        }

        public Dictionary<string, Dictionary<string, string>> GetWMIValuesFromServer(string server, string userName, string password)
        {
            _selectedProvider = new Provider
            {
                ComputerName = server,
                Credential = new UserPasswAuthentication
                {
                    User = userName,
                    Password = password,
                },
            };

            return GetWMIValues();
        }

        public void ExecuteWMIMethodInAllInstancesInLocalMachine(string methodName)
        {
            _selectedProvider = new Provider
            {
                ComputerName = Environment.MachineName,
                Credential = new CurrentUser(),
            };

            ExecuteWMIMethodOfAllInstances(methodName);
        }

        public void ExecuteWMIMethodInAllInstancesInServer(string server, string userName, string password, string methodName)
        {
            _selectedProvider = new Provider
            {
                ComputerName = server,
                Credential = new UserPasswAuthentication
                {
                    User = userName,
                    Password = password,
                },
            };

            ExecuteWMIMethodOfAllInstances(methodName);
        }

        private void ExecuteWMIMethodOfAllInstances(string methodName)
        {
            var instances = GetWMIInstances();

            foreach (var instance in instances)
            {
                ExecuteWMIMethod(instance, methodName);
            }
        }

        private void ExecuteWMIMethod(string wmiInstanceName, string wmiMethodName)
        {
            var wmiInstance = GetManagementObject(_selectedProvider, wmiInstanceName);
            wmiInstance.InvokeMethod(wmiMethodName, null, null); 
        }

        private Dictionary<string, Dictionary<string, string>> GetWMIValues()
        {
            var instances = GetWMIInstances();

            return instances.ToDictionary(instance => instance, instance => ConvertXMLtoDictionary(GetWMIPropertiesOfInstance(instance)));
        }

        private IEnumerable<string> GetWMIInstances()
        {
            var result = new List<string>();

            var scope = CreateNewManagementScope(_selectedProvider.ComputerName, _selectedProvider.Credential, _wmiNamespace);
            var query = new SelectQuery(Query + _wmiClassName);

            var propId = _wmiKeyProperty;

            try
            {
                using (var searcher = new ManagementObjectSearcher(scope, query))
                {
                    var wmiInstances = searcher.Get();

                    if (wmiInstances.Count > 0)
                    {
                        result.AddRange(from ManagementBaseObject service in wmiInstances select (service[propId] != null ? service[propId].ToString() : string.Empty));
                    }
                }
            }
            catch (Exception exception)
            {
                var errorResult = "Error: " + exception.Message;
                result.Add(errorResult);
            }

            return result;
        }

        private string GetWMIPropertiesOfInstance(string classInstanceName)
        {
            var wmiInstance = GetManagementBaseObject(classInstanceName);

            var props = _wmiProperties;

            if (props.Any(p => p == ConfigManager.AllWMIProperties))
            {
                props = GetAllWMIPropertyNamesOfInstance();
            }

            return "<WMIObject " + String.Join(" ", props.Select(prop => prop + "=" + "\"" + wmiInstance[prop] + "\"")) + "/>";
        }

        private List<string> GetAllWMIPropertyNamesOfInstance()
        {
            var result = new List<string>();

            var scope = CreateNewManagementScope(_selectedProvider.ComputerName, _selectedProvider.Credential, _wmiNamespace);
            var queryStr = "SELECT * FROM meta_class WHERE __CLASS = \"" + _wmiClassName + "\"";
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

        private ManagementBaseObject GetManagementBaseObject(string classInstanceName)
        {
            // Example of Query: new ManagementObject("root\\CIMV2", "Win32_Service.Name='PlugPlay'", null);
            // return new ManagementObject(serverString, className + "." + _propId + "='" + classInstance + "'", null);

            var scope = CreateNewManagementScope(_selectedProvider.ComputerName, _selectedProvider.Credential, _wmiNamespace);
            var query = new SelectQuery(Query + _wmiClassName);

            var propId = _wmiKeyProperty;

            using (var searcher = new ManagementObjectSearcher(scope, query))
            {
                var services = searcher.Get();
                foreach (var service in services.Cast<ManagementBaseObject>().Where(service => classInstanceName == (service[propId] != null ? service[propId].ToString() : string.Empty)))
                {
                    return service;
                }
            }

            return null;
        }

        private ManagementObject GetManagementObject(Provider provider, string classInstance)
        {
            var serverString = GetServerString(provider.ComputerName, _wmiNamespace);

            // Example of Query: new ManagementObject("root\\CIMV2", "Win32_Service.Name='PlugPlay'", null);
            return new ManagementObject(serverString, _wmiClassName + "." + _wmiKeyProperty + "='" + classInstance + "'", null);
        }
        
        private static Dictionary<string, string> ConvertXMLtoDictionary(string propsXML)
        {
            var res = new Dictionary<string, string>();

            var parts = propsXML.Split(' ').ToList();
            parts.RemoveAt(0);

            foreach (var att in parts)
            {
                try
                {
                    var name = att.Split('=')[0];
                    var value = att.Split('=')[1];

                    name = name.Replace("\"", "");
                    value = value.Replace("\"", "");

                    res.Add(name, value);
                }
                catch (Exception ex)
                {
                }
            }

            return res;
        }

        private static ManagementScope CreateNewManagementScope(string serverName, Credentials credential, string wmiNamespace)
        {
            var serverString = GetServerString(serverName, wmiNamespace);

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
