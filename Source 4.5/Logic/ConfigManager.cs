namespace Abstracta.WMIMonitor.Logic
{
    using System;
    using System.Configuration;
    using System.Linq;
    using System.Xml;
    using System.Collections.Generic;

    public class ConfigManager
    {
        internal const string AllWMIProperties = "*";

        internal const string DefaultProvidersFileName = "Providers.xml";

        private static volatile ConfigManager _instance;

        private static readonly object Lock = new object();

        internal List<Provider> Providers { get; set; }

        internal Provider SelectedProvider { get; set; }

        internal Provider GetProviderByName(string server)
        {
            return Providers.FirstOrDefault(provider => provider.ComputerName == server);
        }

        internal List<string> ProviderNames
        {
            get { return Providers.Select(provider => provider.ComputerName).ToList(); }
        }

        private ConfigManager()
        {
        }

        internal static ConfigManager GetInstance()
        {
            if (_instance == null)
            {
                lock (Lock)
                {
                    if (_instance == null)
                    {
                        var providersFileName = ConfigurationManager.AppSettings.Get("providersFileName") ??
                                                DefaultProvidersFileName;

                        var props = ConfigurationManager.AppSettings.Get("WMIFilterProperties") ?? AllWMIProperties;
                        var tmp = props.Split(',');
                        var propsList = tmp.Select(res => res.Trim()).ToList();

                        _instance = new ConfigManager
                            {
                                Providers = GetProvidersFromConfigFile(providersFileName),
                                WMIOutputFormatType = OutputFormatType.GUIDetailFormat,
                                WMINamespace = ConfigurationManager.AppSettings.Get("WMINamespace") ?? "cimv2",
                                WMIClassName = ConfigurationManager.AppSettings.Get("WMIClassName") ?? "Win32_Service",
                                WMIKeyProperty = ConfigurationManager.AppSettings.Get("WMIKeyProperty") ?? "Caption",
                                WMIFilterProperties = propsList,
                            };
                    }
                }
            }

            return _instance;
        }

        internal static ConfigManager GetEmptyInstance()
        {
            if (_instance == null)
            {
                lock (Lock)
                {
                    if (_instance == null)
                    {
                        _instance = new ConfigManager();
                    }
                }
            }

            return _instance;
        }

        internal void ReloadConfiguration()
        {
            Providers = GetProvidersFromConfigFile(DefaultProvidersFileName);
        }

        internal string WMINamespace { get; set; }

        internal string WMIClassName { get; set; }

        internal string WMIInstanceName { get; set; }

        internal string WMIKeyProperty { get; set; }

        internal List<string> WMIFilterProperties { get; set; }
       
        internal OutputFormatType WMIOutputFormatType { get; set; }

        internal bool ExecuteWMIMethod { get; set; }

        /// <summary>
        /// Enumerate method names using CSV
        /// </summary>
        internal string WMIMethodsToExecute { get; set; }

        private static List<Provider> GetProvidersFromConfigFile(string configFileName)
        {
            var result = new List<Provider>();

            var doc = new XmlDocument();
            doc.Load(configFileName);

            if (doc.DocumentElement != null)
            {
                foreach (XmlNode node in doc.DocumentElement.ChildNodes)
                {
                    // it can be also a XmlComment element type
                    var element = node as XmlElement;
                    if (element == null) continue;

                    var provider = element;
                    var providerType = provider.GetAttribute("type");

                    string computerName;
                    switch (providerType)
                    {
                        case "local":
                            computerName = Environment.MachineName;
                            break;

                        case "remote":
                            computerName = provider.GetAttribute("name");
                            break;

                        default:
                            throw new ConfigException("Provider type unknown: " + providerType);
                    }

                    if (provider.GetElementsByTagName("credential").Count != 1)
                    {
                        throw new ConfigException(
                            "XML element 'provider' requires one children with tag 'credential'");
                    }

                    var credentials = (XmlElement) provider.GetElementsByTagName("credential")[0];
                    var credentialType = credentials.GetAttribute("type");

                    Credentials credential;
                    switch (credentialType)
                    {
                        case "currentUser":
                            credential = new CurrentUser();
                            break;

                        case "authenticationByPassw":
                            var userName = credentials.GetAttribute("userName");
                            var password = credentials.GetAttribute("password");

                            if (string.IsNullOrEmpty(userName))
                            {
                                throw new ConfigException(
                                    "Credential 'authenticationByPassw' needs a 'userName' value");
                            }
                            if (string.IsNullOrEmpty(password))
                            {
                                throw new ConfigException(
                                    "Credential 'authenticationByPassw' needs a 'password' value");
                            }

                            credential = new UserPasswAuthentication
                                {
                                    User = userName,
                                    Password = password,
                                };
                            break;

                        default:
                            throw new ConfigException("Credential type unknown: " + credentialType);
                    }

                    result.Add(new Provider
                        {
                            ComputerName = computerName,
                            Credential = credential,
                        });
                }
            }

            return result;
        }
    }
}