namespace Abstracta.WMIMonitor.LogicInterface
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Windows;
    using System.Windows.Controls;
    using Logic;

    internal enum Command
    {
        GetWMINamespace,
        GetWMIClasses, 
        GetWMIInstancesOfClass,
        GetWMIPropertiesOfInstance,
        GetWMIMethodsOfInstance,
        GetWMIInstanceAsXML,
        GetWMIKeyProperty,
        GetPropertiesFilter,
        ReloadConfiguration,
        SetSelectedServer,
        SetNewNamespace,
        SetSelectedClass,
        SetSelectedInstance,
        SetNewPropertiesFilter,
        SetNewKeyProperty,
    }

    internal static class CommandManager
    {
        internal const char Separator = ';';

        internal static string ErrorPrefix = WMIWrapper.ErrorPrefix;

        internal static ConfigManager CfgMngr { get { return ConfigManager.GetInstance(); } }

        internal static void Initialize(ItemCollection resultsList, TextBox namespaceText, TextBox propertiesText, TextBox keyProperty)
        {
            resultsList.Add("");
            try
            {
                namespaceText.Text = CfgMngr.WMINamespace;
                
                propertiesText.Text = String.Join(", ", CfgMngr.WMIFilterProperties);

                keyProperty.Text = CfgMngr.WMIKeyProperty;

                var servers = ConfigManager.GetInstance().Providers;
                foreach (var server in servers)
                {
                    resultsList.Add(server.ComputerName);
                }
            }
            catch (ConfigException cex)
            {
                MessageBox.Show(
                    "Configuration error. Fix '" + ConfigManager.DefaultProvidersFileName + "' file and reload it: " +
                    cex.Message, "ERROR", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "ERROR", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        internal static object Execute(Command command, string parameter)
        {
            try
            {
                switch (command)
                {
                    case Command.ReloadConfiguration:
                        ConfigManager.GetInstance().ReloadConfiguration();
                        return null;

                    case Command.SetSelectedServer:
                        ConfigManager.GetInstance().SelectedProvider = ConfigManager.GetInstance().GetProviderByName(parameter);
                        break;

                    case Command.SetNewNamespace:
                        ConfigManager.GetInstance().WMINamespace = parameter;
                        break;

                    case Command.SetSelectedClass:
                        ConfigManager.GetInstance().WMIClassName = parameter;
                        break;

                    case Command.SetSelectedInstance:
                        ConfigManager.GetInstance().WMIInstanceName = parameter;
                        break;

                    case Command.SetNewPropertiesFilter:
                        var tmp = parameter.Split(',');
                        var propsList = tmp.Select(res => res.Trim()).ToList();
                        ConfigManager.GetInstance().WMIFilterProperties = propsList;
                        break;

                    case Command.SetNewKeyProperty:
                        ConfigManager.GetInstance().WMIKeyProperty = parameter;
                        break;

                    case Command.GetPropertiesFilter:
                        return ConfigManager.GetInstance().WMIFilterProperties;

                    case Command.GetWMINamespace:
                        return ConfigManager.GetInstance().WMINamespace;

                    case Command.GetWMIInstancesOfClass:
                        return WMIWrapper.GetInstance().GetWMIInstances();

                    case Command.GetWMIInstanceAsXML:
                        return WMIWrapper.GetInstance().GetWMIInstanceAsXML();

                    case Command.GetWMIKeyProperty:
                        return ConfigManager.GetInstance().WMIKeyProperty;

                    case Command.GetWMIPropertiesOfInstance:
                        return WMIWrapper.GetInstance().GetWMIPropertiesOfInstance();

                    case Command.GetWMIMethodsOfInstance:
                        return WMIWrapper.GetInstance().GetWMIMethodsOfInstance();

                    case Command.GetWMIClasses:
                        return WMIWrapper.GetInstance().GetWMIClassesInNamespace();

                    default:
                        MessageBox.Show("Unknown command", "ERROR", MessageBoxButton.OK, MessageBoxImage.Error);
                        return new List<string>();
                }
            }
            catch (ConfigException cex)
            {
                MessageBox.Show(
                    "Configuration error. Fix '" + ConfigManager.DefaultProvidersFileName + "' file and reload it: " +
                    cex.Message, "ERROR", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "ERROR", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            return new List<string>();
        }
    }
}
