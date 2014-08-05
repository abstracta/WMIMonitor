namespace Abstracta.WMIMonitor.Logic
{
    using System;
    using System.Collections.Generic;
    using System.Windows;
    using System.Windows.Controls;

    internal enum Command
    {
        GetWMIClassesInNamespace, 
        GetWMIInstancesOfClass,
        GetWMIPropertiesOfInstance,
        GetWMIMethodsOfClass,
        GetWMIInstanceAsXML,
        ReloadConfiguration,
    }

    internal static class CommandManager
    {
        internal const char Separator = ';';

        internal static void Initialize(ItemCollection resultsList)
        {
            resultsList.Add("");
            try
            {
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
                string wmiServer, wmiNamespace, wmiClassname, wmiInstance;

                switch (command)
                {
                    case Command.ReloadConfiguration:
                        ConfigManager.GetInstance().ReloadConfiguration();
                        return null;

                    case Command.GetWMIInstancesOfClass:
                        wmiServer = parameter.Split(Separator)[0];
                        wmiClassname = parameter.Split(Separator)[1];
                        return WMIWrapper.GetInstance().GetWMIInstancesOfClass(wmiServer, wmiClassname);

                    case Command.GetWMIInstanceAsXML:
                        wmiServer = parameter.Split(Separator)[0];
                        wmiClassname = parameter.Split(Separator)[1];
                        wmiInstance = parameter.Split(Separator)[2];
                        return WMIWrapper.GetInstance().GetWMIInstanceAsXML(wmiServer, wmiClassname, wmiInstance);

                    case Command.GetWMIPropertiesOfInstance:
                        wmiServer = parameter.Split(Separator)[0];
                        wmiClassname = parameter.Split(Separator)[1];
                        wmiInstance = parameter.Split(Separator)[2];
                        return WMIWrapper.GetInstance().GetWMIPropertiesOfInstance(wmiServer, wmiClassname, wmiInstance);

                    case Command.GetWMIMethodsOfClass:
                        wmiServer = parameter.Split(Separator)[0];
                        wmiClassname = parameter.Split(Separator)[1];
                        wmiInstance = parameter.Split(Separator)[2];
                        return WMIWrapper.GetInstance().GetWMIMethodsOfInstance(wmiServer, wmiClassname, wmiInstance);

                    case Command.GetWMIClassesInNamespace:
                        wmiServer = parameter;
                        wmiNamespace = ConfigManager.GetInstance().WMINamespace;
                        return WMIWrapper.GetInstance().GetWMIClassesInNamespace(wmiServer, wmiNamespace);

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
