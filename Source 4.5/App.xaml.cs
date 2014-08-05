namespace Abstracta.WMIMonitor
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Runtime.InteropServices;
    using System.Windows;

    using Logic;
    using UIClasses;

    internal enum OutPutSelection
    {
        Gui, 
        Console, 
        File,
        Database,
    }

    public partial class App
    {
        // default value if not specifing a process ID
        private const uint AttachParentProcess = 0x0ffffffff;  

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool AttachConsole(uint dwProcessId);

        private const char ValueSplitter = '=';

        internal OutPutSelection StartUi = OutPutSelection.Gui;

        internal OutputFormatType OutFormatType = OutputFormatType.CSVFormat;

        internal string LogToMethodValue { get; set; }

        internal List<ConsoleParameter> Parameters = new List<ConsoleParameter>
            {
                new ConsoleParameter 
                {
                     Name = "/Console", 
                     Syntaxis = "/Console",
                     Description= "Log results to console: 'WMIMonitor.exe /Console'", 
                     Handler = (app, cp) => { app.StartUi = OutPutSelection.Console; }
                },

                new ConsoleParameter
                {
                    Name = "/File", 
                    Syntaxis = "/File", 
                    Description= "Log results to file: 'WMIMonitor.exe /File > fileName.log'", 
                    Handler = (app, cp) => { app.StartUi = OutPutSelection.File; }
                },

                new ConsoleParameter
                {
                    Name = "/LogToMethod", 
                    Syntaxis = "/LogToMethod=<AssemblyName>,<Namespace.ClassName>,<StaticMethodName>", 
                    Description= "Log results to file: 'WMIMonitor.exe /LogToMethod=LogToDatabase.dll,LogToDatabase.LogToDatabaseHelper,LogValues'", 
                    Handler = (app, cp) => 
                    { 
                        app.StartUi = OutPutSelection.Database;
                        app.LogToMethodValue = cp.Value;
                    }
                },

                new ConsoleParameter
                {
                    Name = "/Format:Detail", 
                    Syntaxis = "/Format:Detail", 
                    Description= "Log results as Detail format: 'WMIMonitor.exe /Console /Format:Detail'", 
                    Handler = (app, cp) => { app.OutFormatType = OutputFormatType.GUIDetailFormat; }
                },

                new ConsoleParameter
                {
                    Name = "/Format:CSV", 
                    Syntaxis = "/Format:CSV", 
                    Description= "Log results as CSV format (default value): 'WMIMonitor.exe /Console /Format:CSV'", 
                    Handler = (app, cp) => { app.OutFormatType = OutputFormatType.CSVFormat; }
                },

                new ConsoleParameter
                {
                    Name = "/Format:XML", 
                    Syntaxis = "/Format:XML", 
                    Description= "Log results in XML format: 'WMIMonitor.exe /Console /Format:XML'", 
                    Handler = (app, cp) => { app.OutFormatType = OutputFormatType.XMLFormat; }
                },

                new ConsoleParameter
                {
                    Name = "/MethodAfterQuery", 
                    Syntaxis = "/MethodAfterQuery=Method1,Method2", 
                    Description= "After quering a WMI instance to get the values of its properties, some WMI methods of the instance are executed. No parameters or return value supported. Example: 'WMIMonitor.exe /Console /MethodAfterQuery=Method1,Method2'", 
                    Handler = (app, cp) => 
                        {
                            ConfigManager.GetInstance().ExecuteWMIMethod = true;
                            ConfigManager.GetInstance().WMIMethodsToExecute = cp.Value;
                        }
                },

                new ConsoleParameter
                {
                    Name = "/?", 
                    Syntaxis = "/?", 
                    Description= "Help", 
                    Handler = (app, cp) => 
                        { 
                            // To write in console
                            AttachConsole(AttachParentProcess);

                            Console.WriteLine(app.Help);
                            Current.Shutdown(); 
                        }
                },
            };

        internal string Help 
        { 
            get
            {
                return Parameters.Aggregate("\nParameters:\n", (current, p) => current + ("\t\t" + p.ToString() + "\n"));
            }
        }

        private void AppOnStartUp(object sender, StartupEventArgs e)
        {
            // Processing arguments
            foreach (var arg in e.Args)
            {
                var pName = arg;
                var pValue = string.Empty;
                if (arg.Contains(ValueSplitter.ToString(CultureInfo.InvariantCulture)))
                {
                    var tmp = arg.Split(ValueSplitter);

                    pName = tmp[0];
                    pValue = tmp[1];
                }

                var cp = FindConsoleParameter(pName);
                if (cp != null)
                {
                    cp.Value = pValue;
                    cp.Handler(this, cp);
                }
            }

            ConfigManager.GetInstance().WMIOutputFormatType = OutFormatType;

            if (StartUi == OutPutSelection.Gui)
            {
                StartupUri = new Uri("MainWindow.xaml", UriKind.Relative);
            }
            else 
            {
                const string srvPrefix = "Server: ";
                // const string instPrefix = "\tInstance: ";
                const string itemPrefix = ""; //  "\t\t";

                Logging logTo = new LoggingToConsole();
                switch (StartUi)
                {
                    case OutPutSelection.Console:
                        AttachConsole(AttachParentProcess);
                        break;

                    case OutPutSelection.Database:
                        AttachConsole(AttachParentProcess);
                        logTo = new LoggingToClass(LogToMethodValue);
                        if (!logTo.Initialized)
                        {
                            Console.WriteLine(logTo.ErrorMessage);
                            Current.Shutdown();
                        }
                        break;
                }

                var wmiServers = ConfigManager.GetInstance().ProviderNames;
                foreach (var wmiServer in wmiServers)
                {
                    ConfigManager.GetInstance().SelectedProvider = ConfigManager.GetInstance().GetProviderByName(wmiServer);
                    var wmiInstanceNames = WMIWrapper.GetInstance().GetWMIInstances();
                    
                    if (StartUi != OutPutSelection.Database)
                    {
                        var filterProperties = ConfigManager.GetInstance().WMIFilterProperties;
                        if (filterProperties.Any(fp => fp == ConfigManager.AllWMIProperties))
                        {
                            filterProperties = WMIWrapper.GetInstance().GetAllWMIPropertyNamesOfInstance();
                        }
                        
                        logTo.Log(srvPrefix + wmiServer);
                        logTo.Log(String.Join(WMIWrapper.CSVSeparator.ToString(CultureInfo.InvariantCulture), filterProperties));
                    }

                    foreach (var wmiInstanceName in wmiInstanceNames)
                    {
                        // Console.WriteLine(instPrefix + instance);
                        ConfigManager.GetInstance().WMIInstanceName = wmiInstanceName;

                        if (StartUi != OutPutSelection.Database)
                        {
                            var detail = WMIWrapper.GetInstance().GetWMIPropertiesOfInstance(OutFormatType);
                            foreach (var item in detail)
                            {
                                logTo.Log(itemPrefix + item);
                            }
                        }
                        else
                        {
                            object[] objects = WMIWrapper.GetInstance().GetWMIPropertiesOfInstanceAsObjectArray();
                            var ok = logTo.Log(objects);
                            if (!ok)
                            {
                                Console.WriteLine(logTo.ErrorMessage);
                            }
                        }

                        if (ConfigManager.GetInstance().ExecuteWMIMethod)
                        {
                            var wmiMethodsOfInstance = WMIWrapper.GetInstance().GetWMIMethodsOfInstance();
                            var methodsToExecute = ConfigManager.GetInstance().WMIMethodsToExecute.Split(WMIWrapper.CSVSeparator);

                            foreach (var methodToExecute in methodsToExecute)
                            {
                                var executed = false;
                                var name = methodToExecute.Trim();
                                foreach (var method in wmiMethodsOfInstance.Where(method => method.CanBeExecuted() && (method.Method.Name == name)))
                                {
                                    if (executed) continue;

                                    try
                                    {
                                        method.Execute();
                                        executed = true;
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine(@"Error: " + ex.Message);
                                    }
                                }

                                if (!executed)
                                {
                                    Console.WriteLine(@"Error: Method '" + name + @"' couldn't be executed. Wrong name? Static methods aren't also supported yet.");
                                }
                            }
                        }
                    }
                }

                Current.Shutdown();
            }
        }

        private ConsoleParameter FindConsoleParameter(string name)
        {
            return Parameters.FirstOrDefault(p => name == p.Name);
        }
    }
}
