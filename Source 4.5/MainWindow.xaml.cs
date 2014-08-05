namespace Abstracta.WMIMonitor
{
    using System;
    using System.Collections.Generic;
    using System.Windows;
    using System.Windows.Controls;

    using UIClasses;

    using LogicInterface;
    
    public partial class MainWindow
    {
        public MainWindow()
        {
            InitializeComponent();

            CommandManager.Initialize(Servers.Items, NamespaceText, PropertiesFilter, KeyProperty);

            if (!ServerSelected())
            {
                Classes.IsEnabled = false;
            }
        }

        private const string WMIObjPrefix = "WMI Obj: ";

        private bool ServerSelected()
        {
            return !(Servers.Items.Count == 0 || Servers.SelectedItem == null ||
                   string.Empty == Servers.SelectedItem.ToString());
        }

        private bool ClassSelected()
        {
            return !(Classes.Items.Count == 0 || Classes.SelectedItem == null ||
                   string.Empty == Classes.SelectedItem.ToString());
        }

        private bool InstanceSelected()
        {
            return InstanceList != null && InstanceList.Items.Count != 0 && InstanceList.SelectedItems != null &&
                   !InstanceList.SelectedItem.ToString().StartsWith(CommandManager.ErrorPrefix);
        }
        
        private void Clear(GUIBlocks block)
        {
            switch (block)
            {
                case GUIBlocks.ServerBlock:
                    Servers.Items.Clear();
                    Clear(GUIBlocks.ClassesBlock);
                    break;

                case GUIBlocks.NamespaceBlock:
                    NamespaceText.Text = string.Empty;
                    Clear(GUIBlocks.ClassesBlock);
                    break;

                case GUIBlocks.ClassesBlock:
                    Classes.IsEnabled = false;
                    Classes.Items.Clear();
                    Clear(GUIBlocks.ObjectsBlock);
                    break;

                case GUIBlocks.ObjectsBlock:
                    InstanceList.Items.Clear();
                    Clear(GUIBlocks.PropertiesBlock);
                    Clear(GUIBlocks.MethodsBlock);
                    break;

                case GUIBlocks.PropertiesBlock:
                    PropertiesValues.Text = string.Empty;
                    break;

                case GUIBlocks.MethodsBlock:
                    Methods.RowDefinitions.Clear();
                    Methods.Children.Clear(); 
                    break;
            }
        }

        private void OnWMIServerSelected(object sender, RoutedEventArgs e)
        {
            UpdateWMIClasses();
        }
                
        private void OnWMIClassSelected(object sender, RoutedEventArgs e)
        {
            UpdateWMIObjects();
        }

        private void OnWMIInstanceSelected(object sender, SelectionChangedEventArgs e)
        {
            Clear(GUIBlocks.MethodsBlock);
            Clear(GUIBlocks.PropertiesBlock);
            
            if (ServerSelected() && ClassSelected() && InstanceSelected())
            {
                var instanceName = InstanceList.SelectedItem.ToString().Replace(WMIObjPrefix, string.Empty).Trim(); ;
                CommandManager.Execute(Command.SetSelectedInstance, instanceName);
            }

            UpdateWMIProperties();
            UpdateWMIMethods();
        }

        private void OnNamespaceLostFocus(object sender, RoutedEventArgs e)
        {
            var currentValue = NamespaceText.Text;
            var oldValue = CommandManager.Execute(Command.GetWMINamespace, null) as string;

            if (currentValue != oldValue)
            {
                //MessageBox.Show("Namespace changed");
                CommandManager.Execute(Command.SetNewNamespace, currentValue);
                UpdateWMIClasses();
            }
        }

        private void OnPropertiesLostFocus(object sender, RoutedEventArgs e)
        {
            var currentValue = PropertiesFilter.Text;
            var oldValue = CommandManager.Execute(Command.GetPropertiesFilter, null) as string;

            if (currentValue != oldValue)
            {
                //MessageBox.Show("Properties changed");
                CommandManager.Execute(Command.SetNewPropertiesFilter, currentValue);
                UpdateWMIProperties();
            }
        }

        private void OnKeyPropertyLostFocus(object sender, RoutedEventArgs e)
        {
            var currentValue = KeyProperty.Text;
            var oldValue = CommandManager.Execute(Command.GetWMIKeyProperty, null) as string;

            if (currentValue != oldValue)
            {
                //MessageBox.Show("Key property changed");
                CommandManager.Execute(Command.SetNewKeyProperty, currentValue);
                UpdateWMIObjects();
            }
        }

        private void GetWMIPropsAsXMLAndCopyToClippboard(object sender, RoutedEventArgs e)
        {
            if (!ServerSelected() || !ClassSelected() || !InstanceSelected())
            {
                return;
            }

            var xml = CommandManager.Execute(Command.GetWMIInstanceAsXML, null) as string;

            if (xml != null)
            {
                Clipboard.SetText(xml);
            }
            else
            {
                MessageBox.Show("Empty result", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void UpdateWMIClasses()
        {
            Clear(GUIBlocks.ClassesBlock);

            if (!ServerSelected())
            {
                return;
            }

            CommandManager.Execute(Command.SetSelectedServer, Servers.SelectedItem.ToString());
            var results = CommandManager.Execute(Command.GetWMIClasses, null) as List<string>;

            if (results != null)
            {
                if (results.Count == 0)
                {
                    MessageBox.Show("No classes found in the namespace", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    foreach (var result in results)
                    {
                        Classes.Items.Add(result);
                    }
                }
            }
            else
            {
                MessageBox.Show("Empty result", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            Classes.IsEnabled = true;

            Classes.SelectedItem = CommandManager.CfgMngr.WMIClassName;
            OnWMIClassSelected(null, null);
        }

        private void UpdateWMIObjects(object sender, RoutedEventArgs e)
        {
            UpdateWMIObjects();
        }

        private void UpdateWMIObjects()
        {
            Clear(GUIBlocks.ObjectsBlock);

            if (!ServerSelected() || !ClassSelected())
            {
                return;
            }

            CommandManager.Execute(Command.SetSelectedClass, Classes.SelectedItem.ToString());
            var results = CommandManager.Execute(Command.GetWMIInstancesOfClass, null) as List<string>;

            if (results != null)
            {
                if (results.Count == 0)
                {
                    MessageBox.Show("No data found", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    foreach (var result in results)
                    {
                        InstanceList.Items.Add(WMIObjPrefix + result);
                    }
                }
            }
            else
            {
                MessageBox.Show("Empty result", "Information", MessageBoxButton.OK, MessageBoxImage.Information);
            }

            KeyProperty.Text = CommandManager.Execute(Command.GetWMIKeyProperty, null) as string;
        }

        private void UpdateWMIProperties(object sender, RoutedEventArgs e)
        {
            UpdateWMIProperties();
        }

        private void UpdateWMIProperties()
        {
            Clear(GUIBlocks.PropertiesBlock);

            if (!ServerSelected() || !ClassSelected() || !InstanceSelected())
            {
                return;
            }
            
            var properties = CommandManager.Execute(Command.GetWMIPropertiesOfInstance, null) as List<string>;
            if (properties != null)
            {
                foreach (var prop in properties)
                {
                    PropertiesValues.Text += prop + "\n";
                }
            }
            else
            {
                PropertiesValues.Text += "No properties found\n";
            }
        }

        private void UpdateWMIMethods()
        {
            Clear(GUIBlocks.MethodsBlock);

            if (!ServerSelected() || !ClassSelected() || !InstanceSelected())
            {
                return;
            }
            
            var methods = CommandManager.Execute(Command.GetWMIMethodsOfInstance, null) as List<MethodWapper>;
            if (methods != null)
            {
                var index = 0;
                foreach (var method in methods)
                {
                    Methods.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto, });

                    var methodDefinition = new Label { Content = method.ToString() };
                    Grid.SetRow(methodDefinition, index);
                    Grid.SetColumn(methodDefinition, 0);
                    Methods.Children.Add(methodDefinition);

                    var methodButton = new Button
                    {
                        Content = "Run",
                        IsEnabled = method.CanBeExecuted(),
                    };

                    var methodWapper = method;
                    methodButton.Click += (o, args) =>
                        {
                            try
                            {
                                methodWapper.Execute();
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message, "ERROR", MessageBoxButton.OK, MessageBoxImage.Error);
                            }
                        };

                    Grid.SetRow(methodButton, index);
                    Grid.SetColumn(methodButton, 1);
                    Methods.Children.Add(methodButton);

                    index++;
                }
            }
            else
            {
                Methods.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto, });
                var methodDefinition = new Label { Content = "No methods found" };
                Grid.SetRow(methodDefinition, 0);
                Grid.SetColumn(methodDefinition, 0);
                Methods.Children.Add(methodDefinition);
            }
        }

        private void CopyDetailToClippboard(object sender, RoutedEventArgs e)
        {
            Clipboard.SetText(PropertiesValues.Text);
        }
    }
}
