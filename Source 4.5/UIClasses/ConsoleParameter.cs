namespace Abstracta.WMIMonitor.UIClasses
{
    using System;

    internal class ConsoleParameter
    {
        internal string Name { get; set; }

        internal string Syntaxis { get; set; }

        internal string Description { get; set; }

        internal string Value { get; set; }

        internal Action<App, ConsoleParameter> Handler { get; set; }

        internal new string ToString()
        {
            return Syntaxis + ": \t" + (Name.Length < 6 ? "\t\t" : Name.Length < 12 ? "\t" : "") + Description;
        }
    }
}
