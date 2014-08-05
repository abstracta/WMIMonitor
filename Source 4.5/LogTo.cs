namespace Abstracta.WMIMonitor
{
    using System;
    using System.Collections.Generic;
    using System.Reflection;

    public abstract class Logging
    {
        public bool Initialized = false;

        public string ErrorMessage = string.Empty;

        public abstract bool Log(object o);

        public abstract bool Log(object[] o);
    }

    public class LoggingToConsole : Logging
    {
        public override bool Log(object o)
        {
            Console.WriteLine(o.ToString());
            return true;
        }

        public override bool Log(object[] objects)
        {
            ErrorMessage = "Not initialized";
            return false;
        }
    }

    public class LoggingToClass : Logging
    {
        private readonly List<Type> _expectedTypes = new List<Type> { typeof(string), typeof(long), typeof(long), typeof(long) };
        
        private readonly object _myObject;

        private readonly MethodInfo _myMethod;

        public LoggingToClass(string logToMethodValue)
        {
            // Interface: string key, long maxValue, long averageValue, long countValue

            var tmp = logToMethodValue.Split(',');
            if (tmp.Length != 3)
            {
                ErrorMessage = "Need the tree values: <AssemblyName>,<Namespace.ClassName>,<StaticMethodName>";
                return;
            }

            var assemblyName = tmp[0].Trim();
            var className = tmp[1].Trim();
            var methodName = tmp[2].Trim();

            Assembly myAssembly;

            try
            {
                myAssembly = Assembly.Load(assemblyName);
                if (myAssembly == null)
                {
                    ErrorMessage = "Assembly not found";
                    return;
                }
            }
            catch (Exception e)
            {
                ErrorMessage = "Assembly not found: " + e.Message;
                return;
            }

            var myType = myAssembly.GetType(className);
            if (myType == null)
            {
                ErrorMessage = "Class not found in assembly";
                return;
            }

            _myMethod = myType.GetMethod(methodName, _expectedTypes.ToArray());
            if (_myMethod == null)
            {
                ErrorMessage = "Method not found in class";
                return;
            }

            _myObject = Activator.CreateInstance(myType);
            if (_myObject == null)
            {
                ErrorMessage = "Object couldn't be created";
                return;
            }

            Initialized = true;
        }

        public override bool Log(object o)
        {
            ErrorMessage = "Not implemented";
            return false;
        }

        public override bool Log(object[] objects)
        {
            if (_myObject == null || _myMethod == null || !Initialized)
            {
                ErrorMessage = "Not initialized";
                return false;
            }

            if (objects.Length != _expectedTypes.Count)
            {
                ErrorMessage = "Unexpected parameters[] lenght";
                return false;
            }

            for (var i = 0; i < _expectedTypes.Count; i++)
            {
                if (_expectedTypes[i] != objects[i].GetType()) 
                {
                    ErrorMessage = "Unexpected item type at [" + i + "] " + _expectedTypes[i]  + " vs " + objects[i].GetType();
                    return false;
                }
            }

            var result = _myMethod.Invoke(_myObject, objects);
            
            return true;
        }
    }
}
