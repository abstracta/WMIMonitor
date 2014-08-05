using System.Linq;
using System.Management;

namespace Abstracta.WMIMonitor.Logic
{
    public class MethodWapper
    {
        private ManagementObject WMIInstance { get; set; }

        public MethodWapper(MethodData methodData, ManagementObject wmiInstance)
        {
            Method = methodData;
            WMIInstance = wmiInstance;
        }

        public MethodData Method { get; set; }

        public bool CanBeExecuted()
        {
            return Method.InParameters == null || Method.InParameters.Properties.Count == 0;
        }

        public void Execute()
        {
            try
            {
                WMIInstance.InvokeMethod(Method.Name, null, null);
            }
            catch
            {
            }
        }

        public new string ToString()
        {
            var tmpStr = Method.Name;

            if (Method.InParameters != null)
            {
                tmpStr += "(";
                tmpStr = Method.InParameters.Properties.Cast<PropertyData>()
                               .Aggregate(tmpStr,
                                          (current, propertyData) =>
                                          current + (propertyData.Type + " " + propertyData.Name + ", "));
                tmpStr += ")";

                tmpStr = tmpStr.Replace(", )", ")");
            }
            else
            {
                tmpStr += "()";
            }

            return tmpStr;
        }
    }
}