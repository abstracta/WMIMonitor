namespace Abstracta.WMIMonitor.LogicInterface
{
    using System.Linq;
    using System.Management;
    
    public class MethodWapper
    {
        private ManagementObject WMIInstance { get; set; }

        public MethodData Method { get; set; }

        public MethodWapper(MethodData methodData, ManagementObject wmiInstance)
        {
            Method = methodData;
            WMIInstance = wmiInstance;
        }
        
        public bool CanBeExecuted()
        {
            return Method.InParameters == null || Method.InParameters.Properties.Count == 0;
        }

        public void Execute()
        {
            WMIInstance.InvokeMethod(Method.Name, null, null); 
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