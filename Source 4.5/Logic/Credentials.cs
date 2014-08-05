namespace Abstracta.WMIMonitor.Logic
{
    public class Credentials
    {
    }

    public class CurrentUser : Credentials
    {
    }

    public class UserPasswAuthentication : Credentials
    {
        public string User { get; set; }

        public string Password { get; set; }
    }
}
