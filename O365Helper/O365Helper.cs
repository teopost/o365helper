using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Security;
using System.Management.Automation.Remoting;
using System.Management.Automation.Runspaces;
using System.Management.Automation;
using System.Collections.ObjectModel;

namespace O365Wrapper
{
    class O365Helper
    {
        private string _AdminUser = "";
        private string _AdminPassword = "";

        // UserData
        public string userPrincipalName = "";
        public string displayName = "";
        public string firstName = "";
        public string lastName = "";
        public string licenseAssignment = "";
        public string usageLocation = "";
        public string Password = "";

        public string LastMessage = "";
        public string LastError = "";

        public O365Helper(string AdminUser, string AdminPassword)
        {

            this._AdminUser = AdminUser;
            this._AdminPassword = AdminPassword;
        }

        public bool CreateUser()
        {
            LastError = "";
            LastMessage = "";

            string scriptText = "import-module MSOnline\r\n" +
            "$Username = '" + _AdminUser + "'\r\n" +
            "$Password = ConvertTo-SecureString -AsPlainText '" + _AdminPassword + "' -Force\r\n" +
            "$Livecred = New-Object System.Management.Automation.PSCredential $Username, $Password\r\n" +
            "Connect-MsolService -Credential $Livecred\r\n" +
              "New-MsolUser -UserPrincipalName '" + this.userPrincipalName + "' -DisplayName '" + this.displayName + "' -FirstName '" + this.firstName + "' -LastName '" + this.lastName + "' -UsageLocation '" + this.usageLocation + "' -Password '" + this.Password + "'\r\n";

            //  "New-MsolUser -UserPrincipalName '" + userData.userPrincipalName + "' -DisplayName '" + userData.displayName + "' -FirstName '" + userData.firstName + "' -LastName '" + userData.lastName + "' -LicenseAssignment '" + userData.licenseAssignment + "' -UsageLocation '" + userData.usageLocation + "' -Password '" + userData.Password + "'\r\n";

            PowerShell psExec = PowerShell.Create();
            psExec.AddScript(scriptText);
            //psExec.AddCommand("out-string");
            Collection<PSObject> results;
            Collection<ErrorRecord> errors;
            results = psExec.Invoke();

            StringBuilder tmpResult = new StringBuilder();
            if (results != null && results.Count != 0)
            {
                foreach (PSObject obj in results)
                {
                    //stringBuilder.AppendLine(obj.ToString());
                    tmpResult.AppendLine(obj.Members["DisplayName"].Value.ToString());
                }
                LastMessage = tmpResult.ToString();
                return true;
            }
            else
            {
                errors = psExec.Streams.Error.ReadAll();
                StringBuilder tmpErrorMsg = new StringBuilder();
                if (errors != null)
                {
                    foreach (ErrorRecord obj1 in errors)
                    {
                        //stringBuilder.AppendLine(obj.ToString());
                        tmpErrorMsg.AppendLine(obj1.Exception.Message.ToString());
                    }
                }
                this.LastError = tmpErrorMsg.ToString();
                return false;
            }

        }
        public bool UpdateUser()
        {

            LastError = "";
            LastMessage = "";

            string scriptText = "import-module MSOnline\r\n" +
            "$Username = '" + _AdminUser + "'\r\n" +
            "$Password = ConvertTo-SecureString -AsPlainText '" + _AdminPassword + "' -Force\r\n" +
            "$Livecred = New-Object System.Management.Automation.PSCredential $Username, $Password\r\n" +
            "Connect-MsolService -Credential $Livecred\r\n" +
              "Set-MsolUser -UserPrincipalName '" + this.userPrincipalName + "' -DisplayName '" + this.displayName + "' -FirstName '" + this.firstName + "' -LastName '" + this.lastName + "' -UsageLocation '" + this.usageLocation + "' -Password '" + this.Password + "'\r\n";

            //  "New-MsolUser -UserPrincipalName '" + userData.userPrincipalName + "' -DisplayName '" + userData.displayName + "' -FirstName '" + userData.firstName + "' -LastName '" + userData.lastName + "' -LicenseAssignment '" + userData.licenseAssignment + "' -UsageLocation '" + userData.usageLocation + "' -Password '" + userData.Password + "'\r\n";

            PowerShell psExec = PowerShell.Create();
            psExec.AddScript(scriptText);
            //psExec.AddCommand("out-string");
            Collection<PSObject> results;
            Collection<ErrorRecord> errors;
            results = psExec.Invoke();

            StringBuilder tmpResult = new StringBuilder();
            if (results != null && results.Count != 0)
            {
                foreach (PSObject obj in results)
                {
                    //stringBuilder.AppendLine(obj.ToString());
                    tmpResult.AppendLine(obj.Members["DisplayName"].Value.ToString());
                }
                LastMessage = tmpResult.ToString();
                return true;
            }
            else
            {
                errors = psExec.Streams.Error.ReadAll();
                StringBuilder tmpErrorMsg = new StringBuilder();
                if (errors != null)
                {
                    foreach (ErrorRecord obj1 in errors)
                    {
                        //stringBuilder.AppendLine(obj.ToString());
                        tmpErrorMsg.AppendLine(obj1.Exception.Message.ToString());
                    }
                }
                this.LastError = tmpErrorMsg.ToString();
                return false;
            }

        }

        public bool DeleteUser(string UserPrincipalName)
        {
            LastError = "";
            LastMessage = "";

            string scriptText = "import-module MSOnline\r\n" +
            "$Username = '" + _AdminUser + "'\r\n" +
            "$Password = ConvertTo-SecureString -AsPlainText '" + _AdminPassword + "' -Force\r\n" +
            "$Livecred = New-Object System.Management.Automation.PSCredential $Username, $Password\r\n" +
            "Connect-MsolService -Credential $Livecred\r\n" +
              "Remove-MsolUser -UserPrincipalName '" + this.userPrincipalName + "' -force \r\n";

       
            PowerShell psExec = PowerShell.Create();
            psExec.AddScript(scriptText);
            //psExec.AddCommand("out-string");
            Collection<PSObject> results;
            Collection<ErrorRecord> errors;
            results = psExec.Invoke();

            StringBuilder tmpResult = new StringBuilder();
            if (results != null && results.Count != 0)
            {
                foreach (PSObject obj in results)
                {
                    //stringBuilder.AppendLine(obj.ToString());
                    tmpResult.AppendLine(obj.Members["DisplayName"].Value.ToString());
                }
                LastMessage = tmpResult.ToString();
                return true;
            }
            else
            {
                errors = psExec.Streams.Error.ReadAll();
                StringBuilder tmpErrorMsg = new StringBuilder();
                if (errors != null)
                {
                    foreach (ErrorRecord obj1 in errors)
                    {
                        //stringBuilder.AppendLine(obj.ToString());
                        tmpErrorMsg.AppendLine(obj1.Exception.Message.ToString());
                    }
                    
                }
                this.LastError = tmpErrorMsg.ToString();
                return false;
            }

        }
    }
}
