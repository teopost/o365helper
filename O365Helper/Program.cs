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

/*
 http://o365info.com/using-remote-powershell-to-manage_212/
 http://techathon.mytechlabs.com/create-user-on-microsoft-outlookoffice365-using-c-with-powershell-api/
 http://techathon.mytechlabs.com/microsoft-outlookoffice365-issue-or-error-with-64-bit-with-powershell-api/
 */

namespace O365Wrapper
{

    class Program
    {

        static void Main(string[] args)
        {

            try
            {
                string AdminUser = "";
                string AdminPassword = "";

                var o = new O365Helper(AdminUser, AdminPassword);

                // Insert New user
                // ----------------
                o.userPrincipalName = "User09@didanet.onmicrosoft.com";
                o.displayName = "User from c#";
                o.firstName = "Nome";
                o.lastName = "Cognome 2";
                //  o.licenseAssignment = "";
                o.Password = "pippo123!22";
                o.usageLocation = "IT";

                Console.WriteLine("Creating user...");
                bool RetValCreate = o.CreateUser();

                if (RetValCreate)
                    Console.WriteLine("User is successfully created: " + o.LastMessage);
                else
                    Console.WriteLine(o.LastError);

                Console.WriteLine("Press any key to continue");
                Console.ReadLine();


                // Delete user
                // -----------
                Console.WriteLine("Deleting user...");
                bool RetValDelete = o.DeleteUser("User09@didanet.onmicrosoft.com");

                if (RetValDelete)
                    Console.WriteLine("User deleted: " + o.LastMessage);
                else
                    Console.WriteLine(o.LastError);
                
                
                Console.WriteLine("Press any key to continue");
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
            }

        }
     

    }



}
