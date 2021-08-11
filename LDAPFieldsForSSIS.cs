#region Help:  Introduction to the Script Component
#endregion

#region Namespaces
using System;
using System.Data;
using System.DirectoryServices;
using Microsoft.SqlServer.Dts.Pipeline.Wrapper;
using Microsoft.SqlServer.Dts.Runtime.Wrapper;
#endregion

// Needs to be run in SSIS and connected to fields. This will run through LDAP fields and users posting them to a database using SSIS.

[Microsoft.SqlServer.Dts.Pipeline.SSISScriptComponentEntryPointAttribute]
public class ScriptMain : UserComponent
{
    public override void CreateNewOutputRows()
    {
        String[] companyAffiliate = new string[] { "A", "B", "C", "D" };

        foreach (String Affiliate in companyAffiliate)
        {
            try
            {
                // Sign into Active Directory
                DirectoryEntry myLdapConnection = createDirectoryEntry(Affiliate);

                // Pull Active Directory Fields
                DirectorySearcher search = new DirectorySearcher(myLdapConnection);

                // Prevent Data pull limitations
                search.PageSize = 999999;

                SearchResultCollection allUsers = search.FindAll();

                // Loop through active directory users
                foreach (SearchResult result in allUsers)
                {
                    if (result != null)
                    {
                        // if user exists, cycle through all LDAP fields  
                        ResultPropertyCollection fields = result.Properties;


                        foreach (String ldapField in fields.PropertyNames)
                        {
                            // cycle through objects in each field
                            foreach (Object myCollection in fields[ldapField])
                            {

                                // Pull samAccountName for primary key to identify all fields for a user
                                foreach (Object samAccountName in result.Properties["SamAccountName"])
                                {

                                    Output0Buffer.AddRow();
                                    Output0Buffer.Affiliate = LeftOfString(Affiliate, 255);
                                    Output0Buffer.samAccountName = LeftOfString(samAccountName.ToString(), 200);
                                    Output0Buffer.ldapField = LeftOfString(ldapField, 50);
                                    Output0Buffer.myCollection = LeftOfString(myCollection.ToString(), 3000);
                                }
                            }

                        }
                    }

                }

                search.Dispose();

            }

            catch (ArgumentNullException nullException)
            {
                Console.WriteLine("Exception caught:\n\n" + nullException.ToString());
            }
        }

    }



    static DirectoryEntry createDirectoryEntry(String affiliate)
    {
        // create and return new LDAP connection with desired settings
        DirectoryEntry ldapConnection = new DirectoryEntry("LDAP://");
        ldapConnection.Path = "LDAP:///ou=" + affiliate + ",DC=,DC=com";
        ldapConnection.AuthenticationType = AuthenticationTypes.Secure;
        return ldapConnection;
    }

    public static String LeftOfString(String field, int characterLength)
    {
        if (field.Length >= characterLength)
        {
            return field.Substring(0, characterLength);
        }
        else
        {
            return field;
        }
    }




}
