using System;
using System.Collections.Generic;
using System.Configuration;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TEPL.QMS.Common.Constants;
using TEPL.QMS.Common.Models;

namespace TEPL.QMS.Common
{
    public class ActiveDirectoryHelper
    {
        public static string GetUserProperty(string username, string propertyName)
        {
            using (DirectoryEntry entry = new DirectoryEntry("LDAP://tataelectronics.co.in"))
            {
                using (DirectorySearcher searcher = new DirectorySearcher(entry))
                {
                    searcher.Filter = $"(&(objectClass=user)(sAMAccountName={username}))";
                    SearchResult result = searcher.FindOne();

                    if (result != null)
                    {
                        DirectoryEntry userEntry = result.GetDirectoryEntry();
                        if (userEntry.Properties.Contains(propertyName))
                        {
                            return userEntry.Properties[propertyName].Value.ToString();
                        }
                    }
                }
            }
            return null;
        }
        public static DirectoryEntry GetUserDirectoryEntry(string username)
        {
            using (DirectoryEntry entry = new DirectoryEntry(QMSConstants.LDAPPath))
            {
                using (DirectorySearcher searcher = new DirectorySearcher(entry))
                {
                    searcher.Filter = $"(&(objectClass=user)(sAMAccountName={username}))";
                    SearchResult result = searcher.FindOne();

                    if (result != null)
                    {
                        return result.GetDirectoryEntry();
                    }
                }
            }

            return null;
        }

        public static SearchResultCollection GetUsersFromDirectory(string name)
        {
            try
            {
                using (DirectoryEntry entry = new DirectoryEntry(QMSConstants.LDAPPath))
                {
                    using (DirectorySearcher searcher = new DirectorySearcher(entry))
                    {
                        //searcher.Filter = $"(&(objectClass=user)(displayname=*{name}*))";
                        searcher.Filter = "(&(objectClass=user)(displayname=*" + name + "*))";
                        SearchResultCollection result = searcher.FindAll();

                        if (result != null)
                        {
                            return result;
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                LoggerBlock.WriteLog("Error in GetUsersFromDirectory " + ex.Message.ToString());
                throw ex;
            }
            return null;
        }
        public static SearchResultCollection GetUsersFromADUsingDirectorySearch(string name)
        {
            try
            {
                using (DirectoryEntry entry = new DirectoryEntry(QMSConstants.LDAPPath, ConfigurationManager.AppSettings["ServiceUserName"].ToString(),
                    ConfigurationManager.AppSettings["ServicePassword"].ToString()))
                {
                    object obj1 = entry.NativeObject;
                    using (DirectorySearcher searcher = new DirectorySearcher(entry))
                    {
                        searcher.Filter = "(&(objectClass=user)(displayname=*" + name + "*))";
                        SearchResultCollection result = searcher.FindAll();
                        if (result != null)
                        {
                            return result;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteLog("Error in GetUsersFromADUsingDirectorySearch. Error Message is " + ex.Message.ToString());
                throw ex;
            }
            return null;
        }

        public static List<User> GetUsersFromADUsingPrincipalContext(string name)
        {
            PrincipalSearchResult<Principal> result = null;
            List<User> listUsers = new List<User>();
            try
            {
                LoggerBlock.WriteLog("Message - 1");
                using (var context = new PrincipalContext(ContextType.Domain, "tataelectronics.co.in", "DC=tataelectronics,DC=co,DC=in"))
                {
                    LoggerBlock.WriteLog("Message - 2");
                    var principal = new UserPrincipal(context)
                    {
                        DisplayName = "*"+ name + "*"
                    };
                    LoggerBlock.WriteLog("Message - 3");
                    using (var searcher = new PrincipalSearcher(principal))
                    {
                        result = searcher.FindAll();
                    }
                    LoggerBlock.WriteLog("Message - 4");
                    foreach (Principal child in result)
                    {
                        User obj = new User();
                        obj.LoginID = child.SamAccountName.ToString().Trim();
                        obj.DisplayName = child.DisplayName.ToString().Trim();                       

                        try
                        {
                            if (child.UserPrincipalName != null)
                                obj.EmailID = child.UserPrincipalName.ToString().Trim();

                            var property = "department";
                            var directoryEntry = child.GetUnderlyingObject() as DirectoryEntry;
                            if (directoryEntry.Properties.Contains(property))
                            {
                                string Department = directoryEntry.Properties[property].Value.ToString();
                            }
                        }
                        catch
                        {

                        }

                        listUsers.Add(obj);
                    }
                    LoggerBlock.WriteLog("Message - 5");
                }
            }
            catch (Exception ex)
            {
                LoggerBlock.WriteLog("Error in GetUsersFromADUsingPrincipalContext. Error Message is " + ex.Message.ToString());
                throw ex;
            }
            return listUsers;
        }
        public static User GetUserDirectoryEntryUsingPrinContext(string loginID)
        {
            User obj = new User();
            PrincipalSearchResult<Principal> result = null;
            try
            {
                using (var context = new PrincipalContext(ContextType.Domain, "tataelectronics.co.in", "DC=tataelectronics,DC=co,DC=in"))
                {
                    var principal = new UserPrincipal(context)
                    {
                        SamAccountName = loginID
                    };
                    using (var searcher = new PrincipalSearcher(principal))
                    {
                        result = searcher.FindAll();
                    }
                    foreach (Principal child in result)
                    {
                        obj.LoginID = child.SamAccountName.ToString().Trim();
                        obj.DisplayName = child.DisplayName.ToString().Trim();
                        if (child.UserPrincipalName != null)
                            obj.EmailID = child.UserPrincipalName.ToString().Trim();

                        try
                        {
                            var property = "department";
                            var directoryEntry = child.GetUnderlyingObject() as DirectoryEntry;
                            if (directoryEntry.Properties.Contains(property))
                            {
                                obj.DepartmentName = directoryEntry.Properties[property].Value.ToString();
                            }
                        }
                        catch
                        {

                        }

                    }
                }
            }
            catch(Exception ex)
            {
                LoggerBlock.WriteLog("Error in GetUserDirectoryEntryUsingPrinContext. Error Message is " + ex.Message.ToString());
                throw ex;
            }
            return obj;
        }

        public static DirectoryEntry ValidateUserUsingDirectoryEntry(string username, string password)
        {
            try
            {


                using (DirectoryEntry entry = new DirectoryEntry(QMSConstants.LDAPPath, username, password))
                {
                    using (DirectorySearcher searcher = new DirectorySearcher(entry))
                    {
                        searcher.Filter = $"(&(objectClass=user)(sAMAccountName={username}))";
                        SearchResult result = searcher.FindOne();

                        if (result != null)
                        {
                            return result.GetDirectoryEntry();
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                LoggerBlock.WriteLog("Error in ValidateUserUsingDirectoryEntry. Error Message is " + ex.Message.ToString());
                throw ex;
            }
            return null;
        }

    }
}