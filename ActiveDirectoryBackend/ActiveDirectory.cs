using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;
using System.Net;
using System.DirectoryServices;
using System.Windows.Forms;
using System.DirectoryServices.ActiveDirectory;
using System.Collections;
using System.DirectoryServices.AccountManagement;
using System.IO;

namespace ActiveDirectoryBackend
{
    public class ActiveDirectory
    {
        //Class member variables that are used multiple times.
        #region Member Variables

        private DirectoryEntry de = null;

        private string m_email = null;
        private string m_name = null;
        private string m_password = null;
        private string m_subject = null;
        private string m_body = null;
        private string m_extractedUserData = null;
        private string m_userGroup = null;
        private string m_errorCode = null;
        private bool m_returnValue = false;
        private string m_ldapAddress = "LDAP://something.org";
        #endregion

        //Get and Sets for all member variables. I hate properties.
        #region Get and Sets

        public void setName(string name)
        {
            m_name = name;
        }

        public string getName()
        {
            return m_name;
        }

        public void setPassword(string pass)
        {
            m_password = pass;
        }

        public string getPassword()
        {
            return m_password;
        }

        public void setEmail(string email)
        {
            m_email = email;
        }

        public string getEmail()
        {
            return m_email;
        }

        public void setSubject(string subject)
        {
            m_subject = subject;
        }

        public string getSubject()
        {
            return m_subject;
        }

        public void setBody(string body)
        {
            m_body = body;
        }

        public string getBody()
        {
            return m_body;
        }

        public void setExtractedData(string data)
        {
            m_extractedUserData = data;
        }

        public string getExtractedData()
        {
            return m_extractedUserData;
        }

        public void setErrorCode(string errorCode)
        {
            m_errorCode = errorCode;
        }

        public string getErrorCode()
        {
            return m_errorCode;
        }

        public void setReturnValue(bool returnValue)
        {
            m_returnValue = returnValue;
        }

        public bool getReturnValue()
        {
            return m_returnValue;
        }

        public void setUserGroup(string userGroup)
        {
            m_userGroup = userGroup;
        }

        public string getUserGroup()
        {
            return m_userGroup;
        }

        #endregion

        //****************************************************************
        // Date: 12-2-15 Author: Dominick Raimondi
        // Purpose: Default Constructor
        // Items to change: None
        //*****************************************************************
        public ActiveDirectory()
        {
            de = connectWithActiveDirectoryTLD();
        }

        //****************************************************************
        // Date: 8-18-15 Author: Dominick Raimondi
        // Purpose: Sends an email using SMTP
        // Items to change: None
        //*****************************************************************
        #region sendEmail
        public void sendEmail(string email, string name, string subject, string body, string info)
        {
            m_email = email;
            m_name = name;
            m_subject = subject;
            m_body = body;


            try
            {
                MailAddress fromAddress = new MailAddress("CHANGE", "CHANGE");
                MailAddress toAddress = new MailAddress(m_email, m_name);

                SmtpClient smtp = new SmtpClient();

                smtp.Host = "CHANGE";
                smtp.Port = 25;
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp.UseDefaultCredentials = false;


                MailMessage sendingMail = new MailMessage(fromAddress, toAddress);
                sendingMail.Subject = m_subject;
                sendingMail.Body = m_body;

                smtp.Send(sendingMail);


                Console.WriteLine("Sent");
                m_errorCode = "Email sent.";
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        #endregion

        //****************************************************************
        // Date: 8-18-15 Author: Dominick Raimondi
        // Purpose: Connects to the very top level of the AD. Used in all
        // search functions to completely search AD.
        // Items to change: None
        //*****************************************************************
        #region connectWithActiveDirectoryTLD()
        public DirectoryEntry connectWithActiveDirectoryTLD()
        {
            try
            {
                string ldapAddress = "LDAP://something.com";
                // Can jsut use secure if no SSL is set up
                DirectoryEntry root = new DirectoryEntry(ldapAddress, "USERNAME", "PASSWORD", AuthenticationTypes.SecureSocketsLayer);
                object connected = root.NativeObject;
                Console.WriteLine("Connected");
                return root;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.GetType());
                Console.WriteLine(ex.Message);
                m_errorCode = ex.Message;
                return null;
            }
        }
        #endregion

        //****************************************************************
        // Date: 7-29-15 Author: Dominick Raimondi
        // Purpose: Variations of searching the AD. These functions search 
        // for sAMAccount. One returns a bool another the accout name.
        // Items to change: None
        //*****************************************************************
        #region searchActiveDirectory
        public bool searchActiveDirectory(string samAccount)
        {

            //DirectoryEntry de = connectWithActiveDirectoryTLD();
            DirectorySearcher deSearch = new DirectorySearcher();
            deSearch.SearchRoot = de;

            deSearch.Filter = "(&(objectClass=user)(sAMAccountName=" + samAccount + "))";
            deSearch.SearchScope = SearchScope.Subtree;
            SearchResult results = null;

            try
            {
                results = deSearch.FindOne();
            }
            catch (Exception notFound)
            {
                Console.WriteLine(notFound.ToString());
            }

            if (results == null)
            {
                Console.WriteLine("Not Found!");
                return false;
            }
            else
            {
                Console.WriteLine("Found at least one");
                return true;
            }
        }

        public string searchActiveDirectoryReturn(string samAccount)
        {

            //DirectoryEntry de = connectWithActiveDirectoryTLD();
            DirectorySearcher deSearch = new DirectorySearcher();
            deSearch.SearchRoot = de;

            deSearch.Filter = "(&(objectClass=user)(sAMAccountName=" + samAccount + "))";
            deSearch.SearchScope = SearchScope.Subtree;

            SearchResult results = deSearch.FindOne();

            if (results == null)
            {
                Console.WriteLine("Not Found!");
                return null;
            }
            else
            {
                Console.WriteLine("Found at least one");
                return results.GetDirectoryEntry().Properties["distinguishedName"].Value.ToString();
            }
        }

        #endregion

        //****************************************************************
        // Date: 7-29-15 Author: Dominick Raimondi
        // Purpose: Searches the users current Directory for same names. 
        // Essentially this finds duplicates within the same OU. 
        // Items to change: None
        //*****************************************************************
        #region searchActiveDirectoryCN
        public bool searchActiveDirectoryCN(string commonName)
        {

            //DirectoryEntry de = connectWithActiveDirectoryTLD();
            DirectorySearcher deSearch = new DirectorySearcher();
            deSearch.SearchRoot = de;

            deSearch.Filter = "(&(objectClass=user)(CN=" + commonName + "))";
            deSearch.SearchScope = SearchScope.Subtree;
            SearchResult results = deSearch.FindOne();

            if (results == null)
            {
                Console.WriteLine("No Matches. Name can be used.");
                return false;
            }
            else
            {
                Console.WriteLine("Name found. Name cannot be used.");
                return true;
            }
        }
        #endregion

        //****************************************************************
        // Date: 8-18-15 Author: Dominick Raimondi
        // Purpose: Searches the Active Directory for the users email. Only
        // one specific email should be in the entire AD.
        // Items to change: None
        //*****************************************************************
        #region searchActiveDirectoryEmail
        public bool searchActiveDirectoryEmail(string email)
        {

            //DirectoryEntry de = connectWithActiveDirectoryTLD();
            DirectorySearcher deSearch = new DirectorySearcher();
            deSearch.SearchRoot = de;

            deSearch.Filter = "(&(objectClass=user)(mail=" + email + "))";
            deSearch.SearchScope = SearchScope.Subtree;
            SearchResult results = deSearch.FindOne();

            if (results == null)
            {
                Console.WriteLine("Email not in Active Directory.");
                return false;
            }
            else
            {
                Console.WriteLine("Email in Active Directory.");
                return true;
            }
        }
        #endregion

        //****************************************************************
        // Date: 8-18-15 Author: Dominick Raimondi
        // Purpose: Pulls multiple pieces of information from user
        // Items to change: LDAP Address
        //*****************************************************************
        #region extractDataFromUser

        public void extractDataFromUser()
        {
            m_extractedUserData = null;

            //DirectoryEntry de = new DirectoryEntry(ldapAddress, "USERNAME", "PASSWORD", AuthenticationTypes.SecureSocketsLayer);

            DirectorySearcher deSearch = new DirectorySearcher();
            deSearch.SearchRoot = de;
            deSearch.Filter = "(&(objectClass=user)(sAMAccountName=" + m_name + "))";
            deSearch.SearchScope = SearchScope.Subtree;
            SearchResult results = deSearch.FindOne();

            string[] namesArray;

            namesArray = results.GetDirectoryEntry().Properties["distinguishedName"].Value.ToString().Split(',');
            m_userGroup = results.GetDirectoryEntry().Properties["memberOf"].Value.ToString();
            m_extractedUserData = namesArray.ToString();

        }

        #endregion

        //****************************************************************
        // Date: 8-19-15 Author: Dominick Raimondi
        // Purpose: Checks UserName and Email combination to make sure they
        // match in AD and resets the password. Emails the new password to
        // user using the email function and passing the password. 
        // Items to change: None (Keep an eye on log file).
        //*****************************************************************
        #region resetPassword

        public bool resetPassword(string username, string email)
        {
            // Log file
            StreamWriter file = new StreamWriter("c:\\NamedUserAccountLogFile_PasswordReset.txt", true);
            DateTime date = DateTime.Now;
            // Initial searching of the active directory. Searches the entire AD for the
            // sAMAccount name and email. If either don't exist, well obviously they can't 
            // reset a non existing account.
            if (searchActiveDirectory(username) == true && searchActiveDirectoryEmail(email) == true)
            {
                // Extracts the email associated with the sAMAccount entered
                // to be used later to check if it matches the email entered.

                DirectorySearcher deSearch = new DirectorySearcher();
                deSearch.SearchRoot = de;

                deSearch.Filter = "(&(objectClass=user)(sAMAccountName=" + username + "))";
                deSearch.SearchScope = SearchScope.Subtree;
                SearchResult results = deSearch.FindOne();

                m_email = results.GetDirectoryEntry().Properties["mail"].Value.ToString();
                m_name = results.GetDirectoryEntry().Properties["cn"].Value.ToString();
                m_email = m_email.ToLower();
                email = email.ToLower();

                // Email check. The email was the email extracted from the sAMAccount the user entered.
                // If the user entered the same email that was taken from the account name, the password
                // will be reset.
                if (email == m_email)
                {
                    file.WriteLine(username + "," + email + "," + "PassReset," + "Reset," + date.ToString() + "\n");
                    Console.WriteLine("Match");
                    Console.WriteLine(m_name);

                    string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
                    char[] stringChars = new char[8];
                    Random random = new Random();

                    for (int i = 0; i < stringChars.Length; i++)
                    {
                        stringChars[i] = chars[random.Next(chars.Length)];
                    }

                    m_password = new String(stringChars);

                    m_password = "#" + m_password;

                    results.GetDirectoryEntry().Invoke("SetPassword", new object[] { m_password });
                    results.GetDirectoryEntry().CommitChanges();

                    DirectoryEntry grabUser = results.GetDirectoryEntry();
                    grabUser.Properties["pwdLastSet"].Value = 0;
                    grabUser.CommitChanges();
                    grabUser.Close();

                    results.GetDirectoryEntry().Close();


                    // New password is reset in the same manner a new account password is created.
                    // New password will be sent to the email of the account holder as to protect
                    // from unwanted authorization. 
                    m_subject = "Password Reset";

                    m_body = "Hello,\n\nYou have requested a new password. Your password is " +
                        m_password + ". Contact the Administrator if you are still having issues.\n\nThank you.";

                    sendEmail(m_email, m_name, m_subject, m_body, null);
                    file.Close();
                    m_errorCode = "Please check your email for new password.";
                    return true;
                }
                else
                {
                    // Logging for mismatch username and email, meaning that the user existed
                    // but entered the wrong email associated with the account. Possible security threat.
                    file.WriteLine(username + "," + email + "," + "PassReset," + "MismatchEmail," + date.ToString() + "\n");
                    file.Close();
                    Console.WriteLine("No Match");
                    m_errorCode = "Username and/or email isn't correct.\n";
                    return false;
                }
            }
            else
            {
                // Logging for unknown username or email, meaning that the user entered a username or email
                // that is not in the active directory at all. Possible security threat.
                m_errorCode = "Username and/or email isn't correct.\n";
                file.WriteLine(username + "," + email + "," + "PassReset," + "NoUserNameOrEmail," + date.ToString() + "\n");
                file.Close();
                return false;
            }
        }

        #endregion

    }
}

