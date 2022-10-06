namespace InTouch_AutoFile
{
    using System;
    using System.Collections.Generic;
    using Outlook = Microsoft.Office.Interop.Outlook;
    using Serilog;

    /// <summary>
    /// Provide methods related to Outlook Contacts.
    /// </summary>
    public class Contacts
    {
        // This dictionary is used to provide a fast lookup for all the outlook contacts over all the contact folders.
        private static readonly Dictionary<string, Tuple<string,string>> emailLookup = new Dictionary<string, Tuple<string, string>>();

        private readonly Outlook.Folder folder;
        public Outlook.Folder Folder
        {
            get { return folder; }
        }

        private readonly Outlook.Folders folders;
        public Outlook.Folders Folders
        {
            get { return folders; }
        }

        public Contacts()
        {
            folder = Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts) as Outlook.Folder;
            folders = folder.Folders;
            CreateEmailLookup();          
        }

        public static Outlook.ContactItem FindContactFromEmailAddress(string emailAddress)
        {
            // Validate the emailAddress parameter.
            if(emailAddress is object)
            {
                emailAddress = emailAddress.ToLower();
            }
            else
            {
                throw new Exception("A value for emailAddress must be provided.");
            }

            // Search with the email lookup first.
            if (emailAddress is object)
            {
                emailAddress = emailAddress.ToLower();
                if (emailLookup.ContainsKey(emailAddress))
                {
                    try
                    {
                        Tuple<string, string> IDs = emailLookup[emailAddress];
                        return Globals.ThisAddIn.Application.Session.GetItemFromID(IDs.Item1,IDs.Item2) as Outlook.ContactItem;
                    }
                    catch(Exception ex)
                    {
                        Log.Error(ex.Message,ex);
                        throw;
                    }
                }
            }

            // Lookup has failed to find the email address.
            return null;
        }

        public static bool DoesLookupContain(string emailAddress)
        {
            if (emailAddress is object)
            {
                if (emailLookup.ContainsKey(emailAddress.ToLower()))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

        private void CreateEmailLookup()
        {
            // Clear the EmailLookup before starting.
            emailLookup.Clear();

            // Add the default contacts folder to EmailLookup.
            AddContactsFolderToEmailLookup(folder);

            // Only add visible contact folders. Outlook has several non visible folders.
            foreach (Outlook.Folder nextFolder in folders)
            {
                switch (nextFolder.Name)
                {
                    case "Recipient Cache":
                        break;

                    case "Organizational Contacts":
                        break;

                    case "PeopleCentricConversation Buddies":
                        break;

                    case "GAL Contacts":
                        break;

                    case "{A9E2BC46-B3A0-4243-B315-60D991004455}":
                        break;

                    case "{06967759-274D-40B2-A3EB-D7F9E73727D7}":
                        break;

                    case "Companies":
                        break;

                    default:
                        AddContactsFolderToEmailLookup(nextFolder);
                        break;
                }
            }
        }

        private static void AddContactsFolderToEmailLookup(Outlook.Folder contactsFolder)
        {
            
            try
            {
                foreach (object nextObject in contactsFolder.Items)
                {
                    if (nextObject is Outlook.ContactItem contact)
                    {
                        AddContactToEmailLookup(contact, contactsFolder);
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex.Message, ex);
            }
        }

        private static void AddContactToEmailLookup(Outlook.ContactItem contact, Outlook.Folder contactsFolder)
        {
            try
            {
                if (contact.Email1Address is object)
                {
                    if (!emailLookup.ContainsKey(contact.Email1Address))
                    {
                        emailLookup.Add(contact.Email1Address.ToLower(), new Tuple<string, string>(contact.EntryID, contactsFolder.StoreID));
                    }
                }

                if (contact.Email2Address is object)
                {
                    if (!emailLookup.ContainsKey(contact.Email2Address))
                    {
                        emailLookup.Add(contact.Email2Address.ToLower(), new Tuple<string, string>(contact.EntryID, contactsFolder.StoreID));
                    }
                }

                if (contact.Email3Address is object)
                {
                    if (!emailLookup.ContainsKey(contact.Email3Address))
                    {
                        emailLookup.Add(contact.Email3Address.ToLower(), new Tuple<string, string>(contact.EntryID, contactsFolder.StoreID));
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex.Message, ex);
            }
        }
    }
}
