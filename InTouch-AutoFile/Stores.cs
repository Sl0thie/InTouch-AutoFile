namespace InTouch_AutoFile
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Runtime.InteropServices;
    using System.Text;
    using System.Threading.Tasks;
    using Outlook = Microsoft.Office.Interop.Outlook;
    using Serilog;

    public class ITStore
    {
        public string DisplayName
        {
            get { return displayName; }
            set { displayName = value; }
        }
        private string displayName;

        public string StoreID
        {
            get { return storeID; }
            set { storeID = value; }
        }
        private string storeID;
        
        public Outlook.MAPIFolder RootFolder
        {
            get { return rootFolder; }
            set { rootFolder = value; }
        }
        private Outlook.MAPIFolder rootFolder;
    }

    public class Stores
    {
        private readonly Dictionary<string,ITStore> storesLookup = new Dictionary<string, ITStore>();

        public Dictionary<string,ITStore> StoresLookup
        {
            get { return storesLookup; }
        }

        public Stores()
        {
            Outlook.Stores stores = null;
            Outlook.Store store;
            string storeList = string.Empty;

            try
            {
                stores = Globals.ThisAddIn.Application.Session.Stores;

                for (int i = 1; i <= stores.Count; i++)
                {
                    store = stores[i];

                    ITStore nextStore = new ITStore
                    {
                        DisplayName = store.DisplayName,
                        RootFolder = store.GetRootFolder(),
                        StoreID = store.StoreID
                    };
                    storesLookup.Add(store.DisplayName, nextStore) ;

                    if (store is object)
                    {
                        Marshal.ReleaseComObject(store);
                    }
                }

                Log.Information(storeList);
            }
            finally
            {
                if (stores is object)
                {
                    Marshal.ReleaseComObject(stores);
                }
            }
        }
    }
}
