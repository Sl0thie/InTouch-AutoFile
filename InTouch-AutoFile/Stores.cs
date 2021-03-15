﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;



namespace InTouch_AutoFile
{


    public class ITStore
    {
        private string displayName;

        public string DisplayName
        {
            get { return displayName; }
            set { displayName = value; }
        }

        private string storeID;

        public string StoreID
        {
            get { return storeID; }
            set { storeID = value; }
        }

        private Outlook.MAPIFolder rootFolder;
        public Outlook.MAPIFolder RootFolder
        {
            get { return rootFolder; }
            set { rootFolder = value; }
        }

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
            Outlook.Store store = null;
            string storeList = string.Empty;

            try
            {
                stores = Globals.ThisAddIn.Application.Session.Stores;

                for (int i = 1; i <= stores.Count; i++)
                {
                    store = stores[i];

                    Op.LogMessage("Store : " + store.DisplayName);
                    Op.LogMessage("ExchangeStoreType : " + store.ExchangeStoreType.ToString());
                    Op.LogMessage("FilePath : " + store.FilePath);
                    Op.LogMessage("StoreID : " + store.StoreID);
                    Op.LogMessage("IsDataFileStore : " + store.IsDataFileStore.ToString());

                    ITStore nextStore = new ITStore();
                    nextStore.DisplayName = store.DisplayName;
                    nextStore.RootFolder = store.GetRootFolder();
                    nextStore.StoreID = store.StoreID;
                    storesLookup.Add(store.DisplayName, nextStore) ;

                    if (store is object)
                        Marshal.ReleaseComObject(store);
                }
                Op.LogMessage(storeList);
            }
            finally
            {
                if (stores is object)
                    Marshal.ReleaseComObject(stores);
            }
        }
    }
}