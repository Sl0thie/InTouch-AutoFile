using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;



namespace InTouch_AutoFile
{
    public class Stores
    {
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
