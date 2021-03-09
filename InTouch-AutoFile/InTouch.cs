namespace InTouch_AutoFile
{
    public static class InTouch
    {
        private static readonly Contacts contacts = new Contacts();
        public static Contacts Contacts
        {
            get { return contacts; }
        }

        private static readonly TaskManager taskManager = new TaskManager();
        public static TaskManager TaskManager
        {
            get { return taskManager; }
        }
    }
}
