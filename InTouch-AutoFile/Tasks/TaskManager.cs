namespace InTouch_AutoFile
{
    using System;
    using System.Threading;
    using System.Collections.Concurrent;
    using Serilog;
    using InTouch_AutoFile.Tasks;

    /// <summary>
    /// TaskManager manages the background tasks.
    /// </summary>
    /// <remarks>The TaskManager is used to manage the background tasks that perform 
    /// operations such as moving emails from the Inbox. It provides a queue to store tasks
    /// and executes them one at a time.</remarks>
    public class TaskManager
    {
        private bool taskRunning = false; // Is a task currently running.
        private readonly ConcurrentQueue<Action> backgroundTasks = new ConcurrentQueue<Action>(); // Queue for the tasks.
        public ConcurrentQueue<Action> BackgroundTasks
        {
            get
            {
                return backgroundTasks;
            }
        }
        private Action currentAction;
        public Action CurrentAction
        {
            get
            {
                return currentAction;
            }

            set
            {
                currentAction = value;
            }
        }

        private readonly TaskAddinSetup taskAddinSetup; // A task to manage add-in setup.
        private readonly TaskFileInbox taskFileIndox; // A task to manage sorting the Inbox folder.
        private readonly TaskFileSentItems taskFileSentItems; // A task to manage sorting the Sent Items folder.
        private readonly TaskMonitorAliases taskMonitorAliases; // A task to monitor aliases.
        private readonly TaskFindIcon taskFindIcon;

        public TaskManager()
        {
            // Create task objects.
            taskAddinSetup = new TaskAddinSetup(TaskFinished);
            taskFileIndox = new TaskFileInbox(TaskFinished);
            taskFileSentItems = new TaskFileSentItems(TaskFinished);
            taskMonitorAliases = new TaskMonitorAliases(TaskFinished);
            taskFindIcon = new TaskFindIcon(TaskFinished);


            // Configure and start a background thread.
            Thread backgroundThread = new Thread(new ThreadStart(BackgroundProcess))
            {
                Name = "AF.TaskManager",
                IsBackground = true,
                Priority = ThreadPriority.Normal
            };
            backgroundThread.SetApartmentState(ApartmentState.STA);
            backgroundThread.Start();
        }

        /// <summary>
        /// Callback for when the current task is finished.
        /// </summary>
        public void TaskFinished()
        {
            taskRunning = false;
        }

        /// <summary>
        /// Main loop for managing tasks.
        /// </summary>
        private void BackgroundProcess()
        {
            bool DoLoop = true; // Keeps the main loop running.
            while (DoLoop)
            {
                Thread.Sleep(5000);
                try
                {
                    if ((!taskRunning) && (!BackgroundTasks.IsEmpty))
                    {
                        taskRunning = true;
                        BackgroundTasks.TryDequeue(out currentAction);
                        //Log.Information("TaskManager Starting " + currentAction.Target + "." + currentAction.Method.Name.ToString());
                        currentAction.Invoke();
                    }            
                }
                catch (Exception ex) 
                { 
                    Log.Error(ex.Message, ex);
                    throw; 
                }
            }
        }

        public void EnqueueAddinSetupTask()
        {
            backgroundTasks.Enqueue(taskAddinSetup.RunTask);
        }

        public void EnqueueSentItemsTask()
        {
            // Check if the task is already queued before adding to the queue.
            foreach (var task in backgroundTasks)
            {
                if (task.GetType() == typeof(TaskFileSentItems))
                {
                    return;
                }
            }

            backgroundTasks.Enqueue(taskFileSentItems.RunTask);
        }

        public void EnqueueInboxTask()
        {
            // Check if the task is already queued before adding to the queue.
            foreach (var task in backgroundTasks)
            {
                if(task.GetType() == typeof(TaskFileInbox))
                {
                    return;
                }
            }

            backgroundTasks.Enqueue(taskFileIndox.RunTask);
        }

        public void EnqueueMonitorAliases()
        {
            backgroundTasks.Enqueue(taskMonitorAliases.RunTask);
        }

        public void EnqueueFindIcon()
        {
            backgroundTasks.Enqueue(taskFindIcon.RunTask);
        }
    }
}
