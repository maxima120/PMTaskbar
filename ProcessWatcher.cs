using System;
using System.Data;
using System.Management;

namespace PMTaskbar
{
    // taken from : https://www.codeproject.com/Articles/12138/Process-Information-and-Notifications-using-WMI
    public class ProcessWatcher : IDisposable
    {
        const string scope = @"\\.\root\CIMV2";

        // events to subscribe
        public EventHandler Started = null;
        public EventHandler Terminated = null;

        // WMI event watcher
        private ManagementEventWatcher watcher;
        private bool disposedValue;

        // The constructor uses the application name like notepad.exe
        // And it starts the watcher
        public ProcessWatcher(string appName)
        {
            // querry every 2 seconds
            string poll = "2";

            string queryString = $"SELECT * FROM __InstanceOperationEvent WITHIN {poll} WHERE TargetInstance ISA 'Win32_Process' AND TargetInstance.Name = '{appName}'";

            // create the watcher and start to listen
            watcher = new ManagementEventWatcher(scope, queryString);
            watcher.EventArrived += new EventArrivedEventHandler(OnEventArrived);
            watcher.Start();
        }

        public static DataTable RunningProcesses()
        {
            // The second way of constructing a query
            string queryString =
                "SELECT Name, ProcessId, Caption, ExecutablePath" +
                "  FROM Win32_Process";

            ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, queryString);
            ManagementObjectCollection processes = searcher.Get();

            DataTable result = new DataTable();
            result.Columns.Add("Name", Type.GetType("System.String"));
            result.Columns.Add("ProcessId", Type.GetType("System.Int32"));
            result.Columns.Add("Caption", Type.GetType("System.String"));
            result.Columns.Add("Path", Type.GetType("System.String"));

            foreach (ManagementObject mo in processes)
            {
                DataRow row = result.NewRow();
                row["Name"] = mo["Name"].ToString();
                row["ProcessId"] = Convert.ToInt32(mo["ProcessId"]);
                if (mo["Caption"] != null)
                    row["Caption"] = mo["Caption"].ToString();
                if (mo["ExecutablePath"] != null)
                    row["Path"] = mo["ExecutablePath"].ToString();
                result.Rows.Add(row);
            }
            return result;
        }
        private void OnEventArrived(object sender, System.Management.EventArrivedEventArgs e)
        {
            try
            {
                string eventName = e.NewEvent.ClassPath.ClassName;

                if (eventName.CompareTo("__InstanceCreationEvent") == 0)
                {
                    // Started
                    if (Started != null)
                        Started(this, e);
                }
                else if (eventName.CompareTo("__InstanceDeletionEvent") == 0)
                {
                    // Terminated
                    if (Terminated != null)
                        Terminated(this, e);

                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex.Message);
            }
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                }

                watcher.Stop();
                watcher.Dispose();

                disposedValue = true;
            }
        }

        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
        ~ProcessWatcher()
        {
            Dispose(disposing: false);
        }
    }
}