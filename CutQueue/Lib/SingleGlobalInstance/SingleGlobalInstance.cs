using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Threading;

namespace CutQueue.Lib.SingleGlobalInstance
{
    class SingleGlobalInstance : IDisposable
    {
        private Mutex mutex;

        public bool HasHandle { get; set; }

        private void InitMutex()
        {
            Assembly executingAssembly = Assembly.GetExecutingAssembly();
            string appGuid = ((GuidAttribute)executingAssembly.GetCustomAttributes(typeof(GuidAttribute), false).GetValue(0)).Value;
            string mutexId = string.Format("Global\\{{{0}}}", appGuid);
            mutex = new Mutex(false, mutexId);

            MutexRights mutexRights = MutexRights.FullControl;
            AccessControlType accessControlType = AccessControlType.Allow;
            WellKnownSidType sid = WellKnownSidType.WorldSid;
            var allowEveryoneRule = new MutexAccessRule(new SecurityIdentifier(sid, null), mutexRights, accessControlType);
            var securitySettings = new MutexSecurity();
            securitySettings.AddAccessRule(allowEveryoneRule);
            mutex.SetAccessControl(securitySettings);
        }

        public SingleGlobalInstance(int timeOut)
        {
            InitMutex();
            try
            {
                if (timeOut < 0)
                    HasHandle = mutex.WaitOne(Timeout.Infinite, false);
                else
                    HasHandle = mutex.WaitOne(timeOut, false);

                if (HasHandle == false)
                    throw new TimeoutException("Timeout waiting for exclusive access on SingleInstance");
            }
            catch (AbandonedMutexException)
            {
                HasHandle = true;
            }
        }


        public void Dispose()
        {
            if (mutex != null)
            {
                if (HasHandle)
                {
                    mutex.ReleaseMutex();
                }
                mutex.Close();
            }
        }
    }
}
