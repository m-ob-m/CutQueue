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
        //edit by user "jitbit" - renamed private fields to "_"
        public bool _hasHandle = false;
        Mutex _mutex;

        private void InitMutex()
        {
            Assembly executingAssembly = Assembly.GetExecutingAssembly();
            string appGuid = ((GuidAttribute)executingAssembly.GetCustomAttributes(typeof(GuidAttribute), false).GetValue(0)).Value;
            string mutexId = string.Format("Global\\{{{0}}}", appGuid);
            _mutex = new Mutex(false, mutexId);

            MutexRights mutexRights = MutexRights.FullControl;
            AccessControlType accessControlType = AccessControlType.Allow;
            WellKnownSidType sid = WellKnownSidType.WorldSid;
            var allowEveryoneRule = new MutexAccessRule(new SecurityIdentifier(sid, null), mutexRights, accessControlType);
            var securitySettings = new MutexSecurity();
            securitySettings.AddAccessRule(allowEveryoneRule);
            _mutex.SetAccessControl(securitySettings);
        }

        public SingleGlobalInstance(int timeOut)
        {
            InitMutex();
            try
            {
                if (timeOut < 0)
                    _hasHandle = _mutex.WaitOne(Timeout.Infinite, false);
                else
                    _hasHandle = _mutex.WaitOne(timeOut, false);

                if (_hasHandle == false)
                    throw new TimeoutException("Timeout waiting for exclusive access on SingleInstance");
            }
            catch (AbandonedMutexException)
            {
                _hasHandle = true;
            }
        }


        public void Dispose()
        {
            if (_mutex != null)
            {
                if (_hasHandle)
                {
                    _mutex.ReleaseMutex();
                }
                _mutex.Close();
            }
        }
    }
}
