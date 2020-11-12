using System;
using System.Runtime.InteropServices;
using System.Security.Principal;

namespace CutQueue.Lib.tools
{
    public class Impersonation : IDisposable
    {
        private WindowsImpersonationContext _impersonatedUserContext;

        #region FUNCTIONS (P/INVOKE)

        // Declare signatures for Win32 LogonUser and CloseHandle APIs
        [DllImport("advapi32.dll", SetLastError = true)]
        static extern bool LogonUser(
          string principal,
          string authority,
          string password,
          LogonSessionType logonType,
          LogonProvider logonProvider,
          out IntPtr token);

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool CloseHandle(IntPtr handle);

        [DllImport("advapi32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern int DuplicateToken(IntPtr hToken,
            int impersonationLevel,
            ref IntPtr hNewToken);

        [DllImport("advapi32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern bool RevertToSelf();

        #endregion

        #region ENUMS

        enum LogonSessionType : uint
        {
            Interactive = 2,
            Network,
            Batch,
            Service,
            NetworkCleartext = 8,
            NewCredentials
        }

        enum LogonProvider : uint
        {
            Default = 0, // default for platform (use this!)
            WinNT35,     // sends smoke signals to authority
            WinNT40,     // uses NTLM
            WinNT50      // negotiates Kerb or NTLM
        }

        #endregion


        /// <summary>
        /// Class to allow running a segment of code under a given user login context
        /// </summary>
        /// <param name="user">domain\user</param>
        /// <param name="password">user's domain password</param>
        public Impersonation(string domain, string username, string password)
        {
            var token = ValidateParametersAndGetFirstLoginToken(username, domain, password);

            var duplicateToken = IntPtr.Zero;
            try
            {
                if (DuplicateToken(token, 2, ref duplicateToken) == 0)
                {


                    throw new InvalidOperationException("DuplicateToken call to reset permissions for this token failed");
                }

                var identityForLoggedOnUser = new WindowsIdentity(duplicateToken);
                _impersonatedUserContext = identityForLoggedOnUser.Impersonate();
                if (_impersonatedUserContext == null)
                {
                    throw new InvalidOperationException("WindowsIdentity.Impersonate() failed");
                }
            }
            finally
            {
                if (token != IntPtr.Zero)
                    CloseHandle(token);
                if (duplicateToken != IntPtr.Zero)
                    CloseHandle(duplicateToken);
            }
        }

        private static IntPtr ValidateParametersAndGetFirstLoginToken(string domain, string username, string password)
        {


            if (!RevertToSelf())
            {
                throw new InvalidOperationException("RevertToSelf call to remove any prior impersonations failed");
            }


            var result = LogonUser(username, domain, password, LogonSessionType.NewCredentials, LogonProvider.WinNT50, out IntPtr token);
            if (!result)
            {
                var errorCode = Marshal.GetLastWin32Error();
                throw new InvalidOperationException("Logon for user " + username + " failed.");
            }
            return token;
        }

        public void Dispose()
        {
            // Stop impersonation and revert to the process identity
            if (_impersonatedUserContext != null)
            {
                _impersonatedUserContext.Undo();
                _impersonatedUserContext = null;
            }
        }
    }
}
