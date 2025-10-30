using System;
using System.Runtime.InteropServices;
using System.Security;

namespace MontfoortIT.Library.Extensions
{
    public static class SecureStringExtensions
    {
        public static string ToUnsecureString(this SecureString secure)
        {
            IntPtr valuePtr = IntPtr.Zero;
            try
            {
                valuePtr = Marshal.SecureStringToGlobalAllocUnicode(secure);
                return Marshal.PtrToStringUni(valuePtr);
            }
            finally
            {
                Marshal.ZeroFreeGlobalAllocUnicode(valuePtr);
            }
        }

    }
}
