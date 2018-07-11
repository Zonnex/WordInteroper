using System;
using System.Runtime.CompilerServices;
using System.Reflection;

namespace WordInteroper
{
    public static class Contracts
    {
        public static void Require(bool precondition, string errorMessage = "", [CallerMemberName]string method = null)
        {
            if(!precondition)
            {
                if(method != null)
                    throw new ContractException($"Contract fault: {errorMessage} in {method}");
                else
                    throw new ContractException(errorMessage);
            }
        }

        public static void Require<TException>(bool precondition, string errorMessage = "", [CallerMemberName]string method = null)
            where TException : Exception
        {
            if(!precondition)
            {
                TException ex = (TException)Activator.CreateInstance(typeof(TException));
                typeof(TException)
                    .GetField("_message", BindingFlags.NonPublic | BindingFlags.Instance)
                    .SetValue(ex, errorMessage);
                throw ex;
            }
        }
    }
}
