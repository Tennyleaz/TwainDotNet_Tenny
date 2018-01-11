using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TwainDotNet.TwainNative;

namespace TwainDotNet
{
    public class DefaultDataSourceException : TwainException
    {
        public DefaultDataSourceException(): this(null, null)
        {

        }

        public DefaultDataSourceException(string message) : this(message, null)
        {

        }

        public DefaultDataSourceException(string message, TwainResult returnCode, ConditionCode conditionCode)
            : this(message, null)
        {
            ReturnCode = returnCode;
            ConditionCode = conditionCode;
        }

        public DefaultDataSourceException(string message, Exception innerException) : base(message, innerException)
        {

        }
    }
}
