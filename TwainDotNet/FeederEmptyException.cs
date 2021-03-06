﻿using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.Serialization;

namespace TwainDotNet
{
    public class FeederEmptyException : TwainException
    {
        public FeederEmptyException()
            : this(null, null)
        {
        }

        public FeederEmptyException(string message)
            : this(message, null)
        {
        }

        protected FeederEmptyException(SerializationInfo info, StreamingContext context) :
            base(info, context)
        {
        }

        public FeederEmptyException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }

    public class DeviceOpenExcetion : TwainException
    {
        public DeviceOpenExcetion() : this(null, null)
        {

        }

        public DeviceOpenExcetion(string message) : this(message, null)
        {

        }

        public DeviceOpenExcetion(string message, TwainNative.TwainResult result) : this(message, null)
        {
            base.ReturnCode = result;
        }

        public DeviceOpenExcetion(string message, Exception innerException) : base(message, innerException)
        {

        }
    }
}
