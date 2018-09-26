using System;
using System.Collections.Generic;
using System.Text;

namespace FastStream.Log
{
   public class FlashLogMessage
    {
        public string Message { get; set; }
        public FlashLogLevel Level { get; set; }
        public Exception Exception { get; set; }
    }
}
