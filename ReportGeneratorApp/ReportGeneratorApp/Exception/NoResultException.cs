using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ReportGeneratorApp.CustomizeException
{
    public class NoResultException : Exception
    {
        public NoResultException(String message)
            : base(message)
        {
        }
    }
}