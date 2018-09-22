using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExcelWebProject
{
    public class DefaultValue : Attribute
    {
        public object Value { get; set; }
        public DefaultValue()
        {

        }
    }
}