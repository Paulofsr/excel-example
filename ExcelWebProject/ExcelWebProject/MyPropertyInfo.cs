using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Web;

namespace ExcelWebProject
{
    public class MyPropertyInfo
    {
        public PropertyInfo Property { get; set; }

        public object DefaultValue { get; set; }
    }
}