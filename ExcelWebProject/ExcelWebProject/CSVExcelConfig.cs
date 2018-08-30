using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExcelWebProject
{
    public class CSVExcelConfig
    {
        public CSVExcelConfig()
        {
            GravarEmFormatoTexto = false;
            TrimFields = false;
        }

        public bool GravarEmFormatoTexto { get; set; }

        public bool TrimFields { get; set; }
    }
}