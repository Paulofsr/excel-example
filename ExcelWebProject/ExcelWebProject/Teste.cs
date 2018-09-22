using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.ComponentModel.DataAnnotations;

namespace ExcelWebProject
{
    public class Teste
    {
        [Display(Name = "Nome")]
        public string Nome { get; set; }

        [Display(Name = "Valor")]
        [DefaultValue(Value = -1)]
        public decimal A { get; set; }

        [Display(Name = "Outro")]
        public decimal B { get; set; }
    }
}