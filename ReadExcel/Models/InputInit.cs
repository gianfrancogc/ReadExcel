using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ReadExcel.Models
{
    public class InputInit
    {
        public string Sku { get; set; }
        public string Sucursal { get; set; }
    }

    public class ExportPim
    {
        public string SkuPeru { get; set; }
        public string Tienda { get; set; }
    }
}