using Microsoft.Office.Interop.Excel;
using ReadExcel.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ReadExcel.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {

            return View();
        }

        [HttpPost]
        public ActionResult Index(HttpPostedFileBase excelfile, HttpPostedFileBase excelFile2)
        {

            var listaInput = GetInput(excelfile);
            var listaPim = GetExportPim(excelFile2);

            foreach (var item in listaPim)
            {
                var firsList = listaInput.Where(x => x.Sku == item.SkuPeru).ToList();
                var listTienda = firsList.Select(x => x.Sucursal).ToList();
                var singlePim = listaPim.SingleOrDefault(x => x.SkuPeru == item.SkuPeru).Tienda.Split(';').ToList();
                foreach (var tienda in listTienda)
                {

                }
            }

            var exportP = new List<ExportPim>();
            foreach (var item in listaInput)
            {
                var pim = listaPim.SingleOrDefault(x => x.SkuPeru == item.Sku);
                if (pim!= null)
                {
                    var sucursales = pim.Tienda.Split(';').ToList();
                    sucursales.Remove(item.Sucursal);
                    var newTiendas = String.Join(";", sucursales);
                    exportP.Add(new ExportPim
                    {
                        SkuPeru = pim.SkuPeru,
                        Tienda = newTiendas
                    });
                }
            }

            string FileName = "ExportPimFinal.xlsx";
            string TempFilename = Path.Combine(Path.GetTempPath(), FileName);
            CreateExcelFile.CreateExcelDocument(exportP.GroupBy(x => x.SkuPeru).ToList(), TempFilename);

            byte[] rawData = System.IO.File.ReadAllBytes(TempFilename);

            MemoryStream ms = new MemoryStream(rawData);

            FileStreamResult fr = new FileStreamResult(ms, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            {
                FileDownloadName = FileName
            };

            System.IO.File.Delete(TempFilename);
            return fr;

        }


        public  List<InputInit> GetInput(HttpPostedFileBase excelfile)
        {

            string path = Server.MapPath("~/Content/" + excelfile.FileName);

            if (System.IO.File.Exists(path))
            {
                System.IO.File.Delete(path);
            }

            excelfile.SaveAs(path);
            Application app = new Application();
            Workbook workbook = app.Workbooks.Open(path);
            Worksheet worksheet = workbook.ActiveSheet;
            Range range = worksheet.UsedRange;
            var lista = new List<InputInit>();
            for (int i = 2; i < range.Rows.Count; i++)
            {
                var input = new InputInit
                {
                    Sku = ((Range)range.Cells[i, 1]).Text,
                    Sucursal = ((Range)range.Cells[i, 2]).Text,
                };
                lista.Add(input);
            }

            workbook.Close();
            return lista;
        }


        public List<ExportPim> GetExportPim(HttpPostedFileBase excelFile2)
        {
            string path = Server.MapPath("~/Content/" + excelFile2.FileName);
            if (System.IO.File.Exists(path))
            {
                System.IO.File.Delete(path);
            }
            excelFile2.SaveAs(path);
            Application app = new Application();
            Workbook workbook = app.Workbooks.Open(path);
            Worksheet worksheet = workbook.ActiveSheet;
            Range range = worksheet.UsedRange;
            var lista = new List<ExportPim>();
            for (int i = 2; i < range.Rows.Count; i++)
            {
                var input = new ExportPim
                {
                    SkuPeru = ((Range)range.Cells[i, 1]).Text,
                    Tienda = ((Range)range.Cells[i, 2]).Text,
                };
                lista.Add(input);
            }

            workbook.Close();
            return lista;
        }


        public ActionResult Success()
        {
            ViewBag.Message = "TODO OK.";

            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}