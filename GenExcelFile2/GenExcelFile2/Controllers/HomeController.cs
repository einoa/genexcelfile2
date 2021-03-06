﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace GenExcelFile2.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
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

        [HttpGet]
        public ActionResult GenExcel()
        {
            return View();
        }

        [HttpPost]
        public ActionResult GenExcel(FormCollection collection)
        {
            System.IO.MemoryStream ms = GenExcelFile2.Models.Excel.TestFile();
            return File(ms, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Testfile.xlsx");
        }
    }
}