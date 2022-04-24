using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace BatchPieces.Controllers
{
    public class FileListController : Controller
    {
        public ActionResult Index()
        {
            ViewBag.Name = "MA YABIN";
            return View ();
        }
    }
}
