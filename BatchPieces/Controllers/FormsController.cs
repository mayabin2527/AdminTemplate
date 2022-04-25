using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Microsoft.Office.Interop.Word;

namespace BatchPieces.Controllers
{
    public class FormsController : BaseController
    {
        public ActionResult Index()
        {
            return View ();
        }

        public JsonResult AddDocModel(string username, string age)
        {

            /*无效代码
             * Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            string version = app.Version;
            Document doc = app.Documents.Add("word");
            doc.ActiveWindow.Visible = true;
            foreach (Bookmark bk in doc.Bookmarks)
            {
                bk.Range.Text = GetStrByBookmarkName(bk.Name);
            }
            //保存文档并打开
            object IsSave = true;
            object missing = true;
         
            doc.Close(ref IsSave, ref missing, ref missing);*/

            // return new JsonResult() {  Data=new { State = "Ok" } };
            string path = "http://WWW.AAA.B/SOU/D.PDF"+Server.MapPath("/SOURCE");
            return Json(new { Path= path, State = "Ok" }, JsonRequestBehavior.AllowGet);
        }


               
        public string GetStrByBookmarkName(String name)
        {
            string ret = "";
            switch (name)
            {
                case "name":
                    ret = "Ma Yabin";
                    break;
                default:
                    break;
            }
            return ret;
        }

        
    }
}
