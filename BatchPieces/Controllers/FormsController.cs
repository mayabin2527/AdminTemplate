using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
//using Microsoft.Office.Interop.Word;

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
            //return Json(new { Path = "PATE", State = "Ok" }, JsonRequestBehavior.AllowGet);
            string path = "Deau";
            //如下MAC无法初始化
            object oMissing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document wordDoc = new Microsoft.Office.Interop.Word.Document();
            wordApp.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone;
            wordDoc.Activate();
            try
            {
                wordDoc = wordApp.Documents.Open(username);
                foreach (Microsoft.Office.Interop.Word.Bookmark item in wordDoc.Bookmarks)
                {
                    if (item.Name == "name")
                    {
                        Microsoft.Office.Interop.Word.Range rang = wordDoc.Range(item.Start, item.End);
                        rang.Text = "？？？？";
                    }
                    if (item.Name == "number")
                    {
                        Microsoft.Office.Interop.Word.Range rang = wordDoc.Range(item.Start, item.End);
                        rang.Text = "012345678901234567";
                    }
                }
                wordApp.Visible = false;
                path = System.IO.Path.GetDirectoryName(username);
                wordDoc.SaveAs(path + "\\替换后模板.doc");
                wordDoc.Close();
                wordApp.Quit();
            }
            catch (Exception e1)
            {
                //MessageBox.Show("请重试\n {0}", e1.Message);
                return Json(new { Message = e1.Message, State = "Error" }, JsonRequestBehavior.AllowGet);
                wordDoc.Close();
                wordApp.Quit();
            }


           //path = "http://WWW.AAA.B/SOU/D.PDF"+Server.MapPath("/SOURCE");
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
