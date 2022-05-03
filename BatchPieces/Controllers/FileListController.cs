using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace BatchPieces.Controllers
{
    public class FileListController : Controller
    {
        public ActionResult Index()
        {
            ViewBag.Name = "MA YABIN ";
            return View();
        }
        public JsonResult DeltelWithID(string username, string age)
        {
            return Json(new { FileID = username, State = "Ok" }, JsonRequestBehavior.AllowGet);
        }
        //文件上传方式一  https://www.cnblogs.com/liessay/p/11898214.html?ivk_sa=1024320u
        public ActionResult UpLoad(string fileid,string type)
        {
            Request.Files["File"].SaveAs(Request.MapPath("~/upload/") + Request.Files["File"].FileName);
            int fileCount = Request.Files.Count; //上传数量
            double fileSize = Request.Files["File"].ContentLength; //文件大小（字节）
            string fileName = Request.Files["File"].FileName; //文件名
            string fileType = Request.Files["File"].ContentType;//文件类型
            string fileExt = System.IO.Path.GetExtension(fileName); //文件扩展后缀名
            return Json(new { FileName = Content($"上传数量:{fileCount} 文件名:{fileName} 文件类型：{fileType} 文件格式:{fileExt}") , State = "Ok" });
            //return Content($"上传数量:{fileCount} 文件名:{fileName} 文件类型：{fileType} 文件格式:{fileExt}");
        }
        //文件上传方式2 (GOOGLE \MAC )OK https://blog.huati365.com/0110890bfcaf97ce
        public ActionResult UpLoadOne(HttpPostedFileBase file)
        {
            //判断文件是否存在
            if (file != null)
            {
                //写一个文件保存的路径
                //string imgpath = Server.MapPath(@"\upload\");

                string imgpath = Request.MapPath("~/upload/");
                //判断路径是否存在,不存在则创建
                if (!Directory.Exists(imgpath))
                {
                    Directory.CreateDirectory(imgpath);
                }
                //给文件再起个名
                string ImgName = DateTime.Now.ToString("yyyyMMddHHmmss") + "_" + file.FileName;
                //把文件保存到服务器上
                file.SaveAs(imgpath + ImgName);
                //返回文件的位置
                return Json(new { FileName = imgpath + ImgName, State = "Ok" });
            }
            return Json(new { State = "Error" });
        }

    }
}
