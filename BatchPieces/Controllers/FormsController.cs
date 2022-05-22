using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using BatchPieces.Tools;
using Spire.Doc;
using Spire.Doc.Documents;

//using Microsoft.Office.Interop.Word;

namespace BatchPieces.Controllers
{
    public class FormsController : BaseController
    {
        public ActionResult Index()
        {
            return View ();
        }
        public JsonResult GetCompany()
        {
            /* DAPPER   https://www.jianshu.com/p/cb304bb8cc3e
            using (var connection = new SqlConnection(...)) 
{ 
  var authors = connection.Query<Author>(
    "Select * From Author").ToList(); 
}
             */

            List<Employee> list = new List<Employee>();
            for (int i = 1; i <= 10; i++)
            {
                string ID = "10000"+i.ToString();
                string First = "First" + i.ToString();
             
                list.Add(new Employee() { ID = ID, First = First, LastName = "LastName", Username = "Salary" });
            }
            return Json(new { list = list, State = "Ok" });


        }
        public JsonResult GetCategory(string company)
        {
            List<Employee> list = new List<Employee>();
            for (int i = 1; i <= 20; i++)
            {
                string ID = "10000" + i.ToString();
                string First = "类别" + i.ToString();

                list.Add(new Employee() { ID = ID, First = First, LastName = "", Username = "" });
            }
            //查询List 
            // var newlist = list.Where(a => a.ID == "200001" && a.LastName == "类别1").ToList();
            var newlist = list.Where(a => a.ID == company).ToList();
            return Json(new { list = newlist, State = "Ok" });
        }

        public JsonResult AddDocModel(string username, string age)
        {
            for (int i = 0; i < 5; i++)
            {
                System.Threading.Thread.Sleep(1000);
            }
            
            //使用此插件替换书签，可使用免费版，教程：https://blog.csdn.net/ssw_jack/article/details/80061791
            //官网：https://www.e-iceblue.com/Introduce/word-for-net-introduce.html#.YohSjJNByrc


            //加载模板文档
            Document doc = new Document();
            doc.LoadFromFile(@"C:\Users\Administrator\Desktop\bookmark_template.docx");
            
            

            //初始化Bookmark对象
            DocumentTools bookmark = new DocumentTools(doc);

            //用文本替换书签bookmark_text的内容
            string text = "XXX科技股份有限公司成立于2010年12月，是一家致力于高新技术产品研发、生产、销售的高科技股份制企业，"
                + "公司坚持以技术创新为核心，以知识产权为基础，以人才战略为支撑，经过多年的砺练与发展，公司已逐步成以创新为引导的，"
                + "产品具有竞争力，人才素质优良的新兴科技企业。";
            bookmark.ReplaceContent("bookmark_text", text, true);
            /*
            //用图片替换书签bookmark_picture的内容
            string picPath = @"C:\Users\Administrator\Desktop\company_logo.jpg";
            bookmark.ReplaceContent("bookmark_picture", picPath, 80f, 80f, TextWrappingStyle.TopAndBottom, ShapeHorizontalAlignment.Center);

            //创建模拟数据
            DataTable dt = new DataTable();
            dt.Columns.Add("employee_id", typeof(string));
            dt.Columns.Add("name", typeof(string));
            dt.Columns.Add("age", typeof(string));
            dt.Columns.Add("sex", typeof(string));
            dt.Columns.Add("title", typeof(string));
            dt.Rows.Add(new string[] { "工号", "姓名", "年龄", "性别", "职位" });
            dt.Rows.Add(new string[] { "1023", "Nancy", "28", "女", "Java程序员" });
            dt.Rows.Add(new string[] { "1024", "James", "34", "男", ".NET程序员" });
            dt.Rows.Add(new string[] { "1025", "Kobe", "38", "男", "系统管理员" });

            //创建表格，并填充数据
            Table table = bookmark.CreateTable(dt.Rows.Count, dt.Columns.Count, 100f, RowAlignment.Left, dt);

            //用表格替换书签bookmark_table的内容
            bookmark.ReplaceContent("bookmark_table", table);
            */
            //生成Word文件
            string path = Server.MapPath("upload/") + DateTime.Now.ToString("yyMMddhhmmss") + "output.docx";
            doc.SaveToFile(path, FileFormat.Docx2013);

            return Json(new { Path = path, State = "Ok" }, JsonRequestBehavior.AllowGet);

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
            /*
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

// 根据Word书签赋值 
SetBookMarksValue(null, "T1", "ItemName");


//path = "http://WWW.AAA.B/SOU/D.PDF"+Server.MapPath("/SOURCE");
return Json(new { Path= path, State = "Ok" }, JsonRequestBehavior.AllowGet);*/

        }
        /// <summary>
        /// 根据Word书签赋值 
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="name"></param>
        /// <param name="value"></param>
        private void SetBookMarksValue(Microsoft.Office.Interop.Word.DocumentClass doc, object name, string value)
{
doc.Bookmarks.get_Item(ref name).Range.Text = value;
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

//后端拼接数据返回 前段datatables 使用 ok 建议使用
public JsonResult myData(string username, string age)
{
List<Employee> list = new List<Employee>();
List<TableList> list1 = new List<TableList>();
            for (int i = 1; i <= 10; i++)
{
    string ID = i.ToString();
    string First = "First" + i.ToString();
    string LastName = "LastName" + i.ToString();
    string Username = "Username"+ i.ToString();

    string sa = " <a id = '{0}' href = '#' class='fa fa-edit edit1' data-toggle='modal' data-target='.bs-example-modal-lg'>编辑</a>  <a id = '{1}' href='#' class='fa fa-times delete1' data-toggle='modal' data-target='.bs-example-modal-sm'>删除</a>";
    string Salary = string.Format(sa, "edit" + i.ToString(), "deltel" + i.ToString());
    list.Add(new Employee() { ID = ID, First = First, LastName = LastName, Username = Salary });

     list1.Add(new TableList() { 编号 = ID, 姓什么 = First, 名字 = LastName, 用户名称 = Salary });

}
            return Json(new { data = " <tr>  <td> 5</td> <td>Larry5</td> <td>the Bird5</td>   <td>twitter5</td> </tr >",List= list1, State = "Ok" }, JsonRequestBehavior.AllowGet);
}
public class Employee
{

public string ID { get; set; }
public string First { get; set; }
public string LastName { get; set; }
public string Username { get; set; }
}

        public class TableList
        {

            public string 编号 { get; set; }
            public string 姓什么 { get; set; }
            public string 名字 { get; set; }
            public string 用户名称 { get; set; }
        }
    }
}
