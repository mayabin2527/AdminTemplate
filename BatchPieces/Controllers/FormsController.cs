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
            for (int i = 0; i < 10; i++)
            {
                System.Threading.Thread.Sleep(1000);
            }
            return Json(new { Path = "ttt", State = "Ok" }, JsonRequestBehavior.AllowGet);
            
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
for (int i = 1; i <= 50; i++)
{
    string ID = i.ToString();
    string First = "First" + i.ToString();
    string LastName = "LastName" + i.ToString();
    string Username = "Username"+ i.ToString();

    string sa = " <a id = '{0}' href = '#' class='fa fa-edit edit1' data-toggle='modal' data-target='.bs-example-modal-lg'>编辑</a>  <a id = '{1}' href='#' class='fa fa-times delete1' data-toggle='modal' data-target='.bs-example-modal-sm'>删除</a>";
    string Salary = string.Format(sa, "edit" + i.ToString(), "deltel" + i.ToString());
    list.Add(new Employee() { ID = ID, First = First, LastName = LastName, Username = Salary });
}
return Json(new { data = " <tr>  <td> 5</td> <td>Larry5</td> <td>the Bird5</td>   <td>twitter5</td> </tr >", State = "Ok" }, JsonRequestBehavior.AllowGet);
}
public class Employee
{

public string ID { get; set; }
public string First { get; set; }
public string LastName { get; set; }
public string Username { get; set; }
}
}
}
