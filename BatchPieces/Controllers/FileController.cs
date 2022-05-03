using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Script.Serialization;

namespace BatchPieces.Controllers
{
    public class FileController : Controller
    {
        public ActionResult Index()
        {
            //file list 分页版

            string path1 = Request.MapPath("~/upload/JSONFile.json");
            string str = "";// File.ReadAllText(path);
            using (StreamReader sr = new StreamReader(path1))
            {
                string res = sr.ReadToEnd();
                str = res;
                Console.WriteLine(res);
            }
            JavaScriptSerializer Serializer = new JavaScriptSerializer();
          //  List<data> shoppingList = Serializer.Deserialize<List<data>>(str);
            Entity result = FromJSON<Entity>(str);
            ViewBag.FList = result.sites;
            ViewBag.List = getTableHtml(result.sites);
            return View ();
        }

        //后端拼接数据返回 前段datatables 使用 ok 建议使用
        public JsonResult myData(string username, string age)
        {
            //https://datatables.net/manual/ajax#Loading-data
            List<Employee> list= new List<Employee>();
            Employee i = new Employee() { Name= "Name", Position= "Position", Office= "Office", Age="22", Startdate= "Startdate" , Salary= "Salary" };
            Employee i1 = new Employee() { Name = "Name23", Position = "Position", Office = "Office", Age = "23", Startdate = "Startdate", Salary = "<a href = '#' > View </a> <a id = 'edit88888' href = '#' class='fa fa-edit edit1' data-toggle='modal' data-target='.bs-example-modal-lg'>编辑</a>  <a id = 'delete000001' href='#' class='fa fa-times delete1' data-toggle='modal' data-target='.bs-example-modal-sm'>删除</a>" };
            list.Add(i);
            list.Add(i1);

            Employee i3 = new Employee() { Name = "Name3", Position = "Position", Office = "Office", Age = "22", Startdate = "Startdate", Salary = "Salary" };
            Employee i4 = new Employee() { Name = "Name4", Position = "Position", Office = "Office", Age = "22", Startdate = "Startdate", Salary = "Salary" };
            list.Add(i3);
            list.Add(i4);

            return Json(new { data = list, State = "Ok" }, JsonRequestBehavior.AllowGet);
        }

        public partial class Employee
        {
           
            public string Name { get; set; }
            public string Position { get; set; }
            public string Office { get; set; }
            public string Age { get; set; }
            public string Startdate { get; set; }
            public string Salary { get; set; }
        }

        public string getTableHtml(List<data> list) {
            string trItem = @" <tr>
                                                            <td>Tiger Nixon</td>
                                                            <td>System Architect</td>
                                                            <td>Edinburgh</td>
                                                            <td>61</td>
                                                            <td>2011/04/25</td>
                                                            <!-- <td>$320,800</td>-->
                                                            <td class=' last'>
                                                                < a href = '#' > View </ a >



                                                                 < a id = 'edit123' href = '#' class='fa fa-edit edit1' data-toggle='modal' data-target='.bs-example-modal-lg'>编辑</a>
                                                                <a id = 'delete123' href='#' class='fa fa-times delete1' data-toggle='modal' data-target='.bs-example-modal-sm'>删除</a>

                                                            </td>
                                                        </tr>";
            foreach (var item in list)
            {

            }
            return trItem;
        }

        /// <summary>
        /// Json字符串转内存对象
        /// </summary>
        /// <param name="jsonString"></param>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static T FromJSON<T>(string jsonString)
        {
            JavaScriptSerializer json = new JavaScriptSerializer();
            return json.Deserialize<T>(jsonString);
        }
        [Serializable]
        public partial class Entity
        {
           
            public List<data> sites { get; set; }
        }
        [Serializable]
        public partial class data
        {
            public string name { get; set; }
            public string url { get; set; }
        }
    }
}
