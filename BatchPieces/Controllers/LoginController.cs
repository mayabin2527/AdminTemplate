using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace BatchPieces.Controllers
{
    public class LoginController : Controller
    {
        public ActionResult Index()
        {
            //ADFS 登录参考：https://www.cnblogs.com/hudun/p/5919630.html
            return View ();
        }

        public JsonResult Submit(string username,string age)
        {
            Session["User"] = username;
            Session["Admin"] = "yes";
            var res = new JsonResult();

            //var value = "actionValue"; 

            //db.ContextOptions.ProxyCreationEnabled = false; 

            /*var list = (from a in db.Articles

                        select new

                        {

                            name = a.ArtTitle,

                            yy = a.ArtPublishTime

                        }).Take(5);

            //记得这里要select new 否则会报错：序列化类型 System.Data.Entity.DynamicProxies XXXXX 的对象时检测到循环引用。 

            //不select new 也行的加上这句 //db.ContextOptions.ProxyCreationEnabled = false; 

            res.Data = list;//返回列表 

    */

            //var name = "小华";

            //var age = "12";

            //var name1 = "小华";

            //var age1 = "12";

            //res.Data = new object[] { new { name, age }, new { name1, age1 } };//返回一个自定义的object数组 



            var person = new { Name = "小明", Age = 22, Sex = "男" ,State="Ok"};

            res.Data = person;//返回单个对象； 



            //res.Data = "这是个字符串";//返回一个字符串,意义不大； 



            res.JsonRequestBehavior = JsonRequestBehavior.AllowGet;//允许使用GET方式获取，否则用GET获取是会报错。 

            return res;
        }
    }
}
