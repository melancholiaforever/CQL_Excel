using Newtonsoft.Json;
using System.Data;
using System.Text;
using System.Net;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;
using Newtonsoft.Json.Linq;
using System.Configuration;

namespace ExcelAddIn1
{
    class BLL
    {

        public static DataTable GetDataTable(string dataset,string table,string limit, string offset)
        {
            ////创建Httphelper对象
            //HttpHelper http = new HttpHelper();
            //ServicePointManager.ServerCertificateValidationCallback += (object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) =>
            //{
            //    return true;
            //};
            ////ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls;
            ////创建Httphelper参数对象
            //HttpItem item = new HttpItem()
            //{
            //    Method = "POST",
            //    URL = "https://e.morenodes.com:11108/v1/query",//URL     必需项 
            //    CerPath = "C:\\Cert\\read.data.thunderdb.io.pfx", //
            //    Pwd = "covenantsql",
            //    Postdata = "database = 053d0bb19637ffc7b4a94e3c79cc71b67a768813b09e4b67f1d6159902754a8b & query = "+ strSql + "limit 10000",
            //    ContentType = "text"
            //};
            ////请求的返回值对象
            //HttpResult result = http.GetHtml(item);
            //string html = result.Html;
            //DataTable dataTable1 = MyData.Class_mysql_conn.Get_DataTable(strSql, MyData.Class_mysql_conn.ConnStr,tableName);
            //var keystorefile = @"" + path + "\\read.data.thunderdb.io.pfx";
            string strSql = "select * from quandl_" + dataset + " where quandlcode = '"+ table + "' limit " + limit + " offset " + offset;
            string certpath = ConfigurationManager.AppSettings["certpath"];
            string keypasswd = ConfigurationManager.AppSettings["keypassword"];
            string url = ConfigurationManager.AppSettings["url"];
            string database = ConfigurationManager.AppSettings["database"];
            var data = Encoding.UTF8.GetBytes("database=" + database + "&query=" + strSql);
            var request = (HttpWebRequest)WebRequest.Create(url);
            //X509Certificate2 支持读取.p12格式的证书
            var cer = new X509Certificate2(certpath, keypasswd);
            request.ClientCertificates.Add(cer);
            ServicePointManager.ServerCertificateValidationCallback += (object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) =>
            {
                return true;
            };
            //ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
            request.Method = "POST";
            request.ContentType = "application/x-www-form-urlencoded";
            request.ContentLength = data.Length;
            using (var stream = request.GetRequestStream())
            {
                stream.Write(data, 0, data.Length);
            }
            var response = (HttpWebResponse)request.GetResponse();
            var context = new StreamReader(response.GetResponseStream(), Encoding.UTF8).ReadToEnd();
            DataTable dt = GetObject(context);
            return dt;
        }

        public static DataTable ShowTables()
        {
            ////创建Httphelper对象
            //HttpHelper http = new HttpHelper();
            //ServicePointManager.ServerCertificateValidationCallback += (object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) =>
            //{
            //    return true;
            //};
            ////ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls;
            ////创建Httphelper参数对象
            //HttpItem item = new HttpItem()
            //{
            //    Method = "POST",
            //    URL = "https://e.morenodes.com:11108/v1/query",//URL     必需项 
            //    CerPath = "C:\\Cert\\read.data.thunderdb.io.pfx", //
            //    Pwd = "covenantsql",
            //    Postdata = "database = 053d0bb19637ffc7b4a94e3c79cc71b67a768813b09e4b67f1d6159902754a8b & query = "+ strSql + "limit 10000",
            //    ContentType = "text"
            //};
            ////请求的返回值对象
            //HttpResult result = http.GetHtml(item);
            //string html = result.Html;
            //DataTable dataTable1 = MyData.Class_mysql_conn.Get_DataTable(strSql, MyData.Class_mysql_conn.ConnStr,tableName);
            //var keystorefile = @""+path+"\\read.data.thunderdb.io.pfx";
            string certpath = ConfigurationManager.AppSettings["certpath"];
            string keypasswd = ConfigurationManager.AppSettings["keypassword"];
            string url = ConfigurationManager.AppSettings["url"];
            string database = ConfigurationManager.AppSettings["database"];
            var data = Encoding.UTF8.GetBytes("database=" + database + "&query=show tables");
            var request = (HttpWebRequest)WebRequest.Create(url);
            //X509Certificate2 支持读取.p12格式的证书
            var cer = new X509Certificate2(certpath, keypasswd);
            request.ClientCertificates.Add(cer);
            ServicePointManager.ServerCertificateValidationCallback += (object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) =>
            {
                return true;
            };
            //ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
            request.Method = "POST";
            request.ContentType = "application/x-www-form-urlencoded";
            request.ContentLength = data.Length;
            using (var stream = request.GetRequestStream())
            {
                stream.Write(data, 0, data.Length);
            }
            var response = (HttpWebResponse)request.GetResponse();
            var context = new StreamReader(response.GetResponseStream(), Encoding.UTF8).ReadToEnd();
            DataTable dt = GetObject(context);
            return dt;
        }

        public static DataTable GetDataTableflex(string sql)
        {
            ////创建Httphelper对象
            //HttpHelper http = new HttpHelper();
            //ServicePointManager.ServerCertificateValidationCallback += (object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) =>
            //{
            //    return true;
            //};
            ////ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls;
            ////创建Httphelper参数对象
            //HttpItem item = new HttpItem()
            //{
            //    Method = "POST",
            //    URL = "https://e.morenodes.com:11108/v1/query",//URL     必需项 
            //    CerPath = "C:\\Cert\\read.data.thunderdb.io.pfx", //
            //    Pwd = "covenantsql",
            //    Postdata = "database = 053d0bb19637ffc7b4a94e3c79cc71b67a768813b09e4b67f1d6159902754a8b & query = "+ strSql + "limit 10000",
            //    ContentType = "text"
            //};
            ////请求的返回值对象
            //HttpResult result = http.GetHtml(item);
            //string html = result.Html;
            //DataTable dataTable1 = MyData.Class_mysql_conn.Get_DataTable(strSql, MyData.Class_mysql_conn.ConnStr,tableName);
            //var keystorefile = @"" + path + "\\read.data.thunderdb.io.pfx";
            string certpath = ConfigurationManager.AppSettings["certpath"];
            string keypasswd = ConfigurationManager.AppSettings["keypassword"];
            string url = ConfigurationManager.AppSettings["url"];
            string database = ConfigurationManager.AppSettings["database"];
            var data = Encoding.UTF8.GetBytes("database=" + database + "&query=" + sql);
            var request = (HttpWebRequest)WebRequest.Create(url);
            //X509Certificate2 支持读取.p12格式的证书
            var cer = new X509Certificate2(certpath, keypasswd);
            request.ClientCertificates.Add(cer);
            ServicePointManager.ServerCertificateValidationCallback += (object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) =>
            {
                return true;
            };
            //ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
            request.Method = "POST";
            request.ContentType = "application/x-www-form-urlencoded";
            request.ContentLength = data.Length;
            using (var stream = request.GetRequestStream())
            {
                stream.Write(data, 0, data.Length);
            }
            var response = (HttpWebResponse)request.GetResponse();
            var context = new StreamReader(response.GetResponseStream(), Encoding.UTF8).ReadToEnd();
            DataTable dt = GetObject(context);
            return dt;
        }
        //}
        //
        /// <summary>
        /// Json 字符串 转换为 DataTable数据集合
        /// </summary>
        /// <param name="json"></param>
        /// <returns></returns>
        /// <summary>
        /// Json 字符串 转换为 DataTable数据集合
        /// </summary>
        /// <param name="json"></param>
        /// <returns></returns>


        //public static void getObject(string json)
        //{
        //    JObject obj = JsonConvert.DeserializeObject<JObject>(json);
        //    if ((bool)obj["success"])
        //    {
        //        var data = obj["data"] as JObject;
        //        DataTable dt = new DataTable();
        //        foreach (var col in data["columns"])
        //        {
        //            dt.Columns.Add(col.ToString());
        //        }
        //        JArray rows = data["rows"] as JArray;
        //        foreach (JArray row in rows)
        //        {
        //            DataRow dr = dt.NewRow();
        //            for (var i = 0; i < row.Count; i++)
        //            {
        //                dr[i] = row[i].ToString();
        //            }
        //            dt.Rows.Add(dr);
        //        }
        //    }
        //}
        //public static void GetObject(string json)
        //{
        //    JObject obj = JsonConvert.DeserializeObject<JObject>(json);
        //    if ((bool)obj["success"])
        //    {
        //        var data = obj["data"] as JObject;
        //        DataTable dt = new DataTable();
        //        foreach (var col in data["columns"])
        //        {
        //            dt.Columns.Add(col.ToString());
        //        }
        //        JArray rows = data["rows"] as JArray;
        //        foreach (JArray row in rows)
        //        {
        //            DataRow dr = dt.NewRow();
        //            for (var i = 0; i < row.Count; i++)
        //            {
        //                dr[i] = row[i].ToString();
        //            }
        //            dt.Rows.Add(dr);
        //        }
        //    }
        //}



        public static DataTable GetObject(string json)
        {
            JObject obj = JsonConvert.DeserializeObject<JObject>(json);
            if ((bool)obj["success"])
            {
                var data = obj["data"] as JObject;
                DataTable dt = new DataTable();
                foreach (var col in data["columns"])
                {
                    dt.Columns.Add(col.ToString());
                }
                JArray rows = data["rows"] as JArray;
                foreach (JArray row in rows)
                {
                    DataRow dr = dt.NewRow();
                    for (var i = 0; i < row.Count; i++)
                    {
                        dr[i] = row[i].ToString();
                    }
                    dt.Rows.Add(dr);
                }
                return dt;
            }
            return new DataTable();
        }

    }
}