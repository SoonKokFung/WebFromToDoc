
using Spire.Doc;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace WebFromToDoc
{
    public partial class WebForm1 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void Unnamed_Click(object sender, EventArgs e)
        {
            createDoc(Server.MapPath("Template/Example.docx"));
        }
        private void createDoc(string filePath)
        {
            if (File.Exists(filePath))
            {
                Document document = new Document();
                document.LoadFromFile(filePath);

                document.Replace("<Title>", txtTitle.Text, true, true);
                document.Replace("<name>", txtName.Text, true, true);

                document.Replace("<CompanyName>", "冰冰无限公司", true, true);
                document.Replace("<Date>", DateTime.Now.ToString("dd-MMM-yyyy HH:mm:ss").ToString(), true, true);

                string tempPath = Server.MapPath("Temporarily/" + DateTime.Now.ToString("ddmmyyyyhhmmssffffff") + ".docx");
                document.SaveToFile(tempPath, FileFormat.Docx);

                //let the user download the doc
                Response.ContentType = "Application/msword";
                Response.AddHeader("Content-Disposition", "attachment;filename=" + tempPath);
                Response.TransmitFile(Path.Combine(tempPath.ToString()));
                Response.Flush();
                File.Delete(tempPath);
                Response.End();
            }

        }
    }
}