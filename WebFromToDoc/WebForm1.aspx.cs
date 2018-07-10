using Spire.Doc;
using System;
using System.IO;

namespace WebFromToDoc
{
    public partial class WebForm1 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void Unnamed_Click(object sender, EventArgs e)
        {
            //load the template
            createDoc(Server.MapPath("Template/MyTemplate.dotx"));
        }
        private void createDoc(string filePath)
        {   //check the template is available
            if (File.Exists(filePath))
            {
                //load the doc
                Document document = new Document(filePath);

                //find and replace the content
                document.Replace("<Title>", txtTitle.Text, true, true);
                document.Replace("<name>", txtName.Text, true, true);

                //find and replace the header & footer
                document.Replace("<CompanyName>", "冰冰无限公司", true, true);
                document.Replace("<Date>", DateTime.Now.ToString("dd-MMM-yyyy HH:mm:ss").ToString(), true, true);

                //save the doc in temporarily folder
                string tempPath = Server.MapPath("Temporarily/" + DateTime.Now.ToString("ddmmyyyyhhmmssffffff") + ".docx");
                document.SaveToFile(tempPath, FileFormat.Docx);
                //let the user download and delete it the doc
                Response.ContentType = "Application/msword";
                Response.AddHeader("Content-Disposition", "attachment;filename= Example.docx");
                Response.TransmitFile(Path.Combine(tempPath.ToString()));
                Response.Flush();
                File.Delete(tempPath);
                Response.End();
            }

        }
    }
}