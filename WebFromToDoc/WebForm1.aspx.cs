using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;

namespace WebFromToDoc
{
    public partial class WebForm1 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                DropDownList1.DataSource = get();
                DropDownList1.DataBind();
                GridView1.DataSource = getData(Server.MapPath("Excel/" + DropDownList1.SelectedValue));
                GridView1.DataBind();
            }
        }


        protected void Unnamed_Click(object sender, EventArgs e)
        {
            //load the template
            createDoc(Server.MapPath("Template/MyTemplate.dotx"));
            //

        }
        private List<string> get()
        {   //get all excel file 
            List<string> excelPath = new List<string>();
            foreach (string file in Directory.GetFiles(Server.MapPath("~/Excel")))
            {
                excelPath.Add(file.Substring(file.LastIndexOf("\\") + 1));
            }
            return excelPath;
        }
        private DataTable getData(string filePath)
        {
            //read the excel file
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(filePath);
            //read the excel first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            //return the data as DataTable
            return sheet.ExportDataTable();
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

                Section section = document.Sections[0];
                TextSelection selection = document.FindString("<MyTable>", true, true);
                TextRange range = selection.GetAsOneRange();
                Paragraph paragraph = range.OwnerParagraph;
                Body body = paragraph.OwnerTextBody;
                int index = body.ChildObjects.IndexOf(paragraph);

                Table table = section.AddTable(true);

                DataTable dt = getData(Server.MapPath("Excel/" + DropDownList1.SelectedValue));


                string[] header = dt.Columns.Cast<DataColumn>()
                                 .Select(x => x.ColumnName)
                                 .ToArray();
                //Add Cells
                table.ResetCells(dt.Rows.Count + 1, header.Length);

                //Header Row
                TableRow FRow = table.Rows[0];
                FRow.IsHeader = true;
                //Row Height
                FRow.Height = 23;
                //Header Format
                FRow.RowFormat.BackColor = Color.AliceBlue;
                for (int x = 0; x < header.Length; x++)
                {
                    //Cell Alignment
                    Paragraph p = FRow.Cells[x].AddParagraph();
                    FRow.Cells[x].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                    p.Format.HorizontalAlignment = HorizontalAlignment.Center;
                    //Data Format
                    TextRange tr = p.AppendText(header[x]);
                    tr.CharacterFormat.FontName = "Calibri";
                    tr.CharacterFormat.FontSize = 14;
                    tr.CharacterFormat.TextColor = Color.Teal;
                    tr.CharacterFormat.Bold = true;
                }

                for (int x = 0; x < dt.Rows.Count; x++)
                {
                    TableRow DataRow = table.Rows[x + 1];
                    DataRow.Height = 20;
                    for (int y = 0; y < dt.Columns.Count; y++)
                    {
                        //Cell Alignment
                        DataRow.Cells[y].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                        //Fill Data in Rows
                        Paragraph p2 = DataRow.Cells[y].AddParagraph();
                        TextRange tr2 = p2.AppendText(dt.Rows[x][y].ToString());
                        //Format Cells
                        p2.Format.HorizontalAlignment = HorizontalAlignment.Center;
                        tr2.CharacterFormat.FontName = "Calibri";
                        tr2.CharacterFormat.FontSize = 12;
                        tr2.CharacterFormat.TextColor = Color.Brown;
                    }
                }

                body.ChildObjects.Remove(paragraph);
                body.ChildObjects.Insert(index, table);


                //save the doc in temporarily folder
                string tempPath = Server.MapPath("Temporarily/" + DateTime.Now.ToString("ddmmyyyyhhmmssffffff") + ".docx");
                document.SaveToFile(tempPath, Spire.Doc.FileFormat.Docx);
                //let the user download and delete it the doc
                Response.ContentType = "Application/msword";
                Response.AddHeader("Content-Disposition", "attachment;filename= Example.docx");
                Response.TransmitFile(Path.Combine(tempPath.ToString()));
                Response.Flush();
                File.Delete(tempPath);
                Response.End();
            }

        }

        protected void DropDownList1_SelectedIndexChanged(object sender, EventArgs e)
        {
            GridView1.DataSource = getData(Server.MapPath("Excel/" + DropDownList1.SelectedValue));
            GridView1.DataBind();
        }
    }
}