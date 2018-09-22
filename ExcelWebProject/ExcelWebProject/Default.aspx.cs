using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace ExcelWebProject
{
    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            ExcelIO<Teste> excel = new ExcelWebProject.ExcelIO<ExcelWebProject.Teste>();
            List<ExcelError> erros = new List<ExcelWebProject.ExcelError>();
            var lista = excel.LerArquivo(Request.PhysicalApplicationPath +  @"\Arquivos\Pasta1.xlsx", "Plan1", out erros);
            if(Session["error"] != null && !string.IsNullOrEmpty(Session["error"].ToString()))
            {
                Response.Write($"<script> alert('{Session["error"].ToString()}')</script>");
                Session["error"] = null;
            }
        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            Label1.Text = DateTime.Now.ToString();
            throw new Exception("Teste!!!");
        }

        protected void Button2_Click(object sender, EventArgs e)
        {

        }
    }
}