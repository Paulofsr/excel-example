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