using System;
using System.Data;
using System.Data.OleDb;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Com.IAV.Database;

namespace InventorAccessPortal
{
    public partial class AnventorAccess : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            ConnDbForAcccess mydb = new ConnDbForAcccess();
            Response.Write(mydb.ReturnSqlResultCount("SELECT * From Codes;"));
            Response.Write("HelloWord!");
        }
    }
}