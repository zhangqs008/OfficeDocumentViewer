using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace Whir.Software.DocumentViewer
{
    public partial class Preview : System.Web.UI.Page
    {
        public string Url
        {
            get { return Request.QueryString["url"]; }
        }

        public string Source
        {
            get { return Request.QueryString["source"]; }
        }

        protected void Page_Load(object sender, EventArgs e)
        {

        }
    }
}