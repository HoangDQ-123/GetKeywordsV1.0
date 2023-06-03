using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MySql.Data.MySqlClient; 

namespace GetKeywords.Modules
{
    internal class clsConnectDB
    {
        private string strConnect = "";
        void OpenConnect()
        {
            // MySqlConnection conn = new MySqlConnection();
            strConnect = "server=" + InitVar.v_server + ";uid=" + InitVar.v_UID + ";pwd=" + InitVar.v_pass + ";database=" + InitVar.v_DBName;
            InitVar.conn.ConnectionString = strConnect;
            
        }
    }
}
