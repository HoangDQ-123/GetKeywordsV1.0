using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
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
        public void CloseConnect()
        {
            InitVar.conn.Close();
        }

        int InsertKeys(DataGridViewRow d_row, string d_header) // chú ý khớp header và số cột trong DataGrid
        {
            string strQuery = "";
            string tbName = "t_keys";
            string strHeader = "";
            string strValue = "";

            strQuery += "insert into ";
            strQuery += tbName + "(";
            try
            {
                string[] tmpHeader = d_header.Split('|');
                for(int i=0; i<tmpHeader.Length;i++)
                //foreach (string s in tmpHeader)
                {
                    strHeader += tmpHeader[i] + ",";
                    strValue += d_row.Cells[i] + "','";
                }

                strQuery += "(" + strHeader + ")" + "values('" + strValue.Substring(0, strValue.Length - 1) + "');";
                //string Query = "insert into student.studentinfo(idStudentInfo,Name,Father_Name,Age,Semester) values('" + this.IdTextBox.Text + "','" + this.NameTextBox.Text + "','" + this.FnameTextBox.Text + "','" + this.AgeTextBox.Text + "','" + this.SemesterTextBox.Text + "');";

                return 1;
            }
            catch (Exception ex)
            {
                return 0;
                MessageBox.Show(ex.Message);
            }
        }
    }
}
