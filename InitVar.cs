using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GetKeywords
{
    public static class InitVar
    {
        public static int v_VolMax = 1000; // T.Hoang thêm vào biến, sửa tên textbox, quản lý Volume Max
        public static int v_speed = 1; // Tốc độ timer
        public static int v_LevelSearch = 0; // Độ sâu, đặt =0 là vô tận.


        public static string pathConfig = "config.txt";




        public static void SaveFileConfig(string path)
        {
            StreamWriter fw = new StreamWriter(InitVar.pathConfig);

            fw.WriteLine(v_VolMax);
            fw.WriteLine(v_speed);
            fw.WriteLine(v_LevelSearch);

            fw.Close();
        }
        public static void OpenFileConfig(string path)
        {
            StreamReader fr = new StreamReader(path);

            v_VolMax = Convert.ToInt32(fr.ReadLine()); 
            v_speed = Convert.ToInt32(fr.ReadLine()); 
            v_LevelSearch = Convert.ToInt32(fr.ReadLine());

            fr.Close();
        }
    }
}
