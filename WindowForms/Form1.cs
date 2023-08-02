using GetKeywords.Modules;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System;
using System.Collections;
using System.Data;
using System.Diagnostics.Eventing.Reader;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace GetKeywords
{
    public partial class Form1 : Form
    {

        static System.Windows.Forms.Timer myTimer = new System.Windows.Forms.Timer();
        static int alarmCounter = 0;
        static bool exitFlag = false;
        static System.Drawing.Point pt;
        private int[] StepTimer = new int[100];
        private int[] StepTimer01 = new int[100];
        private int[] StepTimer02 = new int[100];
        private int[] StepTimer03 = new int[100];
        private int[] StepTimer04 = new int[100];
        private int[] NextStepDelay = new int[100];
        private int[] NextStepDelay01 = new int[100];
        private int[] NextStepDelay02 = new int[100];
        private int[] NextStepDelay03 = new int[100];
        private int[] NextStepDelay04 = new int[100];
        

        static int d_errorfile = 0; // Số lần không lấy được file excel
        
        private ExcelConnect f;
        private string CurrentKeywords = null;

        private int KeyIndex = 0;
        private int KeyIndex_backup = 0;
        private int countNextKey = 0;
        private int countBack = 0;

        private int LoadFileExcelOK = 0;

        private int LevelSearch;

        private int NextKeyCount = 0;

        private string ListAccount;
        private string[] Account = new string[5];
        private int AccountIndex = 0;
        private string passLogin;

        private int max_Process_Plan01;
        private int max_Process_Plan02;
        private int max_Process_Plan03; // Số lượng tối đa tiến trình trên thanh trượt kịch bản 3
        private int max_Process_Plan04; // Số lượng tối đa tiến trình trên thanh trượt kịch bản 4

        private string[] ListSuggestKeys;
        private string[] ListNegativeKeys;
        private char[] Separator = { '|' };

        //private string[] strKeyWords= new string[2000000];
        //private string[] strVolume = new string[2000000];
        //private string[] strCheck = new string[2000000];
        //private string[] strComp = new string[2000000];

        private ArrayList strKeyWords = new ArrayList();
        private ArrayList strVolume = new ArrayList();
        private ArrayList strCheck = new ArrayList();
        private ArrayList strComp = new ArrayList();


        //Khởi tạo các nút kịch bản Login
        private System.Drawing.Point _AccountButtonLogin;
        private System.Drawing.Point _LogoutButtonLogin;
        private System.Drawing.Point _LoginButtonLogin;
        private System.Drawing.Point _EmailTextLogin;
        private System.Drawing.Point _PassTextLogin;
        private System.Drawing.Point _TextSearchLogin;
        private System.Drawing.Point _ButtonSearchLogin;

        // Khởi tạo các nút kịch bản 03:
        private System.Drawing.Point _TextSearch03;
        private System.Drawing.Point _ButtonSearch03;
        private System.Drawing.Point _ButtonDownload03;
        private System.Drawing.Point _ButtonExcel03;

        // Khởi tạo các nút kịch bản 04:
        private System.Drawing.Point _TextSearch04;
        private System.Drawing.Point _ButtonSearch04;
        private System.Drawing.Point _ClickCloudFare04;

        private Random rand = new Random();



        public Form1()
        {
            InitializeComponent();
            // Lấy các dữ liệu setting
            // Delay Time after event 1
            NextStepDelay[0] = 1; //focus text search
            NextStepDelay[1] = 1; //input text search
            NextStepDelay[2] = 3; // click search
            NextStepDelay[3] = 1; // click download button
            NextStepDelay[4] = 1; // click excel
            NextStepDelay[5] = 2; // click stop

            StepTimer[0] = 2;
            for (int i = 0; i < 10; i++)
            {
                StepTimer[i + 1] = StepTimer[i] + NextStepDelay[i];
            }
            max_Process_Plan01 = 20;

            // Delay Time after event 2
            NextStepDelay01[0] = 1; //focus to link
            NextStepDelay01[1] = 1; //input text link
            NextStepDelay01[2] = 3; // click enter
            NextStepDelay01[3] = 2; // click Account Menu
            NextStepDelay01[4] = 2; // click to Logout
            NextStepDelay01[5] = 2; // click to Login Again
            NextStepDelay01[6] = 2; // focus to tai khoan
            NextStepDelay01[7] = 1; // input text tai khoan
            NextStepDelay01[8] = 1; // focus to mat khau
            NextStepDelay01[9] = 1; // input text mat khau
            NextStepDelay01[10] = 3; // click to login
            NextStepDelay01[11] = 1; //focus to text keyword
            NextStepDelay01[12] = 1; //input to text keyword
            NextStepDelay01[13] = 3; // click search
            NextStepDelay01[14] = 1; // click download button
            NextStepDelay01[15] = 1; // click excel
            NextStepDelay01[16] = 2; // click stop

            StepTimer01[0] = 2;
            for (int j = 0; j < 20; j++)
            {
                StepTimer01[j + 1] = StepTimer01[j] + NextStepDelay01[j];
            }
            max_Process_Plan02 = 20;

            // Delay Time after event 3
            NextStepDelay02[0] = 1; // + rand.Next(100); //focus text search
            NextStepDelay02[1] = 1; //+ rand.Next(100); //CtrlA
            NextStepDelay02[2] = 1; //+ rand.Next(100); // input text search
            NextStepDelay02[3] = 4; //+ rand.Next(100); // click search
            NextStepDelay02[4] = 1; //+ rand.Next(100); // click download button
            NextStepDelay02[5] = 2; //+ rand.Next(100); // click excel
            NextStepDelay02[6] = 1; //+ rand.Next(100); // click kiem tra file
            NextStepDelay02[7] = 1; //+ rand.Next(100); // click stop

            StepTimer02[0] = 2;
            for (int m = 0; m < 10; m++)
            {
                StepTimer02[m + 1] = StepTimer02[m] + NextStepDelay02[m];
            }
            max_Process_Plan03 = 20;

            // Delay Time after event 4
            NextStepDelay03[0] = 2; //click to link
            NextStepDelay03[1] = 2; //input dia chi website
            NextStepDelay03[2] = 2; //Enter
            NextStepDelay03[3] = 1; //Click CloudFare
            NextStepDelay03[4] = 2; // focus to text
            NextStepDelay03[5] = 2; // input to text
            NextStepDelay03[6] = 2; // click search
            NextStepDelay03[7] = 2; // click stop

            StepTimer03[0] = 2;
            for (int n = 0; n < 10; n++)
            {
                StepTimer03[n + 1] = StepTimer03[n] + NextStepDelay03[n];
            }

            max_Process_Plan04 = 20;
            //event5
            NextStepDelay04[0] = 2; //focus tai khoan
            NextStepDelay04[1] = 2; //input tai khoan
            NextStepDelay04[2] = 3; //focus mat khau
            NextStepDelay04[3] = 4; //input mat khau
            NextStepDelay04[4] = 2; // click to login
            NextStepDelay04[5] = 2; //focus to text keyword
            NextStepDelay04[6] = 2; //input to text keyword
            NextStepDelay04[7] = 2; //click to search
            NextStepDelay04[8] = 1; //click to stop
           
            for (int a = 0; a < 10; a++)
            {
                StepTimer04[a + 1] = StepTimer04[a] + NextStepDelay04[a];
            }

            // Load Toa do Kich ban
            LoadPlanPoint("Plan.xlsx");
        }
        private const int MOUSEEVENTF_LEFTDOWN = 0x02;
        private const int MOUSEEVENTF_LEFTUP = 0x04;
        private const int MOUSEEVENTF_RIGHTDOWN = 0x08;
        private const int MOUSEEVENTF_RIGHTUP = 0x10;

        [DllImport("user32.dll")]

        // Định nghĩa hàm mouse_event() từ thư viện user32.dll
        private static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);

        // ...
        private void btnStart_Click(object sender, EventArgs e)
        {
            if (btnStart.Text == "Start") { 
            string strSuggestKey = txtSuggestKey.Text;
            string strNegativeKey = txtNegativeKey.Text;
            ListSuggestKeys = strSuggestKey.Split(Separator);
            ListNegativeKeys = strNegativeKey.Split(Separator);

                ListAccount = txttaikhoan.Text;
                Account = ListAccount.Split(Separator);
                passLogin = txtmatkhau.Text;

                Cursor.Show();  // Cho phép hiện con trỏ chuột lên

                alarmCounter = 0;
                exitFlag = false;
                d_errorfile = 0;
                // KeyIndex = 0;   //// Lưu ý trường hợp có cần phải tính toàn lại loadkeys Index không 
                LoadFileExcelOK = 0;
                //InitVar.v_speed = Convert.ToInt32(txtSpeed.Text);
                //InitVar.v_VolMax = Convert.ToInt32(txtVolMax.Text);


                if (cboPlan.SelectedIndex == 0) // Lựa chọn Login
                {
                    progressBar1.Maximum = max_Process_Plan02; // số lượng các thao tác trong kế hoạch.
                    progressBar1.Value = 0;
                    tmrPlan02.Interval = InitVar.v_speed;
                    tmrPlan02.Start();
                    cboPlan.Text = "Kịch bản 02";
                }
                if (cboPlan.SelectedIndex == 1) // Lựa chọn Get Keywords
                {

                    progressBar1.Maximum = max_Process_Plan01; // số lượng các thao tác trong kế hoạch.
                    progressBar1.Value = 0;
                    tmrPlan01.Interval = InitVar.v_speed;
                    tmrPlan01.Start();
                    cboPlan.Text = "Kịch bản 01";
                }
                if (cboPlan.SelectedIndex == 2) // Lựa chọn Get Keywords tiep theo
                {

                    progressBar1.Maximum = max_Process_Plan03; // số lượng các thao tác trong kế hoạch 03.
                    progressBar1.Value = 0;
                    tmrPlan03.Interval = InitVar.v_speed;
                    tmrPlan03.Start();
                    cboPlan.Text = "Kịch bản 03";
                }
                if (cboPlan.SelectedIndex == 3) // không tìm thấy file ex
                {
                    progressBar1.Maximum = max_Process_Plan04; // số lượng các thao tác trong kế hoạch 04.
                    progressBar1.Value = 0;
                    tmrPlan04.Interval = InitVar.v_speed;
                    tmrPlan04.Start();
                    cboPlan.Text = "Kịch bản 04";
                }
                if (cboPlan.SelectedIndex == 4) // xen
                {
                    progressBar1.Maximum = 9; // số lượng các thao tác trong kế hoạch.
                    progressBar1.Value = 0;
                    tmrPlan05.Interval = InitVar.v_speed;
                    tmrPlan05.Start();
                    cboPlan.Text = "Kịch bản 05";
                }
                btnStart.Text = "Stop";
            }
            else
            {
                btnStart.Text = "Start";
                cboPlan.SelectedIndex = 2;  // Lựa chọn sẵn kịch bản 03
                tmrPlan01.Stop();
                tmrPlan02.Stop();
                tmrPlan03.Stop();
                tmrPlan04.Stop();
                tmrPlan05.Stop();
                if (alarmCounter >= StepTimer02[2]) KeyIndex--;
                //MessageBox.Show("Vui lòng thiết lập lại trạng thái bắt đầu nếu muốn Start");  // Tạm thời bỏ thông báo này để next key
            }    

        }

        private void Form1_Load(object sender, EventArgs e)
        {

            InitVar.OpenFileConfig(InitVar.pathConfig);

            // Khởi tạo tạm các keyGG & ChatGPT
            InitVar.v_arrKeyGG = new string[100];
            InitVar.v_arrKeyGG[0] = "AIzaSyCSXStlfJDlEoikXv6P7yEOSRb2PsVjZAM";
            InitVar.v_arrKeyGG[1] = "AIzaSyBA5C5BzTPq1Ooi4x7rAytNtqTjjdGJzK8";
            InitVar.v_arrKeyGG[2] = "AIzaSyB-sU_otHwxn2JNwIqI42O0gHLEk-mkZtY";
            InitVar.v_arrKeyGG[3] = "AIzaSyCUZBpGQNUs1AJJdH8lSsjSUv2dxmN1zWI";
            InitVar.v_arrKeyGG[4] = "AIzaSyAk2bQ23muiPHYXf2yFN7GlRic3vpXFh4Y";
            InitVar.v_arrKeyGG[5] = "AIzaSyCQbkcrYYQTLz_IF-SDmFbVefjHHxyTNM8";
            InitVar.v_arrKeyGG[6] = "AIzaSyClUiZTcwp359Kbb-W7WDmGziVCjtWA37M";
            InitVar.v_arrKeyGG[7] = "AIzaSyDft1aB-jQ-Kpk9tE_HjTQm9mHvwlgNaWk";
            InitVar.v_arrKeyGG[8] = "AIzaSyDdNIKC3t7J9elX3QXUSFD7ELgr3I4UPzY";
            InitVar.v_arrKeyGG[9] = "AIzaSyD2rfceyTgDh1QzTH-uEJmOlyk-goIEW38";

            InitVar.v_arrKeyChatGPT = new string[100];
            InitVar.v_arrKeyChatGPT[0] = "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";
            InitVar.v_arrKeyChatGPT[1] = "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";
            InitVar.v_arrKeyChatGPT[2] = "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";
            InitVar.v_arrKeyChatGPT[3] = "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";
            InitVar.v_arrKeyChatGPT[4] = "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";
            InitVar.v_arrKeyChatGPT[5] = "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";
            InitVar.v_arrKeyChatGPT[6] = "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";
            InitVar.v_arrKeyChatGPT[7] = "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";
            InitVar.v_arrKeyChatGPT[8] = "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";
            InitVar.v_arrKeyChatGPT[9] = "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";


            // Lấy các dữ liệu setting
            //InitVar.v_speed = Convert.ToInt32(txtSpeed.Text);
            //InitVar.v_VolMax = Convert.ToInt32(txtVolMax.Text); // T.Hoàng thêm 15:54 20230302

           

            // Thêm kịch bản
            cboPlan.Items.Clear();
            cboPlan.Items.Add("Login");
            cboPlan.Items.Add("Get keywords");
            cboPlan.Items.Add("Dowload Keyword tiep theo");
            cboPlan.Items.Add("Dowload lại nếu không tìm thấy file excel");

            cboPlan.SelectedIndex = 2;

            // Mở kết nối file excel
            //f.fileName = "Keyword Tool Export -Keyword Suggestions - " + CurrentKeywords;

            foreach (DataGridViewColumn column in dgrListKeywords.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
        }

        /// Đoạn code save file config
        /// 


        /// <summary>
        /// Day la doan nhap file excel thu 1
        /// </summary>
        /// <param name="path"></param>
        private void Importexcel(string path)
        {
            using (FileStream fileStream = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                using (ExcelPackage excelPackage = new ExcelPackage(fileStream))
                {
                    // Lấy Sheet đầu tiên trong tệp Excel
                    ExcelWorksheet excelWorksheet = excelPackage.Workbook.Worksheets[0];

                    int i = 1;
                    int index = 0;
                    while ((excelWorksheet.Cells[i + 1, 1].Value != null) && (excelWorksheet.Cells[i + 1, 2].Value != null) && (excelWorksheet.Cells[i + 1, 2].Value.ToString().Contains(".") == false))
                    {
                        // Cột 17 là vị trí của cột Competition;......
                        if (excelWorksheet.Cells[i + 1, 3].Value == null)
                        {
                            strKeyWords.Add(Convert.ToString( excelWorksheet.Cells[i + 1, 1].Value));
                            strVolume.Add(Convert.ToString(excelWorksheet.Cells[i + 1, 2].Value));
                            strCheck.Add("0");
                            strComp.Add(Convert.ToString(excelWorksheet.Cells[i + 1, 17].Value));
                        }
                        else
                        {
                            strKeyWords.Add(Convert.ToString(excelWorksheet.Cells[i + 1, 1].Value));
                            strVolume.Add(Convert.ToString(excelWorksheet.Cells[i + 1, 2].Value));
                            strCheck.Add(Convert.ToString(excelWorksheet.Cells[i + 1, 3].Value));
                            strComp.Add(Convert.ToString(excelWorksheet.Cells[i + 1, 17].Value));
                        }
                        i++; index++;
                    }

                    if (excelWorksheet.Cells[1, 27].Value != null)  //
                    {
                        KeyIndex = Convert.ToInt32(excelWorksheet.Cells[1, 27].Value);
                        txtCur.Text = Convert.ToString(KeyIndex);
                    }

                    if (excelWorksheet.Cells[2, 27].Value != null)  //
                    {
                        txtSuggestKey.Text = Convert.ToString(excelWorksheet.Cells[2, 27].Value);
                    }
                    if (excelWorksheet.Cells[3, 27].Value != null)  //
                    {
                        txtNegativeKey.Text = Convert.ToString(excelWorksheet.Cells[3, 27].Value);
                    }
                    if (excelWorksheet.Cells[4, 27].Value != null)  //
                    {
                        txttaikhoan.Text = Convert.ToString(excelWorksheet.Cells[4, 27].Value);
                    }
                    if (excelWorksheet.Cells[5, 27].Value != null)  //
                    {
                        txtmatkhau.Text = Convert.ToString(excelWorksheet.Cells[5, 27].Value);
                    }

                }
            }


           
            txtTotal.Text = Convert.ToString(strKeyWords.Count);  
        }
        /// <summary>
        /// doan nhap file cho vong lay keywords tiep theo
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// 
        private async void ImportExcelCircle(string path)
        {
            //int kq = 1;
            LoadFileExcelOK = 1;
            using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(path)))
            {
                ExcelWorksheet excelWorksheet = excelPackage.Workbook.Worksheets[0];
                int i = 1;
                while (excelWorksheet.Cells[i + 1, 1].Value != null) //&& (excelWorksheet.Cells[i + 1, 2].Value!= null)) // && Convert.ToInt32(excelWorksheet.Cells[i + 1, 2].Value) > 1000))
                {
                    
                    // THoang code 21:46 20230301
                    if ((excelWorksheet.Cells[i + 1, 2].Value != null) && ( Convert.ToInt32(excelWorksheet.Cells[i + 1, 2].Value) >= InitVar.v_VolMin) && (Convert.ToInt32(excelWorksheet.Cells[i + 1, 17].Value) <= InitVar.v_LevelDif))
                        {
                        //kq = 0;
                        LoadFileExcelOK = 0;
                        string str2 = excelWorksheet.Cells[i + 1, 1].Value.ToString();
                        //if (Convert.ToInt32(excelWorksheet.Cells[i + 1, 2].Value) >= v_VolMax)
                        //{
                        // Kiem tra trung lap trong danh sach
                        bool dup = false;
                            for (int j = 0; j <= strKeyWords.Count - 1; j++)
                            {

                            string str1 = strKeyWords[j].ToString();  
                                if (str1.Equals(str2, StringComparison.InvariantCultureIgnoreCase))
                                {
                                    dup = true;
                                    CurrentKeywords = txtKeywords.Text;
                                    if (CurrentKeywords.Equals(str2, StringComparison.InvariantCultureIgnoreCase))
                                    {
                                        strVolume[KeyIndex-1] = excelWorksheet.Cells[i + 1, 2].Value;
                                    
                                    }
                                break;
                                }
    
                            }
                            if (dup == false)
                            {
                                // Kiểm tra ListSuggest & Negative
                                bool sug = false;

                                for (int t1 = 0; t1 <= ListSuggestKeys.Length - 1; t1++)
                                {
                                int indexSub = str2.IndexOf(ListSuggestKeys[t1]);
                                //if (str2.Contains(ListSuggestKeys[t1]) == true)
                                if (indexSub >= 0)
                                    {
                                    sug = true;
                                        break;
                                    }
                                }

                                bool nega = true;
                                
                                for (int t2 = 0; t2 <= ListNegativeKeys.Length - 1; t2++)
                                {
                                if (ListNegativeKeys[t2] != "")
                                    {
                                        int indexSub = str2.IndexOf(ListNegativeKeys[t2]);
                                        //if(str2.Contains(ListNegativeKeys[t2]) == true)
                                        if (indexSub >= 0 )
                                            {
                                                nega = false;
                                                break;
                                            }
                                    }
                                }




                                //for (int t1 = 0; t1 <= ListSuggestKeys.Length-1; t1++)


                                //for (int t2 = 0; t2 <= ListNegativeKeys.Length - 1; t2++)
                                //{
                                //    if (ListNegativeKeys[t2] == "")
                                //    {
                                //        if (str2.Contains(ListSuggestKeys[t1]) == true)
                                //        {
                                //            sug = true;
                                //            break;
                                //        }
                                //    }
                                //    else
                                //    {
                                //        if ((str2.Contains(ListSuggestKeys[t1]) == true) && str2.Contains(ListNegativeKeys[t2]) == false)
                                //        {
                                //            sug = true;
                                //            break;
                                //        }    
                                //    }    
                                //}

                                if ((sug == true) && (nega == true))
                                {
                                //string InputRequest = "";
                                //string OutputContent = await clsAPI.CallChatGPTAPI(InputRequest); // Gọi hàm từ clsAPI.

                                strKeyWords.Insert(KeyIndex, excelWorksheet.Cells[i + 1, 1].Value);
                                strVolume.Insert(KeyIndex, excelWorksheet.Cells[i + 1, 2].Value);
                                strCheck.Insert(KeyIndex, (LevelSearch + 1));
                                strComp.Insert(KeyIndex, excelWorksheet.Cells[i + 1, 17].Value);
                                }

 
                        }

                        //}
                   
                    }

                    
                    i++;
                }
                File.Delete(path); // THoang: Xóa luôn file sau khi đã nạp
                //return kq;
            }
        }

        /// <summary>
        /// Load dữ liệu tọa độ điểm Kịch bản
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        private void LoadPlanPoint(string path)
        {
            try
            {
                using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(path)))
                {
                    // Load kich ban Login
                    ExcelWorksheet excelWorksheetPlanLogin = excelPackage.Workbook.Worksheets["PlanLogin"];

                    _AccountButtonLogin.X = Convert.ToInt32(excelWorksheetPlanLogin.Cells[2, 2].Value);
                    _AccountButtonLogin.Y = Convert.ToInt32(excelWorksheetPlanLogin.Cells[2, 3].Value);

                    _LogoutButtonLogin.X = Convert.ToInt32(excelWorksheetPlanLogin.Cells[3, 2].Value);
                    _LogoutButtonLogin.Y = Convert.ToInt32(excelWorksheetPlanLogin.Cells[3, 3].Value);

                    _EmailTextLogin.X = Convert.ToInt32(excelWorksheetPlanLogin.Cells[4, 2].Value);
                    _EmailTextLogin.Y = Convert.ToInt32(excelWorksheetPlanLogin.Cells[4, 3].Value);

                    _PassTextLogin.X = Convert.ToInt32(excelWorksheetPlanLogin.Cells[5, 2].Value);
                    _PassTextLogin.Y = Convert.ToInt32(excelWorksheetPlanLogin.Cells[5, 3].Value);

                    _LoginButtonLogin.X = Convert.ToInt32(excelWorksheetPlanLogin.Cells[6, 2].Value);
                    _LoginButtonLogin.Y = Convert.ToInt32(excelWorksheetPlanLogin.Cells[6, 3].Value);

                    _TextSearchLogin.X = Convert.ToInt32(excelWorksheetPlanLogin.Cells[7, 2].Value);
                    _TextSearchLogin.Y = Convert.ToInt32(excelWorksheetPlanLogin.Cells[7, 3].Value);

                    _ButtonSearchLogin.X = Convert.ToInt32(excelWorksheetPlanLogin.Cells[8, 2].Value);
                    _ButtonSearchLogin.Y = Convert.ToInt32(excelWorksheetPlanLogin.Cells[8, 3].Value);

                    /////////////////////////////////////////////////
                    //

                    // Load kich ban 03
                    ExcelWorksheet excelWorksheetPlan03 = excelPackage.Workbook.Worksheets["Plan03"];

                    _TextSearch03.X = Convert.ToInt32(excelWorksheetPlan03.Cells[2, 2].Value);
                    _TextSearch03.Y = Convert.ToInt32(excelWorksheetPlan03.Cells[2, 3].Value);

                    _ButtonSearch03.X = Convert.ToInt32(excelWorksheetPlan03.Cells[3, 2].Value);
                    _ButtonSearch03.Y = Convert.ToInt32(excelWorksheetPlan03.Cells[3, 3].Value);

                    _ButtonDownload03.X = Convert.ToInt32(excelWorksheetPlan03.Cells[4, 2].Value);
                    _ButtonDownload03.Y = Convert.ToInt32(excelWorksheetPlan03.Cells[4, 3].Value);

                    _ButtonExcel03.X = Convert.ToInt32(excelWorksheetPlan03.Cells[5, 2].Value);
                    _ButtonExcel03.Y = Convert.ToInt32(excelWorksheetPlan03.Cells[5, 3].Value);
                    /////////////////////////////////////////////////


                    // Load kich ban 04
                    ExcelWorksheet excelWorksheetPlan04 = excelPackage.Workbook.Worksheets["Plan04"];

                    _TextSearch04.X = Convert.ToInt32(excelWorksheetPlan04.Cells[2, 2].Value);
                    _TextSearch04.Y = Convert.ToInt32(excelWorksheetPlan04.Cells[2, 3].Value);

                    _ButtonSearch04.X = Convert.ToInt32(excelWorksheetPlan04.Cells[3, 2].Value);
                    _ButtonSearch04.Y = Convert.ToInt32(excelWorksheetPlan04.Cells[3, 3].Value);

                    _ClickCloudFare04.X = Convert.ToInt32(excelWorksheetPlan04.Cells[4, 2].Value);
                    _ClickCloudFare04.Y = Convert.ToInt32(excelWorksheetPlan04.Cells[4, 3].Value);

                    /////////////////////////////////////////////////

                    excelPackage.Dispose();
                }
              
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error load PlanPoint: " + ex.Message);
            }
        }

        private void openExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "Import Excel";
            openFileDialog.Filter = "Excel(*.xlsx)|*.xlsx|Excel 2016(*.xls)|*.xls";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    Importexcel(openFileDialog.FileName);
                    MessageBox.Show("Nhap file thanh cong");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Nhap file khong thanh cong \n" + ex.Message);
                }
            }

            if (strKeyWords.Count > 0) {
                txtKeywords.Text = Convert.ToString(strKeyWords[KeyIndex]);
            }
        }
        private void ExportExcel(string path)
        {
            Excel.Application application = new Excel.Application();
            application.Application.Workbooks.Add(Type.Missing);
            for (int i = 0; i < dgrListKeywords.Columns.Count; i++)
            {
                application.Cells[1, i + 1] = dgrListKeywords.Columns[i].HeaderText;
            }
            for (int i = 0; i <= dgrListKeywords.Rows.Count - 1; i++)
            {
                for (int j = 0; j <= dgrListKeywords.Columns.Count - 1; j++)
                {
                    application.Cells[i + 2, j + 1] = dgrListKeywords.Rows[i].Cells[j].Value;
                    //application.Cells[i + 2, j + 1] = dgrListKeywords.Rows[i].Cells[j].Value;
                }
            }
            application.Columns.AutoFit();
            application.ActiveWorkbook.SaveCopyAs(path);
            application.ActiveWorkbook.Saved = true;
            application.Quit();
        }
        private void QuickExportExcel(string path, int noMsg = 0 )
        {
            //DataTable dt = new DataTable();

            if (strKeyWords.Count >= 1)
            {

                // Tạo một đối tượng ExcelPackage
                using (ExcelPackage excelPackage = new ExcelPackage())
                {
                    // Tạo một đối tượng Worksheet
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");

                    // Lấy dữ liệu từ DataGrid và đổ vào worksheet
                    for (int i = 0; i < strKeyWords.Count-1; i++)
                    {
                        worksheet.Cells[i + 1, 1].Value = strKeyWords[i];
                        worksheet.Cells[i + 1, 2].Value = strVolume[i];
                        worksheet.Cells[i + 1, 3].Value = strCheck[i];
                        worksheet.Cells[i + 1, 4].Value = strComp[i];
                    }
                    worksheet.Cells[1, 27].Value = KeyIndex;  // Ghi lại KeyIndex để tiếp tục xử lý tại ô "AA:1"
                    worksheet.Cells[2, 27].Value = txtSuggestKey.Text;
                    worksheet.Cells[3, 27].Value = txtNegativeKey.Text;
                    worksheet.Cells[4, 27].Value = txttaikhoan.Text;
                    worksheet.Cells[5, 27].Value = txtmatkhau.Text;

                    // Lưu workbook vào một MemoryStream
                    using (MemoryStream memoryStream = new MemoryStream())
                    {
                        excelPackage.SaveAs(memoryStream);
                        memoryStream.Position = 0;

                        // Lưu MemoryStream vào file Excel
                        using (FileStream fileStream = new FileStream(path, FileMode.Create, FileAccess.Write))
                        {
                            memoryStream.WriteTo(fileStream);
                        }
                    }
                }

                if (noMsg == 0) MessageBox.Show("Xuat file thanh cong");
            }
            else
            {
                if (noMsg == 0) MessageBox.Show("There is NO keywords to Export");
            }    
        }
        private void saveExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult dlr = MessageBox.Show("Khyến nghị! Bạn nên dùng chức năng Xuất nhanh Excel." + Environment.NewLine + "Bạn có chắc chắn muốn tiếp tục với Xuất Excel bình thường không?", "Viện Tin học Xây dựng", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dlr == DialogResult.Yes)
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Title = "Export Excel";
                saveFileDialog.Filter = "Excel(*.xlsx)|*.xlsx|Excel 2016(*.xls)|*.xls";
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //try
                    {
                        ExportExcel(saveFileDialog.FileName);
                        MessageBox.Show("Xuat file thanh cong");
                    }
                    //catch (Exception ex)
                    //{
                    //    MessageBox.Show("Xuat file khong thanh cong \n" + ex.Message);
                    //}
                }
            }
        }
        private void btnPause_Click(object sender, EventArgs e)
        {
            if (btnPause.Text == "Pause")
            {
                btnPause.Text = "Continue";
                tmrPlan03.Stop();
            }
            else
            {
                btnPause.Text = "Pause";
                tmrPlan03.Start();
            }
        }

        private void tmrPlanGetKeyword(object sender, EventArgs e)
        {
            
            alarmCounter++;

            if (alarmCounter == StepTimer[0]) //focus text search
            {
                pt.X = 382;
                pt.Y = 486;
                Cursor.Position = pt;
                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);


                progressBar1.Value += 1;
            }
            if (alarmCounter == StepTimer[1]) //input text search
            {
                SendKeys.Send(txtKeywords.Text);


                progressBar1.Value += 1;
            }

            if (alarmCounter == StepTimer[2]) // click search
            {
                pt.X = 941;
                pt.Y = 486;
                Cursor.Position = pt;
                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);


                progressBar1.Value += 1;
            }

            if (alarmCounter == StepTimer[3]) // click download button
            {
                pt.X = 971;
                pt.Y = 937;
                Cursor.Position = pt;
                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);


                progressBar1.Value += 1;
            }

            if (alarmCounter == StepTimer[4]) // click export to excel
            {
                pt.X = 918;
                pt.Y = 751;
                Cursor.Position = pt;
                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);


                progressBar1.Value += 1;
            }
            if (alarmCounter == StepTimer[5])
            {
                tmrPlan01.Stop();
                alarmCounter = 0;

                this.WindowState = FormWindowState.Normal;


                progressBar1.Value += 1;
            }
        }

        private void tmrPlanLogin(object sender, EventArgs e)
        
        {
            alarmCounter++;
                if (alarmCounter == StepTimer01[0]) //CLICK TO link
                {
                    pt.X = 610;
                    pt.Y = 60;
                    Cursor.Position = pt;
                    mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);


                    progressBar1.Value += 1;
                }
                if (alarmCounter == StepTimer01[1]) //input text dia chi
                {
                    SendKeys.Send(txtdiachi.Text);


                    progressBar1.Value += 1;
                }
                if (alarmCounter == StepTimer01[2]) //input enter
                {
                    //SendKeys.SendWait("+(CTRL)");
                    //SendKeys.SendWait("+(A)");
                    SendKeys.Send("{ENTER}");


                    progressBar1.Value += 1;
                }
                if (alarmCounter == StepTimer01[3]) //CLICK TO Account Button ở góc
                {
                    pt.X = _AccountButtonLogin.X;
                    pt.Y = _AccountButtonLogin.Y;
                    Cursor.Position = pt;
                    mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);


                    progressBar1.Value += 1;
                }
                if (alarmCounter == StepTimer01[4]) //CLICK TO Logout ở cuối Menu
                {
                    pt.X = _LogoutButtonLogin.X;
                    pt.Y = _LogoutButtonLogin.Y;
                    Cursor.Position = pt;
                    mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);


                    progressBar1.Value += 1;
                }
                if (alarmCounter == StepTimer01[5]) //Click lại nút Login 1 lần nữa ở góc
                {
                    pt.X = _AccountButtonLogin.X;
                    pt.Y = _AccountButtonLogin.Y;
                    Cursor.Position = pt;
                    mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                    progressBar1.Value += 1;
                }

            if (alarmCounter == StepTimer01[6]) //Click Vào ô text nhập Email
            {
                pt.X = _EmailTextLogin.X;
                pt.Y = _EmailTextLogin.Y;
                Cursor.Position = pt;
                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                progressBar1.Value += 1;
            }

            if (alarmCounter == StepTimer01[7]) //input to tai khoan
                {
                AccountIndex++;
                if(AccountIndex >= Account.Length-1) AccountIndex = 0;
                SendKeys.Send(Account[AccountIndex]);

                    progressBar1.Value += 1;
                }
                if (alarmCounter == StepTimer01[8]) //focus text mat khau
                {
                    pt.X = _PassTextLogin.X;
                    pt.Y = _PassTextLogin.Y;
                    Cursor.Position = pt;
                    mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                    progressBar1.Value += 1;
                }
                if (alarmCounter == StepTimer01[9]) //input text mat khau
                {
                    SendKeys.Send(passLogin);
                    progressBar1.Value += 1;
                }
                if (alarmCounter == StepTimer01[10]) // click login
                {
                    //pt.X = _LogoutButtonLogin.X;
                    //pt.Y = _LogoutButtonLogin.Y;
                    //Cursor.Position = pt;
                    //mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                    SendKeys.Send("{ENTER}");
                    progressBar1.Value += 1;
                }
                if (alarmCounter == StepTimer01[11]) //focus text search
                {
                    pt.X = 382;
                    pt.Y = 486;
                    Cursor.Position = pt;
                    mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);

                    progressBar1.Value += 1;
                }
                if (alarmCounter == StepTimer01[12]) //input text search
                {
                    SendKeys.Send(txtKeywords.Text);

                    progressBar1.Value += 1;
                }
                if (alarmCounter == StepTimer01[13]) // click search
                {
                    pt.X = _ButtonSearch04.X;
                    pt.Y = _ButtonSearch04.Y;
                    Cursor.Position = pt;
                    mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);

                    progressBar1.Value += 1;
                }

                //if (alarmCounter == StepTimer01[13]) // click download button
                //{
                //    pt.X = _;
                //    pt.Y = 937;
                //    Cursor.Position = pt;
                //    mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);

                //    progressBar1.Value += 1;
                //}

                //if (alarmCounter == StepTimer01[14]) // click export to excel
                //{
                //    pt.X = 918;
                //    pt.Y = 751;
                //    Cursor.Position = pt;
                //    mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);

                //    progressBar1.Value += 1;
                //}
                if (alarmCounter == StepTimer01[14])
                {
                    tmrPlan02.Stop();
                    alarmCounter = 0;
                    //this.WindowState = FormWindowState.Normal;
                    progressBar1.Value = 0;
                    tmrPlan03.Start();
                }
             }

        private void tmrPlanDowloadKeywordtieptheo(object sender, EventArgs e)
        {
            //try
            {


                alarmCounter++;

                if (alarmCounter == StepTimer02[0]) //focus text search
                {
                    //pt.X = 419;
                    //pt.Y = 238;
                    pt = _TextSearch03;


                    Cursor.Position = pt;
                    mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);

                    progressBar1.Value += 1;
                }
                if (alarmCounter == StepTimer02[1]) //input text search
                {
                    //SendKeys.SendWait("+(CTRL)");
                    //SendKeys.SendWait("+(A)");
                    SendKeys.Send("^(a)");


                    progressBar1.Value += 1;
                }
                if (alarmCounter == StepTimer02[2]) //input text search
                {
                    if (strKeyWords.Count > 0)
                    {
                        do
                        {
                            KeyIndex++;
                            if(KeyIndex % InitVar.v_AutoSave == 0) //Sau 5000 keywords sẽ lưu tạm lại
                            {
                                string strTempPath = "";
                                DateTime now = DateTime.Now;
                                string formattedDateTime = now.ToString("dd/MM/yyyy HH:mm:ss");
                                string downloadPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\Downloads\";
                                strTempPath = downloadPath + "KeyTool" + formattedDateTime.Replace("/","").Replace(":","") + ".xlsx";
                                QuickExportExcel(strTempPath, 1);
                            }

                            //THoang 21:56 20230303
                            if (Convert.ToInt32(strVolume[KeyIndex - 1]) <= InitVar.v_VolMax)
                            {
                                strCheck[KeyIndex - 1] = "-1";
                            }
                            /////////
                            // Kiem tra den khi Vol het bang Grid thi coi nhu la da xong
                            if (KeyIndex >= strKeyWords.Count)  // Bỏ qua keywords cuối cùng, thinking tiếp nếu list có vol tận cuối cùng
                            {
                                // THoang 22:59 20230302
                                tmrPlan03.Stop();
                                alarmCounter = 0;
                                MessageBox.Show("Hoàn thành chiến dịch. Vui lòng Export File"); //Sau chuyển vào Label Trạng thái                        
                                break;
                            }
                            else
                            {
                                txtKeywords.Text = strKeyWords[KeyIndex - 1].ToString();
                                LevelSearch = Convert.ToInt32(strCheck[KeyIndex - 1].ToString());
                                // THoang 18:59 20230301
                               txtCur.Text = Convert.ToString(KeyIndex);
                            }

                            //THoang 21:56 20230303
                        } while ((Convert.ToInt32(strVolume[KeyIndex - 1]) <= InitVar.v_VolMax) || (Convert.ToInt32(strCheck[KeyIndex - 1]) >= InitVar.v_LevelSearch) || (Convert.ToInt32(strComp[KeyIndex - 1]) >= InitVar.v_LevelDif)); // Chi chay các keyword có vol >=1000 || chưa đánh dấu 100

                        //////
                    }
                    else
                    {
                        MessageBox.Show("Import file excel và Start lại");
                        tmrPlan03.Stop();
                        progressBar1.Value = 0;
                        alarmCounter = 0;
                    }

                    try
                    {
                        string sendString;
                        //sendString = "kết+quả+seagame+31+bóng+đá+nam";
                        //sendString = sendString.Replace("+", "{+}");

                        sendString = txtKeywords.Text.Replace("+", "{+}").Replace("^", "{^}").Replace("~", "{~}").Replace("%", "{%}"); //.Replace("(", "{(}").Replace(")", "{)}").Replace("{", "{{}").Replace("}", "{}}").Replace("[", "{[}").Replace("]", "{]}"); //Xử lý ký tự đặc biệt.
                                                                                                                                       //if (txtKeywords.Text.Contains("^"))
                        SendKeys.Send(sendString);

                        progressBar1.Value += 1;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message + " Chụp ảnh màn hình lỗi này. STOP, và có thể cân nhắc tự Nextkey để tiếp tục");
                    }
                    }

                if (alarmCounter == StepTimer02[3]) // click search
                {
                    //pt.X = 982;
                    //pt.Y = 238;
                    pt = _ButtonSearch03;

                    Cursor.Position = pt;
                    mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);

                    progressBar1.Value += 1;
                }

                if (alarmCounter == StepTimer02[4]) // click download button
                {
                    //pt.X = 971;
                    //pt.Y = 937;
                    pt = _ButtonDownload03;

                    Cursor.Position = pt;

                    mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);

                    progressBar1.Value += 1;
                }

                if (alarmCounter == StepTimer02[5]) // click export to excel
                {
                    //pt.X = 918;
                    //pt.Y = 751;
                    pt = _ButtonExcel03;

                    Cursor.Position = pt;
                    mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);

                    progressBar1.Value += 1;
                }
                if (alarmCounter == StepTimer02[6])
                {
                    // Kiem tra file excel co ton tai khong
                    // - Co ton tai: Import file excel them vao grid
                    // - Khong ton tai: ????

                    //tmrPlan03.Stop()
                    string fileName = txtKeywords.Text.Replace(".", " ").Replace("/", " ").Replace(":", " ").Replace("!", " ").Replace("&", "&amp").Replace("'","&#039").Replace("$", " ");
                    //fileName = fileName.Replace("+","{+}");
                    string downloadPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\Downloads\";
                    string filePath = downloadPath + "Keyword Tool Export - Keyword Suggestions - " + fileName + ".xlsx";
                    if ((System.IO.File.Exists(filePath) == true))
                    {

                        ImportExcelCircle(filePath);
                        // THoang 18:59 20230303
                        if (LoadFileExcelOK == 0)
                        {
                            txtTotal.Text = Convert.ToString(strKeyWords.Count);
                            strCheck[KeyIndex - 1] = LevelSearch; //"100"; // Giá trị cao vượt qua LevelSearch
                            //MessageBox.Show("Nhap file thanh cong");
                            //KeyIndex++;
                            d_errorfile = 0; // Khi tim thay file tra error ve =0, de chay vong kich ban 03.
                            if (countNextKey > 0)
                            {
                                KeyIndex = KeyIndex_backup;   //quay trở lại vị trí key đã quét mà bỏ qua trước đó
                                // Đếm số lần quay lại.
                                countBack++;
                            }
                            countNextKey = 0;
                        }
                        else
                        {
                            KeyIndex--;

                            d_errorfile++; // Tăng số lần đếm không lấy được file: 1-chạy 04; 2-chạy 05

                            if(d_errorfile == 1)
                            {
                                alarmCounter = StepTimer02[4] - 1;
                                KeyIndex++;
                            }
                            else if (d_errorfile == 2)   //if (d_errorfile == 1 || d_errorfile == 2)
                            {
                                tmrPlan03.Stop();

                                progressBar1.Value = 0;
                                progressBar1.Maximum = max_Process_Plan04;

                                alarmCounter = 0;
                                tmrPlan04.Start();
                                cboPlan.Text = "Kịch bản 04";
                            }
                            //else if (d_errorfile == 3)  // vao truong hop d_errorfile = 2 (xu ly tiep neu muon lon hon 2)
                            //{
                            //    tmrPlan03.Stop();
                            //    progressBar1.Value = 0;
                            //    progressBar1.Maximum = 9;
                            //    alarmCounter = 0;
                            //    tmrPlan05.Start();
                            //    cboPlan.Text = "Kịch bản 03";
                            //}
                            else
                            {
                                if (countNextKey == 0) KeyIndex_backup = KeyIndex;
                                countNextKey++;
                                KeyIndex++;
                                if (countNextKey >=3)
                                {
                                    tmrPlan03.Stop();
                                    alarmCounter = 0;
                                    progressBar1.Value = 0;
                                    //DialogResult r = MessageBox.Show("Lỗi nghiêm trọng! đã bỏ qua 3 keys vẫn chưa được. Vui lòng chạy lại trạng thái ban đầu", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    d_errorfile = 0;
                                    KeyIndex = KeyIndex_backup; // khi gặp lỗi 3 lần, cũng sẽ quay lại, hoặc khi loadOK excel
                                    //btnStart.Text = "Start";
                                    //cboPlan.SelectedIndex = 2;  // Lựa chọn sẵn kịch bản 03
                                    //tmrPlan01.Stop();
                                    //tmrPlan02.Stop();
                                    //tmrPlan03.Stop();
                                    //tmrPlan04.Stop();
                                    //tmrPlan05.Stop();
                                    //btnStart_Click(btnStart, EventArgs.Empty);
                                    //Viết lại kịch bản Login vào đây (chú ý, có 2 chỗ copy ở dưới nữa). Kiểm tra lại kịch bản 
                                    tmrPlan02.Start();
                                    cboPlan.Text = "Kịch bản Login";
                                }
                                else
                                {
                                    d_errorfile = 0;
                                    btnStart_Click(btnStart, EventArgs.Empty);
                                    btnNextKey_Click(btnNextKey, EventArgs.Empty);
                                    //btnStart.Text = "Start";
                                    if (KeyIndex >= strKeyWords.Count)  // Bỏ qua keywords cuối cùng, thinking tiếp nếu list có vol tận cuối cùng
                                    {
                                        // THoang 22:59 20230302
                                        tmrPlan03.Stop();
                                        alarmCounter = 0;
                                        MessageBox.Show("Hoàn thành chiến dịch. Vui lòng Export File"); //Sau chuyển vào Label Trạng thái                        
                                    }
                                    else
                                    {
                                        btnStart_Click(btnStart, EventArgs.Empty);
                                    }
                                }  
                                

                                //tmrPlan03.Stop();
                                //progressBar1.Value = 0;
                                //DialogResult result = MessageBox.Show("Lỗi nghiêm trọng, Bạn có muốn Next Keywords này không?? ","Thông báo lựa chọn", MessageBoxButtons.YesNo);
                                //if (result == DialogResult.No)
                                //{
                                //    d_errorfile = 0;
                                //    btnStart.Text = "Start";
                                //    cboPlan.SelectedIndex = 2;  // Lựa chọn sẵn kịch bản 03
                                //    tmrPlan01.Stop();
                                //    tmrPlan02.Stop();
                                //    tmrPlan03.Stop();
                                //    tmrPlan04.Stop();
                                //    tmrPlan05.Stop();
                                //    btnStart_Click(btnStart, EventArgs.Empty);
                                //}
                                //else
                                //{
                                //    d_errorfile = 0;
                                //    btnStart_Click(btnStart, EventArgs.Empty);
                                //    btnNextKey_Click(btnNextKey, EventArgs.Empty);
                                //    //btnStart.Text = "Start";
                                //    if (KeyIndex >= dgrListKeywords.Rows.Count)  // Bỏ qua keywords cuối cùng, thinking tiếp nếu list có vol tận cuối cùng
                                //    {
                                //        // THoang 22:59 20230302
                                //        tmrPlan03.Stop();
                                //        alarmCounter = 0;
                                //        MessageBox.Show("Hoàn thành chiến dịch. Vui lòng Export File"); //Sau chuyển vào Label Trạng thái                        
                                //    }
                                //    else
                                //    {
                                //        btnStart_Click(btnStart, EventArgs.Empty);
                                //    }
                                //}
                            }
                        }
                    }
                    else // của if(System.IO.File.Exists(filePath) == true)
                    {
                        // THoang 22:59 20230302

                        KeyIndex--;

                        d_errorfile++; // Tăng số lần đếm không lấy được file: 1-chạy 04; 2-chạy 05

                        if (d_errorfile == 1)
                        {
                            alarmCounter = StepTimer02[4]-1;
                            KeyIndex++;
                        }
                        else if (d_errorfile == 2)    //(d_errorfile == 1 || d_errorfile == 2)
                        {
                            tmrPlan03.Stop();

                            progressBar1.Value = 0;
                            progressBar1.Maximum = max_Process_Plan04;

                            alarmCounter = 0;
                            tmrPlan04.Start();
                            cboPlan.Text = "Kịch bản 04";
                        }
                        else if (d_errorfile == 3)  // vao truong hop d_errorfile = 2 (xu ly tiep neu muon lon hon 2)
                        {
                            tmrPlan03.Stop();
                            progressBar1.Value = 0;
                            progressBar1.Maximum = max_Process_Plan04;
                            alarmCounter = 0;
                            tmrPlan05.Start();
                            cboPlan.Text = "Kịch bản 05";
                        }
                        else
                        {
                            

                            if (countNextKey == 0) KeyIndex_backup = KeyIndex;
                            
                            countNextKey++;
                            KeyIndex++;

                            if (countNextKey >= 3)
                            {
                                tmrPlan03.Stop();
                                alarmCounter = 0;
                                progressBar1.Value = 0;
                                //DialogResult r = MessageBox.Show("Lỗi nghiêm trọng! đã bỏ qua 3 keys vẫn chưa được", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                d_errorfile = 0;
                                KeyIndex = KeyIndex_backup; // khi gặp lỗi 3 lần, cũng sẽ quay lại, hoặc khi loadOK excel
                                //btnStart.Text = "Start";
                                //cboPlan.SelectedIndex = 2;  // Lựa chọn sẵn kịch bản 03
                                //tmrPlan01.Stop();
                                //tmrPlan02.Stop();
                                //tmrPlan03.Stop();
                                //tmrPlan04.Stop();
                                //tmrPlan05.Stop();
                                //btnStart_Click(btnStart, EventArgs.Empty);
                                tmrPlan02.Start();
                                cboPlan.Text = "Kịch bản Login";
                            }
                            else
                            {
                                d_errorfile = 0;
                                btnStart_Click(btnStart, EventArgs.Empty);
                                btnNextKey_Click(btnNextKey, EventArgs.Empty);
                                //btnStart.Text = "Start";
                                if (KeyIndex >= strKeyWords.Count)  // Bỏ qua keywords cuối cùng, thinking tiếp nếu list có vol tận cuối cùng
                                {
                                    // THoang 22:59 20230302
                                    tmrPlan03.Stop();
                                    alarmCounter = 0;
                                    MessageBox.Show("Hoàn thành chiến dịch. Vui lòng Export File"); //Sau chuyển vào Label Trạng thái                        
                                }
                                else
                                {
                                    btnStart_Click(btnStart, EventArgs.Empty);
                                }
                            }
                        }
                    }

                    //alarmCounter = 0;

                    //this.WindowState = FormWindowState.Normal;

                    progressBar1.Value += 1;
                }
                if (alarmCounter == StepTimer02[7])
                {
                    alarmCounter = 0;

                    this.WindowState = FormWindowState.Normal;

                    progressBar1.Value = 0;
                }
            }
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //} 
        }

        private void tmrPlanLoadLaiExcel(object sender, EventArgs e)
        {
            alarmCounter++;

            if (alarmCounter == StepTimer03[0]) //CLICK TO link
            {
                pt.X = 610;
                pt.Y = 60;
                Cursor.Position = pt;
                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                progressBar1.Value += 1;
            }
            if (alarmCounter == StepTimer03[1]) //input text dia chi websites
            {
                SendKeys.Send(txtdiachi.Text);

                progressBar1.Value += 1;
            }
            if (alarmCounter == StepTimer03[2]) //input enter
            {
                //SendKeys.SendWait("+(CTRL)");
                //SendKeys.SendWait("+(A)");
                SendKeys.Send("{ENTER}");

                progressBar1.Value += 1;
            }

            if (alarmCounter == StepTimer03[3]) //Click CloudFare
            {
                ////pt.X = _ClickCloudFare04.X + 10;
                ////pt.Y = _ClickCloudFare04.Y + 10;
                ////Cursor.Position = pt;
                ////pt.X = _ClickCloudFare04.X - 10;
                ////pt.Y = _ClickCloudFare04.Y + 10;
                ////Cursor.Position = pt;
                ////pt.X = _ClickCloudFare04.X + 10;
                ////pt.Y = _ClickCloudFare04.Y - 10;
                ////Cursor.Position = pt;
                ////pt.X = _ClickCloudFare04.X - 10;
                ////pt.Y = _ClickCloudFare04.Y - 10;
                ////Cursor.Position = pt;

                //pt = _ClickCloudFare04;

                //Cursor.Position = pt;
                ////mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                //mouse_event(MOUSEEVENTF_RIGHTDOWN | MOUSEEVENTF_RIGHTUP, pt.X, pt.Y, 0, 0);
                //AddList("Click CloudFare");
                //MessageBox.Show("Đã click vào nút: X = " + pt.X + " Y = " + pt.Y);

                progressBar1.Value += 1;
            }

            if (alarmCounter == StepTimer03[4]) //focus text search
            {
                //pt.X = 382;
                //pt.Y = 486;
                pt = _TextSearch04;

                Cursor.Position = pt;
                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);

                progressBar1.Value += 1;
            }
            if (alarmCounter == StepTimer03[5]) //input text search
            {
                // SendKeys.Send(txtKeywords.Text);
                SendKeys.Send("fsdfsdfsdf sdfsdfasdfsadf sdfsadf sadf");

                progressBar1.Value += 1;
            }
            if (alarmCounter == StepTimer03[6]) // click search
            {
                //pt.X = 941;
                //pt.Y = 486;
                pt = _ButtonSearch04;

                Cursor.Position = pt;
                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);

                progressBar1.Value += 1;
            }
            //if (alarmCounter == StepTimer03[6]) // click download button
            //{
            //    pt.X = 979;
            //    pt.Y = 936;
            //    Cursor.Position = pt;
            //    mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
            //    AddList("Click Download");

            //    progressBar1.Value += 1;
            //}

            //if (alarmCounter == StepTimer03[7]) // click export to excel
            //{
            //    pt.X = 928;
            //    pt.Y = 744;
            //    Cursor.Position = pt;
            //    mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
            //    AddList("Click to excel File");

            //    progressBar1.Value += 1;
            //}
            if (alarmCounter == StepTimer03[7])
            {
                alarmCounter = 0;

                this.WindowState = FormWindowState.Normal;

                progressBar1.Value += 1;

                tmrPlan04.Stop();
                progressBar1.Value = 0;
                progressBar1.Maximum = max_Process_Plan03;

                alarmCounter = 0;
                tmrPlan03.Start();
                cboPlan.Text = "Kịch bản 03";
            }
        }

        private void txtTotal_TextChanged(object sender, EventArgs e)
        {

        }

        // T.Hoàng code 4:28 ngày 20230302
        public static bool IsNumeric(object value)
        {
            try
            {
                double d = Convert.ToDouble(value);
                return true;
            }
            catch (FormatException)
            {
                return false;
            }
        }

        private void tmrPlanXen(object sender, EventArgs e)
        {
            alarmCounter++;
            if (alarmCounter == StepTimer04[0]) //focus text taikhoan
            {
                pt.X = 641;
                pt.Y = 376;
                Cursor.Position = pt;
                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);

                progressBar1.Value += 1;
            }
            if (alarmCounter == StepTimer04[1]) //input to tai khoan
            {
                SendKeys.Send(txttaikhoan.Text);

                progressBar1.Value += 1;
            }
            if (alarmCounter == StepTimer04[2]) //focus text mat khau
            {
                pt.X = 641;
                pt.Y = 461;
                Cursor.Position = pt;
                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);

                progressBar1.Value += 1;
            }
            if (alarmCounter == StepTimer04[3]) //input text mat khau
            {
                SendKeys.Send(txtmatkhau.Text);

                progressBar1.Value += 1;
            }
            if (alarmCounter == StepTimer04[4]) // click login
            {
                pt.X = 255;
                pt.Y = 564;
                Cursor.Position = pt;
                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);

                progressBar1.Value += 1;
            }
            if (alarmCounter == StepTimer04[5]) //focus text search
            {
                pt.X = 382;
                pt.Y = 486;
                Cursor.Position = pt;
                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);

                progressBar1.Value += 1;
            }
            if (alarmCounter == StepTimer04[6]) //input text search
            {
                SendKeys.Send(txtKeywords.Text);

                progressBar1.Value += 1;
            }
            if (alarmCounter == StepTimer04[7]) // click search
            {
                pt.X = 941;
                pt.Y = 486;
                Cursor.Position = pt;
                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);

                progressBar1.Value += 1;
            }

            if (alarmCounter == StepTimer04[8])
            {
                alarmCounter = 0;

                progressBar1.Value += 1;

                tmrPlan05.Stop();
                progressBar1.Value = 0;
                progressBar1.Maximum = max_Process_Plan03;

                alarmCounter = 0;
                tmrPlan03.Start();
                cboPlan.Text = "Kịch bản 03";
            }
        }


        private void quickExportExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Title = "Export Excel";
            saveFileDialog.Filter = "Excel(*.xlsx)|*.xlsx|Excel 2016(*.xls)|*.xls";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                //try
                
                    QuickExportExcel(saveFileDialog.FileName);

                //catch (Exception ex)
                //{
                //    MessageBox.Show("Xuat file khong thanh cong \n" + ex.Message);
                //}
            }
        }

        private void clearDataGridToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Bạn có chắc chắn muốn xóa Danh sách này?", "Xác nhận xóa", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                dgrListKeywords.Rows.Clear();
                KeyIndex = 0;
            }
            
        }

        private void đọcFileKịchBảnToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LoadPlanPoint("Plan.xlsx");
        }

        private void btnNextKey_Click(object sender, EventArgs e)
        {
            KeyIndex++;
            if (countNextKey == 0 || countBack >=2)  // đếm 2 lần quay lại key ko được, coi như key hỏng và bỏ qua hẳn
            {
                strCheck[KeyIndex - 1]= LevelSearch; //"100"; //Giá trị cao, mặc định là 100 để vượt qua Level
                strVolume[KeyIndex - 1] = "0"; // Đưa Vol về = 0, chưa hiểu ý thầy Bình chỗ này!
                countBack = 0; // đếm lại
            }
            txtKeywords.Text = Convert.ToString(strKeyWords[KeyIndex]);
        }

        private void FocusCurrentCell(DataGridView dataGridView, int curRow)
        {
            if (dataGridView.CurrentRow != null && dataGridView.CurrentCell != null)
            {
                // Lưu trữ vị trí ô cell hiện tại
                int currentRowIndex = curRow;
                int currentColumnIndex = 0;

                // Di chuyển tiêu điểm đến ô cell hiện tại
                dataGridView.CurrentCell = dataGridView.Rows[currentRowIndex].Cells[currentColumnIndex];

                // Tập trung vào DataGridView để ô cell hiện tại trở thành tiêu điểm
                dataGridView.Focus();
            }
        }

        private void ghiFileCàiĐặtToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmConfig f_config = new frmConfig();
            f_config.ShowDialog();
        }

        private void dgrListKeywords_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void xóaKeyĐangChọnToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if(dgrListKeywords.CurrentRow.Index >=0)
            {
                dgrListKeywords.Rows.RemoveAt(dgrListKeywords.CurrentRow.Index);
            }    
        }
        //////////////////////////////////////////////////////////////////////
    }
}




//using GetKeywords.Modules;
//using Microsoft.Office.Interop.Excel;
//using OfficeOpenXml;
//using System;
//using System.Collections;
//using System.Data;
//using System.Diagnostics.Eventing.Reader;
//using System.Drawing;
//using System.IO;
//using System.Runtime.InteropServices;
//using System.Threading.Tasks;
//using System.Windows.Forms;
//using Excel = Microsoft.Office.Interop.Excel;

//namespace GetKeywords
//{
//    public partial class Form1 : Form
//    {

//        static System.Windows.Forms.Timer myTimer = new System.Windows.Forms.Timer();
//        static int alarmCounter = 0;
//        static bool exitFlag = false;
//        static System.Drawing.Point pt;
//        private int[] StepTimer = new int[100];
//        private int[] StepTimer01 = new int[100];
//        private int[] StepTimer02 = new int[100];
//        private int[] StepTimer03 = new int[100];
//        private int[] StepTimer04 = new int[100];
//        private int[] NextStepDelay = new int[100];
//        private int[] NextStepDelay01 = new int[100];
//        private int[] NextStepDelay02 = new int[100];
//        private int[] NextStepDelay03 = new int[100];
//        private int[] NextStepDelay04 = new int[100];


//        static int d_errorfile = 0; // Số lần không lấy được file excel

//        private ExcelConnect f;
//        private string CurrentKeywords = null;

//        private int KeyIndex = 0;
//        private int KeyIndex_backup = 0;
//        private int countNextKey = 0;
//        private int countBack = 0;

//        private int LoadFileExcelOK = 0;

//        private int LevelSearch;

//        private int NextKeyCount = 0;

//        private string ListAccount;
//        private string[] Account = new string[5];
//        private int AccountIndex = 0;
//        private string passLogin;

//        private int max_Process_Plan01;
//        private int max_Process_Plan02;
//        private int max_Process_Plan03; // Số lượng tối đa tiến trình trên thanh trượt kịch bản 3
//        private int max_Process_Plan04; // Số lượng tối đa tiến trình trên thanh trượt kịch bản 4

//        private string[] ListSuggestKeys;
//        private string[] ListNegativeKeys;
//        private char[] Separator = { '|' };

//        //private string[] strKeyWords= new string[2000000];
//        //private string[] strVolume = new string[2000000];
//        //private string[] strCheck = new string[2000000];
//        //private string[] strComp = new string[2000000];

//        private ArrayList strKeyWords = new ArrayList();
//        private ArrayList strVolume = new ArrayList();
//        private ArrayList strCheck = new ArrayList();
//        private ArrayList strComp = new ArrayList();


//        //Khởi tạo các nút kịch bản Login
//        private System.Drawing.Point _AccountButtonLogin;
//        private System.Drawing.Point _LogoutButtonLogin;
//        private System.Drawing.Point _LoginButtonLogin;
//        private System.Drawing.Point _EmailTextLogin;
//        private System.Drawing.Point _PassTextLogin;
//        private System.Drawing.Point _TextSearchLogin;
//        private System.Drawing.Point _ButtonSearchLogin;

//        // Khởi tạo các nút kịch bản 03:
//        private System.Drawing.Point _TextSearch03;
//        private System.Drawing.Point _ButtonSearch03;
//        private System.Drawing.Point _ButtonDownload03;
//        private System.Drawing.Point _ButtonExcel03;

//        // Khởi tạo các nút kịch bản 04:
//        private System.Drawing.Point _TextSearch04;
//        private System.Drawing.Point _ButtonSearch04;
//        private System.Drawing.Point _ClickCloudFare04;

//        private Random rand = new Random();



//        public Form1()
//        {
//            InitializeComponent();
//            // Lấy các dữ liệu setting
//            // Delay Time after event 1
//            NextStepDelay[0] = 1; //focus text search
//            NextStepDelay[1] = 1; //input text search
//            NextStepDelay[2] = 3; // click search
//            NextStepDelay[3] = 1; // click download button
//            NextStepDelay[4] = 1; // click excel
//            NextStepDelay[5] = 2; // click stop

//            StepTimer[0] = 2;
//            for (int i = 0; i < 10; i++)
//            {
//                StepTimer[i + 1] = StepTimer[i] + NextStepDelay[i];
//            }
//            max_Process_Plan01 = 20;

//            // Delay Time after event 2
//            NextStepDelay01[0] = 1; //focus to link
//            NextStepDelay01[1] = 1; //input text link
//            NextStepDelay01[2] = 3; // click enter
//            NextStepDelay01[3] = 2; // click Account Menu
//            NextStepDelay01[4] = 2; // click to Logout
//            NextStepDelay01[5] = 2; // click to Login Again
//            NextStepDelay01[6] = 2; // focus to tai khoan
//            NextStepDelay01[7] = 1; // input text tai khoan
//            NextStepDelay01[8] = 1; // focus to mat khau
//            NextStepDelay01[9] = 1; // input text mat khau
//            NextStepDelay01[10] = 3; // click to login
//            NextStepDelay01[11] = 1; //focus to text keyword
//            NextStepDelay01[12] = 1; //input to text keyword
//            NextStepDelay01[13] = 3; // click search
//            NextStepDelay01[14] = 1; // click download button
//            NextStepDelay01[15] = 1; // click excel
//            NextStepDelay01[16] = 2; // click stop

//            StepTimer01[0] = 2;
//            for (int j = 0; j < 20; j++)
//            {
//                StepTimer01[j + 1] = StepTimer01[j] + NextStepDelay01[j];
//            }
//            max_Process_Plan02 = 20;

//            // Delay Time after event 3
//            NextStepDelay02[0] = 1; // + rand.Next(100); //focus text search
//            NextStepDelay02[1] = 1; //+ rand.Next(100); //CtrlA
//            NextStepDelay02[2] = 1; //+ rand.Next(100); // input text search
//            NextStepDelay02[3] = 4; //+ rand.Next(100); // click search
//            NextStepDelay02[4] = 1; //+ rand.Next(100); // click download button
//            NextStepDelay02[5] = 2; //+ rand.Next(100); // click excel
//            NextStepDelay02[6] = 1; //+ rand.Next(100); // click kiem tra file
//            NextStepDelay02[7] = 1; //+ rand.Next(100); // click stop

//            StepTimer02[0] = 2;
//            for (int m = 0; m < 10; m++)
//            {
//                StepTimer02[m + 1] = StepTimer02[m] + NextStepDelay02[m];
//            }
//            max_Process_Plan03 = 20;

//            // Delay Time after event 4
//            NextStepDelay03[0] = 2; //click to link
//            NextStepDelay03[1] = 2; //input dia chi website
//            NextStepDelay03[2] = 2; //Enter
//            NextStepDelay03[3] = 1; //Click CloudFare
//            NextStepDelay03[4] = 2; // focus to text
//            NextStepDelay03[5] = 2; // input to text
//            NextStepDelay03[6] = 2; // click search
//            NextStepDelay03[7] = 2; // click stop

//            StepTimer03[0] = 2;
//            for (int n = 0; n < 10; n++)
//            {
//                StepTimer03[n + 1] = StepTimer03[n] + NextStepDelay03[n];
//            }

//            max_Process_Plan04 = 20;
//            //event5
//            NextStepDelay04[0] = 2; //focus tai khoan
//            NextStepDelay04[1] = 2; //input tai khoan
//            NextStepDelay04[2] = 3; //focus mat khau
//            NextStepDelay04[3] = 4; //input mat khau
//            NextStepDelay04[4] = 2; // click to login
//            NextStepDelay04[5] = 2; //focus to text keyword
//            NextStepDelay04[6] = 2; //input to text keyword
//            NextStepDelay04[7] = 2; //click to search
//            NextStepDelay04[8] = 1; //click to stop

//            for (int a = 0; a < 10; a++)
//            {
//                StepTimer04[a + 1] = StepTimer04[a] + NextStepDelay04[a];
//            }

//            // Load Toa do Kich ban
//            LoadPlanPoint("Plan.xlsx");
//        }
//        private const int MOUSEEVENTF_LEFTDOWN = 0x02;
//        private const int MOUSEEVENTF_LEFTUP = 0x04;
//        private const int MOUSEEVENTF_RIGHTDOWN = 0x08;
//        private const int MOUSEEVENTF_RIGHTUP = 0x10;

//        [DllImport("user32.dll")]

//        // Định nghĩa hàm mouse_event() từ thư viện user32.dll
//        private static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);

//        // ...
//        private void btnStart_Click(object sender, EventArgs e)
//        {
//            if (btnStart.Text == "Start")
//            {
//                string strSuggestKey = txtSuggestKey.Text;
//                string strNegativeKey = txtNegativeKey.Text;
//                ListSuggestKeys = strSuggestKey.Split(Separator);
//                ListNegativeKeys = strNegativeKey.Split(Separator);

//                ListAccount = txttaikhoan.Text;
//                Account = ListAccount.Split(Separator);
//                passLogin = txtmatkhau.Text;

//                Cursor.Show();  // Cho phép hiện con trỏ chuột lên

//                alarmCounter = 0;
//                exitFlag = false;
//                d_errorfile = 0;
//                // KeyIndex = 0;   //// Lưu ý trường hợp có cần phải tính toàn lại loadkeys Index không 
//                LoadFileExcelOK = 0;
//                //InitVar.v_speed = Convert.ToInt32(txtSpeed.Text);
//                //InitVar.v_VolMax = Convert.ToInt32(txtVolMax.Text);


//                if (cboPlan.SelectedIndex == 0) // Lựa chọn Login
//                {
//                    progressBar1.Maximum = max_Process_Plan02; // số lượng các thao tác trong kế hoạch.
//                    progressBar1.Value = 0;
//                    tmrPlan02.Interval = InitVar.v_speed;
//                    tmrPlan02.Start();
//                    cboPlan.Text = "Kịch bản 02";
//                }
//                if (cboPlan.SelectedIndex == 1) // Lựa chọn Get Keywords
//                {

//                    progressBar1.Maximum = max_Process_Plan01; // số lượng các thao tác trong kế hoạch.
//                    progressBar1.Value = 0;
//                    tmrPlan01.Interval = InitVar.v_speed;
//                    tmrPlan01.Start();
//                    cboPlan.Text = "Kịch bản 01";
//                }
//                if (cboPlan.SelectedIndex == 2) // Lựa chọn Get Keywords tiep theo
//                {

//                    progressBar1.Maximum = max_Process_Plan03; // số lượng các thao tác trong kế hoạch 03.
//                    progressBar1.Value = 0;
//                    tmrPlan03.Interval = InitVar.v_speed;
//                    tmrPlan03.Start();
//                    cboPlan.Text = "Kịch bản 03";
//                }
//                if (cboPlan.SelectedIndex == 3) // không tìm thấy file ex
//                {
//                    progressBar1.Maximum = max_Process_Plan04; // số lượng các thao tác trong kế hoạch 04.
//                    progressBar1.Value = 0;
//                    tmrPlan04.Interval = InitVar.v_speed;
//                    tmrPlan04.Start();
//                    cboPlan.Text = "Kịch bản 04";
//                }
//                if (cboPlan.SelectedIndex == 4) // xen
//                {
//                    progressBar1.Maximum = 9; // số lượng các thao tác trong kế hoạch.
//                    progressBar1.Value = 0;
//                    tmrPlan05.Interval = InitVar.v_speed;
//                    tmrPlan05.Start();
//                    cboPlan.Text = "Kịch bản 05";
//                }
//                btnStart.Text = "Stop";
//            }
//            else
//            {
//                btnStart.Text = "Start";
//                cboPlan.SelectedIndex = 2;  // Lựa chọn sẵn kịch bản 03
//                tmrPlan01.Stop();
//                tmrPlan02.Stop();
//                tmrPlan03.Stop();
//                tmrPlan04.Stop();
//                tmrPlan05.Stop();
//                if (alarmCounter >= StepTimer02[2]) KeyIndex--;
//                //MessageBox.Show("Vui lòng thiết lập lại trạng thái bắt đầu nếu muốn Start");  // Tạm thời bỏ thông báo này để next key
//            }

//        }

//        private void Form1_Load(object sender, EventArgs e)
//        {

//            InitVar.OpenFileConfig(InitVar.pathConfig);

//            // Khởi tạo tạm các keyGG & ChatGPT
//            InitVar.v_arrKeyGG = new string[100];
//            InitVar.v_arrKeyGG[0] = "AIzaSyCSXStlfJDlEoikXv6P7yEOSRb2PsVjZAM";
//            InitVar.v_arrKeyGG[1] = "AIzaSyBA5C5BzTPq1Ooi4x7rAytNtqTjjdGJzK8";
//            InitVar.v_arrKeyGG[2] = "AIzaSyB-sU_otHwxn2JNwIqI42O0gHLEk-mkZtY";
//            InitVar.v_arrKeyGG[3] = "AIzaSyCUZBpGQNUs1AJJdH8lSsjSUv2dxmN1zWI";
//            InitVar.v_arrKeyGG[4] = "AIzaSyAk2bQ23muiPHYXf2yFN7GlRic3vpXFh4Y";
//            InitVar.v_arrKeyGG[5] = "AIzaSyCQbkcrYYQTLz_IF-SDmFbVefjHHxyTNM8";
//            InitVar.v_arrKeyGG[6] = "AIzaSyClUiZTcwp359Kbb-W7WDmGziVCjtWA37M";
//            InitVar.v_arrKeyGG[7] = "AIzaSyDft1aB-jQ-Kpk9tE_HjTQm9mHvwlgNaWk";
//            InitVar.v_arrKeyGG[8] = "AIzaSyDdNIKC3t7J9elX3QXUSFD7ELgr3I4UPzY";
//            InitVar.v_arrKeyGG[9] = "AIzaSyD2rfceyTgDh1QzTH-uEJmOlyk-goIEW38";

//            InitVar.v_arrKeyChatGPT = new string[100];
//            InitVar.v_arrKeyChatGPT[0] = "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";
//            InitVar.v_arrKeyChatGPT[1] = "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";
//            InitVar.v_arrKeyChatGPT[2] = "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";
//            InitVar.v_arrKeyChatGPT[3] = "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";
//            InitVar.v_arrKeyChatGPT[4] = "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";
//            InitVar.v_arrKeyChatGPT[5] = "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";
//            InitVar.v_arrKeyChatGPT[6] = "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";
//            InitVar.v_arrKeyChatGPT[7] = "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";
//            InitVar.v_arrKeyChatGPT[8] = "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";
//            InitVar.v_arrKeyChatGPT[9] = "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";


//            // Lấy các dữ liệu setting
//            //InitVar.v_speed = Convert.ToInt32(txtSpeed.Text);
//            //InitVar.v_VolMax = Convert.ToInt32(txtVolMax.Text); // T.Hoàng thêm 15:54 20230302



//            // Thêm kịch bản
//            cboPlan.Items.Clear();
//            cboPlan.Items.Add("Login");
//            cboPlan.Items.Add("Get keywords");
//            cboPlan.Items.Add("Dowload Keyword tiep theo");
//            cboPlan.Items.Add("Dowload lại nếu không tìm thấy file excel");

//            cboPlan.SelectedIndex = 2;

//            // Mở kết nối file excel
//            //f.fileName = "Keyword Tool Export -Keyword Suggestions - " + CurrentKeywords;

//            foreach (DataGridViewColumn column in dgrListKeywords.Columns)
//            {
//                column.SortMode = DataGridViewColumnSortMode.NotSortable;
//            }
//        }

//        /// Đoạn code save file config
//        /// 


//        /// <summary>
//        /// Day la doan nhap file excel thu 1
//        /// </summary>
//        /// <param name="path"></param>
//        private void Importexcel(string path)
//        {
//            using (FileStream fileStream = new FileStream(path, FileMode.Open, FileAccess.Read))
//            {
//                using (ExcelPackage excelPackage = new ExcelPackage(fileStream))
//                {
//                    // Lấy Sheet đầu tiên trong tệp Excel
//                    ExcelWorksheet excelWorksheet = excelPackage.Workbook.Worksheets[0];

//                    int i = 1;
//                    int index = 0;
//                    while ((excelWorksheet.Cells[i + 1, 1].Value != null) && (excelWorksheet.Cells[i + 1, 2].Value != null) && (excelWorksheet.Cells[i + 1, 2].Value.ToString().Contains(".") == false))
//                    {
//                        // Cột 17 là vị trí của cột Competition;......
//                        if (excelWorksheet.Cells[i + 1, 3].Value == null)
//                        {
//                            //dgrListKeywords.Rows.Add(excelWorksheet.Cells[i + 1, 1].Value, excelWorksheet.Cells[i + 1, 2].Value, "0", excelWorksheet.Cells[i + 1, 17].Value);
//                            strKeyWords.Add(Convert.ToString(excelWorksheet.Cells[i + 1, 1].Value));
//                            strVolume.Add(Convert.ToString(excelWorksheet.Cells[i + 1, 2].Value));
//                            strCheck.Add("0");
//                            strComp.Add(Convert.ToString(excelWorksheet.Cells[i + 1, 17].Value));
//                        }
//                        else
//                        {
//                            //dgrListKeywords.Rows.Add(excelWorksheet.Cells[i + 1, 1].Value, excelWorksheet.Cells[i + 1, 2].Value, excelWorksheet.Cells[i + 1, 3].Value, excelWorksheet.Cells[i + 1, 17].Value);
//                            strKeyWords.Add(Convert.ToString(excelWorksheet.Cells[i + 1, 1].Value));
//                            strVolume.Add(Convert.ToString(excelWorksheet.Cells[i + 1, 2].Value));
//                            strCheck.Add(Convert.ToString(excelWorksheet.Cells[i + 1, 3].Value));
//                            strComp.Add(Convert.ToString(excelWorksheet.Cells[i + 1, 17].Value));
//                        }
//                        i++; index++;
//                    }

//                    if (excelWorksheet.Cells[1, 27].Value != null)  //
//                    {
//                        KeyIndex = Convert.ToInt32(excelWorksheet.Cells[1, 27].Value);
//                        txtCur.Text = Convert.ToString(KeyIndex);
//                    }

//                    if (excelWorksheet.Cells[2, 27].Value != null)  //
//                    {
//                        txtSuggestKey.Text = Convert.ToString(excelWorksheet.Cells[2, 27].Value);
//                    }
//                    if (excelWorksheet.Cells[3, 27].Value != null)  //
//                    {
//                        txtNegativeKey.Text = Convert.ToString(excelWorksheet.Cells[3, 27].Value);
//                    }
//                    if (excelWorksheet.Cells[4, 27].Value != null)  //
//                    {
//                        txttaikhoan.Text = Convert.ToString(excelWorksheet.Cells[4, 27].Value);
//                    }
//                    if (excelWorksheet.Cells[5, 27].Value != null)  //
//                    {
//                        txtmatkhau.Text = Convert.ToString(excelWorksheet.Cells[5, 27].Value);
//                    }

//                }
//            }


//            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//            //using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(path)))
//            //{
//            //    ExcelWorksheet excelWorksheet = excelPackage.Workbook.Worksheets[0];
//            //    int i = 1;
//            //    while ((excelWorksheet.Cells[i + 1, 1].Value != null) && (excelWorksheet.Cells[i + 1, 2].Value != null) && (excelWorksheet.Cells[i + 1, 2].Value.ToString().Contains(".") == false))
//            //    {
//            //        // Cột 17 là vị trí của cột Competition;......
//            //        if (excelWorksheet.Cells[i + 1, 3].Value == null)
//            //        {
//            //            dgrListKeywords.Rows.Add(excelWorksheet.Cells[i + 1, 1].Value, excelWorksheet.Cells[i + 1, 2].Value, "0", excelWorksheet.Cells[i + 1, 17].Value);
//            //        }
//            //        else
//            //        {
//            //            dgrListKeywords.Rows.Add(excelWorksheet.Cells[i + 1, 1].Value, excelWorksheet.Cells[i + 1, 2].Value, excelWorksheet.Cells[i + 1, 3].Value, excelWorksheet.Cells[i + 1, 17].Value);
//            //        }
//            //        i++;
//            //    }

//            //    if (excelWorksheet.Cells[1, 27].Value != null)  //
//            //    {
//            //        KeyIndex = Convert.ToInt32(excelWorksheet.Cells[1, 27].Value);
//            //        txtCur.Text = Convert.ToString(KeyIndex);
//            //    }

//            //    if (excelWorksheet.Cells[2, 27].Value != null)  //
//            //    {
//            //        txtSuggestKey.Text = Convert.ToString( excelWorksheet.Cells[2, 27].Value);
//            //    }
//            //    if (excelWorksheet.Cells[3, 27].Value != null)  //
//            //    {
//            //        txtNegativeKey.Text = Convert.ToString(excelWorksheet.Cells[3, 27].Value);
//            //    }
//            //    if (excelWorksheet.Cells[4, 27].Value != null)  //
//            //    {
//            //        txttaikhoan.Text = Convert.ToString(excelWorksheet.Cells[4, 27].Value);
//            //    }
//            //    if (excelWorksheet.Cells[5, 27].Value != null)  //
//            //    {
//            //        txtmatkhau.Text = Convert.ToString(excelWorksheet.Cells[5, 27].Value);
//            //    }



//            //    excelPackage.Dispose();
//            //}
//            txtTotal.Text = Convert.ToString(strKeyWords.Count);  // Convert.ToString(dgrListKeywords.Rows.Count);
//        }
//        /// <summary>
//        /// doan nhap file cho vong lay keywords tiep theo
//        /// </summary>
//        /// <param name="sender"></param>
//        /// <param name="e"></param>
//        /// 
//        private async void ImportExcelCircle(string path)
//        {
//            //int kq = 1;
//            LoadFileExcelOK = 1;
//            using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(path)))
//            {
//                ExcelWorksheet excelWorksheet = excelPackage.Workbook.Worksheets[0];
//                int i = 1;
//                while (excelWorksheet.Cells[i + 1, 1].Value != null) //&& (excelWorksheet.Cells[i + 1, 2].Value!= null)) // && Convert.ToInt32(excelWorksheet.Cells[i + 1, 2].Value) > 1000))
//                {

//                    // THoang code 21:46 20230301
//                    if ((excelWorksheet.Cells[i + 1, 2].Value != null) && (Convert.ToInt32(excelWorksheet.Cells[i + 1, 2].Value) >= InitVar.v_VolMin) && (Convert.ToInt32(excelWorksheet.Cells[i + 1, 17].Value) <= InitVar.v_LevelDif))
//                    {
//                        //kq = 0;
//                        LoadFileExcelOK = 0;
//                        string str2 = excelWorksheet.Cells[i + 1, 1].Value.ToString();
//                        //if (Convert.ToInt32(excelWorksheet.Cells[i + 1, 2].Value) >= v_VolMax)
//                        //{
//                        // Kiem tra trung lap trong danh sach
//                        bool dup = false;
//                        for (int j = 0; j <= dgrListKeywords.Rows.Count - 1; j++)
//                        {

//                            string str1 = strKeyWords[j]; //dgrListKeywords.Rows[j].Cells[0].Value.ToString();

//                            if (str1.Equals(str2, StringComparison.InvariantCultureIgnoreCase))
//                            {
//                                dup = true;
//                                CurrentKeywords = txtKeywords.Text;
//                                if (CurrentKeywords.Equals(str2, StringComparison.InvariantCultureIgnoreCase))
//                                {
//                                    dgrListKeywords.Rows[KeyIndex - 1].Cells[1].Value = excelWorksheet.Cells[i + 1, 2].Value;

//                                }
//                                break;
//                            }

//                        }
//                        if (dup == false)
//                        {
//                            // Kiểm tra ListSuggest & Negative
//                            bool sug = false;

//                            for (int t1 = 0; t1 <= ListSuggestKeys.Length - 1; t1++)
//                            {
//                                int indexSub = str2.IndexOf(ListSuggestKeys[t1]);
//                                //if (str2.Contains(ListSuggestKeys[t1]) == true)
//                                if (indexSub >= 0)
//                                {
//                                    sug = true;
//                                    break;
//                                }
//                            }

//                            bool nega = true;

//                            for (int t2 = 0; t2 <= ListNegativeKeys.Length - 1; t2++)
//                            {
//                                if (ListNegativeKeys[t2] != "")
//                                {
//                                    int indexSub = str2.IndexOf(ListNegativeKeys[t2]);
//                                    //if(str2.Contains(ListNegativeKeys[t2]) == true)
//                                    if (indexSub >= 0)
//                                    {
//                                        nega = false;
//                                        break;
//                                    }
//                                }
//                            }




//                            //for (int t1 = 0; t1 <= ListSuggestKeys.Length-1; t1++)


//                            //for (int t2 = 0; t2 <= ListNegativeKeys.Length - 1; t2++)
//                            //{
//                            //    if (ListNegativeKeys[t2] == "")
//                            //    {
//                            //        if (str2.Contains(ListSuggestKeys[t1]) == true)
//                            //        {
//                            //            sug = true;
//                            //            break;
//                            //        }
//                            //    }
//                            //    else
//                            //    {
//                            //        if ((str2.Contains(ListSuggestKeys[t1]) == true) && str2.Contains(ListNegativeKeys[t2]) == false)
//                            //        {
//                            //            sug = true;
//                            //            break;
//                            //        }    
//                            //    }    
//                            //}

//                            if ((sug == true) && (nega == true))
//                            {
//                                //string InputRequest = "";
//                                //string OutputContent = await clsAPI.CallChatGPTAPI(InputRequest); // Gọi hàm từ clsAPI.


//                                DataGridViewRow newRow = new DataGridViewRow();
//                                newRow.Cells.Add(new DataGridViewTextBoxCell { Value = excelWorksheet.Cells[i + 1, 1].Value });
//                                newRow.Cells.Add(new DataGridViewTextBoxCell { Value = excelWorksheet.Cells[i + 1, 2].Value });
//                                newRow.Cells.Add(new DataGridViewTextBoxCell { Value = (LevelSearch + 1) });
//                                newRow.Cells.Add(new DataGridViewTextBoxCell { Value = excelWorksheet.Cells[i + 1, 17].Value });

//                                // Thay đổi cơ chế chèn file
//                                dgrListKeywords.Rows.Insert(KeyIndex, newRow);
//                                // dgrListKeywords.Rows.Add(excelWorksheet.Cells[i + 1, 1].Value, excelWorksheet.Cells[i + 1, 2].Value);
//                            }


//                        }

//                        //}

//                    }


//                    i++;
//                }
//                File.Delete(path); // THoang: Xóa luôn file sau khi đã nạp
//                //return kq;
//            }
//        }

//        /// <summary>
//        /// Load dữ liệu tọa độ điểm Kịch bản
//        /// </summary>
//        /// <param name="path"></param>
//        /// <returns></returns>
//        private void LoadPlanPoint(string path)
//        {
//            try
//            {
//                using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(path)))
//                {
//                    // Load kich ban Login
//                    ExcelWorksheet excelWorksheetPlanLogin = excelPackage.Workbook.Worksheets["PlanLogin"];

//                    _AccountButtonLogin.X = Convert.ToInt32(excelWorksheetPlanLogin.Cells[2, 2].Value);
//                    _AccountButtonLogin.Y = Convert.ToInt32(excelWorksheetPlanLogin.Cells[2, 3].Value);

//                    _LogoutButtonLogin.X = Convert.ToInt32(excelWorksheetPlanLogin.Cells[3, 2].Value);
//                    _LogoutButtonLogin.Y = Convert.ToInt32(excelWorksheetPlanLogin.Cells[3, 3].Value);

//                    _EmailTextLogin.X = Convert.ToInt32(excelWorksheetPlanLogin.Cells[4, 2].Value);
//                    _EmailTextLogin.Y = Convert.ToInt32(excelWorksheetPlanLogin.Cells[4, 3].Value);

//                    _PassTextLogin.X = Convert.ToInt32(excelWorksheetPlanLogin.Cells[5, 2].Value);
//                    _PassTextLogin.Y = Convert.ToInt32(excelWorksheetPlanLogin.Cells[5, 3].Value);

//                    _LoginButtonLogin.X = Convert.ToInt32(excelWorksheetPlanLogin.Cells[6, 2].Value);
//                    _LoginButtonLogin.Y = Convert.ToInt32(excelWorksheetPlanLogin.Cells[6, 3].Value);

//                    _TextSearchLogin.X = Convert.ToInt32(excelWorksheetPlanLogin.Cells[7, 2].Value);
//                    _TextSearchLogin.Y = Convert.ToInt32(excelWorksheetPlanLogin.Cells[7, 3].Value);

//                    _ButtonSearchLogin.X = Convert.ToInt32(excelWorksheetPlanLogin.Cells[8, 2].Value);
//                    _ButtonSearchLogin.Y = Convert.ToInt32(excelWorksheetPlanLogin.Cells[8, 3].Value);

//                    /////////////////////////////////////////////////
//                    //

//                    // Load kich ban 03
//                    ExcelWorksheet excelWorksheetPlan03 = excelPackage.Workbook.Worksheets["Plan03"];

//                    _TextSearch03.X = Convert.ToInt32(excelWorksheetPlan03.Cells[2, 2].Value);
//                    _TextSearch03.Y = Convert.ToInt32(excelWorksheetPlan03.Cells[2, 3].Value);

//                    _ButtonSearch03.X = Convert.ToInt32(excelWorksheetPlan03.Cells[3, 2].Value);
//                    _ButtonSearch03.Y = Convert.ToInt32(excelWorksheetPlan03.Cells[3, 3].Value);

//                    _ButtonDownload03.X = Convert.ToInt32(excelWorksheetPlan03.Cells[4, 2].Value);
//                    _ButtonDownload03.Y = Convert.ToInt32(excelWorksheetPlan03.Cells[4, 3].Value);

//                    _ButtonExcel03.X = Convert.ToInt32(excelWorksheetPlan03.Cells[5, 2].Value);
//                    _ButtonExcel03.Y = Convert.ToInt32(excelWorksheetPlan03.Cells[5, 3].Value);
//                    /////////////////////////////////////////////////


//                    // Load kich ban 04
//                    ExcelWorksheet excelWorksheetPlan04 = excelPackage.Workbook.Worksheets["Plan04"];

//                    _TextSearch04.X = Convert.ToInt32(excelWorksheetPlan04.Cells[2, 2].Value);
//                    _TextSearch04.Y = Convert.ToInt32(excelWorksheetPlan04.Cells[2, 3].Value);

//                    _ButtonSearch04.X = Convert.ToInt32(excelWorksheetPlan04.Cells[3, 2].Value);
//                    _ButtonSearch04.Y = Convert.ToInt32(excelWorksheetPlan04.Cells[3, 3].Value);

//                    _ClickCloudFare04.X = Convert.ToInt32(excelWorksheetPlan04.Cells[4, 2].Value);
//                    _ClickCloudFare04.Y = Convert.ToInt32(excelWorksheetPlan04.Cells[4, 3].Value);

//                    /////////////////////////////////////////////////

//                    excelPackage.Dispose();
//                }
//                //txtTotal.Text = Convert.ToString(dgrListKeywords.Rows.Count);
//            }
//            catch (Exception ex)
//            {
//                MessageBox.Show("Error load PlanPoint: " + ex.Message);
//            }
//        }

//        private void openExcelToolStripMenuItem_Click(object sender, EventArgs e)
//        {
//            OpenFileDialog openFileDialog = new OpenFileDialog();
//            openFileDialog.Title = "Import Excel";
//            openFileDialog.Filter = "Excel(*.xlsx)|*.xlsx|Excel 2016(*.xls)|*.xls";
//            if (openFileDialog.ShowDialog() == DialogResult.OK)
//            {
//                try
//                {
//                    Importexcel(openFileDialog.FileName);
//                    MessageBox.Show("Nhap file thanh cong");
//                }
//                catch (Exception ex)
//                {
//                    MessageBox.Show("Nhap file khong thanh cong \n" + ex.Message);
//                }
//            }

//            if (dgrListKeywords.RowCount > 0)
//            {
//                txtKeywords.Text = Convert.ToString(dgrListKeywords.Rows[KeyIndex].Cells[0].Value);
//            }
//        }
//        private void ExportExcel(string path)
//        {
//            Excel.Application application = new Excel.Application();
//            application.Application.Workbooks.Add(Type.Missing);
//            for (int i = 0; i < dgrListKeywords.Columns.Count; i++)
//            {
//                application.Cells[1, i + 1] = dgrListKeywords.Columns[i].HeaderText;
//            }
//            for (int i = 0; i <= dgrListKeywords.Rows.Count - 1; i++)
//            {
//                for (int j = 0; j <= dgrListKeywords.Columns.Count - 1; j++)
//                {
//                    application.Cells[i + 2, j + 1] = dgrListKeywords.Rows[i].Cells[j].Value;
//                    //application.Cells[i + 2, j + 1] = dgrListKeywords.Rows[i].Cells[j].Value;
//                }
//            }
//            application.Columns.AutoFit();
//            application.ActiveWorkbook.SaveCopyAs(path);
//            application.ActiveWorkbook.Saved = true;
//            application.Quit();
//        }
//        private void QuickExportExcel(string path, int noMsg = 0)
//        {
//            //DataTable dt = new DataTable();

//            if (dgrListKeywords.RowCount >= 1)
//            {

//                // Tạo một đối tượng ExcelPackage
//                using (ExcelPackage excelPackage = new ExcelPackage())
//                {
//                    // Tạo một đối tượng Worksheet
//                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");

//                    // Lấy dữ liệu từ DataGrid và đổ vào worksheet
//                    for (int i = 0; i < dgrListKeywords.Rows.Count; i++)
//                    {
//                        for (int j = 0; j < dgrListKeywords.Columns.Count; j++)
//                        {
//                            worksheet.Cells[i + 1, j + 1].Value = dgrListKeywords.Rows[i].Cells[j].Value;
//                        }
//                    }
//                    worksheet.Cells[1, 27].Value = KeyIndex;  // Ghi lại KeyIndex để tiếp tục xử lý tại ô "AA:1"
//                    worksheet.Cells[2, 27].Value = txtSuggestKey.Text;
//                    worksheet.Cells[3, 27].Value = txtNegativeKey.Text;
//                    worksheet.Cells[4, 27].Value = txttaikhoan.Text;
//                    worksheet.Cells[5, 27].Value = txtmatkhau.Text;

//                    // Lưu workbook vào một MemoryStream
//                    using (MemoryStream memoryStream = new MemoryStream())
//                    {
//                        excelPackage.SaveAs(memoryStream);
//                        memoryStream.Position = 0;

//                        // Lưu MemoryStream vào file Excel
//                        using (FileStream fileStream = new FileStream(path, FileMode.Create, FileAccess.Write))
//                        {
//                            memoryStream.WriteTo(fileStream);
//                        }
//                    }
//                }

//                if (noMsg == 0) MessageBox.Show("Xuat file thanh cong");
//            }
//            else
//            {
//                if (noMsg == 0) MessageBox.Show("There is NO keywords to Export");
//            }
//        }
//        private void saveExcelToolStripMenuItem_Click(object sender, EventArgs e)
//        {
//            DialogResult dlr = MessageBox.Show("Khyến nghị! Bạn nên dùng chức năng Xuất nhanh Excel." + Environment.NewLine + "Bạn có chắc chắn muốn tiếp tục với Xuất Excel bình thường không?", "Viện Tin học Xây dựng", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
//            if (dlr == DialogResult.Yes)
//            {
//                SaveFileDialog saveFileDialog = new SaveFileDialog();
//                saveFileDialog.Title = "Export Excel";
//                saveFileDialog.Filter = "Excel(*.xlsx)|*.xlsx|Excel 2016(*.xls)|*.xls";
//                if (saveFileDialog.ShowDialog() == DialogResult.OK)
//                {
//                    //try
//                    {
//                        ExportExcel(saveFileDialog.FileName);
//                        MessageBox.Show("Xuat file thanh cong");
//                    }
//                    //catch (Exception ex)
//                    //{
//                    //    MessageBox.Show("Xuat file khong thanh cong \n" + ex.Message);
//                    //}
//                }
//            }
//        }
//        private void btnPause_Click(object sender, EventArgs e)
//        {
//            if (btnPause.Text == "Pause")
//            {
//                btnPause.Text = "Continue";
//                tmrPlan03.Stop();
//            }
//            else
//            {
//                btnPause.Text = "Pause";
//                tmrPlan03.Start();
//            }
//        }

//        private void tmrPlanGetKeyword(object sender, EventArgs e)
//        {

//            alarmCounter++;

//            if (alarmCounter == StepTimer[0]) //focus text search
//            {
//                pt.X = 382;
//                pt.Y = 486;
//                Cursor.Position = pt;
//                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);


//                progressBar1.Value += 1;
//            }
//            if (alarmCounter == StepTimer[1]) //input text search
//            {
//                SendKeys.Send(txtKeywords.Text);


//                progressBar1.Value += 1;
//            }

//            if (alarmCounter == StepTimer[2]) // click search
//            {
//                pt.X = 941;
//                pt.Y = 486;
//                Cursor.Position = pt;
//                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);


//                progressBar1.Value += 1;
//            }

//            if (alarmCounter == StepTimer[3]) // click download button
//            {
//                pt.X = 971;
//                pt.Y = 937;
//                Cursor.Position = pt;
//                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);


//                progressBar1.Value += 1;
//            }

//            if (alarmCounter == StepTimer[4]) // click export to excel
//            {
//                pt.X = 918;
//                pt.Y = 751;
//                Cursor.Position = pt;
//                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);


//                progressBar1.Value += 1;
//            }
//            if (alarmCounter == StepTimer[5])
//            {
//                tmrPlan01.Stop();
//                alarmCounter = 0;

//                this.WindowState = FormWindowState.Normal;


//                progressBar1.Value += 1;
//            }
//        }

//        private void tmrPlanLogin(object sender, EventArgs e)

//        {
//            alarmCounter++;
//            if (alarmCounter == StepTimer01[0]) //CLICK TO link
//            {
//                pt.X = 610;
//                pt.Y = 60;
//                Cursor.Position = pt;
//                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);


//                progressBar1.Value += 1;
//            }
//            if (alarmCounter == StepTimer01[1]) //input text dia chi
//            {
//                SendKeys.Send(txtdiachi.Text);


//                progressBar1.Value += 1;
//            }
//            if (alarmCounter == StepTimer01[2]) //input enter
//            {
//                //SendKeys.SendWait("+(CTRL)");
//                //SendKeys.SendWait("+(A)");
//                SendKeys.Send("{ENTER}");


//                progressBar1.Value += 1;
//            }
//            if (alarmCounter == StepTimer01[3]) //CLICK TO Account Button ở góc
//            {
//                pt.X = _AccountButtonLogin.X;
//                pt.Y = _AccountButtonLogin.Y;
//                Cursor.Position = pt;
//                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);


//                progressBar1.Value += 1;
//            }
//            if (alarmCounter == StepTimer01[4]) //CLICK TO Logout ở cuối Menu
//            {
//                pt.X = _LogoutButtonLogin.X;
//                pt.Y = _LogoutButtonLogin.Y;
//                Cursor.Position = pt;
//                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);


//                progressBar1.Value += 1;
//            }
//            if (alarmCounter == StepTimer01[5]) //Click lại nút Login 1 lần nữa ở góc
//            {
//                pt.X = _AccountButtonLogin.X;
//                pt.Y = _AccountButtonLogin.Y;
//                Cursor.Position = pt;
//                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
//                progressBar1.Value += 1;
//            }

//            if (alarmCounter == StepTimer01[6]) //Click Vào ô text nhập Email
//            {
//                pt.X = _EmailTextLogin.X;
//                pt.Y = _EmailTextLogin.Y;
//                Cursor.Position = pt;
//                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
//                progressBar1.Value += 1;
//            }

//            if (alarmCounter == StepTimer01[7]) //input to tai khoan
//            {
//                AccountIndex++;
//                if (AccountIndex >= Account.Length - 1) AccountIndex = 0;
//                SendKeys.Send(Account[AccountIndex]);

//                progressBar1.Value += 1;
//            }
//            if (alarmCounter == StepTimer01[8]) //focus text mat khau
//            {
//                pt.X = _PassTextLogin.X;
//                pt.Y = _PassTextLogin.Y;
//                Cursor.Position = pt;
//                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
//                progressBar1.Value += 1;
//            }
//            if (alarmCounter == StepTimer01[9]) //input text mat khau
//            {
//                SendKeys.Send(passLogin);
//                progressBar1.Value += 1;
//            }
//            if (alarmCounter == StepTimer01[10]) // click login
//            {
//                //pt.X = _LogoutButtonLogin.X;
//                //pt.Y = _LogoutButtonLogin.Y;
//                //Cursor.Position = pt;
//                //mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
//                SendKeys.Send("{ENTER}");
//                progressBar1.Value += 1;
//            }
//            if (alarmCounter == StepTimer01[11]) //focus text search
//            {
//                pt.X = 382;
//                pt.Y = 486;
//                Cursor.Position = pt;
//                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);

//                progressBar1.Value += 1;
//            }
//            if (alarmCounter == StepTimer01[12]) //input text search
//            {
//                SendKeys.Send(txtKeywords.Text);

//                progressBar1.Value += 1;
//            }
//            if (alarmCounter == StepTimer01[13]) // click search
//            {
//                pt.X = _ButtonSearch04.X;
//                pt.Y = _ButtonSearch04.Y;
//                Cursor.Position = pt;
//                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);

//                progressBar1.Value += 1;
//            }

//            //if (alarmCounter == StepTimer01[13]) // click download button
//            //{
//            //    pt.X = _;
//            //    pt.Y = 937;
//            //    Cursor.Position = pt;
//            //    mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);

//            //    progressBar1.Value += 1;
//            //}

//            //if (alarmCounter == StepTimer01[14]) // click export to excel
//            //{
//            //    pt.X = 918;
//            //    pt.Y = 751;
//            //    Cursor.Position = pt;
//            //    mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);

//            //    progressBar1.Value += 1;
//            //}
//            if (alarmCounter == StepTimer01[14])
//            {
//                tmrPlan02.Stop();
//                alarmCounter = 0;
//                //this.WindowState = FormWindowState.Normal;
//                progressBar1.Value = 0;
//                tmrPlan03.Start();
//            }
//        }

//        private void tmrPlanDowloadKeywordtieptheo(object sender, EventArgs e)
//        {
//            //try
//            {


//                alarmCounter++;

//                if (alarmCounter == StepTimer02[0]) //focus text search
//                {
//                    //pt.X = 419;
//                    //pt.Y = 238;
//                    pt = _TextSearch03;


//                    Cursor.Position = pt;
//                    mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);

//                    progressBar1.Value += 1;
//                }
//                if (alarmCounter == StepTimer02[1]) //input text search
//                {
//                    //SendKeys.SendWait("+(CTRL)");
//                    //SendKeys.SendWait("+(A)");
//                    SendKeys.Send("^(a)");


//                    progressBar1.Value += 1;
//                }
//                if (alarmCounter == StepTimer02[2]) //input text search
//                {
//                    if (dgrListKeywords.Rows.Count > 0)
//                    {
//                        do
//                        {
//                            KeyIndex++;
//                            if (KeyIndex % InitVar.v_AutoSave == 0) //Sau 5000 keywords sẽ lưu tạm lại
//                            {
//                                string strTempPath = "";
//                                DateTime now = DateTime.Now;
//                                string formattedDateTime = now.ToString("dd/MM/yyyy HH:mm:ss");
//                                string downloadPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\Downloads\";
//                                strTempPath = downloadPath + "KeyTool" + formattedDateTime.Replace("/", "").Replace(":", "") + ".xlsx";
//                                QuickExportExcel(strTempPath, 1);
//                            }

//                            //THoang 21:56 20230303
//                            if (Convert.ToInt32(dgrListKeywords.Rows[KeyIndex - 1].Cells[1].Value) <= InitVar.v_VolMax)
//                            {
//                                dgrListKeywords.Rows[KeyIndex - 1].Cells[2].Value = "-1";
//                            }
//                            /////////
//                            // Kiem tra den khi Vol het bang Grid thi coi nhu la da xong
//                            if (KeyIndex >= dgrListKeywords.Rows.Count)  // Bỏ qua keywords cuối cùng, thinking tiếp nếu list có vol tận cuối cùng
//                            {
//                                // THoang 22:59 20230302
//                                tmrPlan03.Stop();
//                                alarmCounter = 0;
//                                MessageBox.Show("Hoàn thành chiến dịch. Vui lòng Export File"); //Sau chuyển vào Label Trạng thái                        
//                                break;
//                            }
//                            else
//                            {
//                                txtKeywords.Text = dgrListKeywords.Rows[KeyIndex - 1].Cells[0].Value.ToString();
//                                LevelSearch = Convert.ToInt32(dgrListKeywords.Rows[KeyIndex - 1].Cells[2].Value.ToString());
//                                // THoang 18:59 20230301
//                                FocusCurrentCell(dgrListKeywords, KeyIndex - 1);
//                                dgrListKeywords.Rows[KeyIndex - 1].Selected = true;
//                                txtCur.Text = Convert.ToString(KeyIndex);

//                            }

//                            //THoang 21:56 20230303
//                        } while ((Convert.ToInt32(dgrListKeywords.Rows[KeyIndex - 1].Cells[1].Value) <= InitVar.v_VolMax) || (Convert.ToInt32(dgrListKeywords.Rows[KeyIndex - 1].Cells[2].Value) >= InitVar.v_LevelSearch) || (Convert.ToInt32(dgrListKeywords.Rows[KeyIndex - 1].Cells[3].Value) >= InitVar.v_LevelDif)); // Chi chay các keyword có vol >=1000 || chưa đánh dấu 100

//                        //////
//                    }
//                    else
//                    {
//                        MessageBox.Show("Import file excel và Start lại");
//                        tmrPlan03.Stop();
//                        progressBar1.Value = 0;
//                        alarmCounter = 0;
//                    }

//                    try
//                    {
//                        string sendString;
//                        //sendString = "kết+quả+seagame+31+bóng+đá+nam";
//                        //sendString = sendString.Replace("+", "{+}");

//                        sendString = txtKeywords.Text.Replace("+", "{+}").Replace("^", "{^}").Replace("~", "{~}").Replace("%", "{%}"); //.Replace("(", "{(}").Replace(")", "{)}").Replace("{", "{{}").Replace("}", "{}}").Replace("[", "{[}").Replace("]", "{]}"); //Xử lý ký tự đặc biệt.
//                                                                                                                                       //if (txtKeywords.Text.Contains("^"))
//                        SendKeys.Send(sendString);

//                        progressBar1.Value += 1;
//                    }
//                    catch (Exception ex)
//                    {
//                        MessageBox.Show(ex.Message + " Chụp ảnh màn hình lỗi này. STOP, và có thể cân nhắc tự Nextkey để tiếp tục");
//                    }
//                }

//                if (alarmCounter == StepTimer02[3]) // click search
//                {
//                    //pt.X = 982;
//                    //pt.Y = 238;
//                    pt = _ButtonSearch03;

//                    Cursor.Position = pt;
//                    mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);

//                    progressBar1.Value += 1;
//                }

//                if (alarmCounter == StepTimer02[4]) // click download button
//                {
//                    //pt.X = 971;
//                    //pt.Y = 937;
//                    pt = _ButtonDownload03;

//                    Cursor.Position = pt;

//                    mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);

//                    progressBar1.Value += 1;
//                }

//                if (alarmCounter == StepTimer02[5]) // click export to excel
//                {
//                    //pt.X = 918;
//                    //pt.Y = 751;
//                    pt = _ButtonExcel03;

//                    Cursor.Position = pt;
//                    mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);

//                    progressBar1.Value += 1;
//                }
//                if (alarmCounter == StepTimer02[6])
//                {
//                    // Kiem tra file excel co ton tai khong
//                    // - Co ton tai: Import file excel them vao grid
//                    // - Khong ton tai: ????

//                    //tmrPlan03.Stop()
//                    string fileName = txtKeywords.Text.Replace(".", " ").Replace("/", " ").Replace(":", " ").Replace("!", " ").Replace("&", "&amp").Replace("'", "&#039").Replace("$", " ");
//                    //fileName = fileName.Replace("+","{+}");
//                    string downloadPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\Downloads\";
//                    string filePath = downloadPath + "Keyword Tool Export - Keyword Suggestions - " + fileName + ".xlsx";
//                    if ((System.IO.File.Exists(filePath) == true))
//                    {

//                        ImportExcelCircle(filePath);
//                        // THoang 18:59 20230303
//                        if (LoadFileExcelOK == 0)
//                        {
//                            txtTotal.Text = Convert.ToString(dgrListKeywords.Rows.Count);
//                            dgrListKeywords.Rows[KeyIndex - 1].Cells[2].Value = LevelSearch; //"100"; // Giá trị cao vượt qua LevelSearch
//                            //MessageBox.Show("Nhap file thanh cong");
//                            //KeyIndex++;
//                            d_errorfile = 0; // Khi tim thay file tra error ve =0, de chay vong kich ban 03.
//                            if (countNextKey > 0)
//                            {
//                                KeyIndex = KeyIndex_backup;   //quay trở lại vị trí key đã quét mà bỏ qua trước đó
//                                // Đếm số lần quay lại.
//                                countBack++;
//                            }
//                            countNextKey = 0;
//                        }
//                        else
//                        {
//                            KeyIndex--;

//                            d_errorfile++; // Tăng số lần đếm không lấy được file: 1-chạy 04; 2-chạy 05

//                            if (d_errorfile == 1)
//                            {
//                                alarmCounter = StepTimer02[4] - 1;
//                                KeyIndex++;
//                            }
//                            else if (d_errorfile == 2)   //if (d_errorfile == 1 || d_errorfile == 2)
//                            {
//                                tmrPlan03.Stop();

//                                progressBar1.Value = 0;
//                                progressBar1.Maximum = max_Process_Plan04;

//                                alarmCounter = 0;
//                                tmrPlan04.Start();
//                                cboPlan.Text = "Kịch bản 04";
//                            }
//                            //else if (d_errorfile == 3)  // vao truong hop d_errorfile = 2 (xu ly tiep neu muon lon hon 2)
//                            //{
//                            //    tmrPlan03.Stop();
//                            //    progressBar1.Value = 0;
//                            //    progressBar1.Maximum = 9;
//                            //    alarmCounter = 0;
//                            //    tmrPlan05.Start();
//                            //    cboPlan.Text = "Kịch bản 03";
//                            //}
//                            else
//                            {
//                                if (countNextKey == 0) KeyIndex_backup = KeyIndex;
//                                countNextKey++;
//                                KeyIndex++;
//                                if (countNextKey >= 3)
//                                {
//                                    tmrPlan03.Stop();
//                                    alarmCounter = 0;
//                                    progressBar1.Value = 0;
//                                    //DialogResult r = MessageBox.Show("Lỗi nghiêm trọng! đã bỏ qua 3 keys vẫn chưa được. Vui lòng chạy lại trạng thái ban đầu", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
//                                    d_errorfile = 0;
//                                    KeyIndex = KeyIndex_backup; // khi gặp lỗi 3 lần, cũng sẽ quay lại, hoặc khi loadOK excel
//                                    //btnStart.Text = "Start";
//                                    //cboPlan.SelectedIndex = 2;  // Lựa chọn sẵn kịch bản 03
//                                    //tmrPlan01.Stop();
//                                    //tmrPlan02.Stop();
//                                    //tmrPlan03.Stop();
//                                    //tmrPlan04.Stop();
//                                    //tmrPlan05.Stop();
//                                    //btnStart_Click(btnStart, EventArgs.Empty);
//                                    //Viết lại kịch bản Login vào đây (chú ý, có 2 chỗ copy ở dưới nữa). Kiểm tra lại kịch bản 
//                                    tmrPlan02.Start();
//                                    cboPlan.Text = "Kịch bản Login";
//                                }
//                                else
//                                {
//                                    d_errorfile = 0;
//                                    btnStart_Click(btnStart, EventArgs.Empty);
//                                    btnNextKey_Click(btnNextKey, EventArgs.Empty);
//                                    //btnStart.Text = "Start";
//                                    if (KeyIndex >= dgrListKeywords.Rows.Count)  // Bỏ qua keywords cuối cùng, thinking tiếp nếu list có vol tận cuối cùng
//                                    {
//                                        // THoang 22:59 20230302
//                                        tmrPlan03.Stop();
//                                        alarmCounter = 0;
//                                        MessageBox.Show("Hoàn thành chiến dịch. Vui lòng Export File"); //Sau chuyển vào Label Trạng thái                        
//                                    }
//                                    else
//                                    {
//                                        btnStart_Click(btnStart, EventArgs.Empty);
//                                    }
//                                }


//                                //tmrPlan03.Stop();
//                                //progressBar1.Value = 0;
//                                //DialogResult result = MessageBox.Show("Lỗi nghiêm trọng, Bạn có muốn Next Keywords này không?? ","Thông báo lựa chọn", MessageBoxButtons.YesNo);
//                                //if (result == DialogResult.No)
//                                //{
//                                //    d_errorfile = 0;
//                                //    btnStart.Text = "Start";
//                                //    cboPlan.SelectedIndex = 2;  // Lựa chọn sẵn kịch bản 03
//                                //    tmrPlan01.Stop();
//                                //    tmrPlan02.Stop();
//                                //    tmrPlan03.Stop();
//                                //    tmrPlan04.Stop();
//                                //    tmrPlan05.Stop();
//                                //    btnStart_Click(btnStart, EventArgs.Empty);
//                                //}
//                                //else
//                                //{
//                                //    d_errorfile = 0;
//                                //    btnStart_Click(btnStart, EventArgs.Empty);
//                                //    btnNextKey_Click(btnNextKey, EventArgs.Empty);
//                                //    //btnStart.Text = "Start";
//                                //    if (KeyIndex >= dgrListKeywords.Rows.Count)  // Bỏ qua keywords cuối cùng, thinking tiếp nếu list có vol tận cuối cùng
//                                //    {
//                                //        // THoang 22:59 20230302
//                                //        tmrPlan03.Stop();
//                                //        alarmCounter = 0;
//                                //        MessageBox.Show("Hoàn thành chiến dịch. Vui lòng Export File"); //Sau chuyển vào Label Trạng thái                        
//                                //    }
//                                //    else
//                                //    {
//                                //        btnStart_Click(btnStart, EventArgs.Empty);
//                                //    }
//                                //}
//                            }
//                        }
//                    }
//                    else // của if(System.IO.File.Exists(filePath) == true)
//                    {
//                        // THoang 22:59 20230302

//                        KeyIndex--;

//                        d_errorfile++; // Tăng số lần đếm không lấy được file: 1-chạy 04; 2-chạy 05

//                        if (d_errorfile == 1)
//                        {
//                            alarmCounter = StepTimer02[4] - 1;
//                            KeyIndex++;
//                        }
//                        else if (d_errorfile == 2)    //(d_errorfile == 1 || d_errorfile == 2)
//                        {
//                            tmrPlan03.Stop();

//                            progressBar1.Value = 0;
//                            progressBar1.Maximum = max_Process_Plan04;

//                            alarmCounter = 0;
//                            tmrPlan04.Start();
//                            cboPlan.Text = "Kịch bản 04";
//                        }
//                        else if (d_errorfile == 3)  // vao truong hop d_errorfile = 2 (xu ly tiep neu muon lon hon 2)
//                        {
//                            tmrPlan03.Stop();
//                            progressBar1.Value = 0;
//                            progressBar1.Maximum = max_Process_Plan04;
//                            alarmCounter = 0;
//                            tmrPlan05.Start();
//                            cboPlan.Text = "Kịch bản 05";
//                        }
//                        else
//                        {
//                            ////tmrPlan03.Stop();
//                            ////progressBar1.Value = 0;
//                            ////MessageBox.Show("Lỗi nghiêm trọng, cập nhật lại email, mật khẩu và lựa chọn chạy lại kịch bản 03");
//                            ////d_errorfile = 0;

//                            //tmrPlan03.Stop();
//                            //progressBar1.Value = 0;
//                            //DialogResult result = MessageBox.Show("Lỗi nghiêm trọng, Bạn có muốn Next Keywords này không?? ", "Thông báo lựa chọn", MessageBoxButtons.YesNo);
//                            //if (result == DialogResult.No)
//                            //{
//                            //    d_errorfile = 0;

//                            //    //btnStart_Click(btnStart, EventArgs.Empty);

//                            //    btnStart.Text = "Start";
//                            //    cboPlan.SelectedIndex = 2;  // Lựa chọn sẵn kịch bản 03
//                            //    tmrPlan01.Stop();
//                            //    tmrPlan02.Stop();
//                            //    tmrPlan03.Stop();
//                            //    tmrPlan04.Stop();
//                            //    tmrPlan05.Stop();
//                            //    btnStart_Click(btnStart, EventArgs.Empty);
//                            //}
//                            //else
//                            //{
//                            //    d_errorfile = 0;
//                            //    btnStart_Click(btnStart, EventArgs.Empty);
//                            //    if (KeyIndex >= dgrListKeywords.Rows.Count)  // Bỏ qua keywords cuối cùng, thinking tiếp nếu list có vol tận cuối cùng
//                            //    {
//                            //        // THoang 22:59 20230302
//                            //        tmrPlan03.Stop();
//                            //        alarmCounter = 0;
//                            //        MessageBox.Show("Hoàn thành chiến dịch. Vui lòng Export File"); //Sau chuyển vào Label Trạng thái
//                            //    }
//                            //    else
//                            //    {
//                            //        btnNextKey_Click(btnNextKey, EventArgs.Empty);

//                            //        //btnStart.Text = "Start";
//                            //        btnStart_Click(btnStart, EventArgs.Empty);
//                            //    }
//                            //}

//                            if (countNextKey == 0) KeyIndex_backup = KeyIndex;

//                            countNextKey++;
//                            KeyIndex++;

//                            if (countNextKey >= 3)
//                            {
//                                tmrPlan03.Stop();
//                                alarmCounter = 0;
//                                progressBar1.Value = 0;
//                                //DialogResult r = MessageBox.Show("Lỗi nghiêm trọng! đã bỏ qua 3 keys vẫn chưa được", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
//                                d_errorfile = 0;
//                                KeyIndex = KeyIndex_backup; // khi gặp lỗi 3 lần, cũng sẽ quay lại, hoặc khi loadOK excel
//                                //btnStart.Text = "Start";
//                                //cboPlan.SelectedIndex = 2;  // Lựa chọn sẵn kịch bản 03
//                                //tmrPlan01.Stop();
//                                //tmrPlan02.Stop();
//                                //tmrPlan03.Stop();
//                                //tmrPlan04.Stop();
//                                //tmrPlan05.Stop();
//                                //btnStart_Click(btnStart, EventArgs.Empty);
//                                tmrPlan02.Start();
//                                cboPlan.Text = "Kịch bản Login";
//                            }
//                            else
//                            {
//                                d_errorfile = 0;
//                                btnStart_Click(btnStart, EventArgs.Empty);
//                                btnNextKey_Click(btnNextKey, EventArgs.Empty);
//                                //btnStart.Text = "Start";
//                                if (KeyIndex >= dgrListKeywords.Rows.Count)  // Bỏ qua keywords cuối cùng, thinking tiếp nếu list có vol tận cuối cùng
//                                {
//                                    // THoang 22:59 20230302
//                                    tmrPlan03.Stop();
//                                    alarmCounter = 0;
//                                    MessageBox.Show("Hoàn thành chiến dịch. Vui lòng Export File"); //Sau chuyển vào Label Trạng thái                        
//                                }
//                                else
//                                {
//                                    btnStart_Click(btnStart, EventArgs.Empty);
//                                }
//                            }
//                        }
//                    }

//                    //alarmCounter = 0;

//                    //this.WindowState = FormWindowState.Normal;

//                    progressBar1.Value += 1;
//                }
//                if (alarmCounter == StepTimer02[7])
//                {
//                    alarmCounter = 0;

//                    this.WindowState = FormWindowState.Normal;

//                    progressBar1.Value = 0;
//                }
//            }
//            //catch (Exception ex)
//            //{
//            //    MessageBox.Show(ex.Message);
//            //} 
//        }

//        private void tmrPlanLoadLaiExcel(object sender, EventArgs e)
//        {
//            alarmCounter++;

//            if (alarmCounter == StepTimer03[0]) //CLICK TO link
//            {
//                pt.X = 610;
//                pt.Y = 60;
//                Cursor.Position = pt;
//                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
//                progressBar1.Value += 1;
//            }
//            if (alarmCounter == StepTimer03[1]) //input text dia chi websites
//            {
//                SendKeys.Send(txtdiachi.Text);

//                progressBar1.Value += 1;
//            }
//            if (alarmCounter == StepTimer03[2]) //input enter
//            {
//                //SendKeys.SendWait("+(CTRL)");
//                //SendKeys.SendWait("+(A)");
//                SendKeys.Send("{ENTER}");

//                progressBar1.Value += 1;
//            }

//            if (alarmCounter == StepTimer03[3]) //Click CloudFare
//            {
//                ////pt.X = _ClickCloudFare04.X + 10;
//                ////pt.Y = _ClickCloudFare04.Y + 10;
//                ////Cursor.Position = pt;
//                ////pt.X = _ClickCloudFare04.X - 10;
//                ////pt.Y = _ClickCloudFare04.Y + 10;
//                ////Cursor.Position = pt;
//                ////pt.X = _ClickCloudFare04.X + 10;
//                ////pt.Y = _ClickCloudFare04.Y - 10;
//                ////Cursor.Position = pt;
//                ////pt.X = _ClickCloudFare04.X - 10;
//                ////pt.Y = _ClickCloudFare04.Y - 10;
//                ////Cursor.Position = pt;

//                //pt = _ClickCloudFare04;

//                //Cursor.Position = pt;
//                ////mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
//                //mouse_event(MOUSEEVENTF_RIGHTDOWN | MOUSEEVENTF_RIGHTUP, pt.X, pt.Y, 0, 0);
//                //AddList("Click CloudFare");
//                //MessageBox.Show("Đã click vào nút: X = " + pt.X + " Y = " + pt.Y);

//                progressBar1.Value += 1;
//            }

//            if (alarmCounter == StepTimer03[4]) //focus text search
//            {
//                //pt.X = 382;
//                //pt.Y = 486;
//                pt = _TextSearch04;

//                Cursor.Position = pt;
//                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);

//                progressBar1.Value += 1;
//            }
//            if (alarmCounter == StepTimer03[5]) //input text search
//            {
//                // SendKeys.Send(txtKeywords.Text);
//                SendKeys.Send("fsdfsdfsdf sdfsdfasdfsadf sdfsadf sadf");

//                progressBar1.Value += 1;
//            }
//            if (alarmCounter == StepTimer03[6]) // click search
//            {
//                //pt.X = 941;
//                //pt.Y = 486;
//                pt = _ButtonSearch04;

//                Cursor.Position = pt;
//                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);

//                progressBar1.Value += 1;
//            }
//            //if (alarmCounter == StepTimer03[6]) // click download button
//            //{
//            //    pt.X = 979;
//            //    pt.Y = 936;
//            //    Cursor.Position = pt;
//            //    mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
//            //    AddList("Click Download");

//            //    progressBar1.Value += 1;
//            //}

//            //if (alarmCounter == StepTimer03[7]) // click export to excel
//            //{
//            //    pt.X = 928;
//            //    pt.Y = 744;
//            //    Cursor.Position = pt;
//            //    mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
//            //    AddList("Click to excel File");

//            //    progressBar1.Value += 1;
//            //}
//            if (alarmCounter == StepTimer03[7])
//            {
//                alarmCounter = 0;

//                this.WindowState = FormWindowState.Normal;

//                progressBar1.Value += 1;

//                tmrPlan04.Stop();
//                progressBar1.Value = 0;
//                progressBar1.Maximum = max_Process_Plan03;

//                alarmCounter = 0;
//                tmrPlan03.Start();
//                cboPlan.Text = "Kịch bản 03";
//            }
//        }

//        private void txtTotal_TextChanged(object sender, EventArgs e)
//        {

//        }

//        // T.Hoàng code 4:28 ngày 20230302
//        public static bool IsNumeric(object value)
//        {
//            try
//            {
//                double d = Convert.ToDouble(value);
//                return true;
//            }
//            catch (FormatException)
//            {
//                return false;
//            }
//        }

//        private void tmrPlanXen(object sender, EventArgs e)
//        {
//            alarmCounter++;
//            if (alarmCounter == StepTimer04[0]) //focus text taikhoan
//            {
//                pt.X = 641;
//                pt.Y = 376;
//                Cursor.Position = pt;
//                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);

//                progressBar1.Value += 1;
//            }
//            if (alarmCounter == StepTimer04[1]) //input to tai khoan
//            {
//                SendKeys.Send(txttaikhoan.Text);

//                progressBar1.Value += 1;
//            }
//            if (alarmCounter == StepTimer04[2]) //focus text mat khau
//            {
//                pt.X = 641;
//                pt.Y = 461;
//                Cursor.Position = pt;
//                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);

//                progressBar1.Value += 1;
//            }
//            if (alarmCounter == StepTimer04[3]) //input text mat khau
//            {
//                SendKeys.Send(txtmatkhau.Text);

//                progressBar1.Value += 1;
//            }
//            if (alarmCounter == StepTimer04[4]) // click login
//            {
//                pt.X = 255;
//                pt.Y = 564;
//                Cursor.Position = pt;
//                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);

//                progressBar1.Value += 1;
//            }
//            if (alarmCounter == StepTimer04[5]) //focus text search
//            {
//                pt.X = 382;
//                pt.Y = 486;
//                Cursor.Position = pt;
//                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);

//                progressBar1.Value += 1;
//            }
//            if (alarmCounter == StepTimer04[6]) //input text search
//            {
//                SendKeys.Send(txtKeywords.Text);

//                progressBar1.Value += 1;
//            }
//            if (alarmCounter == StepTimer04[7]) // click search
//            {
//                pt.X = 941;
//                pt.Y = 486;
//                Cursor.Position = pt;
//                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);

//                progressBar1.Value += 1;
//            }

//            if (alarmCounter == StepTimer04[8])
//            {
//                alarmCounter = 0;

//                progressBar1.Value += 1;

//                tmrPlan05.Stop();
//                progressBar1.Value = 0;
//                progressBar1.Maximum = max_Process_Plan03;

//                alarmCounter = 0;
//                tmrPlan03.Start();
//                cboPlan.Text = "Kịch bản 03";
//            }
//        }


//        private void quickExportExcelToolStripMenuItem_Click(object sender, EventArgs e)
//        {
//            SaveFileDialog saveFileDialog = new SaveFileDialog();
//            saveFileDialog.Title = "Export Excel";
//            saveFileDialog.Filter = "Excel(*.xlsx)|*.xlsx|Excel 2016(*.xls)|*.xls";
//            if (saveFileDialog.ShowDialog() == DialogResult.OK)
//            {
//                //try

//                QuickExportExcel(saveFileDialog.FileName);

//                //catch (Exception ex)
//                //{
//                //    MessageBox.Show("Xuat file khong thanh cong \n" + ex.Message);
//                //}
//            }
//        }

//        private void clearDataGridToolStripMenuItem_Click(object sender, EventArgs e)
//        {
//            DialogResult result = MessageBox.Show("Bạn có chắc chắn muốn xóa Danh sách này?", "Xác nhận xóa", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
//            if (result == DialogResult.Yes)
//            {
//                dgrListKeywords.Rows.Clear();
//                KeyIndex = 0;
//            }

//        }

//        private void đọcFileKịchBảnToolStripMenuItem_Click(object sender, EventArgs e)
//        {
//            LoadPlanPoint("Plan.xlsx");
//        }

//        private void btnNextKey_Click(object sender, EventArgs e)
//        {
//            KeyIndex++;
//            if (countNextKey == 0 || countBack >= 2)  // đếm 2 lần quay lại key ko được, coi như key hỏng và bỏ qua hẳn
//            {
//                dgrListKeywords.Rows[KeyIndex - 1].Cells[2].Value = LevelSearch; //"100"; //Giá trị cao, mặc định là 100 để vượt qua Level
//                dgrListKeywords.Rows[KeyIndex - 1].Cells[1].Value = "0"; // Đưa Vol về = 0, chưa hiểu ý thầy Bình chỗ này!
//                countBack = 0; // đếm lại
//            }
//            txtKeywords.Text = Convert.ToString(dgrListKeywords.Rows[KeyIndex].Cells[0].Value);
//        }

//        private void FocusCurrentCell(DataGridView dataGridView, int curRow)
//        {
//            if (dataGridView.CurrentRow != null && dataGridView.CurrentCell != null)
//            {
//                // Lưu trữ vị trí ô cell hiện tại
//                int currentRowIndex = curRow;
//                int currentColumnIndex = 0;

//                // Di chuyển tiêu điểm đến ô cell hiện tại
//                dataGridView.CurrentCell = dataGridView.Rows[currentRowIndex].Cells[currentColumnIndex];

//                // Tập trung vào DataGridView để ô cell hiện tại trở thành tiêu điểm
//                dataGridView.Focus();
//            }
//        }

//        private void ghiFileCàiĐặtToolStripMenuItem_Click(object sender, EventArgs e)
//        {
//            frmConfig f_config = new frmConfig();
//            f_config.ShowDialog();
//        }

//        private void dgrListKeywords_CellContentClick(object sender, DataGridViewCellEventArgs e)
//        {

//        }

//        private void xóaKeyĐangChọnToolStripMenuItem_Click(object sender, EventArgs e)
//        {
//            if (dgrListKeywords.CurrentRow.Index >= 0)
//            {
//                dgrListKeywords.Rows.RemoveAt(dgrListKeywords.CurrentRow.Index);
//            }
//        }
//        //////////////////////////////////////////////////////////////////////
//    }
//}