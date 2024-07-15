using GetKeywords.Modules;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;

using System;
using System.Collections;
using System.Data;

using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;

using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

using System.Collections;
using System.Linq;
using System.Diagnostics;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Threading.Tasks;
using System.Collections.Generic;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;

namespace GetKeywords
{
    public partial class Form1 : Form
    {
        [DllImport("user32.dll")]
        public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        [DllImport("user32.dll")]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        public const int SWP_SHOWWINDOW = 0x0040;

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

        private int maxKeyExcel = 1000000;

        private string ListAccount;
        private string[] Account = new string[5];
        private int AccountIndex = 0;
        private string passLogin;

        private int CurDownload = 0;

        private int max_Process_Plan01;
        private int max_Process_Plan02;
        private int max_Process_Plan03; // Số lượng tối đa tiến trình trên thanh trượt kịch bản 3
        private int max_Process_Plan04; // Số lượng tối đa tiến trình trên thanh trượt kịch bản 4

        private string[] ListSuggestKeys;
        private string[] ListNegativeKeys;
        private char[] Separator = {'|'};

        //private string[] strKeyWords= new string[2000000];
        //private string[] strVolume = new string[2000000];
        //private string[] strCheck = new string[2000000];
        //private string[] strComp = new string[2000000];

        private ArrayList strKeyWords = new ArrayList();

        private ArrayList strVolume = new ArrayList();
        private ArrayList strCheck = new ArrayList();
        private ArrayList strComp = new ArrayList();
        private ArrayList strKeyCha1 = new ArrayList();
        private ArrayList strKeyCha2 = new ArrayList();
        private ArrayList strKeyCha3 = new ArrayList();
        //private ArrayList strLinkShopee = new ArrayList();

        //Khởi tạo các nút kịch bản Login
        private System.Drawing.Point _AccountButtonLogin;
        private System.Drawing.Point _LogoutButtonLogin;
        private System.Drawing.Point _LoginButtonLogin;
        private System.Drawing.Point _EmailTextLogin;
        private System.Drawing.Point _PassTextLogin;
        private System.Drawing.Point _TextSearchLogin;
        private System.Drawing.Point _ButtonSearchLogin;
        private System.Drawing.Point _AccountMenu;
        private System.Drawing.Point _LogoutMenu;
        private System.Drawing.Point _LoginMenu;
        private System.Drawing.Point _PointHeaderAccLogin;
        private System.Drawing.Point _PointHeaderAccLogout;


        // Khởi tạo các nút kịch bản 03:
        private System.Drawing.Point _TextSearch03;
        private System.Drawing.Point _ButtonSearch03;
        private System.Drawing.Point _ButtonDownload03;
        private System.Drawing.Point _ButtonExcel03;

        // Khởi tạo các nút kịch bản 04:
        private System.Drawing.Point _TextSearch04;
        private System.Drawing.Point _ButtonSearch04;
        private System.Drawing.Point _ClickCloudFare04;
        // Khởi tạo các nút đóng mở hotspot

        private System.Drawing.Point _CloseButton;
        private System.Drawing.Point _Turnon;
        private System.Drawing.Point _Turnoff;
        private System.Drawing.Point _CheckHuman;
        private System.Drawing.Point _CheckHuman2;
        private System.Drawing.Point _KeySug;


        // Khởi tạo các nút Resize:
        private int _OxLogout;
        private int _OyLogout;
        private int _dXLogout;
        private int _dYLogout;

        private int _OxLogin;
        private int _OyLogin;
        private int _dXLogin;
        private int _dYLogin;

        private Random rand = new Random();
        private int countOrigin = 0;

        System.Data.DataTable dtKeywords = new System.Data.DataTable("Keywords");

        HashSet<string> suggestKeySet;
        HashSet<string> nagativeKeySet;

        public void InitVar_Plan()
        {
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
            NextStepDelay01[4] = 6; // click to Logout
            NextStepDelay01[5] = 2; // click to Login Again
            NextStepDelay01[6] = 2; // focus to tai khoan
            NextStepDelay01[7] = 1; // input text tai khoan
            NextStepDelay01[8] = 1; // focus to mat khau
            NextStepDelay01[9] = 1; // input text mat khau
            NextStepDelay01[10] = 6; // click to login
            NextStepDelay01[11] = 1; //focus to text keyword
            NextStepDelay01[12] = 1; //input to text keyword
            NextStepDelay01[13] = 3; // click search
            NextStepDelay01[14] = 1; // click download button
            NextStepDelay01[15] = 1; // click excel
            NextStepDelay01[16] = 2; // click stop

            StepTimer01[0] = 2;
            for (int j = 0; j < 20; j++)
            {
                StepTimer01[j + 1] = StepTimer01[j] + NextStepDelay01[j]*3;
            }
            max_Process_Plan02 = 20;

            // Delay Time after event 3
            NextStepDelay02[0] = 4; // + rand.Next(100); //focus text search
            NextStepDelay02[1] = 1; //+ rand.Next(100); //CtrlA
            NextStepDelay02[2] = 2; //+ rand.Next(100); // input text search
            NextStepDelay02[3] = 18; // InitVar.v_TimeDownload; //+ rand.Next(100); // click search
            NextStepDelay02[4] = 1; //+ rand.Next(100); // click download button
            NextStepDelay02[5] = 5; //+ rand.Next(100); // click excel
            NextStepDelay02[6] = 1; //+ rand.Next(100); // click kiem tra file
            NextStepDelay02[7] = 1; //+ rand.Next(100); // click stop
            
            StepTimer02[0] = 10;
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

             
        }

        public Form1()
        {
            InitializeComponent();
            // Lấy các dữ liệu setting
            // Delay Time after event 1

            InitVar_Plan();

            // Load Toa do Kich ban
            LoadPlanPoint("Plan.xlsx");

            SetSize();
        }
        private const int MOUSEEVENTF_LEFTDOWN = 0x02;
        private const int MOUSEEVENTF_LEFTUP = 0x04;
        private const int MOUSEEVENTF_RIGHTDOWN = 0x08;
        private const int MOUSEEVENTF_RIGHTUP = 0x10;

        private const int MOUSEEVENTF_MOVE = 0x0001;
        //private const int MOUSEEVENTF_LEFTDOWN = 0x0002;
        //private const int MOUSEEVENTF_LEFTUP = 0x0004;
        //private const int MOUSEEVENTF_RIGHTDOWN = 0x0008;
        //private const int MOUSEEVENTF_RIGHTUP = 0x0010;
        private const int MOUSEEVENTF_MIDDLEDOWN = 0x0020;
        private const int MOUSEEVENTF_MIDDLEUP = 0x0040;
        private const int MOUSEEVENTF_XDOWN = 0x0080;
        private const int MOUSEEVENTF_XUP = 0x0100;
        private const int MOUSEEVENTF_WHEEL = 0x0800;
        private const int MOUSEEVENTF_HWHEEL = 0x01000;
        private const int MOUSEEVENTF_ABSOLUTE = 0x8000;

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
            suggestKeySet = new HashSet<string>(ListSuggestKeys);
            nagativeKeySet = new HashSet<string>(ListNegativeKeys);

                if (txtCur.Text != null)
                {
                    KeyIndex = Convert.ToInt32(txtCur.Text);
                }

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
                    alarmCounter = 0;
                    tmrPlan02.Start();
                    cboPlan.Text = "Kịch bản 02";
                }
                if (cboPlan.SelectedIndex == 1) // Lựa chọn Get Keywords
                {

                    progressBar1.Maximum = max_Process_Plan01; // số lượng các thao tác trong kế hoạch.
                    progressBar1.Value = 0;
                    tmrPlan01.Interval = InitVar.v_speed;
                    alarmCounter = 0;
                    tmrPlan01.Start();
                    cboPlan.Text = "Kịch bản 01";
                }
                if (cboPlan.SelectedIndex == 2) // Lựa chọn Get Keywords tiep theo
                {

                    progressBar1.Maximum = max_Process_Plan03; // số lượng các thao tác trong kế hoạch 03.
                    progressBar1.Value = 0;
                    tmrPlan03.Interval = InitVar.v_speed;
                    alarmCounter = 0;
                    tmrPlan03.Start();
                    cboPlan.Text = "Kịch bản 03";
                }
                if (cboPlan.SelectedIndex == 3) // không tìm thấy file ex
                {
                    progressBar1.Maximum = max_Process_Plan04; // số lượng các thao tác trong kế hoạch 04.
                    progressBar1.Value = 0;
                    tmrPlan04.Interval = InitVar.v_speed;
                    alarmCounter = 0;
                    tmrPlan04.Start();
                    cboPlan.Text = "Kịch bản LoadExcel";
                }
                if (cboPlan.SelectedIndex == 4) // xen
                {
                    progressBar1.Maximum = 9; // số lượng các thao tác trong kế hoạch.
                    progressBar1.Value = 0;
                    tmrPlan05.Interval = InitVar.v_speed;
                    alarmCounter = 0;
                    tmrPlan05.Start();
                    alarmCounter = 0;
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
                if (alarmCounter >= StepTimer02[2])
                {
                    KeyIndex--;
                    //FocusCurrentCell(dgrListKeywords, KeyIndex - 1);
                }
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
        /// 


        //Ghi chú - Check human: Dưới: 265 - 620     Trên: 220 - 315;

        // Đặt trạng thái hiển thị của cửa sổ
        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        // Import hàm SetForegroundWindow từ user32.dll
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd); // Hàm đưa cửa sổ lên phía trước.

        public static void AutoMouseClick(System.Drawing.Point pt, int _delay = 200, int clickCount = 1)
        {
            for (int i = 0; i < clickCount; i++)
            {
                //Thread.Sleep(_delay);
                System.Windows.Forms.Cursor.Position = new System.Drawing.Point(pt.X, pt.Y);
                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
            }
        }


        // Phương thức gộp các ArrayList thành DataTable
        static void MergeArrayListsToDataTable(ArrayList[] arrayLists, System.Data.DataTable dt)
        {
            int maxRows = arrayLists.Max(arr => arr.Count); // Lấy số hàng tối đa
            while (dt.Columns.Count > 0)
            {
                dt.Columns.RemoveAt(0);
            }
            for (int i = 0; i < arrayLists.Length; i++)
            {
                dt.Columns.Add("Column" + (i + 1), typeof(object));
                //dt.Columns[i].ExtendedProperties.Add("Width", 50);
                for (int j = 0; j < maxRows; j++)
                {
                    if (j < arrayLists[i].Count)
                    {
                        if (dt.Rows.Count <= j)
                            dt.Rows.Add(dt.NewRow());

                        dt.Rows[j][i] = arrayLists[i][j];
                    }
                }
            }
        }

        private void Importexcel(string path)
        {
            int i = 1;
            int index = 0;

            using (FileStream fileStream = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                using (ExcelPackage excelPackage = new ExcelPackage(fileStream))
                {

                    // Thay đổi cấu trúc file excel, Export done! Tối ưu vị trí dòng khi đọc dữ liệu. Vui lòng giữ nguyên trạng thái.

                    for (int iSheet = 0; iSheet < excelPackage.Workbook.Worksheets.Count; iSheet++)
                    {
                        // Lấy laanf luowtj trong tệp Excel
                        ExcelWorksheet excelWorksheet = excelPackage.Workbook.Worksheets[iSheet];

                        while ((excelWorksheet.Cells[i + 1, 1].Value != null) && (excelWorksheet.Cells[i + 1, 2].Value != null) && (excelWorksheet.Cells[i + 1, 2].Value.ToString().Contains(".") == false))
                        {
                            // Cột 17 là vị trí của cột Competition;......
                            strKeyWords.Add(Convert.ToString(excelWorksheet.Cells[i + 1, 1].Value));
                            strVolume.Add(Convert.ToString(excelWorksheet.Cells[i + 1, 2].Value));
                            if (excelWorksheet.Cells[i + 1, 3].Value == null)
                            {
                                strCheck.Add("0");
                                //strComp.Add(Convert.ToString(excelWorksheet.Cells[i + 1, 4].Value));
                            }
                            else
                            {
                                //strKeyWords.Add(Convert.ToString(excelWorksheet.Cells[i + 1, 1].Value));
                                //strVolume.Add(Convert.ToString(excelWorksheet.Cells[i + 1, 2].Value));
                                strCheck.Add(Convert.ToString(excelWorksheet.Cells[i + 1, 3].Value));
                                //strComp.Add(Convert.ToString(excelWorksheet.Cells[i + 1, 4].Value));
                            }

                            if ((excelWorksheet.Cells[i + 1, 4].Value == null) || (excelWorksheet.Cells[i + 1, 4].Value == ""))
                            {
                                strComp.Add("0");
                            }
                            else
                            {
                                strComp.Add(Convert.ToString(excelWorksheet.Cells[i + 1, 4].Value));
                            }

                            if ((excelWorksheet.Cells[i + 1, 5].Value == null) || (excelWorksheet.Cells[i + 1, 5].Value == ""))
                            {
                                strKeyCha1.Add("");
                            }
                            else
                            {
                                strKeyCha1.Add(Convert.ToString(excelWorksheet.Cells[i + 1, 5].Value));
                            }

                            if ((excelWorksheet.Cells[i + 1, 6].Value == null) || (excelWorksheet.Cells[i + 1, 6].Value == ""))
                            {
                                strKeyCha2.Add("");
                            }
                            else
                            {
                                strKeyCha2.Add(Convert.ToString(excelWorksheet.Cells[i + 1, 6].Value));
                            }
                            if ((excelWorksheet.Cells[i + 1, 7].Value == null) || (excelWorksheet.Cells[i + 1, 7].Value == ""))
                            {
                                strKeyCha3.Add("");
                            }
                            else
                            {
                                strKeyCha3.Add(Convert.ToString(excelWorksheet.Cells[i + 1, 6].Value));
                            }

                            //if ((excelWorksheet.Cells[i + 1, 7].Value == null) || (excelWorksheet.Cells[i + 1, 7].Value == ""))
                            //{
                            //    strLinkShopee.Add("");
                            //}
                            //else
                            //{
                            //    strLinkShopee.Add(Convert.ToString(excelWorksheet.Cells[i + 1, 6].Value));
                            //}



                            i++;
                        }
                        i = i - maxKeyExcel;

                        // Lấy các dữ liệu thông tin chung ở Sheet đầu tiên/
                        if (iSheet == 0)
                        {

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

                    LoadDataToGrid();
                    // Đưa giữ liệu lên trên DataGridView ...
                    //MergeArrayListsToDataTable(new ArrayList[] {strKeyWords, strVolume, strCheck, strComp}, dtKeywords);
                    //dtKeywords.Columns["Column1"].ExtendedProperties.Add("Width", 100); // Độ rộng cho cột 1
                    //dtKeywords.Columns["Column2"].ExtendedProperties.Add("Width", 50); // Độ rộng cho cột 2
                    //dtKeywords.Columns["Column3"].ExtendedProperties.Add("Width", 20); // Độ rộng cho cột 3
                    //dtKeywords.Columns["Column4"].ExtendedProperties.Add("Width", 30); // Độ rộng cho cột 4

                    //dgrListKeywords.DataSource = dtKeywords;

                    foreach (DataGridViewColumn column in dgrListKeywords.Columns)
                    {
                        string columnName = column.HeaderText;
                        int width = Convert.ToInt32(dtKeywords.Columns[columnName].ExtendedProperties["Width"]);
                        column.Width = width;
                    }
                    FocusCurrentCell(dgrListKeywords, KeyIndex - 1);
                }
            }


            txtTotal.Text = Convert.ToString(strKeyWords.Count);  
        }

        private void LoadDataToGrid()
        {
            // Đưa giữ liệu lên trên DataGridView ...
            MergeArrayListsToDataTable(new ArrayList[] { strKeyWords, strVolume, strCheck, strComp }, dtKeywords);
            dtKeywords.Columns["Column1"].ExtendedProperties.Add("Width", 100); // Độ rộng cho cột 1
            dtKeywords.Columns["Column2"].ExtendedProperties.Add("Width", 50); // Độ rộng cho cột 2
            dtKeywords.Columns["Column3"].ExtendedProperties.Add("Width", 20); // Độ rộng cho cột 3
            dtKeywords.Columns["Column4"].ExtendedProperties.Add("Width", 30); // Độ rộng cho cột 4

            dgrListKeywords.DataSource = dtKeywords;
        }

        /// <summary>
        /// doan nhap file cho vong lay keywords tiep theo
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// 
        private async void ImportExcelCircle(string path, string Key)
        {
            prcExcel.Value = 10;
            List<string> listOfKeywords = strKeyWords.Cast<string>().ToList();
            HashSet<string> keyWordsSet = new HashSet<string>(listOfKeywords.Select(k => k.ToLowerInvariant()));
            Dictionary<string, int> keywordDict = new Dictionary<string, int>(StringComparer.InvariantCultureIgnoreCase);

            using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(path)))
            {
                if (excelPackage.Workbook.Worksheets.Count > 0)
                {
                    ExcelWorksheet excelWorksheet = excelPackage.Workbook.Worksheets[0];
                    int i = 1;
                    while (excelWorksheet.Cells[i + 1, 1].Value != null) //&& (excelWorksheet.Cells[i + 1, 2].Value!= null)) // && Convert.ToInt32(excelWorksheet.Cells[i + 1, 2].Value) > 1000))
                    {
                        if (i < 10000)
                        {
                            prcExcel.Value = 20 + i / 10000*60;
                        }
                        else
                        {
                            prcExcel.Value = 80;
                        }
                        // THoang code 21:46 20230301
                        if ((excelWorksheet.Cells[i + 1, 2].Value != null) && (Convert.ToInt32(excelWorksheet.Cells[i + 1, 2].Value) >= InitVar.v_VolMin) && (Convert.ToInt32(excelWorksheet.Cells[i + 1, 17].Value) <= InitVar.v_LevelDif))
                        {
                            string str2 = excelWorksheet.Cells[i + 1, 1].Value.ToString();
                            cboPlan.Text = "Check uniqueKeys";
                            //bool dup = keyWordsSet.Contains(str2);
                            bool dup = false;
                            // Tạo từ điển để tra cứu nhanh
                            for (int j = 0; j < listOfKeywords.Count; j++)
                            {
                                keywordDict[listOfKeywords[j]] = j;
                            }

                            // Kiểm tra và xử lý từ khóa
                            if (keywordDict.TryGetValue(str2, out int index))
                            {
                                CurrentKeywords = txtKeywords.Text;

                                if (strKeyCha2[index].ToString() == "")
                                {
                                    if (!CurrentKeywords.Equals(strKeyCha1[index].ToString(), StringComparison.InvariantCultureIgnoreCase))
                                    {
                                        // Kiểm tra nếu không trùng với KeyCha1 thì mới thêm vào KeyCha2.
                                        strKeyCha2[index] = CurrentKeywords;
                                    }
                                }
                                else if (strKeyCha3[index].ToString() == "")
                                {
                                    if (!CurrentKeywords.Equals(strKeyCha1[index].ToString(), StringComparison.InvariantCultureIgnoreCase) &&
                                        !CurrentKeywords.Equals(strKeyCha2[index].ToString(), StringComparison.InvariantCultureIgnoreCase))
                                    {
                                        strKeyCha3[index] = CurrentKeywords;
                                    }
                                }

                                if (CurrentKeywords.Equals(str2, StringComparison.InvariantCultureIgnoreCase))
                                {
                                    strVolume[KeyIndex - 1] = excelWorksheet.Cells[i + 1, 2].Value;
                                }
                                dup = true;
                            }
                            cboPlan.Text = "Done uniqueKeys";
                            if (dup == false)
                            {
                                // Kiểm tra ListSuggestKeys & ListNegativeKeys
                                bool sug = suggestKeySet.Any(key => str2.Contains(key));
                                bool nega = ListNegativeKeys.Any(key => str2.Contains(key));

                                if ((sug == true) && (nega == false))
                                {
                                    strKeyWords.Insert(KeyIndex, excelWorksheet.Cells[i + 1, 1].Value);
                                    strVolume.Insert(KeyIndex, excelWorksheet.Cells[i + 1, 2].Value);
                                    strCheck.Insert(KeyIndex, (LevelSearch + 1));
                                    strComp.Insert(KeyIndex, excelWorksheet.Cells[i + 1, 17].Value);
                                    strKeyCha1.Insert(KeyIndex, txtKeywords.Text);
                                    strKeyCha2.Insert(KeyIndex, "");
                                    strKeyCha3.Insert(KeyIndex, "");
                                }
                            }
                            //}
                            LoadFileExcelOK = 0;
                            //LoadDataToGrid();
                        }
                        i++;
                    }
                }
                prcExcel.Value = 100;
                File.Delete(path); // THoang: Xóa luôn file sau khi đã nạp
                cboPlan.Text = "Da xoa file";
                //return kq;
            }
        }

        private async void ImportExcelCircleBackup(string path, string Key)
        {
            using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(path)))
            {
                if (excelPackage.Workbook.Worksheets.Count > 0)
                {
                    ExcelWorksheet excelWorksheet = excelPackage.Workbook.Worksheets[0];
                    int i = 1;
                    while (excelWorksheet.Cells[i + 1, 1].Value != null) //&& (excelWorksheet.Cells[i + 1, 2].Value!= null)) // && Convert.ToInt32(excelWorksheet.Cells[i + 1, 2].Value) > 1000))
                    {

                        // THoang code 21:46 20230301
                        if ((excelWorksheet.Cells[i + 1, 2].Value != null) && (Convert.ToInt32(excelWorksheet.Cells[i + 1, 2].Value) >= InitVar.v_VolMin) && (Convert.ToInt32(excelWorksheet.Cells[i + 1, 17].Value) <= InitVar.v_LevelDif))
                        {
                             string str2 = excelWorksheet.Cells[i + 1, 1].Value.ToString();
                            // Kiem tra trung lap trong danh sach
                            cboPlan.Text = "Check uniqueKeys";
                            bool dup = false;
                            for (int j = 0; j <= strKeyWords.Count - 1; j++)
                            {

                                string str1 = strKeyWords[j].ToString();
                                if (str1.Equals(str2, StringComparison.InvariantCultureIgnoreCase))
                                {
                                    dup = true;

                                    CurrentKeywords = txtKeywords.Text;

                                    if (strKeyCha2[j].ToString() == "")  // Keycha2 chưa có mới xử lý
                                    {
                                        if (!CurrentKeywords.Equals(strKeyCha1[j].ToString(), StringComparison.InvariantCultureIgnoreCase))
                                        {
                                            // Kiểm tra nếu không trùng với KeyCha1 thì mới thêm vào KeyCha2.
                                            strKeyCha2[j] = CurrentKeywords;
                                        }
                                    }
                                    else if (strKeyCha3[j].ToString() == "")  // keychar2 có rồi thì mới chuyển sang keycha3
                                    {
                                        if (!CurrentKeywords.Equals(strKeyCha1[j].ToString(), StringComparison.InvariantCultureIgnoreCase) && !CurrentKeywords.Equals(strKeyCha2[j].ToString(), StringComparison.InvariantCultureIgnoreCase))
                                        {
                                            // Kiểm tra nếu không trùng với KeyCha1 & KeyCha2 thì mới thêm vào KeyCha3.
                                            strKeyCha3[j] = CurrentKeywords;
                                        }
                                    }

                                    if (CurrentKeywords.Equals(str2, StringComparison.InvariantCultureIgnoreCase))
                                    {
                                        strVolume[KeyIndex - 1] = excelWorksheet.Cells[i + 1, 2].Value;
                                    }
                                    break;
                                }

                            }
                            cboPlan.Text = "Done uniqueKeys";

                            if (dup == false)
                            {
                                // Kiểm tra ListSuggestKeys & ListNegativeKeys
                                bool sug = suggestKeySet.Any(key => str2.Contains(key));
                                bool nega = ListNegativeKeys.Any(key => str2.Contains(key));

                                if ((sug == true) && (nega == false))
                                {
                                    strKeyWords.Insert(KeyIndex, excelWorksheet.Cells[i + 1, 1].Value);
                                    strVolume.Insert(KeyIndex, excelWorksheet.Cells[i + 1, 2].Value);
                                    strCheck.Insert(KeyIndex, (LevelSearch + 1));
                                    strComp.Insert(KeyIndex, excelWorksheet.Cells[i + 1, 17].Value);
                                    strKeyCha1.Insert(KeyIndex, txtKeywords.Text);
                                    strKeyCha2.Insert(KeyIndex, "");
                                    strKeyCha3.Insert(KeyIndex, "");
                                }
                            }
                            //}
                            LoadFileExcelOK = 0;
                            LoadDataToGrid();
                        }
                        i++;
                    }
                }
                File.Delete(path); // THoang: Xóa luôn file sau khi đã nạp
                cboPlan.Text = "Da xoa file";
                //return kq;
            }
        }


        public void SetSize()
        {
            Process[] processesLogout = Process.GetProcessesByName("firefox"); // Tìm các tiến trình của Edge

            if (processesLogout.Length > 0)
                for (int i = 0; i < processesLogout.Length; i++)
                {
                    IntPtr hWnd = processesLogout[i].MainWindowHandle; // Lấy handle của cửa sổ chính của Edge
                    SetWindowPos(hWnd, IntPtr.Zero, _OxLogout, _OyLogout, _dXLogout, _dYLogout, SWP_SHOWWINDOW);
                }

            Process[] processesLogin = Process.GetProcessesByName("chrome"); // Tìm các tiến trình của Edge

            if (processesLogin.Length > 0)
                for (int i = 0; i < processesLogin.Length; i++)
                {
                    IntPtr hWnd = processesLogin[i].MainWindowHandle; // Lấy handle của cửa sổ chính của Edge
                    SetWindowPos(hWnd, IntPtr.Zero, _OxLogin, _OyLogin, _dXLogin, _dYLogin, SWP_SHOWWINDOW);
                }
            else
            {
                MessageBox.Show("Không tìm thấy trình duyệt đang chạy.");
            }
        }
        private static void OpenApp(string path)
        {
            try
            {
                // Đường dẫn tới ứng dụng muốn mở
                string applicationPath = path;

                // Tạo một đối tượng ProcessStartInfo
                ProcessStartInfo processStartInfo = new ProcessStartInfo
                {
                    FileName = applicationPath,
                    // Thêm các tham số (nếu có)
                    Arguments = "",
                    // Đặt các tùy chọn khác nếu cần
                    UseShellExecute = true,
                    CreateNoWindow = false
                };

                // Mở ứng dụng
                Process process = Process.Start(processStartInfo);
                Console.WriteLine("Ứng dụng đã được mở.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Có lỗi xảy ra: " + ex.Message);
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
                _Turnon.X = 341;
                _Turnon.Y = 285;

                _Turnoff.X = 342;
                _Turnoff.Y = 569;

                _CloseButton.X = 1200;
                _CloseButton.Y = 25;

                _CheckHuman.X = 310;
                _CheckHuman.Y = 285;

                _CheckHuman2.X = 265;
                _CheckHuman2.Y = 580;

                _KeySug.X = 90;
                _KeySug.Y = 560;

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

                    _AccountMenu.X = Convert.ToInt32(excelWorksheetPlanLogin.Cells[9, 2].Value);
                    _AccountMenu.Y = Convert.ToInt32(excelWorksheetPlanLogin.Cells[9, 3].Value);

                    _LogoutMenu.X = Convert.ToInt32(excelWorksheetPlanLogin.Cells[10, 2].Value);
                    _LogoutMenu.Y = Convert.ToInt32(excelWorksheetPlanLogin.Cells[10, 3].Value);

                    _LoginMenu.X = Convert.ToInt32(excelWorksheetPlanLogin.Cells[11, 2].Value);
                    _LoginMenu.Y = Convert.ToInt32(excelWorksheetPlanLogin.Cells[11, 3].Value);


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
                    
                    _PointHeaderAccLogin.X = Convert.ToInt32(excelWorksheetPlan03.Cells[6, 2].Value);
                    _PointHeaderAccLogin.Y = Convert.ToInt32(excelWorksheetPlan03.Cells[6, 3].Value);

                    _PointHeaderAccLogout.X = Convert.ToInt32(excelWorksheetPlan03.Cells[7, 2].Value);
                    _PointHeaderAccLogout.Y = Convert.ToInt32(excelWorksheetPlan03.Cells[7, 3].Value);

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
                    ///// Load Rize
                    ///
                    ExcelWorksheet excelWorksheetSizeBrower = excelPackage.Workbook.Worksheets["Resize"];
                    _OxLogout = Convert.ToInt32(excelWorksheetSizeBrower.Cells[2, 1].Value);
                    _OyLogout = Convert.ToInt32(excelWorksheetSizeBrower.Cells[3, 1].Value);
                    _dXLogout = Convert.ToInt32(excelWorksheetSizeBrower.Cells[4, 1].Value);
                    _dYLogout = Convert.ToInt32(excelWorksheetSizeBrower.Cells[5, 1].Value);

                    _OxLogin = Convert.ToInt32(excelWorksheetSizeBrower.Cells[6, 1].Value);
                    _OyLogin = Convert.ToInt32(excelWorksheetSizeBrower.Cells[7, 1].Value);
                    _dXLogin = Convert.ToInt32(excelWorksheetSizeBrower.Cells[8, 1].Value);
                    _dYLogin = Convert.ToInt32(excelWorksheetSizeBrower.Cells[9, 1].Value);

                    excelPackage.Dispose();
                }
              
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error load PlanPoint: " + ex.Message);
            }
        }

        public static void ByPassUnusualActivity(System.Drawing.Point Turnon, System.Drawing.Point Turnoff, System.Drawing.Point _CloseButton, System.Drawing.Point _CheckHuman)
        {
            const uint SWP_NOSIZE = 0x0001;
            const uint SWP_NOZORDER = 0x0004;
            const uint SWP_SHOWWINDOW = 0x0040;
            //const int SW_MINIMIZE = 6;
            //const int SW_MAXIMIZE = 3;
            const int SW_MINIMIZE = 6;
            const int SW_NORMAL = 1;
            const int SW_MAXIMIZE = 3;

            Process[] processes = Process.GetProcessesByName("hsscp"); // Tìm các tiến trình

            if (processes.Length > 0)
            {
                AutoMouseClick(_CloseButton, 3000, 1);
                Thread.Sleep(3000);
                OpenApp("C:\\Program Files\\Mozilla Firefox\\firefox.exe");
                Thread.Sleep(6000);
                AutoMouseClick(_CheckHuman);
                Thread.Sleep(3000);
                for (int i = 0; i < processes.Length; i++)
                {
                    IntPtr hWnd = processes[i].MainWindowHandle; // Lấy handle của cửa sổ chính của Hotspot
                    ShowWindow(hWnd, SW_NORMAL);// Hiện cứa sổ ở chế độ Normal
                    SetWindowPos(hWnd, IntPtr.Zero, 0, 0, 647, 607, SWP_SHOWWINDOW);
                    SetForegroundWindow(hWnd);
                    Thread.Sleep(2000);
                    AutoMouseClick(Turnoff);
                    Thread.Sleep(2000);
                    AutoMouseClick(Turnoff);
                    Thread.Sleep(20000);
                    AutoMouseClick(Turnon);
                    Thread.Sleep(2000);

                    SetWindowPos(hWnd, IntPtr.Zero, 1263, 0, 647, 607, SWP_SHOWWINDOW);
                    AutoMouseClick(Turnon);
                }
            }
            else
            {
                MessageBox.Show("Không tìm thấy Hostpot đang chạy.");
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

            if ((strKeyWords.Count > 0) && (KeyIndex< strKeyWords.Count))  {
                txtKeywords.Text = Convert.ToString(strKeyWords[KeyIndex]);
                txtVol.Text = strVolume[KeyIndex].ToString();
            }
        }
        private void ExportExcel(string path, int split = 0)
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
        private void QuickExportExcel(string path, int noMsg = 0, int split = 0)
        {
            maxKeyExcel = 1000000;
            //DataTable dt = new DataTable();
            using (ExcelPackage excelPackage = new ExcelPackage())
            
            if (strKeyWords.Count >= 1)
            {
                int VolKey = strKeyWords.Count;
                int numSplit = VolKey / maxKeyExcel;

                for (int iSplit = 0; iSplit <= numSplit; iSplit++)
                {

                    if (VolKey > maxKeyExcel)
                    {
                        VolKey -= maxKeyExcel;
                    }
                    else
                    {
                        maxKeyExcel = VolKey;
                    }


                    // Tạo một đối tượng ExcelPackage

                    {
                        // Tạo một đối tượng Worksheet
                        ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet" + (iSplit + 1));

                        // Lấy dữ liệu từ DataGrid và đổ vào worksheet
                        for (int i = 0; i < maxKeyExcel; i++)
                        {
                            worksheet.Cells[i + 1, 1].Value = strKeyWords[i];
                            worksheet.Cells[i + 1, 2].Value = strVolume[i];
                            worksheet.Cells[i + 1, 3].Value = strCheck[i];
                            worksheet.Cells[i + 1, 4].Value = strComp[i];
                                worksheet.Cells[i + 1, 5].Value = strKeyCha1[i];
                                worksheet.Cells[i + 1, 6].Value = strKeyCha2[i];
                                worksheet.Cells[i + 1, 7].Value = strKeyCha3[i];
                                //worksheet.Cells[i + 1, 7].Value = strLinkShopee[i];
                            }

                            if (iSplit == 0)
                            {
                                worksheet.Cells[1, 27].Value = KeyIndex;  // Ghi lại KeyIndex để tiếp tục xử lý tại ô "AA:1"
                                worksheet.Cells[2, 27].Value = txtSuggestKey.Text;
                                worksheet.Cells[3, 27].Value = txtNegativeKey.Text;
                                worksheet.Cells[4, 27].Value = txttaikhoan.Text;
                                worksheet.Cells[5, 27].Value = txtmatkhau.Text;
                            }

                       
                    }
                }

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

        private void tmrPlanGetKeyword(object sender, EventArgs e)   //tmr01 - Kịch bản lấy keywords - Bỏ
        {
            
            alarmCounter++;

            if (alarmCounter == StepTimer[0]) //focus text search
            {

                pt.X = 382;
                pt.Y = 486;
                //Cursor.Position = pt;
                //mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                AutoMouseClick(pt, 100);

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
                //Cursor.Position = pt;
                //mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                AutoMouseClick(pt, 100);

                progressBar1.Value += 1;
            }

            if (alarmCounter == StepTimer[3]) // click download button
            {
                pt.X = 971;
                pt.Y = 937;
                //Cursor.Position = pt;
                //mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                AutoMouseClick(pt, 100);

                progressBar1.Value += 1;
            }

            if (alarmCounter == StepTimer[4]) // click export to excel
            {
                pt.X = 918;
                pt.Y = 751;
                //Cursor.Position = pt;
                //mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                AutoMouseClick(pt, 100);

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

        private void tmrPlanLogin(object sender, EventArgs e)  // tmrPlan02 Kịch bản Login (kịch bản 01)
        
        {
            alarmCounter++;
                if (alarmCounter == StepTimer01[0]) //CLICK TO link
                {
                    pt.X = 610;
                    pt.Y = 60;
                //Cursor.Position = pt;
                //mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                AutoMouseClick(pt, 100);

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
                //Cursor.Position = pt;
                //mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                AutoMouseClick(pt, 100);

                progressBar1.Value += 1;
                }

                if (alarmCounter == StepTimer01[4]) //CLICK TO Account Button ở góc
                {
                    pt.X = _AccountMenu.X;
                    pt.Y = _AccountMenu.Y;
                //Cursor.Position = pt;
                //mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                AutoMouseClick(pt, 100);
                progressBar1.Value += 1;
                }

                if (alarmCounter == StepTimer01[5]) //CLICK TO Account Button ở góc
                {
                    pt.X = _LoginMenu.X;
                    pt.Y = _LoginMenu.Y;
                //Cursor.Position = pt;
                //mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                AutoMouseClick(pt, 100);
                progressBar1.Value += 1;
                }

                if (alarmCounter == StepTimer01[6]) //CLICK TO Account Button ở góc
                {
                    pt.X = _AccountButtonLogin.X;
                    pt.Y = _AccountButtonLogin.Y;
                //Cursor.Position = pt;
                //mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                AutoMouseClick(pt, 100);
                progressBar1.Value += 1;
                }

                if (alarmCounter == StepTimer01[7]) //CLICK TO Account Button ở góc
                {
                    pt.X = _LoginMenu.X;
                    pt.Y = _LoginMenu.Y;
                    Cursor.Position = pt;
                    mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);

                    progressBar1.Value += 1;
                }


           if (alarmCounter == StepTimer01[8]) //Click Vào ô text nhập Email
            {
                pt.X = _EmailTextLogin.X;
                pt.Y = _EmailTextLogin.Y;
                //Cursor.Position = pt;
                //mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                AutoMouseClick(pt, 100);
                progressBar1.Value += 1;
            }

            if (alarmCounter == StepTimer01[9]) //input to tai khoan
                {
                AccountIndex++;
                if(AccountIndex >= Account.Length-1) AccountIndex = 0;
                SendKeys.Send(Account[AccountIndex]);

                    progressBar1.Value += 1;
                }
            if (alarmCounter == StepTimer01[10]) //focus text mat khau
                {
                    pt.X = _PassTextLogin.X;
                    pt.Y = _PassTextLogin.Y;
                // Cursor.Position = pt;
                //mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                AutoMouseClick(pt, 100);
                progressBar1.Value += 1;
                }
             if (alarmCounter == StepTimer01[11]) //input text mat khau
                {
                    SendKeys.Send(passLogin);
                    progressBar1.Value += 1;
                }
             if (alarmCounter == StepTimer01[12]) // click login
                {
                    //pt.X = _LogoutButtonLogin.X;
                    //pt.Y = _LogoutButtonLogin.Y;
                    //Cursor.Position = pt;
                    //mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                    SendKeys.Send("{ENTER}");
                    progressBar1.Value += 1;
                }
             if (alarmCounter == StepTimer01[13]) //focus text search
                {
                    pt.X = 382;
                    pt.Y = 486;
                //Cursor.Position = pt;
                //mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                AutoMouseClick(pt, 100);
                progressBar1.Value += 1;
                }
             if (alarmCounter == StepTimer01[14]) //input text search
                {
                    SendKeys.Send(txtKeywords.Text);

                    progressBar1.Value += 1;
                }

             if (alarmCounter == StepTimer01[15]) // click search
                {
                    pt.X = _ButtonSearch04.X;
                    pt.Y = _ButtonSearch04.Y;
                //Cursor.Position = pt;
                //mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                AutoMouseClick(pt, 100);
                progressBar1.Value += 1;
                }

                if (alarmCounter == StepTimer01[16])
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

                    int scrollAmount = 120; // Giá trị dương để scroll lên, giá trị âm để scroll xuống
                    int scrollTimes = 10; // Số lần scroll

                    // Lặp lại việc scroll lên cho đến khi đạt đến đầu trang
                    //SendKeys.Send("{PGUP}");
                    //Thread.Sleep(500);

                    for (int i = 0; i < scrollTimes; i++)
                    {
                        pt.X = 1190;
                        pt.Y = 150;
                        AutoMouseClick(pt, 100);
                        mouse_event(MOUSEEVENTF_WHEEL, 1190, 150, scrollAmount, 0);
                        Thread.Sleep(1000); // Thêm độ trễ nhỏ để đảm bảo scroll được thực hiện
                    }

                    AutoMouseClick(_CheckHuman, 100); // Click vào cloudflare phía trên
                    Thread.Sleep(1000);
                    //pt.X = 982;
                    //pt.Y = 238;
                    pt = _PointHeaderAccLogout;
                    //Cursor.Position = pt;
                    //mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                    AutoMouseClick(pt, 100);
                    Thread.Sleep(1000);


                    //pt.X = 419;
                    //pt.Y = 238;
                    pt = _TextSearch03;
                    //Cursor.Position = pt;
                    //mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                    AutoMouseClick(pt, 100);
                    progressBar1.Value += 1;

                    CurDownload = 0; // Khởi tạo lại cho biến đếm download. Chú ý chỗ này
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
                            //FocusCurrentCell(dgrListKeywords, KeyIndex - 1);
                            if (KeyIndex % InitVar.v_AutoSave == 0) //Sau 5000 keywords sẽ lưu tạm lại
                            {
                                string strTempPath = "";
                                DateTime now = DateTime.Now;
                                string formattedDateTime = now.ToString("dd/MM/yyyy HH:mm:ss");
                                string downloadPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\Downloads\";
                                strTempPath = downloadPath + "KeyTool" + formattedDateTime.Replace("/","").Replace(":","") + ".xlsx";
                                QuickExportExcel(strTempPath, 1);
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
                                txtVol.Text = strVolume[KeyIndex - 1].ToString();
                                LevelSearch = Convert.ToInt32(strCheck[KeyIndex - 1].ToString());
                                // THoang 18:59 20230301
                               txtCur.Text = Convert.ToString(KeyIndex);
                            }
                            //THoang 21:56 20230303
                            if (Convert.ToInt32(strVolume[KeyIndex - 1]) <= InitVar.v_VolMax)
                            {
                                strCheck[KeyIndex - 1] = "-1";
                            }

                            //THoang 21:56 20230303
                        } while ((Convert.ToInt32(strVolume[KeyIndex - 1]) <= InitVar.v_VolMax) || (Convert.ToInt32(strCheck[KeyIndex - 1]) >= InitVar.v_LevelSearch) || (Convert.ToInt32(strComp[KeyIndex - 1]) >= InitVar.v_LevelDif)); // Chi chay các keyword có vol >=v_VolMax || chưa đánh dấu 100

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

                    //Cursor.Position = pt;
                    //mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                    AutoMouseClick(pt, 100);
                    progressBar1.Value += 1;

                    // Xử lý chuyển tab sang và đưa địa chỉ vào.
                    // Thử sử dụng delay xem có đạt kết quả như mong muốn không
                    Thread.Sleep(7000);
                    SendKeys.Send("^l");
                    Thread.Sleep(1000);
                    SendKeys.Send("^c");
                    Thread.Sleep(1000);
                    string tempClipBoard = Clipboard.GetText();
                    string OriginAddress = "https://keywordtool.io/search/keywords/google/result/65818107e636f61ebe0ff058?category=web&country=US&country_language=en&country_location=2840&keyword=test&language=en&metrics_country=US&metrics_currency=USD&metrics_is_default_location=0&metrics_is_estimated=0&metrics_language=1000&metrics_location=2840&metrics_network=2&search_type=1&time=1717200000&signature=086dadf94a2583a9192bc4d09edea152769f29826fefb252585701780cd3ecb2";
                    string OriginAddressVN = "https://keywordtool.io/search/keywords/google/result/657fe603f9e1e6c5e10a3460?category=web&country=VN&country_language=vi&country_location=2704&keyword=test&language=vi&metrics_country=VN&metrics_currency=USD&metrics_is_default_location=0&metrics_is_estimated=0&metrics_language=1040&metrics_location=2704&metrics_network=2&search_type=1&time=1717200000&signature=9043dd371914bdbc252733c0f8a53bed4da7bbbe81b3ca6b29d633a53508cf14";
                    Thread.Sleep(500);

                    ///////////////////////////////////
                    countOrigin++;
                    if (countOrigin > 20)  // 21 lần thì xử lý kịch bản Unusual
                    {
                        ByPassUnusualActivity(_Turnon, _Turnoff, _CloseButton, _CheckHuman);

                        if (cmbLanguage.SelectedIndex == 0)
                        {
                            Clipboard.SetText(OriginAddressVN);
                        }
                        else
                        {
                            Clipboard.SetText(OriginAddress);
                        }    
                        Thread.Sleep(500);
                        SendKeys.Send("^v");
                        Thread.Sleep(500);
                        SendKeys.Send("{ENTER}");
                        Thread.Sleep(5000);
                        countOrigin = 0;
                    }
                    ///////////////////////////////////
                    SendKeys.Send("{ESC}");
                    Thread.Sleep(1000);
                    //pt.X = 982;
                    //pt.Y = 238;

                    int scrollAmount = 120; // Giá trị dương để scroll lên, giá trị âm để scroll xuống
                    int scrollTimes = 10; // Số lần scroll

                    // Lặp lại việc scroll lên cho đến khi đạt đến đầu trang
                    for (int i = 0; i < scrollTimes; i++)
                    {
                        pt.X = 1190;
                        pt.Y = 575;
                        AutoMouseClick(pt, 100);
                        mouse_event(MOUSEEVENTF_WHEEL, 1190, 575, scrollAmount, 0);
                        Thread.Sleep(50); // Thêm độ trễ nhỏ để đảm bảo scroll được thực hiện
                    }
                    // Tọa độ cloudflare ở dưới cần phải xem lại
                    //pt.X = 265;
                    //pt.Y = 580;
                    AutoMouseClick(_CheckHuman2, 1000); // Click vào cloudflare phía Dưới
                    Thread.Sleep(300);
                    AutoMouseClick(_KeySug, 1000); // Click vào cloudflare phía Dưới
                    Thread.Sleep(300);
                    // toa độ nút thích share trên các mạng xã hội
                    pt.X = 762;
                    pt.Y = 139;
                    AutoMouseClick(pt, 1000);
                    Thread.Sleep(1000);

                    pt = _PointHeaderAccLogin;
                    //Cursor.Position = pt;
                    //mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                    AutoMouseClick(pt, 100);
                    Thread.Sleep(1000);
                    SendKeys.Send("^l");
                    Thread.Sleep(1000);
                    Clipboard.SetText(tempClipBoard);  // Lấy lại link adress của keyword;
                    SendKeys.Send("^v");
                    Thread.Sleep(1000);
                    SendKeys.Send("{ENTER}");
                    Thread.Sleep(1000);
                }

                if (alarmCounter == StepTimer02[4]) // click download button
                {
                    //pt.X = 971;
                    //pt.Y = 937;
                    //if (LoadFileExcelOK == -1)
                    //{
                    //    alarmCounter--;
                    //}
                    //else
                    //{
                        pt = _ButtonDownload03;

                        //Cursor.Position = pt;

                    //mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                    AutoMouseClick(pt, 100);
                    progressBar1.Value += 1;
                    //}
                }

                if (alarmCounter == StepTimer02[5]) // click export to excel
                {
                    pt = _ButtonExcel03;

                    //Cursor.Position = pt;
                    //mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                    AutoMouseClick(pt, 100);
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
                        LoadFileExcelOK = -1;
                        // -1: đang thực hiện đọc file Excel.
                        // 0: đọc file Excel thành công.
                        // 1: đã đọc xong file excel và không thành công.
                        cboPlan.Text = "Da tim thay file";
                        ImportExcelCircle(filePath, txtKeywords.Text);   
                        // THoang 18:59 20230303
                        if (LoadFileExcelOK == 0)
                        {
                            txtTotal.Text = Convert.ToString(strKeyWords.Count);
                            strCheck[KeyIndex - 1] = InitVar.v_LevelSearch + LevelSearch; //"100"; // Giá trị cao vượt qua LevelSearch
                            //MessageBox.Show("Nhap file thanh cong");
                            KeyIndex++;
                            d_errorfile = 0; // Khi tim thay file tra error ve =0, de chay vong kich ban 03.
                            if (countNextKey > 0)
                            {
                                KeyIndex = KeyIndex_backup;   //quay trở lại vị trí key đã quét mà bỏ qua trước đó
                                // Đếm số lần quay lại.
                                countBack++;
                            }
                            countNextKey = 0;
                            //SendKeys.Send("%{TAB}");
                        }
                        else
                        {
                            KeyIndex--;
                            //FocusCurrentCell(dgrListKeywords, KeyIndex - 1);

                            d_errorfile++; // Tăng số lần đếm không lấy được file: 1-chạy 04; 2-chạy 05
                            //// Thinhking thêm một chút kịch bản, chuẩn bị code chỗ này cho lặp MaxDownload.
                            
                            if(d_errorfile == 1)
                            {
                                cboPlan.Text = "Click lại Download";
                                Thread.Sleep(InitVar.v_TimeDownload2 * InitVar.v_speed);
                                alarmCounter = StepTimer02[4] - 2;
                                progressBar1.Value -= 2;
                                KeyIndex++;
                                //FocusCurrentCell(dgrListKeywords, KeyIndex - 1);
                                CurDownload++;
                                cboPlan.Text = "Down lần" + CurDownload;
                                if (CurDownload < InitVar.v_MaxCountDownload)
                                {
                                    d_errorfile -= 1;
                                }
                                else
                                {
                                    //MessageBox.Show("Đã download " + CurDownload.ToString() + " vẫn chưa được!");
                                    NextKeyFunc();
                                }

                            }
                            else
                            {
                                //Download Max lần rồi vẫn ko được, có cần thông báo hay next key không??
                            }    
                            //else if (d_errorfile == 2)   //if (d_errorfile == 1 || d_errorfile == 2)
                            //{
                            //    tmrPlan03.Stop();

                            //    progressBar1.Value = 0;
                            //    progressBar1.Maximum = max_Process_Plan04;

                            //    alarmCounter = 0;
                            //    tmrPlan04.Start();
                            //    cboPlan.Text = "Kịch bản LoadExcel";
                            //}
                            ////else if (d_errorfile == 3) // gốc =3 rào lại trường hợp ==2 ở trên, đổi giá trị so sánh thành 2  // vao truong hop d_errorfile = 2 (xu ly tiep neu muon lon hon 2)
                            ////{
                            ////    tmrPlan03.Stop();
                            ////    progressBar1.Value = 0;
                            ////    progressBar1.Maximum = 9;
                            ////    alarmCounter = 0;
                            ////    tmrPlan05.Start();
                            ////    cboPlan.Text = "Kịch chưa xác định";
                            ////}
                            //else
                            //{
                            //    if (countNextKey == 0) KeyIndex_backup = KeyIndex;
                            //    countNextKey++;
                            //    KeyIndex++;
                            //    FocusCurrentCell(dgrListKeywords, KeyIndex - 1);
                            //    if (countNextKey >=3)
                            //    {
                            //        tmrPlan03.Stop();
                            //        alarmCounter = 0;
                            //        progressBar1.Value = 0;
                            //         d_errorfile = 0;
                            //        KeyIndex = KeyIndex_backup; // khi gặp lỗi 3 lần, cũng sẽ quay lại, hoặc khi loadOK excel
                                    
                            //        //Viết lại kịch bản Login vào đây (chú ý, có 2 chỗ copy ở dưới nữa). Kiểm tra lại kịch bản 
                            //        tmrPlan02.Start();
                            //        cboPlan.Text = "Kịch bản Login";
                            //    }
                            //    else
                            //    {
                            //        d_errorfile = 0;
                            //        btnStart_Click(btnStart, EventArgs.Empty);  // Stop search key
                            //        //btnNextKey_Click(btnNextKey, EventArgs.Empty);
                            //        NextKeyFunc();
                            //        //btnStart.Text = "Start";
                            //        if (KeyIndex >= strKeyWords.Count)  // Bỏ qua keywords cuối cùng, thinking tiếp nếu list có vol tận cuối cùng
                            //        {
                            //            // THoang 22:59 20230302
                            //            tmrPlan03.Stop();
                            //            alarmCounter = 0;
                            //            MessageBox.Show("Hoàn thành chiến dịch. Vui lòng Export File"); //Sau chuyển vào Label Trạng thái                        
                            //        }
                            //        else
                            //        {
                            //            cboPlan.SelectedIndex = 2;
                            //            btnStart_Click(btnStart, EventArgs.Empty); // Restart
                            //        }
                            //    }  
                            //}
                        }
                    }
                    else // của if(System.IO.File.Exists(filePath) == true)
                    {
                        // THoang 22:59 20230302

                        KeyIndex--;
                        //FocusCurrentCell(dgrListKeywords, KeyIndex - 1);

                        d_errorfile++; // Tăng số lần đếm không lấy được file: 1-chạy 04; 2-chạy 05

                        if (d_errorfile == 1)
                        {
                            cboPlan.Text = "Click lại Download";
                            Thread.Sleep(InitVar.v_TimeDownload2 * InitVar.v_speed);
                            alarmCounter = StepTimer02[4]-2;
                            progressBar1.Value -= 2;
                            KeyIndex++;
                            //FocusCurrentCell(dgrListKeywords, KeyIndex - 1);
                            CurDownload++;
                            cboPlan.Text = "Down lần" + CurDownload;
                            if (CurDownload < InitVar.v_MaxCountDownload)
                            {
                                d_errorfile -= 1;
                            }
                            else
                            {
                                NextKeyFunc();
                            }
                        }
                        //else if (d_errorfile == 2)    //(d_errorfile == 1 || d_errorfile == 2)
                        //{
                        //    tmrPlan03.Stop();

                        //    progressBar1.Value = 0;
                        //    progressBar1.Maximum = max_Process_Plan04;

                        //    alarmCounter = 0;
                        //    tmrPlan04.Start();
                        //    cboPlan.Text = "Kịch bảimpn LoadExcel";
                        //}
                        ////else if (d_errorfile == 3)  //gốc ==3 đổi thành ==2 khi rào trường hợp ==2 ở trên vao truong hop d_errorfile = 2 (xu ly tiep neu muon lon hon 2)
                        ////{
                        ////    tmrPlan03.Stop();
                        ////    progressBar1.Value = 0;
                        ////    progressBar1.Maximum = max_Process_Plan04;
                        ////    alarmCounter = 0;
                        ////    tmrPlan05.Start();
                        ////    alarmCounter = 0;
                        ////    cboPlan.Text = "Kịch bản chưa xác đinh";
                        ////}
                        //else
                        //{
                        //    if (countNextKey == 0) KeyIndex_backup = KeyIndex;
                        //    countNextKey++;
                        //    KeyIndex++;
                        //    FocusCurrentCell(dgrListKeywords, KeyIndex - 1);
                        //    if (countNextKey >= 3)
                        //    {
                        //        tmrPlan03.Stop();
                        //        alarmCounter = 0;
                        //        progressBar1.Value = 0;
                        //        //DialogResult r = MessageBox.Show("Lỗi nghiêm trọng! đã bỏ qua 3 keys vẫn chưa được", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        //        d_errorfile = 0;
                        //        KeyIndex = KeyIndex_backup; // khi gặp lỗi 3 lần, cũng sẽ quay lại, hoặc khi loadOK excel
                        //        //btnStart.Text = "Start";
                        //        //cboPlan.SelectedIndex = 2;  // Lựa chọn sẵn kịch bản 03
                        //        //tmrPlan01.Stop();
                        //        //tmrPlan02.Stop();
                        //        //tmrPlan03.Stop();
                        //        //tmrPlan04.Stop();
                        //        //tmrPlan05.Stop();
                        //        //btnStart_Click(btnStart, EventArgs.Empty);
                        //        tmrPlan02.Start();
                        //        cboPlan.Text = "Kịch bản Login";
                        //    }
                        //    else
                        //    {
                        //        d_errorfile = 0;
                        //        btnStart_Click(btnStart, EventArgs.Empty); // thực hiện stop
                        //        //btnNextKey_Click(btnNextKey, EventArgs.Empty);
                        //        NextKeyFunc();
                        //        //btnStart.Text = "Start";
                        //        if (KeyIndex >= strKeyWords.Count)  // Bỏ qua keywords cuối cùng, thinking tiếp nếu list có vol tận cuối cùng
                        //        {
                        //            // THoang 22:59 20230302
                        //            tmrPlan03.Stop();
                        //            alarmCounter = 0;
                        //            MessageBox.Show("Hoàn thành chiến dịch. Vui lòng Export File"); //Sau chuyển vào Label Trạng thái                        
                        //        }
                        //        else
                        //        {
                        //            cboPlan.SelectedIndex = 2;
                        //            btnStart_Click(btnStart, EventArgs.Empty); // Restart
                        //        }
                        //    }
                        //}
                        else
                        {
                            // Chỗ thứ 2 có cần thông báo hay nextkey không?
                        }
                    }
                    alarmCounter +=1;
                    //this.WindowState = FormWindowState.Normal;
                    progressBar1.Value += 1;
                }

                if (alarmCounter == StepTimer02[7])
                {
                    alarmCounter = 0;

                    //this.WindowState = FormWindowState.Normal;

                    progressBar1.Value = 0;
                }
            }
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //} 
        }  //Kịch bản 03 - Quan trọng nhất

        private void tmrPlanLoadLaiExcel(object sender, EventArgs e) // tmrPlan04 Kịch bản 04
        {
            alarmCounter++;

            if (alarmCounter == StepTimer03[0]) //CLICK TO link
            {
                pt.X = 610;
                pt.Y = 60;
                //Cursor.Position = pt;
                //mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                AutoMouseClick(pt, 100);
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

                //Cursor.Position = pt;
                //mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                AutoMouseClick(pt, 100);

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
                //pt = _ButtonSearch04;

                //Cursor.Position = pt;
                //mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                SendKeys.Send("{ENTER}");

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

        private void tmrPlanXen(object sender, EventArgs e)  // tmrPlan05 Kịch bản 05
        {
            alarmCounter++;
            if (alarmCounter == StepTimer04[0]) //focus text taikhoan
            {
                pt.X = 641;
                pt.Y = 376;
                //Cursor.Position = pt;
                //mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                AutoMouseClick(pt, 100);
                progressBar1.Value += 1;
            }
            if (alarmCounter == StepTimer04[1]) //input to tai khoan
            {
                SendKeys.Send(Account[AccountIndex]);

                progressBar1.Value += 1;
            }
            if (alarmCounter == StepTimer04[2]) //focus text mat khau
            {
                pt.X = 641;
                pt.Y = 461;
                //Cursor.Position = pt;
                //mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                AutoMouseClick(pt, 100);
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
                //Cursor.Position = pt;
                //mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                AutoMouseClick(pt, 100);
                progressBar1.Value += 1;
            }
            if (alarmCounter == StepTimer04[5]) //focus text search
            {
                pt.X = 382;
                pt.Y = 486;
                //Cursor.Position = pt;
                //mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                AutoMouseClick(pt, 100);
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
                //Cursor.Position = pt;
                //mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, pt.X, pt.Y, 0, 0);
                AutoMouseClick(pt, 100);
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

        private void NextKeyFunc()
        {
            if (KeyIndex <= strKeyWords.Count - 1)
            {
                KeyIndex++;
                //FocusCurrentCell(dgrListKeywords, KeyIndex - 1);
                if (countNextKey == 0 || countBack >= 2)  // đếm 2 lần quay lại key ko được, coi như key hỏng và bỏ qua hẳn
                {
                    strCheck[KeyIndex - 1] = LevelSearch; //"100"; //Giá trị cao, mặc định là 100 để vượt qua Level
                    strVolume[KeyIndex - 1] = "0"; // Đưa Vol về = 0, chưa hiểu ý thầy Bình chỗ này!
                    countBack = 0; // đếm lại
                }
                txtKeywords.Text = Convert.ToString(strKeyWords[KeyIndex]);
                txtVol.Text = strVolume[KeyIndex].ToString();
            }
            else
            {
                MessageBox.Show("Đã hết số lượng KeyWords");
            }
        }

        private void btnNextKey_Click(object sender, EventArgs e)
        {
            // Xóa phần tử khi bấm nút NextKey
            strKeyWords.RemoveAt(KeyIndex -1);
            strVolume.RemoveAt(KeyIndex - 1);
            strCheck.RemoveAt(KeyIndex - 1);
            strComp.RemoveAt(KeyIndex - 1);

            //dtKeywords.Rows.RemoveAt(KeyIndex - 1);
            //dgrListKeywords.DataSource = dtKeywords;
            //FocusCurrentCell(dgrListKeywords, KeyIndex-1);

            //if (KeyIndex <= dgrListKeywords.Rows.Count - 1)
            //{
            //    KeyIndex++;
            //    FocusCurrentCell(dgrListKeywords, KeyIndex - 1);
            //    if (countNextKey == 0 || countBack >= 2)  // đếm 2 lần quay lại key ko được, coi như key hỏng và bỏ qua hẳn
            //    {
            //        strCheck[KeyIndex - 1] = LevelSearch; //"100"; //Giá trị cao, mặc định là 100 để vượt qua Level
            //        strVolume[KeyIndex - 1] = "0"; // Đưa Vol về = 0, chưa hiểu ý thầy Bình chỗ này!
            //        countBack = 0; // đếm lại
            //    }
            //    txtKeywords.Text = Convert.ToString(strKeyWords[KeyIndex]);
            //    txtVol.Text = strVolume[KeyIndex].ToString();
            //}
            //else
            //{
            //    MessageBox.Show("Đã hết số lượng KeyWords");
            //}    
        }

        private void FocusCurrentCell(DataGridView dataGridView, int curRow)
        {
            if (dataGridView.CurrentRow != null && dataGridView.CurrentCell != null)
            {
                // Lưu trữ vị trí ô cell hiện tại
                int currentRowIndex = curRow;
                int currentColumnIndex = 0;
                if (curRow < dataGridView.Rows.Count)
                {
                    // Di chuyển tiêu điểm đến ô cell hiện tại
                    dataGridView.CurrentCell = dataGridView.Rows[currentRowIndex].Cells[currentColumnIndex];
                    // Tập trung vào DataGridView để ô cell hiện tại trở thành tiêu điểm
                    dataGridView.Focus();
                    if (currentRowIndex > 3)
                    {
                        dgrListKeywords.FirstDisplayedScrollingRowIndex = currentRowIndex - 3;
                    }
                    
                }
            }
        }

        private void ghiFileCàiĐặtToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmConfig f_config = new frmConfig();
            f_config.ShowDialog();

            InitVar_Plan();
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

        private void btnTestKey_Click(object sender, EventArgs e)
        {
            SendKeys.Send("%{TAB}");
        }

        private void đọcChỉSốToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string strInfo = "";
            strInfo += "Số lượng KeyCha1" + strKeyCha1.Count + "\n";
            strInfo += "Số lượng KeyCha2" + strKeyCha2.Count + "\n";
            strInfo += "Số lượng KeyCha3" + strKeyCha3.Count + "\n";
        }

        private void cboPlan_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void cmbLanguage_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        //////////////////////////////////////////////////////////////////////
    }
}

