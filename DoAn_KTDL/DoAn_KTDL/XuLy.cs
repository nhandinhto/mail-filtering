using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.IO;

namespace DoAn_KTDL
{
    class XuLy
    {
        private static List<List<string>> listBagofNonSpam = new List<List<string>>();

        private static List<List<string>> listBagofSpam = new List<List<string>>();

        private static List<string> bagofwordTest = new List<string>();

        private static Dictionary<string, float> dictSpam = new Dictionary<string, float>();

        private static Dictionary<string, float> dictNonSpam = new Dictionary<string, float>();



        public static List<string> listSpamEmail = new List<string>();

        public static List<string> listNonSpamEmail = new List<string>();

        public static string[][] tanSoTuKep;

        public static string[][] tanSoTuDon;

        public static double P_spam;

        public static double P_nonspam;

        static int totalSpamMessage = 0;
        static int totalNonSpamMessage = 0;

        //Đọc dữ liệu từ file Excel
        public static void getData_spam()
        {
            try
            {
                string link = Application.StartupPath + "\\email_spam.xlsx";
                if (!System.IO.File.Exists(link))
                {
                    Console.WriteLine("Đường dẫn không chính xác");

                }
                else
                {
                    Excel.Application xlApp = new Excel.Application();
                    Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(link);
                    // Lấy sheeet 1 
                    Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkbook.Sheets.get_Item(1);
                    //Lấy phạm vi dữ liệu
                    Excel.Range xlRange = xlWorkSheet.UsedRange;
                    //Tạo mảng lưu trữ dữ liệu
                    object[,] valueArray = (object[,])xlRange.get_Value(Excel.XlRangeValueDataType.xlRangeValueDefault);
                    for (int rows = 1; rows <= xlWorkSheet.UsedRange.Rows.Count; rows++)
                    {
                        listSpamEmail.Add(valueArray[rows, 1].ToString());
                        totalSpamMessage++;
                    }
                    xlWorkbook.Close();
                    xlApp.Quit();
                }
            }
            catch (Exception ex)
            {
            }
        }

        public static void getData_Nonspam()
        {
            try
            {
                string link = Application.StartupPath + "\\email_nonspam.xlsx";
                //string link = @"C:\Users\Nhan\Desktop\email_nonspam.xlsx";
                if (!System.IO.File.Exists(link))
                {
                    Console.WriteLine("Đường dẫn không chính xác");

                }
                else
                {
                    Excel.Application xlApp = new Excel.Application();
                    Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(link);
                    // Lấy sheeet 1 
                    Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkbook.Sheets.get_Item(1);
                    //Lấy phạm vi dữ liệu
                    Excel.Range xlRange = xlWorkSheet.UsedRange;
                    //Tạo mảng lưu trữ dữ liệu
                    object[,] valueArray = (object[,])xlRange.get_Value(Excel.XlRangeValueDataType.xlRangeValueDefault);
                    for (int rows = 1; rows <= xlWorkSheet.UsedRange.Rows.Count; rows++)
                    {
                        listNonSpamEmail.Add(valueArray[rows, 1].ToString());
                        totalNonSpamMessage++;
                    }
                    xlWorkbook.Close();
                    xlApp.Quit();
                }
                
            }
            catch (Exception ex)
            {
            }
        }
        //Các khâu xử lý dữ liệu
        public static string RemoveDigit(string input)
        {
            string output = "";
            for (int i = 0; i < input.Length; i++)
            {
                if (char.IsLetter(input[i]) || input[i] == ' ')
                    output += input[i];
            }
            return output;
        }

        public static string RemoveSpecialChars(string input)
        {
            string[] tuitu = input.Split(new char[] { '\n', '&', ',', '.', '!', '*', '\"', '/', '(', ')', '_', '-', '@', '#', '$', '?', '&', '~', '`', '%', '^', '*', '+', '=' });
            string output = "";
            foreach (string s in tuitu)
                output += s + ".";
            return output;
        }

        public static string RemoveConnectWord(string input)
        {
            string output = "";
            string[] op = input.Split(' ');
            List<string> pattern = new List<string> { "thì", "là", "mà", "và", "các", "những", "nhưng", "tuy nhiên", "mặc dù", "vì thế", "không những", "mà còn", "dù sao", "tóm lại", "tóm tắt", "nói chung là", "nói tóm lại", "tóm gọn", "nói chung" };
            for (int i = 0; i < op.Length; i++)
            {
                if (!pattern.Exists(t => t.Equals(op[i])))
                    output += op[i] + " ";
            }
            return output;
        }

        public static string RemoveLink(string input)
        {
            string output = "";
            string[] filter = input.Split(' ');
            for (int i = 0; i < filter.Length; i++)
            {
                if (!filter[i].Contains("https") && !filter[i].Contains(".com") && !filter[i].Contains(".net") && !filter[i].Contains(".us") && !filter[i].Contains(".vn") && !filter[i].Contains(".org"))
                {
                    output += filter[i] + " ";
                }
            }
            return output;
        }


        //Hàm tách các từ đơn 
        public static List<string> tuiTu1(string dulieu)
        {
            string[] tuitu = dulieu.Split(' ');
            List<string> bagOfwords = new List<string>();
            bagOfwords = tuitu.ToList();
            return bagOfwords;
        }
        //Hamf tah tu ghep ( tu` co 2 am tiet)
        public static List<string> tuiTu2(string dulieu)
        {
            dulieu = dulieu.Trim();
            string[] tuitu = dulieu.Split(' ');
            List<string> bagOfwords = new List<string>();
            for (int i = 0; i < tuitu.Length - 1; i++)
            {
                string word = "";
                int k = i + 1;
                if (tuitu[i] != "" && tuitu[i] != " ")
                {
                    while (tuitu[k] == "" || tuitu[k] == " ")
                    {
                        k++;
                    }
                    word = tuitu[i] + " " + tuitu[k];
                    bagOfwords.Add(word);
                }

            }
            return bagOfwords;
        }

        //Tim cac tu ghep co nghia ( tu` co 2 am tiet)
        public static string[][] locTuiTuKep_xacdinhTanSo(List<string> tuiTuKep, List<string> listData,int totalMessage)
        {
            List<string> loc = tuiTuKep;
            int k = loc.Count;
            tanSoTuKep = new string[k][];
            for (int i = 0; i < k; i++)
            {
                tanSoTuKep[i] = new string[2];
            }
            for (int i = 0; i < k; i++)
            {
                tanSoTuKep[i][0] = loc[i];
            }
            for (int i = 0; i < k; i++)
            {
                string Word = tanSoTuKep[i][0];
                int count = 0;
                foreach (string s in listData)
                {
                    List<string> bagOfWords = chuyenThanhTuiDuLieu(s);
                    foreach (string t in bagOfWords)
                    {
                        if (t.Contains(Word))
                            count++;
                    }
                }
                count++;
                tanSoTuKep[i][1] = count.ToString();
            }
            for (int i = 0; i < k; i++)
            {
                tanSoTuKep[i][1] = (float.Parse(tanSoTuKep[i][1]) / totalMessage).ToString();
            }
            return tanSoTuKep;

        }
        //Xác định tần số
        public static void xacdinhTanSo_tuKep_spam(List<string> tuiTuKep, List<string> listData)
        {
            string[][] vector = locTuiTuKep_xacdinhTanSo(tuiTuKep, listData,totalSpamMessage);
            int k = vector.Length;
            for (int i = 0; i < k; i++)
            {
                if (float.Parse(vector[i][1]) >= 0)
                {
                    string data_string = vector[i][0];
                    float tanSo = float.Parse(vector[i][1]);
                    if (dictSpam.ContainsKey(data_string) == false)
                        dictSpam.Add(data_string, tanSo);
                }
            }
        }

        public static void xacdinhTanSo_tuKep_Nonspam(List<string> tuiTuKep, List<string> listData)
        {
            string[][] vector = locTuiTuKep_xacdinhTanSo(tuiTuKep, listData,totalNonSpamMessage);
            int k = vector.Length;
            for (int i = 0; i < k; i++)
            {
                if (float.Parse(vector[i][1]) >= 0)
                {
                    string data_string = vector[i][0];
                    float tanSo = float.Parse(vector[i][1]);
                    if (dictNonSpam.ContainsKey(data_string) == false)
                        dictNonSpam.Add(data_string, tanSo);
                }
            }
        }

        //Tần số từ đơn
        public static void xacdinhTanSo_tuDon_spam(List<string> tuituDon, List<string> listData)
        {
            List<string> loc = tuituDon;
            int k = loc.Count;
            tanSoTuDon = new string[k][];
            for (int i = 0; i < k; i++)
            {
                tanSoTuDon[i] = new string[2];
            }
            for (int i = 0; i < k; i++)
            {
                tanSoTuDon[i][0] = loc[i];
            }
            for (int i = 0; i < k; i++)
            {
                string Word = tanSoTuDon[i][0];
                int count = 0;
                foreach (string s in listData)
                {
                    List<string> bagOfWords = chuyenThanhTuiDuLieu(s);
                    bool found = false;
                    foreach (string t in bagOfWords)
                    {
                        if (t.Contains(Word))
                            found = true;
                    }
                    if (found == true)
                    {
                        count++;
                    }
                }
                count++;
                tanSoTuDon[i][1] = count.ToString();
            }
            int totalMessage = listData.Count + 1;
            for (int i = 0; i < k; i++)
            {
                tanSoTuDon[i][1] = (float.Parse(tanSoTuDon[i][1]) / totalMessage).ToString();
            }
            for (int i = 0; i < k; i++)
            {
                string data_string = tanSoTuDon[i][0];
                float tanSo = float.Parse(tanSoTuDon[i][1]);
                if (dictSpam.ContainsKey(data_string) == false)
                    dictSpam.Add(data_string, tanSo);

            }
        }

        public static void xacdinhTanSo_tuDon_Nonspam(List<string> tuituDon, List<string> listData)
        {
            List<string> loc = tuituDon;
            int k = loc.Count;
            tanSoTuDon = new string[k][];
            for (int i = 0; i < k; i++)
            {
                tanSoTuDon[i] = new string[2];
            }
            for (int i = 0; i < k; i++)
            {
                tanSoTuDon[i][0] = loc[i];
            }
            for (int i = 0; i < k; i++)
            {
                string Word = tanSoTuDon[i][0];
                int count = 0;
                foreach (string s in listData)
                {
                    List<string> bagOfWords = chuyenThanhTuiDuLieu(s);
                    bool found = false;
                    foreach (string t in bagOfWords)
                    {
                        if (t.Contains(Word))
                            found = true;
                    }
                    if (found == true)
                    {
                        count++;
                    }
                }
                count++;
                tanSoTuDon[i][1] = count.ToString();
            }
            int totalMessage = listData.Count + 1;
            for (int i = 0; i < k; i++)
            {
                tanSoTuDon[i][1] = (float.Parse(tanSoTuDon[i][1]) / totalMessage).ToString();
            }
            for (int i = 0; i < k; i++)
            {
                string data_string = tanSoTuDon[i][0];
                float tanSo = float.Parse(tanSoTuDon[i][1]);
                if (dictNonSpam.ContainsKey(data_string) == false)
                    dictNonSpam.Add(data_string, tanSo);

            }
        }

        //Chuyển dữ liệu thành túi các từ
        public static List<string> chuyenThanhTuiDuLieu(string dulieu)
        {
            dulieu = RemoveLink(dulieu);
            dulieu = RemoveSpecialChars(dulieu);
            string[] tuitu = dulieu.Split('.');
            List<string> toBagofWord = new List<string>();
            for (int i = 0; i < tuitu.Length; i++)
            {
                tuitu[i] = RemoveDigit(tuitu[i]);
                tuitu[i] = RemoveConnectWord(tuitu[i]);
                tuitu[i] = tuitu[i].Trim();
                tuitu[i] = tuitu[i].ToLower();
                if (tuitu[i] != "")
                {
                    if (!toBagofWord.Exists(t => t.Equals(tuitu[i])))
                        toBagofWord.Add(tuitu[i].Trim());
                }
            }
            return toBagofWord;
        }

        public static void xulyDuLieu_spam()
        {
            foreach (string s in listSpamEmail)
            {
                List<string> tuitu = new List<string>();
                tuitu = chuyenThanhTuiDuLieu(s);
                listBagofSpam.Add(tuitu);
            }
            List<string> tuituDon = new List<string>();
            List<string> tuituKep = new List<string>();
            for (int i = 0; i < listBagofSpam.Count; i++)
            {
                for (int j = 0; j < listBagofSpam[i].Count; j++)
                {
                    List<string> tt1 = tuiTu1(listBagofSpam[i][j]);
                    foreach (string s in tt1)
                    {
                        if (s != "")
                        {
                            if (s.Length >= 2)
                                if (!tuituDon.Exists(t => t.Equals(s)))
                                    tuituDon.Add(s);
                        }
                    }
                    List<string> tt2 = tuiTu2(listBagofSpam[i][j]);
                    foreach (string s in tt2)
                    {
                        if (s != "")
                        {
                            if (!tuituKep.Exists(t => t.Equals(s)))
                                tuituKep.Add(s);
                        }
                    }
                }
            }
            xacdinhTanSo_tuDon_spam(tuituDon, listSpamEmail);
            xacdinhTanSo_tuKep_spam(tuituKep, listSpamEmail);
            ghiFile_tudienSpam();
        }

        public static void xulyDuLieu_Nonspam()
        {
            foreach (string s in listNonSpamEmail)
            {
                List<string> tuitu = new List<string>();
                tuitu = chuyenThanhTuiDuLieu(s);
                listBagofNonSpam.Add(tuitu);
            }
            List<string> tuituDon = new List<string>();
            List<string> tuituKep = new List<string>();

            for (int i = 0; i < listBagofNonSpam.Count; i++)
            {
                for (int j = 0; j < listBagofNonSpam[i].Count; j++)
                {
                    List<string> tt1 = tuiTu1(listBagofNonSpam[i][j]);
                    foreach (string s in tt1)
                    {
                        if (s != "")
                        {
                            if (s.Length >= 2)
                                if (!tuituDon.Exists(t => t.Equals(s)))
                                    tuituDon.Add(s);
                        }
                    }
                    List<string> tt2 = tuiTu2(listBagofNonSpam[i][j]);
                    foreach (string s in tt2)
                    {
                        if (s != "")
                        {
                            if (!tuituKep.Exists(t => t.Equals(s)))
                                tuituKep.Add(s);
                        }
                    }
                }
            }
            xacdinhTanSo_tuDon_Nonspam(tuituDon, listNonSpamEmail);
            xacdinhTanSo_tuKep_Nonspam(tuituKep, listNonSpamEmail);
            ghiFile_tudienNonSpam();
        }

        public static void xulyDuLieu_Test(string test)
        {
            List<string> tuitu = chuyenThanhTuiDuLieu(test);
            List<string> tuiTuDon = new List<string>();
            List<string> tuiTuKep = new List<string>();

            for (int i = 0; i < tuitu.Count; i++)
            {
                List<string> tt1 = tuiTu1(tuitu[i]);
                foreach (string s in tt1)
                {
                    if (s != "")
                    {
                        if (s.Length >= 2)
                            tuiTuDon.Add(s);
                    }
                }
                List<string> tt2 = tuiTu2(tuitu[i]);
                foreach (string s in tt2)
                {
                    if (s != "")
                    {
                        tuiTuKep.Add(s);
                    }
                }
            }

            bagofwordTest = new List<string>();
            foreach (string s in tuiTuDon)
                bagofwordTest.Add(s);
            foreach (string s in tuiTuKep)
                bagofwordTest.Add(s);
        }
        //DỰ đoán

        public static int preDict(List<string> bagTest)
        {
            double ham = P_nonspam;
            double spam = P_spam;
            foreach (string t in bagTest)
            {
                if (dictSpam.FirstOrDefault(key => key.Key.Equals(t)).Key == null && dictNonSpam.FirstOrDefault(key => key.Key.Equals(t)).Key == null)
                {
                }
                else
                {
                    double tsSpam = 0;
                    double tsNSpam = 0;
                    try
                    {
                        tsSpam = dictSpam.FirstOrDefault(key => key.Key.Equals(t)).Value;
                    }
                    catch
                    {
                        tsSpam = 0;
                    }
                    try
                    {
                        tsNSpam = dictNonSpam.FirstOrDefault(key => key.Key.Equals(t)).Value;
                    }
                    catch
                    {
                        tsNSpam = 0;
                    }
                    if (tsNSpam == 0)
                    {
                        tsNSpam = (double)1 / (totalNonSpamMessage + 1);
                        tsSpam = (double)(tsSpam * totalSpamMessage);
                        tsSpam = (double)(tsSpam + 1) / (totalSpamMessage + 1);  
                    }
                    else if (tsSpam == 0)
                    {
                        tsSpam = (double)1 / (totalSpamMessage + 1);
                        tsNSpam =(double)( tsNSpam * totalNonSpamMessage);
                        tsNSpam = (double)(tsNSpam + 1) / (totalNonSpamMessage +1);
                    }
                    ham *= tsNSpam;
                    spam *= tsSpam;
                }
            }
            if (ham > spam)
            {
                return 1;
            }
            else
            {
                return 0;
            }
        }

        //Đọc dữ liệu
        public static void ghiFile_tudienSpam()
        {
            string filepath = Application.StartupPath + "\\dictSpam.txt";
            FileStream fs = new FileStream(filepath, FileMode.Create);
            StreamWriter sw = new StreamWriter(fs, Encoding.UTF8);
            foreach (KeyValuePair<string, float> item in dictSpam)
            {
                sw.WriteLine(item.Key);
                sw.WriteLine(item.Value.ToString());
            }
            sw.Flush();
            fs.Close();
        }

        public static void ghiFile_tudienNonSpam()
        {
            string filepath = Application.StartupPath + "\\dictNonSpam.txt";
            FileStream fs = new FileStream(filepath, FileMode.Create);
            StreamWriter sw = new StreamWriter(fs, Encoding.UTF8);
            foreach (KeyValuePair<string, float> item in dictNonSpam)
            {
                sw.WriteLine(item.Key);
                sw.WriteLine(item.Value.ToString());
            }
            sw.Flush();
            fs.Close();
        }

        public static bool docFile_tudienSpam()
        {
            try
            {
                dictSpam = new Dictionary<string, float>();
                string filepath = Application.StartupPath + "\\dictSpam.txt";
                FileStream fs = new FileStream(filepath, FileMode.Open);
                using (StreamReader sr = new StreamReader(fs, Encoding.UTF8))
                {
                    bool keyReaded = false;
                    string line;
                    string key ="";
                    float freq = 0;
                    while ((line = sr.ReadLine()) != null)
                    {
                        if (keyReaded == false)
                        {
                            key = line;
                            keyReaded = true;
                        }
                        else
                        {
                            freq = float.Parse(line);
                            keyReaded = false;
                            dictSpam.Add(key, freq);
                        }
                    }
                }
                return true;
            }
            catch
            {
                Console.WriteLine("Không thể đọc file");
                return false;
            }
            

        }

        public static bool docFile_tudienNonSpam()
        {
            try
            {
                dictNonSpam = new Dictionary<string, float>();
                string filepath = Application.StartupPath + "\\dictNonSpam.txt";
                FileStream fs = new FileStream(filepath, FileMode.Open);
                using (StreamReader sr = new StreamReader(fs, Encoding.UTF8))
                {
                    bool keyReaded = false;
                    string line;
                    string key = "";
                    float freq = 0;
                    while ((line = sr.ReadLine()) != null)
                    {
                        if (keyReaded == false)
                        {
                            key = line;
                            keyReaded = true;
                        }
                        else
                        {
                            freq = float.Parse(line);
                            keyReaded = false;
                            dictNonSpam.Add(key, freq);
                        }
                    }
                }
                return true;
            }
            catch
            {
                Console.WriteLine("Không thể đọc file");
                return false;
            }
            

        }
        //Độ chính xác của model
        public static double Accuracy()
        {
            int countCorrect = 0;
            foreach (string dulieu in listSpamEmail)
            {
                if (duDoan(dulieu) == 0)
                    countCorrect++;
            }
            foreach (string dulieu in listNonSpamEmail)
            {
                if (duDoan(dulieu) == 1)
                    countCorrect++;
            }
            return ((double)countCorrect / (totalNonSpamMessage + totalSpamMessage))*100;
        }
        //Train dữ liêu
        public bool train()
        {
            try
            {
                getData_spam();
                getData_Nonspam();
                P_spam = (float)listSpamEmail.Count / (listSpamEmail.Count + listNonSpamEmail.Count);
                P_nonspam = (float)listNonSpamEmail.Count / (listSpamEmail.Count + listNonSpamEmail.Count);
                if (docFile_tudienNonSpam() == false || docFile_tudienSpam() == false)
                {
                    xulyDuLieu_spam();
                    xulyDuLieu_Nonspam();
                }
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public static int duDoan(string test)
        {
            xulyDuLieu_Test(test);
            return preDict(bagofwordTest);
        }
 
    }
}
