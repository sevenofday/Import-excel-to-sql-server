using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Aspose.Cells;
using System.Collections;
using System.IO;
using System.Text.RegularExpressions;
using System.Net;
using System.Security.Cryptography;
using Newtonsoft.Json;
using System.Globalization;
using System.Threading;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using Newtonsoft.Json.Serialization;
using Newtonsoft.Json.Linq;

namespace GeneralExcelToSql
{
    public partial class FolderFileDialog : Form
    {
        public FolderFileDialog()
        {
            InitializeComponent();
        }

       
        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            //this.openFileDialog1.Filter = "*.*";

            if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                SourceFile.Text  = this.openFileDialog1.FileName;
                // 你的 处理文件路径代码 
            }
            
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            folderBrowserDialog1.ShowDialog();
            textBox2.Text = folderBrowserDialog1.SelectedPath;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (SourceFile.Text.Trim() == "")
            {
                MessageBox.Show("请选择源文件","提示");
                return;
            }
            if (textBox1.Text.Trim() == "")
            {
                MessageBox.Show("请输入表名", "提示");
                return;
            }
            if (textBox2.Text.Trim() == "")
            {
                MessageBox.Show("请选择输出目录", "提示");
                return;
            }
            String str = GeneraFile(SourceFile.Text.Trim(), textBox1.Text.Trim(), textBox2.Text + "\\" + textBox1.Text.Trim() + ".sql");
            MessageBox.Show(str, "提示");
        }

        /// <summary>
        /// 生成sql文件
        /// </summary>
        /// <param name="strFile">文件路径</param>
        /// <param name="TableName">生成表名</param>
        /// <param name="outputFolder">保存文件路径</param>
        /// <returns></returns>
        public String GeneraFile(String strFile, string TableName, string outputFolder)
        {
            StringBuilder strError = new StringBuilder();
            DataTable dtExcel = null;//EXECL文件数据
            StringBuilder sql = new StringBuilder();
            string str = string.Empty;
            try
            {
                if (strFile.ToUpper().EndsWith(".XLSX"))
                {
                    dtExcel = Get07ExcelData(strFile);
                }
                else if (strFile.ToUpper().EndsWith(".XLS"))
                {
                    dtExcel = GetExcelData(strFile);//获取EXECL文件数据
                }
                else
                {
                    strError.Append("对不起，你导入数据失败，失败的原因为：导入的文件格式出错");
                }
            }
            catch (Exception ex)
            {
                strError.Append("对不起，你导入数据失败，失败的原因为：" + ex.Message);
            }

            //try
            //{
            //    if (dtExcel.Rows.Count == 0)
            //    {
            //        strError.Append("对不起，你导入数据失败，失败的原因为：你导入的数据为空！");
            //    }
            //}
            //catch (Exception ex)
            //{
            //    strError.Append("对不起，你导入数据失败，失败的原因为：" + ex.Message);
            //}

            if (strError.ToString() == string.Empty)
            {
                try
                {
                    strError = Delete_Execl_Null_Data(dtExcel);//删除EXCEL导入数据的空白列
                }
                catch (Exception ex)
                {
                    strError.Append("对不起，你导入数据失败，失败的原因为：" + ex.Message);
                }
            }
            if (strError.ToString() == string.Empty)
            {
                try
                {
                    strError.Append(GetSql(dtExcel, TableName, out str)) ;//生成运行的sql
                }
                catch (Exception ex)
                {
                    strError.Append("对不起，你导入数据失败，失败的原因为：" + ex.Message);
                }
            }
            if (strError.ToString() == string.Empty)
            {
                try
                {
                    strError.Append(SaveSqlFile(str, outputFolder)); ;//写入文件中
                }
                catch (Exception ex)
                {
                    strError.Append("对不起，你导入数据失败，失败的原因为：" + ex.Message);
                }
            }
            if (strError.ToString() != string.Empty)
            {
                return strError.ToString();
            }
            else
            {
                return "生成sql脚步成功";
            }
            
        }

        /// <summary>
        /// SQL保存文件
        /// </summary>
        /// <param name="sql"></param>
        /// <param name="directory"></param>
        /// <returns></returns>
        public string SaveSqlFile(string sql, string directory)
        {
            try{
                if (File.Exists(directory))
                {
                    File.Delete(directory);
                }
                StreamWriter sw = new StreamWriter(directory, true, Encoding.Default); 
                sw.Write(sql);             //在文本末尾写入文本 
                sw.Flush();                    //清空 
                sw.Close();                    //关闭 
                return "";
            }catch(Exception ex)
            {
                return ex.ToString();
            }
            return "";
        }

        /// <summary>
        /// 生成sql脚本
        /// </summary>
        /// <param name="dtExcel"></param>
        /// <param name="TableName"></param>
        /// <param name="outputFolder"></param>
        /// <returns></returns>
        public string GetSql(DataTable dtExcel, string TableName,out string str)
        {
            try
            {
                string strField = string.Empty;
                string strDesc = string.Empty;
                string strImport = string.Empty;
                string strtemp = string.Empty;
                strField = "create table " + textBox1.Text.Trim() + " ( \n ID varchar(50) primary key, \n";
                strDesc = "EXEC sys.sp_addextendedproperty @name=N'MS_Description', @value=N'ID' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'" + textBox1.Text.Trim() + "', @level2type=N'COLUMN',@level2name=N'ID'\n GO \n";
                strImport = "delete from Tb_ExcelImport_TableField where TableName='"+textBox1.Text.Trim() +"' \n";
                string field = string.Empty;
                for (int i = 0; i < dtExcel.Columns.Count; i++)
                {
                    field += dtExcel.Columns[i].ColumnName + ";";
                    if (System.Text.Encoding.Default.GetBytes(field).Length >= 128)
                    {
                        field = field.TrimEnd(';');
                        string strError = GetEnglishBaidu(field, out strtemp);
                        if (strError != "")
                        {
                            str = "";
                            return strError;
                        }
                        string[] k = strtemp.Split(';');
                        string[] j = field.Split(';');
                        if (k.Length != j.Length)
                        {
                            str = "";
                            return "翻译生成单词和原来单词数量不一样";
                        }
                        for (int x = 0; x < k.Length; x++)
                        {
                            strField += k[x].Replace(" ", "") + " varchar(200) ,\n";
                            strDesc += "EXEC sys.sp_addextendedproperty @name=N'MS_Description', @value=N'" + j[x] + "' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'" + textBox1.Text.Trim() + "', @level2type=N'COLUMN',@level2name=N'" + k[x].Replace(" ", "") + "'\n GO \n";
                            strImport += "insert into Tb_ExcelImport_TableField(TableName,Field,ChinaField,FieldType) values('" + textBox1.Text.Trim() + "','" + k[x].Replace(" ", "") + "','" + j[x] + "','VARCHAR')\n";
                        }
                        field = "";
                    }
                    else if (i == dtExcel.Columns.Count - 1)
                    {
                        field = field.TrimEnd(';');
                        string strError = GetEnglishBaidu(field, out strtemp);
                        if (strError != "")
                        {
                            str = "";
                            return strError;
                        }
                        string[] k = strtemp.Split(';');
                        string[] j = field.Split(';');
                        if (k.Length!=j.Length)
                        {
                            str = "";
                            return "翻译生成单词和原来单词数量不一样";
                        }
                        for (int x = 0; x < k.Length; x++)
                        {
                            strField += k[x].Replace(" ", "") + " varchar(200) ,\n";
                            strDesc += "EXEC sys.sp_addextendedproperty @name=N'MS_Description', @value=N'" + j[x] + "' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'TABLE',@level1name=N'" + textBox1.Text.Trim() + "', @level2type=N'COLUMN',@level2name=N'" + k[x].Replace(" ", "") + "'\n GO \n";
                            strImport += "insert into Tb_ExcelImport_TableField(TableName,Field,ChinaField,FieldType) values('" + textBox1.Text.Trim() + "','" + k[x].Replace(" ", "") + "','" + j[x] + "','VARCHAR')\n";
                        }
                        field = "";
                    }
                }
                
                char[] replace = { ',', '\n' };
                str = strField.TrimEnd(replace) + "\n)\n" + strDesc + strImport;
                return "";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                str = "";
                return ex.ToString();
            }
            
        }

        /// <summary>
        /// baidu API接口中文翻译为英文
        /// </summary>
        /// <param name="en"></param>
        /// <param name="temp"></param>
        /// <returns></returns>
        public string GetEnglishBaidu(string en,out string temp)
        {
            try
            {
                string str = string.Empty;
                WebClient client = new WebClient();  //引用System.Net
                string fromTranslate = en.Replace("+",""); //翻译前的内容
                //Random ran = new Random();
                //string salt= ((long)(ran.NextDouble() * 9000000000) + 1000000000).ToString();
                string salt = "1435660288";
                string secretkey = "4A0ciqdsBIrbHtNX2Cba";
                string appid = "20180124000118278";

                string sign = string.Empty;
                MD5 md5 = MD5.Create(); //实例化一个md5对像
                // 加密后是一个字节类型的数组，这里要注意编码UTF8/Unicode等的选择　
                byte[] s = md5.ComputeHash(Encoding.UTF8.GetBytes(appid + fromTranslate+ salt + secretkey));
                sign=BitConverter.ToString(s).Replace("-", "").ToLower();; 
                //appid为自己的api_id，q为翻译对象，from为翻译语言，to为翻译后语言,sign为数字签名
                string url = string.Format("https://fanyi-api.baidu.com/api/trans/vip/translate?appid={0}&q={1}&from={2}&to={3}&salt={4}&sign={5}", appid, fromTranslate, "zh", "en", salt, sign);
                var buffer = client.DownloadData(url);
                string result = Encoding.UTF8.GetString(buffer);
                if (result.Contains("error_msg"))
                {
                    JObject l= (JObject)JsonConvert.DeserializeObject(result);
                    JToken k = l["error_msg"];
                    str = k.ToString().Replace("\"", "");
                    temp = "";
                    return str;
                }
                JObject o = (JObject)JsonConvert.DeserializeObject(result);
                JToken r = o["trans_result"][0]["dst"];
                str = r.ToString().Replace("\"", "");  //dst为翻译后的值
                // 首字母大写
                CultureInfo cultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture;
                TextInfo text = cultureInfo.TextInfo;
                string toLower = text.ToLower(str);
                temp = text.ToTitleCase(toLower);
                return "";
            }
            catch (Exception ex)
            {
                temp = "";
                return ex.ToString();
            }
            
        }

        #region Excel 2007 读取
        private DataTable Get07ExcelData(string localFilePath)
        {
            DataTable dt = new DataTable();
            try
            {
                dt = ExcelToDatatalbe(localFilePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show (ex.ToString());
            }
            return dt;
        }

        /// <summary>
        /// Excel 2007 文件读取
        /// </summary>
        /// <param name="fullFilename">路径</param>
        /// <returns>datatable</returns>
        public DataTable ExcelToDatatalbe(string fullFilename)//读取
        {

            Workbook book = new Workbook();
            book.Open(fullFilename);
            DataSet resultDs = new DataSet();
            foreach (Worksheet ws in book.Worksheets)
            {
                if (ws.Cells.Count > 0)
                {
                    Cells cells = ws.Cells;
                    DataTable resultDt = cells.ExportDataTable(0, 0, cells.MaxDataRow + 1, cells.MaxDataColumn + 1);
                    ArrayList list = new ArrayList();
                    for (int i = 0; i < resultDt.Columns.Count; i++)
                    {
                        string columnName = resultDt.Rows[0][i].ToString();
                        columnName = GetConvert(columnName);
                        if (list.Contains(columnName))
                        {
                            columnName += "_1";
                        }
                        list.Add(columnName);

                    }
                    for (int i = 0; i < list.Count; i++)
                    {
                        resultDt.Columns[i].ColumnName = list[i].ToString();
                    }
                    resultDt.Rows.Remove(resultDt.Rows[0]);
                    resultDt.TableName = ws.Name;
                    resultDs.Tables.Add(resultDt);

                }
            }
            return resultDs.Tables[0];
        }

        /// <summary>
        /// 获取EXECL表数据
        /// </summary>
        /// <param name="strFile"></param>
        /// <returns></returns>
        public DataTable GetExcelData(string strFile)
        {
            Workbook book = new Workbook();
            book.Open(strFile);
            DataSet resultDs = new DataSet();
            foreach (Worksheet ws in book.Worksheets)
            {
                if (ws.Cells.Count > 0)
                {
                    Cells cells = ws.Cells;
                    DataTable resultDt = cells.ExportDataTable(0, 0, cells.MaxDataRow + 1, cells.MaxDataColumn + 1);
                    ArrayList list = new ArrayList();
                    for (int i = 0; i < resultDt.Columns.Count; i++)
                    {
                        string columnName = resultDt.Rows[0][i].ToString();
                        columnName = GetConvert(columnName);
                        int j = 1;
                        while (list.Contains(columnName))
                        {
                            columnName = columnName+"_"+j.ToString();//前列名相同，重新命名
                            j++;
                        }
                        list.Add(columnName);

                    }
                    for (int i = 0; i < list.Count; i++)
                    {
                        resultDt.Columns[i].ColumnName = list[i].ToString();
                    }
                    resultDt.Rows.Remove(resultDt.Rows[0]);
                    resultDt.TableName = ws.Name;
                    resultDs.Tables.Add(resultDt);

                }
            }
            return resultDs.Tables[0];
        }
        #endregion

        #region 删除EXCEL导入数据中的空白行与列
        private StringBuilder Delete_Execl_Null_Data(DataTable ExcelDT)
        {
            StringBuilder sbErrorMsg = new StringBuilder();
            try
            {
                //删除空白列
                for (int i = 0; i < ExcelDT.Columns.Count; i++)
                {
                    bool isNull = true;

                    string CNStr = ExcelDT.Columns[i].ColumnName.ToString().Trim();

                    if (CNStr.Substring(0, 1).ToUpper() != "F")
                    {
                        isNull = false;
                    }
                    else
                    {
                        CNStr = CNStr.Substring(1);

                        try
                        {
                            Convert.ToInt32(CNStr);
                        }
                        catch
                        {
                            isNull = false;
                        }
                    }

                    if (isNull)
                    {
                        ExcelDT.Columns.RemoveAt(i);
                        i = i - 1;
                    }
                }

                //删除空白行
                for (int i = 0; i < ExcelDT.Rows.Count; i++)
                {
                    bool isNull = true;
                    for (int j = 0; j < ExcelDT.Columns.Count; j++)
                    {
                        if (ExcelDT.Rows[i][j].ToString().Trim() != string.Empty)
                        {
                            isNull = false;
                            break;
                        }
                    }

                    if (isNull)
                    {
                        ExcelDT.Rows.RemoveAt(i);
                        i = i - 1;
                    }
                }

            }
            catch (Exception ex)
            {
                sbErrorMsg.Append("对不起，删除EXCEL导入数据中的空白行列过程中出现异常，异常信息为：" + ex.Message.ToString() + "！\\n");
            }

            return sbErrorMsg;
        }
        #endregion

        /// <summary>
        /// 通过正则表达式验证是否存在特殊字符，如果存在特殊字符则将其转换为空字符
        /// </summary>
        /// <param name="str">字符串</param>
        /// <returns></returns>
        public string GetConvert(string str)
        {
            string pattern = @"/\{|\{\[|\【|\(|\（|\</";
            string replacement = "_";
            Regex rgx = new Regex(pattern);
            str = rgx.Replace(str, replacement);

            pattern = @"[^a-zA-Z0-9-_\u4e00-\u9fa5]";
            replacement = "";
            rgx = new Regex(pattern);
            str = rgx.Replace(str, replacement);
            return str;
        }
    }

}
