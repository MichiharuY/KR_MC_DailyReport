using System.Collections;
using System.Runtime.InteropServices;
using System.Data.SqlClient;
using System.ComponentModel.Design;
using static System.ComponentModel.Design.ObjectSelectorEditor;
using Microsoft.Office.Interop.Excel;
using System.Xml.Linq;
using System.Windows.Forms;
using System.IO;

namespace KR_DailyReport
{

    public partial class FormMain : Form
    {
        // DB接続文字列
        const string CONNECT_STRING = "Database=KRDairyDB;Integrated Security=SSPI;Persist Security Info=False;Connection Timeout=10";
        //const string CONNECT_STRING = "Integrated Security=SSPI;Persist Security Info=False;Connection Timeout=10";

        // 変数
        public NET10Address? m_NET10Addr = null; // 収録アドレス管理クラス
        public NET10Control? m_NET10Ctrl = null; // NET10通信管理クラス

        // NET10アドレス情報(Excelファイル読出用)
        public struct ExcelInfo
        {
            public List<int> Row;       // Excelシートの行番号
            public List<string> Item;   // 項目名
            public List<short> Addr;    // アドレス
            public List<short> Data;    // 書込データ
        }
        private string m_TemplateName = Properties.Settings.Default.TemplateName; // 設定ファイルから取得
        private string m_SavePath = Properties.Settings.Default.SavePath; // 設定ファイルから取得
        private ExcelInfo[] m_Excel = new ExcelInfo[2]; // Excel情報
        private Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
        private Microsoft.Office.Interop.Excel.Workbook? xlWorkbook = null;
        private Microsoft.Office.Interop.Excel.Worksheet? xlWorksheet = null;

        private SqlConnection m_SQLConnect; // SQL接続クラス（ADO.NET）

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public FormMain()
        {
            InitializeComponent();

            m_Excel[0] = new ExcelInfo();
            m_Excel[1] = new ExcelInfo();
        }

        /// <summary>
        /// デストラクタ
        /// </summary>
        ~FormMain()
        {
            if (xlWorksheet != null)
            {
                Marshal.ReleaseComObject(xlWorksheet);
            }

            if (xlWorkbook != null)
            {
                // Excel Close and Release
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);
            }

            // Excel Quit and Release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }

        /// <summary>
        /// フォームロード
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FormMain_Load(object sender, EventArgs e)
        {
            this.StartPosition = FormStartPosition.CenterScreen;

#if false
            ////////////////////////////////////////
            // NET10通信の初期化
            m_NET10Ctrl = new NET10Control();
            if (m_NET10Ctrl.StartComm() == false)
            {
                MessageBox.Show("NET10通信の初期化に失敗しました。", "NET10通信初期化エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
#endif
            dgvDailyData.ColumnCount = 2;

            dgvDailyData.Columns[0].Width = 270;
            dgvDailyData.Columns[0].HeaderText = "項目名";

            dgvDailyData.Columns[1].Width = 150;
            dgvDailyData.Columns[1].HeaderText = "値";

            dtpDate.Value = DateTime.Now;
        }

        /// <summary>
        /// パラメータのコンボリストを作成する
        /// </summary>
        private void MakeParamCombo()
        {
            lblStatus.Text = "Excelファイル読込中...";

            // パラメータ名、アドレス、設定値、Excelのセル位置
            //ExcelRead(m_ExcelPath, 1, 2, 2, 200, 5);

            lblStatus.Text = "";
        }

        private void FormMain_FormClosing(object sender, FormClosingEventArgs e)
        {
#if false

            if (m_NET10Ctrl != null)
            {
                if (m_NET10Ctrl.EndComm() == false)
                {
                    MessageBox.Show("NET10の通信停止に失敗しました。", "NET10通信停止エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
#endif
        }

        /// <summary>
        /// Excelへ日報出力
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnExcelOutput_Click(object sender, EventArgs e)
        {
            if (dgvDailyData.Rows .Count == 0)
            {
                MessageBox.Show("データが抽出されていません", "データ未選択", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (MessageBox.Show("Excelへ日報を出力します。よろしいですか？\n" + "(保存先：" + m_SavePath + ")", "日報出力", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {

                var data = new ArrayList();
                for (int i = 0; i < dgvDailyData.RowCount; i++)
                {
                    var cell = dgvDailyData[1, i].Value;

                    if (cell == null || cell.ToString() == "")
                    {
                        data.Add("");
                    }
                    else
                    {
                        data.Add(dgvDailyData[1, i].Value.ToString());
                    }
                }

                ExcelWrite(m_SavePath, 1, 3, 4, data, dtpDate.Value.ToString());

                MessageBox.Show("日報出力が完了しました。", "完了確認", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private string MakeSavePath()
        {
            //SaveFileDialogクラスのインスタンスを作成
            SaveFileDialog sfd = new SaveFileDialog();

            //はじめのファイル名を指定する
            //はじめに「ファイル名」で表示される文字列を指定する
            sfd.FileName = "日報.xlsx";
            //はじめに表示されるフォルダを指定する
            sfd.InitialDirectory = m_SavePath;
            //[ファイルの種類]に表示される選択肢を指定する
            //指定しない（空の文字列）の時は、現在のディレクトリが表示される
            sfd.Filter = "Excelファイル(*.xlsx)|*.xlsx|すべてのファイル(*.*)|*.*";
            //[ファイルの種類]ではじめに選択されるものを指定する
            //2番目の「すべてのファイル」が選択されているようにする
            sfd.FilterIndex = 2;
            //タイトルを設定する
            sfd.Title = "保存先のファイルを選択してください";
            //ダイアログボックスを閉じる前に現在のディレクトリを復元するようにする
            sfd.RestoreDirectory = true;
            //既に存在するファイル名を指定したとき警告する
            //デフォルトでTrueなので指定する必要はない
            sfd.OverwritePrompt = true;
            //存在しないパスが指定されたとき警告を表示する
            //デフォルトでTrueなので指定する必要はない
            sfd.CheckPathExists = true;

            //ダイアログを表示する
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                //OKボタンがクリックされたとき、選択されたファイル名を表示する
                Console.WriteLine(sfd.FileName);
            }

            return sfd.FileName;
        }

#if false
        /// <summary>
        /// エクセルファイルの指定したシートを2次元配列に読み込む.
        /// </summary>
        /// <param name="filePath">エクセルファイルのパス</param>
        /// <param name="sheetIndex">シートの番号 (1, 2, 3, ...)</param>
        /// <param name="startRow">最初の行 (>= 1)</param>
        /// <param name="startColmn">最初の列 (>= 1)</param>
        /// <param name="lastRow">最後の行</param>
        /// <param name="lastColmn">最後の列</param>
        /// <returns>シート情報を格納した2次元文字配列. ただしファイル読み込みに失敗したときには null.</returns>
        public ArrayList ExcelRead(string filePath, int sheetIndex,
                              int startRow, int startColmn,
                              int lastRow, int lastColmn)
        {
            var arrOut = new ArrayList();

            m_Excel[sheetIndex - 1].Row = new List<int>();
            m_Excel[sheetIndex - 1].Item = new List<string>();
            m_Excel[sheetIndex - 1].Addr = new List<short>();
            m_Excel[sheetIndex - 1].Data = new List<short>();

            // ワークブックを開く
            if (!ExcelOpen(filePath)) { return arrOut; }

            if (xlWorkbook != null)
            {
                xlWorksheet = xlWorkbook.Sheets[sheetIndex];
                xlWorksheet.Select();


                for (int r = startRow; r <= lastRow; r++)
                {
                    // 一行読み込む
                    var row = new ArrayList();
                    for (int c = startColmn; c <= lastColmn; c++)
                    {
                        var cell = xlWorksheet.Cells[r, c];

                        if (cell == null || cell.Value == null)
                        {
                            row.Add("");
                        }
                        else
                        {
                            row.Add(cell.Value);
                        }
                    }

                    string item = "";
                    short addr = 0;

                    // 項目名
                    if (row[0] != null && row[0] != "")
                    {
                        item = row[0].ToString();
                    }
                    else
                    {
                        // 読込終了
                        break;
                    }

                    // アドレス
                    if (row[3] != null)
                    {
                        if (row[3].ToString() != "")
                        {
                            addr = Convert.ToInt16(row[3].ToString().Substring(1, 4), 16);
                        }
                    }

                    short data = 0;
                    if (sheetIndex == 1)
                    {
                        // 作業データ
                        if (row[2] != null)
                        {
                            if (row[2].ToString() != "")
                            {
                                data = Convert.ToInt16(row[2].ToString(), 10);
                            }
                        }
                    }
                    else
                    {
                        // 日報データ
                        if (row[2] != null)
                        {
                            if (row[2].ToString() != "")
                            {
                                data = Convert.ToInt16(row[2].ToString(), 10);
                            }
                        }

                    }

                    //リストに格納
                    m_Excel[sheetIndex - 1].Row.Add(r);
                    m_Excel[sheetIndex - 1].Item.Add(item);
                    m_Excel[sheetIndex - 1].Addr.Add(addr);
                    m_Excel[sheetIndex - 1].Data.Add(data);

                    arrOut.Add(row);
                }

                // ワークシートを閉じる
                Marshal.ReleaseComObject(xlWorksheet);
                xlWorksheet = null;
            }

            // ワークブックとエクセルのプロセスを閉じる
            ExcelClose();

            return arrOut;
        }
#endif

        /// <summary>
        /// エクセルファイルの指定したシートのセルにデータを書き込む
        /// </summary>
        /// <param name="filePath">エクセルファイルのパス</param>
        /// <param name="sheetIndex">シートの番号 (1, 2, 3, ...)</param>
        /// <param name="row">行 (>= 1)</param>
        /// <param name="colmn">列 (>= 1)</param>
        /// <param name="data">書込データ</param>
        /// <returns>実行結果</returns>
        public bool ExcelWrite(string filePath, int sheetIndex, int row, int colmn, ArrayList data, string date)
        {
            this.Cursor = Cursors.WaitCursor;

            lblStatus.Text = "Excelファイルへ保存中...";

            // テンプレートをコピー
            if (!File.Exists(filePath))
            {
                Directory.CreateDirectory(filePath);
            }

            // データ取得年月日
            string strTime = dtpDate.Value.Date.ToString().Substring(0,11) + cmbStartTimeStamp.SelectedItem.ToString(); ;
            DateTime dTime = DateTime.Parse(strTime);

            string SaveFilePath = filePath + dTime.ToString("日報yyyyMMdd_HHmmss") + ".xlsx";
            File.Copy(AppDomain.CurrentDomain.BaseDirectory + @"Template\" + m_TemplateName, SaveFilePath);

            // ワークブックを開く
            if (!ExcelOpen(SaveFilePath)) { return false; }

            if (xlWorkbook != null)
            {
                xlWorksheet = xlWorkbook.Sheets[sheetIndex];
                xlWorksheet.Select();

                xlWorksheet.Cells[1, 4] = date.Substring(0, 10) + "(1/1)";     // 日付

                // データ
                xlWorksheet.Cells[row, colmn] = data[data.Count - 2];
                xlWorksheet.Cells[row + 1, colmn] = data[data.Count - 1];

                for (int i = 0; i < data.Count - 2; i++)
                {
                    xlWorksheet.Cells[i + row + 2, colmn] = data[i];
                }

                // 保存する
                xlWorkbook.Save();

                // ワークシートを閉じる
                Marshal.ReleaseComObject(xlWorksheet);
                xlWorksheet = null;
            }

            // ワークブックとエクセルのプロセスを閉じる
            ExcelClose();

            lblStatus.Text = "";

            this.Cursor = Cursors.Default;

            return true;
        }

        /// <summary>
        /// 指定されたパスのExcel Workbookを開く
        /// </summary>
        /// <param name="filePath">Excel Workbookのパス(相対パスでも絶対パスでもOK)</param>
        /// <returns>Excel Workbookのオープンに成功したら true. それ以外 false.</returns>
        protected bool ExcelOpen(string filePath)
        {
            if (!System.IO.File.Exists(filePath))
            {
                return false;
            }

            try
            {
                xlApp = new Microsoft.Office.Interop.Excel.Application();
                xlApp.Visible = false;

                // filePath が相対パスのとき例外が発生するので fullPath に変換
                string fullPath = System.IO.Path.GetFullPath(filePath);
                xlWorkbook = xlApp.Workbooks.Open(fullPath);
            }
            catch
            {
                ExcelClose();
                return false;
            }

            return true;
        }

        /// <summary>
        /// 開いているWorkbookとExcelのプロセスを閉じる.
        /// </summary>
        protected void ExcelClose()
        {
            if (xlWorkbook != null)
            {
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);
                xlWorkbook = null;
            }

            if (xlApp != null)
            {
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
                xlApp = null;
            }
        }

        /// <summary>
        /// パラメータ変更（制御パラメータ）
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cmbItemName1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        /// <summary>
        /// 制御パラメータの入力値をNet10の指定アドレスに送信
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSet1_Click(object sender, EventArgs e)
        {
        }

        /// <summary>
        /// 日付変更時
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dtpDate_ValueChanged(object sender, EventArgs e)
        {
            TimeList();
        }

        /// <summary>
        /// 対象日付の時刻をリスト化
        /// </summary>
        private void TimeList()
        {
            SqlConnection con = new SqlConnection(CONNECT_STRING);
            con.Open();

            try
            {
                int i = 0;
                var strList = new List<string>();

                cmbStartTimeStamp.Items.Clear();
                cmbStartTimeStamp.Text = "";
                string st_date = dtpDate.Value.Year + "-" + dtpDate.Value.Month + "-" + dtpDate.Value.Day + " 00:00:00";
                string ed_date = dtpDate.Value.Year + "-" + dtpDate.Value.Month + "-" + dtpDate.Value.Day + " 23:59:59";
                string sqlstr = "SELECT 開始時刻 FROM JOB WHERE 開始時刻 BETWEEN '" + st_date + "' AND '" + ed_date + "'";
                SqlCommand com = new SqlCommand(sqlstr, con);

                using (SqlDataReader reader = com.ExecuteReader())
                {
                    while (reader.Read() == true)
                    {
                        if (reader[0] != null)
                        {
                            cmbStartTimeStamp.Items.Add(reader[0].ToString().Substring(11));
                        }
                        i++;
                    }
                }
                com.Dispose();
            }
            finally
            {
                con.Close();
            }

            if (cmbStartTimeStamp.Items.Count > 0)
            {
                cmbStartTimeStamp.SelectedIndex = 0;
            }
        }

        /// <summary>
        /// データ抽出
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnGetData_Click(object sender, EventArgs e)
        {
            if (dtpDate.Text.ToString() == "")
            {
                MessageBox.Show("日付が入力されていません", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (cmbStartTimeStamp.SelectedItem == null || cmbStartTimeStamp.SelectedItem.ToString() == "")
            {
                MessageBox.Show("開始時刻が選択されていません", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            SqlConnection con = new SqlConnection(CONNECT_STRING);
            con.Open();

            try
            {
                var strHeadList = new List<string>();
                var strList = new List<string>();

                string sqlstr = "SELECT t.name AS テーブル名 ,c.name AS 項目名 ,type_name(user_type_id) AS 属性 , max_length AS 長さ , CASE WHEN is_nullable = 1 THEN 'YES' ELSE 'NO' END AS NULL許可 FROM sys.objects t INNER JOIN sys.columns c ON t.object_id = c.object_id WHERE t.type = 'U' AND t.name = 'JOB' ORDER BY c.column_id";
                SqlCommand com = new SqlCommand(sqlstr, con);

                using (SqlDataReader reader = com.ExecuteReader())
                {
                    while (reader.Read() == true)
                    {
                        strHeadList.Add(reader[1].ToString());
                    }
                }

                string tgt_date = dtpDate.Value.Year + "-" + dtpDate.Value.Month + "-" + dtpDate.Value.Day;
                sqlstr = "SELECT * FROM JOB WHERE 開始時刻 ='" + tgt_date + " " + cmbStartTimeStamp.SelectedItem + "'";
                com = new SqlCommand(sqlstr, con);

                using (SqlDataReader reader = com.ExecuteReader())
                {
                    while (reader.Read() == true)
                    {
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            if (reader[i] != null)
                            {
                                strList.Add(reader[i].ToString());
                            }
                            else
                            {
                                strList.Add("");
                            }
                        }
                    }
                }

                dgvDailyData.Rows.Clear();
                for (int i = 1; i < strHeadList.Count; i++)
                {
                    dgvDailyData.Rows.Add(strHeadList[i], strList[i]);
                }
                com.Dispose();

            }
            finally
            {
                con.Close();
            }
        }
    }
}