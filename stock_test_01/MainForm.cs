using CsvHelper;
using CsvHelper.Configuration;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection.Emit;
using System.Runtime.Remoting.Messaging;
using System.Text.RegularExpressions;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace stock_test_01
{
    public class MainForm : Form
    {
        // メニューバー
        MenuStrip menuStrip;
        // 実行フォルダ
        static string currentDirectoryPath = AppDomain.CurrentDomain.BaseDirectory;
        // データフォルダ
        static string dataDirectoryPath = Path.Combine(currentDirectoryPath, "Data");
        // URL・ファイル・パス
        static string lastDownloadFile = "last_download.txt";
        static string lastDownloadFilePath = Path.Combine(dataDirectoryPath, lastDownloadFile);
        // ダウンロード元のURL
        static string PERChartUrl = "https://nikkeiyosoku.com/nikkeiper/";
        static string PERChartFilePath = Path.Combine(dataDirectoryPath, "PERChart.html");
        static string batchFile = "test.bat";
        static string batchFilePath = Path.Combine(dataDirectoryPath, batchFile);
        static string PERChartTextFile = "Stock.txt";
        static string PERChartTextFilePath = Path.Combine(dataDirectoryPath, PERChartTextFile);
        static string nikkeiIndexCSVFile = "nikkei_stock_average_daily_jp.csv";
        static string nikkeiIndexCSVFilePath = Path.Combine(dataDirectoryPath, nikkeiIndexCSVFile);
        string[] indexHeader;
        string[] indexValue;
        //変数
        static int nikkeiIndexCSVOldedstYear;

        public MainForm()
        {
            //メニューバー表示
            menuStrip = new MenuStrip();
            Controls.Add(menuStrip);
            DataDirectoryExists();
            DateTime lastDownloadDate = GetLastDownloadDate();
            Console.WriteLine(lastDownloadDate);
            if (IsNewDay(lastDownloadDate))
            {
                DownloadIndexFiles();
                SaveDownloadDate(DateTime.Now);
                Console.WriteLine("ファイルをダウンロード");
                ConvertPERChart();
            }
            else
            {
                if (File.Exists(PERChartTextFilePath)) 
                {
                    Console.WriteLine("今日はファイルをダウンロード済み");
                }
                else
                {
                    ConvertPERChart();
                }
            }
            NikkeiIndexCSVSeparate();
            ExtractIndexData();
            NikkeiIndexPerView();
        }

        private void DataDirectoryExists()
        {
            // Dataフォルダが存在するか確認し、無ければ作成
            if (!Directory.Exists(dataDirectoryPath))
            {
                Directory.CreateDirectory(dataDirectoryPath);
                Console.WriteLine("Dataフォルダ作成");
            }
            else
            {

            }

        }

        // 最後のダウンロード日時をファイルから取得
        static DateTime GetLastDownloadDate()
        {
            // ダウンロード日時を記録するファイル
            if (File.Exists(lastDownloadFilePath))
            {
                string savedDate = File.ReadAllText(lastDownloadFilePath);
                if (DateTime.TryParse(savedDate, out DateTime lastDownloadDate))
                {
                    return lastDownloadDate;
                }
            }
            // 初回実行時やファイルが存在しない場合、デフォルトの日付（例えば、1970年1月1日）を返す
            return DateTime.MinValue;
        }

        // ダウンロード日時をファイルに保存
        static void SaveDownloadDate(DateTime currentDate)
        {
            File.WriteAllText(lastDownloadFilePath, currentDate.ToString());
            Console.WriteLine($"ダウンロード日時を保存しました: {currentDate}");
        }

        // 昨日と比較して、まだ新しい日であるかを判定
        static bool IsNewDay(DateTime lastDownloadDate)
        {
            return lastDownloadDate.Date != DateTime.Now.Date;
        }

        private void LoadingIndex()
        {

        }

        // ダウンロード(日経平均・PERチャート)
        private void DownloadIndexFiles()
        {
            WebClient wc = new WebClient();
            wc.DownloadFile(PERChartUrl, PERChartFilePath);
            if (File.Exists(batchFilePath))
            {


            }
            else
            {
                string batContent = @"
                @echo off
                cd /d %~dp0
                setlocal enabledelayedexpansion

                rem ファイル名をフルパスで設定
                rem set FILENAMEFULLPATH=""%~dp0\nikkei_stock_average_daily_jp.csv""
                set FILENAME=nikkei_stock_average_daily_jp.csv
                set FILENAMEFULLPATH=""%~dp0\%FILENAME%""

                rem 日経平均csvをダウンロード
                bitsadmin /transfer ""nikkei"" https://indexes.nikkei.co.jp/nkave/historical/nikkei_stock_average_daily_jp.csv ""%FILENAMEFULLPATH%""

                rem 終了
                endlocal
                exit /b
                ";
                // バッチファイルを作成
                File.WriteAllText(batchFilePath, batContent);
                Console.WriteLine($"バッチファイルを作成しました: {batchFilePath}");
            }
            if (File.Exists(batchFilePath))
            {
                // バッチ実行
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = batchFilePath,
                    CreateNoWindow = true,
                    UseShellExecute = false
                };

                using (Process process = Process.Start(psi))
                {
                    // プロセスが終了するまで待機
                    process.WaitForExit();

                    // バッチファイルの終了コードを取得
                    int exitCode = process.ExitCode;

                    // 終了コードを表示
                    Console.WriteLine($"Batch file exited with code {exitCode}");
                }
            }
            else
            {
                Console.WriteLine("バッチファイルがありません");
            }

        }

        private void NikkeiIndexCSVSeparate()
        {
            var d = new Dictionary<int, List<string>>();
            //List<string> l = new List<string>();

            //エンコード
            System.Text.Encoding enc = System.Text.Encoding.GetEncoding(932);

            // ファイルを読み込む
            StreamReader sr = new StreamReader(nikkeiIndexCSVFilePath, enc);
            string[] lines = File.ReadAllLines(nikkeiIndexCSVFilePath);
            var validLines = lines.Take(lines.Length - 1).ToArray();

            // ヘッダー行を保存
            string header = sr.ReadLine();
            Console.WriteLine(header);

            //変数の初期化
            int year  = 0;

            // 各行を処理
            foreach (var line in validLines.Skip(1))
            {
                string[] columns = line.Split(',');

                // 配列内の各要素からダブルクォートを削除
                for (int i = 0; i < columns.Length; i++)
                {
                    columns[i] = columns[i].Replace("\"", ""); // ダブルクォートを削除
                }
                //l.Add(columns[0]);
                Console.WriteLine($"{columns[0]}");
                DateTime date;
                DateTime.TryParse(columns[0], out date);
                year = date.Year;
                if(d.ContainsKey(year))
                {
                    Console.WriteLine("0");

                }
                else
                {
                    d[year] = new List<string>();
                    d[year].Add(header);
                    Console.WriteLine(string.Join(", ", d[year]));
                }

                d[year].Add(line.Replace("\"", ""));

            }
            foreach(var entry in d)
            {
                string output = Path.Combine(dataDirectoryPath, $"nikkei_stock_average_daily_{entry.Key}.csv");
                File.WriteAllLines(output, entry.Value);
            }
            //nikkeiIndexCSVOldedstYear = DateTime.Parse(l[0]).Year;
            //Console.WriteLine(nikkeiIndexCSVOldedstYear);
        }

        static string ExtractTextInRange(string text, string start, string end)
        {
            // 範囲指定用の正規表現
            string pattern = Regex.Escape(start) + "(.*?)" + Regex.Escape(end);
            Match match = Regex.Match(text, pattern, RegexOptions.Singleline);

            if (match.Success)
            {
                return match.Groups[1].Value.Trim();
            }
            return string.Empty;
        }

        private void ConvertPERChart()
        {


            // HTMLファイルを読み込み
            string content = File.ReadAllText(PERChartFilePath);
            // 抽出範囲の開始文字列と終了文字列
            string startPhrase = "（直近2ヶ月分）";
            string endPhrase = "検索したいチャート名を入力して下さい。";
            //string startPhrase = "<table class=\"table table-bordered table-striped\">";
            //string endPhrase = "<div class=\"btm -selectbox\">";

            // 範囲内のテキストを抽出
            string extractedText = ExtractTextInRange(content, startPhrase, endPhrase);
            // 抽出結果を出力
            /*
            if (!string.IsNullOrEmpty(extractedText))
            {
                Console.WriteLine("抽出されたテキスト:");
                Console.WriteLine(extractedText);
            }
            else
            {
                Console.WriteLine("指定した範囲のテキストが見つかりませんでした。");
            }
            */
            // HTMLファイルに保存
            File.WriteAllText(PERChartFilePath, extractedText);
            File.WriteAllText(PERChartTextFilePath, extractedText);

        }

        private void ExtractIndexData()
        {
            // textファイルを読み込み
            string content = File.ReadAllText(PERChartTextFilePath);
            // 正規表現パターンを定義
            string header = @"<th>(.*?)</th>";
            string value = @"<td>(.*?)<\/td>";
            // マッチするテキストをリストに格納
            List<string> headers = new List<string>();
            foreach (Match match in Regex.Matches(content, header))
            {
                headers.Add(match.Groups[1].Value);
            }
            indexHeader = headers.ToArray();

            List<string> values = new List<string>();
            foreach (Match match in Regex.Matches(content, value))
            {
                // HTMLタグを除去
                string tagRemove = Regex.Replace(match.Groups[1].Value.Trim(), "<.*?>", string.Empty);
                values.Add(tagRemove);
            }
            indexValue = values.ToArray();

        }

        private void NikkeiIndexPerView()
        {
            int labelSizeX = 80;
            int labelSizeY = 25;
            int startPosX = 12;
            int startPosY = 30;
            int fontSize = 12;
            System.Windows.Forms.Label[] indexHeaderLabels = new System.Windows.Forms.Label[indexHeader.Length];
            System.Windows.Forms.Label[] indexValueLabels = new System.Windows.Forms.Label[indexValue.Length];

            //--------------------
            //日経平均ヘッダー
            for (int i = 0; i < indexHeader.Length; i++)
            {
                //Console.WriteLine($"{indexHeader[i]}");
                indexHeaderLabels[i] = new System.Windows.Forms.Label
                {
                    Text = indexHeader[i],
                    TextAlign = ContentAlignment.MiddleCenter,
                    Size = new Size(labelSizeX, labelSizeY),
                    Location = new Point(startPosX + (labelSizeX - 1) * (i + 1), startPosY),
                    BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle,
                    BackColor = Color.FromArgb(255, 242, 204),
                };
                indexHeaderLabels[i].Font = new Font(indexHeaderLabels[i].Font.FontFamily, fontSize);

            }
            Controls.AddRange(indexHeaderLabels);
            
            //日経平均
            for (int i = 0; i < indexHeader.Length; i++)
            {
                //:Console.WriteLine($"{indexValue[i]}");
                indexValueLabels[i] = new System.Windows.Forms.Label
                {
                    Text = indexValue[i],
                    TextAlign = ContentAlignment.MiddleRight,
                    Size = new Size(labelSizeX, labelSizeY),
                    Location = new Point(startPosX + (labelSizeX - 1) * (i + 1), startPosY + (labelSizeY - 1)),
                    BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle,
                    BackColor = Color.FromArgb(255, 255, 255),
                };
                indexValueLabels[i].Font = new Font(indexValueLabels[i].Font.FontFamily, fontSize);
                switch (indexHeader[i])
                {
                    case "前日差":
                        float diff = float.Parse(indexValue[2]);
                        var valueForeColor = Color.FromArgb(0, 0, 0);
                        if (diff > 0)
                        {
                            indexValueLabels[i].ForeColor = Color.FromArgb(255, 0, 0);
                        }
                        else if (diff < 0)
                        {
                            indexValueLabels[i].ForeColor = Color.FromArgb(81, 171, 79);
                        }
                        else
                        {
                            indexValueLabels[i].ForeColor = Color.FromArgb(0, 0, 0);
                        }
                        break;

                    case "PER":
                        float per = float.Parse(indexValue[3]);
                        var valueBackColor = Color.FromArgb(255, 255, 255);
                        if (per < 10)
                        {
                            indexValueLabels[i].ForeColor = Color.FromArgb(217, 217, 217);
                            indexValueLabels[i].BackColor = Color.FromArgb(39, 78, 19);
                        }
                        else if (per >= 10 && per < 11)
                        {
                            indexValueLabels[i].ForeColor = Color.FromArgb(204, 204, 204);
                            indexValueLabels[i].BackColor = Color.FromArgb(56, 118, 29);

                        }
                        else if (per >= 11 && per < 12)
                        {
                            indexValueLabels[i].BackColor = Color.FromArgb(106, 168, 79);
                        }
                        else if (per >= 12 && per < 13)
                        {
                            indexValueLabels[i].BackColor = Color.FromArgb(147, 196, 125);
                        }
                        else if (per >= 13 && per < 14)
                        {
                            indexValueLabels[i].BackColor = Color.FromArgb(182, 215, 168);
                        }
                        else if (per >= 14 && per < 15)
                        {
                            indexValueLabels[i].BackColor = Color.FromArgb(217, 234, 211);
                        }
                        else if (per >= 15 && per < 16)
                        {
                            indexValueLabels[i].BackColor = Color.FromArgb(244, 204, 204);
                        }
                        else if (per >= 16 && per < 17)
                        {
                            indexValueLabels[i].BackColor = Color.FromArgb(234, 153, 153);
                        }
                        else if (per >= 17 && per < 18)
                        {
                            indexValueLabels[i].BackColor = Color.FromArgb(244, 102, 102);
                        }
                        else if (per >= 18)
                        {
                            indexValueLabels[i].ForeColor = Color.FromArgb(204, 204, 204);
                            valueBackColor = Color.FromArgb(204, 0, 0);

                        }
                        else
                        {
                            indexValueLabels[i].BackColor = Color.FromArgb(255, 255, 255);
                        }
                        break;

                    default:

                        break;
                }
            }
            Controls.AddRange(indexValueLabels);

            //--------------------
            //PER倍率ヘッダー
            List<string> PERRangeHeaderLabelsList = new List<string>();
            for(int i = 0; i < 20; i++)
            {
                string s;
                if(i < 19)
                {
                    s = $"PER{9 + 0.5 * i}";
                }
                else
                {
                    s = "PER19";

                }
                PERRangeHeaderLabelsList.Add(s);
                //Console.WriteLine(PERRangeLabelsList[i]);
            }
            System.Windows.Forms.Label[] PERRangeHeaderLabels = new System.Windows.Forms.Label[PERRangeHeaderLabelsList.Count];
            for (int i = 0; i < PERRangeHeaderLabelsList.Count; i++)
            {
                PERRangeHeaderLabels[i] = new System.Windows.Forms.Label()
                {
                    Text = PERRangeHeaderLabelsList[i],
                    TextAlign = ContentAlignment.MiddleCenter,
                    Size = new Size(labelSizeX, labelSizeY),
                    Location = new Point(startPosX + (labelSizeX - 1) * (i + 1), startPosY + (labelSizeY - 1) * 2),
                    BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle,
                };
                PERRangeHeaderLabels[i].Font = new Font(PERRangeHeaderLabels[i].Font.FontFamily, fontSize);
                switch (i)
                {
                    case 0:// PER9
                    case 1:// PER9.5
                        PERRangeHeaderLabels[i].ForeColor = Color.FromArgb(217, 217, 217);
                        PERRangeHeaderLabels[i].BackColor = Color.FromArgb(39, 78, 19);
                        break;

                    case 2:// PER10
                    case 3:// PER10.5
                        PERRangeHeaderLabels[i].ForeColor = Color.FromArgb(217, 217, 217);
                        PERRangeHeaderLabels[i].BackColor = Color.FromArgb(56, 118, 29);
                        break;

                    case 4:// PER11
                    case 5:// PER11.5
                        PERRangeHeaderLabels[i].ForeColor = Color.FromArgb(0, 0, 0);
                        PERRangeHeaderLabels[i].BackColor = Color.FromArgb(106, 168, 79);
                        break;

                    case 6:// PER12
                    case 7:// PER12.5
                        PERRangeHeaderLabels[i].ForeColor = Color.FromArgb(0, 0, 0);
                        PERRangeHeaderLabels[i].BackColor = Color.FromArgb(147, 196, 125);
                        break;

                    case 8:// PER13
                    case 9:// PER13.5
                        PERRangeHeaderLabels[i].ForeColor = Color.FromArgb(0, 0, 0);
                        PERRangeHeaderLabels[i].BackColor = Color.FromArgb(182, 215, 168);
                        break;

                    case 10:// PER14
                    case 11:// PER14.5
                        PERRangeHeaderLabels[i].ForeColor = Color.FromArgb(0, 0, 0);
                        PERRangeHeaderLabels[i].BackColor = Color.FromArgb(217, 234, 211);
                        break;

                    case 12:// PER15
                    case 13:// PER15.5
                        PERRangeHeaderLabels[i].ForeColor = Color.FromArgb(0, 0, 0);
                        PERRangeHeaderLabels[i].BackColor = Color.FromArgb(244, 204, 204);
                        break;

                    case 14:// PER16
                    case 15:// PER16.5
                        PERRangeHeaderLabels[i].ForeColor = Color.FromArgb(0, 0, 0);
                        PERRangeHeaderLabels[i].BackColor = Color.FromArgb(234, 153, 153);
                        break;

                    case 16:// PER17
                    case 17:// PER17.5
                        PERRangeHeaderLabels[i].ForeColor = Color.FromArgb(0, 0, 0);
                        PERRangeHeaderLabels[i].BackColor = Color.FromArgb(244, 102, 102);
                        break;

                    case 18:// PER18
                    case 19:// PER18
                        PERRangeHeaderLabels[i].ForeColor = Color.FromArgb(217, 217, 217);
                        PERRangeHeaderLabels[i].BackColor = Color.FromArgb(204, 0, 0);
                        break;

                    default:
                        break;
                }    
            }
            Controls.AddRange(PERRangeHeaderLabels);

            //PER倍率表側
            string[] PERTableSideText =
            {
                "PER倍率",
                "日経平均",
                "PBR",
            };
            System.Windows.Forms.Label[] PERTableSideLabels = new System.Windows.Forms.Label[PERTableSideText.Length];
            for (int i = 0; i < PERTableSideText.Length; i++)
            {
                PERTableSideLabels[i] = new System.Windows.Forms.Label()
                {
                    Text = PERTableSideText[i],
                    TextAlign = ContentAlignment.MiddleCenter,
                    Size = new Size(labelSizeX, labelSizeY),
                    Location = new Point(startPosX , startPosY + (labelSizeY - 1) * (2 + i)),
                    BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle,
                    BackColor = Color.FromArgb(255, 242, 204),
                };
                PERTableSideLabels[i].Font = new Font(PERTableSideLabels[i].Font.FontFamily, fontSize);
            }
            Controls.AddRange(PERTableSideLabels);

            //PER倍率別日経平均
            List<double> PERRangeValueLabelsList = new List<double>();
            for (int i = 0; i < PERRangeHeaderLabelsList.Count; i++)
            {
                double index = float.Parse(indexValue[5]);
                double d;
                if (i < PERRangeHeaderLabelsList.Count - 1)
                {
                    d = 9 + 0.5 * (double)i;
                }
                else
                {
                    d = 19;
                }
                double result = index * d;
                PERRangeValueLabelsList.Add(result);
                //Console.WriteLine($"{PERRangeHeaderLabelsList[i]}:{Math.Round(result, 2)}");
            }
            System.Windows.Forms.Label[] PERRangeValueLabels = new System.Windows.Forms.Label[PERRangeHeaderLabelsList.Count];
            for (int i = 0; i < PERRangeHeaderLabelsList.Count; i++)
            {
                PERRangeValueLabels[i] = new System.Windows.Forms.Label()
                {
                    Text = Convert.ToString(Math.Round(PERRangeValueLabelsList[i], 2)),
                    TextAlign = ContentAlignment.MiddleRight,
                    Size = new Size(labelSizeX, labelSizeY),
                    Location = new Point(startPosX + (labelSizeX - 1) * (i + 1), startPosY + (labelSizeY - 1) * 3),
                    BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle,
                    BackColor = Color.FromArgb(255, 255, 255),
                };
                PERRangeValueLabels[i].Font = new Font(PERRangeValueLabels[i].Font.FontFamily, fontSize);
            }
            Controls.AddRange(PERRangeValueLabels);

            //PER倍率別日経平均PBR倍率
            List<double> PBRValueLabelsList = new List<double>();
            for (int i = 0; i < PERRangeHeaderLabelsList.Count; i++)
            {
                double result = PERRangeValueLabelsList[i] / float.Parse(indexValue[6]);
                PBRValueLabelsList.Add(result);
                //Console.WriteLine($"{Math.Round(result, 2)}");
            }
            System.Windows.Forms.Label[] PBRValueLabels = new System.Windows.Forms.Label[PERRangeHeaderLabelsList.Count];
            for (int i = 0; i < PERRangeHeaderLabelsList.Count; i++)
            {
                PBRValueLabels[i] = new System.Windows.Forms.Label()
                {
                    Text = Convert.ToString(Math.Round(PBRValueLabelsList[i], 2)),
                    TextAlign = ContentAlignment.MiddleRight,
                    Size = new Size(labelSizeX, labelSizeY),
                    Location = new Point(startPosX + (labelSizeX - 1) * (i + 1), startPosY + (labelSizeY - 1) * 4),
                    BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle,
                    BackColor = Color.FromArgb(255, 255, 255),
                };
                PBRValueLabels[i].Font = new Font(PBRValueLabels[i].Font.FontFamily, fontSize);
                double pbr = PBRValueLabelsList[i];
                if (pbr < 1)
                {
                    PBRValueLabels[i].ForeColor = Color.FromArgb(0, 0, 255);
                }
                else if(pbr >= 1.5)
                {
                    PBRValueLabels[i].ForeColor = Color.FromArgb(255, 0, 0);
                }
            }
            Controls.AddRange(PBRValueLabels);

            //--------------------
            //年別表頭
            DateTime current_year = DateTime.Now;
            int yearPassed = int.Parse(current_year.Year.ToString()) - 2021;
            Console.WriteLine(yearPassed);
            List<string> PERTableHeadText = new List<string>();
            for (int i =0; i < yearPassed; i++)
            {
                string year = (2022 + i).ToString();
                PERTableHeadText.Add(year);
            };
            System.Windows.Forms.Label[] PERTableHeadLabels = new System.Windows.Forms.Label[PERTableHeadText.Count];
            for (int i = 0; i < PERTableHeadText.Count; i++)
            {
                PERTableHeadLabels[i] = new System.Windows.Forms.Label()
                {
                    Text = PERTableHeadText[i],
                    TextAlign = ContentAlignment.MiddleCenter,
                    Size = new Size(labelSizeX, labelSizeY),
                    Location = new Point(startPosX + (labelSizeX - 1) * (i + 1), startPosY + (labelSizeY - 1) * 6),
                    BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle,
                    BackColor = Color.FromArgb(255, 242, 204),
                };
                PERTableHeadLabels[i].Font = new Font(PERTableHeadLabels[i].Font.FontFamily, fontSize);
            }
            Controls.AddRange(PERTableHeadLabels);

            //年別日経平均表側
            List<string> NikkeiRangeTableSideLabelsList = new List<string>();
            string[] NikkeiRangeTableSideText =
            {
                "終値最大",
                "終値最小",
                "終値中央値",
                "終値平均",
                "PER倍率最大",
                "PER倍率最小",
                "成長率",
            };
            System.Windows.Forms.Label[] NikkeiRangeTableSideLabels = new System.Windows.Forms.Label[NikkeiRangeTableSideText.Length];
            for (int i = 0; i < NikkeiRangeTableSideText.Length; i++)
            {
                NikkeiRangeTableSideLabels[i] = new System.Windows.Forms.Label()
                {
                    Text = NikkeiRangeTableSideText[i],
                    TextAlign = ContentAlignment.MiddleCenter,
                    Size = new Size(labelSizeX, labelSizeY),
                    Location = new Point(startPosX, startPosY + (labelSizeY - 1) * (7 + i)),
                    BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle,
                    BackColor = Color.FromArgb(255, 242, 204),
                };
                NikkeiRangeTableSideLabels[i].Font = new Font(NikkeiRangeTableSideLabels[i].Font.FontFamily, fontSize);
            }
            Controls.AddRange(NikkeiRangeTableSideLabels);

            //年別PER
            List<string> PERRangeTableSideLabelsList = new List<string>();
            for (int i = 0; i < 20; i++)
            {
                string s;
                if (i == 0)
                {
                    s = "PER19";
                }
                else
                {
                    s = $"PER{18 - 0.5 * (i - 1)}";
                }
                PERRangeTableSideLabelsList.Add(s);
                //Console.WriteLine(PERRangeLabelsList[i]);
            }
            System.Windows.Forms.Label[] PERRangeTableSideLabels = new System.Windows.Forms.Label[PERRangeTableSideLabelsList.Count];
            for (int i = 0; i < PERRangeTableSideLabelsList.Count; i++)
            {
                PERRangeTableSideLabels[i] = new System.Windows.Forms.Label()
                {
                    Text = PERRangeTableSideLabelsList[i],
                    TextAlign = ContentAlignment.MiddleCenter,
                    Size = new Size(labelSizeX, labelSizeY),
                    Location = new Point(startPosX + (labelSizeX - 1) * 0, startPosY + (labelSizeY - 1) * (7 + NikkeiRangeTableSideText.Length ) + (labelSizeY - 1) * i),
                    BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle,
                };
                PERRangeTableSideLabels[i].Font = new Font(PERRangeTableSideLabels[i].Font.FontFamily, fontSize);
                switch (i)
                {
                    case 0:// PER19
                    case 1:// PER18

                        PERRangeTableSideLabels[i].ForeColor = Color.FromArgb(217, 217, 217);
                        PERRangeTableSideLabels[i].BackColor = Color.FromArgb(204, 0, 0);
                        break;

                    case 2:// PER17.5
                    case 3:// PER17
                        PERRangeTableSideLabels[i].ForeColor = Color.FromArgb(0, 0, 0);
                        PERRangeTableSideLabels[i].BackColor = Color.FromArgb(244, 102, 102);
                        break;

                    case 4:// PER16.5
                    case 5:// PER16
                        PERRangeTableSideLabels[i].ForeColor = Color.FromArgb(0, 0, 0);
                        PERRangeTableSideLabels[i].BackColor = Color.FromArgb(234, 153, 153);

                        break;

                    case 6:// PER15.5
                    case 7:// PER15
                        PERRangeTableSideLabels[i].ForeColor = Color.FromArgb(0, 0, 0);
                        PERRangeTableSideLabels[i].BackColor = Color.FromArgb(244, 204, 204);
                        break;

                    case 8:// PER14.5
                    case 9:// PER14
                        PERRangeTableSideLabels[i].ForeColor = Color.FromArgb(0, 0, 0);
                        PERRangeTableSideLabels[i].BackColor = Color.FromArgb(217, 234, 211);
                        break;

                    case 10:// PER13.5
                    case 11:// PER13
                        PERRangeTableSideLabels[i].ForeColor = Color.FromArgb(0, 0, 0);
                        PERRangeTableSideLabels[i].BackColor = Color.FromArgb(182, 215, 168);
                        break;

                    case 12:// PER12.5
                    case 13:// PER12
                        PERRangeTableSideLabels[i].ForeColor = Color.FromArgb(0, 0, 0);
                        PERRangeTableSideLabels[i].BackColor = Color.FromArgb(147, 196, 125);
                        break;

                    case 14:// PER11.5
                    case 15:// PER11
                        PERRangeTableSideLabels[i].ForeColor = Color.FromArgb(0, 0, 0);
                        PERRangeTableSideLabels[i].BackColor = Color.FromArgb(106, 168, 79);
                        break;

                    case 16:// PER10.5
                    case 17:// PER10
                        PERRangeTableSideLabels[i].ForeColor = Color.FromArgb(217, 217, 217);
                        PERRangeTableSideLabels[i].BackColor = Color.FromArgb(56, 118, 29);
                        break;

                    case 18:// PER9.5
                    case 19:// PER9
                        PERRangeTableSideLabels[i].ForeColor = Color.FromArgb(217, 217, 217);
                        PERRangeTableSideLabels[i].BackColor = Color.FromArgb(39, 78, 19);
                        break;

                    default:
                        break;
                }
            }

            List<string> dataFileList = new List<string>();
            string searchKeyword = "nikkei_stock_average_daily_";
            dataFileList = Directory.GetFiles(dataDirectoryPath).Where(f => f.Contains(searchKeyword)).Select(f => f.Replace(dataDirectoryPath, "").Replace(searchKeyword, "").Replace("\\", "").Replace(".csv", "")).ToList();
            Console.WriteLine(dataFileList[0]);

            Controls.AddRange(PERRangeTableSideLabels);
        }
    }
}
    

    
