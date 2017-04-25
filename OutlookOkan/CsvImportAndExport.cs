using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using CsvHelper;

namespace OutlookOkan
{
    //TODO ReadAndWriteCsv Class との重複処理を無くすように2つのクラスを整理する。
    public class CsvImportAndExport
    {
        /// <summary>
        /// CSVファイルをインポートする。
        /// </summary>
        /// <returns>インポートするCSVファイルのパス</returns>
        public string ImportCsv()
        {
            MessageBox.Show("書式の異なるCSVをインポートするとエラーになります。気を付けてください。");

            var openFileDialog = new OpenFileDialog
            {
                Title = "CSVファイルを選択してください。",
                InitialDirectory = @"C:\",
                Filter = "CSV|*.csv"
            };

            var importPath =  openFileDialog.ShowDialog() == DialogResult.OK ? openFileDialog.FileName : null;
            openFileDialog.Dispose();

            return importPath;
        }

        /// <summary>
        /// CSVファイルを読み込みパースする。
        /// </summary>
        /// <typeparam name="TMaptype">CsvClassMap型</typeparam>
        /// <returns>パースされたCSV</returns>
        public CsvParser ParseCsv<TMaptype>(string filePath) where TMaptype : CsvHelper.Configuration.CsvClassMap
        {
            var csvParser = new CsvParser(new StreamReader(filePath, Encoding.GetEncoding("Shift_JIS")));
            csvParser.Configuration.HasHeaderRecord = false;
            csvParser.Configuration.RegisterClassMap<TMaptype>();

            return csvParser;
        }

        /// <summary>
        /// パースされたCSVをもとに、List<T/>を変えす。
        /// </summary>
        /// <typeparam name="TCsvType"></typeparam>
        /// <param name="csvPerser">パースされたCSV</param>
        /// <returns>CSVデータ(List<T/>)</returns>
        public List<TCsvType> ReadCsv<TCsvType>(CsvParser csvPerser)
        {
            var list = new CsvReader(csvPerser).GetRecords<TCsvType>().ToList();
            csvPerser.Dispose();

            return list;
        }

        /// <summary>
        /// 設定ウィンドウ内で表示されている項目をCSVでエクスポートする。
        /// </summary>
        /// <typeparam name="TMaptype">CsvClassMap型</typeparam>
        /// <param name="bindableData">エクスポートするデータ</param>
        /// <param name="defaultFileName">デフォルトのファイル名(.csvと付けること)</param>
        public void CsvExport<TMaptype>(BindingSource bindableData,string defaultFileName) where TMaptype : CsvHelper.Configuration.CsvClassMap
        {
            var saveFileDialog = new SaveFileDialog
            {
                Title = "保存先を選択してください。",
                InitialDirectory = @"C:\",
                Filter = "CSV|*.csv",
                CreatePrompt = true,
                FileName = defaultFileName
            };

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    var csvWriter = new CsvWriter(new StreamWriter(saveFileDialog.FileName, false, Encoding.GetEncoding("Shift_JIS")));
                    csvWriter.Configuration.HasHeaderRecord = false;
                    csvWriter.Configuration.RegisterClassMap<TMaptype>();

                    csvWriter.WriteRecords(bindableData);

                    csvWriter.Dispose();

                    MessageBox.Show("エクスポートが完了しました。");
                }
                catch (Exception e)
                {
                    MessageBox.Show("エクスポートに失敗しました" + e);
                }
            }
            saveFileDialog.Dispose();
        }
    }
}