using CsvHelper;
using OutlookOkan.Properties;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace OutlookOkan.CsvTools
{
    //TODO To be improved
    public class CsvImportAndExport
    {
        /// <summary>
        /// CSVファイルをインポートする。
        /// </summary>
        /// <returns>インポートするCSVファイルのパス</returns>
        public string ImportCsv()
        {
            MessageBox.Show(Resources.BerforeCSVImportAlert);

            var openFileDialog = new OpenFileDialog
            {
                Title = Resources.SelectCSVFile,
                InitialDirectory = @"C:\",
                Filter = "CSV|*.csv"
            };

            var importPath = openFileDialog.ShowDialog() == DialogResult.OK ? openFileDialog.FileName : null;
            openFileDialog.Dispose();

            return importPath;
        }

        /// <summary>
        /// CSVファイルを読み込みパースする。
        /// </summary>
        /// <typeparam name="TMaptype">CsvClassMap型</typeparam>
        /// <returns>パースされたCSV</returns>
        public CsvReader LoadCsv<TMaptype>(string filePath) where TMaptype : CsvHelper.Configuration.ClassMap
        {
            var csvReader = new CsvReader(new StreamReader(filePath, Encoding.UTF8));
            csvReader.Configuration.HasHeaderRecord = false;
            csvReader.Configuration.RegisterClassMap<TMaptype>();

            return csvReader;
        }

        /// <summary>
        /// 読み込んだCSVから、List<T/>を変えす。
        /// </summary>
        /// <typeparam name="TCsvType"></typeparam>
        /// <param name="loadedCsv">パースされたCSV</param>
        /// <returns>CSVデータ(List<T/>)</returns>
        public List<TCsvType> ReadCsv<TCsvType>(CsvReader loadedCsv)
        {
            var list = loadedCsv.GetRecords<TCsvType>().ToList();
            loadedCsv.Dispose();

            return list;
        }

        /// <summary>
        /// 設定ウィンドウ内で表示されている項目をCSVでエクスポートする。
        /// </summary>
        /// <typeparam name="TMaptype">CsvClassMap型</typeparam>
        /// <param name="bindableData">エクスポートするデータ</param>
        /// <param name="defaultFileName">デフォルトのファイル名(.csvと付けること)</param>
        public void CsvExport<TMaptype>(BindingSource bindableData, string defaultFileName) where TMaptype : CsvHelper.Configuration.ClassMap
        {
            var saveFileDialog = new SaveFileDialog
            {
                Title = Resources.SelectSaveDestination,
                InitialDirectory = @"C:\",
                Filter = "CSV|*.csv",
                CreatePrompt = true,
                FileName = defaultFileName
            };

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    var csvWriter = new CsvWriter(new StreamWriter(saveFileDialog.FileName, false, Encoding.UTF8));
                    csvWriter.Configuration.HasHeaderRecord = false;
                    csvWriter.Configuration.RegisterClassMap<TMaptype>();

                    csvWriter.WriteRecords(bindableData);

                    csvWriter.Dispose();

                    MessageBox.Show(Resources.SuccessfulExport);
                }
                catch (Exception e)
                {
                    MessageBox.Show(Resources.ExportFailed + e);
                }
            }
            saveFileDialog.Dispose();
        }
    }
}