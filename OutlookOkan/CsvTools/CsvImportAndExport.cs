using CsvHelper;
using Microsoft.Win32;
using OutlookOkan.Properties;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows;

namespace OutlookOkan.CsvTools
{
    public sealed class CsvImportAndExport : CsvToolsBase
    {
        private Encoding _fileEncoding;

        /// <summary>
        /// CSVファイルをインポートするパスを取得する。
        /// </summary>
        /// <returns>インポートするCSVファイルのパス</returns>
        public string ImportCsv()
        {
            MessageBox.Show(Resources.CSVImportAlert, Resources.AppName, MessageBoxButton.OK);

            var openFileDialog = new OpenFileDialog
            {
                Title = Resources.SelectCSVFile,
                InitialDirectory = @"C:\",
                Filter = "CSV|*.csv"
            };

            var importPath = openFileDialog.ShowDialog() ?? false ? openFileDialog.FileName : null;

            _fileEncoding = Encoding.GetEncoding(DetectCharset(importPath));

            return importPath;
        }

        /// <summary>
        /// CSVファイルを読み込みパースする。
        /// </summary>
        /// <typeparam name="TMaptype">CsvClassMap型</typeparam>
        /// <param name="filePath">CSVファイルのパス</param>
        /// <returns>パースされたCSV</returns>
        public CsvReader LoadCsv<TMaptype>(string filePath) where TMaptype : CsvHelper.Configuration.ClassMap
        {
            var csvReader = new CsvReader(new StreamReader(filePath, _fileEncoding));
            csvReader.Configuration.HasHeaderRecord = false;
            csvReader.Configuration.RegisterClassMap<TMaptype>();

            return csvReader;
        }

        /// <summary>
        /// 設定ウィンドウ内で表示されている項目をCSVでエクスポートする。
        /// </summary>
        /// <typeparam name="TMaptype">CsvClassMap型</typeparam>
        /// <param name="records">エクスポートするデータ</param>
        /// <param name="defaultFileName">デフォルトのファイル名(.csvと付けること)</param>
        public void CsvExport<TMaptype>(List<object> records, string defaultFileName) where TMaptype : CsvHelper.Configuration.ClassMap
        {
            var saveFileDialog = new SaveFileDialog
            {
                Title = Resources.SelectSaveDestination,
                InitialDirectory = @"C:\",
                Filter = "CSV|*.csv",
                CreatePrompt = true,
                FileName = defaultFileName
            };

            if (!saveFileDialog.ShowDialog() ?? false) return;
            try
            {
                var csvWriter = new CsvWriter(new StreamWriter(saveFileDialog.FileName, false, Encoding.UTF8));
                csvWriter.Configuration.HasHeaderRecord = false;
                csvWriter.Configuration.RegisterClassMap<TMaptype>();
                csvWriter.WriteRecords(records);
                csvWriter.Dispose();

                MessageBox.Show(Resources.SuccessfulExport, Resources.AppName, MessageBoxButton.OK);
            }
            catch (Exception e)
            {
                MessageBox.Show(Resources.ExportFailed + e, Resources.AppName, MessageBoxButton.OK);
            }
        }
    }
}