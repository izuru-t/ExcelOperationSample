using ClosedXML.Excel;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;

namespace ExcelOperationSample
{
    /// <summary>
    /// ClosedXMLを使用したExcelファイル操作サンプル
    /// （DBからデータを取得してExcelに書き出す）
    /// 
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            // テストデータ削除
            deleteTestData();
            // テストデータ登録
            createTestData();
            // エクセル出力
            exportDbExcel(@"c:\temp","User");
        }
        /// <summary>
        /// テーブルデータエクセル出力
        /// </summary>
        /// <param name="batPath">出力パス</param>
        /// <param name="tableName">テーブル名</param>
        private static void exportDbExcel(string batPath, string tableName)
        {
            var connectionString = ConfigurationManager.ConnectionStrings["LocalDB"].ConnectionString;
            using (var connection = new SqlConnection(connectionString)) 
            {
                connection.Open();

                using (var command = connection.CreateCommand())
                {
                    // sys.columnsテーブルからフィールド名情報を取得
                    command.CommandText = string.Format("select name from sys.columns where object_id = (select object_id from sys.objects where name = '[{0}]') order by column_id", tableName);
                    using (var reader = command.ExecuteReader())
                    {
                        // エクセルワークブックのインスタンス生成
                        using (var workbook = new XLWorkbook())
                        {
                            // テーブル名のシートを追加
                            var worksheet = workbook.Worksheets.Add(tableName);

                            int col = 1;
                            while (reader.Read())
                            {
                                // データ行の書き込み
                                worksheet.Cell(1, col).Value = "'" + reader["name"].ToString();

                                col++;
                            }
                            // エクセルワークブックを保存
                            workbook.SaveAs(Path.Combine(batPath, tableName + ".xlsx"));
                        }
                    }

                    // Userテーブルからデータを取得
                    command.CommandText = string.Format("SELECT * FROM [{0}]", tableName);
                    using (var reader = command.ExecuteReader())
                    {
                        // エクセルワークブックのインスタンス生成
                        using (var workbook = new XLWorkbook())
                        {
                            // テーブル名のシートを追加
                            var worksheet = workbook.Worksheets.Add(tableName);

                            int row = 1;
                            while (reader.Read())
                            {
                                // データ行の書き込み
                                worksheet.Cell(row, 1).Value = "'" + reader["id"].ToString();
                                worksheet.Cell(row, 2).Value = "'" + reader["Name"].ToString();
                                worksheet.Cell(row, 3).Value = "'" + reader["NameKana"].ToString();
                                worksheet.Cell(row, 4).Value = "'" + reader["BirthDay"].ToString();

                                row++;
                            }
                            // エクセルワークブックを保存
                            workbook.SaveAs(Path.Combine(batPath, tableName + ".xlsx"));
                        }
                    }
                }

                connection.Close();
            }
        }

        /// <summary>
        /// テストデータ削除
        /// </summary>
        private static void deleteTestData()
        {
            var connectionString = ConfigurationManager.ConnectionStrings["LocalDB"].ConnectionString;
            using (var connection = new SqlConnection(connectionString))
            {
                connection.Open();
                var command = connection.CreateCommand();
                for (int i = 0; i < 10000; i++)
                {
                    command.CommandText = "TRUNCATE TABLE [User]";
                    command.ExecuteNonQuery();
                }

            }
        }
        /// <summary>
        /// テストデータインサート
        /// </summary>
        private static void createTestData()

        {
            var connectionString = ConfigurationManager.ConnectionStrings["LocalDB"].ConnectionString;
            using (var connection = new SqlConnection(connectionString))
            {
                connection.Open();
                var command = connection.CreateCommand();
                for (int i = 0; i < 10; i++)
                {
                    command.CommandText = string.Format("INSERT INTO [User] (id,Name,NameKana,BirthDay)VALUES({0},{0},{0},'1900/1/1')", i);
                    command.ExecuteNonQuery();
                }

            }
        }
    }
}
