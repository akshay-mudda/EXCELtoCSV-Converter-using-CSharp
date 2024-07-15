using System;
using System.Collections.Generic;
using System.Data.SqlClient;

namespace EXCELtoCSV_Converter_using_CSharp
{
    public static class DAL
    {
        public static void InsertMoveHistory(List<FileMoveRecord> moveRecords, string connectionString)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                foreach (var record in moveRecords)
                {
                    string query = "INSERT INTO FileMoveHistory (FileName, FileType, SourcePath, DestinationPath, CreateDate) VALUES (@FileName, @FileType, @SourcePath, @DestinationPath, @CreateDate)";
                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@FileName", record.FileName);
                        command.Parameters.AddWithValue("@FileType", record.FileType);
                        command.Parameters.AddWithValue("@SourcePath", record.SourcePath);
                        command.Parameters.AddWithValue("@DestinationPath", record.DestinationPath);
                        command.Parameters.AddWithValue("@CreateDate", DateTime.Now);

                        command.ExecuteNonQuery();
                    }
                }
            }
        }
    }
}