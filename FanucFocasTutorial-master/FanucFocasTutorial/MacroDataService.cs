using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;

namespace FanucFocasTutorial
{
    public class MacroDataService
    {
        private readonly string _connectionString;
        private readonly string _dbPath;

        public MacroDataService()
        {
            _dbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MacroData.db");
            _connectionString = $"Data Source={_dbPath};Version=3;";

            InitializeDatabase();
        }

        private void InitializeDatabase()
        {
            if (!File.Exists(_dbPath))
            {
                SQLiteConnection.CreateFile(_dbPath);
            }

            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();

                string createTableSql = @"
                    CREATE TABLE IF NOT EXISTS MacroData (
                        IpAddress TEXT NOT NULL,
                        MacroNumber INTEGER NOT NULL,
                        Value REAL NOT NULL,
                        LastUpdated TEXT NOT NULL,
                        PRIMARY KEY (IpAddress, MacroNumber)
                    )";

                using (var cmd = new SQLiteCommand(createTableSql, conn))
                {
                    cmd.ExecuteNonQuery();
                }
            }
        }

        // IP 주소에 해당하는 모든 매크로 데이터 로드
        public Dictionary<short, double> LoadMacroDataFromDb(string ipAddress)
        {
            var result = new Dictionary<short, double>();

            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();

                string sql = "SELECT MacroNumber, Value FROM MacroData WHERE IpAddress = @IpAddress";

                using (var cmd = new SQLiteCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@IpAddress", ipAddress);

                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            short macroNumber = Convert.ToInt16(reader["MacroNumber"]);
                            double value = Convert.ToDouble(reader["Value"]);
                            result[macroNumber] = value;
                        }
                    }
                }
            }

            return result;
        }

        // 단일 매크로 데이터 저장 또는 업데이트
        public void SaveOrUpdateMacroData(string ipAddress, short macroNumber, double value)
        {
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();

                string sql = @"
                    INSERT OR REPLACE INTO MacroData (IpAddress, MacroNumber, Value, LastUpdated)
                    VALUES (@IpAddress, @MacroNumber, @Value, @LastUpdated)";

                using (var cmd = new SQLiteCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@IpAddress", ipAddress);
                    cmd.Parameters.AddWithValue("@MacroNumber", macroNumber);
                    cmd.Parameters.AddWithValue("@Value", value);
                    cmd.Parameters.AddWithValue("@LastUpdated", DateTime.Now.ToString("o"));
                    cmd.ExecuteNonQuery();
                }
            }
        }

        // 여러 매크로 데이터 일괄 저장 또는 업데이트
        public void SaveOrUpdateMacroDataBatch(string ipAddress, Dictionary<short, double> macroData)
        {
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();

                using (var transaction = conn.BeginTransaction())
                {
                    try
                    {
                        string sql = @"
                            INSERT OR REPLACE INTO MacroData (IpAddress, MacroNumber, Value, LastUpdated)
                            VALUES (@IpAddress, @MacroNumber, @Value, @LastUpdated)";

                        using (var cmd = new SQLiteCommand(sql, conn))
                        {
                            foreach (var kvp in macroData)
                            {
                                cmd.Parameters.Clear();
                                cmd.Parameters.AddWithValue("@IpAddress", ipAddress);
                                cmd.Parameters.AddWithValue("@MacroNumber", kvp.Key);
                                cmd.Parameters.AddWithValue("@Value", kvp.Value);
                                cmd.Parameters.AddWithValue("@LastUpdated", DateTime.Now.ToString("o"));
                                cmd.ExecuteNonQuery();
                            }
                        }

                        transaction.Commit();
                    }
                    catch
                    {
                        transaction.Rollback();
                        throw;
                    }
                }
            }
        }

        // IP에 해당하는 매크로 데이터가 DB에 있는지 확인
        public bool HasMacroDataForIp(string ipAddress)
        {
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();

                string sql = "SELECT COUNT(*) FROM MacroData WHERE IpAddress = @IpAddress";

                using (var cmd = new SQLiteCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@IpAddress", ipAddress);
                    long count = (long)cmd.ExecuteScalar();
                    return count > 0;
                }
            }
        }

        // 특정 IP의 모든 매크로 데이터 삭제
        public void DeleteMacroDataForIp(string ipAddress)
        {
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();

                string sql = "DELETE FROM MacroData WHERE IpAddress = @IpAddress";

                using (var cmd = new SQLiteCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@IpAddress", ipAddress);
                    cmd.ExecuteNonQuery();
                }
            }
        }
    }
}
