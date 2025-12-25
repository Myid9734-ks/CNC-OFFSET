using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Linq;

namespace FanucFocasTutorial
{
    public class LogDataService
    {
        private readonly string _connectionString;
        private readonly string _dbPath;

        public LogDataService()
        {
            _dbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "OperationLogs.db");
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

                // 공구 옵셋 로그 테이블
                string createOffsetLogTable = @"
                    CREATE TABLE IF NOT EXISTS ManualOffsetLogs (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT,
                        Timestamp TEXT NOT NULL,
                        IpAddress TEXT NOT NULL,
                        ToolNumber INTEGER NOT NULL,
                        Axis TEXT NOT NULL,
                        UserInput REAL NOT NULL,
                        BeforeValue REAL NOT NULL,
                        AfterValue REAL NOT NULL,
                        SentValue REAL NOT NULL,
                        Success INTEGER NOT NULL,
                        ErrorMessage TEXT
                    )";

                // 매크로 변수 로그 테이블
                string createMacroLogTable = @"
                    CREATE TABLE IF NOT EXISTS MacroVariableLogs (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT,
                        Timestamp TEXT NOT NULL,
                        IpAddress TEXT NOT NULL,
                        MacroNumber INTEGER NOT NULL,
                        UserInput REAL NOT NULL,
                        BeforeValue REAL NOT NULL,
                        AfterValue REAL NOT NULL,
                        SentValue REAL NOT NULL,
                        Success INTEGER NOT NULL,
                        ErrorMessage TEXT
                    )";

                // PMC 제어 로그 테이블
                string createPmcLogTable = @"
                    CREATE TABLE IF NOT EXISTS PmcControlLogs (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT,
                        Timestamp TEXT NOT NULL,
                        IpAddress TEXT NOT NULL,
                        ControlType TEXT NOT NULL,
                        BeforeValue INTEGER NOT NULL,
                        AfterValue INTEGER NOT NULL,
                        Success INTEGER NOT NULL,
                        ErrorMessage TEXT
                    )";

                using (var cmd = new SQLiteCommand(createOffsetLogTable, conn))
                {
                    cmd.ExecuteNonQuery();
                }

                using (var cmd = new SQLiteCommand(createMacroLogTable, conn))
                {
                    cmd.ExecuteNonQuery();
                }

                using (var cmd = new SQLiteCommand(createPmcLogTable, conn))
                {
                    cmd.ExecuteNonQuery();
                }

                // 설비 상태 로그 테이블
                string createStateLogTable = @"
                    CREATE TABLE IF NOT EXISTS MachineStateLogs (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT,
                        Timestamp TEXT NOT NULL,
                        IpAddress TEXT NOT NULL,
                        State TEXT NOT NULL,
                        DurationSeconds INTEGER NOT NULL,
                        PmcF0_0 INTEGER NOT NULL,
                        PmcF0_7 INTEGER NOT NULL DEFAULT 0,
                        PmcF1_0 INTEGER NOT NULL DEFAULT 0,
                        PmcF3_5 INTEGER NOT NULL DEFAULT 0,
                        PmcF7_0 INTEGER NOT NULL DEFAULT 0,
                        PmcF10_0 INTEGER NOT NULL,
                        PmcG4_3 INTEGER NOT NULL,
                        PmcX8_4 INTEGER NOT NULL
                    )";

                using (var cmd = new SQLiteCommand(createStateLogTable, conn))
                {
                    cmd.ExecuteNonQuery();
                }

                // 기존 테이블에 새 PMC 컬럼 추가 (ALTER TABLE)
                try
                {
                    using (var cmd = new SQLiteCommand("ALTER TABLE MachineStateLogs ADD COLUMN PmcF0_7 INTEGER NOT NULL DEFAULT 0", conn))
                    {
                        cmd.ExecuteNonQuery();
                    }
                }
                catch { } // 컬럼이 이미 존재하면 무시

                try
                {
                    using (var cmd = new SQLiteCommand("ALTER TABLE MachineStateLogs ADD COLUMN PmcF1_0 INTEGER NOT NULL DEFAULT 0", conn))
                    {
                        cmd.ExecuteNonQuery();
                    }
                }
                catch { } // 컬럼이 이미 존재하면 무시

                try
                {
                    using (var cmd = new SQLiteCommand("ALTER TABLE MachineStateLogs ADD COLUMN PmcF3_5 INTEGER NOT NULL DEFAULT 0", conn))
                    {
                        cmd.ExecuteNonQuery();
                    }
                }
                catch { } // 컬럼이 이미 존재하면 무시

                // 근무조별 상태 로그 테이블
                string createShiftStateLogTable = @"
                    CREATE TABLE IF NOT EXISTS ShiftStateLogs (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT,
                        IpAddress TEXT NOT NULL,
                        ShiftDate TEXT NOT NULL,
                        ShiftType TEXT NOT NULL,
                        IsExtended INTEGER NOT NULL,
                        RunningSeconds INTEGER NOT NULL,
                        LoadingSeconds INTEGER NOT NULL,
                        AlarmSeconds INTEGER NOT NULL,
                        IdleSeconds INTEGER NOT NULL,
                        ProductionCount INTEGER NOT NULL,
                        CreatedAt TEXT NOT NULL
                    )";

                using (var cmd = new SQLiteCommand(createShiftStateLogTable, conn))
                {
                    cmd.ExecuteNonQuery();
                }

                // 인덱스 생성
                string createIndexes = @"
                    CREATE INDEX IF NOT EXISTS idx_state_logs_timestamp
                    ON MachineStateLogs(Timestamp);

                    CREATE INDEX IF NOT EXISTS idx_state_logs_ip_state
                    ON MachineStateLogs(IpAddress, State);

                    CREATE INDEX IF NOT EXISTS idx_shift_logs_ip_date
                    ON ShiftStateLogs(IpAddress, ShiftDate, ShiftType);
                ";

                using (var cmd = new SQLiteCommand(createIndexes, conn))
                {
                    cmd.ExecuteNonQuery();
                }
            }
        }

        // 공구 옵셋 로그 저장
        public void SaveOffsetLog(ManualOffsetLog log)
        {
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();

                string sql = @"
                    INSERT INTO ManualOffsetLogs
                    (Timestamp, IpAddress, ToolNumber, Axis, UserInput, BeforeValue, AfterValue, SentValue, Success, ErrorMessage)
                    VALUES (@Timestamp, @IpAddress, @ToolNumber, @Axis, @UserInput, @BeforeValue, @AfterValue, @SentValue, @Success, @ErrorMessage)";

                using (var cmd = new SQLiteCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@Timestamp", log.Timestamp.ToString("o"));
                    cmd.Parameters.AddWithValue("@IpAddress", log.IpAddress);
                    cmd.Parameters.AddWithValue("@ToolNumber", log.ToolNumber);
                    cmd.Parameters.AddWithValue("@Axis", log.Axis);
                    cmd.Parameters.AddWithValue("@UserInput", log.UserInput);
                    cmd.Parameters.AddWithValue("@BeforeValue", log.BeforeValue);
                    cmd.Parameters.AddWithValue("@AfterValue", log.AfterValue);
                    cmd.Parameters.AddWithValue("@SentValue", log.SentValue);
                    cmd.Parameters.AddWithValue("@Success", log.Success ? 1 : 0);
                    cmd.Parameters.AddWithValue("@ErrorMessage", log.ErrorMessage ?? "");
                    cmd.ExecuteNonQuery();
                }
            }
        }

        // 매크로 변수 로그 저장
        public void SaveMacroLog(MacroVariableLog log)
        {
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();

                string sql = @"
                    INSERT INTO MacroVariableLogs
                    (Timestamp, IpAddress, MacroNumber, UserInput, BeforeValue, AfterValue, SentValue, Success, ErrorMessage)
                    VALUES (@Timestamp, @IpAddress, @MacroNumber, @UserInput, @BeforeValue, @AfterValue, @SentValue, @Success, @ErrorMessage)";

                using (var cmd = new SQLiteCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@Timestamp", log.Timestamp.ToString("o"));
                    cmd.Parameters.AddWithValue("@IpAddress", log.IpAddress);
                    cmd.Parameters.AddWithValue("@MacroNumber", log.MacroNumber);
                    cmd.Parameters.AddWithValue("@UserInput", log.UserInput);
                    cmd.Parameters.AddWithValue("@BeforeValue", log.BeforeValue);
                    cmd.Parameters.AddWithValue("@AfterValue", log.AfterValue);
                    cmd.Parameters.AddWithValue("@SentValue", log.SentValue);
                    cmd.Parameters.AddWithValue("@Success", log.Success ? 1 : 0);
                    cmd.Parameters.AddWithValue("@ErrorMessage", log.ErrorMessage ?? "");
                    cmd.ExecuteNonQuery();
                }
            }
        }

        // 모든 공구 옵셋 로그 로드
        public List<ManualOffsetLog> LoadAllOffsetLogs()
        {
            var logs = new List<ManualOffsetLog>();

            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();

                string sql = "SELECT * FROM ManualOffsetLogs ORDER BY Timestamp DESC";

                using (var cmd = new SQLiteCommand(sql, conn))
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        logs.Add(new ManualOffsetLog
                        {
                            Timestamp = DateTime.Parse(reader["Timestamp"].ToString()),
                            IpAddress = reader["IpAddress"].ToString(),
                            ToolNumber = Convert.ToInt16(reader["ToolNumber"]),
                            Axis = reader["Axis"].ToString(),
                            UserInput = Convert.ToDouble(reader["UserInput"]),
                            BeforeValue = Convert.ToDouble(reader["BeforeValue"]),
                            AfterValue = Convert.ToDouble(reader["AfterValue"]),
                            SentValue = Convert.ToDouble(reader["SentValue"]),
                            Success = Convert.ToInt32(reader["Success"]) == 1,
                            ErrorMessage = reader["ErrorMessage"].ToString()
                        });
                    }
                }
            }

            return logs;
        }

        // 모든 매크로 변수 로그 로드
        public List<MacroVariableLog> LoadAllMacroLogs()
        {
            var logs = new List<MacroVariableLog>();

            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();

                string sql = "SELECT * FROM MacroVariableLogs ORDER BY Timestamp DESC";

                using (var cmd = new SQLiteCommand(sql, conn))
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        logs.Add(new MacroVariableLog
                        {
                            Timestamp = DateTime.Parse(reader["Timestamp"].ToString()),
                            IpAddress = reader["IpAddress"].ToString(),
                            MacroNumber = Convert.ToInt16(reader["MacroNumber"]),
                            UserInput = Convert.ToDouble(reader["UserInput"]),
                            BeforeValue = Convert.ToDouble(reader["BeforeValue"]),
                            AfterValue = Convert.ToDouble(reader["AfterValue"]),
                            SentValue = Convert.ToDouble(reader["SentValue"]),
                            Success = Convert.ToInt32(reader["Success"]) == 1,
                            ErrorMessage = reader["ErrorMessage"].ToString()
                        });
                    }
                }
            }

            return logs;
        }

        // 모든 로그 삭제
        public void ClearAllLogs()
        {
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();

                using (var transaction = conn.BeginTransaction())
                {
                    try
                    {
                        using (var cmd = new SQLiteCommand("DELETE FROM ManualOffsetLogs", conn))
                        {
                            cmd.ExecuteNonQuery();
                        }

                        using (var cmd = new SQLiteCommand("DELETE FROM MacroVariableLogs", conn))
                        {
                            cmd.ExecuteNonQuery();
                        }

                        using (var cmd = new SQLiteCommand("DELETE FROM PmcControlLogs", conn))
                        {
                            cmd.ExecuteNonQuery();
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

        // PMC 제어 로그 저장
        public void SavePmcLog(PmcControlLog log)
        {
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();

                string sql = @"
                    INSERT INTO PmcControlLogs
                    (Timestamp, IpAddress, ControlType, BeforeValue, AfterValue, Success, ErrorMessage)
                    VALUES (@Timestamp, @IpAddress, @ControlType, @BeforeValue, @AfterValue, @Success, @ErrorMessage)";

                using (var cmd = new SQLiteCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@Timestamp", log.Timestamp.ToString("o"));
                    cmd.Parameters.AddWithValue("@IpAddress", log.IpAddress);
                    cmd.Parameters.AddWithValue("@ControlType", log.ControlType);
                    cmd.Parameters.AddWithValue("@BeforeValue", log.BeforeValue ? 1 : 0);
                    cmd.Parameters.AddWithValue("@AfterValue", log.AfterValue ? 1 : 0);
                    cmd.Parameters.AddWithValue("@Success", log.Success ? 1 : 0);
                    cmd.Parameters.AddWithValue("@ErrorMessage", log.ErrorMessage ?? "");
                    cmd.ExecuteNonQuery();
                }
            }
        }

        // 모든 PMC 제어 로그 로드
        public List<PmcControlLog> LoadAllPmcLogs()
        {
            var logs = new List<PmcControlLog>();

            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();

                string sql = "SELECT * FROM PmcControlLogs ORDER BY Timestamp DESC";

                using (var cmd = new SQLiteCommand(sql, conn))
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        logs.Add(new PmcControlLog
                        {
                            Timestamp = DateTime.Parse(reader["Timestamp"].ToString()),
                            IpAddress = reader["IpAddress"].ToString(),
                            ControlType = reader["ControlType"].ToString(),
                            BeforeValue = Convert.ToInt32(reader["BeforeValue"]) == 1,
                            AfterValue = Convert.ToInt32(reader["AfterValue"]) == 1,
                            Success = Convert.ToInt32(reader["Success"]) == 1,
                            ErrorMessage = reader["ErrorMessage"].ToString()
                        });
                    }
                }
            }

            return logs;
        }

        // 로그 총 개수 조회
        public int GetTotalLogCount()
        {
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();

                string sql = @"
                    SELECT
                        (SELECT COUNT(*) FROM ManualOffsetLogs) +
                        (SELECT COUNT(*) FROM MacroVariableLogs) +
                        (SELECT COUNT(*) FROM PmcControlLogs) AS Total";

                using (var cmd = new SQLiteCommand(sql, conn))
                {
                    return Convert.ToInt32(cmd.ExecuteScalar());
                }
            }
        }

        // ==================== 설비 상태 모니터링 관련 ====================

        // 설비 상태 로그 저장
        public void SaveMachineStateLog(FanucFocasTutorial.MachineStateLog log)
        {
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();

                string sql = @"
                    INSERT INTO MachineStateLogs
                    (Timestamp, IpAddress, State, DurationSeconds,
                     PmcF0_0, PmcF0_7, PmcF1_0, PmcF3_5, PmcF10_0, PmcG4_3, PmcX8_4)
                    VALUES (@Timestamp, @IpAddress, @State, @DurationSeconds,
                            @PmcF0_0, @PmcF0_7, @PmcF1_0, @PmcF3_5, @PmcF10_0, @PmcG4_3, @PmcX8_4)";

                using (var cmd = new SQLiteCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@Timestamp", log.Timestamp.ToString("o"));
                    cmd.Parameters.AddWithValue("@IpAddress", log.IpAddress);
                    cmd.Parameters.AddWithValue("@State", log.State.ToString());
                    cmd.Parameters.AddWithValue("@DurationSeconds", log.DurationSeconds);

                    // PMC 값 파싱
                    var pmcValues = ParseBoolDictionary(log.PmcValues);
                    cmd.Parameters.AddWithValue("@PmcF0_0", pmcValues.ContainsKey("F0.0") && pmcValues["F0.0"] ? 1 : 0);
                    cmd.Parameters.AddWithValue("@PmcF0_7", pmcValues.ContainsKey("F0.7") && pmcValues["F0.7"] ? 1 : 0);
                    cmd.Parameters.AddWithValue("@PmcF1_0", pmcValues.ContainsKey("F1.0") && pmcValues["F1.0"] ? 1 : 0);
                    cmd.Parameters.AddWithValue("@PmcF3_5", pmcValues.ContainsKey("F3.5") && pmcValues["F3.5"] ? 1 : 0);
                    cmd.Parameters.AddWithValue("@PmcF10_0", pmcValues.ContainsKey("F10.0") && pmcValues["F10.0"] ? 1 : 0);
                    cmd.Parameters.AddWithValue("@PmcG4_3", pmcValues.ContainsKey("G4.3") && pmcValues["G4.3"] ? 1 : 0);
                    cmd.Parameters.AddWithValue("@PmcX8_4", pmcValues.ContainsKey("X8.4") && pmcValues["X8.4"] ? 1 : 0);

                    cmd.ExecuteNonQuery();
                }
            }
        }

        // OEE 분석용 일일 상태별 누적 시간 조회
        public Dictionary<string, int> GetDailyStateDurations(string ipAddress, DateTime date)
        {
            var result = new Dictionary<string, int>();

            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();

                string sql = @"
                    SELECT State, SUM(DurationSeconds) as TotalSeconds
                    FROM MachineStateLogs
                    WHERE IpAddress = @IpAddress
                      AND DATE(Timestamp) = DATE(@Date)
                    GROUP BY State";

                using (var cmd = new SQLiteCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@IpAddress", ipAddress);
                    cmd.Parameters.AddWithValue("@Date", date.ToString("yyyy-MM-dd"));

                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            result[reader["State"].ToString()] = Convert.ToInt32(reader["TotalSeconds"]);
                        }
                    }
                }
            }

            return result;
        }

        // 모든 설비 상태 로그 로드
        public List<FanucFocasTutorial.MachineStateLog> LoadAllMachineStateLogs()
        {
            var logs = new List<FanucFocasTutorial.MachineStateLog>();

            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();

                string sql = "SELECT * FROM MachineStateLogs ORDER BY Timestamp DESC LIMIT 1000";

                using (var cmd = new SQLiteCommand(sql, conn))
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        logs.Add(new FanucFocasTutorial.MachineStateLog
                        {
                            Timestamp = DateTime.Parse(reader["Timestamp"].ToString()),
                            IpAddress = reader["IpAddress"].ToString(),
                            State = (FanucFocasTutorial.MachineState)Enum.Parse(typeof(FanucFocasTutorial.MachineState), reader["State"].ToString()),
                            DurationSeconds = Convert.ToInt32(reader["DurationSeconds"]),
                            PmcValues = $"{{\"F0.0\":{reader["PmcF0_0"]},\"F7.0\":{reader["PmcF7_0"]},\"F10.0\":{reader["PmcF10_0"]},\"G4.3\":{reader["PmcG4_3"]},\"X8.4\":{reader["PmcX8_4"]}}}"
                        });
                    }
                }
            }

            return logs;
        }

        // ==================== 근무조 상태 관리 ====================

        /// <summary>
        /// 근무조 상태 데이터 저장
        /// </summary>
        public void SaveShiftStateData(ShiftStateData data)
        {
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();

                string sql = @"
                    INSERT INTO ShiftStateLogs
                    (IpAddress, ShiftDate, ShiftType, IsExtended,
                     RunningSeconds, LoadingSeconds, AlarmSeconds, IdleSeconds,
                     ProductionCount, CreatedAt)
                    VALUES (@IpAddress, @ShiftDate, @ShiftType, @IsExtended,
                            @RunningSeconds, @LoadingSeconds, @AlarmSeconds, @IdleSeconds,
                            @ProductionCount, @CreatedAt)";

                using (var cmd = new SQLiteCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@IpAddress", data.IpAddress);
                    cmd.Parameters.AddWithValue("@ShiftDate", data.ShiftDate.ToString("yyyy-MM-dd"));
                    cmd.Parameters.AddWithValue("@ShiftType", data.ShiftType.ToString());
                    cmd.Parameters.AddWithValue("@IsExtended", data.IsExtended ? 1 : 0);
                    cmd.Parameters.AddWithValue("@RunningSeconds", data.RunningSeconds);
                    cmd.Parameters.AddWithValue("@LoadingSeconds", data.LoadingSeconds);
                    cmd.Parameters.AddWithValue("@AlarmSeconds", data.AlarmSeconds);
                    cmd.Parameters.AddWithValue("@IdleSeconds", data.IdleSeconds);
                    cmd.Parameters.AddWithValue("@ProductionCount", data.ProductionCount);
                    cmd.Parameters.AddWithValue("@CreatedAt", DateTime.Now.ToString("o"));

                    cmd.ExecuteNonQuery();
                }
            }
        }

        /// <summary>
        /// 특정 근무조 상태 데이터 로드
        /// </summary>
        public ShiftStateData LoadShiftStateData(string ipAddress, DateTime shiftDate, ShiftType shiftType)
        {
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();

                string sql = @"
                    SELECT * FROM ShiftStateLogs
                    WHERE IpAddress = @IpAddress
                      AND ShiftDate = @ShiftDate
                      AND ShiftType = @ShiftType
                    ORDER BY CreatedAt DESC
                    LIMIT 1";

                using (var cmd = new SQLiteCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@IpAddress", ipAddress);
                    cmd.Parameters.AddWithValue("@ShiftDate", shiftDate.ToString("yyyy-MM-dd"));
                    cmd.Parameters.AddWithValue("@ShiftType", shiftType.ToString());

                    using (var reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            return new ShiftStateData
                            {
                                IpAddress = reader["IpAddress"].ToString(),
                                ShiftDate = DateTime.Parse(reader["ShiftDate"].ToString()),
                                ShiftType = (ShiftType)Enum.Parse(typeof(ShiftType), reader["ShiftType"].ToString()),
                                IsExtended = Convert.ToInt32(reader["IsExtended"]) == 1,
                                RunningSeconds = Convert.ToInt32(reader["RunningSeconds"]),
                                LoadingSeconds = Convert.ToInt32(reader["LoadingSeconds"]),
                                AlarmSeconds = Convert.ToInt32(reader["AlarmSeconds"]),
                                IdleSeconds = Convert.ToInt32(reader["IdleSeconds"]),
                                ProductionCount = Convert.ToInt32(reader["ProductionCount"])
                            };
                        }
                    }
                }
            }

            return null; // 데이터 없음
        }

        /// <summary>
        /// 근무조 이력 조회 (최근 30일)
        /// </summary>
        public List<ShiftStateData> GetShiftHistory(string ipAddress, int days = 30)
        {
            var result = new List<ShiftStateData>();

            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();

                string sql = @"
                    SELECT * FROM ShiftStateLogs
                    WHERE IpAddress = @IpAddress
                      AND ShiftDate >= DATE('now', @DaysAgo)
                    ORDER BY ShiftDate DESC, ShiftType";

                using (var cmd = new SQLiteCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@IpAddress", ipAddress);
                    cmd.Parameters.AddWithValue("@DaysAgo", $"-{days} days");

                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            result.Add(new ShiftStateData
                            {
                                IpAddress = reader["IpAddress"].ToString(),
                                ShiftDate = DateTime.Parse(reader["ShiftDate"].ToString()),
                                ShiftType = (ShiftType)Enum.Parse(typeof(ShiftType), reader["ShiftType"].ToString()),
                                IsExtended = Convert.ToInt32(reader["IsExtended"]) == 1,
                                RunningSeconds = Convert.ToInt32(reader["RunningSeconds"]),
                                LoadingSeconds = Convert.ToInt32(reader["LoadingSeconds"]),
                                AlarmSeconds = Convert.ToInt32(reader["AlarmSeconds"]),
                                IdleSeconds = Convert.ToInt32(reader["IdleSeconds"]),
                                ProductionCount = Convert.ToInt32(reader["ProductionCount"])
                            });
                        }
                    }
                }
            }

            return result;
        }

        // 간단한 JSON 파서 (Dictionary<string, bool> 전용)
        private Dictionary<string, bool> ParseBoolDictionary(string json)
        {
            var result = new Dictionary<string, bool>();
            if (string.IsNullOrEmpty(json)) return result;

            // Remove { } and split by ,
            json = json.Trim().Trim('{', '}');
            var pairs = json.Split(',');

            foreach (var pair in pairs)
            {
                // Split by : to get key and value
                var parts = pair.Split(':');
                if (parts.Length == 2)
                {
                    var key = parts[0].Trim().Trim('"');
                    var value = parts[1].Trim().ToLower() == "true";
                    result[key] = value;
                }
            }

            return result;
        }

        /// <summary>
        /// 근무조 데이터 조회 (날짜 범위, IP, 근무조 타입 필터)
        /// </summary>
        public List<ShiftStateData> GetShiftDataByFilter(DateTime startDate, DateTime endDate, string ipAddress = null, ShiftType? shiftType = null)
        {
            var result = new List<ShiftStateData>();

            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();

                string sql = @"
                    SELECT * FROM ShiftStateLogs
                    WHERE ShiftDate >= @StartDate AND ShiftDate <= @EndDate";

                if (!string.IsNullOrEmpty(ipAddress))
                {
                    sql += " AND IpAddress = @IpAddress";
                }

                if (shiftType.HasValue)
                {
                    sql += " AND ShiftType = @ShiftType";
                }

                sql += " ORDER BY ShiftDate DESC, IpAddress, ShiftType";

                using (var cmd = new SQLiteCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@StartDate", startDate.ToString("yyyy-MM-dd"));
                    cmd.Parameters.AddWithValue("@EndDate", endDate.ToString("yyyy-MM-dd"));

                    if (!string.IsNullOrEmpty(ipAddress))
                    {
                        cmd.Parameters.AddWithValue("@IpAddress", ipAddress);
                    }

                    if (shiftType.HasValue)
                    {
                        cmd.Parameters.AddWithValue("@ShiftType", shiftType.Value.ToString());
                    }

                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            result.Add(new ShiftStateData
                            {
                                IpAddress = reader["IpAddress"].ToString(),
                                ShiftDate = DateTime.Parse(reader["ShiftDate"].ToString()),
                                ShiftType = (ShiftType)Enum.Parse(typeof(ShiftType), reader["ShiftType"].ToString()),
                                IsExtended = Convert.ToInt32(reader["IsExtended"]) == 1,
                                RunningSeconds = Convert.ToInt32(reader["RunningSeconds"]),
                                LoadingSeconds = Convert.ToInt32(reader["LoadingSeconds"]),
                                AlarmSeconds = Convert.ToInt32(reader["AlarmSeconds"]),
                                IdleSeconds = Convert.ToInt32(reader["IdleSeconds"]),
                                ProductionCount = Convert.ToInt32(reader["ProductionCount"])
                            });
                        }
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// 등록된 모든 IP 주소 목록 조회
        /// </summary>
        public List<string> GetAllIpAddresses()
        {
            var result = new List<string>();

            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();

                string sql = "SELECT DISTINCT IpAddress FROM ShiftStateLogs ORDER BY IpAddress";

                using (var cmd = new SQLiteCommand(sql, conn))
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        result.Add(reader["IpAddress"].ToString());
                    }
                }
            }

            return result;
        }
    }
}
