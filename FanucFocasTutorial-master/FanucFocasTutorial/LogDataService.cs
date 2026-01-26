using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Text;

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
                        UnmeasuredSeconds INTEGER NOT NULL DEFAULT 0,
                        ProductionCount INTEGER NOT NULL,
                        CreatedAt TEXT NOT NULL,
                        LastUpdatedAt TEXT
                    )";

                using (var cmd = new SQLiteCommand(createShiftStateLogTable, conn))
                {
                    cmd.ExecuteNonQuery();
                }

                // 기존 ShiftStateLogs 테이블에 UnmeasuredSeconds, LastUpdatedAt 컬럼 추가
                try
                {
                    using (var cmd = new SQLiteCommand("ALTER TABLE ShiftStateLogs ADD COLUMN UnmeasuredSeconds INTEGER NOT NULL DEFAULT 0", conn))
                    {
                        cmd.ExecuteNonQuery();
                    }
                }
                catch { } // 컬럼이 이미 존재하면 무시

                try
                {
                    using (var cmd = new SQLiteCommand("ALTER TABLE ShiftStateLogs ADD COLUMN LastUpdatedAt TEXT", conn))
                    {
                        cmd.ExecuteNonQuery();
                    }
                }
                catch { } // 컬럼이 이미 존재하면 무시

                // 제품 사이클 로그 테이블 (하루치만 저장)
                string createCycleLogTable = @"
                    CREATE TABLE IF NOT EXISTS ProductCycleLogs (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT,
                        IpAddress TEXT NOT NULL,
                        ShiftDate TEXT NOT NULL,
                        ShiftType TEXT NOT NULL,
                        State TEXT NOT NULL,
                        DurationSeconds INTEGER NOT NULL,
                        Timestamp TEXT NOT NULL,
                        CycleNumber INTEGER
                    )";

                using (var cmd = new SQLiteCommand(createCycleLogTable, conn))
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

                    CREATE INDEX IF NOT EXISTS idx_cycle_logs_ip_date
                    ON ProductCycleLogs(IpAddress, ShiftDate, ShiftType);
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
        public void SaveShiftStateData(ShiftStateData data, DateTime? updateTime = null)
        {
            DateTime saveTime = updateTime ?? DateTime.Now;
            
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();

                string sql = @"
                    INSERT INTO ShiftStateLogs
                    (IpAddress, ShiftDate, ShiftType, IsExtended,
                     RunningSeconds, LoadingSeconds, AlarmSeconds, IdleSeconds, UnmeasuredSeconds,
                     ProductionCount, CreatedAt, LastUpdatedAt)
                    VALUES (@IpAddress, @ShiftDate, @ShiftType, @IsExtended,
                            @RunningSeconds, @LoadingSeconds, @AlarmSeconds, @IdleSeconds, @UnmeasuredSeconds,
                            @ProductionCount, @CreatedAt, @LastUpdatedAt)";

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
                    cmd.Parameters.AddWithValue("@UnmeasuredSeconds", data.UnmeasuredSeconds);
                    cmd.Parameters.AddWithValue("@ProductionCount", data.ProductionCount);
                    cmd.Parameters.AddWithValue("@CreatedAt", saveTime.ToString("o"));
                    cmd.Parameters.AddWithValue("@LastUpdatedAt", saveTime.ToString("o"));

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
                                UnmeasuredSeconds = reader["UnmeasuredSeconds"] != DBNull.Value ? Convert.ToInt32(reader["UnmeasuredSeconds"]) : 0,
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

                // C#에서 날짜 계산 후 SQL에 전달
                DateTime startDate = DateTime.Today.AddDays(-days);
                string startDateStr = startDate.ToString("yyyy-MM-dd");

                string sql = @"
                    SELECT * FROM ShiftStateLogs
                    WHERE IpAddress = @IpAddress
                      AND ShiftDate >= @StartDate
                    ORDER BY ShiftDate DESC, ShiftType";

                using (var cmd = new SQLiteCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@IpAddress", ipAddress);
                    cmd.Parameters.AddWithValue("@StartDate", startDateStr);

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
                                UnmeasuredSeconds = reader["UnmeasuredSeconds"] != DBNull.Value ? Convert.ToInt32(reader["UnmeasuredSeconds"]) : 0,
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
            string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ShiftSearch_Debug.txt");

            try
            {
                // 검색 시작 로그
                string logMessage = $"\n{"=".PadRight(80, '=')}\n";
                logMessage += $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 근무조 데이터 검색 시작\n";
                logMessage += $"  검색 범위: {startDate:yyyy-MM-dd} ~ {endDate:yyyy-MM-dd}\n";
                logMessage += $"  IP 필터: {(string.IsNullOrEmpty(ipAddress) ? "전체" : ipAddress)}\n";
                logMessage += $"  근무조 필터: {(shiftType.HasValue ? shiftType.Value.ToString() : "전체")}\n";
                File.AppendAllText(logPath, logMessage, Encoding.UTF8);

                using (var conn = new SQLiteConnection(_connectionString))
                {
                    conn.Open();

                    // 날짜를 yyyy-MM-dd 형식으로 변환
                    string startDateStr = startDate.Date.ToString("yyyy-MM-dd");
                    string endDateStr = endDate.Date.ToString("yyyy-MM-dd");

                    // 야간 근무조는 자정을 넘어가므로 검색 범위 확장
                    // 야간 근무조: 20:30 ~ 다음날 08:30이므로, ShiftDate는 시작 날짜로 저장됨
                    // 예: 2025-01-26 20:30 시작 → ShiftDate = 2025-01-26
                    // 하지만 실제로는 2025-01-27 08:30까지이므로, 2025-01-27로 검색해도 나와야 함
                    // 따라서 야간 근무조 검색 시 시작일 하루 전부터 검색
                    string extendedStartDateStr = startDate.Date.AddDays(-1).ToString("yyyy-MM-dd");

                    File.AppendAllText(logPath, $"  날짜 변환: StartDate={startDateStr}, EndDate={endDateStr}, ExtendedStartDate={extendedStartDateStr}\n", Encoding.UTF8);

                    // SQL 쿼리: 야간 근무조는 확장된 범위에서 검색
                    string sql = @"
                        SELECT * FROM ShiftStateLogs
                        WHERE (
                            -- 주간 근무조: 일반 범위
                            (ShiftType = 'Day' AND ShiftDate >= @StartDate AND ShiftDate <= @EndDate)
                            OR
                            -- 야간 근무조: 확장된 범위 (시작일 하루 전부터)
                            (ShiftType = 'Night' AND ShiftDate >= @ExtendedStartDate AND ShiftDate <= @EndDate)
                        )";

                    if (!string.IsNullOrEmpty(ipAddress))
                    {
                        sql += " AND IpAddress = @IpAddress";
                    }

                    if (shiftType.HasValue)
                    {
                        sql += " AND ShiftType = @ShiftType";
                    }

                    sql += " ORDER BY ShiftDate DESC, IpAddress, ShiftType";

                    File.AppendAllText(logPath, $"  SQL 쿼리:\n{sql}\n", Encoding.UTF8);

                    int rawCount = 0;
                    int dayCount = 0;
                    int nightCount = 0;
                    int filteredOutCount = 0;

                    using (var cmd = new SQLiteCommand(sql, conn))
                    {
                        cmd.Parameters.AddWithValue("@StartDate", startDateStr);
                        cmd.Parameters.AddWithValue("@EndDate", endDateStr);
                        cmd.Parameters.AddWithValue("@ExtendedStartDate", extendedStartDateStr);

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
                                rawCount++;
                                try
                                {
                                    var shiftData = new ShiftStateData
                                    {
                                        IpAddress = reader["IpAddress"].ToString(),
                                        ShiftDate = DateTime.Parse(reader["ShiftDate"].ToString()),
                                        ShiftType = (ShiftType)Enum.Parse(typeof(ShiftType), reader["ShiftType"].ToString()),
                                        IsExtended = Convert.ToInt32(reader["IsExtended"]) == 1,
                                        RunningSeconds = Convert.ToInt32(reader["RunningSeconds"]),
                                        LoadingSeconds = Convert.ToInt32(reader["LoadingSeconds"]),
                                        AlarmSeconds = Convert.ToInt32(reader["AlarmSeconds"]),
                                        IdleSeconds = Convert.ToInt32(reader["IdleSeconds"]),
                                        UnmeasuredSeconds = reader["UnmeasuredSeconds"] != DBNull.Value ? Convert.ToInt32(reader["UnmeasuredSeconds"]) : 0,
                                        ProductionCount = Convert.ToInt32(reader["ProductionCount"])
                                    };

                                    // 야간 근무조 필터링: ShiftDate+1일이 검색 범위 내에 있는지 확인
                                    if (shiftData.ShiftType == ShiftType.Night)
                                    {
                                        DateTime shiftEndDate = shiftData.ShiftDate.AddDays(1);
                                        // 야간 근무조는 ShiftDate 날짜의 20:30부터 다음날 08:30까지
                                        // 따라서 ShiftDate 또는 ShiftDate+1일이 검색 범위 내에 있으면 포함
                                        if (shiftData.ShiftDate <= endDate.Date && shiftEndDate >= startDate.Date)
                                        {
                                            result.Add(shiftData);
                                            nightCount++;
                                            File.AppendAllText(logPath, $"  [야간] 포함: {shiftData.ShiftDate:yyyy-MM-dd} ({shiftData.IpAddress}) - ShiftDate={shiftData.ShiftDate:yyyy-MM-dd}, ShiftEndDate={shiftEndDate:yyyy-MM-dd}\n", Encoding.UTF8);
                                        }
                                        else
                                        {
                                            filteredOutCount++;
                                            File.AppendAllText(logPath, $"  [야간] 제외: {shiftData.ShiftDate:yyyy-MM-dd} ({shiftData.IpAddress}) - ShiftDate={shiftData.ShiftDate:yyyy-MM-dd}, ShiftEndDate={shiftEndDate:yyyy-MM-dd} (범위 밖)\n", Encoding.UTF8);
                                        }
                                    }
                                    else
                                    {
                                        // 주간 근무조는 일반적인 범위 체크
                                        if (shiftData.ShiftDate >= startDate.Date && shiftData.ShiftDate <= endDate.Date)
                                        {
                                            result.Add(shiftData);
                                            dayCount++;
                                            File.AppendAllText(logPath, $"  [주간] 포함: {shiftData.ShiftDate:yyyy-MM-dd} ({shiftData.IpAddress})\n", Encoding.UTF8);
                                        }
                                        else
                                        {
                                            filteredOutCount++;
                                            File.AppendAllText(logPath, $"  [주간] 제외: {shiftData.ShiftDate:yyyy-MM-dd} ({shiftData.IpAddress}) - 범위 밖\n", Encoding.UTF8);
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    // 개별 레코드 파싱 오류는 로그만 남기고 계속 진행
                                    string errorMsg = $"  [오류] 레코드 파싱 실패: {ex.Message}\n";
                                    File.AppendAllText(logPath, errorMsg, Encoding.UTF8);
                                    System.Diagnostics.Debug.WriteLine($"근무조 데이터 파싱 오류: {ex.Message}");
                                }
                            }
                        }
                    }

                    // 검색 결과 요약 로그
                    string summaryLog = $"\n  검색 결과 요약:\n";
                    summaryLog += $"    SQL 조회 결과: {rawCount}개\n";
                    summaryLog += $"    주간 근무조: {dayCount}개\n";
                    summaryLog += $"    야간 근무조: {nightCount}개\n";
                    summaryLog += $"    필터링 제외: {filteredOutCount}개\n";
                    summaryLog += $"    최종 결과: {result.Count}개\n";
                    summaryLog += $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 근무조 데이터 검색 완료\n";
                    summaryLog += $"{"=".PadRight(80, '=')}\n";
                    File.AppendAllText(logPath, summaryLog, Encoding.UTF8);
                }
            }
            catch (Exception ex)
            {
                // 전체 쿼리 오류는 예외를 다시 던짐
                string errorLog = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 근무조 데이터 조회 오류 발생\n";
                errorLog += $"  오류 메시지: {ex.Message}\n";
                errorLog += $"  스택 트레이스: {ex.StackTrace}\n";
                errorLog += $"{"=".PadRight(80, '=')}\n";
                File.AppendAllText(logPath, errorLog, Encoding.UTF8);
                System.Diagnostics.Debug.WriteLine($"근무조 데이터 조회 오류: {ex.Message}");
                throw;
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

        /// <summary>
        /// 근무조 데이터 총 개수 조회 (디버깅용)
        /// </summary>
        public int GetShiftDataCount()
        {
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();

                string sql = "SELECT COUNT(*) FROM ShiftStateLogs";

                using (var cmd = new SQLiteCommand(sql, conn))
                {
                    return Convert.ToInt32(cmd.ExecuteScalar());
                }
            }
        }

        /// <summary>
        /// 날짜 범위 내 근무조 데이터 개수 조회 (디버깅용)
        /// </summary>
        public int GetShiftDataCountByDateRange(DateTime startDate, DateTime endDate)
        {
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();

                string startDateStr = startDate.Date.ToString("yyyy-MM-dd");
                string endDateStr = endDate.Date.ToString("yyyy-MM-dd");

                string sql = @"
                    SELECT COUNT(*) FROM ShiftStateLogs
                    WHERE ShiftDate >= @StartDate AND ShiftDate <= @EndDate";

                using (var cmd = new SQLiteCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@StartDate", startDateStr);
                    cmd.Parameters.AddWithValue("@EndDate", endDateStr);
                    return Convert.ToInt32(cmd.ExecuteScalar());
                }
            }
        }

        /// <summary>
        /// 근무조 데이터 업데이트 (미측정 시간 포함)
        /// </summary>
        public void UpdateShiftStateData(ShiftStateData data, DateTime? updateTime = null)
        {
            DateTime saveTime = updateTime ?? DateTime.Now;
            
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();

                // 먼저 최신 레코드의 ID를 찾음
                string findIdSql = @"
                    SELECT Id FROM ShiftStateLogs
                    WHERE IpAddress = @IpAddress
                      AND ShiftDate = @ShiftDate
                      AND ShiftType = @ShiftType
                    ORDER BY CreatedAt DESC
                    LIMIT 1";

                int? recordId = null;
                using (var findCmd = new SQLiteCommand(findIdSql, conn))
                {
                    findCmd.Parameters.AddWithValue("@IpAddress", data.IpAddress);
                    findCmd.Parameters.AddWithValue("@ShiftDate", data.ShiftDate.ToString("yyyy-MM-dd"));
                    findCmd.Parameters.AddWithValue("@ShiftType", data.ShiftType.ToString());
                    
                    var result = findCmd.ExecuteScalar();
                    if (result != null && result != DBNull.Value)
                    {
                        recordId = Convert.ToInt32(result);
                    }
                }

                if (recordId.HasValue)
                {
                    // 레코드가 있으면 UPDATE
                    string updateSql = @"
                        UPDATE ShiftStateLogs
                        SET RunningSeconds = @RunningSeconds,
                            LoadingSeconds = @LoadingSeconds,
                            AlarmSeconds = @AlarmSeconds,
                            IdleSeconds = @IdleSeconds,
                            UnmeasuredSeconds = @UnmeasuredSeconds,
                            ProductionCount = @ProductionCount,
                            LastUpdatedAt = @LastUpdatedAt
                        WHERE Id = @Id";

                    using (var cmd = new SQLiteCommand(updateSql, conn))
                    {
                        cmd.Parameters.AddWithValue("@Id", recordId.Value);
                        cmd.Parameters.AddWithValue("@RunningSeconds", data.RunningSeconds);
                        cmd.Parameters.AddWithValue("@LoadingSeconds", data.LoadingSeconds);
                        cmd.Parameters.AddWithValue("@AlarmSeconds", data.AlarmSeconds);
                        cmd.Parameters.AddWithValue("@IdleSeconds", data.IdleSeconds);
                        cmd.Parameters.AddWithValue("@UnmeasuredSeconds", data.UnmeasuredSeconds);
                        cmd.Parameters.AddWithValue("@ProductionCount", data.ProductionCount);
                        cmd.Parameters.AddWithValue("@LastUpdatedAt", saveTime.ToString("o"));
                        cmd.ExecuteNonQuery();
                    }
                }
                else
                {
                    // 레코드가 없으면 INSERT
                    SaveShiftStateData(data, saveTime);
                }
            }
        }

        /// <summary>
        /// 특정 근무조의 마지막 업데이트 시간 조회
        /// </summary>
        public DateTime? GetLastUpdatedTime(string ipAddress, DateTime shiftDate, ShiftType shiftType)
        {
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();

                // LastUpdatedAt이 NULL이면 CreatedAt을 사용
                string sql = @"
                    SELECT COALESCE(LastUpdatedAt, CreatedAt) FROM ShiftStateLogs
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

                    var result = cmd.ExecuteScalar();
                    if (result != null && result != DBNull.Value)
                    {
                        return DateTime.Parse(result.ToString());
                    }
                }
            }

            return null;
        }

        // ==================== 제품 사이클 로그 관리 ====================

        /// <summary>
        /// 제품 사이클 상태 전환 로그 저장 (10초 단위)
        /// </summary>
        public void SaveCycleStateLog(string ipAddress, DateTime shiftDate, ShiftType shiftType, 
            MachineState state, int durationSeconds, DateTime timestamp, int? cycleNumber = null)
        {
            // 10초 단위로 반올림
            int roundedDuration = (int)Math.Round(durationSeconds / 10.0) * 10;
            
            // 10초 미만은 저장하지 않음
            if (roundedDuration < 10)
                return;

            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();

                string sql = @"
                    INSERT INTO ProductCycleLogs
                    (IpAddress, ShiftDate, ShiftType, State, DurationSeconds, Timestamp, CycleNumber)
                    VALUES (@IpAddress, @ShiftDate, @ShiftType, @State, @DurationSeconds, @Timestamp, @CycleNumber)";

                using (var cmd = new SQLiteCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@IpAddress", ipAddress);
                    cmd.Parameters.AddWithValue("@ShiftDate", shiftDate.ToString("yyyy-MM-dd"));
                    cmd.Parameters.AddWithValue("@ShiftType", shiftType.ToString());
                    cmd.Parameters.AddWithValue("@State", state.ToString());
                    cmd.Parameters.AddWithValue("@DurationSeconds", roundedDuration);
                    cmd.Parameters.AddWithValue("@Timestamp", timestamp.ToString("o"));
                    cmd.Parameters.AddWithValue("@CycleNumber", cycleNumber.HasValue ? (object)cycleNumber.Value : DBNull.Value);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        /// <summary>
        /// 특정 근무조의 사이클 로그 조회
        /// </summary>
        public List<CycleStateLog> GetCycleLogs(string ipAddress, DateTime shiftDate, ShiftType shiftType)
        {
            var result = new List<CycleStateLog>();

            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();

                string sql = @"
                    SELECT * FROM ProductCycleLogs
                    WHERE IpAddress = @IpAddress
                      AND ShiftDate = @ShiftDate
                      AND ShiftType = @ShiftType
                    ORDER BY Timestamp ASC";

                using (var cmd = new SQLiteCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@IpAddress", ipAddress);
                    cmd.Parameters.AddWithValue("@ShiftDate", shiftDate.ToString("yyyy-MM-dd"));
                    cmd.Parameters.AddWithValue("@ShiftType", shiftType.ToString());

                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            result.Add(new CycleStateLog
                            {
                                IpAddress = reader["IpAddress"].ToString(),
                                ShiftDate = DateTime.Parse(reader["ShiftDate"].ToString()),
                                ShiftType = (ShiftType)Enum.Parse(typeof(ShiftType), reader["ShiftType"].ToString()),
                                State = (MachineState)Enum.Parse(typeof(MachineState), reader["State"].ToString()),
                                DurationSeconds = Convert.ToInt32(reader["DurationSeconds"]),
                                Timestamp = DateTime.Parse(reader["Timestamp"].ToString()),
                                CycleNumber = reader["CycleNumber"] != DBNull.Value ? (int?)Convert.ToInt32(reader["CycleNumber"]) : null
                            });
                        }
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// 특정 근무조의 사이클 로그 삭제 (다음날 같은 근무조 시작 시)
        /// </summary>
        public void DeleteCycleLogs(string ipAddress, DateTime shiftDate, ShiftType shiftType)
        {
            using (var conn = new SQLiteConnection(_connectionString))
            {
                conn.Open();

                string sql = @"
                    DELETE FROM ProductCycleLogs
                    WHERE IpAddress = @IpAddress
                      AND ShiftDate = @ShiftDate
                      AND ShiftType = @ShiftType";

                using (var cmd = new SQLiteCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@IpAddress", ipAddress);
                    cmd.Parameters.AddWithValue("@ShiftDate", shiftDate.ToString("yyyy-MM-dd"));
                    cmd.Parameters.AddWithValue("@ShiftType", shiftType.ToString());
                    cmd.ExecuteNonQuery();
                }
            }
        }

        /// <summary>
        /// 사이클 리포트 생성 (평균, 특이사항 등)
        /// </summary>
        public CycleReport GenerateCycleReport(string ipAddress, DateTime shiftDate, ShiftType shiftType)
        {
            var logs = GetCycleLogs(ipAddress, shiftDate, shiftType);
            
            if (logs.Count == 0)
                return new CycleReport { HasData = false };

            // 사이클 추출 (Loading → Running 패턴)
            var cycles = new List<CycleData>();
            CycleData currentCycle = null;

            foreach (var log in logs)
            {
                if (log.State == MachineState.Loading)
                {
                    // 새 사이클 시작
                    if (currentCycle != null && currentCycle.RunningSeconds > 0)
                    {
                        cycles.Add(currentCycle);
                    }
                    currentCycle = new CycleData
                    {
                        StartTime = log.Timestamp,
                        LoadingSeconds = log.DurationSeconds
                    };
                }
                else if (log.State == MachineState.Running && currentCycle != null)
                {
                    currentCycle.RunningSeconds = log.DurationSeconds;
                    currentCycle.EndTime = log.Timestamp;
                }
            }

            // 마지막 사이클 추가
            if (currentCycle != null && currentCycle.RunningSeconds > 0)
            {
                cycles.Add(currentCycle);
            }

            if (cycles.Count == 0)
                return new CycleReport { HasData = false };

            // 통계 계산
            var cycleTimes = cycles.Select(c => c.LoadingSeconds + c.RunningSeconds).ToList();
            double avgCycleTime = cycleTimes.Average();
            int maxCycleTime = cycleTimes.Max();
            int minCycleTime = cycleTimes.Min();

            // 특이사항: 평균보다 20% 이상 긴 사이클
            double threshold = avgCycleTime * 1.2;
            var anomalies = cycles.Where(c => (c.LoadingSeconds + c.RunningSeconds) > threshold).ToList();

            return new CycleReport
            {
                HasData = true,
                TotalCycles = cycles.Count,
                AvgCycleTime = (int)Math.Round(avgCycleTime),
                MaxCycleTime = maxCycleTime,
                MinCycleTime = minCycleTime,
                AnomalyCount = anomalies.Count,
                Anomalies = anomalies
            };
        }
    }

    // ==================== 사이클 로그 관련 클래스 ====================

    public class CycleStateLog
    {
        public string IpAddress { get; set; }
        public DateTime ShiftDate { get; set; }
        public ShiftType ShiftType { get; set; }
        public MachineState State { get; set; }
        public int DurationSeconds { get; set; }
        public DateTime Timestamp { get; set; }
        public int? CycleNumber { get; set; }
    }

    public class CycleData
    {
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
        public int LoadingSeconds { get; set; }
        public int RunningSeconds { get; set; }
        public int TotalSeconds => LoadingSeconds + RunningSeconds;
    }

    public class CycleReport
    {
        public bool HasData { get; set; }
        public int TotalCycles { get; set; }
        public int AvgCycleTime { get; set; }
        public int MaxCycleTime { get; set; }
        public int MinCycleTime { get; set; }
        public int AnomalyCount { get; set; }
        public List<CycleData> Anomalies { get; set; } = new List<CycleData>();
    }
}
