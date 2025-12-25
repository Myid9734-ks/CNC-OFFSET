using System;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Data.SQLite;
using Microsoft.Win32;
using System.Text;

namespace FanucFocasTutorial.Forms
{
    public partial class AutoOffsetForm : Form
    {
        private const string REGISTRY_KEY = @"Software\FanucFocasTutorial";
        private const string DB_PATH_VALUE = "DBPath";
        private DataGridView gridOffset;
        private Label lblWarning;
        private Label lblCompensationStatus;
        private Label lblMeasurementTime;
        private Dictionary<string, List<decimal>> measurementHistory;
        private Dictionary<string, List<decimal>> diameterDifferenceHistory;
        private Dictionary<string, int> consecutiveNgCount; // 연속 NG 카운트 추가
        private Dictionary<string, bool> macroSetHistory; // 매크로 설정 이력 추가
        private const decimal SIGNIFICANT_CHANGE_THRESHOLD = 0.1m;
        private const decimal DIAMETER_DIFFERENCE_THRESHOLD = 0.03m;
        private const decimal EMERGENCY_STOP_THRESHOLD = 0.1m; // ±0.1mm 이상 보정 시 즉시 비상정지
        private const int REQUIRED_MEASUREMENTS = 2; // 2회로 변경
        private System.IO.FileSystemWatcher dbWatcher;
        private DateTime lastMeasurementTime = DateTime.MinValue;
        private DateTime lastDbFileModificationTime = DateTime.MinValue; // DB 파일 수정 시간 추가
        private Form parentForm;  // 부모 폼 참조 추가
        private System.Windows.Forms.Timer messageTimer; // 메시지 타이머 추가
        private System.Windows.Forms.Timer fastPollingTimer; // 빠른 폴링 타이머 추가
        private volatile bool isProcessingDbChange = false; // DB 변경 처리 중 플래그
        private string lastMeasurementDataHash = ""; // 마지막 측정 데이터 해시

        /* CNC 장비별 보정 규칙
         * #10 장비 (192.168.0.100:8193): 내경-상 (T8, X축) - 목표 18.000
         * #20 장비 (192.168.0.101:8194):
         *   - 전체높이 (T6, Z축) - 목표 54.70
         *   - 하단높이 (T7, Z축) - 목표 21.70
         *   - 폭 (T9, Z축) - 목표 17.15
         *   - 외경-상 (T9, X축) - 목표 23.05
         *   - 외경-하 (T7, X축) - 목표 23.05
         * 연속 NG 규칙: 1번째 NG는 대기, 2번째 NG는 1차 보정, 3-4번째 NG는 대기(보정값 확인), 5번째 NG는 매크로 #900=1 설정
         * 자동 옵셋 메뉴 클릭 후 첫 데이터는 무시
         */

        // 선택된 장비 연결 정보 (MainForm에서 전달받음)
        private CNCConnectionManager _connectionManager;
        private string _selectedIpAddress;

        // 자동 보정 실행 여부 제어
        private bool _autoCompensationEnabled = false;

        // 초기 로딩 여부 제어 (첫 로딩 시 보정 안함)
        private bool _isInitialLoading = true;

        // 보정 정보 수집을 위한 변수들
        private List<CompensationInfo> _pendingCompensations = new List<CompensationInfo>();
        private System.Windows.Forms.Timer _compensationTimer;

        // 로그 파일 관련 변수들
        private bool _isLoggingActive = false;
        private string _logFilePath;
        private int _currentMeasurementNumber = 0;
        private List<MeasurementLogData> _measurementLogHistory = new List<MeasurementLogData>();
        private DateTime _sessionStartTime;
        private DateTime _lastMeasurementTime;

        // 보정 정보 클래스
        private class CompensationInfo
        {
            public short ToolNo { get; set; }
            public string Item { get; set; }
            public string AxisType { get; set; }
            public decimal CompensationValue { get; set; }
            public string TargetEquipmentIP { get; set; }
            public bool IsWidthLinked { get; set; } = false;
        }

        // 측정 로그 데이터 클래스
        private class MeasurementLogData
        {
            public DateTime MeasurementTime { get; set; }
            public int MeasurementNumber { get; set; }
            public string Item { get; set; }
            public short ToolNumber { get; set; }
            public decimal MeasuredValue { get; set; }
            public decimal TargetValue { get; set; }
            public decimal Deviation { get; set; }
            public string Judgment { get; set; }
            public bool WasCompensated { get; set; }
            public decimal CompensationValue { get; set; }
            public int NgSequence { get; set; } // NG 몇 번째인지 (0=OK, 1=첫번째NG, 2=두번째NG, ...)
            public DateTime SessionStartTime { get; set; } // 세션 시작 시간
            public TimeSpan ElapsedFromStart { get; set; } // 세션 시작부터 경과 시간
            public TimeSpan IntervalFromPrevious { get; set; } // 이전 측정부터 경과 시간
            public int CycleNumber { get; set; } // 사이클 번호 (8개 항목이 1사이클)
        }

        public AutoOffsetForm(Form parentForm, CNCConnectionManager connectionManager, string selectedIpAddress)
        {
            this.parentForm = parentForm;
            _connectionManager = connectionManager;
            _selectedIpAddress = selectedIpAddress;
            measurementHistory = new Dictionary<string, List<decimal>>();
            diameterDifferenceHistory = new Dictionary<string, List<decimal>>();
            consecutiveNgCount = new Dictionary<string, int>(); // 연속 NG 카운트 초기화
            macroSetHistory = new Dictionary<string, bool>(); // 매크로 설정 이력 초기화
            
            // 메시지 타이머 초기화
            messageTimer = new System.Windows.Forms.Timer();
            messageTimer.Interval = 5000; // 5초
            messageTimer.Tick += (s, e) => {
                messageTimer.Stop();
                lblWarning.Text = "";
            };

            // 빠른 폴링 타이머 초기화 (네트워크 공유폴더 대응)
            fastPollingTimer = new System.Windows.Forms.Timer();
            fastPollingTimer.Interval = 1000; // 1초마다 체크
            fastPollingTimer.Tick += FastPollingTimer_Tick;

            // 보정 타이머 초기화 (보정 정보 수집 후 일괄 표시)
            _compensationTimer = new System.Windows.Forms.Timer();
            _compensationTimer.Interval = 500; // 0.5초 대기 후 일괄 처리
            _compensationTimer.Tick += CompensationTimer_Tick;
            
            InitializeComponent();
            LoadData();

            string dbPath = GetSavedDbPath();
            if (!string.IsNullOrEmpty(dbPath))
            {
                // 파일 감시 설정 개선 (네트워크 공유 폴더 대응)
                dbWatcher = new System.IO.FileSystemWatcher(Path.GetDirectoryName(dbPath));
                dbWatcher.Filter = Path.GetFileName(dbPath);
                dbWatcher.NotifyFilter = System.IO.NotifyFilters.LastWrite |
                                        System.IO.NotifyFilters.CreationTime |
                                        System.IO.NotifyFilters.Size |
                                        System.IO.NotifyFilters.Attributes;  // 네트워크 드라이브 속성 변경 감지
                dbWatcher.IncludeSubdirectories = false;
                dbWatcher.InternalBufferSize = 65536;  // 버퍼 크기 증가 (네트워크 지연 대응)

                dbWatcher.Changed += (s, e) => {
                    try {
                        // 즉시 처리 (지연 최소화)
                        if (!isProcessingDbChange)
                        {
                            isProcessingDbChange = true;
                            System.Threading.Timer delayTimer = null;
                            delayTimer = new System.Threading.Timer(_ => {
                                try {
                                    this.BeginInvoke(new Action(() => {
                                        lblWarning.Text = "데이터 변경 감지됨, 로드 중...";
                                        lblWarning.ForeColor = Color.Blue;
                                        CheckAndLoadDataImmediate();
                                        isProcessingDbChange = false;
                                    }));
                                } catch (Exception) {
                                    isProcessingDbChange = false;
                                }
                                delayTimer?.Dispose();
                            }, null, 100, System.Threading.Timeout.Infinite);  // 100ms로 단축
                        }
                    } catch (Exception) {
                        isProcessingDbChange = false;
                    }
                };
                dbWatcher.EnableRaisingEvents = true;
            }

            // 폼 종료 시 이벤트 추가
            this.FormClosing += AutoOffsetForm_FormClosing;

            // 수동으로 시작 버튼을 누를 때까지 대기
        }

        private void AutoOffsetForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;  // 폼이 완전히 닫히는 것을 방지
                this.Hide();      // 대신 숨김 처리
                parentForm.Show();  // 부모 폼 표시
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                // 로깅 중지
                StopLogging();

                if (dbWatcher != null)
                {
                    dbWatcher.Dispose();
                }
                if (messageTimer != null)
                {
                    messageTimer.Dispose();
                }
                if (fastPollingTimer != null)
                {
                    fastPollingTimer.Stop();
                    fastPollingTimer.Dispose();
                }
                if (_compensationTimer != null)
                {
                    _compensationTimer.Stop();
                    _compensationTimer.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        private void InitializeDataGridViewColumns()
        {
            // 그리드 컬럼 설정
            var columnWidths = new Dictionary<string, float>()
            {
                { "항목", 0.20f },
                { "공구번호", 0.10f },
                { "측정값", 0.15f },
                { "목표값", 0.15f },
                { "편차", 0.15f },
                { "판정", 0.12f },
                { "보정량", 0.13f }
            };

            foreach (var col in columnWidths)
            {
                var column = new DataGridViewTextBoxColumn
                {
                    HeaderText = col.Key,
                    Name = col.Key,
                    FillWeight = col.Value * 100,
                    SortMode = DataGridViewColumnSortMode.NotSortable,
                    HeaderCell = { Style = { Alignment = DataGridViewContentAlignment.MiddleCenter } },
                    DefaultCellStyle = new DataGridViewCellStyle
                    {
                        Alignment = DataGridViewContentAlignment.MiddleCenter,
                        Font = new Font("맑은 고딕", 9f),
                        SelectionBackColor = Color.Transparent,  // 선택 시 배경색 투명
                        SelectionForeColor = Color.Black  // 선택 시 글자색 그대로 유지
                    }
                };
                gridOffset.Columns.Add(column);
            }
        }

        private void InitializeComponent()
        {
            this.Text = "자동 옵셋";
            this.BackColor = Color.White;
            this.WindowState = FormWindowState.Maximized;  // DockStyle 대신 WindowState 사용

            // 그리드 설정
            gridOffset = new DataGridView
            {
                Location = new Point(0, 0),
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AllowUserToAddRows = false,
                ReadOnly = true,
                Font = new Font("맑은 고딕", 9f),
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.CellSelect,  // 셀 선택 모드
                MultiSelect = false,
                ColumnHeadersHeight = 25,
                RowTemplate = { Height = 23 },
                EnableHeadersVisualStyles = false,  // 헤더 스타일 비활성화
                StandardTab = false  // 탭 이동도 비활성화
            };

            // 그리드 컬럼 초기화
            InitializeDataGridViewColumns();

            // 마우스 클릭 이벤트 무시 (선택 방지)
            gridOffset.CellClick += (s, e) => {
                gridOffset.ClearSelection();
            };
            gridOffset.SelectionChanged += (s, e) => {
                gridOffset.ClearSelection();
            };

            // 버튼 패널
            var buttonPanel = new Panel
            {
                Height = 45,
                Dock = DockStyle.Bottom,
                Padding = new Padding(5),
                BackColor = Color.WhiteSmoke
            };

            // 보정 상태 표시 레이블
            lblCompensationStatus = new Label
            {
                Text = "대기중",
                Font = new Font("맑은 고딕", 9f, FontStyle.Bold),
                AutoSize = true,
                Location = new Point(110, 12),
                ForeColor = Color.DarkGray
            };

            // 측정 시간 표시 레이블 추가
            lblMeasurementTime = new Label
            {
                Text = "최근 측정: -",
                Font = new Font("맑은 고딕", 9f),
                AutoSize = true,
                Location = new Point(250, 12),
                ForeColor = Color.Black
            };

            buttonPanel.Controls.AddRange(new Control[] { lblCompensationStatus, lblMeasurementTime });

            // 경고 패널
            var warningPanel = new Panel
            {
                Height = 120,
                Dock = DockStyle.Bottom,
                Padding = new Padding(5),
                BackColor = Color.White,
                AutoScroll = true
            };

            // 경고 레이블
            lblWarning = new Label
            {
                Text = "",
                Font = new Font("맑은 고딕", 9f, FontStyle.Bold),
                AutoSize = true,
                Dock = DockStyle.Top,
                ForeColor = Color.Red,
                Padding = new Padding(0, 0, 0, 5),
                MaximumSize = new Size(this.Width - 20, 0)
            };

            warningPanel.Controls.Add(lblWarning);

            // 메인 패널 (그리드를 담을 컨테이너)
            var mainPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(5)
            };
            mainPanel.Controls.Add(gridOffset);

            // 폼에 컨트롤 추가
            this.Controls.AddRange(new Control[] { mainPanel, warningPanel, buttonPanel });
        }

        private string GetAxisType(string item)
        {
            // 항목에 따라 X축 또는 Z축 결정
            switch (item)
            {
                case "외경-상":
                case "외경-하":
                case "내경-상":
                case "내경-하":
                    return "X축";
                case "폭":
                case "하단높이":
                case "전체높이":
                    return "Z축";
                default:
                    return "X축"; // 기본값
            }
        }

        private void UpdateMeasurementTimeLabel()
        {
            if (lastMeasurementTime != DateTime.MinValue)
            {
                lblMeasurementTime.Text = $"최근 측정: {lastMeasurementTime:yyyy-MM-dd HH:mm}";
                lblMeasurementTime.ForeColor = Color.Black;
            }
            else
            {
                lblMeasurementTime.Text = "최근 측정: -";
                lblMeasurementTime.ForeColor = Color.Gray;
            }
        }

        private void CheckAndLoadData()
        {
            try
            {
                // 지연시간 단축 (기존 1000ms -> 200ms)
                System.Threading.Thread.Sleep(200);

                string dbPath = GetSavedDbPath();
                if (string.IsNullOrEmpty(dbPath) || !File.Exists(dbPath))
                {
                    lblWarning.Text = "DB 파일을 찾을 수 없습니다";
                    lblWarning.ForeColor = Color.Red;
                    messageTimer.Stop();
                    messageTimer.Start();
                    return;
                }

                // DB 파일 수정 시간 확인 - 변경된 경우에만 처리
                FileInfo fileInfo = new FileInfo(dbPath);
                DateTime currentDbModificationTime = fileInfo.LastWriteTime;

                if (lastDbFileModificationTime != DateTime.MinValue &&
                    currentDbModificationTime <= lastDbFileModificationTime)
                {
                    // 파일이 실제로 변경되지 않았으면 처리하지 않음
                    if (lblWarning.Text.Contains("로드 중") || lblWarning.Text.Contains("감지됨"))
                    {
                        lblWarning.Text = "";
                    }
                    return;
                }

                // 네트워크 파일 접근 재시도 로직 추가
                int retryCount = 0;
                const int maxRetries = 3;
                bool fileAccessible = false;

                while (retryCount < maxRetries && !fileAccessible)
                {
                    try
                    {
                        // 파일 접근 가능성 테스트
                        using (var testStream = new FileStream(dbPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                        {
                            fileAccessible = true;
                        }
                    }
                    catch (IOException)
                    {
                        retryCount++;
                        if (retryCount < maxRetries)
                        {
                            System.Threading.Thread.Sleep(500); // 500ms 대기 후 재시도
                        }
                    }
                }

                if (!fileAccessible)
                {
                    lblWarning.Text = "DB 파일에 접근할 수 없습니다 (파일이 사용 중)";
                    lblWarning.ForeColor = Color.Red;
                    messageTimer.Stop();
                    messageTimer.Start();
                    return;
                }

                using (var conn = new System.Data.SQLite.SQLiteConnection($"Data Source={dbPath};Version=3;"))
                {
                    conn.Open();
                    using (var cmd = new System.Data.SQLite.SQLiteCommand(conn))
                    {
                        cmd.CommandText = "SELECT measurement_time FROM measurements ORDER BY measurement_time DESC LIMIT 1";
                        var result = cmd.ExecuteScalar();
                        if (result != null && result != DBNull.Value)
                        {
                            string timeStr = result.ToString();
                            DateTime newMeasurementTime = DateTime.ParseExact(timeStr, "yyyyMMddHHmm", null);

                            // 새 측정 데이터가 있으면 즉시 로드
                            if (newMeasurementTime > lastMeasurementTime)
                            {
                                lastMeasurementTime = newMeasurementTime;
                                lastDbFileModificationTime = currentDbModificationTime; // DB 파일 수정 시간 업데이트
                                LoadData();
                                UpdateMeasurementTimeLabel();

                                // 새 데이터 로드 메시지 표시 및 타이머 시작
                                lblWarning.Text = "새로운 측정 데이터가 로드됨";
                                lblWarning.ForeColor = Color.Green;

                                // 타이머 재시작
                                messageTimer.Stop();
                                messageTimer.Start();
                            }
                            else
                            {
                                // 새 데이터가 없을 때 DB 파일 수정 시간만 업데이트
                                lastDbFileModificationTime = currentDbModificationTime;

                                // 메시지 제거
                                if (lblWarning.Text.Contains("로드 중") || lblWarning.Text.Contains("감지됨"))
                                {
                                    lblWarning.Text = "";
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                lblWarning.Text = $"데이터 확인 중 오류 발생: {ex.Message}";
                lblWarning.ForeColor = Color.Red;

                // 오류 메시지도 5초 후 삭제
                messageTimer.Stop();
                messageTimer.Start();
            }
        }

        // 즉시 데이터 로드 (지연 최소화)
        private void CheckAndLoadDataImmediate()
        {
            try
            {
                string dbPath = GetSavedDbPath();
                if (string.IsNullOrEmpty(dbPath) || !File.Exists(dbPath))
                {
                    lblWarning.Text = "DB 파일을 찾을 수 없습니다";
                    lblWarning.ForeColor = Color.Red;
                    messageTimer.Stop();
                    messageTimer.Start();
                    return;
                }

                // 파일 접근 가능성을 빠르게 테스트
                bool fileAccessible = false;
                try
                {
                    using (var testStream = new FileStream(dbPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        fileAccessible = true;
                    }
                }
                catch (IOException)
                {
                    // 파일이 잠겨있으면 빠른 폴링 모드로 전환
                    if (!fastPollingTimer.Enabled)
                    {
                        fastPollingTimer.Start();
                        lblWarning.Text = "DB 파일 접근 대기 중 (폴링 모드)...";
                        lblWarning.ForeColor = Color.Orange;
                    }
                    return;
                }

                if (fileAccessible)
                {
                    // 빠른 폴링 모드 해제
                    if (fastPollingTimer.Enabled)
                    {
                        fastPollingTimer.Stop();
                    }

                    // 데이터 해시 체크로 실제 변경 여부 확인
                    string currentDataHash = GetCurrentMeasurementDataHash(dbPath);
                    if (!string.IsNullOrEmpty(currentDataHash) && currentDataHash != lastMeasurementDataHash)
                    {
                        lastMeasurementDataHash = currentDataHash;
                        LoadData();
                        UpdateMeasurementTimeLabel();

                        lblWarning.Text = "새로운 측정 데이터가 로드됨";
                        lblWarning.ForeColor = Color.Green;
                        messageTimer.Stop();
                        messageTimer.Start();
                    }
                    else if (lblWarning.Text.Contains("로드 중") || lblWarning.Text.Contains("감지됨"))
                    {
                        lblWarning.Text = "";
                    }
                }
            }
            catch (Exception ex)
            {
                lblWarning.Text = $"즉시 데이터 확인 중 오류 발생: {ex.Message}";
                lblWarning.ForeColor = Color.Red;
                messageTimer.Stop();
                messageTimer.Start();
            }
        }

        // 빠른 폴링 타이머 이벤트
        private void FastPollingTimer_Tick(object sender, EventArgs e)
        {
            if (!isProcessingDbChange)
            {
                CheckAndLoadDataImmediate();
            }
        }

        // 현재 측정 데이터의 해시값 계산
        private string GetCurrentMeasurementDataHash(string dbPath)
        {
            try
            {
                using (var conn = new System.Data.SQLite.SQLiteConnection($"Data Source={dbPath};Version=3;"))
                {
                    conn.Open();
                    using (var cmd = new System.Data.SQLite.SQLiteCommand(conn))
                    {
                        cmd.CommandText = "SELECT * FROM measurements ORDER BY measurement_time DESC LIMIT 1";
                        using (var reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                // 모든 측정값을 문자열로 연결하여 해시 생성
                                var dataString = "";
                                for (int i = 0; i < reader.FieldCount; i++)
                                {
                                    dataString += reader.IsDBNull(i) ? "NULL" : reader.GetValue(i).ToString();
                                    dataString += "|";
                                }
                                return dataString.GetHashCode().ToString();
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                // 해시 계산 실패시 빈 문자열 반환
            }
            return "";
        }

        private void LoadData()
        {
            try
            {
                // DataGridView 컬럼이 제대로 초기화되었는지 확인
                if (gridOffset.Columns.Count == 0)
                {
                    InitializeDataGridViewColumns();
                }
                
                gridOffset.Rows.Clear();
                List<string> warningMessages = new List<string>();

                string savedDbPath = GetSavedDbPath();
                string dbPath;

                if (string.IsNullOrEmpty(savedDbPath) || !File.Exists(savedDbPath))
                {
                    using (OpenFileDialog openFileDialog = new OpenFileDialog())
                    {
                        openFileDialog.Filter = "DB 파일|*.db";
                        openFileDialog.Title = "측정 데이터 DB 파일을 선택하세요";
                        openFileDialog.InitialDirectory = !string.IsNullOrEmpty(savedDbPath) 
                            ? Path.GetDirectoryName(savedDbPath) 
                            : @"C:\db_data";

                        if (openFileDialog.ShowDialog() != DialogResult.OK)
                        {
                            MessageBox.Show("DB 파일을 선택하지 않았습니다.\n자동 옵셋 기능을 사용할 수 없습니다.",
                                          "알림",
                                          MessageBoxButtons.OK,
                                          MessageBoxIcon.Warning);
                            return;
                        }

                        dbPath = openFileDialog.FileName;
                        SaveDbPath(dbPath);
                    }
                }
                else
                {
                    dbPath = savedDbPath;
                }

                // 초기 로딩 시 DB 파일 수정 시간 기록
                if (lastDbFileModificationTime == DateTime.MinValue)
                {
                    FileInfo fileInfo = new FileInfo(dbPath);
                    lastDbFileModificationTime = fileInfo.LastWriteTime;
                }

                using (var conn = new System.Data.SQLite.SQLiteConnection($"Data Source={dbPath};Version=3;"))
                {
                    try
                    {
                        conn.Open();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"DB 연결 실패: {ex.Message}\n자동 옵셋 기능을 사용할 수 없습니다.",
                                      "오류",
                                      MessageBoxButtons.OK,
                                      MessageBoxIcon.Error);
                        return;
                    }

                    using (var cmd = new System.Data.SQLite.SQLiteCommand(conn))
                    {
                        cmd.CommandText = "SELECT * FROM measurements ORDER BY measurement_time DESC LIMIT 1";
                        using (var reader = cmd.ExecuteReader())
                        {
                            if (!reader.Read())
                            {
                                MessageBox.Show("측정 데이터가 없습니다.\n자동 옵셋 기능을 사용할 수 없습니다.",
                                              "알림",
                                              MessageBoxButtons.OK,
                                              MessageBoxIcon.Warning);
                                return;
                            }

                            try
                            {
                                // measurement_time 업데이트
                                if (!reader.IsDBNull(reader.GetOrdinal("measurement_time")))
                                {
                                    string timeStr = reader.GetString(reader.GetOrdinal("measurement_time"));
                                    lastMeasurementTime = DateTime.ParseExact(timeStr, "yyyyMMddHHmm", null);
                                    UpdateMeasurementTimeLabel();
                                }

                                // 2호기 데이터
                                AddMeasurementRow("외경-상", 9, 
                                    reader.IsDBNull(reader.GetOrdinal("outer_diameter_top")) ? 0m : reader.GetDecimal(reader.GetOrdinal("outer_diameter_top")), 
                                    23.050m, warningMessages);
                                AddMeasurementRow("외경-하", 7, 
                                    reader.IsDBNull(reader.GetOrdinal("outer_diameter_bottom")) ? 0m : reader.GetDecimal(reader.GetOrdinal("outer_diameter_bottom")), 
                                    23.050m, warningMessages);
                                AddMeasurementRow("폭", 9, 
                                    reader.IsDBNull(reader.GetOrdinal("width")) ? 0m : reader.GetDecimal(reader.GetOrdinal("width")), 
                                    17.187m, warningMessages);
                                AddMeasurementRow("하단높이", 7, 
                                    reader.IsDBNull(reader.GetOrdinal("bottom_height")) ? 0m : reader.GetDecimal(reader.GetOrdinal("bottom_height")), 
                                    21.700m, warningMessages);
                                AddMeasurementRow("전체높이", 6, 
                                    reader.IsDBNull(reader.GetOrdinal("total_height")) ? 0m : reader.GetDecimal(reader.GetOrdinal("total_height")), 
                                    54.700m, warningMessages);

                                // 1호기 데이터
                                AddMeasurementRow("내경-상", 8,
                                    reader.IsDBNull(reader.GetOrdinal("inner_diameter_top")) ? 0m : reader.GetDecimal(reader.GetOrdinal("inner_diameter_top")),
                                    18.000m, warningMessages);
                                AddMeasurementRow("내경-하", 8, 
                                    reader.IsDBNull(reader.GetOrdinal("inner_diameter_bottom")) ? 0m : reader.GetDecimal(reader.GetOrdinal("inner_diameter_bottom")), 
                                    17.9975m, warningMessages);
                                AddMeasurementRow("직각도", 3, 
                                    reader.IsDBNull(reader.GetOrdinal("perpendicularity")) ? 0m : reader.GetDecimal(reader.GetOrdinal("perpendicularity")), 
                                    0.0075m, warningMessages);
                            }
                            catch (Exception ex)
                            {
                                // 프로그램 종료 대신 오류 메시지 표시 후 계속 진행
                                lblWarning.Text = $"데이터 변환 중 오류 발생: {ex.Message}";
                                lblWarning.ForeColor = Color.Red;
                                messageTimer.Stop();
                                messageTimer.Start();
                                
                                // 기본 데이터로 대체
                                LoadDefaultData(warningMessages);
                                return;
                            }
                        }
                    }
                }

                // 첫 번째 데이터 이후부터 NG 카운팅 시작
                // _isInitialLoading은 첫 번째 측정 데이터가 들어올 때 false로 변경

                // 경고 메시지 표시
                if (warningMessages.Count > 0)
                {
                    lblWarning.Text = string.Join("\n", warningMessages);
                    // 경고 메시지는 계속 표시되어야 하므로 타이머 시작하지 않음
                }
                else
                {
                    lblWarning.Text = "데이터 로드 완료";
                    lblWarning.ForeColor = Color.Green;

                    // 로드 완료 메시지는 5초 후 삭제
                    messageTimer.Stop();
                    messageTimer.Start();
                }
            }
            catch (Exception ex)
            {
                // 프로그램 종료 대신 오류 메시지 표시 후 기본 데이터 로드
                lblWarning.Text = $"데이터 로드 중 오류 발생: {ex.Message}";
                lblWarning.ForeColor = Color.Red;
                messageTimer.Stop();
                messageTimer.Start();
                
                // 기본 데이터로 대체
                LoadDefaultData(new List<string>());
            }
        }

        private void LoadDefaultData(List<string> warningMessages)
        {
            // 2호기 데이터
            AddMeasurementRow("외경-상", 9, 23.050m, 23.050m, warningMessages);
            AddMeasurementRow("외경-하", 7, 23.015m, 23.050m, warningMessages);
            AddMeasurementRow("폭", 9, 17.180m, 17.187m, warningMessages);
            AddMeasurementRow("하단높이", 7, 21.650m, 21.700m, warningMessages);
            AddMeasurementRow("전체높이", 6, 54.700m, 54.700m, warningMessages);

            // 1호기 데이터
            AddMeasurementRow("내경-상", 8, 18.000m, 17.9975m, warningMessages);
            AddMeasurementRow("내경-하", 8, 17.985m, 17.9975m, warningMessages);
            AddMeasurementRow("직각도", 3, 0.012m, 0.0075m, warningMessages);
        }

        private void AddMeasurementRow(string item, short toolNo, decimal measuredValue, decimal targetValue, List<string> warningMessages)
        {
            string machine = toolNo <= 3 ? "1호기" : "2호기";
            string key = $"{machine} T{toolNo}번 {item}";
            
            // 이전 측정값과 비교하여 실제 새 측정인지 확인
            bool isNewMeasurement = true;
            if (measurementHistory.ContainsKey(key) && measurementHistory[key].Count > 0)
            {
                // 마지막 측정값과 동일하면 새 측정이 아님
                decimal lastMeasurement = measurementHistory[key].Last();
                if (Math.Abs(lastMeasurement - measuredValue) < 0.0001m) // 소수점 오차 고려
                {
                    isNewMeasurement = false;
                }
            }

            // 초기 로딩 시에는 모든 데이터를 새 측정이 아닌 것으로 처리
            if (_isInitialLoading)
            {
                isNewMeasurement = false;
                // 첫 번째 데이터 처리 후 초기 로딩 완료
                _isInitialLoading = false;
            }
            
            if (!measurementHistory.ContainsKey(key))
            {
                measurementHistory[key] = new List<decimal>();
            }
            
            var history = measurementHistory[key];
            
            // 실제 새 측정인 경우에만 이력에 추가 (단, 초기 로딩 시에는 제외)
            if (isNewMeasurement && !_isInitialLoading)
            {
                history.Add(measuredValue);

                // 최근 3개 측정값만 유지
                if (history.Count > REQUIRED_MEASUREMENTS)
                {
                    history.RemoveAt(0);
                }
            }
            
            // 외경 상하 차이 체크 (2호기 T9, T7)
            if (machine == "2호기" && (toolNo == 9 || toolNo == 7) && item.Contains("외경"))
            {
                CheckDiameterDifference(measuredValue, toolNo, targetValue, warningMessages);
            }

            // 편차 계산 (목표값과의 차이)
            decimal deviation = measuredValue - targetValue;

            // 공차 범위 가져오기
            (decimal lowerLimit, decimal upperLimit) = GetToleranceLimits(item);

            // 판정 및 보정량 결정
            string judgment;
            string compensation;
            bool needCompensation = false;  // 보정 필요 여부 플래그 추가
            bool shouldSetMacro = false;    // 매크로 설정 필요 여부 플래그 추가

            // 기본 보정량 계산 (기준값 - 측정값)
            decimal targetVal = GetTargetValue(item);
            decimal calculatedCompensation = targetVal - measuredValue;

            // 직각도는 보정하지 않는 항목이므로 보정량만 표시
            if (item == "직각도")
            {
                compensation = calculatedCompensation.ToString("F3");

                if (measuredValue >= lowerLimit && measuredValue <= upperLimit)
                {
                    judgment = "OK";
                }
                else
                {
                    judgment = measuredValue > upperLimit ? "HIGH" : "LOW";
                }
                measurementHistory.Remove(key);
            }
            else
            {
                if (measuredValue >= lowerLimit && measuredValue <= upperLimit)
                {
                    judgment = "OK";
                    compensation = calculatedCompensation.ToString("F3");
                    measurementHistory.Remove(key);

                    // OK인 경우 연속 NG 카운트 리셋
                    if (IsCompensationTarget(toolNo, item))
                    {
                        ProcessConsecutiveNG(key, false);
                    }
                }
                else
                {
                    // NG 처리 로직
                    bool shouldCompensate = false;

                    // 보정 대상 항목인지 확인 (초기 로딩 시에는 보정 안함)
                    if (IsCompensationTarget(toolNo, item) && !_isInitialLoading)
                    {
                        // ±0.1mm 이상 보정량인 경우 즉시 비상정지
                        if (Math.Abs(calculatedCompensation) >= EMERGENCY_STOP_THRESHOLD)
                        {
                            shouldSetMacro = true;
                            warningMessages.Add($"⚠️ {machine} T{toolNo}번 {item}: 보정량 {calculatedCompensation:F3}mm ≥ ±{EMERGENCY_STOP_THRESHOLD}mm → 즉시 비상정지");
                        }
                        else
                        {
                            shouldCompensate = ProcessConsecutiveNG(key, true);
                            shouldSetMacro = ShouldSetMacro(key);
                        }
                    }

                    if (shouldSetMacro)
                    {
                        // 비상정지 사유 구분
                        bool isLargeCompensation = Math.Abs(calculatedCompensation) >= EMERGENCY_STOP_THRESHOLD;
                        string emergencyReason = isLargeCompensation ? "과도한보정량" : "5회연속NG";

                        judgment = $"비상정지({emergencyReason})";
                        compensation = "Macro #900=1";

                        // 매크로 설정 이력 확인 후 실행
                        if (!macroSetHistory.ContainsKey(key) || !macroSetHistory[key])
                        {
                            macroSetHistory[key] = true;
                            if (_autoCompensationEnabled)
                            {
                                SimulateMacroSet(toolNo, item);
                            }
                        }
                    }
                    else if (shouldCompensate)
                    {
                        judgment = measuredValue > upperLimit ? "HIGH" : "LOW";
                        compensation = calculatedCompensation.ToString("F3");

                        // 자동 보정이 활성화된 경우에만 보정 실행
                        if (_autoCompensationEnabled)
                        {
                            SimulateCompensation(toolNo, item, calculatedCompensation);
                        }

                        // 하단높이 보정 시 폭 연동 보정
                        if (item == "하단높이")
                        {
                            HandleBottomHeightCompensation(calculatedCompensation);
                        }

                        needCompensation = true;
                    }
                    else
                    {
                        // 보정 대상이 아닌 항목이거나 첫 번째 NG인 경우
                        if (!IsCompensationTarget(toolNo, item))
                        {
                            // 보정하지 않는 항목 (직각도, 내경-하 등)
                            judgment = measuredValue > upperLimit ? "HIGH" : "LOW";
                            compensation = calculatedCompensation.ToString("F3");
                        }
                        else
                        {
                            // 초기 로딩이거나 NG 대기 상태 - 보정하지 않지만 보정량은 표시
                            if (_isInitialLoading)
                            {
                                judgment = measuredValue > upperLimit ? "HIGH" : "LOW";
                                compensation = calculatedCompensation.ToString("F3");
                            }
                            else
                            {
                                // NG 횟수에 따른 메시지 처리
                                int ngCount = consecutiveNgCount.ContainsKey(key) ? consecutiveNgCount[key] : 0;

                                if (ngCount == 1)
                                {
                                    judgment = measuredValue > upperLimit ? "HIGH(대기)" : "LOW(대기)";
                                    compensation = calculatedCompensation.ToString("F3");
                                    warningMessages.Add($"{machine} T{toolNo}번 {item}: 1번째 NG - 다음 측정 후 보정 예정");
                                }
                                else if (ngCount == 3 || ngCount == 4)
                                {
                                    judgment = measuredValue > upperLimit ? "HIGH(보정값확인)" : "LOW(보정값확인)";
                                    compensation = calculatedCompensation.ToString("F3");
                                    warningMessages.Add($"{machine} T{toolNo}번 {item}: {ngCount}번째 NG - 보정값 적용 여부 확인 중");
                                }
                                else
                                {
                                    judgment = measuredValue > upperLimit ? "HIGH(대기)" : "LOW(대기)";
                                    compensation = calculatedCompensation.ToString("F3");
                                }
                            }
                        }
                    }

                    if (needCompensation || shouldSetMacro)
                    {
                        measurementHistory.Remove(key);
                    }
                }
            }

            // 로그 데이터 기록 (첫 번째 데이터는 기본값으로 기록하지 않음)
            if (isNewMeasurement || _currentMeasurementNumber == 0)
            {
                int ngSequence = consecutiveNgCount.ContainsKey(key) ? consecutiveNgCount[key] : 0;
                bool wasCompensated = needCompensation || shouldSetMacro;
                string remarks = "";

                if (shouldSetMacro)
                {
                    remarks = "비상정지 매크로 설정";
                }
                else if (needCompensation)
                {
                    remarks = "자동 보정 실행";
                }
                else if (judgment.Contains("대기"))
                {
                    remarks = "보정 대기 상태";
                }

                LogMeasurementData(item, toolNo, measuredValue, targetValue, judgment,
                    wasCompensated, calculatedCompensation, ngSequence, remarks);
            }

            // DataGridView에 컬럼이 있는지 확인 후 행 추가
            if (gridOffset.Columns.Count == 0)
            {
                InitializeDataGridViewColumns();
            }

            var row = gridOffset.Rows[gridOffset.Rows.Add(
                item,
                toolNo.ToString(),
                measuredValue.ToString("F3"),
                targetValue.ToString("F4"),
                deviation.ToString("F4"),
                judgment,
                compensation
            )];

            if (judgment != "OK" && judgment != "측정중")
            {
                row.DefaultCellStyle.BackColor = Color.MistyRose;
            }

            if (item == "직각도")
            {
                row.DividerHeight = 2;
            }
        }

        private void CheckDiameterDifference(decimal measuredValue, short toolNo, decimal targetValue, List<string> warningMessages)
        {
            string key = "외경차";
            if (!diameterDifferenceHistory.ContainsKey(key))
            {
                diameterDifferenceHistory[key] = new List<decimal>();
            }

            // 상하 외경 측정값 가져오기
            decimal upperDiameter = 0;
            decimal lowerDiameter = 0;
            foreach (DataGridViewRow row in gridOffset.Rows)
            {
                if (row.Cells["공구번호"].Value == null) continue;
                
                string currentToolNo = row.Cells["공구번호"].Value.ToString();
                string item = row.Cells["항목"].Value.ToString();
                if (currentToolNo == "9" && item == "외경-상")
                {
                    upperDiameter = Convert.ToDecimal(row.Cells["측정값"].Value);
                }
                else if (currentToolNo == "7" && item == "외경-하")
                {
                    lowerDiameter = Convert.ToDecimal(row.Cells["측정값"].Value);
                }
            }

            if (upperDiameter != 0 && lowerDiameter != 0)
            {
                decimal difference = Math.Abs(upperDiameter - lowerDiameter);
                diameterDifferenceHistory[key].Add(difference);

                if (diameterDifferenceHistory[key].Count > REQUIRED_MEASUREMENTS)
                {
                    diameterDifferenceHistory[key].RemoveAt(0);
                }

                // 3회 연속 측정에서 차이가 임계값을 초과하는지 확인
                if (diameterDifferenceHistory[key].Count == REQUIRED_MEASUREMENTS &&
                    diameterDifferenceHistory[key].All(d => d > DIAMETER_DIFFERENCE_THRESHOLD))
                {
                    decimal avgDifference = diameterDifferenceHistory[key].Average();
                    warningMessages.Add($"⚠ 외경 상하 차이 과다: {avgDifference:F3}mm (기준: {DIAMETER_DIFFERENCE_THRESHOLD}mm)");
                }
            }
        }

        private bool CheckDiameterCompensationNeeded(short toolNo, decimal measuredValue, decimal targetValue)
        {
            if (diameterDifferenceHistory.ContainsKey("외경차") && 
                diameterDifferenceHistory["외경차"].Count == REQUIRED_MEASUREMENTS)
            {
                decimal avgDifference = diameterDifferenceHistory["외경차"].Average();
                if (avgDifference > DIAMETER_DIFFERENCE_THRESHOLD)
                {
                    // 다른 외경 측정값이 목표값에 더 가까운 경우에만 보정
                    decimal otherDiameter = GetOtherDiameterValue(toolNo);
                    decimal currentDeviation = Math.Abs(measuredValue - targetValue);
                    decimal otherDeviation = Math.Abs(otherDiameter - targetValue);
                    
                    return currentDeviation > otherDeviation;
                }
            }
            return false;
        }

        private decimal GetOtherDiameterValue(short toolNo)
        {
            foreach (DataGridViewRow row in gridOffset.Rows)
            {
                if (row.Cells["공구번호"].Value == null) continue;
                
                string currentToolNo = row.Cells["공구번호"].Value.ToString();
                string item = row.Cells["항목"].Value.ToString();
                
                // T9면 T7의 값을, T7이면 T9의 값을 찾음
                if ((toolNo == 9 && currentToolNo == "7" && item == "외경-하") ||
                    (toolNo == 7 && currentToolNo == "9" && item == "외경-상"))
                {
                    return Convert.ToDecimal(row.Cells["측정값"].Value);
                }
            }
            return 0;
        }

        private string GetSavedDbPath()
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(REGISTRY_KEY))
                {
                    if (key != null)
                    {
                        return key.GetValue(DB_PATH_VALUE) as string;
                    }
                }
            }
            catch (Exception)
            {
                // 레지스트리 읽기 실패시 무시
            }
            return null;
        }

        private void SaveDbPath(string dbPath)
        {
            try
            {
                using (var key = Registry.CurrentUser.CreateSubKey(REGISTRY_KEY))
                {
                    if (key != null)
                    {
                        key.SetValue(DB_PATH_VALUE, dbPath);
                    }
                }
            }
            catch (Exception)
            {
                // 레지스트리 저장 실패시 무시
            }
        }

        // 장비별 보정 대상 항목 확인
        private bool IsCompensationTarget(short toolNo, string item)
        {
            // #10 장비 - 내경-상만 보정
            if (toolNo == 8 && item == "내경-상") return true;

            // #20 장비 - 전체높이, 하단높이, 폭, 외경-상, 외경-하 보정
            if ((toolNo == 6 && item == "전체높이") ||
                (toolNo == 7 && item == "하단높이") ||
                (toolNo == 9 && item == "폭") ||
                (toolNo == 9 && item == "외경-상") ||
                (toolNo == 7 && item == "외경-하"))
                return true;

            return false;
        }

        // 선택된 설비 정보 반환
        private string GetSelectedEquipmentInfo()
        {
            return $"설비 ({_selectedIpAddress})";
        }

        // 항목별 목표값 반환
        private decimal GetTargetValue(string item)
        {
            switch (item)
            {
                case "내경-상": return 18.000m;
                case "내경-하": return 17.9975m;
                case "전체높이": return 54.70m;
                case "하단높이": return 21.70m;
                case "폭": return 17.15m;
                case "외경-상":
                case "외경-하": return 23.05m;
                default: return 0m;
            }
        }

        // 항목별 공차 범위 반환 메서드
        private (decimal lowerLimit, decimal upperLimit) GetToleranceLimits(string item)
        {
            switch (item)
            {
                case "외경-상":
                case "외경-하":
                    return (23.020m, 23.080m);
                case "폭":
                    return (17.15m, 17.199m);
                case "하단높이":
                    return (21.665m, 21.735m);
                case "전체높이":
                    return (54.650m, 54.750m);
                case "내경-상":
                case "내경-하":
                    return (17.990m, 18.005m);
                case "직각도":
                    return (0.000m, 0.015m);
                default:
                    return (0m, 0m); // 기본값
            }
        }

        // 자동 프로세스 시작 메서드
        public void StartAutoProcess()
        {
            try
            {
                // 로그 파일 생성 및 로깅 시작
                StartLogging();

                // 자동 보정 활성화
                _autoCompensationEnabled = true;

                // 데이터 로드 (이미 LoadData 메서드가 있음)
                LoadData();

                // 상태 업데이트
                lblCompensationStatus.Text = "자동 옵셋 처리 중... (로깅 활성화)";
                lblCompensationStatus.ForeColor = Color.Blue;

                // 파일 감시 시작
                if (dbWatcher != null)
                {
                    dbWatcher.EnableRaisingEvents = true;
                }

                // 주기적으로 데이터 확인 (네트워크 공유 폴더를 위해 간격 단축)
                System.Windows.Forms.Timer checkTimer = new System.Windows.Forms.Timer();
                checkTimer.Interval = 30000; // 30초마다
                checkTimer.Tick += (s, e) => CheckAndLoadData();
                checkTimer.Start();

                // 빠른 폴링 타이머도 함께 시작 (실시간 감지 강화)
                if (!fastPollingTimer.Enabled)
                {
                    fastPollingTimer.Start();
                }
            }
            catch (Exception ex)
            {
                lblWarning.Text = $"자동 시작 중 오류 발생: {ex.Message}";
                lblWarning.ForeColor = Color.Red;
            }
        }

        // 자동 프로세스 중지 메서드
        public void StopAutoProcess()
        {
            // 로깅 중지
            StopLogging();

            _autoCompensationEnabled = false;
            lblCompensationStatus.Text = "자동 옵셋 중지됨 (로깅 중지됨)";
            lblCompensationStatus.ForeColor = Color.Gray;

            // 빠른 폴링 타이머도 중지
            if (fastPollingTimer.Enabled)
            {
                fastPollingTimer.Stop();
            }
        }

        // CNC 보정 시뮬레이션 메서드 (보정 정보 수집)
        private void SimulateCompensation(short toolNo, string item, decimal compensationValue)
        {
            try
            {
                string axisType = GetAxisType(item);
                string targetEquipmentIP = GetTargetEquipmentIP(toolNo);

                // 보정 정보를 리스트에 추가
                _pendingCompensations.Add(new CompensationInfo
                {
                    ToolNo = toolNo,
                    Item = item,
                    AxisType = axisType,
                    CompensationValue = compensationValue,
                    TargetEquipmentIP = targetEquipmentIP,
                    IsWidthLinked = false
                });

                // 타이머 재시작 (추가 보정이 있을 수 있으므로 대기)
                _compensationTimer.Stop();
                _compensationTimer.Start();
            }
            catch (Exception ex)
            {
                lblWarning.Text = $"보정 시뮬레이션 중 오류 발생: {ex.Message}";
                lblWarning.ForeColor = Color.Red;
                messageTimer.Stop();
                messageTimer.Start();
            }
        }

        // 매크로 변수 관리 (A부터 시작, 필요시 B,C 추가)
        private List<MacroSlotInfo> activeMacroSlots = new List<MacroSlotInfo>(); // 현재 활성 슬롯들
        private DateTime lastCompensationTime = DateTime.MinValue; // 마지막 보정 시간

        // 매크로 슬롯 정보
        private class MacroSlotInfo
        {
            public int SlotIndex { get; set; } // 0=A(#100~102), 1=B(#103~105), 2=C(#106~108)
            public short ToolNo { get; set; }
            public DateTime UsedTime { get; set; }
            public decimal XValue { get; set; }
            public decimal ZValue { get; set; }
        }

        // 공구번호에 따른 장비 IP 결정
        private string GetTargetEquipmentIP(short toolNo)
        {
            // #10 장비 (192.168.0.100): T8 (내경-상)
            // #20 장비 (192.168.0.101): T6(전체높이), T7(하단높이, 외경-하), T9(폭, 외경-상)

            if (toolNo == 8)
            {
                return "192.168.0.100"; // #10 장비
            }
            else if (toolNo == 6 || toolNo == 7 || toolNo == 9)
            {
                return "192.168.0.101"; // #20 장비
            }

            // 기본값 (혹시 모를 다른 공구번호)
            return _selectedIpAddress;
        }

        // 매크로 슬롯 할당 (항상 A부터 시작, 필요시 B,C 추가)
        private int GetMacroSlot(short toolNo)
        {
            DateTime now = DateTime.Now;

            // 새로운 보정 사이클인지 확인 (예: 5분 이상 지났으면 새 사이클)
            bool isNewCycle = (now - lastCompensationTime).TotalMinutes > 5;

            if (isNewCycle)
            {
                // 새 사이클 시작 - A부터 다시 시작
                activeMacroSlots.Clear();
                lastCompensationTime = now;
            }

            // 이미 해당 공구가 활성 슬롯에 있는지 확인
            var existingSlot = activeMacroSlots.FirstOrDefault(slot => slot.ToolNo == toolNo);
            if (existingSlot != null)
            {
                existingSlot.UsedTime = now;
                return existingSlot.SlotIndex;
            }

            // 새로운 슬롯 추가 (A=0, B=1, C=2 순서)
            int nextSlotIndex = activeMacroSlots.Count; // 0부터 시작
            const int maxSlots = 4; // A, B, C, D까지

            if (nextSlotIndex < maxSlots)
            {
                activeMacroSlots.Add(new MacroSlotInfo
                {
                    SlotIndex = nextSlotIndex,
                    ToolNo = toolNo,
                    UsedTime = now,
                    XValue = 0,
                    ZValue = 0
                });

                return nextSlotIndex;
            }

            // 최대 슬롯 수 초과시 가장 오래된 슬롯 재사용
            var oldestSlot = activeMacroSlots.OrderBy(slot => slot.UsedTime).First();
            oldestSlot.ToolNo = toolNo;
            oldestSlot.UsedTime = now;
            oldestSlot.XValue = 0;
            oldestSlot.ZValue = 0;

            return oldestSlot.SlotIndex;
        }

        // 보정 타이머 이벤트 (수집된 보정 정보를 일괄 처리)
        private void CompensationTimer_Tick(object sender, EventArgs e)
        {
            _compensationTimer.Stop();

            if (_pendingCompensations.Count == 0)
                return;

            try
            {
                var allCompensationResults = new List<string>();
                bool anySuccess = false;

                foreach (var compensation in _pendingCompensations)
                {
                    var result = ProcessSingleCompensation(compensation);
                    allCompensationResults.Add(result.Message);
                    if (result.Success)
                        anySuccess = true;
                }

                // 통합 메시지 박스 표시
                string fullMessage = string.Join("\n\n" + new string('-', 50) + "\n\n", allCompensationResults);

                this.BeginInvoke(new Action(() => {
                    MessageBox.Show(
                        fullMessage,
                        anySuccess ? "매크로 변수 설정 완료" : "매크로 변수 설정 시뮬레이션",
                        MessageBoxButtons.OK,
                        anySuccess ? MessageBoxIcon.Information : MessageBoxIcon.Warning
                    );
                }));
            }
            catch (Exception ex)
            {
                lblWarning.Text = $"보정 처리 중 오류 발생: {ex.Message}";
                lblWarning.ForeColor = Color.Red;
                messageTimer.Stop();
                messageTimer.Start();
            }
            finally
            {
                _pendingCompensations.Clear();
            }
        }

        // 개별 보정 처리 메서드
        private (bool Success, string Message) ProcessSingleCompensation(CompensationInfo compensation)
        {
            try
            {
                // 비상정지 매크로 설정인 경우 별도 처리
                if (compensation.AxisType == "매크로 #900")
                {
                    return ProcessEmergencyStopMacro(compensation);
                }

                // 매크로 슬롯 할당 (A부터 시작, 필요시 B,C 추가)
                int slotIndex = GetMacroSlot(compensation.ToolNo);
                int baseAddress = 100 + (slotIndex * 3); // #100부터 시작

                // 해당 공구의 슬롯 정보 가져오기
                var slotInfo = activeMacroSlots.First(slot => slot.ToolNo == compensation.ToolNo && slot.SlotIndex == slotIndex);

                // 현재 보정값 설정 및 저장
                if (compensation.AxisType.Contains("X"))
                {
                    slotInfo.XValue = compensation.CompensationValue;
                }
                else
                {
                    slotInfo.ZValue = compensation.CompensationValue;
                }

                // 슬롯 이름 결정 (A, B, C, D)
                string slotName = ((char)('A' + slotIndex)).ToString();
                string allActiveSlots = string.Join(",", activeMacroSlots.OrderBy(s => s.SlotIndex).Select(s => ((char)('A' + s.SlotIndex)).ToString()));

                string macroMessage = $"T{compensation.ToolNo} {compensation.Item} ({compensation.AxisType}) 보정";
                if (compensation.IsWidthLinked)
                {
                    macroMessage += " [하단높이 연동]";
                }
                macroMessage += $"\n매크로 변수 설정:" +
                               $"\n  #{baseAddress} = {compensation.ToolNo} (공구번호)" +
                               $"\n  #{baseAddress + 1} = {slotInfo.XValue:F3} (X축 보정량)" +
                               $"\n  #{baseAddress + 2} = {slotInfo.ZValue:F3} (Z축 보정량)" +
                               $"\n  [슬롯 {slotName} 사용] - 현재 활성: {allActiveSlots}";

                // 실제 CNC에 매크로 변수 설정
                bool macroSetSuccess = WriteMacroVariables(baseAddress, compensation.ToolNo, slotInfo.XValue, slotInfo.ZValue, compensation.TargetEquipmentIP);

                // 더 상세한 디버깅 정보 가져오기
                var connection = _connectionManager?.GetConnection(compensation.TargetEquipmentIP);
                string connectionStatus = connection == null ? "null" :
                                        connection.IsConnected ? "Connected" : "Disconnected";

                string debugInfo = $"\n\n디버깅 정보:" +
                                 $"\n- 선택된 IP: {_selectedIpAddress}" +
                                 $"\n- 대상 장비 IP: {compensation.TargetEquipmentIP}" +
                                 $"\n- 연결 상태: {connectionStatus}" +
                                 $"\n- 핸들: {connection?.Handle ?? 0}" +
                                 $"\n- 매크로 주소: #{baseAddress}, #{baseAddress + 1}, #{baseAddress + 2}" +
                                 $"\n- 설정할 값: {compensation.ToolNo}, {(int)(slotInfo.XValue * 1000)}, {(int)(slotInfo.ZValue * 1000)}" +
                                 $"\n- 슬롯 재사용: {slotIndex}번 슬롯 (시간: {slotInfo.UsedTime:HH:mm:ss})";

                // 결과에 따라 메시지 수정
                if (macroSetSuccess)
                {
                    macroMessage += "\n\n✅ 실제 CNC 매크로 변수 설정 완료" + debugInfo;
                }
                else
                {
                    macroMessage += "\n\n⚠️ CNC 연결 실패 - 시뮬레이션 모드" + debugInfo;
                }

                return (macroSetSuccess, macroMessage);
            }
            catch (Exception ex)
            {
                return (false, $"T{compensation.ToolNo} {compensation.Item} 보정 중 오류: {ex.Message}");
            }
        }

        // 비상정지 매크로 처리 메서드
        private (bool Success, string Message) ProcessEmergencyStopMacro(CompensationInfo compensation)
        {
            try
            {
                var connection = _connectionManager?.GetConnection(compensation.TargetEquipmentIP);
                string connectionStatus = connection == null ? "null" :
                                        connection.IsConnected ? "Connected" : "Disconnected";

                string macroMessage = $"T{compensation.ToolNo} {compensation.Item} - 비상정지 매크로 설정" +
                                     $"\n매크로 변수 설정:" +
                                     $"\n  #900 = 1 (비상정지 플래그)";

                bool macroSetSuccess = false;

                // CNC 연결이 가능한 경우 실제 매크로 설정
                if (connection != null && connection.IsConnected)
                {
                    try
                    {
                        ushort handle = connection.Handle;
                        short result = Focas1.cnc_wrmacro(handle, 900, 10, 1, 0); // #900 = 1 설정
                        macroSetSuccess = (result == Focas1.EW_OK);

                        if (!macroSetSuccess)
                        {
                            macroMessage += $"\n\n⚠️ 매크로 #900 설정 실패: {GetErrorMessage(result)} (코드: {result})";
                        }
                    }
                    catch (Exception ex)
                    {
                        macroMessage += $"\n\n⚠️ 매크로 설정 중 예외 발생: {ex.Message}";
                    }
                }

                string debugInfo = $"\n\n디버깅 정보:" +
                                 $"\n- 선택된 IP: {_selectedIpAddress}" +
                                 $"\n- 대상 장비 IP: {compensation.TargetEquipmentIP}" +
                                 $"\n- 연결 상태: {connectionStatus}" +
                                 $"\n- 핸들: {connection?.Handle ?? 0}" +
                                 $"\n- 매크로 주소: #900" +
                                 $"\n- 설정할 값: 1 (비상정지)";

                // 결과에 따라 메시지 수정
                if (macroSetSuccess)
                {
                    macroMessage += "\n\n🚨 실제 CNC 비상정지 매크로 설정 완료" + debugInfo;
                }
                else
                {
                    macroMessage += "\n\n⚠️ CNC 연결 실패 - 비상정지 시뮬레이션 모드" + debugInfo;
                }

                return (macroSetSuccess, macroMessage);
            }
            catch (Exception ex)
            {
                return (false, $"T{compensation.ToolNo} {compensation.Item} 비상정지 매크로 처리 중 오류: {ex.Message}");
            }
        }


        // 실제 CNC에 매크로 변수 쓰기
        private bool WriteMacroVariables(int baseAddress, short toolNo, decimal xValue, decimal zValue, string targetEquipmentIP)
        {
            try
            {
                // 대상 장비 IP 주소로 연결 가져오기
                var connection = _connectionManager?.GetConnection(targetEquipmentIP);
                if (connection == null || !connection.IsConnected)
                {
                    return false; // 연결되지 않음
                }

                ushort handle = connection.Handle;

                // 매크로 변수 3개 설정
                // #baseAddress = 공구번호 (정수)
                // #baseAddress+1 = X축 보정량 (0.001mm 단위)
                // #baseAddress+2 = Z축 보정량 (0.001mm 단위)

                // FOCAS cnc_wrmacro 함수 호출 - 기존 읽기 방식과 동일하게 10 사용
                // MainForm에서 cnc_rdmacro는 두 번째 매개변수로 10을 사용함
                short result1 = Focas1.cnc_wrmacro(handle, (short)baseAddress, 10, toolNo, 0);
                short result2 = Focas1.cnc_wrmacro(handle, (short)(baseAddress + 1), 10, (int)(xValue * 1000), 3); // 소수점 3자리
                short result3 = Focas1.cnc_wrmacro(handle, (short)(baseAddress + 2), 10, (int)(zValue * 1000), 3); // 소수점 3자리

                if (result1 == Focas1.EW_OK && result2 == Focas1.EW_OK && result3 == Focas1.EW_OK)
                {
                    // 로그 기록
                    lblWarning.Text = $"매크로 변수 설정 성공: T{toolNo} #{baseAddress}~{baseAddress + 2}";
                    lblWarning.ForeColor = Color.Green;
                    return true;
                }
                else
                {
                    string errorMsg = $"매크로 변수 설정 실패:\n" +
                                    $"#{baseAddress}(공구): {GetErrorMessage(result1)} (코드: {result1})\n" +
                                    $"#{baseAddress + 1}(X축): {GetErrorMessage(result2)} (코드: {result2})\n" +
                                    $"#{baseAddress + 2}(Z축): {GetErrorMessage(result3)} (코드: {result3})";
                    lblWarning.Text = errorMsg;
                    lblWarning.ForeColor = Color.Red;

                    // 더 상세한 오류 정보를 디버깅 정보에 추가
                    this.BeginInvoke(new Action(() => {
                        MessageBox.Show(
                            $"FOCAS 오류 상세:\n" +
                            $"핸들: {handle}\n" +
                            $"매크로 #{baseAddress} = {toolNo} → 결과: {result1} ({GetErrorMessage(result1)})\n" +
                            $"매크로 #{baseAddress + 1} = {(int)(xValue * 1000)} → 결과: {result2} ({GetErrorMessage(result2)})\n" +
                            $"매크로 #{baseAddress + 2} = {(int)(zValue * 1000)} → 결과: {result3} ({GetErrorMessage(result3)})",
                            "FOCAS 오류 분석",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error
                        );
                    }));

                    return false;
                }
            }
            catch (Exception ex)
            {
                lblWarning.Text = $"매크로 변수 쓰기 오류: {ex.Message}";
                lblWarning.ForeColor = Color.Red;
                return false;
            }
        }

        // FOCAS 오류 코드를 문자열로 변환
        private string GetErrorMessage(short errorCode)
        {
            switch (errorCode)
            {
                case 0: // EW_OK
                    return "성공";
                case 1: // EW_FUNC or EW_NOPMC
                    return "함수 사용 불가 또는 PMC 없음";
                case 2: // EW_LENGTH
                    return "데이터 블록 길이 오류";
                case 3: // EW_NUMBER
                    return "데이터 번호 오류";
                case 4: // EW_ATTRIB
                    return "데이터 속성 오류";
                case 5: // EW_DATA
                    return "데이터 값 오류";
                case 6: // EW_NOOPT
                    return "옵션 없음";
                case 7: // EW_PROT
                    return "쓰기 보호됨";
                case 8: // EW_OVRFLOW
                    return "메모리 오버플로";
                case 9: // EW_PARAM
                    return "매개변수 오류";
                case 10: // EW_BUFFER
                    return "버퍼 오류";
                case 11: // EW_PATH
                    return "경로 오류";
                case 12: // EW_MODE
                    return "모드 오류";
                case 13: // EW_REJECT
                    return "실행 거부";
                case 14: // EW_DTSRVR
                    return "데이터 서버 오류";
                case -1: // EW_BUSY
                    return "처리 중";
                case -2: // EW_RESET
                    return "리셋 또는 정지 발생";
                case -8: // EW_HANDLE
                    return "핸들 오류";
                case -16: // EW_SOCKET
                    return "소켓 오류";
                case -17: // EW_PROTOCOL
                    return "프로토콜 오류";
                default:
                    return $"알 수 없는 오류({errorCode})";
            }
        }

        // 매크로 설정 시뮬레이션 메서드 (비상정지 매크로 설정)
        private void SimulateMacroSet(short toolNo, string item)
        {
            try
            {
                string equipmentInfo = GetSelectedEquipmentInfo();
                string targetEquipmentIP = GetTargetEquipmentIP(toolNo);

                // 비상정지 매크로 정보를 보정 리스트에 추가
                _pendingCompensations.Add(new CompensationInfo
                {
                    ToolNo = toolNo,
                    Item = $"{item} (비상정지)",
                    AxisType = "매크로 #900",
                    CompensationValue = 1, // #900 = 1 설정
                    TargetEquipmentIP = targetEquipmentIP,
                    IsWidthLinked = false
                });

                // 타이머 재시작
                _compensationTimer.Stop();
                _compensationTimer.Start();

                // 로그 기록
                lblWarning.Text = $"비상정지 매크로 설정 예약: T{toolNo} {item}";
                lblWarning.ForeColor = Color.Orange;
            }
            catch (Exception ex)
            {
                lblWarning.Text = $"매크로 설정 시뮬레이션 중 오류 발생: {ex.Message}";
                lblWarning.ForeColor = Color.Red;
                messageTimer.Stop();
                messageTimer.Start();
            }
        }

        // 연속 NG 처리 로직
        private bool ProcessConsecutiveNG(string key, bool isNG)
        {
            if (!consecutiveNgCount.ContainsKey(key))
            {
                consecutiveNgCount[key] = 0;
            }

            if (isNG)
            {
                consecutiveNgCount[key]++;

                // 첫 번째 NG - 대기 (보정 안함)
                if (consecutiveNgCount[key] == 1)
                {
                    return false; // 보정 안함
                }
                // 두 번째 NG - 1차 보정 실행
                else if (consecutiveNgCount[key] == 2)
                {
                    return true; // 1차 보정 실시
                }
                // 세 번째, 네 번째 NG - 대기 (보정값 적용 여부 확인, 보정 안함)
                else if (consecutiveNgCount[key] == 3 || consecutiveNgCount[key] == 4)
                {
                    return false; // 대기 (보정 안함)
                }
                // 다섯 번째 NG - 매크로 설정
                else if (consecutiveNgCount[key] >= 5)
                {
                    return false; // 매크로 설정을 위해 false 반환
                }
            }
            else
            {
                // OK인 경우 카운트 리셋
                consecutiveNgCount[key] = 0;
            }

            return false;
        }

        // 매크로 설정 필요 여부 확인
        private bool ShouldSetMacro(string key)
        {
            if (!consecutiveNgCount.ContainsKey(key))
                return false;

            // 보정 후에도 3회 연속 NG인 경우
            return consecutiveNgCount[key] >= 5;
        }

        // 하단높이 보정 시 폭 자동 조정 처리
        private void HandleBottomHeightCompensation(decimal bottomHeightCompensation)
        {
            try
            {
                // 하단높이 보정에 따른 폭의 연동 보정
                // 하단높이가 -방향 보정이면 폭도 -방향 보정 (같은 방향)
                // 하단높이가 +방향 보정이면 폭도 +방향 보정 (같은 방향)
                decimal widthAdjustment = bottomHeightCompensation;

                // 실제 CNC에 폭 보정 적용 (T9번 도구)
                if (_autoCompensationEnabled)
                {
                    ApplyWidthCompensation(widthAdjustment);
                }
            }
            catch (Exception ex)
            {
                lblWarning.Text = $"폭 연동 보정 중 오류 발생: {ex.Message}";
                lblWarning.ForeColor = Color.Red;
                messageTimer.Stop();
                messageTimer.Start();
            }
        }

        // 폭 보정 적용 (T9번 도구)
        private void ApplyWidthCompensation(decimal compensationValue)
        {
            try
            {
                short toolNo = 9; // 폭은 T9번 도구
                string axisType = "Z축"; // 폭은 Z축
                string targetEquipmentIP = GetTargetEquipmentIP(toolNo);

                // 폭 연동 보정 정보를 리스트에 추가
                _pendingCompensations.Add(new CompensationInfo
                {
                    ToolNo = toolNo,
                    Item = "폭",
                    AxisType = axisType,
                    CompensationValue = compensationValue,
                    TargetEquipmentIP = targetEquipmentIP,
                    IsWidthLinked = true
                });

                // 타이머 재시작
                _compensationTimer.Stop();
                _compensationTimer.Start();

                // 로그 기록
                lblWarning.Text = $"폭 연동 보정 적용: T9 Z축 {compensationValue:F3}mm";
                lblWarning.ForeColor = Color.Blue;
            }
            catch (Exception ex)
            {
                lblWarning.Text = $"폭 보정 적용 중 오류: {ex.Message}";
                lblWarning.ForeColor = Color.Red;
            }
        }

        // 로그 파일 관련 메서드들

        /// <summary>
        /// 로깅 시작 - 로그 파일 생성 및 헤더 작성
        /// </summary>
        private void StartLogging()
        {
            try
            {
                // 로그 파일 경로 생성
                string logDirectory = Path.Combine(Application.StartupPath, "Logs");
                if (!Directory.Exists(logDirectory))
                {
                    Directory.CreateDirectory(logDirectory);
                }

                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                _logFilePath = Path.Combine(logDirectory, $"AutoOffset_Log_{timestamp}.csv");

                // 로그 파일 헤더 작성
                WriteLogHeader();

                // 로깅 활성화
                _isLoggingActive = true;
                _currentMeasurementNumber = 0;
                _sessionStartTime = DateTime.Now;
                _lastMeasurementTime = DateTime.MinValue;

                lblWarning.Text = $"로그 파일 생성됨: {Path.GetFileName(_logFilePath)}";
                lblWarning.ForeColor = Color.Green;
                messageTimer.Stop();
                messageTimer.Start();
            }
            catch (Exception ex)
            {
                lblWarning.Text = $"로그 파일 생성 실패: {ex.Message}";
                lblWarning.ForeColor = Color.Red;
                messageTimer.Stop();
                messageTimer.Start();
            }
        }

        /// <summary>
        /// 로깅 중지
        /// </summary>
        private void StopLogging()
        {
            try
            {
                if (_isLoggingActive)
                {
                    // 로그 파일 마무리 작성
                    WriteLogFooter();
                    _isLoggingActive = false;

                    lblWarning.Text = $"로그 파일 저장 완료: {Path.GetFileName(_logFilePath)}";
                    lblWarning.ForeColor = Color.Green;
                    messageTimer.Stop();
                    messageTimer.Start();
                }
            }
            catch (Exception ex)
            {
                lblWarning.Text = $"로그 파일 저장 실패: {ex.Message}";
                lblWarning.ForeColor = Color.Red;
                messageTimer.Stop();
                messageTimer.Start();
            }
        }

        /// <summary>
        /// 로그 파일 헤더 작성
        /// </summary>
        private void WriteLogHeader()
        {
            var sb = new StringBuilder();
            sb.AppendLine("=== 자동 옵셋 로그 파일 ===");
            sb.AppendLine($"생성 시간: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine($"설비 IP: {_selectedIpAddress}");
            sb.AppendLine("");
            sb.AppendLine("측정번호,측정시간,항목,공구번호,측정값,목표값,편차,판정,보정여부,보정값,NG순서,비고,세션경과시간(초),측정간격(초),사이클번호");

            File.WriteAllText(_logFilePath, sb.ToString(), Encoding.UTF8);
        }

        /// <summary>
        /// 로그 파일 마무리 작성
        /// </summary>
        private void WriteLogFooter()
        {
            var sb = new StringBuilder();
            sb.AppendLine("");
            sb.AppendLine("=== 로그 파일 종료 ===");
            sb.AppendLine($"종료 시간: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine($"총 측정 횟수: {_currentMeasurementNumber}");

            File.AppendAllText(_logFilePath, sb.ToString(), Encoding.UTF8);
        }

        /// <summary>
        /// 측정 데이터를 로그 파일에 기록
        /// </summary>
        private void LogMeasurementData(string item, short toolNo, decimal measuredValue, decimal targetValue,
            string judgment, bool wasCompensated, decimal compensationValue, int ngSequence, string remarks = "")
        {
            if (!_isLoggingActive)
                return;

            try
            {
                _currentMeasurementNumber++;
                DateTime currentTime = DateTime.Now;

                // 시간 계산
                TimeSpan elapsedFromStart = currentTime - _sessionStartTime;
                TimeSpan intervalFromPrevious = _lastMeasurementTime == DateTime.MinValue ?
                    TimeSpan.Zero : currentTime - _lastMeasurementTime;

                // 사이클 번호 계산 (8개 항목이 1사이클)
                int cycleNumber = (int)Math.Ceiling(_currentMeasurementNumber / 8.0);

                var logData = new MeasurementLogData
                {
                    MeasurementTime = currentTime,
                    MeasurementNumber = _currentMeasurementNumber,
                    Item = item,
                    ToolNumber = toolNo,
                    MeasuredValue = measuredValue,
                    TargetValue = targetValue,
                    Deviation = measuredValue - targetValue,
                    Judgment = judgment,
                    WasCompensated = wasCompensated,
                    CompensationValue = compensationValue,
                    NgSequence = ngSequence,
                    SessionStartTime = _sessionStartTime,
                    ElapsedFromStart = elapsedFromStart,
                    IntervalFromPrevious = intervalFromPrevious,
                    CycleNumber = cycleNumber
                };

                _measurementLogHistory.Add(logData);

                // CSV 형식으로 로그 기록
                var sb = new StringBuilder();
                sb.AppendFormat("{0},{1},{2},{3},{4:F3},{5:F3},{6:F3},{7},{8},{9:F3},{10},{11},{12:F1},{13:F1},{14}",
                    logData.MeasurementNumber,
                    logData.MeasurementTime.ToString("yyyy-MM-dd HH:mm:ss"),
                    logData.Item,
                    logData.ToolNumber,
                    logData.MeasuredValue,
                    logData.TargetValue,
                    logData.Deviation,
                    logData.Judgment,
                    wasCompensated ? "보정됨" : "대기",
                    logData.CompensationValue,
                    ngSequence == 0 ? "OK" : $"{ngSequence}번째NG",
                    remarks,
                    logData.ElapsedFromStart.TotalSeconds,
                    logData.IntervalFromPrevious.TotalSeconds,
                    logData.CycleNumber
                );
                sb.AppendLine();

                File.AppendAllText(_logFilePath, sb.ToString(), Encoding.UTF8);

                // 마지막 측정 시간 업데이트
                _lastMeasurementTime = currentTime;
            }
            catch (Exception ex)
            {
                lblWarning.Text = $"로그 기록 실패: {ex.Message}";
                lblWarning.ForeColor = Color.Red;
                messageTimer.Stop();
                messageTimer.Start();
            }
        }
    }
}