using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Drawing;
using System.ComponentModel;
using System.IO;

namespace FanucFocasTutorial
{
    public partial class AutoMonitoringForm : Form
    {
        private static string GetLogFilePath()
        {
            string fileName = $"MachineState_Debug_{DateTime.Now:yyyyMMdd}.txt";
            return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName);
        }
        private static readonly object _logLock = new object();

        private CNCConnection _connection;
        private MainForm.IPConfig _config;
        private System.Windows.Forms.Timer _updateTimer;
        private TableLayoutPanel _mainLayout;

        // PMC 타입 상수
        private const short PMC_TYPE_G = 0;
        private const short PMC_TYPE_F = 1;
        private const short PMC_TYPE_R = 5;
        private const short PMC_TYPE_D = 7;
        private const short PMC_TYPE_X = 4;

        // 설비 상태 모니터링 UI 필드
        private Button _btnStateRunning;
        private Button _btnStateLoading;
        private Button _btnStateAlarm;
        private Button _btnStateIdle;
        private Label _lblStateRunningTime;
        private Label _lblStateLoadingTime;
        private Label _lblStateAlarmTime;
        private Label _lblStateIdleTime;
        private Label _lblPmcDebugInfo;

        // IP별 상태 추적
        private Dictionary<string, StateTransition> _stateTransitions;
        private Dictionary<string, Dictionary<MachineState, int>> _dailyStateDurations;
        private PmcStateValues _currentPmcValues;
        private bool _isStateMonitoringUpdating = false;

        // 근무조 관리
        private ShiftInfo _currentShift;
        private Dictionary<string, ShiftStateData> _shiftStateData;

        // M1 옵셔널 스톱 타이머
        private DateTime? _m1StartTime = null;

        // 시간 정보
        private Label _lblCurrentTime;
        private Label _lblShiftInfo;

        // 운전 정보
        private Label _lblProgram;
        private Label _lblSpindleInfo;
        private Label _lblFeedrateInfo;
        private Label _lblAlarmInfo;

        // 통합 지표
        private Label _lblActualWorkingTime;  // 실가공시간
        private Label _lblInputTime;          // 투입시간
        private Label _lblStopTime;           // 정지시간
        private Label _lblUnmeasuredTime;     // 미측정시간
        private Label _lblOperationRate;      // 가동률
        private Label _lblProductionCount;    // 생산수량

        // 상태 카운터 (메모리에서만 관리)
        private DateTime _shiftStartTime;
        private TimeSpan _autoTime = TimeSpan.Zero;
        private TimeSpan _manualTime = TimeSpan.Zero;
        private TimeSpan _alarmTime = TimeSpan.Zero;
        private TimeSpan _idleTime = TimeSpan.Zero;
        private string _lastStatus = "Idle";
        private DateTime _lastStatusTime = DateTime.Now;

        // 이전 값들을 저장하여 변경된 경우에만 업데이트
        private string _prevCurrentTime = "";
        private string _prevEquipmentStatus = "";
        private string _prevProgram = "";
        private string _prevSpindleInfo = "";
        private string _prevFeedrateInfo = "";
        private string _prevAlarmInfo = "";
        private string _prevOperatingTime = "";
        private string _prevCuttingTime = "";
        private string _prevIdleTime = "";
        private string _prevProductionCount = "";
        private string _prevGoodCount = "";
        private string _prevNGCount = "";
        private string _prevCycleTime = "";
        private Color _prevEquipmentStatusColor = Color.White;

        public AutoMonitoringForm(CNCConnection connection, MainForm.IPConfig config = null)
        {
            _connection = connection;
            _config = config ?? new MainForm.IPConfig
            {
                IpAddress = connection.IpAddress,
                Port = 8193,
                LoadingMCode = 0
            };

            // 깜빡임 방지
            this.DoubleBuffered = true;
            this.SetStyle(ControlStyles.OptimizedDoubleBuffer |
                         ControlStyles.AllPaintingInWmPaint |
                         ControlStyles.UserPaint, true);
            this.UpdateStyles();

            // IP별 상태 추적 초기화
            _stateTransitions = new Dictionary<string, StateTransition>();
            _dailyStateDurations = new Dictionary<string, Dictionary<MachineState, int>>();
            _shiftStateData = new Dictionary<string, ShiftStateData>();

            // 근무조 초기화
            _currentShift = ShiftManager.GetCurrentShift(DateTime.Now);

            // 현재 connection의 상태 초기화
            if (_connection != null)
            {
                InitializeStateForIp(_connection.IpAddress);
                InitializeShiftDataForIp(_connection.IpAddress);
            }

            InitializeComponent();
            InitializeShiftTime();
            SetupTimer();
            
            // 프로그램 시작 로그
            WriteStateLog($"[{_connection?.IpAddress ?? "N/A"}] ========== 자동모니터링 시작 ==========");
            WriteStateLog($"[{_connection?.IpAddress ?? "N/A"}] 프로그램 초기화 완료 - IP: {_connection?.IpAddress}, 근무조: {ShiftManager.GetShiftDisplayName(_currentShift)}");
        }

        // 연결 변경 메서드 추가
        public void UpdateConnection(CNCConnection newConnection)
        {
            _connection = newConnection;

            // 새 connection의 IP 상태 초기화
            if (_connection != null && !_stateTransitions.ContainsKey(_connection.IpAddress))
            {
                InitializeStateForIp(_connection.IpAddress);
            }
        }

        /// <summary>
        /// PMC 주소 문자열을 파싱하여 타입, 주소, 비트 위치로 변환
        /// </summary>
        private bool ParsePmcAddress(string address, out short pmcType, out ushort pmcAddress, out int bitPosition)
        {
            pmcType = 0;
            pmcAddress = 0;
            bitPosition = 0;

            if (string.IsNullOrWhiteSpace(address))
                return false;

            // 타입 문자 추출 (F, G, R, D, X 등)
            char typeChar = address[0];
            pmcType = typeChar switch
            {
                'F' => PMC_TYPE_F,
                'G' => PMC_TYPE_G,
                'R' => PMC_TYPE_R,
                'D' => PMC_TYPE_D,
                'X' => PMC_TYPE_X,
                _ => (short)0
            };

            if (pmcType == 0)
                return false;

            // 주소와 비트 파싱
            string remaining = address.Substring(1);
            string[] parts = remaining.Split('.');

            if (parts.Length >= 1)
            {
                if (!ushort.TryParse(parts[0], out pmcAddress))
                    return false;
            }

            if (parts.Length >= 2)
            {
                if (!int.TryParse(parts[1], out bitPosition))
                    return false;
            }

            return true;
        }

        // IP별 상태 초기화
        private void InitializeStateForIp(string ip)
        {
            // StateTransition 초기화
            _stateTransitions[ip] = new StateTransition
            {
                StartTime = DateTime.Now,
                CurrentState = MachineState.Idle,
                ElapsedSeconds = 0
            };

            // 일일 누적 시간 초기화
            _dailyStateDurations[ip] = new Dictionary<MachineState, int>
            {
                { MachineState.Running, 0 },
                { MachineState.Loading, 0 },
                { MachineState.Alarm, 0 },
                { MachineState.Idle, 0 }
            };
        }

        // IP별 근무조 데이터 초기화
        private void InitializeShiftDataForIp(string ip)
        {
            try
            {
                var logService = new LogDataService();
                var savedData = logService.LoadShiftStateData(
                    ip,
                    _currentShift.StartTime.Date,
                    _currentShift.Type);

                if (savedData != null)
                {
                    // 마지막 업데이트 시간 조회
                    var lastUpdatedTime = logService.GetLastUpdatedTime(
                        ip,
                        _currentShift.StartTime.Date,
                        _currentShift.Type);

                    // 미측정 시간 계산: 마지막 업데이트 ~ 현재 시간
                    WriteStateLog($"[{ip}] ========== 미측정 시간 계산 시작 ==========");
                    WriteStateLog($"[{ip}] 현재 근무조: {ShiftManager.GetShiftDisplayName(_currentShift)}");
                    WriteStateLog($"[{ip}] 근무조 시간: {_currentShift.StartTime:yyyy-MM-dd HH:mm:ss} ~ {_currentShift.EndTime:yyyy-MM-dd HH:mm:ss}");
                    
                    if (lastUpdatedTime.HasValue)
                    {
                        DateTime now = DateTime.Now;
                        WriteStateLog($"[{ip}] 마지막 업데이트 시간: {lastUpdatedTime.Value:yyyy-MM-dd HH:mm:ss}");
                        WriteStateLog($"[{ip}] 현재 시간: {now:yyyy-MM-dd HH:mm:ss}");

                        // 현재 시간이 같은 근무조 내에 있는지 확인
                        if (now >= _currentShift.StartTime && now < _currentShift.EndTime)
                        {
                            TimeSpan unmeasuredGap = now - lastUpdatedTime.Value;
                            WriteStateLog($"[{ip}] 시간 차이: {unmeasuredGap.TotalSeconds}초 ({FormatDuration((int)unmeasuredGap.TotalSeconds)})");

                            // 음수 방지 (시간이 잘못된 경우)
                            if (unmeasuredGap.TotalSeconds > 0)
                            {
                                int additionalUnmeasured = (int)unmeasuredGap.TotalSeconds;
                                int beforeUnmeasured = savedData.UnmeasuredSeconds;
                                savedData.UnmeasuredSeconds += additionalUnmeasured;
                                
                                WriteStateLog($"[{ip}] 미측정 시간 추가 - 이전: {FormatDuration(beforeUnmeasured)}, 추가: {FormatDuration(additionalUnmeasured)}, 합계: {FormatDuration(savedData.UnmeasuredSeconds)}");
                                
                                // 미측정 시간 계산 후 즉시 DB에 저장
                                try
                                {
                                    logService.UpdateShiftStateData(savedData);
                                    WriteStateLog($"[{ip}] ★ 미측정 시간 DB 저장 완료");
                                }
                                catch (Exception ex)
                                {
                                    WriteStateLog($"[{ip}] ★ 미측정 시간 DB 저장 실패: {ex.Message}");
                                }
                            }
                            else
                            {
                                WriteStateLog($"[{ip}] 미측정 시간 계산 스킵 (시간 차이 <= 0)");
                            }
                        }
                        else
                        {
                            WriteStateLog($"[{ip}] 미측정 시간 계산 스킵 (다른 근무조)");
                        }
                    }
                    else
                    {
                        WriteStateLog($"[{ip}] 마지막 업데이트 시간 없음 - 미측정 시간 계산 불가");
                    }
                    WriteStateLog($"[{ip}] ========== 미측정 시간 계산 완료 ==========");

                    // DB에서 복원
                    _shiftStateData[ip] = savedData;
                    WriteStateLog($"[{ip}] ========== DB 데이터 복원 완료 ==========");
                    WriteStateLog($"[{ip}] 복원된 데이터 - Running:{FormatDuration(savedData.RunningSeconds)} " +
                        $"Loading:{FormatDuration(savedData.LoadingSeconds)} " +
                        $"Alarm:{FormatDuration(savedData.AlarmSeconds)} " +
                        $"Idle:{FormatDuration(savedData.IdleSeconds)} " +
                        $"Unmeasured:{FormatDuration(savedData.UnmeasuredSeconds)} " +
                        $"생산:{savedData.ProductionCount}개");
                    _dailyStateDurations[ip][MachineState.Running] = savedData.RunningSeconds;
                    _dailyStateDurations[ip][MachineState.Loading] = savedData.LoadingSeconds;
                    _dailyStateDurations[ip][MachineState.Alarm] = savedData.AlarmSeconds;
                    _dailyStateDurations[ip][MachineState.Idle] = savedData.IdleSeconds;
                }
                else
                {
                    // DB에 데이터 없음 - 새로 시작
                    WriteStateLog($"[{ip}] ========== 새 근무조 시작 (DB 데이터 없음) ==========");
                    _shiftStateData[ip] = new ShiftStateData
                    {
                        IpAddress = ip,
                        ShiftType = _currentShift.Type,
                        ShiftDate = _currentShift.StartTime.Date,
                        IsExtended = _currentShift.IsExtended,
                        RunningSeconds = 0,
                        LoadingSeconds = 0,
                        AlarmSeconds = 0,
                        IdleSeconds = 0,
                        UnmeasuredSeconds = 0,
                        ProductionCount = 0
                    };
                }
            }
            catch (Exception)
            {
                // DB 복원 실패 시 0으로 시작
                _shiftStateData[ip] = new ShiftStateData
                {
                    IpAddress = ip,
                    ShiftType = _currentShift.Type,
                    ShiftDate = _currentShift.StartTime.Date,
                    IsExtended = _currentShift.IsExtended,
                    RunningSeconds = 0,
                    LoadingSeconds = 0,
                    AlarmSeconds = 0,
                    IdleSeconds = 0,
                    UnmeasuredSeconds = 0,
                    ProductionCount = 0
                };
            }
        }

        private void InitializeComponent()
        {
            this.Text = "자동 모니터링";
            this.Size = new Size(1000, 700);
            this.FormBorderStyle = FormBorderStyle.Sizable;

            InitializeControls();  // 설비 상태 모니터링 UI 통합
        }

        private void InitializeControls()
        {
            // 메인 레이아웃 패널 (5열 구조)
            _mainLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 5,
                RowCount = 7,  // 8행에서 7행으로 축소
                Padding = new Padding(10),
                CellBorderStyle = TableLayoutPanelCellBorderStyle.Single
            };

            // 컬럼 비율 설정 (5열)
            for (int i = 0; i < 5; i++)
            {
                _mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20F));
            }

            // 행 비율 설정
            for (int i = 0; i < 7; i++)
            {
                _mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 14.29F));
            }

            // Row 0: 현재 시간 + 근무조 정보 (3열 + 2열)
            _lblCurrentTime = CreateStyledLabel($"현재 시간: {DateTime.Now:HH:mm:ss}", Color.LightBlue);
            _mainLayout.Controls.Add(_lblCurrentTime, 0, 0);
            _mainLayout.SetColumnSpan(_lblCurrentTime, 3);

            _lblShiftInfo = CreateStyledLabel(ShiftManager.GetShiftDisplayName(_currentShift), Color.LightGreen);
            _mainLayout.Controls.Add(_lblShiftInfo, 3, 0);
            _mainLayout.SetColumnSpan(_lblShiftInfo, 2);

            // 설비 상태 모니터링 LED 버튼
            _btnStateRunning = CreateStateLedButton("● 가공중", Color.LightGray);
            _btnStateLoading = CreateStateLedButton("● 제품교체", Color.LightGray);
            _btnStateAlarm = CreateStateLedButton("● 장애정지", Color.LightGray);
            _btnStateIdle = CreateStateLedButton("● 유휴", Color.LightGray);

            // 상태별 누적 시간 레이블
            _lblStateRunningTime = CreateStateTimeLabel("누적: 00:00:00");
            _lblStateLoadingTime = CreateStateTimeLabel("누적: 00:00:00");
            _lblStateAlarmTime = CreateStateTimeLabel("누적: 00:00:00");
            _lblStateIdleTime = CreateStateTimeLabel("누적: 00:00:00");

            // PMC 디버그 정보
            _lblPmcDebugInfo = new Label
            {
                Text = "PMC: F0.0=? F7.0=? F10.0=? G4.3=? X8.4=?",
                Font = new Font("Consolas", 9f),
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(240, 240, 240),
                BorderStyle = BorderStyle.FixedSingle
            };

            // Row 1: 4개 상태 LED 버튼
            _mainLayout.Controls.Add(_btnStateRunning, 0, 1);
            _mainLayout.Controls.Add(_btnStateLoading, 1, 1);
            _mainLayout.Controls.Add(_btnStateAlarm, 2, 1);
            _mainLayout.Controls.Add(_btnStateIdle, 3, 1);

            // Row 2: 4개 상태 시간 레이블
            _mainLayout.Controls.Add(_lblStateRunningTime, 0, 2);
            _mainLayout.Controls.Add(_lblStateLoadingTime, 1, 2);
            _mainLayout.Controls.Add(_lblStateAlarmTime, 2, 2);
            _mainLayout.Controls.Add(_lblStateIdleTime, 3, 2);

            // Row 3: PMC 디버그 정보 (5열 합치기)
            _mainLayout.Controls.Add(_lblPmcDebugInfo, 0, 3);
            _mainLayout.SetColumnSpan(_lblPmcDebugInfo, 5);

            // 운전 정보
            _lblProgram = CreateStyledLabel("프로그램: -", Color.LightYellow);
            _lblSpindleInfo = CreateStyledLabel("스핀들: 0 RPM (0%)", Color.LightYellow);
            _lblFeedrateInfo = CreateStyledLabel("이송속도: 0 mm/min", Color.LightYellow);
            _lblAlarmInfo = CreateStyledLabel("알람 없음", Color.LightCoral);

            // Row 4: 운전 정보
            _mainLayout.Controls.Add(_lblProgram, 0, 4);
            _mainLayout.SetColumnSpan(_lblProgram, 3);
            _mainLayout.Controls.Add(_lblSpindleInfo, 3, 4);
            _mainLayout.SetColumnSpan(_lblSpindleInfo, 2);

            // Row 5: 이송속도, 알람
            _mainLayout.Controls.Add(_lblFeedrateInfo, 0, 5);
            _mainLayout.SetColumnSpan(_lblFeedrateInfo, 3);
            _mainLayout.Controls.Add(_lblAlarmInfo, 3, 5);
            _mainLayout.SetColumnSpan(_lblAlarmInfo, 2);

            // 통합 지표 레이블
            _lblActualWorkingTime = CreateStyledLabel("실가공: 00:00:00", Color.FromArgb(144, 238, 144));
            _lblInputTime = CreateStyledLabel("투입: 00:00:00", Color.FromArgb(173, 216, 230));
            _lblStopTime = CreateStyledLabel("정지: 00:00:00", Color.FromArgb(255, 182, 193));
            _lblUnmeasuredTime = CreateStyledLabel("미측정: 00:00:00", Color.FromArgb(220, 220, 220));
            _lblOperationRate = CreateStyledLabel("가동률: 0%", Color.FromArgb(255, 255, 153));
            _lblProductionCount = CreateStyledLabel("생산: 0개", Color.LightCyan);

            // Row 6: 통합 지표
            _mainLayout.Controls.Add(_lblActualWorkingTime, 0, 6);
            _mainLayout.Controls.Add(_lblInputTime, 1, 6);
            _mainLayout.Controls.Add(_lblStopTime, 2, 6);
            _mainLayout.Controls.Add(_lblUnmeasuredTime, 3, 6);

            // 가동률과 생산수량을 하나의 셀에 표시하기 위해 패널 생성
            TableLayoutPanel metricsPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 2,
                ColumnCount = 1,
                Padding = new Padding(0)
            };
            metricsPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
            metricsPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
            metricsPanel.Controls.Add(_lblOperationRate, 0, 0);
            metricsPanel.Controls.Add(_lblProductionCount, 0, 1);
            _mainLayout.Controls.Add(metricsPanel, 4, 6);

            this.Controls.Add(_mainLayout);
        }

        private Label CreateStyledLabel(string text, Color backColor, bool isHeader = false)
        {
            return new Label
            {
                Text = text,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("맑은 고딕", isHeader ? 12f : 11f, isHeader ? FontStyle.Bold : FontStyle.Regular),
                BackColor = backColor,
                BorderStyle = BorderStyle.FixedSingle,
                AutoSize = false
            };
        }

        private void InitializeShiftTime()
        {
            DateTime now = DateTime.Now;
            TimeSpan currentTime = now.TimeOfDay;
            TimeSpan dayStart = new TimeSpan(8, 30, 0);
            TimeSpan dayEnd = new TimeSpan(20, 30, 0);

            if (currentTime >= dayStart && currentTime < dayEnd)
            {
                _shiftStartTime = new DateTime(now.Year, now.Month, now.Day, 8, 30, 0);
            }
            else
            {
                if (currentTime >= new TimeSpan(0, 0, 0) && currentTime < dayStart)
                {
                    // 야간 교대조 (전날 20:30부터)
                    _shiftStartTime = new DateTime(now.Year, now.Month, now.Day - 1, 20, 30, 0);
                }
                else
                {
                    // 야간 교대조 (당일 20:30부터)
                    _shiftStartTime = new DateTime(now.Year, now.Month, now.Day, 20, 30, 0);
                }
            }
        }

        private void SetupTimer()
        {
            _updateTimer = new System.Windows.Forms.Timer
            {
                Interval = 1000 // 1초마다 업데이트
            };
            _updateTimer.Tick += UpdateTimer_Tick;
            _updateTimer.Start();
            WriteStateLog($"[{_connection?.IpAddress ?? "N/A"}] 타이머 시작 - 1초 간격 업데이트");
        }

        private void UpdateTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                UpdateDisplayData();
                UpdateStatusTime();
                UpdateMachineStateMonitoring();  // 설비 상태 모니터링 업데이트
            }
            catch (Exception ex)
            {
                WriteStateLog($"[{_connection?.IpAddress ?? "N/A"}] 타이머 틱 오류: {ex.Message}");
            }
        }

        private void UpdateDisplayData()
        {
            if (_connection == null || !_connection.IsConnected)
            {
                // 연결 안됨 - PMC 기반 상태 모니터링에서 처리
                return;
            }

            try
            {
                // 현재 시간 업데이트 (변경된 경우에만)
                string currentTimeText = $"현재 시간: {DateTime.Now:HH:mm:ss}";
                UpdateLabelIfChanged(_lblCurrentTime, currentTimeText);

                // 설비 상태는 이제 PMC 기반 상태 모니터링 (UpdateMachineStateMonitoring)에서 처리됨

                // 프로그램 정보
                string programText = $"프로그램: {_connection.GetCurrentProgram()}";
                UpdateLabelIfChanged(_lblProgram, programText);

                // 스핀들 정보
                int spindleSpeed = _connection.GetSpindleSpeed();
                double spindleLoad = _connection.GetSpindleLoad();
                string spindleText = $"스핀들: {spindleSpeed} RPM ({spindleLoad:F1}%)";
                UpdateLabelIfChanged(_lblSpindleInfo, spindleText);

                // 이송속도 정보
                int feedrate = _connection.GetFeedrate();
                string feedrateText = $"이송속도: {feedrate} mm/min";
                UpdateLabelIfChanged(_lblFeedrateInfo, feedrateText);

                // 알람 정보
                string alarmText = _connection.GetCurrentAlarm();
                UpdateLabelIfChanged(_lblAlarmInfo, alarmText);

                // 알람 정보 클릭 이벤트 (한 번만 등록)
                if (_lblAlarmInfo.Tag == null)
                {
                    _lblAlarmInfo.Cursor = Cursors.Hand;
                    _lblAlarmInfo.Click += (s, args) => {
                        var alarmForm = new AlarmDetailsForm(_connection);
                        alarmForm.Show();
                    };
                    _lblAlarmInfo.Tag = "EventAdded";
                }

                // M99 감지 및 생산수량 업데이트 (근무조별 관리)
                _connection.CheckM99AndUpdateProduction();

                // 생산수량은 근무조 데이터에서 관리
                if (_connection != null && _shiftStateData.ContainsKey(_connection.IpAddress))
                {
                    int productionCount = _connection.GetProductionCount();
                    _shiftStateData[_connection.IpAddress].ProductionCount = productionCount;
                }
            }
            catch (Exception)
            {
                // 오류 발생 시 처리 (설비 상태는 PMC 기반 모니터링에서 처리)
            }
        }

        private void UpdateStatusTime()
        {
            if (_connection == null || !_connection.IsConnected) return;

            string currentStatus = _connection.GetEquipmentStatus();
            DateTime now = DateTime.Now;
            TimeSpan elapsed = now - _lastStatusTime;

            // 이전 상태 시간 누적
            switch (_lastStatus)
            {
                case "Auto":
                    _autoTime = _autoTime.Add(elapsed);
                    break;
                case "Manual":
                    _manualTime = _manualTime.Add(elapsed);
                    break;
                case "Alarm":
                    _alarmTime = _alarmTime.Add(elapsed);
                    break;
                case "Idle":
                    _idleTime = _idleTime.Add(elapsed);
                    break;
            }

            _lastStatus = currentStatus;
            _lastStatusTime = now;
        }


        // 값이 변경된 경우에만 Label을 업데이트하여 깜빡임 방지
        private void UpdateLabelIfChanged(Label label, string newText, Color? newBackColor = null)
        {
            if (label.Text != newText)
            {
                label.Text = newText;
            }

            if (newBackColor.HasValue && label.BackColor != newBackColor.Value)
            {
                label.BackColor = newBackColor.Value;
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            WriteStateLog($"[{_connection?.IpAddress ?? "N/A"}] ========== 프로그램 종료 시작 ==========");
            
            // 현재 근무조 데이터 저장
            if (_connection != null && _connection.IsConnected)
            {
                WriteStateLog($"[{_connection.IpAddress}] 종료 시 데이터 저장 시작");
                SaveCurrentShiftData(_connection.IpAddress);
            }
            else
            {
                WriteStateLog($"[{_connection?.IpAddress ?? "N/A"}] 연결 안됨 - 데이터 저장 스킵");
            }

            _updateTimer?.Stop();
            _updateTimer?.Dispose();
            WriteStateLog($"[{_connection?.IpAddress ?? "N/A"}] 타이머 정지 및 해제 완료");
            WriteStateLog($"[{_connection?.IpAddress ?? "N/A"}] ========== 프로그램 종료 완료 ==========");
            base.OnFormClosing(e);
        }

        // ==================== 설비 상태 모니터링 ====================
        // (InitializeControls()에 통합됨)

        private Button CreateStateLedButton(string text, Color backColor)
        {
            return new Button
            {
                Text = text,
                Font = new Font("맑은 고딕", 11f, FontStyle.Bold),
                BackColor = backColor,
                ForeColor = Color.DarkGray,
                Dock = DockStyle.Fill,
                FlatStyle = FlatStyle.Flat,
                TextAlign = ContentAlignment.MiddleCenter,
                Margin = new Padding(3)
            };
        }

        private Label CreateStateTimeLabel(string text)
        {
            return new Label
            {
                Text = text,
                Font = new Font("Consolas", 10f),
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Margin = new Padding(3)
            };
        }

        private async void UpdateMachineStateMonitoring()
        {
            if (_connection == null || !_connection.IsConnected)
            {
                WriteStateLog($"[{_connection?.IpAddress ?? "N/A"}] 연결 안됨 - 상태 모니터링 스킵");
                return;
            }

            if (_isStateMonitoringUpdating)
            {
                WriteStateLog($"[{_connection.IpAddress}] 이미 업데이트 중 - 스킵");
                return;
            }

            _isStateMonitoringUpdating = true;

            try
            {
                string ip = _connection.IpAddress;

                // IP별 상태가 초기화되지 않았으면 초기화
                if (!_stateTransitions.ContainsKey(ip))
                {
                    WriteStateLog($"[{ip}] 상태 추적 초기화");
                    InitializeStateForIp(ip);
                }

                // 근무조 전환 체크 및 처리
                CheckAndHandleShiftTransition(ip);

                var pmcResult = await System.Threading.Tasks.Task.Run(() =>
                {
                    if (ReadPmcStateValues(out PmcStateValues pmc))
                        return (success: true, values: pmc);
                    WriteStateLog($"[{ip}] PMC 읽기 실패");
                    return (success: false, values: default(PmcStateValues));
                });

                if (!pmcResult.success)
                {
                    WriteStateLog($"[{ip}] PMC 읽기 실패로 상태 모니터링 중단");
                    return;
                }

                _currentPmcValues = pmcResult.values;
                MachineState currentState = DetermineMachineState(pmcResult.values);
                
                WriteStateLog($"[{ip}] PMC 읽기 성공 - F0.7={B(pmcResult.values.F0_7)} F1.0={B(pmcResult.values.F1_0)} F3.5={B(pmcResult.values.F3_5)} F10={pmcResult.values.F10_value} G4.3={B(pmcResult.values.G4_3)} G5.0={B(pmcResult.values.G5_0)} → 상태: {currentState}");

                var transition = _stateTransitions[ip];
                var dailyDurations = _dailyStateDurations[ip];

                if (transition.CurrentState != currentState)
                {
                    int duration = (int)(DateTime.Now - transition.StartTime).TotalSeconds;
                    dailyDurations[transition.CurrentState] += duration;

                    // 상태 변경 로그 작성
                    WriteStateLog($"[{ip}] ★★★ 상태변경 ★★★ {transition.CurrentState} → {currentState} " +
                        $"(지속시간: {FormatDuration(duration)}) | PMC: F0.7={B(pmcResult.values.F0_7)} F1.0={B(pmcResult.values.F1_0)} " +
                        $"F3.5={B(pmcResult.values.F3_5)} F10={pmcResult.values.F10_value} G4.3={B(pmcResult.values.G4_3)} G5.0={B(pmcResult.values.G5_0)}");
                    
                    WriteStateLog($"[{ip}] 이전 상태 누적 시간 - Running:{FormatDuration(dailyDurations[MachineState.Running])} " +
                        $"Loading:{FormatDuration(dailyDurations[MachineState.Loading])} " +
                        $"Alarm:{FormatDuration(dailyDurations[MachineState.Alarm])} " +
                        $"Idle:{FormatDuration(dailyDurations[MachineState.Idle])}");

                    // 투입 → 실가공 전환 시 생산 수량 증가
                    if (transition.CurrentState == MachineState.Loading && currentState == MachineState.Running)
                    {
                        _connection.IncrementProduction();
                        WriteStateLog($"[{ip}] 생산수량 증가: {_connection.GetProductionCount()}개");
                    }

                    // 근무조 데이터 업데이트 (메모리)
                    if (_shiftStateData.ContainsKey(ip))
                    {
                        var shiftData = _shiftStateData[ip];
                        shiftData.RunningSeconds = dailyDurations[MachineState.Running];
                        shiftData.LoadingSeconds = dailyDurations[MachineState.Loading];
                        shiftData.AlarmSeconds = dailyDurations[MachineState.Alarm];
                        shiftData.IdleSeconds = dailyDurations[MachineState.Idle];
                        shiftData.ProductionCount = _connection.GetProductionCount();
                        
                        WriteStateLog($"[{ip}] 메모리 데이터 업데이트 - Running:{FormatDuration(shiftData.RunningSeconds)} " +
                            $"Loading:{FormatDuration(shiftData.LoadingSeconds)} " +
                            $"Alarm:{FormatDuration(shiftData.AlarmSeconds)} " +
                            $"Idle:{FormatDuration(shiftData.IdleSeconds)} " +
                            $"생산:{shiftData.ProductionCount}개");
                        
                        // 상태 변경 시 즉시 DB 업데이트
                        try
                        {
                            var logService = new LogDataService();
                            logService.UpdateShiftStateData(shiftData);
                            WriteStateLog($"[{ip}] ★ DB 업데이트 완료 (상태 변경 시)");
                        }
                        catch (Exception ex)
                        {
                            WriteStateLog($"[{ip}] ★ DB 업데이트 실패: {ex.Message}");
                        }
                    }

                    transition.StartTime = DateTime.Now;
                    transition.CurrentState = currentState;
                    transition.ElapsedSeconds = 0;
                    
                    WriteStateLog($"[{ip}] 새 상태 시작: {currentState} (시작시간: {transition.StartTime:HH:mm:ss})");
                }
                else
                {
                    transition.ElapsedSeconds = (int)(DateTime.Now - transition.StartTime).TotalSeconds;
                }

                UpdateStateMonitoringUI(currentState, transition, dailyDurations);
                UpdatePmcDebugUI(pmcResult.values);
            }
            catch (Exception ex)
            {
                WriteStateLog($"[{_connection?.IpAddress ?? "N/A"}] 상태 모니터링 오류: {ex.Message}");
                WriteStateLog($"[{_connection?.IpAddress ?? "N/A"}] 스택 트레이스: {ex.StackTrace}");
            }
            finally
            {
                _isStateMonitoringUpdating = false;
            }
        }

        /// <summary>
        /// 근무조 전환 감지 및 처리
        /// </summary>
        private void CheckAndHandleShiftTransition(string ip)
        {
            ShiftInfo newShift = ShiftManager.GetCurrentShift(DateTime.Now);

            // 근무조 변경 감지
            if (ShiftManager.IsShiftChanged(_currentShift, newShift))
            {
                WriteStateLog($"[{ip}] ========== 근무조 전환 감지 ==========");
                WriteStateLog($"[{ip}] 이전 근무조: {ShiftManager.GetShiftDisplayName(_currentShift)}");
                WriteStateLog($"[{ip}] 새 근무조: {ShiftManager.GetShiftDisplayName(newShift)}");
                
                // 현재 근무조 데이터 저장
                WriteStateLog($"[{ip}] 이전 근무조 데이터 저장 시작");
                SaveCurrentShiftData(ip);

                // 새 근무조로 전환
                _currentShift = newShift;
                WriteStateLog($"[{ip}] 새 근무조로 전환 완료");

                // 상태 추적 초기화
                _dailyStateDurations[ip] = new Dictionary<MachineState, int>
                {
                    { MachineState.Running, 0 },
                    { MachineState.Loading, 0 },
                    { MachineState.Alarm, 0 },
                    { MachineState.Idle, 0 }
                };

                // 근무조 데이터 초기화
                _shiftStateData[ip] = new ShiftStateData
                {
                    IpAddress = ip,
                    ShiftType = newShift.Type,
                    ShiftDate = newShift.StartTime.Date,
                    IsExtended = newShift.IsExtended,
                    RunningSeconds = 0,
                    LoadingSeconds = 0,
                    AlarmSeconds = 0,
                    IdleSeconds = 0,
                    UnmeasuredSeconds = 0,
                    ProductionCount = 0
                };

                // 생산수량 초기화
                if (_connection != null)
                {
                    _connection.ResetProductionCount();
                }

                // UI 업데이트
                UpdateLabelIfChanged(_lblShiftInfo, ShiftManager.GetShiftDisplayName(newShift));
            }
        }

        /// <summary>
        /// 현재 근무조 데이터 DB 저장
        /// </summary>
        private void SaveCurrentShiftData(string ip)
        {
            if (!_shiftStateData.ContainsKey(ip))
                return;

            try
            {
                var shiftData = _shiftStateData[ip];
                var dailyDurations = _dailyStateDurations[ip];

                // 현재 진행 중인 상태의 시간을 누적에 반영
                if (_stateTransitions.ContainsKey(ip))
                {
                    var transition = _stateTransitions[ip];
                    int elapsedSeconds = (int)(DateTime.Now - transition.StartTime).TotalSeconds;

                    // 현재 상태에 진행 중인 시간 추가
                    dailyDurations[transition.CurrentState] += elapsedSeconds;

                    WriteStateLog($"[{ip}] 앱 종료 시 데이터 저장: {transition.CurrentState} 상태 {elapsedSeconds}초 추가 반영");
                }

                // 누적 시간을 shiftData에 반영
                shiftData.RunningSeconds = dailyDurations[MachineState.Running];
                shiftData.LoadingSeconds = dailyDurations[MachineState.Loading];
                shiftData.AlarmSeconds = dailyDurations[MachineState.Alarm];
                shiftData.IdleSeconds = dailyDurations[MachineState.Idle];

                WriteStateLog($"[{ip}] ========== DB 저장 시작 ==========");
                WriteStateLog($"[{ip}] 저장되는 데이터:");
                WriteStateLog($"[{ip}]   - Running: {FormatDuration(shiftData.RunningSeconds)} ({shiftData.RunningSeconds}초)");
                WriteStateLog($"[{ip}]   - Loading: {FormatDuration(shiftData.LoadingSeconds)} ({shiftData.LoadingSeconds}초)");
                WriteStateLog($"[{ip}]   - Alarm: {FormatDuration(shiftData.AlarmSeconds)} ({shiftData.AlarmSeconds}초)");
                WriteStateLog($"[{ip}]   - Idle: {FormatDuration(shiftData.IdleSeconds)} ({shiftData.IdleSeconds}초)");
                WriteStateLog($"[{ip}]   - Unmeasured: {FormatDuration(shiftData.UnmeasuredSeconds)} ({shiftData.UnmeasuredSeconds}초)");
                WriteStateLog($"[{ip}]   - 생산수량: {shiftData.ProductionCount}개");
                WriteStateLog($"[{ip}]   - 근무조: {shiftData.ShiftType}, 날짜: {shiftData.ShiftDate:yyyy-MM-dd}");

                // DB 저장
                var logService = new LogDataService();
                try
                {
                    logService.SaveShiftStateData(shiftData);
                    WriteStateLog($"[{ip}] ★ DB 저장 성공 - 마지막 업데이트 시간: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                }
                catch (Exception ex)
                {
                    WriteStateLog($"[{ip}] ★ DB 저장 실패: {ex.Message}");
                    WriteStateLog($"[{ip}] 스택 트레이스: {ex.StackTrace}");
                }
                WriteStateLog($"[{ip}] ========== DB 저장 완료 ==========");
            }
            catch (Exception ex)
            {
                WriteStateLog($"[{ip}] 저장 실패: {ex.Message}");
            }
        }

        private bool ReadPmcStateValues(out PmcStateValues values)
        {
            values = new PmcStateValues();

            if (_connection == null || !_connection.IsConnected)
                return false;

            try
            {
                // F0.0 (ST_OP 가동)
                if (ParsePmcAddress(_config.PmcF0_0, out short type_f0_0, out ushort addr_f0_0, out int bit_f0_0))
                {
                    if (!ReadPmcBit(type_f0_0, (short)addr_f0_0, (short)bit_f0_0, out values.F0_0)) return false;
                }

                // F0.7 (스타트 실행)
                if (ParsePmcAddress(_config.PmcF0_7, out short type_f0_7, out ushort addr_f0_7, out int bit_f0_7))
                {
                    if (!ReadPmcBit(type_f0_7, (short)addr_f0_7, (short)bit_f0_7, out values.F0_7)) return false;
                }

                // F1.0 (알람 신호)
                if (ParsePmcAddress(_config.PmcF1_0, out short type_f1_0, out ushort addr_f1_0, out int bit_f1_0))
                {
                    if (!ReadPmcBit(type_f1_0, (short)addr_f1_0, (short)bit_f1_0, out values.F1_0)) return false;
                }

                // F3.5 (메모리 모드)
                if (ParsePmcAddress(_config.PmcF3_5, out short type_f3_5, out ushort addr_f3_5, out int bit_f3_5))
                {
                    if (!ReadPmcBit(type_f3_5, (short)addr_f3_5, (short)bit_f3_5, out values.F3_5)) return false;
                }

                // F10 (M코드 번호 - Word)
                if (ParsePmcAddress(_config.PmcF10, out short type_f10, out ushort addr_f10, out int _))
                {
                    if (!ReadPmcWord(type_f10, (short)addr_f10, out values.F10_value)) return false;
                }

                // G4.3 (M핀 처리)
                if (ParsePmcAddress(_config.PmcG4_3, out short type_g4_3, out ushort addr_g4_3, out int bit_g4_3))
                {
                    if (!ReadPmcBit(type_g4_3, (short)addr_g4_3, (short)bit_g4_3, out values.G4_3)) return false;
                }

                // G5.0 (투입 조건)
                if (ParsePmcAddress(_config.PmcG5_0, out short type_g5_0, out ushort addr_g5_0, out int bit_g5_0))
                {
                    if (!ReadPmcBit(type_g5_0, (short)addr_g5_0, (short)bit_g5_0, out values.G5_0)) return false;
                }

                // X8.4 (비상정지)
                if (ParsePmcAddress(_config.PmcX8_4, out short type_x8_4, out ushort addr_x8_4, out int bit_x8_4))
                {
                    if (!ReadPmcBit(type_x8_4, (short)addr_x8_4, (short)bit_x8_4, out values.X8_4)) return false;
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        private bool ReadPmcBit(short pmcType, short address, short bitPosition, out bool value)
        {
            value = false;

            try
            {
                Focas1.IODBPMC0 pmc = new Focas1.IODBPMC0();
                short ret = Focas1.pmc_rdpmcrng(_connection.Handle, pmcType, 0, (ushort)address, (ushort)address, 9, pmc);

                if (ret == Focas1.EW_OK)
                {
                    byte byteValue = pmc.cdata[0];
                    value = ((byteValue >> bitPosition) & 1) == 1;
                    return true;
                }

                return false;
            }
            catch
            {
                return false;
            }
        }

        private bool ReadPmcWord(short pmcType, short address, out short value)
        {
            value = 0;

            try
            {
                Focas1.IODBPMC0 pmc = new Focas1.IODBPMC0();
                short ret = Focas1.pmc_rdpmcrng(_connection.Handle, pmcType, 0, (ushort)address, (ushort)(address + 1), 10, pmc);

                if (ret == Focas1.EW_OK)
                {
                    value = (short)(pmc.cdata[0] | (pmc.cdata[1] << 8));
                    return true;
                }

                return false;
            }
            catch
            {
                return false;
            }
        }

        private MachineState DetermineMachineState(PmcStateValues pmc)
        {
            // 알람 체크: F1.0 (PMC 신호) OR CNC 알람
            if (pmc.F1_0 || _connection.HasAlarm())
            {
                _m1StartTime = null;  // M1 타이머 리셋
                return MachineState.Alarm;
            }

            // M1 옵셔널 스톱 체크 (3초 이상 지속 시 유휴)
            if (pmc.F10_value == 1)
            {
                if (_m1StartTime == null)
                    _m1StartTime = DateTime.Now;
                else if ((DateTime.Now - _m1StartTime.Value).TotalSeconds >= 3)
                    return MachineState.Idle;  // 3초 이상이면 유휴
            }
            else
            {
                _m1StartTime = null;  // M1 아니면 리셋
            }

            // 투입중: M코드가 설정된 경우만 체크 (0이 아닌 경우)
            if (_config.LoadingMCode > 0 && pmc.F0_7 && pmc.F3_5 && pmc.F10_value == _config.LoadingMCode && (pmc.G4_3 || pmc.G5_0))
                return MachineState.Loading;

            // 가공중: 자동 운전 신호 AND 메모리 모드
            if (pmc.F0_7 && pmc.F3_5)
                return MachineState.Running;

            return MachineState.Idle;
        }

        private void UpdateStateMonitoringUI(MachineState currentState, StateTransition transition, Dictionary<MachineState, int> dailyDurations)
        {
            ResetAllStateLeds();

            // 현재 상태 LED 활성화 및 누적 시간 표시
            switch (currentState)
            {
                case MachineState.Running:
                    _btnStateRunning.BackColor = Color.FromArgb(76, 175, 80);
                    _btnStateRunning.ForeColor = Color.White;
                    _lblStateRunningTime.Text = $"누적: {FormatDuration(dailyDurations[MachineState.Running] + transition.ElapsedSeconds)}";
                    break;

                case MachineState.Loading:
                    _btnStateLoading.BackColor = Color.FromArgb(255, 235, 59);
                    _btnStateLoading.ForeColor = Color.Black;
                    _lblStateLoadingTime.Text = $"누적: {FormatDuration(dailyDurations[MachineState.Loading] + transition.ElapsedSeconds)}";
                    break;

                case MachineState.Alarm:
                    _btnStateAlarm.BackColor = Color.FromArgb(244, 67, 54);
                    _btnStateAlarm.ForeColor = Color.White;
                    _lblStateAlarmTime.Text = $"누적: {FormatDuration(dailyDurations[MachineState.Alarm] + transition.ElapsedSeconds)}";
                    break;

                case MachineState.Idle:
                    _btnStateIdle.BackColor = Color.FromArgb(158, 158, 158);
                    _btnStateIdle.ForeColor = Color.White;
                    _lblStateIdleTime.Text = $"누적: {FormatDuration(dailyDurations[MachineState.Idle] + transition.ElapsedSeconds)}";
                    break;
            }

            // 비활성 상태는 누적 시간만 표시
            UpdateInactiveStateTimes(currentState, dailyDurations);

            // 통합 지표 계산 및 표시
            UpdateIntegratedMetrics(dailyDurations, transition, currentState);
        }

        private void UpdateInactiveStateTimes(MachineState currentState, Dictionary<MachineState, int> dailyDurations)
        {
            // 비활성 상태는 누적 시간만 표시
            if (currentState != MachineState.Running)
                _lblStateRunningTime.Text = $"누적: {FormatDuration(dailyDurations[MachineState.Running])}";
            if (currentState != MachineState.Loading)
                _lblStateLoadingTime.Text = $"누적: {FormatDuration(dailyDurations[MachineState.Loading])}";
            if (currentState != MachineState.Alarm)
                _lblStateAlarmTime.Text = $"누적: {FormatDuration(dailyDurations[MachineState.Alarm])}";
            if (currentState != MachineState.Idle)
                _lblStateIdleTime.Text = $"누적: {FormatDuration(dailyDurations[MachineState.Idle])}";
        }

        // 통합 지표 계산 및 표시
        private void UpdateIntegratedMetrics(Dictionary<MachineState, int> dailyDurations, StateTransition transition, MachineState currentState)
        {
            // 현재 진행 중인 상태의 시간 포함
            int runningTotal = dailyDurations[MachineState.Running] + (currentState == MachineState.Running ? transition.ElapsedSeconds : 0);
            int loadingTotal = dailyDurations[MachineState.Loading] + (currentState == MachineState.Loading ? transition.ElapsedSeconds : 0);
            int alarmTotal = dailyDurations[MachineState.Alarm] + (currentState == MachineState.Alarm ? transition.ElapsedSeconds : 0);
            int idleTotal = dailyDurations[MachineState.Idle] + (currentState == MachineState.Idle ? transition.ElapsedSeconds : 0);

            // 실가공시간 = Running
            int actualWorking = runningTotal;

            // 투입시간 = Loading (제품교체 시간만)
            int input = loadingTotal;

            // 정지시간 = Idle + Alarm
            int stop = idleTotal + alarmTotal;

            // 미측정시간
            int unmeasured = 0;
            if (_connection != null && _shiftStateData.ContainsKey(_connection.IpAddress))
            {
                unmeasured = _shiftStateData[_connection.IpAddress].UnmeasuredSeconds;
            }

            // 가동률 = 실가공시간 / 전체시간(미측정 포함) * 100
            int total = runningTotal + loadingTotal + idleTotal + alarmTotal + unmeasured;
            double operationRate = total > 0 ? (double)actualWorking / total * 100 : 0;

            // UI 업데이트
            _lblActualWorkingTime.Text = $"실가공: {FormatDuration(actualWorking)}";
            _lblInputTime.Text = $"투입: {FormatDuration(input)}";
            _lblStopTime.Text = $"정지: {FormatDuration(stop)}";
            _lblUnmeasuredTime.Text = $"미측정: {FormatDuration(unmeasured)}";
            _lblOperationRate.Text = $"가동률: {operationRate:F1}%";

            // 생산수량 (근무조 데이터에서 가져오기)
            if (_connection != null && _shiftStateData.ContainsKey(_connection.IpAddress))
            {
                int productionCount = _shiftStateData[_connection.IpAddress].ProductionCount;
                _lblProductionCount.Text = $"생산: {productionCount}개";
            }
        }

        private void ResetAllStateLeds()
        {
            _btnStateRunning.BackColor = Color.LightGray;
            _btnStateRunning.ForeColor = Color.DarkGray;

            _btnStateLoading.BackColor = Color.LightGray;
            _btnStateLoading.ForeColor = Color.DarkGray;

            _btnStateAlarm.BackColor = Color.LightGray;
            _btnStateAlarm.ForeColor = Color.DarkGray;

            _btnStateIdle.BackColor = Color.LightGray;
            _btnStateIdle.ForeColor = Color.DarkGray;
        }

        private void UpdatePmcDebugUI(PmcStateValues pmc)
        {
            _lblPmcDebugInfo.Text =
                $"PMC: F0.7={B(pmc.F0_7)} F1.0={B(pmc.F1_0)} F3.5={B(pmc.F3_5)} G4.3={B(pmc.G4_3)}";
        }

        private string B(bool value) => value ? "1" : "0";

        private string FormatDuration(int seconds)
        {
            int hours = seconds / 3600;
            int minutes = (seconds % 3600) / 60;
            int secs = seconds % 60;
            return $"{hours:D2}:{minutes:D2}:{secs:D2}";
        }

        private void WriteStateLog(string message)
        {
            try
            {
                lock (_logLock)
                {
                    string logFilePath = GetLogFilePath();
                    string logMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {message}";
                    
                    // 파일이 없으면 헤더 작성
                    if (!File.Exists(logFilePath))
                    {
                        string header = $"========================================\n";
                        header += $"자동모니터링 로그 파일\n";
                        header += $"생성일시: {DateTime.Now:yyyy-MM-dd HH:mm:ss}\n";
                        header += $"========================================\n\n";
                        File.WriteAllText(logFilePath, header);
                    }
                    
                    File.AppendAllText(logFilePath, logMessage + Environment.NewLine);
                }
            }
            catch
            {
                // 로그 쓰기 실패 시 무시
            }
        }
    }
}