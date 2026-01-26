using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Drawing;
using System.Linq;
using System.IO;
using System.Diagnostics;

namespace FanucFocasTutorial
{
    /// <summary>
    /// 여러 설비를 동시에 모니터링하는 폼
    /// </summary>
    public partial class MultiMonitoringForm : Form
    {
        private List<CNCConnection> _connections;
        private FlowLayoutPanel _mainLayout;
        private System.Windows.Forms.Timer _updateTimer;
        private Dictionary<string, EquipmentMonitorPanel> _monitorPanels;
        private Dictionary<string, MainForm.IPConfig> _ipConfigs;

        public MultiMonitoringForm(List<CNCConnection> connections, Dictionary<string, MainForm.IPConfig> ipConfigs = null)
        {
            _connections = connections;
            _ipConfigs = ipConfigs ?? new Dictionary<string, MainForm.IPConfig>();
            _monitorPanels = new Dictionary<string, EquipmentMonitorPanel>();

            // 깜빡임 방지
            this.DoubleBuffered = true;
            this.SetStyle(ControlStyles.OptimizedDoubleBuffer |
                         ControlStyles.AllPaintingInWmPaint |
                         ControlStyles.UserPaint, true);
            this.UpdateStyles();

            InitializeComponent();
            CreateMonitorPanels();
            SetupTimer();

            // 폼 닫힘 이벤트 등록 (근무조 데이터 저장)
            this.FormClosing += MultiMonitoringForm_FormClosing;
        }

        private void InitializeComponent()
        {
            this.Text = "전체 설비 모니터링";
            this.Size = new Size(1400, 900);
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.StartPosition = FormStartPosition.CenterScreen;

            // 메인 레이아웃 (자동 줄바꿈, 여백 최소화)
            _mainLayout = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                Padding = new Padding(5),  // 외부 여백 감소
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = true
            };

            this.Controls.Add(_mainLayout);
        }

        private void CreateMonitorPanels()
        {
            foreach (var connection in _connections)
            {
                if (connection == null) continue;

                // IPConfig 가져오기
                MainForm.IPConfig config = _ipConfigs.ContainsKey(connection.IpAddress)
                    ? _ipConfigs[connection.IpAddress]
                    : null;

                var panel = new EquipmentMonitorPanel(connection, config);
                _monitorPanels[connection.IpAddress] = panel;
                _mainLayout.Controls.Add(panel);
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
            
            // 타이머 시작 로그
            WriteTimerLog("타이머 시작 - 1초 간격 업데이트");
        }

        private void UpdateTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                WriteTimerLog($"타이머 틱 시작 - {DateTime.Now:HH:mm:ss.fff} (패널 수: {_monitorPanels.Count}개)");
                
                foreach (var panel in _monitorPanels.Values)
                {
                    panel.UpdateDisplay();
                }
                
                WriteTimerLog($"타이머 틱 완료 - {DateTime.Now:HH:mm:ss.fff}");
            }
            catch (Exception ex)
            {
                WriteTimerLog($"타이머 틱 오류: {ex.Message}");
            }
        }

        private void WriteTimerLog(string message)
        {
            try
            {
                string logFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "R854_Debug.txt");
                string logMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [MultiMonitoringForm] {message}";
                File.AppendAllText(logFilePath, logMessage + Environment.NewLine, System.Text.Encoding.UTF8);
            }
            catch
            {
                // 로그 쓰기 실패 시 무시
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            // 모든 패널의 근무조 데이터 저장
            SaveAllShiftData();

            _updateTimer?.Stop();
            _updateTimer?.Dispose();
            base.OnFormClosing(e);
        }

        private void MultiMonitoringForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            // FormClosing 이벤트 핸들러 (OnFormClosing에서 처리)
        }

        /// <summary>
        /// 모든 설비의 근무조 데이터 저장
        /// </summary>
        private void SaveAllShiftData()
        {
            foreach (var panel in _monitorPanels.Values)
            {
                try
                {
                    panel.SaveCurrentShiftData();
                }
                catch
                {
                    // 저장 실패 시 무시
                }
            }
        }
    }

    /// <summary>
    /// 개별 설비 모니터링 패널
    /// </summary>
    public class EquipmentMonitorPanel : GroupBox
    {
        private static readonly string _logFilePath = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory,
            "R854_Debug.txt");
        private static readonly object _logLock = new object();

        private CNCConnection _connection;
        private MainForm.IPConfig _config;

        // PMC 타입 상수
        private const short PMC_TYPE_G = 0;
        private const short PMC_TYPE_F = 1;
        private const short PMC_TYPE_R = 5;
        private const short PMC_TYPE_D = 7;
        private const short PMC_TYPE_X = 4;

        // 근무조 관리
        private ShiftInfo _currentShift;
        private Dictionary<MachineState, int> _stateDurations;
        private StateTransition _stateTransition;
        private PmcStateValues _currentPmcValues;
        private ShiftStateData _shiftStateData;  // 근무조 데이터 (DB 저장용)

        // M1 옵셔널 스톱 타이머
        private DateTime? _m1StartTime = null;

        // 이전 값 캐싱 (깜빡임 방지)
        private string _prevActualWorking = "";
        private string _prevInput = "";
        private string _prevAlarm = "";
        private string _prevIdle = "";
        private string _prevUnmeasured = "";
        private string _prevOperationRate = "";
        private string _prevProduction = "";

        // UI 컨트롤
        private Label _lblHeader;
        private Label _lblShiftInfo;
        private TableLayoutPanel _contentLayout;

        // 통합 지표
        private TableLayoutPanel _pnlActualWorking;
        private TableLayoutPanel _pnlInput;
        private TableLayoutPanel _pnlAlarm;
        private TableLayoutPanel _pnlIdle;
        private TableLayoutPanel _pnlUnmeasured;
        private Label _lblActualWorking;  // 상태 라벨 (색상 변경용)
        private Label _lblInput;
        private Label _lblAlarm;
        private Label _lblIdle;
        private Label _lblUnmeasured;
        private Label _lblActualWorkingTime;
        private Label _lblInputTime;
        private Label _lblAlarmTime;
        private Label _lblIdleTime;
        private Label _lblUnmeasuredTime;
        private Label _lblOperationRate;
        private Label _lblProduction;

        public EquipmentMonitorPanel(CNCConnection connection, MainForm.IPConfig config = null)
        {
            _connection = connection;
            _config = config ?? new MainForm.IPConfig
            {
                IpAddress = connection.IpAddress,
                LoadingMCode = 140  // 기본값
            };

            // 근무조 초기화
            _currentShift = ShiftManager.GetCurrentShift(DateTime.Now);
            _stateDurations = new Dictionary<MachineState, int>
            {
                { MachineState.Running, 0 },
                { MachineState.Loading, 0 },
                { MachineState.Alarm, 0 },
                { MachineState.Idle, 0 }
            };
            _stateTransition = new StateTransition
            {
                StartTime = DateTime.Now,
                CurrentState = MachineState.Idle,
                ElapsedSeconds = 0
            };

            // 근무조 데이터 초기화 (DB에서 복원 시도)
            LoadShiftDataFromDb();

            InitializeUI();
        }

        // WS_EX_COMPOSITED 스타일 적용 (깜박임 방지)
        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= 0x02000000;  // WS_EX_COMPOSITED
                return cp;
            }
        }

        /// <summary>
        /// PMC 주소 문자열을 파싱하여 타입, 주소, 비트 위치로 변환
        /// </summary>
        private bool ParsePmcAddress(string address, out short pmcType, out ushort pmcAddress, out int bitPosition)
        {
            pmcType = -1;  // -1을 sentinel 값으로 사용 (PMC_TYPE_G가 0이므로 0 사용 불가)
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
                _ => (short)-1
            };

            if (pmcType == -1)  // -1 체크 (0은 유효한 PMC_TYPE_G)
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

        private void InitializeUI()
        {
            // GroupBox 설정
            this.Size = new Size(720, 180);  // 높이 감소
            this.Margin = new Padding(8, 8, 8, 8);  // 패널 간 간격
            this.Font = new Font("맑은 고딕", 9.5f);  // 폰트 약간 증가
            this.DoubleBuffered = true;  // 깜박임 방지

            // 헤더 (별칭 또는 IP + 시간)
            string displayName = string.IsNullOrWhiteSpace(_config.Alias) ? _connection.IpAddress : $"{_config.Alias} ({_connection.IpAddress})";
            _lblHeader = new Label
            {
                Text = $"{displayName} | {DateTime.Now:HH:mm:ss}",
                Font = new Font("맑은 고딕", 11f, FontStyle.Bold),  // 폰트 증가
                Dock = DockStyle.Top,
                Height = 32,  // 높이 약간 증가
                TextAlign = ContentAlignment.MiddleLeft,
                BackColor = Color.FromArgb(60, 120, 180),
                ForeColor = Color.White,
                Padding = new Padding(10, 0, 0, 0)
            };

            // 근무조 정보
            _lblShiftInfo = new Label
            {
                Text = ShiftManager.GetShiftDisplayName(_currentShift),
                Font = new Font("맑은 고딕", 9.5f),  // 폰트 증가
                Dock = DockStyle.Top,
                Height = 26,  // 높이 약간 증가
                TextAlign = ContentAlignment.MiddleLeft,
                BackColor = Color.FromArgb(144, 238, 144),
                Padding = new Padding(10, 0, 0, 0)
            };

            // 컨텐츠 레이아웃 (1행)
            _contentLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 1,
                ColumnCount = 6,
                CellBorderStyle = TableLayoutPanelCellBorderStyle.Single,
                Padding = new Padding(5)
            };

            // 행 비율
            _contentLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));  // 통합 지표

            // 열 비율
            for (int i = 0; i < 6; i++)
            {
                _contentLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 16.67F));
            }

            // Row 0: 통합 지표 (6개 카드)
            _pnlActualWorking = CreateMetricCard("가공", "00:00:00", Color.LightGray, out _lblActualWorking, out _lblActualWorkingTime);
            _pnlInput = CreateMetricCard("투입", "00:00:00", Color.LightGray, out _lblInput, out _lblInputTime);
            _pnlAlarm = CreateMetricCard("알람", "00:00:00", Color.LightGray, out _lblAlarm, out _lblAlarmTime);
            _pnlIdle = CreateMetricCard("유휴", "00:00:00", Color.LightGray, out _lblIdle, out _lblIdleTime);
            _pnlUnmeasured = CreateMetricCard("미측정", "00:00:00", Color.FromArgb(220, 220, 220), out _lblUnmeasured, out _lblUnmeasuredTime);

            // 가동률 + 생산수량을 하나의 패널에
            var metricsPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 2,
                ColumnCount = 1
            };
            metricsPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
            metricsPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));

            _lblOperationRate = CreateMetricLabel("가동률: 0%", Color.FromArgb(255, 255, 153));
            _lblProduction = CreateMetricLabel("생산: 0개", Color.LightCyan);

            metricsPanel.Controls.Add(_lblOperationRate, 0, 0);
            metricsPanel.Controls.Add(_lblProduction, 0, 1);

            _contentLayout.Controls.Add(_pnlActualWorking, 0, 0);
            _contentLayout.Controls.Add(_pnlInput, 1, 0);
            _contentLayout.Controls.Add(_pnlAlarm, 2, 0);
            _contentLayout.Controls.Add(_pnlIdle, 3, 0);
            _contentLayout.Controls.Add(_pnlUnmeasured, 4, 0);
            _contentLayout.Controls.Add(metricsPanel, 5, 0);

            // 컨트롤 추가
            this.Controls.Add(_contentLayout);
            this.Controls.Add(_lblShiftInfo);
            this.Controls.Add(_lblHeader);
        }

        private TableLayoutPanel CreateMetricCard(string labelText, string timeText, Color backColor, out Label stateLabel, out Label timeLabel)
        {
            var panel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 2,
                ColumnCount = 1,
                BackColor = Color.Transparent
            };
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 40F));  // 라벨
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 60F));  // 시간

            stateLabel = new Label
            {
                Text = labelText,
                Font = new Font("맑은 고딕", 11f, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Fill,
                BackColor = backColor  // 라벨만 색상 적용
            };

            timeLabel = new Label
            {
                Text = timeText,
                Font = new Font("Consolas", 14f, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Fill,
                BackColor = Color.Transparent  // 투명 (깜박임 방지)
            };

            panel.Controls.Add(stateLabel, 0, 0);
            panel.Controls.Add(timeLabel, 0, 1);

            return panel;
        }

        private Label CreateMetricLabel(string text, Color backColor)
        {
            return new Label
            {
                Text = text,
                Font = new Font("맑은 고딕", 10f, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Fill,
                BackColor = backColor,
                BorderStyle = BorderStyle.FixedSingle
            };
        }

        public void UpdateDisplay()
        {
            if (_connection == null || !_connection.IsConnected)
            {
                string displayName = string.IsNullOrWhiteSpace(_config.Alias) ? (_connection?.IpAddress ?? "N/A") : $"{_config.Alias} ({_connection?.IpAddress ?? "N/A"})";
                _lblHeader.Text = $"{displayName} | 연결 안됨";
                _lblHeader.BackColor = Color.Gray;
                return;
            }

            try
            {
                // 헤더 업데이트
                string displayName = string.IsNullOrWhiteSpace(_config.Alias) ? _connection.IpAddress : $"{_config.Alias} ({_connection.IpAddress})";
                _lblHeader.Text = $"{displayName} | {DateTime.Now:HH:mm:ss}";

                // PMC 값 읽기 및 상태 판별
                if (ReadPmcStateValues(out PmcStateValues pmc))
                {
                    _currentPmcValues = pmc;

                    // 알람 상태 체크 (설정된 PMC 알람 주소 사용)
                    bool hasAlarmPmcA = false;
                    string alarmDetail = "";
                    try
                    {
                        // 설정된 알람 PMC 주소 파싱
                        if (ParsePmcAddress(_config.PmcR854_2, out short alarmType, out ushort alarmAddr, out int alarmBit))
                        {
                            Focas1.IODBPMC0 pmcAlarm = new Focas1.IODBPMC0();
                            short retAlarm = Focas1.pmc_rdpmcrng(_connection.Handle, alarmType, 0, alarmAddr, alarmAddr, 9, pmcAlarm);

                            if (retAlarm == Focas1.EW_OK)
                            {
                                byte byteValue = pmcAlarm.cdata[0];
                                bool bit0 = (byteValue & (1 << 0)) != 0;
                                bool bit1 = (byteValue & (1 << 1)) != 0;
                                bool bit2 = (byteValue & (1 << 2)) != 0;
                                bool bit3 = (byteValue & (1 << 3)) != 0;
                                bool bit4 = (byteValue & (1 << 4)) != 0;
                                bool bit5 = (byteValue & (1 << 5)) != 0;
                                bool bit6 = (byteValue & (1 << 6)) != 0;
                                bool bit7 = (byteValue & (1 << 7)) != 0;

                                // 설정된 비트 위치의 값을 사용
                                hasAlarmPmcA = ((byteValue >> alarmBit) & 1) == 1;
                                alarmDetail = $"{_config.PmcR854_2} byte=0x{byteValue:X2} (Bits: 7={B(bit7)} 6={B(bit6)} 5={B(bit5)} 4={B(bit4)} 3={B(bit3)} 2={B(bit2)} 1={B(bit1)} 0={B(bit0)})";

                                // 로그 파일에 기록
                                WriteStateLog($"[{_connection.IpAddress}] {alarmDetail}, 알람판정={hasAlarmPmcA}, 현재상태={_stateTransition.CurrentState}");

                                // 상태 전환 시 특별 로그
                                if (_stateTransition.CurrentState != MachineState.Alarm && hasAlarmPmcA)
                                {
                                    WriteStateLog($"[{_connection.IpAddress}] ★★★ 알람 감지! {_stateTransition.CurrentState} -> ALARM ★★★");
                                }
                                else if (_stateTransition.CurrentState == MachineState.Alarm && !hasAlarmPmcA)
                                {
                                    WriteStateLog($"[{_connection.IpAddress}] ★★★ 알람 해제! ALARM -> ? ★★★");
                                }
                            }
                            else
                            {
                                alarmDetail = $"{_config.PmcR854_2} 읽기 실패 (ret={retAlarm})";
                                WriteStateLog($"[{_connection.IpAddress}] {alarmDetail}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        alarmDetail = $"예외발생: {ex.Message}";
                    }

                    bool isAlarmState = hasAlarmPmcA;

                    MachineState currentState = DetermineMachineState(pmc, hasAlarmPmcA);

                    // 상태 전환 처리
                    if (_stateTransition.CurrentState != currentState)
                    {
                        int duration = (int)(DateTime.Now - _stateTransition.StartTime).TotalSeconds;
                        _stateDurations[_stateTransition.CurrentState] += duration;

                        // 상태 변경 로그 작성
                        WriteStateLog($"[{_connection.IpAddress}] 상태변경: {_stateTransition.CurrentState} → {currentState} " +
                            $"(지속시간: {duration}초) | PMC: F0.7={B(pmc.F0_7)} F1.0={B(pmc.F1_0)} " +
                            $"F3.5={B(pmc.F3_5)} F10={pmc.F10_value} G4.3={B(pmc.G4_3)} G5.0={B(pmc.G5_0)} AlarmPMC={hasAlarmPmcA}");

                        // 사이클 로그 저장 (10초 단위, Loading/Running만)
                        if (_stateTransition.CurrentState == MachineState.Loading || 
                            _stateTransition.CurrentState == MachineState.Running)
                        {
                            try
                            {
                                var logService = new LogDataService();
                                int? cycleNumber = null;
                                
                                // Loading → Running 전환 시 사이클 번호 부여
                                if (_stateTransition.CurrentState == MachineState.Loading && currentState == MachineState.Running)
                                {
                                    cycleNumber = _connection.GetProductionCount() + 1; // 다음 생산수량
                                }
                                
                                logService.SaveCycleStateLog(
                                    _connection.IpAddress,
                                    _currentShift.StartTime.Date,
                                    _currentShift.Type,
                                    _stateTransition.CurrentState,
                                    duration,
                                    _stateTransition.StartTime,
                                    cycleNumber
                                );
                            }
                            catch (Exception ex)
                            {
                                WriteStateLog($"[{_connection.IpAddress}] 사이클 로그 저장 실패: {ex.Message}");
                            }
                        }

                        // 투입 → 실가공 전환 시 생산 수량 증가
                        if (_stateTransition.CurrentState == MachineState.Loading && currentState == MachineState.Running)
                        {
                            _connection.IncrementProduction();
                            WriteStateLog($"[{_connection.IpAddress}] 생산수량 증가: {_connection.GetProductionCount()}개");
                        }

                        // 근무조 데이터 업데이트
                        _shiftStateData.RunningSeconds = _stateDurations[MachineState.Running];
                        _shiftStateData.LoadingSeconds = _stateDurations[MachineState.Loading];
                        _shiftStateData.AlarmSeconds = _stateDurations[MachineState.Alarm];
                        _shiftStateData.IdleSeconds = _stateDurations[MachineState.Idle];
                        _shiftStateData.ProductionCount = _connection.GetProductionCount();

                        _stateTransition.StartTime = DateTime.Now;
                        _stateTransition.CurrentState = currentState;
                        _stateTransition.ElapsedSeconds = 0;
                    }
                    else
                    {
                        _stateTransition.ElapsedSeconds = (int)(DateTime.Now - _stateTransition.StartTime).TotalSeconds;
                    }

                    // 현재 진행 중인 시간을 포함하여 근무조 데이터 매번 업데이트
                    _shiftStateData.RunningSeconds = _stateDurations[MachineState.Running] +
                        (currentState == MachineState.Running ? _stateTransition.ElapsedSeconds : 0);
                    _shiftStateData.LoadingSeconds = _stateDurations[MachineState.Loading] +
                        (currentState == MachineState.Loading ? _stateTransition.ElapsedSeconds : 0);
                    _shiftStateData.AlarmSeconds = _stateDurations[MachineState.Alarm] +
                        (currentState == MachineState.Alarm ? _stateTransition.ElapsedSeconds : 0);
                    _shiftStateData.IdleSeconds = _stateDurations[MachineState.Idle] +
                        (currentState == MachineState.Idle ? _stateTransition.ElapsedSeconds : 0);
                    _shiftStateData.ProductionCount = _connection.GetProductionCount();

                    // 근무조 전환 감지
                    CheckAndHandleShiftTransition();

                    // UI 업데이트
                    UpdateStateUI(currentState);
                }
            }
            catch (Exception ex)
            {
                // 오류 무시
            }
        }

        private void UpdateStateUI(MachineState currentState)
        {
            this.SuspendLayout();  // 레이아웃 업데이트 일시 중지 (깜박임 방지)

            // 현재 상태 시간 포함
            int runningTotal = _stateDurations[MachineState.Running] + (currentState == MachineState.Running ? _stateTransition.ElapsedSeconds : 0);
            int loadingTotal = _stateDurations[MachineState.Loading] + (currentState == MachineState.Loading ? _stateTransition.ElapsedSeconds : 0);
            int alarmTotal = _stateDurations[MachineState.Alarm] + (currentState == MachineState.Alarm ? _stateTransition.ElapsedSeconds : 0);
            int idleTotal = _stateDurations[MachineState.Idle] + (currentState == MachineState.Idle ? _stateTransition.ElapsedSeconds : 0);

            // 통합 지표 계산
            int actualWorking = runningTotal;
            int input = loadingTotal;  // 투입시간 = Loading (제품교체 시간만)
            int alarm = alarmTotal;
            int idle = idleTotal;
            int unmeasured = _shiftStateData.UnmeasuredSeconds;
            int total = runningTotal + loadingTotal + idleTotal + alarmTotal + unmeasured;
            double operationRate = total > 0 ? (double)actualWorking / total * 100 : 0;

            // 시간 텍스트
            string actualWorkingText = FormatDuration(actualWorking);
            string inputText = FormatDuration(input);
            string alarmText = FormatDuration(alarm);
            string idleText = FormatDuration(idle);
            string unmeasuredText = FormatDuration(unmeasured);
            string operationRateText = $"가동률: {operationRate:F1}%";
            int productionCount = _connection.GetProductionCount();
            string productionText = $"생산: {productionCount}개";

            // 카드 라벨 색상 초기화 (모두 회색)
            _lblActualWorking.BackColor = Color.LightGray;
            _lblInput.BackColor = Color.LightGray;
            _lblAlarm.BackColor = Color.LightGray;
            _lblIdle.BackColor = Color.LightGray;
            _lblUnmeasured.BackColor = Color.FromArgb(220, 220, 220);

            // 현재 상태에 따라 해당 라벨만 강조
            switch (currentState)
            {
                case MachineState.Running:
                    _lblActualWorking.BackColor = Color.FromArgb(76, 175, 80);  // 녹색
                    break;
                case MachineState.Loading:
                    _lblInput.BackColor = Color.FromArgb(255, 193, 7);  // 노란색
                    break;
                case MachineState.Alarm:
                    _lblAlarm.BackColor = Color.FromArgb(244, 67, 54);  // 빨간색
                    break;
                case MachineState.Idle:
                    _lblIdle.BackColor = Color.FromArgb(158, 158, 158);  // 회색
                    break;
            }

            // 시간 업데이트 (변경된 경우에만)
            if (actualWorkingText != _prevActualWorking)
            {
                _lblActualWorkingTime.Text = actualWorkingText;
                _prevActualWorking = actualWorkingText;
            }
            if (inputText != _prevInput)
            {
                _lblInputTime.Text = inputText;
                _prevInput = inputText;
            }
            if (alarmText != _prevAlarm)
            {
                _lblAlarmTime.Text = alarmText;
                _prevAlarm = alarmText;
            }
            if (idleText != _prevIdle)
            {
                _lblIdleTime.Text = idleText;
                _prevIdle = idleText;
            }
            if (unmeasuredText != _prevUnmeasured)
            {
                _lblUnmeasuredTime.Text = unmeasuredText;
                _prevUnmeasured = unmeasuredText;
            }
            if (operationRateText != _prevOperationRate)
            {
                _lblOperationRate.Text = operationRateText;
                _prevOperationRate = operationRateText;
            }
            if (productionText != _prevProduction)
            {
                _lblProduction.Text = productionText;
                _prevProduction = productionText;
            }

            this.ResumeLayout();  // 레이아웃 업데이트 재개
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
                    WriteStateLog($"[{_connection.IpAddress}] ★ G5.0 파싱: 문자열='{_config.PmcG5_0}', Type={type_g5_0}, Addr={addr_g5_0}, Bit={bit_g5_0}");
                    if (!ReadPmcBit(type_g5_0, (short)addr_g5_0, (short)bit_g5_0, out values.G5_0))
                    {
                        WriteStateLog($"[{_connection.IpAddress}] ★ G5.0 읽기 실패!");
                        return false;
                    }
                    WriteStateLog($"[{_connection.IpAddress}] ★ G5.0 읽기 성공: {values.G5_0}");
                }
                else
                {
                    WriteStateLog($"[{_connection.IpAddress}] ★ G5.0 파싱 실패: '{_config.PmcG5_0}'");
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

                    // G5.0 읽을 때만 상세 로그
                    if (pmcType == 0 && address == 5 && bitPosition == 0)
                    {
                        WriteStateLog($"[{_connection.IpAddress}] ★★ ReadPmcBit 상세: Type={pmcType}, Addr={address}, Bit={bitPosition}, Byte=0x{byteValue:X2}, Result={value}");
                    }

                    return true;
                }
                else
                {
                    WriteStateLog($"[{_connection.IpAddress}] ★★ ReadPmcBit 실패: Type={pmcType}, Addr={address}, Bit={bitPosition}, ret={ret}");
                }

                return false;
            }
            catch (Exception ex)
            {
                WriteStateLog($"[{_connection.IpAddress}] ★★ ReadPmcBit 예외: Type={pmcType}, Addr={address}, Bit={bitPosition}, Error={ex.Message}");
                return false;
            }
        }

        private string B(bool value)
        {
            return value ? "1" : "0";
        }

        private void WriteStateLog(string message)
        {
            try
            {
                lock (_logLock)
                {
                    string logMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {message}";
                    File.AppendAllText(_logFilePath, logMessage + Environment.NewLine);
                }
            }
            catch
            {
                // 로그 쓰기 실패 시 무시
            }
        }

        /// <summary>
        /// 근무조 전환 감지 및 처리
        /// </summary>
        private void CheckAndHandleShiftTransition()
        {
            ShiftInfo newShift = ShiftManager.GetCurrentShift(DateTime.Now);

            // 근무조 변경 감지
            if (ShiftManager.IsShiftChanged(_currentShift, newShift))
            {
                // 현재 근무조 데이터 저장
                SaveShiftData();

                // 이전 근무조 리포트 생성 및 저장 (리포트는 ShiftDataViewForm에서 표시)
                try
                {
                    var logService = new LogDataService();
                    var report = logService.GenerateCycleReport(
                        _connection.IpAddress,
                        _currentShift.StartTime.Date,
                        _currentShift.Type
                    );
                    WriteStateLog($"[{_connection.IpAddress}] 이전 근무조 리포트 생성 완료 - 총 사이클: {report.TotalCycles}개, 특이사항: {report.AnomalyCount}개");
                }
                catch (Exception ex)
                {
                    WriteStateLog($"[{_connection.IpAddress}] 리포트 생성 실패: {ex.Message}");
                }

                // 새 근무조로 전환
                _currentShift = newShift;

                // 다음날 같은 근무조 시작 시 이전 데이터 삭제
                try
                {
                    var logService = new LogDataService();
                    DateTime previousShiftDate = _currentShift.StartTime.Date.AddDays(-1);
                    logService.DeleteCycleLogs(_connection.IpAddress, previousShiftDate, _currentShift.Type);
                    WriteStateLog($"[{_connection.IpAddress}] 이전 근무조 사이클 로그 삭제 완료 ({previousShiftDate:yyyy-MM-dd} {_currentShift.Type})");
                }
                catch (Exception ex)
                {
                    WriteStateLog($"[{_connection.IpAddress}] 사이클 로그 삭제 실패: {ex.Message}");
                }

                // 상태 추적 초기화
                _stateDurations[MachineState.Running] = 0;
                _stateDurations[MachineState.Loading] = 0;
                _stateDurations[MachineState.Alarm] = 0;
                _stateDurations[MachineState.Idle] = 0;

                // 근무조 데이터 초기화
                _shiftStateData = new ShiftStateData
                {
                    IpAddress = _connection.IpAddress,
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
                _connection.ResetProductionCount();

                // 상태 전환 시작 시간 초기화 (UI 시간 표시 초기화)
                _stateTransition.StartTime = DateTime.Now;
                _stateTransition.ElapsedSeconds = 0;

                // UI 업데이트
                if (_lblShiftInfo != null)
                {
                    _lblShiftInfo.Text = ShiftManager.GetShiftDisplayName(newShift);
                }

                WriteStateLog($"[{_connection.IpAddress}] 근무조 전환: {_currentShift.Type} → {newShift.Type}, UI 시간 초기화");
            }
        }

        /// <summary>
        /// 현재 근무조 데이터 DB 저장
        /// </summary>
        private void SaveShiftData()
        {
            try
            {
                DateTime updateTime = DateTime.Now;
                
                // 현재 진행 중인 상태의 시간을 먼저 누적에 반영
                MachineState currentState = _stateTransition.CurrentState;
                int elapsedSeconds = (int)(updateTime - _stateTransition.StartTime).TotalSeconds;
                
                // 현재 상태에 진행 중인 시간 추가
                if (elapsedSeconds > 0)
                {
                    _stateDurations[currentState] += elapsedSeconds;
                    WriteStateLog($"[{_connection.IpAddress}] 근무조 변경 시 현재 상태({currentState}) {elapsedSeconds}초 추가 반영");
                }
                
                // 현재 상태의 누적 시간 최종 반영
                _shiftStateData.RunningSeconds = _stateDurations[MachineState.Running];
                _shiftStateData.LoadingSeconds = _stateDurations[MachineState.Loading];
                _shiftStateData.AlarmSeconds = _stateDurations[MachineState.Alarm];
                _shiftStateData.IdleSeconds = _stateDurations[MachineState.Idle];
                _shiftStateData.ProductionCount = _connection.GetProductionCount();

                // DB 저장
                var logService = new LogDataService();
                logService.SaveShiftStateData(_shiftStateData);

                WriteStateLog($"[{_connection.IpAddress}] ★★★ 근무조 데이터 저장 완료 ★★★ " +
                    $"실가공={FormatDuration(_shiftStateData.RunningSeconds)}, " +
                    $"투입={FormatDuration(_shiftStateData.LoadingSeconds)}, " +
                    $"알람={FormatDuration(_shiftStateData.AlarmSeconds)}, " +
                    $"유휴={FormatDuration(_shiftStateData.IdleSeconds)}, " +
                    $"생산={_shiftStateData.ProductionCount}개");
            }
            catch (Exception ex)
            {
                WriteStateLog($"[{_connection.IpAddress}] 근무조 데이터 저장 실패: {ex.Message}");
            }
        }

        /// <summary>
        /// 수동으로 근무조 데이터 저장 (외부 호출용)
        /// </summary>
        public void SaveCurrentShiftData()
        {
            SaveShiftData();
        }

        /// <summary>
        /// DB에서 오늘 현재 근무조의 누적 데이터 복원
        /// </summary>
        private void LoadShiftDataFromDb()
        {
            try
            {
                var logService = new LogDataService();
                var savedData = logService.LoadShiftStateData(
                    _connection.IpAddress,
                    _currentShift.StartTime.Date,
                    _currentShift.Type);

                if (savedData != null)
                {
                    // 마지막 업데이트 시간 조회
                    var lastUpdatedTime = logService.GetLastUpdatedTime(
                        _connection.IpAddress,
                        _currentShift.StartTime.Date,
                        _currentShift.Type);

                    // 미측정 시간 계산: 마지막 업데이트 ~ 현재 시간
                    if (lastUpdatedTime.HasValue)
                    {
                        DateTime now = DateTime.Now;

                        // 현재 시간이 같은 근무조 내에 있는지 확인
                        if (now >= _currentShift.StartTime && now < _currentShift.EndTime)
                        {
                            TimeSpan unmeasuredGap = now - lastUpdatedTime.Value;

                            // 음수 방지 (시간이 잘못된 경우)
                            if (unmeasuredGap.TotalSeconds > 0)
                            {
                                int additionalUnmeasured = (int)unmeasuredGap.TotalSeconds;
                                savedData.UnmeasuredSeconds += additionalUnmeasured;
                                
                                // 미측정 시간 계산 후 즉시 DB에 저장
                                logService.UpdateShiftStateData(savedData);
                                
                                WriteStateLog($"[{_connection.IpAddress}] 미측정 시간 추가 및 DB 저장: " +
                                    $"{FormatDuration(additionalUnmeasured)} " +
                                    $"(마지막 업데이트: {lastUpdatedTime.Value:yyyy-MM-dd HH:mm:ss})");
                            }
                        }
                    }

                    // DB에서 복원
                    _shiftStateData = savedData;
                    _stateDurations[MachineState.Running] = savedData.RunningSeconds;
                    _stateDurations[MachineState.Loading] = savedData.LoadingSeconds;
                    _stateDurations[MachineState.Alarm] = savedData.AlarmSeconds;
                    _stateDurations[MachineState.Idle] = savedData.IdleSeconds;

                    // 생산 수량도 복원
                    _connection.ResetProductionCount();
                    for (int i = 0; i < savedData.ProductionCount; i++)
                    {
                        _connection.IncrementProduction();
                    }

                    WriteStateLog($"[{_connection.IpAddress}] DB에서 근무조 데이터 복원 완료: " +
                        $"실가공={FormatDuration(savedData.RunningSeconds)}, " +
                        $"투입={FormatDuration(savedData.LoadingSeconds)}, " +
                        $"알람={FormatDuration(savedData.AlarmSeconds)}, " +
                        $"유휴={FormatDuration(savedData.IdleSeconds)}, " +
                        $"미측정={FormatDuration(savedData.UnmeasuredSeconds)}, " +
                        $"생산={savedData.ProductionCount}개");
                }
                else
                {
                    // DB에 데이터 없음 - 새로 시작
                    _shiftStateData = new ShiftStateData
                    {
                        IpAddress = _connection.IpAddress,
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

                    WriteStateLog($"[{_connection.IpAddress}] 새 근무조 시작");
                }
            }
            catch (Exception ex)
            {
                // DB 복원 실패 시 0으로 시작
                _shiftStateData = new ShiftStateData
                {
                    IpAddress = _connection.IpAddress,
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

                WriteStateLog($"[{_connection.IpAddress}] DB 복원 실패, 0으로 시작: {ex.Message}");
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

        private MachineState DetermineMachineState(PmcStateValues pmc, bool hasAlarmCnc)
        {
            // 알람 체크: F1.0 (PMC 신호) OR CNC 알람
            if (pmc.F1_0 || hasAlarmCnc)
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
            // 스타트 실행 중 AND 메모리 모드 AND M코드 실행 AND (M핀 처리 중 OR G5.0)
            bool loadingCondition1 = _config.LoadingMCode > 0;
            bool loadingCondition2 = pmc.F0_7;
            bool loadingCondition3 = pmc.F3_5;
            bool loadingCondition4 = pmc.F10_value == _config.LoadingMCode;
            bool loadingCondition5 = pmc.G4_3 || pmc.G5_0;

            // 로딩 조건 디버깅
            if (_config.LoadingMCode > 0)  // M코드가 설정되어 있을 때만 로그
            {
                WriteStateLog($"[{_connection.IpAddress}] 로딩체크: MCode설정={loadingCondition1}({_config.LoadingMCode}), " +
                    $"F0.7={loadingCondition2}, F3.5={loadingCondition3}, " +
                    $"F10={pmc.F10_value}(설정={_config.LoadingMCode},일치={loadingCondition4}), " +
                    $"G4.3/G5.0={loadingCondition5}(G4.3={pmc.G4_3},G5.0={pmc.G5_0})");
            }

            if (loadingCondition1 && loadingCondition2 && loadingCondition3 && loadingCondition4 && loadingCondition5)
                return MachineState.Loading;

            // 가공중: 자동 운전 신호 AND 메모리 모드
            if (pmc.F0_7 && pmc.F3_5)
                return MachineState.Running;

            return MachineState.Idle;
        }

        private string FormatDuration(int seconds)
        {
            int hours = seconds / 3600;
            int minutes = (seconds % 3600) / 60;
            int secs = seconds % 60;
            return $"{hours:D2}:{minutes:D2}:{secs:D2}";
        }
    }

    // ==================== 설비 상태 모니터링 관련 ====================

    // 설비 상태 Enum (우선순위 순)
    public enum MachineState
    {
        Alarm,      // 장애 정지 (최우선)
        Loading,    // 제품 교체 중
        Running,    // 가공 중
        Idle        // 유휴 상태
    }

    // PMC 상태 값 구조체
    public struct PmcStateValues
    {
        public bool F0_0;   // ST_OP 가동 신호
        public bool F0_7;   // 스타트 실행 중
        public bool F1_0;   // 알람 신호
        public bool F3_5;   // 메모리 모드
        public short F10_value;  // F10 주소 값 (M코드 번호)
        public bool G4_3;   // M핀 처리 중
        public bool G5_0;   // 추가 투입 조건
        public bool X8_4;   // 비상정지 신호
    }

    // 상태 전환 추적 클래스
    public class StateTransition
    {
        public DateTime StartTime { get; set; }
        public MachineState CurrentState { get; set; }
        public int ElapsedSeconds { get; set; }
    }

    // 설비 상태 로그 클래스
    public class MachineStateLog
    {
        public DateTime Timestamp { get; set; }
        public string IpAddress { get; set; }
        public MachineState State { get; set; }
        public int DurationSeconds { get; set; }
        public string PmcValues { get; set; }  // JSON 형식
    }
}
