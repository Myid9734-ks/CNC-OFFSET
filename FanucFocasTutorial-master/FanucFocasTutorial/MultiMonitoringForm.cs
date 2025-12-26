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
        private Dictionary<string, string> _ipAliases;

        public MultiMonitoringForm(List<CNCConnection> connections, Dictionary<string, string> ipAliases = null)
        {
            _connections = connections;
            _ipAliases = ipAliases ?? new Dictionary<string, string>();
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

                // 별칭 가져오기 (없으면 null)
                string alias = _ipAliases.ContainsKey(connection.IpAddress)
                    ? _ipAliases[connection.IpAddress]
                    : null;

                var panel = new EquipmentMonitorPanel(connection, alias);
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
        }

        private void UpdateTimer_Tick(object sender, EventArgs e)
        {
            foreach (var panel in _monitorPanels.Values)
            {
                panel.UpdateDisplay();
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
        private string _alias;

        // 근무조 관리
        private ShiftInfo _currentShift;
        private Dictionary<MachineState, int> _stateDurations;
        private StateTransition _stateTransition;
        private PmcStateValues _currentPmcValues;
        private ShiftStateData _shiftStateData;  // 근무조 데이터 (DB 저장용)

        // 이전 값 캐싱 (깜빡임 방지)
        private string _prevActualWorking = "";
        private string _prevInput = "";
        private string _prevAlarm = "";
        private string _prevIdle = "";
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
        private Label _lblActualWorkingTime;
        private Label _lblInputTime;
        private Label _lblAlarmTime;
        private Label _lblIdleTime;
        private Label _lblOperationRate;
        private Label _lblProduction;

        public EquipmentMonitorPanel(CNCConnection connection, string alias = null)
        {
            _connection = connection;
            _alias = alias;

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

        private void InitializeUI()
        {
            // GroupBox 설정
            this.Size = new Size(720, 180);  // 높이 감소
            this.Margin = new Padding(8, 8, 8, 8);  // 패널 간 간격
            this.Font = new Font("맑은 고딕", 9.5f);  // 폰트 약간 증가
            this.DoubleBuffered = true;  // 깜박임 방지

            // 헤더 (별칭 또는 IP + 시간)
            string displayName = string.IsNullOrWhiteSpace(_alias) ? _connection.IpAddress : $"{_alias} ({_connection.IpAddress})";
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
                ColumnCount = 5,
                CellBorderStyle = TableLayoutPanelCellBorderStyle.Single,
                Padding = new Padding(5)
            };

            // 행 비율
            _contentLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));  // 통합 지표

            // 열 비율
            for (int i = 0; i < 5; i++)
            {
                _contentLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20F));
            }

            // Row 0: 통합 지표 (5개 카드)
            _pnlActualWorking = CreateMetricCard("가공", "00:00:00", Color.LightGray, out _lblActualWorkingTime);
            _pnlInput = CreateMetricCard("투입", "00:00:00", Color.LightGray, out _lblInputTime);
            _pnlAlarm = CreateMetricCard("알람", "00:00:00", Color.LightGray, out _lblAlarmTime);
            _pnlIdle = CreateMetricCard("유휴", "00:00:00", Color.LightGray, out _lblIdleTime);

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
            _contentLayout.Controls.Add(metricsPanel, 4, 0);

            // 컨트롤 추가
            this.Controls.Add(_contentLayout);
            this.Controls.Add(_lblShiftInfo);
            this.Controls.Add(_lblHeader);
        }

        private TableLayoutPanel CreateMetricCard(string labelText, string timeText, Color backColor, out Label timeLabel)
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

            var label = new Label
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
                BackColor = Color.White  // 타이머는 항상 흰색
            };

            panel.Controls.Add(label, 0, 0);
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
                string displayName = string.IsNullOrWhiteSpace(_alias) ? (_connection?.IpAddress ?? "N/A") : $"{_alias} ({_connection?.IpAddress ?? "N/A"})";
                _lblHeader.Text = $"{displayName} | 연결 안됨";
                _lblHeader.BackColor = Color.Gray;
                return;
            }

            try
            {
                // 헤더 업데이트
                string displayName = string.IsNullOrWhiteSpace(_alias) ? _connection.IpAddress : $"{_alias} ({_connection.IpAddress})";
                _lblHeader.Text = $"{displayName} | {DateTime.Now:HH:mm:ss}";

                // PMC 값 읽기 및 상태 판별
                if (ReadPmcStateValues(out PmcStateValues pmc))
                {
                    _currentPmcValues = pmc;

                    // 알람 상태 체크 (PMC R854.2만 사용)
                    // PMC R854.2 체크 (알람)
                    bool hasAlarmPmcA = false;
                    string alarmDetail = "";
                    try
                    {
                        const short PMC_TYPE_R = 5;  // R-type (Internal relay)

                        // R854.2 체크
                        Focas1.IODBPMC0 pmcR854 = new Focas1.IODBPMC0();
                        short ret854 = Focas1.pmc_rdpmcrng(_connection.Handle, PMC_TYPE_R, 0, 854, 854, 9, pmcR854);

                        if (ret854 == Focas1.EW_OK)
                        {
                            byte byteValue = pmcR854.cdata[0];
                            bool bit0 = (byteValue & (1 << 0)) != 0;
                            bool bit1 = (byteValue & (1 << 1)) != 0;
                            bool bit2 = (byteValue & (1 << 2)) != 0;
                            bool bit3 = (byteValue & (1 << 3)) != 0;
                            bool bit4 = (byteValue & (1 << 4)) != 0;
                            bool bit5 = (byteValue & (1 << 5)) != 0;
                            bool bit6 = (byteValue & (1 << 6)) != 0;
                            bool bit7 = (byteValue & (1 << 7)) != 0;

                            hasAlarmPmcA = bit2;
                            alarmDetail = $"R854 byte=0x{byteValue:X2} (Bits: 7={B(bit7)} 6={B(bit6)} 5={B(bit5)} 4={B(bit4)} 3={B(bit3)} 2={B(bit2)} 1={B(bit1)} 0={B(bit0)})";

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
                            alarmDetail = $"R854 읽기 실패 (ret={ret854})";
                            WriteStateLog($"[{_connection.IpAddress}] {alarmDetail}");
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

                        // 투입 → 실가공 전환 시 생산 수량 증가
                        if (_stateTransition.CurrentState == MachineState.Loading && currentState == MachineState.Running)
                        {
                            _connection.IncrementProduction();
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
            // 현재 상태 시간 포함
            int runningTotal = _stateDurations[MachineState.Running] + (currentState == MachineState.Running ? _stateTransition.ElapsedSeconds : 0);
            int loadingTotal = _stateDurations[MachineState.Loading] + (currentState == MachineState.Loading ? _stateTransition.ElapsedSeconds : 0);
            int alarmTotal = _stateDurations[MachineState.Alarm] + (currentState == MachineState.Alarm ? _stateTransition.ElapsedSeconds : 0);
            int idleTotal = _stateDurations[MachineState.Idle] + (currentState == MachineState.Idle ? _stateTransition.ElapsedSeconds : 0);

            // 통합 지표 계산
            int actualWorking = runningTotal;
            int input = runningTotal + loadingTotal;
            int alarm = alarmTotal;
            int idle = idleTotal;
            int total = runningTotal + loadingTotal + idleTotal + alarmTotal;
            double operationRate = total > 0 ? (double)actualWorking / total * 100 : 0;

            // 시간 텍스트
            string actualWorkingText = FormatDuration(actualWorking);
            string inputText = FormatDuration(input);
            string alarmText = FormatDuration(alarm);
            string idleText = FormatDuration(idle);
            string operationRateText = $"가동률: {operationRate:F1}%";
            int productionCount = _connection.GetProductionCount();
            string productionText = $"생산: {productionCount}개";

            // 카드 색상 초기화
            _pnlActualWorking.BackColor = Color.LightGray;
            _pnlInput.BackColor = Color.LightGray;
            _pnlAlarm.BackColor = Color.LightGray;
            _pnlIdle.BackColor = Color.LightGray;

            // 현재 상태에 따라 카드 강조
            switch (currentState)
            {
                case MachineState.Running:
                    _pnlActualWorking.BackColor = Color.FromArgb(76, 175, 80);  // 녹색
                    break;
                case MachineState.Loading:
                    _pnlInput.BackColor = Color.FromArgb(255, 193, 7);  // 노란색
                    break;
                case MachineState.Alarm:
                    _pnlAlarm.BackColor = Color.FromArgb(244, 67, 54);  // 빨간색
                    break;
                case MachineState.Idle:
                    _pnlIdle.BackColor = Color.FromArgb(158, 158, 158);  // 회색
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
        }

        private bool ReadPmcStateValues(out PmcStateValues values)
        {
            values = new PmcStateValues();

            if (_connection == null || !_connection.IsConnected)
                return false;

            try
            {
                if (!ReadPmcBit(1, 0, 0, out values.F0_0)) return false;
                if (!ReadPmcBit(1, 0, 7, out values.F0_7)) return false;
                if (!ReadPmcBit(1, 1, 0, out values.F1_0)) return false;
                if (!ReadPmcBit(1, 3, 5, out values.F3_5)) return false;
                if (!ReadPmcWord(1, 10, out values.F10_value)) return false;
                if (!ReadPmcBit(0, 4, 3, out values.G4_3)) return false;
                if (!ReadPmcBit(0, 5, 0, out values.G5_0)) return false;
                if (!ReadPmcBit(4, 8, 4, out values.X8_4)) return false;

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

                // 새 근무조로 전환
                _currentShift = newShift;

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
                return MachineState.Alarm;

            // 투입중: 스타트 실행 중 AND 메모리 모드 AND M140 실행 AND (M핀 처리 중 OR G5.0)
            if (pmc.F0_7 && pmc.F3_5 && pmc.F10_value == 140 && (pmc.G4_3 || pmc.G5_0))
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
