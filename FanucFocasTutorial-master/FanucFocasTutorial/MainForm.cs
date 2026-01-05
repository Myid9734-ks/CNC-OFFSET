// ㅇㅇㄹㄴ
using System;
using System.Windows.Forms;
using System.Threading;
using System.Threading.Tasks;
using System.Drawing;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;
using System.Text.RegularExpressions;
// using Connecting.Forms;
using FanucFocasTutorial.Forms;
using System.Linq;
using System.IO;
using System.Xml.Serialization;

namespace FanucFocasTutorial
{
    public partial class MainForm : Form
    {
        private readonly CNCConnectionManager _connectionManager;
        private System.Windows.Forms.Timer _updateTimer;
        private DataGridView _gridTools;
        private TextBox _txtXWear;
        private TextBox _txtZWear;
        private Dictionary<string, bool> _connectionStates;
        private Dictionary<string, Color> _ipGridColors;
        private Dictionary<string, string> _ipAliases; // IP별 별칭 저장
        private Dictionary<string, int> _ipLoadingMCodes; // IP별 로딩 M코드 저장
        private Dictionary<string, IPConfig> _ipConfigs; // IP별 전체 설정 저장
        private ListBox _ipList;
        private Button _registerButton; // 등록/수정 버튼
        private const int DEFAULT_TIMEOUT = 2;
        private CNCConnection _currentConnection;
        private GroupBox _mainGroupBox;
        private DataGridView _macroGridView;
        private TextBox _macroValueInput;
        private TextBox _syncIntervalInput; // 동기화 간격 입력
        private ProgressBar _macroLoadProgress;
        private Label _macroLoadLabel;
        private short _selectedMacroNo = 0; // 선택된 매크로 번호 저장
        private int _syncIntervalSeconds = 60; // 동기화 간격 (초), 기본값 60초

        // PMC 제어 버튼 (LED 스타일)
        private Button _btnBlockSkip;      // Block Skip (R0101.2 버튼 시뮬레이션, R0201.2에서 상태 읽기)
        private Button _btnOptStop;        // Optional Stop (R0101.1 버튼 시뮬레이션, R0201.1에서 상태 읽기)

        // PMC 주소 상수
        private const short PMC_TYPE_F = 1;  // F-type (Signal from CNC)
        private const short PMC_TYPE_G = 0;  // G-type (PMC → CNC)
        private const short PMC_TYPE_R = 5;  // R-type (Internal relay)
        private const short PMC_TYPE_D = 10; // D-type (Data table)
        private const short PMC_TYPE_X = 4;  // X-type (PMC input signal)
        private const short PMC_DATA_BYTE = 0; // Byte type

        // DB 관련
        private MacroDataService _macroDataService;
        private LogDataService _logDataService;
        private System.Windows.Forms.Timer _syncTimer;
        private Dictionary<string, int> _syncIndexPerIp; // IP별 동기화 인덱스
        private bool _isSyncing = false;
        private bool _isMonitoringUpdating = false; // 모니터링 업데이트 중 플래그

        // 랜덤 색상 생성을 위한 Random 객체 추가
        private readonly Random _random = new Random();

        // 수동 옵셋 로그 저장
        private List<ManualOffsetLog> _manualOffsetLogs = new List<ManualOffsetLog>();

        // 매크로 변수 로그 저장
        private List<MacroVariableLog> _macroVariableLogs = new List<MacroVariableLog>();

        // 디버그 로그 파일 경로
        private string _debugLogPath;

        // PMC 제어 로그 파일 경로
        private string _pmcLogPath;

        // 랜덤 파스텔 색상을 생성하는 메서드 추가
        private Color GenerateRandomPastelColor()
        {
            // 파스텔 색상을 위해 기본 밝기값을 높게 설정 (200-255)
            int red = _random.Next(200, 256);
            int green = _random.Next(200, 256);
            int blue = _random.Next(200, 256);
            return Color.FromArgb(red, green, blue);
        }

        // 모니터링 화면 컨트롤
        private GroupBox _monitoringGroup;
        private Label _lblProgramInfo;      // 프로그램 정보
        private Label _lblToolInfo;         // 공구/옵셋 정보
        private Label _lblSpindleInfo;      // 스핀들 정보
        private Label _lblAxisInfo;         // 축 정보
        private Label _lblFeedrateInfo;     // 이송 정보
        private Label _lblMachineTime;      // 가공/운전 시간
        private TableLayoutPanel _monitoringLayout;
        private Label _lblMode;             // 동작 모드
        private Label _lblStatus;           // 상태
        private Label _lblOpSignal;         // 운전 신호
        private Label _lblCurrentTool;      // 현재 공구

        // 중앙 패널 필드 추가
        private Panel _centerPanel;
        private MultiMonitoringForm _backgroundMonitoringForm; // 백그라운드에서 계속 실행되는 모니터링 폼

        // IP 저장 파일 경로
        private readonly string _ipConfigFile = Path.Combine(Application.StartupPath, "ip_config.xml");

        // 설정 저장 파일 경로
        private readonly string _settingsFile = Path.Combine(Application.StartupPath, "settings.xml");

        public MainForm()
        {
            InitializeComponent();
            _connectionManager = new CNCConnectionManager();
            _connectionStates = new Dictionary<string, bool>();
            _ipGridColors = new Dictionary<string, Color>();
            _ipAliases = new Dictionary<string, string>(); // 별칭 딕셔너리 초기화
            _ipLoadingMCodes = new Dictionary<string, int>(); // 로딩 M코드 딕셔너리 초기화
            _ipConfigs = new Dictionary<string, IPConfig>(); // IP별 설정 딕셔너리 초기화
            _macroDataService = new MacroDataService(); // DB 서비스 초기화
            _logDataService = new LogDataService(); // 로그 DB 서비스 초기화
            _syncIndexPerIp = new Dictionary<string, int>(); // IP별 동기화 인덱스 초기화
            _debugLogPath = Path.Combine(Application.StartupPath, "SyncDebugLog.txt"); // 디버그 로그 파일 경로
            _pmcLogPath = Path.Combine(Application.StartupPath, "pmc_control_log.txt"); // PMC 제어 로그 파일 경로

            LoadLogsFromDatabase(); // DB에서 로그 로드
            InitializeControls();
            LoadSettings(); // 설정 로드
            LoadSavedIPs(); // 저장된 IP 로드
            InitializeSyncTimer(); // 동기화 타이머 초기화
        }

        private void InitializeControls()
        {
            // 폼 설정
            this.Text = "Fanuc Tool Offset Monitor";
            this.Size = new System.Drawing.Size(1000, 700);
            this.WindowState = FormWindowState.Maximized;  // 시작 시 최대화
            this.AutoScroll = true;  // 스크롤바 활성화 (작은 화면 대응)
            this.MinimumSize = new Size(1280, 720);  // 최소 크기 설정

            // 중앙 패널 초기화
            _centerPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10)
            };

            // 메인 GroupBox 초기화
            _mainGroupBox = new GroupBox
            {
                Dock = DockStyle.Fill,
                Text = "공구 오프셋 모니터링",
                Font = new Font("맑은 고딕", 12f, FontStyle.Bold),
                Padding = new Padding(10, 10, 20, 10),  // 오른쪽 패딩 증가
                Visible = false  // 초기에는 숨김 상태로 설정
            };

            // 수동 옵셋 화면을 중앙 패널에 추가
            _centerPanel.Controls.Add(_mainGroupBox);

            InitializeMonitoringControls();  // 모니터링 컨트롤 초기화 추가

            // 왼쪽 패널 설정
            Panel leftPanel = new Panel
            {
                Dock = DockStyle.Left,
                Width = 300,
                Padding = new Padding(15)
            };

            // IP 관리 및 메뉴를 위한 통합 GroupBox 생성
            GroupBox combinedGroup = new GroupBox
            {
                Dock = DockStyle.Fill,
                Text = "설비 관리 및 메뉴",
                Font = new Font("맑은 고딕", 12f, FontStyle.Bold),
                Padding = new Padding(10)
            };

            // 전체 내용을 담을 스크롤 가능한 Panel
            Panel mainPanel = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true  // 스크롤 활성화 (낮은 해상도 대응)
            };

            // IP 리스트박스
            _ipList = new ListBox
            {
                Dock = DockStyle.Top,
                Height = 210,  // 7개 항목 표시 (30px * 7 = 210px)
                DrawMode = DrawMode.OwnerDrawFixed,
                ItemHeight = 30,
                Font = new Font("맑은 고딕", 12f),
                Margin = new Padding(0, 0, 0, 5),  // 5픽셀 간격
                TabIndex = 0
            };
            _ipList.DrawItem += IpList_DrawItem;
            _ipList.SelectedIndexChanged += IpList_SelectedIndexChanged;

            // 등록 버튼
            _registerButton = new Button
            {
                Text = "등록",
                Dock = DockStyle.Top,
                Height = 40,
                Font = new Font("맑은 고딕", 12f),
                Margin = new Padding(0, 0, 0, 5),  // 5픽셀 간격
                TabIndex = 4
            };
            _registerButton.Click += RegisterButton_Click;

            // 삭제 버튼
            Button deleteButton = new Button
            {
                Text = "삭제",
                Dock = DockStyle.Top,
                Height = 40,
                Font = new Font("맑은 고딕", 12f),
                Margin = new Padding(0, 0, 0, 5),  // 5픽셀 간격
                TabIndex = 5
            };
            deleteButton.Click += DeleteButton_Click;

            // IP 관리와 메뉴 사이 구분선
            Label separator = new Label
            {
                Text = "━━━━━━━━━━━━━━━━━━━━━━",
                Dock = DockStyle.Top,
                Height = 25,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("맑은 고딕", 10f),
                ForeColor = Color.Gray
            };

            // 왼쪽 패널에 컨트롤 추가
            Label spacer1 = new Label { Height = 5, Dock = DockStyle.Top };
            Label spacer2 = new Label { Height = 5, Dock = DockStyle.Top };
            Label spacer4 = new Label { Height = 10, Dock = DockStyle.Top };

            // 수동 옵셋 버튼
            Button manualOffsetButton = new Button
            {
                Text = "수동 옵셋",
                Dock = DockStyle.Top,
                Height = 40,
                Font = new Font("맑은 고딕", 12f),
                Margin = new Padding(0, 0, 0, 5),
                TabIndex = 7
            };
            manualOffsetButton.Click += (s, e) => {
                _monitoringGroup.Visible = false;  // 모니터링 화면 숨김

                // 기존 자동 옵셋 폼이 있다면 로깅 중지 후 제거
                var existingAutoOffsetForm = _centerPanel.Controls.OfType<AutoOffsetForm>().FirstOrDefault();
                if (existingAutoOffsetForm != null)
                {
                    existingAutoOffsetForm.StopAutoProcess(); // 로깅 중지
                    _centerPanel.Controls.Remove(existingAutoOffsetForm);
                    existingAutoOffsetForm.Dispose();
                }

                _mainGroupBox.Text = "공구 오프셋";  // GroupBox 텍스트 변경
                _mainGroupBox.Visible = true;  // 수동 옵셋 화면 표시
                _mainGroupBox.BringToFront();

                // 매크로 값 백그라운드 로드
                LoadAllMacroValuesAsync();

                // PMC 상태 새로고침
                RefreshPmcStates();
            };

            // 자동 옵셋 버튼
            Button autoOffsetButton = new Button
            {
                Text = "자동 옵셋",
                Dock = DockStyle.Top,
                Height = 40,
                Font = new Font("맑은 고딕", 12f),
                Margin = new Padding(0, 0, 0, 5),
                TabIndex = 8
            };
            autoOffsetButton.Click += (s, e) => {
                _mainGroupBox.Visible = false;  // 수동 옵셋 화면 숨김
                _monitoringGroup.Visible = false;  // 모니터링 화면 숨김

                // 기존 자동 옵셋 폼이 있다면 제거
                var existingForm = _centerPanel.Controls.OfType<AutoOffsetForm>().FirstOrDefault();
                if (existingForm != null)
                {
                    _centerPanel.Controls.Remove(existingForm);
                    existingForm.Dispose();
                }
                
                // 선택된 IP 주소 가져오기
                string selectedIp = GetSelectedIpAddress();
                if (string.IsNullOrEmpty(selectedIp))
                {
                    MessageBox.Show("연결할 설비를 선택해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var autoOffsetForm = new AutoOffsetForm(this, _connectionManager, selectedIp);
                autoOffsetForm.TopLevel = false;
                autoOffsetForm.FormBorderStyle = FormBorderStyle.None;
                autoOffsetForm.Dock = DockStyle.Fill;
                
                _centerPanel.Controls.Add(autoOffsetForm);
                autoOffsetForm.Show();
                autoOffsetForm.BringToFront();
                
                // 자동시작 버튼 자동 클릭
                autoOffsetForm.StartAutoProcess();
            };


            // 자동 모니터링 버튼
            Button autoMonitoringButton = new Button
            {
                Text = "자동 모니터링",
                Dock = DockStyle.Top,
                Height = 40,
                Font = new Font("맑은 고딕", 12f),
                Margin = new Padding(0, 0, 0, 5),
                TabIndex = 9
            };
            autoMonitoringButton.Click += (s, e) => {
                // AutoOffsetForm, AutoMonitoringForm은 제거
                var existingAutoOffsetForm = _centerPanel.Controls.OfType<AutoOffsetForm>().FirstOrDefault();
                if (existingAutoOffsetForm != null)
                {
                    existingAutoOffsetForm.StopAutoProcess();
                    _centerPanel.Controls.Remove(existingAutoOffsetForm);
                    existingAutoOffsetForm.Dispose();
                }

                var existingAutoMonitoringForm = _centerPanel.Controls.OfType<AutoMonitoringForm>().FirstOrDefault();
                if (existingAutoMonitoringForm != null)
                {
                    _centerPanel.Controls.Remove(existingAutoMonitoringForm);
                    existingAutoMonitoringForm.Dispose();
                }

                // 다른 폼들 숨기기
                var existingShiftDataForm = _centerPanel.Controls.OfType<ShiftDataViewForm>().FirstOrDefault();
                if (existingShiftDataForm != null)
                {
                    _centerPanel.Controls.Remove(existingShiftDataForm);
                }

                _mainGroupBox.Visible = false;
                _monitoringGroup.Visible = false;

                // IP 관리 목록 순서대로 설비 수집
                var allConnections = new List<CNCConnection>();
                var connectionDict = _connectionManager.GetAllConnections().ToDictionary(c => c.IpAddress);

                foreach (var item in _ipList.Items)
                {
                    string ipPort = item.ToString();
                    string ip = ipPort.Split(':')[0];

                    if (connectionDict.ContainsKey(ip))
                    {
                        allConnections.Add(connectionDict[ip]);
                    }
                }

                if (allConnections.Count == 0)
                {
                    MessageBox.Show("등록된 설비가 없습니다.\nIP 관리에서 설비를 먼저 등록하세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // 백그라운드 모니터링 폼 재사용 또는 새로 생성
                if (_backgroundMonitoringForm == null || _backgroundMonitoringForm.IsDisposed)
                {
                    _backgroundMonitoringForm = new MultiMonitoringForm(allConnections, _ipConfigs);
                    _backgroundMonitoringForm.TopLevel = false;
                    _backgroundMonitoringForm.FormBorderStyle = FormBorderStyle.None;
                    _backgroundMonitoringForm.Dock = DockStyle.Fill;
                }

                // 화면에 표시
                if (!_centerPanel.Controls.Contains(_backgroundMonitoringForm))
                {
                    _centerPanel.Controls.Add(_backgroundMonitoringForm);
                }
                _backgroundMonitoringForm.Show();
                _backgroundMonitoringForm.BringToFront();
            };

            // 근무조 현황 버튼
            Button shiftDataButton = new Button
            {
                Text = "근무조 현황",
                Dock = DockStyle.Top,
                Height = 40,
                Font = new Font("맑은 고딕", 12f),
                Margin = new Padding(0, 0, 0, 5),
                TabIndex = 10
            };
            shiftDataButton.Click += (s, e) => {
                // AutoOffsetForm, AutoMonitoringForm은 제거
                var existingAutoOffsetForm = _centerPanel.Controls.OfType<AutoOffsetForm>().FirstOrDefault();
                if (existingAutoOffsetForm != null)
                {
                    existingAutoOffsetForm.StopAutoProcess();
                    _centerPanel.Controls.Remove(existingAutoOffsetForm);
                    existingAutoOffsetForm.Dispose();
                }

                var existingAutoMonitoringForm = _centerPanel.Controls.OfType<AutoMonitoringForm>().FirstOrDefault();
                if (existingAutoMonitoringForm != null)
                {
                    _centerPanel.Controls.Remove(existingAutoMonitoringForm);
                    existingAutoMonitoringForm.Dispose();
                }

                // MultiMonitoringForm은 백그라운드에서 유지 (화면에서만 제거)
                if (_backgroundMonitoringForm != null && _centerPanel.Controls.Contains(_backgroundMonitoringForm))
                {
                    _centerPanel.Controls.Remove(_backgroundMonitoringForm);
                }

                var existingShiftDataForm = _centerPanel.Controls.OfType<ShiftDataViewForm>().FirstOrDefault();
                if (existingShiftDataForm != null)
                {
                    _centerPanel.Controls.Remove(existingShiftDataForm);
                    existingShiftDataForm.Dispose();
                }

                _mainGroupBox.Visible = false;
                _monitoringGroup.Visible = false;

                // 근무조 현황 폼을 메인 화면에 embedded로 표시
                var shiftDataForm = new ShiftDataViewForm();
                shiftDataForm.TopLevel = false;
                shiftDataForm.FormBorderStyle = FormBorderStyle.None;
                shiftDataForm.Dock = DockStyle.Fill;

                _centerPanel.Controls.Add(shiftDataForm);
                shiftDataForm.Show();
                shiftDataForm.BringToFront();
            };

            // 로그기록 버튼
            Button logRecordButton = new Button
            {
                Text = "로그기록",
                Dock = DockStyle.Top,
                Height = 40,
                Font = new Font("맑은 고딕", 12f),
                Margin = new Padding(0, 0, 0, 5),
                TabIndex = 11
            };
            logRecordButton.Click += (s, e) => {
                // AutoOffsetForm, AutoMonitoringForm은 제거
                var existingAutoOffsetForm = _centerPanel.Controls.OfType<AutoOffsetForm>().FirstOrDefault();
                if (existingAutoOffsetForm != null)
                {
                    existingAutoOffsetForm.StopAutoProcess();
                    _centerPanel.Controls.Remove(existingAutoOffsetForm);
                    existingAutoOffsetForm.Dispose();
                }

                var existingMonitoringForm = _centerPanel.Controls.OfType<AutoMonitoringForm>().FirstOrDefault();
                if (existingMonitoringForm != null)
                {
                    _centerPanel.Controls.Remove(existingMonitoringForm);
                    existingMonitoringForm.Dispose();
                }

                // MultiMonitoringForm은 백그라운드에서 유지 (화면에서만 제거)
                if (_backgroundMonitoringForm != null && _centerPanel.Controls.Contains(_backgroundMonitoringForm))
                {
                    _centerPanel.Controls.Remove(_backgroundMonitoringForm);
                }

                _mainGroupBox.Visible = false;
                _monitoringGroup.Visible = false;

                // 로그 뷰어 표시
                ShowLogViewer();
            };

            // 모든 컨트롤을 mainPanel에 추가 (역순으로 추가 - Dock.Top이므로)
            mainPanel.Controls.AddRange(new Control[] {
                logRecordButton,
                shiftDataButton,
                autoMonitoringButton,
                autoOffsetButton,
                manualOffsetButton,
                spacer4,
                separator,
                deleteButton,
                spacer2,
                _registerButton,
                spacer1,
                _ipList
            });

            combinedGroup.Controls.Add(mainPanel);
            leftPanel.Controls.Add(combinedGroup);

            // 오른쪽 패널 설정 수정
            Panel rightPanel = new Panel
            {
                Dock = DockStyle.Right,
                Width = 250,  // 너비 증가
                Padding = new Padding(5)
            };

            // 데이터그리드 패널 설정
            Panel toolGridPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(0, 0, 10, 0)  // 오른쪽 여백 조정
            };

            // 데이터그리드 설정
            _gridTools = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AllowUserToResizeColumns = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                EnableHeadersVisualStyles = false,
                BackgroundColor = Color.White,
                GridColor = Color.LightGray,
                BorderStyle = BorderStyle.None
            };

            // 헤더 스타일 설정
            _gridTools.ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = SystemColors.Control,
                ForeColor = SystemColors.ControlText,
                SelectionBackColor = SystemColors.Control,
                SelectionForeColor = SystemColors.ControlText,
                Font = new Font("맑은 고딕", 11f, FontStyle.Regular),
                Alignment = DataGridViewContentAlignment.MiddleCenter  // 가운데 정렬
            };

            // 셀 스타일 설정
            _gridTools.DefaultCellStyle = new DataGridViewCellStyle
            {
                Font = new Font("맑은 고딕", 11f),
                Alignment = DataGridViewContentAlignment.MiddleRight,  // 오른쪽 정렬
                SelectionBackColor = Color.FromArgb(51, 153, 255),
                SelectionForeColor = Color.White
            };

            // 데이터그리드 컬럼 추가
            DataGridViewTextBoxColumn toolColumn = new DataGridViewTextBoxColumn
            {
                Name = "Tool",
                HeaderText = "공구 번호",
                FillWeight = 15,
                MinimumWidth = 80
            };
            _gridTools.Columns.Add(toolColumn);

            DataGridViewTextBoxColumn xGeometryColumn = new DataGridViewTextBoxColumn
            {
                Name = "XGeometry",
                HeaderText = "X축 형상",
                FillWeight = 20,
                MinimumWidth = 100
            };
            _gridTools.Columns.Add(xGeometryColumn);

            DataGridViewTextBoxColumn xWearColumn = new DataGridViewTextBoxColumn
            {
                Name = "XWear",
                HeaderText = "X축 마모",
                FillWeight = 20,
                MinimumWidth = 100
            };
            _gridTools.Columns.Add(xWearColumn);

            DataGridViewTextBoxColumn zGeometryColumn = new DataGridViewTextBoxColumn
            {
                Name = "ZGeometry",
                HeaderText = "Z축 형상",
                FillWeight = 20,
                MinimumWidth = 100
            };
            _gridTools.Columns.Add(zGeometryColumn);

            DataGridViewTextBoxColumn zWearColumn = new DataGridViewTextBoxColumn
            {
                Name = "ZWear",
                HeaderText = "Z축 마모",
                FillWeight = 25,  // 더 큰 비중으로 설정
                MinimumWidth = 100
            };
            _gridTools.Columns.Add(zWearColumn);

            // 이벤트 핸들러 연결
            _gridTools.SelectionChanged += GridTools_SelectionChanged;

            // X축 마모 입력
            Label lblXWearInput = new Label
            { 
                Text = "X축 마모 입력:", 
                AutoSize = true, 
                Dock = DockStyle.Top
            };

            _txtXWear = new TextBox
            {
                Dock = DockStyle.Top,
                Margin = new Padding(0, 5, 0, 5),
                Height = 40,
                Font = new Font("맑은 고딕", 12f),
                Name = "txtXWear"  // 이름 명시적 설정
            };
            _txtXWear.GotFocus += (s, e) => _updateTimer?.Stop();
            _txtXWear.LostFocus += (s, e) => _updateTimer?.Start();
            _txtXWear.KeyDown += (s, e) => {
                if (e.KeyCode == Keys.Enter)
                {
                    e.Handled = true;
                    e.SuppressKeyPress = true;  // 비프음 방지
                    UpdateWearValue(0);
                    _updateTimer?.Start();
                }
            };

            Button btnXWear = new Button
            { 
                Text = "X축 입력",
                Dock = DockStyle.Top,
                Height = 40,
                Font = new Font("맑은 고딕", 12f),
                Name = "btnXWear"  // 이름 명시적 설정
            };
            btnXWear.Click += (s, e) => {
                UpdateWearValue(0);
                _updateTimer?.Start();
            };

            // 공백용 패널
            Panel spacerPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 40
            };

            // Z축 마모 입력
            Label lblZWearInput = new Label
            { 
                Text = "Z축 마모 입력:", 
                AutoSize = true, 
                Dock = DockStyle.Top
            };

            _txtZWear = new TextBox
            {
                Dock = DockStyle.Top,
                Margin = new Padding(0, 5, 0, 5),
                Height = 40,
                Font = new Font("맑은 고딕", 12f),
                Name = "txtZWear"  // 이름 명시적 설정
            };
            _txtZWear.GotFocus += (s, e) => _updateTimer?.Stop();
            _txtZWear.LostFocus += (s, e) => _updateTimer?.Start();
            _txtZWear.KeyDown += (s, e) => {
                if (e.KeyCode == Keys.Enter)
                {
                    e.Handled = true;
                    e.SuppressKeyPress = true;  // 비프음 방지
                    UpdateWearValue(1);
                    _updateTimer?.Start();
                }
            };

            Button btnZWear = new Button
            {
                Text = "Z축 입력",
                Dock = DockStyle.Top,
                Height = 40,
                Font = new Font("맑은 고딕", 12f),
                Name = "btnZWear"  // 이름 명시적 설정
            };
            btnZWear.Click += (s, e) => {
                UpdateWearValue(1);
                _updateTimer?.Start();
            };

            // PMC 제어 버튼 (LED 스타일)
            _btnBlockSkip = new Button
            {
                Text = "● Block Skip",
                Dock = DockStyle.Top,
                Height = 40,
                Font = new Font("맑은 고딕", 11f, FontStyle.Bold),
                Name = "btnBlockSkip",
                Margin = new Padding(5, 5, 5, 5),
                FlatStyle = FlatStyle.Flat,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(10, 0, 0, 0),
                BackColor = Color.FromArgb(240, 240, 240),
                ForeColor = Color.Gray,
                Cursor = Cursors.Hand
            };
            _btnBlockSkip.FlatAppearance.BorderColor = Color.FromArgb(200, 200, 200);
            _btnBlockSkip.FlatAppearance.BorderSize = 1;
            _btnBlockSkip.Click += ChkBlockSkip_Click;

            _btnOptStop = new Button
            {
                Text = "● Optional Stop",
                Dock = DockStyle.Top,
                Height = 40,
                Font = new Font("맑은 고딕", 11f, FontStyle.Bold),
                Name = "btnOptStop",
                Margin = new Padding(5, 5, 5, 5),
                FlatStyle = FlatStyle.Flat,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(10, 0, 0, 0),
                BackColor = Color.FromArgb(240, 240, 240),
                ForeColor = Color.Gray,
                Cursor = Cursors.Hand
            };
            _btnOptStop.FlatAppearance.BorderColor = Color.FromArgb(200, 200, 200);
            _btnOptStop.FlatAppearance.BorderSize = 1;
            _btnOptStop.Click += ChkOptStop_Click;

            // PMC 버튼 구분용 간격 패널 (Z축과 PMC 버튼 사이)
            Panel pmcSpacerPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 20,
                BackColor = Color.Transparent
            };

            // PMC 버튼 간 간격 패널 (Block Skip과 Optional Stop 사이)
            Panel pmcButtonSpacerPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 5,
                BackColor = Color.Transparent
            };

            // X축 입력 간격 패널 (텍스트박스와 버튼 사이)
            Panel xWearSpacerPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 5,
                BackColor = Color.Transparent
            };

            // Z축 입력 간격 패널 (텍스트박스와 버튼 사이)
            Panel zWearSpacerPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 5,
                BackColor = Color.Transparent
            };

            // 매크로 변수 쓰기 버튼
            Button btnMacroWrite = new Button
            {
                Text = "매크로 쓰기",
                Dock = DockStyle.Bottom,
                Height = 40,
                Font = new Font("맑은 고딕", 12f),
                Name = "btnMacroWrite",
                Margin = new Padding(0, 0, 0, 10)
            };
            btnMacroWrite.Click += BtnMacroWrite_Click;

            // 매크로 값 입력
            _macroValueInput = new TextBox
            {
                Dock = DockStyle.Bottom,
                Margin = new Padding(0, 5, 0, 15),
                Height = 40,
                Font = new Font("맑은 고딕", 12f),
                Name = "txtMacroValue"
            };
            _macroValueInput.KeyDown += (s, e) => {
                if (e.KeyCode == Keys.Enter)
                {
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                    BtnMacroWrite_Click(s, e);
                }
            };

            // 매크로 값 입력 레이블
            Label lblMacroValueInput = new Label
            {
                Text = "매크로 값 입력:",
                AutoSize = true,
                Dock = DockStyle.Bottom,
                Font = new Font("맑은 고딕", 10f)
            };

            // 동기화 간격 설정 버튼
            Button btnApplySyncInterval = new Button
            {
                Text = "적용",
                Dock = DockStyle.Bottom,
                Height = 35,
                Font = new Font("맑은 고딕", 10f),
                Name = "btnApplySyncInterval",
                Margin = new Padding(0, 0, 0, 10)
            };
            btnApplySyncInterval.Click += BtnApplySyncInterval_Click;

            // 동기화 간격 입력
            _syncIntervalInput = new TextBox
            {
                Dock = DockStyle.Bottom,
                Margin = new Padding(0, 5, 0, 5),
                Height = 35,
                Font = new Font("맑은 고딕", 11f),
                Name = "txtSyncInterval",
                Text = "60"
            };
            _syncIntervalInput.KeyDown += (s, e) => {
                if (e.KeyCode == Keys.Enter)
                {
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                    BtnApplySyncInterval_Click(s, e);
                }
            };

            // 동기화 간격 레이블
            Label lblSyncInterval = new Label
            {
                Text = "동기화 간격 (초):",
                AutoSize = true,
                Dock = DockStyle.Bottom,
                Font = new Font("맑은 고딕", 10f),
                Margin = new Padding(0, 10, 0, 0)
            };

            // 매크로 섹션 제목
            Label lblMacroSection = new Label
            {
                Text = "=== 매크로 변수 ===",
                AutoSize = true,
                Dock = DockStyle.Bottom,
                Font = new Font("맑은 고딕", 11f, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter,
                Padding = new Padding(0, 10, 0, 10)
            };

            // 오른쪽 패널에 컨트롤 추가 (Dock.Top은 나중 추가가 위에 표시되므로 역순)
            rightPanel.Controls.AddRange(new Control[] {
                _btnOptStop,            // 13. Optional Stop 버튼 (가장 아래)
                pmcButtonSpacerPanel,   // 12. PMC 버튼 간 간격
                _btnBlockSkip,          // 11. Block Skip 버튼
                pmcSpacerPanel,         // 10. PMC 버튼 구분 간격 (Z축과의 간격)
                btnZWear,               // 9. Z축 입력 버튼
                zWearSpacerPanel,       // 8. Z축 입력 간격
                _txtZWear,              // 7. Z축 텍스트박스
                lblZWearInput,          // 6. Z축 레이블
                spacerPanel,            // 5. 구분선
                btnXWear,               // 4. X축 입력 버튼
                xWearSpacerPanel,       // 3. X축 입력 간격
                _txtXWear,              // 2. X축 텍스트박스
                lblXWearInput           // 1. X축 레이블 (가장 위)
            });

            // 매크로 컨트롤 추가 (Bottom이므로 역순)
            rightPanel.Controls.Add(lblMacroSection);
            rightPanel.Controls.Add(lblSyncInterval);
            rightPanel.Controls.Add(_syncIntervalInput);
            rightPanel.Controls.Add(btnApplySyncInterval);
            rightPanel.Controls.Add(lblMacroValueInput);
            rightPanel.Controls.Add(_macroValueInput);
            rightPanel.Controls.Add(btnMacroWrite);

            // 상태 패널 설정
            Panel statusPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 40,
                BackColor = Color.WhiteSmoke,
                Padding = new Padding(10, 5, 10, 5)
            };

            // 상태 정보 레이블들
            Label lblMode = new Label
            {
                Text = "동작 모드: -",
                AutoSize = true,
                Font = new Font("맑은 고딕", 10f),
                Margin = new Padding(0, 0, 20, 0)
            };

            Label lblStatus = new Label
            {
                Text = "상태: -",
                AutoSize = true,
                Font = new Font("맑은 고딕", 10f),
                Margin = new Padding(0, 0, 20, 0)
            };

            Label lblOpSignal = new Label
            {
                Text = "운전 신호: -",
                AutoSize = true,
                Font = new Font("맑은 고딕", 10f),
                Margin = new Padding(0, 0, 20, 0)
            };

            Label lblCurrentTool = new Label
            {
                Text = "선택된 공구: -",
                AutoSize = true,
                Font = new Font("맑은 고딕", 10f)
            };

            // 상태 패널에 레이블 추가
            FlowLayoutPanel statusFlow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false
            };
            statusFlow.Controls.AddRange(new Control[] { lblMode, lblStatus, lblOpSignal, lblCurrentTool });
            statusPanel.Controls.Add(statusFlow);

            // 필드에 레이블 할당
            _lblMode = lblMode;
            _lblStatus = lblStatus;
            _lblOpSignal = lblOpSignal;
            _lblCurrentTool = lblCurrentTool;

            // 데이터그리드뷰뷰를 담을 패널 생성
            toolGridPanel.Controls.Add(_gridTools);

            // 매크로 데이터를 표시할 GroupBox 생성
            GroupBox macroGroup = new GroupBox
            {
                Dock = DockStyle.Bottom,
                Height = 250,
                Text = "=== 매크로 변수 ===",
                Font = new Font("맑은 고딕", 12f, FontStyle.Bold),
                Padding = new Padding(5)
            };

            // 매크로 데이터용 DataGridView 생성
            _macroGridView = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true, // 읽기 전용으로 변경
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                EnableHeadersVisualStyles = false,
                BackgroundColor = Color.White,
                GridColor = Color.LightGray,
                BorderStyle = BorderStyle.None,
                ColumnHeadersHeight = 30,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
                ScrollBars = ScrollBars.Vertical, // 세로 스크롤바 활성화
                RowHeadersVisible = false
            };

            // 매크로 그리드뷰 스타일 설정
            _macroGridView.ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = SystemColors.Control,
                ForeColor = SystemColors.ControlText,
                SelectionBackColor = SystemColors.Control,
                SelectionForeColor = SystemColors.ControlText,
                Font = new Font("맑은 고딕", 11f, FontStyle.Regular),
                Alignment = DataGridViewContentAlignment.MiddleCenter
            };

            _macroGridView.DefaultCellStyle = new DataGridViewCellStyle
            {
                Font = new Font("맑은 고딕", 11f),
                Alignment = DataGridViewContentAlignment.MiddleRight,
                SelectionBackColor = Color.FromArgb(51, 153, 255),
                SelectionForeColor = Color.White
            };

            // 매크로 그리드뷰 컬럼 추가
            DataGridViewTextBoxColumn macroNoColumn = new DataGridViewTextBoxColumn
            {
                Name = "MacroNo",
                HeaderText = "매크로 번호",
                FillWeight = 40,
                MinimumWidth = 120,
                ReadOnly = true, // 매크로 번호는 편집 불가
                SortMode = DataGridViewColumnSortMode.NotSortable // 정렬 비활성화
            };
            _macroGridView.Columns.Add(macroNoColumn);

            DataGridViewTextBoxColumn macroValueColumn = new DataGridViewTextBoxColumn
            {
                Name = "MacroValue",
                HeaderText = "데이터",
                FillWeight = 60,
                MinimumWidth = 150,
                ReadOnly = true, // 읽기 전용
                SortMode = DataGridViewColumnSortMode.NotSortable // 정렬 비활성화
            };
            _macroGridView.Columns.Add(macroValueColumn);

            // #500부터 #999까지 행 생성 (500개)
            for (int i = 500; i <= 999; i++)
            {
                _macroGridView.Rows.Add($"#{i}", "");
            }

            // 이벤트 핸들러 추가
            _macroGridView.CellClick += MacroGridView_CellClick; // 클릭 시 매크로 선택

            // 진행률 표시 패널 생성
            Panel progressPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 50,
                Padding = new Padding(5)
            };

            // 진행률 레이블
            _macroLoadLabel = new Label
            {
                Dock = DockStyle.Top,
                Height = 20,
                Text = "",
                Font = new Font("맑은 고딕", 9f),
                TextAlign = ContentAlignment.MiddleCenter,
                Visible = false
            };

            // 진행률 바
            _macroLoadProgress = new ProgressBar
            {
                Dock = DockStyle.Top,
                Height = 25,
                Minimum = 0,
                Maximum = 500,
                Value = 0,
                Visible = false
            };

            progressPanel.Controls.Add(_macroLoadLabel);
            progressPanel.Controls.Add(_macroLoadProgress);

            // 매크로 GroupBox에 컨트롤 추가
            macroGroup.Controls.Add(_macroGridView);
            macroGroup.Controls.Add(progressPanel);

            // 패널에 컨트롤 추가
            toolGridPanel.Controls.Add(macroGroup);

            // 패널에 컨트롤 추가 순서 변경
            _mainGroupBox.Controls.Add(toolGridPanel);
            _mainGroupBox.Controls.Add(rightPanel);

            // 메인 폼에 컨트롤 추가
            this.Controls.AddRange(new Control[] { _centerPanel, leftPanel });

            // 타이머 설정
            _updateTimer = new System.Windows.Forms.Timer
            {
                Interval = 10000 // 10초 간격으로 수정 (UI 지연 개선)
            };
            _updateTimer.Tick += UpdateTimer_Tick;

            // 데이터그리드 로드 후 크기 조정 - 수정된 방법
            this.Load += (s, e) => {
                // 폼이 완전히 로드된 후 타이머를 사용하여 지연 실행
                System.Windows.Forms.Timer resizeTimer = new System.Windows.Forms.Timer();
                resizeTimer.Interval = 100; // 100ms 후 실행
                resizeTimer.Tick += (st, et) => {
                    if (toolGridPanel.Width > 0) {
                        _gridTools.Width = toolGridPanel.Width - 30;
                        resizeTimer.Stop();
                        resizeTimer.Dispose();
                    }
                };
                resizeTimer.Start();
            };

            // 레이아웃 이벤트 사용
            toolGridPanel.Layout += (s, e) => {
                if (toolGridPanel.Width > 0) {
                    _gridTools.Width = toolGridPanel.Width - 30;
                }
            };
        }

        private void InitializeMonitoringControls()
        {
            // 모니터링 그룹 생성
            _monitoringGroup = new GroupBox
            {
                Dock = DockStyle.Fill,
                Text = "장비 모니터링",
                Font = new Font("맑은 고딕", 12f, FontStyle.Bold),
                Padding = new Padding(10),
                Visible = false
            };

            // 테이블 레이아웃 설정
            _monitoringLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 6,
                Padding = new Padding(5),
                CellBorderStyle = TableLayoutPanelCellBorderStyle.Single
            };

            // 컬럼 설정
            _monitoringLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            _monitoringLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));

            // 모니터링 레이블 초기화
            _lblProgramInfo = CreateMonitoringLabel("프로그램 정보:\nO0000\nN0000");
            _lblToolInfo = CreateMonitoringLabel("공구/옵셋 정보:\nT00 D00\n공구 수명: 0%");
            _lblSpindleInfo = CreateMonitoringLabel("스핀들 정보:\n회전수: 0 RPM\n부하: 0%");
            _lblAxisInfo = CreateMonitoringLabel("축 정보:\nX: 0.000\nZ: 0.000");
            _lblFeedrateInfo = CreateMonitoringLabel("이송 정보:\n실속도: 0 mm/min\n지령: 0 mm/min");
            _lblMachineTime = CreateMonitoringLabel("시간 정보:\n절삭: 00:00:00\n운전: 00:00:00");

            // 레이아웃에 컨트롤 추가
            _monitoringLayout.Controls.Add(_lblProgramInfo, 0, 0);
            _monitoringLayout.Controls.Add(_lblToolInfo, 1, 0);
            _monitoringLayout.Controls.Add(_lblSpindleInfo, 0, 1);
            _monitoringLayout.Controls.Add(_lblAxisInfo, 1, 1);
            _monitoringLayout.Controls.Add(_lblFeedrateInfo, 0, 2);
            _monitoringLayout.Controls.Add(_lblMachineTime, 1, 2);

            // 모니터링 그룹에 레이아웃 추가
            _monitoringGroup.Controls.Add(_monitoringLayout);

            // 메인 폼에 모니터링 그룹 추가
            this.Controls.Add(_monitoringGroup);
        }

        private Label CreateMonitoringLabel(string text)
        {
            return new Label
            {
                Text = text,
                Dock = DockStyle.Fill,
                Font = new Font("맑은 고딕", 11f),
                TextAlign = ContentAlignment.MiddleCenter,
                AutoSize = false,
                BackColor = Color.White
            };
        }

        // 설비 상태 모니터링 UI 초기화
        private void IpList_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;

            e.DrawBackground();

            string item = _ipList.Items[e.Index].ToString();
            string ip = item.Split(':')[0];

            // 별칭이 있으면 표시할 텍스트 구성
            string displayText = item;
            if (_ipAliases.ContainsKey(ip) && !string.IsNullOrEmpty(_ipAliases[ip]))
            {
                displayText = $"[{_ipAliases[ip]}] {item}";
            }

            // 상태 아이콘 그리기
            using (var brush = new SolidBrush(Color.Black))  // 텍스트 색상을 검은색으로 변경
            using (var circleBrush = new SolidBrush(_connectionStates.ContainsKey(ip) && _connectionStates[ip] ? Color.Green : Color.Red))
            {
                // 배경을 흰색으로 채우기
                using (var backBrush = new SolidBrush(Color.White))
                {
                    e.Graphics.FillRectangle(backBrush, e.Bounds);
                }

                // 선택된 항목의 배경색 처리
                if ((e.State & DrawItemState.Selected) == DrawItemState.Selected)
                {
                    using (var selectedBrush = new SolidBrush(SystemColors.Highlight))
                    {
                        e.Graphics.FillRectangle(selectedBrush, e.Bounds);
                        brush.Color = SystemColors.HighlightText;  // 선택된 항목의 텍스트는 흰색으로
                    }
                }

                // 텍스트 그리기 (별칭 포함)
                e.Graphics.DrawString(displayText, e.Font, brush, e.Bounds.Left + 25, e.Bounds.Top + 2);

                // 원형 상태 표시등 그리기
                e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                e.Graphics.FillEllipse(circleBrush, e.Bounds.Left + 5, e.Bounds.Top + 5, 10, 10);
            }

            e.DrawFocusRectangle();
        }

        private void AddDefaultConnection()
        {
            // 기본 연결 설정 제거 - IP 관리에서 수동으로 추가하도록 변경
        }

        // IP 설정 클래스
        [Serializable]
        public class IPConfig
        {
            public string IpAddress { get; set; }
            public ushort Port { get; set; }
            public string ColorHex { get; set; }
            public string Alias { get; set; } // 설비 별칭 추가
            public int LoadingMCode { get; set; } // 로딩 M코드 추가

            // PMC 상태 모니터링 어드레스
            public string PmcF0_0 { get; set; } = "F0.0";  // ST_OP 가동 신호
            public string PmcF0_7 { get; set; } = "F0.7";  // 스타트 실행 중
            public string PmcF1_0 { get; set; } = "F1.0";  // 알람 신호
            public string PmcF3_5 { get; set; } = "F3.5";  // 메모리 모드
            public string PmcF10 { get; set; } = "F10";    // M코드 번호
            public string PmcG4_3 { get; set; } = "G4.3";  // M핀 처리 중
            public string PmcG5_0 { get; set; } = "G5.0";  // 추가 투입 조건
            public string PmcX8_4 { get; set; } = "X8.4";  // 비상정지 신호
            public string PmcR854_2 { get; set; } = "R854.2"; // 알람 상태 PMC-A

            // PMC 제어 어드레스 - Block Skip
            public string BlockSkipInputAddr { get; set; } = "R101.2";  // 버튼 입력
            public string BlockSkipStateAddr { get; set; } = "R201.2";  // 상태 확인

            // PMC 제어 어드레스 - Optional Stop
            public string OptStopInputAddr { get; set; } = "R101.1";  // 버튼 입력
            public string OptStopStateAddr { get; set; } = "R201.1";  // 상태 확인

            public IPConfig() { } // XML 직렬화를 위한 기본 생성자
        }

        [Serializable]
        public class IPConfigList
        {
            public List<IPConfig> Configs { get; set; }

            public IPConfigList()
            {
                Configs = new List<IPConfig>();
            }
        }

        // 애플리케이션 설정 클래스
        [Serializable]
        public class AppSettings
        {
            public int SyncIntervalSeconds { get; set; }

            public AppSettings()
            {
                SyncIntervalSeconds = 60; // 기본값 60초
            }
        }

        // 저장된 IP 로드
        private void LoadSavedIPs()
        {
            try
            {
                if (File.Exists(_ipConfigFile))
                {
                    var serializer = new XmlSerializer(typeof(IPConfigList));
                    IPConfigList configList;

                    using (var reader = new FileStream(_ipConfigFile, FileMode.Open))
                    {
                        configList = (IPConfigList)serializer.Deserialize(reader);
                    }

                    foreach (var config in configList.Configs)
                    {
                        if (_connectionManager.AddConnection(config.IpAddress, config.Port, DEFAULT_TIMEOUT))
                        {
                            _ipList.Items.Add($"{config.IpAddress}:{config.Port}");
                            _connectionStates[config.IpAddress] = false;

                            // 저장된 색상 복원
                            if (!string.IsNullOrEmpty(config.ColorHex))
                            {
                                try
                                {
                                    _ipGridColors[config.IpAddress] = ColorTranslator.FromHtml(config.ColorHex);
                                }
                                catch
                                {
                                    _ipGridColors[config.IpAddress] = GenerateRandomPastelColor();
                                }
                            }
                            else
                            {
                                _ipGridColors[config.IpAddress] = GenerateRandomPastelColor();
                            }

                            // 저장된 별칭 복원
                            if (!string.IsNullOrEmpty(config.Alias))
                            {
                                _ipAliases[config.IpAddress] = config.Alias;
                            }

                            // 저장된 M코드 복원 (0 포함 - 0은 자동화 없음을 의미)
                            _ipLoadingMCodes[config.IpAddress] = config.LoadingMCode;

                            // 전체 설정 복원
                            _ipConfigs[config.IpAddress] = config;
                        }
                    }

                    if (configList.Configs.Count > 0)
                    {
                        // 자동으로 모든 IP 연결 시도
                        ConnectAllSavedIPs();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"IP 설정 로드 중 오류 발생: {ex.Message}", "로드 오류",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // 저장된 IP들 자동 연결
        private void ConnectAllSavedIPs()
        {
            foreach (var connection in _connectionManager.GetAllConnections())
            {
                try
                {
                    if (connection.Connect())
                    {
                        _connectionStates[connection.IpAddress] = true;

                        // 첫 번째 성공한 연결을 현재 연결로 설정
                        if (_currentConnection == null)
                        {
                            _currentConnection = connection;
                            UpdateMachineStatus(connection.IpAddress);
                        }
                    }
                    else
                    {
                        _connectionStates[connection.IpAddress] = false;
                    }
                }
                catch
                {
                    _connectionStates[connection.IpAddress] = false;
                }
            }

            // 연결된 IP가 있으면 타이머 시작
            if (_connectionStates.Values.Any(connected => connected))
            {
                _updateTimer.Start();
            }

            _ipList.Invalidate();
        }

        // IP 설정 저장
        private void SaveIPsToFile()
        {
            try
            {
                var configList = new IPConfigList();

                foreach (string item in _ipList.Items)
                {
                    string[] parts = item.Split(':');
                    if (parts.Length == 2)
                    {
                        string ip = parts[0];

                        // _ipConfigs에서 전체 설정 가져오기
                        if (_ipConfigs.ContainsKey(ip))
                        {
                            var config = _ipConfigs[ip];
                            // 색상 정보만 추가로 설정
                            config.ColorHex = _ipGridColors.ContainsKey(ip) ?
                                ColorTranslator.ToHtml(_ipGridColors[ip]) : "";
                            configList.Configs.Add(config);
                        }
                        else
                        {
                            // _ipConfigs에 없는 경우 (하위 호환성)
                            ushort port = ushort.Parse(parts[1]);
                            string colorHex = _ipGridColors.ContainsKey(ip) ?
                                ColorTranslator.ToHtml(_ipGridColors[ip]) : "";
                            string alias = _ipAliases.ContainsKey(ip) ? _ipAliases[ip] : "";
                            int loadingMCode = _ipLoadingMCodes.ContainsKey(ip) ? _ipLoadingMCodes[ip] : 0;

                            configList.Configs.Add(new IPConfig
                            {
                                IpAddress = ip,
                                Port = port,
                                ColorHex = colorHex,
                                Alias = alias,
                                LoadingMCode = loadingMCode
                            });
                        }
                    }
                }

                var serializer = new XmlSerializer(typeof(IPConfigList));
                using (var writer = new FileStream(_ipConfigFile, FileMode.Create))
                {
                    serializer.Serialize(writer, configList);
                }

                // 저장 성공 디버그 메시지 추가
                System.Diagnostics.Debug.WriteLine($"IP 설정 저장 완료: {_ipConfigFile}");
                System.Diagnostics.Debug.WriteLine($"저장된 IP 개수: {configList.Configs.Count}");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"IP 설정 저장 중 오류 발생:\n{ex.Message}\n\n저장 경로: {_ipConfigFile}", "저장 오류",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // 설정 로드
        private void LoadSettings()
        {
            try
            {
                if (File.Exists(_settingsFile))
                {
                    var serializer = new XmlSerializer(typeof(AppSettings));
                    using (var reader = new FileStream(_settingsFile, FileMode.Open))
                    {
                        var settings = (AppSettings)serializer.Deserialize(reader);
                        _syncIntervalSeconds = settings.SyncIntervalSeconds;

                        // 유효성 검사: 0 이하면 기본값 60초 사용
                        if (_syncIntervalSeconds <= 0)
                        {
                            _syncIntervalSeconds = 60;
                        }

                        // UI 업데이트
                        if (_syncIntervalInput != null)
                        {
                            _syncIntervalInput.Text = _syncIntervalSeconds.ToString();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"설정 로드 중 오류: {ex.Message}");
                // 로드 실패 시 기본값 사용
                _syncIntervalSeconds = 60;
            }
        }

        // 설정 저장
        private void SaveSettings()
        {
            try
            {
                var settings = new AppSettings
                {
                    SyncIntervalSeconds = _syncIntervalSeconds
                };

                var serializer = new XmlSerializer(typeof(AppSettings));
                using (var writer = new FileStream(_settingsFile, FileMode.Create))
                {
                    serializer.Serialize(writer, settings);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"설정 저장 중 오류 발생: {ex.Message}", "저장 오류",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void AddDefaultConnectionWithParams(string ip, ushort port)
        {
            if (_connectionManager.AddConnection(ip, port, DEFAULT_TIMEOUT))
            {
                _ipList.Items.Add($"{ip}:{port}");
                var connection = _connectionManager.GetConnection(ip);
                if (connection.Connect())
                {
                    _currentConnection = connection;
                    _connectionStates[ip] = true;
                    UpdateMachineStatus(ip);
                    _updateTimer.Start();
                }
                else
                {
                    _connectionStates[ip] = false;
                }
                _ipList.Invalidate();
            }
        }

        private void UpdateTimer_Tick(object sender, EventArgs e)
        {
            // 모든 등록된 연결에 대해 상태 확인 및 유지
            foreach (var connection in _connectionManager.GetAllConnections())
            {
                try
                {
                    if (connection.IsConnected && connection.CheckConnection())
                    {
                        _connectionStates[connection.IpAddress] = true;

                        // 현재 선택된 연결인 경우만 화면 업데이트
                        if (_currentConnection != null && _currentConnection.IpAddress == connection.IpAddress)
                        {
                            // 수동 옵셋 페이지가 표시중일 때만 업데이트
                            if (_mainGroupBox != null && _mainGroupBox.Visible)
                            {
                                UpdateMachineStatus(connection.IpAddress);
                                RefreshPmcStates(); // PMC 상태도 함께 갱신
                            }

                            // 모니터링 화면이 표시중일 때만 업데이트
                            if (_monitoringGroup.Visible)
                            {
                                UpdateMonitoringData();
                            }
                        }
                    }
                    else
                    {
                        _connectionStates[connection.IpAddress] = false;
                        // 백그라운드에서 재연결 시도
                        if (connection.Connect())
                        {
                            _connectionStates[connection.IpAddress] = true;

                            // 현재 선택된 연결인 경우만 화면 업데이트
                            if (_currentConnection != null && _currentConnection.IpAddress == connection.IpAddress)
                            {
                                // 수동 옵셋 페이지가 표시중일 때만 업데이트
                                if (_mainGroupBox != null && _mainGroupBox.Visible)
                                {
                                    UpdateMachineStatus(connection.IpAddress);
                                }
                            }
                        }
                    }
                }
                catch
                {
                    _connectionStates[connection.IpAddress] = false;
                    connection.Disconnect();
                    if (connection.Connect())
                    {
                        _connectionStates[connection.IpAddress] = true;

                        // 현재 선택된 연결인 경우만 화면 업데이트
                        if (_currentConnection != null && _currentConnection.IpAddress == connection.IpAddress)
                        {
                            UpdateMachineStatus(connection.IpAddress);
                        }
                    }
                }
            }
            _ipList.Invalidate();
        }

        private async void UpdateMonitoringData()
        {
            if (_currentConnection == null || !_currentConnection.IsConnected) return;

            // 이미 업데이트 중이면 스킵
            if (_isMonitoringUpdating) return;

            _isMonitoringUpdating = true;

            try
            {
                var handle = _currentConnection.Handle;

                // 모든 FOCAS 호출을 병렬로 실행
                var programTask = Task.Run(() => GetProgramInfo(handle));
                var toolTask = Task.Run(() => GetToolInfo(handle));
                var spindleTask = Task.Run(() => GetSpindleInfo(handle));
                var axisTask = Task.Run(() => GetAxisInfo(handle));
                var feedTask = Task.Run(() => GetFeedInfo(handle));
                var timeTask = Task.Run(() => GetMachineTime(handle));

                // 모든 작업이 완료될 때까지 대기
                await Task.WhenAll(programTask, toolTask, spindleTask, axisTask, feedTask, timeTask);

                // UI 업데이트 (UI 스레드에서 실행)
                if (programTask.Result != null) _lblProgramInfo.Text = programTask.Result;
                if (toolTask.Result != null) _lblToolInfo.Text = toolTask.Result;
                if (spindleTask.Result != null) _lblSpindleInfo.Text = spindleTask.Result;
                if (axisTask.Result != null) _lblAxisInfo.Text = axisTask.Result;
                if (feedTask.Result != null) _lblFeedrateInfo.Text = feedTask.Result;
                if (timeTask.Result != null) _lblMachineTime.Text = timeTask.Result;
            }
            catch (Exception ex)
            {
                // 오류 발생 시 로그만 남기고 계속 진행 (MessageBox 제거)
                WritePmcLog($"모니터링 데이터 업데이트 오류: {ex.Message}");
            }
            finally
            {
                _isMonitoringUpdating = false;
            }
        }

        // 개별 FOCAS 호출 메서드들 (백그라운드 스레드에서 실행)
        private string GetProgramInfo(ushort handle)
        {
            try
            {
                Focas1.ODBPRO prog = new Focas1.ODBPRO();
                Focas1.ODBEXEPRG exeprog = new Focas1.ODBEXEPRG();
                if (Focas1.cnc_rdprgnum(handle, prog) == Focas1.EW_OK &&
                    Focas1.cnc_exeprgname(handle, exeprog) == Focas1.EW_OK)
                {
                    return $"프로그램 정보:\nO{exeprog.name}\nN{prog.mdata:D4}";
                }
            }
            catch { }
            return null;
        }

        private string GetToolInfo(ushort handle)
        {
            try
            {
                Focas1.ODBMDL_1 modal = new Focas1.ODBMDL_1();
                Focas1.ODBSYS sysInfo = new Focas1.ODBSYS();
                if (Focas1.cnc_sysinfo(handle, sysInfo) == Focas1.EW_OK)
                {
                    string mtType = sysInfo.mt_type[0].ToString() + sysInfo.mt_type[1].ToString();

                    if (Focas1.cnc_modal(handle, 2, 1, modal) == Focas1.EW_OK)
                    {
                        if (mtType == "MM")
                            return $"공구/옵셋 정보:\nT{modal.g_data:D2}\nD{modal.g_data:D2}";
                        else
                            return $"공구 정보:\nT{modal.g_data:D2}";
                    }
                }
            }
            catch { }
            return null;
        }

        private string GetSpindleInfo(ushort handle)
        {
            try
            {
                Focas1.ODBACT speed = new Focas1.ODBACT();
                Focas1.ODBSPLOAD spload = new Focas1.ODBSPLOAD();
                short type = 0;
                short datano = 1;
                if (Focas1.cnc_acts(handle, speed) == Focas1.EW_OK &&
                    Focas1.cnc_rdspmeter(handle, type, ref datano, spload) == Focas1.EW_OK)
                {
                    return $"스핀들 정보:\n회전수: {speed.data} RPM\n부하: {spload.spload1.spload}%";
                }
            }
            catch { }
            return null;
        }

        private string GetAxisInfo(ushort handle)
        {
            try
            {
                Focas1.ODBAXIS axes = new Focas1.ODBAXIS();
                if (Focas1.cnc_absolute(handle, 1, 2, axes) == Focas1.EW_OK)
                {
                    return $"축 정보:\nX: {axes.data[0] / 1000.0:F3}\nZ: {axes.data[1] / 1000.0:F3}";
                }
            }
            catch { }
            return null;
        }

        private string GetFeedInfo(ushort handle)
        {
            try
            {
                Focas1.ODBACT feed = new Focas1.ODBACT();
                Focas1.ODBCMD cmd = new Focas1.ODBCMD();
                short num = 1;
                if (Focas1.cnc_actf(handle, feed) == Focas1.EW_OK &&
                    Focas1.cnc_rdcommand(handle, 1, 1, ref num, cmd) == Focas1.EW_OK)
                {
                    return $"이송 정보:\n실속도: {feed.data} mm/min\n지령: {cmd.cmd0.cmd_val} mm/min";
                }
            }
            catch { }
            return null;
        }

        private string GetMachineTime(ushort handle)
        {
            try
            {
                Focas1.IODBPSD_1 param = new Focas1.IODBPSD_1();
                if (Focas1.cnc_rdparam(handle, 6750, -1, 8, param) == Focas1.EW_OK)
                {
                    int minutes = param.idata;
                    return $"시간 정보:\n운전: {minutes / 60:D2}:{minutes % 60:D2}:00";
                }
            }
            catch { }
            return null;
        }

        private void UpdateOffsetLabels(short toolNo)
        {
            if (_currentConnection == null || !_currentConnection.IsConnected) return;

            double xGeometry = _currentConnection.GetToolOffset(toolNo, 0, true);
            double zGeometry = _currentConnection.GetToolOffset(toolNo, 1, true);
            double xWear = _currentConnection.GetToolOffset(toolNo, 0, false);
            double zWear = _currentConnection.GetToolOffset(toolNo, 1, false);

            _lblCurrentTool.Text = $"선택된 공구: {toolNo}번";
        }

        private void UpdateToolData(ushort handle, DataGridView grid)
        {
            if (_currentConnection == null || !_currentConnection.IsConnected) return;

            // 현재 선택된 행 인덱스 저장
            int selectedIndex = grid.CurrentRow?.Index ?? -1;
            short selectedToolNo = selectedIndex >= 0 ? 
                Convert.ToInt16(grid.CurrentRow.Cells["Tool"].Value) : (short)0;

            // IP에 따른 배경색 설정
            Color backgroundColor = _ipGridColors.ContainsKey(_currentConnection.IpAddress) 
                ? _ipGridColors[_currentConnection.IpAddress] 
                : Color.White;
            
            grid.BackgroundColor = backgroundColor;
            grid.DefaultCellStyle.BackColor = backgroundColor;
            
            grid.Rows.Clear();
            for (short toolNo = 1; toolNo <= 30; toolNo++)
            {
                double xGeometry = _currentConnection.GetToolOffset(toolNo, 0, true);
                double zGeometry = _currentConnection.GetToolOffset(toolNo, 1, true);
                double xWear = _currentConnection.GetToolOffset(toolNo, 0, false);
                double zWear = _currentConnection.GetToolOffset(toolNo, 1, false);

                var row = grid.Rows[grid.Rows.Add(toolNo, 
                    xGeometry.ToString("F3"),
                    xWear.ToString("F3"),
                    zGeometry.ToString("F3"),
                    zWear.ToString("F3"))];
                    
                // 각 셀의 배경색 설정
                foreach (DataGridViewCell cell in row.Cells)
                {
                    cell.Style.BackColor = backgroundColor;
                }
            }

            // 이전 선택 복원
            if (selectedIndex >= 0 && selectedIndex < grid.Rows.Count)
            {
                grid.CurrentCell = grid.Rows[selectedIndex].Cells[0];
            }
        }

        private void UpdateMachineStatus(string ipAddress)
        {
            var connection = _connectionManager.GetConnection(ipAddress);
            if (connection?.IsConnected == true)
            {
                _currentConnection = connection;
                _lblMode.Text = $"동작 모드: {connection.GetMode()}";
                _lblStatus.Text = $"상태: {connection.GetStatus()}";
                _lblOpSignal.Text = $"운전 신호: {connection.GetOpSignal()}";
                UpdateToolData(0, _gridTools);  // handle 파라미터는 사용하지 않으므로 0으로 전달
            }
        }

        private void UpdateWearValue(int axis)
        {
            // 선택된 IP 주소 가져오기
            string selectedIp = GetSelectedIpAddress();
            if (string.IsNullOrEmpty(selectedIp))
            {
                MessageBox.Show("연결할 설비를 선택해주세요.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var connection = _connectionManager.GetConnection(selectedIp);
            if (connection == null || !connection.IsConnected)
            {
                MessageBox.Show("선택한 설비가 연결되지 않았습니다.", "연결 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            TextBox txtInput = axis == 0 ? _txtXWear : _txtZWear;
            string axisName = axis == 0 ? "X" : "Z";
            
            if (_gridTools == null || _gridTools.CurrentRow == null)
            {
                MessageBox.Show("공구를 선택해주세요.", "알림", 
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string inputText = txtInput.Text;
            double inputValue;

            // 정수 입력 처리
            if (double.TryParse(inputText, out double value))
            {
                // 정수가 입력된 경우 자동 변환
                if (inputText.IndexOf('.') == -1)
                {
                    inputValue = value / 1000.0;  // 정수를 소수로 변환 (예: 100 -> 0.100)
                }
                else
                {
                    inputValue = value;  // 이미 소수점이 있는 경우 그대로 사용
                }

                // 입력값 범위 체크
                if (Math.Abs(inputValue) > 0.11)
                {
                    MessageBox.Show($"입력값({inputValue:F3})이 허용 범위(±0.11)를 초과합니다.", 
                        "입력값 범위 초과", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    txtInput.Text = "";
                    return;
                }

                try
                {
                    short toolNo = Convert.ToInt16(_gridTools.CurrentRow.Cells["Tool"].Value);

                    // 현재 마모값 읽기 (보정 전 데이터)
                    double currentWear = connection.GetToolOffset(toolNo, (short)axis, false);

                    // 새로운 마모값 계산
                    double newWear = currentWear + inputValue;

                    // 로그 엔트리 생성
                    var logEntry = new ManualOffsetLog
                    {
                        Timestamp = DateTime.Now,
                        IpAddress = selectedIp,
                        ToolNumber = toolNo,
                        Axis = axisName,
                        UserInput = inputValue,
                        BeforeValue = currentWear,
                        AfterValue = newWear,
                        SentValue = newWear,
                        Success = false
                    };

                    // 마모값 설정
                    if (connection.SetToolOffset(toolNo, (short)axis, false, newWear))
                    {
                        logEntry.Success = true;

                        // 로그 저장 (메모리 + DB)
                        _manualOffsetLogs.Add(logEntry);
                        _logDataService.SaveOffsetLog(logEntry);

                        // 데이터 갱신
                        UpdateToolData(0, _gridTools);

                        // 입력창 초기화
                        txtInput.Text = "";
                    }
                    else
                    {
                        logEntry.Success = false;
                        logEntry.ErrorMessage = "마모값 설정 실패";
                        _manualOffsetLogs.Add(logEntry);
                        _logDataService.SaveOffsetLog(logEntry);

                        MessageBox.Show("마모값 설정에 실패했습니다.", "오류",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                catch (Exception ex)
                {
                    // 예외 발생 시에도 로그 기록
                    var errorLog = new ManualOffsetLog
                    {
                        Timestamp = DateTime.Now,
                        IpAddress = selectedIp,
                        ToolNumber = _gridTools.CurrentRow != null ? Convert.ToInt16(_gridTools.CurrentRow.Cells["Tool"].Value) : (short)0,
                        Axis = axisName,
                        UserInput = inputValue,
                        BeforeValue = 0,
                        AfterValue = 0,
                        SentValue = 0,
                        Success = false,
                        ErrorMessage = ex.Message
                    };
                    _manualOffsetLogs.Add(errorLog);
                    _logDataService.SaveOffsetLog(errorLog);

                    MessageBox.Show($"마모값 설정 중 오류 발생: {ex.Message}", "오류",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("유효한 숫자를 입력해주세요.", "입력 오류", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtInput.Text = "";
            }
        }

        private void GridTools_SelectionChanged(object sender, EventArgs e)
        {
            if (_gridTools.CurrentRow != null)  // 모든 행 선택 가능
            {
                try
                {
                    var toolValue = _gridTools.CurrentRow.Cells["Tool"].Value;
                    if (toolValue != null && short.TryParse(toolValue.ToString(), out short toolNo))
                    {
                        _lblCurrentTool.Text = $"선택된 공구: {toolNo}번";
                        UpdateOffsetLabels(toolNo);
                    }
                }
                catch (Exception)
                {
                    // 변환 실패 시 무시
                }
            }
        }

        private void RegisterButton_Click(object sender, EventArgs e)
        {
            // 선택된 항목이 있으면 수정 모드, 없으면 신규 등록 모드
            bool isEditMode = _ipList.SelectedItem != null;
            IPConfig existingConfig = null;
            string oldIp = null;
            ushort oldPort = 0;

            if (isEditMode)
            {
                string selected = _ipList.SelectedItem.ToString();
                oldIp = selected.Split(':')[0];
                oldPort = ushort.Parse(selected.Split(':')[1]);

                // 기존 설정이 있으면 로드, 없으면 기본값으로 새로 생성
                if (_ipConfigs.ContainsKey(oldIp))
                {
                    existingConfig = _ipConfigs[oldIp];
                }
                else
                {
                    // 기존 데이터에서 설정 복원
                    existingConfig = new IPConfig
                    {
                        IpAddress = oldIp,
                        Port = oldPort,
                        Alias = _ipAliases.ContainsKey(oldIp) ? _ipAliases[oldIp] : "",
                        LoadingMCode = _ipLoadingMCodes.ContainsKey(oldIp) ? _ipLoadingMCodes[oldIp] : 0
                    };
                }
            }

            // 모달창 표시
            using (var form = new MachineRegistrationForm(existingConfig))
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    IPConfig config = form.Config;
                    string newIp = config.IpAddress;
                    ushort newPort = config.Port;

                    // 수정 모드인 경우
                    if (isEditMode)
                    {
                        bool isSameIp = (oldIp == newIp && oldPort == newPort);
                        Color existingColor = _ipGridColors.ContainsKey(oldIp) ? _ipGridColors[oldIp] : GenerateRandomPastelColor();

                        // IP가 변경된 경우 기존 연결 제거
                        if (!isSameIp)
                        {
                            _connectionManager.RemoveConnection(oldIp);
                            _connectionStates.Remove(oldIp);
                            _ipGridColors.Remove(oldIp);
                            _ipAliases.Remove(oldIp);
                            _ipLoadingMCodes.Remove(oldIp);
                            _ipConfigs.Remove(oldIp);

                            // 새 IP로 연결 추가
                            if (!_connectionManager.AddConnection(newIp, newPort, DEFAULT_TIMEOUT))
                            {
                                MessageBox.Show("새로운 IP로 연결을 생성하는데 실패했습니다.", "연결 오류",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }

                            // 새 IP에 색상 적용
                            _ipGridColors[newIp] = existingColor;
                        }
                        else
                        {
                            // IP가 같으면 기존 데이터만 제거
                            _ipAliases.Remove(oldIp);
                            _ipLoadingMCodes.Remove(oldIp);
                            _ipConfigs.Remove(oldIp);
                        }

                        // 새 설정 저장
                        _ipConfigs[newIp] = config;
                        _ipAliases[newIp] = config.Alias;
                        _ipLoadingMCodes[newIp] = config.LoadingMCode;

                        // 리스트 항목 업데이트
                        int selectedIndex = _ipList.SelectedIndex;
                        _ipList.Items[selectedIndex] = $"{newIp}:{newPort}";

                        SaveIPsToFile();
                        _ipList.Invalidate();

                        MessageBox.Show("설비 정보가 수정되었습니다.", "수정 완료",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);

                        _ipList.SelectedIndex = -1;
                    }
                    else
                    {
                        // 신규 등록 모드
                        if (_connectionManager.AddConnection(newIp, newPort, DEFAULT_TIMEOUT))
                        {
                            // 설정 저장
                            _ipConfigs[newIp] = config;
                            _ipAliases[newIp] = config.Alias;
                            _ipLoadingMCodes[newIp] = config.LoadingMCode;

                            // 랜덤 색상 생성
                            if (!_ipGridColors.ContainsKey(newIp))
                            {
                                _ipGridColors[newIp] = GenerateRandomPastelColor();
                            }

                            _ipList.Items.Add($"{newIp}:{newPort}");

                            var connection = _connectionManager.GetConnection(newIp);
                            if (connection.Connect())
                            {
                                if (_currentConnection == null)
                                {
                                    _currentConnection = connection;
                                    UpdateMachineStatus(newIp);
                                }

                                _connectionStates[newIp] = true;

                                if (!_updateTimer.Enabled)
                                {
                                    _updateTimer.Start();
                                }

                                MessageBox.Show("연결에 성공했습니다.", "연결 성공",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            else
                            {
                                _connectionStates[newIp] = false;
                                MessageBox.Show($"연결에 실패했습니다.\n{connection.ConnectionStatus}", "연결 실패",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }

                            SaveIPsToFile();
                            _ipList.Invalidate();
                        }
                        else
                        {
                            MessageBox.Show("이미 등록된 IP 주소입니다.", "등록 오류",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                }
            }
        }

        private void DeleteButton_Click(object sender, EventArgs e)
        {
            if (_ipList.SelectedItem == null)
            {
                MessageBox.Show("삭제할 IP를 선택해주세요.", "선택 오류",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string selected = _ipList.SelectedItem.ToString();
            string ip = selected.Split(':')[0];

            if (_connectionManager.RemoveConnection(ip))
            {
                _ipList.Items.Remove(selected);
                _connectionStates.Remove(ip);
                _ipGridColors.Remove(ip);
                _ipAliases.Remove(ip);
                _ipLoadingMCodes.Remove(ip);
                _ipConfigs.Remove(ip);

                if (_currentConnection?.IpAddress == ip)
                {
                    _currentConnection = null;
                }

                // IP 설정 저장
                SaveIPsToFile();
                _ipList.Invalidate();
            }
        }

        private void IpList_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_ipList.SelectedItem == null)
            {
                _registerButton.Text = "등록";
                return;
            }

            string selected = _ipList.SelectedItem.ToString();
            string ip = selected.Split(':')[0];

            // 버튼 텍스트를 "수정"으로 변경
            _registerButton.Text = "수정";

            var connection = _connectionManager.GetConnection(ip);
            if (connection != null)
            {
                // 기존 연결은 유지하고 현재 표시용 연결만 변경
                _currentConnection = connection;

                if (!connection.IsConnected)
                {
                    // 연결이 끊어진 상태면 재연결 시도 (조용히)
                    if (!connection.Connect())
                    {
                        _connectionStates[ip] = false;
                        _ipList.Invalidate();
                        return;
                    }
                }

                _connectionStates[ip] = true;
                UpdateMachineStatus(ip);

                // 자동 모니터링 폼이 열려있으면 연결 업데이트
                var autoMonitoringForm = _centerPanel.Controls.OfType<AutoMonitoringForm>().FirstOrDefault();
                if (autoMonitoringForm != null)
                {
                    autoMonitoringForm.UpdateConnection(connection);
                }

                // 수동 옵셋 화면이 표시중이면 매크로 데이터 로드
                if (_mainGroupBox.Visible)
                {
                    LoadAllMacroValuesAsync();
                }

                _ipList.Invalidate();
            }
        }

        // 선택된 IP 주소 가져오기
        private string GetSelectedIpAddress()
        {
            if (_ipList.SelectedItem != null)
            {
                string selectedItem = _ipList.SelectedItem.ToString();
                return selectedItem.Split(':')[0];
            }
            return null;
        }

        // 매크로 그리드 셀 클릭 이벤트 (매크로 번호 선택)
        private void MacroGridView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            // 헤더 클릭 무시
            if (e.RowIndex < 0) return;

            // 매크로 번호 추출 (#500 -> 500)
            string macroNoStr = _macroGridView.Rows[e.RowIndex].Cells["MacroNo"].Value.ToString().Replace("#", "");
            if (short.TryParse(macroNoStr, out short macroNo))
            {
                _selectedMacroNo = macroNo;

                // 텍스트박스는 비워둠 (증감값 입력용)
                _macroValueInput.Text = "";

                // 값 입력란으로 포커스 이동
                _macroValueInput.Focus();
            }
        }

        // 매크로 그리드 셀 더블클릭 이벤트 (값 읽기)
        private void MacroGridView_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            // 헤더 클릭 무시
            if (e.RowIndex < 0) return;

            if (_currentConnection == null || !_currentConnection.IsConnected)
            {
                MessageBox.Show("CNC에 연결되지 않았습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 매크로 번호 추출 (#500 -> 500)
            string macroNoStr = _macroGridView.Rows[e.RowIndex].Cells["MacroNo"].Value.ToString().Replace("#", "");
            if (!short.TryParse(macroNoStr, out short macroNo))
            {
                MessageBox.Show("매크로 번호를 읽을 수 없습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // CNC에서 매크로 값 읽기
            Focas1.ODBM macro = new Focas1.ODBM();
            short ret = Focas1.cnc_rdmacro(_currentConnection.Handle, macroNo, 10, macro);

            if (ret == Focas1.EW_OK)
            {
                double value = macro.mcr_val * Math.Pow(10, -macro.dec_val);
                _macroGridView.Rows[e.RowIndex].Cells["MacroValue"].Value = value.ToString("F3");

                // 선택된 매크로 저장 및 텍스트박스에도 로드
                _selectedMacroNo = macroNo;
                _macroValueInput.Text = value.ToString("F3");
            }
            else
            {
                MessageBox.Show($"매크로 #{macroNo} 읽기 실패 (오류코드: {ret})", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 동기화 간격 적용 버튼 클릭 이벤트
        private void BtnApplySyncInterval_Click(object sender, EventArgs e)
        {
            // 입력값 검증
            if (string.IsNullOrWhiteSpace(_syncIntervalInput.Text))
            {
                MessageBox.Show("동기화 간격을 입력하세요.", "입력 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                _syncIntervalInput.Focus();
                return;
            }

            if (!int.TryParse(_syncIntervalInput.Text, out int newInterval))
            {
                MessageBox.Show("올바른 숫자를 입력하세요.", "입력 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                _syncIntervalInput.Focus();
                return;
            }

            // 범위 검증 (0초(중지) ~ 3600초 / 1시간)
            if (newInterval < 0 || newInterval > 3600)
            {
                MessageBox.Show("동기화 간격은 0초(중지) ~ 3600초(1시간) 범위 내에서 입력하세요.", "입력 범위 초과",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                _syncIntervalInput.Focus();
                _syncIntervalInput.SelectAll();
                return;
            }

            // 이전 값과 동일하면 무시
            if (newInterval == _syncIntervalSeconds)
            {
                MessageBox.Show("현재 설정과 동일한 값입니다.", "정보", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 설정 업데이트
            _syncIntervalSeconds = newInterval;

            // 0초 입력 시 타이머 중지
            if (_syncIntervalSeconds == 0)
            {
                _syncTimer.Stop();
                SaveSettings();
                MessageBox.Show("동기화가 중지되었습니다.", "설정 완료",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 타이머 간격 업데이트 (타이머가 실행 중이면 재시작)
            bool wasRunning = _syncTimer.Enabled;
            if (wasRunning)
            {
                _syncTimer.Stop();
            }

            _syncTimer.Interval = _syncIntervalSeconds * 1000;

            if (wasRunning)
            {
                _syncTimer.Start();
            }

            // 설정 저장
            SaveSettings();

            MessageBox.Show($"동기화 간격이 {_syncIntervalSeconds}초로 변경되었습니다.", "설정 완료",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // ========== PMC 제어 메서드 ==========

        // Block Skip 버튼 상태 업데이트 (LED 스타일)
        private void UpdateBlockSkipButton(bool isOn)
        {
            if (_btnBlockSkip.InvokeRequired)
            {
                _btnBlockSkip.Invoke(new Action(() => UpdateBlockSkipButton(isOn)));
                return;
            }

            if (isOn)
            {
                // ON 상태: 녹색 LED
                _btnBlockSkip.ForeColor = Color.LimeGreen;
                _btnBlockSkip.BackColor = Color.FromArgb(230, 255, 230); // 연한 녹색 배경
            }
            else
            {
                // OFF 상태: 회색 LED
                _btnBlockSkip.ForeColor = Color.Gray;
                _btnBlockSkip.BackColor = Color.FromArgb(240, 240, 240); // 회색 배경
            }
        }

        // Optional Stop 버튼 상태 업데이트 (LED 스타일)
        private void UpdateOptStopButton(bool isOn)
        {
            if (_btnOptStop.InvokeRequired)
            {
                _btnOptStop.Invoke(new Action(() => UpdateOptStopButton(isOn)));
                return;
            }

            if (isOn)
            {
                // ON 상태: 녹색 LED
                _btnOptStop.ForeColor = Color.LimeGreen;
                _btnOptStop.BackColor = Color.FromArgb(230, 255, 230); // 연한 녹색 배경
            }
            else
            {
                // OFF 상태: 회색 LED
                _btnOptStop.ForeColor = Color.Gray;
                _btnOptStop.BackColor = Color.FromArgb(240, 240, 240); // 회색 배경
            }
        }

        // PMC 로그 기록
        private void WritePmcLog(string message)
        {
            try
            {
                string logMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {message}";
                File.AppendAllText(_pmcLogPath, logMessage + Environment.NewLine);
            }
            catch
            {
                // 로그 기록 실패 시 무시
            }
        }

        // PMC 비트 읽기
        private bool ReadPmcBit(short pmcType, ushort address, int bitPosition, out bool bitValue)
        {
            bitValue = false;
            string pmcTypeName = pmcType == PMC_TYPE_G ? "G" :
                                (pmcType == PMC_TYPE_R ? "R" :
                                (pmcType == PMC_TYPE_D ? "D" : pmcType.ToString()));

            WritePmcLog($"===== ReadPmcBit 시작 =====");
            WritePmcLog($"  요청: PMC_{pmcTypeName}{address}.{bitPosition}");
            WritePmcLog($"  연결 상태: {(_currentConnection != null ? _currentConnection.IpAddress : "NULL")} / " +
                       $"IsConnected: {(_currentConnection?.IsConnected ?? false)}");

            if (_currentConnection == null || !_currentConnection.IsConnected)
            {
                WritePmcLog($"  ❌ 실패: 연결 안됨");
                return false;
            }

            try
            {
                Focas1.IODBPMC0 pmcData = new Focas1.IODBPMC0();
                short ret = Focas1.pmc_rdpmcrng(_currentConnection.Handle, pmcType, PMC_DATA_BYTE,
                    address, address, 9, pmcData);

                WritePmcLog($"  pmc_rdpmcrng 반환 코드: {ret} (0=성공)");

                if (ret == Focas1.EW_OK)
                {
                    byte byteValue = pmcData.cdata[0];
                    bitValue = (byteValue & (1 << bitPosition)) != 0;
                    WritePmcLog($"  바이트 값: 0x{byteValue:X2} ({Convert.ToString(byteValue, 2).PadLeft(8, '0')}b)");
                    WritePmcLog($"  비트 {bitPosition} 값: {bitValue}");
                    WritePmcLog($"  ✅ 읽기 성공");
                    return true;
                }

                WritePmcLog($"  ❌ 읽기 실패: 오류 코드 {ret}");
                return false;
            }
            catch (Exception ex)
            {
                WritePmcLog($"  ❌ 예외 발생: {ex.Message}");
                WritePmcLog($"  스택: {ex.StackTrace}");
                return false;
            }
        }

        // ==================== 설비 상태 모니터링용 PMC 읽기 ====================

        // 다중 PMC 주소 일괄 읽기
        // PMC 비트 쓰기
        private bool WritePmcBit(short pmcType, ushort address, int bitPosition, bool bitValue)
        {
            string pmcTypeName = pmcType == PMC_TYPE_G ? "G" :
                                (pmcType == PMC_TYPE_R ? "R" :
                                (pmcType == PMC_TYPE_D ? "D" : pmcType.ToString()));

            WritePmcLog($"");
            WritePmcLog($"===== WritePmcBit 시작 =====");
            WritePmcLog($"  요청: PMC_{pmcTypeName}{address}.{bitPosition} = {bitValue}");
            WritePmcLog($"  연결 상태: {(_currentConnection != null ? _currentConnection.IpAddress : "NULL")} / " +
                       $"IsConnected: {(_currentConnection?.IsConnected ?? false)}");
            WritePmcLog($"  Handle: {_currentConnection?.Handle ?? 0}");

            if (_currentConnection == null || !_currentConnection.IsConnected)
            {
                WritePmcLog($"  ❌ 실패: 연결 안됨");
                MessageBox.Show("CNC에 연결되지 않았습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            try
            {
                // 1. 현재 값 읽기
                WritePmcLog($"  [단계 1] 현재 값 읽기 시작...");
                Focas1.IODBPMC0 pmcData = new Focas1.IODBPMC0();
                short ret = Focas1.pmc_rdpmcrng(_currentConnection.Handle, pmcType, PMC_DATA_BYTE,
                    address, address, 9, pmcData);

                WritePmcLog($"  pmc_rdpmcrng 반환 코드: {ret} (0=성공)");

                if (ret != Focas1.EW_OK)
                {
                    WritePmcLog($"  ❌ PMC 읽기 실패 (오류코드: {ret})");
                    MessageBox.Show($"PMC 읽기 실패 (오류코드: {ret})", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }

                byte currentValue = pmcData.cdata[0];
                WritePmcLog($"  현재 바이트 값: 0x{currentValue:X2} ({Convert.ToString(currentValue, 2).PadLeft(8, '0')}b)");
                WritePmcLog($"  현재 비트 {bitPosition} 값: {((currentValue & (1 << bitPosition)) != 0)}");

                // 2. 비트 수정
                WritePmcLog($"  [단계 2] 비트 수정...");
                byte newValue;
                if (bitValue)
                {
                    newValue = (byte)(currentValue | (1 << bitPosition)); // 비트 켜기
                    WritePmcLog($"  동작: 비트 {bitPosition} ON (OR 연산)");
                }
                else
                {
                    newValue = (byte)(currentValue & ~(1 << bitPosition)); // 비트 끄기
                    WritePmcLog($"  동작: 비트 {bitPosition} OFF (AND NOT 연산)");
                }

                WritePmcLog($"  새 바이트 값: 0x{newValue:X2} ({Convert.ToString(newValue, 2).PadLeft(8, '0')}b)");
                WritePmcLog($"  값 변경 여부: {currentValue != newValue} (이전: {currentValue}, 새값: {newValue})");

                // 3. 쓰기
                WritePmcLog($"  [단계 3] PMC 쓰기 시작...");
                pmcData.type_a = pmcType;
                pmcData.type_d = PMC_DATA_BYTE;
                pmcData.datano_s = (short)address;
                pmcData.datano_e = (short)address;
                pmcData.cdata = new byte[5];
                pmcData.cdata[0] = newValue;

                WritePmcLog($"  쓰기 파라미터:");
                WritePmcLog($"    - type_a: {pmcData.type_a}");
                WritePmcLog($"    - type_d: {pmcData.type_d}");
                WritePmcLog($"    - datano_s: {pmcData.datano_s}");
                WritePmcLog($"    - datano_e: {pmcData.datano_e}");
                WritePmcLog($"    - cdata[0]: 0x{pmcData.cdata[0]:X2}");

                ret = Focas1.pmc_wrpmcrng(_currentConnection.Handle, 9, pmcData);

                WritePmcLog($"  pmc_wrpmcrng 반환 코드: {ret} (0=성공)");

                if (ret != Focas1.EW_OK)
                {
                    WritePmcLog($"  ❌ PMC 쓰기 실패 (오류코드: {ret})");
                    MessageBox.Show($"PMC 쓰기 실패 (오류코드: {ret})", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }

                // 4. 쓰기 후 검증 (읽기)
                WritePmcLog($"  [단계 4] 쓰기 검증 (재읽기)...");
                System.Threading.Thread.Sleep(50); // 약간의 대기
                Focas1.IODBPMC0 verifyData = new Focas1.IODBPMC0();
                ret = Focas1.pmc_rdpmcrng(_currentConnection.Handle, pmcType, PMC_DATA_BYTE,
                    address, address, 9, verifyData);

                if (ret == Focas1.EW_OK)
                {
                    byte verifyValue = verifyData.cdata[0];
                    bool verifyBit = (verifyValue & (1 << bitPosition)) != 0;
                    WritePmcLog($"  검증 바이트 값: 0x{verifyValue:X2} ({Convert.ToString(verifyValue, 2).PadLeft(8, '0')}b)");
                    WritePmcLog($"  검증 비트 {bitPosition} 값: {verifyBit}");
                    WritePmcLog($"  쓰기 반영 확인: {(verifyBit == bitValue ? "✅ 성공" : "❌ 실패 - 값 불일치")}");

                    if (verifyBit != bitValue)
                    {
                        WritePmcLog($"  ⚠️ 경고: 쓰기 성공했으나 검증 시 값이 일치하지 않음!");
                        WritePmcLog($"    요청값: {bitValue}, 실제값: {verifyBit}");
                    }
                }
                else
                {
                    WritePmcLog($"  ⚠️ 검증 읽기 실패 (오류코드: {ret})");
                }

                WritePmcLog($"  ✅ WritePmcBit 완료");
                return true;
            }
            catch (Exception ex)
            {
                WritePmcLog($"  ❌ 예외 발생: {ex.Message}");
                WritePmcLog($"  타입: {ex.GetType().FullName}");
                WritePmcLog($"  스택: {ex.StackTrace}");
                MessageBox.Show($"PMC 제어 중 오류 발생: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        // PMC 주소 파싱 헬퍼 함수
        private bool ParsePmcAddress(string address, out short pmcType, out ushort pmcAddress, out int bitPosition)
        {
            pmcType = 0;
            pmcAddress = 0;
            bitPosition = 0;

            try
            {
                if (string.IsNullOrWhiteSpace(address))
                    return false;

                // Type 추출 (첫 문자)
                char typeChar = address[0];
                string typeMap = typeChar switch
                {
                    'F' => "F",
                    'G' => "G",
                    'R' => "R",
                    'D' => "D",
                    'X' => "X",
                    _ => null
                };

                if (typeMap == null)
                    return false;

                pmcType = typeChar switch
                {
                    'F' => PMC_TYPE_F,
                    'G' => PMC_TYPE_G,
                    'R' => PMC_TYPE_R,
                    'D' => PMC_TYPE_D,
                    'X' => PMC_TYPE_X,
                    _ => (short)0
                };

                // Address와 Bit 추출
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
            catch
            {
                return false;
            }
        }

        // Block Skip 버튼 클릭
        private void ChkBlockSkip_Click(object sender, EventArgs e)
        {
            if (_currentConnection == null || !_ipConfigs.ContainsKey(_currentConnection.IpAddress))
                return;

            var config = _ipConfigs[_currentConnection.IpAddress];

            // PMC 주소 파싱
            if (!ParsePmcAddress(config.BlockSkipStateAddr, out short stateType, out ushort stateAddr, out int stateBit))
            {
                MessageBox.Show("Block Skip 상태 주소 설정이 올바르지 않습니다.", "설정 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!ParsePmcAddress(config.BlockSkipInputAddr, out short inputType, out ushort inputAddr, out int inputBit))
            {
                MessageBox.Show("Block Skip 입력 주소 설정이 올바르지 않습니다.", "설정 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 현재 상태 읽기
            ReadPmcBit(stateType, stateAddr, stateBit, out bool beforeValue);
            bool success = false;

            // 버튼 입력에 펄스 전송 (OFF → ON → OFF 시퀀스)
            if (WritePmcBit(inputType, inputAddr, inputBit, true))
            {
                System.Threading.Thread.Sleep(300); // 300ms 대기 (래더 스캔 보장)
                WritePmcBit(inputType, inputAddr, inputBit, false);

                // 래더 처리 대기
                System.Threading.Thread.Sleep(200);
                RefreshPmcStates(); // 즉시 상태 갱신
                success = true;
            }

            // 변경 후 상태 읽기
            ReadPmcBit(stateType, stateAddr, stateBit, out bool afterValue);

            // 로그 저장
            try
            {
                var log = new PmcControlLog
                {
                    Timestamp = DateTime.Now,
                    IpAddress = _currentConnection?.IpAddress ?? "알 수 없음",
                    ControlType = "Block Skip",
                    BeforeValue = beforeValue,
                    AfterValue = afterValue,
                    Success = success,
                    ErrorMessage = success ? null : "PMC 쓰기 실패"
                };
                _logDataService.SavePmcLog(log);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Block Skip 로그 저장 실패: {ex.Message}");
            }
        }

        // Optional Stop 버튼 클릭
        private void ChkOptStop_Click(object sender, EventArgs e)
        {
            if (_currentConnection == null || !_ipConfigs.ContainsKey(_currentConnection.IpAddress))
                return;

            var config = _ipConfigs[_currentConnection.IpAddress];

            // PMC 주소 파싱
            if (!ParsePmcAddress(config.OptStopStateAddr, out short stateType, out ushort stateAddr, out int stateBit))
            {
                MessageBox.Show("Optional Stop 상태 주소 설정이 올바르지 않습니다.", "설정 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!ParsePmcAddress(config.OptStopInputAddr, out short inputType, out ushort inputAddr, out int inputBit))
            {
                MessageBox.Show("Optional Stop 입력 주소 설정이 올바르지 않습니다.", "설정 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 현재 상태 읽기
            ReadPmcBit(stateType, stateAddr, stateBit, out bool beforeValue);
            bool success = false;

            // 버튼 입력에 펄스 전송 (OFF → ON → OFF 시퀀스)
            if (WritePmcBit(inputType, inputAddr, inputBit, true))
            {
                System.Threading.Thread.Sleep(300); // 300ms 대기 (래더 스캔 보장)
                WritePmcBit(inputType, inputAddr, inputBit, false);

                // 래더 처리 대기
                System.Threading.Thread.Sleep(200);
                RefreshPmcStates(); // 즉시 상태 갱신
                success = true;
            }

            // 변경 후 상태 읽기
            ReadPmcBit(stateType, stateAddr, stateBit, out bool afterValue);

            // 로그 저장
            try
            {
                var log = new PmcControlLog
                {
                    Timestamp = DateTime.Now,
                    IpAddress = _currentConnection?.IpAddress ?? "알 수 없음",
                    ControlType = "Optional Stop",
                    BeforeValue = beforeValue,
                    AfterValue = afterValue,
                    Success = success,
                    ErrorMessage = success ? null : "PMC 쓰기 실패"
                };
                _logDataService.SavePmcLog(log);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Optional Stop 로그 저장 실패: {ex.Message}");
            }
        }

        // PMC 상태 새로고침
        private void RefreshPmcStates()
        {
            if (_currentConnection == null || !_currentConnection.IsConnected)
            {
                return;
            }

            if (!_ipConfigs.ContainsKey(_currentConnection.IpAddress))
            {
                return;
            }

            var config = _ipConfigs[_currentConnection.IpAddress];

            // Block Skip 상태 읽기
            if (ParsePmcAddress(config.BlockSkipStateAddr, out short bsType, out ushort bsAddr, out int bsBit))
            {
                if (ReadPmcBit(bsType, bsAddr, bsBit, out bool blockSkipState))
                {
                    UpdateBlockSkipButton(blockSkipState);
                }
            }

            // Optional Stop 상태 읽기
            if (ParsePmcAddress(config.OptStopStateAddr, out short osType, out ushort osAddr, out int osBit))
            {
                if (ReadPmcBit(osType, osAddr, osBit, out bool optStopState))
                {
                    UpdateOptStopButton(optStopState);
                }
            }
        }

        // 매크로 쓰기 버튼 클릭 이벤트
        private void BtnMacroWrite_Click(object sender, EventArgs e)
        {
            if (_currentConnection == null || !_currentConnection.IsConnected)
            {
                MessageBox.Show("CNC에 연결되지 않았습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 그리드에서 현재 선택된 행의 매크로 번호 가져오기
            short macroNo = 0;
            if (_macroGridView.CurrentRow != null && _macroGridView.CurrentRow.Index >= 0)
            {
                string macroNoStr = _macroGridView.CurrentRow.Cells["MacroNo"].Value.ToString().Replace("#", "");
                if (short.TryParse(macroNoStr, out macroNo))
                {
                    _selectedMacroNo = macroNo; // 선택된 매크로 번호 업데이트
                }
            }

            // 매크로 선택 확인
            if (macroNo == 0)
            {
                MessageBox.Show("왼쪽 리스트에서 매크로를 선택하세요.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 증감값 입력 검증
            if (string.IsNullOrWhiteSpace(_macroValueInput.Text))
            {
                MessageBox.Show("증감값을 입력하세요.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                _macroValueInput.Focus();
                return;
            }

            // 정수 입력 처리 (마모 입력과 동일한 로직)
            string inputText = _macroValueInput.Text;
            double incrementValue;

            if (double.TryParse(inputText, out double value))
            {
                // 정수가 입력된 경우 자동 변환
                if (inputText.IndexOf('.') == -1)
                {
                    incrementValue = value / 1000.0;  // 정수를 소수로 변환 (예: 50 -> 0.050)
                }
                else
                {
                    incrementValue = value;  // 이미 소수점이 있는 경우 그대로 사용
                }
            }
            else
            {
                MessageBox.Show("올바른 값을 입력하세요. (숫자만 입력 가능)", "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                _macroValueInput.Focus();
                return;
            }

            // 증감값 범위 제한: -0.1 ~ +0.1
            if (incrementValue < -0.1 || incrementValue > 0.1)
            {
                MessageBox.Show("증감값은 -0.1 ~ +0.1 범위 내에서 입력하세요.", "입력 범위 초과", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                _macroValueInput.Focus();
                _macroValueInput.SelectAll();
                return;
            }

            // 기존 매크로 값 읽기
            Focas1.ODBM macro = new Focas1.ODBM();
            short readRet = Focas1.cnc_rdmacro(_currentConnection.Handle, macroNo, 10, macro);

            if (readRet != Focas1.EW_OK)
            {
                MessageBox.Show($"매크로 #{macroNo} 읽기 실패 (오류코드: {readRet})", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // 기존값 + 입력값 계산
            double currentValue = macro.mcr_val * Math.Pow(10, -macro.dec_val);
            double newValue = currentValue + incrementValue;

            // 새 값을 정수로 변환 (소수점 3자리 기준, 반올림 적용)
            int intValue = (int)Math.Round(newValue * 1000);
            short decValue = 3;

            // 로그 엔트리 생성
            var logEntry = new MacroVariableLog
            {
                Timestamp = DateTime.Now,
                IpAddress = _currentConnection.IpAddress,
                MacroNumber = macroNo,
                UserInput = incrementValue,
                BeforeValue = currentValue,
                AfterValue = newValue,
                SentValue = newValue,
                Success = false
            };

            short ret = Focas1.cnc_wrmacro(_currentConnection.Handle, macroNo, 10, intValue, decValue);

            if (ret == Focas1.EW_OK)
            {
                // 로그 성공 표시 (메모리 + DB)
                logEntry.Success = true;
                _macroVariableLogs.Add(logEntry);
                _logDataService.SaveMacroLog(logEntry);

                // 그리드뷰에도 업데이트
                int rowIndex = macroNo - 500;
                _macroGridView.Rows[rowIndex].Cells["MacroValue"].Value = newValue.ToString("F3");

                // DB에도 업데이트
                _macroDataService.SaveOrUpdateMacroData(_currentConnection.IpAddress, macroNo, newValue);

                // 텍스트박스 초기화
                _macroValueInput.Text = "";
            }
            else
            {
                // 로그 실패 표시 (메모리 + DB)
                logEntry.ErrorMessage = $"오류코드: {ret}";
                _macroVariableLogs.Add(logEntry);
                _logDataService.SaveMacroLog(logEntry);

                MessageBox.Show($"매크로 #{macroNo} 쓰기 실패 (오류코드: {ret})", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 모든 매크로 값을 DB에서 로드 후 CNC와 동기화
        private void LoadAllMacroValuesAsync()
        {
            // 현재 연결이 없으면 IP 리스트에서 선택된 항목의 연결을 사용
            if (_currentConnection == null || !_currentConnection.IsConnected)
            {
                string selectedIp = GetSelectedIpAddress();

                if (!string.IsNullOrEmpty(selectedIp))
                {
                    var connection = _connectionManager.GetConnection(selectedIp);

                    if (connection != null && connection.IsConnected)
                    {
                        _currentConnection = connection;
                    }
                    else
                    {
                        return;
                    }
                }
                else
                {
                    return;
                }
            }

            string currentIp = _currentConnection.IpAddress;

            // DB에 데이터가 있는지 확인
            bool hasDbData = _macroDataService.HasMacroDataForIp(currentIp);

            if (hasDbData)
            {
                // DB에서 로드 (즉시 표시)
                var dbData = _macroDataService.LoadMacroDataFromDb(currentIp);

                foreach (var kvp in dbData)
                {
                    int rowIndex = kvp.Key - 500;
                    if (rowIndex >= 0 && rowIndex < 500)
                    {
                        _macroGridView.Rows[rowIndex].Cells["MacroValue"].Value = kvp.Value.ToString("F3");
                    }
                }

                // 백그라운드 동기화 시작
                StartBackgroundSync();
            }
            else
            {
                // DB에 데이터 없음 - CNC에서 초기 로드
                LoadFromCncAndSaveToDb(currentIp);
            }
        }

        // CNC에서 초기 로드하고 DB에 저장
        private void LoadFromCncAndSaveToDb(string ipAddress)
        {
            _macroLoadProgress.Value = 0;
            _macroLoadProgress.Maximum = 500;
            _macroLoadProgress.Visible = true;
            _macroLoadLabel.Text = "매크로 데이터 로딩 중... 0/500 (0%)";
            _macroLoadLabel.Visible = true;

            var macroData = new Dictionary<short, double>();

            for (int i = 500; i <= 999; i++)
            {
                if (_currentConnection == null || !_currentConnection.IsConnected)
                    break;

                short macroNo = (short)i;
                int rowIndex = i - 500;

                try
                {
                    Focas1.ODBM macro = new Focas1.ODBM();
                    short ret = Focas1.cnc_rdmacro(_currentConnection.Handle, macroNo, 10, macro);

                    if (ret == Focas1.EW_OK)
                    {
                        double value = macro.mcr_val * Math.Pow(10, -macro.dec_val);

                        // 그리드 업데이트
                        _macroGridView.Rows[rowIndex].Cells["MacroValue"].Value = value.ToString("F3");

                        // DB 저장용 데이터 수집
                        macroData[macroNo] = value;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"매크로 #{macroNo} 읽기 오류: {ex.Message}");
                }

                // 진행률 업데이트
                if ((i - 500 + 1) % 10 == 0)
                {
                    int progress = i - 500 + 1;
                    int percent = (progress * 100) / 500;

                    _macroLoadProgress.Value = progress;
                    _macroLoadLabel.Text = $"매크로 데이터 로딩 중... {progress}/500 ({percent}%)";
                    Application.DoEvents();
                }
            }

            // DB에 일괄 저장
            if (macroData.Count > 0)
            {
                _macroDataService.SaveOrUpdateMacroDataBatch(ipAddress, macroData);
            }

            _macroLoadProgress.Visible = false;
            _macroLoadLabel.Visible = false;
        }

        // 동기화 타이머 초기화
        private void InitializeSyncTimer()
        {
            // _syncIntervalSeconds가 0이면 기본값 60초 사용
            if (_syncIntervalSeconds <= 0)
            {
                _syncIntervalSeconds = 60;
            }

            _syncTimer = new System.Windows.Forms.Timer
            {
                Interval = _syncIntervalSeconds * 1000 // 설정된 초 단위를 밀리초로 변환
            };
            _syncTimer.Tick += SyncTimer_Tick;
        }

        // 백그라운드 동기화 시작
        private void StartBackgroundSync()
        {
            if (!_syncTimer.Enabled)
            {
                _syncTimer.Start();
            }
        }

        // 타이머 틱 이벤트 - 모든 연결된 IP에 대해 동기화
        private async void SyncTimer_Tick(object sender, EventArgs e)
        {
            // 이미 동기화 중이면 스킵
            if (_isSyncing)
                return;

            _isSyncing = true;

            try
            {
                // 연결된 모든 IP 가져오기
                var connectedIps = _connectionManager.GetAllConnections()
                    .Where(c => c.IsConnected)
                    .Select(c => c.IpAddress)
                    .ToList();

                if (connectedIps.Count == 0)
                    return;

                // 각 IP별로 비동기 작업 생성
                var tasks = new List<Task>();

                foreach (string ip in connectedIps)
                {
                    // DB에 데이터가 없으면 건너뛰기
                    if (!_macroDataService.HasMacroDataForIp(ip))
                        continue;

                    var connection = _connectionManager.GetConnection(ip);
                    if (connection == null || !connection.IsConnected)
                        continue;

                    // 각 IP의 매크로 동기화를 백그라운드 작업으로 추가
                    tasks.Add(SyncMacroValuesAsync(ip, connection));
                }

                // 모든 IP의 동기화를 병렬로 실행
                await Task.WhenAll(tasks);
            }
            finally
            {
                _isSyncing = false;
            }
        }

        // 단일 IP의 매크로 값 동기화 (백그라운드 처리)
        private async Task SyncMacroValuesAsync(string ip, CNCConnection connection)
        {
            await Task.Run(() =>
            {
                int startIndex = 500;
                int endIndex = 1000;

                // DB에서 현재 값 한 번만 로드 (매번 로드하지 않음)
                var dbData = _macroDataService.LoadMacroDataFromDb(ip);

                for (int i = startIndex; i < endIndex; i++)
                {
                    short macroNo = (short)i;

                    try
                    {
                        Focas1.ODBM macro = new Focas1.ODBM();
                        short ret = Focas1.cnc_rdmacro(connection.Handle, macroNo, 10, macro);

                        if (ret == Focas1.EW_OK)
                        {
                            double cncValue = macro.mcr_val * Math.Pow(10, -macro.dec_val);

                            if (dbData.ContainsKey(macroNo))
                            {
                                double dbValue = dbData[macroNo];

                                // 값이 다르면 DB 업데이트
                                if (Math.Abs(cncValue - dbValue) > 0.0001)
                                {
                                    _macroDataService.SaveOrUpdateMacroData(ip, macroNo, cncValue);

                                    // 현재 표시 중인 IP라면 그리드도 업데이트 (UI 스레드에서)
                                    if (_currentConnection?.IpAddress == ip && _mainGroupBox.Visible)
                                    {
                                        int rowIndex = i - 500;
                                        if (rowIndex >= 0 && rowIndex < 500)
                                        {
                                            // UI 스레드로 전환
                                            this.Invoke(new Action(() =>
                                            {
                                                if (_macroGridView.Rows.Count > rowIndex)
                                                {
                                                    _macroGridView.Rows[rowIndex].Cells["MacroValue"].Value = cncValue.ToString("F3");
                                                }
                                            }));
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch
                    {
                        // 오류 무시
                    }
                }
            });
        }

        // FOCAS 에러 코드 설명 반환
        private string GetFocasErrorDescription(short errorCode)
        {
            switch (errorCode)
            {
                case 0:  // EW_OK
                    return "정상";
                case 1:  // EW_FUNC / EW_NOPMC
                    return "기능 오류 - 명령 준비 오류 또는 PMC 없음";
                case 2:  // EW_LENGTH
                    return "데이터 블록 길이 오류";
                case 3:  // EW_NUMBER / EW_RANGE
                    return "데이터 번호 오류 또는 주소 범위 오류";
                case 4:  // EW_ATTRIB / EW_TYPE
                    return "데이터 속성 오류 또는 데이터 타입 오류";
                case 5:  // EW_DATA
                    return "데이터 오류";
                case 6:  // EW_NOOPT
                    return "옵션 없음 - 해당 기능이 장비에 없음";
                case 7:  // EW_PROT
                    return "쓰기 보호 오류";
                case 8:  // EW_OVRFLOW
                    return "메모리 오버플로우 오류";
                case 9:  // EW_PARAM
                    return "CNC 파라미터 오류";
                case 10: // EW_BUFFER
                    return "버퍼 오류";
                case 11: // EW_PATH
                    return "경로 오류";
                case 12: // EW_MODE
                    return "CNC 모드 오류 - 현재 CNC 모드에서 실행 불가";
                case 13: // EW_REJECT
                    return "실행 거부 오류";
                case 14: // EW_DTSRVR
                    return "데이터 서버 오류";
                case 15: // EW_ALARM
                    return "알람 발생 중";
                case 16: // EW_STOP
                    return "CNC가 실행중이 아님";
                case 17: // EW_PASSWD
                    return "보호 데이터 오류 - 비밀번호 필요";
                case -1: // EW_BUSY
                    return "비지 오류 - CNC가 다른 작업 처리 중";
                case -2: // EW_RESET
                    return "리셋 또는 정지 발생";
                case -3: // EW_MMCSYS
                    return "emm386 또는 mmcsys 설치 오류";
                case -4: // EW_PARITY
                    return "공유 RAM 패리티 오류";
                case -5: // EW_SYSTEM
                    return "시스템 오류";
                case -6: // EW_UNEXP
                    return "비정상 오류";
                case -7: // EW_VERSION
                    return "CNC/PMC 버전 불일치";
                case -8: // EW_HANDLE
                    return "Windows 라이브러리 핸들 오류";
                case -9: // EW_HSSB
                    return "HSSB 통신 오류";
                case -10: // EW_SYSTEM2
                    return "시스템 오류 2";
                case -11: // EW_BUS
                    return "버스 오류";
                case -15: // EW_NODLL
                    return "DLL 없음 오류";
                case -16: // EW_SOCKET
                    return "Windows 소켓 오류";
                case -17: // EW_PROTOCOL
                    return "프로토콜 오류";
                default:
                    return $"알 수 없는 오류 코드: {errorCode}";
            }
        }

        private void ShowLogViewer()
        {
            // 기존 로그 뷰어가 있다면 제거
            var existingLogViewer = _centerPanel.Controls.OfType<Panel>().FirstOrDefault(p => p.Name == "LogViewerPanel");
            if (existingLogViewer != null)
            {
                _centerPanel.Controls.Remove(existingLogViewer);
                existingLogViewer.Dispose();
            }

            // 로그 뷰어 패널 생성
            Panel logViewerPanel = new Panel
            {
                Name = "LogViewerPanel",
                Dock = DockStyle.Fill,
                BackColor = Color.White
            };

            // 제목 라벨
            Label titleLabel = new Label
            {
                Text = "로그 기록",
                Font = new Font("맑은 고딕", 16f, FontStyle.Bold),
                Dock = DockStyle.Top,
                Height = 50,
                TextAlign = ContentAlignment.MiddleCenter,
                BackColor = Color.LightSteelBlue
            };

            // 로그 텍스트박스
            TextBox logTextBox = new TextBox
            {
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 10f),
                ReadOnly = true,
                BackColor = Color.White
            };

            // 버튼 패널
            Panel buttonPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 60,
                Padding = new Padding(10)
            };

            Button refreshButton = new Button
            {
                Text = "새로고침",
                Width = 120,
                Height = 40,
                Font = new Font("맑은 고딕", 11f),
                Location = new Point(10, 10)
            };
            refreshButton.Click += (s, e) => LoadLogRecords(logTextBox);

            Button clearButton = new Button
            {
                Text = "로그 삭제",
                Width = 120,
                Height = 40,
                Font = new Font("맑은 고딕", 11f),
                Location = new Point(140, 10)
            };
            clearButton.Click += (s, e) => {
                var result = MessageBox.Show("모든 로그를 삭제하시겠습니까?", "확인",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    // 메모리와 DB에서 모두 삭제
                    _manualOffsetLogs.Clear();
                    _macroVariableLogs.Clear();
                    _logDataService.ClearAllLogs();

                    // 화면 갱신
                    LoadLogRecords(logTextBox);
                    MessageBox.Show("로그가 삭제되었습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            };

            buttonPanel.Controls.Add(refreshButton);
            buttonPanel.Controls.Add(clearButton);

            logViewerPanel.Controls.Add(logTextBox);
            logViewerPanel.Controls.Add(titleLabel);
            logViewerPanel.Controls.Add(buttonPanel);

            _centerPanel.Controls.Add(logViewerPanel);
            logViewerPanel.BringToFront();

            // 초기 로그 로드
            LoadLogRecords(logTextBox);
        }

        private void LoadLogRecords(TextBox logTextBox)
        {
            StringBuilder logBuilder = new StringBuilder();
            logBuilder.AppendLine("=====================================================");
            logBuilder.AppendLine("        수동 옵셋 & 매크로 변수 작업 로그 기록");
            logBuilder.AppendLine("=====================================================");
            logBuilder.AppendLine($"조회 시간: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            logBuilder.AppendLine($"공구 옵셋 로그: {_manualOffsetLogs.Count}개 | 매크로 변수 로그: {_macroVariableLogs.Count}개");
            logBuilder.AppendLine($"총 로그 수: {_manualOffsetLogs.Count + _macroVariableLogs.Count}개");
            logBuilder.AppendLine($"연결된 설비: {_connectionStates.Count(x => x.Value)}개");
            logBuilder.AppendLine("=====================================================");
            logBuilder.AppendLine();

            int totalLogs = _manualOffsetLogs.Count + _macroVariableLogs.Count;

            if (totalLogs == 0)
            {
                logBuilder.AppendLine("[정보] 아직 작업 기록이 없습니다.");
                logBuilder.AppendLine();
                logBuilder.AppendLine("수동 옵셋 페이지에서 다음 작업을 수행하면 기록됩니다:");
                logBuilder.AppendLine("- X축/Z축 마모값 입력 → 공구 옵셋 로그 기록");
                logBuilder.AppendLine("- 매크로 값 입력 → 매크로 변수 로그 기록");
            }
            else
            {
                // 공구 옵셋 로그와 매크로 변수 로그를 합쳐서 시간순으로 정렬
                var allLogs = new List<(DateTime timestamp, string logText)>();

                foreach (var log in _manualOffsetLogs)
                {
                    allLogs.Add((log.Timestamp, log.ToString()));
                }

                foreach (var log in _macroVariableLogs)
                {
                    allLogs.Add((log.Timestamp, log.ToString()));
                }

                // 최신 로그부터 표시 (역순 정렬)
                var sortedLogs = allLogs.OrderByDescending(x => x.timestamp);

                foreach (var log in sortedLogs)
                {
                    logBuilder.AppendLine(log.logText);
                }

                logBuilder.AppendLine();
                logBuilder.AppendLine("-----------------------------------------------------");
                int offsetSuccess = _manualOffsetLogs.Count(x => x.Success);
                int offsetFail = _manualOffsetLogs.Count(x => !x.Success);
                int macroSuccess = _macroVariableLogs.Count(x => x.Success);
                int macroFail = _macroVariableLogs.Count(x => !x.Success);

                logBuilder.AppendLine($"[공구 옵셋] 성공: {offsetSuccess}건 | 실패: {offsetFail}건");
                logBuilder.AppendLine($"[매크로 변수] 성공: {macroSuccess}건 | 실패: {macroFail}건");
                logBuilder.AppendLine($"[전체 합계] 성공: {offsetSuccess + macroSuccess}건 | 실패: {offsetFail + macroFail}건");
            }

            logTextBox.Text = logBuilder.ToString();
        }

        // DB에서 로그 로드
        private void LoadLogsFromDatabase()
        {
            try
            {
                _manualOffsetLogs = _logDataService.LoadAllOffsetLogs();
                _macroVariableLogs = _logDataService.LoadAllMacroLogs();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"로그 로드 중 오류 발생: {ex.Message}", "오류",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // 디버그 로그 파일에 기록
        private void WriteDebugLog(string message)
        {
            try
            {
                string logMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {message}";
                File.AppendAllText(_debugLogPath, logMessage + Environment.NewLine);
            }
            catch
            {
                // 로그 실패는 무시 (프로그램 실행에 영향 없도록)
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            _updateTimer?.Stop();
            _connectionManager?.Dispose();
            base.OnFormClosing(e);
        }
    }

    // 수동 옵셋 로그 클래스
    public class ManualOffsetLog
    {
        public DateTime Timestamp { get; set; }
        public string IpAddress { get; set; }
        public short ToolNumber { get; set; }
        public string Axis { get; set; }
        public double UserInput { get; set; }
        public double BeforeValue { get; set; }
        public double AfterValue { get; set; }
        public double SentValue { get; set; }
        public bool Success { get; set; }
        public string ErrorMessage { get; set; }

        public override string ToString()
        {
            string status = Success ? "[성공]" : "[실패]";
            string error = Success ? "" : $" - 오류: {ErrorMessage}";
            return $"[{Timestamp:yyyy-MM-dd HH:mm:ss}] {status} IP:{IpAddress} | 공구:T{ToolNumber} | 축:{Axis} | " +
                   $"입력:{UserInput:+0.000;-0.000;0.000} | 이전:{BeforeValue:F3} | 이후:{AfterValue:F3} | 전송:{SentValue:F3}{error}";
        }
    }

    // 매크로 변수 로그 클래스
    public class MacroVariableLog
    {
        public DateTime Timestamp { get; set; }
        public string IpAddress { get; set; }
        public short MacroNumber { get; set; }
        public double UserInput { get; set; }
        public double BeforeValue { get; set; }
        public double AfterValue { get; set; }
        public double SentValue { get; set; }
        public bool Success { get; set; }
        public string ErrorMessage { get; set; }

        public override string ToString()
        {
            string status = Success ? "[성공]" : "[실패]";
            string error = Success ? "" : $" - 오류: {ErrorMessage}";
            return $"[{Timestamp:yyyy-MM-dd HH:mm:ss}] {status} IP:{IpAddress} | 매크로:#500-{MacroNumber} | " +
                   $"입력:{UserInput:+0.000;-0.000;0.000} | 이전:{BeforeValue:F3} | 이후:{AfterValue:F3} | 전송:{SentValue:F3}{error}";
        }
    }

    // PMC 제어 로그 클래스
    public class PmcControlLog
    {
        public DateTime Timestamp { get; set; }
        public string IpAddress { get; set; }
        public string ControlType { get; set; }  // "Block Skip" or "Optional Stop"
        public bool BeforeValue { get; set; }
        public bool AfterValue { get; set; }
        public bool Success { get; set; }
        public string ErrorMessage { get; set; }

        public override string ToString()
        {
            string status = Success ? "[성공]" : "[실패]";
            string error = Success ? "" : $" - 오류: {ErrorMessage}";
            string before = BeforeValue ? "ON" : "OFF";
            string after = AfterValue ? "ON" : "OFF";
            return $"[{Timestamp:yyyy-MM-dd HH:mm:ss}] {status} IP:{IpAddress} | {ControlType} | " +
                   $"이전:{before} → 이후:{after}{error}";
        }
    }
}