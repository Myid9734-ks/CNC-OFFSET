using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace FanucFocasTutorial
{
    public class MachineRegistrationForm : Form
    {
        private TabControl _tabControl;
        private Button _btnOk;
        private Button _btnCancel;

        // 기본 정보 탭
        private TextBox _txtAlias;
        private TextBox _txtIpAddress;
        private TextBox _txtPort;
        private TextBox _txtLoadingMCode;

        // PMC 상태 모니터링 어드레스 탭
        private TextBox _txtF0_0;
        private TextBox _txtF0_7;
        private TextBox _txtF1_0;
        private TextBox _txtF3_5;
        private TextBox _txtF10;
        private TextBox _txtG4_3;
        private TextBox _txtG5_0;
        private TextBox _txtX8_4;
        private TextBox _txtR854_2;

        // PMC 제어 어드레스 탭
        private TextBox _txtBlockSkipInput;
        private TextBox _txtBlockSkipState;
        private TextBox _txtOptStopInput;
        private TextBox _txtOptStopState;

        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public MainForm.IPConfig Config { get; private set; }
        private bool _isEditMode;

        public MachineRegistrationForm(MainForm.IPConfig existingConfig = null)
        {
            _isEditMode = existingConfig != null;
            Config = existingConfig ?? new MainForm.IPConfig
            {
                Port = 8193,
                LoadingMCode = 0
            };

            InitializeComponent();
            LoadConfigToUI();
        }

        private void InitializeComponent()
        {
            this.Text = _isEditMode ? "설비 정보 수정" : "설비 등록";
            this.Size = new Size(600, 550);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterParent;

            // 탭 컨트롤 생성
            _tabControl = new TabControl
            {
                Dock = DockStyle.Fill,
                Font = new Font("맑은 고딕", 10f)
            };

            // 탭 페이지 생성
            var basicTab = CreateBasicInfoTab();
            var monitoringTab = CreateMonitoringAddressTab();
            var controlTab = CreateControlAddressTab();

            _tabControl.TabPages.Add(basicTab);
            _tabControl.TabPages.Add(monitoringTab);
            _tabControl.TabPages.Add(controlTab);

            // 하단 버튼 패널
            Panel buttonPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 50,
                Padding = new Padding(10)
            };

            _btnOk = new Button
            {
                Text = "확인",
                DialogResult = DialogResult.OK,
                Width = 100,
                Height = 35,
                Location = new Point(this.ClientSize.Width - 220, 7),
                Font = new Font("맑은 고딕", 10f)
            };
            _btnOk.Click += BtnOk_Click;

            _btnCancel = new Button
            {
                Text = "취소",
                DialogResult = DialogResult.Cancel,
                Width = 100,
                Height = 35,
                Location = new Point(this.ClientSize.Width - 110, 7),
                Font = new Font("맑은 고딕", 10f)
            };

            buttonPanel.Controls.Add(_btnOk);
            buttonPanel.Controls.Add(_btnCancel);

            this.Controls.Add(_tabControl);
            this.Controls.Add(buttonPanel);
            this.AcceptButton = _btnOk;
            this.CancelButton = _btnCancel;
        }

        private TabPage CreateBasicInfoTab()
        {
            TabPage tab = new TabPage("기본 정보");
            Panel panel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(20),
                AutoScroll = true
            };

            int yPos = 10;
            int labelWidth = 120;
            int textBoxWidth = 380;
            int rowHeight = 40;

            // 별칭
            panel.Controls.Add(CreateLabel("설비 별칭:", 10, yPos, labelWidth));
            _txtAlias = CreateTextBox(labelWidth + 10, yPos, textBoxWidth);
            panel.Controls.Add(_txtAlias);
            yPos += rowHeight;

            // IP 주소
            panel.Controls.Add(CreateLabel("IP 주소:", 10, yPos, labelWidth));
            _txtIpAddress = CreateTextBox(labelWidth + 10, yPos, textBoxWidth);
            panel.Controls.Add(_txtIpAddress);
            yPos += rowHeight;

            // 포트
            panel.Controls.Add(CreateLabel("포트:", 10, yPos, labelWidth));
            _txtPort = CreateTextBox(labelWidth + 10, yPos, textBoxWidth);
            _txtPort.Text = "8193";
            panel.Controls.Add(_txtPort);
            yPos += rowHeight;

            // 로딩 M코드
            panel.Controls.Add(CreateLabel("로딩 M코드:", 10, yPos, labelWidth));
            _txtLoadingMCode = CreateTextBox(labelWidth + 10, yPos, textBoxWidth);
            _txtLoadingMCode.Text = "0";
            panel.Controls.Add(_txtLoadingMCode);  // 텍스트박스 추가
            Label lblMCodeDesc = new Label
            {
                Text = "* 0 = 로딩 감지 안 함",
                AutoSize = true,
                Location = new Point(labelWidth + 10, yPos + 30),
                ForeColor = Color.Gray,
                Font = new Font("맑은 고딕", 8f)
            };
            panel.Controls.Add(lblMCodeDesc);

            tab.Controls.Add(panel);
            return tab;
        }

        private TabPage CreateMonitoringAddressTab()
        {
            TabPage tab = new TabPage("상태 모니터링 주소");
            Panel panel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(20),
                AutoScroll = true
            };

            int yPos = 10;
            int labelWidth = 180;
            int textBoxWidth = 150;
            int rowHeight = 35;

            // F 영역
            Label lblF = new Label { Text = "[ F 영역 - CNC → PMC ]", Font = new Font("맑은 고딕", 9f, FontStyle.Bold), Location = new Point(10, yPos), AutoSize = true };
            panel.Controls.Add(lblF);
            yPos += 25;

            panel.Controls.Add(CreateLabel("F0.0 (ST_OP 가동):", 10, yPos, labelWidth));
            _txtF0_0 = CreateTextBox(labelWidth + 10, yPos, textBoxWidth);
            panel.Controls.Add(_txtF0_0);
            yPos += rowHeight;

            panel.Controls.Add(CreateLabel("F0.7 (스타트 실행):", 10, yPos, labelWidth));
            _txtF0_7 = CreateTextBox(labelWidth + 10, yPos, textBoxWidth);
            panel.Controls.Add(_txtF0_7);
            yPos += rowHeight;

            panel.Controls.Add(CreateLabel("F1.0 (알람 신호):", 10, yPos, labelWidth));
            _txtF1_0 = CreateTextBox(labelWidth + 10, yPos, textBoxWidth);
            panel.Controls.Add(_txtF1_0);
            yPos += rowHeight;

            panel.Controls.Add(CreateLabel("F3.5 (메모리 모드):", 10, yPos, labelWidth));
            _txtF3_5 = CreateTextBox(labelWidth + 10, yPos, textBoxWidth);
            panel.Controls.Add(_txtF3_5);
            yPos += rowHeight;

            panel.Controls.Add(CreateLabel("F10 (M코드 번호):", 10, yPos, labelWidth));
            _txtF10 = CreateTextBox(labelWidth + 10, yPos, textBoxWidth);
            panel.Controls.Add(_txtF10);
            yPos += rowHeight + 10;

            // G 영역
            Label lblG = new Label { Text = "[ G 영역 - PMC → CNC ]", Font = new Font("맑은 고딕", 9f, FontStyle.Bold), Location = new Point(10, yPos), AutoSize = true };
            panel.Controls.Add(lblG);
            yPos += 25;

            panel.Controls.Add(CreateLabel("G4.3 (M핀 처리):", 10, yPos, labelWidth));
            _txtG4_3 = CreateTextBox(labelWidth + 10, yPos, textBoxWidth);
            panel.Controls.Add(_txtG4_3);
            yPos += rowHeight;

            panel.Controls.Add(CreateLabel("G5.0 (투입 조건):", 10, yPos, labelWidth));
            _txtG5_0 = CreateTextBox(labelWidth + 10, yPos, textBoxWidth);
            panel.Controls.Add(_txtG5_0);
            yPos += rowHeight + 10;

            // X 영역
            Label lblX = new Label { Text = "[ X 영역 - PMC 입력 ]", Font = new Font("맑은 고딕", 9f, FontStyle.Bold), Location = new Point(10, yPos), AutoSize = true };
            panel.Controls.Add(lblX);
            yPos += 25;

            panel.Controls.Add(CreateLabel("X8.4 (비상정지):", 10, yPos, labelWidth));
            _txtX8_4 = CreateTextBox(labelWidth + 10, yPos, textBoxWidth);
            panel.Controls.Add(_txtX8_4);
            yPos += rowHeight + 10;

            // R 영역
            Label lblR = new Label { Text = "[ R 영역 - 내부 릴레이 ]", Font = new Font("맑은 고딕", 9f, FontStyle.Bold), Location = new Point(10, yPos), AutoSize = true };
            panel.Controls.Add(lblR);
            yPos += 25;

            panel.Controls.Add(CreateLabel("R854.2 (알람 PMC-A):", 10, yPos, labelWidth));
            _txtR854_2 = CreateTextBox(labelWidth + 10, yPos, textBoxWidth);
            panel.Controls.Add(_txtR854_2);

            tab.Controls.Add(panel);
            return tab;
        }

        private TabPage CreateControlAddressTab()
        {
            TabPage tab = new TabPage("PMC 제어 주소");
            Panel panel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(20),
                AutoScroll = true
            };

            int yPos = 10;
            int labelWidth = 180;
            int textBoxWidth = 150;
            int rowHeight = 35;

            // Block Skip
            Label lblBlockSkip = new Label { Text = "[ Block Skip ]", Font = new Font("맑은 고딕", 9f, FontStyle.Bold), Location = new Point(10, yPos), AutoSize = true };
            panel.Controls.Add(lblBlockSkip);
            yPos += 25;

            panel.Controls.Add(CreateLabel("버튼 입력 주소:", 10, yPos, labelWidth));
            _txtBlockSkipInput = CreateTextBox(labelWidth + 10, yPos, textBoxWidth);
            panel.Controls.Add(_txtBlockSkipInput);
            yPos += rowHeight;

            panel.Controls.Add(CreateLabel("상태 확인 주소:", 10, yPos, labelWidth));
            _txtBlockSkipState = CreateTextBox(labelWidth + 10, yPos, textBoxWidth);
            panel.Controls.Add(_txtBlockSkipState);
            yPos += rowHeight + 15;

            // Optional Stop
            Label lblOptStop = new Label { Text = "[ Optional Stop ]", Font = new Font("맑은 고딕", 9f, FontStyle.Bold), Location = new Point(10, yPos), AutoSize = true };
            panel.Controls.Add(lblOptStop);
            yPos += 25;

            panel.Controls.Add(CreateLabel("버튼 입력 주소:", 10, yPos, labelWidth));
            _txtOptStopInput = CreateTextBox(labelWidth + 10, yPos, textBoxWidth);
            panel.Controls.Add(_txtOptStopInput);
            yPos += rowHeight;

            panel.Controls.Add(CreateLabel("상태 확인 주소:", 10, yPos, labelWidth));
            _txtOptStopState = CreateTextBox(labelWidth + 10, yPos, textBoxWidth);
            panel.Controls.Add(_txtOptStopState);
            yPos += rowHeight + 15;

            // 설명
            Label lblDesc = new Label
            {
                Text = "* 주소 형식 예: R101.2, F1.0, G4.3 등\n* Type.Address.Bit 형식으로 입력하세요.",
                Location = new Point(10, yPos),
                Size = new Size(500, 40),
                ForeColor = Color.Gray,
                Font = new Font("맑은 고딕", 8f)
            };
            panel.Controls.Add(lblDesc);

            tab.Controls.Add(panel);
            return tab;
        }

        private Label CreateLabel(string text, int x, int y, int width)
        {
            return new Label
            {
                Text = text,
                Location = new Point(x, y + 5),
                Width = width,
                Font = new Font("맑은 고딕", 9f)
            };
        }

        private TextBox CreateTextBox(int x, int y, int width)
        {
            return new TextBox
            {
                Location = new Point(x, y),
                Width = width,
                Font = new Font("맑은 고딕", 9f)
            };
        }

        private void LoadConfigToUI()
        {
            // 기본 정보
            _txtAlias.Text = Config.Alias ?? "";
            _txtIpAddress.Text = Config.IpAddress ?? "";
            _txtPort.Text = Config.Port.ToString();
            _txtLoadingMCode.Text = Config.LoadingMCode.ToString();

            // PMC 상태 모니터링 어드레스
            _txtF0_0.Text = Config.PmcF0_0 ?? "F0.0";
            _txtF0_7.Text = Config.PmcF0_7 ?? "F0.7";
            _txtF1_0.Text = Config.PmcF1_0 ?? "F1.0";
            _txtF3_5.Text = Config.PmcF3_5 ?? "F3.5";
            _txtF10.Text = Config.PmcF10 ?? "F10";
            _txtG4_3.Text = Config.PmcG4_3 ?? "G4.3";
            _txtG5_0.Text = Config.PmcG5_0 ?? "G5.0";
            _txtX8_4.Text = Config.PmcX8_4 ?? "X8.4";
            _txtR854_2.Text = Config.PmcR854_2 ?? "R854.2";

            // PMC 제어 어드레스
            _txtBlockSkipInput.Text = Config.BlockSkipInputAddr ?? "R101.2";
            _txtBlockSkipState.Text = Config.BlockSkipStateAddr ?? "R201.2";
            _txtOptStopInput.Text = Config.OptStopInputAddr ?? "R101.1";
            _txtOptStopState.Text = Config.OptStopStateAddr ?? "R201.1";
        }

        private void BtnOk_Click(object sender, EventArgs e)
        {
            // 유효성 검사
            if (string.IsNullOrWhiteSpace(_txtIpAddress.Text))
            {
                MessageBox.Show("IP 주소를 입력하세요.", "입력 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                _tabControl.SelectedIndex = 0;
                _txtIpAddress.Focus();
                return;
            }

            if (!ushort.TryParse(_txtPort.Text, out ushort port))
            {
                MessageBox.Show("올바른 포트 번호를 입력하세요.", "입력 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                _tabControl.SelectedIndex = 0;
                _txtPort.Focus();
                return;
            }

            if (!int.TryParse(_txtLoadingMCode.Text, out int loadingMCode))
            {
                MessageBox.Show("올바른 로딩 M코드를 입력하세요.", "입력 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                _tabControl.SelectedIndex = 0;
                _txtLoadingMCode.Focus();
                return;
            }

            // Config에 저장
            Config.Alias = _txtAlias.Text.Trim();
            Config.IpAddress = _txtIpAddress.Text.Trim();
            Config.Port = port;
            Config.LoadingMCode = loadingMCode;

            // PMC 상태 모니터링 어드레스
            Config.PmcF0_0 = _txtF0_0.Text.Trim();
            Config.PmcF0_7 = _txtF0_7.Text.Trim();
            Config.PmcF1_0 = _txtF1_0.Text.Trim();
            Config.PmcF3_5 = _txtF3_5.Text.Trim();
            Config.PmcF10 = _txtF10.Text.Trim();
            Config.PmcG4_3 = _txtG4_3.Text.Trim();
            Config.PmcG5_0 = _txtG5_0.Text.Trim();
            Config.PmcX8_4 = _txtX8_4.Text.Trim();
            Config.PmcR854_2 = _txtR854_2.Text.Trim();

            // PMC 제어 어드레스
            Config.BlockSkipInputAddr = _txtBlockSkipInput.Text.Trim();
            Config.BlockSkipStateAddr = _txtBlockSkipState.Text.Trim();
            Config.OptStopInputAddr = _txtOptStopInput.Text.Trim();
            Config.OptStopStateAddr = _txtOptStopState.Text.Trim();

            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
