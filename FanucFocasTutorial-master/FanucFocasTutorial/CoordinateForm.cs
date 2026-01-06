using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace FanucFocasTutorial
{
    public class CoordinateForm : Form
    {
        private CNCConnection _connection;
        private string _currentIP;
        private string _machineType; // "MCT" or "LATHE"
        private short _selectedCoordNo = 1; // 기본값 G54 (1)
        private Dictionary<short, string> _coordNames = new Dictionary<short, string>
        {
            { 0, "00 (EXT)" },
            { 1, "01 (G54)" },
            { 2, "02 (G55)" },
            { 3, "03 (G56)" },
            { 4, "04 (G57)" },
            { 5, "05 (G58)" },
            { 6, "06 (G59)" }
        };

        // UI 컨트롤
        private Label _lblCurrentIP;
        private Label _lblMachineType;
        private Label _lblSelectedCoord; // 선택된 좌표계 표시

        // 축 값 표시 및 입력 (X, Y, Z 또는 X, Z, C)
        private Dictionary<string, Label> _currentValueLabels = new Dictionary<string, Label>();
        private Dictionary<string, TextBox> _inputTextBoxes = new Dictionary<string, TextBox>();

        private Button _btnReadAll;
        private Button _btnWrite;

        private DataGridView _dgvAllCoordinates;
        private TextBox _txtHistory;

        public CoordinateForm(CNCConnection connection, string ip, string machineType)
        {
            _connection = connection;
            _currentIP = ip;
            _machineType = machineType;

            InitializeComponents();
            LoadCoordinateData(_selectedCoordNo);

            // 최초 진입 시 전체 좌표계 자동 로드
            BtnReadAll_Click(null, EventArgs.Empty);
        }

        private void InitializeComponents()
        {
            Text = "좌표계 관리";
            Size = new Size(1200, 800);
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(1000, 600);

            // 메인 레이아웃: 좌측 패널 + 우측 패널
            TableLayoutPanel mainLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                Padding = new Padding(10)
            };
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35F)); // 좌측
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 65F)); // 우측

            // === 좌측 패널 ===
            Panel leftPanel = CreateLeftPanel();
            mainLayout.Controls.Add(leftPanel, 0, 0);

            // === 우측 패널 ===
            Panel rightPanel = CreateRightPanel();
            mainLayout.Controls.Add(rightPanel, 1, 0);

            Controls.Add(mainLayout);
        }

        private Panel CreateLeftPanel()
        {
            Panel panel = new Panel { Dock = DockStyle.Fill };

            TableLayoutPanel layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 3,
                ColumnCount = 1
            };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 80F));  // 정보 영역
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));  // 축 값 입력
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 50F));  // 버튼

            // 1. 정보 영역
            Panel infoPanel = CreateInfoPanel();
            layout.Controls.Add(infoPanel, 0, 0);

            // 2. 축 값 입력 영역
            GroupBox axisInputGroup = CreateAxisInputPanel();
            layout.Controls.Add(axisInputGroup, 0, 1);

            // 3. 버튼 영역
            Panel buttonPanel = CreateButtonPanel();
            layout.Controls.Add(buttonPanel, 0, 2);

            panel.Controls.Add(layout);
            return panel;
        }

        private Panel CreateInfoPanel()
        {
            Panel panel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(5)
            };

            _lblCurrentIP = new Label
            {
                Text = $"IP: {_currentIP}",
                Location = new Point(10, 10),
                AutoSize = true,
                Font = new Font("맑은 고딕", 10F, FontStyle.Bold)
            };

            _lblMachineType = new Label
            {
                Text = $"기계 타입: {_machineType} ({GetAxisString()})",
                Location = new Point(10, 40),
                AutoSize = true,
                Font = new Font("맑은 고딕", 9F)
            };

            _lblSelectedCoord = new Label
            {
                Text = $"선택된 좌표계: {_coordNames[_selectedCoordNo]}",
                Location = new Point(10, 60),
                AutoSize = true,
                Font = new Font("맑은 고딕", 9F, FontStyle.Bold),
                ForeColor = Color.Blue
            };

            panel.Controls.Add(_lblCurrentIP);
            panel.Controls.Add(_lblMachineType);
            panel.Controls.Add(_lblSelectedCoord);

            return panel;
        }


        private GroupBox CreateAxisInputPanel()
        {
            GroupBox group = new GroupBox
            {
                Text = "축 값 입력 (상대값)",
                Dock = DockStyle.Fill,
                Font = new Font("맑은 고딕", 9F, FontStyle.Bold),
                Padding = new Padding(10)
            };

            Panel panel = new Panel { Dock = DockStyle.Fill };

            string[] axes = GetAxes();
            int yPos = 25;

            foreach (string axis in axes)
            {
                // 축 레이블
                Label lblAxis = new Label
                {
                    Text = $"{axis}축:",
                    Location = new Point(10, yPos + 5),
                    AutoSize = true,
                    Font = new Font("맑은 고딕", 9F, FontStyle.Bold)
                };

                // 현재 값 레이블
                Label lblCurrent = new Label
                {
                    Text = "현재: 0.000",
                    Location = new Point(50, yPos + 5),
                    AutoSize = true,
                    Font = new Font("맑은 고딕", 9F),
                    ForeColor = Color.Blue
                };
                _currentValueLabels[axis] = lblCurrent;

                // 입력 텍스트박스
                Label lblInput = new Label
                {
                    Text = "입력:",
                    Location = new Point(10, yPos + 35),
                    AutoSize = true,
                    Font = new Font("맑은 고딕", 8F)
                };

                TextBox txtInput = new TextBox
                {
                    Location = new Point(50, yPos + 32),
                    Width = 200,
                    Font = new Font("맑은 고딕", 9F),
                    Text = "" // 빈 텍스트박스
                };
                txtInput.KeyPress += InputTextBox_KeyPress;
                txtInput.KeyDown += InputTextBox_KeyDown;
                _inputTextBoxes[axis] = txtInput;

                // 설명 레이블
                Label lblDesc = new Label
                {
                    Text = "(-0.3 ~ +0.3 mm)",
                    Location = new Point(50, yPos + 60),
                    AutoSize = true,
                    Font = new Font("맑은 고딕", 7F),
                    ForeColor = Color.Gray
                };

                panel.Controls.Add(lblAxis);
                panel.Controls.Add(lblCurrent);
                panel.Controls.Add(lblInput);
                panel.Controls.Add(txtInput);
                panel.Controls.Add(lblDesc);

                yPos += 90;
            }

            group.Controls.Add(panel);
            return group;
        }

        private Panel CreateButtonPanel()
        {
            Panel panel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(5)
            };

            _btnReadAll = new Button
            {
                Text = "전체 읽기",
                Location = new Point(10, 10),
                Size = new Size(135, 35),
                Font = new Font("맑은 고딕", 9F, FontStyle.Bold)
            };
            _btnReadAll.Click += BtnReadAll_Click;

            _btnWrite = new Button
            {
                Text = "입력",
                Location = new Point(155, 10),
                Size = new Size(135, 35),
                Font = new Font("맑은 고딕", 9F, FontStyle.Bold),
                BackColor = Color.LightCoral
            };
            _btnWrite.Click += BtnWrite_Click;

            panel.Controls.Add(_btnReadAll);
            panel.Controls.Add(_btnWrite);

            return panel;
        }

        private Panel CreateRightPanel()
        {
            Panel panel = new Panel { Dock = DockStyle.Fill };

            TableLayoutPanel layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 2,
                ColumnCount = 1
            };
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 60F)); // 좌표계 테이블
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 40F)); // 작업 이력

            // 1. 좌표계 전체 테이블
            GroupBox tableGroup = CreateCoordinateTablePanel();
            layout.Controls.Add(tableGroup, 0, 0);

            // 2. 작업 이력
            GroupBox historyGroup = CreateHistoryPanel();
            layout.Controls.Add(historyGroup, 0, 1);

            panel.Controls.Add(layout);
            return panel;
        }

        private GroupBox CreateCoordinateTablePanel()
        {
            GroupBox group = new GroupBox
            {
                Text = "좌표계 전체 보기",
                Dock = DockStyle.Fill,
                Font = new Font("맑은 고딕", 9F, FontStyle.Bold),
                Padding = new Padding(10)
            };

            _dgvAllCoordinates = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                RowHeadersVisible = false,
                BackgroundColor = Color.White,
                Font = new Font("맑은 고딕", 9F),
                ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
                {
                    Font = new Font("맑은 고딕", 9F, FontStyle.Bold),
                    Alignment = DataGridViewContentAlignment.MiddleCenter
                },
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    Alignment = DataGridViewContentAlignment.MiddleCenter
                }
            };

            // 컬럼 설정
            _dgvAllCoordinates.Columns.Add(new DataGridViewTextBoxColumn { Name = "좌표계", HeaderText = "좌표계", Width = 100 });
            foreach (string axis in GetAxes())
            {
                _dgvAllCoordinates.Columns.Add(new DataGridViewTextBoxColumn
                {
                    Name = axis,
                    HeaderText = $"{axis}축 (mm)"
                });
            }

            // 테이블 행 클릭 이벤트
            _dgvAllCoordinates.CellClick += DgvAllCoordinates_CellClick;

            group.Controls.Add(_dgvAllCoordinates);
            return group;
        }

        private GroupBox CreateHistoryPanel()
        {
            GroupBox group = new GroupBox
            {
                Text = "작업 이력",
                Dock = DockStyle.Fill,
                Font = new Font("맑은 고딕", 9F, FontStyle.Bold),
                Padding = new Padding(10)
            };

            _txtHistory = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Font = new Font("Consolas", 8F),
                BackColor = Color.WhiteSmoke
            };

            group.Controls.Add(_txtHistory);
            return group;
        }

        // === 이벤트 핸들러 ===

        private void DgvAllCoordinates_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            // 헤더 클릭은 무시
            if (e.RowIndex < 0) return;

            try
            {
                // 클릭한 행의 좌표계 번호 가져오기
                DataGridViewRow row = _dgvAllCoordinates.Rows[e.RowIndex];
                string coordName = row.Cells[0].Value?.ToString();

                // 좌표계 이름에서 번호 추출 (예: "01 (G54)" → 1)
                short coordNo = _coordNames.FirstOrDefault(x => x.Value == coordName).Key;

                // 선택된 좌표계 업데이트
                _selectedCoordNo = coordNo;
                _lblSelectedCoord.Text = $"선택된 좌표계: {_coordNames[_selectedCoordNo]}";

                // 좌표 데이터 로드
                LoadCoordinateData(coordNo);

                // 선택된 행 하이라이트
                _dgvAllCoordinates.ClearSelection();
                _dgvAllCoordinates.Rows[e.RowIndex].Selected = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"좌표계 선택 오류: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnReadAll_Click(object sender, EventArgs e)
        {
            try
            {
                _dgvAllCoordinates.Rows.Clear();

                foreach (var kvp in _coordNames.OrderBy(x => x.Key))
                {
                    short coordNo = kvp.Key;
                    string coordName = kvp.Value;

                    Dictionary<string, double> values = _connection.GetAllWorkCoordinates(coordNo, _machineType);

                    DataGridViewRow row = new DataGridViewRow();
                    row.CreateCells(_dgvAllCoordinates);
                    row.Cells[0].Value = coordName;

                    string[] axes = GetAxes();
                    for (int i = 0; i < axes.Length; i++)
                    {
                        row.Cells[i + 1].Value = values[axes[i]].ToString("F3");
                    }

                    _dgvAllCoordinates.Rows.Add(row);
                }

                AddHistory($"전체 좌표계 읽기 완료");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"전체 읽기 오류: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnWrite_Click(object sender, EventArgs e)
        {
            try
            {
                string[] axes = GetAxes();
                List<string> updatedAxes = new List<string>();

                foreach (string axis in axes)
                {
                    if (!_inputTextBoxes.ContainsKey(axis)) continue;

                    string inputText = _inputTextBoxes[axis].Text.Trim();
                    if (string.IsNullOrEmpty(inputText)) continue;

                    if (!double.TryParse(inputText, out double value))
                    {
                        MessageBox.Show($"{axis}축 입력값이 올바르지 않습니다: {inputText}", "입력 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    // 범위 체크: -0.3 ~ +0.3
                    if (value < -0.3 || value > 0.3)
                    {
                        MessageBox.Show($"{axis}축 입력값은 -0.3 ~ +0.3 범위여야 합니다: {value}", "범위 오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    // 입력값이 0이면 건너뛰기
                    if (Math.Abs(value) < 0.0001)
                    {
                        continue;
                    }

                    // 현재 값 읽기
                    short axisNo = GetAxisNumber(axis);
                    double currentValue = _connection.GetWorkCoordinate(_selectedCoordNo, axisNo);

                    // 새로운 값 = 현재 값 + 입력 값
                    double newValue = currentValue + value;

                    // CNC에 쓰기
                    bool success = _connection.SetWorkCoordinate(_selectedCoordNo, axisNo, newValue);

                    if (success)
                    {
                        updatedAxes.Add($"{axis}:{value:+0.000;-0.000}");

                        // 입력 텍스트박스 초기화
                        _inputTextBoxes[axis].Text = "";
                    }
                    else
                    {
                        MessageBox.Show($"{axis}축 쓰기 실패", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }

                if (updatedAxes.Count > 0)
                {
                    AddHistory($"{_coordNames[_selectedCoordNo]} 업데이트: {string.Join(", ", updatedAxes)}");

                    // 현재 좌표 다시 읽기
                    LoadCoordinateData(_selectedCoordNo);

                    // 전체 좌표계 테이블 자동 업데이트
                    BtnReadAll_Click(sender, e);

                    MessageBox.Show("좌표계 업데이트 완료", "완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("업데이트할 값이 없습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"쓰기 오류: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void InputTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            // 숫자, 소수점, 마이너스, 백스페이스만 허용
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && e.KeyChar != '.' && e.KeyChar != '-')
            {
                e.Handled = true;
            }

            // 소수점은 하나만
            TextBox txt = sender as TextBox;
            if (e.KeyChar == '.' && txt != null && txt.Text.Contains("."))
            {
                e.Handled = true;
            }

            // 마이너스는 맨 앞에만
            if (e.KeyChar == '-' && txt != null && txt.SelectionStart != 0)
            {
                e.Handled = true;
            }
        }

        private void InputTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            // 엔터키 입력 시 입력 버튼 클릭 실행
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                e.SuppressKeyPress = true; // 경고음 방지
                BtnWrite_Click(sender, EventArgs.Empty);
            }
        }

        // === 헬퍼 메서드 ===

        private void LoadCoordinateData(short coordNo)
        {
            try
            {
                Dictionary<string, double> values = _connection.GetAllWorkCoordinates(coordNo, _machineType);

                foreach (var kvp in values)
                {
                    string axis = kvp.Key;
                    double value = kvp.Value;

                    if (_currentValueLabels.ContainsKey(axis))
                    {
                        _currentValueLabels[axis].Text = $"현재: {value:F3}";
                    }
                }

                AddHistory($"{_coordNames[coordNo]} 읽기 완료");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"좌표 읽기 오류: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string[] GetAxes()
        {
            // LATHE는 C축 제거 (X, Z만 사용)
            return _machineType == "LATHE" ? new[] { "X", "Z" } : new[] { "X", "Y", "Z" };
        }

        private string GetAxisString()
        {
            return _machineType == "LATHE" ? "X, Z" : "X, Y, Z";
        }

        private short GetAxisNumber(string axis)
        {
            // LATHE의 경우 축 번호가 다름: X=1, Z=2
            if (_machineType == "LATHE")
            {
                switch (axis)
                {
                    case "X": return 1; // LATHE X축은 1번
                    case "Z": return 2; // LATHE Z축은 2번
                    default: return 0;
                }
            }
            else // MCT
            {
                switch (axis)
                {
                    case "X": return 0;
                    case "Y": return 1;
                    case "Z": return 2;
                    default: return 0;
                }
            }
        }

        private void AddHistory(string message)
        {
            string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            string logMessage = $"[{timestamp}] {message}";

            if (_txtHistory.Text.Length > 0)
            {
                _txtHistory.AppendText(Environment.NewLine);
            }
            _txtHistory.AppendText(logMessage);

            // 스크롤을 맨 아래로
            _txtHistory.SelectionStart = _txtHistory.Text.Length;
            _txtHistory.ScrollToCaret();
        }

        public void UpdateConnection(CNCConnection connection, string ip, string machineType)
        {
            _connection = connection;
            _currentIP = ip;
            _machineType = machineType;

            _lblCurrentIP.Text = $"IP: {_currentIP}";
            _lblMachineType.Text = $"기계 타입: {_machineType} ({GetAxisString()})";

            // IP 변경 시 현재 선택된 좌표 다시 읽기
            LoadCoordinateData(_selectedCoordNo);

            // 전체 좌표계 테이블도 자동 업데이트
            BtnReadAll_Click(null, EventArgs.Empty);

            AddHistory($"IP 변경: {ip} ({machineType})");
        }
    }
}
