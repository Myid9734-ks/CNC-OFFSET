using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace FanucFocasTutorial
{
    public class ShiftDataViewForm : Form
    {
        private DateTimePicker _startDatePicker;
        private DateTimePicker _endDatePicker;
        private ComboBox _ipComboBox;
        private ComboBox _shiftTypeComboBox;
        private Button _searchButton;
        private DataGridView _dataGridView;
        private Label _lblTotalRunning;
        private Label _lblAvgOperationRate;
        private Label _lblTotalProduction;
        private TextBox _txtCycleReport;
        private LogDataService _logService;

        public ShiftDataViewForm()
        {
            _logService = new LogDataService();
            InitializeComponents();
            LoadIpAddresses();
            LoadDefaultData();
        }

        private void InitializeComponents()
        {
            this.Text = "근무조 현황";
            this.Size = new Size(1200, 700);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.Sizable;

            // 상단 패널 (조회 조건)
            Panel topPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 80,
                Padding = new Padding(10),
                BackColor = Color.FromArgb(240, 240, 240)
            };

            // 날짜 선택
            Label lblStartDate = new Label
            {
                Text = "시작일:",
                Location = new Point(10, 15),
                AutoSize = true,
                Font = new Font("맑은 고딕", 10f)
            };

            _startDatePicker = new DateTimePicker
            {
                Location = new Point(70, 12),
                Width = 150,
                Format = DateTimePickerFormat.Short,
                Value = DateTime.Today.AddDays(-7)
            };

            Label lblEndDate = new Label
            {
                Text = "종료일:",
                Location = new Point(230, 15),
                AutoSize = true,
                Font = new Font("맑은 고딕", 10f)
            };

            _endDatePicker = new DateTimePicker
            {
                Location = new Point(290, 12),
                Width = 150,
                Format = DateTimePickerFormat.Short,
                Value = DateTime.Today
            };

            // IP 선택
            Label lblIp = new Label
            {
                Text = "설비:",
                Location = new Point(450, 15),
                AutoSize = true,
                Font = new Font("맑은 고딕", 10f)
            };

            _ipComboBox = new ComboBox
            {
                Location = new Point(495, 12),
                Width = 150,
                DropDownStyle = ComboBoxStyle.DropDownList
            };

            // 근무조 선택
            Label lblShiftType = new Label
            {
                Text = "근무조:",
                Location = new Point(655, 15),
                AutoSize = true,
                Font = new Font("맑은 고딕", 10f)
            };

            _shiftTypeComboBox = new ComboBox
            {
                Location = new Point(715, 12),
                Width = 120,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            _shiftTypeComboBox.Items.AddRange(new object[] { "전체", "주간", "야간" });
            _shiftTypeComboBox.SelectedIndex = 0;

            // 조회 버튼
            _searchButton = new Button
            {
                Text = "조회",
                Location = new Point(845, 10),
                Width = 80,
                Height = 30,
                Font = new Font("맑은 고딕", 10f, FontStyle.Bold),
                BackColor = Color.FromArgb(76, 175, 80),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            _searchButton.Click += SearchButton_Click;

            topPanel.Controls.AddRange(new Control[] {
                lblStartDate, _startDatePicker,
                lblEndDate, _endDatePicker,
                lblIp, _ipComboBox,
                lblShiftType, _shiftTypeComboBox,
                _searchButton
            });

            // DataGridView
            _dataGridView = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                RowHeadersVisible = false,
                Font = new Font("맑은 고딕", 9.5f),
                ColumnHeadersHeight = 40,
                RowTemplate = { Height = 35 }
            };

            _dataGridView.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(60, 120, 180);
            _dataGridView.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            _dataGridView.ColumnHeadersDefaultCellStyle.Font = new Font("맑은 고딕", 10f, FontStyle.Bold);
            _dataGridView.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            _dataGridView.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            _dataGridView.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(245, 245, 245);

            // 컬럼 정의
            _dataGridView.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Date",
                HeaderText = "날짜",
                DataPropertyName = "Date",
                Width = 100
            });

            _dataGridView.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "ShiftType",
                HeaderText = "근무조",
                DataPropertyName = "ShiftType",
                Width = 80
            });

            _dataGridView.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "IpAddress",
                HeaderText = "IP 주소",
                DataPropertyName = "IpAddress",
                Width = 120
            });

            _dataGridView.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Running",
                HeaderText = "실가공",
                DataPropertyName = "Running",
                Width = 100
            });

            _dataGridView.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Loading",
                HeaderText = "투입",
                DataPropertyName = "Loading",
                Width = 100
            });

            _dataGridView.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Alarm",
                HeaderText = "알람",
                DataPropertyName = "Alarm",
                Width = 100
            });

            _dataGridView.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Idle",
                HeaderText = "유휴",
                DataPropertyName = "Idle",
                Width = 100
            });

            _dataGridView.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "OperationRate",
                HeaderText = "가동율",
                DataPropertyName = "OperationRate",
                Width = 80
            });

            _dataGridView.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Production",
                HeaderText = "생산수량",
                DataPropertyName = "Production",
                Width = 100
            });

            // 하단 패널 (요약 통계)
            Panel bottomPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 60,
                BackColor = Color.FromArgb(250, 250, 250),
                Padding = new Padding(20, 10, 20, 10)
            };

            // 리포트 텍스트 박스 패널 (데이터 그리드와 리포트 비율: 36% : 64%)
            // 남은 공간 계산: 700(폼 높이) - 80(상단) - 60(하단) = 560px
            // 리포트: 560 * 0.64 = 358px
            Panel reportPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 358,
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Padding = new Padding(10)
            };

            Label lblReport = new Label
            {
                Text = "제품 사이클 리포트:",
                Dock = DockStyle.Top,
                Height = 25,
                Font = new Font("맑은 고딕", 10f, FontStyle.Bold),
                ForeColor = Color.FromArgb(60, 120, 180)
            };

            _txtCycleReport = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Font = new Font("맑은 고딕", 9f),
                BackColor = Color.White,
                BorderStyle = BorderStyle.None,
                Text = "조회 버튼을 클릭하면 제품 사이클 리포트가 표시됩니다."
            };

            reportPanel.Controls.Add(_txtCycleReport);
            reportPanel.Controls.Add(lblReport);

            _lblTotalRunning = new Label
            {
                Text = "총 실가공: 00:00:00",
                Location = new Point(20, 20),
                AutoSize = true,
                Font = new Font("맑은 고딕", 11f, FontStyle.Bold),
                ForeColor = Color.FromArgb(76, 175, 80)
            };

            _lblAvgOperationRate = new Label
            {
                Text = "평균 가동율: 0.0%",
                Location = new Point(250, 20),
                AutoSize = true,
                Font = new Font("맑은 고딕", 11f, FontStyle.Bold),
                ForeColor = Color.FromArgb(33, 150, 243)
            };

            _lblTotalProduction = new Label
            {
                Text = "총 생산: 0개",
                Location = new Point(450, 20),
                AutoSize = true,
                Font = new Font("맑은 고딕", 11f, FontStyle.Bold),
                ForeColor = Color.FromArgb(255, 152, 0)
            };

            bottomPanel.Controls.AddRange(new Control[] {
                _lblTotalRunning, _lblAvgOperationRate, _lblTotalProduction
            });

            // 폼에 컨트롤 추가 (순서 중요: Bottom Dock은 나중에 추가된 것이 위에 표시됨)
            this.Controls.Add(_dataGridView);
            this.Controls.Add(bottomPanel);  // 가장 아래
            this.Controls.Add(reportPanel);  // bottomPanel 위
            this.Controls.Add(topPanel);     // 가장 위
        }

        private void LoadIpAddresses()
        {
            try
            {
                var ipList = _logService.GetAllIpAddresses();
                _ipComboBox.Items.Add("전체");
                _ipComboBox.Items.AddRange(ipList.ToArray());
                _ipComboBox.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"IP 주소 목록 로드 실패: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadDefaultData()
        {
            SearchButton_Click(null, null);
        }

        private void SearchButton_Click(object sender, EventArgs e)
        {
            try
            {
                DateTime startDate = _startDatePicker.Value.Date;
                DateTime endDate = _endDatePicker.Value.Date;

                if (startDate > endDate)
                {
                    MessageBox.Show("시작일은 종료일보다 이전이어야 합니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string selectedIp = _ipComboBox.SelectedItem?.ToString();
                string ipFilter = (selectedIp == "전체" || string.IsNullOrEmpty(selectedIp)) ? null : selectedIp;

                ShiftType? shiftTypeFilter = null;
                if (_shiftTypeComboBox.SelectedItem?.ToString() == "주간")
                    shiftTypeFilter = ShiftType.Day;
                else if (_shiftTypeComboBox.SelectedItem?.ToString() == "야간")
                    shiftTypeFilter = ShiftType.Night;

                // 데이터 조회
                var data = _logService.GetShiftDataByFilter(startDate, endDate, ipFilter, shiftTypeFilter);

                // DataGridView에 바인딩
                var displayData = data.Select(d => new
                {
                    Date = d.ShiftDate.ToString("yyyy-MM-dd"),
                    ShiftType = d.ShiftType == ShiftType.Day ? "주간" : "야간",
                    IpAddress = d.IpAddress,
                    Running = FormatTime(d.RunningSeconds),
                    Loading = FormatTime(d.LoadingSeconds),
                    Alarm = FormatTime(d.AlarmSeconds),
                    Idle = FormatTime(d.IdleSeconds),
                    OperationRate = CalculateOperationRate(d).ToString("F1") + "%",
                    Production = d.ProductionCount.ToString() + "개"
                }).ToList();

                _dataGridView.DataSource = displayData;

                // 요약 통계 업데이트
                UpdateSummary(data);

                // 사이클 리포트 생성 및 표시
                UpdateCycleReport(data, ipFilter, shiftTypeFilter);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"데이터 조회 실패: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void UpdateCycleReport(List<ShiftStateData> shiftData, string ipFilter, ShiftType? shiftTypeFilter)
        {
            try
            {
                var reportText = new System.Text.StringBuilder();
                reportText.AppendLine("=".PadRight(80, '='));
                reportText.AppendLine("제품 사이클 리포트");
                reportText.AppendLine("=".PadRight(80, '='));
                reportText.AppendLine();

                // IP별, 근무조별 리포트 생성
                var ipGroups = shiftData.GroupBy(d => d.IpAddress);
                bool hasReport = false;

                foreach (var ipGroup in ipGroups)
                {
                    string ip = ipGroup.Key;
                    if (ipFilter != null && ip != ipFilter) continue;

                    var shiftGroups = ipGroup.GroupBy(d => new { d.ShiftDate, d.ShiftType });
                    
                    foreach (var shiftGroup in shiftGroups)
                    {
                        if (shiftTypeFilter.HasValue && shiftGroup.Key.ShiftType != shiftTypeFilter.Value) continue;

                        var report = _logService.GenerateCycleReport(ip, shiftGroup.Key.ShiftDate, shiftGroup.Key.ShiftType);
                        
                        if (report.HasData)
                        {
                            hasReport = true;
                            reportText.AppendLine($"[{ip}] {shiftGroup.Key.ShiftDate:yyyy-MM-dd} {shiftGroup.Key.ShiftType}");
                            reportText.AppendLine($"  총 사이클 수: {report.TotalCycles}개");
                            reportText.AppendLine($"  평균 사이클 시간: {FormatTime(report.AvgCycleTime)}");
                            reportText.AppendLine($"  최대 사이클 시간: {FormatTime(report.MaxCycleTime)}");
                            reportText.AppendLine($"  최소 사이클 시간: {FormatTime(report.MinCycleTime)}");
                            reportText.AppendLine($"  특이사항 (평균의 120% 초과): {report.AnomalyCount}건");
                            
                            if (report.Anomalies.Count > 0)
                            {
                                reportText.AppendLine();
                                reportText.AppendLine("  [특이사항 상세]");
                                foreach (var anomaly in report.Anomalies.Take(10)) // 최대 10개만 표시
                                {
                                    reportText.AppendLine($"    - {anomaly.StartTime:HH:mm:ss} ~ {anomaly.EndTime:HH:mm:ss} " +
                                        $"(로딩: {FormatTime(anomaly.LoadingSeconds)}, 가공: {FormatTime(anomaly.RunningSeconds)}, " +
                                        $"총: {FormatTime(anomaly.TotalSeconds)})");
                                }
                                if (report.Anomalies.Count > 10)
                                {
                                    reportText.AppendLine($"    ... 외 {report.Anomalies.Count - 10}건");
                                }
                            }
                            reportText.AppendLine();
                        }
                    }
                }

                if (!hasReport)
                {
                    reportText.AppendLine("리포트 데이터가 없습니다.");
                }

                _txtCycleReport.Text = reportText.ToString();
            }
            catch (Exception ex)
            {
                _txtCycleReport.Text = $"리포트 생성 오류: {ex.Message}";
            }
        }

        private void UpdateSummary(List<ShiftStateData> data)
        {
            if (data.Count == 0)
            {
                _lblTotalRunning.Text = "총 실가공: 00:00:00";
                _lblAvgOperationRate.Text = "평균 가동율: 0.0%";
                _lblTotalProduction.Text = "총 생산: 0개";
                return;
            }

            int totalRunning = data.Sum(d => d.RunningSeconds);
            double avgOperationRate = data.Average(d => CalculateOperationRate(d));
            int totalProduction = data.Sum(d => d.ProductionCount);

            _lblTotalRunning.Text = $"총 실가공: {FormatTime(totalRunning)}";
            _lblAvgOperationRate.Text = $"평균 가동율: {avgOperationRate:F1}%";
            _lblTotalProduction.Text = $"총 생산: {totalProduction}개";
        }

        private string FormatTime(int seconds)
        {
            int hours = seconds / 3600;
            int minutes = (seconds % 3600) / 60;
            int secs = seconds % 60;
            return $"{hours:D2}:{minutes:D2}:{secs:D2}";
        }

        private double CalculateOperationRate(ShiftStateData data)
        {
            int total = data.RunningSeconds + data.LoadingSeconds + data.AlarmSeconds + data.IdleSeconds + data.UnmeasuredSeconds;
            if (total == 0) return 0;
            return (double)data.RunningSeconds / total * 100;
        }
    }
}
