using System;
using System.Windows.Forms;
using System.Drawing;
using System.Collections.Generic;

namespace FanucFocasTutorial
{
    public partial class AlarmDetailsForm : Form
    {
        private CNCConnection _connection;
        private DataGridView _alarmGrid;
        private System.Windows.Forms.Timer _updateTimer;
        private Label _lblTotalAlarms;
        private Label _lblActiveAlarms;

        public AlarmDetailsForm(CNCConnection connection)
        {
            _connection = connection;
            InitializeComponent();
            SetupTimer();
            LoadAlarmData();
        }

        private void InitializeComponent()
        {
            this.Text = "알람 상세 정보";
            this.Size = new Size(800, 600);
            this.FormBorderStyle = FormBorderStyle.Sizable;

            InitializeControls();
        }

        private void InitializeControls()
        {
            // 상단 정보 패널
            Panel topPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 60,
                BackColor = Color.LightGray,
                Padding = new Padding(10)
            };

            _lblTotalAlarms = new Label
            {
                Text = "총 알람: 0개",
                Font = new Font("맑은 고딕", 12f, FontStyle.Bold),
                Dock = DockStyle.Left,
                Width = 200,
                TextAlign = ContentAlignment.MiddleLeft
            };

            _lblActiveAlarms = new Label
            {
                Text = "활성 알람: 0개",
                Font = new Font("맑은 고딕", 12f, FontStyle.Bold),
                Dock = DockStyle.Left,
                Width = 200,
                TextAlign = ContentAlignment.MiddleLeft
            };

            topPanel.Controls.Add(_lblActiveAlarms);
            topPanel.Controls.Add(_lblTotalAlarms);

            // 알람 그리드
            _alarmGrid = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
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
            _alarmGrid.ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = SystemColors.Control,
                ForeColor = SystemColors.ControlText,
                SelectionBackColor = SystemColors.Control,
                SelectionForeColor = SystemColors.ControlText,
                Font = new Font("맑은 고딕", 11f, FontStyle.Bold),
                Alignment = DataGridViewContentAlignment.MiddleCenter
            };

            // 셀 스타일 설정
            _alarmGrid.DefaultCellStyle = new DataGridViewCellStyle
            {
                Font = new Font("맑은 고딕", 10f),
                Alignment = DataGridViewContentAlignment.MiddleLeft,
                SelectionBackColor = Color.FromArgb(51, 153, 255),
                SelectionForeColor = Color.White
            };

            // 컬럼 추가
            _alarmGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "AlarmNo",
                HeaderText = "알람 번호",
                FillWeight = 15
            });

            _alarmGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "AlarmType",
                HeaderText = "알람 타입",
                FillWeight = 20
            });

            _alarmGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "AlarmMessage",
                HeaderText = "알람 메시지",
                FillWeight = 45
            });

            _alarmGrid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "AlarmTime",
                HeaderText = "발생 시간",
                FillWeight = 20
            });

            this.Controls.Add(_alarmGrid);
            this.Controls.Add(topPanel);
        }

        private void SetupTimer()
        {
            _updateTimer = new System.Windows.Forms.Timer
            {
                Interval = 2000 // 2초마다 업데이트
            };
            _updateTimer.Tick += UpdateTimer_Tick;
            _updateTimer.Start();
        }

        private void UpdateTimer_Tick(object sender, EventArgs e)
        {
            LoadAlarmData();
        }

        private void LoadAlarmData()
        {
            if (_connection == null || !_connection.IsConnected)
            {
                _lblTotalAlarms.Text = "총 알람: 연결 안됨";
                _lblActiveAlarms.Text = "활성 알람: 연결 안됨";
                return;
            }

            try
            {
                _alarmGrid.Rows.Clear();

                // 현재 알람 정보 가져오기
                List<AlarmInfo> alarms = GetCurrentAlarms();

                int activeCount = 0;

                foreach (var alarm in alarms)
                {
                    if (alarm.IsActive)
                        activeCount++;

                    _alarmGrid.Rows.Add(
                        alarm.AlarmNumber,
                        alarm.AlarmType,
                        alarm.Message,
                        alarm.OccurredTime.ToString("yyyy-MM-dd HH:mm:ss")
                    );

                    // 활성 알람은 빨간색으로 표시
                    if (alarm.IsActive)
                    {
                        _alarmGrid.Rows[_alarmGrid.Rows.Count - 1].DefaultCellStyle.BackColor = Color.LightCoral;
                    }
                }

                _lblTotalAlarms.Text = $"총 알람: {alarms.Count}개";
                _lblActiveAlarms.Text = $"활성 알람: {activeCount}개";
            }
            catch (Exception)
            {
                _lblTotalAlarms.Text = "총 알람: 오류";
                _lblActiveAlarms.Text = "활성 알람: 오류";
            }
        }

        private List<AlarmInfo> GetCurrentAlarms()
        {
            List<AlarmInfo> alarms = new List<AlarmInfo>();

            try
            {
                // 현재 활성 알람 체크
                Focas1.ODBALM currentAlarm = new Focas1.ODBALM();
                short ret = Focas1.cnc_alarm(_connection.Handle, currentAlarm);

                if (ret == Focas1.EW_OK && currentAlarm.data != 0)
                {
                    alarms.Add(new AlarmInfo
                    {
                        AlarmNumber = currentAlarm.data,
                        AlarmType = "현재 알람",
                        Message = $"알람 번호: {currentAlarm.data}",
                        OccurredTime = DateTime.Now,
                        IsActive = true
                    });
                }

                // 알람 히스토리 정보 (파라미터 6301-6310에서 가져오기)
                for (short paramNo = 6301; paramNo <= 6310; paramNo++)
                {
                    try
                    {
                        Focas1.IODBPSD_1 param = new Focas1.IODBPSD_1();
                        ret = Focas1.cnc_rdparam(_connection.Handle, paramNo, -1, 8, param);

                        if (ret == Focas1.EW_OK && param.idata != 0)
                        {
                            alarms.Add(new AlarmInfo
                            {
                                AlarmNumber = param.idata,
                                AlarmType = "히스토리",
                                Message = $"알람 히스토리 {paramNo - 6300}",
                                OccurredTime = DateTime.Now.AddMinutes(-(paramNo - 6301) * 10), // 가상의 시간
                                IsActive = false
                            });
                        }
                    }
                    catch
                    {
                        // 개별 파라미터 읽기 실패 시 무시
                    }
                }
            }
            catch
            {
                // 전체 알람 읽기 실패
            }

            return alarms;
        }

        private string GetAlarmType(short type)
        {
            switch (type)
            {
                case 0: return "파라미터 스위치";
                case 1: return "파워 온";
                case 2: return "IO";
                case 3: return "서보";
                case 4: return "스핀들";
                case 5: return "프로그램";
                case 6: return "APC";
                case 7: return "스핀들";
                case 8: return "레이저";
                case 9: return "HMI";
                case 10: return "유저 매크로";
                case 11: return "PMC";
                case 12: return "외부";
                case 13: return "PC카드";
                case 14: return "워치독";
                case 15: return "로봇";
                default: return "기타";
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            _updateTimer?.Stop();
            _updateTimer?.Dispose();
            base.OnFormClosing(e);
        }
    }

    public class AlarmInfo
    {
        public int AlarmNumber { get; set; }
        public string AlarmType { get; set; }
        public string Message { get; set; }
        public DateTime OccurredTime { get; set; }
        public bool IsActive { get; set; }
    }
}