using System;
using System.Drawing;
using System.Windows.Forms;

namespace FanucFocasTutorial.Forms
{
    public class MessageBoxForm : Form
    {
        private static Point? LastPosition = null;
        private System.Windows.Forms.Timer autoCloseTimer;
        private Label messageLabel;
        private int remainingSeconds = 10;

        public MessageBoxForm(string message)
        {
            InitializeComponents(message);
            SetupTimer();
        }

        private void InitializeComponents(string message)
        {
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Size = new Size(400, 200);
            this.Text = "보정 처리 결과";
            this.StartPosition = FormStartPosition.CenterScreen;
            this.TopMost = true;
            
            if (LastPosition.HasValue)
            {
                this.StartPosition = FormStartPosition.Manual;
                this.Location = LastPosition.Value;
            }

            messageLabel = new Label
            {
                Text = message + $"\n\n{remainingSeconds}초 후 자동으로 닫힙니다.",
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Fill,
                Font = new Font("맑은 고딕", 9f)
            };

            this.Controls.Add(messageLabel);

            this.FormClosing += (s, e) =>
            {
                LastPosition = this.Location;
            };
        }

        private void SetupTimer()
        {
            autoCloseTimer = new System.Windows.Forms.Timer();
            autoCloseTimer.Interval = 1000; // 1초
            autoCloseTimer.Tick += (s, e) =>
            {
                remainingSeconds--;
                messageLabel.Text = messageLabel.Text.Replace($"{remainingSeconds + 1}초", $"{remainingSeconds}초");
                
                if (remainingSeconds <= 0)
                {
                    autoCloseTimer.Stop();
                    this.Close();
                }
            };
            autoCloseTimer.Start();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && autoCloseTimer != null)
            {
                autoCloseTimer.Stop();
                autoCloseTimer.Dispose();
            }
            base.Dispose(disposing);
        }
    }
} 