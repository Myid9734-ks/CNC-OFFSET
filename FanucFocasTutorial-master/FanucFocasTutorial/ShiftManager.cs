using System;
using System.Collections.Generic;

namespace FanucFocasTutorial
{
    /// <summary>
    /// 근무조 타입
    /// </summary>
    public enum ShiftType
    {
        Day,    // 주간
        Night   // 야간
    }

    /// <summary>
    /// 근무조 정보
    /// </summary>
    public class ShiftInfo
    {
        public ShiftType Type { get; set; }
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
        public bool IsExtended { get; set; }  // 연장 근무 여부
        public TimeSpan WorkingHours { get; set; }  // 실 근무 시간 (휴게 제외)

        // 휴게 시간대
        public List<(TimeSpan Start, TimeSpan End)> BreakTimes { get; set; }
    }

    /// <summary>
    /// 근무조 관리 클래스
    /// </summary>
    public class ShiftManager
    {
        // 시간 상수
        private static readonly TimeSpan DAY_START = new TimeSpan(8, 30, 0);      // 08:30
        private static readonly TimeSpan DAY_END_BASIC = new TimeSpan(17, 30, 0); // 17:30
        private static readonly TimeSpan DAY_END_EXTENDED = new TimeSpan(20, 30, 0); // 20:30
        private static readonly TimeSpan NIGHT_START = new TimeSpan(20, 30, 0);   // 20:30
        private static readonly TimeSpan NEXT_DAY_START = new TimeSpan(8, 30, 0); // 다음날 08:30

        // 휴게 시간
        private static readonly TimeSpan LUNCH_START = new TimeSpan(12, 0, 0);    // 12:00
        private static readonly TimeSpan LUNCH_END = new TimeSpan(13, 0, 0);      // 13:00
        private static readonly TimeSpan DINNER_START = new TimeSpan(18, 0, 0);   // 18:00
        private static readonly TimeSpan DINNER_END = new TimeSpan(18, 30, 0);    // 18:30

        /// <summary>
        /// 현재 시각의 근무조 정보 반환
        /// </summary>
        public static ShiftInfo GetCurrentShift(DateTime now)
        {
            TimeSpan currentTime = now.TimeOfDay;

            // 주간 근무조 (08:30 ~ 20:30)
            if (currentTime >= DAY_START && currentTime < NIGHT_START)
            {
                DateTime shiftStart = new DateTime(now.Year, now.Month, now.Day, 8, 30, 0);
                DateTime shiftEnd = new DateTime(now.Year, now.Month, now.Day, 20, 30, 0);

                // 연장 근무 여부는 나중에 자동 판별 (17:30 이후 가동 시)
                bool isExtended = currentTime >= DAY_END_BASIC;

                var breakTimes = new List<(TimeSpan, TimeSpan)>
                {
                    (LUNCH_START, LUNCH_END)  // 점심 1시간
                };

                if (isExtended)
                {
                    breakTimes.Add((DINNER_START, DINNER_END));  // 석식 30분
                }

                TimeSpan workingHours = isExtended
                    ? new TimeSpan(10, 30, 0)  // 12시간 - 1.5시간 = 10.5시간
                    : new TimeSpan(8, 0, 0);   // 9시간 - 1시간 = 8시간

                return new ShiftInfo
                {
                    Type = ShiftType.Day,
                    StartTime = shiftStart,
                    EndTime = shiftEnd,
                    IsExtended = isExtended,
                    WorkingHours = workingHours,
                    BreakTimes = breakTimes
                };
            }
            // 야간 근무조 (20:30 ~ 다음날 08:30)
            else
            {
                DateTime shiftStart;
                DateTime shiftEnd;

                if (currentTime >= NIGHT_START)
                {
                    // 오늘 20:30 ~ 내일 08:30
                    shiftStart = new DateTime(now.Year, now.Month, now.Day, 20, 30, 0);
                    shiftEnd = shiftStart.AddDays(1).Date.AddHours(8).AddMinutes(30);
                }
                else
                {
                    // 어제 20:30 ~ 오늘 08:30
                    shiftStart = new DateTime(now.Year, now.Month, now.Day, 20, 30, 0).AddDays(-1);
                    shiftEnd = new DateTime(now.Year, now.Month, now.Day, 8, 30, 0);
                }

                // 야간 휴게시간 (예: 00:00~00:30, 04:00~05:00)
                var breakTimes = new List<(TimeSpan, TimeSpan)>
                {
                    (new TimeSpan(0, 0, 0), new TimeSpan(0, 30, 0)),   // 30분
                    (new TimeSpan(4, 0, 0), new TimeSpan(5, 0, 0))     // 1시간
                };

                return new ShiftInfo
                {
                    Type = ShiftType.Night,
                    StartTime = shiftStart,
                    EndTime = shiftEnd,
                    IsExtended = false,  // 야간은 항상 고정
                    WorkingHours = new TimeSpan(10, 30, 0),  // 12시간 - 1.5시간 = 10.5시간
                    BreakTimes = breakTimes
                };
            }
        }

        /// <summary>
        /// 특정 시각이 휴게 시간인지 확인
        /// </summary>
        public static bool IsBreakTime(DateTime time, ShiftInfo shift)
        {
            TimeSpan current = time.TimeOfDay;

            foreach (var (start, end) in shift.BreakTimes)
            {
                if (current >= start && current < end)
                    return true;
            }

            return false;
        }

        /// <summary>
        /// 근무조 변경 여부 확인
        /// </summary>
        public static bool IsShiftChanged(ShiftInfo previous, ShiftInfo current)
        {
            if (previous == null) return true;

            return previous.Type != current.Type ||
                   previous.StartTime.Date != current.StartTime.Date;
        }

        /// <summary>
        /// 근무조 문자열 반환
        /// </summary>
        public static string GetShiftDisplayName(ShiftInfo shift)
        {
            if (shift.Type == ShiftType.Day)
            {
                return shift.IsExtended
                    ? "주간 (08:30~20:30, 연장)"
                    : "주간 (08:30~17:30, 기본)";
            }
            else
            {
                return "야간 (20:30~08:30)";
            }
        }

        /// <summary>
        /// 근무조 경과 시간 계산 (휴게 시간 제외)
        /// </summary>
        public static TimeSpan GetElapsedWorkingTime(ShiftInfo shift, DateTime now)
        {
            if (now < shift.StartTime) return TimeSpan.Zero;

            DateTime end = now > shift.EndTime ? shift.EndTime : now;
            TimeSpan total = end - shift.StartTime;

            // 휴게 시간 제외
            TimeSpan breakTime = TimeSpan.Zero;
            foreach (var (start, breakEnd) in shift.BreakTimes)
            {
                DateTime breakStart = shift.StartTime.Date.Add(start);
                DateTime breakEndTime = shift.StartTime.Date.Add(breakEnd);

                // 야간 근무조의 경우 다음날로 넘어가는 휴게 시간 처리
                if (shift.Type == ShiftType.Night && start < NIGHT_START)
                {
                    breakStart = breakStart.AddDays(1);
                    breakEndTime = breakEndTime.AddDays(1);
                }

                if (breakEndTime <= shift.StartTime || breakStart >= end)
                    continue;

                DateTime effectiveStart = breakStart < shift.StartTime ? shift.StartTime : breakStart;
                DateTime effectiveEnd = breakEndTime > end ? end : breakEndTime;

                if (effectiveEnd > effectiveStart)
                {
                    breakTime += effectiveEnd - effectiveStart;
                }
            }

            return total - breakTime;
        }
    }

    /// <summary>
    /// 근무조별 상태 누적 데이터
    /// </summary>
    public class ShiftStateData
    {
        public string IpAddress { get; set; }
        public ShiftType ShiftType { get; set; }
        public DateTime ShiftDate { get; set; }
        public bool IsExtended { get; set; }

        // 상태별 누적 시간 (초)
        public int RunningSeconds { get; set; }
        public int LoadingSeconds { get; set; }
        public int AlarmSeconds { get; set; }
        public int IdleSeconds { get; set; }
        public int UnmeasuredSeconds { get; set; }  // 미측정 시간 (앱 종료 기간)

        // 생산수량
        public int ProductionCount { get; set; }

        // 통합 지표
        public int ActualWorkingSeconds => RunningSeconds;  // 실가공시간
        public int InputSeconds => LoadingSeconds;  // 투입시간 (제품교체 시간만)
        public int StopSeconds => IdleSeconds + AlarmSeconds;  // 정지시간

        public double OperationRate
        {
            get
            {
                // 가동률 = 실가공시간 / 전체시간(미측정 포함) * 100
                int total = RunningSeconds + LoadingSeconds + IdleSeconds + AlarmSeconds + UnmeasuredSeconds;
                return total > 0 ? (double)RunningSeconds / total * 100 : 0;
            }
        }
    }
}
