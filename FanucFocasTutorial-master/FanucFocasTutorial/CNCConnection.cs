using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.IO;
using System.Runtime.InteropServices;

namespace FanucFocasTutorial
{
    public class CNCConnection : IDisposable
    {
        // DLL 로딩 진단용 Win32 API
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr LoadLibrary(string lpFileName);

        [DllImport("kernel32.dll")]
        private static extern uint GetLastError();

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool SetDllDirectory(string lpPathName);

        private ushort _handle;
        private readonly string _ipAddress;
        private readonly ushort _port;
        private readonly int _timeout;
        private bool _isConnected;
        private readonly double _scale = 1000;

        // 생산 카운트 관련 필드
        private int _productionCount = 0;
        private int _goodCount = 0;
        private int _ngCount = 0;
        private DateTime _lastM99Time = DateTime.MinValue;
        private TimeSpan _currentCycleTime = TimeSpan.Zero;
        private TimeSpan _averageCycleTime = TimeSpan.Zero;
        private List<TimeSpan> _cycleTimes = new List<TimeSpan>();
        private string _lastSequenceNumber = "";
        private bool _isInM99Sequence = false;

        public string IpAddress => _ipAddress;
        public bool IsConnected => _isConnected;
        public string ConnectionStatus { get; private set; }
        public ushort Handle => _handle;

        public CNCConnection(string ipAddress, ushort port = 8193, int timeout = 2)
        {
            _ipAddress = ipAddress;
            _port = port;
            _timeout = timeout;
            _handle = 0;
            _isConnected = false;
            ConnectionStatus = "연결되지 않음";
        }

        public bool Connect()
        {
            try
            {
                string appDir = AppDomain.CurrentDomain.BaseDirectory;

                // DLL 검색 경로 설정
                SetDllDirectory(appDir);

                LogDiagnostic($"[{DateTime.Now:HH:mm:ss}] 연결 시도: {_ipAddress}:{_port}");

                if (_handle != 0)
                {
                    Focas1.cnc_freelibhndl(_handle);
                    _handle = 0;
                }

                short ret = Focas1.cnc_allclibhndl3(_ipAddress, _port, _timeout, out _handle);

                LogDiagnostic($"cnc_allclibhndl3 반환값: {ret}");

                if (ret == Focas1.EW_OK)
                {
                    _isConnected = true;
                    ConnectionStatus = "연결됨";
                    LogDiagnostic($"✓ CNC 연결 성공");
                    return true;
                }
                else
                {
                    _isConnected = false;
                    string errorMsg = GetFocasErrorMessage(ret);
                    ConnectionStatus = $"연결 실패 (에러 코드: {ret}) - {errorMsg}";
                    LogDiagnostic($"✗ CNC 연결 실패: {errorMsg}");
                    return false;
                }
            }
            catch (Exception ex)
            {
                _isConnected = false;
                ConnectionStatus = $"연결 오류: {ex.Message}";
                LogDiagnostic($"✗ 예외 발생: {ex.Message}");
                LogDiagnostic($"  스택 추적: {ex.StackTrace}");
                return false;
            }
        }

        private void LogDiagnostic(string message)
        {
            string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "cnc_diagnostic.log");
            try
            {
                File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}\n");
            }
            catch
            {
                // 로그 실패 무시
            }
        }

        private string GetFocasErrorMessage(short errorCode)
        {
            switch (errorCode)
            {
                case -15:
                    return "EW_NODLL: DLL 파일이 없거나 로드할 수 없음";
                case -16:
                    return "EW_BUS: CNC와 통신 버스 오류";
                case -8:
                    return "EW_HANDLE: 잘못된 핸들";
                case -17:
                    return "EW_SYSTEM: 시스템 오류";
                case 0:
                    return "EW_OK: 성공";
                default:
                    return $"알 수 없는 에러 ({errorCode})";
            }
        }

        public void Disconnect()
        {
            if (_handle != 0)
            {
                Focas1.cnc_freelibhndl(_handle);
                _handle = 0;
            }
            _isConnected = false;
            ConnectionStatus = "연결 해제됨";
        }

        public string GetMode()
        {
            if (_handle == 0)
                return "사용불가";

            Focas1.ODBST mode = new Focas1.ODBST();
            short ret = Focas1.cnc_statinfo(_handle, mode);

            if (ret == Focas1.EW_OK)
                return GetModeString(mode.aut);
            return "사용불가";
        }

        public string GetStatus()
        {
            if (_handle == 0)
                return "사용불가";

            Focas1.ODBST status = new Focas1.ODBST();
            short ret = Focas1.cnc_statinfo(_handle, status);

            if (ret == Focas1.EW_OK)
                return GetStatusString(status.run);
            return "사용불가";
        }

        private string GetModeString(short mode)
        {
            switch (mode)
            {
                case 0: return "MDI";
                case 1: return "MEM";
                case 2: return "****";
                case 3: return "EDIT";
                case 4: return "HND";
                case 5: return "JOG";
                case 6: return "T-JOG";
                case 7: return "T-HND";
                case 8: return "INC";
                case 9: return "REF";
                case 10: return "RMT";
                default: return "사용불가";
            }
        }

        private string GetStatusString(short status)
        {
            switch (status)
            {
                case 0: return "STOP";
                case 1: return "HOLD";
                case 2: return "START";
                case 3: return "MSTR";
                case 4: return "RESTART";
                default: return "사용불가";
            }
        }

        public double GetToolOffset(short toolNo, short axis, bool isGeometry)
        {
            if (_handle == 0) return 0.0;

            Focas1.ODBTOFS offset_data = new Focas1.ODBTOFS();
            
            short type;
            if (axis == 0) // X축
                type = (short)(isGeometry ? 1 : 0);  // X축 형상(1) 또는 마모(0)
            else // Z축
                type = (short)(isGeometry ? 3 : 2);  // Z축 형상(3) 또는 마모(2)

            short length = 10;

            try 
            {
                short ret = Focas1.cnc_rdtofs(_handle, toolNo, type, length, offset_data);
                if (ret != Focas1.EW_OK) return 0.0;
                return offset_data.data / _scale;  // 1/1000mm 단위를 mm 단위로 변환
            }
            catch
            {
                return 0.0;
            }
        }

        public bool SetToolOffset(short toolNo, short axis, bool isGeometry, double value)
        {
            if (_handle == 0) return false;

            try 
            {
                short type;
                if (axis == 0) // X축
                    type = (short)(isGeometry ? 1 : 0);  // X축 형상(1) 또는 마모(0)
                else // Z축
                    type = (short)(isGeometry ? 3 : 2);  // Z축 형상(3) 또는 마모(2)

                int ivalue = (int)(value * _scale); // mm 단위를 1/1000mm 단위로 변환
                short length = 10;

                short ret = Focas1.cnc_wrtofs(_handle, toolNo, type, length, ivalue);
                return ret == Focas1.EW_OK;
            }
            catch
            {
                return false;
            }
        }

        public string GetOpSignal()
        {
            if (_handle == 0) return "사용불가";

            short addr_kind = 1; // F
            short data_type = 0; // Byte
            ushort start = 0;
            ushort end = 0;
            ushort data_length = 9; // 8 + N
            Focas1.IODBPMC0 pmc = new Focas1.IODBPMC0();

            short ret = Focas1.pmc_rdpmcrng(_handle, addr_kind, data_type, start, end, data_length, pmc);
            if (ret != Focas1.EW_OK) return "사용불가";

            return pmc.cdata[0].GetBit(7) ? "ON" : "OFF";
        }

        // 교대조 구분 메서드
        public string GetCurrentShift()
        {
            DateTime now = DateTime.Now;
            TimeSpan currentTime = now.TimeOfDay;
            TimeSpan dayStart = new TimeSpan(8, 30, 0);    // 08:30
            TimeSpan dayEnd = new TimeSpan(20, 30, 0);     // 20:30

            if (currentTime >= dayStart && currentTime < dayEnd)
            {
                return "Day Shift";
            }
            else
            {
                return "Night Shift";
            }
        }

        // 설비 상태 분류 (Auto/Manual/Alarm/Idle)
        public string GetEquipmentStatus()
        {
            if (_handle == 0) return "Idle";

            try
            {
                Focas1.ODBST status = new Focas1.ODBST();
                short ret = Focas1.cnc_statinfo(_handle, status);

                if (ret != Focas1.EW_OK) return "Idle";

                // 알람 상태 체크
                if (HasAlarm())
                {
                    return "Alarm";
                }

                // 운전 상태 체크
                if (status.run == 2) // START 상태
                {
                    // 자동 모드인지 체크
                    if (status.aut == 1) // MEM 모드
                    {
                        return "Auto";
                    }
                    else
                    {
                        return "Manual";
                    }
                }
                else if (status.run == 1) // HOLD 상태
                {
                    return "Manual";
                }
                else
                {
                    return "Idle";
                }
            }
            catch
            {
                return "Idle";
            }
        }

        // 알람 존재 여부 체크
        public bool HasAlarm()
        {
            if (_handle == 0) return false;

            try
            {
                Focas1.ODBALM alarm = new Focas1.ODBALM();
                short ret = Focas1.cnc_alarm(_handle, alarm);

                return ret == Focas1.EW_OK && alarm.data != 0;
            }
            catch
            {
                return false;
            }
        }

        // 현재 알람 정보 가져오기
        public string GetCurrentAlarm()
        {
            if (_handle == 0) return "알람 없음";

            try
            {
                Focas1.ODBALM alarm = new Focas1.ODBALM();
                short ret = Focas1.cnc_alarm(_handle, alarm);

                if (ret == Focas1.EW_OK && alarm.data != 0)
                {
                    return $"알람: {alarm.data}";
                }
                else
                {
                    return "알람 없음";
                }
            }
            catch
            {
                return "알람 확인 불가";
            }
        }

        // 스핀들 부하율 가져오기
        public double GetSpindleLoad()
        {
            if (_handle == 0) return 0.0;

            try
            {
                Focas1.ODBSPLOAD spload = new Focas1.ODBSPLOAD();
                short type = 0;
                short datano = 1;
                short ret = Focas1.cnc_rdspmeter(_handle, type, ref datano, spload);

                if (ret == Focas1.EW_OK)
                {
                    return spload.spload1.spload.data;
                }
                return 0.0;
            }
            catch
            {
                return 0.0;
            }
        }

        // 스핀들 속도 가져오기
        public int GetSpindleSpeed()
        {
            if (_handle == 0) return 0;

            try
            {
                Focas1.ODBACT speed = new Focas1.ODBACT();
                short ret = Focas1.cnc_acts(_handle, speed);

                if (ret == Focas1.EW_OK)
                {
                    return speed.data;
                }
                return 0;
            }
            catch
            {
                return 0;
            }
        }

        // 이송 속도 가져오기
        public int GetFeedrate()
        {
            if (_handle == 0) return 0;

            try
            {
                Focas1.ODBACT feed = new Focas1.ODBACT();
                short ret = Focas1.cnc_actf(_handle, feed);

                if (ret == Focas1.EW_OK)
                {
                    return feed.data;
                }
                return 0;
            }
            catch
            {
                return 0;
            }
        }

        // 현재 실행 중인 프로그램 번호
        public string GetCurrentProgram()
        {
            if (_handle == 0) return "N/A";

            try
            {
                Focas1.ODBEXEPRG exeprog = new Focas1.ODBEXEPRG();
                short ret = Focas1.cnc_exeprgname(_handle, exeprog);

                if (ret == Focas1.EW_OK)
                {
                    return new string(exeprog.name);
                }
                return "N/A";
            }
            catch
            {
                return "N/A";
            }
        }

        // 운전 시간 (분 단위)
        public int GetOperatingTime()
        {
            if (_handle == 0) return 0;

            try
            {
                Focas1.IODBPSD_1 param = new Focas1.IODBPSD_1();
                short ret = Focas1.cnc_rdparam(_handle, 6750, -1, 8, param);

                if (ret == Focas1.EW_OK)
                {
                    return param.idata;
                }
                return 0;
            }
            catch
            {
                return 0;
            }
        }

        // 절삭 시간 (분 단위)
        public int GetCuttingTime()
        {
            if (_handle == 0) return 0;

            try
            {
                Focas1.IODBPSD_1 param = new Focas1.IODBPSD_1();
                short ret = Focas1.cnc_rdparam(_handle, 6751, -1, 8, param);

                if (ret == Focas1.EW_OK)
                {
                    return param.idata;
                }
                return 0;
            }
            catch
            {
                return 0;
            }
        }

        // M99 감지 및 생산 카운트 메서드
        public void CheckM99AndUpdateProduction()
        {
            if (_handle == 0) return;

            try
            {
                // 현재 실행 중인 프로그램 정보 가져오기
                Focas1.ODBPRO prog = new Focas1.ODBPRO();
                short ret = Focas1.cnc_rdprgnum(_handle, prog);

                if (ret == Focas1.EW_OK)
                {
                    string currentSeq = prog.mdata.ToString();

                    // 현재 실행 중인 블록 내용 확인
                    if (IsM99Executed(currentSeq))
                    {
                        if (!_isInM99Sequence)
                        {
                            _isInM99Sequence = true;
                            OnM99Detected();
                        }
                    }
                    else
                    {
                        _isInM99Sequence = false;
                    }

                    _lastSequenceNumber = currentSeq;
                }
            }
            catch
            {
                // M99 감지 실패
            }
        }

        private bool IsM99Executed(string sequenceNumber)
        {
            if (_handle == 0) return false;

            try
            {
                // 현재 실행 블록의 NC 코드 읽기
                Focas1.ODBEXEPRG exeprog = new Focas1.ODBEXEPRG();
                short ret = Focas1.cnc_exeprgname(_handle, exeprog);

                if (ret == Focas1.EW_OK)
                {
                    // 프로그램이 시작 부분으로 돌아갔는지 확인 (M99의 특징)
                    if (_lastSequenceNumber != "" &&
                        !string.IsNullOrEmpty(sequenceNumber) &&
                        int.TryParse(sequenceNumber, out int currentSeq) &&
                        int.TryParse(_lastSequenceNumber, out int lastSeq))
                    {
                        // 시퀀스 번호가 큰 값에서 작은 값으로 돌아간 경우 (M99 실행됨)
                        return currentSeq < lastSeq && lastSeq - currentSeq > 100;
                    }
                }
            }
            catch
            {
                // 에러 발생
            }

            return false;
        }

        private void OnM99Detected()
        {
            DateTime now = DateTime.Now;

            // 사이클 타임 계산
            if (_lastM99Time != DateTime.MinValue)
            {
                _currentCycleTime = now - _lastM99Time;
                _cycleTimes.Add(_currentCycleTime);

                // 최근 10개 사이클의 평균 계산
                if (_cycleTimes.Count > 10)
                {
                    _cycleTimes.RemoveAt(0);
                }

                _averageCycleTime = TimeSpan.FromMilliseconds(_cycleTimes.Average(t => t.TotalMilliseconds));
            }

            _lastM99Time = now;

            // 생산 카운트 증가
            _productionCount++;

            // 알람 상태에 따라 양품/불량 구분
            if (HasAlarm())
            {
                _ngCount++;
            }
            else
            {
                _goodCount++;
            }
        }

        // 생산 정보 조회 메서드들
        public int GetProductionCount() => _productionCount;
        public int GetGoodCount() => _goodCount;
        public int GetNGCount() => _ngCount;
        public TimeSpan GetCurrentCycleTime() => _currentCycleTime;
        public TimeSpan GetAverageCycleTime() => _averageCycleTime;

        // 생산 수량 증가 (상태 전환 시 사용)
        public void IncrementProduction()
        {
            _productionCount++;
            _goodCount++;  // 기본적으로 양품으로 카운트
        }

        // 생산 카운트 리셋 (교대조 변경 시 사용)
        public void ResetProductionCount()
        {
            _productionCount = 0;
            _goodCount = 0;
            _ngCount = 0;
            _cycleTimes.Clear();
            _lastM99Time = DateTime.MinValue;
        }

        public bool CheckConnection()
        {
            if (_handle == 0) return false;

            try
            {
                Focas1.ODBST status = new Focas1.ODBST();
                short ret = Focas1.cnc_statinfo(_handle, status);
                
                if (ret != Focas1.EW_OK)
                {
                    Disconnect();
                    return false;
                }
                
                return true;
            }
            catch
            {
                Disconnect();
                return false;
            }
        }

        public void Dispose()
        {
            Disconnect();
        }
    }
} 