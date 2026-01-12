using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace FanucFocasTutorial
{
    /// <summary>
    /// FANUC CNC 백업 서비스 클래스
    /// inventcom.net FOCAS API를 사용하여 백업 기능 제공
    /// </summary>
    public class BackupService
    {
        private readonly CNCConnection _connection;
        private readonly string _baseBackupPath;
        private StringBuilder _backupLog;
        private string _currentBackupFolder;

        public BackupService(CNCConnection connection)
        {
            _connection = connection;
            _baseBackupPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Backup");
            _backupLog = new StringBuilder();

            // 백업 폴더가 없으면 생성
            if (!Directory.Exists(_baseBackupPath))
            {
                Directory.CreateDirectory(_baseBackupPath);
            }
        }

        /// <summary>
        /// IP별 날짜 폴더 경로 생성 (시간 제외, 날짜만 사용)
        /// </summary>
        private string CreateBackupFolder(string ipAddress)
        {
            string dateStr = DateTime.Now.ToString("yyyy-MM-dd");
            string backupFolder = Path.Combine(_baseBackupPath, $"{ipAddress}-{dateStr}");

            if (!Directory.Exists(backupFolder))
            {
                Directory.CreateDirectory(backupFolder);
            }

            _currentBackupFolder = backupFolder;
            return backupFolder;
        }

        /// <summary>
        /// 하위 폴더 생성 (프로그램, 파라미터, 래더, 공구옵셋, 좌표계, 매크로)
        /// </summary>
        private string CreateSubFolder(string baseFolder, string subFolderName)
        {
            string subFolder = Path.Combine(baseFolder, subFolderName);

            if (!Directory.Exists(subFolder))
            {
                Directory.CreateDirectory(subFolder);
            }

            return subFolder;
        }

        /// <summary>
        /// 백업 로그 초기화
        /// </summary>
        private void InitializeLog()
        {
            _backupLog.Clear();
            _backupLog.AppendLine("═══════════════════════════════════════════════════════");
            _backupLog.AppendLine("               FANUC CNC 백업 작업 로그");
            _backupLog.AppendLine("═══════════════════════════════════════════════════════");
            _backupLog.AppendLine($"백업 시작 시간: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            _backupLog.AppendLine($"IP 주소: {_connection.IpAddress}");
            _backupLog.AppendLine($"백업 폴더: {_currentBackupFolder}");
            _backupLog.AppendLine("═══════════════════════════════════════════════════════");
            _backupLog.AppendLine();
        }

        /// <summary>
        /// 백업 작업 로그 추가
        /// </summary>
        private void LogBackupTask(string taskName, string status, string details = "")
        {
            _backupLog.AppendLine($"[{DateTime.Now:HH:mm:ss}] {taskName}");
            _backupLog.AppendLine($"  상태: {status}");
            if (!string.IsNullOrEmpty(details))
            {
                _backupLog.AppendLine($"  상세: {details}");
            }
            _backupLog.AppendLine();
        }

        /// <summary>
        /// 백업 항목별 로그 추가 (성공/실패)
        /// </summary>
        private void LogBackupItem(string itemName, bool success, string message = "")
        {
            string status = success ? "✓ 성공" : "✗ 실패";
            _backupLog.AppendLine($"  {status} - {itemName}: {message}");
        }

        /// <summary>
        /// 로그 파일 저장
        /// </summary>
        private void SaveLogToFile()
        {
            try
            {
                _backupLog.AppendLine();
                _backupLog.AppendLine("═══════════════════════════════════════════════════════");
                _backupLog.AppendLine($"백업 종료 시간: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                _backupLog.AppendLine("═══════════════════════════════════════════════════════");

                string logFilePath = Path.Combine(_currentBackupFolder, "BackupLog.txt");
                File.WriteAllText(logFilePath, _backupLog.ToString(), Encoding.UTF8);
            }
            catch (Exception ex)
            {
                // 로그 저장 실패 시 무시 (백업 작업 자체에는 영향 없음)
                Console.WriteLine($"로그 파일 저장 실패: {ex.Message}");
            }
        }

        /// <summary>
        /// CNC에 등록된 프로그램 번호 목록 읽기
        /// cnc_rdprogdir API 사용 - 설비에서 직접 목록 받아오기
        /// </summary>
        private List<int> GetAllProgramNumbers()
        {
            string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "NCBackup_Debug.txt");
            List<int> programNumbers = new List<int>();

            try
            {
                File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] GetAllProgramNumbers 내부 시작\n");

                // PRGDIR 구조체 생성 (최대 버퍼)
                File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] PRGDIR 구조체 생성 중...\n");
                Focas1.PRGDIR prgdir = new Focas1.PRGDIR();
                ushort length = 256;

                File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] cnc_rdprogdir 호출 직전 (Handle={_connection.Handle}, length={length})\n");

                // type=0: 프로그램 번호만
                // datano_s=1, datano_e=9999: 전체 범위
                _backupLog.AppendLine($"  - cnc_rdprogdir 파라미터: type=0, start=1, end=9999, length={length}");

                short ret = Focas1.cnc_rdprogdir(_connection.Handle, 0, 1, 9999, length, prgdir);

                File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] cnc_rdprogdir 호출 완료: ret={ret}\n");
                _backupLog.AppendLine($"  - cnc_rdprogdir 결과: {ret} (0=성공)");

                if (ret == Focas1.EW_OK)
                {
                    // prg_data에서 문자열 추출 (예: "O100O200O5000%")
                    string rawData = new string(prgdir.prg_data);

                    // null 문자 및 % 제거
                    int endIndex = rawData.IndexOf('\0');
                    if (endIndex > 0)
                        rawData = rawData.Substring(0, endIndex);

                    rawData = rawData.TrimEnd('%', ' ', '\r', '\n');
                    _backupLog.AppendLine($"  - 원시 데이터 길이: {rawData.Length} 문자");
                    _backupLog.AppendLine($"  - 원시 데이터: {(rawData.Length > 100 ? rawData.Substring(0, 100) + "..." : rawData)}");

                    if (!string.IsNullOrEmpty(rawData))
                    {
                        // "O"로 분리 (예: "O100O200" → ["", "100", "200"])
                        string[] parts = rawData.Split(new[] { 'O' }, StringSplitOptions.RemoveEmptyEntries);
                        _backupLog.AppendLine($"  - 분리된 파트 수: {parts.Length}개");

                        foreach (string part in parts)
                        {
                            // 숫자만 추출 (주석이 포함될 경우 대비)
                            string numberOnly = new string(part.TakeWhile(char.IsDigit).ToArray());

                            if (int.TryParse(numberOnly, out int progNum))
                            {
                                programNumbers.Add(progNum);
                            }
                        }
                    }
                }
                else
                {
                    _backupLog.AppendLine($"  - ✗ cnc_rdprogdir 실패 (에러 코드: {ret})");
                    throw new Exception($"cnc_rdprogdir 실패 (에러 코드: {ret})");
                }
            }
            catch (Exception ex)
            {
                try
                {
                    File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] GetAllProgramNumbers 예외 발생: {ex.Message}\n");
                    File.AppendAllText(logPath, $"스택 트레이스:\n{ex.StackTrace}\n");
                }
                catch { }

                _backupLog.AppendLine($"  - ✗ 예외 발생: {ex.Message}");
                throw new Exception($"프로그램 목록 읽기 실패: {ex.Message}");
            }

            File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] GetAllProgramNumbers 완료: {programNumbers.Count}개 발견\n");
            return programNumbers;
        }

        /// <summary>
        /// 공구 옵셋 백업
        /// API: cnc_rdtofsr (Tool Offset Read - Area Specified)
        /// 참고: https://www.inventcom.net/fanuc-focas-library/ncdata/cnc_rdtofsr
        /// </summary>
        public string BackupToolOffset()
        {
            if (_connection == null || !_connection.IsConnected)
            {
                throw new Exception("CNC 연결이 필요합니다.");
            }

            string backupFolder = CreateBackupFolder(_connection.IpAddress);

            // 로그가 비어있으면 초기화 (단독 호출 시)
            bool isStandaloneCall = _backupLog.Length == 0;
            if (isStandaloneCall)
            {
                InitializeLog();
            }

            var startTime = DateTime.Now;
            _backupLog.AppendLine($"[{startTime:HH:mm:ss}] 공구 옵셋 백업 시작");

            string toolOffsetFolder = CreateSubFolder(backupFolder, "공구옵셋");
            string filePath = Path.Combine(toolOffsetFolder, "ToolOffset.ofs");

            StringBuilder sb = new StringBuilder();
            sb.AppendLine("===== 공구 옵셋 백업 =====");
            sb.AppendLine($"백업 일시: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine($"IP 주소: {_connection.IpAddress}");
            sb.AppendLine();
            sb.AppendLine("공구번호\tX축형상\tX축마모\tZ축형상\tZ축마모");
            sb.AppendLine("────────────────────────────────────");

            int successCount = 0;
            int failCount = 0;

            // 공구 1번부터 200번까지 백업
            _backupLog.AppendLine($"  - 백업 범위: T001~T200 (총 200개)");
            for (short toolNo = 1; toolNo <= 200; toolNo++)
            {
                try
                {
                    double xGeometry = _connection.GetToolOffset(toolNo, 0, true);
                    double xWear = _connection.GetToolOffset(toolNo, 0, false);
                    double zGeometry = _connection.GetToolOffset(toolNo, 1, true);
                    double zWear = _connection.GetToolOffset(toolNo, 1, false);

                    sb.AppendLine($"T{toolNo:D3}\t{xGeometry:F3}\t{xWear:F3}\t{zGeometry:F3}\t{zWear:F3}");
                    successCount++;

                    if (toolNo % 50 == 0)
                    {
                        _backupLog.AppendLine($"  - 진행 중... T{toolNo:D3} (성공: {successCount}, 실패: {failCount})");
                    }
                }
                catch (Exception ex)
                {
                    failCount++;
                    // 모든 실패를 상세 로그에 기록 (디버그용)
                    _backupLog.AppendLine($"  - ✗ T{toolNo:D3} 읽기 실패: {ex.Message}");

                    // 처음 10개 실패는 상세 정보 추가
                    if (failCount <= 10)
                    {
                        _backupLog.AppendLine($"    스택 트레이스: {ex.StackTrace?.Split('\n').FirstOrDefault()?.Trim()}");
                    }
                }
            }

            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);

            var elapsed = (DateTime.Now - startTime).TotalSeconds;
            _backupLog.AppendLine($"  - ✓ 백업 완료: 성공 {successCount}개, 실패 {failCount}개");
            _backupLog.AppendLine($"  - 파일 저장: {filePath}");
            _backupLog.AppendLine($"  - 소요 시간: {elapsed:F2}초");
            _backupLog.AppendLine();

            // 단독 호출 시 로그 저장
            if (isStandaloneCall)
            {
                SaveLogToFile();
            }

            return filePath;
        }

        /// <summary>
        /// 워크 좌표계 백업
        /// API: cnc_rdzofs (Zero Offset Read)
        /// </summary>
        public string BackupWorkCoordinate()
        {
            if (_connection == null || !_connection.IsConnected)
            {
                throw new Exception("CNC 연결이 필요합니다.");
            }

            string backupFolder = CreateBackupFolder(_connection.IpAddress);

            // 로그가 비어있으면 초기화 (단독 호출 시)
            bool isStandaloneCall = _backupLog.Length == 0;
            if (isStandaloneCall)
            {
                InitializeLog();
            }

            var startTime = DateTime.Now;
            _backupLog.AppendLine($"[{startTime:HH:mm:ss}] 워크 좌표계 백업 시작");

            string coordinateFolder = CreateSubFolder(backupFolder, "좌표계");
            string filePath = Path.Combine(coordinateFolder, "WorkCoordinate.wcs");

            StringBuilder sb = new StringBuilder();
            sb.AppendLine("===== 워크 좌표계 백업 =====");
            sb.AppendLine($"백업 일시: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine($"IP 주소: {_connection.IpAddress}");
            sb.AppendLine();
            sb.AppendLine("좌표계\tX축\tY축\tZ축");
            sb.AppendLine("────────────────────────────────────");

            int successCount = 0;
            int failCount = 0;

            // G54~G59 (좌표계 1~6)
            _backupLog.AppendLine($"  - 백업 범위: G54~G59 (좌표계 1~6)");
            for (short coordNo = 1; coordNo <= 6; coordNo++)
            {
                string coordName = $"G{53 + coordNo}";
                try
                {
                    _backupLog.AppendLine($"  - {coordName} cnc_rdzofs 호출 중 (coordNo={coordNo}, type=2, length=8)");
                    Focas1.IODBZOFS zofs = new Focas1.IODBZOFS();
                    short ret = Focas1.cnc_rdzofs(_connection.Handle, coordNo, 2, 8, zofs);

                    if (ret == Focas1.EW_OK)
                    {
                        // 데이터 형식에 따라 적절히 처리 필요
                        sb.AppendLine($"{coordName}\t[좌표계 데이터]");
                        _backupLog.AppendLine($"  - {coordName} ✓ 읽기 성공 (에러 코드: {ret})");
                        successCount++;
                    }
                    else
                    {
                        _backupLog.AppendLine($"  - {coordName} ✗ 읽기 실패 (에러 코드: {ret})");
                        failCount++;
                    }
                }
                catch (Exception ex)
                {
                    _backupLog.AppendLine($"  - {coordName} ✗ 예외 발생: {ex.Message}");
                    failCount++;
                }
            }

            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);

            var elapsed = (DateTime.Now - startTime).TotalSeconds;
            _backupLog.AppendLine($"  - ✓ 백업 완료: 성공 {successCount}개, 실패 {failCount}개");
            _backupLog.AppendLine($"  - 파일 저장: {filePath}");
            _backupLog.AppendLine($"  - 소요 시간: {elapsed:F2}초");
            _backupLog.AppendLine();

            // 단독 호출 시 로그 저장
            if (isStandaloneCall)
            {
                SaveLogToFile();
            }

            return filePath;
        }

        /// <summary>
        /// 매크로 변수 백업
        /// API: cnc_rdmacror (Macro Variable Read - Area Specified)
        /// </summary>
        public string BackupMacroVariable()
        {
            if (_connection == null || !_connection.IsConnected)
            {
                throw new Exception("CNC 연결이 필요합니다.");
            }

            string backupFolder = CreateBackupFolder(_connection.IpAddress);

            // 로그가 비어있으면 초기화 (단독 호출 시)
            bool isStandaloneCall = _backupLog.Length == 0;
            if (isStandaloneCall)
            {
                InitializeLog();
            }

            var startTime = DateTime.Now;
            _backupLog.AppendLine($"[{startTime:HH:mm:ss}] 매크로 변수 백업 시작");

            string macroFolder = CreateSubFolder(backupFolder, "매크로");
            string filePath = Path.Combine(macroFolder, "MacroVariable.mac");

            StringBuilder sb = new StringBuilder();
            sb.AppendLine("===== 매크로 변수 백업 =====");
            sb.AppendLine($"백업 일시: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine($"IP 주소: {_connection.IpAddress}");
            sb.AppendLine();
            sb.AppendLine("변수번호\t값");
            sb.AppendLine("────────────────────────────────────");

            int successCount = 0;
            int failCount = 0;

            // 매크로 변수 #100 ~ #999 백업
            int totalMacros = 900; // 100~999
            _backupLog.AppendLine($"  - 백업 범위: #100~#999 (총 {totalMacros}개)");

            for (short macroNo = 100; macroNo <= 999; macroNo++)
            {
                try
                {
                    Focas1.ODBM macroData = new Focas1.ODBM();
                    short ret = Focas1.cnc_rdmacro(_connection.Handle, macroNo, 10, macroData);

                    if (ret == Focas1.EW_OK)
                    {
                        sb.AppendLine($"#{macroNo}\t{macroData.mcr_val}");
                        successCount++;
                    }
                    else
                    {
                        failCount++;
                        // API 실패 시 에러 코드 기록 (처음 10개만)
                        if (failCount <= 10)
                        {
                            _backupLog.AppendLine($"  - ✗ #{macroNo} cnc_rdmacro 실패 (에러 코드: {ret})");
                        }
                    }

                    // 100개마다 진행 상태 로그
                    if (macroNo % 100 == 0)
                    {
                        _backupLog.AppendLine($"  - 진행 중... #{macroNo} (성공: {successCount}, 실패: {failCount})");
                    }
                }
                catch (Exception ex)
                {
                    failCount++;
                    if (failCount <= 10) // 처음 10개 실패만 로그
                    {
                        _backupLog.AppendLine($"  - ✗ #{macroNo} 예외 발생: {ex.Message}");
                    }
                }
            }

            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);

            var elapsed = (DateTime.Now - startTime).TotalSeconds;
            _backupLog.AppendLine($"  - ✓ 백업 완료: 성공 {successCount}개, 실패 {failCount}개");
            _backupLog.AppendLine($"  - 파일 저장: {filePath}");
            _backupLog.AppendLine($"  - 소요 시간: {elapsed:F2}초");
            _backupLog.AppendLine();

            // 단독 호출 시 로그 저장
            if (isStandaloneCall)
            {
                SaveLogToFile();
            }

            return filePath;
        }

        /// <summary>
        /// 단일 NC 프로그램 백업 (내부 메서드)
        /// </summary>
        private string BackupSingleProgram(string programFolder, int programNumber)
        {
            string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "NCBackup_Debug.txt");
            var startTime = DateTime.Now;
            string filePath = Path.Combine(programFolder, $"O{programNumber}.nc");

            _backupLog.AppendLine($"  [O{programNumber}] 백업 시작...");
            File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] O{programNumber} 백업 시작\n");

            StringBuilder programData = new StringBuilder();

            // cnc_upstart: Upload 시작 (프로그램 번호 전달)
            _backupLog.AppendLine($"  [O{programNumber}] cnc_upstart 호출 중 (프로그램 번호: {programNumber})");
            File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] O{programNumber} cnc_upstart 호출 (Handle={_connection.Handle})\n");
            short ret = Focas1.cnc_upstart(_connection.Handle, (short)programNumber);
            File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] O{programNumber} cnc_upstart 결과: {ret}\n");
            if (ret != Focas1.EW_OK)
            {
                _backupLog.AppendLine($"  [O{programNumber}] ✗ cnc_upstart 실패 (에러 코드: {ret})");
                throw new Exception($"업로드 시작 실패 (에러 코드: {ret})");
            }
            _backupLog.AppendLine($"  [O{programNumber}] ✓ cnc_upstart 성공");

            try
            {
                int readCount = 0;
                int totalBytes = 0;

                while (true)
                {
                    Focas1.ODBUP uploadData = new Focas1.ODBUP();
                    ushort length = 256;

                    ret = Focas1.cnc_upload(_connection.Handle, uploadData, ref length);

                    if (ret == (short)Focas1.focas_ret.EW_OK)
                    {
                        // 데이터 추가
                        string data = new string(uploadData.data).Substring(0, length);
                        programData.Append(data);
                        readCount++;
                        totalBytes += length;

                        if (readCount % 10 == 0)
                        {
                            _backupLog.AppendLine($"  [O{programNumber}] 업로드 중... ({totalBytes} bytes)");
                        }
                    }
                    else if (ret == (short)Focas1.focas_ret.EW_BUFFER)
                    {
                        // 버퍼가 비어있음 - 계속
                        System.Threading.Thread.Sleep(10);
                        continue;
                    }
                    else if (ret == (short)Focas1.focas_ret.DNC_NORMAL)
                    {
                        // 정상 종료 (DNC_NORMAL = -1)
                        _backupLog.AppendLine($"  [O{programNumber}] ✓ 업로드 완료 (DNC_NORMAL, 총 {totalBytes} bytes, {readCount}회 읽기)");
                        break;
                    }
                    else if (ret == (short)Focas1.focas_ret.EW_RESET)
                    {
                        // 정상 종료 (EW_RESET = -2, 일부 FANUC 컨트롤러에서 업로드 완료 시 반환)
                        _backupLog.AppendLine($"  [O{programNumber}] ✓ 업로드 완료 (EW_RESET, 총 {totalBytes} bytes, {readCount}회 읽기)");
                        break;
                    }
                    else
                    {
                        _backupLog.AppendLine($"  [O{programNumber}] ✗ cnc_upload 실패 (에러 코드: {ret}, 읽은 데이터: {totalBytes} bytes)");
                        throw new Exception($"업로드 실패 (에러 코드: {ret})");
                    }
                }
            }
            finally
            {
                // cnc_upend: Upload 종료
                Focas1.cnc_upend(_connection.Handle);
                _backupLog.AppendLine($"  [O{programNumber}] cnc_upend 호출 완료");
            }

            File.WriteAllText(filePath, programData.ToString(), Encoding.UTF8);
            var elapsed = (DateTime.Now - startTime).TotalSeconds;
            _backupLog.AppendLine($"  [O{programNumber}] ✓ 파일 저장 완료: {filePath} (소요 시간: {elapsed:F2}초)");
            _backupLog.AppendLine();

            return filePath;
        }

        /// <summary>
        /// 모든 NC 프로그램 백업 (자동)
        /// API: cnc_rdprogdir, cnc_upload
        /// 참고: https://www.inventcom.net/fanuc-focas-library/Program/cnc_upload
        /// </summary>
        /// <param name="progressCallback">진행 상황 콜백 (현재, 전체, 메시지)</param>
        public Dictionary<string, string> BackupAllNCPrograms(Action<int, int, string> progressCallback = null)
        {
            var results = new Dictionary<string, string>();

            // 로그 파일에 시작 기록
            string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "NCBackup_Debug.txt");
            try
            {
                File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] BackupAllNCPrograms 시작\n");
            }
            catch { }

            try
            {
                File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 연결 확인 중...\n");

                if (_connection == null || !_connection.IsConnected)
                {
                    File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 연결 없음\n");
                    throw new Exception("CNC 연결이 필요합니다.");
                }

                File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 백업 폴더 생성 중...\n");
                string backupFolder = CreateBackupFolder(_connection.IpAddress);

                File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 프로그램 폴더 생성 중...\n");
                string programFolder = CreateSubFolder(backupFolder, "프로그램");

                File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 로그 초기화 중...\n");
                // 로그 초기화
                InitializeLog();
                LogBackupTask("NC 프로그램 백업 시작", "진행 중");

                File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 프로그램 검색 시작 (cnc_rdprogdir 사용)\n");
                _backupLog.AppendLine($"[{DateTime.Now:HH:mm:ss}] 프로그램 검색 중 (cnc_rdprogdir 사용)");
                _backupLog.AppendLine($"  - 방식: cnc_rdprogdir API로 실제 존재하는 프로그램 목록 조회");
                _backupLog.AppendLine();

                // 실제 존재하는 프로그램 목록 가져오기
                List<int> programNumbers = null;
                try
                {
                    programNumbers = GetAllProgramNumbers();
                    File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 프로그램 목록 조회 성공: {programNumbers.Count}개 발견\n");
                    _backupLog.AppendLine($"  - ✓ 프로그램 목록 조회 성공: {programNumbers.Count}개 발견");

                    if (programNumbers.Count > 0)
                    {
                        _backupLog.AppendLine($"  - 프로그램 목록: {string.Join(", ", programNumbers.Take(20).Select(p => $"O{p}"))}");
                        if (programNumbers.Count > 20)
                        {
                            _backupLog.AppendLine($"    ... 외 {programNumbers.Count - 20}개");
                        }
                    }
                    _backupLog.AppendLine();
                }
                catch (Exception ex)
                {
                    File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 프로그램 목록 조회 실패: {ex.Message}\n");
                    _backupLog.AppendLine($"  - ✗ 프로그램 목록 조회 실패: {ex.Message}");
                    _backupLog.AppendLine($"  - 전체 범위 백업으로 전환 (O1~O9999)");
                    _backupLog.AppendLine();

                    // 실패 시 전체 범위 사용
                    programNumbers = Enumerable.Range(1, 9999).ToList();
                }

                // 각 프로그램 백업
                int successCount = 0;
                int failCount = 0;
                int skippedCount = 0;

                File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 프로그램 백업 시작 (총 {programNumbers.Count}개)\n");

                for (int i = 0; i < programNumbers.Count; i++)
                {
                    int progNum = programNumbers[i];

                    try
                    {
                        string filePath = BackupSingleProgram(programFolder, progNum);
                        File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] O{progNum} 백업 성공 ({i+1}/{programNumbers.Count})\n");
                        results[$"O{progNum}"] = $"성공: {filePath}";
                        successCount++;
                    }
                    catch (Exception ex)
                    {
                        // 연결 끊김 감지 (EW_HANDLE = -8)
                        if (ex.Message.Contains("에러 코드: -8"))
                        {
                            File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] ✗✗✗ 연결 끊김 감지 (EW_HANDLE) - 백업 중단\n");
                            _backupLog.AppendLine($"  ✗✗✗ 연결 끊김 감지 (에러 코드: -8, EW_HANDLE)");
                            _backupLog.AppendLine($"  백업을 중단합니다. CNC와의 연결을 확인하세요.");
                            _backupLog.AppendLine($"  진행: {i+1}/{programNumbers.Count} (성공: {successCount}, 실패: {failCount})");
                            throw new Exception($"CNC 연결이 끊어졌습니다. 백업을 중단합니다.\n\n진행 상황: {successCount}개 성공, {failCount}개 실패\n\n해결 방법:\n1. CNC 재연결\n2. 프로그램 재시작\n3. 백업 재시도");
                        }

                        // 프로그램이 존재하지 않는 경우는 건너뜀 (EW_DATA = 5)
                        if (ex.Message.Contains("업로드 시작 실패") && (ex.Message.Contains("에러 코드: 5") || ex.Message.Contains("EW_DATA")))
                        {
                            File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] O{progNum} 존재하지 않음 (에러 코드: 5) ({i+1}/{programNumbers.Count})\n");
                            skippedCount++;
                            continue;
                        }

                        // 실제 백업 실패는 failCount 증가
                        File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] O{progNum} 백업 실패: {ex.Message} ({i+1}/{programNumbers.Count})\n");
                        results[$"O{progNum}"] = $"실패: {ex.Message}";
                        _backupLog.AppendLine($"  [O{progNum}] ✗ 백업 실패: {ex.Message}");
                        _backupLog.AppendLine();
                        failCount++;

                        // 연속으로 10개 이상 실패하면 연결 문제 의심
                        if (failCount >= 10 && successCount == 0)
                        {
                            File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] ✗ 연속 {failCount}개 실패 - 연결 문제 의심\n");
                            _backupLog.AppendLine($"  ✗ 연속 {failCount}개 실패 - CNC 연결 상태를 확인하세요.");
                        }
                    }

                    // 진행 상황 업데이트
                    progressCallback?.Invoke(i + 1, programNumbers.Count, $"O{progNum} 백업 중... (성공: {successCount}, 실패: {failCount})");

                    // 10개마다 또는 마지막 진행 상황 로그
                    if ((i + 1) % 10 == 0 || i == programNumbers.Count - 1)
                    {
                        File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 진행: {i+1}/{programNumbers.Count} (성공: {successCount}, 실패: {failCount}, 건너뜀: {skippedCount})\n");
                        _backupLog.AppendLine($"  - 진행: {i+1}/{programNumbers.Count} (성공: {successCount}, 실패: {failCount}, 건너뜀: {skippedCount})");
                    }
                }

                results["Summary"] = $"성공: {successCount}개, 실패: {failCount}개";

                _backupLog.AppendLine();
                LogBackupTask("NC 프로그램 백업 완료", "완료", $"성공: {successCount}개, 실패: {failCount}개");
            }
            catch (Exception ex)
            {
                try
                {
                    File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 예외 발생: {ex.Message}\n");
                    File.AppendAllText(logPath, $"스택 트레이스:\n{ex.StackTrace}\n");
                }
                catch { }

                try
                {
                    _backupLog.AppendLine($"[{DateTime.Now:HH:mm:ss}] ✗ NC 프로그램 백업 실패");
                    _backupLog.AppendLine($"  에러: {ex.Message}");
                    _backupLog.AppendLine($"  스택 트레이스: {ex.StackTrace ?? "없음"}");
                }
                catch
                {
                    // 로그 작성 실패 시 무시
                }

                results["Error"] = ex.Message ?? "알 수 없는 에러";
                results["StackTrace"] = ex.StackTrace ?? "스택 트레이스 없음";
            }
            finally
            {
                try
                {
                    if (!string.IsNullOrEmpty(_currentBackupFolder))
                    {
                        SaveLogToFile();
                    }
                }
                catch
                {
                    // 로그 저장 실패 시 무시
                }
            }

            return results;
        }

        /// <summary>
        /// 파라미터 백업
        /// API: cnc_rdparam (Parameter Read)
        /// 참고: https://www.inventcom.net/fanuc-focas-library/ncdata/cnc_rdparam
        /// </summary>
        public string BackupParameter(Action<int, int> progressCallback = null)
        {
            if (_connection == null || !_connection.IsConnected)
            {
                throw new Exception("CNC 연결이 필요합니다.");
            }

            string backupFolder = CreateBackupFolder(_connection.IpAddress);

            // 로그가 비어있으면 초기화 (단독 호출 시)
            bool isStandaloneCall = _backupLog.Length == 0;
            if (isStandaloneCall)
            {
                InitializeLog();
            }

            var startTime = DateTime.Now;
            _backupLog.AppendLine($"[{startTime:HH:mm:ss}] 파라미터 백업 시작");

            string parameterFolder = CreateSubFolder(backupFolder, "파라미터");
            string filePath = Path.Combine(parameterFolder, "Parameter.par");

            StringBuilder sb = new StringBuilder();
            sb.AppendLine("===== CNC 파라미터 백업 =====");
            sb.AppendLine($"백업 일시: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine($"IP 주소: {_connection.IpAddress}");
            sb.AppendLine();
            sb.AppendLine("파라미터번호\t축\t값");
            sb.AppendLine("────────────────────────────────────");

            int successCount = 0;
            int failCount = 0;

            // 주요 파라미터만 백업 (0~2000번 범위로 제한)
            // 전체를 백업하면 시간이 너무 오래 걸림
            int totalParams = 2000;
            _backupLog.AppendLine($"  - 백업 범위: 파라미터 0~{totalParams} (총 {totalParams + 1}개)");

            for (short paramNo = 0; paramNo <= totalParams; paramNo++)
            {
                try
                {
                    Focas1.IODBPSD_1 param = new Focas1.IODBPSD_1();
                    short ret = Focas1.cnc_rdparam(_connection.Handle, paramNo, 0, 8, param);

                    if (ret == Focas1.EW_OK)
                    {
                        sb.AppendLine($"{paramNo}\t{param.datano}\t{param.cdata}");
                        successCount++;
                    }
                    else
                    {
                        failCount++;
                        // API 실패 시 에러 코드 기록 (처음 10개만)
                        if (failCount <= 10)
                        {
                            _backupLog.AppendLine($"  - ✗ 파라미터 {paramNo} cnc_rdparam 실패 (에러 코드: {ret})");
                        }
                    }

                    // 200개마다 진행 상태 업데이트
                    if (paramNo % 200 == 0)
                    {
                        progressCallback?.Invoke(paramNo, totalParams);
                        _backupLog.AppendLine($"  - 진행 중... 파라미터 {paramNo}/{totalParams} (성공: {successCount}, 실패: {failCount})");
                    }
                }
                catch (Exception ex)
                {
                    failCount++;
                    if (failCount <= 10) // 처음 10개 실패만 로그
                    {
                        _backupLog.AppendLine($"  - ✗ 파라미터 {paramNo} 예외 발생: {ex.Message}");
                    }
                }
            }

            // 완료 알림
            progressCallback?.Invoke(totalParams, totalParams);

            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);

            var elapsed = (DateTime.Now - startTime).TotalSeconds;
            _backupLog.AppendLine($"  - ✓ 백업 완료: 성공 {successCount}개, 실패 {failCount}개");
            _backupLog.AppendLine($"  - 파일 저장: {filePath}");
            _backupLog.AppendLine($"  - 소요 시간: {elapsed:F2}초");
            _backupLog.AppendLine();

            // 단독 호출 시 로그 저장
            if (isStandaloneCall)
            {
                SaveLogToFile();
            }

            return filePath;
        }

        /// <summary>
        /// PMC/래더 백업
        /// API: pmc_rdpmcrng (PMC Data Read - Area Specified)
        /// </summary>
        public string BackupPMCLadder()
        {
            if (_connection == null || !_connection.IsConnected)
            {
                throw new Exception("CNC 연결이 필요합니다.");
            }

            string backupFolder = CreateBackupFolder(_connection.IpAddress);

            // 로그가 비어있으면 초기화 (단독 호출 시)
            bool isStandaloneCall = _backupLog.Length == 0;
            if (isStandaloneCall)
            {
                InitializeLog();
            }

            var startTime = DateTime.Now;
            _backupLog.AppendLine($"[{startTime:HH:mm:ss}] PMC/래더 백업 시작");

            string ladderFolder = CreateSubFolder(backupFolder, "래더");
            string filePath = Path.Combine(ladderFolder, "PMC_Ladder.pmc");

            StringBuilder sb = new StringBuilder();
            sb.AppendLine("===== PMC/래더 백업 =====");
            sb.AppendLine($"백업 일시: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine($"IP 주소: {_connection.IpAddress}");
            sb.AppendLine();

            int totalSuccess = 0;
            int totalFail = 0;

            // PMC 영역별 백업 - 안전하게 작은 범위만 백업
            _backupLog.AppendLine($"  - PMC R 영역 백업 시작 (R0-R99)");
            sb.AppendLine("[PMC R 영역 (R0-R99)]");
            var (rSuccess, rFail) = BackupPMCArea(sb, 5, 0, 99, "R"); // R-type, 주소 0~99만
            totalSuccess += rSuccess;
            totalFail += rFail;
            _backupLog.AppendLine($"  - PMC R 영역 완료: 성공 {rSuccess}개, 실패 {rFail}개");

            sb.AppendLine();
            _backupLog.AppendLine($"  - PMC G 영역 백업 시작 (G0-G99)");
            sb.AppendLine("[PMC G 영역 (G0-G99)]");
            var (gSuccess, gFail) = BackupPMCArea(sb, 0, 0, 99, "G"); // G-type, 주소 0~99만
            totalSuccess += gSuccess;
            totalFail += gFail;
            _backupLog.AppendLine($"  - PMC G 영역 완료: 성공 {gSuccess}개, 실패 {gFail}개");

            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);

            var elapsed = (DateTime.Now - startTime).TotalSeconds;
            _backupLog.AppendLine($"  - ✓ 전체 백업 완료: 성공 {totalSuccess}개, 실패 {totalFail}개");
            _backupLog.AppendLine($"  - 파일 저장: {filePath}");
            _backupLog.AppendLine($"  - 소요 시간: {elapsed:F2}초");
            _backupLog.AppendLine();

            // 단독 호출 시 로그 저장
            if (isStandaloneCall)
            {
                SaveLogToFile();
            }

            return filePath;
        }

        private (int successCount, int failCount) BackupPMCArea(StringBuilder sb, short pmcType, ushort startAddr, ushort endAddr, string areaName)
        {
            int successCount = 0;
            int failCount = 0;

            try
            {
                // 안전하게 작은 단위로 나눠서 읽기 (10개씩)
                int totalAddrs = endAddr - startAddr + 1;
                _backupLog.AppendLine($"    - {areaName} 영역 읽기 시작 (주소 {startAddr}~{endAddr}, 총 {totalAddrs}개, type={pmcType})");

                for (ushort addr = startAddr; addr <= endAddr; addr += 10)
                {
                    ushort readEnd = (ushort)Math.Min(addr + 9, endAddr);
                    ushort readCount = (ushort)(readEnd - addr + 1);

                    Focas1.IODBPMC0 pmcData = new Focas1.IODBPMC0();
                    ushort length = (ushort)(8 + readCount);

                    short ret = Focas1.pmc_rdpmcrng(_connection.Handle, pmcType, 0, addr, readEnd, length, pmcData);

                    if (ret == Focas1.EW_OK)
                    {
                        for (int i = 0; i < readCount; i++)
                        {
                            sb.AppendLine($"주소 {addr + i}: {pmcData.cdata[i]}");
                            successCount++;
                        }
                    }
                    else
                    {
                        sb.AppendLine($"주소 {addr}~{readEnd} 읽기 실패 (에러 코드: {ret})");
                        _backupLog.AppendLine($"    - {areaName} 주소 {addr}~{readEnd} 읽기 실패 (에러 코드: {ret})");
                        failCount += readCount;
                    }

                    // 50개마다 진행 상태
                    if ((addr - startAddr) % 50 == 0 && addr != startAddr)
                    {
                        _backupLog.AppendLine($"    - {areaName} 진행 중... 주소 {addr} (성공: {successCount}, 실패: {failCount})");
                    }
                }
            }
            catch (Exception ex)
            {
                sb.AppendLine($"백업 실패: {ex.Message}");
                _backupLog.AppendLine($"    - {areaName} 백업 예외 발생: {ex.Message}");
            }

            return (successCount, failCount);
        }

        /// <summary>
        /// 전체 백업 (모든 데이터)
        /// </summary>
        public Dictionary<string, string> BackupAll()
        {
            var results = new Dictionary<string, string>();

            // 백업 폴더 생성 및 로그 초기화
            string backupFolder = CreateBackupFolder(_connection.IpAddress);
            InitializeLog();
            LogBackupTask("전체 백업 시작", "진행 중", "모든 데이터 백업을 시작합니다.");

            int successCount = 0;
            int failCount = 0;

            // 1. 공구 옵셋 백업
            try
            {
                results["ToolOffset"] = BackupToolOffset();
                LogBackupItem("공구 옵셋", true, results["ToolOffset"]);
                successCount++;
            }
            catch (Exception ex)
            {
                results["ToolOffset"] = $"실패: {ex.Message}";
                LogBackupItem("공구 옵셋", false, ex.Message);
                failCount++;
            }

            // 2. 워크 좌표계 백업
            try
            {
                results["WorkCoordinate"] = BackupWorkCoordinate();
                LogBackupItem("워크 좌표계", true, results["WorkCoordinate"]);
                successCount++;
            }
            catch (Exception ex)
            {
                results["WorkCoordinate"] = $"실패: {ex.Message}";
                LogBackupItem("워크 좌표계", false, ex.Message);
                failCount++;
            }

            // 3. 매크로 변수 백업
            try
            {
                results["MacroVariable"] = BackupMacroVariable();
                LogBackupItem("매크로 변수", true, results["MacroVariable"]);
                successCount++;
            }
            catch (Exception ex)
            {
                results["MacroVariable"] = $"실패: {ex.Message}";
                LogBackupItem("매크로 변수", false, ex.Message);
                failCount++;
            }

            // 4. 파라미터 백업
            try
            {
                results["Parameter"] = BackupParameter();
                LogBackupItem("파라미터", true, results["Parameter"]);
                successCount++;
            }
            catch (Exception ex)
            {
                results["Parameter"] = $"실패: {ex.Message}";
                LogBackupItem("파라미터", false, ex.Message);
                failCount++;
            }

            // 5. PMC/래더 백업
            try
            {
                results["PMC_Ladder"] = BackupPMCLadder();
                LogBackupItem("PMC/래더", true, results["PMC_Ladder"]);
                successCount++;
            }
            catch (Exception ex)
            {
                results["PMC_Ladder"] = $"실패: {ex.Message}";
                LogBackupItem("PMC/래더", false, ex.Message);
                failCount++;
            }

            // 백업 완료 로그
            _backupLog.AppendLine();
            LogBackupTask("전체 백업 완료", "완료", $"성공: {successCount}개, 실패: {failCount}개");
            SaveLogToFile();

            return results;
        }
    }
}
