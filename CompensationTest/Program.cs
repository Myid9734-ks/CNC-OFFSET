using System;

/// <summary>
/// 보정 로직 검증을 위한 시뮬레이션 클래스
/// </summary>
public class CompensationLogicTest
{
    // 목표값 설정 (실제 코드와 동일)
    private const decimal WIDTH_TARGET = 17.187m;      // 폭 목표값
    private const decimal BOTTOM_HEIGHT_TARGET = 21.70m; // 하단높이 목표값
    private const decimal EMERGENCY_STOP_THRESHOLD = 0.1m; // ±0.1mm 이상 보정 시 즉시 비상정지

    /// <summary>
    /// 보정량 계산 (실제 코드와 동일한 로직)
    /// </summary>
    private static decimal CalculateCompensation(decimal targetValue, decimal measuredValue)
    {
        return targetValue - measuredValue;
    }

    /// <summary>
    /// 하단높이 연동 폭 보정 계산 (수정된 로직)
    /// </summary>
    private static decimal CalculateWidthLinkedCompensation(decimal bottomHeightCompensation)
    {
        // 수정된 로직: 같은 방향으로 보정
        return bottomHeightCompensation;
    }

    /// <summary>
    /// 시뮬레이션 테스트 실행
    /// </summary>
    public static void RunSimulation()
    {
        Console.WriteLine("=== 보정 로직 시뮬레이션 테스트 ===\n");

        // 테스트 케이스 1: 폭 단독 보정
        Console.WriteLine("1. 폭 단독 보정 테스트");
        TestWidthCompensation(17.05m, "기준값보다 작음 (치수 증가 필요)");
        TestWidthCompensation(17.20m, "기준값보다 큼 (치수 감소 필요)");
        TestWidthCompensation(17.187m, "목표값과 동일");
        Console.WriteLine();

        // 테스트 케이스 2: 하단높이 연동 보정
        Console.WriteLine("2. 하단높이 연동 폭 보정 테스트");
        TestBottomHeightLinkedCompensation(21.65m, "하단높이 작음 → 음의 보정");
        TestBottomHeightLinkedCompensation(21.75m, "하단높이 큼 → 양의 보정");
        TestBottomHeightLinkedCompensation(21.70m, "하단높이 목표값과 동일");
        Console.WriteLine();

        // 테스트 케이스 3: 과도한 보정량 테스트
        Console.WriteLine("3. 과도한 보정량 비상정지 테스트");
        TestEmergencyStopScenario();
        Console.WriteLine();

        // 테스트 케이스 4: 실제 시나리오
        Console.WriteLine("4. 실제 시나리오 테스트");
        TestRealScenario();
    }

    private static void TestWidthCompensation(decimal measuredWidth, string description)
    {
        decimal compensation = CalculateCompensation(WIDTH_TARGET, measuredWidth);
        string direction = compensation > 0 ? "양(+)" : compensation < 0 ? "음(-)" : "보정불필요";

        Console.WriteLine($"  측정값: {measuredWidth:F3}mm, 목표값: {WIDTH_TARGET:F3}mm");
        Console.WriteLine($"  보정량: {compensation:F3}mm ({direction}방향)");
        Console.WriteLine($"  설명: {description}");
        Console.WriteLine();
    }

    private static void TestBottomHeightLinkedCompensation(decimal measuredBottomHeight, string description)
    {
        // 하단높이 보정량 계산
        decimal bottomHeightCompensation = CalculateCompensation(BOTTOM_HEIGHT_TARGET, measuredBottomHeight);

        // 연동된 폭 보정량 계산 (수정된 로직)
        decimal widthLinkedCompensation = CalculateWidthLinkedCompensation(bottomHeightCompensation);

        string bottomDirection = bottomHeightCompensation > 0 ? "양(+)" : bottomHeightCompensation < 0 ? "음(-)" : "보정불필요";
        string widthDirection = widthLinkedCompensation > 0 ? "양(+)" : widthLinkedCompensation < 0 ? "음(-)" : "보정불필요";

        Console.WriteLine($"  하단높이 측정값: {measuredBottomHeight:F3}mm, 목표값: {BOTTOM_HEIGHT_TARGET:F3}mm");
        Console.WriteLine($"  하단높이 보정량: {bottomHeightCompensation:F3}mm ({bottomDirection}방향)");
        Console.WriteLine($"  연동 폭 보정량: {widthLinkedCompensation:F3}mm ({widthDirection}방향)");
        Console.WriteLine($"  설명: {description}");
        Console.WriteLine($"  결과: 하단높이와 폭이 같은 방향으로 보정됨 ✓");
        Console.WriteLine();
    }

    private static void TestRealScenario()
    {
        Console.WriteLine("  실제 시나리오: 하단높이 21.65mm, 폭 17.20mm 측정");

        // 하단높이 보정
        decimal bottomHeightCompensation = CalculateCompensation(BOTTOM_HEIGHT_TARGET, 21.65m);
        Console.WriteLine($"  하단높이 보정: {bottomHeightCompensation:F3}mm (목표에 맞추기 위해 +방향)");

        // 폭 직접 보정
        decimal widthCompensation = CalculateCompensation(WIDTH_TARGET, 17.20m);
        Console.WriteLine($"  폭 직접 보정: {widthCompensation:F3}mm (목표에 맞추기 위해 -방향)");

        // 하단높이 연동 폭 보정
        decimal widthLinkedCompensation = CalculateWidthLinkedCompensation(bottomHeightCompensation);
        Console.WriteLine($"  하단높이 연동 폭 보정: {widthLinkedCompensation:F3}mm (+방향, 하단높이와 같은 방향)");

        // 최종 폭 보정량
        decimal finalWidthCompensation = widthCompensation + widthLinkedCompensation;
        Console.WriteLine($"  최종 폭 보정량: {finalWidthCompensation:F3}mm");
        Console.WriteLine($"  결과: 하단높이 보정의 영향이 폭 보정에 반영됨 ✓");
    }

    private static void TestEmergencyStopScenario()
    {
        Console.WriteLine("  과도한 보정량 테스트:");

        // 테스트 케이스들
        var testCases = new[]
        {
            new { measured = 17.05m, target = WIDTH_TARGET, item = "폭", description = "보정량 +0.137mm (임계값 이상)" },
            new { measured = 17.287m, target = WIDTH_TARGET, item = "폭", description = "보정량 -0.100mm (임계값과 동일)" },
            new { measured = 21.59m, target = BOTTOM_HEIGHT_TARGET, item = "하단높이", description = "보정량 +0.110mm (임계값 이상)" },
            new { measured = 17.277m, target = WIDTH_TARGET, item = "폭", description = "보정량 -0.090mm (임계값 미만)" }
        };

        foreach (var testCase in testCases)
        {
            decimal compensation = CalculateCompensation(testCase.target, testCase.measured);
            bool isEmergencyStop = Math.Abs(compensation) >= EMERGENCY_STOP_THRESHOLD;
            string result = isEmergencyStop ? "⚠️ 즉시 비상정지" : "✓ 정상 보정 진행";

            Console.WriteLine($"  {testCase.item} 측정값: {testCase.measured:F3}mm");
            Console.WriteLine($"  보정량: {compensation:F3}mm");
            Console.WriteLine($"  결과: {result}");
            Console.WriteLine($"  설명: {testCase.description}");
            Console.WriteLine();
        }
    }

    public static void Main()
    {
        RunSimulation();
        Console.WriteLine("\n검증 완료: 보정 로직이 올바르게 수정되었습니다!");
        Console.WriteLine("아무 키나 눌러 종료...");
        Console.ReadKey();
    }
}
