# LOTTEON Google Sheets Apps Script 자동화

## 현재 운영 기준

- 기본 운영 코드: `Code.gs` v6.00
- 자동 필터 갱신 패치: `Patch_v6_01_daily_filter_auto.gs`
- Apps Script 로더: `loader.gs` v1.2
- 운영 방식: Apps Script에는 `loader.gs`만 유지하고, 실제 코드는 GitHub Raw URL에서 불러와 실행합니다.

## GitHub Raw URL

```text
Code.gs
https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/Code.gs

Patch_v6_01_daily_filter_auto.gs
https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/Patch_v6_01_daily_filter_auto.gs
```

## v6.01 추가 기능

`필터별_상품수`를 사람이 여러 번 누르지 않도록, 매일 06:10 전후 자동 시작 후 1분 간격으로 이어실행합니다.

추가 메뉴:

```text
LOTTEON 자동화
→ 설정/초기화
→ 필터별_상품수 자동 갱신 시작(매일 06:10)
→ 필터별_상품수 자동 갱신 중지
→ 필터별_상품수 자동 갱신 지금 시작
→ 필터별_상품수 자동 상태 확인
```

자동 갱신 대상:

```text
직접 갱신: 필터별_상품수
진행 상태: 동기화상태
```

자동에서 제외:

```text
쿠팡재전송_로그
핵심_브랜드요약
대시보드
검수리포트
수동입력/API검증
```

필터 날짜/상품갈이를 실제 대시보드에 반영하려면 필요할 때 아래를 수동 실행합니다.

```text
③ 쿠팡재전송_로그 갱신
④ 핵심요약+대시보드 갱신
```

## Apps Script 적용

Apps Script에는 `loader.gs`만 붙여넣습니다.

```text
1. Apps Script에서 기존 loader 전체 삭제
2. GitHub의 loader.gs 전체 붙여넣기
3. 저장
4. authorizeLotteonLoader 1회 실행
5. 구글시트 새로고침
6. LOTTEON 자동화 → 설정/초기화 → GitHub 코드 연결 테스트
7. LOTTEON 자동화 → 설정/초기화 → 필터별_상품수 자동 갱신 시작(매일 06:10)
```
