# LOTTEON Google Sheets Apps Script Automation

LOTTEON 구글 스프레드시트 자동화 운영 코드 저장소입니다.

## 파일 구조

```text
Code.gs     실제 운영 코드 전체본
loader.gs   Apps Script에 1회 붙여넣는 GitHub 원격 로더
README.md   운영 안내
```

## 운영 방식

1. GitHub 저장소의 `Code.gs`에 최신 Apps Script 전체 코드를 보관합니다.
2. 구글 Apps Script에는 `loader.gs`만 붙여넣습니다.
3. 구글시트 메뉴를 누르면 `loader.gs`가 GitHub Raw URL의 `Code.gs`를 불러와 실행합니다.

## Raw URL

```text
https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/Code.gs
```

## 주의사항

- API Key, X-API-SENDER 같은 인증값은 GitHub 코드에 넣지 않습니다.
- 인증값은 기존처럼 Apps Script Properties에 저장합니다.
- `Code.gs` 수정 후 Apps Script에 다시 붙여넣을 필요는 없습니다.
- `Code.gs`가 약 400KB 이상이므로 loader는 CacheService에 저장하지 않고 실행 때마다 Raw URL에서 새로 불러옵니다.

## 현재 기준

```text
최신 기준: v6.00
핵심 변경: ① 변경사항 반영 실행을 빠른 안전 경로로 분리하여 시간초과 방지
```
