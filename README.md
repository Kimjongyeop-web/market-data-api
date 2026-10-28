# Market Data API | 시장 데이터 수집

김종엽의 금융 연구용 데이터 수집 코드입니다. 데이터 수집과 모델 실험을 구분합니다.

| 연동 | 코드 | 역할 |
|---|---|---|
| Upbit 공개 시세 API | `upbit/fetch_daily.py` | BTC/KRW 일봉 페이지 순회 |
| Upbit 공개 시세 API | `upbit/fetch_minutes.py` | 1분봉 월별 저장, 재시도, 속도 제한, 재개 |
| Naver 시세 HTTP | `naver/kospi200_full_history.py` | 편입 이력 종목 전체 가격 수집·기존 자료 보완 |
| GDELT 이벤트·기사 HTTP | 별도 `commodity-risk-dgri` | 뉴스 수집과 시장별 coverage |

Naver HTTP 수집과 인증형 거래 API는 구분합니다. Upbit 두 스크립트에는 키가 필요하지 않습니다. 실계좌 주문 기능은 없습니다. 추가 증권사·Discord 봇 원본은 이번 확인 범위에서 확보하지 못했습니다.

## 실행

새 실행 폴더에 `work/`를 만든 뒤 그 폴더를 현재 디렉터리로 두고 Upbit 스크립트를 실행합니다. 1분봉 전체 수집에는 상당한 시간과 네트워크 요청이 필요합니다. 보존할 원응답과 임시 partial 파일을 구분하세요. 원코드는 완료된 partial 파일을 제거하므로 기존 원본 폴더에서 직접 실행하지 마세요.

Naver 코드는 별도 실행 폴더에 복사하고 코드 상단에 명시된 기존 분석·QC Excel 복사본을 같은 폴더에 준비합니다. 서버 응답과 제한 변경 가능성이 있으며 이번 정리에서 API 호출을 재검증하지 않았습니다.

## 격리 실행 폴더 만들기

먼저 python prepare_run.py --run-id research_001 을 실행하면 소스 복사본과 빈 work/outputs 폴더만 생성됩니다. 연구·다운로드·삭제는 실행하지 않습니다. 이후 위 재현 순서에 맞춰 입력과 작업 디렉터리를 준비하세요.

