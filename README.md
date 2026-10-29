# Market Data API

금융 연구에 사용하는 Upbit·Naver 시세 수집 코드입니다. Upbit 코드는 인증키 없이 공개 시세를 조회하며 주문 기능은 없습니다.

## 수집 코드

| 소스 | 대상 | 코드 | 처리 내용 |
|---|---|---|---|
| Upbit | BTC/KRW 일봉 | [fetch_daily.py](upbit/fetch_daily.py) | 페이지 순회 |
| Upbit | BTC/KRW 1분봉 | [fetch_minutes.py](upbit/fetch_minutes.py) | 월별 저장, 재시도, 속도 제한, 중단 후 재개 |
| Naver | KOSPI200 편입 이력 종목 | [kospi200_full_history.py](naver/kospi200_full_history.py) | 가격 이력 수집과 기존 자료 보완 |

## 실행 전 확인

필요한 패키지는 [requirements.txt](requirements.txt)에 정리했습니다. 서버 응답 형식과 호출 제한은 변경될 수 있어 실행 시 현재 동작을 확인해야 합니다.

| 코드 | 입력·저장 위치 |
|---|---|
| Upbit | 현재 작업 폴더의 `work/` 사용 |
| Naver | 스크립트 상단에서 지정한 분석·QC Excel 필요 |

`python prepare_run.py --run-id research_001`은 실행별 소스 복사본과 빈 작업 폴더를 만듭니다. 수집 작업은 보존 원본과 분리된 폴더에서 진행합니다.

분봉 수집기에는 완료된 임시 partial 파일을 제거하는 동작이 있습니다. 보존할 파일을 해당 임시 경로에 두지 않아야 합니다.

## 사용하는 연구

- [Bitcoin 변동성 예측](https://github.com/Kimjongyeop-web/bitcoin-volatility-research) — BTC 실현변동성과 HAR / HAR-X 비교
- [KOSPI200 단기반전과 PCA](https://github.com/Kimjongyeop-web/kospi200-pca-reversal) — 종목별 가격자료를 이용한 공동연구

GDELT 뉴스 수집 코드는 [Commodity Risk DGRI](https://github.com/Kimjongyeop-web/commodity-risk-dgri)에 별도로 있습니다.

수집한 원자료는 이 저장소에 포함하지 않습니다. [데이터 범위](DATA_POLICY.md)
