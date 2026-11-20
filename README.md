# 밤비의 도장찍기

(옛 이름: 이미지 삽입기) PDF · 엑셀(XLSX) · 한글(HWP · HWPX) 문서에 도장·서명·로고·사진을 얹어 **같은 형식 그대로** 저장하는 웹 도구입니다.

- 주소: https://eugene-shin9.github.io/bambi-image-inserter/
- 파일은 브라우저 안에서만 처리되고, 어디로도 전송되지 않습니다(페이지에 외부 전송을 막는 보안 정책 적용).

## 기능
- 문서 위에 이미지를 끌어다 놓고 이동·크기 조절
- 스마트 가이드: 쪽 가장자리·가운데, 본문 글자 줄, 엑셀 셀 경계, 다른 이미지에 자동 맞춤 (Alt 누르면 해제)
- 누끼(배경 지우기) 체크
  - 단색 배경: 도장·서명·로고용. 배경색 자동 추정 또는 스포이드
  - AI: 사진용. 모델이 사이트에 함께 들어 있어 외부 서버를 쓰지 않음
- 실행 취소/다시 실행, 복제, mm 단위 크기 입력

## 형식별 방식
| 형식 | 저장 방식 |
|---|---|
| PDF | 원본 쪽 위에 이미지를 그림(쪽 회전·잘림 상자 반영). 보안 PDF는 불가, 전자서명은 무효가 됨 |
| XLSX | 원본 파일에 그림 파트만 추가(나머지 그대로). 이미지가 놓인 셀에 고정 |
| HWP / HWPX | rhwp 엔진으로 "글 앞으로" 그림을 쪽 기준 위치에 넣고 같은 형식으로 다시 저장 |

## 사용한 오픈소스
- [pdf.js](https://github.com/mozilla/pdf.js) (Apache-2.0)
- [pdf-lib](https://github.com/Hopding/pdf-lib) (MIT)
- [JSZip](https://github.com/Stuk/jszip) (MIT)
- [rhwp](https://github.com/edwardkim/rhwp) (MIT)
- [@imgly/background-removal](https://github.com/imgly/background-removal-js) (AGPL-3.0) 및 onnxruntime-web (MIT)

`@imgly/background-removal`이 AGPL-3.0이므로 이 저장소 전체를 AGPL-3.0으로 공개합니다. 각 라이브러리의 라이선스 원문은 `vendor/` 폴더에 있습니다.
