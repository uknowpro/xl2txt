# xl2txt

Windows에 설치된 Microsoft Excel을 통해 파일을 열고, 표시된 워크시트마다 **쉼표 구분 UTF-8 TXT**를 생성하는 데스크톱 도구입니다.

## 다운로드 및 실행

[Releases](https://github.com/uknowpro/xl2txt/releases)에서 `xl2txt-win-x64.exe`를 받아 실행하세요. 32비트 Windows에는 `xl2txt-win-x86.exe`를 사용하세요. Excel은 별도 프로세스 COM 서버이므로 일반적으로 Excel 비트 수와 맞출 필요가 없습니다.

- Windows 10 1607 이상 / Windows 11. Windows 10 대상 PC에서 별도 실사용 확인이 필요합니다.
- **Microsoft Excel 데스크톱 앱 필수**. Excel CSV 내보내기 방식은 Excel 2016 이상 (`xlCSVUTF8` 지원 버전)이 필요합니다. 웹 Excel/뷰어만으로는 동작하지 않습니다.
- Python 또는 .NET 별도 설치 불필요. 단일 EXE에 런타임 포함. 실행 시 .NET이 네이티브 구성요소를 사용자 임시 폴더에 풀 수 있습니다.
- 코드 서명되지 않은 초기 릴리즈입니다. 회사 실행 정책에 따라 차단될 수 있습니다.

## 사용법

1. 파일 선택으로 `.xlsx`, `.xls`, `.xlsm`, `.xlsb`를 지정합니다.
2. 저장 폴더와 변환 방식을 선택하고 **TXT로 변환**을 누릅니다.
   - **셀 표시값 직접 읽기 (기본)**: Excel에서 셀의 표시 텍스트를 읽어 TXT로 직접 기록합니다. Excel의 CSV 저장 기능을 사용하지 않습니다.
   - **Excel CSV 내보내기**: Excel의 기본 내보내기로 변환합니다. 큰 파일에는 이 방식이 더 빠를 수 있습니다.
3. Excel 인증/암호/형식 확인 창이 나타나면 Excel에서 직접 처리합니다.
4. 완료 후 **결과 폴더 열기**를 누릅니다.

파일명과 시간/임의 식별자로 새 결과 폴더를 생성합니다. `001_시트명.txt` 형태로 저장하며 이전 결과는 덮어쓰지 않습니다. 처리 도중에는 선택한 저장 폴더 아래 `.xl2txt-*.partial`에 임시 CSV/TXT를 기록합니다. 실패·취소 시 삭제를 시도하며, 정리 실패는 화면 로그에 표시합니다. 강제 종료된 경우 해당 폴더가 남을 수 있습니다.

## 변환 규칙

- 직접 읽기: 공식 Excel COM `Range.Text`로 셀 표시 텍스트를 읽습니다. A1부터 마지막 값/수식의 행·열까지 기록합니다. 서식만 있는 빈 영역은 제외합니다. 열 너비에 따른 `###`, 반올림, 과학적 표기 등도 표시된 그대로 저장되므로 필요한 열 너비를 원본에서 미리 조정하세요.
- CSV 내보내기: 공식 `Worksheet.Copy`와 `Workbook.SaveAs(xlCSVUTF8)`로 내보낸 후, 지역 설정의 목록 구분자를 쉼표로 정규화합니다.
- UTF-8 BOM, Windows 줄바꿈(CRLF), CSV 따옴표 규칙을 사용합니다. 셀 내 쉼표·따옴표·줄바꿈을 보존합니다.
- 날짜/숫자 표현은 직접 읽기에서는 **셀 표시 텍스트**, CSV 방식에서는 **Excel의 CSV 저장 결과**를 따릅니다. 수식은 결과 값을 기록하며 Excel이 열면서 재계산할 수 있습니다.
- 숨김/매우 숨김 시트와 차트 시트는 제외합니다. 표시된 시트 내 숨김 행·열은 포함될 수 있습니다.
- 병합 셀, 서식, 이미지, 차트, 매크로는 TXT로 보존되지 않습니다.
- Excel 기본 CSV 저장 확인 창이 표시될 수 있습니다. 취소는 현재 Excel 호출/대화상자가 끝난 후 적용됩니다.

## 권한과 원본 보호

현재 Windows 사용자의 Excel로 읽기 전용으로 엽니다. 매크로 실행을 비활성화하고 외부 링크 갱신을 요청하지 않으며, 원본을 저장하지 않습니다. 기존 Excel 창에 연결하지 않고 생성한 Excel 인스턴스를 종료합니다.

암호는 Excel에서 직접 입력합니다. 보호된 보기나 보안 설정을 완화하지 않습니다. 사용 권한이 제한된 Office 문서는 직접 읽기 모드에서 중단합니다. Excel이 접근·복사·내보내기를 거부하면 오류를 표시하며, 실패한 작업을 다른 방식으로 자동 재시도하지 않습니다. 접근과 텍스트 추출이 허용된 문서에 사용하세요.

## 개발 및 검증

.NET 8 SDK:

```powershell
dotnet run --project tests/Tests.csproj -c Release
dotnet publish src/Xl2Txt/Xl2Txt.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:DebugType=None -o dist/win-x64
```

GitHub Actions는 CSV 정규화 테스트, Windows x64/x86 빌드, 두 EXE의 창 열기·닫기 검사를 수행합니다. GitHub 호스팅 러너에는 Excel이 없어 실제 Excel 통합 테스트는 실행하지 않습니다.

Windows 10 + Excel 실사용 확인 항목:

- 한글, 쉼표, 큰따옴표, 여러 줄 셀, 날짜/숫자/수식, 빈 시트와 여러 시트.
- 원본 파일의 내용과 수정 시각이 유지되는지, 기존 Excel 창이 닫히지 않는지.
- 암호/인증 및 내보내기 허용/거부 시 동작.
- 쓰기 불가 폴더, 변환 취소, 기존 결과 보존, Excel 프로세스 종료.
- 목록 구분자가 세미콜론인 지역 설정에서도 쉼표 TXT 생성.

릴리즈는 main에 커밋·푸시한 뒤 성공한 빌드 산출물을 해당 커밋의 태그에 첨부합니다.
