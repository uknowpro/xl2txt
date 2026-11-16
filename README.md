# xl2txt

Windows에 설치된 Microsoft Excel을 통해 파일을 열고, 표시된 워크시트마다 **쉼표 구분 UTF-8 TXT**를 생성하는 데스크톱 도구입니다.

## 다운로드 및 실행

[Releases](https://github.com/uknowpro/xl2txt/releases)에서 `xl2txt-win-x64.exe`를 받아 실행하세요. 32비트 Windows에는 `xl2txt-win-x86.exe`를 사용하세요. Excel은 별도 프로세스 COM 서버이므로 일반적으로 Excel 비트 수와 맞출 필요가 없습니다.

- Windows 10 1607 이상 / Windows 11. Windows 10 대상 PC에서 별도 실사용 확인이 필요합니다.
- **Microsoft Excel 2016 이상** 데스크톱 앱 필수 (`xlCSVUTF8` 지원 버전). 웹 Excel/뷰어만으로는 동작하지 않습니다.
- Python 또는 .NET 별도 설치 불필요. 단일 EXE에 런타임 포함. 실행 시 .NET이 네이티브 구성요소를 사용자 임시 폴더에 풀 수 있습니다.
- 코드 서명되지 않은 초기 릴리즈입니다. 회사 실행 정책에 따라 차단될 수 있습니다.

## 사용법

1. 파일 선택으로 `.xlsx`, `.xls`, `.xlsm`, `.xlsb`를 지정합니다.
2. 저장 폴더를 선택하고 **TXT로 변환**을 누릅니다.
3. Excel 인증/암호/형식 확인 창이 나타나면 Excel에서 직접 처리합니다.
4. 완료 후 **결과 폴더 열기**를 누릅니다.

파일명과 시간/임의 식별자로 새 결과 폴더를 생성합니다. `001_시트명.txt` 형태로 저장하며 이전 결과는 덮어쓰지 않습니다. 처리 도중에는 선택한 저장 폴더 아래 `.xl2txt-*.partial`에 임시 CSV/TXT를 기록합니다. 실패·취소 시 삭제를 시도하며, 정리 실패는 화면 로그에 표시합니다. 강제 종료된 경우 해당 폴더가 남을 수 있습니다.

## 변환 규칙

- Excel의 공식 `Worksheet.Copy`와 `Workbook.SaveAs(xlCSVUTF8)`로 내보낸 후, 지역 설정의 목록 구분자를 쉼표로 정규화합니다.
- UTF-8 BOM, Windows 줄바꿈(CRLF), CSV 따옴표 규칙을 사용합니다. 셀 내 쉼표·따옴표·줄바꿈을 보존합니다.
- 값과 날짜/숫자 표현은 **Excel의 CSV 저장 결과**를 따릅니다. 수식은 수식 문자열이 아닌 Excel이 내보낸 결과 값입니다. Excel이 열면서 재계산할 수 있습니다.
- 숨김/매우 숨김 시트와 차트 시트는 제외합니다. 표시된 시트 내 숨김 행·열은 Excel CSV 내보내기 규칙에 따라 포함될 수 있습니다.
- 병합 셀, 서식, 이미지, 차트, 매크로는 TXT로 보존되지 않습니다.
- Excel 기본 CSV 저장 확인 창이 표시될 수 있습니다. 취소는 현재 Excel 호출/대화상자가 끝난 후 적용됩니다.

## 권한과 원본 보호

현재 Windows 사용자의 Excel로 읽기 전용으로 엽니다. 매크로 실행을 비활성화하고 외부 링크 갱신을 요청하지 않으며, 원본을 저장하지 않습니다. 기존 Excel 창에 연결하지 않고 생성한 Excel 인스턴스를 종료합니다.

DRM 해제, 암호 해독, 보호된 보기 해제, 보안 설정 완화 기능은 없습니다. Excel이 복사/내보내기를 거부하면 실패 처리합니다. 화면에서 읽을 수 있어도 프로그램 접근이나 TXT 내보내기는 별도 제한될 수 있습니다. 모든 외부 DRM 제품의 정책을 자체 판정하는 기능은 없으며, 해당 환경에서 정상적인 Excel 내보내기가 허용되는 문서에 사용해야 합니다. 원본의 DRM이 TXT에 자동 적용되는 것은 아닙니다.

## 개발 및 검증

.NET 8 SDK:

```powershell
dotnet run --project tests/Tests.csproj -c Release
dotnet publish src/Xl2Txt/Xl2Txt.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:DebugType=None -o dist/win-x64
```

GitHub Actions는 CSV 정규화 테스트와 Windows x64/x86 빌드를 수행합니다. GitHub 호스팅 러너에는 Excel이 없어 Excel/DRM 통합 테스트는 실행하지 않습니다.

Windows 10 + Excel 실사용 확인 항목:

- 한글, 쉼표, 큰따옴표, 여러 줄 셀, 날짜/숫자/수식, 빈 시트와 여러 시트.
- 원본 파일의 내용과 수정 시각이 유지되는지, 기존 Excel 창이 닫히지 않는지.
- 암호/인증 및 DRM에서 내보내기 허용/거부 시 동작.
- 쓰기 불가 폴더, 변환 취소, 기존 결과 보존, Excel 프로세스 종료.
- 목록 구분자가 세미콜론인 지역 설정에서도 쉼표 TXT 생성.

릴리즈는 main에 커밋·푸시한 뒤 성공한 빌드 산출물을 해당 커밋의 태그에 첨부합니다.
