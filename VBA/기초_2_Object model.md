# Object model

## 1. Application [[reference]](https://docs.microsoft.com/ko-kr/office/vba/api/excel.application(object) "microsoft docs")

Application은 엑셀의 process를 가리킨다.

실행되어 있는 엑셀을 제어하거나 종료하려면 Application 개체를 이용한다.


    `book1.xls 파일 활성화
    Application.Windows("book1.xls").Activate
    
    `액셀 종료
    Application.Quit()


특이사항으로 Application의 property들 중 대부분의 UI 객체들(ActiveCell, ActiveSheet 등)은 Application 객체 한정자(Object Qualifier)를 사용하지 않고도 작성할 수 있다.


    Application.ActiveCell.Font.Bold = True
    ActiveCell.Font.Bold = True

## 2. Workbooks [[reference]](https://docs.microsoft.com/ko-kr/office/vba/api/excel.workbooks "microsoft docs")

하나의 액셀 process에서 열고 있는 파일 목록을 관리한다.

새로운 파일을 생성하거나, 파일을 열거나 닫는 등의 용도로 이용한다.

파일 목록에서 특정 파일을 접근할 때 Workbooks(1), Workbooks(2)... 등의 방식으로 접근할 수 있다.

    `모든 파일 닫기
    Workbooks.Close
    
    `새로운 파일 생성하기
    Workbooks.Add
    
    `파일 열기
    Workbooks.Open FileName:="File.xlsx", Readonly=True

## 3. Workbook [[reference]](https://docs.microsoft.com/ko-kr/office/vba/api/excel.workbook "microsoft docs")

하나의 엑셀 process에서 열고 있는 파일 목록 중 특정 파일 한개를 가리킨다.

파일을 닫거나, 저장하는 등의 용도로 이용한다.

    파일 닫기
    Workbook.Close
    
    저장하기
    Workbook.Save
    
    다른 이름으로 저장하기
    Workbook.SaveAs Filename:="NewFile.xlsx"

## 4. Worksheets [[reference]](https://docs.microsoft.com/ko-kr/office/vba/api/excel.worksheets "microsoft docs")

한 파일 내의 시트 목록(Collection)을 관리한다.

새로운 시트를 생성하거나, 삭제하는 등의 용도로 사용한다.

    '첫번째 시트 앞에 2개 시트 추가
    WorkSheets.Add Count:=2, Before:=Sheets(1)
    
    `시트 개수 출력
    Debug.print Worksheets.Count
    
    `Sheet3 뒤에 Sheet1 복사
    Worksheets("Sheets1").Copy After:=Worksheets("sheet3")
    
## 5. Worksheet [[reference]](https://docs.microsoft.com/ko-kr/office/vba/api/excel.worksheet "microsoft docs")
