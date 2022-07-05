# Object model

## Application [reference](https://docs.microsoft.com/ko-kr/office/vba/api/excel.application(object) "microsoft docs")

Application은 엑셀의 process를 가리킨다.

실행되어 있는 엑셀을 제어하거나 종료하려면 Application 개체를 이용한다.


    `book1.xls 파일 활성화
    Application.Windows("book1.xls").Activate
    
    `액셀 종료
    Application.Quit()


특이사항으로는 Application의 property들 중 대부분의 UI 객체들(ActiveCell, ActiveSheet 등)은 Application 객체 한정자(Object Qualifier)를 사용하지 않고도 작성할 수 있다.


    Application.ActiveCell.Font.Bold = True
    ActiveCell.Font.Bold = True

## Workbooks

하나의 액셀 process에서 열고 있는 파일 목록을 관리한다.

새로운 파일을 생성하거나, 파일을 열거나 닫는 등의 용도로 이용한다.

파일 목록에서 특정 파일을 접근할 때 Workbooks(1), Workbooks(2)... 등의 방식으로 접근할 수 있다.

    `모든 파일 닫기
    Workbooks.Close
    `새로운 파일 생성하기
    Workbooks.Add
    `파일 열기
    Workbooks.Open FileName:="File.xlsx", Readonly=True

[reference](https://docs.microsoft.com/ko-kr/office/vba/api/excel.workbooks "microsoft docs")

## Workbook



