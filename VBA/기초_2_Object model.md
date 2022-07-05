> ##### 목차
>  [1. Object model(객체 모델) 개요](#1-object-model객체-모델-개요-1)
> 
>  [1) Application](#1-application-reference)
>  [2) Workbooks](#2-workbooks-reference)
>  [3) Workbook](#3-workbook-reference)
>  [4) Worksheets](#4-worksheets-reference)
>  [5) Worksheet](#5-worksheet-reference)
>  [6) Range](#6-range-reference)
> 
>  [2. Object 모델 다루기](#2-object-모델-다루기)

# Object model

# 1. Object model(객체 모델) 개요

### 1) Application [[reference]](https://docs.microsoft.com/ko-kr/office/vba/api/excel.application(object) "microsoft docs")

Application은 엑셀의 process를 가리킨다.

실행되어 있는 엑셀을 제어하거나 종료하려면 Application 개체를 이용한다.


    `book1.xls 파일 활성화
    Application.Windows("book1.xls").Activate
    
    `액셀 종료
    Application.Quit()


특이사항으로 Application의 property들 중 대부분의 UI 객체들(ActiveCell, ActiveSheet 등)은 Application 객체 한정자(Object Qualifier)를 사용하지 않고도 작성할 수 있다.


    Application.ActiveCell.Font.Bold = True
    ActiveCell.Font.Bold = True

### 2) Workbooks [[reference]](https://docs.microsoft.com/ko-kr/office/vba/api/excel.workbooks "microsoft docs")

하나의 액셀 process에서 열고 있는 파일 목록을 관리한다.

새로운 파일을 생성하거나, 파일을 열거나 닫는 등의 용도로 이용한다.

파일 목록에서 특정 파일을 접근할 때 Workbooks(1), Workbooks(2)... 등의 방식으로 접근할 수 있다.

    `모든 파일 닫기
    Workbooks.Close
    
    `새로운 파일 생성하기
    Workbooks.Add
    
    `파일 열기
    Workbooks.Open FileName:="File.xlsx", Readonly=True

### 3) Workbook [[reference]](https://docs.microsoft.com/ko-kr/office/vba/api/excel.workbook "microsoft docs")

하나의 엑셀 process에서 열고 있는 파일 목록 중 특정 파일 한개를 가리킨다.

파일을 닫거나, 저장하는 등의 용도로 이용한다.

    `파일 닫기
    Workbook.Close
    
    `저장하기
    Workbook.Save
    
    `다른 이름으로 저장하기
    Workbook.SaveAs Filename:="NewFile.xlsx"

### 4) Worksheets [[reference]](https://docs.microsoft.com/ko-kr/office/vba/api/excel.worksheets "microsoft docs")

한 파일 내의 **시트 목록(Collection)을 관리**한다.

새로운 시트를 생성하거나, 삭제하는 등의 용도로 사용한다.

    '첫번째 시트 앞에 2개 시트 추가
    WorkSheets.Add Count:=2, Before:=Sheets(1)
    
    `시트 개수 출력
    Debug.print Worksheets.Count
    
    `Sheet3 뒤에 Sheet1 복사
    Worksheets("Sheets1").Copy After:=Worksheets("sheet3")
    
### 5) Worksheet [[reference]](https://docs.microsoft.com/ko-kr/office/vba/api/excel.worksheet "microsoft docs")

한 파일 내의 **시트 하나**를 가리킨다.

Worksheets와 Sheets Collection의 member이다.

시트를 활성화하거나, 삭제하기, 숨기기, 보이기 등의 용도로 사용한다.

    `워크시트 활성화
    Worksheets("Sheet1").Activate
    
    `워크시트 삭제
    Worksheets("Sheet1").Delete
    
    `워크시트 숨기기
    Worksheets(1).Visible = False
    
    `워크시트 보이기
    Worksheets(1).Visible = True
    
    `Worksheets()의 return Object가 Worksheet.
    
    
### 6) Range [[reference]](https://docs.microsoft.com/en-us/office/vba/api/excel.range(object) "microsoft docs")

한 시트 내의 하나의 cell 또는 여러 cell을 가리킨다.

cell에 값을 입력하거나, 삭제하거나, cell들을 병합하는 등의 용도로 사용한다.

액셀 VBA 코딩 시 가장 많이 사용하는 객체

    'A1 cell의 값을 A5에 입력
    Worksheets("Sheet1").Range("A5").Value = _ 
    Worksheets("Sheet1").Range("A1").Value

    'A1:E10 영역의 내용 삭제
    Worksheets(1).Range("A1:E10").ClearContents

    'A2(2행, 1열)에 B1:B5 합계 수식 입력
    Worksheets(1).Range("A2").Formula = "=Sum(B1:B5)"
    Worksheets(1).Cells(2, 1).Formula = "=Sum(B1:B5)"

    '5행 삭제하기
    Worksheets(1).Rows(5).Delete

    '3열(C) 삭제하기
    Worksheets(1).Columns("C").Delete
    Worksheets(1).Columns(3).Delete

    'A1:A2 cell 병합
    Worksheets(1).Range("A1:A2").Merge

## 2. Object 모델 다루기

### 1) 객체 접근 기본형

1. 이름이 t1.xlsx인 파일
2. 이 파일의 Sheet1 시트
3. 이 시트의 A1 cell
4. 이 시트의 A1 cell에 "TEST" 문자열 입력
5. 이 시트의 A1 cell의 글꼴을 굵게 설정
6. 

---
참고 사이트

[ * 취미로 코딩하는 DA(Data Architect)(티스토리)](https://prodtool.tistory.com/)
