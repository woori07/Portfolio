# 엑셀 데이터 정제 자동화

- 하나의 엑셀 파일을 두 개의 엑셀 데이터로 정제하며, 원하는 정제 유형을 선택하여 진행하도록 구현하였습니다.

## 코드 실행
- weekly 작업 코드
    ```
    python excel_macro.py  -t weekly
    # --세부 경로 입력
    ```
- monthly 작업 코드
    ```
    python excel_macro.py  -t monthly
    # 세부 경로 입력
    ```

## 작업 Tree
```
.
├── Excel
│   ├── 결과
│   │   ├── 저널홍보제품목록_{yyyy-mm}_{yyyy-mm-dd}.xlsx
│   │   └── 저작권_{yyyy-mm}_{yyyy-mm-dd}.xlsx
│   └── 제품목록
│       ├── 2023-05
│       │   ├── file1.xlsx
│       │   ├── file2.xlsx
│       │   └── file2.xlsx
│       ├── 2023-06
│       │   ├── file1.xlsx
│       │   ├── file2.xlsx
│       │   └── file3.xlsx
│       └── ..
├── README.md
├── requirements.txt
└── excel_macro.py
```

## 코드 로직
### excel_macro.py [type : monthly]
```
input : 세부 파일 경로
output : 엑셀 파일 ['저널홍보제품목록_{year_month}_{today}.xlsx']
```
1. os로 해당 경로에 어떤 파일 목록이 있는지 가져온다.
2. 그 파일 목록을 읽는다.
3. 작업할 파일을 openpyxl로 읽는다.
 3-1. DataFrame으로 변경한다.
 3-2. 컬러를 파악할 파일을 workbook 형태로 넘겨준다
    -> 가져올 타겟 row 수를 체크한다.
4. 필요한 데이터를 가져온다.
 4-1. 필요한 columns, rows 만큼 가져온다.                      -> result_journal(저널복붙)
 4-2. 필요한 데이터를 변형해주고, 필요한 columns, rows 만큼 가져온다. -> result_distribution_list(배포처목록)
5. 파일을 각 시트에 저장한다.


### excel_macro_weekly.py
1. os로 해당 경로에 어떤 파일 목록이 있는지 가져온다.
2. 그 파일 목록을 읽는다.
3. 작업할 파일을 openpyxl로 읽는다.
 3-1. DataFrame으로 변경한다.
 3-2. 컬러를 파악할 파일을 workbook 형태로 넘겨준다
    -> 가져올 타겟 row 수를 체크한다.
4. 필요한 데이터를 가져온다.
5. 파일을 저장한다.
