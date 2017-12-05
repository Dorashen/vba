'需求和目标：将项目计划表内容+日期 对应到另一张日程表里（日程表类似于一个日历，在日历里填写相应的内容）
'目前只能做到自己做一张日历表，然后跟项目计划表放在同一个文件夹里，每次修改计划表，点击excel上的更新按钮后，自动打开日历表并更新到上面
'结合发布计划，对后续日程表进行实时更新，即将预计发布时间，填写到日程表中对应的日期里
'实现步骤： 1.修改好后，点击按钮
'           2.清空后续列表中除了第一行的所有发布内容
'           3.截取发布计划中填写的日期
'           4.查询后续列表中等于该日期的 日程单元格
'           5.数据插入
'           6.美化（添加序号，字体加粗，变小）
'注: Sheets (1) sheets(2)等 中间的数字表示表所在的位置序号，而不是建立表格的顺序，以下全程识别的是 表所在的位置，所以发布计划表需一直放在第一个位置

Sub click()
  '0 定义变量，
  Set date_table = Workbooks.Open(ThisWorkbook.Path & "\项目发布日程表.xlsx") '设置路程表的路径 和此表放在同一文件夹内
  Dim i As Integer '日程表的循环参数
  Dim j As Integer '发布计划中的循环参数，同时作为行号
  Dim a As Integer '日程表 对应单元格的行号
  Dim b As Integer '日程表 对应单元格的列号
  Dim all_time_date As String '输入的预计发布时间，将日期提取出来
    
  '1.先清空所有数据
  For i = 1 To date_table.Worksheets.Count  '从第1个表格开始 循环查询后续列表，逐一清空
    For Each date_time In date_table.Sheets(i).Range("A2:G8")  '查询每个表A2 到G8的元素 ，每个月表不会超过此区间
      If date_time.Value <> "" Then  '如果单元格不为空
        a = date_time.Row  '获取行列值
        b = date_time.Column
        date_table.Sheets(i).Cells(a, b) = Split(date_time, Chr(10))(0) '只保留第一行数据（出现的一个换行符前的字符，提取出来，（0）表示数组第一个元素，split为分离函数）
      End If
    Next
  Next
  '2.数据和日期匹配
  '2(1)获取输入的日期
  RowCount = ThisWorkbook.Sheets(1).UsedRange.Rows.Count '获取计划表中 表1当前一共有多少行
    For j = 2 To RowCount   '从第2行开始，全部行查询过去，逐一更新过去 
      all_time = ThisWorkbook.Sheets(1).Cells(j, 8) '获取 预计发布时间的值
      blank_order = InStr(1, all_time, " ", 1) '获取 预计发布时间中 空格的位置，从第一个字符开始，搜索空格，末尾1表示 用原文比较的方式 instr函数作用：获取某字符的位置
      If all_time <> "" And blank_order <> 0 Then  '如果预计发布时间不为空，且有空格，比如8.17 22：30   
        all_time_date = Mid(all_time, 1, blank_order - 1)  '截取空格前的日期值  从第一个字符开始，截取到空格前的那个字符
      ElseIf all_time = "" Then  '如果为空，则all_time_date 为空
        all_time_date = ""
      ElseIf all_time <> "" And blank_order = 0 Then '如果查找不到空格 即只输入了日期 没输入时间则 如8.17
        all_time_date = all_time  '直接赋值过去
      End If 
    '2(2)输入的日期与后续表格的日期相匹配，插入数据
     If all_time_date <> "" Then  '当预计发布时间不为空时，更新到后续表格中
        For i = 1 To date_table.Worksheets.Count   ' i=1 从第1个表格开始，直到最后，做循环查询
          For Each date_time In date_table.Sheets(i).Range("A2:G8")  '查询每个表A2 到G8的元素
            If date_time.Value <> "" Then  '如果单元格不为空，再继续，否则可能出错
                If (Split(date_time, Chr(10))(0) = all_time_date) Then  '判断日期是否和预计发布时间相等，相等则插入数据
                    a = date_time.Row '获取行列号
                    b = date_time.Column
                    count_array = Split(date_table.Sheets(i).Cells(a, b), Chr(10))   '先判断 用换行符分隔出来的数组，一共有多少元素，将split函数得到的是数组
                    Count = UBound(count_array) + 1  '查看这个单元格一共有几行，因为数组是从0开始，所以要+1
                    date_table.Sheets(i).Cells(a, b) = date_table.Sheets(i).Cells(a, b) & Chr(10) & Count & ")" & ThisWorkbook.Sheets(1).Range("A" & j) '在当前行的下一列，插入数据 并添加序号
                    '以下是美化，对第一行以后的数据 ，进行字体加粗，字体变小，颜色变黑
                    alt_order = InStr(1, date_time, Chr(10), 1) '获取日程表中每个日期中第一个 换行符的位置, 从第一个字符开始，搜索第一个换行符，用原文比较的方式
                    date_table.Sheets(i).Cells(a, b).Characters(Start:=alt_order + 1, Length:=Len(date_time)).Font.Bold = True '字体加粗
                    date_table.Sheets(i).Cells(a, b).Characters(Start:=alt_order + 1, Length:=Len(date_time)).Font.Color = vbBlack '字体变黑
                    date_table.Sheets(i).Cells(a, b).Characters(Start:=alt_order + 1, Length:=Len(date_time)).Font.Size = 10 '字体变小
                End If
            End If
          Next
        Next
      End If
    Next
  date_table.Save
End Sub
