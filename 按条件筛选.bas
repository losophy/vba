Attribute VB_Name = "模块11"
Sub test()
  Debug.Print "F1帮助 F5运行 F8单步调试"    ''我加了调试打印

  Dim r%, i%                ''定义两个Integer变量
  Dim arr, brr              ''定义两个Variant变量
  Dim d As Object           ''定义一个Object对象变量
  
  Set d1 = CreateObject("scripting.dictionary")   ''通过VBS的方法创建一个字典对象
  Set d2 = CreateObject("scripting.dictionary")
  
  With Worksheets("结果表")         ''在结果表中执行
    nf = .Range("k2")               ''把k2数据赋给nf
    bh = Split(.Range("k3"), ",")   ''根据逗号拆分k3字符到bh数组
  End With
  
  For j = 0 To UBound(bh)           ''遍历bh数组
    d1(Val(bh(j))) = ""             ''用k3各数据作为d1字典的key
  Next
  
  With Worksheets("数据表")                             ''在数据表中执行
    r = .Cells(.Rows.Count, 1).End(xlUp).Row            ''取第一列第一次出现的非空单元格的行号,即数据行数
    arr = .Range("a2:d" & r)                            ''将“A2到D2最大已用行号”区域内的所有单元格赋值给arr变量，arr被赋值后，是一个包含这个区域内所有单元格值的数组
    For i = 1 To UBound(arr)                            ''遍历arr数组
      If arr(i, 1) = nf And d1.exists(arr(i, 2)) Then   ''如果有对应年份和编号
        If Not d2.exists(arr(i, 2)) Then                ''检查d2编号是不是为空
          m = 1                                         ''初始化遍历游标为1
          ReDim brr(1 To UBound(arr, 2), 1 To m)        ''初始化brr容器，大小为arr的大小（UBound(arr, 2)为第2列的行数）
        Else                                            ''如果为空就跳过然后遍历下一行
          brr = d2(arr(i, 2))
          m = UBound(brr, 2) + 1
          ReDim Preserve brr(1 To UBound(arr, 2), 1 To m)
        End If
        For j = 1 To UBound(arr, 2)                     ''相当于遍历这一行,把这行数据赋给brr
          brr(j, m) = arr(i, j)
        Next
        d2(arr(i, 2)) = brr                              ''把brr数据赋给d2
      End If
    Next
  End With
  
  With Worksheets("结果表")
    .Range("a3:d" & .Rows.Count).Clear                           ''把A3到D3清空
    For Each aa In d1.keys
      If d2.exists(aa) Then                                      ''如果d2中有d1的key
        brr = d2(aa)
        ReDim crr(1 To UBound(brr, 2), 1 To UBound(brr))
        For i = 1 To UBound(brr)
          For j = 1 To UBound(brr, 2)
            crr(j, i) = brr(i, j)
          Next
        Next
        r = .Cells(.Rows.Count, 1).End(xlUp).Row                 ''数据行数
        If r = 2 Then                                            ''在第下面空白行开始写
          r = r + 1
        Else
          r = r + 2
        End If
        With .Cells(r, 1).Resize(UBound(crr), UBound(crr, 2))    ''填入crr数据，即条件筛选后的数据
          .Value = crr
          .Borders.LineStyle = xlContinuous
        End With
        
      End If
    Next
  End With
  
End Sub
