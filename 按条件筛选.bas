Attribute VB_Name = "ģ��11"
Sub test()
  Debug.Print "F1���� F5���� F8��������"    ''�Ҽ��˵��Դ�ӡ

  Dim r%, i%                ''��������Integer����
  Dim arr, brr              ''��������Variant����
  Dim d As Object           ''����һ��Object�������
  
  Set d1 = CreateObject("scripting.dictionary")   ''ͨ��VBS�ķ�������һ���ֵ����
  Set d2 = CreateObject("scripting.dictionary")
  
  With Worksheets("�����")         ''�ڽ������ִ��
    nf = .Range("k2")               ''��k2���ݸ���nf
    bh = Split(.Range("k3"), ",")   ''���ݶ��Ų��k3�ַ���bh����
  End With
  
  For j = 0 To UBound(bh)           ''����bh����
    d1(Val(bh(j))) = ""             ''��k3��������Ϊd1�ֵ��key
  Next
  
  With Worksheets("���ݱ�")                             ''�����ݱ���ִ��
    r = .Cells(.Rows.Count, 1).End(xlUp).Row            ''ȡ��һ�е�һ�γ��ֵķǿյ�Ԫ����к�,����������
    arr = .Range("a2:d" & r)                            ''����A2��D2��������кš������ڵ����е�Ԫ��ֵ��arr������arr����ֵ����һ������������������е�Ԫ��ֵ������
    For i = 1 To UBound(arr)                            ''����arr����
      If arr(i, 1) = nf And d1.exists(arr(i, 2)) Then   ''����ж�Ӧ��ݺͱ��
        If Not d2.exists(arr(i, 2)) Then                ''���d2����ǲ���Ϊ��
          m = 1                                         ''��ʼ�������α�Ϊ1
          ReDim brr(1 To UBound(arr, 2), 1 To m)        ''��ʼ��brr��������СΪarr�Ĵ�С��UBound(arr, 2)Ϊ��2�е�������
        Else                                            ''���Ϊ�վ�����Ȼ�������һ��
          brr = d2(arr(i, 2))
          m = UBound(brr, 2) + 1
          ReDim Preserve brr(1 To UBound(arr, 2), 1 To m)
        End If
        For j = 1 To UBound(arr, 2)                     ''�൱�ڱ�����һ��,���������ݸ���brr
          brr(j, m) = arr(i, j)
        Next
        d2(arr(i, 2)) = brr                              ''��brr���ݸ���d2
      End If
    Next
  End With
  
  With Worksheets("�����")
    .Range("a3:d" & .Rows.Count).Clear                           ''��A3��D3���
    For Each aa In d1.keys
      If d2.exists(aa) Then                                      ''���d2����d1��key
        brr = d2(aa)
        ReDim crr(1 To UBound(brr, 2), 1 To UBound(brr))
        For i = 1 To UBound(brr)
          For j = 1 To UBound(brr, 2)
            crr(j, i) = brr(i, j)
          Next
        Next
        r = .Cells(.Rows.Count, 1).End(xlUp).Row                 ''��������
        If r = 2 Then                                            ''�ڵ�����հ��п�ʼд
          r = r + 1
        Else
          r = r + 2
        End If
        With .Cells(r, 1).Resize(UBound(crr), UBound(crr, 2))    ''����crr���ݣ�������ɸѡ�������
          .Value = crr
          .Borders.LineStyle = xlContinuous
        End With
        
      End If
    Next
  End With
  
End Sub
