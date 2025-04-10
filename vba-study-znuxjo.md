---
title: VBA学习
date: '2025-04-09 20:25:41'
updated: '2025-04-10 10:13:26'
permalink: /post/vba-study-znuxjo.html
comments: true
toc: true
---



![image](https://raw.githubusercontent.com/DevilGrape2/hexo-blog/main/images/pixabay-436498-20250409203946-eoefftt.png)

# VBA学习

EXCEL VBA 基础[^1]

EXCEL VBA 代码[^2]

Word VBA 基础[^3]

Word VBA 代码[^4]

Catia vba[^5]

‍

[^1]: # EXCEL VBA 基础

    # 1.变量的数据类型

    |数据类型|储存空间|范围|简写|
    | ----------| ----------------------| -----------------------------| ------|
    |byte|1个字节|0到255||
    |Integer|2个字节|-32768到32767|%|
    |Long|4个字节|-2147483648到2147483647|&|
    |String|10个字节加字符串长度|0到大约20亿|$|
    |Date|8个字节|100年1月1日到9999年12月31日||
    |Variant|16个字节|||

    ## 2.声明变量的数据类型

    声明的格式：dim 变量名  as 数据类型

    ```vbnet
    dim n as interger
    ```
    简写：

    ```vbnet
    dim n%
    ```
    声明对个变量用逗号隔开：

    ```vbnet
    Dim s&,n%,x%
    ```
    例子：

    ```vbnet
    Sub test()
    Dim n as Integer
    n = InputBox("请输入总分数")
    MsgBox "你的总分为："&n &"分"
    End Sub
    ```
    ## 3.常用对象的表示方法

    |Workbooks(”工作簿名”)||备注|
    | -------------------------| ----------------| ----------------------------------|
    |ActiveWorkbook|活动工作簿||
    |ThisWorkbook|代码所在工作簿|按工作表的顺序|
    |Sheet(n)|第n个工作表|按系统工作表名|
    |Sheet(”工作表名”)||按工作表名称|
    |ActiveSheet|活动工作表||
    |Range（“单元格地址”）||一个单元格，一行，一列，一个区域|
    |Cells(行，列)|||
    |[A1]单元格简写|||
    |Activecell|活动单元格||
    |Selection|选择的区域||

    ## 4.属性的表达方式

    对象名在前，属性名在后

    语句格式：对象.属性   父对象.子对象.属性

    例子：

    ```vbnet
    Sub test()
    MsgBox Range("a1").value  'a1单元格的值
    MsgBox ActiveWorkbook.Path  '当前工作簿的路径
    MsgBox ActiveCell.Adress  '当前单元格的地址
    Range("a1").Interior.ColorIndex = 35  'a1单元格的颜色改为红色
    End Sub
    ```
    ## 5.对象的操作方法

    对象名在前，方法在后

    语句格式：对象.方法

    例子：

    ```vbnet
    Sub test()
    Workbooks.Add    '新增工作簿
    Workbooks.Open   '打开工作簿
    ActiveWorkbook.Close   '关闭当前激活工作簿
    Worksheets.Add    '新增工作表
    ThisWorkbook.Sheets("演示").Copy ActiveWorkbook  '代码所在工作簿中的演示工作表复制到当前激活工作簿
    Range("a1").Activate              '激活a1单元格
    Range("b1").Copy [a1]                   '将b1单元格复制到a1
    Range("b1").Copy:Range("a1").PasteSpecial XlpasteValues  '仅将b1单元格的值复制到a1
    Range("b1").Clear          '清除b1单元格
    Range("b1").Delete         '删除b1单元格
    Range("b1").Cut [a1]   '将b1单元格的剪切到a1
    End Sub
    ```
    ## 6. IF语句的使用

    ### if(TRUE或者FALSE,”成立”,”不成立”)

    例子：

    ```vbnet
    Sub test()
    Dim n%,x%
    n = 2
    x = 1
    If n>x Then 
       Msgbox "n比x大"
    Else
       Msgbox "x比n大"
    End If
    End Sub

    Sub test()
    Dim n as Byte

    n = InputBox("请输入你的分数")
    If n>60 Then 
       Msgbox "及格"
    Else
       Msgbox "不及格"
    End If
    End Sub
    ```
    ### if嵌套

    例子

    ```vbnet
    Sub test()
    If Range("t2") >= 15000 Then
    Range("g2") = "贵宾"
    ElseIf  Range("t2") >= 10000 Then
    Range("g2") = "高级"
    ElseIf  Range("t2") >= 5000 Then
    Range("g2") = "中级"
    Else
    Range("g2") = "普通"
    End IF
    End Sub
    ```
    ## 7. FOR循环语句

    ### 定义

    for 变量名=x to x

    “循环内容”

    next

    例子：

    ```vbnet
    Sub test()
    Dim n%
    For n = 2 To 19
    If Cells(n,2)<60 Then                '步长为1
    Cells(n,2).Interior.ColorIndex = 3
    End IF
    Next
    End Sub
    ```
    ### 修改步长

    例子

    ```vbnet
    Sub test()
    Dim n%
    For n = 4 To 50 Step 4      '步长为4
      cj =cj + Cells(n,3)
    Next
    MsgBox "英语成绩为："&cj&"分"
    End Sub
    ```
    ### for循环嵌套

    例子

    ```vbnet
    Sub test()
    Dim n%,y%
    For n = 1 To 3
        For y = 1 To 10
         msgbox"外层循环第" & n &"次" & "内层循环第" & y &"次"
        Next y
    Next n
    End Sub
    ```
    ## 8. End动态数据区域（不够智能）

    <aside>  
    💡 一旦有空格出现，会定位到空格的前一个单元格

    </aside>

    |End(xlUp)|上||
    | ----------------| ----------------------------------------------------------| --|
    |End(xlDown)|下||
    |End(xlToLeft)|左||
    |End(xlToRight)|右||
    |row|返回单元格所在行号，如果是区域，就返回这个区域首行的行号||
    |column|列号||
    |rows|代表行的集合，返回rang对象||
    |rows.count|获取最大行号||
    |columns.count|获取最大列号||

    例子

    ```vbnet
    Sub test()
    x = Range("a1").End(xlToRight).Column  '以a1单元格为基准向右获得最右侧有内容的单元格的列号
    h=Range("a1").End(xlDown).Row   '以a1单元格为基准向下获得最下侧有内容的单元格的行号
    End Sub
    ```
    ## 9 UsedRange（较智能）

    <aside>  
    💡 是worksheet的一个属性，代表指定工作表上的所用区域（可能误判）

    </aside>

    格式：工作表.UsedRange.方法或属性

    例子：

    ```vbnet
    Sub test()
    MsgBox ActiveSheet.UsedRange.Rows.count  '当前工作表活动区域的最大行号
    MsgBox ActiveSheet.UsedRange.columns.count  '当前工作表活动区域的最大列号
    End Sub
    ```
    ## 10. Current Region（较智能）

    <aside>  
    💡 需要空行或空列与主表数据隔开

    </aside>

    例子：

    ## 11. for each 循环语句

    ```mermaid
    graph LR
    range -->Range("区域")
    range -->Selection
    range -->usedrang或currentregion返回的区域
    ```
    循环对象合集 workbooks  worksheets

    ```vbnet
    Sub test()
    n = Range("a1").CurrentRegion.Rows.Count
    MsgBox n
    End Sub
    ```
    > for each 变量名 in 对象集合  
    > 循环的内容  
    > next
    >

    例子

    ```vbnet
    Sub test()
    Dim s As Workbook
    For Each s In workbooks      '循环工作簿
    MsgBox s.Name
    Next
    End Sub

    Sub test()
    Dim s As Worksheet
    For Each s In worksheets      '循环工作表
    MsgBox s.Name
    Next
    End Sub

    Sub test()
    Dim s As Range
    For Each s In Range("a1:f14")     '循环单元格
    MsgBox s
    Next
    End Sub

    Sub test()
    Dim s As Range
    For Each s In Selection    '循环选择区域单元格
    MsgBox s
    Next
    End Sub

    Sub test()
    Dim s As Range
    For Each s In Sheets("2").UsedRange   '在工作表2中循环自动选择区域单元格
    MsgBox s
    Next
    End Sub

    Sub test()
    Dim ss as Range,n%
    For Each ss In Range(Sheet1.[b2],Sheet1.Cells(Rows.Count,2).End(xlUp)) 
    n = n + 1
    If ss.value= "男" Then
        Worksheets.Add(after:=Sheets(Sheets.Count)).Name = Sheet1.Cells(n+1,n)
    Next
    End Sub
    ```
    ## 12. 偏移

    以一个单元格为基准，进行偏移，返回的是单元格

    编写格式

    |单元格.offset(偏移行，偏移列)|从0开始（本单元格的行列号为0起算）|上负下正|
    | -------------------------------| ------------------------------------| ----------|
    |单元格(偏移行，偏移列)|从1开始|左负右正|

    例子：

    ```vbnet
    Sub test()
    Range("a1").Offset(8,4).Select    '以（8，4）单元格为原点（0，0）偏移）
    End Sub

    Sub test()
    Range("a1")(8,4).Select   '向左偏移了8向下移动了4（不包括本单元）
    End Sub

    Sub test()
    Dim ss as Range
    For Each ss In Range(Sheet1.[b2],Sheet1.Cells(Rows.Count,2).End(xlUp)) 
    n = n + 1
    If ss.value= "男" Then
        Worksheets.Add(after:=Sheets(Sheets.Count)).Name = Sheet1.Cells.Offset(0,-1)
    Next
    End Sub
    ```
    ## 13. Resize用法

    调整指定选择区域的大小，返回range对象，该对象表示重新定义的区域

    格式：单元格.resize(新区域行数,新区域列数)  从1开始

    例子

    ```vbnet
    Sub test()
    Range("a5","c10").Resize(8,5).Select
    End Sub

    Sub test()
    Dim ss As Range
    For Each ss In Range("c2",Cells(Rows.Count,3).End(xlUp))
        If ss.value < 60 Then
           ss.Offset(0,-2).Resize(1,3).Interior.ColorIndex = 35 '向左偏移了1个单元格后将选定的1个单元格改为1行1列单元格
         End If
    Next ss
    End Sub
    ```
    ## 14. 结束语句Exit

    <aside>  
    💡 Exit语句和End语句不能彼此代替  
    Exit不定义结构的末尾

    </aside>

    编写格式

    |Exit Do|只能写在DO循环里面|
    | ----------| -----------------------|
    |Exit For|只能写在FOR循环里面|
    |Exit Sub|只能写在sub子过程里面|

    例子

    ```vbnet
    Sub test()
    For i =1 To 10
       If i = 5 Then
       Exit For
    Else
    msgbox i
    End If
    Next i
    End Sub
    ```
    ## 15. DO LOOP

    无限循环语句

    编写格式：DO  
                         循环内容……  
                     LOOP

    例子

    ```vbnet
    Sub test()
    On Error Resume Next   '当代码运行错误时忽略，继续向下运行
    Do
    n = n + 1
    if n = 5 Then Exit Do
    MsgBox n
    Loop
    End Sub
    ```
    ## 16. GOTO

    跳转语句

    编写格式：GOTO  1000  
                      .其他内容  
                      100：

    例子：

    ```vbnet
    Sub test()
    Dim n As Date
    On Error Resume Next   '当代码运行错误时忽略，继续向下运行
    Do
    n = InputBox("输入我的生日（yyyy/mm/dd）")
    If Err.number <> 0 Then MsgBox "你输入的格式有误！！"：GoTo 100
    If n =[d1] Then
    MsgBox "回答正确，爱你哦，么么哒"
    Exit Do
    Else
    MsgBox "你连我的生日都忘了，你完蛋了，重新回答
    End IF
    100:

    Err.Clear
    Loop
    End Sub
    ```
    ## 17 Do While loop与 Do Until loop

    编写格式：Do While条件（成立才循环）  
                      循环内容  
                      LOOP  
                    Do Until条件（成立退出循环）  
                    循环内容  
                     LOOP

    ## 18. 使用工作表内函数

    例子：

    ```vbnet
    Sub test()
    Dim n%,i%
    n = 2
    Do While i <> 3
    If cells(n,3) = 100 Then
     cells(n,3).Interior.ColorIndex = 3
     i = i +1
    End If
    n = n +1
    Loop
    End Sub

    Sub test()
    Dim n%,i%
    n = 2
    Do Until i =3
    If cells(n,3) = 100 Then
     cells(n,3).Interior.ColorIndex = 3
     i = i +1
    End If
    n = n +1
    Loop
    End Sub
    ```
    ```vbnet
    Sub test()
    [g2]=Application.WorksheetFunction.AverageIF([b:b],"女",[c:c])
    End Sub
    ```
    ## 19. **在VBA中使用自定义函数**

    ```vbnet
    Function 称呼(x)
    If x = "男" Then
        称呼 = "先生"
    Else
        称呼 = "女士"
    End If
    End Function

    Sub test()
    Dim i,s
    For i =2 To 7
    Set s = Range("B"&i)
    Range("C"&i) = 称呼(s)
    Next

    End Sub

    ```
    ## 20. Rnd 随机数函数

    返回一个小于1但大于等于0的值

    整数区间随机数公式：Int((最大值-最小值+1)*Rnd+最小值)

    例子

    ```vbnet
    Sub test()
    Dim ss As Range
    For Each ss In Range("C2:c500")
    ss = INT((90-35+1)*RND+35)
    Next ss
    End Sub
    ```
    ## 21.排序

    ### 语法

    单元格对象.Sort(Key1,Order1,Key2,Type,Order2,Key2,Type,Order3,Header,OrderCustom,MatchCase,Orientation,SortMethod,DataOption1,,DataOption2,DataOption3)

    ### 参数讲解

    1. Key1、Key2、Key3排序关键列 可以用这一列的某个单元格表示，比如排序A列，用range(”a1”)。至少使用一个key，最多使用3个，最多可以3列多重排序
    2. Order1、Order2、Order3排序模式，默认升序，Order1:=xlAscending 则key1升序，简写Order1:=1,Order1:=xlDescending 则key1降序，简写Order1:=2
    3. Type 指定要排序的元素，排序数据透视表时使用，xlSortLabels按标签对数据透视表排序，xlSortvalues按值对数据透视表排序
    4. Header排序区域是否有表头？Header:=xlGuess 让软件自己辨认，简写Header:=0 ,Header:=xlYes 有表头，简写Header:=1（第一行不参与排序），Header:=xlNo 没有表头，简写Header:=0（第一行参与排序）

    ## 22. **清除**

    |**代码**|**作用**|
    | --| --|
    |**r.Clear**|**清除所有内容（包括批注、内容、格式、超链等）**|
    |**r.ClearComments**|**清除批注**|
    |**r.ClearContents**|**清除内容**|
    |**r.ClearFormats**|**清除格式**|
    |**r.ClearHyperlinks**|**清除超链接**|

    ### **字体**

    ```vbnet
    **r.Font.Clolr=RGB(255,0,0)     '文字颜色
    r.Font.Size =24           '文字大小
    r.Font.Italic = True     '是否斜体
    r.Font.Bold = True      '是否粗体**
    ```
    ### **使用With精简代码**

    ```vbnet
    **Sub a()
    Dim r
    Set r = Range("A1:A10")
    With r.Font
          .Clolr = RGB(255, 0, 0)
          .Size = 24
          .Italic = True
          .Bold = True
    End With
    End Sub**
    ```
    ### **内部属性**

    ```vbnet
    **r.Interior.Color=RGB(255,0,0)**
    ```
    **MsgBox与InputBo**

    ```vbnet
    Sub f()
    Dim i
    i = InputBox("请输入您的姓名：")
    Range("K1") = i
    MsgBox "您好" & i & "欢迎回来！"
    End Sub
    ```
    **VBA中调用Excel公式和错误处理**

    ```vbnet
    一、四舍五入
    Sub a()
    Dim i, j
    i = 3.1415926
    j = Excel.Application.WorksheetFunction.Round(i, 2)
    MsgBox j
    End Sub

    二、统计数量（多张工作表，用for循环，sheet(i)）
    Sub a()
    Dim a
    a = Excel.Application.WorksheetFunction.CountA(Range("A:A")) - 1
    MsgBox a
    End Sub

    三、条件计数
    Sub a()
    Dim i, a, b, c, x, y
    For i = 2 To Sheets.Count
        Set x = Sheets(i).Range("A:A")
        Set y = Sheets(i).Range("B:B")
        With Excel.Application.WorksheetFunction
            a = a + .CountA(x) - 1
            b = b + .CountIf(y, "男")
            c = c + .CountIf(y, "女")
        End With
    Next
    Range("B1") = a
    Range("B2") = b
    Range("B3") = c
    End Sub

    四、VLOOKUP
    Sub a()
    On Error Resume Next
    Dim j, i
    j = 2
    Do While Range("A" & j) <> 0
        For i = 2 To Sheets.Count
            Range("B" & j) = Excel.Application.WorksheetFunction.VLookup(Range("A" & j), Sheets(i).Range("A:B"), 2, 0)
        Next
        j = j + 1
    Loop
    End Sub

    拓展：考生成绩统计&查询系统
    Sub 查询()
    On Error Resume Next
    Dim i, a, b, c
    Sheets("汇总").Range("D14").ClearContents
    For i = 2 To Sheets.Count
        With Excel.Application.WorksheetFunction
            Set a = Sheets("汇总").Range("D9")
            Set b = Sheets(i).Range("A:H")
            Set c = Sheets("汇总")
            c.Range("D14") = .VLookup(a, b, 5, 0) '姓名
            c.Range("D16") = .VLookup(a, b, 6, 0) '性别
            c.Range("D18") = .VLookup(a, b, 3, 0) '专业类
            c.Range("D20") = .VLookup(a, b, 8, 0) '总分
            '在哪张表上找到数据就显示他的表名
            c.Range("D22") = Sheets(i).Name
            '如果汇总表的D14姓名不为空时就停止循环
            If c.Range("D14") <> "" Then
                Exit For
            End If
        End With
    Next
    End Sub

    Sub 统计()
    Dim i, a, b
    For i = 2 To Sheets.Count
        With Excel.Application.WorksheetFunction
        Set a = Sheets("汇总")
        Set b = Sheets(i)
        a.Range("D26") = .CountA(b.Range("A:A")) - 1
        a.Range("D27") = .CountIf(b.Range("F:F"), "男")
        a.Range("D28") = .CountIf(b.Range("F:F"), "女")
        End With
    Next
    End Sub
    ```
    ## **正则表达式**

    ### **元字符与特殊字符**

    |**元字符**|**描述**|
    | -------| ------------------------------------------------------------------------------|
    | **.**|**句号匹配任意单个字符除了换行符**|
    | **[]**|**字符种类，匹配方括号内的任意字符，中括号内每个字符是或(or)的关系**|
    | **[^]**|**否定的字符种类，匹配除了方括号里的任意字符**|
    |*****|**匹配0次或无限次，重复在*号之前的字符**|
    | **+**|**匹配1次或无限次，重复在+号之前的字符**|
    | **?**|**匹配0次或1次，重复在?号之前的字符**|
    | **{n}**|**正好出现n次**|
    | **{n,m}**|**匹配num个大括号之前的字符，出现n到m次（n&lt;=num&lt;=m）**|
    | **(xyz)**|**字符集又称做组，匹配与xyz完全相等的字符串，每个字符是且(and)的关系**|
    |**|**|
    |* ***|**转义字符，用于匹配一些保留字符 [ ]、( )、{ }、. 、 * 、+ 、? 、^ 、$、\ 、|
    | **^**|**从字符串开始位置开始匹配**|
    | **$**|**从字符串末端开始匹配**|

    **反斜杠后面跟普通字符实现特殊功能**

    |**特殊字符**|**描述**|
    | --| --|
    | **\d**|**匹配数字，相当于[0-9]**|
    | **\D**|**不匹配数字，相当于[^0-9]**|
    | **\s**|**匹配空白字符(包括空格、换行符、制表符等)，相当于 [\t\n\r\f\v]**|
    | **\S**|**与\s相反，相当于 [^\t\n\r\f\v]**|
    | **\w**|**匹配中文，下划线，数字，英文，相当于[a-zA-z0-9_]**|
    | **\W**|**与\w相反，匹配特殊字符，如$、&amp;、空格、\n、\t等**|

[^2]: # EXCEL VBA 代码

    记录单元格操作

    ```vb.net
    Dim yz
    Private Sub 
    ```
    自动保存

    ```vb.net
    Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ThisWorkbook.Save
    End Sub
    ```
    ![捕获.png](https://raw.githubusercontent.com/DevilGrape2/hexo-blog/main/images/0de63de173a511901d134dbbfc51d2ad.png)

    自动备份

    ```vb.net
    Private Sub Workbook_Open()
    Dim wb As Workbook
    Set wb = ThisWorkbook
    t = Format(Now(), "hhmmss")
    wb.SaveAs wb.Path & "\" & t & wb.Name

    End Sub
    ```
    打开工作簿时调用窗体

    ```vb.net
    Private Sub Workbook_Open()
    UserForm1.Show
    End Sub
    ```
    ```plain
    Dim MatLab As Object
    Dim Result As String
    Dim MReal(1, 3) As Double
    Dim MImag(1, 3) As Double

    MatLab = CreateObject("Matlab.Application")

    'Call MATLAB function from VB
    Result = MatLab.Execute("surf(peaks)")

    'Execute simple computation
    Result = MatLab.Execute("a = [1 2 3 4; 5 6 7 8]")
    Result = MatLab.Execute("b = a + a ")

    'Bring matrix b into VB program
    MatLab.GetFullMatrix("b", "base", MReal, MImag)
    ```

[^3]: # Word VBA 基础

    # **Application的常用对象**

    |**常用对象**|**说明**|
    | --| --|
    |**Application.ActiveDocument**|**当前文档，可以简写为ActiveDocument**|
    |**Application.ActivePrinter**|**获取当前打印机**|
    |**Application.ActiveWindows**|**当前窗口**|
    |**Application.Height**|**当前应用文档的高度**|
    |**Application.Width**|**当前应用文档的宽度**|
    |**Application.Build**|**获取Word版本号和编译序号**|
    |**Application.Caption**|**当前应用程序名**|
    |**Application.DefaultSaveFormat**|**返回空字符串，表示Word文档**|
    |**Application.DisplayRecentFiles**|**返回是否显示最近使用的文档的状态**|
    |**Application.Documents.Count**|**返回打开的文档数**|
    |**Application.FontNames.Count**|**返回当前可用的字体数**|
    |**Application.Left**|**返回当前文档的水平位置**|
    |**Application.MacroContainer.FullName**|**返回当前文档名，包括所在路径**|
    |**Application.NormalTemplate.FullName**|**返回文档标准模版名称及所在位置**|
    |**Application.Path**|**显示活动文档的路径和文件名**|
    |**Application.RecentFiles.Count**|**返回最近打开的文档数目**|
    |**Application.System.FreeDiskSpace**|**返回应用程序所在磁盘可用空间**|
    |**Application.Templates.Count**|**返回应用程序所使用的模板数**|
    |**Application.UserName**|**返回应用程序用户名**|
    |**Application.Version**|**返回应用程序的版本号**|
    |**Application.Activate**|**激活指定对象**|
    |**Application.Move**|**设置任务窗口或活动文档窗口的位置**|
    |**Application.GoForward**|**将插入在活动文档中进行编辑的最后三个位置之间向前移动**|
    |**Application.PrintOut**|**该方法可打印指定文档的全部或部分**|
    |**Application.Resize**|**调整Word窗口大小。如果该窗口被最大化或最小化将导致出错**|
    |**Application.Quit**|**退出Word，并可选择保存或传送打开的文档**|

    # **Document的常用对象**

    |**参数**|**中文**|
    | --| --|
    |**ActiveDocument.AttachedTemplate.FullName**|**返回当前文档采用模板名及模板所在位置**|
    |**ActiveDocument.Bookmarks.Count**|**返回当前文档中的书签数**|
    |**ActiveDocument.Characters.Count**|**返回当前文档的字符数**|
    |**ActiveDocument.Comments.Count**|**返回当前文档的批注数**|
    |**ActiveDocument.Endnotes.Count**|**返回当前文档的尾注数**|
    |**ActiveDocument.Fields.Count**|**返回当前文档的域数目**|
    |**ActiveDocument.Footnotes.Count**|**返回当前文档中的脚注数**|
    |**ActiveDocument.FullName**|**返回当前文档的全名及所在位置**|
    |**ActiveDocument.HasPassword**|**判断当前文档是否有密码保护**|
    |**ActiveDocument.Hyperlinks.Count**|**返回当前文档中的链接数**|
    |**ActiveDocument.Indexes.Count**|**返回当前文档中的索引数**|
    |**ActiveDocument.ListParagraphs.Count**|**返回当前文档中项目编号或项目符号数**|
    |**ActiveDocument.PageSetup**|**文档内的页面设置**|
    |**ActiveDocument.Paragraphs.Count**|**返回当前文档中的段落数**|
    |**ActiveDocument.Password = xxx**|**设置打开文件使用的密码**|
    |**ActiveDocument.Path**|**文档所在路径**|
    |**ActiveDocument.ReadOnly**|**获取当前文档是否为只读属性**|
    |**ActiveDocument.Saved**|**当前文档是否被保存**|
    |**ActiveDocument.Sections.Count**|**当前文档中的节数**|
    |**ActiveDocument.Sentences.Count**|**当前文档中的语句数**|
    |**ActiveDocument.Shapes.Count**|**当前文档中的形状数**|
    |**ActiveDocument.Styles.Count**|**当前文档中的样式数**|
    |**ActiveDocument.Tables.Count**|**当前文档中的表格数**|
    |**ActiveDocument.TablesOfAuthorities.Count**|**返回当前文档中的引文目录数**|
    |**ActiveDocument.TableOfContents.Count**|**返回当前文档中的目录数**|
    |**ActiveDocument.TablesOfFigures.Count**|**返回当前文档中的图表目录数**|
    |**ActiveDocument.Words.Count**|**返回当前文档中字词数**|

    # **Documents的常用对象**

    |**参数**|**中文**|
    | --| --|
    |**Documents.Add**|**表示添加至打开的文档集合中新建空文档**|
    |**Documents.Close**|**关闭指定的一个或多个文档**|
    |**Documents.Item (indexs)**|**表示第indexs文档**|
    |**Documents.Open**|**打开指定的文档并将其添加至Documents集合**|
    |**Documents.Save**|**保存指定文档及其说明**|

[^4]: # Word VBA 代码

    **Word批量转PDF**

    ```vb.net
    Sub 批量()
    名称 = Dir("C:\孙兴华\")
    Do While 名称 <> ""
        Set 文档 = Word.Application.Documents.Open("C:\孙兴华\" & 名称)
        路径 = 文档.Path
        文档.ExportAsFixedFormat OutputFileName:=(路径 & "\" & 文档.Name & ".pdf"), _
                ExportFormat:=wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:= _
                wdExportOptimizeForPrint, Range:=wdExportFromTo, From:=1, To:=文档.Range.Information(wdNumberOfPagesInDocument), _
                Item:=wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, _
                CreateBookmarks:=wdExportCreateHeadingBookmarks, DocStructureTags:=True, _
                BitmapMissingFonts:=True
        
        文档.Save
        文档.Close
        名称 = Dir
    Loop
    End Sub
    ```
    **使用Excel公式**

    ```vb.net
    Sub 使用公式()
    Set 起始位置 = ActiveDocument.Range(0, 0)
    Set 表格 = ActiveDocument.Tables.Add(起始位置, 3, 3)
    With 表格
        .Cell(1, 1).Range.InsertAfter "10"
        .Cell(2, 1).Range.InsertAfter "20"
        .Cell(3, 1).Formula "=Average(Above)"
    End With
    End Sub
    ```

[^5]: # Catia vba

    提取工程图中的尺寸并保存到 excel

    ```vb.net
    Sub catiadaochuchicun()

    '定义数据类型
    'Catia文档类型
    Dim doc As DrawingDocument
    Dim sheets As DrawingSheets
    Dim sheet As DrawingSheet
    Dim views As DrawingViews
    Dim view As DrawingView
    Dim dimensions As DrawingDimensions

    '初始化
    Dim dn  As Integer
    Dim ex As Object
    Dim dX  As Integer

    '定义公差数据类型
    Dim oTolType As Long
    Dim oDisplayMode As Long
    Dim oTolName As String
    Dim oUpTolS As String
    Dim oLowTolS As String
    Dim oUpTolD As Double
    Dim oLowTolD As Double


    Set doc = CATIA.ActiveDocument
    Set sheets = doc.sheets
    sheetscount = sheets.Count
    Set sheet = sheets.ActiveSheet
    sheetscount = sheets.Count
    Set views = sheet.views
    viewscount = views.Count '视图数量


    '计算当前页面中尺寸的数量
    dn = 0
    For i = 1 To viewscount
        Set view = views.Item(i)
        Set dimensions = view.dimensions
        dn = dimensions.Count + dn
    Next

    '定义动态数组用于存储尺寸数据
    Dim myvlaue() As Double
    ReDim myvlaue(1 To dn, 1 To dn)

    Dim shangcha() As String
    ReDim shangcha(1 To dn, 1 To dn)

    Dim xiacha() As String
    ReDim xiacha(1 To dn, 1 To dn)


    '在动态数组中存储数据
    Set ex = CreateObject("Excel.Application")

    Set exwbook = ex.Workbooks().Add
    Set exsheet = exwbook.Worksheets("sheet1")

    '在excel里表格的表头

    ex.Range("a2").Value = "序号"
    ex.Range("b2").Value = "尺寸数据"
    ex.Range("c2").Value = "上差"
    ex.Range("d2").Value = "下差"
    'ex.Range("e2").Value = "单位"


    '提取尺寸数据及公差并写入excel
    dX = 0
    For J = 1 To viewscount
        Set view = views.Item(J)
        Set dimensions = view.dimensions
        DT = dimensions.Count
        For A = 1 To DT
            Set dimension = dimensions.Item(A)
            Number = dimension.GetValue.Value
            oUpTolD = 0
            oLowTolD = 0
            dimension.GetTolerances oTolType, oTolName, oUpTolS, oLowTolS, oUpTolD, oLowTolD, oDisplayMode
            myvlaue(J, A) = Number
            shangcha(J, A) = oUpTolD
            xiacha(J, A) = oLowTolD
             'MsgBox myvlaue(J, A)
            ' MsgBox A
            ex.Range("b" & A + dX).Value = myvlaue(J, A)
            ex.Range("c" & A + dX).Value = shangcha(J, A)
            ex.Range("d" & A + dX).Value = xiacha(J, A)
            'If A = DT Then
            'ex.Range("b" & A + 3 + dX).Clear
            'ex.Range("c" & A + 3 + dX).Clear
            'ex.Range("d" & A + 3 + dX).Clear
            'End If
        Next
        dX = dimensions.Count + dX + 1
    Next



    For i = 1 To dn + viewscount - 3
        ex.Range("a" & i + 2).Value = i

    Next

    exwbook.SaveAs "F:\OneDrive\工作\catia插件\CATIA 插件\记录.xls"
    ex.Quit

    End


    End Sub
    ```
    ‍
