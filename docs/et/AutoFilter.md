**AutoFilter**



代表对指定工作表的自动筛选。

**说明**

使用 **AutoFilter** 属性可返回 **AutoFilter** 对象。使用 **Filters** 属性可返回由各个列筛选组成的集合。使用 **Range** 属性可返回代表整个筛选区域的 **Range** 对象。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Dim w As Worksheet Dim filterArray() Dim currentFiltRange As String  Sub ChangeFilters()  Set w = Worksheets("Crew") With w.AutoFilter     currentFiltRange = .Range.Address     With .Filters         ReDim filterArray(1 To .Count, 1 To 3)         For f = 1 To .Count             With .Item(f)                 If .On Then                     filterArray(f, 1) = .Criteria1                     If .Operator Then                         filterArray(f, 2) = .Operator                         filterArray(f, 3) = .Criteria2                     End If                 End If             End With         Next     End With End With  w.AutoFilterMode = False w.Range("A1").AutoFilter field:=1, Criteria1:="S"  End Sub` |

要为工作表创建 **AutoFilter** 对象，必须手动打开或使用 **Range** 对象的 **AutoFilter** 方法打开工作表上某个区域上的自动筛选功能。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Sub RestoreFilters() Set w = Worksheets("Crew") w.AutoFilterMode = False For col = 1 To UBound(filterArray(), 1)     If Not IsEmpty(filterArray(col, 1)) Then         If filterArray(col, 2) Then             w.Range(currentFiltRange).AutoFilter field:=col, _                 Criteria1:=filterArray(col, 1), _                     Operator:=filterArray(col, 2), _                 Criteria2:=filterArray(col, 3)         Else             w.Range(currentFiltRange).AutoFilter field:=col, _                 Criteria1:=filterArray(col, 1)         End If     End If Next End ` |

| 注释                                                         |
| ------------------------------------------------------------ |
| When using **AutoFilter** with dates, the format should be consistent with English date separators ("/") instead of local settings ("."). A valid date would be "2/2/2007", whereas "2.2.2007" is invalid. |

**方法**

|                                                              | 名称            | 说明                                     |
| ------------------------------------------------------------ | --------------- | ---------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ApplyFilter** | 应用指定的 **Autofilter** 对象。         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ShowAllData** | 显示 **AutoFilter** 对象返回的所有数据。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FilterMode**  | 如果工作表的筛选模式为自动筛选，则返回 **True**。只读 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Filters**     | 返回一个 **Filters** 集合，该集合表示自动筛选区域中的所有筛选器。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Range**       | 返回一个 **Range** 对象，它代表应用指定的“自动筛选”的区域。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Sort**        | 获取 **AutoFilter** 集合的排序列和排序次序。                 |

**成员方法**

#### **AutoFilter.ApplyFilter**

应用指定的 **Autofilter** 对象。

**语法**

**express.ApplyFilter()**

*express*   一个代表 **AutoFilter** 对象的变量。

#### **AutoFilter.ShowAllData**

显示 **AutoFilter** 对象返回的所有数据。

**语法**

**express.ShowAllData()**

*express*   一个代表 **AutoFilter** 对象的变量。

**成员属性**

#### **AutoFilter.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **AutoFilter** 对象的变量。

#### **AutoFilter.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **AutoFilter** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **AutoFilter.FilterMode**

如果工作表的筛选模式为自动筛选，则返回 **True**。只读 **Boolean** 类型。

**语法**

**express.FilterMode**

*express*   一个代表 **AutoFilter** 对象的变量。

#### **AutoFilter.Filters**

返回一个 **Filters** 集合，该集合表示自动筛选区域中的所有筛选器。只读。

**语法**

**express.Filters**

*express*   一个代表 **AutoFilter** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例将变量设置为工作表“Crew”上的筛选区域中第一列的筛选的 Criteria1 属性值。*/ function test() {   let mySheet = Application.Worksheets.Item("Crew")   if(mySheet.AutoFilterMode){       if(mySheet.AutoFilter.Filters.Item(1).On){           let c1 = mySheet.AutoFilter.Filters.Item(1).Criteria1       }   } }` |

#### **AutoFilter.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **AutoFilter** 对象的变量。

#### **AutoFilter.Range**

返回一个 **Range** 对象，它代表应用指定的“自动筛选”的区域。

**语法**

**express.Range**

*express*   一个代表 **AutoFilter** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例将应用于“Crew”工作表的“自动筛选”的地址保存在变量中。*/ let rAddress = Application.Worksheets.Item("Crew").AutoFilter.Range.Address /*此示例滚动工作簿窗口，直至超链接区域位于活动窗口的左上角。*/ function test() {   Application.Workbooks.Item(1).Activate()      let myRange = Application.ActiveSheet.Hyperlinks.Item(1).Range   Application.ActiveWindow.ScrollRow = myRange.Row   Application.ActiveWindow.ScrollColumn = myRange.Column }` |

#### **AutoFilter.Sort**

获取 **AutoFilter** 集合的排序列和排序次序。

**语法**

**express.Sort**

*express*   一个代表 **AutoFilter** 对象的变量。

适用环境：web

适用平台：windows/linux