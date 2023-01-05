**Filter**



代表单个列的筛选。

**说明**

**Filter** 对象是 [**Filters** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Filters/Filters%20.htm#jsObject_Filters)集合的成员。**Filters** 集合包含自动筛选区域中的所有筛选。

使用 **Filters**( *index*)（其中 *index* 是筛选的标题或索引号）可返回单个 **Filter** 对象。下例将一个变量设置为工作表 Crew 上的筛选区域中第一列的筛选的 **On** 属性的值。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let w = Application.Worksheets.Item("Crew")     if(w.AutoFilterMode) {         filterIsOn = w.AutoFilter.Filters.Item(1).On     } }` |

注意：**Filter** 对象的所有属性都是只读的。要设置这些属性，请手动应用自动筛选，或使用 [**Range** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)对象的 **AutoFilter** 方法，如下例中所示。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let w = Application.Worksheets.Item("Crew")          w.Cells.AutoFilter(2, "Crucial", xlOr, "Important") }` |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回集合中对象的数目。只读 **Long** 类型。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Criteria1**   | 返回筛选区域内指定列上的第一个筛选值。**Variant** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Criteria2**   | 返回已筛选区域中指定列的第二个筛选值。**Variant** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **On**          | 如果打开指定的筛选，则该属性值为 **True**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Operator**    | 返回一个 **XlAutoFilterOperator**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/XlAutoFilterOperator%20%E6%9E%9A%E4%B8%BE.html)值，它代表与指定筛选所应用的两个条件相关的操作符。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |

**成员属性**

#### **Filter.Application**

如果不使用对象识别符，则该属性返回一个 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **Filter** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例显示一条有关创建 myObject 的应用程序的消息*/ function test() {     let myObject = Application.ActiveWorkbook     if(myObject.Application.Value == "ET") {         alert("This is an ET Application object.")     }     else {         alert("This is not an ET Application object.")     } }` |

#### **Filter.Count**

返回集合中对象的数目。只读 **Long** 类型。

**语法**

**express.Count**

*express*   一个代表 **Filter** 对象的变量。

#### **Filter.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Filter** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **Filter.Criteria1**

返回筛选区域内指定列上的第一个筛选值。**Variant** 类型，只读。

**语法**

**express.Criteria1**

*express*   一个代表 **Filter** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 下例将变量设置为工作表“Crew”上的筛选区域中第一列的筛选的 Criteria1 属性值*/ function test() {     let c1 = Application.Worksheets.Item("Crew")     if(c1.AutoFilterMode) {         let c2 = c1.AutoFilter.Filters.Item(1)             if(c2.On) {                  c1 = c2.Criteria1             }     } }` |

#### **Filter.Criteria2**

返回已筛选区域中指定列的第二个筛选值。**Variant** 类型，只读。

**语法**

**express.Criteria2**

*express*   一个代表 **Filter** 对象的变量。

**说明**

使用筛选的 **Criteria2** 属性时，如果该筛选没有使用两个条件，则将出现错误。在试图使用 **Operator** 属性之前，需保证 **Filter** 对象的 **Criteria2** 属性不为零 (0)。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 下面的示例将一个变量设置为 Criteria2 属性值。该属性是 Crew 工作表上已筛选区域中的第一列筛选的*/ function test() {     let c1 = Application.Worksheets.Item("Crew")     if(c1.AutoFilterMode) {         let c2 = c1.AutoFilter.Filters.Item(1)             if(c2.On && c2.Operator) {                 let c3 = c2.Criteria2             }             else {                 let c3 = "Not set"             }     } }` |

#### **Filter.On**

如果打开指定的筛选，则该属性值为 **True**。**Boolean** 类型，只读。

**语法**

**express.On**

*express*   一个代表 **Filter** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 下例将变量设置为工作表“Crew”上的筛选区域中第一列的筛选的 Criteria1 属性值*/ function test() {     let on1 = Application.Worksheets.Item("Crew")     if(on1.AutoFilterMode) {         let on2 = on1.AutoFilter.Filters.Item(1)             if(on2.On) {                 let c1 = on2.Criteria1             }     } } ` |

#### **Filter.Operator**

返回一个 **XlAutoFilterOperator**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/XlAutoFilterOperator%20%E6%9E%9A%E4%B8%BE.html)值，它代表与指定筛选所应用的两个条件相关的操作符。

**语法**

**express.Operator**

*express*   一个代表 **Filter** 对象的变量。

#### **Filter.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **Filter** 对象的变量。

适用环境：web

适用平台：windows/linux