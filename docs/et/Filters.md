**Filters**



由多个 **Filter**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Filter/Filter%20.htm#jsObject_Filter)对象组成的集合，这些对象代表自动筛选区域内的所有筛选。

**说明**

使用 **Filters** 属性可返回 **Filters** 集合。下例创建一个列表，其中包含工作表 Crew 中已自动筛选的区域内所有筛选的条件和运算符。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let f     let op     let c1     let c2     let ns = "Not set"      let w = Application.Worksheets.Item("Sheet1")     let w2 = Application.Worksheets.Item("Sheet2")     let rw = 1     for(let i = 1 ; i <= w.AutoFilter.Filters.Count; i++) {         f = w.AutoFilter.Filters.Item(i)         if(f.On) {             let c1 = f.Criteria1.substr(f.Criteria1.length - 1,1)             if(f.Operator) {                 op = f.Operator                 c2 = f.Criteria2.substr(f.Criteria2.length - 1,1)             } else {                 op = ns                 c2 = ns             }         } else {             c1 = ns             op = ns             c2 = ns         }         w2.Cells.Item(rw, 1).Value2 = c1         w2.Cells.Item(rw, 2).Value2 = op         w2.Cells.Item(rw, 3).Value2 = c2         rw = rw + 1     } } ` |

使用 **Filters**(*index*)（其中 *index* 是筛选的标题或索引号）可返回单个 **Filter** 对象。下例将一个变量设置为工作表 Crew 上的筛选区域中第一列的筛选的 **On** 属性的值。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let w = Worksheets.Item("Crew")     if(w.AutoFilterMode) {         let filterIsOn = w.AutoFilter.Filters.Item(1).On     } }` |

**方法**

|                                                              | 名称     | 说明                   |
| ------------------------------------------------------------ | -------- | ---------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item** | 从集合中返回一个对象。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回一个 **Long** 值，它代表集合中对象的数量。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |

**成员方法**

#### **Filters.Item**

从集合中返回一个对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **Filters** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**       |
| -------- | ------------- | ------------ | -------------- |
| *Index*  | 必选          | **Long**     | 对象的索引号。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 下例将一个变量设置为工作表 Crew 上的筛选区域中第一列的筛选的 On 属性的值*/ function test() {     let w = Worksheets.Item("Crew")     if(w.AutoFilterMode) {         filterIsOn = w.AutoFilter.Filters.Item(1).On     } } ` |

**成员属性**

#### **Filters.Application**

如果不使用对象识别符，则该属性返回一个 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **Filters** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例显示一条有关创建 myObject 的应用程序的消息*/ function test() {     let myObject = Application.ActiveWorkbook     if(myObject.Application.Value == "ET") {         alert("This is an ET Application object.")     }     else {         alert("This is not an ET Application object.")     } }` |

#### **Filters.Count**

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

*express*   一个代表 **Filters** 对象的变量。

#### **Filters.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Filters** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **Filters.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **Filters** 对象的变量。

适用环境：web

适用平台：windows/linux