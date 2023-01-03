**PivotLayout**



代表数据透视图报表中字段的位置。

**说明**

使用 **PivotLayout** 属性可返回一个 **PivotLayout** 对象。下例创建在第一个数据透视图报表中所使用的数据透视表字段名称的列表。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `function ListFieldNames() {    let objNewSheet = Worksheets.Add()    let intRow = 1    for(let i=1;i <= Charts.Item("Chart1").PivotLayout.PivotFields().Count;i++) {                  objNewSheet.Cells.Item(intRow, 1).Value2 = Charts.Item("Chart1").PivotLayout.PivotFields(i).Caption            intRow++    }  }` |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | table { word-break:break-all; }返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | table { word-break:break-all; }返回指定对象的父对象。只读。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PivotTable**  | 返回一个 **PivotTable** 对象，它代表与数据透视图相关联的数据透视表。 |

**成员属性**

#### **PivotLayout.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **PivotLayout** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() { 	let myObject = Application.ActiveWorkbook 	if (myObject.Application.Value == "ET") { 		alert("This is an ET Application object.") 	} else { 		alert("This is not an ET Application object.") 	} }` |

#### **PivotLayout.Creator**

table { word-break:break-all; }

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **PivotLayout** 对象的变量。

**说明**

table { word-break:break-all; }

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **PivotLayout.Parent**

table { word-break:break-all; }

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **PivotLayout** 对象的变量。

#### **PivotLayout.PivotTable**

返回一个 **PivotTable** 对象，它代表与数据透视图相关联的数据透视表。

**语法**

**express.PivotTable**

*express*   一个代表 **PivotLayout** 对象的变量。

**示例**

此示例将 Sheet1 上的数据透视表的当前页设置为名为“Canada”的页。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `let pvtTable = Worksheets.Item("Sheet1").Range("A3").**PivotTable** pvtTable.PivotFields("Country").CurrentPage = "Canada"` |

此示例确定与“Sales”图表相关联的数据透视表，然后将名为“Oregon”的页设置为数据透视表的当前页。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `let objPT = Charts.Item("Sales").PivotLayout.**PivotTable** objPT.PivotFields("State").CurrentPageName = "Oregon"` |

适用环境：web

适用平台：windows/linux