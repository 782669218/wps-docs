**DataTable**



代表一张图表模拟运算表。

**说明**

使用 **DataTable** 属性可返回 **DataTable** 对象。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例向嵌入式图表中添加带有外边框的模拟运算表。*/ function test(){ let myChart = Application.Worksheets.Item(1).ChartObjects(1).Chart     myChart.HasDataTable = true     myChart.DataTable.HasBorderOutline = true }` |

**方法**

|                                                              | 名称       | 说明       |
| ------------------------------------------------------------ | ---------- | ---------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete** | 删除对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Select** | 选择对象。 |

**属性**

|                                                              | 名称                    | 说明                                                         |
| ------------------------------------------------------------ | ----------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**         | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Border**              | 返回一个 **Border** 对象，它代表对象的边框。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**             | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Font**                | 返回一个 **Font** 对象，它代表指定对象的字体。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Format**              | 返回 **ChartFormat** 对象。只读。                            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HasBorderHorizontal** | 如果图表模拟运算表具有水平单元格边框，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HasBorderOutline**    | 如果图表模拟运算表具有外部边框，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HasBorderVertical**   | 如果图表模拟运算表具有垂直单元格边框，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**              | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowLegendKey**       | 如果数据标签图例项标示可见，则为 **True**。**Boolean** 类型，可读写。 |

**成员方法**

#### **DataTable.Delete**

删除对象。

**语法**

**express.Delete()**

*express*   一个代表 **DataTable** 对象的变量。

#### **DataTable.Select**

选择对象。

**语法**

**express.Select()**

*express*   一个代表 **DataTable** 对象的变量。

**成员属性**

#### **DataTable.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **DataTable** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() { 	let myObject = Application.ActiveWorkbook 	if (myObject.Application.Value == "ET") { 		alert("This is an ET Application object.") 	} else { 		alert("This is not an ET Application object.") 	} }` |

#### **DataTable.Border**

返回一个 **Border** 对象，它代表对象的边框。

**语法**

**express.Border**

*express*   一个代表 **DataTable** 对象的变量。

#### **DataTable.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **DataTable** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **DataTable.Font**

返回一个 **Font** 对象，它代表指定对象的字体。

**语法**

**express.Font**

*express*   一个代表 **DataTable** 对象的变量。

#### **DataTable.Format**

返回 **ChartFormat** 对象。只读。

**语法**

**express.Format**

*express*   一个代表 **DataTable** 对象的变量。

#### **DataTable.HasBorderHorizontal**

如果图表模拟运算表具有水平单元格边框，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.HasBorderHorizontal**

*express*   一个代表 **DataTable** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例使嵌入图表模拟运算表显示时带有外边框，但无单元格边框。*/ function test(){ let myChart = Application.Worksheets.Item(1).ChartObjects(1).Chart     myChart.HasDataTable = true     let myDataTable = myChart.DataTable         myDataTable.HasBorderHorizontal = false         myDataTable.HasBorderVertical = false         myDataTable.HasBorderOutline = true }` |

#### **DataTable.HasBorderOutline**

如果图表模拟运算表具有外部边框，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.HasBorderOutline**

*express*   一个代表 **DataTable** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例使嵌入图表模拟运算表显示时带有外边框，但无单元格边框。*/ function test(){ let myChart = Application.Worksheets.Item(1).ChartObjects(1).Chart     myChart.HasDataTable = true     let myDataTable = myChart.DataTable         myDataTable.HasBorderHorizontal = false         myDataTable.HasBorderVertical = false         myDataTable.HasBorderOutline = true }` |

#### **DataTable.HasBorderVertical**

如果图表模拟运算表具有垂直单元格边框，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.HasBorderVertical**

*express*   一个代表 **DataTable** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例使嵌入图表模拟运算表显示时带有外边框，但无单元格边框。*/ function test(){ let myChart = Application.Worksheets.Item(1).ChartObjects(1).Chart     myChart.HasDataTable = true     let myDataTable = myChart.DataTable         myDataTable.HasBorderHorizontal = false         myDataTable.HasBorderVertical = false         myDataTable.HasBorderOutline = true }` |

#### **DataTable.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **DataTable** 对象的变量。

#### **DataTable.ShowLegendKey**

如果数据标签图例项标示可见，则为 **True**。**Boolean** 类型，可读写。

**语法**

**express.ShowLegendKey**

*express*   一个代表 **DataTable** 对象的变量。

适用环境：web

适用平台：windows/linux