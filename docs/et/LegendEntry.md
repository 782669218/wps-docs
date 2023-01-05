**LegendEntry**



代表图表图例中的图例项。

**说明**

**LegendEntry** 对象是 **LegendEntries** 集合的成员。**LegendEntries** 集合包含图例中所有的 **LegendEntry** 对象。

每个图例项都有两部分：一部分是该项的文本，它是与该图例项相关联的数据系列或趋势线的名称；另一部分是项标志，它在图表中以直观的方式将图例项以及与之相关联的数据系列或趋势线链接起来。项标志及其相关数据系列或趋势线的格式设置属性包含在 **LegendKey** 对象中。

不能修改图例项的文字。**LegendEntry** 对象支持字体格式，且可被删除。图例项不支持图案格式。图例项的位置和尺寸是固定的。

没有返回相应于某图例项的数据系列或趋势线的直接方法。

在删除了图例项之后，唯一的恢复方法是：通过将图表的 **HasLegend** 属性设置为 **False**，然后再设回 **True**，从而删除并重新创建包含这些图例项的图例。

使用 **LegendEntries**( *index*)（其中 *index* 是图例项索引号）可返回一个 **LegendEntry** 对象。不能按名称返回图例项。

图例项的编号代表图例项在图例中的位置。`LegendEntries(1)` 位于图例的顶部； `LegendEntries(LegendEntries.Count)` 位于图例的底部。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例更改工作表“Sheet1”上嵌入的第一个图表中图例顶部的图例项（这通常是第一个数据系列的图例）的文字字体。*/ Application.Worksheets.Item("sheet1").ChartObjects(1).Chart.Legend.LegendEntries(1).Font.Italic = true` |

**方法**

|                                                              | 名称       | 说明       |
| ------------------------------------------------------------ | ---------- | ---------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete** | 删除对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Select** | 选择对象。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Font**        | 返回一个 **Font** 对象，它代表指定对象的字体。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Format**      | 返回 **ChartFormat** 对象。只读。                            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Height**      | 返回一个 **Double** 值，它代表对象的高度（以磅为单位）。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Index**       | 返回 **Long** 值，它代表对象在其同类对象所组成的集合内的索引号。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Left**        | 返回一个 **Double** 值，它代表从对象左边缘到图表区左边缘的距离（以磅为单位）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LegendKey**   | 返回一个 **LegendKey** 对象，该对象表示与图例项相关的图例标示。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Top**         | 返回一个 **Double** 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Width**       | 返回一个 **Double** 值，它代表对象的宽度（以磅为单位）。     |

**成员方法**

#### **LegendEntry.Delete**

删除对象。

**语法**

**express.Delete()**

*express*   一个代表 **LegendEntry** 对象的变量。

**返回值**

Variant

#### **LegendEntry.Select**

选择对象。

**语法**

**express.Select()**

*express*   一个代表 **LegendEntry** 对象的变量。

**返回值**

Variant

**成员属性**

#### **LegendEntry.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **LegendEntry** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() { 	let myObject = Application.ActiveWorkbook 	if (myObject.Application.Value == "ET") { 		alert("This is an ET Application object.") 	} else { 		alert("This is not an ET Application object.") 	} }` |

#### **LegendEntry.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **LegendEntry** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **LegendEntry.Font**

返回一个 **Font** 对象，它代表指定对象的字体。

**语法**

**express.Font**

*express*   一个代表 **LegendEntry** 对象的变量。

#### **LegendEntry.Format**

返回 **ChartFormat** 对象。只读。

**语法**

**express.Format**

*express*   一个代表 **LegendEntry** 对象的变量。

#### **LegendEntry.Height**

返回一个 **Double** 值，它代表对象的高度（以磅为单位）。

**语法**

**express.Height**

*express*   一个代表 **LegendEntry** 对象的变量。

#### **LegendEntry.Index**

返回 **Long** 值，它代表对象在其同类对象所组成的集合内的索引号。

**语法**

**express.Index**

*express*   一个代表 **LegendEntry** 对象的变量。

#### **LegendEntry.Left**

返回一个 **Double** 值，它代表从对象左边缘到图表区左边缘的距离（以磅为单位）。

**语法**

**express.Left**

*express*   一个代表 **LegendEntry** 对象的变量。

#### **LegendEntry.LegendKey**

返回一个 **LegendKey** 对象，该对象表示与图例项相关的图例标示。

**语法**

**express.LegendKey**

*express*   一个代表 **LegendEntry** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 Chart1 上图例项一的图例标示设置为三角形。本示例应在二维折线图上运行。*/ Application.Charts.Item("Chart1").Legend.LegendEntries(1).LegendKey.MarkerStyle = xlMarkerStyleTriangle` |

#### **LegendEntry.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **LegendEntry** 对象的变量。

#### **LegendEntry.Top**

返回一个 **Double** 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。

**语法**

**express.Top**

*express*   一个代表 **LegendEntry** 对象的变量。

#### **LegendEntry.Width**

返回一个 **Double** 值，它代表对象的宽度（以磅为单位）。

**语法**

**express.Width**

*express*   一个代表 **LegendEntry** 对象的变量。

适用环境：web

适用平台：windows/linux