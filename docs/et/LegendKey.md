**LegendKey**



代表图表图例中的图例标示。

**说明**

每个图例项标示都是一个图形，它将图例项标示以及与之相关联的图表中的系列或趋势线链接起来。图例项标示链接到与之相关联的系列或趋势线的方式为：更改一个的格式会同时更改另一个的格式。

使用 **LegendKey** 属性可返回 **LegendKey** 对象。下例更改名为“Sheet1”的工作表上嵌入式图表一的图例顶部的图例项标示的标志背景色。这样会同时更改与该图例项相关的系列中每一个数据点的格式。相关的系列必须支持数据标志。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.Worksheets.Item("sheet1").ChartObjects(1).Chart.Legend.LegendEntries(1).LegendKey.MarkerBackgroundColorIndex = 5 ` |

**方法**

|                                                              | 名称             | 说明                 |
| ------------------------------------------------------------ | ---------------- | -------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ClearFormats** | 清除对象的格式设置。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**       | 删除对象。           |

**属性**

|                                                              | 名称                           | 说明                                                         |
| ------------------------------------------------------------ | ------------------------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**                | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**                    | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Format**                     | 返回 **ChartFormat** 对象。只读。                            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Height**                     | 返回一个 **Double** 值，它代表对象的高度（以磅为单位）。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **InvertIfNegative**           | 如果指定项与一个负数相对应时 ET 就将其反色，则为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Left**                       | 返回一个 **Double** 值，它代表从对象左边缘到图表区左边缘的距离（以磅为单位）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MarkerBackgroundColor**      | 将数据标志的背景色设置为 RGB 值或返回对应的颜色索引值。仅适用于折线图、散点图和雷达图。**Long** 型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MarkerBackgroundColorIndex** | 返回或设置数据标志的背景色，表示为当前调色板中的索引或下列 **XlColorIndex** 常量之一：**xlColorIndexAutomatic** 或 **xlColorIndexNone**。仅适用于折线图、散点图和雷达图。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MarkerForegroundColor**      | 将数据标志的背景色设置为 RGB 值或返回对应的颜色索引值。仅适用于折线图、散点图和雷达图。**Long** 型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MarkerForegroundColorIndex** | 返回或设置数据标志的前景色，表示为当前调色板中的索引或下列 **XlColorIndex** 常量之一：**xlColorIndexAutomatic** 或 **xlColorIndexNone**。仅适用于折线图、散点图和雷达图。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MarkerSize**                 | 返回或设置数据标志的大小，以磅?（磅：指打印的字符的高度的度量单位。1 磅等于 1/72 英寸，或大约等于 1 厘米的 1/28。）为单位。可以是 2 到 72 之间的一个值。**Long** 型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MarkerStyle**                | 返回或设置折线图、散点图或雷达图中数据点或数据系列的数据标志样式。**XlMarkerStyle** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**                     | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PictureType**                | 返回或设置一个 **XlChartPictureType** 值，它代表图片在图例项标示上的显示方式。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PictureUnit2**               | 如果 **PictureType** 属性设置为 **xlStackScale**，则返回或设置图表上每一个图片的单位。否则忽略此属性。可读/写 **Double** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Shadow**                     | 返回或设置一个 **Boolean** 值，它确定对象是否有阴影。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Smooth**                     | 如果图例项标示有平滑线，则为 **True**。可读写。              |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Top**                        | 返回一个 **Double** 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Width**                      | 返回一个 **Double** 值，它代表对象的宽度（以磅为单位）。     |

**成员方法**

#### **LegendKey.ClearFormats**

清除对象的格式设置。

**语法**

**express.ClearFormats()**

*express*   一个代表 **LegendKey** 对象的变量。

**返回值**

Variant

#### **LegendKey.Delete**

删除对象。

**语法**

**express.Delete()**

*express*   一个代表 **LegendKey** 对象的变量。

**返回值**

Variant

**说明**

删除 **LegendKey** 对象将删除整个数据系列。

**成员属性**

#### **LegendKey.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **LegendKey** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() { 	let myObject = Application.ActiveWorkbook 	if (myObject.Application.Value == "ET") { 		alert("This is an ET Application object.") 	} else { 		alert("This is not an ET Application object.") 	} }` |

#### **LegendKey.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **LegendKey** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **LegendKey.Format**

返回 **ChartFormat** 对象。只读。

**语法**

**express.Format**

*express*   一个代表 **LegendKey** 对象的变量。

#### **LegendKey.Height**

返回一个 **Double** 值，它代表对象的高度（以磅为单位）。

**语法**

**express.Height**

*express*   一个代表 **LegendKey** 对象的变量。

#### **LegendKey.InvertIfNegative**

如果指定项与一个负数相对应时 ET 就将其反色，则为 **True**。**Boolean** 类型，可读写。

**语法**

**express.InvertIfNegative**

*express*   一个代表 **LegendKey** 对象的变量。

#### **LegendKey.Left**

返回一个 **Double** 值，它代表从对象左边缘到图表区左边缘的距离（以磅为单位）。

**语法**

**express.Left**

*express*   一个代表 **LegendKey** 对象的变量。

#### **LegendKey.MarkerBackgroundColor**

将数据标志的背景色设置为 RGB 值或返回对应的颜色索引值。仅适用于折线图、散点图和雷达图。**Long** 型，可读写。

**语法**

**express.MarkerBackgroundColor**

*express*   一个代表 **LegendKey** 对象的变量。

#### **LegendKey.MarkerBackgroundColorIndex**

返回或设置数据标志的背景色，表示为当前调色板中的索引或下列 **XlColorIndex** 常量之一：**xlColorIndexAutomatic** 或 **xlColorIndexNone**。仅适用于折线图、散点图和雷达图。**Long** 类型，可读写。

**语法**

**express.MarkerBackgroundColorIndex**

*express*   一个代表 **LegendKey** 对象的变量。

#### **LegendKey.MarkerForegroundColor**

将数据标志的背景色设置为 RGB 值或返回对应的颜色索引值。仅适用于折线图、散点图和雷达图。**Long** 型，可读写。

**语法**

**express.MarkerForegroundColor**

*express*   一个代表 **LegendKey** 对象的变量。

#### **LegendKey.MarkerForegroundColorIndex**

返回或设置数据标志的前景色，表示为当前调色板中的索引或下列 **XlColorIndex** 常量之一：**xlColorIndexAutomatic** 或 **xlColorIndexNone**。仅适用于折线图、散点图和雷达图。**Long** 类型，可读写。

**语法**

**express.MarkerForegroundColorIndex**

*express*   一个代表 **LegendKey** 对象的变量。

#### **LegendKey.MarkerSize**

返回或设置数据标志的大小，以磅?（磅：指打印的字符的高度的度量单位。1 磅等于 1/72 英寸，或大约等于 1 厘米的 1/28。）为单位。可以是 2 到 72 之间的一个值。**Long** 型，可读写。

**语法**

**express.MarkerSize**

*express*   一个代表 **LegendKey** 对象的变量。

#### **LegendKey.MarkerStyle**

返回或设置折线图、散点图或雷达图中数据点或数据系列的数据标志样式。**XlMarkerStyle** 类型，可读写。

**语法**

**express.MarkerStyle**

*express*   一个代表 **LegendKey** 对象的变量。

**说明**

| **XlMarkerStyle** 可为以下 **XlMarkerStyle** 常量之一。 |
| ------------------------------------------------------- |
| **xlMarkerStyleAutomatic**。自动设置标记                |
| **xlMarkerStyleCircle**。圆形标记                       |
| **xlMarkerStyleDash**。长条形标记                       |
| **xlMarkerStyleDiamond**。菱形标记                      |
| **xlMarkerStyleDot**。短条形标记                        |
| **xlMarkerStyleNone**。无标记                           |
| **xlMarkerStylePicture**。图片标记                      |
| **xlMarkerStylePlus**。带加号的方形标记                 |
| **xlMarkerStyleSquare**。方形标记                       |
| **xlMarkerStyleStar**。带星号的方形标记                 |
| **xlMarkerStyleTriangle**。三角形标记                   |
| **xlMarkerStyleX**。带 X 记号的方形标记                 |

#### **LegendKey.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **LegendKey** 对象的变量。

#### **LegendKey.PictureType**

返回或设置一个 **XlChartPictureType** 值，它代表图片在图例项标示上的显示方式。

**语法**

**express.PictureType**

*express*   一个代表 **LegendKey** 对象的变量。

#### **LegendKey.PictureUnit2**

如果 **PictureType** 属性设置为 **xlStackScale**，则返回或设置图表上每一个图片的单位。否则忽略此属性。可读/写 **Double** 类型。

**语法**

**express.PictureUnit2**

*express*   一个代表 **LegendKey** 对象的变量。

#### **LegendKey.Shadow**

返回或设置一个 **Boolean** 值，它确定对象是否有阴影。

**语法**

**express.Shadow**

*express*   一个代表 **LegendKey** 对象的变量。

#### **LegendKey.Smooth**

如果图例项标示有平滑线，则为 **True**。可读写。

**语法**

**express.Smooth**

*express*   一个代表 **LegendKey** 对象的变量。

#### **LegendKey.Top**

返回一个 **Double** 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。

**语法**

**express.Top**

*express*   一个代表 **LegendKey** 对象的变量。

#### **LegendKey.Width**

返回一个 **Double** 值，它代表对象的宽度（以磅为单位）。

**语法**

**express.Width**

*express*   一个代表 **LegendKey** 对象的变量。

适用环境：web

适用平台：windows/linux