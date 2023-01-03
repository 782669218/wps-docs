**Border**



代表对象的边框。

**说明**

大多数具有边框的对象（除 **Range**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)和 **Style**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Style/Style%20.htm#jsObject_Style)对象外）都将边框作为单一实体处理，而不管边框有几个边。整个边框必须作为一个整体单位返回。例如，使用 **TrendLine**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/XlTrendlineType%20%E6%9E%9A%E4%B8%BE.html)对象的 **Border** 属性可返回此类对象的 **Border** 对象。

下例更改活动图表中趋势线的类型和线型。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {      let trendlines = Application.ActiveChart.SeriesCollection(1).Trendlines(1)      trendlines.Type = xlLinear  trendlines.Border.LineStyle = xlDash } ` |

[**Range** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)和 [**Style** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Style/Style%20.htm#jsObject_Style)对象具有四个分立的边框：左边框、右边框、顶部边框和底部边框，这四个边框既可单独返回，也可作为一个组同时返回。使用 **Borders** 属性可返回 **Borders** 集合，该集合包含所有四个边框，并将这些边框视为一个单位。下例向第一张工作表上的单元格 A1 添加双边框。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.Worksheets.Item(1).Range("A1").Borders.LineStyle = xlDouble` |

使用 **Borders**(*index*)（其中 *index* 指定边框）可返回单个 **Border** 对象。下例设置单元格区域 A1:G1 的底部边框的颜色。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.Worksheets.Item("Sheet1").Range("A1:G1").Borders(xlEdgeBottom).Color = (255, 0, 0)` |

*Index* 可为以下 [XlBordersIndex ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/XlBordersIndex%20%E6%9E%9A%E4%B8%BE.html)常量之一：**xlDiagonalDown**、**xlDiagonalUp**、**xlEdgeBottom**、**xlEdgeLeft**、**xlEdgeRight**、**xlEdgeTop**、**xlInsideHorizontal** 或 **xlInsideVertical**。

**属性**

|                                                              | 名称             | 说明                                                         |
| ------------------------------------------------------------ | ---------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**  | 如果不使用对象识别符，则该属性返回一个 [Application ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Color**        | 返回或设置对象的主要颜色，如注释部分中的表格所示。**Variant** 型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ColorIndex**   | 返回或设置一个 **Variant** 值，它代表边框的颜色。            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**      | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LineStyle**    | 返回或设置边框的线型。**XlLineStyle**、**xlGray25**、**xlGray50**、**xlGray75** 或 **xlAutomatic** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**       | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ThemeColor**   | 返回或设置已应用的配色方案中的主题颜色，该配色方案与指定对象相关联。可读/写 **Variant** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TintAndShade** | 返回或设置一个 **Single**，使颜色变深或变浅。                |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Weight**       | 返回或设置一个 **XlBorderWeight**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/XlBorderWeight%20%E6%9E%9A%E4%B8%BE.html)值，它代表边框的粗细。 |

**成员属性**

#### **Border.Application**

如果不使用对象识别符，则该属性返回一个 [Application ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **Border** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例显示一条有关创建 myObject 的应用程序的消息*/ function test() {     let myObject = Application.ActiveWorkbook     if(myObject.Application.Value == "ET") {         alert("This is an ET Application object.")     }     else {         alert("This is not an ET Application object.")     } }` |

#### **Border.Color**

返回或设置对象的主要颜色，如注释部分中的表格所示。**Variant** 型，可读写。

**语法**

**express.Color**

*express*   一个代表 **Border** 对象的变量。

**说明**

| 对象         | 对应颜色                                                     |
| ------------ | ------------------------------------------------------------ |
| **边框**     | 边框的颜色。                                                 |
| **Borders**  | 一个区域的所有四条边的颜色。如果四边不是同一种颜色，则 **Color** 返回的是 0（零）。 |
| **Font**     | 字体的颜色。                                                 |
| **Interior** | 单元格底纹的颜色或图形对象的填充颜色。                       |
| **Tab**      | 选项卡的颜色。                                               |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 此示例对 Chart1 中数值坐标轴的刻度线标志颜色进行设置*/ Application.Charts.Item("Chart1").Axes(xlValue).TickLabels.Font.Color = 0x0000ff00` |

#### **Border.ColorIndex**

返回或设置一个 **Variant** 值，它代表边框的颜色。

**语法**

**express.ColorIndex**

*express*   一个代表 **Border** 对象的变量。

**说明**

颜色可指定为当前调色板中颜色的索引值，也可指定为下列 [XlColorIndex ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/XlColorIndex%20%E6%9E%9A%E4%B8%BE.html)常量之一：

- **xlColorIndexAutomatic**
- **xlColorIndexNone**

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例为 Chart1 中数值坐标轴的主要网格线设置了颜色*/ function test() {   let axes = Application.Charts.Item("Chart1").Axes(xlValue)   if(axes.HasMajorGridlines){ 	axes.MajorGridlines.Border.ColorIndex = 5    //set color to blue   } }` |

#### **Border.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Border** 对象的变量。

#### **Border.LineStyle**

返回或设置边框的线型。**XlLineStyle**、**xlGray25**、**xlGray50**、**xlGray75** 或 **xlAutomatic** 类型，可读写。

**语法**

**express.LineStyle**

*express*   一个代表 **Border** 对象的变量。

**说明**

**xlDouble** 和 **xlSlantDashDot** 不适用于图表。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例为 Chart1 的图表区和绘图区域设置边框*/ function test() { 	let charts = Application.Charts.Item("Chart1") 	charts.ChartArea.Border.LineStyle = xlDashDot 	let border = charts.PlotArea.Border 	border.LineStyle = xlDashDotDot 	border.Weight = xlThick }` |

#### **Border.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **Border** 对象的变量。

#### **Border.ThemeColor**

返回或设置已应用的配色方案中的主题颜色，该配色方案与指定对象相关联。可读/写 **Variant** 类型。

**语法**

**express.ThemeColor**

*express*   一个代表 **Border** 对象的变量。

**说明**

如果对象当前未应用主题颜色，试图访问其主题颜色则会引起无效请求运行时错误。

#### **Border.TintAndShade**

返回或设置一个 **Single**，使颜色变深或变浅。

**语法**

**express.TintAndShade**

*express*   一个代表 **Border** 对象的变量。

**说明**

可以为 **TintAndShade** 属性输入 -1（最暗）到 1（最亮）之间的数字，零 (0) 为中间值。

如果将此属性设置为小于 -1 或大于 1 的值，则会引起运行时错误：“指定的值超出了范围”。此属性用于主题颜色和非主题颜色。

#### **Border.Weight**

返回或设置一个 **XlBorderWeight**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/XlBorderWeight%20%E6%9E%9A%E4%B8%BE.html)值，它代表边框的粗细。

**语法**

**express.Weight**

*express*   一个代表 **Border** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例对 Sheet1 上椭圆一的边框粗细进行设置*/ Application.Worksheets.Item("Sheet1").Ovals(1).Border.Weight = xlMedium` |

适用环境：web

适用平台：windows/linux