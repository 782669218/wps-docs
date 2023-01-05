**ChartGroup**



代表图表中用同一格式绘制的一个或多个数据系列。

**说明**

一张图表包含一个或多个图表组，每个图表组包含一个或多个 **Series** 对象，每个数据系列包含一个或多个 **Points** 对象。例如，单张图表可能既包含折线图图表组（其中包含所有用折线图格式绘制的数据系列），也包含条形图图表组（其中包含所有用条形图格式绘制的数据系列）。**ChartGroup** 对象是 **ChartGroups** 集合的成员。

使用 **ChartGroups**(*index*)（其中 *index* 是图表组的索引号）可以返回单个 **ChartGroup** 对象。

因为当特定图表组所用的图表格式更改时，该图表组的索引号可能会更改，所以使用命名图表组快捷方法之一来返回特定的图表组会更加容易。**PieGroups** 方法返回图表中饼图图表组的集合，**LineGroups** 方法返回图表中折线图图表组的集合，依此类推。这些方法中的每一个都可以与索引号配合使用以返回单个 **ChartGroup** 对象，或不指定索引号而返回 **ChartGroups** 集合。

**示例**

以下示例向图表工作表 1 上的图表组 1 中添加垂直线。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.Charts.Item(1).ChartGroups.Item(1).HasDropLines = true` |

如果图表已被激活，就可使用 **ActiveChart** 属性。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.Charts.Item(1).Activate() Application.ActiveChart.ChartGroups(1).HasDropLines = true` |

**方法**

|                                                              | 名称                 | 说明                                                         |
| ------------------------------------------------------------ | -------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SeriesCollection** | 返回一个对象，它代表图表或图表组中的一个系列（**Series** 对象）或所有系列的集合（**SeriesCollection** 集合）。 |

**属性**

|                                                              | 名称                    | 说明                                                         |
| ------------------------------------------------------------ | ----------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**         | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AxisGroup**           | 返回或设置指定图表的组。可读/写。                            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BubbleScale**         | 返回或设置指定图表组中气泡的缩放比例。可为从 0（零）到 300 之间的整数，表示相对于默认大小的百分率。仅适用于气泡图。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**             | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DoughnutHoleSize**    | 返回或设置圆环图图表组的内径大小。内径大小以图表大小的百分比表示，有效取值范围为 10% 到 90%。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DownBars**            | 返回一个 **DownBars** 对象，该对象表示折线图中的跌柱线。仅应用于折线图。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DropLines**           | 返回一个 **DropLines** 对象，该对象表示折线图或面积图上数据系列的垂直线。仅适用于折线图或面积图。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FirstSliceAngle**     | 以度为单位返回或设置饼图或圆环图的第一扇区的角度（从垂直方向顺时针计算）。仅应用于饼图、三维饼图和圆环图。值的范围从 0 到 360。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **GapWidth**            | 对于条形图和柱形图：以条形或柱形宽度百分数的形式返回或设置条形簇或柱形簇之间的距离。对于饼图中的扇形和饼图中的条形：返回或设置图表的第一节和第二节之间的间距。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Has3DShading**        | 如果图表组具有三维阴影，则该值为 **True**。本属性仅适合曲面图，如果要将其设置给非曲面图，则将返回运行时错误。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HasDropLines**        | 如果折线图或者面积图中有垂直线，则该值为 **True**。仅应用于折线图和面积图。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HasHiLoLines**        | 如果折线图有高低点连线，则为 **True**。仅应用于折线图。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Index**               | 返回 **Long** 值，它代表对象在其同类对象所组成的集合内的索引号。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Overlap**             | 指定放置条形和柱形的方式。取值范围在 -100 到 100 之间。仅应用于二维条形图和二维柱形图。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**              | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RadarAxisLabels**     | 返回一个 **TickLabels** 对象，该对象表示指定图表组的雷达图坐标轴标签。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SecondPlotSize**      | 以主饼图大小的百分比形式返回或设置复合饼图或复合条饼图中第二部分的大小。可为 5 到 200 之间的值。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SeriesLines**         |                                                              |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowNegativeBubbles** | 如果在图表组中显示表示负值的气泡，则该值为 **True**。仅对气泡图有效。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SizeRepresents**      | 返回或设置气泡图中气泡的大小所表示的含义。可为以下 **XlSizeRepresents** 常量之一：**xlSizeIsArea** 或 **xlSizeIsWidth**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SplitType**           | 返回或设置在复合饼图或复合条饼图中拆分两部分的方式。**XlChartSplitType** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SplitValue**          | 返回或设置复合饼图或复合条饼图中分隔两部分的阈值。可读/写 **Variant** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **UpBars**              | 返回 **UpBars** 对象，该对象表示折线图上的涨柱线。仅应用于折线图。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **VaryByCategories**    | 如果 ET 对每个数据标志指定不同的颜色或模式，则该值为 **True**。图表中必须只包含一个数据系列。**Boolean** 类型，可读写。 |

**成员方法**

#### **ChartGroup.SeriesCollection**

返回一个对象，它代表图表或图表组中的一个系列（**Series** 对象）或所有系列的集合（**SeriesCollection** 集合）。

**语法**

**express.SeriesCollection(Index)**

*express*   一个代表 **ChartGroup** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**               |
| -------- | ------------- | ------------ | ---------------------- |
| *Index*  | 可选          | **Variant**  | 数据系列的名称或编号。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例打开 Chart1 中第一个数据系列的数据标签。*/ Application.Charts.Item("Chart1").SeriesCollection(1).HasDataLabels = true` |

**成员属性**

#### **ChartGroup.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **ChartGroup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() {     if(ActiveWorkbook.Application.Value == "ET"){         alert("This is an ET Application object.")     }     else{         alert("This is not an ET Application object.")     } }` |

#### **ChartGroup.AxisGroup**

返回或设置指定图表的组。可读/写。

**语法**

**express.AxisGroup**

*express*   一个代表 **ChartGroup** 对象的变量。

**说明**

对于三维图表，仅 **xlPrimary** 有效。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例在值轴处于辅助组时将其删除。*/ function test() {     if(myChart.Axes.Item(xlValue).AxisGroup == xlSecondary){         myChart.Axes.Item(xlValue).Delete()     } }` |

#### **ChartGroup.BubbleScale**

返回或设置指定图表组中气泡的缩放比例。可为从 0（零）到 300 之间的整数，表示相对于默认大小的百分率。仅适用于气泡图。**Long** 类型，可读写。

**语法**

**express.BubbleScale**

*express*   一个代表 **ChartGroup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将第一个图表组中的气泡大小设置为默认大小的 200%。*/ function test() {     let chart = Application.Worksheets.Item(1).ChartObjects(1).Chart     chart.ChartGroups(1).BubbleScale = 200 }` |

#### **ChartGroup.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **ChartGroup** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **ChartGroup.DoughnutHoleSize**

返回或设置圆环图图表组的内径大小。内径大小以图表大小的百分比表示，有效取值范围为 10% 到 90%。**Long** 类型，可读写。

**语法**

**express.DoughnutHoleSize**

*express*   一个代表 **ChartGroup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例对 Chart1 中第一个圆环图组的内径大小进行设置。本示例应在二维圆环图上运行。*/ Application.Charts.Item("Chart1").DoughnutGroups(1).DoughnutHoleSize = 10` |

#### **ChartGroup.DownBars**

返回一个 **DownBars** 对象，该对象表示折线图中的跌柱线。仅应用于折线图。只读。

**语法**

**express.DownBars**

*express*   一个代表 **ChartGroup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示 Chart1 的涨跌柱线，并对其颜色进行设置。本示例应在包含两组彼此间有一个或多个相交数据点的数据系列的二维折线图上运行。*/ function test() {     let chartgroup = Application.Charts.Item("Chart1").ChartGroups(1)     chartgroup.HasUpDownBars = true     chartgroup.DownBars.Interior.ColorIndex = 3     chartgroup.UpBars.Interior.ColorIndex = 5 }` |

#### **ChartGroup.DropLines**

返回一个 **DropLines** 对象，该对象表示折线图或面积图上数据系列的垂直线。仅适用于折线图或面积图。只读。

**语法**

**express.DropLines**

*express*   一个代表 **ChartGroup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示 Chart1 中第一个图表组的垂直线，并对其线型、粗细和颜色进行设置。本示例应在包含一个数据系列的二维折线图上运行。*/ function test() {     let chartgroup = Application.Charts.Item("Chart1").ChartGroups(1)     chartgroup.HasDropLines = true     let border = chartgroup.DropLines.Border         border.LineStyle = xlThin         border.Weight = xlMedium         border.ColorIndex = 3 }` |

#### **ChartGroup.FirstSliceAngle**

以度为单位返回或设置饼图或圆环图的第一扇区的角度（从垂直方向顺时针计算）。仅应用于饼图、三维饼图和圆环图。值的范围从 0 到 360。**Long** 类型，可读写。

**语法**

**express.FirstSliceAngle**

*express*   一个代表 **ChartGroup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 Chart1 中的第一个图表组中的第一个扇区的角度。本示例应在二维饼图上运行。*/ Application.Charts.Item("Chart1").ChartGroups(1).FirstSliceAngle = 15` |

#### **ChartGroup.GapWidth**

对于条形图和柱形图：以条形或柱形宽度百分数的形式返回或设置条形簇或柱形簇之间的距离。对于饼图中的扇形和饼图中的条形：返回或设置图表的第一节和第二节之间的间距。**Long** 类型，可读写。

**语法**

**express.GapWidth**

*express*   一个代表 **ChartGroup** 对象的变量。

**说明**

该属性的取值必须介于 0 到 500 之间。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 Chart1 的列簇间距设置为列宽的 50%。*/ Application.Charts.Item("Chart1").ChartGroups(1).GapWidth = 50` |

#### **ChartGroup.Has3DShading**

如果图表组具有三维阴影，则该值为 **True**。本属性仅适合曲面图，如果要将其设置给非曲面图，则将返回运行时错误。**Boolean** 类型，可读写。

**语法**

**express.Has3DShading**

*express*   一个代表 **ChartGroup** 对象的变量。

**说明**

注释：Setting Has3DShading = False has the effect of setting the chart material to "Flat", or removing the effect, whereas setting the Has3DShading = True sets the chart content to the default.

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为图表组添加三维阴影。*/ Application.Charts.Item(1).ChartGroups(1).Has3DShading = true` |

#### **ChartGroup.HasDropLines**

如果折线图或者面积图中有垂直线，则该值为 **True**。仅应用于折线图和面积图。**Boolean** 类型，可读写。

**语法**

**express.HasDropLines**

*express*   一个代表 **ChartGroup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示 Chart1 中第一个图表组的垂直线，并对其线型、粗细和颜色进行设置。本示例应在包含一个数据系列的二维折线图上运行。*/ function test() {     let chartgroup = Application.Charts.Item("Chart1").ChartGroups(1)     chartgroup.HasDropLines = true     let border = chartgroup.DropLines.Border         border.LineStyle = xlThin         border.Weight = xlMedium         border.ColorIndex = 3 }` |

#### **ChartGroup.HasHiLoLines**

如果折线图有高低点连线，则为 **True**。仅应用于折线图。**Boolean** 类型，可读写。

**语法**

**express.HasHiLoLines**

*express*   一个代表 **ChartGroup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示 Chart1 中第一个图表组的高低线，并对其线型、粗细和颜色进行设置。本示例应在具有三组类似股市行情（盘高－盘低－收盘）数据系列的二维折线图上运行。*/ function test() {     let chartgroup = Application.Charts.Item("Chart1").ChartGroups(1)     chartgroup.HasHiLoLines = true     let border = chartgroup.HiLoLines.Border         border.LineStyle = xlThin         border.Weight = xlMedium         border.ColorIndex = 3 }` |

#### **ChartGroup.Index**

返回 **Long** 值，它代表对象在其同类对象所组成的集合内的索引号。

**语法**

**express.Index**

*express*   一个代表 **ChartGroup** 对象的变量。

#### **ChartGroup.Overlap**

指定放置条形和柱形的方式。取值范围在 -100 到 100 之间。仅应用于二维条形图和二维柱形图。**Long** 类型，可读写。

**语法**

**express.Overlap**

*express*   一个代表 **ChartGroup** 对象的变量。

**说明**

如果将该属性设置为 -100，条形之间有一个条形的间距。如果重叠比例为 0（零），条形之间没有间距（一个条形紧挨着另一条形）。如果重叠比例为 100，所有条形将完全重叠。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将第一个图表组的重叠比例设置为 -50。本示例应在包含两个或更多系列的二维柱形图上运行。*/ Application.Charts.Item("Chart1").ChartGroups(1).Overlap = -50` |

#### **ChartGroup.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **ChartGroup** 对象的变量。

#### **ChartGroup.RadarAxisLabels**

返回一个 **TickLabels** 对象，该对象表示指定图表组的雷达图坐标轴标签。只读。

**语法**

**express.RadarAxisLabels**

*express*   一个代表 **ChartGroup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示 Chart1 中第一个图表组的雷达图坐标轴标注，并对这些标注的颜色进行设置。本示例应在雷达图上运行。*/ function test() {     let chartgroup = Application.Charts.Item("Chart1").ChartGroups(1)     chartgroup.HasRadarAxisLabels = true     chartgroup.RadarAxisLabels.Font.ColorIndex = 3 }` |

#### **ChartGroup.SecondPlotSize**

以主饼图大小的百分比形式返回或设置复合饼图或复合条饼图中第二部分的大小。可为 5 到 200 之间的值。**Long** 类型，可读写。

**语法**

**express.SecondPlotSize**

*express*   一个代表 **ChartGroup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例必须在复合饼图或复合条饼图中运行。本示例以数值拆分图表的两部分，在主饼图部分中合并所有小于 10 的数值，并在第二部分中显示这些数值。第二部分的大小为主条形图大小的百分之五十。*/ function test() {     let chartobjects = Application.Worksheets.Item(1).ChartObjects(1).Chart.ChartGroups(1)     chartobjects.SplitType = xlSplitByValue     chartobjects.SplitValue = 10     chartobjects.VaryByCategories = true     chartobjects.SecondPlotSize = 50 }` |

#### **ChartGroup.SeriesLines**

**语法**

**express.SeriesLines**

*express*   一个代表 **ChartGroup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示 Chart1 中图表组一的系列线，并对其线型、粗细和颜色进行设置。本示例应在有两个或多个数据系列的二维堆积柱形图上运行。*/ function test() {     let myChartGroups = Application.Charts.Item("Chart1").ChartGroups(1)     myChartGroups.HasSeriesLines = true     let myBorder = myChartGroups.SeriesLines.Border         myBorder.LineStyle = xlThin         myBorder.Weight = xlMedium         myBorder.ColorIndex = 3 }` |

#### **ChartGroup.ShowNegativeBubbles**

如果在图表组中显示表示负值的气泡，则该值为 **True**。仅对气泡图有效。**Boolean** 类型，可读写。

**语法**

**express.ShowNegativeBubbles**

*express*   一个代表 **ChartGroup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.Worksheets.Item(1).ChartObjects(1).Chart.ChartGroups(1).ShowNegativeBubbles = true` |

#### **ChartGroup.SizeRepresents**

返回或设置气泡图中气泡的大小所表示的含义。可为以下 **XlSizeRepresents** 常量之一：**xlSizeIsArea** 或 **xlSizeIsWidth**。**Long** 类型，可读写。

**语法**

**express.SizeRepresents**

*express*   一个代表 **ChartGroup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置图表组一中气泡的大小所表示的含义。*/ Application.Charts.Item(1).ChartGroups(1).SizeRepresents = xlSizeIsWidth` |

#### **ChartGroup.SplitType**

返回或设置在复合饼图或复合条饼图中拆分两部分的方式。**XlChartSplitType** 类型，可读写。

**语法**

**express.SplitType**

*express*   一个代表 **ChartGroup** 对象的变量。

**说明**

| **XlChartSplitType** 可为以下 **XlChartSplitType** 常量之一。 |
| ------------------------------------------------------------ |
| **xlSplitByCustomSplit**                                     |
| **xlSplitByPercentValue**                                    |
| **xlSplitByPosition**                                        |
| **xlSplitByValue**                                           |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例必须在复合饼图或复合条饼图中运行。本示例以数值拆分图表的两部分，在主饼图部分中合并所有小于 10 的数值，并在第二部分中显示这些数值。*/ function test() {     let myChart = Application.Worksheets.Item(1).ChartObjects(1).Chart.ChartGroups(1)     myChart.SplitType = xlSplitByValue     myChart.SplitValue = 10     myChart.VaryByCategories = true }` |

#### **ChartGroup.SplitValue**

返回或设置复合饼图或复合条饼图中分隔两部分的阈值。可读/写 **Variant** 类型。

**语法**

**express.SplitValue**

*express*   一个代表 **ChartGroup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例必须在复合饼图或复合条饼图中运行。本示例以数值拆分图表的两部分，在主饼图部分中合并所有小于 10 的数值，并在第二部分中显示这些数值。*/ function test() {     let myChart = Application.Worksheets.Item(1).ChartObjects(1).Chart.ChartGroups(1)     myChart.SplitType = xlSplitByValue     myChart.SplitValue = 10     myChart.VaryByCategories = true }` |

#### **ChartGroup.UpBars**

返回 **UpBars** 对象，该对象表示折线图上的涨柱线。仅应用于折线图。只读。

**语法**

**express.UpBars**

*express*   一个代表 **ChartGroup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示 Chart1 中图表组一的涨跌柱线，并对其颜色进行设置。本示例应在二维折线图上运行，该折线图应包含两组彼此间有一个或多个相交数据点的系列。*/ function test() {     let myChart = Application.Charts.Item("Chart1").ChartGroups(1)     myChart.HasUpDownBars = True     myChart.DownBars.Interior.ColorIndex = 3     myChart.UpBars.Interior.ColorIndex = 5 }` |

#### **ChartGroup.VaryByCategories**

如果 ET 对每个数据标志指定不同的颜色或模式，则该值为 **True**。图表中必须只包含一个数据系列。**Boolean** 类型，可读写。

**语法**

**express.VaryByCategories**

*express*   一个代表 **ChartGroup** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例对第一个图表组中的每个数据标志指定一种不同的颜色或图案。本示例应在数据系列上有数据标志的二维折线图上运行。*/ Application.Charts.Item("Chart1").ChartGroups(1).VaryByCategories = true` |

适用环境：web

适用平台：windows/linux