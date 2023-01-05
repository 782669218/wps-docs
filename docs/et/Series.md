**Series**



代表图表上的系列。

**说明**

**Series** 对象是 **SeriesCollection** 集合的成员。

使用 **SeriesCollection**(*index*)（其中 *index* 是系列索引号或名称）可返回一个 **Series** 对象。下例设置 Sheet1 上嵌入式图表一中第一个系列的内部颜色。

系列索引号指明了系列添加到图表中的顺序。`SeriesCollection(1)` 是第一个添加到图表中的系列，而 `SeriesCollection(SeriesCollection.Count)` 是最后一个添加到图表中的系列。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `Application.Worksheets.Item("sheet1").ChartObjects(1).Chart.SeriesCollection(1).Interior.Color = (255, 0, 0)` |

**方法**

|                                                              | 名称                | 说明                                                         |
| ------------------------------------------------------------ | ------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ApplyDataLabels** | 向系列应用数据标签。                                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AxisGroup**       | 返回或设置指定系列的组。可读/写                              |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ClearFormats**    | 清除对象的格式设置。                                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Copy**            | 如果系列具有图片填充，此方法将把图片复制到剪贴板。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **DataLabels**      | 返回代表数据系列中单个数据标签（**DataLabel** 对象）或所有数据标签的集合（**DataLabels** 集合）的对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**          | 删除对象。                                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ErrorBar**        | 将误差线应用于系列。**Variant** 类型。                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Paste**           | 从剪贴板中粘贴图片，将其作为所选系列的标志。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Points**          | 返回一个对象，该对象表示数据系列中单个数据点（**Point** 对象）或所有数据点的集合（**Points** 集合）。只读 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Select**          | 选择对象。                                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Trendlines**      | 返回一个对象，该对象表示单个趋势线（**Trendline** 对象）或所有趋势线的集合（**Trendlines** 集合）。 |

**属性**

|                                                              | 名称                           | 说明                                                         |
| ------------------------------------------------------------ | ------------------------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**                | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ApplyPictToEnd**             | 如果图片应用于系列中数据点或所有数据点之后，则为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ApplyPictToFront**           | 如果图片应用于系列中数据点或所有数据点之前，则为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ApplyPictToSides**           | 如果将图片应用于系列中某个数据点或所有数据点的旁边，则为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BarShape**                   | 返回或设置用于三维条形图或柱形图的形状。**XlBarShape** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BubbleSizes**                | 返回或设置一个字符串，该字符串引用包含气泡图 x 值、y 值和大小数据的工作表单元格。在返回单元格引用时，它将返回以 A1 样式注释描述单元格的字符串。要为气泡图设置大小数据，必须使用 R1 样式注释。仅适用于气泡图。**Variant** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ChartType**                  | 返回或设置图表类型。**XlChartType** 类型，可读写。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**                    | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ErrorBars**                  | 返回一个 **ErrorBars** 对象，该对象表示数据系列中的误差线。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Explosion**                  | 返回或设置饼图或圆环图的扇区的分离程度值。如果没有分离，则返回 0（零），即扇区的中心与饼图中心重合。**Long** 型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Format**                     | 返回 **ChartFormat** 对象。只读。                            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Formula**                    | 返回或设置一个 **String** 值，它代表 A1 样式表示法和宏语言中的对象的公式。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FormulaLocal**               | 返回或设置指定对象的公式，使用用户语言 A1 格式引用。**String** 型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FormulaR1C1**                | 返回或设置指定对象的公式，使用宏语言 R1C1 格式符号表示。**String** 型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FormulaR1C1Local**           | 返回或设置指定对象的公式，使用用户语言 R1C1 格式符号表示。**String** 型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Has3DEffect**                | 如果数据系列有三维外观，则为 **True**。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HasDataLabels**              | 如果数据系列具有数据标签，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HasErrorBars**               | 如果序列有错误条时，该值为 **True**。该属性对三维图表无效。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HasLeaderLines**             | 如果数据系列有引导线，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **InvertColor**                | 返回或设置系列中负数据点的填充色。可读/写                    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **InvertColorIndex**           | 返回或设置系列中负数据点的填充色。可读/写                    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **InvertIfNegative**           | 如果指定项与一个负数相对应时 ET 就将其反色，则为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LeaderLines**                | 返回一个 **LeaderLines** 对象，该对象表示系列的引导线。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MarkerBackgroundColor**      | 将数据标志的背景色设置为 RGB 值或返回对应的颜色索引值。仅适用于折线图、散点图和雷达图。**Long** 型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MarkerBackgroundColorIndex** | 返回或设置数据标志的背景色，表示为当前调色板中的索引或下列 **XlColorIndex** 常量之一：**xlColorIndexAutomatic** 或 **xlColorIndexNone**。仅适用于折线图、散点图和雷达图。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MarkerForegroundColor**      | 将数据标志的背景色设置为 RGB 值或返回对应的颜色索引值。仅适用于折线图、散点图和雷达图。**Long** 型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MarkerForegroundColorIndex** | 返回或设置数据标志的前景色，表示为当前调色板中的索引或下列 **XlColorIndex** 常量之一：**xlColorIndexAutomatic** 或 **xlColorIndexNone**。仅适用于折线图、散点图和雷达图。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MarkerSize**                 | 返回或设置数据标志的大小，以磅（磅：指打印的字符的高度的度量单位。1 磅等于 1/72 英寸，或大约等于 1 厘米的 1/28。）为单位。可以是 2 到 72 之间的一个值。**Long** 型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MarkerStyle**                | 返回或设置折线图、散点图或雷达图中数据点或数据系列的数据标志样式。**XlMarkerStyle** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**                       | 返回或设置一个 **String** 值，它代表对象的名称。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**                     | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PictureType**                | 返回或设置一个 **XlChartPictureType** 值，它代表在柱形或条形图上图片的显示方式。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PictureUnit2**               | 如果 **PictureType** 属性设置为 **xlStackScale**，则返回或设置图表上每一个图片的单位。否则忽略此属性。可读/写 **Double** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PlotColorIndex**             | 返回内部用于将系列格式设置与图表元素关联的索引值。只读       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PlotOrder**                  | 返回或设置图表组中选定数据系列的绘制顺序。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Shadow**                     | 返回或设置一个 **Boolean** 值，它确定对象是否有阴影。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Smooth**                     | 如果折线图或散点图有平滑线，则为 **True**。仅适用于折线图和散点图。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**                       | 返回或设置一个代表系列类型的 Long 值。                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Values**                     | 返回或设置一个 **Variant** 值，它代表系列中所有值的集合。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **XValues**                    | 返回或设置图表系列中 x 值的数组。**XValues** 属性可设置为工作表区域或数值数组，但不能为二者的组合。**Variant** 类型，可读写。 |

**成员方法**

#### **Series.ApplyDataLabels**

向系列应用数据标签。

**语法**

**express.ApplyDataLabels(Type, LegendKey, AutoText, HasLeaderLines, ShowSeriesName, ShowCategoryName, ShowValue, ShowPercentage, ShowBubbleSize, Separator)**

*express*   一个代表 **Series** 对象的变量。

**参数**

| **名称**           | **必选/可选** | **数据类型**         | **说明**                                                     |
| ------------------ | ------------- | -------------------- | ------------------------------------------------------------ |
| *Type*             | 可选          | **XlDataLabelsType** | 要应用的数据标签的类型。                                     |
| *LegendKey*        | 可选          | **Variant**          | 如果为 True，则在数据点旁边显示图例项标示。默认值为 False。  |
| *AutoText*         | 可选          | **Variant**          | 如果对象根据内容自动生成相应的文字，则为 True。              |
| *HasLeaderLines*   | 可选          | **Variant**          | 对于 Chart 和 Series 对象，如果数据系列有引导线，则为 True。 |
| *ShowSeriesName*   | 可选          | **Variant**          | 传递布尔值以启用或禁用数据标签的系列名称。                   |
| *ShowCategoryName* | 可选          | **Variant**          | 传递布尔值以启用或禁用数据标签的分类名称。                   |
| *ShowValue*        | 可选          | **Variant**          | 传递布尔值以启用或禁用数据标签的值。                         |
| *ShowPercentage*   | 可选          | **Variant**          | 传递布尔值以启用或禁用数据标签的百分比。                     |
| *ShowBubbleSize*   | 可选          | **Variant**          | 传递布尔值以启用或禁用数据标签的气泡大小。                   |
| *Separator*        | 可选          | **Variant**          | 数据标签的分隔符。                                           |

**说明**

| **XlDataLabelsType** 可为以下 **XlDataLabelsType** 常量之一。 |
| ------------------------------------------------------------ |
| **xlDataLabelsShowBubbleSizes**。数据标签的气泡尺寸。        |
| **xlDataLabelsShowLabelAndPercent**。占总数的百分比及数据点所属的分类。仅用于饼图和圆环图。 |
| **xlDataLabelsShowPercent**。占总数的百分比。仅用于饼图和圆环图。 |
| **xlDataLabelsShowLabel**。数据点所属的分类。                |
| **xlDataLabelsShowNone**。无数据标签。                       |
| **xlDataLabelsShowValue**。默认。数据点的值（若未指定此参数，则假定为此值）。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//此示例对 Chart1 上的第一个数据系列应用分类标签。 Application.Charts.Item("Chart1").SeriesCollection().Item(1).ApplyDataLabels(xlDataLabelsShowLabel)` |

#### **Series.AxisGroup**

返回或设置指定系列的组。可读/写

**语法**

**express.AxisGroup()**

*express*   一个代表 **Series** 对象的变量。

**返回值**

XlAxisGroup

#### **Series.ClearFormats**

清除对象的格式设置。

**语法**

**express.ClearFormats()**

*express*   一个代表 **Series** 对象的变量。

#### **Series.Copy**

如果系列具有图片填充，此方法将把图片复制到剪贴板。

**语法**

**express.Copy()**

*express*   一个代表 **Series** 对象的变量。

#### **Series.DataLabels**

返回代表数据系列中单个数据标签（**DataLabel** 对象）或所有数据标签的集合（**DataLabels** 集合）的对象。

**语法**

**express.DataLabels()**

*express*   一个代表 **Series** 对象的变量。

**说明**

如果指定数据标签的“值”选项处于打开状态，则返回的集合中每个数据点最多包含一个数据标签。可分别设置系列中数据点的数据标签开关选项。

如果指定系列处于面积图中，并且数据标签的“系列名称”选项处于打开状态，则返回的集合中仅包含单个数据标签，即面积系列的标签。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例对 Chart1 的第一个数据系列的数据标签进行设置，显示其关键字段的值。本示例假定在运行时这些值是可见的。 function test() {     let series = Application.Charts.Item("Chart1").SeriesCollection().Item(1)     series.HasDataLabels = true     let bel= series.DataLabels()     bel.ShowLegendKey = true }` |

#### **Series.Delete**

删除对象。

**语法**

**express.Delete()**

*express*   一个代表 **Series** 对象的变量。

#### **Series.ErrorBar**

将误差线应用于系列。**Variant** 类型。

**语法**

**express.ErrorBar(Direction, Include, Type, Amount, MinusValues)**

*express*   一个代表 **Series** 对象的变量。

**参数**

| **名称**      | **必选/可选** | **数据类型**            | **说明**                                                   |
| ------------- | ------------- | ----------------------- | ---------------------------------------------------------- |
| *Direction*   | 必选          | **XlErrorBarDirection** | 误差线方向。                                               |
| *Include*     | 必选          | **XlErrorBarInclude**   | 要包含的误差线部分。                                       |
| *Type*        | 必选          | **XlErrorBarType**      | 误差线类型。                                               |
| *Amount*      | 可选          | **Variant**             | 误差量。当 Type 为 xlErrorBarTypeCustom 时只用于正误差量。 |
| *MinusValues* | 可选          | **Variant**             | 当 Type 为 xlErrorBarTypeCustom 时，为负误差量。           |

**返回值**

Variant

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例显示 Chart1 中的第一个数据系列在 Y 方向上的标准误差线。误差线是正负两个方向的。本示例应在二维折线图上运行。 Application.Charts.Item("Chart1").SeriesCollection().Item(1).ErrorBar(xlY,xlErrorBarIncludeBoth,xlErrorBarTypeStError)` |

#### **Series.Paste**

从剪贴板中粘贴图片，将其作为所选系列的标志。

**语法**

**express.Paste()**

*express*   一个代表 **Series** 对象的变量。

**说明**

此方法可用于柱形图、条形图、折线图或雷达图，并且将 **MarkerStyle** 属性设为 **xlMarkerStylePicture**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例将剪贴板中的图片粘贴到 Chart1 上的第一个系列中。 Application.Charts.Item("Chart1").SeriesCollection().Item(1).Paste()` |

#### **Series.Points**

返回一个对象，该对象表示数据系列中单个数据点（**Point** 对象）或所有数据点的集合（**Points** 集合）。只读

**语法**

**express.Points(Index)**

*express*   一个代表 **Series** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**             |
| -------- | ------------- | ------------ | -------------------- |
| *Index*  | 可选          | **Variant**  | 数据点的名称或编号。 |

**返回值**

Object

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例对 Chart1 中第一个数据系列的第一个数据点应用数据标签。 Application.Charts.Item("Chart1").SeriesCollection().Item(1).Points(1).ApplyDataLabels()` |

#### **Series.Select**

选择对象。

**语法**

**express.Select()**

*express*   一个代表 **Series** 对象的变量。

**返回值**

Variant

#### **Series.Trendlines**

返回一个对象，该对象表示单个趋势线（**Trendline** 对象）或所有趋势线的集合（**Trendlines** 集合）。

**语法**

**express.Trendlines(Index)**

*express*   一个代表 **Series** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**             |
| -------- | ------------- | ------------ | -------------------- |
| *Index*  | 可选          | **Variant**  | 趋势线的名称或编号。 |

**返回值**

Object

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例向 Chart1 中第一个数据系列添加线性趋势线。 Application.Charts.Item("Chart1").SeriesCollection().Item(1).Trendlines.Add(xlLinear)` |

**成员属性**

#### **Series.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **Series** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例显示一条有关创建 myObject 的应用程序的消息。 function test() {     let myObject = Application.ActiveWorkbook     if(myObject.Application.Value == "ET")      {         alert("This is an ET Application object.")     }     else      {         alert("This is not an ET Application object.")     } }` |

#### **Series.ApplyPictToEnd**

如果图片应用于系列中数据点或所有数据点之后，则为 **True**。**Boolean** 类型，可读写。

**语法**

**express.ApplyPictToEnd**

*express*   一个代表 **Series** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//此示例将图片应用于第一个数据系列中所有数据点之后。图片必须已应用于该系列（设置此属性将更改图片的方向）。 Application.Charts.Item(1).SeriesCollection(1).ApplyPictToEnd = true` |

#### **Series.ApplyPictToFront**

如果图片应用于系列中数据点或所有数据点之前，则为 **True**。**Boolean** 类型，可读写。

**语法**

**express.ApplyPictToFront**

*express*   一个代表 **Series** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//此示例将图片应用于第一个数据系列中所有数据点之前。图片必须已应用于该系列（设置此属性将更改图片的方向）。 Application.Charts.Item(1).SeriesCollection(1).ApplyPictToFront = true` |

#### **Series.ApplyPictToSides**

如果将图片应用于系列中某个数据点或所有数据点的旁边，则为 **True**。**Boolean** 类型，可读写。

**语法**

**express.ApplyPictToSides**

*express*   一个代表 **Series** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//此示例将图片应用于第一个数据系列中所有数据点的旁边。图片必须已应用于该系列（设置此属性将更改图片的方向）。 Application.Charts.Item(1).SeriesCollection(1).ApplyPictToSides = true` |

#### **Series.BarShape**

返回或设置用于三维条形图或柱形图的形状。**XlBarShape** 类型，可读写。

**语法**

**express.BarShape**

*express*   一个代表 **Series** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例设置用于第一个图表的第一个数据系列的形状。 Application.Charts.Item(1).SeriesCollection(1).BarShape = xlConeToPoint` |

#### **Series.BubbleSizes**

返回或设置一个字符串，该字符串引用包含气泡图 x 值、y 值和大小数据的工作表单元格。在返回单元格引用时，它将返回以 A1 样式注释描述单元格的字符串。要为气泡图设置大小数据，必须使用 R1 样式注释。仅适用于气泡图。**Variant** 类型，可读写。

**语法**

**express.BubbleSizes**

*express*   一个代表 **Series** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例显示单元格引用，该单元格中包含气泡图 x 值、y 值和大小数据。 alert(Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(1).BubbleSizes)  //此示例说明如何使用 R1 样式标住设置该属性。 Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(1).BubbleSizes = "=Sheet1!r1c5:r5c5"` |

#### **Series.ChartType**

返回或设置图表类型。**XlChartType** 类型，可读写。

**语法**

**express.ChartType**

*express*   一个代表 **Series** 对象的变量。

**说明**

某些图表类型对数据透视图报表无效。

**示例**

#### **Series.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Series** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **Series.ErrorBars**

返回一个 **ErrorBars** 对象，该对象表示数据系列中的误差线。只读。

**语法**

**express.ErrorBars**

*express*   一个代表 **Series** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例对 Chart1 中第一个系列的误差线颜色进行设置。本示例应在其第一个系列有误差线的二维折线图上运行。 function test() {     let Ser = Application.Charts.Item("Chart1").SeriesCollection(1)     Ser.ErrorBars.Border.ColorIndex = 8 }` |

#### **Series.Explosion**

返回或设置饼图或圆环图的扇区的分离程度值。如果没有分离，则返回 0（零），即扇区的中心与饼图中心重合。**Long** 型，可读写。

**语法**

**express.Explosion**

*express*   一个代表 **Series** 对象的变量。

#### **Series.Format**

返回 **ChartFormat** 对象。只读。

**语法**

**express.Format**

*express*   一个代表 **Series** 对象的变量。

#### **Series.Formula**

 返回或设置一个 **String** 值，它代表 A1 样式表示法和宏语言中的对象的公式。

**语法**

**express.Formula**

*express*   一个代表 **Series** 对象的变量。

**说明**

此属性对于 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源无效。

如果单元格包含一个常量，此属性返回该常量。如果单元格为空，此 **Formula** 属性返回一个空字符串。如果单元格包含公式，**Formula** 属性将该公式作为字符串返回，所用格式与在编辑栏（包括等号）中显示时的格式相同。

如果将单元格的值或者公式设置为日期类型，则 ET 将检查此单元格的数字格式是否符合日期或者时间格式。如果不符合，ET 将把数字格式设置为默认的短日期格式。

如果指定区域是一维或二维区域，则可将公式指定为 Visual Basic 中相同维数的数组。同样，也可在 Visual Basic 数组中使用公式。

如果为多单元格区域设置公式，则会用公式填充该区域所有的单元格。

#### **Series.FormulaLocal**

返回或设置指定对象的公式，使用用户语言 A1 格式引用。**String** 型，可读写。

**语法**

**express.FormulaLocal**

*express*   一个代表 **Series** 对象的变量。

**说明**

如果指定单元格包含常量，此属性返回的就是该常量。如果单元格为空，此属性返回一个空字符串。如果单元格包含公式，此属性将该公式作为字符串返回，所用格式与在编辑栏（包括等号）中显示时的格式相同。

如果将单元格的格式的值或公式设为日期类型，ET 将检查该单元格的格式是否符合某个日期或时间数组格式，如果不符合，将采用默认的短日期数字格式。

如果指定区域是一维或二维区域，则可将公式指定为 Visual Basic 中相同维数的数组。同样，也可在 Visual Basic 数组中使用公式。

对多重单元格区域设置公式，则该区域中所有单元格都用此公式填充。

#### **Series.FormulaR1C1**

返回或设置指定对象的公式，使用宏语言 R1C1 格式符号表示。**String** 型，可读写。

**语法**

**express.FormulaR1C1**

*express*   一个代表 **Series** 对象的变量。

**说明**

如果单元格包含一个常量，此属性返回该常量。如果单元格为空，此属性返回一个空字符串。如果单元格包含公式，此属性将该公式作为字符串返回，所用格式与在编辑栏（包括等号）中显示时的格式相同。

如果将单元格的格式的值或公式设为日期类型，ET 将检查该单元格的格式是否符合某个日期或时间数组格式，如果不符合，将采用默认的短日期数字格式。

如果指定区域是一维或二维区域，则可将公式指定为 Visual Basic 中相同维数的数组。同样，也可在 Visual Basic 数组中使用公式。

对多重单元格区域设置公式，则该区域中所有单元格都用此公式填充。

#### **Series.FormulaR1C1Local**

返回或设置指定对象的公式，使用用户语言 R1C1 格式符号表示。**String** 型，可读写。

**语法**

**express.FormulaR1C1Local**

*express*   一个代表 **Series** 对象的变量。

**说明**

如果指定单元格包含常量，此属性返回的就是该常量。如果单元格为空，此属性返回一个空字符串。如果单元格包含公式，此属性将该公式作为字符串返回，所用格式与在编辑栏（包括等号）中显示时的格式相同。

如果将单元格的格式的值或公式设为日期类型，ET 将检查该单元格的格式是否符合某个日期或时间数组格式，如果不符合，将采用默认的短日期数字格式。

如果指定区域是一维或二维区域，则可将公式指定为 Visual Basic 中相同维数的数组。同样，也可在 Visual Basic 数组中使用公式。

对多重单元格区域设置公式，则该区域中所有单元格都用此公式填充。

#### **Series.Has3DEffect**

如果数据系列有三维外观，则为 **True**。可读/写 **Boolean** 类型。

**语法**

**express.Has3DEffect**

*express*   一个代表 **Series** 对象的变量。

**说明**

该属性仅适用于气泡图。

#### **Series.HasDataLabels**

如果数据系列具有数据标签，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.HasDataLabels**

*express*   一个代表 **Series** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例显示 Chart1 中第三个数据系列的数据标签。 function test() {     let Ser = Application.Charts.Item("Chart1").SeriesCollection(3)     Ser.HasDataLabels = true     Ser.ApplyDataLabels(xlValue) }` |

#### **Series.HasErrorBars**

如果序列有错误条时，该值为 **True**。该属性对三维图表无效。**Boolean** 类型，可读写。

**语法**

**express.HasErrorBars**

*express*   一个代表 **Series** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例删除 Chart1 中第一个数据系列的误差线。本示例应在其第一个系列有误差线的二维折线图上运行。 Application.Charts.Item("Chart1").SeriesCollection(1).HasErrorBars = false` |

#### **Series.HasLeaderLines**

如果数据系列有引导线，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.HasLeaderLines**

*express*   一个代表 **Series** 对象的变量。

**说明**

该属性仅应用于饼图。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例向饼图上的数据系列 1 添加数据标签和蓝色引导线。如果看不到引导线，则本示例代码将失败。这种情况下，可以手动将数据标签之一拖出饼图以便让引导线显示。 function test() {     let Ser = Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(1)     Ser.HasDataLabels = true     Ser.DataLabels.Position = xlLabelPositionBestFit     Ser.HasLeaderLines = true     Ser.LeaderLines.Border.ColorIndex = 5 }` |

#### **Series.InvertColor**

返回或设置系列中负数据点的填充色。可读/写

**语法**

**express.InvertColor**

*express*   一个代表 **Series** 对象的变量。

**说明**

**InvertColor** 属性可用来设置负数据点的填充色，颜色可设置为特定的数值、十六进制值、八进制值或 RGB 颜色值。若要将该值设置为 RGB 值，请使用 Visual Basic **RGB** 函数。如果不使用 **InvertColor** 属性，还可以使用 **InvertColorIndex** 属性，它使用当前调色板上更简单的整数值集合。

为了使 **InvertColor** 属性生效，还必须将 **Series** 对象的 **InvertIfNegative** 属性设置为 **True**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//下面的代码示例将“Chart 2”第一个系列中的负数据点的填充色设置为洋红。 function test() {     Application.ActiveSheet.ChartObjects("Chart 2").Activate()     Application.ActiveChart.SeriesCollection(1).InvertIfNegative = true     Application.ActiveChart.SeriesCollection(1).InvertColor = 255, 0, 255 }` |

#### **Series.InvertColorIndex**

返回或设置系列中负数据点的填充色。可读/写

**语法**

**express.InvertColorIndex**

*express*   一个代表 **Series** 对象的变量。

**说明**

**InvertColorIndex** 属性可用来将负数据点的填充色设置为 0 到 56 之间的颜色索引值。有关颜色索引值的详细信息，请参阅使用 ColorIndex 属性将颜色添加到 ET工作表。如果不使用 **InvertColorIndex** 属性，还可以使用 **InvertColor** 属性，该属性可用来将颜色值作为特定的数值、十六进制值、八进制值和 RGB 颜色值来设置颜色。

为了使 **InvertColorIndex** 属性生效，还必须将 **Series** 对象的 **InvertIfNegative** 属性设置为 **True**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//下面的代码示例将“Chart 2”第一个系列中的负数据点的填充色设置为洋红。 function test() {     Application.ActiveSheet.ChartObjects("Chart 2").Activate()     Application.ActiveChart.SeriesCollection(1).InvertIfNegative = true     Application.ActiveChart.SeriesCollection(1).InvertColorIndex = 7 }` |

#### **Series.InvertIfNegative**

如果指定项与一个负数相对应时 ET 就将其反色，则为 **True**。**Boolean** 类型，可读写。

**语法**

**express.InvertIfNegative**

*express*   一个代表 **Series** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//此示例反转 Chart1 中第一个数据系列的负值的图案。此示例须在二维柱形图上运行。 Application.Charts.Item("Chart1").SeriesCollection(1).InvertIfNegative = true` |

#### **Series.LeaderLines**

返回一个 **LeaderLines** 对象，该对象表示系列的引导线。只读。

**语法**

**express.LeaderLines**

*express*   一个代表 **Series** 对象的变量。

**说明**

该属性仅应用于饼图。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例向饼图上的数据系列 1 添加数据标签和蓝色引导线。如果看不到引导线，则本示例代码将失败。这种情况下，可以手动将数据标签之一拖出饼图以便让引导线显示。 function test() {     let Ser = Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(1)     Ser.HasDataLabels = true     Ser.DataLabels.Position = xlLabelPositionBestFit     Ser.HasLeaderLines = true     Ser.LeaderLines.Border.ColorIndex = 5 }` |

#### **Series.MarkerBackgroundColor**

将数据标志的背景色设置为 RGB 值或返回对应的颜色索引值。仅适用于折线图、散点图和雷达图。**Long** 型，可读写。

**语法**

**express.MarkerBackgroundColor**

*express*   一个代表 **Series** 对象的变量。

#### **Series.MarkerBackgroundColorIndex**

返回或设置数据标志的背景色，表示为当前调色板中的索引或下列 **XlColorIndex** 常量之一：**xlColorIndexAutomatic** 或 **xlColorIndexNone**。仅适用于折线图、散点图和雷达图。**Long** 类型，可读写。

**语法**

**express.MarkerBackgroundColorIndex**

*express*   一个代表 **Series** 对象的变量。

#### **Series.MarkerForegroundColor**

将数据标志的背景色设置为 RGB 值或返回对应的颜色索引值。仅适用于折线图、散点图和雷达图。**Long** 型，可读写。

**语法**

**express.MarkerForegroundColor**

*express*   一个代表 **Series** 对象的变量。

#### **Series.MarkerForegroundColorIndex**

返回或设置数据标志的前景色，表示为当前调色板中的索引或下列 **XlColorIndex** 常量之一：**xlColorIndexAutomatic** 或 **xlColorIndexNone**。仅适用于折线图、散点图和雷达图。**Long** 类型，可读写。

**语法**

**express.MarkerForegroundColorIndex**

*express*   一个代表 **Series** 对象的变量。

#### **Series.MarkerSize**

返回或设置数据标志的大小，以磅（磅：指打印的字符的高度的度量单位。1 磅等于 1/72 英寸，或大约等于 1 厘米的 1/28。）为单位。可以是 2 到 72 之间的一个值。**Long** 型，可读写。

**语法**

**express.MarkerSize**

*express*   一个代表 **Series** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//此示例设置系列一中所有数据标志的尺寸。 Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(1).MarkerSize = 10` |

#### **Series.MarkerStyle**

返回或设置折线图、散点图或雷达图中数据点或数据系列的数据标志样式。**XlMarkerStyle** 类型，可读写。

**语法**

**express.MarkerStyle**

*express*   一个代表 **Series** 对象的变量。

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

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例设置 Chart1 中第一个数据系列的标记样式。本示例应在二维折线图上运行。 Application.Charts.Item("Chart1").SeriesCollection(1).MarkerStyle = xlMarkerStyleCircle` |

#### **Series.Name**

返回或设置一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

*express*   一个代表 **Series** 对象的变量。

**说明**

您可以使用 R1C1 符号进行引用，例如“=Sheet1!R1C1”。

#### **Series.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **Series** 对象的变量。

#### **Series.PictureType**

返回或设置一个 **XlChartPictureType** 值，它代表在柱形或条形图上图片的显示方式。

**语法**

**express.PictureType**

*express*   一个代表 **Series** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例设置 Chart1 中的系列一伸展图片。本示例应在带图片数据标志的二维柱形图上运行。 Application.Charts.Item("Chart1").SeriesCollection(1).PictureType = xlStretch` |

#### **Series.PictureUnit2**

如果 **PictureType** 属性设置为 **xlStackScale**，则返回或设置图表上每一个图片的单位。否则忽略此属性。可读/写 **Double** 类型。

**语法**

**express.PictureUnit2**

*express*   一个代表 **Series** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//此示例将 Chart1 中的系列一设为堆积图片，设置每一图片代表五个单位。此示例应在带图片数据标志的二维柱形图上运行。 function test() {     let Ser = Application.Charts.Item("Chart1").SeriesCollection(1)     Ser.PictureType = xlScale     Ser.PictureUnit2 = 5 }` |

#### **Series.PlotColorIndex**

返回内部用于将系列格式设置与图表元素关联的索引值。只读

**语法**

**express.PlotColorIndex**

*express*   一个代表 **Series** 对象的变量。

**说明**

此属性由 ET 内部使用，并非用于从代码中调用。

#### **Series.PlotOrder**

返回或设置图表组中选定数据系列的绘制顺序。**Long** 类型，可读写。

**语法**

**express.PlotOrder**

*express*   一个代表 **Series** 对象的变量。

**说明**

只能设置图表组内的绘制顺序（在图表中包含若干图表类型的情况下，不能设置整个图表的绘制顺序）。图表组指具有相同图表类型和子类型的数据系列的集合。

更改某个系列的绘制顺序将导致同一图表组内其他数据系列的绘制顺序的必要调整。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例使 Chart1 中的数据系列二在绘制时第三个出现。本示例应在至少包含三个数据系列的二维柱形图上运行。 Application.Charts.Item("Chart1").ChartGroups(1).SeriesCollection(2).PlotOrder = 3` |

#### **Series.Shadow**

返回或设置一个 **Boolean** 值，它确定对象是否有阴影。

**语法**

**express.Shadow**

*express*   一个代表 **Series** 对象的变量。

#### **Series.Smooth**

如果折线图或散点图有平滑线，则为 **True**。仅适用于折线图和散点图。可读写。

**语法**

**express.Smooth**

*express*   一个代表 **Series** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//此示例使 Chart1 中的系列一具有平滑线。此示例应在二维折线图上运行。 Application.Charts.Item("Chart1").SeriesCollection(1).Smooth = true` |

#### **Series.Type**

返回或设置一个代表系列类型的 Long 值。

**语法**

**express.Type**

*express*   一个代表 **Series** 对象的变量。

**说明**

| 以下常量可以用于此属性。 |
| ------------------------ |
| **xlColumn**             |
| **xlBar**                |
| **xl3DBar**              |
| **xlLine**               |
| **xlPie**                |
| **xlXYScatter**          |
| **xlArea**               |
| **xl3DArea**             |
| **xlDoughnut**           |
| **xlRadar**              |
| **xl3DSurface**          |
| **xl3DColumn**           |
| **xl3DBar**              |

#### **Series.Values**

返回或设置一个 **Variant** 值，它代表系列中所有值的集合。

**语法**

**express.Values**

*express*   一个代表 **Series** 对象的变量。

**说明**

此属性值可以是工作表中的某一区域或常量值的数组，但不能是两者的组合。有关详细信息，请参阅此示例。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//此示例设置区域的系列值。 Application.Charts.Item("Chart1").SeriesCollection(1).Values = Application.Worksheets.Item("Sheet1").Range("C5:T5")  //若要为每个单独的数据点赋以常量值，则必须使用数组。 Application.Charts.Item("Chart1").SeriesCollection(1).Values = [1, 3, 5, 7, 11, 13, 17, 19]` |

#### **Series.XValues**

返回或设置图表系列中 x 值的数组。**XValues** 属性可设置为工作表区域或数值数组，但不能为二者的组合。**Variant** 类型，可读写。

**语法**

**express.XValues**

*express*   一个代表 **Series** 对象的变量。

**说明**

对于数据透视图，该属性为只读。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例将 Chart1 中系列一的 x 值设置为 Sheet1 上 B1:B5 区域的值。 Application.Charts.Item("Chart1").SeriesCollection(1).XValues = Application.Worksheets.Item("Sheet1").Range("B1:B5")  //本示例使用数组为 Chart1 中系列一的各点逐一赋值。 Application.Charts.Item("Chart1").SeriesCollection(1).XValues = [5.0, 6.3, 12.6, 28, 50]` |

适用环境：web

适用平台：windows/linux