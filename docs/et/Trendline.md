**Trendline**



代表图表上的趋势线。

**说明**

趋势线显示系列中数据的趋势或方向。**Trendline** 对象是 **Trendlines** 集合的成员。**Trendlines** 集合包含某一个系列的所有 **Trendline** 对象。

使用 **Trendlines**(*index*)（其中 *index* 是趋势线索引号）可返回一个 **Trendline** 对象。

索引号指出趋势线添加到系列中的顺序。`Trendlines(1)` 是第一个添加到系列中的趋势线，而 `Trendlines(Trendlines.Count)` 是最后一个。

下例更改工作表一上嵌入式图表一中第一个系列的趋势线类型。如果该系列没有趋势线，则此示例会失败。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `function test() {     Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(1).Trendlines(1).Type = Application.Enum.xlMovingAvg }` |

**方法**

|                                                              | 名称             | 说明                 |
| ------------------------------------------------------------ | ---------------- | -------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ClearFormats** | 清除对象的格式设置。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**       | 删除对象。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Select**       | 选择对象。           |

**属性**

|                                                              | 名称                | 说明                                                         |
| ------------------------------------------------------------ | ------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**     | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Backward2**       | 返回或设置趋势线向后延伸的周期数目（或散点图中的单位个数）。可读/写 **Double** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Border**          | 返回一个 **Border** 对象，它代表对象的边框。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DataLabel**       | 返回一个 **DataLabel** 对象，它代表与趋势线相关的数据标签。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayEquation** | 如果显示图表中的趋势线公式，则该值为 **True**（其数据标签与 R-平方值相同）。将该属性设为 **True** 可自动显示数据标签。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayRSquared** | 如果显示图表中趋势线的 R-平方值，则该值为 **True**（其数据标签与公式的相同）。将该属性设为 **True** 可自动显示数据标签。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Format**          | 返回 **ChartFormat** 对象。只读。                            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Forward2**        | 返回或设置趋势线向前延伸的周期数目（或散点图中的单位个数）。可读/写 **Double** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Index**           | 返回 **Long** 值，它代表对象在其同类对象所组成的集合内的索引号。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Intercept**       | 返回或设置趋势线与数值轴的交叉点。可读/写 **Double** 类型。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **InterceptIsAuto** | 如果趋势线与数值轴的交叉点由回归分析自动确定，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**            | 返回或设置一个 **String** 值，它代表对象的名称。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NameIsAuto**      | 如果 ET 自动确定趋势线的名称，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Order**           | 返回或设置一个 **Long** 值，它代表趋势线类型为 **xlPolynomial** 时趋势线的顺序号（大于 1 的整数）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**          | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Period**          | 返回或设置移动平均趋势线的周期。可以是从 2 至 255 的值。可读/写 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**            | 返回或设置一个 **XlTrendlineType** 值，它代表趋势线的类型。  |

**成员方法**

#### **Trendline.ClearFormats**

清除对象的格式设置。

**语法**

**express.ClearFormats()**

*express*   一个代表 **Trendline** 对象的变量。

**返回值**

Variant

#### **Trendline.Delete**

删除对象。

**语法**

**express.Delete()**

*express*   一个代表 **Trendline** 对象的变量。

**返回值**

Variant

#### **Trendline.Select**

选择对象。

**语法**

**express.Select()**

*express*   一个代表 **Trendline** 对象的变量。

**返回值**

Variant

**成员属性**

#### **Trendline.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **Trendline** 对象的变量。

**示例**

本示例显示一条有关创建 `myObject` 的应用程序的消息。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let myObject = Application.ActiveWorkbook     if (myObject.Application.Value == "ET") {         alert("This is an ET Application object.")     }     else {         alert("This is not an ET Application object.")     } }` |

#### **Trendline.Backward2**

返回或设置趋势线向后延伸的周期数目（或散点图中的单位个数）。可读/写 **Double** 类型。

**语法**

**express.Backward2**

*express*   一个代表 **Trendline** 对象的变量。

**示例**

本示例对 Chart1 中的趋势线向前和向后延伸的单位数进行设置。本示例应在包含单个带趋势线系列的二维柱形图上运行。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let rng = Application.Charts.Item("Chart1").SeriesCollection(1).Trendlines(1)     rng.Forward2 = 5     rng.Backward2 = 0.5 }` |

#### **Trendline.Border**

返回一个 **Border** 对象，它代表对象的边框。

**语法**

**express.Border**

*express*   一个代表 **Trendline** 对象的变量。

#### **Trendline.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Trendline** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **Trendline.DataLabel**

返回一个 **DataLabel** 对象，它代表与趋势线相关的数据标签。只读。

**语法**

**express.DataLabel**

*express*   一个代表 **Trendline** 对象的变量。

**说明**

若要打开趋势线的数据标签，需要将 **DisplayEquation** 属性和 **DisplayRSquared** 属性设置为 **True**。

#### **Trendline.DisplayEquation**

如果显示图表中的趋势线公式，则该值为 **True**（其数据标签与 R-平方值相同）。将该属性设为 **True** 可自动显示数据标签。**Boolean** 类型，可读写。

**语法**

**express.DisplayEquation**

*express*   一个代表 **Trendline** 对象的变量。

**示例**

本示例显示 Chart1 中趋势线一的 R-平方值和公式。本示例应在其第一个系列有趋势线的二维柱形图上运行。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let rng = Application.Charts.Item("Chart1").SeriesCollection(1).Trendlines(1)     rng.DisplayRSquared = true     rng.DisplayEquation = true }` |

#### **Trendline.DisplayRSquared**

如果显示图表中趋势线的 R-平方值，则该值为 **True**（其数据标签与公式的相同）。将该属性设为 **True** 可自动显示数据标签。**Boolean** 类型，可读写。

**语法**

**express.DisplayRSquared**

*express*   一个代表 **Trendline** 对象的变量。

**示例**

本示例显示 Chart1 中趋势线一的 R-平方值和公式。本示例应在其第一个系列有趋势线的二维柱形图上运行。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let rng = Application.Charts.Item("Chart1").SeriesCollection(1).Trendlines(1)     rng.DisplayRSquared = true     rng.DisplayEquation = true }` |

#### **Trendline.Format**

返回 **ChartFormat** 对象。只读。

**语法**

**express.Format**

*express*   一个代表 **Trendline** 对象的变量。

#### **Trendline.Forward2**

返回或设置趋势线向前延伸的周期数目（或散点图中的单位个数）。可读/写 **Double** 类型。

**语法**

**express.Forward2**

*express*   一个代表 **Trendline** 对象的变量。

**示例**

本示例对 Chart1 中的趋势线向前和向后延伸的单位数进行设置。本示例应在包含单个带趋势线系列的二维柱形图上运行。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let rng = Application.Charts.Item("Chart1").SeriesCollection(1).Trendlines(1)     rng.Forward2 = 5     rng.Backward2 = 0.5 }` |

#### **Trendline.Index**

返回 **Long** 值，它代表对象在其同类对象所组成的集合内的索引号。

**语法**

**express.Index**

*express*   一个代表 **Trendline** 对象的变量。

#### **Trendline.Intercept**

返回或设置趋势线与数值轴的交叉点。可读/写 **Double** 类型。

**语法**

**express.Intercept**

*express*   一个代表 **Trendline** 对象的变量。

**说明**

设置该属性会将 **InterceptIsAuto** 属性设置为 **False**。

**示例**

本示例设置 Chart1 的第一条趋势线与数值轴相交于 5。本示例应在包含单个带趋势线系列的二维柱形图上运行。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     Application.Charts.Item("Chart1").SeriesCollection(1).Trendlines(1).Intercept = 5 }` |

#### **Trendline.InterceptIsAuto**

如果趋势线与数值轴的交叉点由回归分析自动确定，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.InterceptIsAuto**

*express*   一个代表 **Trendline** 对象的变量。

**说明**

设置 **Intercept** 属性会将该属性设置为 **False**。

**示例**

本示例设置 ET 自动判断 Chart1 的趋势线与数值坐标轴的交点。本示例应在包含单个带趋势线系列的二维柱形图上运行。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     Application.Charts.Item("Chart1").SeriesCollection(1).Trendlines(1).InterceptIsAuto = true }` |

#### **Trendline.Name**

返回或设置一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

*express*   一个代表 **Trendline** 对象的变量。

#### **Trendline.NameIsAuto**

如果 ET 自动确定趋势线的名称，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.NameIsAuto**

*express*   一个代表 **Trendline** 对象的变量。

**示例**

本示例设置 ET 自动确定 Chart1 的趋势线一的名称。本示例应在包含单个带趋势线系列的二维柱形图上运行。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     Application.Charts.Item("Chart1").SeriesCollection(1).Trendlines(1).NameIsAuto = true }` |

#### **Trendline.Order**

返回或设置一个 **Long** 值，它代表趋势线类型为 **xlPolynomial** 时趋势线的顺序号（大于 1 的整数）。

**语法**

**express.Order**

*express*   一个代表 **Trendline** 对象的变量。

#### **Trendline.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **Trendline** 对象的变量。

#### **Trendline.Period**

返回或设置移动平均趋势线的周期。可以是从 2 至 255 的值。可读/写 **Long** 类型。

**语法**

**express.Period**

*express*   一个代表 **Trendline** 对象的变量。

**示例**

本示例设置 Chart1 上移动平均趋势线的周期。本示例应在二维柱形图上运行，该柱形图中只有一个包含 10 个数据点的系列，并且有移动平均趋势线。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let rng = Application.Charts.Item("Chart1").SeriesCollection(1).Trendlines(1)     if (rng.Type == Application.Enum.xlMovingAvg) {         rng.Period = 5     } }` |

#### **Trendline.Type**

返回或设置一个 **XlTrendlineType** 值，它代表趋势线的类型。

**语法**

**express.Type**

*express*   一个代表 **Trendline** 对象的变量。

**示例**

本示例更改第一个系列（第一张工作表的第一个嵌入图表中）的趋势线类型。如果该系列没有趋势线，则本示例无效。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(1).Trendlines(1).Type = Application.Enum.xlMovingAvg }` |

适用环境：web

适用平台：windows/linux