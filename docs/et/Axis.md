**Axis**



代表图表中的单个坐标轴。

**说明**

**Axis** 对象是 **Axes** 集合的成员。

使用 **Axes**(*type*, *group* )（其中 *type* 为坐标轴类型，而 *group* 为坐标轴组）可返回单个 **Axis** 对象。*Type* 可为以下 **XlAxisType** 常量之一：**xlCategory**、**xlSeries** 或 **xlValue**。*Group* 可为以下 **XlAxisGroup** 常量之一：**xlPrimary** 或 **xlSecondary**。有关详细信息，请参阅 **Axes** 方法。

示例

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例在名为“Chart1”的图表工作表中设置分类轴的标题文本。*/ function test() {     let myChart = Application.Charts.Item("Chart1").Axes.Item(xlCategory)     myChart.HasTitle = true     myChart.AxisTitle.Caption = "1994" }` |

**方法**

|                                                              | 名称       | 说明       |
| ------------------------------------------------------------ | ---------- | ---------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete** | 删除对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Select** | 选择对象。 |

**属性**

|                                                              | 名称                       | 说明                                                         |
| ------------------------------------------------------------ | -------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**            | 如果不使用对象识别符，则该属性返回一个 **App****lication** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AxisBetweenCategories**  | 如果数值轴与分类坐标轴相交于分类之间，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AxisGroup**              | 返回指定轴的组。只读                                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AxisTitle**              | 返回一个 **AxisTitle** 对象，该对象表示指定坐标轴的标题。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BaseUnit**               | 返回或设置指定分类轴的基本单位。**XlTim****eUnit** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BaseUnitIsAuto**         | 如果由 ET 为指定的分类轴选取适当的基本单位，则该值为 **True**。默认值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Border**                 | 返回一个 **Border** 对象，它代表对象的边框。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CategoryNames**          | 以文本数组形式返回或设置指定坐标轴中所有分类的名称。该属性既可设为一个数组，也可设为一个包含所有分类名称的 **Range** 对象。**Variant** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CategoryType**           | 返回或设置分类轴类型。**XlCategoryType** 类型，可读写。      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**                | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Crosses**                | 返回或设置指定坐标轴与其他坐标轴相交的点。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CrossesAt**              | 返回或设置数值轴中与分类坐标轴的交点。仅应用于数值轴。**Double** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayUnit**            | 返回或设置数值轴的单位标签。**XlDisplayUnit**、**xlCustom** 或 **xlNone** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayUnitCustom**      | 如果 **DisplayUnit** 属性的值是 **xlCustom**，则 **DisplayUnitCustom** 返回或设置显示的单位的值。该值范围必须是从 0 到 10E307。**Double** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayUnitLabel**       | 返回指定坐标轴的 **DisplayUnitLabel** 对象。如果 **HasDisplayUnitLabel** 属性设置为 **False**，则返回 **null**。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Format**                 | 返回 ChartFormat 对象。只读。                                |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HasDisplayUnitLabel**    | 如果由 **DisplayUnit** 或 **DisplayUnitCustom** 属性指定的标签显示在指定坐标轴上，则该属性值为 **True**。默认值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HasMajorGridlines**      | 如果坐标轴有主要网格线，则该值为 **True**。只有主要坐标轴组中的坐标轴才能有网格线。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HasMinorGridlines**      | 如果坐标轴有次要网格线，则该值为 **True**。只有主要坐标轴组中的坐标轴才能有网格线。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HasTitle**               | 如果坐标轴或图表有可见标题，则为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Height**                 | 返回一个 **Double** 值，它代表对象的高度（以磅为单位）。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Left**                   | 返回一个 **Double** 值，它代表从对象左边缘到图表区左边缘的距离（以磅为单位）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LogBase**                | 在使用对数刻度时返回或设置对数的底。可读/写 **Double** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MajorGridlines**         | 返回一个 **Gridlines** 对象，该对象表示指定坐标轴的主要网格线。只有主要坐标轴组中的坐标轴才能有网格线。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MajorTickMark**          | 返回或设置指定坐标轴的主刻度线类型。**XlTickMark** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MajorUnit**              | 返回或设置数值轴的主要单位。**Double** 类型，可读写。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MajorUnitIsAuto**        | 如果 ET 计算数值轴的主要单位，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MajorUnitScale**         | 当 **CategoryType** 属性设置为 **xlTimeScale** 时，返回或设置分类轴主要单位的刻度值。**XlTimeUnit** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MaximumScale**           | 返回或设置数值轴上的最大值。**Double** 类型，可读写。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MaximumScaleIsAuto**     | 如果 ET 计算该数值轴的最大值，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MinimumScale**           | 返回或设置数值轴上的最小值。**Double** 类型，可读写。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MinimumScaleIsAuto**     | 如果 ET 为数值轴计算最小值，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MinorGridlines**         | 返回 **Gridlines** 对象，该对象表示指定坐标轴的次要网格线。只有主要坐标轴组中的坐标轴才能有网格线。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MinorTickMark**          | 返回或设置指定坐标轴的次要刻度线类型。**XlTickMark** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MinorUnit**              | 返回或设置数值轴上的次要单位。**Double** 类型，可读写。      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MinorUnitIsAuto**        | 如果 ET 为数值轴计算次要单位，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MinorUnitScale**         | 返回或设置当 **CategoryType** 属性设置为 **xlTimeScale** 时分类轴次要单位的刻度值。**XlTimeUnit** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**                 | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ReversePlotOrder**       | 如果 ET 的绘图区数据点的顺序为从后往前，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ScaleType**              | 返回或设置数值轴的刻度类型。**XlScaleType** 类型，可读写。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TickLabelPosition**      | 说明指定坐标轴上刻度线标志的位置。**XlTickLabelPosition** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TickLabelSpacing**       | 返回或设置刻度线标签之间的分类数或数据系列数。仅应用于分类轴和系列轴。可以是 1 到 31999 之间的一个数值。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TickLabelSpacingIsAuto** | 返回或设置刻度标签间距是否是自动设置的。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TickLabels**             | 返回一个 **TickLabels** 对象，该对象表示指定坐标轴的刻度线标签。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TickMarkSpacing**        | 返回或设置刻度线之间的分类数或数据系列数。仅应用于分类轴和系列轴。可以是 1 到 31999 之间的一个数值。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Top**                    | 返回一个 **Double** 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**                   | 返回一个 **XlAxisType** 值，它代表“坐标轴”类型。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Width**                  | 返回一个 **Double** 值，它代表对象的宽度（以磅为单位）。     |

**成员方法**

#### **Axis.Delete**

删除对象。

**语法**

**express.Delete()**

*express*   一个代表 **Axis** 对象的变量。

#### **Axis.Select**

选择对象。

**语法**

**express.Select()**

*express*   一个代表 **Axis** 对象的变量。

**成员属性**

#### **Axis.Application**

如果不使用对象识别符，则该属性返回一个 **App****lication** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **Axis** 对象的变量。

#### **Axis.AxisBetweenCategories**

如果数值轴与分类坐标轴相交于分类之间，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.AxisBetweenCategories**

*express*   一个代表 **Axis** 对象的变量。

**说明**

该属性仅适用于分类轴，而不适用于三维图表。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 Chart1 的数值轴与分类坐标轴相交于两个分类之间。*/ Application.Charts.Item("Chart1").Axes(xlCategory).AxisBetweenCategories = true` |

#### **Axis.AxisGroup**

返回指定轴的组。只读

**语法**

**express.AxisGroup**

*express*   一个代表 **Axis** 对象的变量。

#### **Axis.AxisTitle**

返回一个 **AxisTitle** 对象，该对象表示指定坐标轴的标题。只读。

**语法**

**express.AxisTitle**

*express*   一个代表 **Axis** 对象的变量。

**说明**

本示例为 Chart1 的分类轴添加轴标签。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {   let myChart = Application.Charts.Item("Chart1").Axes(xlCategory)   myChart.HasTitle = true   myChart.AxisTitle.Text = "July Sales" }` |

#### **Axis.BaseUnit**

返回或设置指定分类轴的基本单位。**XlTim****eUnit** 类型，可读写。

**语法**

**express.BaseUnit**

*express*   一个代表 **Axis** 对象的变量。

**说明**

如果指定坐标轴的 **CategoryType** 属性设置为 **xlCategoryScale**，则设置该属性不会看到任何效果。不过，将保留设置值，并在 **CategoryType** 属性设置为 **xlTimeScale** 时起作用。
不能对数值轴设置本属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例使第一张工作表上嵌入的第一个图表中的分类轴使用时间刻度，并以月为基本单位。*/ function test() {   let axes = Application.Worksheets.Item(1).ChartObjects(1).Chart.Axes(xlCategory)      axes.CategoryType = xlTimeScale   axes.BaseUnit = xlMonths }` |

#### **Axis.BaseUnitIsAuto**

如果由 ET 为指定的分类轴选取适当的基本单位，则该值为 **True**。默认值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.BaseUnitIsAuto**

*express*   一个代表 **Axis** 对象的变量。

**说明**

不能对数值轴设置本属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例使第一张工作表上嵌入的第一个图表中的分类轴使用时间刻度，并使用自动选取的基本单位。*/ function test() {   let axes = Application.Worksheets.Item(1).ChartObjects(1).Chart.Axes(xlCategory)      axes.CategoryType = xlTimeScale   axes.BaseUnitIsAuto = true }` |

#### **Axis.Border**

返回一个 **Border** 对象，它代表对象的边框。

**语法**

**express.Border**

*express*   一个代表 **Axis** 对象的变量。

#### **Axis.CategoryNames**

以文本数组形式返回或设置指定坐标轴中所有分类的名称。该属性既可设为一个数组，也可设为一个包含所有分类名称的 **Range** 对象。**Variant** 类型，可读写。

**语法**

**express.CategoryNames**

*express*   一个代表 **Axis** 对象的变量。

**说明**

该属性仅适用于分类轴。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 Chart1 的分类名设为 Sheet1 中 B1:B5 单元格的值。*/ Application.Charts("Chart1").Axes(xlCategory).CategoryNames = Application.Worksheets("Sheet1").Range("B1:B5")  /*本示例使用数组对 Chart1 的个别分类名进行设置。*/ Charts("Chart1").Axes(xlCategory).CategoryNames = Array("1985", "1986", "1987", "1988", "1989")` |

#### **Axis.CategoryType**

返回或设置分类轴类型。**XlCategoryType** 类型，可读写。

**语法**

**express.CategoryType**

*express*   一个代表 **Axis** 对象的变量。

**说明**

不能对数值轴设置本属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例使第一张工作表上嵌入的第一个图表中的分类轴使用时间刻度，并以月为基本单位。*/ function test() {   let chart = Application.Worksheets(1).CharObjects(1).Chart   let axes = chart.Axes(xlCategory)   axes.CategoryType = xlTimeScale   axes.BaseUnit = xlMonths }` |

#### **Axis.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Axis** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **Axis.Crosses**

返回或设置指定坐标轴与其他坐标轴相交的点。**Long** 类型，可读写。

**语法**

**express.Crosses**

*express*   一个代表 **Axis** 对象的变量。

**说明**

可以是下表所列的 **XlAxisCrosses** 常量之一。

| 常量                       | 含义                               |
| -------------------------- | ---------------------------------- |
| **xlAxisCrossesAutomatic** | 由 ET 设置坐标轴交点。             |
| **xlMinimum**              | 坐标轴在最小值处相交。             |
| **xlMaximum**              | 坐标轴在最大值处相交。             |
| **xlAxisCrossesCustom**    | **CrossesAt** 属性指定坐标轴交点。 |

该属性对雷达图无效。对于三维图表，该属性只能应用于数值轴并指示由分类轴所定义的平面与数值轴相交的位置。
本属性既可用于分类轴，也可用于数值轴。在分类轴上，xlMinimum 使数值轴在第一个分类处与之相交，而 xlMaximum 使数值轴在最后一个分类处与之相交。
请注意，根据坐标轴的不同，xlMinimum 和 xlMaximum 可能有不同的含义。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 Chart1 的数值轴与分类坐标轴的交点在最大的 x 值处。*/ Application.Charts.Item("Chart1").Axes(xlCategory).Crosses = xlMaximum` |

#### **Axis.CrossesAt**

返回或设置数值轴中与分类坐标轴的交点。仅应用于数值轴。**Double** 类型，可读写。

**语法**

**express.CrossesAt**

*express*   一个代表 **Axis** 对象的变量。

**说明**

设置该属性将使 **Crosses** 属性更改为 **xlAxisCrossesCustom**。

该属性不能用于雷达图。对于三维图表，本属性指示由分类轴和数值轴交叉所定义的平面的位置。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置活动图表中的分类坐标轴与数值轴的交点位于数值坐标轴上数值 3 的位置。*/ function Chart(){     // Create a sample source of data.     Application.Range("A1").Value2 = "2"     Application.Range("A2").Value2 = "4"     Application.Range("A3").Value2 = "6"     Application.Range("A4").Value2 = "3"      // Create a chart based on the sample source of data.     Application.Charts.Add()      Application.ActiveChart.ChartType = xlLineMarkersStacked     Application.ActiveChart.SetSourceData(Sheets.Item("Sheet1").Range("A1:A4"), xlColumns)     Application.ActiveChart.Location(xlLocationAsObject, "Sheet1")      // Set the category axis to cross the value axis at value 3.     Application.ActiveChart.Axes(xlValue).Select()     Application.Selection.CrossesAt = 3 }` |

#### **Axis.DisplayUnit**

返回或设置数值轴的单位标签。**XlDisplayUnit**、**xlCustom** 或 **xlNone** 类型，可读写。

**语法**

**express.DisplayUnit**

*express*   一个代表 **Axis** 对象的变量。

**说明**

| **XlDisplayUnit** 可为以下 **XlDisplayUnit** 常量之一。 |
| ------------------------------------------------------- |
| **xlHundredMillions**                                   |
| **xlHundredThousands**                                  |
| **xlMillions**                                          |
| **xlTenThousands**                                      |
| **xlThousands**                                         |
| **xlHundreds**                                          |
| **xlMillionMillions**                                   |
| **xlTenMillions**                                       |
| **xlThousandMillions**                                  |
| 此单位标签还可为以下常量之一                            |
| **xlCustom**                                            |
| **xlNone**                                              |

在绘制大数值图表时，使用单位标志可使刻度线标志易于阅读。例如，如果将数值轴的单位设置为百、千或百万，则可以在坐标轴的刻度标志上使用较小的数字值。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 Chart1 中数值轴的显示单位设置为百。*/ function test() {   let axes = Application.Charts.Item("Chart1").Axes(xlValue)   axes.DisplayUnit = xlHundreds   axes.HasTitle = true   axes.AxisTitle.Caption = "Rebate Amounts" }` |

#### **Axis.DisplayUnitCustom**

如果 **DisplayUnit** 属性的值是 **xlCustom**，则 **DisplayUnitCustom** 返回或设置显示的单位的值。该值范围必须是从 0 到 10E307。**Double** 类型，可读写。

**语法**

**express.DisplayUnitCustom**

*express*   一个代表 **Axis** 对象的变量。

**说明**

在绘制大数值图表时，使用单位标志可使刻度线标志易于阅读。例如，如果将数值轴的单位设置为百、千或百万，则可以在坐标轴的刻度标志上使用较小的数字值。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 Chart1 中数值轴上显示的单位设置为 500。*/ function test() {   let axes = Application.Charts.Item("Chart1").Axes(xlValue)   axes.DisplayUnit = xlCustom   axes.DisplayUnitCustom = 500   axes.HasTitle = true   axes.AxisTitle.Caption = "Rebate Amounts" }` |

#### **Axis.DisplayUnitLabel**

返回指定坐标轴的 **DisplayUnitLabel** 对象。如果 **HasDisplayUnitLabel** 属性设置为 **False**，则返回 **null**。只读。

**语法**

**express.DisplayUnitLabel**

*express*   一个代表 **Axis** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 Chart1 中数值轴的标签标题设置为“Millions”，然后关闭自动字体缩放。*/ function test() {   let myLabel = Application.Charts.Item("Chart1").Axes(xlValue).DisplayUnitLabel   myLabel.Caption = "Millions"   myLabel.AutoScaleFont = false }` |

#### **Axis.Format**

返回 ChartFormat 对象。只读。

**语法**

**express.Format**

*express*   一个代表 **Axis** 对象的变量。

#### **Axis.HasDisplayUnitLabel**

如果由 **DisplayUnit** 或 **DisplayUnitCustom** 属性指定的标签显示在指定坐标轴上，则该属性值为 **True**。默认值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.HasDisplayUnitLabel**

*express*   一个代表 **Axis** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 Chart1 中数值轴上的单位设置为 500，但隐藏单位标志。*/ function test() {   let axes = Application.Charts.Item("Chart1").Axes(xlValue)   axes.DisplayUnit = xlCustom   axes.DisplayUnitCustom = 500   axes.AxisTitle.Caption = "Rebate Amounts"   axes.HasDisplayUnitLabel = false }` |

#### **Axis.HasMajorGridlines**

如果坐标轴有主要网格线，则该值为 **True**。只有主要坐标轴组中的坐标轴才能有网格线。**Boolean** 类型，可读写。

**语法**

**express.HasMajorGridlines**

*express*   一个代表 **Axis** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为 Chart1 中数值轴的主要网格线设置了颜色。*/ function test() {   let axes = Application.Charts.Item("Chart1").Axes(xlValue)      if(axes.HasMajorGridlines){       // Set color to red       axes.MajorGridlines.Border.ColorIndex = 3   } }` |

#### **Axis.HasMinorGridlines**

如果坐标轴有次要网格线，则该值为 **True**。只有主要坐标轴组中的坐标轴才能有网格线。**Boolean** 类型，可读写。

**语法**

**express.HasMinorGridlines**

*express*   一个代表 **Axis** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例对 Chart1 数值轴的次要网格线的颜色进行设置。*/ function test() {   let axes = Application.Charts.Item("Chart1").Axes(xlValue)      if(axes.HasMinorGridlines){       // Set color to green       axes.MinorGridlines.Border.ColorIndex = 4   } }` |

#### **Axis.HasTitle**

如果坐标轴或图表有可见标题，则为 **True**。**Boolean** 类型，可读写。

**语法**

**express.HasTitle**

*express*   一个代表 **Axis** 对象的变量。

**说明**

坐标轴标题由 **AxisTitle** 对象代表。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例为 Chart1 的分类坐标轴添加坐标轴标志。*/ function test() {   Application.Charts.Item("Chart1").Axes().Item(xlCategory).HasTitle = true   Application.Charts.Item("Chart1").Axes().Item(xlCategory).AxisTitle.Text = "July Sales" }` |

#### **Axis.Height**

返回一个 **Double** 值，它代表对象的高度（以磅为单位）。

**语法**

**express.Height**

*express*   一个代表 **Axis** 对象的变量。

#### **Axis.Left**

返回一个 **Double** 值，它代表从对象左边缘到图表区左边缘的距离（以磅为单位）。

**语法**

**express.Left**

*express*   一个代表 **Axis** 对象的变量。

#### **Axis.LogBase**

在使用对数刻度时返回或设置对数的底。可读/写 **Double** 类型。

**语法**

**express.LogBase**

*express*   一个代表 **Axis** 对象的变量。

**说明**

试图将此属性设置为小于或等于零 (0) 的值将导致错误。默认值是 10。

#### **Axis.MajorGridlines**

返回一个 **Gridlines** 对象，该对象表示指定坐标轴的主要网格线。只有主要坐标轴组中的坐标轴才能有网格线。只读。

**语法**

**express.MajorGridlines**

*express*   一个代表 **Axis** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为 Chart1 中数值轴的主要网格线设置了颜色。*/ function test() {   let axes = Application.Charts.Item("Chart1").Axes().Item(xlValue)   if(axes.HasMajorGridlines){       axes.MajorGridlines.Border.ColorIndex = 5    //set color to blue   } }` |

#### **Axis.MajorTickMark**

返回或设置指定坐标轴的主刻度线类型。**XlTickMark** 类型，可读写。

**语法**

**express.MajorTickMark**

*express*   一个代表 **Axis** 对象的变量。

**说明**

| **XlTickMark** 可为以下 **XlTickMark** 常量之一。 |
| ------------------------------------------------- |
| **xlTickMarkInside**                              |
| **xlTickMarkOutside**                             |
| **xlTickMarkCross**                               |
| **xlTickMarkNone**                                |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 Chart1 中的数值轴主要刻度线设置为在轴外侧。*/ Application.Charts.Item("Chart1").Axes().Item(xlValue).MajorTickMark = xlTickMarkOutside` |

#### **Axis.MajorUnit**

返回或设置数值轴的主要单位。**Double** 类型，可读写。

**语法**

**express.MajorUnit**

*express*   一个代表 **Axis** 对象的变量。

**说明**

设置该属性会将 **MajorUnitIsAuto** 属性设置为 False。
使用 **TickMarkSpacing** 属性可设置分类轴上的刻度线间距。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为 Chart1 中的数值轴设置主要单位和次要单位。*/ function test() {   Application.Charts.Item("Chart1").Axes().Item(xlValue).MajorUnit = 100   Application.Charts.Item("Chart1").Axes().Item(xlValue).MinorUnit = 20 }` |

#### **Axis.MajorUnitIsAuto**

如果 ET 计算数值轴的主要单位，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.MajorUnitIsAuto**

*express*   一个代表 **Axis** 对象的变量。

**说明**

设置 **MajorUnit** 属性会将该属性设置为 False。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例自动设置 Chart1 中的数值轴的主要单位和次要单位。*/ function test() {   Application.Charts.Item("Chart1").Axes().Item(xlValue).MajorUnitIsAuto = true   Application.Charts.Item("Chart1").Axes().Item(xlValue).MinorUnitIsAuto = true }` |

#### **Axis.MajorUnitScale**

当 **CategoryType** 属性设置为 **xlTimeScale** 时，返回或设置分类轴主要单位的刻度值。**XlTimeUnit** 类型，可读写。

**语法**

**express.MajorUnitScale**

*express*   一个代表 **Axis** 对象的变量。

**说明**

| **XlTimeUnit** 可为以下 **XlTimeUnit** 常量之一。 |
| ------------------------------------------------- |
| **xlMonths**                                      |
| **xlDays**                                        |
| **xlYears**                                       |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将分类轴设置为使用时间刻度，并设置主要刻度单位和次要刻度单位。*/ function test() {   let axes = Application.Charts.Item(1).Axes().Item(xlCategory)   axes.CategoryType = xlTimeScale   axes.MajorUnit = 5   axes.MajorUnitScale = xlDays   axes.MinorUnit = 1   axes.MinorUnitScale = xlDays }` |

#### **Axis.MaximumScale**

返回或设置数值轴上的最大值。**Double** 类型，可读写。

**语法**

**express.MaximumScale**

*express*   一个代表 **Axis** 对象的变量。

**说明**

设置该属性会将 **MaximumScaleIsAuto** 属性设置为 **False**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例对 Chart1 中的数值轴的最小值和最大值进行设置。*/ function test() {   Application.Charts.Item("Chart1").Axes().Item(xlValue).MinimumScale = 10   Application.Charts.Item("Chart1").Axes().Item(xlValue).MaximumScale = 120 }` |

#### **Axis.MaximumScaleIsAuto**

如果 ET 计算该数值轴的最大值，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.MaximumScaleIsAuto**

*express*   一个代表 **Axis** 对象的变量。

**说明**

设置 **MaximumScale** 属性会将该属性设置为 False。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例自动计算 Chart1 中的数值轴的最小刻度和最大刻度。*/ function test() {   Application.Charts.Item("Chart1").Axes().Item(xlValue).MinimumScaleIsAuto = true   Application.Charts.Item("Chart1").Axes().Item(xlValue).MaximumScaleIsAuto = true }` |

#### **Axis.MinimumScale**

返回或设置数值轴上的最小值。**Double** 类型，可读写。

**语法**

**express.MinimumScale**

*express*   一个代表 **Axis** 对象的变量。

**说明**

设置该属性会将 **MinimumScaleIsAuto** 属性设置为 **False**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例对 Chart1 中的数值轴的最小值和最大值进行设置。*/ function test() { Application.Charts.Item("Chart1").Axes().Item(xlValue).MinimumScale = 10 Application.Charts.Item("Chart1").Axes().Item(xlValue).MaximumScale = 120 }` |

#### **Axis.MinimumScaleIsAuto**

如果 ET 为数值轴计算最小值，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.MinimumScaleIsAuto**

*express*   一个代表 **Axis** 对象的变量。

**说明**

设置 **MinimumScale** 属性会将该属性设置为 False。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例自动计算 Chart1 中的数值轴的最小刻度和最大刻度。*/ function test() {   Application.Charts.Item("Chart1").Axes().Item(xlValue).MinimumScaleIsAuto = true   Application.Charts.Item("Chart1").Axes().Item(xlValue).MaximumScaleIsAuto = true }` |

#### **Axis.MinorGridlines**

返回 **Gridlines** 对象，该对象表示指定坐标轴的次要网格线。只有主要坐标轴组中的坐标轴才能有网格线。只读。

**语法**

**express.MinorGridlines**

*express*   一个代表 **Axis** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例对 Chart1 数值轴的次要网格线的颜色进行设置。*/ function test() {   let minorGridlines  = Application.Charts.Item("Chart1").Axes().Item(xlValue)   if(minorGridlines.HasMinorGridlines){       minorGridlines.MinorGridlines.Border.ColorIndex = 5    //set color to blue   } }` |

#### **Axis.MinorTickMark**

返回或设置指定坐标轴的次要刻度线类型。**XlTickMark** 类型，可读写。

**语法**

**express.MinorTickMark**

*express*   一个代表 **Axis** 对象的变量。

**说明**

| **XlTickMark** 可为以下 **XlTickMark** 常量之一。 |
| ------------------------------------------------- |
| **xlTickMarkInside**                              |
| **xlTickMarkOutside**                             |
| **xlTickMarkCross**                               |
| **xlTickMarkNone**                                |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为 Chart1 数值轴内侧设置次要刻度线。*/ Application.Charts.Item("Chart1").Axes().Item(xlValue).MinorTickMark = xlTickMarkInside` |

#### **Axis.MinorUnit**

返回或设置数值轴上的次要单位。**Double** 类型，可读写。

**语法**

**express.MinorUnit**

*express*   一个代表 **Axis** 对象的变量。

**说明**

设置该属性会将 **MinorUnitIsAuto** 属性设置为 **False**。
使用 **TickMarkSpacing** 属性可设置分类轴上的刻度线间距。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为 Chart1 中的数值轴设置主要单位和次要单位。*/ function test() {   Application.Charts.Item("Chart1").Axes().Item(xlValue).MajorUnit = 100   Application.Charts.Item("Chart1").Axes().Item(xlValue).MinorUnit = 20 }` |

#### **Axis.MinorUnitIsAuto**

如果 ET 为数值轴计算次要单位，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.MinorUnitIsAuto**

*express*   一个代表 **Axis** 对象的变量。

**说明**

设置 **MinorUnit** 属性会将该属性设置为 False。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例自动计算 Chart1 数值轴的主要单位和次要单位。*/ function test() {   Charts.Item("Chart1").Axes().Item(xlValue).MajorUnitIsAuto = true   Charts.Item("Chart1").Axes().Item(xlValue).MinorUnitIsAuto = true }` |

#### **Axis.MinorUnitScale**

返回或设置当 **CategoryType** 属性设置为 **xlTimeScale** 时分类轴次要单位的刻度值。**XlTimeUnit** 类型，可读写。

**语法**

**express.MinorUnitScale**

*express*   一个代表 **Axis** 对象的变量。

**说明**

| **XlTimeUnit** 可为以下 **XlTimeUnit** 常量之一。 |
| ------------------------------------------------- |
| **xlMonths**                                      |
| **xlDays**                                        |
| **xlYears**                                       |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将分类轴设置为使用时间刻度，并设置主要刻度单位和次要刻度单位。*/ function test() {   let axes = Application.Charts.Item(1).Axes().Item(xlCategory)   axes.CategoryType = xlTimeScale   axes.MajorUnit = 5   axes.MajorUnitScale = xlDays   axes.MinorUnit = 1   axes.MinorUnitScale = xlDays }` |

#### **Axis.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **Axis** 对象的变量。

#### **Axis.ReversePlotOrder**

如果 ET 的绘图区数据点的顺序为从后往前，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.ReversePlotOrder**

*express*   一个代表 **Axis** 对象的变量。

**说明**

该属性不能用于雷达图。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 Chart1 数值轴的绘图区数据点的顺序设置为从后往前。*/ Application.Charts.Item("Chart1").Axes().Item(xlValue).ReversePlotOrder = true` |

#### **Axis.ScaleType**

返回或设置数值轴的刻度类型。**XlScaleType** 类型，可读写。

**语法**

**express.ScaleType**

*express*   一个代表 **Axis** 对象的变量。

**说明**

| **XlScaleType** 可为以下 **XlScaleType** 常量之一。 |
| --------------------------------------------------- |
| **xlScaleLinear**                                   |
| **xlScaleLogarithmic**                              |

对数刻度使用以 10 为底的对数。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 Chart1 中的数值轴设置为使用对数刻度。*/ Application.Charts.Item("Chart1").Axes().Item(xlValue).ScaleType = xlScaleLogarithmic` |

#### **Axis.TickLabelPosition**

说明指定坐标轴上刻度线标志的位置。**XlTickLabelPosition** 类型，可读写。

**语法**

**express.TickLabelPosition**

*express*   一个代表 **Axis** 对象的变量。

**说明**

| **XlTickLabelPosition** 可为以下 **XlTickLabelPosition** 常量之一。 |
| ------------------------------------------------------------ |
| **xlTickLabelPositionLow**                                   |
| **xlTickLabelPositionNone**                                  |
| **xlTickLabelPositionHigh**                                  |
| **xlTickLabelPositionNextToAxis**                            |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 Chart1 中分类轴的刻度线标签设置为高位（在图表之上）。*/ Application.Charts.Item("Chart1").Axes().Item(xlCategory).TickLabelPosition = xlTickLabelPositionHigh` |

#### **Axis.TickLabelSpacing**

返回或设置刻度线标签之间的分类数或数据系列数。仅应用于分类轴和系列轴。可以是 1 到 31999 之间的一个数值。**Long** 类型，可读写。

**语法**

**express.TickLabelSpacing**

*express*   一个代表 **Axis** 对象的变量。

**说明**

数值轴上刻度线标签间距总是由 ET 进行计算。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 Chart1 中分类轴上刻度线标签之间的分类数。*/ Application.Charts.Item("Chart1").Axes().Item(xlCategory).TickLabelSpacing = 10` |

#### **Axis.TickLabelSpacingIsAuto**

返回或设置刻度标签间距是否是自动设置的。可读/写 **Boolean** 类型。

**语法**

**express.TickLabelSpacingIsAuto**

*express*   一个代表 **Axis** 对象的变量。

**说明**

**TickLabelSpacing** 属性返回当前刻度标签间距是否是自动设置的。

#### **Axis.TickLabels**

返回一个 **TickLabels** 对象，该对象表示指定坐标轴的刻度线标签。只读。

**语法**

**express.TickLabels**

*express*   一个代表 **Axis** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为 Chart1 中数值轴刻度线标签的字体设置颜色。*/ Application.Charts.Item("Chart1").Axes().Item(xlValue).TickLabels.Font.ColorIndex = 3` |

#### **Axis.TickMarkSpacing**

返回或设置刻度线之间的分类数或数据系列数。仅应用于分类轴和系列轴。可以是 1 到 31999 之间的一个数值。**Long** 类型，可读写。

**语法**

**express.TickMarkSpacing**

*express*   一个代表 **Axis** 对象的变量。

**说明**

可用 **MajorUnit** 属性和 **MinorUnit** 属性设置数值轴上的刻度线间距。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 Chart1 中分类轴上刻度线之间的分类数。*/ Application.Charts.Item("Chart1").Axes().Item(xlCategory).TickMarkSpacing = 10` |

#### **Axis.Top**

返回一个 **Double** 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。

**语法**

**express.Top**

*express*   一个代表 **Axis** 对象的变量。

#### **Axis.Type**

返回一个 **XlAxisType** 值，它代表“坐标轴”类型。

**语法**

**express.Type**

*express*   一个代表 **Axis** 对象的变量。

**说明**

在对散点图使用此属性时，**xlCategory** 将被返回为“坐标轴”类型。

#### **Axis.Width**

返回一个 **Double** 值，它代表对象的宽度（以磅为单位）。

**语法**

**express.Width**

*express*   一个代表 **Axis** 对象的变量。

适用环境：web

适用平台：windows/linux