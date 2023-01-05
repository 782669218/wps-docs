**DataLabels**



由数据系列中所有 **DataLabel** 对象组成的集合。

**说明**

每个 **DataLabel** 对象代表一个数据点或趋势线的数据标签。对于没有可定义数据点的数据系列（例如面积图数据系列），**DataLabels** 集合包含单个数据标签。

使用 **DataLabels** 方法可返回单个 **DataLabels** 集合。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例设置第一个图表的第一个数据系列中数据标签的数字格式。*/ function test(){ let myCollection = Application.Charts.Item(1).SeriesCollection(1)     myCollection.HasDataLabels = true     myCollection.DataLabels.NumberFormat = "##.##" }` |

使用 **DataLabels**(*index*)（其中 *index* 为数据标签的索引号）可返回单个 **DataLabel** 对象。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例在第一张工作表上嵌入的第一个图表上，设置第一个数据系列中的第五个数据标签的数字格式。*/ Application.Worksheets.Item(1).ChartObjects(1).Chart      .SeriesCollection(1).DataLabels(5).NumberFormat = "0.000"` |

**方法**

|                                                              | 名称       | 说明                   |
| ------------------------------------------------------------ | ---------- | ---------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete** | 删除对象。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**   | 从集合中返回一个对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Select** | 选择对象。             |

**属性**

|                                                              | 名称                    | 说明                                                         |
| ------------------------------------------------------------ | ----------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**         | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoText**            | 如果对象会根据内容自动生成合适的文字，则为 **True**。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**               | 返回一个 **Long** 值，它代表集合中对象的数量。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**             | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Format**              | 返回 **ChartFormat** 对象。只读。                            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HorizontalAlignment** | 返回或设置一个 **Variant** 值，它代表指定对象的水平对齐方式。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**                | 返回一个 **String** 值，它代表对象的名称。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NumberFormat**        | 返回或设置一个 **String** 值，它代表对象的格式代码。         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NumberFormatLinked**  | 如果指定数字格式指向单元格（以便当单元格的格式更改时数据标签的格式也作相应的改动），则为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NumberFormatLocal**   | 以采用用户语言字符串的形式返回或设置一个 **Variant** 值，它代表对象的格式代码。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Orientation**         | 返回或设置一个 **Variant** 值，它代表文本方向。              |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**              | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Position**            | 返回或设置一个 **XlDataLabelPosition** 值，它代表数据标签的位置。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ReadingOrder**        | 返回或设置指定对象的阅读次序。可为以下常量之一：**xlRTL**（从右到左）、**xlLTR**（从左到右）或 **xlContext**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Separator**           | 设置或返回一个 **Variant** 值，它代表用于图表中数据标签的分隔符。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Shadow**              | 返回或设置一个 **Boolean** 值，它确定对象是否有阴影。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowBubbleSize**      | 如果为 **True**，则在图表中显示数据标签的气泡的大小。如果为 **False**，则隐藏气泡的大小。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowCategoryName**    | 如果为 **True**，则在图表中显示数据标签的分类名称。如果为 **False**，则隐藏该分类的名称。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowLegendKey**       | 如果数据标签图例项标示可见，则为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowPercentage**      | 如果为 **True**，则在图表中显示数据标签的百分比数值。如果为 **False**，则隐藏该百分比数值。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowSeriesName**      | 返回或设置一个 **Boolean** 值，它指明图表上数据标签的系列名称显示方式。如果为 **True**，则显示系列名称。如果为 **False**，则隐藏系列名称。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **VerticalAlignment**   | 返回或设置一个 **Variant** 值，它代表指定对象的垂直对齐方式。 |

**成员方法**

#### **DataLabels.Delete**

删除对象。

**语法**

**express.Delete()**

*express*   一个代表 **DataLabels** 对象的变量。

**返回值**

Variant

#### **DataLabels.Item**

从集合中返回一个对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **DataLabels** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**       |
| -------- | ------------- | ------------ | -------------- |
| *Index*  | 必选          | **Variant**  | 对象的索引号。 |

**返回值**

包含在集合中的一个 DataLabel 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例对数据标签的数字格式进行设置，该数据标签是第一张工作表中第一张嵌入式图表的第一个数据系列的第五个数据标签。*/ function test(){ Application.Worksheets.Item(1).ChartObjects(1).Chart      .SeriesCollection(1).DataLabels.Item(5).NumberFormat = "0.000" }` |

#### **DataLabels.Select**

选择对象。

**语法**

**express.Select()**

*express*   一个代表 **DataLabels** 对象的变量。

**返回值**

Variant

**成员属性**

#### **DataLabels.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **DataLabels** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() { 	let myObject = Application.ActiveWorkbook 	if (myObject.Application.Value == "ET") { 		alert("This is an ET Application object.") 	} else { 		alert("This is not an ET Application object.") 	} }` |

#### **DataLabels.AutoText**

如果对象会根据内容自动生成合适的文字，则为 **True**。可读/写 **Boolean** 类型。

**语法**

**express.AutoText**

*express*   一个代表 **DataLabels** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例对 Chart1 的第一个数据序列的数据标签进行设置，自动生成合适的文字*/ Application.Charts.Item("Chart1").SeriesCollection(1).DataLabels.AutoText = true` |

| 注释                                                         |
| ------------------------------------------------------------ |
| 如果在**“即时窗口”**中运行 `?ActiveChart.SeriesCollection(1).DataLabels.AutoText`，您将收到以下结果： ET 2003：不返回任何内容。 ET 2007 及更高版本：只有在所有 **DataLabels** 的 **AutoText** 都为 **True** 时返回 **True**，如果所有 **DataLabels** 的 **AutoText** 都为 **False**，或者部分 **DataLabels** 的 **AutoText** 为 **False**，那么返回 **False**。 |

#### **DataLabels.Count**

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

*express*   一个代表 **DataLabels** 对象的变量。

#### **DataLabels.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **DataLabels** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **DataLabels.Format**

返回 **ChartFormat** 对象。只读。

**语法**

**express.Format**

*express*   一个代表 **DataLabels** 对象的变量。

#### **DataLabels.HorizontalAlignment**

返回或设置一个 **Variant** 值，它代表指定对象的水平对齐方式。

**语法**

**express.HorizontalAlignment**

*express*   一个代表 **DataLabels** 对象的变量。

**说明**

此属性的值可设为以下常量之一：

| **xlCenter**      |
| ----------------- |
| **xlDistributed** |
| **xlJustify**     |
| **xlLeft**        |
| **xlRight**       |

某些常量可能不可用，这取决于所选择或安装的语言支持（例如，美国英语）。

#### **DataLabels.Name**

返回一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

*express*   一个代表 **DataLabels** 对象的变量。

#### **DataLabels.NumberFormat**

返回或设置一个 **String** 值，它代表对象的格式代码。

**语法**

**express.NumberFormat**

*express*   一个代表 **DataLabels** 对象的变量。

**说明**

格式代码与**“设置单元格格式”**对话框中的**“格式代码”**选项是同一个字符串。**Format** 函数使用的格式代码字符串与 **NumberFormat** 和 **NumberFormatLocal** 属性使用的格式代码字符串不同。

#### **DataLabels.NumberFormatLinked**

如果指定数字格式指向单元格（以便当单元格的格式更改时数据标签的格式也作相应的改动），则为 **True**。**Boolean** 类型，可读写。

**语法**

**express.NumberFormatLinked**

*express*   一个代表 **DataLabels** 对象的变量。

#### **DataLabels.NumberFormatLocal**

以采用用户语言字符串的形式返回或设置一个 **Variant** 值，它代表对象的格式代码。

**语法**

**express.NumberFormatLocal**

*express*   一个代表 **DataLabels** 对象的变量。

**说明**

**Format** 函数使用的格式代码字符串与 **NumberFormat** 和 **NumberFormatLocal** 属性使用的格式代码字符串不同。

#### **DataLabels.Orientation**

返回或设置一个 **Variant** 值，它代表文本方向。

**语法**

**express.Orientation**

*express*   一个代表 **DataLabels** 对象的变量。

**说明**

此属性的值可设为

#### **DataLabels.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **DataLabels** 对象的变量。

#### **DataLabels.Position**

返回或设置一个 **XlDataLabelPosition** 值，它代表数据标签的位置。

**语法**

**express.Position**

*express*   一个代表 **DataLabels** 对象的变量。

#### **DataLabels.ReadingOrder**

返回或设置指定对象的阅读次序。可为以下常量之一：**xlRTL**（从右到左）、**xlLTR**（从左到右）或 **xlContext**。**Long** 类型，可读写。

**语法**

**express.ReadingOrder**

*express*   一个代表 **DataLabels** 对象的变量。

#### **DataLabels.Separator**

设置或返回一个 **Variant** 值，它代表用于图表中数据标签的分隔符。可读/写。

**语法**

**express.Separator**

*express*   一个代表 **DataLabels** 对象的变量。

**说明**

如果使用字符串，则用一个字符串作为分隔符。如果使用 **xlDataLabelSeparatorDefault** (= 1)，则用默认数据标签分隔符，即逗号或换行符，具体取决于数据标签。

当返回值“1”时，表明用户未更改默认的分隔符逗号“,”。您也可通过传递值“1”将分隔符更改回默认值。

以程序方式访问数据标签之前必须先激活图表，否则将出现运行错误。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例将第一个图表上的第一个系列的数据标签分隔符设置为分号。此示例假定活动工作表上存在图表。*/ Application.ActiveSheet.ChartObjects(1).Chart.SeriesCollection(1)          .DataLabels.Separator = ";"` |

#### **DataLabels.Shadow**

返回或设置一个 **Boolean** 值，它确定对象是否有阴影。

**语法**

**express.Shadow**

*express*   一个代表 **DataLabels** 对象的变量。

#### **DataLabels.ShowBubbleSize**

如果为 **True**，则在图表中显示数据标签的气泡的大小。如果为 **False**，则隐藏气泡的大小。**Boolean** 类型，可读写。

**语法**

**express.ShowBubbleSize**

*express*   一个代表 **DataLabels** 对象的变量。

**说明**

以程序方式访问数据标签之前，必须先激活图表，否则将出现运行错误。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例显示第一个图表上第一个系列的数据标签的气泡大小。此示例假定活动工作表上存在图表。*/ function test(){  Application.ActiveSheet.ChartObjects(1).Activate()     ActiveChart.SeriesCollection(1)          .DataLabels.ShowBubbleSize = true }` |

#### **DataLabels.ShowCategoryName**

如果为 **True**，则在图表中显示数据标签的分类名称。如果为 **False**，则隐藏该分类的名称。**Boolean** 类型，可读写。

**语法**

**express.ShowCategoryName**

*express*   一个代表 **DataLabels** 对象的变量。

**说明**

以程序方式访问数据标签之前，必须先激活图表，否则将出现运行错误。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例显示第一个图表上第一个系列的数据标签的分类名称。此示例假定活动工作表上存在图表。*/ function test(){   Application.ActiveSheet.ChartObjects(1).Activate()     ActiveChart.SeriesCollection(1)          .DataLabels.ShowCategoryName = true }` |

#### **DataLabels.ShowLegendKey**

如果数据标签图例项标示可见，则为 **True**。**Boolean** 类型，可读写。

**语法**

**express.ShowLegendKey**

*express*   一个代表 **DataLabels** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例设置 Chart1 中系列一的数据标签，以显示数值和图例标示*/ function test(){ let myDataLabel = Application.Charts.Item("Chart1").SeriesCollection(1).DataLabels     myDataLabel.ShowLegendKey = true     myDataLabel.Type = xlShowValue }` |

#### **DataLabels.ShowPercentage**

如果为 **True**，则在图表中显示数据标签的百分比数值。如果为 **False**，则隐藏该百分比数值。**Boolean** 类型，可读写。

**语法**

**express.ShowPercentage**

*express*   一个代表 **DataLabels** 对象的变量。

**说明**

以程序方式访问数据标签之前，必须先激活图表，否则将出现运行错误。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例显示了第一个图表上第一个数据系列的数据标签的百分比数值。此示例假定活动工作表上存在图表。*/ function test(){   Application.ActiveSheet.ChartObjects(1).Activate()     ActiveChart.SeriesCollection(1)          .DataLabels.ShowPercentage = true }` |

#### **DataLabels.ShowSeriesName**

返回或设置一个 **Boolean** 值，它指明图表上数据标签的系列名称显示方式。如果为 **True**，则显示系列名称。如果为 **False**，则隐藏系列名称。可读写。

**语法**

**express.ShowSeriesName**

*express*   一个代表 **DataLabels** 对象的变量。

**说明**

以程序方式访问数据标签之前，必须先激活图表，否则将出现运行错误。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例显示了第一个图表上第一个数据系列的数据标签的系列名称。此示例假定活动工作表上存在图表。*/ function test(){ Application.ActiveSheet.ChartObjects(1).Activate     ActiveChart.SeriesCollection(1)          .DataLabels.ShowSeriesName = true }` |

#### **DataLabels.VerticalAlignment**

返回或设置一个 **Variant** 值，它代表指定对象的垂直对齐方式。

**语法**

**express.VerticalAlignment**

*express*   一个代表 **DataLabels** 对象的变量。

**说明**

此属性的值可设为以下常量之一：

| **xlBottom**      |
| ----------------- |
| **xlCenter**      |
| **xlDistributed** |
| **xlJustify**     |
| **xlTop**         |

适用环境：web

适用平台：windows/linux