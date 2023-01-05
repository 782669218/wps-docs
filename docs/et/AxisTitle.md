**AxisTitle**



代表图表坐标轴标题。

**说明**

使用 **AxisTitle** 属性可返回 **AxisTitle** 对象。
只有当坐标轴的 **HasTitle** 属性为 **True** 时，**AxisTitle** 对象才存在，从而才能使用该对象。

示例

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例激活第一个嵌入式图表，设置其数值轴标题文本，将其字体设为 10 磅的“Bookman”，并将单词“millions”设为倾斜。*/ function test() {     Application.Worksheets.Item("sheet1").ChartObjects(1).Activate()     let axes = Application.ActiveChart.Axes(xlValue)     axes.HasTitle = true      let axistitle = axes.AxisTitle     axistitle.Caption = "Revenue (millions)"     axistitle.Font.Name = "bookman"     axistitle.Font.Size = 10     axistitle.Characters(10, 8).Font.Italic = true }` |

**方法**

|                                                              | 名称       | 说明       |
| ------------------------------------------------------------ | ---------- | ---------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete** | 删除对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Select** | 选择对象。 |

**属性**

|                                                              | 名称                    | 说明                                                         |
| ------------------------------------------------------------ | ----------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**         | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Caption**             | 返回或设置一个 **String** 值，它代表坐标轴标题文本。         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Characters**          | 返回 **Characters** 对象，它代表对象文本内某个区域的字符。使用 **Characters** 对象可为文本字符串内的字符设置格式。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**             | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Format**              | 返回 **ChartFormat** 对象。只读。                            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Formula**             | 获取或设置一个 **String** 值，该值以英语使用 A1 样式表示法来表示对象的公式。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FormulaLocal**        | 获取或设置一个 **String** 值，该值以用户的语言使用 A1 样式表示法来表示对象的公式。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FormulaR1C1**         | 获取或设置一个 **String** 值，该值以英语使用 R1C1 样式表示法来表示对象的公式。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FormulaR1C1Local**    | 获取或设置一个 **String** 值，该值以用户的语言使用 R1C1 样式表示法来表示对象的公式。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Height**              | 返回对象的高度（以磅为单位）。只读。                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HorizontalAlignment** | 返回或设置一个 **Variant** 值，它代表指定对象的水平对齐方式。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IncludeInLayout**     | 如果在确定图表布局时轴标题将占用图表布局空间，则为 **True**。默认值是 **True**。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Left**                | 返回或设置 **Double** 值，它代表从对象左边缘到工作表的 A 列左边缘或图表上的图表区左边缘的距离（以磅为单位）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**                | 返回一个 **String** 值，它代表对象的名称。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Orientation**         | 返回或设置一个 **Variant** 值，它代表文本方向。              |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**              | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Position**            | 返回或设置图表上轴标题的位置。可读/写 **XlChartElementPosition** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ReadingOrder**        | 返回或设置指定对象的阅读次序。可为以下常量之一：**xlRTL**（从右到左）、**xlLTR**（从左到右）或 **xlContext**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Shadow**              | 返回或设置一个 **Boolean** 值，它确定对象是否有阴影。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Text**                | 返回或设置指定对象中的文本。**String** 型，可读写。          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Top**                 | 返回或设置一个 **Double** 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **VerticalAlignment**   | 返回或设置一个 **Variant** 值，它代表指定对象的垂直对齐方式。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Width**               | 返回对象的宽度（以磅为单位）。只读。                         |

**成员方法**

#### **AxisTitle.Delete**

删除对象。

**语法**

**express.Delete()**

*express*   一个代表 **AxisTitle** 对象的变量。

#### **AxisTitle.Select**

选择对象。

**语法**

**express.Select()**

*express*   一个代表 **AxisTitle** 对象的变量。

**说明**

返回值一个代表所选对象的 Variant 值。

**成员属性**

#### **AxisTitle.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **AxisTitle** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() {   if(Application.ActiveWorkbook.Application.Value == "ET" ){       alert("This is an ET Application object.")   }   else{       alert("This is not an ET Application object.")   } }` |

#### **AxisTitle.Caption**

返回或设置一个 **String** 值，它代表坐标轴标题文本。

**语法**

**express.Caption**

*express*   一个代表 **AxisTitle** 对象的变量。

#### **AxisTitle.Characters**

返回 **Characters** 对象，它代表对象文本内某个区域的字符。使用 **Characters** 对象可为文本字符串内的字符设置格式。

**语法**

**express.Characters**

*express*   一个代表 **AxisTitle** 对象的变量。

**说明**

参数

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Start*  | 可选          | **Variant**  | 要返回的第一个字符。如果此参数是 1 或被省略，则此属性返回一个以第一个字符为开头的字符区域。 |
| *Length* | 可选          | **Variant**  | 要返回的字符数。如果省略此参数，则此属性返回字符串的后半部分（*Start* 字符之后的所有字符）。 |

**Characters** 对象不是集合。

#### **AxisTitle.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **AxisTitle** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **AxisTitle.Format**

返回 **ChartFormat** 对象。只读。

**语法**

**express.Format**

*express*   一个代表 **AxisTitle** 对象的变量。

#### **AxisTitle.Formula**

获取或设置一个 **String** 值，该值以英语使用 A1 样式表示法来表示对象的公式。可读/写。

**语法**

**express.Formula**

*express*   一个代表 **AxisTitle** 对象的变量。

**说明**

有关 A1 样式表示法的详细信息，请参阅操作方法：使用 A1 表示法引用单元格和区域。

#### **AxisTitle.FormulaLocal**

获取或设置一个 **String** 值，该值以用户的语言使用 A1 样式表示法来表示对象的公式。可读/写。

**语法**

**express.FormulaLocal**

*express*   一个代表 **AxisTitle** 对象的变量。

**说明**

有关 A1 样式表示法的详细信息，请参阅操作方法：使用 A1 表示法引用单元格和区域。

#### **AxisTitle.FormulaR1C1**

获取或设置一个 **String** 值，该值以英语使用 R1C1 样式表示法来表示对象的公式。可读/写。

**语法**

**express.FormulaR1C1**

*express*   一个代表 **AxisTitle** 对象的变量。

#### **AxisTitle.FormulaR1C1Local**

获取或设置一个 **String** 值，该值以用户的语言使用 R1C1 样式表示法来表示对象的公式。可读/写。

**语法**

**express.FormulaR1C1Local**

*express*   一个代表 **AxisTitle** 对象的变量。

#### **AxisTitle.Height**

返回对象的高度（以磅为单位）。只读。

**语法**

**express.Height**

*express*   一个代表 **AxisTitle** 对象的变量。

#### **AxisTitle.HorizontalAlignment**

返回或设置一个 **Variant** 值，它代表指定对象的水平对齐方式。

**语法**

**express.HorizontalAlignment**

*express*   一个代表 **AxisTitle** 对象的变量。

**说明**

此属性的值可设为以下常量之一：

| **xlCenter**      |
| ----------------- |
| **xlDistributed** |
| **xlJustify**     |
| **xlLeft**        |
| **xlRight**       |

某些常量可能不可用，这取决于所选择或安装的语言支持（例如，美国英语）。

#### **AxisTitle.IncludeInLayout**

如果在确定图表布局时轴标题将占用图表布局空间，则为 **True**。默认值是 **True**。可读/写 **Boolean** 类型。

**语法**

**express.IncludeInLayout**

*express*   一个代表 **AxisTitle** 对象的变量。

#### **AxisTitle.Left**

返回或设置 **Double** 值，它代表从对象左边缘到工作表的 A 列左边缘或图表上的图表区左边缘的距离（以磅为单位）。

**语法**

**express.Left**

*express*   一个代表 **AxisTitle** 对象的变量。

#### **AxisTitle.Name**

返回一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

*express*   一个代表 **AxisTitle** 对象的变量。

#### **AxisTitle.Orientation**

返回或设置一个 **Variant** 值，它代表文本方向。

**语法**

**express.Orientation**

*express*   一个代表 **AxisTitle** 对象的变量。

#### **AxisTitle.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **AxisTitle** 对象的变量。

#### **AxisTitle.Position**

返回或设置图表上轴标题的位置。可读/写 **XlChartElementPosition** 类型。

**语法**

**express.Position**

*express*   一个代表 **AxisTitle** 对象的变量。

#### **AxisTitle.ReadingOrder**

返回或设置指定对象的阅读次序。可为以下常量之一：**xlRTL**（从右到左）、**xlLTR**（从左到右）或 **xlContext**。**Long** 类型，可读写。

**语法**

**express.ReadingOrder**

*express*   一个代表 **AxisTitle** 对象的变量。

#### **AxisTitle.Shadow**

返回或设置一个 **Boolean** 值，它确定对象是否有阴影。

**语法**

**express.Shadow**

*express*   一个代表 **AxisTitle** 对象的变量。

#### **AxisTitle.Text**

返回或设置指定对象中的文本。**String** 型，可读写。

**语法**

**express.Text**

*express*   一个代表 **AxisTitle** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例设置 Chart1 中的分类轴标题文本。*/ function test() {   Application.Charts.Item("Chart1").Axes(xlCategory).HasTitle = true   Application.Charts.Item("Chart1").Axes(xlCategory).AxisTitle.Text = "Month" }` |

#### **AxisTitle.Top**

返回或设置一个 **Double** 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。

**语法**

**express.Top**

*express*   一个代表 **AxisTitle** 对象的变量。

#### **AxisTitle.VerticalAlignment**

返回或设置一个 **Variant** 值，它代表指定对象的垂直对齐方式。

**语法**

**express.VerticalAlignment**

*express*   一个代表 **AxisTitle** 对象的变量。

**说明**

此属性的值可设为以下常量之一：

| **xlBottom**      |
| ----------------- |
| **xlCenter**      |
| **xlDistributed** |
| **xlJustify**     |
| **xlTop**         |

#### **AxisTitle.Width**

返回对象的宽度（以磅为单位）。只读。

**语法**

**express.Width**

*express*   一个代表 **AxisTitle** 对象的变量。

**说明**

返回值
Double

适用环境：web

适用平台：windows/linux