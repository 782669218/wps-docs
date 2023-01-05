**ChartTitle**



代表图表标题。

**说明**

使用 **ChartTitle** 属性可返回 **ChartTitle** 对象。

只有图表的 **HasTitle** 属性为 **True** 时，**ChartTitle** 对象才存在，从而才能使用该对象。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例向工作表 Sheet1 上嵌入的第一个图表添加标题。*/ function test() { 	let myChart = application.Worksheets.Item("Sheet1").ChartObjects(1).Chart 		myChart.HasTitle = true 		myChart.ChartTitle.Text = "February Sales" }` |

**方法**

|                                                              | 名称       | 说明       |
| ------------------------------------------------------------ | ---------- | ---------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete** | 删除对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Select** | 选择对象。 |

**属性**

|                                                              | 名称                    | 说明                                                         |
| ------------------------------------------------------------ | ----------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**         | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Caption**             | 返回或设置一个 **String** 值，它代表图表标题文本的方向。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Characters**          | 返回 **Characters** 对象，它代表对象文本内某个区域的字符。使用 **Characters** 对象可为文本字符串内的字符设置格式。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**             | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Format**              | 返回 **ChartFormat** 对象。只读。                            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Formula**             | 获取或设置一个 **String** 值，该值以英语使用 A1 样式表示法来表示对象的公式。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FormulaLocal**        | 获取或设置一个 **String** 值，该值以用户的语言使用 A1 样式表示法来表示对象的公式。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FormulaR1C1**         | 获取或设置一个 **String** 值，该值以英语使用 R1C1 样式表示法来表示对象的公式。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FormulaR1C1Local**    | 获取或设置一个 **String** 值，该值以用户的语言使用 R1C1 样式表示法来表示对象的公式。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Height**              | 返回对象的高度（以磅为单位）。只读。                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HorizontalAlignment** | 返回或设置一个 **Variant** 值，它代表指定对象的水平对齐方式。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IncludeInLayout**     | 如果在确定图表布局时图表标题将占用图表布局空间，则为 **True**。默认值是 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Left**                | 返回或设置 **Double** 值，它代表从对象左边缘到工作表的 A 列左边缘或图表上的图表区左边缘的距离（以磅为单位）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**                | 返回一个 **String** 值，它代表对象的名称。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Orientation**         | 返回或设置一个 **Variant** 值，它代表文本方向。              |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**              | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Position**            | 返回或设置图表上图表标题的位置。可读/写 **XlChartElementPosition** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ReadingOrder**        | 返回或设置指定对象的阅读次序。可为以下常量之一：**xlRTL**（从右到左）、**xlLTR**（从左到右）或 **xlContext**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Shadow**              | 返回或设置一个 **Boolean** 值，它确定对象是否有阴影。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Text**                | 返回或设置指定对象中的文本。**String** 型，可读写。          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Top**                 | 返回或设置一个 **Double** 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **VerticalAlignment**   | 返回或设置一个 **Variant** 值，它代表指定对象的垂直对齐方式。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Width**               | 返回对象的宽度（以磅为单位）。只读。                         |

**成员方法**

#### **ChartTitle.Delete**

删除对象。

**语法**

**express.Delete()**

*express*   一个代表 **ChartTitle** 对象的变量。

**返回值**

Variant

#### **ChartTitle.Select**

选择对象。

**语法**

**express.Select()**

*express*   一个代表 **ChartTitle** 对象的变量。

**返回值**

Variant

**成员属性**

#### **ChartTitle.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **ChartTitle** 对象的变量。

**说明**

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/  function test() { 	let myObject = ActiveWorkbook 		if (myObject.Application.Value == "ET") { 			MsgBox("This is an ET Application object.") 		} else { 			MsgBox("This is not an ET Application object.") 		} }` |

#### **ChartTitle.Caption**

返回或设置一个 **String** 值，它代表图表标题文本的方向。

**语法**

**express.Caption**

*express*   一个代表 **ChartTitle** 对象的变量。

#### **ChartTitle.Characters**

返回 **Characters** 对象，它代表对象文本内某个区域的字符。使用 **Characters** 对象可为文本字符串内的字符设置格式。

**语法**

**express.Characters**

*express*   一个代表 **ChartTitle** 对象的变量。

**说明**

**Characters** 对象不是集合。

#### **ChartTitle.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **ChartTitle** 对象的变量。

#### **ChartTitle.Format**

返回 **ChartFormat** 对象。只读。

**语法**

**express.Format**

*express*   一个代表 **ChartTitle** 对象的变量。

**说明**

**ChartFormat** 对象包含图表区的线条、填充、效果和文本格式。

#### **ChartTitle.Formula**

获取或设置一个 **String** 值，该值以英语使用 A1 样式表示法来表示对象的公式。可读/写。

**语法**

**express.Formula**

*express*   一个代表 **ChartTitle** 对象的变量。

#### **ChartTitle.FormulaLocal**

获取或设置一个 **String** 值，该值以用户的语言使用 A1 样式表示法来表示对象的公式。可读/写。

**语法**

**express.FormulaLocal**

*express*   一个代表 **ChartTitle** 对象的变量。

#### **ChartTitle.FormulaR1C1**

获取或设置一个 **String** 值，该值以英语使用 R1C1 样式表示法来表示对象的公式。可读/写。

**语法**

**express.FormulaR1C1**

*express*   一个代表 **ChartTitle** 对象的变量。

#### **ChartTitle.FormulaR1C1Local**

获取或设置一个 **String** 值，该值以用户的语言使用 R1C1 样式表示法来表示对象的公式。可读/写。

**语法**

**express.FormulaR1C1Local**

*express*   一个代表 **ChartTitle** 对象的变量。

#### **ChartTitle.Height**

返回对象的高度（以磅为单位）。只读。

**语法**

**express.Height**

*express*   一个代表 **ChartTitle** 对象的变量。

#### **ChartTitle.HorizontalAlignment**

返回或设置一个 **Variant** 值，它代表指定对象的水平对齐方式。

**语法**

**express.HorizontalAlignment**

*express*   一个代表 **ChartTitle** 对象的变量。

**说明**

此属性的值可设为以下常量之一：

| **xlCenter**      |
| ----------------- |
| **xlDistributed** |
| **xlJustify**     |
| **xlLeft**        |
| **xlRight**       |

某些常量可能不可用，这取决于所选择或安装的语言支持（例如，美国英语）。

#### **ChartTitle.IncludeInLayout**

如果在确定图表布局时图表标题将占用图表布局空间，则为 **True**。默认值是 **True**。**Boolean** 类型，可读写。

**语法**

**express.IncludeInLayout**

*express*   一个代表 **ChartTitle** 对象的变量。

**说明**

此属性对于图表是否处于自动版式模式无影响。如果用户使用**“图表上方”**命令添加标题，则图表将变得较小。如果用户随后删除标题，或者选择一个覆盖标题选项，则图表将变得较大，就好像标题不在图表上一样

#### **ChartTitle.Left**

返回或设置 **Double** 值，它代表从对象左边缘到工作表的 A 列左边缘或图表上的图表区左边缘的距离（以磅为单位）。

**语法**

**express.Left**

*express*   一个代表 **ChartTitle** 对象的变量。

#### **ChartTitle.Name**

返回一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

*express*   一个代表 **ChartTitle** 对象的变量。

#### **ChartTitle.Orientation**

返回或设置一个 **Variant** 值，它代表文本方向。

**语法**

**express.Orientation**

*express*   一个代表 **ChartTitle** 对象的变量。

**说明**

此属性的值可设为

#### **ChartTitle.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **ChartTitle** 对象的变量。

#### **ChartTitle.Position**

返回或设置图表上图表标题的位置。可读/写 **XlChartElementPosition** 类型。

**语法**

**express.Position**

*express*   一个代表 **ChartTitle** 对象的变量。

#### **ChartTitle.ReadingOrder**

返回或设置指定对象的阅读次序。可为以下常量之一：**xlRTL**（从右到左）、**xlLTR**（从左到右）或 **xlContext**。**Long** 类型，可读写。

**语法**

**express.ReadingOrder**

*express*   一个代表 **ChartTitle** 对象的变量。

#### **ChartTitle.Shadow**

返回或设置一个 **Boolean** 值，它确定对象是否有阴影。

**语法**

**express.Shadow**

*express*   一个代表 **ChartTitle** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例向 myChart 的标题添加阴影。*/ function test() { 	Charts.Item("Chart1").ChartTitle.Shadow = true }` |

#### **ChartTitle.Text**

返回或设置指定对象中的文本。**String** 型，可读写。

**语法**

**express.Text**

*express*   一个代表 **ChartTitle** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例设置 Chart1 的图表标题文本。*/  function test() { let myChart = Charts.Item("Chart1") myChart.HasTitle = true myChart.ChartTitle.Text = "First Quarter Sales" }` |

#### **ChartTitle.Top**

返回或设置一个 **Double** 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。

**语法**

**express.Top**

*express*   一个代表 **ChartTitle** 对象的变量。

#### **ChartTitle.VerticalAlignment**

返回或设置一个 **Variant** 值，它代表指定对象的垂直对齐方式。

**语法**

**express.VerticalAlignment**

*express*   一个代表 **ChartTitle** 对象的变量。

**说明**

此属性的值可设为以下常量之一：

| **xlBottom**      |
| ----------------- |
| **xlCenter**      |
| **xlDistributed** |
| **xlJustify**     |
| **xlTop**         |

#### **ChartTitle.Width**

返回对象的宽度（以磅为单位）。只读。

**语法**

**express.Width**

*express*   一个代表 **ChartTitle** 对象的变量。

适用环境：web

适用平台：windows/linux