**Legend**



代表图标中的图例。每个图表只能有一个图例。

**说明**

**Legend** 对象包含一个或多个 **LegendEntry** 对象；每个 **LegendEntry** 对象都包含一个 **LegendKey** 对象。

除非 **HasLegend** 属性是 **True**，否则无法看到图表的图例。如果该属性为 **False**，**Legend** 对象的属性和方法将会失败。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*使用 Legend 属性可返回 Legend 对象。下例将工作表一上嵌入图表一的图例字形设置为加粗。*/ function test(){ Application.Worksheets.Item(1).ChartObjects(1).Chart.Legend.Font.Bold = true }` |

**方法**

|                                                              | 名称              | 说明                                                         |
| ------------------------------------------------------------ | ----------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Clear**         | 清除整个对象。                                               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**        | 删除对象。                                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **LegendEntries** | 返回表示图例中的单个图例项（**LegendEntry** 对象）或图例项集合（**LegendEntries** 对象）的对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Select**        | 选择对象。                                                   |

**属性**

|                                                              | 名称                | 说明                                                         |
| ------------------------------------------------------------ | ------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**     | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Format**          | 返回 **ChartFormat** 对象。只读。                            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Heigh**           | 返回或设置一个 **Double** 值，它代表对象的宽度（以磅为单位）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IncludeInLayout** | 如果在确定图表布局时图例将占用图表布局空间，则为 **True**。默认值是 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Left**            | 返回或设置 **Double** 值，它代表从对象左边缘到工作表的 A 列左边缘或图表上的图表区左边缘的距离（以磅为单位）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**            | 返回一个 **String** 值，它代表对象的名称。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**          | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Position**        | 返回或设置一个 **XlLegendPosition** 值，它代表图表上图例的位置。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Shadow**          | 返回或设置一个 **Boolean** 值，它确定对象是否有阴影。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Top**             | 返回或设置一个 **Double** 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Width**           | 返回或设置一个 **Double** 值，它代表对象的宽度（以磅为单位）。 |

**成员方法**

#### **Legend.Clear**

清除整个对象。

**语法**

**express.Clear()**

*express*   一个代表 **Legend** 对象的变量。

**返回值**

Variant

#### **Legend.Delete**

删除对象。

**语法**

**express.Delete()**

*express*   一个代表 **Legend** 对象的变量。

**返回值**

Variant

#### **Legend.LegendEntries**

返回表示图例中的单个图例项（**LegendEntry** 对象）或图例项集合（**LegendEntries** 对象）的对象。

**语法**

**express.LegendEntries(Index)**

*express*   一个代表 **Legend** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**           |
| -------- | ------------- | ------------ | ------------------ |
| *Index*  | 可选          | **Variant**  | 图例输入项的编号。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 Chart1 上图例输入项一的字体。*/ Application.Charts.Item("Chart1").Legend.LegendEntries(1).Font.Name = "Arial"` |

#### **Legend.Select**

选择对象。

**语法**

**express.Select()**

*express*   一个代表 **Legend** 对象的变量。

**返回值**

Variant

**成员属性**

#### **Legend.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **Legend** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() { 	let myObject = Application.ActiveWorkbook 	if (myObject.Application.Value == "ET") { 		alert("This is an ET Application object.") 	} else { 		alert("This is not an ET Application object.") 	} }` |

#### **Legend.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Legend** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **Legend.Format**

返回 **ChartFormat** 对象。只读。

**语法**

**express.Format**

*express*   一个代表 **Legend** 对象的变量。

#### **Legend.Heigh**

返回或设置一个 **Double** 值，它代表对象的宽度（以磅为单位）。

**语法**

**express.Heigh**

*express*   一个代表 **Legend** 对象的变量。

#### **Legend.IncludeInLayout**

如果在确定图表布局时图例将占用图表布局空间，则为 **True**。默认值是 **True**。**Boolean** 类型，可读写。

**语法**

**express.IncludeInLayout**

*express*   一个代表 **Legend** 对象的变量。

**说明**

此属性对于图表是否处于自动版式模式无影响。如果用户使用**“图表上方”**命令添加标题，则图表将变得较小。如果用户随后删除标题，或者选择一个覆盖标题选项，则图表将变得较大，就好像标题不在图表上一样。

#### **Legend.Left**

返回或设置 **Double** 值，它代表从对象左边缘到工作表的 A 列左边缘或图表上的图表区左边缘的距离（以磅为单位）。

**语法**

**express.Left**

*express*   一个代表 **Legend** 对象的变量。

#### **Legend.Name**

返回一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

*express*   一个代表 **Legend** 对象的变量。

#### **Legend.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **Legend** 对象的变量。

#### **Legend.Position**

返回或设置一个 **XlLegendPosition** 值，它代表图表上图例的位置。

**语法**

**express.Position**

*express*   一个代表 **Legend** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将图表中的图例移到图表的底部。*/ Application.Charts.Item(1).Legend.Position = xlLegendPositionBottom` |

#### **Legend.Shadow**

返回或设置一个 **Boolean** 值，它确定对象是否有阴影。

**语法**

**express.Shadow**

*express*   一个代表 **Legend** 对象的变量。

#### **Legend.Top**

返回或设置一个 **Double** 值，它代表从对象的上边缘到工作表第一行顶部或图表上的图表区顶部的距离（以磅为单位）。

**语法**

**express.Top**

*express*   一个代表 **Legend** 对象的变量。

#### **Legend.Width**

返回或设置一个 **Double** 值，它代表对象的宽度（以磅为单位）。

**语法**

**express.Width**

*express*   一个代表 **Legend** 对象的变量。

适用环境：web

适用平台：windows/linux