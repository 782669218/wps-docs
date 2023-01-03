**LegendEntries**



指定的图表图例中所有 **LegendEntry** 对象的集合。

**说明**

每个图例项都有两部分：一部分是该项的文本，它是与该图例项相关联的数据系列或趋势线的名称；另一部分是项标志，它在图表中以直观的方式将图例项以及与之相关联的数据系列或趋势线链接起来。项标志及其相关系列或趋势线的格式属性包含在 **LegendKey** 对象中。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*使用 LegendEntries 方法可返回 LegendEntries 集合。下例在嵌入式图表一中的图例项集合中循环，并更改这些图例项的字体颜色。*/ function test(){ let lgd = Application.Worksheets.Item("sheet1").ChartObjects(1).Chart.Legend for(let i =   1;  i < = lgd.LegendEntries.Count; i++){     lgd.LegendEntries(i).Font.ColorIndex = 5 } }` |

使用 **LegendEntries**( *index*)（其中 *index* 是图例项索引号）可返回一个 **LegendEntry** 对象。不能按名称返回图例项。

图例项的编号代表图例项在图例中的位置。`LegendEntries(1)` 位于图例的顶部；`LegendEntries(LegendEntries.Count)` 位于图例的底部。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例将嵌入的第一个图表中图例顶部的图例项（这通常是第一个数据系列的图例）的文字字体设置为斜体。*/ Application.Worksheets.Item("sheet1").ChartObjects(1).Chart.Legend.LegendEntries(1).Font.Italic = true` |

**方法**

|                                                              | 名称     | 说明                   |
| ------------------------------------------------------------ | -------- | ---------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item** | 从集合中返回一个对象。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回一个 **Long** 值，它代表集合中对象的数量。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |

**成员方法**

#### **LegendEntries.Item**

从集合中返回一个对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **LegendEntries** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**       |
| -------- | ------------- | ------------ | -------------- |
| *Index*  | 必选          | **Variant**  | 对象的索引号。 |

**返回值**

包含在集合中的一个 LegendEntry 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例对 Sheet1 中第一张嵌入式图表的图例（通常是第一个数据系列的图例）顶部的图例项的文本字体进行更改。*/ function test(){ Application.Worksheets.Item("sheet1").ChartObjects(1).Chart.Legend.LegendEntries.(1).Font.Italic = true }` |

**成员属性**

#### **LegendEntries.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **LegendEntries** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ *本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() { 	let myObject = Application.ActiveWorkbook 	if (myObject.Application.Value == "ET") { 		alert("This is an ET Application object.") 	} else { 		alert("This is not an ET Application object.") 	} }` |

#### **LegendEntries.Count**

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

*express*   一个代表 **LegendEntries** 对象的变量。

#### **LegendEntries.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **LegendEntries** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **LegendEntries.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **LegendEntries** 对象的变量。

适用环境：web

适用平台：windows/linux