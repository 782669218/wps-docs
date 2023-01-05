**SeriesCollection**



指定的图表或图表组中所有 **Series** 对象的集合。

**说明**

使用 **SeriesCollection** 方法可返回 **SeriesCollection** 集合。

**方法**

|                                                              | 名称          | 说明                                                   |
| ------------------------------------------------------------ | ------------- | ------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**       | 向 **SeriesCollection** 集合添加一个或多个新数据系列。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Extend**    | 向已存在的系列集合中添加新的数据点。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**      | 从集合中返回一个对象。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **NewSeries** | 创建新系列。返回代表该新系列的 **Series** 对象。       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Paste**     | 将剪贴板中的数据粘贴到指定的系列集合中。               |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回一个 **Long** 值，它代表集合中对象的数量。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |

**成员方法**

#### **SeriesCollection.Add**

向 **SeriesCollection** 集合添加一个或多个新数据系列。

**语法**

**express.Add(Source, Rowcol, SeriesLabels, CategoryLabels, Replace)**

*express*   一个代表 **SeriesCollection** 对象的变量。

**参数**

| **名称**         | **必选/可选** | **数据类型** | **说明**                                                     |
| ---------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Source*         | 必选          | **Variant**  | 作为 Range 对象的新数据。                                    |
| *Rowcol*         | 可选          | **XlRowCol** | 指定新值是位于指定区域的行中还是列中。                       |
| *SeriesLabels*   | 可选          | **Variant**  | 如果第一行或第一列包含数据系列的名称，则为 True。如果第一行或第一列包含数据系列的第一个数据点，则为 False。如果省略此参数，ET 就尝试根据第一行或第一列中的内容确定系列名称的位置。 |
| *CategoryLabels* | 可选          | **Variant**  | 如果第一行或第一列包含分类标签的名称，则为 True。如果第一行或第一列包含数据系列的第一个数据点，则为 False。如果省略此参数，ET 就尝试根据第一行或第一列中的内容确定分类标签的位置。 |
| *Replace*        | 可选          | **Variant**  | 如果 CategoryLabels 为 True 且 Replace 为 True，那么指定的分类将替换当前系列中存在的分类。如果 Replace 为 False，现有的分类将保留。默认值为 False。 |

**返回值**

一个代表新数据系列的 Series 对象。

**说明**

实际上，此方法不会像“对象浏览器”中所述的那样返回 **Series** 对象。此方法对数据透视表无效。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例在 Chart1 中新建一个数据系列。新系列的数据源是 Sheet1 上的单元格区域 B1:B10。 Application.Charts.Item("Chart1").SeriesCollection().Add(Application.ActiveWorkbook.Worksheets.Item("Sheet1").Range("B1:B10"))  //本示例在工作表 Sheet1 上的嵌入式图表中新建一个数据系列。 function test() {     Application.Worksheets.Item("Sheet1").ChartObjects(1).Activate()     Application.ActiveChart.SeriesCollection().Add(Application.Worksheets.Item("Sheet1").Range("B1:B10")) }` |

#### **SeriesCollection.Extend**

向已存在的系列集合中添加新的数据点。

**语法**

**express.Extend(Source, Rowcol, CategoryLabels)**

*express*   一个代表 **SeriesCollection** 对象的变量。

**参数**

| **名称**         | **必选/可选** | **数据类型** | **说明**                                                     |
| ---------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Source*         | 必选          | **Variant**  | 要以 Range 对象形式添加到 SeriesCollection 对象的新数据。    |
| *Rowcol*         | 可选          | **Variant**  | 指定新值是位于给定区域源的行中还是列中。可以是下列 XlRowCol 常量之一：xlRows 或 xlColumns。如果省略本参数，ET 将依据选定区域的大小和方向，或数组的维度确定数值的位置。 |
| *CategoryLabels* | 可选          | **Variant**  | 如果为 True，则第一行或列包含类别标签的名称。如果为 False，则第一行或列包含系列的第一个数据点。如果省略该参数，则 ET 将依据第一行或列的内容确定类别标签的位置。 |

**返回值**

Variant

**说明**

此方法对数据透视图无效。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例将工作表 Sheet1 中单元格区域 B1:B6 中的数据添加到 Chart1 中，以延伸其中的系列。 Application.Charts.Item("Chart1").SeriesCollection().Extend(Application.Worksheets.Item("Sheet1").Range("B1:B6"))` |

#### **SeriesCollection.Item**

从集合中返回一个对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **SeriesCollection** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**             |
| -------- | ------------- | ------------ | -------------------- |
| *Index*  | 必选          | **Variant**  | 对象的名称或索引号。 |

**返回值**

包含在集合中的一个 Series 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例对 Chart1 中的趋势线向前和向后延伸的单位数进行设置。本示例应在包含单个带趋势线系列的二维柱形图上运行。 function test() {     let trendlines = Application.Charts.Item("Chart1").SeriesCollection().Item(1).Trendlines().Item(1)     trendlines.Forward = 5     trendlines.Backward = .5 }` |

#### **SeriesCollection.NewSeries**

创建新系列。返回代表该新系列的 **Series** 对象。

**语法**

**express.NewSeries()**

*express*   一个代表 **SeriesCollection** 对象的变量。

**返回值**

Series

**说明**

本方法对数据透视图无效。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例向第一个图表中添加新系列。 let ns = Application.Charts.Item(1).SeriesCollection().NewSeries()` |

#### **SeriesCollection.Paste**

将剪贴板中的数据粘贴到指定的系列集合中。

**语法**

**express.Paste(Rowcol, SeriesLabels, CategoryLabels, Replace, NewSeries)**

*express*   一个代表 **SeriesCollection** 对象的变量。

**参数**

| **名称**         | **必选/可选** | **数据类型** | **说明**                                                     |
| ---------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Rowcol*         | 可选          | **XlRowCol** | 指定与特定数据系列对应的值是在行中还是在列中。               |
| *SeriesLabels*   | 可选          | **Variant**  | 如果为 True，则用每一行的第一列（或每一列的第一行）中的单元格内容作为该行（或列）中数据系列的名称。如果为 False，则用每一行的第一列（或每一列的第一行）中的单元格内容作为数据系列的第一个数据点。默认值为 False。 |
| *CategoryLabels* | 可选          | **Variant**  | 如果为 True，则用选定区域的第一行（或第一列）中的内容作为图表的分类。如果为 False，则用选定区域的第一行（或第一列）中的内容作为图表的第一个数据系列。默认值为 False。 |
| *Replace*        | 可选          | **Variant**  | 如果为 True，则在用所复制区域的信息替换现有的分类时应用分类。如果为 False，则插入新分类而不替换任何原有分类。默认值为 True。 |
| *NewSeries*      | 可选          | **Variant**  | 如果为 True，则将数据作为新系列粘贴。如果为 False，则将数据作为现有系列的新数据点粘贴。默认值为 True。 |

**返回值**

Variant

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例将剪贴板中的数据粘贴到 Chart1 上的新数据系列中。 function test() {     Application.Worksheets.Item("Sheet1").Range("C1:C5").Copy()     Application.Charts.Item("Chart1").SeriesCollection().Paste() }` |

**成员属性**

#### **SeriesCollection.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **SeriesCollection** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例显示一条有关创建 myObject 的应用程序的消息。 function test() {     if(Application.ActiveWorkbook.Application.Value == "ET")      {     alert("This is an ET Application object.")     }     else     {     alert("This is not an ET Application object.")     } }` |

#### **SeriesCollection.Count**

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

*express*   一个代表 **SeriesCollection** 对象的变量。

#### **SeriesCollection.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **SeriesCollection** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **SeriesCollection.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **SeriesCollection** 对象的变量。

适用环境：web

适用平台：windows/linux