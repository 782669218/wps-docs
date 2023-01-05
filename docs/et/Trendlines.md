**Trendlines**



指定的数据系列中所有 **Trendline** 对象的集合。

**说明**

每一个 **Trendline** 对象都代表图表中的趋势线。趋势线显示系列中数据的趋势或方向。

**方法**

|                                                              | 名称     | 说明                   |
| ------------------------------------------------------------ | -------- | ---------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**  | 创建新的趋势线。       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item** | 从集合中返回一个对象。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |

**成员方法**

#### **Trendlines.Add**

创建新的趋势线。

**语法**

**express.Add(Type, Order, Period, Forward, Backward, Intercept, DisplayEquation, DisplayRSquared, Name)**

*express*   一个代表 **Trendlines** 对象的变量。

**参数**

| **名称**          | **必选/可选** | **数据类型**        | **说明**                                                     |
| ----------------- | ------------- | ------------------- | ------------------------------------------------------------ |
| *Type*            | 可选          | **XlTrendlineType** | 趋势线类型。                                                 |
| *Order*           | 可选          | **Variant**         | Variant。如果 Type 为 xlPolynomial。趋势线顺序号。必须为 2 到 6 之间的整数（包括 2 和 6）。 |
| *Period*          | 可选          | **Variant**         | 如果 Type 为 xlMovingAvg。趋势线周期。必须为大于 1，而小于正添加趋势线的数据系列中数据点个数的整数。 |
| *Forward*         | 可选          | **Variant**         | 趋势线向前延伸的周期数目（或散点图中的单位个数）。           |
| *Backward*        | 可选          | **Variant**         | 趋势线向后延伸的周期数目（或散点图中的单位个数）。           |
| *Intercept*       | 可选          | **Variant**         | 趋势线的截距。如果省略此参数，则由回归分析自动设置截距。     |
| *DisplayEquation* | 可选          | **Variant**         | 如果为 True，则在图表中显示趋势线的公式（与 R 平方值显示在同一数据标签中）。默认值为 False。 |
| *DisplayRSquared* | 可选          | **Variant**         | 如果为 True，则在图表中显示趋势线的 R 平方值（与公式显示在同一数据标签中）。默认值为 False。 |
| *Name*            | 可选          | **Variant**         | 作为文本的趋势线名称。如果省略此参数，则由 ET 生成名称。     |

**返回值**

一个代表新趋势线的 Trendline 对象。

**示例**

本示例在 Chart1 中新建线性趋势线。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActiveWorkbook.Charts.Item("Chart1").SeriesCollection(1).Trendlines().Add()` |

#### **Trendlines.Item**

从集合中返回一个对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **Trendlines** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**       |
| -------- | ------------- | ------------ | -------------- |
| *Index*  | 可选          | **Variant**  | 对象的索引号。 |

**返回值**

包含在集合中的一个 Trendline 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {   let rng = Application.Charts.Item("Chart1").SeriesCollection(1).Trendlines().Item(1)   rng.Forward = 5   rng.Backward = 0.5 }` |

**成员属性**

#### **Trendlines.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **Trendlines** 对象的变量。

**示例**

本示例显示一条有关创建 `myObject` 的应用程序的消息。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function TEST() {    let myObject = ActiveWorkbook    if(myObject.Application.Value == "ET"){        MsgBox("This is an ET Application object.")    }    else{        MsgBox("This is not an ET Application object.")    } }` |

适用环境：web

适用平台：windows/linux