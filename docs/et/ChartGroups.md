**ChartGroups**



代表图表中用同一格式绘制的一个或多个数据系列。

**说明**

**ChartGroups** 集合是指定图表中的所有 **ChartGroup** 对象的集合。每张图表包含一个或多个图表组，每个图表组包含一个或多个数据系列，每个数据系列包含一个或多个数据点。例如，单张图表可能既包含折线图图表组（其中包含所有用折线图格式绘制的数据系列），也包含条形图图表组（其中包含所有用条形图格式绘制的数据系列）。

使用 **ChartGroups** 方法可返回 **ChartGroups** 集合。以下示例显示工作表 1 上嵌入图表 1 中图表组的个数。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `alert(Application.Worksheets.Item(1).ChartObjects(1).Chart.ChartGroups().Count)` |

使用 **ChartGroups**(*index*)（其中 *index* 是图表组的索引号）可以返回单个 **ChartGroup** 对象。以下示例向图表工作表 1 上的图表组 1 中添加垂直线。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.Charts.Item(1).ChartGroups(1).HasDropLines = true` |

如果图表已被激活，可使用 **ActiveChart**：

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.Charts.Item(1).Activate() Application.ActiveChart.ChartGroups(1).HasDropLines = true` |

因为当特定图表组所用的图表格式更改时，该图表组的索引号可能会更改，所以使用命名图表组快捷方法之一来返回特定的图表组会更加容易。**PieGroups** 方法返回图表中饼图图表组的集合，**LineGroups** 方法返回图表中折线图图表组的集合，依此类推。这些方法中的每一个都可以与索引号配合使用以返回单个 **ChartGroup** 对象，或不指定索引号而返回 **ChartGroups** 集合。

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

#### **ChartGroups.Item**

从集合中返回一个对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **ChartGroups** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**       |
| -------- | ------------- | ------------ | -------------- |
| *Index*  | 必选          | **Variant**  | 对象的索引号。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例为第一个图表工作表中的第一个图表组添加垂直线。*/ Application.Charts.Item(1).ChartGroups.Item(1).HasDropLines = true` |

**成员属性**

#### **ChartGroups.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **ChartGroups** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() {     let myObject = Application.ActiveWorkbook     if(myObject.Application.Value == "ET") {         alert("This is an ET Application object.")     }     else {         alert("This is not an ET Application object.")     } }` |

#### **ChartGroups.Count**

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

*express*   一个代表 **ChartGroups** 对象的变量。

#### **ChartGroups.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **ChartGroups** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **ChartGroups.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **ChartGroups** 对象的变量。

适用环境：web

适用平台：windows/linux