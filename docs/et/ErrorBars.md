**ErrorBars**



代表图表数据系列上的误差线。

**说明**

误差线表明图表数据的不确定程度。只有二维图表上的面积、条形、柱形、折线和散点图组中的数据系列可以有误差线。只有散点图组中的数据系列可以有 x 误差线和 y 误差线。此对象不是集合。没有代表单个误差线的对象；或者打开系列中所有数据点的 x 误差线或 y 误差线，或者将其全部关闭。

**ErrorBar** 方法更改误差线的格式和类型。

使用 **ErrorBars** 属性可返回 **ErrorBars** 对象。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例打开嵌入的第一个图表的第一个数据系列的误差线，并设置误差线的尾部样式。*/ function test(){ Application.Worksheets.Item("sheet1").ChartObjects(1).Activate() ActiveChart.SeriesCollection(1).HasErrorBars = true ActiveChart.SeriesCollection(1).ErrorBars.EndStyle = xlNoCap }` |

**方法**

|                                                              | 名称             | 说明                 |
| ------------------------------------------------------------ | ---------------- | -------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ClearFormats** | 清除对象的格式设置。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**       | 删除对象。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Select**       | 选择对象。           |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Border**      | 返回一个 **Border** 对象，它代表对象的边框。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EndStyle**    | 返回或设置误差线的尾端样式。可为以下 **XlEndStyleCap** 常量之一：**xlCap** 或 **xlNoCap**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Format**      | 返回 **ChartFormat** 对象。只读。                            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**        | 返回一个 **String** 值，它代表对象的名称。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |

**成员方法**

#### **ErrorBars.ClearFormats**

清除对象的格式设置。

**语法**

**express.ClearFormats()**

*express*   一个代表 **ErrorBars** 对象的变量。

**返回值**

Variant

#### **ErrorBars.Delete**

删除对象。

**语法**

**express.Delete()**

*express*   一个代表 **ErrorBars** 对象的变量。

**返回值**

Variant

#### **ErrorBars.Select**

选择对象。

**语法**

**express.Select()**

*express*   一个代表 **ErrorBars** 对象的变量。

**返回值**

Variant

**成员属性**

#### **ErrorBars.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **ErrorBars** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() { 	let myObject = Application.ActiveWorkbook 	if (myObject.Application.Value == "ET") { 		alert("This is an ET Application object.") 	} else { 		alert("This is not an ET Application object.") 	} }` |

#### **ErrorBars.Border**

返回一个 **Border** 对象，它代表对象的边框。

**语法**

**express.Border**

*express*   一个代表 **ErrorBars** 对象的变量。

#### **ErrorBars.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **ErrorBars** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **ErrorBars.EndStyle**

返回或设置误差线的尾端样式。可为以下 **XlEndStyleCap** 常量之一：**xlCap** 或 **xlNoCap**。**Long** 类型，可读写。

**语法**

**express.EndStyle**

*express*   一个代表 **ErrorBars** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例对 Chart1 中第一个系列的误差线的尾端样式进行设置。本示例必须在其第一个系列带 Y 误差线的二维折线图上运行。*/ Application.Charts.Item("Chart1").SeriesCollection(1).ErrorBars.EndStyle = xlCap` |

#### **ErrorBars.Format**

返回 **ChartFormat** 对象。只读。

**语法**

**express.Format**

*express*   一个代表 **ErrorBars** 对象的变量。

#### **ErrorBars.Name**

返回一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

*express*   一个代表 **ErrorBars** 对象的变量。

#### **ErrorBars.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **ErrorBars** 对象的变量。

适用环境：web

适用平台：windows/linux