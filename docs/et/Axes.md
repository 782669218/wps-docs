**Axes**



指定图表中所有 **Axis** 对象的集合。

**方法**

|                                                              | 名称     | 说明                                       |
| ------------------------------------------------------------ | -------- | ------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item** | 从 **Axes** 集合中返回一个 **Axis** 对象。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回一个 **Long** 值，它代表集合中对象的数量。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |

**成员方法**

#### **Axes.Item**

从 **Axes** 集合中返回一个 **Axis** 对象。

**语法**

**express.Item(Type, AxisGroup)**

*express*   一个代表 **Axes** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型**    | **说明**     |
| ----------- | ------------- | --------------- | ------------ |
| *Type*      | 必选          | **XlAxisType**  | 坐标轴类型。 |
| *AxisGroup* | 可选          | **XlAxisGroup** | 坐标轴。     |

**说明**

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为 Chart1 中分类轴设置标题文本。*/ function test() {   let myChart = Application.Charts.Item("Chart1").Axes.Item(xlCategory)   myChart.HasTitle = true   myChart.AxisTitle.Caption = "1994" }` |

**成员属性**

#### **Axes.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个代表指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **Axes** 对象的变量。

#### **Axes.Count**

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

*express*   一个代表 **Axes** 对象的变量。

#### **Axes.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Axes** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **Axes.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **Axes** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示包含 myAxis 的图表的名称。*/ function test(){     let myAxis = Application.Charts.Item(1).Axes(xlValue)     alert(myAxis.Parent.Name) }` |

适用环境：web

适用平台：windows/linux