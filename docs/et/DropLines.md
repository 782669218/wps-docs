**DropLines**



代表图表组中的垂直线。

**说明**

垂直线将图表中的数据点与 x 轴连接起来。只有折线图和面积图组可以有垂直线。此对象不是集合。没有代表单个垂直线的对象；或者打开图表组中所有数据点的垂直线，或者将其全部关闭。

如果 **HasDropLines** 属性为 **False**，**DropLines** 对象的绝大部分属性将被禁用。

使用 **DropLines** 属性可返回 **DropLines** 对象。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例打开嵌入的第一个图表的第一个图表组的垂直线，并将垂直线的颜色设置为红色。*/ function test(){ Application.Worksheets.Item("Sheet1").ChartObjects(1).Activate() ActiveChart.ChartGroups(1).HasDropLines = true ActiveChart.ChartGroups(1).DropLines.Border.ColorIndex = 3 }` |

**方法**

|                                                              | 名称       | 说明       |
| ------------------------------------------------------------ | ---------- | ---------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete** | 删除对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Select** | 选择对象。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Border**      | 返回一个 **Border** 对象，它代表对象的边框。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Format**      | 返回 **ChartFormat** 对象。只读。                            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**        | 返回一个 **String** 值，它代表对象的名称。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |

**成员方法**

#### **DropLines.Delete**

删除对象。

**语法**

**express.Delete()**

*express*   一个代表 **DropLines** 对象的变量。

**返回值**

Variant

#### **DropLines.Select**

选择对象。

**语法**

**express.Select()**

*express*   一个代表 **DropLines** 对象的变量。

**返回值**

Variant

**成员属性**

#### **DropLines.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **DropLines** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() { 	let myObject = Application.ActiveWorkbook 	if (myObject.Application.Value == "ET") { 		alert("This is an ET Application object.") 	} else { 		alert("This is not an ET Application object.") 	} }` |

#### **DropLines.Border**

返回一个 **Border** 对象，它代表对象的边框。

**语法**

**express.Border**

*express*   一个代表 **DropLines** 对象的变量。

#### **DropLines.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **DropLines** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **DropLines.Format**

返回 **ChartFormat** 对象。只读。

**语法**

**express.Format**

*express*   一个代表 **DropLines** 对象的变量。

#### **DropLines.Name**

返回一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

*express*   一个代表 **DropLines** 对象的变量。

#### **DropLines.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **DropLines** 对象的变量。

适用环境：web

适用平台：windows/linux