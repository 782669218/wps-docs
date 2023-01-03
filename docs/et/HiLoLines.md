**HiLoLines**



代表图表组中的高低点连线。

**说明**

高低点连线连接图表组内每一分类中的最高数据点和最底数据点。只有二维折线图组可以有高低点连线。此对象不是集合。没有代表单个高低点连线的对象；或者打开图表组中所有数据点的高低点连线，或者将其全部关闭。

如果 **HasHiLoLines** 属性是 **False**，**HiLoLines** 对象的大多数属性都会被禁用。

使用 **HiLoLines** 属性可返回 **HiLoLines** 对象。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例使用 HasHiLowLines 属性将 HiLowLines 添加到工作表一上的嵌入图表一（该图表必须是折线图）。之后，该示例会将高低点连线设为蓝色。*/ function test(){ Application.Worksheets.Item(1).ChartObjects(1).Activate ActiveChart.ChartGroups(1).HasHiLoLines = true ActiveChart.ChartGroups(1).HiLoLines.Border.Color = (0, 0, 255) }` |

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

#### **HiLoLines.Delete**

删除对象。

**语法**

**express.Delete()**

*express*   一个代表 **HiLoLines** 对象的变量。

**返回值**

Variant

#### **HiLoLines.Select**

选择对象。

**语法**

**express.Select()**

*express*   一个代表 **HiLoLines** 对象的变量。

**返回值**

Variant

**成员属性**

#### **HiLoLines.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **HiLoLines** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() { 	let myObject = Application.ActiveWorkbook 	if (myObject.Application.Value == "ET") { 		alert("This is an ET Application object.") 	} else { 		alert("This is not an ET Application object.") 	} }` |

#### **HiLoLines.Border**

返回一个 **Border** 对象，它代表对象的边框。

**语法**

**express.Border**

*express*   一个代表 **HiLoLines** 对象的变量。

#### **HiLoLines.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **HiLoLines** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **HiLoLines.Format**

返回 **ChartFormat** 对象。只读。

**语法**

**express.Format**

*express*   一个代表 **HiLoLines** 对象的变量。

#### **HiLoLines.Name**

返回一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

*express*   一个代表 **HiLoLines** 对象的变量。

#### **HiLoLines.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **HiLoLines** 对象的变量。

适用环境：web

适用平台：windows/linux