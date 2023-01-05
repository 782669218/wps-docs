**LeaderLines**



代表图表的引导线。引导线将数据标签连接到数据点。

**说明**

该对象不是一个集合；没有表示单个引导线的对象。

此对象只适用于饼图。

使用 **LeaderLines** 属性可返回 **LeaderLines** 对象。下例图表一上的数据系列一添加数据标签和蓝色引导线。如果看不到引导线，此示例代码将会失败。在这种情况下，您可以手动将一个数据标签从饼图上拖离，以便显示出引导线。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){ let chartObjects2 = Application.Worksheets.Item(1).ChartObjects(1).Chart.SeriesCollection(1) chartObjects2.HasDataLabels = true chartObjects2.DataLabels.Position = xlLabelPositionBestFit chartObjects2.HasLeaderLines = true chartObjects2.LeaderLines.Border.ColorIndex = 5 }` |

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
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |

**成员方法**

#### **LeaderLines.Delete**

删除对象。

**语法**

**express.Delete()**

*express*   一个代表 **LeaderLines** 对象的变量。

#### **LeaderLines.Select**

选择对象。

**语法**

**express.Select()**

*express*   一个代表 **LeaderLines** 对象的变量。

**成员属性**

#### **LeaderLines.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **LeaderLines** 对象的变量。

**示例**

*本示例显示一条有关创建 myObject 的应用程序的消息。*/
function test() {
let myObject = Application.ActiveWorkbook
if (myObject.Application.Value == "ET") {
alert("This is an ET Application object.")
} else {
alert("This is not an ET Application object.")
}
}

#### **LeaderLines.Border**

返回一个 **Border** 对象，它代表对象的边框。

**语法**

**express.Border**

*express*   一个代表 **LeaderLines** 对象的变量。

#### **LeaderLines.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **LeaderLines** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **LeaderLines.Format**

返回 **ChartFormat** 对象。只读。

**语法**

**express.Format**

*express*   一个代表 **LeaderLines** 对象的变量。

#### **LeaderLines.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **LeaderLines** 对象的变量。

适用环境：web

适用平台：windows/linux