**ListColumn**



代表表格中的一列。

**说明**

**ListColumn** 对象是 **ListColumns** 集合的成员。**ListColumns** 集合包含表格中的所有列（**ListObject** 对象）。

使用 **ListObject** 对象的 ListColumns 属性可返回一个 **ListColumns** 集合。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例给活动工作簿的第一个工作表的默认 ListObject 对象添加新 ListColumn 对象。由于未指定位置，因此在最右边添加一个新列。*/ function test(){  let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")     let objListCol = wrksht.ListObjects.Item(1).ListColumns.Add() }` |

**方法**

|                                                              | 名称       | 说明                 |
| ------------------------------------------------------------ | ---------- | -------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete** | 在列表中删除数据列。 |

**属性**

|                                                              | 名称                  | 说明                                                         |
| ------------------------------------------------------------ | --------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**       | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**           | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DataBodyRange**     | 返回一个 **Range** 对象，该对象为列的数据部分的大小。只读。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Index**             | 返回一个 **Long** 值，它代表 **ListColumn** 对象在 **ListColumns** 集合中的索引号。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**              | 返回或设置一个 **String** 值，它代表列表列的名称。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**            | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Range**             | 返回一个 **Range** 对象，它代表在以上列表中的指定列表对象所应用的区域。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Total**             | 返回 **ListColumn** 对象的总计行。只读。                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TotalsCalculation** | 基于 **XlTotalsCalculation** 枚举的值，确定列表列的“总计”行中的计算类型。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **XPath**             | 返回一个 **XPath** 对象，它代表映射到指定 **Range** 对象的元素的 XPath。该区域的上下文确定操作是否成功，或返回空对象。只读。 |

**成员方法**

#### **ListColumn.Delete**

在列表中删除数据列。

**语法**

**express.Delete()**

*express*   一个代表 **ListColumn** 对象的变量。

**说明**

此方法不从工作表中删除列。如果列表链接到 Microsoft SharePoint Foundation 网站，则无法从服务器中删除该列，并产生一个错误。

**成员属性**

#### **ListColumn.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **ListColumn** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() { 	let myObject = Application.ActiveWorkbook 	if (myObject.Application.Value == "ET") { 		alert("This is an ET Application object.") 	} else { 		alert("This is not an ET Application object.") 	} }` |

#### **ListColumn.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **ListColumn** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **ListColumn.DataBodyRange**

返回一个 **Range** 对象，该对象为列的数据部分的大小。只读。

**语法**

**express.DataBodyRange**

*express*   一个代表 **ListColumn** 对象的变量。

**说明**

返回的对象不包含标题单元格和总计单元格。

#### **ListColumn.Index**

返回一个 **Long** 值，它代表 **ListColumn** 对象在 **ListColumns** 集合中的索引号。

**语法**

**express.Index**

*express*   一个代表 **ListColumn** 对象的变量。

#### **ListColumn.Name**

返回或设置一个 **String** 值，它代表列表列的名称。

**语法**

**express.Name**

*express*   一个代表 **ListColumn** 对象的变量。

**说明**

该属性值还被用作列表列的显示名称。此名称必须在列表中唯一。

#### **ListColumn.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **ListColumn** 对象的变量。

#### **ListColumn.Range**

返回一个 **Range** 对象，它代表在以上列表中的指定列表对象所应用的区域。

**语法**

**express.Range**

*express*   一个代表 **ListColumn** 对象的变量。

#### **ListColumn.Total**

返回 **ListColumn** 对象的总计行。只读。

**语法**

**express.Total**

*express*   一个代表 **ListColumn** 对象的变量。

#### **ListColumn.TotalsCalculation**

基于 **XlTotalsCalculation** 枚举的值，确定列表列的“总计”行中的计算类型。可读写。

**语法**

**express.TotalsCalculation**

*express*   一个代表 **ListColumn** 对象的变量。

**说明**

| **XlTotalsCalculation** 可为以下 **XlTotalsCalculation** 常量之一。 |
| ------------------------------------------------------------ |
| **xlTotalsCalculationNone**                                  |
| **xlTotalsCalculationSum**                                   |
| **xlTotalsCalculationAverage**                               |
| **xlTotalsCalculationCount**                                 |
| **xlTotalsCalculationCountNums**                             |
| **xlTotalsCalculationMin**                                   |
| **xlTotalsCalculationStdDev**                                |
| **xlTotalsCalculationVar**                                   |
| **xlTotalsCalculationMax**                                   |

无须显示“总计”行即可设置此属性。此属性没有固定的“默认”值。随着对其他列的添加或删除，ET 会更改该属性的状态。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActiveSheet.ListColumns.Item(1).TotalsCalculation=xlTotalsCalculationSum ` |

#### **ListColumn.XPath**

返回一个 **XPath** 对象，它代表映射到指定 **Range** 对象的元素的 XPath。该区域的上下文确定操作是否成功，或返回空对象。只读。

**语法**

**express.XPath**

*express*   一个代表 **ListColumn** 对象的变量。

**说明**

当 **XPath** 属性包含的区域满足以下条件时，该属性有效：

区域中仅有一个单元格。
如果区域中包含两个或更多个单元格，那么必须满足以下条件之一：
如果单元格包含 XPath 信息，区域内的所有单元格都必须包含 XPath 信息（也即，每个单元格与一个或多个数据映射关联），所有单元格还必须具有相同的 XPath 内容（也即，每个单元格向同一组数据映射提供数据）。
所有单元格都不包含 XPath 信息。
区域内没有不连续的区域。

![img]()注释表格的标题和总计行被视为包含 XPath 信息。 

任何不满足上述条件的区域将返回运行时错误。

如果区域选择有效，但未映射任何单元格，ET 将返回一个 **XPath** 对象，这样您就能通过访问 **SetValue** 方法来创建映射。

适用环境：web

适用平台：windows/linux