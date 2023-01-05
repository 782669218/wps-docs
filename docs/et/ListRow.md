**ListRow**



代表表格中的一行。**ListRow** 对象是 **ListRows** 集合的成员。

**说明**

ListRows 集合包含列表对象中的所有行。**ListRows**

使用 **ListObject** 对象的 **ListRows** 属性可返回一个 **ListRows** 集合。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例给活动工作簿的第一张工作表的默认 ListObject 对象添加新 ListRow 对象。由于未指定位置，因此在表格结束处添加一个新行。*/ function test(){ let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1") let oListRow = wrksht.ListObjects.Item(1).ListRows.Add() }` |

**方法**

|                                                              | 名称       | 说明                                                         |
| ------------------------------------------------------------ | ---------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete** | 删除列表行的单元格并向上移动被删除行下面的任何剩余单元格。即使列表被链接到 SharePoint 网站，您也可以删除该列表中的行。但是，在同步更改之前，SharePoint 网站上的列表将不会进行更新。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Index**       | 返回一个 **Long** 值，它代表 **ListRow** 对象在 **ListRows** 集合中的索引号。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Range**       | 返回一个 **Range** 对象，它代表在以上列表中的指定列表对象所应用的区域。 |

**成员方法**

#### **ListRow.Delete**

删除列表行的单元格并向上移动被删除行下面的任何剩余单元格。即使列表被链接到 SharePoint 网站，您也可以删除该列表中的行。但是，在同步更改之前，SharePoint 网站上的列表将不会进行更新。

**语法**

**express.Delete()**

*express*   一个代表 **ListRow** 对象的变量。

**成员属性**

#### **ListRow.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **ListRow** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() { 	let myObject = Application.ActiveWorkbook 	if (myObject.Application.Value == "ET") { 		alert("This is an ET Application object.") 	} else { 		alert("This is not an ET Application object.") 	} }` |

#### **ListRow.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **ListRow** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **ListRow.Index**

返回一个 **Long** 值，它代表 **ListRow** 对象在 **ListRows** 集合中的索引号。

**语法**

**express.Index**

*express*   一个代表 **ListRow** 对象的变量。

#### **ListRow.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **ListRow** 对象的变量。

#### **ListRow.Range**

返回一个 **Range** 对象，它代表在以上列表中的指定列表对象所应用的区域。

**语法**

**express.Range**

*express*   一个代表 **ListRow** 对象的变量。

适用环境：web

适用平台：windows/linux