**ListRows**



指定的 **ListObject** 对象中 **ListRow** 对象的集合。

**说明**

每一个 **ListRow** 对象都代表表格中的一行。

使用 **ListObject** 对象的 **ListRows** 属性可返回 **ListRows** 集合。下例给工作簿的第一张工作表的默认 **ListObject** 对象添加新行。由于未指定位置，因此在表格结束处添加一个新行。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `let myNewRow = Application.Worksheets.Item(1).ListObjects.Item(1).ListRows.Add()` |

**方法**

|                                                              | 名称    | 说明                                             |
| ------------------------------------------------------------ | ------- | ------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add** | 向指定的 **ListObject** 所代表的表格中添加新行。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回一个 **Integer** 值，它代表集合中对象的数量。            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Item**        | 从集合中返回一个对象。                                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |

**成员方法**

#### **ListRows.Add**

向指定的 **ListObject** 所代表的表格中添加新行。

**语法**

**express.Add(Position, AlwaysInsert)**

*express*   一个代表 **ListRows** 对象的变量。

**参数**

| **名称**       | **必选/可选** | **数据类型** | **说明**                                                     |
| -------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Position*     | 可选          | **Variant**  | Integer 类型。指定新行的相对位置。                           |
| *AlwaysInsert* | 可选          | **Variant**  | Boolean。指定在插入新行时，是否始终移动表格中最后一行下面的单元格中的数据，而不考虑表格下面的行是否为空行。如果为 True，则表格下面的单元格将下移一行。如果为 False，并且表格下面的行为空行，则表格将扩展以占用该空行而不移动它下面的单元格；但如果表格下面的行包含数据，则在插入新行时，这些单元格将下移。 |

**返回值**

一个代表新行的 ListRow 对象。

**说明**

如果不指定 *Position*，新行将添加到底部。如果不指定 *AlwaysInsert*，表格下面的单元格将下移一行（与指定 **True** 的作用相同）。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例向工作簿的第一张工作表中的默认 ListObject 对象添加新行。由于未指定位置，新行将添加到列表的底部。*/ let myNewColumn = Application.ActiveWorkbook.Worksheets.Item(1).ListObject.Item(1).ListRows.Add()` |

**成员属性**

#### **ListRows.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **ListRows** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() { 	let myObject = Application.ActiveWorkbook 	if (myObject.Application.Value == "ET") { 		alert("This is an ET Application object.") 	} else { 		alert("This is not an ET Application object.") 	} }` |

#### **ListRows.Count**

返回一个 **Integer** 值，它代表集合中对象的数量。

**语法**

**express.Count**

*express*   一个代表 **ListRows** 对象的变量。

#### **ListRows.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **ListRows** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **ListRows.Item**

从集合中返回一个对象。

**语法**

**express.Item**

*express*   一个代表 **ListRows** 对象的变量。

#### **ListRows.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **ListRows** 对象的变量。

适用环境：web

适用平台：windows/linux