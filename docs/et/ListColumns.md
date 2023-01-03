**ListColumns**



指定的 **ListObject** 对象中所有 **ListColumn** 对象的集合。

**说明**

每个 **ListColumn** 对象都代表表格中的一列。

| ![img]()注释                                       |
| -------------------------------------------------- |
| 该列的名称会自动生成。在添加完该列后可更改其名称。 |

使用 ListObject 对象的 **ListColumns** 属性可返回 **ListColumns** 集合。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例给工作簿的第一张工作表的默认 ListObject 对象添加一个新列。由于未指定位置，因此在最右边添加一个新列。*/ let myNewColumn = Application.Worksheets.Item(1).ListObjects.Item(1).ListColumns.Add()` |

**方法**

|                                                              | 名称    | 说明                   |
| ------------------------------------------------------------ | ------- | ---------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add** | 向列表对象中添加新列。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回一个 **Integer** 值，它代表集合中对象的数量。            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Item**        | 从集合中返回一个对象。                                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |

**成员方法**

#### **ListColumns.Add**

向列表对象中添加新列。

**语法**

**express.Add(Position)**

*express*   一个代表 **ListColumns** 对象的变量。

**参数**

| **名称**   | **必选/可选** | **数据类型** | **说明**                                                     |
| ---------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Position* | 可选          | **Variant**  | Integer 类型。从 1 开始指定新列的相对位置。以前位于此位置的列则向后移动。 |

**返回值**

一个代表新列的 ListColumn 对象。

**说明**

如果不指定 *Position*，就在最右边添加一列，并自动为该列生成名称。该列的名称可在添加后更改。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例给工作簿的第一张工作表的默认 ListObject 对象添加一个新列。由于未指定位置，因此在最右边添加一个新列。*/ function test(){ let myNewColumn = Application.Worksheets.Item(1).ListObjects.Item(1).ListColumns.Add() }` |

**成员属性**

#### **ListColumns.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **ListColumns** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() { 	let myObject = Application.ActiveWorkbook 	if (myObject.Application.Value == "ET") { 		alert("This is an ET Application object.") 	} else { 		alert("This is not an ET Application object.") 	} }` |

#### **ListColumns.Count**

返回一个 **Integer** 值，它代表集合中对象的数量。

**语法**

**express.Count**

*express*   一个代表 **ListColumns** 对象的变量。

#### **ListColumns.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **ListColumns** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **ListColumns.Item**

从集合中返回一个对象。

**语法**

**express.Item**

*express*   一个代表 **ListColumns** 对象的变量。

#### **ListColumns.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **ListColumns** 对象的变量。

适用环境：web

适用平台：windows/linux