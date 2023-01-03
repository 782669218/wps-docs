**ListObjects**



工作表上所有 **ListObject** 对象的集合。每个 **ListObject** 对象都代表工作表中的一个表格。

**说明**

使用 Worksheet 对象的 **ListObjects** 属性可返回 **ListObjects** 集合。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例创建一个新 ListObjects 集合，该集合代表工作表中所有的表格.*/ let myWorksheetLists = Application.Worksheets.Item(1).ListObjects` |

**方法**

|                                                              | 名称    | 说明               |
| ------------------------------------------------------------ | ------- | ------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add** | 创建新的列表对象。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回一个 **Integer** 值，它代表集合中对象的数量。            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Item**        | 从集合中返回一个对象。                                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |

**成员方法**

#### **ListObjects.Add**

创建新的列表对象。

**语法**

**express.Add(SourceType, Source, LinkSource, TableStyleName, Destination)**

*express*   一个代表 **ListObjects** 对象的变量。

**参数**

| **名称**         | **必选/可选** | **数据类型**               | **说明**                                                     |
| ---------------- | ------------- | -------------------------- | ------------------------------------------------------------ |
| *SourceType*     | 可选          | **XlListObjectSourceType** | 表示查询的来源类型。                                         |
| *Source*         | 可选          | **Variant**                | 当 SourceType 为 xlSrcRange 时，它是代表数据源的 Range 对象。如果省略该参数，Source 的默认值将是列表区域检测代码返回的区域。当 SourceType 为 xlSrcExternal 时，它是一个 String 值数组，用于指定与数据源的连接，包含以下元素： 0 ― SharePoint 网站的 URL 1 ― ListName 2 ― ViewGUID |
| *LinkSource*     | 可选          | **Variant**                | Boolean 型。表示外部数据源是否链接到 ListObject 对象。如果 SourceType 为 xlSrcExternal，则默认值为 True。如果 SourceType 为 xlSrcRange，则此参数无效，如果不省略它，将会返回一个错误。 |
| *TableStyleName* | 可选          | **Variant**                | 一个 XlYesNoGuess 常量，它指示正在导入的数据是否有列标签。如果 Source 没有标题，ET 将自动生成标题。 |
| *Destination*    | 可选          | **Variant**                | 一个 Range 对象，作用是将一个单元格引用指定为新列表对象左上角的目标区域。如果 Range 对象引用多个单元格，则会产生错误。当 SourceType 设置为 xlSrcExternal 时，必须指定 Destination 参数。如果 SourceType 设置为 xlSrcRange，则忽略 Destination 参数。目标区域所在的工作表必须是包含 expression 所指定的 ListObjects 集合的工作表。新列的插入位置将是 Destination 以适应新列表。因此，现有数据不会被覆盖。 |

**返回值**

一个 ListObject 对象，它代表新的列表对象。

**说明**

当列表有标题时，第一行单元格将转换为 **Text**（如果还未被设为文本）。转换将基于单元格的可见文本。这意味着，如果有一个日期值，该日期值的 **Date** 格式随区域设置的更改而更改，则对列表的转换可能产生不同的结果，具体取决于当前的系统区域设置。而且，如果标题行中有两个单元格包含相同的可见文本，则会追加递增的 **Integer** 以使每个列标题唯一。

**示例**

下例在默认 **ListObject** 集合中添加新的 **ListObjects** 对象（该对象基于 Microsoft SharePoint Foundation 网站的数据），并将列表放在工作簿中第一个工作表的 A1 单元格中。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 以下代码示例假设您会将 `strServerName` 和 `strListGUID` 变量替换为有效的服务器名称和列表 GUID。此外，服务器名称后面必须是“/_vti_bin”，否则本示例将不运行。 |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `let objListObject = Application.ActiveWorkbook.Worksheets.Item(1).ListObjects.Add(xlSrcExternal, Array(strServerName, strListName, strListGUID), true, xlGuess, Range("A10"))` |

**成员属性**

#### **ListObjects.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **ListObjects** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test() { 	let myObject = Application.ActiveWorkbook 	if (myObject.Application.Value == "ET") { 		alert("This is an ET Application object.") 	} else { 		alert("This is not an ET Application object.") 	} }` |

#### **ListObjects.Count**

返回一个 **Integer** 值，它代表集合中对象的数量。

**语法**

**express.Count**

*express*   一个代表 **ListObjects** 对象的变量。

#### **ListObjects.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **ListObjects** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **ListObjects.Item**

从集合中返回一个对象。

**语法**

**express.Item**

*express*   一个代表 **ListObjects** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示活动工作簿的 Sheet1 上的默认列表对象的名称。*/ function test(){ let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1") let oListObj = wrksht.ListObjects.Item(1).Name }` |

#### **ListObjects.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **ListObjects** 对象的变量。

适用环境：web

适用平台：windows/linux