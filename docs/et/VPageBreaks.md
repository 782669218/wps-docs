**VPageBreaks**



打印区域中垂直分页符的集合。

**说明**

每一个垂直分页符都由一个 **VPageBreak** 对象代表。

当 [Application](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application) 属性、**Count** 属性、**Creator** 属性、**Item** 属性、**Parent** 属性或 **Add** 方法与 **VPageBreaks** 属性一起使用时：

- 对于自动打印区域，**VPageBreaks** 属性只应用于打印区域内的分页符。
- 对于同一区域中的用户定义的打印区域，**VPageBreaks** 属性应用于所有的分页符。

使用 **VPageBreaks** 属性可返回 **VPageBreaks** 集合。使用 **Add** 方法可添加一个垂直分页符。

如果添加的分页符不和打印区域交叠，则新添加的 **VPageBreak** 对象将不会出现在打印区域的 **VPageBreaks** 集合中。如果重新调整打印区域或重新定义打印区域，将会改变集合的内容。

下例给活动单元格左边添加一个垂直分页符。

| 示例代码复制                                          |
| ----------------------------------------------------- |
| `Application.ActiveSheet.VPageBreaks.Add(ActiveCell)` |

**方法**

|                                                              | 名称     | 说明                                                         |
| ------------------------------------------------------------ | -------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**  | 添加垂直分页符。返回一个代表新垂直分页符的 **VPageBreak** 对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item** | 从集合中返回一个对象。                                       |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回一个 **Long** 值，它代表集合中对象的数量。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |

**成员方法**

#### **VPageBreaks.Add**

添加垂直分页符。返回一个代表新垂直分页符的 **VPageBreak** 对象。

**语法**

**express.Add(Before )**

*express*   一个代表 **VPageBreaks** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                        |
| -------- | ------------- | ------------ | ----------------------------------------------- |
| *Before* | 必选          | **Object**   | 一个 Range 对象。新的分页符将添加到该区域左侧。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例在单元格 F25 上方添加水平分页符，在其左侧添加垂直分页符*/ function test() {     Application.Worksheets.Item(1).HPageBreaks.Add(Worksheets.Item(1).Range("F25"))     Application.Worksheets.Item(1).VPageBreaks.Add(Worksheets.Item(1).Range("F25")) }` |

#### **VPageBreaks.Item**

从集合中返回一个对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **VPageBreaks** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**       |
| -------- | ------------- | ------------ | -------------- |
| *Index*  | 必选          | **Long**     | 对象的索引号。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 此示例更改第一个垂直分页符的位置*/ Application.Worksheets.Item(1).VPageBreaks.Item(1).Location = Worksheets.Item(1).Range("e5")` |

**成员属性**

#### **VPageBreaks.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **VPageBreaks** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例显示一条有关创建 myObject 的应用程序的消息*/ function test() {     let myObject = Application.ActiveWorkbook     if(myObject.Application.Value == "ET") {         alert("This is an ET Application object.")     }     else {         alert("This is not an ET Application object.")     } }` |

#### **VPageBreaks.Count**

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

*express*   一个代表 **VPageBreaks** 对象的变量。

#### **VPageBreaks.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **VPageBreaks** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **VPageBreaks.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **VPageBreaks** 对象的变量。

适用环境：web

适用平台：windows/linux