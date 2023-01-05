**ODBCErrors**



**ODBCError** 对象的集合。

**说明**

每一个 **ODBCError** 对象都代表一个由最近的 ODBC 查询返回的错误。如果指定的 ODBC 查询运行时没有出现错误，则 **ODBCErrors** 集合为空。集合中错误的索引次序与 ODBC 数据源生成它们时的次序相同。您不能给该集合添加成员。

**方法**

|                                                              | 名称     | 说明                   |
| ------------------------------------------------------------ | -------- | ---------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item** | 从集合中返回一个对象。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回一个 **Long** 值，它代表集合中对象的数量。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |

**成员方法**

#### **ODBCErrors.Item**

从集合中返回一个对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **ODBCErrors** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**       |
| -------- | ------------- | ------------ | -------------- |
| *Index*  | 必选          | **Long**     | 对象的索引号。 |

**返回值**

包含在集合中的一个 ODBCError 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一个 ODBC 错误。*/ function test(){ 　　　　let er = Application.ODBCErrors.Item(1) 　　　　MsgBox ("The following error occurred:" + 　　　　   er.ErrorString + " : " + er.SqlState) }` |

**成员属性**

#### **ODBCErrors.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **ODBCErrors** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test(){ 　　　　let myObject = ActiveWorkbook     if (myObject.Application.Value == "ET"){ 　　　　    MsgBox ("This is an ET Application object.") 　　　　} 　　　　else{ 　　　　    MsgBox ("This is not an ET Application object.") 　　　　} }` |

#### **ODBCErrors.Count**

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

*express*   一个代表 **ODBCErrors** 对象的变量。

#### **ODBCErrors.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **ODBCErrors** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **ODBCErrors.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **ODBCErrors** 对象的变量。

适用环境：web

适用平台：windows/linux