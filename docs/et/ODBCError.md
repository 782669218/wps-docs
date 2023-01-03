**ODBCError**



代表一个由最近的 ODBC 查询生成的 ODBC 错误。

**说明**

**ODBCError** 对象是 **ODBCErrors** 集合的成员。如果指定的 ODBC 查询运行时没有出现错误，则 **ODBCErrors** 集合为空。集合中错误的索引次序与 ODBC 数据源生成它们时的次序相同。

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ErrorString** | 返回一个 **String** 值，它代表 ODBC 错误字符串。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SqlState**    | 返回 SQL 状态错误。**String** 型，只读。                     |

**成员属性**

#### **ODBCError.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **ODBCError** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test(){ 　　　　let myObject = ActiveWorkbook     if (myObject.Application.Value == "ET"){ 　　　　    MsgBox ("This is an ET Application object.") 　　　　} 　　　　else{ 　　　　    MsgBox ("This is not an ET Application object.") 　　　　} }` |

#### **ODBCError.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **ODBCError** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **ODBCError.ErrorString**

返回一个 **String** 值，它代表 ODBC 错误字符串。

**语法**

**express.ErrorString**

*express*   一个代表 **ODBCError** 对象的变量。

#### **ODBCError.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **ODBCError** 对象的变量。

#### **ODBCError.SqlState**

返回 SQL 状态错误。**String** 型，只读。

**语法**

**express.SqlState**

*express*   一个代表 **ODBCError** 对象的变量。

**说明**

有关特定错误的解释，请参阅 SQL 文档。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例刷新查询表一，并显示产生的所有 ODBC 错误。*/ function test(){ 　　　　let qts = Worksheets.Item(1).QueryTables.Item(1) 　　　　qts.Refresh() 　　　　let errs = Application.ODBCErrors     if (errs.Count > 0){ 　　　　    let r = qts.Destination.Cells.Item(1) 　　　　    r.Value2 = "The following errors occurred:" 　　　　    let c = 0 　　　　    for(let er = 1; er <= errs.Count; er++){ 　　　　        c = c + 1  　　　　       r.offset.Item(c, 0).value2 = errs.Item(er).ErrorString  　　　　       r.offset.Item(c, 1).value2 = errs.Item(er).SqlState  　　　　   } 　　　　}   　　　　else{ 　　　　    MsgBox ("Query complete: all records returned.") 　　　　} }` |

适用环境：web

适用平台：windows/linux