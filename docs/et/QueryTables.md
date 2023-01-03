**QueryTables**



**QueryTable** 对象的集合。

**说明**

每一个 **QueryTable** 对象都代表一个利用从外部数据源返回的数据生成的工作表表格。

**方法**

|                                                              | 名称     | 说明                   |
| ------------------------------------------------------------ | -------- | ---------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**  | 新建一个查询表。       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item** | 从集合中返回一个对象。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回一个 **Long** 值，它代表集合中对象的数量。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |

**成员方法**

#### **QueryTables.Add**

新建一个查询表。

**语法**

**express.Add(Connection, Destination, Sql)**

*express*   一个代表 **QueryTables** 对象的变量。

**参数**

| **名称**      | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Connection*  | 必选          | **Variant**  | 查询表的数据源。可为以下情形之一： 一个包含 OLE DB 或 ODBC 连接字符串的字符串。ODBC 连接字符串的格式为“ODBC;<连接字符串>”。 一个 QueryTable 对象，该对象表示查询信息的初始复制源，包括连接字符串和 SQL 文本，但不包括 Destination 区域。指定 QueryTable 对象会导致 Sql 参数被忽略。 一个 ADO 或 DAO Recordset 对象。从 ADO 或 DAO 记录集中读取数据。ET 会保留该记录集，直到该查询表被删除或 SQL 连接发生更改。不能对产生的查询表进行编辑。 一个 Web 查询，“URL;”格式的字符串，其中“URL;”是必需的，但不进行本地化，字符串的其余部分作为 Web 查询的 URL。 数据查找程序。“FINDER;<数据查找程序文件路径>”格式的字符串，其中“FINDER;”是必需的，但不能本地化。字符串的其余部分为数据查找程序文件（*.dqy 或 *.iqy）的路径和名称。使用 Add 方法时将读取该文件；之后，对查询表的 Connection 属性的调用将相应地返回以“ODBC”或“URL”开头的字符串。 一个文本文件。“TEXT;<文本文件路径和名称>”格式的字符串，其中 TEXT 是必需的，但不能本地化。 |
| *Destination* | 必选          | **Range**    | 查询表目标区域（生成的查询表的放置区域）左上角的单元格。目标区域必须位于 expression 指定的 QueryTables 对象所在的工作表中。 |
| *Sql*         | 可选          | **Variant**  | 在 ODBC 数据源上运行的 SQL 查询字符串。当使用的数据源为 ODBC 数据源时，该参数可选（如果不在此处指定该参数，则应该在查询表刷新之前使用查询表的 Sql 属性进行设置）。当将 QueryTable 对象、文本文件、ADO 或 DAO Recordset 对象指定为数据源时，不能使用该参数。 |

**返回值**

一个代表新查询表的 QueryTable 对象。

**说明**

由本方法创建的查询将在调用 **Refresh** 方法后运行。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例基于 ADO 记录集创建一个查询表。为了向后兼容，此示例保留了现有的列排序和筛选设置以及布局信息。*/ function test（）{ 　　　　Dim cnnConnect As ADODB.Connection 　　　　Dim rstRecordset As ADODB.Recordset  　　　　Set cnnConnect = New ADODB.Connection 　　　　cnnConnect.Open "Provider=SQLOLEDB;" & _ 　　　　    "Data Source=srvdata;" & _ 　　　　    "User ID=testac;Password=4me2no;"  　　　　Set rstRecordset = New ADODB.Recordset 　　　　rstRecordset.Open _ 　　　　    Source:="Select Name, Quantity, Price From Products", _ 　　　　    ActiveConnection:=cnnConnect, _  　　　　   CursorType:=adOpenDynamic, _   　　　　  LockType:=adLockReadOnly, _  　　　　   Options:=adCmdText  　　　　With ActiveSheet.QueryTables.Add( _       　　　　  Connection:=rstRecordset, _       　　　　  Destination:=Range("A1")) 　　　　    .Name = "Contact List"    　　　　 .FieldNames = True    　　　　 .RowNumbers = False    　　　　 .FillAdjacentFormulas = False     　　　　.PreserveFormatting = True    　　　　 .RefreshOnFileOpen = False    　　　　 .BackgroundQuery = True    　　　　 .RefreshStyle = xlInsertDeleteCells    　　　　 .SavePassword = True    　　　　 .SaveData = True   　　　　  .AdjustColumnWidth = True    　　　　 .RefreshPeriod = 0   　　　　  .PreserveColumnInfo = True    　　　　 .Refresh BackgroundQuery:=False 　　　　End With }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例向新的查询表中导入固定宽度的文本文件。该文本文件的第一列为 5 个字符宽，作为文本导入。第二列为 4 个字符宽，被跳过。其余部分则导入第三列中，并对其应用常规格式。*/ function test(){ 　　　　let shFirstQtr = Workbooks.Item(1).Worksheets.Item(1) 　　　　let qtQtrResults = shFirstQtr.QueryTables.Add("TEXT;C:\\My Documents\\19980331.txt",shFirstQtr.Cells.Item(1,1))    　　　　 qtQtrResults.TextFileParsingType = xlFixedWidth    　　　　 qtQtrResults.TextFileFixedColumnWidths = [5,4]    　　　　 qtQtrResults.TextFileColumnDataTypes = [xlTextFormat, xlSkipColumn, xlGeneralFormat]   　　　　  qtQtrResults.Refresh() }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例在活动工作表上新建查询表。*/ function test(){ 　　　　let sqlstring = "select 96Sales.totals from 96Sales where profit < 5" 　　　　let connstring = "ODBC;DSN=96SalesData;UID=Rep21;PWD=NUyHwYQI;Database=96Sales" 　　　　let qtadd = ActiveSheet.QueryTables.Add(connstring,Range("B1"),sqlstring)  　　　　   qtadd.Refresh() }` |

#### **QueryTables.Item**

从集合中返回一个对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **QueryTables** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**             |
| -------- | ------------- | ------------ | -------------------- |
| *Index*  | 必选          | **Variant**  | 对象的名称或索引号。 |

**返回值**

包含在集合中的一个 QueryTable 对象。

**示例**

本示例对查询表进行设置，以便查询表右侧的公式在每次刷新时都可以进行自动更新。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Sheets.Item("sheet1").QueryTables.Item(1).FillAdjacentFormulas = true` |

**成员属性**

#### **QueryTables.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **QueryTables** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test(){ 　　　　let myObject = ActiveWorkbook     if(myObject.Application.Value == "ET") { 　　　　    MsgBox("This is an ET Application object.") 　　　　}else { 　　　　    MsgBox("This is not an ET Application object.") 　　　　} }` |

#### **QueryTables.Count**

返回一个 **Long** 值，它代表集合中对象的数量。

**语法**

**express.Count**

*express*   一个代表 **QueryTables** 对象的变量。

#### **QueryTables.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **QueryTables** 对象的变量。

**说明**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

#### **QueryTables.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **QueryTables** 对象的变量。

适用环境：web

适用平台：windows/linux