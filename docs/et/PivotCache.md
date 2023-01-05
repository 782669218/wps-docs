**PivotCache**



代表数据透视表的缓存。

**说明**

**PivotCache** 对象是 **PivotCaches** 集合的成员。

**方法**

|                                                              | 名称                 | 说明                                                         |
| ------------------------------------------------------------ | -------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CreatePivotTable** | 创建一个基于 **PivotCache** 对象的数据透视表。返回一个 **PivotTable** 对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **MakeConnection**   | 为指定的数据透视表缓存建立连接。                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Refresh**          | 立即重新绘制指定的图表。                                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ResetTimer**       | 重新设置指定的查询表或数据透视表的刷新计时器，使其时间间隔等于上次使用 **RefreshPeriod** 属性设置的时间间隔。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SaveAsODC**        | 将数据透视表缓存的源保存为“Microsoft Office 数据连接”文件。  |

**属性**

|                                                              | 名称                | 说明                                                         |
| ------------------------------------------------------------ | ------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ADOConnection**   | 如果将数据透视表缓存连接到 OLE DB 数据源，则返回一个 **ADO Connection** 对象。**ADOConnection** 属性表明 ET 与数据提供程序相连，允许用户在 ET 正在使用的同一会话的上下文内编写代码， ET 将该会话与 ADO（关系源）或 ADO MD（OLAP 源）一起使用。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**     | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BackgroundQuery** | 如果数据透视表的查询是异步执行（在后台执行）的，则为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CommandText**     | 返回或设置指定数据源的命令串。**Variant** 型，可读写。       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CommandType**     | 返回或设置说明部分的下表中列出的 **XlCmdType** 常量之一。返回或设置的常量用于描述 **CommandText** 属性的值。默认值为 **xlCmdSQL**。**XlCmdType** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Connection**      | 返回或设置包含下列某项的字符串：允许 ET 连接到 OLE DB 数据源的 OLE DB 设置；允许 ET 连接到 ODBC 数据源的 ODBC 设置；允许 ET 连接到 Web 数据源的 URL；或者文本文件的名称或路径，或是指定某个数据库或 Web 查询的文件名称或路径。**Variant** 型，可读写。 |

**成员方法**

#### **PivotCache.CreatePivotTable**

创建一个基于 **PivotCache** 对象的数据透视表。返回一个 **PivotTable** 对象。

**语法**

**express.CreatePivotTable(TableDestination, TableName, ReadData, DefaultVersion)**

*express*   一个代表 **PivotCache** 对象的变量。

**参数**

| **名称**           | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------------ | ------------- | ------------ | ------------------------------------------------------------ |
| *TableDestination* | 必选          | **Variant**  | 数据透视表目标区域（工作表中用于放置所生成的数据透视表的区域）左上角的单元格。目标区域必须位于工作簿（此工作簿包含由 expression 指定的 PivotCache 对象）的某个工作表中。 |
| *TableName*        | 可选          | **Variant**  | 新的数据透视表的名称。                                       |
| *ReadData*         | 可选          | **Variant**  | 如果该值为 True，则创建一个包含外部数据库中所有记录的数据透视表高速缓存；此高速缓存可以很大。如果为 False，则允许在实际读取数据之前将某些字段设置为基于服务器的页字段。 |
| *DefaultVersion*   | 可选          | **Variant**  | 数据透视表的默认版本。                                       |

**返回值**

PivotTable

**说明**

有关创建基于数据透视表高速缓存的数据透视表的另一种方法，请参阅 **PivotTables** 对象的 **Add** 方法。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在活动工作表的 A3 单元格上新建一个基于 OLAP 提供程序?（OLAP 提供程序：对特定类型的 OLAP 数据库提供访问功能的一组软件。该软件包括数据源驱动程序以及与数据库连接所必需的其他客户端软件。）的数据透视表高速缓存，然后基于该高速缓存新建一个数据透视表。*/ function test(){ 　　　　With ActiveWorkbook.PivotCaches.Add(SourceType:=xlExternal) 　　　　    .Connection = _  　　　　       "OLEDB;Provider=MSOLAP;Location=srvdata;Initial Catalog=National" 　　　　   .CommandType = xlCmdCube 　　　　   .CommandText = Array("Sales")   　　　　  .MaintainConnection = True 　　　　    .CreatePivotTable TableDestination:=Range("A3"), _ 　　　　        TableName:= "PivotTable1" 　　　　End With 　　　　With ActiveSheet.PivotTables("PivotTable1") 　　　　    .SmallGrid = False 　　　　    .PivotCache.RefreshPeriod = 0 　　　　    With .CubeFields("[state]") 　　　　        .Orientation = xlColumnField   　　　　      .Position = 1 　　　　    End With 　　　　    With .CubeFields("[Measures].[Count Of au_id]") 　　　　        .Orientation = xlDataField   　　　　      .Position = 1   　　　　  End With 　　　　End With}` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在活动工作表的 A3 单元格上，通过连接到 Microsoft Jet 上的 ADO 创建一个新的数据透视表高速缓存，然后再基于该高速缓存新建一个数据透视表。*/ function test(){ 　　　　Dim cnnConn As ADODB.Connection 　　　　Dim rstRecordset As ADODB.Recordset 　　　　Dim cmdCommand As ADODB.Command  　　　　' Open the connection. 　　　　Set cnnConn = New ADODB.Connection 　　　　With cnnConn  　　　　   .ConnectionString = _ 　　　　        "Provider=Microsoft.Jet.OLEDB.4.0" 　　　　    .Open "C:\perfdate\record.mdb" 　　　　End With  　　　　' Set the command text. 　　　　Set cmdCommand = New ADODB.Command 　　　　Set cmdCommand.ActiveConnection = cnnConn 　　　　With cmdCommand 　　　　    .CommandText = "Select Speed, Pressure, Time From DynoRun" 　　　　    .CommandType = adCmdText 　　　　    .Execute 　　　　End With  　　　　' Open the recordset. 　　　　Set rstRecordset = New ADODB.Recordset 　　　　Set rstRecordset.ActiveConnection = cnnConn 　　　　rstRecordset.Open cmdCommand 　　　　 　　　　' Create a PivotTable cache and report. 　　　　Set objPivotCache = ActiveWorkbook.PivotCaches.Add( _ 　　　　    SourceType:=xlExternal) 　　　　Set objPivotCache.Recordset = rstRecordset 　　　　With objPivotCache 　　　　    .CreatePivotTable TableDestination:=Range("A3"), _ 　　　　        TableName:="Performance" 　　　　End With  　　　　With ActiveSheet.PivotTables("Performance") 　　　　    .SmallGrid = False 　　　　    With .PivotFields("Pressure") 　　　　        .Orientation = xlRowField 　　　　        .Position = 1 　　　　    End With 　　　　    With .PivotFields("Speed")   　　　　      .Orientation = xlColumnField 　　　　        .Position = 1    　　　　 End With  　　　　   With .PivotFields("Time")   　　　　      .Orientation = xlDataField 　　　　        .Position = 1  　　　　   End With 　　　　End With  　　　　' Close the connections and clean up. 　　　　cnnConn.Close 　　　　Set cmdCommand = Nothing 　　　　Set rstRecordSet = Nothing 　　　　Set cnnConn = Nothing　　　　 }` |

#### **PivotCache.MakeConnection**

为指定的数据透视表缓存建立连接。

**语法**

**express.MakeConnection()**

*express*   一个代表 **PivotCache** 对象的变量。

**说明**

**MakeConnection** 方法可在缓存断开连接和用户想要重新建立连接之后使用。

如果没有连接缓存，各种对象和方法可能都会返回运行错误。使用该方法可确保在执行其他对象或方法之前已建立一个连接。

如果指定的数据透视表缓存的 **MaintainConnection** 属性已设置为 **False**，指定的数据透视表缓存的 **SourceType** 属性尚未设置为 *xlExternal*，或者连接不是 OLE DB 数据源连接，则该方法将生成一个运行时错误。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例确定是否将缓存连接到其源，如有必要，将与源建立连接。本示例假定数据透视表缓存位于活动工作表上。 */ function UseMakeConnection() {     let pvtCache = Application.ActiveWorkbook.PivotCaches().Item(1)     // Handle run-time error if external source is not an OLE DB data source.     try {     // Check connection setting and make connection if necessary.         if(pvtCache.IsConnected == true) {             MsgBox("The MakeConnection method is not needed.")         }         else {             pvtCache.MakeConnection()             MsgBox("A connection has been made.")         }     }     catch(exception) {         MsgBox("The data source is not an OLE DB data source")     } }` |

#### **PivotCache.Refresh**

立即重新绘制指定的图表。

**语法**

**express.Refresh()**

*express*   一个代表 **PivotCache** 对象的变量。

**示例**

此示例刷新工作簿中第一个工作表上的第一个数据透视表的数据透视表缓存。

| 示例代码复制                                               |
| ---------------------------------------------------------- |
| `Worksheets.Item(1).PivotTables(1).PivotCache().Refresh()` |

#### **PivotCache.ResetTimer**

重新设置指定的查询表或数据透视表的刷新计时器，使其时间间隔等于上次使用 **RefreshPeriod** 属性设置的时间间隔。

**语法**

**express.ResetTimer()**

*express*   一个代表 **PivotCache** 对象的变量。

#### **PivotCache.SaveAsODC**

将数据透视表缓存的源保存为“Microsoft Office 数据连接”文件。

**语法**

**express.SaveAsODC(ODCFileName, Description, Keywords)**

*express*   一个代表 **PivotCache** 对象的变量。

**参数**

| **名称**      | **必选/可选** | **数据类型** | **说明**                               |
| ------------- | ------------- | ------------ | -------------------------------------- |
| *ODCFileName* | 必选          | **String**   | 保存文件的位置。                       |
| *Description* | 可选          | **Variant**  | 要保存到文件中的说明。                 |
| *Keywords*    | 可选          | **Variant**  | 可用于搜索该文件的以空格分隔的关键字。 |

**示例**

下例将缓存的源保存为名为“ODCFile”的 ODC 文件。此示例假定数据透视表缓存位于活动工作表上。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function UseSaveAsODC() { 　　　　Application.ActiveWorkbook.PivotCaches().Item(1).SaveAsODC ("ODCFile")						 }` |

**成员属性**

#### **PivotCache.ADOConnection**

如果将数据透视表缓存连接到 OLE DB 数据源，则返回一个 **ADO Connection** 对象。**ADOConnection** 属性表明 ET 与数据提供程序相连，允许用户在 ET 正在使用的同一会话的上下文内编写代码， ET 将该会话与 ADO（关系源）或 ADO MD（OLAP 源）一起使用。只读。

**语法**

**express.ADOConnection**

*express*   一个代表 **PivotCache** 对象的变量。

**说明**

**ADOConnection** 属性仅供具有 OLE DB 数据源的会话使用。如果没有 ADO 会话，查询将导致运行时错误。

**ADOConnection** 属性可用于具有 ADO 的任何基于 OLE DB 的缓存。**ADO Connection** 对象可与 ADO MD 一起使用，以查找有关缓存所基于的 OLAP 多维数据集的信息。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 ADO DB Connection 对象设置为数据透视表缓存的 ADOConnection 属性。本示例假设数据透视表位于活动工作表上。 */ function test(){ 　　　　Sub UseADOConnection()  　　　　    Dim ptOne As PivotTable  　　　　   Dim cmdOne As New ADODB.Command 　　　　    Dim cfOne As CubeField    　　　　  Set ptOne = Sheet1.PivotTables(1)  　　　　   ptOne.PivotCache.MaintainConnection = True   　　　　  Set cmdOne.ActiveConnection = ptOne.PivotCache.ADOConnection    　　　　  ptOne.PivotCache.MakeConnection      ' Create a set.     cmdOne.CommandText = "Create Set [Warehouse].[My Set] as '{[Product].[All Products].Children}'"     cmdOne.CommandType = adCmdUnknown     cmdOne.Execute     　　　　 ' Add a set to the CubeField. 　　　　   Set cfOne = ptOne.CubeFields.AddSet("My Set", "My Set")  　　　　End Sub }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将添加一个计算成员，并假设数据透视表位于活动工作表上。 */ function test(){ 　　　　Sub AddMember()    　　　　  Dim cmd As New ADODB.Command     　　　　 If Not ActiveSheet.PivotTables(1).PivotCache.IsConnected Then     　　　　    ActiveSheet.PivotTables(1).PivotCache.MakeConnection   　　　　  End If    　　　　  Set cmd.ActiveConnection = ActiveSheet.PivotTables(1).PivotCache.ADOConnection   　　　　   ' Add a calculated member.  　　　　   cmd.CommandText = "CREATE MEMBER [Warehouse].[Product].[All Products].[Drink and Non-Consumable] AS '[Product].[All Products].[Drink] + [Product].[All Products].[Non-Consumable]'"    　　　　 cmd.CommandType = adCmdUnknown   　　　　  cmd.Execute     　　　　 ActiveSheet.PivotTables(1).PivotCache.Refresh  　　　　End Sub }` |

#### **PivotCache.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **PivotCache** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条有关创建 myObject 的应用程序的消息。*/ function test(){ 　　　　let myObject = ActiveWorkbook        if(myObject.Application.Value == "ET") {  　　　　   MsgBox("This is an ET Application object.") 　　　　} 　　　　else { 　　　　    MsgBox("This is not an ET Application object.") 　　　　} }` |

#### **PivotCache.BackgroundQuery**

如果数据透视表的查询是异步执行（在后台执行）的，则为 **True**。**Boolean** 类型，可读写。

**语法**

**express.BackgroundQuery**

*express*   一个代表 **PivotCache** 对象的变量。

**说明**

对于 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，此属性是只读的，并且始终返回 **False**。

**示例**

此示例使第一张工作表上的第一张数据透视表的查询在后台执行。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Worksheets.Item(1).PivotTables("Pivot1").PivotCache.BackgroundQuery = true` |

#### **PivotCache.CommandText**

返回或设置指定数据源的命令串。**Variant** 型，可读写。

**语法**

**express.CommandText**

*express*   一个代表 **PivotCache** 对象的变量。

**说明**

对于 OLE DB 数据源，**CommandType** 属性描述了 **CommandText** 属性的值。

对于 ODBC 源，设置 **CommandText** 将导致刷新数据。

#### **PivotCache.CommandType**

返回或设置说明部分的下表中列出的 **XlCmdType** 常量之一。返回或设置的常量用于描述 **CommandText** 属性的值。默认值为 **xlCmdSQL**。**XlCmdType** 类型，可读写。

**语法**

**express.CommandType**

*express*   一个代表 **PivotCache** 对象的变量。

**说明**

| XlCmdType 可为以下这些 XlCmdType 常量之一。                  |
| ------------------------------------------------------------ |
| xlCmdCube。包含一个 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源的多维数据集?（多维数据集：一种 OLAP 数据结构。多维数据集包含维度，如“国家/地区/省（或市/自治区）/市（或县）”，还包括数据字段，如“销售额”。维度将各种类型的数据组织到带有明细数据级别的分层结构中，而数据字段度量数量。）名称。 |
| xlCmdDefault。包含 OLE DB 提供者可识别的命令文本。           |
| xlCmdSql。包含一个 SQL 语句。                                |
| xlCmdTable。包含用于访问 OLE DB 数据源的表名称。             |

只有当查询表或数据透视表缓存的 **QueryType** 属性的值为 **xlOLEDBQuery** 时，才可以设置 **CommandType** 属性。

当 **CommandType** 属性的值为 **xlCmdCube** 时，如果没有与查询表相关联的数据透视表，则不能更改该值。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置第一张查询表的 ODBC 数据源的命令串。该命令串是一个 SQL 语句。*/ function test(){ 　　　　let qtQtrResults = Workbooks.Item(1).Worksheets.Item(1).QueryTables.Item(1) 　　　　qtQtrResults.CommandType = xlCmdSql 　　　　qtQtrResults.CommandText = "Select ProductID From Products Where ProductID < 10" 　　　　qtQtrResults.Refresh() }` |

#### **PivotCache.Connection**

返回或设置包含下列某项的字符串：允许 ET 连接到 OLE DB 数据源的 OLE DB 设置；允许 ET 连接到 ODBC 数据源的 ODBC 设置；允许 ET 连接到 Web 数据源的 URL；或者文本文件的名称或路径，或是指定某个数据库或 Web 查询的文件名称或路径。**Variant** 型，可读写。

**语法**

**express.Connection**

*express*   一个代表 **PivotCache** 对象的变量。

**说明**

在使用脱机多维数据集文件?（脱机多维数据集文件：创建于硬盘或网络共享位置上的文件，用于存储数据透视表或数据透视图的 OLAP 源数据。脱机多维数据集文件允许用户在断开与 OLAP 服务器的连接后继续进行操作。）时，请将 **UseLocalConnection** 属性设置为 **True**，并使用 **LocalConnection** 属性，而不是用 **Connection** 属性。

另外，也可以通过选择 Microsoft ActiveX 数据对象 (ADO) 库直接访问数据源。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例在活动工作表的 A3 单元格上新建一个基于 OLAP 提供程序?（OLAP 提供程序：对特定类型的 OLAP 数据库提供访问功能的一组软件。该软件包括数据源驱动程序以及与数据库连接所必需的其他客户端软件。）的数据透视表高速缓存，然后基于该高速缓存新建一个数据透视表。 */ function test(){ 　　　　let pi = ActiveWorkbook.PivotCaches().Add(xlExternal)    　　　　 pi.Connection = ("OLEDB;Provider=MSOLAP;Location=srvdata;Initial Catalog=National")    　　　　 pi.MaintainConnection = true    　　　　 pi.CreatePivotTable(Range("A3"),"PivotTable1") 										 　　　　let pt = ActiveSheet.PivotTables("PivotTable1")     　　　　pt.SmallGrid = false    　　　　 pt.PivotCache.RefreshPeriod = 0 										   　　　　  let cu = pt.CubeFields("[state]")      　　　　   cu.Orientation = xlColumnField     　　　　    cu.Position = 0 										   　　　　  let cub = pt.CubeFields("[Measures].[Count Of au_id]")      　　　　   cub.Orientation = xlDataField     　　　　    cub.Position = 0 }` |

适用环境：web

适用平台：windows/linux