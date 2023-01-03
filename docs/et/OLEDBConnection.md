**OLEDBConnection**



代表 OLE DB 连接。

**说明**

OLE DB 连接可存储在 ET 工作簿中。当 Micrososft ET 打开此工作簿时， ET 会创建一个称为 OLEDBConnection 对象的 OLE DB 连接的内存副本。
OLEDBConnection 对象包含与连接相关的信息，例如，要连接到的服务器的名称和要在该服务器上打开的对象的名称。（可选）OLEDBConnection 对象也可以包括身份验证凭据信息或要传递到服务器并执行的命令（例如，要由 SQL Server 执行的 SELECT 语句）。

**方法**

|                                                              | 名称               | 说明                                             |
| ------------------------------------------------------------ | ------------------ | ------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CancelRefresh**  | 对指定的 OLE DB 连接取消正在进行的所有刷新操作。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **MakeConnection** | 为指定的 OLE DB 连接建立连接。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Reconnect**      | 先放弃再重新连接指定的连接。                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Refresh**        | 刷新 OLE DB 连接。                               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SaveAsODC**      | 将 OLE DB 连接保存为 WPS Office 数据连接文件。   |

**属性**

|                                                              | 名称                        | 说明                                                         |
| ------------------------------------------------------------ | --------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ADOConnection**           | 如果将数据透视表缓存连接到 OLE DB 数据源，则返回一个 ADO Connection 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AlwaysUseConnectionFile** | 如果连接文件始终用于建立与数据源的连接，则为 **True** 。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**             | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BackgroundQuery**         | 如果 OLE DB 连接的查询是异步执行（在后台执行）的，则为 **True**。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CalculatedMembers**       | 返回指定连接的 **CalculatedMembers** 集合。只读              |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CommandText**             | 返回或设置指定数据源的命令字符串。可读/写 **Variant** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CommandType**             | 返回或设置 **XlCmdType** 常量之一。可读/写 **XlCmdType**。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Connection**              | 返回或设置一个包含 OLE DB 设置的字符串，这些设置使 ET 可以连接 OLE DB 数据源。可读/写 **Variant** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**                 | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EnableRefresh**           | 如果用户可刷新连接，则为 **True**。默认值是 **True**。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IsConnected**             | 如果 **MaintainConnection** 属性为 **True**，则返回 **True**，如果它当前未与其源相连，则返回 **False**。只读 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LocalConnection**         | 返回或设置脱机多维数据集文件的连接字符串。可读/写 **String** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LocaleID**                | 返回或设置指定的连接的区域设置标识符。可读/写。              |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MaintainConnection**      | 如果从刷新操作开始直至关闭工作簿，都一直保留与指定数据源的连接，则返回 **True**。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MaxDrillthroughRecords**  | 返回或设置要检索的记录的最大数目。可读/写 **Long** 类型。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OLAP**                    | 如果 OLE DB 连接与联机分析处理 (OLAP) 服务器相连，则返回 **True**，只读 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**                  | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RefreshDate**             | 返回上次刷新 OLE DB 连接的日期。只读 **Date** 类型。         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Refreshing**              | 如果正在进行指定 OLE DB 连接的后台 OLE DB 查询，则为 **True**。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RefreshOnFileOpen**       | 如果每次打开工作簿都自动更新连接，则为 **True**。默认值为 **False**。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RefreshPeriod**           | 返回或设置两次刷新之间的时间间隔（单位为分钟）。可读/写 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RetrieveInOfficeUILang**  | 如果要以 Office 用户界面显示语言（如果可用）检索数据和错误，则为 **True**。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RobustConnect**           | 返回或设置 OLE DB 连接与其数据源连接的方式。可读/写 **XlRobustConnect**。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SavePassword**            | 如果将 OLE DB 连接字符串中的密码信息保存在连接字符串中，则为 **True**。如果删除密码，则为 **False**。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ServerCredentialsMethod** | 返回或设置应该用于服务器验证的凭据的类型。可读/写 **XlCredentialsMethod**。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ServerFillColor**         | 如果在使用指定连接时从服务器中检索 OLAP 服务器的填充色格式，则为 **True**。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ServerFontStyle**         | 如果在使用指定连接时从服务器中检索 OLAP 服务器的字体样式格式，则为 **True**。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ServerNumberFormat**      | 如果在使用指定连接时从服务器中检索 OLAP 服务器的数字格式，则为 **True**。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ServerSSOApplicationID**  | 返回或设置单一登录 (SSO) 应用程序标识符，该标识符用于在 SSO 数据库中查找凭据。可读/写 **String** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ServerTextColor**         | 如果在使用指定连接时从服务器中检索 OLAP 服务器的文本颜色格式，则为 **True**。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SourceConnectionFile**    | 返回或设置一个 **String** 类型的值，该值指示用于创建连接的 WPS Office 数据连接文件或类似文件。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SourceDataFile**          | 返回或设置一个 **String** 类型的值，该值指示 OLE DB 连接的源数据文件。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **UseLocalConnection**      | 如果 **LocalConnection** 属性用于指定使 ET 能够连接到数据源的字符串，则为 **True**。如果使用 **Connection** 属性指定的连接字符串，则为 **False**。可读/写 **Boolean** 类型。 |

**成员方法**

#### **OLEDBConnection.CancelRefresh**

对指定的 OLE DB 连接取消正在进行的所有刷新操作。

**语法**

**express.CancelRefresh()**

*express*   一个代表 **OLEDBConnection** 对象的变量。

**说明**

使用 **Refreshing** 属性可以确定当前是否正在进行刷新操作。

#### **OLEDBConnection.MakeConnection**

为指定的 OLE DB 连接建立连接。

**语法**

**express.MakeConnection()**

*express*   一个代表 **OLEDBConnection** 对象的变量。

**返回值**

无

**说明**

**MakeConnection** 方法可在断开连接和用户想要重新建立连接时使用。

如果断开连接，则各种对象和方法可能都会返回运行时错误。使用该方法可确保在执行其他对象或方法之前已建立一个连接。

#### **OLEDBConnection.Reconnect**

先放弃再重新连接指定的连接。

**语法**

**express.Reconnect()**

*express*   一个代表 **OLEDBConnection** 对象的变量。

**示例**

下面的代码示例将使指定的连接重新连接。

| 示例代码复制                                |
| ------------------------------------------- |
| `ActiveWorkbook.Connections(1).Reconnect()` |

#### **OLEDBConnection.Refresh**

刷新 OLE DB 连接。

**语法**

**express.Refresh()**

*express*   一个代表 **OLEDBConnection** 对象的变量。

**说明**

创建到 OLE DB 数据源的连接时，ET 使用 **Connection** 属性指定的连接字符串。如果指定的连接字符串缺少必需的值，则显示对话框提示用户输入必需的信息。如果 **DisplayAlerts** 属性为 **False**，则不显示对话框，**Refresh** 方法将失败，并引发“连接信息无效”异常。

ET 成功创建连接之后，将存储完整的连接字符串，这样，以后在同一编辑会话中调用 **Refresh** 方法时就不会显示提示。通过检查 **Connection** 属性的值，可以获得完整的连接字符串。

建立数据库连接后，将检查 SQL 查询的有效性。如果该查询无效，**Refresh** 方法将失败并引发“SQL 语法错误”异常。

如果查询成功完成或成功启动，则 **Refresh** 方法返回 **True**；如果用户取消连接，则返回 **False**。

要查看取回的行数是否超过了工作表中的可用行数，请检查 **FetchedRowOverflow** 属性。每次调用 **Refresh** 方法之前，该属性都将被初始化。

#### **OLEDBConnection.SaveAsODC**

将 OLE DB 连接保存为 WPS Office 数据连接文件。

**语法**

**express.SaveAsODC(ODCFileName, Description, Keywords)**

*express*   一个代表 **OLEDBConnection** 对象的变量。

**参数**

| **名称**      | **必选/可选** | **数据类型** | **说明**                               |
| ------------- | ------------- | ------------ | -------------------------------------- |
| *ODCFileName* | 必选          | **String**   | 保存文件的位置。                       |
| *Description* | 可选          | **Variant**  | 将保存在文件中的说明。                 |
| *Keywords*    | 可选          | **Variant**  | 可用于搜索该文件的以空格分隔的关键字。 |

**示例**

以下示例将连接保存为名为“ODCFile”的 ODC 文件。此示例假定 OLE DB 连接位于活动工作表上。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function UseSaveAsODC(){     Application.ActiveWorkbook.OLEDBConnection.SaveAsODC ("ODCFile") }` |

**成员属性**

#### **OLEDBConnection.ADOConnection**

如果将数据透视表缓存连接到 OLE DB 数据源，则返回一个 ADO Connection 对象。只读。

**语法**

**express.ADOConnection**

*express*   一个代表 **OLEDBConnection** 对象的变量。

**说明**

**ADOConnection** 属性表明 ET 与数据提供程序相连，允许用户在 ET 正在使用的同一会话的上下文中编写代码。

**ADOConnection** 属性只可用于数据源是 OLE DB 数据源的会话。如果没有 ADO 会话，查询将导致运行时错误。**ADOConnection** 属性可用于任意包含 ADO 的基于 OLE DB 的缓存。ADO Connection 对象可与 ADO MD 一起使用，以查找有关缓存所基于的 OLAP 多维数据集的信息。

#### **OLEDBConnection.AlwaysUseConnectionFile**

如果连接文件始终用于建立与数据源的连接，则为 **True** 。可读/写 **Boolean** 类型。

**语法**

**express.AlwaysUseConnectionFile**

*express*   一个代表 **OLEDBConnection** 对象的变量。

**说明**

当此属性为 **True** 时，连接文件将始终用于建立与数据源的连接。如果在工作簿中嵌入的连接不同于外部连接文件，嵌入的连接将被忽略，外部连接文件将是唯一考虑的版本。

#### **OLEDBConnection.Application**

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **OLEDBConnection** 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

#### **OLEDBConnection.BackgroundQuery**

如果 OLE DB 连接的查询是异步执行（在后台执行）的，则为 **True**。可读/写 **Boolean** 类型。

**语法**

**express.BackgroundQuery**

*express*   一个代表 **OLEDBConnection** 对象的变量。

**说明**

对于 OLAP 数据源，此属性是只读的，并且始终返回 **False**。

#### **OLEDBConnection.CalculatedMembers**

返回指定连接的 **CalculatedMembers** 集合。只读

**语法**

**express.CalculatedMembers**

*express*   一个代表 **OLEDBConnection** 对象的变量。

**说明**

使用 **OLEDBConnection** 集合的 **CalculatedMembers** 属性可在使用同一连接的多个数据透视表和多维数据集函数之间共享计算的成员。

#### **OLEDBConnection.CommandText**

返回或设置指定数据源的命令字符串。可读/写 **Variant** 类型。

**语法**

**express.CommandText**

*express*   一个代表 **OLEDBConnection** 对象的变量。

**说明**

应使用 **CommandText** 属性，而不应使用 **SQL** 属性，现在 SQL 属性的存在主要是为了与 ET 的早期版本兼容。如果您同时使用这两个属性，则 **CommandText** 属性的值优先。

**CommandType** 属性描述了 **CommandText** 属性的值。

#### **OLEDBConnection.CommandType**

返回或设置 **XlCmdType** 常量之一。可读/写 **XlCmdType**。

**语法**

**express.CommandType**

*express*   一个代表 **OLEDBConnection** 对象的变量。

**说明**

返回或设置的常量描述了 **CommandText** 属性的值。默认值是 **xlCmdSQL**。

#### **OLEDBConnection.Connection**

返回或设置一个包含 OLE DB 设置的字符串，这些设置使 ET 可以连接 OLE DB 数据源。可读/写 **Variant** 类型。

**语法**

**express.Connection**

*express*   一个代表 **OLEDBConnection** 对象的变量。

**说明**

设置 **Connection** 属性并不会立即初始化到数据源的连接。必须使用 **Refresh** 方法建立连接并检索数据。 在使用脱机多维数据集文件时，请将 **UseLocalConnection** 属性设置为 **True** 并使用 **LocalConnection** 属性，而不是 **Connection** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在活动工作表的 A3 单元格上创建一个基于 OLAP 提供程序的数据透视表缓存，然后基于该缓存创建一个数据透视表。 */ function test(){ 　　　　let apadd = ActiveWorkbook.PivotCaches.Add(xlExternal) 　　　　apadd.Connection = "OLEDB;Provider=MSOLAP;Location=srvdata;Initial Catalog=National" 　　　　apadd.MaintainConnection = true 　　　　apadd.CreatePivotTable (Range("A3"), "PivotTable1") 　　　　 　　　　let pts = ActiveSheet.PivotTables.Item("PivotTable1") 　　　　pts.SmallGrid = false 　　　　pts.PivotCache.RefreshPeriod = 0  　　　　let cfs = pts.CubeFields.Item("[state]") 　　　　cfs.Orientation = xlColumnField 　　　　cfs.Position = 0      　　　　let cfs2 = pts.CubeFields.Item("[Measures].[Count Of au_id]") 　　　　cfs2.Orientation = xlDataField 　　　　cfs2.Position = 0 }` |

#### **OLEDBConnection.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **OLEDBConnection** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **OLEDBConnection.EnableRefresh**

如果用户可刷新连接，则为 **True**。默认值是 **True**。可读/写 **Boolean** 类型。

**语法**

**express.EnableRefresh**

*express*   一个代表 **OLEDBConnection** 对象的变量。

**说明**

如果 **EnableRefresh** 属性设置为 **False**，则 **RefreshOnFileOpen** 属性被忽略。

#### **OLEDBConnection.IsConnected**

如果 **MaintainConnection** 属性为 **True**，则返回 **True**，如果它当前未与其源相连，则返回 **False**。只读 **Boolean** 类型。

**语法**

**express.IsConnected**

*express*   一个代表 **OLEDBConnection** 对象的变量。

**说明**

**IsConnected** 属性不检查是否建立了连接。即使该属性返回 **True**，如果连接不再有效，则向提供程序发送命令将导致错误

#### **OLEDBConnection.LocalConnection**

返回或设置脱机多维数据集文件的连接字符串。可读/写 **String** 类型。

**语法**

**express.LocalConnection**

*express*   一个代表 **OLEDBConnection** 对象的变量。

**说明**

对于非 OLAP 数据源，**LocalConnection** 属性的值为空字符串，并且 **UseLocalConnection** 属性设置为 **False**。

设置 **LocalConnection** 属性并不会立即初始化到数据源的连接。必须首先使用 **Refresh** 方法建立连接并检索数据。

如果 **UseLocalConnection** 属性设置为 **True**，则使用 **LocalConnection** 属性的值。如果 **UseLocalConnection** 属性设置为 **False**，则 **Connection** 属性将为基于源而不是基于本地多维数据集文件的查询表指定连接字符串。

#### **OLEDBConnection.LocaleID**

返回或设置指定的连接的区域设置标识符。可读/写。

**语法**

**express.LocaleID**

*express*   一个代表 **OLEDBConnection** 对象的变量。

**说明**

在将 **LocaleID** 属性设置为新的区域设置之前，必须将 **OLEDBConnection** 对象的 **RetrieveInOfficeUILang** 属性设置为 **False**。有关有效的区域设置 ID (LCID) 值的详细信息，请搜索 MSDN 网站查找“Locale IDs Assigned by Microsoft”（Microsoft 分配的区域设置 ID）。

#### **OLEDBConnection.MaintainConnection**

如果从刷新操作开始直至关闭工作簿，都一直保留与指定数据源的连接，则返回 **True**。可读/写 **Boolean** 类型。

**语法**

**express.MaintainConnection**

*express*   一个代表 **OLEDBConnection** 对象的变量。

**说明**

默认值是 **True**。如果预计会频繁对服务器进行查询，则可以将此属性设置为 **True**，这样能减少重新连接的时间，因而可提高性能。如果将此属性设置为 **False**，则会使打开的连接关闭。

#### **OLEDBConnection.MaxDrillthroughRecords**

返回或设置要检索的记录的最大数目。可读/写 **Long** 类型。

**语法**

**express.MaxDrillthroughRecords**

*express*   一个代表 **OLEDBConnection** 对象的变量。

#### **OLEDBConnection.OLAP**

如果 OLE DB 连接与联机分析处理 (OLAP) 服务器相连，则返回 **True**，只读 **Boolean** 类型。

**语法**

**express.OLAP**

*express*   一个代表 **OLEDBConnection** 对象的变量。

#### **OLEDBConnection.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **OLEDBConnection** 对象的变量。

#### **OLEDBConnection.RefreshDate**

返回上次刷新 OLE DB 连接的日期。只读 **Date** 类型。

**语法**

**express.RefreshDate**

*express*   一个代表 **OLEDBConnection** 对象的变量。

**说明**

对于 OLAP 数据源，在每次查询后都会更新此属性。

#### **OLEDBConnection.Refreshing**

如果正在进行指定 OLE DB 连接的后台 OLE DB 查询，则为 **True**。可读/写 **Boolean** 类型。

**语法**

**express.Refreshing**

*express*   一个代表 **OLEDBConnection** 对象的变量。

**说明**

使用 **CancelRefresh** 方法可取消后台查询。

#### **OLEDBConnection.RefreshOnFileOpen**

 如果每次打开工作簿都自动更新连接，则为 **True**。默认值为 **False**。可读/写 **Boolean** 类型。

**语法**

**express.RefreshOnFileOpen**

*express*   一个代表 **OLEDBConnection** 对象的变量。

**说明**

在 Visual Basic 中使用 **Open** 方法打开工作簿时，不会自动刷新连接。打开工作簿后，可使用 **Refresh** 方法刷新数据。

#### **OLEDBConnection.RefreshPeriod**

返回或设置两次刷新之间的时间间隔（单位为分钟）。可读/写 **Long** 类型。

**语法**

**express.RefreshPeriod**

*express*   一个代表 **OLEDBConnection** 对象的变量。

**说明**

将时间段设置为 0（零）会禁用定时自动刷新，等同于将此属性设置为 **Null**。**RefreshPeriod** 属性的值可以是从 0 到 32767 的整数。

#### **OLEDBConnection.RetrieveInOfficeUILang**

如果要以 Office 用户界面显示语言（如果可用）检索数据和错误，则为 **True**。可读/写 **Boolean** 类型。

**语法**

**express.RetrieveInOfficeUILang**

*express*   一个代表 **OLEDBConnection** 对象的变量。

**说明**

如果此属性设置为 **False**，则改用连接字符串中的 LCID 值。如果未指定 LCID，则使用服务器中的默认 LCID。

#### **OLEDBConnection.RobustConnect**

返回或设置 OLE DB 连接与其数据源连接的方式。可读/写 **XlRobustConnect**。

**语法**

**express.RobustConnect**

*express*   一个代表 **OLEDBConnection** 对象的变量。

**说明**

如果您使用可靠的连接，那么当 ET 无法使用工作簿连接信息建立连接时， ET 将检查工作簿连接是否具有外部连接文件的路径。如果有， ET 将打开外部连接文件并尝试使用外部连接文件中包含的信息进行连接。如果在使用外部连接文件之后能够建立连接，则 ET 将更新工作簿的连接对象。

在信息技术部门在一个中央位置维护和更新连接的情况下，这可以提供可靠的连接，每当上一版本的连接（缓存在工作簿中）失败时，允许用户的工作簿自动获取那些更新。

#### **OLEDBConnection.SavePassword**

如果将 OLE DB 连接字符串中的密码信息保存在连接字符串中，则为 **True**。如果删除密码，则为 **False**。可读/写 **Boolean** 类型。

**语法**

**express.SavePassword**

*express*   一个代表 **OLEDBConnection** 对象的变量。

**说明**

此属性由数据透视表和查询表在 ODBC 和 OLEDB 查询中使用。

#### **OLEDBConnection.ServerCredentialsMethod**

返回或设置应该用于服务器验证的凭据的类型。可读/写 **XlCredentialsMethod**。

**语法**

**express.ServerCredentialsMethod**

*express*   一个代表 **OLEDBConnection** 对象的变量。

#### **OLEDBConnection.ServerFillColor**

如果在使用指定连接时从服务器中检索 OLAP 服务器的填充色格式，则为 **True**。可读/写 **Boolean** 类型。

**语法**

**express.ServerFillColor**

*express*   一个代表 **OLEDBConnection** 对象的变量。

#### **OLEDBConnection.ServerFontStyle**

如果在使用指定连接时从服务器中检索 OLAP 服务器的字体样式格式，则为 **True**。可读/写 **Boolean** 类型。

**语法**

**express.ServerFontStyle**

*express*   一个代表 **OLEDBConnection** 对象的变量。

#### **OLEDBConnection.ServerNumberFormat**

如果在使用指定连接时从服务器中检索 OLAP 服务器的数字格式，则为 **True**。可读/写 **Boolean** 类型。

**语法**

**express.ServerNumberFormat**

*express*   一个代表 **OLEDBConnection** 对象的变量。

#### **OLEDBConnection.ServerSSOApplicationID**

返回或设置单一登录 (SSO) 应用程序标识符，该标识符用于在 SSO 数据库中查找凭据。可读/写 **String** 类型。

**语法**

**express.ServerSSOApplicationID**

*express*   一个代表 **OLEDBConnection** 对象的变量。

#### **OLEDBConnection.ServerTextColor**

如果在使用指定连接时从服务器中检索 OLAP 服务器的文本颜色格式，则为 **True**。可读/写 **Boolean** 类型。

**语法**

**express.ServerTextColor**

*express*   一个代表 **OLEDBConnection** 对象的变量。

#### **OLEDBConnection.SourceConnectionFile**

返回或设置一个 **String** 类型的值，该值指示用于创建连接的 WPS Office 数据连接文件或类似文件。可读/写。

**语法**

**express.SourceConnectionFile**

*express*   一个代表 **OLEDBConnection** 对象的变量。

#### **OLEDBConnection.SourceDataFile**

返回或设置一个 **String** 类型的值，该值指示 OLE DB 连接的源数据文件。可读/写。

**语法**

**express.SourceDataFile**

*express*   一个代表 **OLEDBConnection** 对象的变量。

**说明**

对于基于文件的数据源（如 Access），**SourceDataFile** 属性包含源数据文件的完全限定路径。对于基于服务器的数据源（如 SQL Server），该属性为 null。如果通过编程方式更改 **Connection** 属性，则 **SourceDataFile** 属性设置为 null。

#### **OLEDBConnection.UseLocalConnection**

如果 **LocalConnection** 属性用于指定使 ET 能够连接到数据源的字符串，则为 **True**。如果使用 **Connection** 属性指定的连接字符串，则为 **False**。可读/写 **Boolean** 类型。

**语法**

**express.UseLocalConnection**

*express*   一个代表 **OLEDBConnection** 对象的变量。

**说明**

本示例将第一个数据透视表缓存的连接字符串设置为引用脱机多维数据集文件

适用环境：web

适用平台：windows/linux