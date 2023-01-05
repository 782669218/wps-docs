**ODBCConnection**



代表 ODBC 连接。

**说明**

ODBC 连接可存储在 ET 工作簿中。当 Micrososft ET 打开此工作簿时， ET 会创建一个称为 **ODBCConnection** 对象的 ODBC 连接的内存副本。

**ODBCConnection** 对象包含与连接相关的信息，例如，要连接到的服务器的名称和要在该服务器上打开的对象的名称。（可选）**ODBCConnection** 对象也可以包括身份验证凭据信息或要传递到服务器并执行的命令（例如，要由 SQL Server 执行的 `SELECT` 语句）。

**方法**

|                                                              | 名称              | 说明                                             |
| ------------------------------------------------------------ | ----------------- | ------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CancelRefresh** | 对指定的 ODBC 连接取消正在进行的所有刷新操作。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Refresh**       | 刷新 ODBC 连接。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SaveAsODC**     | 将 ODBC 连接保存为一个 WPS Office 数据连接文件。 |

**属性**

|                                                              | 名称                        | 说明                                                         |
| ------------------------------------------------------------ | --------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AlwaysUseConnectionFile** | 如果连接文件始终用于建立与数据源的连接，则为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**             | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BackgroundQuery**         | 如果 ODBC 连接的查询是异步执行（在后台执行）的，则为 **True**。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CommandText**             | 返回或设置指定数据源的命令字符串。可读/写 **Variant** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CommandType**             | 返回或设置 **XlCmdType** 常量之一。可读/写 **XlCmdType**。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Connection**              | 返回或设置一个包含 ODBC 设置的字符串，这些设置使 ET 可以连接 ODBC 数据源。可读/写 **Variant** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**                 | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EnableRefresh**           | 如果用户可刷新连接，则为 **True**。默认值是 **True**。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**                  | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RefreshDate**             | 返回上次刷新 ODBC 连接的日期。只读 **Date** 类型。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Refreshing**              | 如果正在进行指定 ODBC 连接的后台 ODBC 查询，则为 **True**。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RefreshOnFileOpen**       | 如果每次打开工作簿都自动更新连接，则为 **True**。默认值为 **False**。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RefreshPeriod**           | 返回或设置两次刷新之间的时间间隔（单位为分钟）。可读/写 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RobustConnect**           | 返回或设置 ODBC 连接与其数据源连接的方式。可读/写 **XlRobustConnect**。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SavePassword**            | 如果将 ODBC 连接字符串中的密码信息保存在连接字符串中，则为 **True**。如果删除密码，则为 **False**。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ServerCredentialsMethod** | 返回或设置应该用于服务器验证的凭据的类型。可读/写            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ServerSSOApplicationID**  | 返回或设置单一登录 (SSO) 应用程序标识符，该标识符用于在 SSO 数据库中查找凭据。可读/写 **String** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SourceConnectionFile**    | 返回或设置一个 **String** 类型的值，该值指示用于创建连接的 WPS Office 数据连接文件或类似文件。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SourceData**              | 返回 ODBC 连接的数据源，如表中所示。可读/写 **Variant** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SourceDataFile**          | 返回或设置一个 **String** 类型的值，该值指示 ODBC 连接的源数据文件。可读/写。 |

**成员方法**

#### **ODBCConnection.CancelRefresh**

对指定的 ODBC 连接取消正在进行的所有刷新操作。

**语法**

**express.CancelRefresh()**

*express*   一个代表 **ODBCConnection** 对象的变量。

**说明**

使用 **Refreshing** 属性可以确定当前是否正在进行刷新操作。

#### **ODBCConnection.Refresh**

刷新 ODBC 连接。

**语法**

**express.Refresh()**

*express*   一个代表 **ODBCConnection** 对象的变量。

**说明**

与 ODBC 数据源建立连接时，ET 使用 **Connection** 属性指定的连接字符串。如果指定的连接字符串缺少必需的值，将显示对话框，提示用户提供必需的信息。如果 **DisplayAlerts** 属性为 **False**，则不显示对话框，**Refresh** 方法失败并导致“连接信息无效”异常。

在 ET 成功建立一个连接之后，将存储完整的连接字符串，以便以后在同一编辑会话中调用 **Refresh** 方法时将不再显示提示。通过检查 **Connection** 属性的值可以获得完整的连接字符串。

建立数据库连接后，将检查 SQL 查询的有效性。如果该查询无效，**Refresh** 方法将失败并导致“SQL 语法错误”异常。

如果查询需要参数，则必须在调用 **Refresh** 方法之前，用参数绑定信息初始化 **Parameters** 集合。如果未绑定足够的参数，**Refresh** 方法将失败并导致“参数错误”异常。如果将参数设置为提示用户输入参数值，则无论 **DisplayAlerts** 属性的设置如何，都会向用户显示对话框。如果用户取消参数对话框，**Refresh** 方法将停止并返回 **False**。如果对 **Parameters** 集合绑定了额外的参数，则这些额外参数将被忽略。

如果成功地完成或启动查询，则 **Refresh** 方法将返回 **True**；如果用户取消连接或参数对话框，则该方法返回 **False**。

要查看取回的行数是否超过了工作表中的可用行数，请检查 **FetchedRowOverflow** 属性。每次调用 **Refresh** 方法之前，该属性都将被初始化。

#### **ODBCConnection.SaveAsODC**

将 ODBC 连接保存为一个 WPS Office 数据连接文件。

**语法**

**express.SaveAsODC(ODCFileName, Description, Keywords)**

*express*   一个代表 **ODBCConnection** 对象的变量。

**参数**

| **名称**      | **必选/可选** | **数据类型** | **说明**                               |
| ------------- | ------------- | ------------ | -------------------------------------- |
| *ODCFileName* | 必选          | **String**   | 保存文件的位置。                       |
| *Description* | 可选          | **Variant**  | 将保存在文件中的说明。                 |
| *Keywords*    | 可选          | **Variant**  | 可用于搜索该文件的以空格分隔的关键字。 |

**返回值**

无

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将连接保存为名为“ODCFile”的 ODC 文件。此示例假定 ODBC 连接位于活动工作表上。 */ function UseSaveAsODC(){     Application.ActiveWorkbook.ODBCConnection.SaveAsODC ("ODCFile") }` |

**成员属性**

#### **ODBCConnection.AlwaysUseConnectionFile**

如果连接文件始终用于建立与数据源的连接，则为 **True**。**Boolean** 类型，可读写。

**语法**

**express.AlwaysUseConnectionFile**

*express*   一个代表 **ODBCConnection** 对象的变量。

**说明**

当此属性为 **True** 时，连接文件将用于建立与数据源的连接。如果在工作簿中嵌入的连接不同于外部连接文件，则忽略嵌入的连接，且外部连接文件将是唯一考虑的版本。

#### **ODBCConnection.Application**

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **ODBCConnection** 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

#### **ODBCConnection.BackgroundQuery**

如果 ODBC 连接的查询是异步执行（在后台执行）的，则为 **True**。可读/写 **Boolean** 类型。

**语法**

**express.BackgroundQuery**

*express*   一个代表 **ODBCConnection** 对象的变量。

#### **ODBCConnection.CommandText**

返回或设置指定数据源的命令字符串。可读/写 **Variant** 类型。

**语法**

**express.CommandText**

*express*   一个代表 **ODBCConnection** 对象的变量。

**说明**

应使用 **CommandText** 属性，而不应使用 **SQL** 属性，现在 SQL 属性的存在主要是为了与 ET 的早期版本兼容。如果您同时使用这两个属性，则 **CommandText** 属性的值优先。

**CommandType** 属性描述了 **CommandText** 属性的值。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为第一张查询表的 ODBC 数据源设置命令字符串。注意，该命令字符串是一个 SQL 语句。 */ function test(){ 　　　　let qtQtrResults = Workbooks.Item(1).Worksheets.Item(1).QueryTables.Item(1) 　　　　qtQtrResults.CommandType = xlCmdSQL 　　　　qtQtrResults.CommandText = "Select ProductID From Products Where ProductID < 10" 　　　　qtQtrResults.Refresh() }` |

#### **ODBCConnection.CommandType**

返回或设置 **XlCmdType** 常量之一。可读/写 **XlCmdType**。

**语法**

**express.CommandType**

*express*   一个代表 **ODBCConnection** 对象的变量。

**说明**

返回或设置的常量描述了 **CommandText** 属性的值。默认值是 **xlCmdSQL**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为第一张查询表的 ODBC 数据源设置命令字符串。该命令字符串是一个 SQL 语句。 */ function test(){ 　　　　let qtQtrResults = Workbooks.Item(1).Worksheets.Item(1).QueryTables.Item(1) 　　　　qtQtrResults.CommandType = xlCmdSQL 　　　　qtQtrResults.CommandText = "Select ProductID From Products Where ProductID < 10" 　　　　qtQtrResults.Refresh() }` |

#### **ODBCConnection.Connection**

返回或设置一个包含 ODBC 设置的字符串，这些设置使 ET 可以连接 ODBC 数据源。可读/写 **Variant** 类型。

**语法**

**express.Connection**

*express*   一个代表 **ODBCConnection** 对象的变量。

**说明**

设置 **Connection** 属性并不会立即初始化到数据源的连接。必须使用 **Refresh** 方法建立连接并检索数据。

#### **ODBCConnection.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **ODBCConnection** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **ODBCConnection.EnableRefresh**

如果用户可刷新连接，则为 **True**。默认值是 **True**。可读/写 **Boolean** 类型。

**语法**

**express.EnableRefresh**

*express*   一个代表 **ODBCConnection** 对象的变量。

**说明**

如果 **EnableRefresh** 属性设置为 **False**，则 **RefreshOnFileOpen** 属性被忽略。对于 OLAP 数据源，如果将此属性设置为 **False**，则会禁用更新。

#### **ODBCConnection.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **ODBCConnection** 对象的变量。

#### **ODBCConnection.RefreshDate**

返回上次刷新 ODBC 连接的日期。只读 **Date** 类型。

**语法**

**express.RefreshDate**

*express*   一个代表 **ODBCConnection** 对象的变量。

**说明**

对于 OLAP 数据源，在每次查询后都会更新此属性。

#### **ODBCConnection.Refreshing**

如果正在进行指定 ODBC 连接的后台 ODBC 查询，则为 **True**。可读/写 **Boolean** 类型。

**语法**

**express.Refreshing**

*express*   一个代表 **ODBCConnection** 对象的变量。

**说明**

使用 **CancelRefresh** 方法可取消后台查询。

#### **ODBCConnection.RefreshOnFileOpen**

如果每次打开工作簿都自动更新连接，则为 **True**。默认值为 **False**。可读/写 **Boolean** 类型。

**语法**

**express.RefreshOnFileOpen**

*express*   一个代表 **ODBCConnection** 对象的变量。

**说明**

在 Visual Basic 中使用 **Open** 方法打开工作簿时，不会自动刷新连接。打开工作簿后，可使用 **Refresh** 方法刷新数据。

#### **ODBCConnection.RefreshPeriod**

返回或设置两次刷新之间的时间间隔（单位为分钟）。可读/写 **Long** 类型。

**语法**

**express.RefreshPeriod**

*express*   一个代表 **ODBCConnection** 对象的变量。

**说明**

将时间段设置为 0（零）会禁用定时自动刷新，等同于将此属性设置为 **Null**。**RefreshPeriod** 属性的值可以是从 0 到 32767 的整数。

#### **ODBCConnection.RobustConnect**

返回或设置 ODBC 连接与其数据源连接的方式。可读/写 **XlRobustConnect**。

**语法**

**express.RobustConnect**

*express*   一个代表 **ODBCConnection** 对象的变量。

**说明**

如果您使用可靠的连接，那么当 ET 无法使用工作簿连接信息建立连接时， ET 将检查工作簿连接是否具有外部连接文件的路径。如果有， ET 将打开外部连接文件并尝试使用外部连接文件中包含的信息进行连接。如果在使用外部连接文件之后能够建立连接，则 ET 将更新工作簿的连接对象。

在信息技术部门在一个中央位置维护和更新连接的情况下，这可以提供可靠的连接，每当上一版本的连接（缓存在工作簿中）失败时，允许用户的工作簿自动获取那些更新。

#### **ODBCConnection.SavePassword**

如果将 ODBC 连接字符串中的密码信息保存在连接字符串中，则为 **True**。如果删除密码，则为 **False**。可读/写 **Boolean** 类型。

**语法**

**express.SavePassword**

*express*   一个代表 **ODBCConnection** 对象的变量。

**说明**

此属性由数据透视表和查询表在 ODBC 和 OLE DB 查询中使用。

#### **ODBCConnection.ServerCredentialsMethod**

返回或设置应该用于服务器验证的凭据的类型。可读/写

**语法**

**express.ServerCredentialsMethod**

*express*   一个代表 **ODBCConnection** 对象的变量。

#### **ODBCConnection.ServerSSOApplicationID**

返回或设置单一登录 (SSO) 应用程序标识符，该标识符用于在 SSO 数据库中查找凭据。可读/写 **String** 类型。

**语法**

**express.ServerSSOApplicationID**

*express*   一个代表 **ODBCConnection** 对象的变量。

#### **ODBCConnection.SourceConnectionFile**

返回或设置一个 **String** 类型的值，该值指示用于创建连接的 WPS Office 数据连接文件或类似文件。可读/写。

**语法**

**express.SourceConnectionFile**

*express*   一个代表 **ODBCConnection** 对象的变量。

#### **ODBCConnection.SourceData**

返回 ODBC 连接的数据源，如表中所示。可读/写 **Variant** 类型。

**语法**

**express.SourceData**

*express*   一个代表 **ODBCConnection** 对象的变量。

**说明**

| 数据源          | 返回值                                                       |
| --------------- | ------------------------------------------------------------ |
| ET 列表或数据库 | 文本形式的单元格引用。                                       |
| 外部数据源      | 数组。每行包含一个 SQL 连接字符串，该行的剩余元素作为查询字符串，以 255 个字符为单位分段。 |
| 多个合并区域    | 二维数组。每行由一个引用及与其相关联的页字段项组成。         |
| 其他数据透视表  | 以上三种信息之一。                                           |

#### **ODBCConnection.SourceDataFile**

返回或设置一个 **String** 类型的值，该值指示 ODBC 连接的源数据文件。可读/写。

**语法**

**express.SourceDataFile**

*express*   一个代表 **ODBCConnection** 对象的变量。

**说明**

对于基于文件的数据源（如 Access），**SourceDataFile** 属性包含源数据文件的完全限定路径。对于基于服务器的数据源（如 SQL Server），该属性为 null。如果通过编程方式更改 **Connection** 属性，则 **SourceDataFile** 属性设置为 null。

适用环境：web

适用平台：windows/linux