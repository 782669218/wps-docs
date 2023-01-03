**QueryTable**



代表一个利用从外部数据源（如 SQL Server 或 Microsoft Access 数据库）返回的数据生成的工作表表格。

**说明**

**QueryTable** 对象是 **QueryTables** 集合的成员。

**方法**

|                                                              | 名称              | 说明                                                         |
| ------------------------------------------------------------ | ----------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CancelRefresh** | 取消指定查询表的所有后台查询。使用 **Refreshing** 属性可以确定当前是否正在进行后台查询。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**        | 删除对象。                                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Refresh**       | 更新外部数据区域 (**QueryTable**)。                          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ResetTimer**    | 重新设置指定的查询表或数据透视表的刷新计时器，使其时间间隔等于上次使用 **RefreshPeriod** 属性设置的时间间隔。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SaveAsODC**     | 将查询表缓存的源保存为“Microsoft Office 数据连接”文件。      |

**属性**

|                                                              | 名称                              | 说明                                                         |
| ------------------------------------------------------------ | --------------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AdjustColumnWidth**             | 如果每次刷新指定的查询表时列宽都会自动调整为最适合的宽度，则为 **True**。如果每次刷新时列宽不进行自动调整，则为 **False**。默认值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**                   | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象（可对 OLE 自动操作对象使用此属性来返回该对象的应用程序）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BackgroundQuery**               | 如果查询表的查询是异步执行（在后台执行）的，则为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CommandText**                   | 返回或设置指定数据源的命令串。**Variant** 型，可读写。       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CommandType**                   | 返回或设置备注部分的下表中列出的 **XlCmdType** 常量之一。返回或设置的常量用于描述 **CommandText** 属性的值。默认值为 **xlCmdSQL**。**XlCmdType** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Connection**                    | 返回或设置包含下列某项的字符串：允许 ET 连接到 OLE DB 数据源的 OLE DB 设置；允许 ET 连接到 ODBC 数据源的 ODBC 设置；允许 ET 连接到 Web 数据源的 URL；或者文本文件的名称或路径，或是指定某个数据库或 Web 查询的文件名称或路径。**Variant** 型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**                       | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Destination**                   | 返回查询表目标区域（查询结果表放置的区域）的左上角单元格。目标区域必须位于包含 **QueryTable** 对象的工作表中。**Range** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EditWebPage**                   | 返回或设置用于 Web 查询的网页的统一资源定位符 (URL)。**Variant** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EnableEditing**                 | 如果允许用户对指定查询表进行编辑，则该值为 **True**；如果用户只能刷新查询表，则该值为 **False**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EnableRefresh**                 | 如果用户可刷新数据透视表高速缓存或查询表，则为 **True**。默认值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FetchedRowOverflow**            | 如果上次使用 **Refresh** 方法返回的行数比工作表中可用行数大，则该值为 **True**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FieldNames**                    | 如果数据源的字段名称作为返回数据的列标题显示，则该值为 **True**。默认值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FillAdjacentFormulas**          | 如果数据源的字段名称作为返回数据的列标题显示，则该值为 **True**。默认值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ListObject**                    | 为 **QueryTable** 对象返回一个 **ListObject** 对象。**ListObject** 对象类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MaintainConnection**            | 如果从刷新数据开始直至关闭工作簿，都一直保留指向指定数据源的连接，则为 **True**。默认值是 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**                          | 返回或设置一个 **String** 值，它代表对象的名称。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parameters**                    | 返回一个 **Parameters** 集合，该集合表示查询表参数。只读。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**                        | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PostText**                      | 返回或设置用于 post 方法的字符串，post 方法用于向 Web 服务器输入数据以从 Web 查询中返回数据。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PreserveColumnInfo**            | 如果每次刷新查询表时，列排序、筛选和布局信息都会保留，则该值为 **True**。默认值为 **False**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PreserveFormatting**            | 如果将数据前五行的任何常用格式设置应用到查询表的新数据行，则为 **True**。对未使用的单元格不进行格式设置。如果将应用到查询表的最新一次自动套用格式应用于新数据行，则属性为 **False**。默认值是 **True**。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **QueryType**                     | 表示 ET 填充查询表时所使用的查询类型。**XlQueryType** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Recordset**                     | 返回或设置一个 **Recordset** 对象，该对象用作指定查询表的数据源。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Refreshing**                    | 如果指定的查询表正在进行后台查询，则为 **True**。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RefreshOnFileOpen**             | 如果每次打开工作簿时，数据透视表高速缓存或查询表自动更新，则为 **True**。默认值为 **False**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RefreshPeriod**                 | 返回或设置两次刷新之间的时间间隔。**Long** 型，可读写。      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RefreshStyle**                  | 返回或设置为了容纳查询返回的记录集中的行数而在指定工作表中插入或删除行时所使用的方式。**XlCellInsertionMode** 类型。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ResultRange**                   | 返回一个 **Range** 对象，该对象代表指定查询表所覆盖的工作表区域。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RobustConnect**                 | 返回或设置数据透视表缓存与其数据源连接的方式。**XlRobustConnect** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RowNumbers**                    | 如果行号作为第一列添加到指定查询表，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SaveData**                      | 如果将数据透视表的数据随工作簿一起保存，则为 **True**。如果仅保存数据透视表的定义，则为 **False**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SavePassword**                  | 如果将 ODBC 连接字符串中的密码信息与指定查询一起保存，则为 **True**。如果不保存密码信息，则该值为 **False**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Sort**                          | 返回查询表范围的排序条件。只读。                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SourceConnectionFile**          | 返回或设置一个 **String** 值，它指明用于创建查询表的 Microsoft Office 数据连接文件或类似文件。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SourceDataFile**                | 返回或设置一个 **String** 值，它指明查询表的源数据文件。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextFileColumnDataTypes**       | 返回或设置一个有序的常量数组，用于指定文本文件中相应列的数据类型，而该文本文件则是正要导入查询表中的文本文件。每一列的默认常量为 **xlGeneral**。**Variant** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextFileCommaDelimiter**        | 如果将文本文件导入查询表中时，以逗号作为分隔符，则该值为 **True**。如果以其他字符作为分隔符，则该值为 **False**。默认值为 **False**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextFileDecimalSeparator**      | 返回或设置小数分隔符，在将文本文件导入查询表中时，ET 将使用小数分隔符。默认值为系统小数分隔符。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextFileFixedColumnWidths**     | 返回或设置一个整数数组，该数组对应于正要向查询表中导入的文本文件的列宽（按字符）。有效的宽度为 1 到 32767 个字符。**Variant** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextFileOtherDelimiter**        | 返回或设置在向查询表中导入文本文件时用作分隔符的字符。默认值为 **null**。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextFileParseType**             | 返回或设置要导入查询表的文本文件中数据的列格式。**XlTextParsingType** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextFilePlatform**              | 返回或设置正向查询表中导入的文本文件的原始格式。该属性确定在数据导入过程中使用何种代码页。**XlPlatform** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextFilePromptOnRefresh**       | 如果每次刷新查询表时都要指定导入文本文件的名称，则该属性值为 **True**。**“导入文本文件”**对话框允许用户指定路径和文件名。默认值为 **False**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextFileSemicolonDelimiter**    | 如果在将文本文件导入查询表时使用分号作为分隔符，并且 **TextFileParseType** 属性的值为 **xlDelimited**，则为 **True**。默认值为 **False**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextFileSpaceDelimiter**        | 如果向查询表中导入文本文件时，使用空格字符作为分隔符，则该值为 **True**。默认值为 **False**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextFileStartRow**              | 返回或设置向查询表中导入文本文件时进行文本分列的起始行号。其有效值为 1 到 32767 之间的整数。默认值为 1。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextFileTabDelimiter**          | 如果向查询表中导入文本文件时使用 Tab 作为分隔符，则该值为 **True**。默认值为 **False**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextFileTextQualifier**         | 返回或设置向查询表中导入文本文件时的文本识别符。文本识别符用于指定包含的数据是文本格式。**XlTextQualifier** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextFileThousandsSeparator**    | 返回或设置向查询表中导入文本文件时 ET 所使用的千位分隔符。默认为系统千位分隔符。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextFileTrailingMinusNumbers**  | 如果为 **True**，则表示 ET 将导入的数字作为以“-”符号开头的文本，“-”符号为负号。如果为 **False**，则表示 ET 将导入的数字作为以“-”符号开头的文本，“-”符号为文本。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextFileVisualLayout**          | 返回或设置一个 **XlTextVisualLayoutType** 枚举，该枚举表示导入的文本的可见布局是从左到右还是从右到左。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **WebConsecutiveDelimitersAsOne** | 当从网页的 HTML <PRE> 标记中向查询表导入数据时，如果将连续多个分隔符看作单个分隔符，并且数据将被分列，则该值为 **True**。如果将连续多个分隔符看作多个分隔符，则该值为 **False**。默认值是 **True**。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **WebDisableDateRecognition**     | 向查询表中导入网页时，如果将类似日期的数据当作文本进行处理，则该值为 **True**。如果使用了日期识别，则该值为 **False**。默认值为 **False**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **WebDisableRedirections**        | 如果对 **QueryTable** 对象禁用 Web 查询重定向，则该属性的值为 **True**。默认值为 **False**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **WebFormatting**                 | 返回或设置一个值，该值确定向查询表中导入网页时网页中应用了多少格式设置（如果有）。**XlWebFormatting** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **WebPreFormattedTextToColumns**  | 返回或设置向查询表中导入网页时，是否对网页 HTML <PRE> 标记内的数据进行分列。默认值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **WebSelectionType**              | 返回或设置一个值，该值决定是向查询表中导入整个网页、网页上的所有表格还是仅网页上的特定表格。**XlWebSelectionType** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **WebSingleBlockTextImport**      | 向查询表中导入网页时，如果位于指定网页的 HTML <PRE> 标记中的数据是同时进行处理的，则该值为 **True**。如果数据是以连续行的数据块方式导入的，以便能识别标题行，则该值为 **False**。默认值为 **False**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **WebTables**                     | 向查询表中导入网页时，返回或设置由逗号分隔的表格名称或表格索引号的列表。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **WorkbookConnection**            | 返回查询表所使用的 **WorkbookConnection** 对象。只读。/p>    |

**成员方法**

#### **QueryTable.CancelRefresh**

取消指定查询表的所有后台查询。使用 **Refreshing** 属性可以确定当前是否正在进行后台查询。

**语法**

**express.CancelRefresh()**

*express*   一个代表 **QueryTable** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例取消查询表的刷新操作。 function test() {     let Qtable = Application.Worksheets.Item(1).QueryTables.Item(1)     if(Qtable.Refreshing)     {       Qtable.CancelRefresh()     } }` |

#### **QueryTable.Delete**

删除对象。

**语法**

**express.Delete()**

*express*   一个代表 **QueryTable** 对象的变量。

#### **QueryTable.Refresh**

更新外部数据区域 (**QueryTable**)。

**语法**

**express.Refresh(BackgroundQuery)**

*express*   一个代表 **QueryTable** 对象的变量。

**参数**

| **名称**          | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *BackgroundQuery* | 可选          | **Variant**  | 只用于基于 SQL 查询结果的 QueryTables。如果为 True，则在数据库建立连接并提交查询之后，将控制返回给过程。QueryTable 在后台进行更新。如果为 False，则在所有数据被取回到工作表之后，将控制返回给过程。如果没有指定该参数，则由 BackgroundQuery 属性的设置决定查询模式。 |

**返回值**

Boolean

**说明**

下列说明适用于基于 SQL 查询结果的 **QueryTable** 对象。

**Refresh** 方法使 ET 连接到 **QueryTable** 对象的数据源，执行 SQL 查询，并将数据返回到基于 **QueryTable** 对象的区域。除非调用该方法，否则 **QueryTable** 对象不与数据源通信。

当与 OLE DB 或 ODBC 数据源建立连接时，ET 使用由 **Connection** 属性指定的连接字符串。如果指定的连接字符串缺少必需的值，将显示对话框，提示用户提供必需的信息。如果 **DisplayAlerts** 属性为 **False**，则不显示对话框，**Refresh** 方法失败并导致“连接信息无效”异常。

在 ET 建立一个成功的连接之后，将存储完整的连接字符串，这样，以后在同一编辑会话中调用 **Refresh** 方法时就不会再显示提示。通过检查 **Connection** 属性的值可以获得完整的连接字符串。

完成数据库连接后，将检查 SQL 查询的有效性。如果该查询无效，**Refresh** 方法将失败并导致“SQL 语法错误”异常。

如果查询需要参数，则必须在调用 **Refresh** 方法之前，用参数绑定信息初始化 **Parameters** 集合。如果未绑定足够的参数，**Refresh** 方法将失败并导致“参数错误”异常。如果将参数设置为提示用户输入参数值，则无论 **DisplayAlerts** 属性的设置如何，都会向用户显示对话框。如果用户取消参数对话框，**Refresh** 将停止并返回 **False**。如果对 **Parameters** 绑定了额外的参数，则这些额外参数将被忽略。

如果成功地完成或启动查询，则 **Refresh** 方法返回 **True**；如果用户取消连接或参数对话框，该方法返回 **False**。

要查看取回的行数是否超过了工作表中的可用行数，请检查 **FetchedRowOverflow** 属性。每次调用 **Refresh** 方法之前，该属性都将被初始化。

#### **QueryTable.ResetTimer**

重新设置指定的查询表或数据透视表的刷新计时器，使其时间间隔等于上次使用 **RefreshPeriod** 属性设置的时间间隔。

**语法**

**express.ResetTimer()**

*express*   一个代表 **QueryTable** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//此示例重新设置当前活动工作表上第一张查询表的刷新计时器。 Application.ActiveSheet.QueryTables.Item(1).ResetTimer()` |

#### **QueryTable.SaveAsODC**

将查询表缓存的源保存为“Microsoft Office 数据连接”文件。

**语法**

**express.SaveAsODC(ODCFileName, Description, Keywords)**

*express*   一个代表 **QueryTable** 对象的变量。

**参数**

| **名称**      | **必选/可选** | **数据类型** | **说明**                               |
| ------------- | ------------- | ------------ | -------------------------------------- |
| *ODCFileName* | 必选          | **String**   | 保存文件的位置。                       |
| *Description* | 可选          | **Variant**  | 要保存到文件中的说明。                 |
| *Keywords*    | 可选          | **Variant**  | 可用于搜索该文件的以空格分隔的关键字。 |

**成员属性**

#### **QueryTable.AdjustColumnWidth**

如果每次刷新指定的查询表时列宽都会自动调整为最适合的宽度，则为 **True**。如果每次刷新时列宽不进行自动调整，则为 **False**。默认值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.AdjustColumnWidth**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

最大列宽为屏幕宽度的三分之二。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **AdjustColumnWidth** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//此示例关闭对新增查询表（第一个工作簿中的第一张工作表上）的列宽自动调整。 function test() {     let qyTab = Application.Workbooks.Item(1).Worksheets.Item(1).QueryTables.Add(varDBConnStr,Range("B1"),"Select Price From CurrentStocks Where Symbol = 'MSFT'")     qyTab.AdjustColumnWidth = false     qyTab.Refresh() }` |

#### **QueryTable.Application**

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象（可对 OLE 自动操作对象使用此属性来返回该对象的应用程序）。只读。

**语法**

**express.Application**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **Application** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例显示一条有关创建 myObject 的应用程序的消息。 function test() {     let myObject = Application.ActiveWorkbook     if(myObject.Application.Value == "ET")      { 	alert("This is an ET Application object.")     }     else      { 	alert("This is not an ET Application object.")     } }` |

#### **QueryTable.BackgroundQuery**

如果查询表的查询是异步执行（在后台执行）的，则为 **True**。**Boolean** 类型，可读写。

**语法**

**express.BackgroundQuery**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

对于 OLAP（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，此属性是只读的，并且始终返回 **False**。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **BackgroundQuery** 属性。

#### **QueryTable.CommandText**

返回或设置指定数据源的命令串。**Variant** 型，可读写。

**语法**

**express.CommandText**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

对于 OLE DB 源，**CommandType** 属性描述 **CommandText** 属性的值。

对于 ODBC 源，设置 **CommandText** 将导致刷新数据。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **CommandText** 属性。

包含相应查询表的工作表必须为活动状态，才可访问此属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//此示例设置第一张查询表的 ODBC 数据源的命令串。注意，该命令串是一个 SQL 语句。 function test() {     let qtQtrResults = Application.Workbooks.Item(1).Worksheets.Item(1).QueryTables.Item(1)     qtQtrResults.CommandType = xlCmdSql     qtQtrResults.CommandText = "Select ProductID From Products Where ProductID < 10"     qtQtrResults.Refresh() }` |

#### **QueryTable.CommandType**

返回或设置备注部分的下表中列出的 **XlCmdType** 常量之一。返回或设置的常量用于描述 **CommandText** 属性的值。默认值为 **xlCmdSQL**。**XlCmdType** 类型，可读写。

**语法**

**express.CommandType**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

| **XlCmdType** 可为以下这些 **XlCmdType** 常量之一。          |
| ------------------------------------------------------------ |
| **xlCmdCube**。包含 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源的多维数据集?（多维数据集：一种 OLAP 数据结构。多维数据集包含维度，如“国家/地区/省（或市/自治区）/市（或县）”，还包括数据字段，如“销售额”。维度将各种类型的数据组织到带有明细数据级别的分层结构中，而数据字段度量数量。）名称。 |
| **xlCmdDefault**。包含 OLE DB 提供者可识别的命令文本。       |
| **xlCmdSql**。包含一个 SQL 语句。                            |
| **xlCmdTable**。包含用于访问 OLE DB 数据源的表名称。         |

只有当查询表或数据透视表缓存的 **QueryType** 属性值为 **xlOLEDBQuery** 时，才可以设置 **CommandType** 属性。

当 **CommandType** 属性的值为 **xlCmdCube** 时，如果没有与查询表相关联的数据透视表，则不能更改该值。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **CommandType** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//此示例设置第一张查询表的 ODBC 数据源的命令串。注意，该命令串是一个 SQL 语句。 function test() {     let qtQtrResults = Application.Workbooks.Item(1).Worksheets.Item(1).QueryTables.Item(1)     qtQtrResults.CommandType = xlCmdSql     qtQtrResults.CommandText = "Select ProductID From Products Where ProductID < 10"     qtQtrResults.Refresh() }` |

#### **QueryTable.Connection**

返回或设置包含下列某项的字符串：允许 ET 连接到 OLE DB 数据源的 OLE DB 设置；允许 ET 连接到 ODBC 数据源的 ODBC 设置；允许 ET 连接到 Web 数据源的 URL；或者文本文件的名称或路径，或是指定某个数据库或 Web 查询的文件名称或路径。**Variant** 型，可读写。

**语法**

**express.Connection**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

设置 **Connection** 属性不会立即与数据源建立连接。必须使用 **Refresh** 方法建立连接并检索数据。

有关连接字符串语法的详细信息，请参阅 **QueryTables** 集合的 **Add** 方法。

另外，也可以通过选择 Microsoft ActiveX 数据对象 (ADO) 库直接访问数据源。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **Connection** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//此示例为第一张工作表上的第一个查询表提供新的 ODBC 连接信息。 Application.Worksheets.Item(1).QueryTables.Item(1).Connection = "ODBC;DSN=96SalesData;UID=Rep21;PWD=NUyHwYQI;"  //此示例指定一个文本文件。 Application.Worksheets.Item(1).QueryTables.Item(1).Connection = "TEXT;C:\\My Documents\\19980331.txt"` |

#### **QueryTable.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据则将作为 **ListObject** 对象导入。您可以使用 **ListObject** 的 **QueryTable** 属性访问 **Creator** 属性。

#### **QueryTable.Destination**

返回查询表目标区域（查询结果表放置的区域）的左上角单元格。目标区域必须位于包含 **QueryTable** 对象的工作表中。**Range** 类型，只读。

**语法**

**express.Destination**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **Destination** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例滚动活动窗口，将第一张查询表的左上角单元格滚动至窗口的左上角。 function test() {     let addr = Application.Worksheets.Item(1).QueryTables.Item(1).Destination     Application.ActiveWindow.ScrollColumn = addr.Column     Application.ActiveWindow.ScrollRow = addr.Row }` |

#### **QueryTable.EditWebPage**

返回或设置用于 Web 查询的网页的统一资源定位符 (URL)。**Variant** 类型，可读写。

**语法**

**express.EditWebPage**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

如果 **EditWebPage** 属性没有设置，则返回 **null**。只有当查询类型为 Web 或 OLE 时，**EditWebPage** 属性才有意义。

如果 **EditWebPage** 不是空值，则忽略用于刷新的 **WebTables** 属性。因此，XML 查询和 **WebTable** 属性将引用原始网页中的表，并且只用于编辑情况以预先填充**“Web 查询”**对话框。

如果使用用户界面导入数据，则会将来自 Web 查询或文本查询的数据作为 **QueryTable** 对象导入，而将所有其他外部数据作为 **ListObject** 对象导入。

如果使用对象模型导入数据，则必须将来自 Web 查询或文本查询的数据作为 **QueryTable** 导入，同时可将所有其他外部数据作为 **ListObject** 或 **QueryTable** 导入。

**EditWebPage** 属性只应用于 **QueryTable** 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//在本示例中，ET 将一个网页的 URL 显示给用户。本示例假定单元格 A1 中的 QueryTable 对象位于活动的工作表上，并且名为“MyhomePage.htm”的文件位于 C: 驱动器上。 function test() {     //Set the EditWebPage property to a source.     Range("A1").QueryTable.EditWebPage = "C:\\MyHomepage.htm"      //Display the source to the user.     alert(Range("A1").QueryTable.EditWebPage) }` |

#### **QueryTable.EnableEditing**

如果允许用户对指定查询表进行编辑，则该值为 **True**；如果用户只能刷新查询表，则该值为 **False**。**Boolean** 类型，可读写。

**语法**

**express.EnableEditing**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

本示例对第一张查询表进行设置，使用户不能对其进行编辑。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **EnableEditing** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.Worksheets.Item(1).QueryTables.Item(1).EnableEditing = false` |

#### **QueryTable.EnableRefresh**

如果用户可刷新数据透视表高速缓存或查询表，则为 **True**。默认值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.EnableRefresh**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

如果 **EnableRefresh** 属性设置为 **False**，则忽略 **RefreshOnFileOpen** 属性。

对于 OLAP（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，将此属性设置为 **False** 会禁用更新。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **EnableRefresh** 属性。

#### **QueryTable.FetchedRowOverflow**

如果上次使用 **Refresh** 方法返回的行数比工作表中可用行数大，则该值为 **True**。**Boolean** 类型，只读。

**语法**

**express.FetchedRowOverflow**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **FetchedRowOverflow** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例对第一张查询表进行刷新。如果由查询返回的行数超过工作表中的可用行数，则显示一条错误消息。 function test() {     let Qtable = Application.Worksheets.Item(1).QueryTables.Item(1)     Qtable.Refresh()     if(Qtable.FetchedRowOverflow)      {        alert("Query too large: please redefine.")     } }` |

#### **QueryTable.FieldNames**

如果数据源的字段名称作为返回数据的列标题显示，则该值为 **True**。默认值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.FieldNames**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**FieldNames** 属性仅适用于 **QueryTable** 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例对第一张查询表进行设置，使表中不显示字段名称。 Application.Worksheets.Item(1).QueryTables.Item(1).FieldNames = false` |

#### **QueryTable.FillAdjacentFormulas**

如果数据源的字段名称作为返回数据的列标题显示，则该值为 **True**。默认值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.FillAdjacentFormulas**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**FillAdjacentFormulas** 属性仅适用于 **QueryTable** 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例对第一张查询表进行设置，每当查询表刷新时就自动更新右侧的公式。 Application.Sheets.Item("sheet1").QueryTables.Item(1).FillAdjacentFormulas = true` |

#### **QueryTable.ListObject**

为 **QueryTable** 对象返回一个 **ListObject** 对象。**ListObject** 对象类型，只读。

**语法**

**express.ListObject**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**ListObject** 属性仅适用于 **ListObject** 对象。

#### **QueryTable.MaintainConnection**

如果从刷新数据开始直至关闭工作簿，都一直保留指向指定数据源的连接，则为 **True**。默认值是 **True**。**Boolean** 类型，可读写。

**语法**

**express.MaintainConnection**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

只有当查询表或数据透视表缓存的 **QueryType** 属性值为 **xlOLEDBQuery** 时，才可以设置 **MaintainConnection** 属性。

如果预计会频繁对服务器进行查询，则可将此属性设置为 **True**，这样能减少重新连接的时间因而可提高性能。将此属性设置为 **False**，将会关闭一个打开的连接。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **MaintainConnection** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例在活动工作表的 A3 单元格上新建一个基于 OLAP 提供程序 //（OLAP 提供程序：对特定类型的 OLAP 数据库提供访问功能的一组软件。该软件包括数据源驱动程序以及与数据库连接所必需的其他客户端软件。） //的数据透视表缓存，然后基于该缓存新建一个数据透视表。本示例将在初始刷新后终止连接。 function test() {     let PTcache = Application.ActiveWorkbook.PivotCaches().Add(xlExternal)     PTcache.Connection = "OLEDB;Provider=MSOLAP;Location=srvdata;Initial Catalog=National"     PTcache.MaintainConnection = false     PTcache.CreatePivotTable(Range("A3"),"PivotTable1")      let PT = Application.ActiveSheet.PivotTables("PivotTable1")	     PT.SmallGrid = false     PT.PivotCache.RefreshPeriod = 0     PT.CubeFields.Item("[state]").Orientation = xlColumnField     PT.CubeFields.Item("[state]").Position = 0     PT.CubeFields.Item("[Measures].[Count Of au_id]").Orientation = xlDataField     PT.CubeFields.Item("[Measures].[Count Of au_id]").Position = 0 }` |

#### **QueryTable.Name**

返回或设置一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

下表列出了 **Name** 属性和相关属性的示例值，同时假设有一个具有唯一名称“[Europe].[France].[Paris]”的 OLAP?（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，还有一个包含项名称“Paris”的非 OLAP 数据源。

| 属性       | 值（OLAP 数据源）                 | 值（非 OLAP 数据源）        |
| ---------- | --------------------------------- | --------------------------- |
| **题注**   | Paris                             | Paris                       |
| **名称**   | [Europe].[France].[Paris]（只读） | Paris                       |
| **源名称** | [Europe].[France].[Paris]（只读） | （与 SQL 属性值相同，只读） |
| **值**     | [Europe].[France].[Paris]（只读） | Paris                       |

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**Name** 属性仅适用于 **QueryTable** 对象。

#### **QueryTable.Parameters**

返回一个 **Parameters** 集合，该集合表示查询表参数。只读。

**语法**

**express.Parameters**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **Parameters** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例从现有带参数查询中返回 Parameters 集合。如果第一个参数使用字符数据类型，则用户只须在提示对话框中输入字符。 function test() {     let Qtable = Application.Sheets.Item("Sheet1").QueryTables.Item(1).Parameters.Item(1)     if(Qtable.DataType == xlParamTypeVarChar)      {        Qtable.SetParam(xlPrompt, "Enter a character only")     } }` |

#### **QueryTable.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **Parent** 属性。

#### **QueryTable.PostText**

返回或设置用于 post 方法的字符串，post 方法用于向 Web 服务器输入数据以从 Web 查询中返回数据。**String** 类型，可读写。

**语法**

**express.PostText**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

ET 带有 Web 查询示例，使用 WordPad 或其他文本编辑器可对其中的 HTML 代码进行更改。查询示例在 Microsoft Office 的安装文件夹的“Queries”文件夹中。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**PostText** 属性仅适用于 **QueryTable** 对象。

#### **QueryTable.PreserveColumnInfo**

如果每次刷新查询表时，列排序、筛选和布局信息都会保留，则该值为 **True**。默认值为 **False**。**Boolean** 类型，可读写。

**语法**

**express.PreserveColumnInfo**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

只有当查询表使用数据库连接时，该属性才有效。

可以将该属性设置为 **False**，以便与 ET 的早期版本兼容。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **PreserveColumnInfo** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例保留列排序、筛选和布局信息，以便与 ET 的早期版本兼容。 function test() {     let cnnConnect = Application.ADODB.Connection     cnnConnect.Open "Provider=SQLOLEDB;" & _     "Data Source=srvdata;" & _     "User ID=wadet;Password=4me2no;"      let rstRecordset = Application.ADODB.Recordset     rstRecordset.Open _     Source:="Select Name, Quantity, Price From Products", _     ActiveConnection:=cnnConnect, _     CursorType:=adOpenDynamic, _     LockType:=adLockReadOnly, _     Options:=adCmdText      let worsheet = Application.ActiveSheet.QueryTables.Add( Connection:=rstRecordset, Destination:=Range("A1"))     worsheet.Name = "Contact List"     worsheet.FieldNames = True     worsheet.RowNumbers = False     worsheet.FillAdjacentFormulas = False     worsheet.PreserveFormatting = True     worsheet.RefreshOnFileOpen = False     worsheet.BackgroundQuery = True     worsheet.RefreshStyle = xlInsertDeleteCells     worsheet.SavePassword = True     worsheet.SaveData = True     worsheet.AdjustColumnWidth = True     worsheet.RefreshPeriod = 0     worsheet.PreserveColumnInfo = True     worsheet.Refresh BackgroundQuery:=False  }` |

#### **QueryTable.PreserveFormatting**

如果将数据前五行的任何常用格式设置应用到查询表的新数据行，则为 **True**。对未使用的单元格不进行格式设置。如果将应用到查询表的最新一次自动套用格式应用于新数据行，则属性为 **False**。默认值是 **True**。

**语法**

**express.PreserveFormatting**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

对于数据库查询表，默认的格式设置为 **xlSimple**。

刷新查询表时，将对查询表应用新的自动套用格式样式。只要 **PreserveFormatting** 的值为 **False**，则 AutoFormat（自动套用格式）就会被设置为 **None**。这样，任何在 **PreserveFormatting** 被设置为 **False** 或在查询表刷新之前设置的自动套用格式都不会起作用，且相应产生的查询表也不会被应用任何格式。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **PreserveFormatting** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//此示例保留第一张工作表上的第一个数据透视表的格式。 Application.Worksheets.Item(1).PivotTables("Pivot1").PreserveFormatting = true  //此示例演示了将 PreserveFormatting 设置为 False 后，如何操作才能将“自动套用格式”设置为 xlRangeAutoFormatNone，而并不是指定的 xlRangeAutoFormatColor1 格式。 function test() {     let Qtable = Application.Workbooks.Item(1).Worksheets.Item(1).QueryTables.Item(1)     Qtable.Range.AutoFormat = xlRangeAutoFormatColor1     Qtable.PreserveFormatting = false     Qtable.Refresh() } ` |

#### **QueryTable.QueryType**

表示 ET 填充查询表时所使用的查询类型。**XlQueryType** 类型，只读。

**语法**

**express.QueryType**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

| **XlQueryType** 可为以下这些 **XlQueryType** 常量之一。      |
| ------------------------------------------------------------ |
| **xlTextImport**。基于文本文件，仅用于查询表                 |
| **xlOLEDBQuery**。基于 OLE DB 查询，包括 OLAP（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源 |
| **xlWebQuery**。基于网页，仅用于查询表                       |
| **xlADORecordset**。基于 ADO 记录集查询                      |
| **xlDAORecordSet**。基于 DAO 记录集查询，只用于查询表        |
| **xlODBCQuery**。基于 ODBC 数据源                            |

在 **Connection** 属性值的前缀中指定数据源。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **QueryType** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//如果表是基于网页的，则本示例将刷新第一张工作表上的第一个查询表。 function test() {     let qtQtrResults = Application.Workbooks.Item(1).Worksheets.Item(1).QueryTables.Item(1)     if(qtQtrResults.QueryType == xlWebQuery)     {        qtQtrResults.Refresh()     } }` |

#### **QueryTable.Recordset**

返回或设置一个 **Recordset** 对象，该对象用作指定查询表的数据源。

**语法**

**express.Recordset**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

如果使用此属性来覆盖现有记录集，则更改将在运行 **Refresh** 方法时生效。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **RecordSet** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//此示例对用于第一张查询表（第一张工作表上）的 Recordset 对象进行更改，然后刷新该查询表。 function test() {     Application.Worksheets.Item(1).QueryTables.Item(1).Recordset = Workbooks.OpenDatabase("C:\\Nwind.mdb").OpenRecordset("employees")     Application.Worksheets.Item(1).QueryTables.Item(1).Refresh() }` |

#### **QueryTable.Refreshing**

如果指定的查询表正在进行后台查询，则为 **True**。可读/写 **Boolean** 类型。

**语法**

**express.Refreshing**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

使用 **CancelRefresh** 方法可以取消后台查询。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **Refreshing** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例在查询表一正在进行后台查询的情况下显示了一条信息。 function test() {     if(Application.Worksheets.Item(1).QueryTables.Item(1).Refreshing)     {         alert("Query is currently refreshing: please wait")     }     else     {         Application.Worksheets.Item(1).QueryTables.Item(1).Refresh(false)         Application.Worksheets.Item(1).QueryTables.Item(1).ResultRange.Select()     } }` |

#### **QueryTable.RefreshOnFileOpen**

如果每次打开工作簿时，数据透视表高速缓存或查询表自动更新，则为 **True**。默认值为 **False**。**Boolean** 类型，可读写。

**语法**

**express.RefreshOnFileOpen**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

在 Visual Basic 中使用 **Open** 方法打开工作簿时，不会自动刷新查询表和数据透视表。请在打开工作簿后，使用 **Refresh** 方法刷新数据。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **RefreshOnFileOpen** 属性。

#### **QueryTable.RefreshPeriod**

返回或设置两次刷新之间的时间间隔。**Long** 型，可读写。

**语法**

**express.RefreshPeriod**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

将周期设置为 0（零），则会禁用自动定时刷新，并且等同于将该属性设置为 **Null**。

**RefreshPeriod** 属性的值可以是一个 0 到 32767 之间的整数。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **RefreshPeriod** 属性。

#### **QueryTable.RefreshStyle**

返回或设置为了容纳查询返回的记录集中的行数而在指定工作表中插入或删除行时所使用的方式。**XlCellInsertionMode** 类型。可读写。

**语法**

**express.RefreshStyle**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

| **XlCellInsertionMode** 可为以下 **XlCellInsertionMode** 常量之一。 |
| ------------------------------------------------------------ |
| **xlInsertDeleteCells** 插入或者删除部分行以适应新记录集所需要的实际行数。 |
| **xlOverwriteCells** 不向工作表添加新的单元格或行。如果溢出则覆盖周围单元格的数据。 |
| **xlInsertEntireRows** 必要时插入所有行以允许溢出。不从工作表删除单元格或行。 |

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **RefreshStyle** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例向 Sheet1 添加查询表，并对 RefreshStyle 属性进行设置，插入两个新行以放置查询结果数据。 function test() {     let Qtable = Application.Sheets.Item("sheet1").QueryTables.Add("Finder;c:\\myfile.dqy",Range("sheet1!a1"))     Qtable.RefreshStyle = xlInsertEntireRows     Qtable.Refresh() }` |

#### **QueryTable.ResultRange**

返回一个 **Range** 对象，该对象代表指定查询表所覆盖的工作表区域。只读。

**语法**

**express.ResultRange**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

该区域不包括字段名称行或行号列。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **ResultRange** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例对查询表一中的第一列数据进行汇总，并在数据区域下方显示第一列数据的总和。 function test() {     let Qtable = Application.Sheets.Item("sheet1").QueryTables.Item(1).ResultRange.Columns.Item(1)     Qtable.Name = "Column1"     Qtable.End(xlDown).Offset(2, 0).Formula = "=sum(Column1)" }` |

#### **QueryTable.RobustConnect**

返回或设置数据透视表缓存与其数据源连接的方式。**XlRobustConnect** 类型，可读写。

**语法**

**express.RobustConnect**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

| **XlRobustConnect** 可为以下这些 **XlRobustConnect** 常量之一。 |
| ------------------------------------------------------------ |
| **xlAlways** 。缓存始终使用外部源信息（由 **SourceConnectionFile** 或 **SourceDataFile** 属性定义）进行重新连接。 |
| **xlAsRequired** 。缓存通过 **Connection** 属性使用外部源信息进行重新连接。 |
| **xlNever** 。缓存从不使用源信息进行重新连接。               |

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **RobustConnect** 属性。

#### **QueryTable.RowNumbers**

如果行号作为第一列添加到指定查询表，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.RowNumbers**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

设置该属性为 **True** 并不会使行号立即显示。在查询表下一次刷新时行号将显示，而且在查询表每次刷新时将重新配置这些行号。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **RowNumbers** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例为查询表添加行号和字段名。 function test() {     let Qtable = Application.Worksheets.Item(1).QueryTables.Item("ExternalData1")     Qtable.RowNumbers = true     Qtable.FieldNames = true     Qtable.Refresh() }` |

#### **QueryTable.SaveData**

如果将数据透视表的数据随工作簿一起保存，则为 **True**。如果仅保存数据透视表的定义，则为 **False**。**Boolean** 类型，可读写。

**语法**

**express.SaveData**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

对于 OLAP（OLAP：为查询和报表（而不是处理事务）而进行了优化的数据库技术。OLAP 数据是按分级结构组织的，它存储在多维数据集而不是表中。） 数据源，此属性始终设置为 **False**。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **SaveData** 属性。

#### **QueryTable.SavePassword**

如果将 ODBC 连接字符串中的密码信息与指定查询一起保存，则为 **True**。如果不保存密码信息，则该值为 **False**。**Boolean** 类型，可读写。

**语法**

**express.SavePassword**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

此属性由数据透视表和查询表在 ODBC 和 OLEDB 查询中使用。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **SavePassword** 属性。

如果 **ListObject** 已连接到 SharePoint 列表，此属性将被忽略。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//此示例在每次保存查询表一时，都不保存 ODBC 连接字符串的密码信息。 Application.Worksheets.Item(1).QueryTables.Item(1).SavePassword = false` |

#### **QueryTable.Sort**

返回查询表范围的排序条件。只读。

**语法**

**express.Sort**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将导入为 **QueryTable** 对象，而所有其他数据将导入为 **ListObject** object.

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **Sort** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//下面的示例将刷新查询表，并获取排序条件。 function test() {     Application.QueryTable.Refresh()     let Qtable = Application.QueryTable.Sort }` |

#### **QueryTable.SourceConnectionFile**

返回或设置一个 **String** 值，它指明用于创建查询表的 Microsoft Office 数据连接文件或类似文件。可读写。

**语法**

**express.SourceConnectionFile**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。可以使用 **ListObject** 的 **QueryTable** 属性访问 **SourceConnectionFile** 属性。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **SourceConnectionFile** 属性。

#### **QueryTable.SourceDataFile**

返回或设置一个 **String** 值，它指明查询表的源数据文件。

**语法**

**express.SourceDataFile**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

对于基于文件的数据源（如 Access），**SourceDataFile** 属性包含源数据文件的完全限定路径。对于基于服务器的数据源（如 SQL Server），此属性设置为 **Null**。如果通过编程方式更改 **Connection** 属性，则 **SourceDataFile** 属性设置为 **Null**。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

可以使用 **ListObject** 的 **QueryTable** 属性访问 **SourceDataFile** 属性。

#### **QueryTable.TextFileColumnDataTypes**

返回或设置一个有序的常量数组，用于指定文本文件中相应列的数据类型，而该文本文件则是正要导入查询表中的文本文件。每一列的默认常量为 **xlGeneral**。**Variant** 类型，可读写。

**语法**

**express.TextFileColumnDataTypes**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

可以使用下表列出的 **xlColumnDataType** 常量来指定在数据导入过程中所使用的列数据类型或执行的操作。

| 常量                | 说明               |
| ------------------- | ------------------ |
| **xlGeneralFormat** | 概要               |
| **xlTextFormat**    | 文本               |
| **xlSkipColumn**    | 跳过列             |
| **xlDMYFormat**     | “日-月-年”日期格式 |
| **xlDYMFormat**     | “日-年-月”日期格式 |
| **xlEMDFormat**     | EMD 日期           |
| **xlMDYFormat**     | “月-日-年”日期格式 |
| **xlMYDFormat**     | “月-年-日”日期格式 |
| **xlYDMFormat**     | “年-日-月”日期格式 |
| **xlYMDFormat**     | “年-月-日”日期格式 |

仅当查询表基于文本文件数据（**QueryType** 属性设置为 **xlTextImport**）时，才使用此属性。

如果在含有多列的数组中指定了多个元素，则将忽略那些值。

只有在安装并选定了中国台湾地区语言时才可以使用 **xlEMDFormat**。**xlEMDFormat** 常量指定使用中国台湾地区纪元日期。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**TextFileColumnDataTypes** 属性仅适用于 **QueryTable** 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例向第一个工作簿中第一张工作表上的新查询表中导入一个固定宽度的文本文件。该文本文件的第一列为五个字符宽度，作为文本导入。第二列为四个字符宽度，被跳过。该文本文件的剩余部分被导入第三列，并对其应用常规格式。 function test() {     let Qtable = Application.Workbooks.Item(1).Worksheets.Item(1)     let Qtable1 = Qtable.QueryTables.Add("TEXT;C:\\My Documents\\19980331.txt",Qtable.Cells.Item(1, 1))     Qtable1.TextFileParseType = xlFixedWidth     Qtable1.TextFileFixedColumnWidths = [5, 4]     Qtable1.TextFileColumnDataTypes = [xlTextFormat, xlSkipColumn, xlGeneralFormat]     Qtable1.Refresh() }` |

#### **QueryTable.TextFileCommaDelimiter**

如果将文本文件导入查询表中时，以逗号作为分隔符，则该值为 **True**。如果以其他字符作为分隔符，则该值为 **False**。默认值为 **False**。**Boolean** 类型，可读写。

**语法**

**express.TextFileCommaDelimiter**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

仅当查询表基于文本文件数据（**QueryType** 属性设置为 **xlTextImport**）并且 **TextFileParseType** 属性的值为 **xlDelimited** 时，才使用此属性。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**TextFileCommaDelimiter** 属性仅适用于 **QueryTable** 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例将第一个工作簿中第一张工作表上的查询表的分隔符设置为逗号，然后刷新该查询表。 function test() {     let Qtable = Application.Workbooks.Item(1).Worksheets.Item(1)     let Qtable1 = Qtable.QueryTables.Add("TEXT;C:\\My Documents\\19980331.txt",Qtable.Cells.Item(1, 1))     Qtable1.TextFileParseType = xlDelimited     Qtable1.TextFileCommaDelimiter = true     Qtable1.Refresh() }` |

#### **QueryTable.TextFileDecimalSeparator**

返回或设置小数分隔符，在将文本文件导入查询表中时，ET 将使用小数分隔符。默认值为系统小数分隔符。**String** 类型，可读写。

**语法**

**express.TextFileDecimalSeparator**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

仅当查询表基于文本文件数据（**QueryType** 属性设置为 **xlTextImport**）并且该文件包含的小数位分隔符和千位分隔符与计算机上使用的不同（由于使用不同的语言设置）时，才使用此属性。

下表显示当使用不同分隔符向 ET 中导入文本时得到的不同结果。数字结果显示在最右边的列中。

| 系统小数分隔符 | 系统千位分隔符 | TextFileDecimalSeparator 值 | TextFileThousandsSeparator 值 | 导入的文本 | 单元格的值（数据类型） |
| -------------- | -------------- | --------------------------- | ----------------------------- | ---------- | ---------------------- |
| 句号           | 逗号           | 逗号                        | 句号                          | 123.123,45 | 123,123.45（数字）     |
| 句号           | 逗号           | 逗号                        | 逗号                          | 123.123,45 | 123.123,45（文本）     |
| 逗号           | 句号           | 逗号                        | 句号                          | 123,123.45 | 123,123.45（数字）     |
| 句号           | 逗号           | 句号                        | 逗号                          | 123 123.45 | 123 123.45（文本）     |
| 句号           | 逗号           | 句号                        | 空格                          | 123 123.45 | 123,123.45（数字）     |

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**TextFileDecimalSeparator** 属性仅适用于 **QueryTable** 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例为工作表 Sheet1 上第一份查询表保存最初的小数分隔符，并将其设置为逗号，以准备将一个法语文本文件（举例）导入 ET 的美国英语版中。 function test() {     let Qtable = Application.Worksheets.Item("Sheet1").QueryTables.Item(1).TextFileDecimalSeparator     Application.Worksheets.Item("Sheet1").QueryTables.Item(1).TextFileDecimalSeparator = "," }` |

#### **QueryTable.TextFileFixedColumnWidths**

返回或设置一个整数数组，该数组对应于正要向查询表中导入的文本文件的列宽（按字符）。有效的宽度为 1 到 32767 个字符。**Variant** 类型，可读写。

**语法**

**express.TextFileFixedColumnWidths**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

仅当查询表基于文本文件数据（**QueryType** 属性设置为 **xlTextImport**）并且 **TextFileParseType** 属性的值为 **xlFixedWidth** 时，才使用此属性。

必须指定一个有效的非负值列宽。如果指定的列超出了文本文件的宽度，则会忽略那些值。如果文本文件的宽度大于指定的列的总宽度，则文本文件的剩余部分将会导入到一个附加列中。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**TextFileFixedColumnWidths** 属性仅适用于 **QueryTable** 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例向第一个工作簿中第一张工作表上的新查询表中导入一个固定宽度的文本文件。该文本文件的第一列为五个字符宽度，作为文本导入。第二列为四个字符宽度，被跳过。该文本文件的剩余部分被导入第三列，并对其应用常规格式。 function test() {     let Qtable = Application.Workbooks.Item(1).Worksheets.Item(1)     let Qtable1 = Qtable.QueryTables.Add("TEXT;C:\\My Documents\\19980331.txt",Qtable.Cells.Item(1, 1))     Qtable1.TextFileParseType = xlFixedWidth     Qtable1.TextFileFixedColumnWidths = [5, 4]     Qtable1.TextFileColumnDataTypes = [xlTextFormat, xlSkipColumn, xlGeneralFormat]     Qtable1.Refresh() }` |

#### **QueryTable.TextFileOtherDelimiter**

返回或设置在向查询表中导入文本文件时用作分隔符的字符。默认值为 **null**。**String** 类型，可读写。

**语法**

**express.TextFileOtherDelimiter**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

仅当查询表基于文本文件数据（**QueryType** 属性设置为 **xlTextImport**）并且 **TextFileParseType** 属性的值为 **xlDelimited** 时，才使用此属性。

如果在字符串中指定了多个字符，则只使用第一个字符。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**TextFileOtherDelimiter** 属性仅适用于 **QueryTable** 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例将第一个工作簿中第一张工作表上的查询表的分隔符设置为井号 (#)，然后刷新该查询表。 function test() {     let Qtable = Application.Workbooks.Item(1).Worksheets.Item(1)     let Qtable1 = Qtable.QueryTables.Add("TEXT;C:\\My Documents\\19980331.txt",Qtable.Cells.Item(1, 1))     Qtable1.TextFileParseType = xlDelimited     Qtable1.TextFileOtherDelimiter = "#"     Qtable1.Refresh() }` |

#### **QueryTable.TextFileParseType**

返回或设置要导入查询表的文本文件中数据的列格式。**XlTextParsingType** 类型，可读写。

**语法**

**express.TextFileParseType**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

| **XlTextParsingType** 可为以下 **XlTextParsingType** 常量之一。 |
| ------------------------------------------------------------ |
| **xlFixedWidth** 表示将文件中的数据安排在固定宽度的列中。    |
| **xlDelimited** 默认值表示文件由分隔符分隔。                 |

仅当查询表基于文本文件数据（**QueryType** 属性设置为 **xlTextImport**）时，才使用此属性。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**TextFileParseType** 属性仅适用于 **QueryTable** 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例向第一个工作簿中第一张工作表上的新查询表中导入一个固定宽度的文本文件。该文本文件的第一列为五个字符宽度，作为文本导入。第二列为四个字符宽度，被跳过。该文本文件的剩余部分被导入第三列，并对其应用常规格式。 function test() {     let Qtable = Application.Workbooks.Item(1).Worksheets.Item(1)     let Qtable1 = Qtable.QueryTables.Add("TEXT;C:\\My Documents\\19980331.txt",Qtable.Cells.Item(1, 1))     Qtable1.TextFileParseType = xlFixedWidth     Qtable1.TextFileFixedColumnWidths = [5, 4]     Qtable1.TextFileColumnDataTypes = [xlTextFormat, xlSkipColumn, xlGeneralFormat]     Qtable1.Refresh() }` |

#### **QueryTable.TextFilePlatform**

返回或设置正向查询表中导入的文本文件的原始格式。该属性确定在数据导入过程中使用何种代码页。**XlPlatform** 类型，可读写。

**语法**

**express.TextFilePlatform**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

默认值是在**“文本文件导入向导”**的**“文件原始格式”**选项中的当前设置。

| **XlPlatform** 可为以下 **XlPlatform** 常量之一。 |
| ------------------------------------------------- |
| **xlMacintosh**                                   |
| **xlMSDOS**                                       |
| **xlWindows**                                     |

仅当查询表基于文本文件数据（**QueryType** 属性设置为 **xlTextImport**）时才使用该属性。

如果使用用户界面导入数据，则会将来自 Web 查询或文本查询的数据作为 **QueryTable** 对象导入，而将所有其他外部数据作为 **ListObject** 对象导入。

如果使用对象模型导入数据，则必须将来自 Web 查询或文本查询的数据作为 **QueryTable** 导入，同时可将所有其他外部数据作为 **ListObject** 或 **QueryTable** 导入。

**TextFilePlatform** 属性只应用于 **QueryTable** 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例向第一个工作簿中第一张工作表上的查询表中导入一个 MS-DOS 文本文件，然后刷新该查询表。 function test() {     let Qtable = Application.Workbooks.Item(1).Worksheets.Item(1)     let Qtable1 = Qtable.QueryTables.Add("TEXT;C:\\My Documents\\19980331.txt",Qtable.Cells.Item(1, 1))     Qtable1.TextFilePlatform = xlMSDOS     Qtable1.TextFileParseType = xlDelimited     Qtable1.TextFileTabDelimiter = true     Qtable1.Refresh() } ` |

#### **QueryTable.TextFilePromptOnRefresh**

如果每次刷新查询表时都要指定导入文本文件的名称，则该属性值为 **True**。**“导入文本文件”**对话框允许用户指定路径和文件名。默认值为 **False**。**Boolean** 类型，可读写。

**语法**

**express.TextFilePromptOnRefresh**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

仅当查询表基于文本文件数据（**QueryType** 属性设置为 **xlTextImport**）时才使用该属性。

如果该属性值为 **True**，则第一次刷新查询表时，并不显示对话框。

在用户界面中，其默认值为 **True**。

如果使用用户界面导入数据，则会将来自 Web 查询或文本查询的数据作为 **QueryTable** 对象导入，而将所有其他外部数据作为 **ListObject** 对象导入。

如果使用对象模型导入数据，则必须将来自 Web 查询或文本查询的数据作为 **QueryTable** 导入，同时可将所有其他外部数据作为 **ListObject** 或 **QueryTable** 导入。

**TextFilePromptOnRefresh** 属性只应用于 **QueryTable** 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例在每次刷新第一个工作簿中第一张工作表上的查询表时，都提示用户输入文本文件的名称。 function test() {     let Qtable = Application.Workbooks.Item(1).Worksheets.Item(1)     let Qtable1 = Qtable.QueryTables.Add("TEXT;C:\\My Documents\\19980331.txt",Qtable.Cells.Item(1, 1))     Qtable1.TextFileParseType = xlDelimited     Qtable1.TextFilePromptOnRefresh = true     Qtable1.TextFileTabDelimiter = true     Qtable1.Refresh() }` |

#### **QueryTable.TextFileSemicolonDelimiter**

如果在将文本文件导入查询表时使用分号作为分隔符，并且 **TextFileParseType** 属性的值为 **xlDelimited**，则为 **True**。默认值为 **False**。**Boolean** 类型，可读写。

**语法**

**express.TextFileSemicolonDelimiter**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

仅当查询表基于文本文件数据（**QueryType** 属性设置为 **xlTextImport**）时，才使用此属性。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**TextFileSemicolonDelimiter** 属性仅适用于 **QueryTable** 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例将第一个工作簿中第一张工作表上的查询表的分隔符设置为分号，然后刷新该查询表。 function test() {     let Qtable = Application.Workbooks.Item(1).Worksheets.Item(1)     let Qtable1 = Qtable.QueryTables.Add("TEXT;C:\\My Documents\\19980331.txt",Qtable.Cells.Item(1, 1))     Qtable1.TextFileParseType = xlDelimited     Qtable1.TextFileSemicolonDelimiter = true     Qtable1.Refresh() }` |

#### **QueryTable.TextFileSpaceDelimiter**

如果向查询表中导入文本文件时，使用空格字符作为分隔符，则该值为 **True**。默认值为 **False**。**Boolean** 类型，可读写。

**语法**

**express.TextFileSpaceDelimiter**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

仅当查询表基于文本文件数据（**QueryType** 属性设置为 **xlTextImport**）并且 **TextFileParseType** 属性的值为 **xlDelimited** 时，才使用此属性。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**TextFileSpaceDelimiter** 属性仅适用于 **QueryTable** 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例将空格字符设置为第一个工作簿中第一张工作表上的查询表的分隔符，然后刷新该查询表。 function test() {     let Qtable = Application.Workbooks.Item(1).Worksheets.Item(1)     let Qtable1 = Qtable.QueryTables.Add("TEXT;C:\\My Documents\\19980331.txt",Qtable.Cells.Item(1, 1))     Qtable1.TextFileParseType = xlDelimited     Qtable1.TextFileSemicolonDelimiter = true     Qtable1.Refresh() }` |

#### **QueryTable.TextFileStartRow**

返回或设置向查询表中导入文本文件时进行文本分列的起始行号。其有效值为 1 到 32767 之间的整数。默认值为 1。**Long** 类型，可读写。

**语法**

**express.TextFileStartRow**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

仅当查询表基于文本文件数据（**QueryType** 属性设置为 **xlTextImport**）时，才使用此属性。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**TextFileStartRow** 属性仅适用于 **QueryTable** 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例将第一个工作簿中第一张工作表上查询表的文本分列起始行设置为第五行，然后刷新该查询表。 function test() {     let Qtable = Application.Workbooks.Item(1).Worksheets.Item(1)     let Qtable1 = Qtable.QueryTables.Add("TEXT;C:\\My Documents\\19980331.txt",Qtable.Cells.Item(1, 1))     Qtable1.TextFileParseType = xlDelimited     Qtable1.TextFileStartRow = 5     Qtable1.TextFileTabDelimiter = true     Qtable1.Refresh() }` |

#### **QueryTable.TextFileTabDelimiter**

如果向查询表中导入文本文件时使用 Tab 作为分隔符，则该值为 **True**。默认值为 **False**。**Boolean** 类型，可读写。

**语法**

**express.TextFileTabDelimiter**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

仅当查询表基于文本文件数据（**QueryType** 属性设置为 **xlTextImport**）并且 **TextFileParseType** 属性的值为 **xlDelimited** 时，才使用此属性。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**TextFileTabDelimiter** 属性仅适用于 **QueryTable** 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例将第一个工作簿中第一张工作表上的查询表的分隔符设置为 Tab，然后刷新该查询表。 function test() {     let Qtable = Application.Workbooks.Item(1).Worksheets.Item(1)     let Qtable1 = Qtable.QueryTables.Add("TEXT;C:\\My Documents\\19980331.txt",Qtable.Cells.Item(1, 1))     Qtable1.TextFileParseType = xlDelimited     Qtable1.TextFileTabDelimiter = true     Qtable1.Refresh() }` |

#### **QueryTable.TextFileTextQualifier**

返回或设置向查询表中导入文本文件时的文本识别符。文本识别符用于指定包含的数据是文本格式。**XlTextQualifier** 类型，可读写。

**语法**

**express.TextFileTextQualifier**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

| **XlTextQualifier** 可为以下 **XlTextQualifier** 常量之一。 |
| ----------------------------------------------------------- |
| **xlTextQualifierNone**                                     |
| **xlTextQualifierDoubleQuote** 默认值                       |
| **xlTextQualifierSingleQuote**                              |

仅当查询表基于文本文件数据（**QueryType** 属性设置为 **xlTextImport**）时，才使用此属性。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**TextFileTextQualifier** 属性仅适用于 **QueryTable** 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例将第一个工作簿中第一张工作表上的查询表的文本识别符设置为单引号。 function test() {     let Qtable = Application.Workbooks.Item(1).Worksheets.Item(1)     let Qtable1 = Qtable.QueryTables.Add("TEXT;C:\\My Documents\\19980331.txt",Qtable.Cells.Item(1, 1))     Qtable1.TextFileParseType = xlDelimited     Qtable1.TextFileTextQualifier = xlTextQualifierSingleQuote     Qtable1.Refresh() }` |

#### **QueryTable.TextFileThousandsSeparator**

返回或设置向查询表中导入文本文件时 ET 所使用的千位分隔符。默认为系统千位分隔符。**String** 类型，可读写。

**语法**

**express.TextFileThousandsSeparator**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

仅当查询表基于文本文件数据（**QueryTable** 属性设置为 **xlTextImport**）时，尤其是当文件包含的小数位分隔符和千位分隔符与计算机上使用的不同（由于使用不同的语言设置）时，才使用此属性。

下表显示当使用不同分隔符向 ET 中导入文本时得到的不同结果。数字结果显示在最右边的列中。

| 系统小数分隔符 | 系统千位分隔符 | TextFileDecimalSeparator 值 | TextFileThousandsSeparator 值 | 导入的文本 | 单元格的值（数据类型） |
| -------------- | -------------- | --------------------------- | ----------------------------- | ---------- | ---------------------- |
| 句号           | 逗号           | 逗号                        | 句号                          | 123.123,45 | 123,123.45（数字）     |
| 句号           | 逗号           | 逗号                        | 逗号                          | 123.123,45 | 123.123,45（文本）     |
| 逗号           | 句号           | 逗号                        | 句号                          | 123,123.45 | 123,123.45（数字）     |
| 句号           | 逗号           | 句号                        | 逗号                          | 123 123.45 | 123 123.45（文本）     |
| 句号           | 逗号           | 句号                        | 空格                          | 123 123.45 | 123,123.45（数字）     |

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**TextFileThousandsSeparator** 属性仅适用于 **QueryTable** 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例保存工作表 Sheet1 上第一个查询表最初的千位分隔符，并将其设置为句号，以准备将法语文本文件（举例）导入 ET 的美国英语版中。 function test() {     let strDecSep = Application.Worksheets.Item("Sheet1").QueryTables.Item(1).TextFileThousandsSeparator     Application.Worksheets.Item("Sheet1").QueryTables.Item(1).TextFileThousandsSeparator = "." }` |

#### **QueryTable.TextFileTrailingMinusNumbers**

如果为 **True**，则表示 ET 将导入的数字作为以“-”符号开头的文本，“-”符号为负号。如果为 **False**，则表示 ET 将导入的数字作为以“-”符号开头的文本，“-”符号为文本。**Boolean** 类型，可读写。

**语法**

**express.TextFileTrailingMinusNumbers**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**TextFileTrailingMinusNumbers** 属性仅适用于 **QueryTable** 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//在本示例中，ET 确定单元格 A1 的设置，然后将导入的数值处理为以“-”符号开头的文本。本示例假定在活动工作表上存在 QueryTable 对象。 function test() { // Determine setting for TextFileTrailingMinusNumbers     if(Range("A1").QueryTable.TextFileTrailingMinusNumbers == true)     {         alert("Numbers imported as text that begin with a '-' symbol will be treated as a negative symbol.")     }     else      {         alert("Numbers imported as text that begin with a '-' symbol will not be treated as a negative symbol.")     } }` |

#### **QueryTable.TextFileVisualLayout**

返回或设置一个 **XlTextVisualLayoutType** 枚举，该枚举表示导入的文本的可见布局是从左到右还是从右到左。

**语法**

**express.TextFileVisualLayout**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

| **XlTextVisualLayoutType** 可为以下 **XlTextVisualLayoutType** 常量之一。 |
| ------------------------------------------------------------ |
| **xlTextVisualLTR**                                          |
| **xlTextVisualRTL**                                          |

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**TextFileVisualLayout** 属性仅适用于 **QueryTable** 对象。

#### **QueryTable.WebConsecutiveDelimitersAsOne**

当从网页的 HTML <PRE> 标记中向查询表导入数据时，如果将连续多个分隔符看作单个分隔符，并且数据将被分列，则该值为 **True**。如果将连续多个分隔符看作多个分隔符，则该值为 **False**。默认值是 **True**。可读/写 **Boolean** 类型。

**语法**

**express.WebConsecutiveDelimitersAsOne**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

仅当查询表的 **QueryType** 属性设置为 **xlWebQuery**，查询返回 HTML 文档，并且 **WebPreFormattedTextToColumns** 属性设置为 **True** 时，才使用此属性。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**WebConsecutiveDelimitersAsOne** 属性仅适用于 **QueryTable** 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例将第一个工作簿中第一张工作表上的查询表的分隔符设置为空格字符，然后刷新该查询表。连续多个空格作为单个空格处理。 function test() {     let shFirstQtr = Application.Workbooks.Item(1).Worksheets.Item(1)     let qtQtrResults = shFirstQtr.QueryTables.Add("URL;http://datasvr/98q1/19980331.htm",shFirstQtr.Cells.Item(1,1))     qtQtrResults.WebConsecutiveDelimitersAsOne = true     qtQtrResults.Refresh() }` |

#### **QueryTable.WebDisableDateRecognition**

向查询表中导入网页时，如果将类似日期的数据当作文本进行处理，则该值为 **True**。如果使用了日期识别，则该值为 **False**。默认值为 **False**。**Boolean** 类型，可读写。

**语法**

**express.WebDisableDateRecognition**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

仅当查询表的 **QueryType** 属性设置为 **xlWebQuery** 并且查询返回 HTML 文档时，才使用此属性。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**WebDisableDateRecognition** 属性仅适用于 **QueryTable** 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例关闭日期识别，以便类似于日期的网页数据将作为文本导入。然后刷新该查询表。 function test() {     let shFirstQtr = Application.Workbooks.Item(1).Worksheets.Item(1)      let qtQtrResults = shFirstQtr.QueryTables.Add("URL;http://datasvr/98q1/19980331.htm",shFirstQtr.Cells.Item(1,1))     qtQtrResults.WebDisableDateRecognition = true     qtQtrResults.Refresh() }` |

#### **QueryTable.WebDisableRedirections**

如果对 **QueryTable** 对象禁用 Web 查询重定向，则该属性的值为 **True**。默认值为 **False**。**Boolean** 类型，可读写。

**语法**

**express.WebDisableRedirections**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**WebDisableRedirections** 属性仅适用于 **QueryTable** 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//在本示例中，ET 确定工作簿中第一张工作表的 Web 查询重定向的设置。本示例假定 QueryTable 对象位于第一张工作表上，否则将出现运行错误。 function test() {     let wksSheet = Application.ActiveSheet     alert(wksSheet.QueryTables.Item(1).WebDisableRedirections) }` |

#### **QueryTable.WebFormatting**

返回或设置一个值，该值确定向查询表中导入网页时网页中应用了多少格式设置（如果有）。**XlWebFormatting** 类型，可读写。

**语法**

**express.WebFormatting**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

仅当查询表的 **QueryType** 属性设置为 **xlWebQuery** 并且查询返回 HTML 文档时，才使用此属性。

| XlWebFormatting 可为以下 XlWebFormatting 常量之一。 |
| --------------------------------------------------- |
| **xlWebFormattingAll**                              |
| **xlWebFormattingRTF**                              |
| **xlWebFormattingNone** 默认值                      |

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**WebFormatting** 属性仅适用于 **QueryTable** 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例在第一个工作簿中第一张工作表上添加一个新的 Web 查询表，并导入应用于数据的所有网页格式，然后刷新该查询表。 function test() {     let shFirstQtr = Application.Workbooks.Item(1).Worksheets.Item(1)     let qtQtrResults = shFirstQtr.QueryTables.Add("URL;http://datasvr/98q1/19980331.htm",shFirstQtr.Cells.Item(1,1))     qtQtrResults.WebFormatting = xlWebFormattingAll     qtQtrResults.Refresh() }` |

#### **QueryTable.WebPreFormattedTextToColumns**

返回或设置向查询表中导入网页时，是否对网页 HTML <PRE> 标记内的数据进行分列。默认值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.WebPreFormattedTextToColumns**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

仅当查询表的 **QueryType** 属性为 **xlWebQuery** 并且查询返回 HTML 文档时，才使用此属性。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**WebPreFormattedTextToColumns** 属性仅适用于 **QueryTable** 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例在第一个工作簿中第一张工作表上添加一个新的 Web 查询表。注意，本示例并未对 HTML <PRE> 标记之间的数据进行分列。 function test() {     let shFirstQtr = Application.Workbooks.Item(1).Worksheets.Item(1)     let qtQtrResults = shFirstQtr.QueryTables.Add("URL;http://datasvr/98q1/19980331.htm",shFirstQtr.Cells.Item(1,1))     qtQtrResults.WebFormatting = xlNone     qtQtrResults.WebPreFormattedTextToColumns = false     qtQtrResults.Refresh() }` |

#### **QueryTable.WebSelectionType**

返回或设置一个值，该值决定是向查询表中导入整个网页、网页上的所有表格还是仅网页上的特定表格。**XlWebSelectionType** 类型，可读写。

**语法**

**express.WebSelectionType**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

仅当查询表的 **QueryType** 属性设置为 **xlWebQuery** 并且查询返回 HTML 文档时，才使用此属性。

如果此属性的值为 **xlSpecifiedTables**，则可以使用 **WebTables** 属性指定要导入的表格。

| XlWebSelectionType 可为以下 XlWebSelectionType 常量之一。 |
| --------------------------------------------------------- |
| **xlEntirePage**                                          |
| **xlAllTables** 默认值                                    |
| **xlSpecifiedTables**                                     |

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**WebSelectionType** 属性仅适用于 **QueryTable** 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例在第一个工作簿中第一张工作表上添加一个新的 Web 查询表，然后导入网页中第一个和第二个表格中的数据。 function test() {     let shFirstQtr = Application.Workbooks.Item(1).Worksheets.Item(1)     let qtQtrResults = shFirstQtr.QueryTables.Add("URL;http://datasvr/98q1/19980331.htm",shFirstQtr.Cells.Item(1,1))     qtQtrResults.WebFormatting = xlNone     qtQtrResults.WebSelectionType = xlSpecifiedTables     qtQtrResults.WebTables = "1,2"     qtQtrResults.Refresh() } ` |

#### **QueryTable.WebSingleBlockTextImport**

向查询表中导入网页时，如果位于指定网页的 HTML <PRE> 标记中的数据是同时进行处理的，则该值为 **True**。如果数据是以连续行的数据块方式导入的，以便能识别标题行，则该值为 **False**。默认值为 **False**。**Boolean** 类型，可读写。

**语法**

**express.WebSingleBlockTextImport**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

仅当查询表的 **QueryType** 属性设置为 **xlWebQuery** 并且查询返回 HTML 文档时，才使用此属性。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**WebSingleBlockTextImport** 属性仅适用于 **QueryTable** 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例在第一个工作簿中第一张工作表上添加一个新的 Web 查询表，然后同时导入 HTML <PRE> 标记中的所有数据。 function test() {     let shFirstQtr = Application.Workbooks.Item(1).Worksheets.Item(1)     let qtQtrResults = shFirstQtr.QueryTables.Add("URL;http://datasvr/98q1/19980331.htm",shFirstQtr.Cells.Item(1,1))     qtQtrResults.WebSingleBlockTextImport = true     qtQtrResults.Refresh() }` |

#### **QueryTable.WebTables**

向查询表中导入网页时，返回或设置由逗号分隔的表格名称或表格索引号的列表。**String** 类型，可读写。

**语法**

**express.WebTables**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

仅当查询表的 **QueryType** 属性设置为 **xlWebQuery**，查询返回 HTML 文档，并且 **WebSelectionType** 属性的值为 **xlSpecifiedTables** 时，才使用此属性。

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**WebTables** 属性仅适用于 **QueryTable** 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例在第一个工作簿中第一张工作表上添加一个新的 Web 查询表，然后导入网页中第一个和第二个表格中的数据。 function test() {     let shFirstQtr = Application.Workbooks.Item(1).Worksheets.Item(1)     let qtQtrResults = shFirstQtr.QueryTables.Add("URL;http://datasvr/98q1/19980331.htm",shFirstQtr.Cells.Item(1,1))     qtQtrResults.WebFormatting = xlNone     qtQtrResults.WebSelectionType = xlSpecifiedTables     qtQtrResults.WebTables = "1,2"     qtQtrResults.Refresh() }` |

#### **QueryTable.WorkbookConnection**

返回查询表所使用的 **WorkbookConnection** 对象。只读。/p>

**语法**

**express.WorkbookConnection**

*express*   一个代表 **QueryTable** 对象的变量。

**说明**

如果使用用户界面导入数据，来自 Web 查询或文本查询的数据将作为 **QueryTable** 对象导入，而所有其他外部数据将作为 **ListObject** 对象导入。

如果使用对象模型导入数据，来自 Web 查询或文本查询的数据必须作为 **QueryTable** 导入，而所有其他外部数据既可作为 **ListObject** 导入，也可作为 **QueryTable** 导入。

**WorkbookConnection** 属性仅适用于 **QueryTable** 对象。

适用环境：web

适用平台：windows/linux