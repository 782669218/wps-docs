**WorkbookConnection**



连接是从 WPS Office ET 2007 工作簿之外的外部数据源获取数据所需要的一组信息。

**说明**

连接可存储在 ET 工作簿中。工作簿打开时，ET 创建连接的内存副本，称为连接对象。连接对象包含服务器名称和服务器上打开的对象的名称等信息。连接对象也可以选择包含身份验证凭据和/或一个命令，该命令可传递给服务器执行（例如：由 SQL Server 执行的 SELECT 语句）。

| ![img]()注释                       |
| ---------------------------------- |
| 连接的确切形式取决于数据检索的机制 |

**方法**

|                                                              | 名称        | 说明             |
| ------------------------------------------------------------ | ----------- | ---------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**  | 删除工作簿连接。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Refresh** | 刷新工作簿连接。 |

**属性**

|                                                              | 名称                | 说明                                                         |
| ------------------------------------------------------------ | ------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**     | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**         | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Description**     | 返回或设置 **WorkbookConnection** 对象的简要说明。可读/写 **String** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**            | 返回或设置工作簿连接对象的名称。可读/写 **String** 类型。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ODBCConnection**  | 返回指定的 **WorkbookConnection** 对象的 ODBC 连接详细信息。只读 **ODBCConnection** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OLEDBConnection** | 返回指定的 **WorkbookConnection** 对象的 OLEDB 连接详细信息。只读 **OLEDBConnection** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**          | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Ranges**          | 返回指定的 **WorkbookConnection** 对象的对象区域。只读 **Ranges** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**            | 返回工作簿连接类型。只读 **XlConnectionType**。              |

**成员方法**

#### **WorkbookConnection.Delete**

删除工作簿连接。

**语法**

**express.Delete()**

*express*   一个代表 **WorkbookConnection** 对象的变量。

**说明**

使用此方法可以删除一个外部数据连接。此方法不适用于与其他工作簿之间的链接。

删除连接将不会删除正在使用该连接的任何对象。删除连接将不会导致从文件系统中删除任何连接文件。如果编辑任何这些对象以使用另一个连接，则一切都将重新开始工作。

那些使用已删除链接的对象所表现出的行为就好像无法建立连接一样。

#### **WorkbookConnection.Refresh**

刷新工作簿连接。

**语法**

**express.Refresh()**

*express*   一个代表 **WorkbookConnection** 对象的变量。

**说明**

如果 **DisplayAlerts** 属性为 **False**，则不显示对话框，**Refresh** 方法将失败，并引发“连接信息无效”异常。

刷新连接失败对刷新其他操作没有任何影响。

**成员属性**

#### **WorkbookConnection.Application**

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **WorkbookConnection** 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

#### **WorkbookConnection.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **WorkbookConnection** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **WorkbookConnection.Description**

返回或设置 **WorkbookConnection** 对象的简要说明。可读/写 **String** 类型。

**语法**

**express.Description**

*express*   一个代表 **WorkbookConnection** 对象的变量。

**说明**

在**“连接属性”**对话框中，用户可以编辑连接的名称和/或说明。在该对话框中更改名称和说明只会更改 ET 连接对象中的相应字段。

说明的最大长度为 255 个字符。如果用户指定的连接文件的说明长于 255 个字符，则说明会被截断，以满足 255 个字符的限制。

#### **WorkbookConnection.Name**

返回或设置工作簿连接对象的名称。可读/写 **String** 类型。

**语法**

**express.Name**

*express*   一个代表 **WorkbookConnection** 对象的变量。

#### **WorkbookConnection.ODBCConnection**

返回指定的 **WorkbookConnection** 对象的 ODBC 连接详细信息。只读 **ODBCConnection** 类型。

**语法**

**express.ODBCConnection**

*express*   一个代表 **WorkbookConnection** 对象的变量。

#### **WorkbookConnection.OLEDBConnection**

返回指定的 **WorkbookConnection** 对象的 OLEDB 连接详细信息。只读 **OLEDBConnection** 类型。

**语法**

**express.OLEDBConnection**

*express*   一个代表 **WorkbookConnection** 对象的变量。

#### **WorkbookConnection.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **WorkbookConnection** 对象的变量。

#### **WorkbookConnection.Ranges**

返回指定的 **WorkbookConnection** 对象的对象区域。只读 **Ranges** 类型。

**语法**

**express.Ranges**

*express*   一个代表 **WorkbookConnection** 对象的变量。

#### **WorkbookConnection.Type**

返回工作簿连接类型。只读 **XlConnectionType**。

**语法**

**express.Type**

*express*   一个代表 **WorkbookConnection** 对象的变量。

适用环境：web

适用平台：windows/linux