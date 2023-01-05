**Connections**



指定工作簿的 Connection 对象的集合。

**说明**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例演示如何从现有文件向工作簿添加连接。*/ Application.ActiveWorkbook.Connections.AddFromFile(     "C:\\Documents and Settings\\myComputer\\My Documents\\My Data Sources\\Northwind 2007 Customers.odc")  ` |

**方法**

|                                                              | 名称            | 说明                   |
| ------------------------------------------------------------ | --------------- | ---------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**         | 添加到工作簿的新连接。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddFromFile** | 添加从指定文件的连接。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**        | 此方法创建一个连接项。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回集合中对象的数目。只读 **Long** 类型。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回指定对象的父对象。只读。                                 |

**成员方法**

#### **Connections.Add**

添加到工作簿的新连接。

**语法**

**express.Add(Name, Description, ConnectionString, CommandText, lCmdtype)**

*express*   一个代表 **Connections** 对象的变量。

**参数**

| **名称**           | **必选/可选** | **数据类型** | **说明**                 |
| ------------------ | ------------- | ------------ | ------------------------ |
| *Name*             | 必选          | **String**   | 连接的名称。             |
| *Description*      | 必选          | **String**   | 连接的简要说明。         |
| *ConnectionString* | 必选          | **String**   | 连接字符串。             |
| *CommandText*      | 必选          | **String**   | 用于创建连接的命令文本。 |
| *lCmdtype*         | 可选          | **String**   | 命令类型。               |

**返回值**

WorkbookConnection

#### **Connections.AddFromFile**

添加从指定文件的连接。

**语法**

**express.AddFromFile(Filename)**

*express*   一个代表 **Connections** 对象的变量。

**参数**

| **名称**   | **必选/可选** | **数据类型** | **说明**     |
| ---------- | ------------- | ------------ | ------------ |
| *Filename* | 必选          | **String**   | 文件的名称。 |

**返回值**

WorkbookConnection

#### **Connections.Item**

此方法创建一个连接项。

**语法**

**express.Item(Index)**

*express*   一个代表 **Connections** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**     |
| -------- | ------------- | ------------ | ------------ |
| *Index*  | 必选          | **Variant**  | 项的索引值。 |

**返回值**

WorkbookConnection

**成员属性**

#### **Connections.Application**

如果不与对象识别符一起使用，则此属性返回代表 ET 应用程序的 **Application** 对象。如果与对象识别符一起使用，则此属性返回代表指定对象的创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **Connections** 对象的变量。

**说明**

您可以将此属性与 OLE 自动化对象一起使用以返回该对象的应用程序。

#### **Connections.Count**

返回集合中对象的数目。只读 **Long** 类型。

**语法**

**express.Count**

*express*   一个代表 **Connections** 对象的变量。

#### **Connections.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Connections** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **Connections.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **Connections** 对象的变量。

适用环境：web

适用平台：windows/linux