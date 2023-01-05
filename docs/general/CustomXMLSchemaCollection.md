#### **CustomXMLSchemaCollection**



代表附加到数据流的 **CustomXMLSchema** 对象的集合。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

以下示例向 **CustomXMLSchemaCollection** 对象添加 **CustomXMLSchema** 对象。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `let SchemaCollection SchemaCollection.Add("http://tempuri.org/XMLSchema.xsd")` |

**方法**

|                                                              | 名称              | 说明                                                         |
| ------------------------------------------------------------ | ----------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**           | 通过该方法，可以将一个或多个架构添加到架构集合中，该集合然后可被添加到数据存储中的流以及架构库中。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddCollection** | 将现有的架构集合添加到当前架构集合。                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Validate**      | 指定架构集合中的架构是否有效（符合 XML 的语法规则以及指定词汇的规则；用于结构化 XML 的标准）。 |

**属性**

|                                                              | 名称             | 说明                                                         |
| ------------------------------------------------------------ | ---------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**  | 获取一个代表 **CustomXMLSchemaCollection** 对象的容器应用程序的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**        | 获取一个 **Long** 类型的值，指示 **CustomXMLSchemaCollection** 集合中的项数。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**      | 获取一个 32 位整数，指示创建 **CustomXMLSchemaCollection** 对象时所使用的应用程序。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Item**         | 获取 **CustomXMLSchemaCollection** 集合中的一个 **CustomXMLSchema** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NamespaceURI** | 获取 **CustomXMLSchemaCollection** 对象的命名空间的唯一地址标识符。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**       | 获取 **CustomXMLSchemaCollection** 对象的 **Parent** 对象。只读。 |

**成员方法**

#### **CustomXMLSchemaCollection.Add**

通过该方法，可以将一个或多个架构添加到架构集合中，该集合然后可被添加到数据存储中的流以及架构库中。

**语法**

**express.Add(NamespaceURI, Alias, FileName, InstallForAllUsers)**

*express*   一个代表 **CustomXMLSchemaCollection** 对象的变量。

**参数**

| **名称**             | **必选/可选** | **数据类型** | **说明**                                                     |
| -------------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *NamespaceURI*       | 可选          | **String**   | 包含要添加到集合中的架构的命名空间。如果架构在架构库中已存在，则方法将从该处检索架构。 |
| *Alias*              | 可选          | **String**   | 包含要添加到集合中的架构的别名。如果别名在架构库中已存在，则方法可以使用此参数找到它。 |
| *FileName*           | 可选          | **String**   | 包含架构在磁盘上的位置。如果指定此参数，则架构将被添加到集合以及架构库。 |
| *InstallForAllUsers* | 可选          | **Boolean**  | 指定在该方法将架构添加到架构库的情况下，是否应将架构库项写入注册表（针对所有用户的 HKey_Local_Machine，或仅针对当前用户的 HKey_Current_User）。该参数默认为 False 并写入 HKey_Current_User。 |

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例将架构添加到架构集合，从集合中选择单一的节点，然后将该节点返回到调用过程。 function test() {     let cxp1     let cxn     // Adds a schema to the collection.     cxp1 = Application.ActivePresentation.CustomXMLParts.Item(1).SchemaCollection.Add("urn:invoice:namespace", "coreDefinitions", "wdCore.xsd", true)     cxn = Application.ActivePresentation.CustomXMLParts.Item(1).SelectSingleNode("//*[@quantity < 4]")     return cxn }  }` |

#### **CustomXMLSchemaCollection.AddCollection**

将现有的架构集合添加到当前架构集合。

**语法**

**express.AddCollection(SchemaCollection)**

*express*   一个代表 **CustomXMLSchemaCollection** 对象的变量。

**参数**

| **名称**           | **必选/可选** | **数据类型**                  | **说明**                               |
| ------------------ | ------------- | ----------------------------- | -------------------------------------- |
| *SchemaCollection* | 必选          | **CustomXMLSchemaCollection** | 代表要导入到当前架构集合中的架构集合。 |

**说明**

如果在导入集合时命名空间之间存在冲突（例如，如果已将 x.xsd 链接到“urn:invoice:namespace”，但是传入的集合具有同一命名空间的 z.xsd），则优先采用传入的集合。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例接收目标架构集合和传入架构集合的参数，然后将一个集合添加到另一个集合。 function test(objTargetCustomXMLSchemaCollection, objIncomingCustomXMLSchemaCollection)  {     // Adds a schema collection to another schema the collection.     objTargetCustomXMLSchemaCollection.AddCollection(objIncomingCustomXMLSchemaCollection)           }` |

#### **CustomXMLSchemaCollection.Validate**

指定架构集合中的架构是否有效（符合 XML 的语法规则以及指定词汇的规则；用于结构化 XML 的标准）。

**语法**

**express.Validate()**

*express*   一个代表 **CustomXMLSchemaCollection** 对象的变量。

**说明**

除了确定架构是否有效之外，此方法还遍历集合中每个架构的 **include** 语句，并将引用的架构添加到源架构。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//下面的示例验证架构集合，并将 Boolean 结果返回给调用过程。 function test(objSourceCustomXMLSchemaCollection) {     let boolValid     // Validates the schemas in a schema collection.     boolValid = objSourceCustomXMLSchemaCollection.Validate()     ValidateSchemas = boolValid }` |

**成员属性**

#### **CustomXMLSchemaCollection.Application**

获取一个代表 **CustomXMLSchemaCollection** 对象的容器应用程序的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **CustomXMLSchemaCollection** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

#### **CustomXMLSchemaCollection.Count**

获取一个 **Long** 类型的值，指示 **CustomXMLSchemaCollection** 集合中的项数。只读。

**语法**

**express.Count**

*express*   一个代表 **CustomXMLSchemaCollection** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

#### **CustomXMLSchemaCollection.Creator**

获取一个 32 位整数，指示创建 **CustomXMLSchemaCollection** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

*express*   一个代表 **CustomXMLSchemaCollection** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

#### **CustomXMLSchemaCollection.Item**

获取 **CustomXMLSchemaCollection** 集合中的一个 **CustomXMLSchema** 对象。只读。

**语法**

**express.Item**

*express*   一个代表 **CustomXMLSchemaCollection** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

#### **CustomXMLSchemaCollection.NamespaceURI**

获取 **CustomXMLSchemaCollection** 对象的命名空间的唯一地址标识符。只读。

**语法**

**express.NamespaceURI**

*express*   一个代表 **CustomXMLSchemaCollection** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

#### **CustomXMLSchemaCollection.Parent**

获取 **CustomXMLSchemaCollection** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **CustomXMLSchemaCollection** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

适用环境：web

适用平台：windows/linux