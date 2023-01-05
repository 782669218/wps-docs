#### **CustomXMLPrefixMappings**



代表 **CustomXMLPrefixMapping** 对象的集合。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

以下示例通过向 **CustomXMLPrefixMapping** 集合添加命名空间和前缀，创建一个 **CustomXMLPrefixMapping** 对象。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `Application.ActivePresentation.CustomXMLParts.Item(1).NamespaceManager.AddNamespace("xs", "urn:invoice:namespace")` |

**方法**

|                                                              | 名称                | 说明                                                    |
| ------------------------------------------------------------ | ------------------- | ------------------------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddNamespace**    | 允许您添加要在查询项目时使用的自定义命名空间/前缀映射。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **LookupNamespace** | 允许您获取对应于指定前缀的命名空间。                    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **LookupPrefix**    | 允许您获取对应于指定命名空间的前缀。                    |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 获取一个代表 **CustomXMLPrefixMappings** 对象的容器应用程序的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 获取一个 **Long** 类型的值，指示 **CustomXMLPrefixMappings** 集合中的项数。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 获取一个 32 位整数，指示创建 **CustomXMLPrefixMappings** 对象时所使用的应用程序。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Item**        | 获取 **CustomXMLPrefixMappings** 集合中的一个 **CustomXMLPrefixMapping** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 获取 **CustomXMLPrefixMappings** 对象的 **Parent** 对象。只读。 |

**成员方法**

#### **CustomXMLPrefixMappings.AddNamespace**

允许您添加要在查询项目时使用的自定义命名空间/前缀映射。

**语法**

**express.AddNamespace(Prefix, NamespaceURI)**

*express*   一个代表 **CustomXMLPrefixMappings** 对象的变量。

**参数**

| **名称**       | **必选/可选** | **数据类型** | **说明**                           |
| -------------- | ------------- | ------------ | ---------------------------------- |
| *Prefix*       | 必选          | **String**   | 包含要添加到前缀映射列表的前缀。   |
| *NamespaceURI* | 必选          | **String**   | 包含要分配给新添加前缀的命名空间。 |

**说明**

如果前缀在命名空间管理器中已存在，则此方法将覆盖该前缀的含义，除非前缀是数据存储（**IXMLDataStore** 接口）在内部添加或使用的（在这种情况下它将返回错误）。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

#### **CustomXMLPrefixMappings.LookupNamespace**

允许您获取对应于指定前缀的命名空间。

**语法**

**express.LookupNamespace(Prefix)**

*express*   一个代表 **CustomXMLPrefixMappings** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                   |
| -------- | ------------- | ------------ | -------------------------- |
| *Prefix* | 必选          | **String**   | 包含前缀映射列表中的前缀。 |

**说明**

如果未将命名空间分配给请求的前缀，则该方法返回空字符串 ("")。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

#### **CustomXMLPrefixMappings.LookupPrefix**

允许您获取对应于指定命名空间的前缀。

**语法**

**express.LookupPrefix(NamespaceURI)**

*express*   一个代表 **CustomXMLPrefixMappings** 对象的变量。

**参数**

| **名称**       | **必选/可选** | **数据类型** | **说明**           |
| -------------- | ------------- | ------------ | ------------------ |
| *NamespaceURI* | 必选          | **String**   | 包含命名空间 URI。 |

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**成员属性**

#### **CustomXMLPrefixMappings.Application**

获取一个代表 **CustomXMLPrefixMappings** 对象的容器应用程序的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **CustomXMLPrefixMappings** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

#### **CustomXMLPrefixMappings.Count**

获取一个 **Long** 类型的值，指示 **CustomXMLPrefixMappings** 集合中的项数。只读。

**语法**

**express.Count**

*express*   一个代表 **CustomXMLPrefixMappings** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

#### **CustomXMLPrefixMappings.Creator**

获取一个 32 位整数，指示创建 **CustomXMLPrefixMappings** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

*express*   一个代表 **CustomXMLPrefixMappings** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

#### **CustomXMLPrefixMappings.Item**

获取 **CustomXMLPrefixMappings** 集合中的一个 **CustomXMLPrefixMapping** 对象。只读。

**语法**

**express.Item**

*express*   一个代表 **CustomXMLPrefixMappings** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

#### **CustomXMLPrefixMappings.Parent**

获取 **CustomXMLPrefixMappings** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **CustomXMLPrefixMappings** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

适用环境：web

适用平台：windows/linux