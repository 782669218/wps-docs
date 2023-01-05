#### **CustomXMLSchema**



代表 **CustomXMLSchemaCollection** 集合中的一个架构。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**方法**

|                                                              | 名称       | 说明                                        |
| ------------------------------------------------------------ | ---------- | ------------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete** | 从 **CustomXMLSchema** 集合中删除指定架构。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Reload** | 从文件重新加载架构。                        |

**属性**

|                                                              | 名称             | 说明                                                         |
| ------------------------------------------------------------ | ---------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**  | 获取一个代表 **CustomXMLSchema** 对象的容器应用程序的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**      | 获取一个 32 位整数，指示创建 **CustomXMLSchema** 对象时所使用的应用程序。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Location**     | 获取一个 **String** 类型的值，代表计算机上架构的位置。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NamespaceURI** | 获取 **CustomXMLSchema** 对象的命名空间的唯一地址标识符。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**       | 获取 **CustomXMLSchema** 对象的 **Parent** 对象。只读。      |

**成员方法**

#### **CustomXMLSchema.Delete**

从 **CustomXMLSchema** 集合中删除指定架构。

**语法**

**express.Delete()**

*express*   一个代表 **CustomXMLSchema** 对象的变量。

**说明**

如果尝试对已经过验证或附加到数据流的集合中的架构进行此操作，则不会执行操作并会显示一条错误消息。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例将架构添加到集合，然后删除该架构。 function test() {      try      {         // Adds a schema to the collection.     　　Application.ActivePresentation.CustomXMLParts.Item(1).SchemaCollection.Add("urn:invoice:namespace")          // Deletes the schema.         Application.ActivePresentation.CustomXMLParts.Item(1).SchemaCollection.Item(1).Delete()     }                 // Exception handling. Show the message and resume.     catch(exception)  	{         alert(exception.Description)     } }` |

#### **CustomXMLSchema.Reload**

从文件重新加载架构。

**语法**

**express.Reload()**

*express*   一个代表 **CustomXMLSchema** 对象的变量。

**说明**

通常，此方法用于更新架构的位置或确定架构是否仍然有效。它还用于重新加载频繁更改的架构。如果尝试对集合中的已经过验证或绑定到数据流的架构执行此操作，则不执行操作，并显示一条错误消息。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例指定了架构的位置，然后重新加载它。 function test() {     try{         let objCustomXMLSchema = Application.ActivePresentation.CustomXMLParts.Item(1).SchemaCollection.Item(1)         // Reload the schema.         bjCustomXMLSchema.Reload()     }catch(exception){ 	alert(exception.Description) 	} }` |

**成员属性**

#### **CustomXMLSchema.Application**

获取一个代表 **CustomXMLSchema** 对象的容器应用程序的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **CustomXMLSchema** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

#### **CustomXMLSchema.Creator**

获取一个 32 位整数，指示创建 **CustomXMLSchema** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

*express*   一个代表 **CustomXMLSchema** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

#### **CustomXMLSchema.Location**

获取一个 **String** 类型的值，代表计算机上架构的位置。只读。

**语法**

**express.Location**

*express*   一个代表 **CustomXMLSchema** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

#### **CustomXMLSchema.NamespaceURI**

获取 **CustomXMLSchema** 对象的命名空间的唯一地址标识符。只读。

**语法**

**express.NamespaceURI**

*express*   一个代表 **CustomXMLSchema** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

#### **CustomXMLSchema.Parent**

获取 **CustomXMLSchema** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **CustomXMLSchema** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

适用环境：web

适用平台：windows/linux