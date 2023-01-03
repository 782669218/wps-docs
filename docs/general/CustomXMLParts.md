#### **CustomXMLParts**



代表 **CustomXMLPart** 对象的集合。

**说明**

存在三个始终与文档一起创建的默认部件。它们是“封面”、“文档属性”和“应用程序属性”。最后两个部件在 WPS Office 的以前版本中，但现在是在 **CustomXMLParts** 对象集合中以 XML 形式提供

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例将一个节点添加到 CustomXMLPart 对象（它是 CustomXMLParts 对象集合的一部分）。 let myPart = Application.ActiveDocument.CustomXMLParts.Add("<author>Mark Twain</author>") ` |

**方法**

|                                                              | 名称                  | 说明                                                  |
| ------------------------------------------------------------ | --------------------- | ----------------------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**               | 允许您向文件添加新的 **CustomXMLPart**。              |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SelectByID**        | 选择与 GUID 匹配的自定义 XML 部件。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SelectByNamespace** | 选择其命名空间与搜索条件匹配的自定义 XML 部件的集合。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 获取一个代表 **CustomXMLParts** 对象的容器应用程序的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 获取一个 **Long** 类型的值，指示 **CustomXMLParts** 集合中的项数。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 获取一个 32 位整数，指示创建 **CustomXMLParts** 对象时所使用应用程序。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Item**        | 获取 **CustomXMLParts** 集合中的一个 **CustomXMLPart** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 获取 **CustomXMLParts** 对象的 **Parent** 对象。只读。       |

**成员方法**

#### **CustomXMLParts.Add**

允许您向文件添加新的 **CustomXMLPart**。

**语法**

**express.Add(XML, SchemaCollection)**

*express*   一个代表 **CustomXMLParts** 对象的变量。

**参数**

| **名称**           | **必选/可选** | **数据类型**                  | **说明**                                      |
| ------------------ | ------------- | ----------------------------- | --------------------------------------------- |
| *XML*              | 可选          | **String**                    | 包含要添加到新创建的 CustomXMLPart 中的 XML。 |
| *SchemaCollection* | 可选          | **CustomXMLSchemaCollection** | 代表要用于验证此流的架构集。                  |

**返回值**

CustomXMLPart

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例将添加新的 CustomXMLPart，使用搜索标准选择 CustomXMLPart，然后从该部件中选择单一节点。需要有对应命名空间的custom XML part function test() {     try {     let cxp1     let cxn      // Example written for Word.      // Add a custom XML part.     Application.ActiveDocument.CustomXMLParts.Add("<custXMLPart />")      // Returns the first custom XML part with the given root namespace.     cxp1 = Application.ActiveDocument.CustomXMLParts.Item("urn:invoice:namespace")              // Get a node using XPath.                                  cxn = cxp1.SelectSingleNode("//*[@quantity < 4]")     }          // Exception handling. Show the message and resume.     catch(Err)  	{         alert(Err.Description)     }  }` |

#### **CustomXMLParts.SelectByID**

选择与 GUID 匹配的自定义 XML 部件。

**语法**

**express.SelectByID(Id)**

*express*   一个代表 **CustomXMLParts** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                     |
| -------- | ------------- | ------------ | ---------------------------- |
| *Id*     | 必选          | **String**   | 包含自定义 XML 部件的 GUID。 |

**说明**

如果具有此 ID 的自定义 XML 部件不存在，则该方法返回 **Nothing**。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例说明了检索与 GUID 匹配的自定义 XML 部件，然后在该部件中搜索与 XPath 表达式匹配的节点。 function test() {     // Returns a custom xml part by its ID.Get the node matching the XPath expression.     let cxp1 = Application.ActivePresentation.CustomXMLParts.SelectByID("{6C3C8BC8-F283-45AE-878A-BAB7291924A1}")          // Get the node matching the XPath expression.                                  let cxn = cxp1.SelectSingleNode("//*[@supplierID = 1]") }` |

#### **CustomXMLParts.SelectByNamespace**

选择其命名空间与搜索条件匹配的自定义 XML 部件的集合。

**语法**

**express.SelectByNamespace(NamespaceURI)**

*express*   一个代表 **CustomXMLParts** 对象的变量。

**参数**

| **名称**       | **必选/可选** | **数据类型** | **说明**           |
| -------------- | ------------- | ------------ | ------------------ |
| *NamespaceURI* | 必选          | **String**   | 包含命名空间 URI。 |

**说明**

如果具有此命名空间的自定义 XML 部件不存在，则该方法将返回空的 **CustomXMLParts** 集合对象。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例说明了检索与 GUID 匹配的自定义 XML 部件，然后在该部件中搜索与 XPath 表达式匹配的节点。 function test() {     // Returns all of the custom xml parts with the given namespace.     let cxp1 = Application.ActiveDocument.CustomXMLParts.SelectByNamespace("urn:invoice:namespace")                    // Get the node matching the XPath expression.                                  let cxn = cxp1.SelectSingleNode("//*[@supplierID = 1]") }` |

**成员属性**

#### **CustomXMLParts.Application**

获取一个代表 **CustomXMLParts** 对象的容器应用程序的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **CustomXMLParts** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

#### **CustomXMLParts.Count**

获取一个 **Long** 类型的值，指示 **CustomXMLParts** 集合中的项数。只读。

**语法**

**express.Count**

*express*   一个代表 **CustomXMLParts** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

#### **CustomXMLParts.Creator**

获取一个 32 位整数，指示创建 **CustomXMLParts** 对象时所使用应用程序。只读。

**语法**

**express.Creator**

*express*   一个代表 **CustomXMLParts** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

#### **CustomXMLParts.Item**

获取 **CustomXMLParts** 集合中的一个 **CustomXMLPart** 对象。只读。

**语法**

**express.Item**

*express*   一个代表 **CustomXMLParts** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

#### **CustomXMLParts.Parent**

获取 **CustomXMLParts** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **CustomXMLParts** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

适用环境：web

适用平台：windows/linux