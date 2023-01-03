#### **CustomXMLPart**



代表 **CustomXMLParts** 集合中的单个 **CustomXMLPart**。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

以下示例将部件添加到 **CustomXMLPart** 对象。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `function AddPartToCollection() { let myPart = Application.ActiveDocument.CustomXMLParts.Add("Mark Twain")                  }` |

**方法**

|                                                              | 名称                 | 说明                                                         |
| ------------------------------------------------------------ | -------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddNode**          | 将节点添加到 XML 树。                                        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**           | 从数据存储（**IXMLDataStore** 接口）中删除当前 **CustomXMLPart**。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Load**             | 使模板作者可以通过现有文件填充 **CustomXMLPart**。如果加载成功，则返回 **True**。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **LoadXML**          | 允许模板作者依据 XML 字符串填充 **CustomXMLPart** 对象。如果加载成功，则返回 **True**。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SelectNodes**      | 从自定义 XML 部件中选择节点的集合。                          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SelectSingleNode** | 选择自定义 XML 部件内与 XPath 表达式匹配的单个节点。         |

**属性**

|                                                              | 名称                 | 说明                                                         |
| ------------------------------------------------------------ | -------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**      | 获取一个代表 **CustomXMLPart** 对象的容器应用程序的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BuiltIn**          | 获取一个指示 **CustomXMLPart** 是否为内置的值。只读          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**          | 获取一个 32 位整数，指示创建 **CustomXMLPart** 对象时所使用的应用程序。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DocumentElement**  | 获取文档中数据绑定区域的根元素。如果该区域为空，则该属性将返回 **Nothing**。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Errors**           | 获取一个提供对任何 XML 验证错误（如果有）的访问的 **CustomXMLValidationErrors** 对象。如果不存在验证错误，则此属性将返回 **Nothing**。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Id**               | 获取一个 **String** 类型的值，其中包含分配给当前 **CustomXMLPart** 对象的 GUID。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NamespaceManager** | 获取对当前 **CustomXMLPart** 对象使用的命名空间前缀映射集。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NamespaceURI**     | 获取 **CustomXMLPart** 对象的命名空间的唯一地址标识符。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**           | 获取 **CustomXMLPart** 对象的 **Parent** 对象。只读。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SchemaCollection** | 获取或设置一个代表附加到文档中数据绑定区域的架构集的 **CustomXMLSchemaCollection** 对象。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **XML**              | 获取当前 **CustomXMLPart** 对象的 XML 表示形式。只读。       |

**成员方法**

#### **CustomXMLPart.AddNode**

将节点添加到 XML 树。

**语法**

**express.AddNode(Parent, Name, NamespaceURI, NextSibling, NodeType, NodeValue)**

*express*   一个代表 **CustomXMLPart** 对象的变量。

**参数**

| **名称**       | **必选/可选** | **数据类型**             | **说明**                                                     |
| -------------- | ------------- | ------------------------ | ------------------------------------------------------------ |
| *Parent*       | 必选          | **CustomXMLNode**        | 代表应该在其下添加此节点的节点。如果添加属性，则该参数指示应该向其添加属性的元素。 |
| *Name*         | 可选          | **String**               | 代表要添加的节点的基本名称。                                 |
| *NamespaceURI* | 可选          | **String**               | 代表要追加的元素的命名空间。如果要追加 msoCustomXMLNodeElement 或 msoCustomXMLNodeAttribute 类型的节点，则此参数是必需的，否则它将被忽略。 |
| *NextSibling*  | 可选          | **CustomXMLNode**        | 代表应该成为新节点的下一个同级项的节点。如果未指定此参数，则将该节点添加到父节点的子节点的末尾。如果添加 msoXMLNodeAttribute 类型的节点，则忽略此参数。如果该节点不是父节点的子节点，则显示一个错误。 |
| *NodeType*     | 可选          | **MsoCustomXMLNodeType** | 指定要追加的节点的类型。如果未指定该参数，则假定节点类型为 msoCustomXMLNodeElement。 |
| *NodeValue*    | 可选          | **String**               | 用于为允许文本的那些节点设置所追加节点的值。如果节点不允许文本，则忽略该参数。 |

**说明**

如果 **AddNode** 操作将导致树结构无效，则不执行追加操作，并显示一条错误消息。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例说明了如何将节点添加到 CustomXMLPart 对象。  function test() {     let cxp1     let cxn     let ad = Application.ActiveDocument     // Add and populate a custom xml part     cxp1 = ad.CustomXMLParts.Add("<invoice />")              // Set the parent node      cxn = cxp1.SelectSingleNode("/invoice")              // Add a node under the parent node     cxp1.AddNode(cxn, "upccode", "urn:invoice:namespace") }` |

#### **CustomXMLPart.Delete**

从数据存储（**IXMLDataStore** 接口）中删除当前 **CustomXMLPart**。

**语法**

**express.Delete()**

*express*   一个代表 **CustomXMLPart** 对象的变量。

**说明**

如果尝试删除包含核心属性的部件，则不会执行操作并显示一条错误消息。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例将添加一个自定义 XML 部件、使用标准选择节点，并删除部件和节点。 function test() {     let ad = Application.ActiveDocument     // Example written for Word.     // Add a custom XML part and then load the XML from a file.     let cxp1 = ad.CustomXMLParts.Add()                        // Delete custom XML part.     cxp1.Delete() }` |

#### **CustomXMLPart.Load**

使模板作者可以通过现有文件填充 **CustomXMLPart**。如果加载成功，则返回 **True**。

**语法**

**express.Load(FilePath)**

*express*   一个代表 **CustomXMLPart** 对象的变量。

**参数**

| **名称**   | **必选/可选** | **数据类型** | **说明**                                            |
| ---------- | ------------- | ------------ | --------------------------------------------------- |
| *FilePath* | 必选          | **String**   | 指向用户计算机上或包含要加载的 XML 的网络上的文件。 |

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例将添加一个自定义 XML 部件、使用标准选择节点，并删除部件和节点。 function test() {     try      {         let ad = Application.ActiveDocument         // Example written for Word.         // Add a custom XML part and then load the XML from a file.         let cxp1 = ad.CustomXMLParts.Add()         cxp1.Load("c:\\invoice.xml") 	             let cxn = cxp1.SelectSingleNode("//*[@quantity < 4]")          // Insert a subtree before the single node selected previously.         cxn.InsertSubTreeBefore("<discounts><discount>0.10</discount></discounts>")                              // Delete custom XML part.         cxp1.Delete()         cxn.Delete()     }                      // Exception handling. Show the message and resume.     catch(Err)  	{         alert(Err.Description)     } }` |

#### **CustomXMLPart.LoadXML**

允许模板作者依据 XML 字符串填充 **CustomXMLPart** 对象。如果加载成功，则返回 **True**。

**语法**

**express.LoadXML(XML)**

*express*   一个代表 **CustomXMLPart** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**           |
| -------- | ------------- | ------------ | ------------------ |
| *XML*    | 必选          | **String**   | 包含要加载的 XML。 |

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例通过字符串将 XML 加载到自定义 XML 部件中。 function test() {     try      {         // Add a custom XML part and then load the XML.         let cxp1 = Application.ActiveDocument.CustomXMLParts.Add()         cxp1.LoadXML("<discounts><discount>0.10</discount></discounts>")     }                      // Exception handling. Show the message and resume.     catch(Err)     {         alert(Err.Description)     } }` |

#### **CustomXMLPart.SelectNodes**

从自定义 XML 部件中选择节点的集合。

**语法**

**express.SelectNodes(XPath)**

*express*   一个代表 **CustomXMLPart** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**            |
| -------- | ------------- | ------------ | ------------------- |
| *XPath*  | 必选          | **String**   | 包含 XPath 表达式。 |

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例说明了添加自定义 XML 部件，选择与命名空间 URI 匹配的部件，然后在该部件内选择与 XPath 表达式匹配的节点。 function test() {     let cxp1     let cxn          // Add a custom xml part.     Application.ActiveDocument.CustomXMLParts.Add()          // Return the first custom xml part with the given index.     cxp1 = Application.ActiveDocument.CustomXMLParts.Item(1)           // Get all of the nodes matching an XPath expression.     cxns = cxp1.SelectNodes("//*[@unitPrice > 20]")  }` |

#### **CustomXMLPart.SelectSingleNode**

选择自定义 XML 部件内与 XPath 表达式匹配的单个节点。

**语法**

**express.SelectSingleNode(XPath)**

*express*   一个代表 **CustomXMLPart** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                |
| -------- | ------------- | ------------ | ----------------------- |
| *XPath*  | 必选          | **String**   | 包含一个 XPath 表达式。 |

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例说明了添加自定义 XML 部件，选择与命名空间 URI 匹配的部件，然后在该部件内选择与 XPath 表达式匹配的节点。 function test() {     let cxp1     let cxn     // Add a custom xml part.     Application.ActiveDocument.CustomXMLParts.Add()          // Return the first custom xml part with the given index.     cxp1 = Application.ActiveDocument.CustomXMLParts.Item(1)                  // Get a node using XPath.                                  cxn = cxp1.SelectSingleNode("//*[@supplierID = 1]") }` |

**成员属性**

#### **CustomXMLPart.Application**

获取一个代表 **CustomXMLPart** 对象的容器应用程序的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **CustomXMLPart** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

#### **CustomXMLPart.BuiltIn**

获取一个指示 **CustomXMLPart** 是否为内置的值。只读

**语法**

**express.BuiltIn**

*express*   一个代表 **CustomXMLPart** 对象的变量。

#### **CustomXMLPart.Creator**

获取一个 32 位整数，指示创建 **CustomXMLPart** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

*express*   一个代表 **CustomXMLPart** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

#### **CustomXMLPart.DocumentElement**

获取文档中数据绑定区域的根元素。如果该区域为空，则该属性将返回 **Nothing**。只读。

**语法**

**express.DocumentElement**

*express*   一个代表 **CustomXMLPart** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

#### **CustomXMLPart.Errors**

获取一个提供对任何 XML 验证错误（如果有）的访问的 **CustomXMLValidationErrors** 对象。如果不存在验证错误，则此属性将返回 **Nothing**。只读。

**语法**

**express.Errors**

*express*   一个代表 **CustomXMLPart** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

#### **CustomXMLPart.Id**

获取一个 **String** 类型的值，其中包含分配给当前 **CustomXMLPart** 对象的 GUID。只读。

**语法**

**express.Id**

*express*   一个代表 **CustomXMLPart** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

#### **CustomXMLPart.NamespaceManager**

获取对当前 **CustomXMLPart** 对象使用的命名空间前缀映射集。只读。

**语法**

**express.NamespaceManager**

*express*   一个代表 **CustomXMLPart** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

#### **CustomXMLPart.NamespaceURI**

获取 **CustomXMLPart** 对象的命名空间的唯一地址标识符。只读。

**语法**

**express.NamespaceURI**

*express*   一个代表 **CustomXMLPart** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

#### **CustomXMLPart.Parent**

获取 **CustomXMLPart** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **CustomXMLPart** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

#### **CustomXMLPart.SchemaCollection**

获取或设置一个代表附加到文档中数据绑定区域的架构集的 **CustomXMLSchemaCollection** 对象。可读/写。

**语法**

**express.SchemaCollection**

*express*   一个代表 **CustomXMLPart** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

#### **CustomXMLPart.XML**

获取当前 **CustomXMLPart** 对象的 XML 表示形式。只读。

**语法**

**express.XML**

*express*   一个代表 **CustomXMLPart** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

适用环境：web

适用平台：windows/linux