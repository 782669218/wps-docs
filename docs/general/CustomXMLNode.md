#### **CustomXMLNode**



代表文档中树的一个 XML 节点。**CustomXMLNode** 对象是 **CustomXMLNodes** 集合的一个成员。

**说明**

**CustomXMLNode** 对象设计为具有与 **IXMLDOMNode** 接口等效的功能。此外，它包含 **XPath** 属性，这是对 MSXML 提供的对象的重大改进。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

以下示例通过使用 XPath 表达式从 **CustomXMLPart** 对象中选择单个节点，并将它分配给 **CustomXMLNode** 对象。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `function test() { let ment = Application.ActiveDocument // Returns the first custom xml part with the given root namespace. let cxp1 = ment.CustomXMLParts.Item("urn:invoice:namespace")      // Get the first node matching the XPath expression.                              let cxn = cxp1.SelectSingleNode("//*[@quantity < 4]") }` |

**方法**

|                                                              | 名称                    | 说明                                                         |
| ------------------------------------------------------------ | ----------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AppendChildNode**     | 将单个节点作为树中上下文元素节点下的最后一个子节点追加。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AppendChildSubtree**  | 将子树作为树中上下文元素节点下的最后一个子节点添加。         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**              | 从树中删除当前节点（包括其所有子项，如果有）。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **HasChildNodes**       | 如果当前元素节点包含子元素节点，则获取 **True**。            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **InsertNodeBefore**    | 紧邻树中上下文节点之前插入一个新节点。                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **InsertSubtreeBefore** | 将指定的子树插入到紧邻上下文节点之前的位置。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **RemoveChild**         | 从树中删除指定的子节点。                                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ReplaceChildNode**    | 从主树中删除指定的子节点（及其子树），并将它替换为相同位置中的其他节点。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ReplaceChildSubtree** | 从主树中删除指定节点（及其子树），并将它替换为相同位置中的其他子树。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SelectNodes**         | 选择与 XPath 表达式匹配的节点的集合。此方法与 **CustomXMLPart**.**SelectNodes** 方法的不同之处在于，它将从作为上下文节点的“表达式”节点开始计算 XPath 表达式。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SelectSingleNode**    | 从与 XPath 表达式匹配的集合中选择一个节点。此方法与 **CustomXMLPart**.**SelectSingleNode** 方法的不同之处在于，它将从作为上下文节点的“表达式”节点开始计算 XPath 表达式。 |

**属性**

|                                                              | 名称                | 说明                                                         |
| ------------------------------------------------------------ | ------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**     | 获取一个代表 **CustomXMLNode** 的容器应用程序的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Attributes**      | 获取一个代表当前节点中的当前元素的属性的 **CustomXMLNodes** 集合。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BaseName**        | 获取文档对象模型 (DOM) 中节点的不带命名空间前缀的基本名称（如果存在）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ChildNodes**      | 获取一个包含当前节点中所有子元素的 **CustomXMLNodes** 集合。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**         | 获取一个 32 位整数，指示创建 **CustomXMLNode** 对象时所使用的应用程序。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FirstChild**      | 获取与当前节点中第一个子元素对应的 **CustomXMLNode** 对象。如果该节点没有子元素（或者其类型不为 **msoCustomXMLNodeElement**），则返回 **Nothing**。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LastChild**       | 获取与当前节点中的最后一个子元素对应的 **CustomXMLNode** 对象。如果该节点没有子元素（或者其类型不为 **msoCustomXMLNodeElement**），则该属性将返回 **Nothing**。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NamespaceURI**    | 获取 **CustomXMLNode** 对象的命名空间的唯一地址标识符。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NextSibling**     | 获取当前节点的下一个同级节点（元素、注释或处理指令）。如果该节点是其所在级别的最后一个同级节点，则该属性将返回 **Nothing**。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NodeType**        | 获取当前节点的类型。只读。                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NodeValue**       | 获取或设置当前节点的值。可读/写。                            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OwnerDocument**   | 获取代表与此节点关联的 ET 工作簿、Microsoft PowerPoint 演示文稿或 WPS 文档的对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OwnerPart**       | 获取代表与此节点关联的部件的对象。只读。                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**          | 获取 **CustomXMLNode** 对象的 **Parent** 对象。只读。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ParentNode**      | 获取当前节点的父元素节点。如果当前节点位于根级，则该属性将返回 **Nothing**。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PreviousSibling** | 获取当前节点的前一个同级节点（元素、注释或处理指令）。如果当前节点为其所在级别的第一个节点，则该属性将返回 **Nothing**。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Text**            | 获取或设置当前节点的文本。可读/写。                          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **XML**             | 获取当前节点及其子项（如果存在）的 XML 表示形式。只读。      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **XPath**           | 获取一个 **String** 类型的值，其中包含当前节点的规范化 XPath。如果该节点已不在文档对象模型 (DOM) 中，则该属性将返回一条错误消息。只读。 |

**成员方法**

#### **CustomXMLNode.AppendChildNode**

将单个节点作为树中上下文元素节点下的最后一个子节点追加。

**语法**

**express.AppendChildNode(Name, NamespaceURI, NodeType, NodeValue)**

*express*   一个代表 **CustomXMLNode** 对象的变量。

**参数**

| **名称**       | **必选/可选** | **数据类型**             | **说明**                                                     |
| -------------- | ------------- | ------------------------ | ------------------------------------------------------------ |
| *Name*         | 可选          | **String**               | 代表要追加的元素的基本名称。                                 |
| *NamespaceURI* | 可选          | **String**               | 代表要追加的元素的命名空间。如果要追加 msoCustomXMLNodeElement 或 msoCustomXMLNodeAttribute 类型的节点，则此参数是必需的，否则它将被忽略。 |
| *NodeType*     | 可选          | **MsoCustomXMLNodeType** | 指定要追加的节点的类型。如果未指定该参数，则假定节点类型为 msoCustomXMLNodeElement。 |
| *NodeValue*    | 可选          | **String**               | 用于为允许文本的那些节点设置所追加节点的值。如果节点不允许文本，则忽略该参数。 |

**说明**

如果上下文节点是除 **msoXMLNodeElement** 之外的任何类型，或者如果操作将导致树结构无效，则不执行追加操作，并显示一条错误消息。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例说明了如何将 CustomXMLNode 对象追加到其他节点。  function test() {                       let cxn = Application.ActivePresentation.CustomXMLParts.Item(1).SelectSingleNode("/ns1:coreProperties[1]")      // Append a child node to the single node selected previously.     cxn.AppendChildNode("discount", "urn:invoice:namespace", Application.Enum.msoCustomXMLNodeElement, "0.10")  }` |

#### **CustomXMLNode.AppendChildSubtree**

将子树作为树中上下文元素节点下的最后一个子节点添加。

**语法**

**express.AppendChildSubtree(XML)**

*express*   一个代表 **CustomXMLNode** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**           |
| -------- | ------------- | ------------ | ------------------ |
| *XML*    | 必选          | **String**   | 代表要添加的子树。 |

**说明**

如果上下文节点是除 **msoXMLNodeElement** 之外的任何类型，则不执行追加操作，并显示一条错误消息。如果根据架构验证 CustomXMLNode，但该操作将导致树结构无效，则不执行追加操作，并显示一条错误消息。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例说明了如何将节点追加到现有节点。 function test() {     let tion = Application.ActiveDocument     // Add and populate a custom xml part     let cxp1 = tion.CustomXMLParts.Add("<invoice />")             // Get nodes using XPath.                                  let cxn = cxp1.SelectSingleNode("/ns1:coreProperties[1]")      // Append a child subtree to the single node selected previously.     cxn.AppendChildSubtree("0.10")   }` |

#### **CustomXMLNode.Delete**

从树中删除当前节点（包括其所有子项，如果有）。

**语法**

**express.Delete()**

*express*   一个代表 **CustomXMLNode** 对象的变量。

**说明**

如果操作将导致树结构无效，则不会执行删除操作并显示一条错误消息。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例演示了使用各种方法添加自定义 XML 部件、使用不同的标准选择部件和节点、附加子树，以及删除部件和节点的过程。 function test() {     try  	{         let tion = Application.ActiveDocument         // Example written for Word.         // Adding a custom XML part.         tion.CustomXMLParts.Add("")                 // Add and then load from a file.         let cxp1 = tion.CustomXMLParts.Add()         // Returns the first custom XML part with the given root namespace.         let cxp2 = tion.CustomXMLParts.Item(1)                     // Access all with a given root namespace; returns the entire collection.         let cxps = tion.CustomXMLParts.SelectByNamespace("urn:invoice:namespace")               // DOM-type operations.         // Get the XML.         strXml = cxp2.XML         // Get the root namespace         strUri = cxp2.NamespaceURI          // Get nodes using XPath.                                      let cxn = cxp2.SelectSingleNode("//*[@quantity < 4]")          let cxns = cxp2.SelectNodes("//*[@unitPrice > 20]")         // Append a child subtree to the single node selected previously.         cxn.AppendChildSubtree("<rebates><rebate>0.10</rebate></rebates>")                         // Delete custom XML part and node and its children.         cxp2.Delete()         cxn.Delete()     }     // Exception handling. Show the message and resume.    catch(exception)     {        alert(exception.Description)    }   } ` |

#### **CustomXMLNode.HasChildNodes**

如果当前元素节点包含子元素节点，则获取 **True**。

**语法**

**express.HasChildNodes()**

*express*   一个代表 **CustomXMLNode** 对象的变量。

**说明**

当 **CustomXMLNode** 的节点类型不是 **msoCustomXMLNodeElement** 时，此方法将始终返回 **False**。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例演示了使用各种方法添加自定义 XML 部件、使用不同的标准选择部件和节点、附加子树，以及删除部件和节点的过程。 function test() {     try      {         let tion = Application.ActiveDocument         // Example written for Word.         // Adding a custom XML part.         tion.CustomXMLParts.Add("")                 // Add and then load from a file.         let cxp1 = tion.CustomXMLParts.Add()            // Returns the first custom XML part with the given root namespace.         let cxp2 = tion.CustomXMLParts.Item(1)                     // Access all with a given root namespace; returns the entire collection.         let cxps = tion.CustomXMLParts.SelectByNamespace("urn:invoice:namespace")               // DOM-type operations.         // Get the XML.         strXml = cxp2.XML         // Get the root namespace         strUri = cxp2.NamespaceURI          // Get nodes using XPath.                                      let cxn = cxp2.SelectSingleNode("//*[@quantity < 4]")          let cxns = cxp2.SelectNodes("//*[@unitPrice > 20]")         // Append a child subtree to the single node selected previously.         cxn.AppendChildSubtree("<rebates><rebate>0.10</rebate></rebates>")                         // Delete custom XML part and node and its children.         cxp2.Delete()         cxn.Delete()     }     // Exception handling. Show the message and resume.    catch(exception)     {        alert(exception.Description)    }   } ` |

#### **CustomXMLNode.InsertNodeBefore**

紧邻树中上下文节点之前插入一个新节点。

**语法**

**express.InsertNodeBefore(Name, NamespaceURI, NodeType, NodeValue, NextSibling)**

*express*   一个代表 **CustomXMLNode** 对象的变量。

**参数**

| **名称**       | **必选/可选** | **数据类型**             | **说明**                                                     |
| -------------- | ------------- | ------------------------ | ------------------------------------------------------------ |
| *Name*         | 可选          | **String**               | 代表要添加的节点的基本名称。                                 |
| *NamespaceURI* | 可选          | **String**               | 代表要添加的元素的命名空间。如果要添加 msoCustomXMLNodeElement 或 msoCustomXMLNodeAttribute 类型的节点，则此参数是必需的，否则它将被忽略。 |
| *NodeType*     | 可选          | **MsoCustomXMLNodeType** | 指定要添加的节点的类型。如果未指定参数，则假定它是 msoCustomXMLNodeElement 类型的节点。 |
| *NodeValue*    | 可选          | **String**               | 用于为允许文本的那些节点设置要添加节点的值。如果节点不允许文本，则忽略该参数。 |
| *NextSibling*  | 可选          | **CustomXMLNode**        | 代表上下文节点。                                             |

**说明**

如果在添加 **msoCustomXMLNodeElement**、**msoCustomXMLNodeComment** 或 **msoCustomXMLNodeProcessingInstruction** 类型的节点时上下文节点不存在，则将新节点添加到上下文节点的最后一个子节点。如果操作将导致树结构无效，则不执行插入操作，并显示一条错误消息。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例添加一个自定义部件，然后通过使用 XPath 表达式查找该部件中的一个节点。代码然后在找到的节点之前插入一个节点。 function test() {     let tion = Application.ActiveDocument     // Add a custom xml part.     tion.CustomXMLParts.Add("<invoice>")             // Returns the first custom xml part with the given root namespace.     let cxp1 = tion.CustomXMLParts.Item("urn:invoice:namespace")                  // Get node using XPath.                                  let cxn = cxp1.SelectSingleNode("//*[@supplier = "Contoso"]")      // Insert a node before the single node selected previously.     cxn.InsertNodeBefore("discount", "urn:invoice:namespace") ` |

#### **CustomXMLNode.InsertSubtreeBefore**

将指定的子树插入到紧邻上下文节点之前的位置。

**语法**

**express.InsertSubtreeBefore(XML, NextSibling)**

*express*   一个代表 **CustomXMLNode** 对象的变量。

**参数**

| **名称**      | **必选/可选** | **数据类型**      | **说明**           |
| ------------- | ------------- | ----------------- | ------------------ |
| *XML*         | 必选          | **String**        | 代表要添加的子树。 |
| *NextSibling* | 可选          | **CustomXMLNode** | 指定上下文节点。   |

**说明**

如果 *NextSibling* 参数不是上下文节点的子节点，或者操作将导致树结构无效，则不执行插入操作，并显示一条错误消息。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例添加一个自定义部件，然后通过使用 XPath 表达式查找该部件中的一个节点。代码然后在找到的节点之后插入一个节点。 function test() {     let tion = Application.ActiveDocument     // Add a custom xml part.     tion.CustomXMLParts.Add("<invoice>"）            // Returns the first custom xml part with the given root namespace.     let cxp1 = tion.CustomXMLParts.Item("urn:invoice:namespace")                  // Get nodes using XPath.                                  Set cxn = cxp1.SelectSingleNode("//*[@supplier = "Contoso"]")      // Insert a node before the single node selected previously.     cxn.InsertNodeAfter("discount", "urn:invoice:namespace") }` |

#### **CustomXMLNode.RemoveChild**

从树中删除指定的子节点。

**语法**

**express.RemoveChild(Child)**

*express*   一个代表 **CustomXMLNode** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型**      | **说明**                 |
| -------- | ------------- | ----------------- | ------------------------ |
| *Child*  | 必选          | **CustomXMLNode** | 代表上下文节点的子节点。 |

**说明**

如果在 *Child* 参数中指定的节点不是上下文节点的子节点，或者操作将导致树无效，则不执行删除操作，并显示一条错误消息

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例选择一个自定义部件，然后选择该部件中的一个节点。代码然后删除该节点的子节点。 function test() {     let tion = Application.ActiveDocument     // Return the first part with the given root namespace.     let cxp1 = tion.CustomXMLParts.Item("urn:invoice:namespace")              // Get node using XPath expression.                                  let cxn = cxp1.SelectSingleNode("//*[@supplierID = 1]")      // Remove a child node.     cxn.RemoveChild(cxn.SelectSingleNode("//discount"))  }` |

#### **CustomXMLNode.ReplaceChildNode**

从主树中删除指定的子节点（及其子树），并将它替换为相同位置中的其他节点。

**语法**

**express.ReplaceChildNode(OldNode, Name, NamespaceURI, NodeType, NodeValue)**

*express*   一个代表 **CustomXMLNode** 对象的变量。

**参数**

| **名称**       | **必选/可选** | **数据类型**             | **说明**                                                     |
| -------------- | ------------- | ------------------------ | ------------------------------------------------------------ |
| *OldNode*      | 必选          | **CustomXMLNode**        | 代表要替换的子节点。                                         |
| *Name*         | 可选          | **String**               | 代表要添加的元素的基本名称。                                 |
| *NamespaceURI* | 可选          | **String**               | 代表要添加的元素的命名空间。如果要添加 msoCustomXMLNodeElement 或 msoCustomXMLNodeAttribute 类型的节点，则此参数是必需的，否则它将被忽略。 |
| *NodeType*     | 可选          | **MsoCustomXMLNodeType** | 指定要添加的节点的类型。如果未指定参数，则假定节点类型为 msoCustomXMLNodeElement。 |
| *NodeValue*    | 可选          | **String**               | 用于为允许文本的那些节点设置要添加的节点值。如果节点不允许文本，则忽略该参数。 |

**说明**

如果 *OldNode* 参数不是上下文节点的子节点，或者操作将导致树结构无效，则不执行替换操作，并显示一条错误消息。此外，如果要添加的节点已存在，则不执行替换操作，并显示一条错误消息。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例选择一个自定义部件，然后选择该部件中的一个节点。代码然后将该节点的子节点替换为其他节点。 function test() {     let tion = Application.ActiveDocument     // Return the first custom xml part with the given root namespace.     let cxp1 = tion.CustomXMLParts.Item("urn:invoice:namespace")                                      let cxn = cxp1.SelectSingleNode("//*[@supplierID = 1]")      // Replace a child node.     cxn.ReplaceChildNode(cxn.SelectSingleNode("//discount", "rebate") }` |

#### **CustomXMLNode.ReplaceChildSubtree**

从主树中删除指定节点（及其子树），并将它替换为相同位置中的其他子树。

**语法**

**express.ReplaceChildSubtree(XML, OldNode)**

*express*   一个代表 **CustomXMLNode** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型**      | **说明**             |
| --------- | ------------- | ----------------- | -------------------- |
| *XML*     | 必选          | **String**        | 代表要添加的子树。   |
| *OldNode* | 必选          | **CustomXMLNode** | 代表要替换的子节点。 |

**说明**

如果操作将导致树结构无效，则不执行替换操作，并显示一条错误消息。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例选择一个自定义部件，然后选择该部件中的一个节点。代码然后将该节点的子子树替换为其他子树。 function test() {     let tion = Application.ActiveDocument     // Return the first custom xml part with the given root namespace.     let cxp1 = tion.CustomXMLParts.Item("urn:invoice:namespace")                // Get node using XPath expression.                                  let cxn = cxp1.SelectSingleNode("//*[@supplierID = 1]")      // Replace one subtree and its children with another.     cxn.ReplaceChildSubtree("<rebates><rebate>0.10</rebate></rebates>", "//discounts") }` |

#### **CustomXMLNode.SelectNodes**

选择与 XPath 表达式匹配的节点的集合。此方法与 **CustomXMLPart**.**SelectNodes** 方法的不同之处在于，它将从作为上下文节点的“表达式”节点开始计算 XPath 表达式。

**语法**

**express.SelectNodes(XPath)**

*express*   一个代表 **CustomXMLNode** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                |
| -------- | ------------- | ------------ | ----------------------- |
| *XPath*  | 必选          | **String**   | 包含一个 XPath 表达式。 |

**返回值**

CustomXMLNodes

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例说明了添加自定义 XML 部件，选择与命名空间 URI 匹配的部件，然后在该部件内选择与 XPath 表达式匹配的节点。 function test() {     // Add a custom xml part.     Application.ActiveDocument.CustomXMLParts.Add("")     // Return the first custom xml part with the given namespace.     let cxp1 = Application.ActiveDocument.CustomXMLParts.Item("urn:invoice:namespace")      // Get all of the nodes matching an XPath expression.     let cxns = cxp1.SelectNodes("//*[@unitPrice > 20]") }` |

#### **CustomXMLNode.SelectSingleNode**

从与 XPath 表达式匹配的集合中选择一个节点。此方法与 **CustomXMLPart**.**SelectSingleNode** 方法的不同之处在于，它将从作为上下文节点的“表达式”节点开始计算 XPath 表达式。

**语法**

**express.SelectSingleNode(XPath)**

*express*   一个代表 **CustomXMLNode** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                |
| -------- | ------------- | ------------ | ----------------------- |
| *XPath*  | 必选          | **String**   | 包含一个 XPath 表达式。 |

**说明**

XPath 表达式的前缀映射是从 **NamespaceManager** 属性检索的。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例说明了添加自定义 XML 部件，选择与命名空间 URI 匹配的部件，然后在该部件内选择与 XPath 表达式匹配的节点。 function test() {     // Add a custom xml part.     Application.ActiveDocument.CustomXMLParts.Add("")     // Return the first custom xml part with the given namespace.     let cxp1 = Application.ActiveDocument.CustomXMLParts.Item("urn:invoice:namespace")             // Get a node using XPath.                                  let cxn = cxp1.Item(1).SelectSingleNode("//*[@supplierID = 1]") }` |

**成员属性**

#### **CustomXMLNode.Application**

获取一个代表 **CustomXMLNode** 的容器应用程序的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **CustomXMLNode** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

#### **CustomXMLNode.Attributes**

获取一个代表当前节点中的当前元素的属性的 **CustomXMLNodes** 集合。只读。

**语法**

**express.Attributes**

*express*   一个代表 **CustomXMLNode** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

#### **CustomXMLNode.BaseName**

获取文档对象模型 (DOM) 中节点的不带命名空间前缀的基本名称（如果存在）。只读。

**语法**

**express.BaseName**

*express*   一个代表 **CustomXMLNode** 对象的变量。

**说明**

这是 **CustomXMLNode** 对象的默认成员。

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

#### **CustomXMLNode.ChildNodes**

获取一个包含当前节点中所有子元素的 **CustomXMLNodes** 集合。只读。

**语法**

**express.ChildNodes**

*express*   一个代表 **CustomXMLNode** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

#### **CustomXMLNode.Creator**

获取一个 32 位整数，指示创建 **CustomXMLNode** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

*express*   一个代表 **CustomXMLNode** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

#### **CustomXMLNode.FirstChild**

获取与当前节点中第一个子元素对应的 **CustomXMLNode** 对象。如果该节点没有子元素（或者其类型不为 **msoCustomXMLNodeElement**），则返回 **Nothing**。只读。

**语法**

**express.FirstChild**

*express*   一个代表 **CustomXMLNode** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

#### **CustomXMLNode.LastChild**

获取与当前节点中的最后一个子元素对应的 **CustomXMLNode** 对象。如果该节点没有子元素（或者其类型不为 **msoCustomXMLNodeElement**），则该属性将返回 **Nothing**。只读。

**语法**

**express.LastChild**

*express*   一个代表 **CustomXMLNode** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

#### **CustomXMLNode.NamespaceURI**

获取 **CustomXMLNode** 对象的命名空间的唯一地址标识符。只读。

**语法**

**express.NamespaceURI**

*express*   一个代表 **CustomXMLNode** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

#### **CustomXMLNode.NextSibling**

获取当前节点的下一个同级节点（元素、注释或处理指令）。如果该节点是其所在级别的最后一个同级节点，则该属性将返回 **Nothing**。只读。

**语法**

**express.NextSibling**

*express*   一个代表 **CustomXMLNode** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

#### **CustomXMLNode.NodeType**

获取当前节点的类型。只读。

**语法**

**express.NodeType**

*express*   一个代表 **CustomXMLNode** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

#### **CustomXMLNode.NodeValue**

获取或设置当前节点的值。可读/写。

**语法**

**express.NodeValue**

*express*   一个代表 **CustomXMLNode** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常 |

#### **CustomXMLNode.OwnerDocument**

获取代表与此节点关联的 ET 工作簿、Microsoft PowerPoint 演示文稿或 WPS 文档的对象。只读。

**语法**

**express.OwnerDocument**

*express*   一个代表 **CustomXMLNode** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

#### **CustomXMLNode.OwnerPart**

获取代表与此节点关联的部件的对象。只读。

**语法**

**express.OwnerPart**

*express*   一个代表 **CustomXMLNode** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

#### **CustomXMLNode.Parent**

获取 **CustomXMLNode** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **CustomXMLNode** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

#### **CustomXMLNode.ParentNode**

获取当前节点的父元素节点。如果当前节点位于根级，则该属性将返回 **Nothing**。只读。

**语法**

**express.ParentNode**

*express*   一个代表 **CustomXMLNode** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

#### **CustomXMLNode.PreviousSibling**

获取当前节点的前一个同级节点（元素、注释或处理指令）。如果当前节点为其所在级别的第一个节点，则该属性将返回 **Nothing**。只读。

**语法**

**express.PreviousSibling**

*express*   一个代表 **CustomXMLNode** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

#### **CustomXMLNode.Text**

获取或设置当前节点的文本。可读/写。

**语法**

**express.Text**

*express*   一个代表 **CustomXMLNode** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

#### **CustomXMLNode.XML**

获取当前节点及其子项（如果存在）的 XML 表示形式。只读。

**语法**

**express.XML**

*express*   一个代表 **CustomXMLNode** 对象的变量。

**说明**

| 注释                                                         |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

#### **CustomXMLNode.XPath**

获取一个 **String** 类型的值，其中包含当前节点的规范化 XPath。如果该节点已不在文档对象模型 (DOM) 中，则该属性将返回一条错误消息。只读。

**语法**

**express.XPath**

*express*   一个代表 **CustomXMLNode** 对象的变量。

**说明**

| ![img]()注释                                                 |
| ------------------------------------------------------------ |
| 不支持从自定义 XML 部件中引用 DTD。自定义 XML 部件中的 DTD 引用将无法解析，并且在尝试将文件的内容保存到平面 XML 文件中时，包含 DTD 引用的自定义 XML 部件将产生异常。 |

适用环境：web

适用平台：windows/linux