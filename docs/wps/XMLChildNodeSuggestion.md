#### **XMLChildNodeSuggestion**



表示根据架构可能作为当前元素子元素的节点，但不保证其有效性。

**说明**

在用户将元素插入文档之前，WPS 无法验证元素的显示顺序（或出现次数的上限或下限或其他约束）。如果选中**“仅列出当前元素的子元素”**复选框，则每个 **XMLChildNodeSuggestion** 对象都是**“XML 结构”**任务窗格底部 XML 元素列表中的项目。

使用 **XMLSchemaReference** 属性可访问节点所引用的单个架构。使用 **Item** 方法可从 **XMLChildNodeSuggestion** 对象的集合中返回一个 **XMLChildNodeSuggestion** 对象。以下代码将所有允许的元素插入活动文档中。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() { let objNode = Selection.XMLParentNode   for(let i = 1; i <= objNode.ChildNodeSuggestions.Count; i++){     objNode.ChildNodeSuggestions.Item(i).Insert()     Selection.MoveRight() } }` |

使用 **NamespaceURI** 属性可返回 **XMLChildNodeSuggestion** 对象中引用的架构的命名空间。使用 **BaseName** 属性可返回代表 **XMLChildNodeSuggestion** 对象的元素的名称。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() { let objSuggestion = ActiveDocument.ChildNodeSuggestions      for(let i = 1; i <= objSuggestion.Count; i++){              if(objSuggestion.Item(i).BaseName == "example"){         objSuggestion.Item(i).Insert()     }              } }` |

**方法**

|                                                              | 名称       | 说明                                       |
| ------------------------------------------------------------ | ---------- | ------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Insert** | 插入代表 XML 元素节点的 **XMLNode** 对象。 |

**属性**

|                                                              | 名称                   | 说明                                                         |
| ------------------------------------------------------------ | ---------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**        | 返回一个代表 WPS 应用程序的 **Application** 对象。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BaseName**           | 返回一个 **String** 类型的值，该值代表不带任何前缀的元素名称。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**            | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NamespaceURI**       | 返回一个 **String** 类型的值，该值代表指定对象的架构命名空间的统一资源标识符 (URI)。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**             | 返回一个 **Object** 类型值，该值代表指定 **XMLChildNodeSuggestion** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **XMLSchemaReference** | 返回一个 **XMLSchemaReference** 对象，代表指定的 **XMLChildNodeSuggestion** 对象所属的 XML 架构。 |

**成员方法**

#### **XMLChildNodeSuggestion.Insert**

插入代表 XML 元素节点的 **XMLNode** 对象。

**语法**

**express.Insert(Range)**

*express*   一个代表 **XMLChildNodeSuggestion** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                  |
| -------- | ------------- | ------------ | ----------------------------------------- |
| *Range*  | 可选          | **Variant**  | 环绕其放置打开和关闭 XML 元素的文字范围。 |

**成员属性**

#### **XMLChildNodeSuggestion.Application**

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

*express*   一个代表 **XMLChildNodeSuggestion** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **XMLChildNodeSuggestion.BaseName**

返回一个 **String** 类型的值，该值代表不带任何前缀的元素名称。

**语法**

**express.BaseName**

*express*   一个代表 **XMLChildNodeSuggestion** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例将 author 属性添加到活动文档中的 book 元素，然后设置属性的值。*/ function AddIDAttribute(){     let objElement = ActiveDocument.XMLNodes      for(let i = 1; i <= objElement.Count; i++){         if(objElement.Item(i).NodeType == wdXMLNodeElement){ 	    if(objElement.Item(i).BaseName == "book"){                                  let objAttribute = objElement.Item(i).Attributes.Add("author", objElement.Item(i).NamespaceURI)                 objAttribute.NodeValue = "David Barber"                              }         }     } }` |

#### **XMLChildNodeSuggestion.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **XMLChildNodeSuggestion** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **XMLChildNodeSuggestion.NamespaceURI**

返回一个 **String** 类型的值，该值代表指定对象的架构命名空间的统一资源标识符 (URI)。只读。

**语法**

**express.NamespaceURI**

*express*   一个代表 **XMLChildNodeSuggestion** 对象的变量。

**说明**

| 注释                                                         |
| ------------------------------------------------------------ |
| 如果创建的是用于 WPS 的 XML 架构，则强烈建议在架构中指定 targetNamespace 设置。 |

#### **XMLChildNodeSuggestion.Parent**

返回一个 **Object** 类型值，该值代表指定 **XMLChildNodeSuggestion** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **XMLChildNodeSuggestion** 对象的变量。

**说明**

------

#### **XMLChildNodeSuggestion.XMLSchemaReference**

返回一个 **XMLSchemaReference** 对象，代表指定的 **XMLChildNodeSuggestion** 对象所属的 XML 架构。

**语法**

**express.XMLSchemaReference**

*express*   一个代表 **XMLChildNodeSuggestion** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下列示例为活动文档中的第一个 XMLChildNodeSuggestion 对象重新加载架构。*/ function test() { let objSuggestion = ActiveDocument.XMLChildNodeSuggestions.Item(1) objSuggestion.XMLSchemaReference.Reload() }  ` |

适用环境：web

适用平台：windows/linux