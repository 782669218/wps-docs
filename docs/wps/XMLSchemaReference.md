#### **XMLSchemaReference**



代表附加到文档的单个架构。

**说明**

使用 **XMLSchemaReference** 属性可返回用于 **XMLSchemaReference** 对象的 **ChildNodeSuggestion** 对象。如果引用的 XML 架构是 SimpleSample 架构，则以下示例插入建议的 XML 子元素。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() { let objSuggestion = ActiveDocument.ChildNodeSuggestions for(let i =   1;  i < =  objSuggestion.Count; i++) {         if(objSuggestion.XMLSchemaReference = "SimpleSample") {         objSuggestion.Item(i).Insert()     } } }` |

**方法**

|                                                              | 名称       | 说明                              |
| ------------------------------------------------------------ | ---------- | --------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete** | 删除指定的 XML 架构引用。         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Reload** | 重新加载在文档中引用的 XML 架构。 |

**属性**

|                                                              | 名称             | 说明                                                         |
| ------------------------------------------------------------ | ---------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**  | 返回一个代表 WPS 应用程序的 **Application** 对象。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**      | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Location**     | 返回一个 **String** 类型的值，该值代表 XML 架构的物理位置。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NamespaceURI** | 返回 **String** 值，该值表示指定对象的架构命名空间的统一资源标识符 (URI)。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**       | 返回一个 **Object** 类型值，该值代表指定 **XMLSchemaReference** 对象的父对象。 |

**成员方法**

#### **XMLSchemaReference.Delete**

删除指定的 XML 架构引用。

**语法**

**express.Delete()**

*express*   一个代表 **XMLSchemaReference** 对象的变量。

#### **XMLSchemaReference.Reload**

重新加载在文档中引用的 XML 架构。

**语法**

**express.Reload()**

*express*   一个代表 **XMLSchemaReference** 对象的变量。

**返回值**

无

**成员属性**

#### **XMLSchemaReference.Application**

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

*express*   一个代表 **XMLSchemaReference** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **XMLSchemaReference.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **XMLSchemaReference** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **XMLSchemaReference.Location**

返回一个 **String** 类型的值，该值代表 XML 架构的物理位置。只读。

**语法**

**express.Location**

*express*   一个代表 **XMLSchemaReference** 对象的变量。

#### **XMLSchemaReference.NamespaceURI**

返回 **String** 值，该值表示指定对象的架构命名空间的统一资源标识符 (URI)。只读。

**语法**

**express.NamespaceURI**

*express*   一个代表 **XMLSchemaReference** 对象的变量。

**说明**

如果创建的是用于 WPS 的 XML 架构，则强烈建议在架构中指定 targetNamespace 设置。

**示例**

以下示例重新加载 SimpleSample 架构，如果该架构没有附加到活动文档，则附加该架构。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*SimpleSample 架构包含在 Smart Document Software Development Kit (SDK) 中。有关详细信息，请参阅 Microsoft Developer Network (MSDN) 网站上的 Smart Document SDK。*/ function test() { if(ActiveDocument.XMLSchemaReferences.Item(1).NamespaceURI != "SimpleSample") {     Application.XMLNamespaces.Item("SimpleSample").AttachToDocument(ActiveDocument) } }` |

#### **XMLSchemaReference.Parent**

返回一个 **Object** 类型值，该值代表指定 **XMLSchemaReference** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **XMLSchemaReference** 对象的变量。

适用环境：web

适用平台：windows/linux