#### **XMLNamespace**



代表架构库中的单个架构。

**说明**

在 WPS 中，可以通过**“模板和加载项”**对话框中的**“XML 架构”**选项卡来访问架构库。架构库表示安装在用户计算机上的架构，并且用户已将这些架构应用于 WPS 文档或者用户已使用**“架构库”**对话框将这些架构显式添加到架构库中。

使用 **XMLNamespaces** 集合的 **Item** 方法可返回一个 **XMLNameSpace** 对象。**Item** 方法的索引值可以是 **Long** 类型，表明架构在架构库中的位置，也可以是 **String** 类型，代表使用 **URI** 属性（在架构中定义的 TargetNamespace 设置）返回的架构的名称。

以下示例将一个名为 SimpleSample 的架构附加到活动文档。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function ApplySampleSchema(){     let objSchema = Application.XMLNamespaces          for(let i =1; i <= objSchema.Count; i++){         if(objSchema.Item(i).URI == "SimpleSample"){             objSchema.Item(i).AttachToDocument (ActiveDocument)             break         }     } }` |

| 注释                                                         |
| ------------------------------------------------------------ |
| SimpleSample 架构包含在 Smart Document Software Development Kit (SDK) 中。有关详细信息，请参阅 Microsoft Developer Network (MSDN) 网站上的 Smart Document SDK。 |

**方法**

|                                                              | 名称                 | 说明                                       |
| ------------------------------------------------------------ | -------------------- | ------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AttachToDocument** | 将 XML 架构附加到文档。                    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**           | 删除可用的 XML 架构列表中指定的 XML 架构。 |

**属性**

|                                                              | 名称                 | 说明                                                         |
| ------------------------------------------------------------ | -------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Alias**            | 返回一个 **String** 类型的值，该值代表指定对象的显示名称。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**      | 返回一个代表 WPS 应用程序的 **Application** 对象。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**          | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DefaultTransform** | 该属性返回一个 **XSLTransform** 对象，该对象表示对于特殊命名空间从 XML 架构中打开文档时要使用的默认 Extensible Stylesheet Language Transformation (XSLT) 文件。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Location**         | 返回或设置 **String** 值，该值表示 XML 架构的命名空间在架构库中的物理位置。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**           | 返回一个 **Object** 类型值，该值代表指定 **XMLNamespace** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **URI**              | 返回一个 **String** 类型的值，代表相关命名空间的“统一资源标识符”(URI)。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **XSLTransforms**    | 返回一个 XSLTransforms 集合，代表指定用于架构的 Extensible Stylesheet Language Transformation (XSLT) 文件。 |

**成员方法**

#### **XMLNamespace.AttachToDocument**

将 XML 架构附加到文档。

**语法**

**express.AttachToDocument(Document)**

*express*   一个代表 **XMLNamespace** 对象的变量。

**参数**

| **名称**   | **必选/可选** | **数据类型** | **说明**                      |
| ---------- | ------------- | ------------ | ----------------------------- |
| *Document* | 必选          | **Document** | 将指定 XML 架构附加到的文档。 |

**示例**

以下示例将 SimpleSample 架构添加到 Schema Library，再附加到活动文档。

| 注释                                                         |
| ------------------------------------------------------------ |
| SimpleSample 架构包含在 Smart Document Software Development Kit (SDK) 中。有关详细信息，请参阅 Microsoft Developer Network (MSDN) 网站上的 Smart Document SDK。 |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() { let objSchema = Application.XMLNamespaces.Add("c:\\schemas\\simplesample.xsd") objSchema.AttachToDocument (ActiveDocument) }  ` |

#### **XMLNamespace.Delete**

删除可用的 XML 架构列表中指定的 XML 架构。

**语法**

**express.Delete()**

*express*   一个代表 **XMLNamespace** 对象的变量。

**成员属性**

#### **XMLNamespace.Alias**

返回一个 **String** 类型的值，该值代表指定对象的显示名称。

**语法**

**express.Alias**

*express*   一个代表 **XMLNamespace** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示活动文档所附加的第一个架构的显示名称。*/ MsgBox(Application.XMLNamespaces.Item(1).Alias)` |

#### **XMLNamespace.Application**

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

*express*   一个代表 **XMLNamespace** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **XMLNamespace.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **XMLNamespace** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **XMLNamespace.DefaultTransform**

该属性返回一个 **XSLTransform** 对象，该对象表示对于特殊命名空间从 XML 架构中打开文档时要使用的默认 Extensible Stylesheet Language Transformation (XSLT) 文件。

**语法**

**express.DefaultTransform**

*express*   一个代表 **XMLNamespace** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例返回架构库中第一个架构的默认 XSLT，WPS 将使用 XSLT 打开与该架构命名空间相关联的 XML 文件。本示例假设第一个架构有一个或多个已应用的 XSLT 文件。*/ let objXSLT = Application.XMLNamespaces.Item(1).DefaultTransform` |

#### **XMLNamespace.Location**

返回或设置 **String** 值，该值表示 XML 架构的命名空间在架构库中的物理位置。可读/写。

**语法**

**express.Location**

*express*   一个代表 **XMLNamespace** 对象的变量。

#### **XMLNamespace.Parent**

返回一个 **Object** 类型值，该值代表指定 **XMLNamespace** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **XMLNamespace** 对象的变量。

#### **XMLNamespace.URI**

返回一个 **String** 类型的值，代表相关命名空间的“统一资源标识符”(URI)。

**语法**

**express.URI**

*express*   一个代表 **XMLNamespace** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下列示例显示“架构库”中第一个架构的 URI。*/ MsgBox(Application.XMLNamespaces.Item(1).URI)` |

#### **XMLNamespace.XSLTransforms**

返回一个 XSLTransforms 集合，代表指定用于架构的 Extensible Stylesheet Language Transformation (XSLT) 文件。

**语法**

**express.XSLTransforms**

*express*   一个代表 **XMLNamespace** 对象的变量。

**说明**

下列示例将 simplesample.xsl 转换添加到 SimpleSample 架构的转换中。

| 注释                                                         |
| ------------------------------------------------------------ |
| SimpleSample 架构包含在 Smart Document Software Development Kit (SDK) 中；但没有相应的 simplesample.xsl 文件。此代码仅为示例而创建。有关详细信息，请参阅 Microsoft Developer Network (MSDN) 网站上的 Smart Document SDK。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() { let objSchema = Application.XMLNamespaces.Item("SimpleSample") let objTransform = objSchema.XSLTransforms.Add("c:\\schemas\\simplesample.xsl") }` |

适用环境：web

适用平台：windows/linux