#### **DocumentProperties**



**DocumentProperty** 对象的集合。每个 **DocumentProperty** 对象均代表容器文档的一个内置属性或自定义属性。

**说明**

使用 **Add** 方法可创建一个新的自定义属性，并将其添加到 **DocumentProperties** 集合中。不能使用 **Add** 方法创建内置文档属性。

使用 **BuiltinDocumentProperties(index)** 可返回一个代表特定内置文档属性的 **DocumentProperty** 对象；其中 *index* 是该内置文档属性的索引号。使用 **CustomDocumentProperties(index)** 可返回一个代表特定自定义文档属性的 **DocumentProperty** 对象；其中 *index* 是该自定义文档属性的编号。

**方法**

|                                                              | 名称    | 说明                                                         |
| ------------------------------------------------------------ | ------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add** | 新建一个自定义的文档属性。只能在自定义 **DocumentProperties** 集合中新添文档属性。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 获取一个 **Application** 对象，代表 **DocumentProperties** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 获取一个 **Long** 类型的值，指示 **DocumentProperties** 集合中的项数。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 获取一个 32 位整数，指示创建 **DocumentProperties** 对象时所使用的应用程序。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Item**        | 获取 **DocumentProperties** 集合中的一个 **DocumentProperty** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 获取 **DocumentProperties** 对象的 **Parent** 对象。只读。   |

**成员方法**

#### **DocumentProperties.Add**

新建一个自定义的文档属性。只能在自定义 **DocumentProperties** 集合中新添文档属性。

**语法**

**express.Add(Name, LinkToContent, Type, Value, LinkSource)**

*express*   一个代表 **DocumentProperties** 对象的变量。

**参数**

| **名称**        | **必选/可选** | **数据类型** | **说明**                                                     |
| --------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Name*          | 必选          | **String**   | 属性的名称。                                                 |
| *LinkToContent* | 必选          | **Boolean**  | 指定属性是否链接到容器文档中的内容。如果参数为 True，则必需设置 LinkSource 参数，如果为 False，则需要数值参数。 |
| *Type*          | 可选          | **Variant**  | 该属性的数据类型。可以是以下 MsoDocProperties 常量之一：msoPropertyBoolean、msoPropertyDate、msoPropertyFloat、msoPropertyNumber 或 msoPropertyString。 |
| *Value*         | 可选          | **Variant**  | 如果未链接到容器文档中的内容，则为属性的值。该数值将进行转换以匹配类型参数指定的数据类型，如果不能转换，则将发生错误。如果 LinkToContent 为 True，将忽略该参数，新文档属性将被赋予默认值，直到容器应用程序更新链接的属性值（通常在保存文档时进行）。 |
| *LinkSource*    | 可选          | **Variant**  | 如果 LinkToContent 为 False，则忽略此参数。链接属性的来源。可使用的源链接的类型由容器应用程序决定。 |

**说明**

如果在 **DocumentProperties** 集合中添加了自定义文档属性，该属性链接到 WPS Office文档中指定的值，则必须保存文档以查看 **DocumentProperty** 对象的更改。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例设计为在 WPS 中运行，它将向 DocumentProperties 集合中添加三个自定义文档属性。 function ValidateSchemas(objSourceCustomXMLSchemaCollection) {     let customProps = Application.ActiveDocument.CustomDocumentProperties     customProps.Add("CustomNumber", false, Application.Enum.msoPropertyTypeNumber, 1000)     customProps.Add("CustomString", false, Application.Enum.msoPropertyTypeString, "This is a custom property.")     customProps.Add("CustomDate", false, Application.Enum.msoPropertyTypeDate, new Date()) }` |

**成员属性**

#### **DocumentProperties.Application**

获取一个 **Application** 对象，代表 **DocumentProperties** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。

**语法**

**express.Application**

*express*   一个代表 **DocumentProperties** 对象的变量。

#### **DocumentProperties.Count**

获取一个 **Long** 类型的值，指示 **DocumentProperties** 集合中的项数。只读。

**语法**

**express.Count**

*express*   一个代表 **DocumentProperties** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例显示活动文档中自定义文档属性的数量。 alert("There are " + Application.ActiveDocument.CustomDocumentProperties.Count +     " custom document properties in the active document.")` |

#### **DocumentProperties.Creator**

获取一个 32 位整数，指示创建 **DocumentProperties** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

*express*   一个代表 **DocumentProperties** 对象的变量。

#### **DocumentProperties.Item**

获取 **DocumentProperties** 集合中的一个 **DocumentProperty** 对象。只读。

**语法**

**express.Item**

*express*   一个代表 **DocumentProperties** 对象的变量。

#### **DocumentProperties.Parent**

获取 **DocumentProperties** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **DocumentProperties** 对象的变量。

**示例**

适用环境：web

适用平台：windows/linux