#### **XMLNodes**



**XMLNode** 对象的集合，该集合代表**“XML 结构”**任务窗格的树视图中的所有节点，这些节点表示用户已应用于文档的元素。树视图中的每个节点都是 **XMLNode** 对象的一个实例。树视图的层次结构表明节点是否包含子节点。

**方法**

|                                                              | 名称     | 说明                                                |
| ------------------------------------------------------------ | -------- | --------------------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**  | 返回一个 **XMLNode** 对象，该对象代表新添加的元素。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item** | 返回集合中的单个 **XMLNode** 对象。                 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个代表 WPS 应用程序的 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回一个 **Number** 类型的值，该值代表集合中 XML 元素的数量。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |

**成员方法**

#### **XMLNodes.Add**

返回一个 **XMLNode** 对象，该对象代表新添加的元素。

**语法**

**express.Add(Name, Namespace, Range)**

*express*   一个代表 **XMLNodes** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Name*      | 必选          | **String**   | 在 Namespace 参数指定的 XML 架构中的元素的名称。由于 XML 区分大小写，因此，Name 参数所指定的元素的拼写必须和它在架构中所显示的完全一样。如果与 Namespace 参数中指定的架构中的任何元素名称都不匹配，则显示错误。 |
| *Namespace* | 必选          | **String**   | 架构中定义的架构名称。Namespace 参数区分大小写，因此必须严格按照它在架构中所显示的那样进行拼写。如果在附加到文档的任何架构中都找不到指定的命名空间，则显示错误。 |
| *Range*     | 可选          | **Range**    | 要应用元素的区域。默认为将元素标记置于插入点或选定内容的周围（如果已选择文本）。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 以下示例将选定文本添加到示例元素，该示例元素来自活动文档中所引用的第一个架构*/ function test(){     Application.ActiveDocument.XMLNodes.Add ("example",ActiveDocument.XMLSchemaReferences.Item(1).NamespaceURI,Selection.Range) }` |

#### **XMLNodes.Item**

返回集合中的单个 **XMLNode** 对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **XMLNodes** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Index*  | 必选          | **Number**   | 要返回的单个对象。可以是代表单个对象序号位置的 Number 类型值。 |

**成员属性**

#### **XMLNodes.Application**

返回一个代表 WPS 应用程序的 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象。

**语法**

**express.Application**

*express*   一个代表 **XMLNodes** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **XMLNodes.Count**

返回一个 **Number** 类型的值，该值代表集合中 XML 元素的数量。只读。

**语法**

**express.Count**

*express*   一个代表 **XMLNodes** 对象的变量。

#### **XMLNodes.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **XMLNodes** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **XMLNodes.Parent**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Parent**

*express*   一个代表 **XMLNodes** 对象的变量。

适用环境：web

适用平台：windows/linux