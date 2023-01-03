**CustomProperty** 



代表智能标记的自定义属性的单个实例。**CustomProperty** 对象是 **CustomProperties**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/CustomProperties/CustomProperties%20.htm#jsObject_CustomProperties)集合的一个成员。

**说明**

使用 **CustomProperties** 集合的 [Item](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/CustomProperties/CustomProperties%20.htm#CustomProperties.Item)方法或 **Properties**(*Index*) 属性（其中 *Index* 是属性编号）返回一个 **CustomProperty** 对象。

使用**Nam****e** 和 **Value**属性返回与智能标记的自定义属性相关的信息。本示例显示一条消息，其中包含当前文档中第一个智能标记的第一个自定义属性的名称和值。本示例假定当前文档包含至少一个智能标记，并且第一个智能标记至少具有一个自定义属性。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function SmartTagsProps(){   let rng = Application.ActiveDocument.SmartTags.Item(1).Properties.Item(1)   alert("Smart Tag Name: " + rng.Name + "\n" + "Smart Tag Value: " + rng.Value) }` |

**方法**

|                                                              | 名称       | 说明                 |
| ------------------------------------------------------------ | ---------- | -------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete** | 删除指定自定义属性。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个代表 WPS 应用程序的 [Application](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application) 对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**        | 返回指定对象的名称。**String** 类型，只读。                  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回一个 **Object** 类型值，该值代表指定 **CustomProperty** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Value**       | 返回或设置自定义属性的值。可读/写 **String** 类型。          |

**成员方法**

#### **CustomProperty.Delete**

删除指定自定义属性。

**语法**

**express.Delete()**

*express*   一个代表 **CustomProperty** 对象的变量。

**成员属性**

#### **CustomProperty.Application**

返回一个代表 WPS 应用程序的 [Application](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application) 对象。

**语法**

**express.Application**

*express*   一个代表 **CustomProperty** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **CustomProperty.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **CustomProperty** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **CustomProperty.Name**

返回指定对象的名称。**String** 类型，只读。

**语法**

**express.Name**

*express*   一个代表 **CustomProperty** 对象的变量。

#### **CustomProperty.Parent**

返回一个 **Object** 类型值，该值代表指定 **CustomProperty** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **CustomProperty** 对象的变量。

#### **CustomProperty.Value**

返回或设置自定义属性的值。可读/写 **String** 类型。

**语法**

**express.Value**

*express*   一个代表 **CustomProperty** 对象的变量。

适用环境：web

适用平台：windows/linux