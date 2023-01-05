#### **DocumentProperty**



代表容器文档的一个自定义或内置文档属性。**DocumentProperty** 对象是 **DocumentProperties** 集合的成员。

**说明**

使用 BuiltinDocumentProperties(index) 可返回一个代表特定内置文档属性的 DocumentProperty 对象，其中 index 是该内置文档属性的名称或索引号。使用 CustomDocumentProperties(index) 可返回一个代表特定自定义文档属性的 **DocumentProperty** 对象，其中 *index* 是该自定义文档属性的名称或索引号。下面的列表包含所有可用内置文档属性的名称：

| 标题         | 字数         |
| ------------ | ------------ |
| 主题         | 字符数       |
| 作者         | 安全性       |
| 关键字       | 类别         |
| 批注         | 格式         |
| 批注         | 经理         |
| 上一个作者   | 单位         |
| 修订次数     | 字节数       |
| 应用程序名   | 行数         |
| 上次打印日期 | 段落数       |
| 创建日期     | 幻灯片数     |
| 上次保存时间 | 备注数       |
| 编辑时间总计 | 隐藏幻灯片数 |
| 页数         | 多媒体剪辑数 |

容器应用程序不一定为每个内置文档属性都定义一个属性值。如果所给的应用程序没有为某内置文档属性定义一个属性值，那么返回该文档属性的 Value 属性将产生错误。

**方法**

|                                                              | 名称       | 说明                   |
| ------------------------------------------------------------ | ---------- | ---------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete** | 删除自定义的文档属性。 |

**属性**

|                                                              | 名称              | 说明                                                         |
| ------------------------------------------------------------ | ----------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**   | 获取一个 **Application** 对象，代表 **DocumentProperty** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**       | 获取一个 32 位整数，指示创建 **DocumentProperty** 对象时所使用的应用程序。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LinkSource**    | 获取或设置所链接的自定义文档属性的来源。可读/写。            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LinkToContent** | 如果自定义文档属性的值链接到容器文档的内容，则为 **True**。如果该值是静态的，则为 **False**。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**          | 获取或设置文档属性的名称。可读写。                           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**        | 获取 **DocumentProperty** 对象的 **Parent** 对象。只读。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**          | 获取或设置文档属性类型。对于内置文档属性为只读；对于自定义文档属性为可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Value**         | 获取或设置文档属性的值。可读/写。                            |

**成员方法**

#### **DocumentProperty.Delete**

删除自定义的文档属性。

**语法**

**express.Delete()**

*express*   一个代表 **DocumentProperty** 对象的变量。

**说明**

不能删除内置的文档属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例删除一个自定义文档属性。为了使该示例正确运行，必须有名为“CustomNumber”的自定义 DocumentProperty 对象。 Application.ActiveDocument.CustomDocumentProperties.Item("CustomNumber").Delete()` |

**成员属性**

#### **DocumentProperty.Application**

获取一个 **Application** 对象，代表 **DocumentProperty** 对象的容器应用程序（可以使用 **Automation** 对象的此属性返回该对象的容器应用程序）。只读。

**语法**

**express.Application**

*express*   一个代表 **DocumentProperty** 对象的变量。

#### **DocumentProperty.Creator**

获取一个 32 位整数，指示创建 **DocumentProperty** 对象时所使用的应用程序。只读。

**语法**

**express.Creator**

*express*   一个代表 **DocumentProperty** 对象的变量。

#### **DocumentProperty.LinkSource**

获取或设置所链接的自定义文档属性的来源。可读/写。

**语法**

**express.LinkSource**

*express*   一个代表 **DocumentProperty** 对象的变量。

**说明**

本属性只用于自定义文档属性，不能用于内置文档属性。

指定链接的链接源由容器应用程序定义。

设置 **LinkSource** 属性会将 **LinkToContent** 属性设置为 **True**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例显示自定义文档属性的链接状态。要使本示例正常工作，dp 必须是一个有效的 DocumentProperty 对象。 function test(dp){     let tf = dp.LinkToContent ? "" : "not "     let stat = "This property is " + tf + "linked"      if(dp.LinkToContent){         stat = stat + "\n" + "The link source is " + dp.LinkSource     }      alert(stat) }` |

#### **DocumentProperty.LinkToContent**

如果自定义文档属性的值链接到容器文档的内容，则为 **True**。如果该值是静态的，则为 **False**。可读/写。

**语法**

**express.LinkToContent**

*express*   一个代表 **DocumentProperty** 对象的变量。

**说明**

该属性只用于自定义文档属性。对内置文档属性，该属性值为 **False**。

使用 **LinkSource** 属性设置指定链接属性的来源。设置 **LinkSource** 属性会将 **LinkToContent** 属性设置为 **True**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例显示自定义文档属性的链接状态。要使本示例正常工作，dp 必须是一个有效的 DocumentProperty 对象。 function test(dp){     let tf = dp.LinkToContent ? "" : "not "     let stat = "This property is " + tf + "linked"      if(dp.LinkToContent){         stat = stat + "\n" + "The link source is " + dp.LinkSource     }      alert(stat) }` |

#### **DocumentProperty.Name**

获取或设置文档属性的名称。可读写。

**语法**

**express.Name**

*express*   一个代表 **DocumentProperty** 对象的变量。

**说明**

**DocumentProperty** 对象代表容器文档的自定义或内置文档属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例显示一个文档属性的名称、类型和值。必须向该过程传递一个有效的 DocumentProperty 对象。  function test(dp){     alert("value = " + dp.Value + "\n" +         "type = " + dp.Type + "\n" +         "name = " + dp.Name) }` |

#### **DocumentProperty.Parent**

获取 **DocumentProperty** 对象的 **Parent** 对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **DocumentProperty** 对象的变量。

**示例**

#### **DocumentProperty.Type**

获取或设置文档属性类型。对于内置文档属性为只读；对于自定义文档属性为可读/写。

**语法**

**express.Type**

*express*   一个代表 **DocumentProperty** 对象的变量。

**说明**

返回值将是一个 **MsoDocProperties** 常量。

#### **DocumentProperty.Value**

获取或设置文档属性的值。可读/写。

**语法**

**express.Value**

*express*   一个代表 **DocumentProperty** 对象的变量。

**说明**

如果容器应用程序没有定义某个内置文档属性的值，则读取该文档属性的 **Value** 属性将导致出错。

适用环境：web

适用平台：windows/linux