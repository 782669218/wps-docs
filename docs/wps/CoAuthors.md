**CoAuthors** 



文档中所有 **CoAuthor** 对象的集合。

**说明**

**CoAuthors** 集合包含文档中的所有共同作者（当前正在编辑文档的作者）。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码示例获取活动文档中的共同作者的数量。*/ function test() { let i i = ActiveDocument.CoAuthoring.Authors.Count MsgBox("The number of co-authors is " + i) }` |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个代表 WPS 应用程序的 Application 对象。只读。         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回 CoAuthors 集合中的项目数。只读。                        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回表示用于创建指定对象的应用程序的 32 位整数。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回一个 **Object** 类型值，该值代表指定 **CoAuthors** 对象的父对象。 |

**成员属性**

#### **CoAuthors.Application**

返回一个代表 WPS 应用程序的 Application 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **CoAuthors** 对象的变量。

#### **CoAuthors.Count**

返回 CoAuthors 集合中的项目数。只读。

**语法**

**express.Count**

*express*   一个代表 **CoAuthors** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码示例显示活动文档中的共同作者的数量。*/ MsgBox("The active document contains " + ActiveDocument.CoAuthoring.Authors.Count + " authors.")` |

#### **CoAuthors.Creator**

返回表示用于创建指定对象的应用程序的 32 位整数。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **CoAuthors** 对象的变量。

**说明**

如果对象是在 WPS 中创建的，则 **Creator** 属性返回十六进制数 4D535744，代表 **string**“WPS”。该属性主要设计用于 Apple Macintosh 平台，在该平台上，每个应用程序都有一个由四个字符组成的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的详细信息，请参考 WPS OfficeMacintosh Edition 附带的语言参考帮助。

| 注释                                    |
| --------------------------------------- |
| 该值也可用常量 **wdCreatorCode** 表示。 |

#### **CoAuthors.Parent**

返回一个 **Object** 类型值，该值代表指定 **CoAuthors** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **CoAuthors** 对象的变量。

适用环境：web

适用平台：windows/linux