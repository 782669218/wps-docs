**CoAuthor** 



代表文档中的单个共同作者。**CoAuthor** 对象是 **CoAuthors** 集合的一个成员。**CoAuthors** 集合包含文档中的所有共同作者（当前正在编辑文档的作者）。

**说明**

使用 **CoAuthors**(*Index*) 可以返回单个 **CoAuthor** 对象，其中 *Index* 是索引号。

| 注释                                                         |
| ------------------------------------------------------------ |
| 当一个新的共同作者开始编辑文档时，该作者可能需要一分钟或更长时间才会显示在文档中。 |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码示例返回活动文档中第一个共同作者的名称。*/ function test() {     let author = Application.ActiveDocument.CoAuthoring.Authors.Item(1)     alert("The name of the first co-author in this document is " + author.Name) }` |

**属性**

|                                                              | 名称             | 说明                                                         |
| ------------------------------------------------------------ | ---------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**  | 返回一个代表 WPS 应用程序的 Application 对象。只读。         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**      | 返回表示用于创建指定对象的应用程序的 32 位整数。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EmailAddress** | 返回一个字符串，指明指定合著者的电子邮件地址。只读。         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ID**           | 返回一个 **String** 类型的值，该值指定某个指定作者的唯一标识符。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IsMe**         | 如果此作者代表当前用户，则返回 true。只读。                  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Locks**        | 返回一个 CoAuthLocks 集合，该集合代表文档中与指定共同作者相关联的锁。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**         | 返回一个 **String** 类型的值，该值包含指定的共同作者的显示名称。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**       | 返回一个 **Object** 类型值，该值代表指定 **CoAuthor** 对象的父对象。 |

**成员属性**

#### **CoAuthor.Application**

返回一个代表 WPS 应用程序的 Application 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **CoAuthor** 对象的变量。

#### **CoAuthor.Creator**

返回表示用于创建指定对象的应用程序的 32 位整数。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **CoAuthor** 对象的变量。

**说明**

如果对象是在 WPS 中创建的，则 **Creator** 属性返回十六进制数 4D535744，代表 **string**“WPS”。该属性主要设计用于 Apple Macintosh 平台，在该平台上，每个应用程序都有一个由四个字符组成的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的详细信息，请参考 WPS OfficeMacintosh Edition 附带的语言参考帮助。

| 注释                                    |
| --------------------------------------- |
| 该值也可用常量 **wdCreatorCode** 表示。 |

#### **CoAuthor.EmailAddress**

返回一个字符串，指明指定合著者的电子邮件地址。只读。

**语法**

**express.EmailAddress**

*express*   一个代表 **CoAuthor** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码示例显示活动文档中第一个合著者的电子邮件地址。*/ function test() {     if (!Application.ActiveDocument.CoAuthoring.Authors.Count) {         alert(Application.ActiveDocument.CoAuthoring.Authors.Item(1).EmailAddress)     }     alert("There are no co-authors in this document.") }` |

#### **CoAuthor.ID**

返回一个 **String** 类型的值，该值指定某个指定作者的唯一标识符。只读。

**语法**

**express.ID**

*express*   一个代表 **CoAuthor** 对象的变量。

**说明**

不应假定 **ID** 属性返回的唯一标识符具有特定的长度或格式。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码示例显示活动文档中的每个共同作者的唯一标识符。*/ function test() {     let allAuthors = Application.ActiveDocument.CoAuthoring.Authors     for (let i = 1; i <= allAuthors.Count; i++) {         alert("The ID for  " + allAuthors.Item(i).Name + " is " + allAuthors.Item(i).ID + ".")     } }` |

#### **CoAuthor.IsMe**

如果此作者代表当前用户，则返回 true。只读。

**语法**

**express.IsMe**

*express*   一个代表 **CoAuthor** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码示例检查活动文档，以查看 CoAuthors 集合中的第一个共同作者是否为当前用户。*/ function test() {     if (Application.ActiveDocument.CoAuthoring.Authors.Item(1).IsMe) {         alert("The current user is the first coauthor.")     } }` |

#### **CoAuthor.Locks**

返回一个 CoAuthLocks 集合，该集合代表文档中与指定共同作者相关联的锁。只读。

**语法**

**express.Locks**

*express*   一个代表 **CoAuthor** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码示例显示活动文档中与第一个共同作者相关联的锁的数量。*/ function test() {     let lockCount     let coAuth = Application.ActiveDocument.CoAuthoring.Authors.Item(1)     lockCount = coAuth.Locks.Count     alert("There are " + lockCount + " locks in the active document for " + coAuth.Name + ".") }` |

#### **CoAuthor.Name**

返回一个 **String** 类型的值，该值包含指定的共同作者的显示名称。只读。

**语法**

**express.Name**

*express*   一个代表 **CoAuthor** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码示例显示活动文档中第一个共同作者的名称。*/ function test() {     let author = Application.ActiveDocument.CoAuthoring.Authors.Item(1)     alert("The name of the first co-author in this document is " + author.Name) }` |

#### **CoAuthor.Parent**

返回一个 **Object** 类型值，该值代表指定 **CoAuthor** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **CoAuthor** 对象的变量。

适用环境：web

适用平台：windows/linux