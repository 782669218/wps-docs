#### **Sections**



所选内容、范围或文档中的 **Section**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Section/Section%20.htm#jsObject_Section)对象的集合。

**说明**

表示指定文档的Section对象集合

**方法**

|                                                              | 名称     | 说明                                                      |
| ------------------------------------------------------------ | -------- | --------------------------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**  | 返回一个 **Section** 对象，该对象代表添加到文档中的新节。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item** | 返回集合中的单个 **Section** 对象。                       |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个代表 WPS 应用程序的 **Application** 对象。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回一个 **Long** 类型的值，该值代表集合中的节数。只读。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **First**       | 返回一个 **Section** 对象，该对象代表 **Sections** 集合中的第一个项目。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Last**        | 将 **Sections** 集合中的最后一项作为 **Section** 对象返回。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PageSetup**   | 返回一个 **PageSetup** 对象，该对象与指定文档、区域、节或选定内容相关联。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回一个 **Object** 类型值，该值代表指定 **Sections** 对象的父对象。 |

**成员方法**

#### **Sections.Add**

返回一个 **Section** 对象，该对象代表添加到文档中的新节。

**语法**

**express.Add(Range, Start)**

*express*   一个代表 **Sections** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Range*  | 可选          | **Variant**  | 要在其之前插入分节符的区域。如果省略该参数，则将分节符插至文档末尾。 |
| *Start*  | 可选          | **Variant**  | 要添加的分节符的类型。可以是 **WdSectionStart** 常量之一。如果省略该参数，则添加“下一页”类型的分节符。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在活动文档第三段之前添加一个“下一页”的分节符*/ function test() {     let myRange = Application.ActiveDocument.Paragraphs.Item(3).Range     Application.ActiveDocument.Sections.Add(myRange) }  /*本示例在选定内容中添加一个“连续”的分节符*/ function test() {     let myRange = Application.Selection.Range     Application.ActiveDocument.Sections.Add(myRange, wdSectionContinuous) }  /*本示例在活动文档的末尾添加一个“下一页”分节符*/ function test() {     Application.ActiveDocument.Sections.Add() }` |

#### **Sections.Item**

返回集合中的单个 **Section** 对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **Sections** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Index*  | 必选          | **Long**     | 要返回的单个对象。可以是代表单个对象序号位置的 **Long** 类型值。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*返回当前活动文档第一节Section对象*/ Application.ActiveDocument.Sections.Item(1)` |

**成员属性**

#### **Sections.Application**

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

*express*   一个代表 **Sections** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **Sections.Count**

返回一个 **Long** 类型的值，该值代表集合中的节数。只读。

**语法**

**express.Count**

*express*   一个代表 **Sections** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*取当前活动文档有多少节*/ Application.ActiveDocument.Sections.Count` |

#### **Sections.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Sections** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **Sections.First**

返回一个 **Section** 对象，该对象代表 **Sections** 集合中的第一个项目。

**语法**

**express.First**

*express*   一个代表 **Sections** 对象的变量。

#### **Sections.Last**

将 **Sections** 集合中的最后一项作为 **Section** 对象返回。

**语法**

**express.Last**

*express*   一个代表 **Sections** 对象的变量。

#### **Sections.PageSetup**

返回一个 **PageSetup** 对象，该对象与指定文档、区域、节或选定内容相关联。

**语法**

**express.PageSetup**

*express*   一个代表 **Sections** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将 Summary.doc 中第一节的装订线设置为 36 磅（0.5 英寸）。*/ Documents.Item("Summary.doc").Sections.Item(1).PageSetup.Gutter = 36` |

#### **Sections.Parent**

返回一个 **Object** 类型值，该值代表指定 **Sections** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **Sections** 对象的变量。

适用环境：web

适用平台：windows/linux