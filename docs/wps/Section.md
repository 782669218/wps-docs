#### **Section**



代表所选内容、范围或文档中的一节

**说明**

代表所选内容、范围或文档中的一节。**Section** 对象是 **Sections** 集合的一个成员。**Sections** 集合包括所选内容、范围或文档中的所有节。

**属性**

|                                                              | 名称                  | 说明                                                         |
| ------------------------------------------------------------ | --------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**       | 返回一个代表 WPS 应用程序的 **Application** 对象。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Borders**           | 返回 **Borders** 集合，该集合代表节中的所有边框。            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**           | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Footers**           | 返回一个 **HeadersFooters** 集合，该集合代表指定节的页脚。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Headers**           | 返回一个 **HeadersFooters** 集合，该集合代表指定节的页眉。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Index**             | 返回一个 **Long** 类型的值，该值代表项目在集合中的位置。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Pagesetup**         | 代表页面设置说明                                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**            | 返回一个 **Object** 类型值，该值代表指定 **Section** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ProtectedForForms** | 如果保护指定节中的窗体，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Range**             | 返回一个 **Range** 对象，该对象代表指定对象中包含的文档部分。 |

**成员属性**

#### **Section.Application**

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

*express*   一个代表 **Section** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **Section.Borders**

返回 **Borders** 集合，该集合代表节中的所有边框。

**语法**

**express.Borders**

*express*   一个代表 **Section** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

#### **Section.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Section** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **Section.Footers**

返回一个 **HeadersFooters** 集合，该集合代表指定节的页脚。只读。

**语法**

**express.Footers**

*express*   一个代表 **Section** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。要返回代表指定节的页眉的 **HeadersFooters** 集合，请使用 **Headers** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例向活动文档第一节中的主页脚添加一个右对齐的页码。*/ function test() { let myFooters = ActiveDocument.Sections.Item(1).Footers.Item(wdHeaderFooterPrimary) myFooters.PageNumbers.Add(wdAlignPageNumberRight) }` |

#### **Section.Headers**

返回一个 **HeadersFooters** 集合，该集合代表指定节的页眉。只读。

**语法**

**express.Headers**

*express*   一个代表 **Section** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。要返回代表指定节的页脚的 **HeadersFooters** 集合，请使用 **Footers** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为活动文档中除首页之外的每页添加居中的页码（为首页创建一个独立的页眉）。*/ function test() { let myHeaders = ActiveDocument.Sections.Item(1).Headers.Item(wdHeaderFooterPrimary) myHeaders.PageNumbers.Add(wdAlignPageNumberCenter, false) }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将向活动文档首页页眉添加文本。*/ function test() { ActiveDocument.PageSetup.DifferentFirstPageHeaderFooter = true let myHeaders = ActiveDocument.Sections.Item(1).Headers.Item(wdHeaderFooterFirstPage) myHeaders.Range.InsertAfter("First Page Text") myHeaders.Range.Paragraphs.Alignment = wdAlignParagraphRight }` |

#### **Section.Index**

返回一个 **Long** 类型的值，该值代表项目在集合中的位置。只读。

**语法**

**express.Index**

*express*   一个代表 **Section** 对象的变量。

**示例**

| 示例代码复制                                        |
| --------------------------------------------------- |
| `Application.ActiveDocument.Sections.Item(1).Index` |

#### **Section.Pagesetup**

代表页面设置说明

**语法**

**express.Pagesetup**

*express*   一个代表 **Section** 对象的变量。

**说明**

**代表页面设置说明。PageSetup** 对象包含文档的所有页面设置属性（如左边距、下边距和纸张大小）。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例将活动文档的第一节设为横向 Application.ActiveDocument.Sections.Item(1).PageSetup.Orientation =  wdOrientLandscape ` |

#### **Section.Parent**

返回一个 **Object** 类型值，该值代表指定 **Section** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **Section** 对象的变量。

#### **Section.ProtectedForForms**

如果保护指定节中的窗体，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.ProtectedForForms**

*express*   一个代表 **Section** 对象的变量。

**说明**

如果保护节中的窗体，则只能选择和修改窗体域中的文本。要保护整个文档，请使用 **Document** 对象的 **Protect** 方法。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例打开活动文档中第二节的窗体保护。*/ function test() { if(ActiveDocument.Sections.Count >= 2) {     ActiveDocument.Sections.Item(2).ProtectedForForms = true } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例取消对选定内容中第一节的保护。*/ Selection.Sections.Item(1).ProtectedForForms = false` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例切换对选定内容中第一节的保护状态。*/ Selection.Sections.Item(1).ProtectedForForms = !(Selection.Sections.Item(1).ProtectedForForms)` |

#### **Section.Range**

返回一个 **Range** 对象，该对象代表指定对象中包含的文档部分。

**语法**

**express.Range**

*express*   一个代表 **Section** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//下示例在活动文档的开头将文本作为一个新段落插入 function test() {     var myRange = Application.ActiveDocument.Sections.Item(1).Range     myRange.MoveEnd(wdCharacter,-1)     myRange.Collapse(wdCollapseEnd)     myRange.InsertParagraphAfter()     myRange.InsertAfter("End of section") }` |

适用环境：web

适用平台：windows/linux