#### **HeadersFooters**



**方法**

|                                                              | 名称     | 说明                                                         |
| ------------------------------------------------------------ | -------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item** | 返回一个 **HeaderFooter** 对象，该对象代表一个区域或部分中的页眉或页脚。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个代表 WPS 应用程序的 **Application** 对象。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回一个 **Long** 类型的值，该值代表集合中页眉和/或页脚的数目。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回一个 **Object** 类型值，该值代表指定 **HeadersFooters** 对象的父对象。 |

**成员方法**

#### **HeadersFooters.Item**

返回一个 **HeaderFooter** 对象，该对象代表一个区域或部分中的页眉或页脚。

**语法**

**express.Item(Index)**

*express*   一个代表 **HeadersFooters** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型**            | **说明**                           |
| -------- | ------------- | ----------------------- | ---------------------------------- |
| *Index*  | 必选          | **WdHeaderFooterIndex** | 指定区域或部分中页眉或页脚的常量。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在活动文档中创建首页页眉并设置其格式。*/ function test(){     Application.ActiveDocument.PageSetup.DifferentFirstPageHeaderFooter = true      let rng = Application.ActiveDocument.Sections.Item(1).Headers.Item(wdHeaderFooterFirstPage).Range         rng.InsertBefore("Sales Report")         rng.Font.Bold = true         rng.Font.Size = "15"         rng.Font.Color = wdColorBlue         rng.Paragraphs.Alignment = wdAlignParagraphCenter }` |

**成员属性**

#### **HeadersFooters.Application**

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

*express*   一个代表 **HeadersFooters** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **HeadersFooters.Count**

返回一个 **Long** 类型的值，该值代表集合中页眉和/或页脚的数目。只读。

**语法**

**express.Count**

*express*   一个代表 **HeadersFooters** 对象的变量。

#### **HeadersFooters.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **HeadersFooters** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **HeadersFooters.Parent**

返回一个 **Object** 类型值，该值代表指定 **HeadersFooters** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **HeadersFooters** 对象的变量。

适用环境：web

适用平台：windows/linux