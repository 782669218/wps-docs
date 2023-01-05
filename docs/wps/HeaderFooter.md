#### **HeaderFooter**



代表单一页眉或页脚。**HeaderFooter** 对象是 **HeadersFooters** 集合的一个成员。**HeadersFooters** 集合包含指定文档节中的所有页眉和页脚。

**说明**

使用 **Headers**(*Index*) 或 **Footers**(*Index*) 可返回一个 **HeaderFooter** 对象，其中 index 为 **WdHeaderFooterIndex** 常量（**wdHeaderFooterEvenPages**、**wdHeaderFooterFirstPage** 或 **wdHeaderFooterPrimary**）之一。以下示例更改活动文档第一节中主页眉和主页脚的文字。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let mySection = Application.ActiveDocument.Sections.Item(1)         mySection.Headers.Item(wdHeaderFooterPrimary).Range.Text = "Header text"         mySection.Footers.Item(wdHeaderFooterPrimary).Range.Text = "Footer text" }` |

您也可以使用 

HeaderFooter

 属性和 

Selection

 对象来返回单个 

HeaderFooter

 对象。

使用 **DifferentFirstPageHeaderFooter** 属性和 **PageSetup** 对象可指定不同的首页。以下示例在活动文档的首页页脚中插入文字。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     Application.ActiveDocument.PageSetup.DifferentFirstPageHeaderFooter = true     Application.ActiveDocument.Sections.Item(1).Footers.Item(wdHeaderFooterFirstPage).Range.InsertBefore("Written by Joe Smith") }` |

使用 **OddAndEvenPagesHeaderFooter** 属性和 **PageSetup** 对象可为奇数页和偶数页设置不同的页眉和页脚。如果 **OddAndEvenPagesHeaderFooter** 属性为 **True**，则可使用 **wdHeaderFooterPrimary** 返回奇数页的页眉或页脚，使用 **wdHeaderFooterEvenPages** 返回偶数页的页眉和页脚。

使用 **Add** 方法和 **PageNumbers** 对象可在页眉或页脚中添加页码。以下示例在活动文档第一节的主页脚中添加页码。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActiveDocument.Sections.Item(1).Footers.Item(wdHeaderFooterPrimary).PageNumbers.Add()` |

**属性**

|                                                              | 名称               | 说明                                                         |
| ------------------------------------------------------------ | ------------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**    | 返回一个代表 WPS 应用程序的 **Application** 对象。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**        | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Exists**         | 如果存在指定的 **HeaderFooter** 对象，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Index**          | 返回 **WdHeaderFooterIndex**，代表文档或节中指定的页眉或页脚。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IsHeader**       | 如果指定的 **HeaderFooter** 对象是页眉，则该属性值为 **True**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LinkToPrevious** | 如果指定页眉或页脚链接至前一节中相应的页眉或页脚，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PageNumbers**    | 返回一个 **PageNumbers** 集合，该集合代表指定页眉或页脚中包含的所有页码域。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**         | 返回一个 **Object** 类型值，该值代表指定 **HeaderFooter** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Range**          | 返回一个 **Range** 对象，该对象代表指包含在定页眉或页脚中的文档部分。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Shapes**         | 返回一个 **Shapes** 集合，该集合代表页眉或页脚中的所有的 **Shape** 对象。只读。 |

**成员属性**

#### **HeaderFooter.Application**

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

*express*   一个代表 **HeaderFooter** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **HeaderFooter.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **HeaderFooter** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **HeaderFooter.Exists**

如果存在指定的 **HeaderFooter** 对象，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.Exists**

*express*   一个代表 **HeaderFooter** 对象的变量。

**说明**

默认情况下，主页眉和页脚存在于所有新文档中。使用此方法可以确定首页或奇数页是否使用页眉或页脚。还可以使用**DifferentFirstPageHeaderFooter** 或 **OddAndEvenPagesHeaderFooter**属性返回或设置指定文档或节中的页眉和页脚数目。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果第一节包含首页页眉，则本示例设置页眉的文字。*/ function test() {   let secTemp = Application.ActiveDocument.Sections.Item(1)   if(secTemp.Headers.Item(wdHeaderFooterFirstPage).Exists == true){      secTemp.Headers.Item(wdHeaderFooterFirstPage).Range.Text = "First Page"   } }` |

#### **HeaderFooter.Index**

返回 **WdHeaderFooterIndex**，代表文档或节中指定的页眉或页脚。只读。

**语法**

**express.Index**

*express*   一个代表 **HeaderFooter** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果指定的变量引用第一个页眉，则本示例将一个图形添至活动文档中的第一个页眉。*/ function test() {     let hdrFirstPage = Application.ActiveDocument.Sections.Item(1).Headers.Item(wdHeaderFooterFirstPage)     if(hdrFirstPage.Index == wdHeaderFooterFirstPage){        hdrFirstPage.Shapes.AddShape(msoShapeHeart, 36, 36, 36, 36).Fill.ForeColor.RGB = (255, 0, 0)     } }` |

#### **HeaderFooter.IsHeader**

如果指定的 **HeaderFooter** 对象是页眉，则该属性值为 **True**。**Boolean** 类型，只读。

**语法**

**express.IsHeader**

*express*   一个代表 **HeaderFooter** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例选择页脚并添加页码。*/ function test() {   let view = Application.ActiveDocument.ActiveWindow.ActivePane.View   view.Type = wdPrintView   view.SeekView = wdSeekCurrentPageHeader      if(Application.Selection.HeaderFooter.IsHeader == true){       Application.ActiveDocument.ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter   }      Application.Selection.HeaderFooter.PageNumbers.Add() }` |

#### **HeaderFooter.LinkToPrevious**

如果指定页眉或页脚链接至前一节中相应的页眉或页脚，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.LinkToPrevious**

*express*   一个代表 **HeaderFooter** 对象的变量。

**说明**

链接页眉或页脚时，其内容与前一个页眉或页脚中的内容相同。因为默认情况下，**LinkToPrevious** 属性设为 **True**，所以可以通过设置第一节的页眉、页脚和页码来为整个文档添加页眉、页脚和页码。例如，下面的示例将页码添加至活动文档所有节中全部页的页眉中。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Apploication.ActiveDocument.Sections.Item(1).Headers.Item(wdHeaderFooterPrimary).PageNumbers.Add() ` |

LinkToPrevious

 属性可以分别应用于每个页眉或页脚。例如，可将偶数页的页眉的 

LinkToPrevious

 属性设置为 

True

，而将偶数页的页脚设置为 

False

。

 

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例的第一部分创建一个有两节的新文档，第二部分为新文档的第一和第二节的偶数页和奇数页创建不同的页眉。*/ function test() {   Application.Documents.Add()      for(let j = 1;j <= 4;j++){       Application.Selection.TypeParagraph()       Application.Selection.InsertBreak()       Application.Selection.TypeParagraph()   }      Application.ActiveDocument.Paragraphs.Item(5).Range.InsertBreak(wdSectionBreakNextPage)   Application.ActiveDocument.PageSetup.OddAndEvenPagesHeaderFooter = true      let oddHeader = Application.ActiveDocument.Sections.Item(2).Headers.Item(wdHeaderFooterPrimary)       oddHeader.LinkToPrevious = false       oddHeader.Range.InsertBefore("Section 2 Odd Header")      let evenHeader = Application.ActiveDocument.Sections.Item(2).Headers.Item(wdHeaderFooterEvenPages)       evenHeader.LinkToPrevious = false       evenHeader.Range.InsertBefore("Section 2 Even Header")      Application.ActiveDocument.Sections.Item(1).Headers.Item(wdHeaderFooterPrimary).Range.InsertBefore("Section 1 Odd Header")   Application.ActiveDocument.Sections.Item(1).Headers.Item(wdHeaderFooterEvenPages).Range.InsertBefore("Section 1 Even Header") }` |

#### **HeaderFooter.PageNumbers**

返回一个 **PageNumbers** 集合，该集合代表指定页眉或页脚中包含的所有页码域。

**语法**

**express.PageNumbers**

*express*   一个代表 **HeaderFooter** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例创建一个新文档，并向页脚中添加页码。*/ function test() {   let myDoc = Application.Documents.Add()   myDoc.Sections.Item(1).Footers.Item(wdHeaderFooterPrimary).PageNumbers.Add(wdAlignPageNumberCenter) }` |

#### **HeaderFooter.Parent**

返回一个 **Object** 类型值，该值代表指定 **HeaderFooter** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **HeaderFooter** 对象的变量。

#### **HeaderFooter.Range**

返回一个 **Range** 对象，该对象代表指包含在定页眉或页脚中的文档部分。

**语法**

**express.Range**

*express*   一个代表 **HeaderFooter** 对象的变量。

#### **HeaderFooter.Shapes**

返回一个 **Shapes** 集合，该集合代表页眉或页脚中的所有的 **Shape** 对象。只读。

**语法**

**express.Shapes**

*express*   一个代表 **HeaderFooter** 对象的变量。

**说明**

该集合可以包含绘图、形状、图片、OLE 对象、ActiveX 控件、文本对象和标注。有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

当 **Shapes** 属性应用于文档时，该属性将返回正文中的所有 **Shape** 对象，不包括页眉和页脚。当将其应用到 **HeaderFooter** 对象时，**Shapes** 属性返回文档中所有页眉和页脚中的 **Shape** 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示活动文档第一节的主页眉和页脚中所有形状的总数。*/ alert(Application.ActiveDocument.Sections.Item(1).Headers.Item(wdHeaderFooterPrimary).Shapes.Count)` |

适用环境：web

适用平台：windows/linux