#### **Endnotes**



**Endnote** 对象的集合，该集合中的对象代表所选内容、范围或文档中的所有尾注。

**方法**

|                                                              | 名称                           | 说明                                                      |
| ------------------------------------------------------------ | ------------------------------ | --------------------------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**                        | 返回一个 **Endnote** 对象，该对象代表添加到区域中的尾注。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Convert**                    | 将尾注转换为脚注。                                        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**                       | 返回集合中的单个 **Endnote** 对象。                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ResetContinuationNotice**    | 将尾注延续标记重新设置为默认标记。                        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ResetContinuationSeparator** | 将尾注延续分隔符重新设置为默认分隔符。                    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ResetSeparator**             | 将尾注分隔符重新设置为默认分隔符。                        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SwapWithFootnotes**          | 将一篇文档中的所有尾注转换成脚注或将脚注转换为尾注。      |

**属性**

|                                                              | 名称                      | 说明                                                         |
| ------------------------------------------------------------ | ------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**           | 返回一个代表 WPS 应用程序的 [Application ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ContinuationNotice**    | 返回一个 **Range** 对象，该对象代表脚注的接续注明。只读。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ContinuationSeparator** | 返回一个 **Range** 对象，该对象代表尾注的接续分隔符。只读。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**                 | 返回一个 **Long** 类型的值，该值代表集合中尾注的数量。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**               | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Location**              | 返回或设置所有尾注的位置。[WdEndnoteLocation ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdEndnoteLocation%20%E6%9E%9A%E4%B8%BE.html)类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NumberStyle**           | 返回或设置编号样式。可读/写 **WdNoteNumberStyle** 类型。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NumberingRule**         | 返回或设置在分页符或分节符后对尾注进行编号的方式。可读写 [WdNumberingRule ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdNumberingRule%20%E6%9E%9A%E4%B8%BE.html)类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**                | 返回一个 **Object** 类型值，该值代表指定 **Endnotes** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Separator**             | 返回一个 **Range** 对象，该对象代表尾注分隔符。              |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **StartingNumber**        | 返回或设置注释编号、行号或页码的起始值。可读/写 **Long** 类型。 |

**成员方法**

#### **Endnotes.Add**

返回一个 **Endnote** 对象，该对象代表添加到区域中的尾注。

**语法**

**express.Add(Range, Reference, Text)**

*express*   一个代表 **Endnotes** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型**     | **说明**                                                     |
| ----------- | ------------- | ---------------- | ------------------------------------------------------------ |
| *Range*     | 必选          | **Range object** | 标记用于尾注或脚注的区域。可以是折叠区域。                   |
| *Reference* | 可选          | **Variant**      | 自定义引用标记的文字。如果省略该参数，WPS 将插入自动编号的引用标记。 |
| *Text*      | 可选          | **Variant**      | 尾注或脚注文本。                                             |

**说明**

要为 *Reference* 参数指定一个符号，可用 `{FontName CharNum}` 语法。FontName 是包含该符号的字体名称。修饰字体的名称显示在**“插入”**菜单**“符号”**对话框的**“字体”**框中。CharNum 是要插入符号的相应位置序数与 31 之和，在符号表中从左向右计数。例如，如果指定的符号是“Symbol”字体的 omega (ω)，它在符号表格中的位置为 56，该参数就应为“{Symbol 87}”。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为活动文档的第三段添加一条尾注。*/ function test() {   let myRange = Application.ActiveDocument.Paragraphs.Item(3).Range   Application.ActiveDocument.Endnotes.Add(myRange, undefined, "Ibid., 314.") }` |

#### **Endnotes.Convert**

将尾注转换为脚注。

**语法**

**express.Convert()**

*express*   一个代表 **Endnotes** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档中所有的尾注转换为脚注。*/ function test() {   let endDocEndnotes = Application.ActiveDocument.Endnotes   if(endDocEndnotes.Count > 0) {       myEndnotes.Convert()   } }` |

#### **Endnotes.Item**

返回集合中的单个 **Endnote** 对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **Endnotes** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Index*  | 必选          | **Long**     | 要返回的单个对象。可以是代表单个对象序号位置的 Long 类型值。 |

#### **Endnotes.ResetContinuationNotice**

将尾注延续标记重新设置为默认标记。

**语法**

**express.ResetContinuationNotice()**

*express*   一个代表 **Endnotes** 对象的变量。

**说明**

默认标记为空（无文本）。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例重新设置活动文档的尾注延续标记。*/ Application.ActiveDocument.Endnotes.ResetContinuationNotice()` |

#### **Endnotes.ResetContinuationSeparator**

将尾注延续分隔符重新设置为默认分隔符。

**语法**

**express.ResetContinuationSeparator()**

*express*   一个代表 **Endnotes** 对象的变量。

**说明**

默认分隔符是一条长横线，该横线分隔文档文字与上接前一页的注释。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例重新设置每个打开的文档中第一节的尾注延续分隔符。*/ function test() {   for(let i = 1; i <= Application.Documents.Count; i++) {       Application.Documents.Item(i).Sections.Item(1).Range.Endnotes.ResetContinuationSeparator()   } }` |

#### **Endnotes.ResetSeparator**

将尾注分隔符重新设置为默认分隔符。

**语法**

**express.ResetSeparator()**

*express*   一个代表 **Endnotes** 对象的变量。

**说明**

/*默认分隔符为一条短横线，该横线分隔文档文字和注释。*/

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例为所选内容所在的文档中的注释重新设置尾注分隔符。*/ Application.Selection.Endnotes.ResetSeparator()` |

#### **Endnotes.SwapWithFootnotes**

将一篇文档中的所有尾注转换成脚注或将脚注转换为尾注。

**语法**

**express.SwapWithFootnotes()**

*express*   一个代表 **Endnotes** 对象的变量。

**说明**

要将某一区域的尾注转换为脚注，请使用 **Convert**方法。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档中的尾注转换为脚注，并将脚注转换为尾注。*/ Application.ActiveDocument.Endnotes.SwapWithFootnotes()` |

**成员属性**

#### **Endnotes.Application**

返回一个代表 WPS 应用程序的 [Application ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象。

**语法**

**express.Application**

*express*   一个代表 **Endnotes** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **Endnotes.ContinuationNotice**

返回一个 **Range** 对象，该对象代表脚注的接续注明。只读。

**语法**

**express.ContinuationNotice**

*express*   一个代表 **Endnotes** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例用文字“Continued...”来代替当前的脚注接续注明。*/ function test() {   let notice = Application.ActiveDocument.Footnotes.ContinuationNotice   notice.Delete()   notice.InsertBefore("Continued...") }` |

#### **Endnotes.ContinuationSeparator**

返回一个 **Range** 对象，该对象代表尾注的接续分隔符。只读。

**语法**

**express.ContinuationSeparator**

*express*   一个代表 **Endnotes** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例用一组下划线字符来代替当前的尾注接续分隔符。*/ function test() {   let separator = Application.ActiveDocument.Endnotes.ContinuationSeparator   separator.Delete()   separator.InsertBefore("____") }` |

#### **Endnotes.Count**

返回一个 **Long** 类型的值，该值代表集合中尾注的数量。只读。

**语法**

**express.Count**

*express*   一个代表 **Endnotes** 对象的变量。

#### **Endnotes.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Endnotes** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **Endnotes.Location**

返回或设置所有尾注的位置。[WdEndnoteLocation ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdEndnoteLocation%20%E6%9E%9A%E4%B8%BE.html)类型，可读写。

**语法**

**express.Location**

*express*   一个代表 **Endnotes** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将所有的尾注置于节的结尾。*/ Application.ActiveDocument.Endnotes.Location = wdEndOfSection` |

#### **Endnotes.NumberStyle**

返回或设置编号样式。可读/写 **WdNoteNumberStyle** 类型。

**语法**

**express.NumberStyle**

*express*   一个代表 **Endnotes** 对象的变量。

**说明**

根据所选择或安装的语言支持（例如，英语（美国））的不同，以上某些常量可能不可用。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置活动文档中脚注和尾注的格式。*/ function test() {   Application.ActiveDocument.Footnotes.NumberStyle = wdNoteNumberStyleLowercaseRoman   Application.ActiveDocument.Endnotes.NumberStyle = wdNoteNumberStyleUppercaseRoman }` |

#### **Endnotes.NumberingRule**

返回或设置在分页符或分节符后对尾注进行编号的方式。可读写 [WdNumberingRule ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdNumberingRule%20%E6%9E%9A%E4%B8%BE.html)类型。

**语法**

**express.NumberingRule**

*express*   一个代表 **Endnotes** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在活动文档中的每个分节符后对尾注重新开始编号。*/ Application.ActiveDocument.Endnotes.NumberingRule = wdRestartSection` |

#### **Endnotes.Parent**

返回一个 **Object** 类型值，该值代表指定 **Endnotes** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **Endnotes** 对象的变量。

#### **Endnotes.Separator**

返回一个 **Range** 对象，该对象代表尾注分隔符。

**语法**

**express.Separator**

*express*   一个代表 **Endnotes** 对象的变量。

#### **Endnotes.StartingNumber**

返回或设置注释编号、行号或页码的起始值。可读/写 **Long** 类型。

**语法**

**express.StartingNumber**

*express*   一个代表 **Endnotes** 对象的变量。

**说明**

只有在页面视图中才能看到行号。

当应用于页码时，该属性返回或设置指定 **HeaderFooter** 对象的起始页码。该页码是否在第一页显示取决于 **ShowFirstPageNumber** 属性的设置。如果 **RestartNumberingAtSection** 属性设置为 **False**，它将覆盖 **StartingNumber** 属性，以便可以从前一节继续编排页码。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例创建一篇新文档，将脚注的起始编号设置为 10，然后在插入点添加一个脚注。*/ function test() {   let myDoc = Application.Documents.Add()   myDoc.Footnotes.StartingNumber=10   myDoc.Footnotes.Add(Selection.Range,"Text for a footnote") }  /*本示例将为活动文档设置行号。起始编号设置为 5，编号间隔设置为 5，并且文档每节的开头独立编号。*/ function test() {   let numbering= Application.ActiveDocument.PageSetup.LineNumbering   numbering.Active = true   numbering.StartingNumber = 5   numbering.CountBy = 5   numbering.RestartMode = wdRestartSection }  /*本示例设置页码属性，然后为活动文档的页眉添加页码。*/ function test() {   let Starting= Application.ActiveDocument.Sections.Item(1).Headers.Item(wdHeaderFooterPrimary).PageNumbers   Starting.NumberStyle = wdPageNumberStyleArabic   Starting.IncludeChapterNumber = false   Starting.RestartNumberingAtSection = true   Starting.StartingNumber = 5   Starting.Add(wdAlignPageNumberCenter,true) }` |

适用环境：web

适用平台：windows/linux