#### **Style**



代表单个内置样式或用户定义的样式。**Style** 对象将样式属性（如字体、字形、段落间距）添加为 **Style** 对象的属性。**Style** 对象是 **Styles** 集合的成员。 **Styles** 集合包含指定文档中的所有样式。

**说明**

使用 **Styles**(*Index*) 可返回一个 **Style** 对象，其中 *Index* 为样式名、**WdBuiltinStyle** 常量或索引号。样式名的拼写和间距必须完全匹配，但是其大小写不必完全匹配。以下示例更改活动文档中名为“Color”的用户定义的样式的字体名称。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActiveDocument.Styles.Item("Color").Font.Name = "Arial"` |

以下示例将内置“标题 1”样式设置为非加粗。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActiveDocument.Styles.Item(wdStyleHeading1).Font.Bold = false` |

样式索引编号代表样式在按字母顺序排列的样式名列表中的位置。请注意，

 是按字母顺序排列的列表中的第一个样式。以下示例显示 

Styles

 集合中第一个样式的基准样式和样式名。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `alert("Base style= " + Application.ActiveDocument.Styles.Item(1).BaseStyle + "\r" + "Style name= " + ActiveDocument.Styles.Item(1).NameLocal)` |

要将样式应用于范围、段落或多个段落，请将 

Style

 属性设为用户定义的样式名或内置样式名。以下示例对活动文档的前四段应用“正文”样式。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `let myRange = Application.ActiveDocument.Range(ActiveDocument.Paragraphs.Item(1).Range.Start,ActiveDocument.Paragraphs.Item(4).Range.End) myRange.Style = wdStyleNormal` |

以下示例对所选内容的第一段应用“标题 1”样式。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.Selection.Paragraphs.Item(1).Style = wdStyleHeading1` |

以下示例创建一个名为“Bolded”的字符样式，并将其应用于所选内容。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `let myStyle = Application.ActiveDocument.Styles.Add("Bolded",wdStyleTypeCharacter) myStyle.Font.Bold = true Application.Selection.Range.Style = "Bolded"` |

使用 

OrganizerCopy

 方法可在文档和模板之间复制样式。使用 

UpdateStyles

 方法可更新活动文档中的样式，使其与附加模板中的样式定义相匹配。使用 

OpenAsDocument

 方法可将模板作为文档打开，以便可以修改该模板的样式。

 

**方法**

|                                                              | 名称                   | 说明                                                         |
| ------------------------------------------------------------ | ---------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**             | 删除指定的样式。                                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **LinkToListTemplate** | 将指定的样式与一个列表模板链接，如此该样式中的格式就可以应用于列表。 |

**属性**

|                                                              | 名称                                    | 说明                                                         |
| ------------------------------------------------------------ | --------------------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**                         | 返回一个代表 WPS 应用程序的 **Application** 对象。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutomaticallyUpdate**                 | 如果样式根据选定内容重新定义，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BaseStyle**                           | 返回或设置一种现有样式，根据此样式可以设置其他样式。**Variant** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Borders**                             | 返回一个 **Borders** 集合，该集合代表指定样式的所有边框。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BuiltIn**                             | 如果指定对象是 WPS 的内置样式或题注标签之一，则该属性值为 **True**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**                             | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Description**                         | 返回指定样式的说明。只读 **String** 类型。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Font**                                | 返回或设置 **Font** 对象，该对象代表指定样式的字符格式。**Font** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Frame**                               | 返回一个 **Frame** 对象，该对象代表指定样式的图文框格式。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **InUse**                               | 如果为 **True**，则指定的样式是一个已在文档中修改或应用的内置样式，或者是在文档中新建的样式。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LanguageID**                          | 返回或设置一个 **WdLanguageID** 常量，该常量代表指定范围的语言。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LanguageIDFarEast**                   | 返回或设置指定对象的东亚语言。可读写 **WdLanguageID** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Linked**                              | 返回或设置 **Boolean** 值，该值表示某样式是否是段落和字符格式都可以使用的链接样式。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LinkStyle**                           | 设置或返回一个 **Variant** 类型的值，该值代表段落和字符样式之间的链接。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ListLevelNumber**                     | 返回指定样式的列表级别。只读 **Long** 类型。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ListTemplate**                        | 返回一个 **ListTemplate** 对象，该对象代表指定 **Style** 对象的列表格式。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Locked**                              | 如果不能更改或编辑样式，则该属性值为 **True**。可读写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NameLocal**                           | 返回由用户语言表示的内置样式的名称。可读写 **String** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NextParagraphStyle**                  | 返回或设置自动应用于在格式设置为指定样式的段落之后插入的新段落的样式。**Variant** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NoProofing**                          | 如果拼写和语法检查程序忽略设置为此样式格式的文本，则该属性值为 **True**。可读写 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NoSpaceBetweenParagraphsOfSameStyle** | 如果为 **True**，则 WPS 会删除设置为使用相同样式的段落的间距。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ParagraphFormat**                     | 返回或设置一个 **ParagraphFormat** 对象，该对象代表指定样式的段落设置。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**                              | 返回一个 **Object** 类型值，该值代表指定 **Style** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Priority**                            | 返回或设置一个 **Long** 类型的值，该值代表**“样式”**任务窗格中排序样式的优先级。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **QuickStyle**                          | 返回或设置一个 **Boolean** 类型的值，该值代表样式是否对应于可用的快速样式。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Shading**                             | 返回一个 **Shading** 对象，该对象代表指定对象的底纹格式。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Table**                               | 返回一个 **TableStyle** 对象，该对象代表可通过使用表格样式应用于表格的属性。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**                                | 返回样式的类型。只读 **WdStyleType** 类型。                  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **UnhideWhenUsed**                      | 如果指定样式在用于文档中后显示为 WPS 2015 中**“样式”**和**“样式”**任务窗格中的推荐样式，则该属性值为 **True**。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Visibility**                          | 如果指定样式显示为**“样式”**库和**“样式”**任务窗格中的推荐样式，则该属性值为 **True**。可读/写。 |

**成员方法**

#### **Style.Delete**

删除指定的样式。

**语法**

**express.Delete()**

*express*   一个代表 **Style** 对象的变量。

#### **Style.LinkToListTemplate**

将指定的样式与一个列表模板链接，如此该样式中的格式就可以应用于列表。

**语法**

**express.LinkToListTemplate(ListTemplate, ListLevelNumber)**

*express*   一个代表 **Style** 对象的变量。

**参数**

| **名称**          | **必选/可选** | **数据类型**            | **说明**                                                     |
| ----------------- | ------------- | ----------------------- | ------------------------------------------------------------ |
| *ListTemplate*    | 必选          | **ListTemplate object** | 与样式相链接的列表模板。                                     |
| *ListLevelNumber* | 可选          | **Variant**             | 对应样式所链接到的列表级的整数。如果省略该参数，则使用样式的级别。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例创建一个新的列表模板，接着将 1 到 9 的标题样式与 1 到 9 的列表级别相链接。然后将新列表模板应用于文档。任何应用标题样式的段落都根据列表模板进行编号。*/ function test() { let Lis let Sty let ltTemp = ActiveDocument.ListTemplates.Add(true) for(let i=1;i <= 9;i++) {     Lis = ltTemp.ListLevels.Item(i)     Lis.NumberStyle = wdListNumberStyleArabic     Lis.NumberPosition = InchesToPoints(0.25 * (i - 1))     Lis.TextPosition = InchesToPoints(0.25 * i)     Lis.NumberFormat = "%" + i + "."     Sty = ActiveDocument.Styles.Item("标题 " + i)     Sty.LinkToListTemplate(ltTemp) } ActiveDocument.Content.ListFormat.ApplyListTemplate(ltTemp) }` |

**成员属性**

#### **Style.Application**

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

*express*   一个代表 **Style** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **Style.AutomaticallyUpdate**

如果样式根据选定内容重新定义，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.AutomaticallyUpdate**

*express*   一个代表 **Style** 对象的变量。

**说明**

如果将 **AutomaticallyUpdate** 属性设置为 **False**，则 WPS 会在根据选定内容重新定义样式之前，提示确认。在将样式应用于具有相同样式、但手动格式设置不同的选定内容时，可以重新定义样式。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例创建名为“Style1”的样式，该样式不需进行确认即可重新定义。*/ function test() { let docNew = Documents.Add() let styleNew = docNew.Styles.Add("Style1") styleNew.BaseStyle = docNew.Styles.Item(wdStyleNormal) styleNew.ParagraphFormat.LineSpacingRule = wdLineSpaceDouble styleNew.AutomaticallyUpdate = true }` |

#### **Style.BaseStyle**

返回或设置一种现有样式，根据此样式可以设置其他样式。**Variant** 类型，可读写。

**语法**

**express.BaseStyle**

*express*   一个代表 **Style** 对象的变量。

**说明**

要设置 **BaseStyle** 属性，请指定基本样式的本地名称（一个整数或 **wdBuiltinStyle** 常量），或代表基本样式的对象。有关 **wdBuiltinStyle** 常量的列表，请参阅要设置的对象的 **Style** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例新建一篇文档，然后为其新添一个名为“myHeading”的段落样式，并以样式“标题 1”为新样式的基本样式。然后将新样式的左缩进量设置为 1 英寸（72 磅）。*/ function test() { let docNew = Documents.Add() let styleNew = docNew.Styles.Add("NewHeading1") styleNew.BaseStyle = docNew.Styles.Item(wdStyleHeading1) styleNew.ParagraphFormat.LeftIndent = 72 }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回文档正文段落样式所用的基本样式。*/ function test() { let styleBase = ActiveDocument.Styles.Item(wdStyleBodyText).BaseStyle alert(styleBase) }  ` |

#### **Style.Borders**

返回一个 **Borders** 集合，该集合代表指定样式的所有边框。

**语法**

**express.Borders**

*express*   一个代表 **Style** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

#### **Style.BuiltIn**

如果指定对象是 WPS 的内置样式或题注标签之一，则该属性值为 **True**。**Boolean** 类型，只读。

**语法**

**express.BuiltIn**

*express*   一个代表 **Style** 对象的变量。

**说明**

可用 **WdBuiltinStyle** 常量指定所有语言通用的内置样式，也可用 WPS 的某个语言版本的样式名来指定某种语言使用的内置样式。例如，如果在 WPS Office的语言设置中指定美式英语，则下列语句是等效的：

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() { ActiveDocument.Styles.Item(wdStyleHeading1) ActiveDocument.Styles.Item("Heading 1") }` |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例实现的功能是：检查活动文档中的所有样式，如果检查到一个非内置样式，则显示该样式的名称。*/ function test() { for(let i=1;i <= ActiveDocument.Styles.Count;i++) {     if(ActiveDocument.Styles.Item(i).BuiltIn == false) {         alert(ActiveDocument.Styles.Item(i).NameLocal)     } } }` |

#### **Style.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Style** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **Style.Description**

返回指定样式的说明。只读 **String** 类型。

**语法**

**express.Description**

*express*   一个代表 **Style** 对象的变量。

**说明**

样式说明的典型示例可以是：“正文+字体：Arial、12 磅、加粗、倾斜、段前 12 磅、段后 3 磅、KeepWithNext、2 级”。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例新建一个文档，并插入一个制表符分隔的活动文档样式及其说明的列表*/ function test() { 	let docActive = Application.ActiveDocument 	let docNew = Application.Documents.Add() 	for(let i=1;i <= docActive.Styles.Count;i++) { 	    let Rng = docNew.Range() 	    Rng.InsertAfter(ActiveDocument.Styles.Item(i).NameLocal + "\t" + ActiveDocument.Styles.Item(i).Description) 	    Rng.InsertParagraphAfter() 	} }` |

#### **Style.Font**

返回或设置 **Font** 对象，该对象代表指定样式的字符格式。**Font** 类型，可读写。

**语法**

**express.Font**

*express*   一个代表 **Style** 对象的变量。

**说明**

要设置该属性，需指定一个返回 **Font** 对象的表达式。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将取消活动文档的“标题 1”样式中的加粗格式*/ Application.ActiveDocument.Styles.Item(wdStyleHeading1).Font.Bold = false` |

#### **Style.Frame**

返回一个 **Frame** 对象，该对象代表指定样式的图文框格式。只读。

**语法**

**express.Frame**

*express*   一个代表 **Style** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例创建一个带图文框格式的样式，然后将该样式用于选定内容的第一段。*/ function test() { let styleNew = ActiveDocument.Styles.Add("frame",wdStyleTypeParagraph) let Fra = styleNew.Frame Fra.RelativeHorizontalPosition = wdRelativeHorizontalPositionMargin Fra.HeightRule = wdFrameAuto Fra.WidthRule = wdFrameAuto Fra.TextWrap = true Selection.Paragraphs.Item(1).Range.Style = "frame" }` |

#### **Style.InUse**

如果为 **True**，则指定的样式是一个已在文档中修改或应用的内置样式，或者是在文档中新建的样式。**Boolean** 类型，只读。

**语法**

**express.InUse**

*express*   一个代表 **Style** 对象的变量。

**说明**

**InUse** 属性并不一定表明该样式当前是否已应用于文档中的任何文本。例如，删除一个应用了该样式的文本后，该样式的 **InUse** 属性仍然为 **True**。对于尚未在文档中使用过的内置样式，该属性返回 **False**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一个消息框，列出活动文档中当前正在使用的所有样式的名称。*/ function test() { let docActive = ActiveDocument let strMessage = "Styles in use:" + "\r" for(let i = 1;i < = docActive.Styles.Count; i++)  {      if(docActive.Styles.Item(i).InUse == true) {         docActive.Content.Find         docActive.Content.Find.ClearFormatting()         docActive.Content.Find.Text = ""         docActive.Content.Find.Style = docActive.Styles.Item(i)         docActive.Content.Find.Execute(null,null,null,null,null,null,null,null,true)         if(docActive.Content.Find.Found == true) {             strMessage = strMessage + docActive.Styles.Item(i).Name + "\r"         }     } } alert(strMessage) }` |

#### **Style.LanguageID**

返回或设置一个 **WdLanguageID** 常量，该常量代表指定范围的语言。可读写。

**语法**

**express.LanguageID**

*express*   一个代表 **Style** 对象的变量。

**说明**

根据您所选择或安装的语言支持（例如，美国英语），某些 **WdLanguageID** 常量可能不可用。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将 Title 样式重定义为使用西班牙语校对工具，然后在消息框中显示新样式的说明。*/ function test() { ActiveDocument.Styles.Item("Title").LanguageID = wdSpanish alert(ActiveDocument.Styles.Item("Title").Description) }` |

#### **Style.LanguageIDFarEast**

返回或设置指定对象的东亚语言。可读写 **WdLanguageID** 类型。

**语法**

**express.LanguageIDFarEast**

*express*   一个代表 **Style** 对象的变量。

**说明**

推荐使用本方法，以返回或设置在 WPS 东亚版中创建的文档中的东亚文字的语言。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将所选内容的语言设置为朝鲜语。*/ Selection.LanguageIDFarEast = wdKorean` |

#### **Style.Linked**

返回或设置 **Boolean** 值，该值表示某样式是否是段落和字符格式都可以使用的链接样式。只读。

**语法**

**express.Linked**

*express*   一个代表 **Style** 对象的变量。

**说明**

如果为 **True**，表示**“修改样式”**对话框中的**“样式类型”**选项设置为**“链接(段落和字符)”**。

#### **Style.LinkStyle**

设置或返回一个 **Variant** 类型的值，该值代表段落和字符样式之间的链接。可读写。

**语法**

**express.LinkStyle**

*express*   一个代表 **Style** 对象的变量。

**说明**

一旦将字符样式与段落样式链接起来，这两种样式就采取相同的字符格式。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例创建并设置新字符样式的格式，并将字符样式与内置标题样式“标题 1”链接起来，这样“标题 1”样式就可以采用新创建样式的字符格式。*/ function LinkHeadStyle() {     let styChar1 = ActiveDocument.Styles.Add("标题 1 Characters",wdStyleTypeCharacter)     styChar1.Font.Name = "Verdana"     styChar1.Font.Bold = true     styChar1.Font.Shadow = true     let Bor = styChar1.Font.Borders.Item(1)     Bor.LineStyle = wdLineStyleDot     Bor.LineWidth = wdLineWidth300pt     Bor.Color = wdColorDarkRed     ActiveDocument.Styles.Item("标题 1").LinkStyle = ActiveDocument.Styles.Item("标题 1 Characters")     let Con = ActiveDocument.Content     Con.InsertParagraphAfter()     Con.InsertAfter("New Linked Style")     Con.Select()     Selection.Collapse(wdCollapseEnd)     Selection.Style = ActiveDocument.Styles.Item("标题 1") }` |

#### **Style.ListLevelNumber**

返回指定样式的列表级别。只读 **Long** 类型。

**语法**

**express.ListLevelNumber**

*express*   一个代表 **Style** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示“标题 3”样式的列表级别。*/ alert(ActiveDocument.Styles.Item(wdStyleHeading3).ListLevelNumber)` |

#### **Style.ListTemplate**

返回一个 **ListTemplate** 对象，该对象代表指定 **Style** 对象的列表格式。

**语法**

**express.ListTemplate**

*express*   一个代表 **Style** 对象的变量。

**说明**

一个列表模板包含定义特定列表的所有格式。**“项目符号和编号”**对话框（**“格式”**菜单）中的每个选项卡上有七种格式（不包括**“无”**），每一种格式对应于一个列表模板。文档和模板也可以包含列表模板的集合。

#### **Style.Locked**

如果不能更改或编辑样式，则该属性值为 **True**。可读写 **Boolean** 类型。

**语法**

**express.Locked**

*express*   一个代表 **Style** 对象的变量。

#### **Style.NameLocal**

返回由用户语言表示的内置样式的名称。可读写 **String** 类型。

**语法**

**express.NameLocal**

*express*   一个代表 **Style** 对象的变量。

**说明**

通过设置该属性可为用户定义的样式重新命名，或者为内置样式添加别名。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示应用于选定段落的样式名（由用户语言表示）。如果应用于选定段落的样式不止一个，则显示第一个样式名。*/ alert(Selection.Paragraphs.Style.NameLocal)` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在活动文档中为“标题 1”样式添加别名“MyH1”。*/ ActiveDocument.Styles.Item("标题 1").NameLocal = "MyH1"` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将名为“Test”的样式重新命名为“Intro”。*/ ActiveDocument.Styles.Item("Test").NameLocal = "Intro"` |

#### **Style.NextParagraphStyle**

返回或设置自动应用于在格式设置为指定样式的段落之后插入的新段落的样式。**Variant** 类型，可读写。

**语法**

**express.NextParagraphStyle**

*express*   一个代表 **Style** 对象的变量。

**说明**

您可以使用样式的本地名称（一个整数或 **WdBuiltinStyle** 常量）或代表新样式的对象，来设置 **NextParagraphStyle** 属性。有关 **WdBuiltinStyle** 常量的列表，请参阅要设置的对象的 **Style** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在活动文档中设置“标题 1”样式后接“标题 2”样式。*/ ActiveDocument.Styles.Item(wdStyleHeading1).NextParagraphStyle = ActiveDocument.Styles.Item(wdStyleHeading2)` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例新建一篇文档并添加一个名为“MyStyle”的段落样式。该新样式根据“正文”样式创建，后接“标题 3”样式，左缩进量为 1 英寸（72 磅），设为加粗格式。*/ function test() { let myDoc = Documents.Add() let myStyle = myDoc.Styles.Add("MyStyle") myStyle.BaseStyle = wdStyleNormal myStyle.NextParagraphStyle = wdStyleHeading3 myStyle.ParagraphFormat.LeftIndent = 72 myStyle.Font.Bold = true }` |

#### **Style.NoProofing**

如果拼写和语法检查程序忽略设置为此样式格式的文本，则该属性值为 **True**。可读写 **Long** 类型。

**语法**

**express.NoProofing**

*express*   一个代表 **Style** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例设置拼写和语法检查程序，使其忽略活动文档中设置为“Test”样式格式的任何文字。*/ ActiveDocument.Styles.Item("Test").NoProofing = true` |

#### **Style.NoSpaceBetweenParagraphsOfSameStyle**

如果为 **True**，则 WPS 会删除设置为使用相同样式的段落的间距。**Boolean** 类型，可读写。

**语法**

**express.NoSpaceBetweenParagraphsOfSameStyle**

*express*   一个代表 **Style** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function NoSpace() {     ActiveDocument.Styles.Item("List 1").NoSpaceBetweenParagraphsOfSameStyle = true }` |

#### **Style.ParagraphFormat**

返回或设置一个 **ParagraphFormat** 对象，该对象代表指定样式的段落设置。可读/写。

**语法**

**express.ParagraphFormat**

*express*   一个代表 **Style** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例修改活动文档的标题 2 样式。应用该样式的段落将根据第一个制表位缩进，并且采用 2 倍行距。*/ function test() { let Par = ActiveDocument.Styles.Item(wdStyleHeading2).ParagraphFormat Par.TabIndent(1) Par.Space2() }` |

#### **Style.Parent**

返回一个 **Object** 类型值，该值代表指定 **Style** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **Style** 对象的变量。

#### **Style.Priority**

返回或设置一个 **Long** 类型的值，该值代表**“样式”**任务窗格中排序样式的优先级。可读写。

**语法**

**express.Priority**

*express*   一个代表 **Style** 对象的变量。

#### **Style.QuickStyle**

返回或设置一个 **Boolean** 类型的值，该值代表样式是否对应于可用的快速样式。可读写。

**语法**

**express.QuickStyle**

*express*   一个代表 **Style** 对象的变量。

#### **Style.Shading**

返回一个 **Shading** 对象，该对象代表指定对象的底纹格式。

**语法**

**express.Shading**

*express*   一个代表 **Style** 对象的变量。

#### **Style.Table**

返回一个 **TableStyle** 对象，该对象代表可通过使用表格样式应用于表格的属性。

**语法**

**express.Table**

*express*   一个代表 **Style** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例新建一个表格样式，该样式仅为第一行、最后一行和最后一列指定四周边框以及特殊的边框和底纹。*/ function NewTableStyle() {     let styTable = ActiveDocument.Styles.Add("TableStyle 1", wdStyleTypeTable)     let Tab = styTable.Table     //Apply borders around table, a double border to the heading row,     //a double border to the last column, and shading to last row     Tab.Borders.Item(wdBorderTop).LineStyle = wdLineStyleSingle     Tab.Borders.Item(wdBorderBottom).LineStyle = wdLineStyleSingle     Tab.Borders.Item(wdBorderLeft).LineStyle = wdLineStyleSingle     Tab.Borders.Item(wdBorderRight).LineStyle = wdLineStyleSingle     Tab.Condition(wdFirstRow).Borders.Item(wdBorderBottom).LineStyle = wdLineStyleDouble     Tab.Condition(wdLastColumn).Borders.Item(wdBorderLeft).LineStyle = wdLineStyleDouble     Tab.Condition(wdLastRow).Shading.BackgroundPatternColor = wdColorGray125 }` |

#### **Style.Type**

返回样式的类型。只读 **WdStyleType** 类型。

**语法**

**express.Type**

*express*   一个代表 **Style** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条消息，显示活动文档中名为“SubTitle”的样式的样式类型。*/ function test() { if(ActiveDocument.Styles.Item("SubTitle").Type == wdStyleTypeParagraph) {     alert("Paragraph style") } else if(ActiveDocument.Styles.Item("SubTitle").Type == wdStyleTypeCharacter) {     alert("Character style") } }  ` |

#### **Style.UnhideWhenUsed**

如果指定样式在用于文档中后显示为 WPS 2015 中**“样式”**和**“样式”**任务窗格中的推荐样式，则该属性值为 **True**。可读/写。

**语法**

**express.UnhideWhenUsed**

*express*   一个代表 **Style** 对象的变量。

#### **Style.Visibility**

如果指定样式显示为**“样式”**库和**“样式”**任务窗格中的推荐样式，则该属性值为 **True**。可读/写。

**语法**

**express.Visibility**

*express*   一个代表 **Style** 对象的变量。

适用环境：web

适用平台：windows/linux