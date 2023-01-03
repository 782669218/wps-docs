#### **ParagraphFormat**



代表段落的所有格式。

**说明**

使用 **Format** 属性可返回一个或多个段落的 **ParagraphFormat** 对象。**ParagraphFormat** 属性返回所选内容、范围、样式、 **Find** 对象或 **Replacement** 对象的 **ParagraphFormat** 对象。以下示例将活动文档中的第三段居中。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActiveDocument.Paragraphs.Item(3).Format.Alignment = wdAlignParagraphCenter` |

可以使用 JS 的 **n****ew** 关键词新建一个单独的 **ParagraphFormat** 对象。以下示例创建一个 **ParagraphFormat** 对象，为其设置部分格式属性，然后将其所有属性应用于活动文档中的第一段。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let myParaF = Application.ActiveDocument.Paragraphs.Item(1).Format     myParaF.Alignment = wdAlignParagraphCenter     myParaF.Borders.Enable = true } ` |

**方法**

|                                                              | 名称                         | 说明                                       |
| ------------------------------------------------------------ | ---------------------------- | ------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CloseUp**                  | 删除指定段落格式中段落前的任何间距。       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **IndentCharWidth**          | 将一个或多个段落缩进指定的字符数。         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **IndentFirstLineCharWidth** | 将一个或多个段落的首行缩进指定的字符数。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **OpenOrCloseUp**            | 切换指定段落的段前间距。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **OpenUp**                   | 为指定段落设置 12 磅的段前间距。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Reset**                    | 删除手动段落格式（不使用样式应用的格式）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Space1**                   | 为指定段落设置单倍行距。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Space15**                  | 为指定段落设置 1.5 倍行距。                |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Space2**                   | 为指定段落设置 2 倍行距。                  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **TabHangingIndent**         | 将悬挂缩进量设置为指定的制表位数。         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **TabIndent**                | 将指定段落的左缩进量设置为指定的制表位数。 |

**属性**

|                                                              | 名称                                | 说明                                                         |
| ------------------------------------------------------------ | ----------------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AddSpaceBetweenFarEastAndAlpha**  | 如果 WPS 将自动在指定段落的日文和拉丁文文字之间添加空格，则该属性值为 **True**。如果仅对于某些指定段落将该属性设置为 **True**，则该属性会返回 **wdUndefined**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AddSpaceBetweenFarEastAndDigit**  | 如果 WPS 将自动在指定段落的日文文字和数字之间添加空格，则该属性值为 **True**。如果仅对于某些指定段落将该属性设置为 **True**，则该属性会返回 **wdUndefined**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Alignment**                       | 返回或设置一个 **WdParagraphAlignment** 常量，该常量代表指定段落的对齐方式，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**                     | 返回一个代表 WPS 应用程序的 **Application** 对象。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoAdjustRightIndent**           | 如果 WPS 会根据您指定的每行字符数自动调整指定段落的右缩进，则该属性值为 **True**。如果只将某些指定段落的 **AutoAdjustRightIndent** 属性设置为 **True**，则该属性会返回 **wdUndefined**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BaseLineAlignment**               | 返回或设置一个 **WdBaselineAlignment** 常量，该常量代表行中字体的垂直位置，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Borders**                         | 返回一个 **Borders** 集合，该集合代表指定对象的所有边框。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CharacterUnitFirstLineIndent**    | 返回或设置首行或悬挂缩进的值（以字符为单位）。用正值设置首行缩进，用负值设置悬挂缩进。**Single** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CharacterUnitLeftIndent**         | 该属性返回或设置指定段落的左缩进量（以字符为单位）。**Single** 类 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CharacterUnitRightIndent**        | 该属性返回或设置指定段落的右缩进量（以字符为单位）。**Single** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**                         | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisableLineHeightGrid**           | 如果该属性的值为 **True**，则当指定每页的行数时，WPS 会将指定段落中的字符与行网格对齐。如果只将某些指定段落的 **DisableLineHeightGrid** 属性设置为 **True**，则返回 **wdUndefined**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Duplicate**                       | 返回一个只读 **ParagraphFormat** 对象，该对象代表指定段落的段落格式。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FarEastLineBreakControl**         | 如果为 **True**，则 WPS 会将东亚语言文字的换行规则应用于指定的段落。如果只将某些指定段落的 **FarEastLineBreakControl** 属性设定为 **True**，则返回 **wdUndefined**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FirstLineIndent**                 | 返回或设置首行缩进或悬挂缩进的大小（以磅值表示）。用正数设置首行缩进的尺寸，用负数设置悬挂缩进的尺寸。**Single** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HalfWidthPunctuationOnTopOfLine** | 如果为 **True**，则 WPS 会将指定段落行首的标点符号改为半角字符。如果仅将某些指定段落的该属性设置为 **True**，则此属性将返回 **wdUndefined**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HangingPunctuation**              | 如果为 **True**，则指定段落中的标点将可以溢出边界。如果仅将某些指定段落的该属性设置为 **True**，则返回 **wdUndefined**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Hyphenation**                     | 如果指定的段落进行自动断字，则该属性值为 **True**。如果指定的段落不进行自动断字，则该属性值为 **False**。可读写 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **KeepTogether**                    | 在 WPS 对文档重新分页时，如果指定段落中的所有行都位于同一页上，则该属性值为 **True**。可读写 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **KeepWithNext**                    | 在 WPS 对文档重新分页时，如果指定段落与它的下一段位于同一页上，则该属性值为 **True**。可读写 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LeftIndent**                      | 返回或设置一个 **Single** 类型的值，该值代表指定段落格式的左缩进值（以磅为单位）。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LineSpacing**                     | 返回或设置指定段落的行距（以磅为单位）。**Single** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LineSpacingRule**                 | 返回或设置指定段落格式的行距。可读写 **WdLineSpacing** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LineUnitAfter**                   | 返回或设置指定段落的段后间距（以网格线为单位）。可读写 **Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LineUnitBefore**                  | 返回或设置指定段落的段前间距（以网格线为单位）。可读写 **Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MirrorIndents**                   | 返回或设置一个 **Long** 类型的值，该值代表左缩进和右缩进的宽度是否相同。该属性值可以是 **True**、**False** 或 **wdUndefined**。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NoLineNumber**                    | 如果取消指定段的行号，则该属性值为 **True**。该属性值可以是 **True**、**False** 或 **wdUndefined**。可读写 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OutlineLevel**                    | 返回或设置指定段落的大纲级别。可读写 **WdOutlineLevel** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PageBreakBefore**                 | 如果在指定段落前插入了分页符，则该属性值为 **True**。该属性值可以是 **True**、**False** 或 **wdUndefined**。可读写 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**                          | 返回一个 **Object** 类型值，该值代表指定 **ParagraphFormat** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ReadingOrder**                    | 返回或设置指定段落的读取次序而不改变其对齐方式。可读写 **WdReadingOrder** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RightIndent**                     | 返回或设置指定段落的右缩进量（以磅为单位）。可读写 **Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Shading**                         | 返回一个 **Shading** 对象，该对象代表指定对象的底纹格式。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SpaceAfter**                      | 返回或设置指定段落或文本栏后面的间距（以磅为单位）。可读/写 **Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SpaceAfterAuto**                  | 如果 WPS 自动设置指定段落的段后间距，则该属性为 **True**。可读/写 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SpaceBefore**                     | 返回或设置指定段落的段前间距（以磅为单位）。可读/写 **Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SpaceBeforeAuto**                 | 如果 WPS 自动设置指定段落的段前间距，则该属性为 **True**。可读/写 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Style**                           | 返回或设置指定对象的样式。可读写 **Variant** 类型。          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TabStops**                        | 返回或设置一个 **TabStops** 集合，该集合代表指定段落中的所有自定义制表位。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextboxTightWrap**                | 返回或设置一个 **WdTextboxTightWrap** 常量，该常量代表文本环绕形状或文本框的紧密程度。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **WidowControl**                    | 在 WPS 对文档重新分页时，如果指定段落的首行和末行与段落的其他各行同页，则该属性值为 **True**。该属性值可以是 **True**、**False** 或 **wdUndefined**。可读写 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **WordWrap**                        | 如果 WPS 在指定段落或文本框架的西文单词中间断字换行，则该属性值为 **True**。可读写 **Long** 类型。 |

**成员方法**

#### **ParagraphFormat.CloseUp**

删除指定段落格式中段落前的任何间距。

**语法**

**express.CloseUp()**

*express*   一个代表 **ParagraphFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下两个语句是等价的：*/ function test() { let pgfm = ActiveDocument.Paragraphs.Item(1).Format pgfm.CloseUp() pgfm.SpaceBefore = 0 }` |

#### **ParagraphFormat.IndentCharWidth**

将一个或多个段落缩进指定的字符数。

**语法**

**express.IndentCharWidth(Count)**

*express*   一个代表 **ParagraphFormat** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                 |
| -------- | ------------- | ------------ | ------------------------ |
| *Count*  | 必选          | **Integer**  | 指定段落要缩进的字符数。 |

**说明**

此方法等效于单击**“格式”**工具栏上的**“增加缩进量”**按钮。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动文档的第一段缩进 10 个字符。*/ ActiveDocument.Paragraphs.Item(1).Format.IndentCharWidth (10)` |

#### **ParagraphFormat.IndentFirstLineCharWidth**

将一个或多个段落的首行缩进指定的字符数。

**语法**

**express.IndentFirstLineCharWidth(Count)**

*express*   一个代表 **ParagraphFormat** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                           |
| -------- | ------------- | ------------ | ---------------------------------- |
| *Count*  | 必选          | **Integer**  | 每个指定段落的首行要缩进的字符数。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动文档中第一段的首行缩进 10 个字符。*/ Application.ActiveDocument.Paragraphs.Item(1).Format.IndentFirstLineCharWidth (10)` |

#### **ParagraphFormat.OpenOrCloseUp**

切换指定段落的段前间距。

**语法**

**express.OpenOrCloseUp()**

*express*   一个代表 **ParagraphFormat** 对象的变量。

**说明**

如果指定段落的段前间距等于 0（零），则此方法将其设置为 12 磅；如果段落的段前间距大于 0（零），则此方法将其设置为 0（零）。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例切换活动文档中第一段的格式，或增加 12 磅的段前间距，或不设置段前间距。*/ ActiveDocument.Paragraphs.Item(1).Format.OpenOrCloseUp()` |

#### **ParagraphFormat.OpenUp**

为指定段落设置 12 磅的段前间距。

**语法**

**express.OpenUp()**

*express*   一个代表 **ParagraphFormat** 对象的变量。

**说明**

您还可以使用 **SpaceBefore** 属性设置段落间距。下列两个语句是等效的：

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() { Selection.ParagraphFormat.OpenUp() Selection.ParagraphFormat.SpaceBefore = 12 }` |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例更改活动文档的第二段的格式，为该段落设置 12 磅的段前间距。*/ ActiveDocument.Paragraphs.Item(2).Format.OpenUp()` |

#### **ParagraphFormat.Reset**

删除手动段落格式（不使用样式应用的格式）。

**语法**

**express.Reset()**

*express*   一个代表 **ParagraphFormat** 对象的变量。

**说明**

如果手动将段落设置为右对齐，而基本样式为其他对齐方式，则用 **Reset** 方法将手动对齐方式改为与基本样式相同。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例删除活动文档第二段的手动设定的段落格式。*/ Application.ActiveDocument.Paragraphs.Item(2).Reset()` |

#### **ParagraphFormat.Space1**

为指定段落设置单倍行距。

**语法**

**express.Space1()**

*express*   一个代表 **ParagraphFormat** 对象的变量。

**说明**

精确间距取决于各段内最大字符的字号。

您还可以使用 **LineSpacingRule** 属性设置段落间距。下列两个语句是等效的：

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() { Selection.ParagraphFormat.Space1() Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle }` |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动文档中第一段的行距更改为单倍行距。*/ ActiveDocument.Paragraphs.Item(1).Format.Space1()` |

#### **ParagraphFormat.Space15**

为指定段落设置 1.5 倍行距。

**语法**

**express.Space15()**

*express*   一个代表 **ParagraphFormat** 对象的变量。

**说明**

精确间距为各段内最大字符的字号加上 6 磅。

您还可以使用 **LineSpacingRule** 属性设置段落间距。下列两个语句是等效的：

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() { Selection.ParagraphFormat.Space15() Selection.ParagraphFormat.LineSpacingRule = wdLineSpace1pt5 }` |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动文档中第一段的行距更改为 1.5 倍行距。*/ ActiveDocument.Paragraphs.Item(1).Format.Space15()` |

#### **ParagraphFormat.Space2**

为指定段落设置 2 倍行距。

**语法**

**express.Space2()**

*express*   一个代表 **ParagraphFormat** 对象的变量。

**说明**

精确间距为各段内最大字符的字号加上 12 磅。

您还可以使用 **LineSpacingRule** 属性设置段落间距。例如，下列两个语句是等效的：

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() { Selection.ParagraphFormat.Space2() Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceDouble     }` |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将所选内容中第一段的行距更改为 2 倍行距。*/ ActiveDocument.Paragraphs.Item(1).Format.Space2()` |

#### **ParagraphFormat.TabHangingIndent**

将悬挂缩进量设置为指定的制表位数。

**语法**

**express.TabHangingIndent(Count)**

*express*   一个代表 **ParagraphFormat** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Count*  | 必选          | **Integer**  | 如果为正，则表示要缩进的制表位数；如果为负，则表示要删除的缩进制表位数。 |

**说明**

如果 *Count* 的值为负数，您还可以使用此方法删除悬挂缩进的制表位。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将所选段落的悬挂缩进设置为第二个制表位。*/ Selection.ParagraphFormat.TabHangingIndent(2)` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将所选段落的悬挂缩进量减少一个制表位。*/ Selection.ParagraphFormat.TabHangingIndent(-1)` |

#### **ParagraphFormat.TabIndent**

将指定段落的左缩进量设置为指定的制表位数。

**语法**

**express.TabIndent(Count)**

*express*   一个代表 **ParagraphFormat** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Count*  | 必选          | **Integer**  | 如果为正，则表示要缩进的制表位数；如果为负，则表示要删除的缩进制表位数。 |

**说明**

如果 *Count* 的值为负数，您还可以使用此方法删除缩进。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将所选段落缩进到第二个制表位。*/ Selection.ParagraphFormat.TabIndent(2)` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将所选段落的缩进量减少一个制表位。*/ Selection.ParagraphFormat.TabIndent(-1)` |

**成员属性**

#### **ParagraphFormat.AddSpaceBetweenFarEastAndAlpha**

如果 WPS 将自动在指定段落的日文和拉丁文文字之间添加空格，则该属性值为 **True**。如果仅对于某些指定段落将该属性设置为 **True**，则该属性会返回 **wdUndefined**。**Long** 类型，可读写。

**语法**

**express.AddSpaceBetweenFarEastAndAlpha**

*express*   一个代表 **ParagraphFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例实现的功能是：设置 WPS 在活动文档第一段的日文文字和拉丁文字之间自动添加空格。*/ ActiveDocument.Paragraphs.Item(1).AddSpaceBetweenFarEastAndAlpha = true` |

#### **ParagraphFormat.AddSpaceBetweenFarEastAndDigit**

如果 WPS 将自动在指定段落的日文文字和数字之间添加空格，则该属性值为 **True**。如果仅对于某些指定段落将该属性设置为 **True**，则该属性会返回 **wdUndefined**。**Long** 类型，可读写。

**语法**

**express.AddSpaceBetweenFarEastAndDigit**

*express*   一个代表 **ParagraphFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例可实现的功能是：设定 WPS 自动在活动文档第一段的中文文字和数字之间添加空格。*/ ActiveDocument.Paragraphs.Item(1).AddSpaceBetweenFarEastAndDigit = true` |

#### **ParagraphFormat.Alignment**

返回或设置一个 **WdParagraphAlignment** 常量，该常量代表指定段落的对齐方式，可读写。

**语法**

**express.Alignment**

*express*   一个代表 **ParagraphFormat** 对象的变量。

**说明**

由于您选择或安装的语言支持不同（例如美国英语），上述部分常量可能无法使用。

#### **ParagraphFormat.Application**

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

*express*   一个代表 **ParagraphFormat** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **ParagraphFormat.AutoAdjustRightIndent**

如果 WPS 会根据您指定的每行字符数自动调整指定段落的右缩进，则该属性值为 **True**。如果只将某些指定段落的 **AutoAdjustRightIndent** 属性设置为 **True**，则该属性会返回 **wdUndefined**。**Long** 类型，可读写。

**语法**

**express.AutoAdjustRightIndent**

*express*   一个代表 **ParagraphFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例实现的功能是：设定 WPS 根据指定的每行字符数，自动调整所选段落的右缩进。*/ Selection.ParagraphFormat.AutoAdjustRightIndent = true` |

#### **ParagraphFormat.BaseLineAlignment**

返回或设置一个 **WdBaselineAlignment** 常量，该常量代表行中字体的垂直位置，可读写。

**语法**

**express.BaseLineAlignment**

*express*   一个代表 **ParagraphFormat** 对象的变量。

#### **ParagraphFormat.Borders**

返回一个 **Borders** 集合，该集合代表指定对象的所有边框。

**语法**

**express.Borders**

*express*   一个代表 **ParagraphFormat** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

#### **ParagraphFormat.CharacterUnitFirstLineIndent**

返回或设置首行或悬挂缩进的值（以字符为单位）。用正值设置首行缩进，用负值设置悬挂缩进。**Single** 类型，可读写。

**语法**

**express.CharacterUnitFirstLineIndent**

*express*   一个代表 **ParagraphFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档中第一段的首行缩进设为一个字符。*/ ActiveDocument.Paragraphs.Item(1).CharacterUnitFirstLineIndent = 1` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档中第二段的悬挂缩进设为 1.5 个字符。*/ ActiveDocument.Paragraphs.Item(2).CharacterUnitFirstLineIndent = -1.5` |

#### **ParagraphFormat.CharacterUnitLeftIndent**

该属性返回或设置指定段落的左缩进量（以字符为单位）。**Single** 类

**语法**

**express.CharacterUnitLeftIndent**

*express*   一个代表 **ParagraphFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档中第一段的左缩进设为从左边距缩进一个字符。*/ ActiveDocument.Paragraphs.Item(1).CharacterUnitLeftIndent = 1` |

#### **ParagraphFormat.CharacterUnitRightIndent**

该属性返回或设置指定段落的右缩进量（以字符为单位）。**Single** 类型，可读写。

**语法**

**express.CharacterUnitRightIndent**

*express*   一个代表 **ParagraphFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档中的所有段落的右缩进设为从右边距缩进一个字符。*/ ActiveDocument.Paragraphs.CharacterUnitRightIndent = 1` |

#### **ParagraphFormat.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **ParagraphFormat** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **ParagraphFormat.DisableLineHeightGrid**

如果该属性的值为 **True**，则当指定每页的行数时，WPS 会将指定段落中的字符与行网格对齐。如果只将某些指定段落的 **DisableLineHeightGrid** 属性设置为 **True**，则返回 **wdUndefined**。**Long** 类型，可读写。

**语法**

**express.DisableLineHeightGrid**

*express*   一个代表 **ParagraphFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例实现的功能是：设定 WPS 在已指定每页行数的情况下，将选定段落中的字符与行网格对齐。*/ Selection.ParagraphFormat.DisableLineHeightGrid = true` |

#### **ParagraphFormat.Duplicate**

返回一个只读 **ParagraphFormat** 对象，该对象代表指定段落的段落格式。

**语法**

**express.Duplicate**

*express*   一个代表 **ParagraphFormat** 对象的变量。

**说明**

可以使用 **Duplicate** 属性来获取重复的 **ParagraphFormat** 对象所有属性的设置。可将 **Duplicate** 属性返回的对象赋给另一个相同类型的对象，从而一次应用该对象的所有设置。在将重复对象赋给另一个对象之前，可更改重复对象的任何属性，而不会影响源对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例复制活动文档第一段的段落格式，并将该格式保存在 myDup 变量中，然后将 myDup 的左缩进量改为 1 英寸。再创建一篇新文档，插入文本，将 myDup 中保存的段落格式应用于该文本。*/ function test() { ActiveDocument.Range(0, 0).InsertAfter ("Paragraph Number 1") let myDup = ActiveDocument.Paragraphs.Item(1).Format.Duplicate myDup.LeftIndent = InchesToPoints(1) Documents.Add() Selection.InsertAfter ("This is a new paragraph.") Selection.Paragraphs.Format = myDup }  ` |

#### **ParagraphFormat.FarEastLineBreakControl**

如果为 **True**，则 WPS 会将东亚语言文字的换行规则应用于指定的段落。如果只将某些指定段落的 **FarEastLineBreakControl** 属性设定为 **True**，则返回 **wdUndefined**。**Long** 类型，可读写。

**语法**

**express.FarEastLineBreakControl**

*express*   一个代表 **ParagraphFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设定 WPS 将东亚语言文字的换行规则应用于活动文档的第一段。*/ function test() { ActiveDocument.Paragraphs.Item(1).FarEastLineBreakControl = true } ` |

#### **ParagraphFormat.FirstLineIndent**

返回或设置首行缩进或悬挂缩进的大小（以磅值表示）。用正数设置首行缩进的尺寸，用负数设置悬挂缩进的尺寸。**Single** 类型，可读写。

**语法**

**express.FirstLineIndent**

*express*   一个代表 **ParagraphFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为活动文档的首段设置 1 英寸的首行缩进。*/ Application.ActiveDocument.Paragraphs.Item(1).FirstLineIndent = InchesToPoints(1)  /*本示例为活动文档的第二段设置 0.5 英寸的悬挂缩进。InchesToPoints 方法用来将英寸转化为磅值。*/ Application.ActiveDocument.Paragraphs.Item(2).FirstLineIndent = InchesToPoints(-0.5)` |

#### **ParagraphFormat.HalfWidthPunctuationOnTopOfLine**

如果为 **True**，则 WPS 会将指定段落行首的标点符号改为半角字符。如果仅将某些指定段落的该属性设置为 **True**，则此属性将返回 **wdUndefined**。**Long** 类型，可读写。

**语法**

**express.HalfWidthPunctuationOnTopOfLine**

*express*   一个代表 **ParagraphFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 将活动文档第一段的行首标点符号改为半角字符。*/ ActiveDocument.Paragraphs.Item(1).HalfWidthPunctuationOnTopOfLine = true` |

#### **ParagraphFormat.HangingPunctuation**

如果为 **True**，则指定段落中的标点将可以溢出边界。如果仅将某些指定段落的该属性设置为 **True**，则返回 **wdUndefined**。**Long** 类型，可读写。

**语法**

**express.HangingPunctuation**

*express*   一个代表 **ParagraphFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例实现的功能是：允许活动文档第一段的标点可以溢出边界。*/ ActiveDocument.Paragraphs.Item(1).HangingPunctuation = true` |

#### **ParagraphFormat.Hyphenation**

如果指定的段落进行自动断字，则该属性值为 **True**。如果指定的段落不进行自动断字，则该属性值为 **False**。可读写 **Long** 类型。

**语法**

**express.Hyphenation**

*express*   一个代表 **ParagraphFormat** 对象的变量。

**说明**

此属性可以是 **True**、**False** 或 **wdUndefined**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例关闭活动文档中具有“正文”样式的所有段落的自动断字功能。*/ ActiveDocument.Styles.Item("正文").ParagraphFormat.Hyphenation = false` |

#### **ParagraphFormat.KeepTogether**

在 WPS 对文档重新分页时，如果指定段落中的所有行都位于同一页上，则该属性值为 **True**。可读写 **Long** 类型。

**语法**

**express.KeepTogether**

*express*   一个代表 **ParagraphFormat** 对象的变量。

**说明**

此属性可以是 **True**、**False** 或 **wdUndefined**。

#### **ParagraphFormat.KeepWithNext**

在 WPS 对文档重新分页时，如果指定段落与它的下一段位于同一页上，则该属性值为 **True**。可读写 **Long** 类型。

**语法**

**express.KeepWithNext**

*express*   一个代表 **ParagraphFormat** 对象的变量。

**说明**

此属性可以是 **True**、**False** 或 **wdUndefined**。

#### **ParagraphFormat.LeftIndent**

返回或设置一个 **Single** 类型的值，该值代表指定段落格式的左缩进值（以磅为单位）。可读写。

**语法**

**express.LeftIndent**

*express*   一个代表 **ParagraphFormat** 对象的变量。

#### **ParagraphFormat.LineSpacing**

返回或设置指定段落的行距（以磅为单位）。**Single** 类型，可读写。

**语法**

**express.LineSpacing**

*express*   一个代表 **ParagraphFormat** 对象的变量。

**说明**

使用**LinesToPoints**方法可将行数转换为相应的值（以磅为单位）。例如，`LinesToPoints(2)` 返回值 24。

可以在 **LineSpacingRule**属性已设置为下列值之后对**LineSpacing**属性进行设置：

- **wdLineSpaceAtLeast** 行距可以大于或等于指定的 **LineSpacing** 值，但绝不能小于该值。
- **wdLineSpaceExactly** 指定的 **LineSpacing**行距值绝不能更改，即使段落中使用了较大的字体。
- **wdLineSpaceMultiple** 必须指定**LineSpacing**属性值，以磅为单位。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将选定段落的行距设置为至少 24 磅。*/ function test() {   let pgfm = Application.Selection.ParagraphFormat   pgfm.LineSpacingRule = wdLineSpaceAtLeast   pgfm.LineSpacing = 24 }` |

#### **ParagraphFormat.LineSpacingRule**

返回或设置指定段落格式的行距。可读写 **WdLineSpacing** 类型。

**语法**

**express.LineSpacingRule**

*express*   一个代表 **ParagraphFormat** 对象的变量。

**说明**

使用 **wdLineSpaceSingle**、**wdLineSpace1pt5** 或 **wdLineSpaceDouble** 可将行距设置为这些值之一。要将行距设置为固定磅值或多倍行距，还必须设置 **LineSpacing** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例为活动文档的第一段设置 2 倍行距。*/ ActiveDocument.Paragraphs.Item(1).LineSpacingRule = wdLineSpaceDouble` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例返回所选内容的第一段所用的行距标准。*/ let lrule = Selection.Paragraphs.Item(1).LineSpacingRule` |

#### **ParagraphFormat.LineUnitAfter**

返回或设置指定段落的段后间距（以网格线为单位）。可读写 **Single** 类型。

**语法**

**express.LineUnitAfter**

*express*   一个代表 **ParagraphFormat** 对象的变量。

#### **ParagraphFormat.LineUnitBefore**

返回或设置指定段落的段前间距（以网格线为单位）。可读写 **Single** 类型。

**语法**

**express.LineUnitBefore**

*express*   一个代表 **ParagraphFormat** 对象的变量。

#### **ParagraphFormat.MirrorIndents**

返回或设置一个 **Long** 类型的值，该值代表左缩进和右缩进的宽度是否相同。该属性值可以是 **True**、**False** 或 **wdUndefined**。可读写。

**语法**

**express.MirrorIndents**

*express*   一个代表 **ParagraphFormat** 对象的变量。

#### **ParagraphFormat.NoLineNumber**

如果取消指定段的行号，则该属性值为 **True**。该属性值可以是 **True**、**False** 或 **wdUndefined**。可读写 **Long** 类型。

**语法**

**express.NoLineNumber**

*express*   一个代表 **ParagraphFormat** 对象的变量。

**说明**

使用 **PageSetup** 对象的 **LineNumbering** 属性可打开行号。

#### **ParagraphFormat.OutlineLevel**

返回或设置指定段落的大纲级别。可读写 **WdOutlineLevel** 类型。

**语法**

**express.OutlineLevel**

*express*   一个代表 **ParagraphFormat** 对象的变量。

**说明**

如果段落应用了一种标题样式（从“标题 1”到“标题 9”），则大纲级别与该标题样式相同，并且不能更改。大纲级别仅在大纲视图或文档结构图窗格中可见。

#### **ParagraphFormat.PageBreakBefore**

如果在指定段落前插入了分页符，则该属性值为 **True**。该属性值可以是 **True**、**False** 或 **wdUndefined**。可读写 **Long** 类型。

**语法**

**express.PageBreakBefore**

*express*   一个代表 **ParagraphFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在所选内容的第一段前插入分页符。*/ Selection.Paragraphs.Item(1).PageBreakBefore = true` |

#### **ParagraphFormat.Parent**

返回一个 **Object** 类型值，该值代表指定 **ParagraphFormat** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **ParagraphFormat** 对象的变量。

#### **ParagraphFormat.ReadingOrder**

返回或设置指定段落的读取次序而不改变其对齐方式。可读写 **WdReadingOrder** 类型。

**语法**

**express.ReadingOrder**

*express*   一个代表 **ParagraphFormat** 对象的变量。

**说明**

使用 **Selection** 对象的 **LtrPara**、**LtrRun**、**RtlPara** 和 **RtlRun** 方法可更改段落的对齐方式和读取次序。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将第一段的读取次序设置为从右向左。*/ ActiveDocument.Paragraphs.Item(1).ReadingOrder = wdReadingOrderRtl` |

#### **ParagraphFormat.RightIndent**

返回或设置指定段落的右缩进量（以磅为单位）。可读写 **Single** 类型。

**语法**

**express.RightIndent**

*express*   一个代表 **ParagraphFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动文档中所有段落的右缩进都设置为距右边距 1 英寸。InchesToPoints 方法用来将英寸转换为磅。*/ ActiveDocument.Paragraphs.RightIndent = InchesToPoints(1)` |

#### **ParagraphFormat.Shading**

返回一个 **Shading** 对象，该对象代表指定对象的底纹格式。

**语法**

**express.Shading**

*express*   一个代表 **ParagraphFormat** 对象的变量。

#### **ParagraphFormat.SpaceAfter**

返回或设置指定段落或文本栏后面的间距（以磅为单位）。可读/写 **Single** 类型。

**语法**

**express.SpaceAfter**

*express*   一个代表 **ParagraphFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档中所有段落的段后间距设置为 12 磅。*/ ActiveDocument.Range().ParagraphFormat.SpaceAfter = 12` |

#### **ParagraphFormat.SpaceAfterAuto**

如果 WPS 自动设置指定段落的段后间距，则该属性为 **True**。可读/写 **Long** 类型。

**语法**

**express.SpaceAfterAuto**

*express*   一个代表 **ParagraphFormat** 对象的变量。

**说明**

如果仅将部分指定段落的 **SpaceAfterAuto** 属性设置为 **True**，则返回 **wdUndefined**。该属性可被设置为 **True** 或 **False**。

如果将 **SpaceAfterAuto** 设置为 **True**，则 **SpaceAfter** 属性将被忽略。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示关于活动文档的 SpaceAfterAuto 属性设置的报告。*/ function test() { let saa = ActiveDocument.Range().ParagraphFormat.SpaceAfterAuto switch(saa) {     case 1: 	MsgBox("Spacing after paragraphs is handled automatically for all paragraphs.")     case 0: 	MsgBox("Spacing after paragraphs is handled manually for all paragraphs.")     case null: 	MsgBox("Spacing after paragraphs is handled automatically for some paragraphs,manually for some paragraphs.") } }` |

#### **ParagraphFormat.SpaceBefore**

返回或设置指定段落的段前间距（以磅为单位）。可读/写 **Single** 类型。

**语法**

**express.SpaceBefore**

*express*   一个代表 **ParagraphFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档中所有段落的段前间距设置为 12 磅。*/ ActiveDocument.Range().ParagraphFormat.SpaceBefore = 12` |

#### **ParagraphFormat.SpaceBeforeAuto**

如果 WPS 自动设置指定段落的段前间距，则该属性为 **True**。可读/写 **Long** 类型。

**语法**

**express.SpaceBeforeAuto**

*express*   一个代表 **ParagraphFormat** 对象的变量。

**说明**

如果仅将部分指定段落的 **SpaceBeforeAuto** 属性设置为 **True**，则返回 **wdUndefined**。该属性可被设置为 **True** 或 **False**。

如果将 **SpaceBeforeAuto** 设置为 **True**，则 **SpaceBefore** 属性将被忽略。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示关于活动文档的 SpaceBeforeAuto 属性设置的报告。*/ function test() { let saa = ActiveDocument.Range().ParagraphFormat.SpaceBeforeAuto switch(saa) {     case 1: 	MsgBox( "Spacing after paragraphs is handled automatically for all paragraphs.")     case 0: 	MsgBox( "Spacing after paragraphs is handled manually for all paragraphs.")     case null: 	MsgBox("Spacing after paragraphs is handled automatically for some paragraphs,manually for some paragraphs.") } }` |

#### **ParagraphFormat.Style**

返回或设置指定对象的样式。可读写 **Variant** 类型。

**语法**

**express.Style**

*express*   一个代表 **ParagraphFormat** 对象的变量。

**说明**

要设置该属性，请指定样式的本地名称、一个整数、一个 **WdBuiltinStyle** 常量或一个代表样式的对象。

#### **ParagraphFormat.TabStops**

返回或设置一个 **TabStops** 集合，该集合代表指定段落中的所有自定义制表位。可读写。

**语法**

**express.TabStops**

*express*   一个代表 **ParagraphFormat** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在活动文档各段的 2 英寸处添加居中的制表位。InchesToPoints 方法用来将英寸转换为磅。*/ ActiveDocument.Paragraphs.TabStops.Add(InchesToPoints(2), wdAlignTabCenter)` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将文档中各段的制表位设置为与第一段的制表位一致。*/ function test() { let para1Tabs = ActiveDocument.Paragraphs.Item(1).TabStops ActiveDocument.Paragraphs.TabStops = para1Tabs }` |

#### **ParagraphFormat.TextboxTightWrap**

返回或设置一个 **WdTextboxTightWrap** 常量，该常量代表文本环绕形状或文本框的紧密程度。可读写。

**语法**

**express.TextboxTightWrap**

*express*   一个代表 **ParagraphFormat** 对象的变量。

#### **ParagraphFormat.WidowControl**

在 WPS 对文档重新分页时，如果指定段落的首行和末行与段落的其他各行同页，则该属性值为 **True**。该属性值可以是 **True**、**False** 或 **wdUndefined**。可读写 **Long** 类型。

**语法**

**express.WidowControl**

*express*   一个代表 **ParagraphFormat** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例给活动文档中的各段落设置格式，从而使段中的首行或末行有可能单独位于上页的页尾或下页的页首。*/ ActiveDocument.Paragraphs.WidowControl = false` |

#### **ParagraphFormat.WordWrap**

如果 WPS 在指定段落或文本框架的西文单词中间断字换行，则该属性值为 **True**。可读写 **Long** 类型。

**语法**

**express.WordWrap**

*express*   一个代表 **ParagraphFormat** 对象的变量。

**说明**

如果只有一些指定段落或文本框架的此属性设置为 **True**，则此属性返回 **wdUndefined**。根据您所选择或安装的语言支持（例如，美国英语），这种用法可能不可用。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例设置 WPS 在活动文档的第一段中的西文单词中间断字换行。*/ ActiveDocument.Paragraphs.Item(1).WordWrap = true` |

适用环境：web

适用平台：windows/linux