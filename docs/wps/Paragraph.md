#### **Paragraph**



代表所选内容、范围或文档中的一个段落。**Paragraph** 对象是 **Paragraphs** 集合的一个成员。**Paragraphs** 集合包含所选内容、范围或文档中的所有段落。

**说明**

使用 **Paragraphs**(*Index*) 可返回一个 **Paragraph** 对象，其中 *Index* 为索引号。以下示例将活动文档中的第一段右对齐。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActiveDocument.Paragraphs.Item(1).Alignment = wdAlignParagraphRight` |

使用 

Add

、

InsertParagraph

、

InsertParagraphAfter

 或 

InsertParagraphBefore

 方法可在文档中添加一个新的空白段落。以下示例在所选内容的第一段前添加一个段落标记。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {  let S = Application.Selection.Paragraphs.Item(1).Range    Application.Selection.Paragraphs.Add(S) } ` |

以下示例同样在所选内容的第一段前添加一个段落标记。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.Selection.Paragraphs.Item(1).Range.InsertParagraphBefore()` |

**方法**

|                                                              | 名称                         | 说明                                                         |
| ------------------------------------------------------------ | ---------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CloseUp**                  | 清除指定段落前的任何间距。                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Indent**                   | 为一个或多个段落增加一个级别的缩进。                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **IndentCharWidth**          | 将段落缩进指定的字符数。                                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **IndentFirstLineCharWidth** | 将一个或多个段落的首行缩进指定的字符数。                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **JoinList**                 | 将列表段落与距指定段落前后最近的列表进行连接。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ListAdvanceTo**            | 为列表中的段落设置列表级别。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ListNumberOriginal**       | 返回一个 **Integer** 类型的值，该值代表段落的原始列表级别。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Next**                     | 返回一个 **Paragraph** 对象，该对象代表下一段。              |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **OpenOrCloseUp**            | 切换段前间距。                                               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **OpenUp**                   | 为指定段落设置 12 磅的段前间距。                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Outdent**                  | 为一个或多个段落删除一个级别的缩进。                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **OutlineDemote**            | 对指定段落应用下一个标题级别样式（从“标题 1”到“标题 8”）。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **OutlineDemoteToBody**      | 通过应用“正文”样式将指定段落降级为正文文本。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **OutlinePromote**           | 对指定段落应用上一个标题级别样式（从“标题 1”到“标题 8”）。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Previous**                 | 将上一段作为一个 **Paragraph** 对象返回。                    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Reset**                    | 删除手动段落格式（不使用样式应用的格式）。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ResetAdvanceTo**           | 将使用自定义列表级别的段落重置为原始级别设置。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SelectNumber**             | 选择列表中的编号或项目符号。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SeparateList**             | 将列表分隔为两个单独的列表。对于编号的列表，新列表将从起始编号（通常为 1）开始重新编号。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Space1**                   | 为指定段落设置单倍行距。                                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Space15**                  | 为指定段落设置 1.5 倍行距。                                  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Space2**                   | 为指定段落设置 2 倍行距。                                    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **TabHangingIndent**         | 将悬挂缩进量设置为指定的制表位数。                           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **TabIndent**                | 将指定段落的左缩进量设置为指定的制表位数。                   |

**属性**

|                                                              | 名称                                | 说明                                                         |
| ------------------------------------------------------------ | ----------------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AddSpaceBetweenFarEastAndAlpha**  | 如果 WPS 将自动在指定段落的日文和拉丁文文字之间添加空格，则该属性值为 **True**。如果仅对于某些指定段落将该属性设置为 **True**，则该属性会返回 **wdUndefined**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AddSpaceBetweenFarEastAndDigit**  | 如果 WPS 将自动在指定段落的日文文字和数字之间添加空格，则该属性值为 **True**。如果仅对于某些指定段落将该属性设置为 **True**，则该属性会返回 **wdUndefined**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Alignment**                       | 返回或设置一个 **WdParagraphAlignment** 常量，该常量代表指定段落的对齐方式，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**                     | 返回一个代表 WPS 应用程序的 **Application** 对象。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoAdjustRightIndent**           | 如果 WPS 会根据您指定的每行字符数自动调整指定段落的右缩进，则该属性值为 **True**。如果只将某些指定段落的 **AutoAdjustRightIndent** 属性设置为 **True**，则该属性会返回 **wdUndefined**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BaseLineAlignment**               | 返回或设置一个 **WdBaselineAlignment** 常量，该常量代表行中字体的垂直位置，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Borders**                         | 返回一个 **Borders** 集合，该集合代表指定段落的所有边框。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CharacterUnitFirstLineIndent**    | 返回或设置首行或悬挂缩进的值（以字符为单位）。用正值设置首行缩进，用负值设置悬挂缩进。**Single** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CharacterUnitLeftIndent**         | 该属性返回或设置指定段落的左缩进量（以字符为单位）。**Single** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CharacterUnitRightIndent**        | 该属性返回或设置指定段落的右缩进量（以字符为单位）。**Single** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**                         | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisableLineHeightGrid**           | 如果该属性的值为 **True**，则当指定每页的行数时，WPS 会将指定段落中的字符与行网格对齐。如果只将某些指定段落的 **DisableLineHeightGrid** 属性设置为 **True**，则返回 **wdUndefined**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DropCap**                         | 返回一个 **DropCap** 对象，该对象代表指定段落中格式为首字下沉的大写字母。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FarEastLineBreakControl**         | 如果为 **True**，则 WPS 会将东亚语言文字的换行规则应用于指定的段落。如果只将某些指定段落的 **FarEastLineBreakControl** 属性设定为 **True**，则返回 **wdUndefined**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FirstLineIndent**                 | 返回或设置首行缩进或悬挂缩进的大小（以磅值表示）。用正数设置首行缩进的尺寸，用负数设置悬挂缩进的尺寸。**Single** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Format**                          | 返回或设置一个 **ParagraphFormat** 对象，该对象代表指定的一个或多个段落的格式。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HalfWidthPunctuationOnTopOfLine** | 如果为 **True**，则 WPS 会将指定段落行首的标点符号改为半角字符。如果仅将某些指定段落的该属性设置为 **True**，则此属性将返回 **wdUndefined**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HangingPunctuation**              | 如果为 **True**，则指定段落中的标点将可以溢出边界。如果仅将某些指定段落的该属性设置为 **True**，则返回 **wdUndefined**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Hyphenation**                     | 如果指定的段落进行自动断字，则该属性值为 **True**。如果指定的段落不进行自动断字，则该属性值为 **False**。可读写 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ID**                              | 在当前文档另存为网页时，返回或设置指定对象的标识标签。可读写 **String** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IsStyleSeparator**                | 如果段落包含特殊隐藏段落标记，而利用该标记 WPS 可以显示合并不同段落样式的段落，则该属性值为 **True**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **KeepTogether**                    | 在 WPS 对文档重新分页时，如果指定段落中的所有行都位于同一页上，则该属性值为 **True**。可读写 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **KeepWithNext**                    | 在 WPS 对文档重新分页时，如果指定段落与它的下一段位于同一页上，则该属性值为 **True**。可读写 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LeftIndent**                      | 返回或设置一个 **Single** 类型的值，该值代表指定段落的左缩进值（以磅为单位）。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LineSpacing**                     | 返回或设置指定段落的行距（以磅为单位）。**Single** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LineSpacingRule**                 | 返回或设置指定段落的行距。可读写 **WdLineSpacing** 类型。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LineUnitAfter**                   | 返回或设置指定段落的段后间距（以网格线为单位）。可读写 **Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LineUnitBefore**                  | 返回或设置指定段落的段前间距（以网格线为单位）。可读写 **Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MirrorIndents**                   | 返回或设置一个 **Long** 类型的值，该值代表左缩进和右缩进的宽度是否相同。该属性值可以是 **True**、**False** 或 **wdUndefined**。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NoLineNumber**                    | 如果取消指定段落的行号，则该属性值为 **True**。可读写 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OutlineLevel**                    | 返回或设置指定段落的大纲级别。可读写 **WdOutlineLevel** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PageBreakBefore**                 | 如果在指定段落前插入了分页符，则该属性值为 **True**。可读写 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**                          | 返回一个 **Object** 类型值，该值代表指定 **Paragraph** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Range**                           | 返回一个 **Range** 对象，该对象代表指定段落中包含的文档部分。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ReadingOrder**                    | 返回或设置指定段落的读取次序而不改变其对齐方式。可读写 **WdReadingOrder** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RightIndent**                     | 返回或设置指定段落的右缩进（以磅为单位）。可读写 **Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Shading**                         | 返回一个 **Shading** 对象，该对象引用指定段落的底纹格式。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SpaceAfter**                      | 返回或设置指定段落或文本栏后面的间距（以磅为单位）。可读/写 **Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SpaceAfterAuto**                  | 如果 WPS 自动设置指定段落的段后间距，则该属性为 **True**。可读/写 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SpaceBefore**                     | 返回或设置指定段落的段前间距（以磅为单位）。可读/写 **Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SpaceBeforeAuto**                 | 如果 WPS 自动设置指定段落的段前间距，则该属性为 **True**。可读/写 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Style**                           | 返回或设置指定对象的样式。可读写 **Variant** 类型。          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TabStops**                        | 返回或设置 **TabStops** 集合，该集合表示指定段落的所有自定义制表位。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextboxTightWrap**                | 返回或设置一个 **WdTextboxTightWrap** 常量，该常量代表文本环绕形状或文本框的紧密程度。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **WidowControl**                    | 在 WPS 对文档重新分页时，如果指定段落的首行和末行与段落的其他各行同页，则该属性值为 **True**。可读写 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **WordWrap**                        | 如果 WPS 在指定段落或文本框架的西文单词中间断字换行，则该属性值为 **True**。可读写 **Long** 类型。 |

**成员方法**

#### **Paragraph.CloseUp**

清除指定段落前的任何间距。

**语法**

**express.CloseUp()**

*express*   一个代表 **Paragraph** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下两个语句是等价的：*/ function test() { ActiveDocument.Paragraphs.Item(1).CloseUp() ActiveDocument.Paragraphs.Item(1).SpaceBefore = 0 }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例清除选定内容中首段的段前间距。*/ Selection.Paragraphs.Item(1).CloseUp()` |

#### **Paragraph.Indent**

为一个或多个段落增加一个级别的缩进。

**语法**

**express.Indent()**

*express*   一个代表 **Paragraph** 对象的变量。

**说明**

此方法等效于单击**“格式”**工具栏上的**“增加缩进量”**按钮。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例为活动文档中的所有段落增加两个级别的缩进，然后为第一段删除一个级别的缩进。*/ function test() { ActiveDocument.Paragraphs.Indent() ActiveDocument.Paragraphs.Indent() ActiveDocument.Paragraphs.Item(1).Outdent() }` |

#### **Paragraph.IndentCharWidth**

将段落缩进指定的字符数。

**语法**

**express.IndentCharWidth(Count)**

*express*   一个代表 **Paragraph** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                 |
| -------- | ------------- | ------------ | ------------------------ |
| *Count*  | 必选          | **Integer**  | 指定段落要缩进的字符数。 |

**说明**

此方法等效于单击**“格式”**工具栏上的**“增加缩进量”**按钮。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动文档的第一段缩进 10 个字符。*/ ActiveDocument.Paragraphs.Item(1).IndentCharWidth(10)` |

#### **Paragraph.IndentFirstLineCharWidth**

将一个或多个段落的首行缩进指定的字符数。

**语法**

**express.IndentFirstLineCharWidth(Count)**

*express*   一个代表 **Paragraph** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                           |
| -------- | ------------- | ------------ | ---------------------------------- |
| *Count*  | 必选          | **Integer**  | 每个指定段落的首行要缩进的字符数。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动文档中第一段的首行缩进 10 个字符。*/ Application.ActiveDocument.Paragraphs.Item(1).IndentCharWidth(10)` |

#### **Paragraph.JoinList**

将列表段落与距指定段落前后最近的列表进行连接。

**语法**

**express.JoinList()**

*express*   一个代表 **Paragraph** 对象的变量。

#### **Paragraph.ListAdvanceTo**

为列表中的段落设置列表级别。

**语法**

**express.ListAdvanceTo(Level1 to Level9)**

*express*   一个代表 **Paragraph** 对象的变量。

**参数**

| **名称**           | **必选/可选** | **数据类型** | **说明**       |
| ------------------ | ------------- | ------------ | -------------- |
| *Level1 to Level9* | 可选          | **Integer**  | 指定列表级别。 |

**说明**

WPS 通过在指定段落与上一段落之间以隐藏文字形式插入必要的干预段落，来调整用于对指定段落的每个列表级别进行编号的编号值。

#### **Paragraph.ListNumberOriginal**

返回一个 **Integer** 类型的值，该值代表段落的原始列表级别。只读。

**语法**

**express.ListNumberOriginal(Level)**

*express*   一个代表 **Paragraph** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**       |
| -------- | ------------- | ------------ | -------------- |
| *Level*  | 必选          | **Integer**  | 指定列表级别。 |

#### **Paragraph.Next**

返回一个 **Paragraph** 对象，该对象代表下一段。

**语法**

**express.Next(Count)**

*express*   一个代表 **Paragraph** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                     |
| -------- | ------------- | ------------ | ---------------------------- |
| *Count*  | 可选          | **Variant**  | 要前移的段落数。默认值为 1。 |

**返回值**

Paragraph

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在活动文档中的前九段前面分别插入一个数字和一个制表符。*/ function test() { for(let n = 0; n <= 8; n++) {     let myRange = ActiveDocument.Paragraphs.Item(1).Next(n).Range     myRange.Collapse(wdCollapseStart)     myRange.InsertAfter(n + 1 + "\t") } }` |

#### **Paragraph.OpenOrCloseUp**

切换段前间距。

**语法**

**express.OpenOrCloseUp()**

*express*   一个代表 **Paragraph** 对象的变量。

**说明**

如果指定段落的段前间距等于 0（零），则此方法将其设置为 12 磅；如果指定段落的段前间距大于 0（零），则此方法将其设置为 0（零）。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例切换活动文档中第一段的格式，或增加 12 磅的段前间距，或不设置段前间距。*/ ActiveDocument.Paragraphs.Item(1).OpenOrCloseUp()` |

#### **Paragraph.OpenUp**

为指定段落设置 12 磅的段前间距。

**语法**

**express.OpenUp()**

*express*   一个代表 **Paragraph** 对象的变量。

**说明**

您还可以使用 **SpaceBefore** 属性设置段落间距。下列两个语句是等效的：

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() { ActiveDocument.Paragraphs.Item(1).OpenUp() ActiveDocument.Paragraphs.Item(1).SpaceBefore = 12 }  ` |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例更改活动文档的第二段的格式，为该段落设置 12 磅的段前间距。*/ ActiveDocument.Paragraphs.Item(2).OpenUp()` |

#### **Paragraph.Outdent**

为一个或多个段落删除一个级别的缩进。

**语法**

**express.Outdent()**

*express*   一个代表 **Paragraph** 对象的变量。

**说明**

此方法等效于单击**“格式”**工具栏上的**“减少缩进量”**按钮。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例为活动文档中的所有段落增加两个级别的缩进，然后为第一段删除一个级别的缩进。*/ function test() { ActiveDocument.Paragraphs.Indent() ActiveDocument.Paragraphs.Indent() ActiveDocument.Paragraphs.Item(1).Outdent() }` |

#### **Paragraph.OutlineDemote**

对指定段落应用下一个标题级别样式（从“标题 1”到“标题 8”）。

**语法**

**express.OutlineDemote()**

*express*   一个代表 **Paragraph** 对象的变量。

**说明**

如果一个段落的样式为“标题 2”，则此方法将该段落的标题样式降级到“标题 3”。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将所选内容中的第一段降级。*/ Selection.Paragraphs.Item(1).OutlineDemote()` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动文档中的第三段降级。*/ ActiveDocument.Paragraphs.Item(3).OutlineDemote()` |

#### **Paragraph.OutlineDemoteToBody**

通过应用“正文”样式将指定段落降级为正文文本。

**语法**

**express.OutlineDemoteToBody()**

*express*   一个代表 **Paragraph** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将所选内容中的第一段降级为正文文本。*/ Selection.Paragraphs.Item(1).OutlineDemoteToBody()` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动窗口切换至大纲视图，并将所选内容中的第一段降级为正文文本。*/ function test() { ActiveDocument.ActiveWindow.View.Type = wdOutlineView Selection.Paragraphs.Item(1).OutlineDemoteToBody() }` |

#### **Paragraph.OutlinePromote**

对指定段落应用上一个标题级别样式（从“标题 1”到“标题 8”）。

**语法**

**express.OutlinePromote()**

*express*   一个代表 **Paragraph** 对象的变量。

**说明**

如果一个段落的样式为“标题 2”，则此方法将该段落的样式升级到“标题 1”。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将所选内容中的第一段升级。*/ Selection.Paragraphs.Item(1).OutlinePromote()` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动窗口切换至大纲视图，并将活动文档中的第一段升级。*/ function test() { ActiveDocument.ActiveWindow.View.Type = wdOutlineView ActiveDocument.Paragraphs.Item(1).OutlinePromote() }` |

#### **Paragraph.Previous**

将上一段作为一个 **Paragraph** 对象返回。

**语法**

**express.Previous(Count)**

*express*   一个代表 **Paragraph** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                     |
| -------- | ------------- | ------------ | ---------------------------- |
| *Count*  | 可选          | **Variant**  | 要后移的段落数。默认值为 1。 |

**返回值**

Paragraph

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例选择活动文档中所选内容的上一段。*/ Selection.Previous(wdParagraph, 1).Select()` |

#### **Paragraph.Reset**

删除手动段落格式（不使用样式应用的格式）。

**语法**

**express.Reset()**

*express*   一个代表 **Paragraph** 对象的变量。

**说明**

如果手动将段落设置为右对齐，而基本样式为其他对齐方式，则用 **Reset** 方法将手动对齐方式改为与基本样式相同。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例删除活动文档第二段的手动设定的段落格式。*/ ActiveDocument.Paragraphs.Item(2).Reset()` |

#### **Paragraph.ResetAdvanceTo**

将使用自定义列表级别的段落重置为原始级别设置。

**语法**

**express.ResetAdvanceTo()**

*express*   一个代表 **Paragraph** 对象的变量。

**说明**

WPS 通过删除指定段落与上一段落之间所有隐藏的干预段落，来调整用于对指定段落的每个列表级别进行编号的编号值。

#### **Paragraph.SelectNumber**

选择列表中的编号或项目符号。

**语法**

**express.SelectNumber()**

*express*   一个代表 **Paragraph** 对象的变量。

**说明**

如果在不包括列表的段落、选定内容或区域中调用 **SelectNumber** 方法，则显示一条错误消息。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例选定列表中的项目符号或编号（该列表包括活动文档中的选定段落），并将项目符号或编号的字号增加至 17 磅。本示例假定选定内容中的第一段被设为项目符号或编号列表。*/ function SelectParaNumber(){     Selection.Paragraphs.Item(1).SelectNumber()     Selection.Font.Size = 17 }` |

#### **Paragraph.SeparateList**

将列表分隔为两个单独的列表。对于编号的列表，新列表将从起始编号（通常为 1）开始重新编号。

**语法**

**express.SeparateList()**

*express*   一个代表 **Paragraph** 对象的变量。

#### **Paragraph.Space1**

为指定段落设置单倍行距。

**语法**

**express.Space1()**

*express*   一个代表 **Paragraph** 对象的变量。

**说明**

精确间距取决于各段内最大字符的字号。

您还可以使用 **LineSpacingRule** 属性设置段落的行距。下列两个语句是等效的：

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() { ActiveDocument.Paragraphs.Item(1).Space1() ActiveDocument.Paragraphs.Item(1).LineSpacingRule = wdLineSpaceSingle }` |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动文档中第一段的行距更改为单倍行距。*/ ActiveDocument.Paragraphs.Item(1).Space1()` |

#### **Paragraph.Space15**

为指定段落设置 1.5 倍行距。

**语法**

**express.Space15()**

*express*   一个代表 **Paragraph** 对象的变量。

**说明**

精确间距为各段内最大字符的字号加上 6 磅。

您还可以使用 **LineSpacingRule** 属性设置段落的行距。下列两个语句是等效的：

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() { ActiveDocument.Paragraphs.Item(1).Space15() ActiveDocument.Paragraphs.Item(1).LineSpacingRule = wdLineSpace1pt5 }  ` |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动文档中第一段的行距更改为 1.5 倍行距。*/ ActiveDocument.Paragraphs.Item(1).Space15()` |

#### **Paragraph.Space2**

为指定段落设置 2 倍行距。

**语法**

**express.Space2()**

*express*   一个代表 **Paragraph** 对象的变量。

**说明**

精确间距为各段内最大字符的字号加上 12 磅。

您还可以使用 **LineSpacingRule** 属性设置段落的行距。下列两个语句是等效的：

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() { ActiveDocument.Paragraphs.Item(1).Space2() ActiveDocument.Paragraphs.Item(1).LineSpacingRule = wdLineSpaceDouble }  ` |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将所选内容中第一段的行距更改为 2 倍行距。*/ Selection.Paragraphs.Item(1).Space2()` |

#### **Paragraph.TabHangingIndent**

将悬挂缩进量设置为指定的制表位数。

**语法**

**express.TabHangingIndent(Count)**

*express*   一个代表 **Paragraph** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Count*  | 必选          | **Integer**  | 如果为正，则表示要缩进的制表位数；如果为负，则表示要删除的缩进制表位数。 |

**说明**

如果 *Count* 的值为负数，您还可以使用此方法删除悬挂缩进的制表位。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动文档中的第一段悬挂缩进到第二个制表位。*/ ActiveDocument.Paragraphs.Item(1).TabHangingIndent(2)` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动文档中第一段的悬挂缩进量减少一个制表位。*/ ActiveDocument.Paragraphs.Item(1).TabHangingIndent(-1)` |

#### **Paragraph.TabIndent**

将指定段落的左缩进量设置为指定的制表位数。

**语法**

**express.TabIndent(Count)**

*express*   一个代表 **Paragraph** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Count*  | 必选          | **Integer**  | 如果为正，则表示要缩进的制表位数；如果为负，则表示要删除的缩进制表位数。 |

**说明**

如果 *Count* 的值为负数，您还可以使用此方法删除缩进。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动文档中的第一段缩进到第二个制表位。*/ ActiveDocument.Paragraphs.Item(1).TabIndent(2)` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动文档中第一段的缩进量减少一个制表位。*/ ActiveDocument.Paragraphs.Item(1).TabIndent(-1)` |

**成员属性**

#### **Paragraph.AddSpaceBetweenFarEastAndAlpha**

如果 WPS 将自动在指定段落的日文和拉丁文文字之间添加空格，则该属性值为 **True**。如果仅对于某些指定段落将该属性设置为 **True**，则该属性会返回 **wdUndefined**。**Long** 类型，可读写。

**语法**

**express.AddSpaceBetweenFarEastAndAlpha**

*express*   一个代表 **Paragraph** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例实现的功能是：设置 WPS 在活动文档第一段的日文文字和拉丁文字之间自动添加空格。*/ ActiveDocument.Paragraphs.Item(1).AddSpaceBetweenFarEastAndAlpha = true` |

#### **Paragraph.AddSpaceBetweenFarEastAndDigit**

如果 WPS 将自动在指定段落的日文文字和数字之间添加空格，则该属性值为 **True**。如果仅对于某些指定段落将该属性设置为 **True**，则该属性会返回 **wdUndefined**。**Long** 类型，可读写。

**语法**

**express.AddSpaceBetweenFarEastAndDigit**

*express*   一个代表 **Paragraph** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例可实现的功能是：设定 WPS 自动在活动文档第一段的中文文字和数字之间添加空格。*/ ActiveDocument.Paragraphs.Item(1).AddSpaceBetweenFarEastAndDigit = true` |

#### **Paragraph.Alignment**

返回或设置一个 **WdParagraphAlignment** 常量，该常量代表指定段落的对齐方式，可读写。

**语法**

**express.Alignment**

*express*   一个代表 **Paragraph** 对象的变量。

**说明**

对于所选择或安装的不同语言支持（例如美国英语），有些 **WdParagraphAlignment** 常量可能不可用。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档中的首段设置为右对齐。*/ function test(){     Application.ActiveDocument.Paragraphs.Item(1).Alignment = wdAlignParagraphRight }` |

#### **Paragraph.Application**

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

*express*   一个代表 **Paragraph** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **Paragraph.AutoAdjustRightIndent**

如果 WPS 会根据您指定的每行字符数自动调整指定段落的右缩进，则该属性值为 **True**。如果只将某些指定段落的 **AutoAdjustRightIndent** 属性设置为 **True**，则该属性会返回 **wdUndefined**。**Long** 类型，可读写。

**语法**

**express.AutoAdjustRightIndent**

*express*   一个代表 **Paragraph** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例实现的功能是：设定 WPS 根据指定的每行字符数，自动调整所选段落的右缩进。*/ Selection.ParagraphFormat.AutoAdjustRightIndent = true` |

#### **Paragraph.BaseLineAlignment**

返回或设置一个 **WdBaselineAlignment** 常量，该常量代表行中字体的垂直位置，可读写。

**语法**

**express.BaseLineAlignment**

*express*   一个代表 **Paragraph** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 自动调整活动文档中的基线字体对齐方式。*/ ActiveDocument.BaseLineAlignment = wdBaselineAlignAuto` |

#### **Paragraph.Borders**

返回一个 **Borders** 集合，该集合代表指定段落的所有边框。

**语法**

**express.Borders**

*express*   一个代表 **Paragraph** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

#### **Paragraph.CharacterUnitFirstLineIndent**

返回或设置首行或悬挂缩进的值（以字符为单位）。用正值设置首行缩进，用负值设置悬挂缩进。**Single** 类型，可读写。

**语法**

**express.CharacterUnitFirstLineIndent**

*express*   一个代表 **Paragraph** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档中第一段的首行缩进设为一个字符。*/ Application.ActiveDocument.Paragraphs.Item(1).CharacterUnitFirstLineIndent = 1  /*本示例将活动文档中第二段的悬挂缩进设为 1.5 个字符。*/ Application.ActiveDocument.Paragraphs.Item(2).CharacterUnitFirstLineIndent = -1.5` |

#### **Paragraph.CharacterUnitLeftIndent**

该属性返回或设置指定段落的左缩进量（以字符为单位）。**Single** 类型，可读写。

**语法**

**express.CharacterUnitLeftIndent**

*express*   一个代表 **Paragraph** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档中第一段的左缩进设为从左边距缩进一个字符。*/ ActiveDocument.Paragraphs.Item(1).CharacterUnitLeftIndent = 1 ` |

#### **Paragraph.CharacterUnitRightIndent**

该属性返回或设置指定段落的右缩进量（以字符为单位）。**Single** 类型，可读写。

**语法**

**express.CharacterUnitRightIndent**

*express*   一个代表 **Paragraph** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档中的所有段落的右缩进设为从右边距缩进一个字符。*/ ActiveDocument.Paragraphs.CharacterUnitRightIndent = 1` |

#### **Paragraph.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Paragraph** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **Paragraph.DisableLineHeightGrid**

如果该属性的值为 **True**，则当指定每页的行数时，WPS 会将指定段落中的字符与行网格对齐。如果只将某些指定段落的 **DisableLineHeightGrid** 属性设置为 **True**，则返回 **wdUndefined**。**Long** 类型，可读写。

**语法**

**express.DisableLineHeightGrid**

*express*   一个代表 **Paragraph** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例实现的功能是：设定 WPS 在已指定每页行数的情况下，将选定段落中的字符与行网格对齐。*/ Selection.ParagraphFormat.DisableLineHeightGrid = true` |

#### **Paragraph.DropCap**

返回一个 **DropCap** 对象，该对象代表指定段落中格式为首字下沉的大写字母。只读。

**语法**

**express.DropCap**

*express*   一个代表 **Paragraph** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为活动文档首段设置首字下沉。*/ function test() { let dropcap =  ActiveDocument.Paragraphs.Item(1).DropCap dropcap.FontName = "Arial" dropcap.Position = wdDropNormal dropcap.LinesToDrop = 3 dropcap.DistanceFromText = InchesToPoints(0.1) }` |

#### **Paragraph.FarEastLineBreakControl**

如果为 **True**，则 WPS 会将东亚语言文字的换行规则应用于指定的段落。如果只将某些指定段落的 **FarEastLineBreakControl** 属性设定为 **True**，则返回 **wdUndefined**。**Long** 类型，可读写。

**语法**

**express.FarEastLineBreakControl**

*express*   一个代表 **Paragraph** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设定 WPS 将东亚语言文字的换行规则应用于活动文档的第一段。*/ ActiveDocument.Paragraphs.Item(1).FarEastLineBreakControl = true ` |

#### **Paragraph.FirstLineIndent**

返回或设置首行缩进或悬挂缩进的大小（以磅值表示）。用正数设置首行缩进的尺寸，用负数设置悬挂缩进的尺寸。**Single** 类型，可读写。

**语法**

**express.FirstLineIndent**

*express*   一个代表 **Paragraph** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为活动文档的首段设置 1 英寸的首行缩进。*/ ActiveDocument.Paragraphs.Item(1).FirstLineIndent = InchesToPoints(1)` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为活动文档的第二段设置 0.5 英寸的悬挂缩进。InchesToPoints 方法用来将英寸转化为磅值。*/ ActiveDocument.Paragraphs.Item(2).FirstLineIndent = InchesToPoints(-0.5)` |

#### **Paragraph.Format**

返回或设置一个 **ParagraphFormat** 对象，该对象代表指定的一个或多个段落的格式。

**语法**

**express.Format**

*express*   一个代表 **Paragraph** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回活动文档中第一段的格式，然后将该格式应用于文档的选定内容。*/ function test() { let paraFormat = ActiveDocument.Paragraphs.Item(1).Format.Duplicate Selection.Paragraphs.Format = paraFormat }` |

#### **Paragraph.HalfWidthPunctuationOnTopOfLine**

如果为 **True**，则 WPS 会将指定段落行首的标点符号改为半角字符。如果仅将某些指定段落的该属性设置为 **True**，则此属性将返回 **wdUndefined**。**Long** 类型，可读写。

**语法**

**express.HalfWidthPunctuationOnTopOfLine**

*express*   一个代表 **Paragraph** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 将活动文档第一段的行首标点符号改为半角字符。*/ ActiveDocument.Paragraphs.Item(1).HalfWidthPunctuationOnTopOfLine = true` |

#### **Paragraph.HangingPunctuation**

如果为 **True**，则指定段落中的标点将可以溢出边界。如果仅将某些指定段落的该属性设置为 **True**，则返回 **wdUndefined**。**Long** 类型，可读写。

**语法**

**express.HangingPunctuation**

*express*   一个代表 **Paragraph** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例实现的功能是：允许活动文档第一段的标点可以溢出边界。*/ ActiveDocument.Paragraphs.Item(1).HangingPunctuation = true` |

#### **Paragraph.Hyphenation**

如果指定的段落进行自动断字，则该属性值为 **True**。如果指定的段落不进行自动断字，则该属性值为 **False**。可读写 **Long** 类型。

**语法**

**express.Hyphenation**

*express*   一个代表 **Paragraph** 对象的变量。

**说明**

此属性可以是 **True**、**False** 或 **wdUndefined**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例关闭活动文档中第一段的自动断字功能。*/ ActiveDocument.Paragraphs.Item(1).Hyphenation = false` |

#### **Paragraph.ID**

在当前文档另存为网页时，返回或设置指定对象的标识标签。可读写 **String** 类型。

**语法**

**express.ID**

*express*   一个代表 **Paragraph** 对象的变量。

#### **Paragraph.IsStyleSeparator**

如果段落包含特殊隐藏段落标记，而利用该标记 WPS 可以显示合并不同段落样式的段落，则该属性值为 **True**。**Boolean** 类型，只读。

**语法**

**express.IsStyleSeparator**

*express*   一个代表 **Paragraph** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为包含内置“Normal”样式的样式分隔符的所有段落设置格式。*/ function StyleSep(){     let pghDoc = ActiveDocument.Paragraphs     for(let i = 1; i <= pghDoc.Count; i++ ){         if(pghDoc.Item(i).IsStyleSeparator == true){             pghDoc.Item(i).Range.Select()             Selection.Style = "Normal"         }     } }` |

#### **Paragraph.KeepTogether**

在 WPS 对文档重新分页时，如果指定段落中的所有行都位于同一页上，则该属性值为 **True**。可读写 **Long** 类型。

**语法**

**express.KeepTogether**

*express*   一个代表 **Paragraph** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例设置活动文档中第一段的格式，以使在 WPS 对文档重新分页时同一段中的各行都位于同一页上。*/ ActiveDocument.Paragraphs.Item(1).KeepTogether = true` |

#### **Paragraph.KeepWithNext**

在 WPS 对文档重新分页时，如果指定段落与它的下一段位于同一页上，则该属性值为 **True**。可读写 **Long** 类型。

**语法**

**express.KeepWithNext**

*express*   一个代表 **Paragraph** 对象的变量。

**说明**

此属性可以是 **True**、**False** 或 **wdUndefined**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动文档中的第三段至第六段保持在同一页上。*/ function test() { for(let i = 3; i <= 5; i++){     ActiveDocument.Paragraphs.Item(i).KeepWithNext = true } }` |

#### **Paragraph.LeftIndent**

返回或设置一个 **Single** 类型的值，该值代表指定段落的左缩进值（以磅为单位）。可读写。

**语法**

**express.LeftIndent**

*express*   一个代表 **Paragraph** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例为活动文档中的第一段设置 1 英寸的左缩进。InchesToPoints 方法用来将英寸转换为磅。*/ ActiveDocument.Paragraphs.Item(1).LeftIndent = InchesToPoints(1)    ` |

#### **Paragraph.LineSpacing**

返回或设置指定段落的行距（以磅为单位）。**Single** 类型，可读写。

**语法**

**express.LineSpacing**

*express*   一个代表 **Paragraph** 对象的变量。

**说明**

使用 **LinesToPoints** 方法可将行数转换为相应的值（以磅为单位）。例如，`LinesToPoints(2)` 返回值 24。

可以在 **LineSpacingRule** 属性已设置为下列值之后对 **LineSpacing** 属性进行设置：

- **wdLineSpaceAtLeast** 行距可以大于或等于指定的 **LineSpacing** 值，但绝不能小于该值。
- **wdLineSpaceExactly** 指定的 **LineSpacing** 行距值绝不能更改，即使段落中使用了较大的字体。
- **wdLineSpaceMultiple** 必须指定 **LineSpacing** 属性值，以磅为单位。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动文档中第一段的行距设置为总是至少 12 磅。*/ function test() { let pgh = ActiveDocument.Paragraphs.Item(1) pgh.LineSpacingRule = wdLineSpaceAtLeast pgh.LineSpacing = 12 }  ` |

#### **Paragraph.LineSpacingRule**

返回或设置指定段落的行距。可读写 **WdLineSpacing** 类型。

**语法**

**express.LineSpacingRule**

*express*   一个代表 **Paragraph** 对象的变量。

**说明**

使用 **wdLineSpaceSingle**、**wdLineSpace1pt5** 或 **wdLineSpaceDouble** 可将行距设置为这些值之一。要将行距设置为固定磅值或多倍行距，还必须设置 **LineSpacing** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例为活动文档的第一段设置 2 倍行距。*/ ActiveDocument.Paragraphs.Item(1).LineSpacingRule = wdLineSpaceDouble ` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例返回所选内容的第一段所用的行距标准。*/ lrule = Selection.Paragraphs.Item(1).LineSpacingRule` |

#### **Paragraph.LineUnitAfter**

返回或设置指定段落的段后间距（以网格线为单位）。可读写 **Single** 类型。

**语法**

**express.LineUnitAfter**

*express*   一个代表 **Paragraph** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动文档中第一段的段后间距设置为 1 个网格线。*/ ActiveDocument.Paragraphs.Item(1).LineUnitAfter = 1` |

#### **Paragraph.LineUnitBefore**

返回或设置指定段落的段前间距（以网格线为单位）。可读写 **Single** 类型。

**语法**

**express.LineUnitBefore**

*express*   一个代表 **Paragraph** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动文档第二段的段前间距设置为 1 个网格线。*/ ActiveDocument.Paragraphs.Item(2).LineUnitBefore = 1` |

#### **Paragraph.MirrorIndents**

返回或设置一个 **Long** 类型的值，该值代表左缩进和右缩进的宽度是否相同。该属性值可以是 **True**、**False** 或 **wdUndefined**。可读写。

**语法**

**express.MirrorIndents**

*express*   一个代表 **Paragraph** 对象的变量。

#### **Paragraph.NoLineNumber**

如果取消指定段落的行号，则该属性值为 **True**。可读写 **Long** 类型。

**语法**

**express.NoLineNumber**

*express*   一个代表 **Paragraph** 对象的变量。

**说明**

此属性可以是 **True**、**False** 或 **wdUndefined**。使用 **PageSetup** 对象的 **LineNumbering** 属性可设置行号。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例为活动文档设置行号。起始编号设置为 1，并且在文档中的所有节都连续编号。然后取消第二段的行号。*/ function test() { let lnbg = ActiveDocument.PageSetup.LineNumbering lnbg.Active = true lnbg.StartingNumber = 1 lnbg.CountBy = 1 lnbg.RestartMode = wdRestartContinuous ActiveDocument.Paragraphs.Item(2).NoLineNumber = true }` |

#### **Paragraph.OutlineLevel**

返回或设置指定段落的大纲级别。可读写 **WdOutlineLevel** 类型。

**语法**

**express.OutlineLevel**

*express*   一个代表 **Paragraph** 对象的变量。

**说明**

如果段落应用了一种标题样式（从“标题 1”到“标题 9”），则大纲级别与该标题样式相同，并且不能更改。大纲级别仅在大纲视图或文档结构图窗格中可见。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例返回活动文档中第一段的大纲级别。*/ temp = ActiveDocument.Paragraphs.Item(1).OutlineLevel` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例设置活动文档每一段的大纲级别。首先将正文样式应用于所有段落。Mod 操作符用来确定文档中的后继段落采用何种大纲级别（1、2 或 3），然后将视图切换至大纲视图。*/ function test() { let myParas = ActiveDocument.Paragraphs myParas.Style = wdStyleNormal for(let x = 1; x <= myParas.Count; x++){     if(x % 3 == 1){ 	myParas.Item(x).OutlineLevel = wdOutlineLevel1 	}     else if(x % 3 == 2){ 	myParas.Item(x).OutlineLevel = wdOutlineLevel2 	}     else{         myParas.Item(x).OutlineLevel = wdOutlineLevel3     } } ActiveDocument.ActiveWindow.View.Type = wdOutlineView }` |

#### **Paragraph.PageBreakBefore**

如果在指定段落前插入了分页符，则该属性值为 **True**。可读写 **Long** 类型。

**语法**

**express.PageBreakBefore**

*express*   一个代表 **Paragraph** 对象的变量。

**说明**

此属性可以是 **True**、**False** 或 **wdUndefined**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在所选内容的第一段前插入分页符。*/ Selection.Paragraphs.Item(1).PageBreakBefore = true` |

#### **Paragraph.Parent**

返回一个 **Object** 类型值，该值代表指定 **Paragraph** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **Paragraph** 对象的变量。

#### **Paragraph.Range**

返回一个 **Range** 对象，该对象代表指定段落中包含的文档部分。

**语法**

**express.Range**

*express*   一个代表 **Paragraph** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例为活动文档中的第一段应用“标题 1”样式。*/ ActiveDocument.Paragraphs.Item(1).Range.Style = wdStyleHeading1` |

#### **Paragraph.ReadingOrder**

返回或设置指定段落的读取次序而不改变其对齐方式。可读写 **WdReadingOrder** 类型。

**语法**

**express.ReadingOrder**

*express*   一个代表 **Paragraph** 对象的变量。

**说明**

使用 **Selection** 对象的 **LtrPara**、**LtrRun**、**RtlPara** 和 **RtlRun** 方法可更改段落的对齐方式和读取次序。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将第一段的读取次序设置为从右向左。*/ ActiveDocument.Paragraphs.Item(1).ReadingOrder = wdReadingOrderRtl` |

#### **Paragraph.RightIndent**

返回或设置指定段落的右缩进（以磅为单位）。可读写 **Single** 类型。

**语法**

**express.RightIndent**

*express*   一个代表 **Paragraph** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动文档中第一段的右缩进设置为距右边距 1 英寸。InchesToPoints 方法用来将英寸转换为磅。*/ ActiveDocument.Paragraphs.Item(1).RightIndent = InchesToPoints(1)` |

#### **Paragraph.Shading**

返回一个 **Shading** 对象，该对象引用指定段落的底纹格式。

**语法**

**express.Shading**

*express*   一个代表 **Paragraph** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例对所选内容的第一段应用黄色底纹。*/ function test() { let shad = Selection.Paragraphs.Item(1).Shading shad.Texture = wdTexture12Pt5Percent shad.BackgroundPatternColorIndex = wdYellow shad.ForegroundPatternColorIndex = wdBlack }` |

#### **Paragraph.SpaceAfter**

返回或设置指定段落或文本栏后面的间距（以磅为单位）。可读/写 **Single** 类型。

**语法**

**express.SpaceAfter**

*express*   一个代表 **Paragraph** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档第一段的段后间距设置为 12 磅。*/ ActiveDocument.Paragraphs.Item(1).SpaceAfter = 12` |

#### **Paragraph.SpaceAfterAuto**

如果 WPS 自动设置指定段落的段后间距，则该属性为 **True**。可读/写 **Long** 类型。

**语法**

**express.SpaceAfterAuto**

*express*   一个代表 **Paragraph** 对象的变量。

**说明**

如果仅将部分指定段落的 **SpaceAfterAuto** 属性设置为 **True**，则返回 **wdUndefined**。该属性可被设置为 **True** 或 **False**。

如果将 **SpaceAfterAuto** 设置为 **True**，则 **SpaceAfter** 属性将被忽略。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示关于活动文档的第一段的 SpaceAfterAuto 设置的报告。*/ function test() { let saa = ActiveDocument.Paragraphs.Item(1).SpaceAfterAuto switch(saa) {     case 1:         MsgBox ( "Spacing after paragraphs is handled automatically for all paragraphs.")         break     case 0:         MsgBox ( "Spacing after paragraphs is handled manually for all paragraphs.")         break     case undefined:         MsgBox ( "Spacing after paragraphs is handled automatically for some paragraphs, manually for some paragraphs.")         break } }` |

#### **Paragraph.SpaceBefore**

返回或设置指定段落的段前间距（以磅为单位）。可读/写 **Single** 类型。

**语法**

**express.SpaceBefore**

*express*   一个代表 **Paragraph** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档中第二段的段前间距设置为 12 磅。*/ ActiveDocument.Paragraphs.Item(2).SpaceBefore = 12` |

#### **Paragraph.SpaceBeforeAuto**

如果 WPS 自动设置指定段落的段前间距，则该属性为 **True**。可读/写 **Long** 类型。

**语法**

**express.SpaceBeforeAuto**

*express*   一个代表 **Paragraph** 对象的变量。

**说明**

如果仅将部分指定段落的 **SpaceBeforeAuto** 属性设置为 **True**，则该属性返回 **wdUndefined**。该属性可被设置为 **True** 或 **False**。

如果将 **SpaceBeforeAuto** 设置为 **True**，则 **SpaceBefore** 属性将被忽略。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示活动文档中第一段的 SpaceBeforeAuto 设置的报告。*/ function test() { let saa = ActiveDocument.Paragraphs.Item(1).SpaceBeforeAuto switch(saa) {     case 1:         MsgBox ( "Spacing before paragraphs is handled automatically for all paragraphs.")         break     case 0:         MsgBox ( "Spacing before paragraphs is handled manually for all paragraphs.")         break     case undefined:         MsgBox ( "Spacing before paragraphs is handled automatically for some paragraphs, manually for some paragraphs.")         break } }  ` |

#### **Paragraph.Style**

返回或设置指定对象的样式。可读写 **Variant** 类型。

**语法**

**express.Style**

*express*   一个代表 **Paragraph** 对象的变量。

**说明**

要设置该属性，请指定样式的本地名称、一个整数、一个 **WdBuiltinStyle** 常量或一个代表样式的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示活动文档中各段的样式。*/ function test() { let para = ActiveDocument.Paragraphs for(let i = 1; i <= para.Count; i++) {     MsgBox (para.Item(i).Style) } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动文档中所有段落交替设置为“标题 3”和“正文”样式。*/ function test() { let pghc = ActiveDocument.Paragraphs for(let i = 1; i<= pghc.Count; i++){     if(i % 2 ==0){         pghc.Item(i).Style = wdStyleNormal     }     else{         pghc.Item(i).Style = wdStyleHeading3     } } }` |

#### **Paragraph.TabStops**

返回或设置 **TabStops** 集合，该集合表示指定段落的所有自定义制表位。可读写。

**语法**

**express.TabStops**

*express*   一个代表 **Paragraph** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将文档中各段的制表位设置为与第一段的制表位一致。*/ function test() { let para1Tabs = ActiveDocument.Paragraphs.Item(1).TabStops ActiveDocument.Paragraphs.TabStops = para1Tabs }` |

#### **Paragraph.TextboxTightWrap**

返回或设置一个 **WdTextboxTightWrap** 常量，该常量代表文本环绕形状或文本框的紧密程度。可读写。

**语法**

**express.TextboxTightWrap**

*express*   一个代表 **Paragraph** 对象的变量。

#### **Paragraph.WidowControl**

在 WPS 对文档重新分页时，如果指定段落的首行和末行与段落的其他各行同页，则该属性值为 **True**。可读写 **Long** 类型。

**语法**

**express.WidowControl**

*express*   一个代表 **Paragraph** 对象的变量。

**说明**

此属性可以是 **True**、**False** 或 **wdUndefined**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例设置活动文档中第一段的格式，以便段落中的首行或末行可以单独出现在页面顶端或底端。*/ ActiveDocument.Paragraphs.Item(1).WidowControl = false` |

#### **Paragraph.WordWrap**

如果 WPS 在指定段落或文本框架的西文单词中间断字换行，则该属性值为 **True**。可读写 **Long** 类型。

**语法**

**express.WordWrap**

*express*   一个代表 **Paragraph** 对象的变量。

**说明**

如果仅将某些指定段落或文本框架的此属性设置为 **True**，则此属性返回 **wdUndefined**。根据您所选择或安装的语言支持（例如，美国英语），这种用法可能不可用。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例设置 WPS 在活动文档的第一段中的西文单词中间断字换行。*/ ActiveDocument.Paragraphs.Item(1).WordWrap = false` |

适用环境：web

适用平台：windows/linux