#### **Paragraphs**



**方法**

|                                                              | 名称                         | 说明                                                         |
| ------------------------------------------------------------ | ---------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**                      | 返回一个 **Paragraph** 对象，该对象代表添加到文档中的新的空白段落。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Indent**                   | 为一个或多个段落增加一个级别的缩进。                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **IndentFirstLineCharWidth** | 将一个或多个段落的首行缩进指定的字符数。                     |

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
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CharacterUnitLeftIndent**         | 该属性返回或设置指定段落的左缩进量（以字符为单位）。**Single** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CharacterUnitRightIndent**        | 该属性返回或设置指定段落的右缩进量（以字符为单位）。**Single** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**                           | 返回一个 **Long** 类型的值，该值代表集合中的段数。只读。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**                         | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisableLineHeightGrid**           | 如果该属性的值为 **True**，则当指定每页的行数时，WPS 会将指定段落中的字符与行网格对齐。如果只将某些指定段落的 **DisableLineHeightGrid** 属性设置为 **True**，则返回 **wdUndefined**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FarEastLineBreakControl**         | 如果为 **True**，则 WPS 会将东亚语言文字的换行规则应用于指定的段落。如果只将某些指定段落的 **FarEastLineBreakControl** 属性设定为 **True**，则返回 **wdUndefined**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **First**                           | 返回一个 **Paragraph** 对象，该对象代表在 **Paragraphs** 集合中的第一个项目。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FirstLineIndent**                 | 返回或设置首行缩进或悬挂缩进的大小（以磅值表示）。用正数设置首行缩进的尺寸，用负数设置悬挂缩进的尺寸。**Single** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Format**                          | 返回或设置一个 **ParagraphFormat** 对象，该对象代表指定的一个或多个段落的格式。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HalfWidthPunctuationOnTopOfLine** | 如果为 **True**，则 WPS 会将指定段落行首的标点符号改为半角字符。如果仅将某些指定段落的该属性设置为 **True**，则此属性将返回 **wdUndefined**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HangingPunctuation**              | 如果为 **True**，则指定段落中的标点将可以溢出边界。如果仅将某些指定段落的该属性设置为 **True**，则返回 **wdUndefined**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Hyphenation**                     | 如果指定的段落进行自动断字，则该属性值为 **True**。如果指定的段落不进行自动断字，则该属性值为 **False**。可读写 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **KeepTogether**                    | 在 WPS 对文档重新分页时，如果指定段落中的所有行都位于同一页上，则该属性值为 **True**。可读写 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **KeepWithNext**                    | 在 WPS 对文档重新分页时，如果指定段落与其下一段位于同一页上，则该属性值为 **True**。可读写 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Last**                            | 返回一个 **Paragraph** 对象，该对象代表段落集合中的最后一个项目。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LeftIndent**                      | 返回或设置一个 **Single** 类型的值，该值代表指定段落的左缩进值（以磅为单位）。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LineSpacing**                     | 返回或设置指定段落的行距（以磅为单位）。**Single** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LineSpacingRule**                 | 返回或设置指定段落的行距。可读写 **WdLineSpacing** 类型。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LineUnitAfter**                   | 返回或设置指定段落的段后间距（以网格线为单位）。可读写 **Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LineUnitBefore**                  | 返回或设置指定段落的段前间距（以网格线为单位）。可读写 **Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NoLineNumber**                    | 如果取消指定段的行号，则该属性值为 **True**。该属性值可以是 **True**、**False** 或 **wdUndefined**。可读写 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OutlineLevel**                    | 返回或设置指定段落的大纲级别。可读写 **WdOutlineLevel** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PageBreakBefore**                 | 如果在指定段落前插入了分页符，则该属性值为 **True**。该属性值可以是 **True**、**False** 或 **wdUndefined**。可读写 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**                          | 返回一个 **Object** 类型值，该值代表指定 **Paragraphs** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ReadingOrder**                    | 返回或设置指定段落的读取次序而不改变其对齐方式。可读写 **WdReadingOrder** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RightIndent**                     | 返回或设置指定段落的右缩进量（以磅为单位）。可读写 **Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Shading**                         | 返回一个 **Shading** 对象，该对象代表指定段落的底纹格式。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SpaceAfter**                      | 返回或设置指定段落或文本栏后面的间距（以磅为单位）。可读/写 **Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SpaceAfterAuto**                  | 如果 WPS 自动设置指定段落的段后间距，则该属性为 **True**。可读/写 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SpaceBefore**                     | 返回或设置指定段落的段前间距（以磅为单位）。可读/写 **Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SpaceBeforeAuto**                 | 如果 WPS 自动设置指定段落的段前间距，则该属性为 **True**。可读/写 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Style**                           | 返回或设置指定段落的样式。可读写 **Variant** 类型。          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TabStops**                        | 返回或设置一个 **TabStops** 集合，该集合代表指定段落中的所有自定义制表位。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **WidowControl**                    | 在 WPS 对文档重新分页时，如果指定段落的首行和末行与段落的其他各行同页，则该属性值为 **True**。该属性值可以是 **True**、**False** 或 **wdUndefined**。可读写 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **WordWrap**                        | 如果 WPS 在指定段落中的西文单词中间断字换行，则该属性值为 **True**。可读写 **Long** 类型。 |

**成员方法**

#### **Paragraphs.Add**

返回一个 **Paragraph** 对象，该对象代表添加到文档中的新的空白段落。

**语法**

**express.Add(Range)**

*express*   一个代表 **Paragraphs** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                         |
| -------- | ------------- | ------------ | ------------------------------------------------ |
| *Range*  | 可选          | **Variant**  | 要在其前添加新段落的区域。新的段落不替换该区域。 |

**说明**

如果不指定 *Range*，则将新段落添加到选定内容或区域之后，或者添加到文档最后，具体情况取决于 expression 的设置。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在选定内容之后添加一个段落。*/ Application.Selection.Paragraphs.Add()  /*本示例在选定内容中第一段之前添加一个段落标记。*/ Application.Selection.Paragraphs.Add(Selection.Paragraphs.Item(1).Range)  /*本示例在活动文档第二段之前添加一个段落标记。*/ Application.ActiveDocument.Paragraphs.Add(ActiveDocument.Paragraphs.Item(2).Range)  /*本示例在活动文档的末尾添加一个新的段落标记。*/ Application.ActiveDocument.Paragraphs.Add()` |

#### **Paragraphs.Indent**

为一个或多个段落增加一个级别的缩进。

**语法**

**express.Indent()**

*express*   一个代表 **Paragraphs** 对象的变量。

**说明**

此方法等效于单击**“格式”**工具栏上的**“增加缩进量”**按钮。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例为活动文档中的所有段落增加两个级别的缩进，然后为第一段删除一个级别的缩进。*/ function test() {   Application.ActiveDocument.Paragraphs.Indent()   Application.ActiveDocument.Paragraphs.Item(1).Outdent() }` |

#### **Paragraphs.IndentFirstLineCharWidth**

将一个或多个段落的首行缩进指定的字符数。

**语法**

**express.IndentFirstLineCharWidth(Count)**

*express*   一个代表 **Paragraphs** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                           |
| -------- | ------------- | ------------ | ---------------------------------- |
| *Count*  | 必选          | **Integer**  | 每个指定段落的首行要缩进的字符数。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动文档中所有段落的首行缩进 10 个字符。*/ Application.ActiveDocument.Paragraphs.IndentFirstLineCharWidth(10)` |

**成员属性**

#### **Paragraphs.AddSpaceBetweenFarEastAndAlpha**

如果 WPS 将自动在指定段落的日文和拉丁文文字之间添加空格，则该属性值为 **True**。如果仅对于某些指定段落将该属性设置为 **True**，则该属性会返回 **wdUndefined**。**Long** 类型，可读写。

**语法**

**express.AddSpaceBetweenFarEastAndAlpha**

*express*   一个代表 **Paragraphs** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例实现的功能是：设置 WPS 在活动文档第一段的日文文字和拉丁文字之间自动添加空格。*/ ActiveDocument.Paragraphs.Item(1).AddSpaceBetweenFarEastAndAlpha = true` |

#### **Paragraphs.AddSpaceBetweenFarEastAndDigit**

如果 WPS 将自动在指定段落的日文文字和数字之间添加空格，则该属性值为 **True**。如果仅对于某些指定段落将该属性设置为 **True**，则该属性会返回 **wdUndefined**。**Long** 类型，可读写。

**语法**

**express.AddSpaceBetweenFarEastAndDigit**

*express*   一个代表 **Paragraphs** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例可实现的功能是：设定 WPS 自动在活动文档第一段的中文文字和数字之间添加空格。*/ ActiveDocument.Paragraphs.Item(1).AddSpaceBetweenFarEastAndDigit = true ` |

#### **Paragraphs.Alignment**

返回或设置一个 **WdParagraphAlignment** 常量，该常量代表指定段落的对齐方式，可读写。

**语法**

**express.Alignment**

*express*   一个代表 **Paragraphs** 对象的变量。

**说明**

对于所选择或安装的不同语言支持（例如美国英语），某些 **WdParagraphAlignment** 常量可能不可用。

#### **Paragraphs.Application**

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

*express*   一个代表 **Paragraphs** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **Paragraphs.AutoAdjustRightIndent**

如果 WPS 会根据您指定的每行字符数自动调整指定段落的右缩进，则该属性值为 **True**。如果只将某些指定段落的 **AutoAdjustRightIndent** 属性设置为 **True**，则该属性会返回 **wdUndefined**。**Long** 类型，可读写。

**语法**

**express.AutoAdjustRightIndent**

*express*   一个代表 **Paragraphs** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 根据指定的每行字符数，自动调整所选段落的右缩进。*/ Selection.ParagraphFormat.AutoAdjustRightIndent = true` |

#### **Paragraphs.BaseLineAlignment**

返回或设置一个 **WdBaselineAlignment** 常量，该常量代表行中字体的垂直位置，可读写。

**语法**

**express.BaseLineAlignment**

*express*   一个代表 **Paragraphs** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 自动调整活动文档中的基线字体对齐方式。*/ ActiveDocument.BaseLineAlignment = wdBaselineAlignAuto` |

#### **Paragraphs.Borders**

返回一个 **Borders** 集合，该集合代表指定对象的所有边框。

**语法**

**express.Borders**

*express*   一个代表 **Paragraphs** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

#### **Paragraphs.CharacterUnitFirstLineIndent**

返回或设置首行或悬挂缩进的值（以字符为单位）。用正值设置首行缩进，用负值设置悬挂缩进。**Single** 类型，可读写。

**语法**

**express.CharacterUnitFirstLineIndent**

*express*   一个代表 **Paragraphs** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档中第一段的首行缩进设为一个字符。*/ ActiveDocument.Paragraphs.Item(1).CharacterUnitFirstLineIndent = 1 ` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档中第二段的悬挂缩进设为 1.5 个字符。*/ ActiveDocument.Paragraphs.Item(2).CharacterUnitFirstLineIndent = -1.5` |

#### **Paragraphs.CharacterUnitLeftIndent**

该属性返回或设置指定段落的左缩进量（以字符为单位）。**Single** 类型，可读写。

**语法**

**express.CharacterUnitLeftIndent**

*express*   一个代表 **Paragraphs** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档中第一段的左缩进设为从左边距缩进一个字符。*/ ActiveDocument.Paragraphs.Item(1).CharacterUnitLeftIndent = 1  ` |

#### **Paragraphs.CharacterUnitRightIndent**

该属性返回或设置指定段落的右缩进量（以字符为单位）。**Single** 类型，可读写。

**语法**

**express.CharacterUnitRightIndent**

*express*   一个代表 **Paragraphs** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档中的所有段落的右缩进设为从右边距缩进一个字符。*/ ActiveDocument.Paragraphs.CharacterUnitRightIndent = 1` |

#### **Paragraphs.Count**

返回一个 **Long** 类型的值，该值代表集合中的段数。只读。

**语法**

**express.Count**

*express*   一个代表 **Paragraphs** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示活动文档的段数。*/ alert("The active document contains " + ActiveDocument.Paragraphs.Count + " paragraphs.") ` |

#### **Paragraphs.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Paragraphs** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **Paragraphs.DisableLineHeightGrid**

如果该属性的值为 **True**，则当指定每页的行数时，WPS 会将指定段落中的字符与行网格对齐。如果只将某些指定段落的 **DisableLineHeightGrid** 属性设置为 **True**，则返回 **wdUndefined**。**Long** 类型，可读写。

**语法**

**express.DisableLineHeightGrid**

*express*   一个代表 **Paragraphs** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 在已指定每页行数的情况下，将选定段落中的字符与行网格对齐。*/ Selection.ParagraphFormat.DisableLineHeightGrid = true` |

#### **Paragraphs.FarEastLineBreakControl**

如果为 **True**，则 WPS 会将东亚语言文字的换行规则应用于指定的段落。如果只将某些指定段落的 **FarEastLineBreakControl** 属性设定为 **True**，则返回 **wdUndefined**。**Long** 类型，可读写。

**语法**

**express.FarEastLineBreakControl**

*express*   一个代表 **Paragraphs** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设定 WPS 将东亚语言文字的换行规则应用于活动文档的第一段。*/ ActiveDocument.Paragraphs.Item(1).FarEastLineBreakControl = true` |

#### **Paragraphs.First**

返回一个 **Paragraph** 对象，该对象代表在 **Paragraphs** 集合中的第一个项目。

**语法**

**express.First**

*express*   一个代表 **Paragraphs** 对象的变量。

#### **Paragraphs.FirstLineIndent**

返回或设置首行缩进或悬挂缩进的大小（以磅值表示）。用正数设置首行缩进的尺寸，用负数设置悬挂缩进的尺寸。**Single** 类型，可读写。

**语法**

**express.FirstLineIndent**

*express*   一个代表 **Paragraphs** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为活动文档的首段设置 1 英寸的首行缩进。*/ Application.ActiveDocument.Paragraphs.FirstLineIndent = InchesToPoints(1)  /*本示例为活动文档的第二段设置 0.5 英寸的悬挂缩进。InchesToPoints 方法用来将英寸转化为磅值。*/ Application.ActiveDocument.Paragraphs.FirstLineIndent = InchesToPoints(-0.5) ` |

#### **Paragraphs.Format**

返回或设置一个 **ParagraphFormat** 对象，该对象代表指定的一个或多个段落的格式。

**语法**

**express.Format**

*express*   一个代表 **Paragraphs** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动文档的所有段落靠左对齐。*/ ActiveDocument.Paragraphs.Format.Alignment = wdAlignParagraphLeft` |

#### **Paragraphs.HalfWidthPunctuationOnTopOfLine**

如果为 **True**，则 WPS 会将指定段落行首的标点符号改为半角字符。如果仅将某些指定段落的该属性设置为 **True**，则此属性将返回 **wdUndefined**。**Long** 类型，可读写。

**语法**

**express.HalfWidthPunctuationOnTopOfLine**

*express*   一个代表 **Paragraphs** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 将活动文档第一段的行首标点符号改为半角字符。*/ ActiveDocument.Paragraphs.HalfWidthPunctuationOnTopOfLine = true` |

#### **Paragraphs.HangingPunctuation**

如果为 **True**，则指定段落中的标点将可以溢出边界。如果仅将某些指定段落的该属性设置为 **True**，则返回 **wdUndefined**。**Long** 类型，可读写。

**语法**

**express.HangingPunctuation**

*express*   一个代表 **Paragraphs** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例实现的功能是：允许活动文档第一段的标点可以溢出边界。*/ ActiveDocument.Paragraphs.HangingPunctuation = true` |

#### **Paragraphs.Hyphenation**

如果指定的段落进行自动断字，则该属性值为 **True**。如果指定的段落不进行自动断字，则该属性值为 **False**。可读写 **Long** 类型。

**语法**

**express.Hyphenation**

*express*   一个代表 **Paragraphs** 对象的变量。

**说明**

此属性可以是 **True**、**False** 或 **wdUndefined**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例关闭活动文档中所有段落的自动断字功能。*/ ActiveDocument.Paragraphs.Hyphenation = false` |

#### **Paragraphs.KeepTogether**

在 WPS 对文档重新分页时，如果指定段落中的所有行都位于同一页上，则该属性值为 **True**。可读写 **Long** 类型。

**语法**

**express.KeepTogether**

*express*   一个代表 **Paragraphs** 对象的变量。

**说明**

此属性可以是 **True**、**False** 或 **wdUndefined**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例为活动文档中的各段落设置格式，以使在 WPS 对文档重新分页时同一段中的各行都位于同一页上。*/ ActiveDocument.Paragraphs.KeepTogether = true` |

#### **Paragraphs.KeepWithNext**

在 WPS 对文档重新分页时，如果指定段落与其下一段位于同一页上，则该属性值为 **True**。可读写 **Long** 类型。

**语法**

**express.KeepWithNext**

*express*   一个代表 **Paragraphs** 对象的变量。

**说明**

此属性可以是 **True**、**False** 或 **wdUndefined**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将当前所选内容中的所有段落设置为位于同一页上。*/ Selection.Paragraphs.KeepWithNext = true` |

#### **Paragraphs.Last**

返回一个 **Paragraph** 对象，该对象代表段落集合中的最后一个项目。

**语法**

**express.Last**

*express*   一个代表 **Paragraphs** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动文档中的最后一个段落设置为右对齐。*/ ActiveDocument.Paragraphs.Last.Alignment = wdAlignParagraphRight` |

#### **Paragraphs.LeftIndent**

返回或设置一个 **Single** 类型的值，该值代表指定段落的左缩进值（以磅为单位）。可读写。

**语法**

**express.LeftIndent**

*express*   一个代表 **Paragraphs** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动文档中所有段落的左缩进都设置为 1 英寸。InchesToPoints 方法用来将英寸转换为磅。*/ ActiveDocument.Paragraphs.LeftIndent = InchesToPoints(1)` |

#### **Paragraphs.LineSpacing**

返回或设置指定段落的行距（以磅为单位）。**Single** 类型，可读写。

**语法**

**express.LineSpacing**

*express*   一个代表 **Paragraphs** 对象的变量。

**说明**

使用 **LinesToPoints** 方法可将行数转换为相应的值（以磅为单位）。例如，`LinesToPoints(2)` 返回值 24。

可以在 **LineSpacingRule** 属性已设置为下列值之后对 **LineSpacing** 属性进行设置：

- **wdLineSpaceAtLeast** 行距可以大于或等于指定的 **LineSpacing** 值，但绝不能小于该值。
- **wdLineSpaceExactly** 指定的 **LineSpacing** 行距值绝不能更改，即使段落中使用了较大的字体。
- **wdLineSpaceMultiple** 必须指定 **LineSpacing** 属性值，以磅为单位。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例为所选段落设置 3 倍行距。*/ function test() { Selection.Paragraphs.LineSpacingRule = wdLineSpaceMultiple Selection.Paragraphs.LineSpacing = LinesToPoints(3) }` |

#### **Paragraphs.LineSpacingRule**

返回或设置指定段落的行距。可读写 **WdLineSpacing** 类型。

**语法**

**express.LineSpacingRule**

*express*   一个代表 **Paragraphs** 对象的变量。

**说明**

使用 **wdLineSpaceSingle**、**wdLineSpace1pt5** 或 **wdLineSpaceDouble** 可将行距设置为这些值之一。要将行距设置为固定磅值或多倍行距，还必须设置 **LineSpacing** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例为活动文档的所有段落设置 2 倍行距。*/ ActiveDocument.Paragraphs.LineSpacingRule = wdLineSpaceDouble` |

#### **Paragraphs.LineUnitAfter**

返回或设置指定段落的段后间距（以网格线为单位）。可读写 **Single** 类型。

**语法**

**express.LineUnitAfter**

*express*   一个代表 **Paragraphs** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动文档中所有段落的段后间距设置为 1 个网格线。*/ ActiveDocument.Paragraphs.LineUnitAfter = 1` |

#### **Paragraphs.LineUnitBefore**

返回或设置指定段落的段前间距（以网格线为单位）。可读写 **Single** 类型。

**语法**

**express.LineUnitBefore**

*express*   一个代表 **Paragraphs** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动文档中所有段落的段前间距设置为 1 个网格线。*/ ActiveDocument.Paragraphs.LineUnitBefore = 1` |

#### **Paragraphs.NoLineNumber**

如果取消指定段的行号，则该属性值为 **True**。该属性值可以是 **True**、**False** 或 **wdUndefined**。可读写 **Long** 类型。

**语法**

**express.NoLineNumber**

*express*   一个代表 **Paragraphs** 对象的变量。

**说明**

使用 **PageSetup** 对象的 **LineNumbering** 属性可设置行号。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例为活动文档设置行号。起始编号设置为 1，并且在文档中的所有节都连续编号。然后取消第二段的行号。*/ function test() { let linenum = ActiveDocument.PageSetup.LineNumbering     linenum.Active = true     linenum.StartingNumber = 1     linenum.CountBy = 1     linenum.RestartMode = wdRestartContinuous ActiveDocument.Paragraphs.Item(2).NoLineNumber = true }` |

#### **Paragraphs.OutlineLevel**

返回或设置指定段落的大纲级别。可读写 **WdOutlineLevel** 类型。

**语法**

**express.OutlineLevel**

*express*   一个代表 **Paragraphs** 对象的变量。

**说明**

如果段落应用了一种标题样式（从“标题 1”到“标题 9”），则大纲级别与该标题样式相同，并且不能更改。大纲级别仅在大纲视图或文档结构图窗格中可见。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例返回活动文档中所有段落的大纲级别。*/ let temp = ActiveDocument.Paragraphs.OutlineLevel` |

#### **Paragraphs.PageBreakBefore**

如果在指定段落前插入了分页符，则该属性值为 **True**。该属性值可以是 **True**、**False** 或 **wdUndefined**。可读写 **Long** 类型。

**语法**

**express.PageBreakBefore**

*express*   一个代表 **Paragraphs** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在所选内容的第一段前插入分页符。*/ Selection.Paragraphs.PageBreakBefore = true` |

#### **Paragraphs.Parent**

返回一个 **Object** 类型值，该值代表指定 **Paragraphs** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **Paragraphs** 对象的变量。

#### **Paragraphs.ReadingOrder**

返回或设置指定段落的读取次序而不改变其对齐方式。可读写 **WdReadingOrder** 类型。

**语法**

**express.ReadingOrder**

*express*   一个代表 **Paragraphs** 对象的变量。

**说明**

使用 **Selection** 对象的 **LtrPara**、**LtrRun**、**RtlPara** 和 **RtlRun** 方法可更改段落的对齐方式和读取次序。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动文档中所有段落的读取次序设置为从右向左。*/ ActiveDocument.Paragraphs.ReadingOrder = wdReadingOrderRtl` |

#### **Paragraphs.RightIndent**

返回或设置指定段落的右缩进量（以磅为单位）。可读写 **Single** 类型。

**语法**

**express.RightIndent**

*express*   一个代表 **Paragraphs** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将活动文档中所有段落的右缩进都设置为距右边距 1 英寸。InchesToPoints 方法用来将英寸转换为磅。*/ ActiveDocument.Paragraphs.RightIndent = InchesToPoints(1)` |

#### **Paragraphs.Shading**

返回一个 **Shading** 对象，该对象代表指定段落的底纹格式。

**语法**

**express.Shading**

*express*   一个代表 **Paragraphs** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例对所选内容的所有段落应用黄色底纹。*/ function test() { let shadingParagraph = Selection.Paragraphs.Shading     shadingParagraph.Texture = wdTexture12Pt5Percent     shadingParagraph.BackgroundPatternColorIndex = wdYellow     shadingParagraph.ForegroundPatternColorIndex = wdBlack }` |

#### **Paragraphs.SpaceAfter**

返回或设置指定段落或文本栏后面的间距（以磅为单位）。可读/写 **Single** 类型。

**语法**

**express.SpaceAfter**

*express*   一个代表 **Paragraphs** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档中所有段落的段后间距设置为 12 磅。*/ ActiveDocument.Paragraphs.SpaceAfter = 12` |

#### **Paragraphs.SpaceAfterAuto**

如果 WPS 自动设置指定段落的段后间距，则该属性为 **True**。可读/写 **Long** 类型。

**语法**

**express.SpaceAfterAuto**

*express*   一个代表 **Paragraphs** 对象的变量。

**说明**

如果仅将部分指定段落的 **SpaceAfterAuto** 属性设置为 **True**，则该属性返回 **wdUndefined**。该属性可被设置为 **True** 或 **False**。

如果将 **SpaceAfterAuto** 设置为 **True**，则 **SpaceAfter** 属性将被忽略。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示关于活动文档的 SpaceAfterAuto 属性设置的报告。*/ function test() { let mySpaceAfterAuto = ActiveDocument.Paragraphs.SpaceAfterAuto switch(mySpaceAfterAuto){     case 0:         let x = "Spacing after paragraphs is handled " + "automatically for all paragraphs."         break     case 1:         let x = "Spacing after paragraphs is handled " + "manually for all paragraphs."         break     case undefined:         let x = "Spacing after paragraphs is handled " + "automatically for some paragraphs, " + "manually for some paragraphs." } }` |

#### **Paragraphs.SpaceBefore**

返回或设置指定段落的段前间距（以磅为单位）。可读/写 **Single** 类型。

**语法**

**express.SpaceBefore**

*express*   一个代表 **Paragraphs** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档中所有段落的段前间距设置为 12 磅。*/ ActiveDocument.Paragraphs.SpaceBefore = 12` |

#### **Paragraphs.SpaceBeforeAuto**

如果 WPS 自动设置指定段落的段前间距，则该属性为 **True**。可读/写 **Long** 类型。

**语法**

**express.SpaceBeforeAuto**

*express*   一个代表 **Paragraphs** 对象的变量。

**说明**

如果仅将部分指定段落的 **SpaceBeforeAuto** 属性设置为 **True**，则该属性返回 **wdUndefined**。该属性可被设置为 **True** 或 **False**。

如果将 **SpaceBeforeAuto** 设置为 **True**，则 **SpaceBefore** 属性将被忽略。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示关于活动文档的 SpaceBeforeAuto 属性设置的报告。*/ function test() { let mySpaceBeforeAuto = ActiveDocument.Paragraphs.SpaceBeforeAuto switch(mySpaceBeforeAuto){     case 0:         let x = "Spacing before paragraphs is handled " + "automatically for all paragraphs."         break     case 1:         let x = "Spacing before paragraphs is handled " + "manually for all paragraphs."         break     case undefined:         let x = "Spacing before paragraphs is handled " + "automatically for some paragraphs, " + "manually for some paragraphs." } }` |

#### **Paragraphs.Style**

返回或设置指定段落的样式。可读写 **Variant** 类型。

**语法**

**express.Style**

*express*   一个代表 **Paragraphs** 对象的变量。

**说明**

要设置该属性，请指定样式的本地名称、一个整数、一个 **WdBuiltinStyle** 常量或一个代表样式的对象。

如果返回包含多个样式的范围的样式，则只返回第一个字符样式或段落样式。

#### **Paragraphs.TabStops**

返回或设置一个 **TabStops** 集合，该集合代表指定段落中的所有自定义制表位。可读写。

**语法**

**express.TabStops**

*express*   一个代表 **Paragraphs** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在活动文档各段的 2 英寸处添加居中的制表位。InchesToPoints 方法用来将英寸转换为磅。*/ ActiveDocument.Paragraphs.TabStops.Add(InchesToPoints(2),wdAlignTabCenter,null)` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将文档中各段的制表位设置为与第一段的制表位一致。*/ function test() { let para1Tabs = ActiveDocument.Paragraphs.Item(1).TabStops ActiveDocument.Paragraphs.TabStops = para1Tabs }` |

#### **Paragraphs.WidowControl**

在 WPS 对文档重新分页时，如果指定段落的首行和末行与段落的其他各行同页，则该属性值为 **True**。该属性值可以是 **True**、**False** 或 **wdUndefined**。可读写 **Long** 类型。

**语法**

**express.WidowControl**

*express*   一个代表 **Paragraphs** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例给活动文档中的各段落设置格式，从而使段中的首行或末行有可能单独位于上页的页尾或下页的页首。*/ ActiveDocument.Paragraphs.WidowControl = false` |

#### **Paragraphs.WordWrap**

如果 WPS 在指定段落中的西文单词中间断字换行，则该属性值为 **True**。可读写 **Long** 类型。

**语法**

**express.WordWrap**

*express*   一个代表 **Paragraphs** 对象的变量。

**说明**

如果仅将某些指定段落的此属性设置为 **True**，则此属性返回 **wdUndefined**。根据您所选择或安装的语言支持（例如，美国英语），该属性可能不可用。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例设置 WPS 在活动文档的第一段中的西文单词中间断字换行。*/ ActiveDocument.Paragraphs.Item(1).WordWrap = true` |

适用环境：web

适用平台：windows/linux