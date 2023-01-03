#### **Font**



包含对象的字体属性（如字体名称、字号、颜色等）。

**说明**

使用 **Font** 属性可返回 **Font** 对象。以下指令对所选内容应用加粗格式。

| 示例代码复制                             |
| ---------------------------------------- |
| `Application.Selection.Font.Bold = true` |

以下示例将活动文档中第一个段落的格式设置为 24 磅、Arial 和加粗。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let myRange = ActiveDocument.Paragraphs.Item(1).Range     myRange.Font.Bold = true     myRange.Font.Name = "Arial"     myRange.Font.Size = 24 }` |

以下示例将活动文档中“标题 2”样式的格式更改为 Arial 字体和斜体。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     Application.ActiveDocument.Styles.Item(wdStyleHeading2).Font.Name = "Arial"     Application.ActiveDocument.Styles.Item(wdStyleHeading2).Font.Italic = true }` |

您可以使用 **n**

ew

 关键字来创建一个独立的新 

Font

 对象。以下示例创建一个 

Font

 对象，设置一些格式属性，然后将该 

Font

 对象应用于活动文档的第一个段落。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let myFont = Application.ActiveDocument.Paragraphs(1).Range.Font     myFont.Bold = true     myFont.Name = "Arial" } ` |

**方法**

|                                                              | 名称                     | 说明                                                         |
| ------------------------------------------------------------ | ------------------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Grow**                 | 将字号增大为下一个可用的字号。                               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Reset**                | 删除手动字符格式（不使用样式应用的格式）。例如，如果手动将单词设置为加粗格式，但基本样式为纯文本（不是粗体），则可用 **Reset** 方法删除加粗格式。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SetAsTemplateDefault** | 将指定的字体格式设置为活动文档和所有基于活动模板创建的新文档的默认格式。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Shrink**               | 将字号减小为下一个可用的字号。                               |

**属性**

|                                                              | 名称                          | 说明                                                         |
| ------------------------------------------------------------ | ----------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AllCaps**                   | 如果字体格式为全部字母大写，则该属性值为 **True**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Animation**                 | 返回或设置应用于字体的动态效果类型。**WdAnimation** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**               | 返回一个代表 WPS 应用程序的 **Application** 对象。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Bold**                      | 如果字体格式为加粗，则该属性值为 **True**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BoldBi**                    | 如果字体格式为加粗，则该属性值为 **True**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Borders**                   | 返回一个 **Borders** 集合，该集合代表指定字体的所有边框。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ColorIndex**                | 返回或设置一个 **WdColorIndex** 常量，该常量代表指定字体的颜色。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ColorIndexBi**              | 返回或设置从右向左语言文档中指定 **Font** 对象的颜色。**WdColorIndex** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ContextualAlternates**      | 指定是否为指定字体启用上下文替换。可读/写 **Long** 类型。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**                   | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DiacriticColor**            | 返回或设置用于指定 **Font** 对象的音调符号的 24 位颜色。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisableCharacterSpaceGrid** | 如果 WPS 忽略相应 **Font** 对象的每行字符数，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DoubleStrikeThrough**       | 如果将指定字体的格式设置为双删除线文本，则该属性值为 **True**。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Duplicate**                 | 返回一个只读 **Font** 对象，该对象代表指定字体的字符格式。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Emboss**                    | 如果指定字体格式为阳文，则该属性值为 **True**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EmphasisMark**              | 返回或设置 **WdEmphasisMark** 常量，该常量代表字符或指定字符串的着重号。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Engrave**                   | 如果字体格式为阴文，则该属性值为 **True**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Fill**                      | 返回 FillFormat 对象，该对象包含指定范围文本使用的字体的填充格式属性。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Glow**                      | 返回 GlowFormat 对象，该对象代表指定范围文本使用的字体的发光格式。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Hidden**                    | 如果字体格式为隐藏文字，则该属性值为 **True**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Italic**                    | 如果将字体或范围设置为倾斜格式，则该属性值为 **True**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ItalicBi**                  | 如果将字体或范围设置为倾斜格式，则该属性值为 **True**。可读写 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Kerning**                   | 返回或设置 WPS 根据字号自动调整字距时的最小字号。**Single** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Ligatures**                 | 返回或设置指定 **Font** 对象的连字设置。可读/写 WdLigatures。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Line**                      | 返回 LineFormat 对象，该对象指定行的格式。可读/写。          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**                      | 返回或设置指定对象的名称。**String** 类型，可读写。          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NameAscii**                 | 返回或设置拉丁语（字符代码从 0（零）到 127 的字符）所用的字体。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NameBi**                    | 返回或设置使用从右向左语言的文档中的字体的名称。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NameFarEast**               | 返回或设置一种东亚字体名称。**String** 类型，可读写。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NameOther**                 | 返回或设置字符代码从 128 到 255 的字符所用的字体。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NumberForm**                | 返回或设置 OpenType 字体的数字形式设置。可读/写 WdNumberForm。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NumberSpacing**             | 返回或设置字体的数字间距设置。可读/写 WdNumberSpacing。      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Outline**                   | 如果字体格式为镂空，则该属性值为 **True**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**                    | 返回一个 **Object** 类型值，该值代表指定 **Font** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Position**                  | 返回或设置相对于基准线的文本位置（以磅为单位）。可读写 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Reflection**                | 返回一个 ReflectionFormat 对象，该对象代表形状的反射格式。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Scaling**                   | 返回或设置应用于字体的缩放百分比。**Long** 类型，可读写。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Shading**                   | 返回一个 **Shading** 对象，该对象代表指定字体的底纹格式。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Shadow**                    | 如果将指定字体设置为阴影格式，则该属性值为 **True**。可读写 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Size**                      | 返回或设置字号（以磅为单位）。**Single** 类型，可读写。      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SizeBi**                    | 返回或设置字体大小（以磅为单位）。**Single** 类型，可读写。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SmallCaps**                 | 如果字体格式为小型大写字母，则该属性值为 **True**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Spacing**                   | 返回或设置字符的间距（以磅为单位）。可读写 **Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **StrikeThrough**             | 如果字体格式为带有删除线，则该属性值为 **True**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **StylisticSet**              | 指明指定字体的样式集。可读/写 WdStylisticSet。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Subscript**                 | 如果字体格式为下标，则该属性值为 **True**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Superscript**               | 如果字体格式为上标，则该属性值为 **True**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextColor**                 | 返回一个代表指定字体的颜色的 [ColorFormat](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/ColorFormat/ColorFormat%20.htm#jsObject_ColorFormat)[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/ColorFormat/ColorFormat%20.htm#jsObject_ColorFormat)对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextShadow**                | 返回 ShadowFormat 对象，该对象指明指定字体的阴影格式。       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ThreeD**                    | 返回 ThreeDFormat 对象，该对象包含指定字体的三维效果格式属性。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Underline**                 | 返回或设置应用于字体的下划线类型。可读写 **WdUnderline**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdUnderline%20%E6%9E%9A%E4%B8%BE.html)类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **UnderlineColor**            | 返回或设置指定 **Font** 对象的下划线的 24 位颜色。           |

**成员方法**

#### **Font.Grow**

将字号增大为下一个可用的字号。

**语法**

**express.Grow()**

*express*   一个代表 **Font** 对象的变量。

**说明**

如果所选内容或范围包含多种字号，则将每一个字号都增大为下一个可用的设置。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例增大新文档中第四个单词的字号。*/ function test() {   let rngTemp = Application.Documents.Add().Content   rngTemp.InsertAfter("This is a test of the Grow method.")   alert("Click OK to increase the font size of the fourth word.")   rngTemp.Words.Item(4).Font.Grow() }  /*以下示例增大选定文本的字号。*/ function test() {   if(Application.Selection.Type == wdSelectionNormal){       Application.Selection.Font.Grow()   }   else{       alert("You need to select some text.")   } }` |

#### **Font.Reset**

删除手动字符格式（不使用样式应用的格式）。例如，如果手动将单词设置为加粗格式，但基本样式为纯文本（不是粗体），则可用 **Reset** 方法删除加粗格式。

**语法**

**express.Reset()**

*express*   一个代表 **Font** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例删除选定部分中的手动设定的格式。*/ Application.Selection.Font.Reset()` |

#### **Font.SetAsTemplateDefault**

将指定的字体格式设置为活动文档和所有基于活动模板创建的新文档的默认格式。

**语法**

**express.SetAsTemplateDefault()**

*express*   一个代表 **Font** 对象的变量。

**说明**

默认字体格式存储在“正文”样式中。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将所选内容中的字符格式设置为默认格式。*/ Selection.Font.SetAsTemplateDefault()` |

#### **Font.Shrink**

将字号减小为下一个可用的字号。

**语法**

**express.Shrink()**

*express*   一个代表 **Font** 对象的变量。

**说明**

如果所选内容或范围中包括多种字号，则将每一个字号都减小为下一个可用的设置。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在新文档中插入一行字号逐渐减小的字母 Z。*/ function test() { let myRange = Documents.Add().Content myRange.Font.Size = 45  for(let i = 1; i <= 5; i++){     myRange.InsertAfter("Z")      for(let r = 1; r <= 3; r++){         myRange.Characters.Item(i).Font.Shrink()     } } }` |

**成员属性**

#### **Font.AllCaps**

如果字体格式为全部字母大写，则该属性值为 **True**。**Long** 类型，可读写。

**语法**

**express.AllCaps**

*express*   一个代表 **Font** 对象的变量。

**说明**

返回 **True**、**False** 或 wdUndefined（当返回值既可为 **True**，也可为 **False** 时取该值）。该属性可设置为 **True**、**False** 或 **wdToggle**（与当前设置相反）。

如果设置 **AllCaps** 为 **True**，则 **SmallCaps** 属性的值为 **False**，反之亦然。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例检查活动文档第三段的文本格式是否设置为全部字母大写。*/ function test() { if(ActiveDocument.Paragraphs.Item(3).Range.Font.AllCaps == true){      MsgBox("Text is all caps.") } else{     MsgBox("Text is not all caps.") } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将选定文本的字体格式设置为全部字母大写。*/ function test() { if(Selection.Type == wdSelectionNormal){     Selection.Font.AllCaps = true } else{     MsgBox("You need to select some text.") } }  ` |

#### **Font.Animation**

返回或设置应用于字体的动态效果类型。**WdAnimation** 类型，可读写。

**语法**

**express.Animation**

*express*   一个代表 **Font** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为新文档中的文本设置动态效果。*/ function test() { let docNew = Documents.Add() docNew.Content.InsertAfter("This is a test of animation.") docNew.Content.Font.Animation = wdAnimationLasVegasLights }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为选定文本设置动态效果。*/ function test() { if(Selection.Type == wdSelectionNormal) {     Selection.Font.Animation = wdAnimationShimmer } else{     MsgBox("You need to select some text.") } }` |

#### **Font.Application**

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

*express*   一个代表 **Font** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **Font.Bold**

如果字体格式为加粗，则该属性值为 **True**。**Long** 类型，可读写。

**语法**

**express.Bold**

*express*   一个代表 **Font** 对象的变量。

**说明**

返回 **True**、**False** 或 **wdUndefined**（当返回值既可为 **True**，也可为 **False** 时取该值）。可设置为 **True**、**False** 或 **wdToggle**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例实现的功能是：如果选定内容中有一部分字体为加粗格式，则将全部选定内容的格式都设为加粗格式。*/ function test() {   if(Application.Selection.Type == wdSelectionNormal) {       if(Application.Selection.Font.Bold == wdUndefined) {           Application.Selection.Font.Bold = true       }   }   else{       alert("You need to select some text.")   } }` |

#### **Font.BoldBi**

如果字体格式为加粗，则该属性值为 **True**。**Long** 类型，可读写。

**语法**

**express.BoldBi**

*express*   一个代表 **Font** 对象的变量。

**说明**

该属性返回 **True**、**False** 或 **wdUndefined**（用于加粗和非加粗混合的文本）。该属性可设为 **True**、**False** 或 **wdToggle**。

**BoldBi** 属性应用于从右向左语言的文字。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将从右向左语言的活动文档中第一段的格式设为加粗。*/ ActiveDocument.Paragraphs.Item(1).Range.Font.BoldBi = true` |

#### **Font.Borders**

返回一个 **Borders** 集合，该集合代表指定字体的所有边框。

**语法**

**express.Borders**

*express*   一个代表 **Font** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

#### **Font.ColorIndex**

返回或设置一个 **WdColorIndex** 常量，该常量代表指定字体的颜色。可读写。

**语法**

**express.ColorIndex**

*express*   一个代表 **Font** 对象的变量。

**说明**

**wdByAuthor** 常量不是有效的字体颜色。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例更改活动文档首段的文字颜色。*/ Application.ActiveDocument.Paragraphs.Item(1).Range.Font.ColorIndex = wdGreen  /*本示例将所选文本的格式设置为红色。*/ Application.Selection.Font.ColorIndex = wdRed` |

#### **Font.ColorIndexBi**

返回或设置从右向左语言文档中指定 **Font** 对象的颜色。**WdColorIndex** 类型，可读写。

**语法**

**express.ColorIndexBi**

*express*   一个代表 **Font** 对象的变量。

**说明**

**wdByAuthor** 常量对 **Font** 对象是无效的。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 Font 对象的颜色设置为蓝绿色。*/ Selection.Font.ColorIndexBi = wdTeal` |

#### **Font.ContextualAlternates**

指定是否为指定字体启用上下文替换。可读/写 **Long** 类型。

**语法**

**express.ContextualAlternates**

*express*   一个代表 **Font** 对象的变量。

**说明**

上下文替换是根据单个字符周围的字母（其上下文）应用于这些字符的连字。上下文替换还可以应用于特定上下文中的所有词，例如常常用于标题中的词（如“of”和“the”）。为字体启用上下文替换后，将在字体设计器定义的那些上下文中使用上下文替换，而不使用标准连字。

设置该属性与选中**“使用上下文替换”**（位于 WPS 2015 中**“字体”**对话框中**“高级”**选项卡上的**“OpenType 功能”**组中）旁边的复选框具有相同作用。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码示例为活动文档中的字体启用上下文替换。*/ ActiveDocument.Range.Font.ContextualAlternates = true` |

#### **Font.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Font** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **Font.DiacriticColor**

返回或设置用于指定 **Font** 对象的音调符号的 24 位颜色。

**语法**

**express.DiacriticColor**

*express*   一个代表 **Font** 对象的变量。

**说明**

此属性可以是任何有效的 **WdColor** 常量或 Visual Basic 的 **RGB** 函数所返回的值。必须将 **UseDiffDiacColor** 属性的值设为 **True** 才能使用此属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将当前选定内容中的音调符号设为蓝色。*/ function test() { if(Options.UseDiffDiacColor == true) {     Selection.Font.DiacriticColor = wdColorBlue } }` |

#### **Font.DisableCharacterSpaceGrid**

如果 WPS 忽略相应 **Font** 对象的每行字符数，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.DisableCharacterSpaceGrid**

*express*   一个代表 **Font** 对象的变量。

**说明**

如果只将某些指定文字的 **DisableCharacterSpaceGrid** 属性设为 **True**，则该属性会返回 **wdUndefined**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*该示例设置 WPS 忽略选定文本每行中的字符数。*/ Selection.Font.DisableCharacterSpaceGrid = true` |

#### **Font.DoubleStrikeThrough**

如果将指定字体的格式设置为双删除线文本，则该属性值为 **True**。

**语法**

**express.DoubleStrikeThrough**

*express*   一个代表 **Font** 对象的变量。

**说明**

返回 **True**、**False** 或 **wdUndefined**（既可为 **True**，也可为 **False**）。可设置为 **True**、**False** 或 **wdToggle**。**Long** 类型，可读写。要设置或返回单删除线格式，可使用 **StrikeThrough** 属性。如果将 **DoubleStrikeThrough** 设置为 **True**，则 **StrikeThrough** 将设置为 **False**，反之亦然。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例对所选文字应用双删除线格式。*/ function test() { if(Selection.Type == wdSelectionNormal) {     Selection.Font.DoubleStrikeThrough = true } else{     MsgBox("You need to select some text.") } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例取消活动文档中第一个单词的双删除线格式，并将该单词的第一个字母大写。*/ function test() { ActiveDocument.Words.Item(1).Font.DoubleStrikeThrough = false ActiveDocument.Words.Item(1).Case = wdTitleSentence }` |

#### **Font.Duplicate**

返回一个只读 **Font** 对象，该对象代表指定字体的字符格式。

**语法**

**express.Duplicate**

*express*   一个代表 **Font** 对象的变量。

**说明**

可以使用 **Duplicate** 属性获取重复的 **Font** 对象所有属性的设置。可将 **Duplicate** 属性返回的对象赋给另一个 **Font** 对象，从而一次应用该对象的所有设置。在将重复对象赋给另一个对象之前，可更改重复对象的任何属性，而不会影响源对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 MyDupFont 变量设为选定内容的字符格式，从 MyDupFont 中删除加粗格式，并添加倾斜格式。然后，本示例会创建一个新文档，插入文本，再将 MyDupFont 中保存的格式应用于该文本。*/ function test() { let myDupFont = Selection.Font.Duplicate myDupFont.Bold = false myDupFont.Italic = true  Documents.Add() Selection.InsertAfter("This is some text.") Selection.Font = myDupFont }` |

#### **Font.Emboss**

如果指定字体格式为阳文，则该属性值为 **True**。**Long** 类型，可读写。

**语法**

**express.Emboss**

*express*   一个代表 **Font** 对象的变量。

**说明**

返回 **True**、**False** 或 **wdUndefined**。该属性可设为 **True**、**False** 或 **wdToggle**。如果将 **Emboss** 设置为 **True**，则 **Engrave** 将设置为 **False**，反之亦然。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将新文档中第二个句子用阳文显示。*/ function test() { let con = Documents.Add().Content con.InsertAfter("This is the first sentence. ") con.InsertAfter("This is the second sentence. ") con.Sentences.Item(2).Font.Emboss = true }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将选定的文本用阳文显示。*/ function test() { if(Selection.Type == wdSelectionNormal) {     Selection.Font.Emboss = true } else{     MsgBox("You need to select some text.") } }` |

#### **Font.EmphasisMark**

返回或设置 **WdEmphasisMark** 常量，该常量代表字符或指定字符串的着重号。可读写。

**语法**

**express.EmphasisMark**

*express*   一个代表 **Font** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例实现的功能是：在活动文档的第四个单词上面设置着重号：逗号。*/ ActiveDocument.Words.Item(4).EmphasisMark = wdEmphasisMarkOverComma` |

#### **Font.Engrave**

如果字体格式为阴文，则该属性值为 **True**。**Long** 类型，可读写。

**语法**

**express.Engrave**

*express*   一个代表 **Font** 对象的变量。

**说明**

返回 **True**、**False** 或 **wdUndefined**（当返回值既可为 **True**，也可为 **False** 时取该值）。可设置为 **True**、**False** 或 **wdToggle**。如果将 **Engrave** 设置为 **True**，则 **Emboss** 将设置为 **False**，反之亦然。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档中的第一个字母的格式设置为阴文。*/ function test() { let rngTemp = ActiveDocument.Characters.Item(1) rngTemp.Font.Size = 20 rngTemp.Font.Engrave = true }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将选定内容的格式设置为阴文。*/ function test() { if(Selection.Type == wdSelectionNormal){     Selection.Font.Engrave = true } else{     MsgBox("You need to select some text.") } }` |

#### **Font.Fill**

返回 FillFormat 对象，该对象包含指定范围文本使用的字体的填充格式属性。只读。

**语法**

**express.Fill**

*express*   一个代表 **Font** 对象的变量。

#### **Font.Glow**

返回 GlowFormat 对象，该对象代表指定范围文本使用的字体的发光格式。只读。

**语法**

**express.Glow**

*express*   一个代表 **Font** 对象的变量。

#### **Font.Hidden**

如果字体格式为隐藏文字，则该属性值为 **True**。**Long** 类型，可读写。

**语法**

**express.Hidden**

*express*   一个代表 **Font** 对象的变量。

**说明**

该属性返回 **True**、**False** 或 **wdUndefined**（当返回值既可为 **True**，也可为 **False** 时取该值）。可设置为 **True**、**False** 或 **wdToggle**。

若要控制隐藏文字的显示，请使用 **View** 对象的 **ShowHiddenText** 属性。

当不显示隐藏文字时，用 **TextRetrievalMode** 对象的 **IncludeHiddenText** 属性可控制返回 **Range** 对象的属性或方法是否包含隐藏文字。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例检查选定内容中的隐藏文字。*/ function test() {   if(Application.Selection.Type == wdSelectionNormal){       if(Application.Selection.Font.Hidden == wdUndefined || Application.Selection.Font.Hidden == true){           alert("There is hidden text in the selection.")       }       else{           alert("No hidden text in the selection.")       }   }   else{       alert("You need to select some text.")   } }  /*本示例显示活动窗口中的所有隐藏文字并将选定内容的格式设置为隐藏文字。*/ function test() {   Application.ActiveDocument.ActiveWindow.View.ShowHiddenText = true      if(Application.Selection.Type == wdSelectionNormal){       Application.Selection.Font.Hidden = true   } }` |

#### **Font.Italic**

如果将字体或范围设置为倾斜格式，则该属性值为 **True**。**Long** 类型，可读写。

**语法**

**express.Italic**

*express*   一个代表 **Font** 对象的变量。

**说明**

该属性返回 **True**、**False** 或 **wdUndefined**（当返回值既可为 **True**，也可为 **False** 时取该值）。可设置为 **True**、**False** 或 **wdToggle**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例检查所选内容中的倾斜格式并删除所选内容中存在的倾斜格式。*/ function test() {   if(Application.Selection.Type == wdSelectionNormal){       let mySel = Application.Selection.Font.Italic          if(mySel == wdUndefined || mySel == true){           alert("there is italic text in selection. " + "Click OK to remove.")           Application.Selection.Font.Italic = false       }       else{           alert("No italic text in the selection.")       }   }   else{       alert("You need to select some text.")   } }` |

#### **Font.ItalicBi**

如果将字体或范围设置为倾斜格式，则该属性值为 **True**。可读写 **Long** 类型。

**语法**

**express.ItalicBi**

*express*   一个代表 **Font** 对象的变量。

**说明**

该属性返回 **True**、**False** 或 **wdUndefined**（用于倾斜和非倾斜混合格式的文本）。可设置为 **True**、**False** 或 **wdToggle**。

**ItalicBi** 属性用于从右向左的语言。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将从右向左语言的活动文档中的第一段设置为倾斜格式。*/ ActiveDocument.Paragraphs.Item(1).Range.ItalicBi = true` |

#### **Font.Kerning**

返回或设置 WPS 根据字号自动调整字距时的最小字号。**Single** 类型，可读写。

**语法**

**express.Kerning**

*express*   一个代表 **Font** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档进行自动字距调整的最小字号设为 12 磅或更大。*/ ActiveDocument.Content.Font.Kerning = 12` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示 WPS 根据字号自动调整选定文本的字距时的最小字号。*/ function test() { if(Selection.Type == wdSelectionNormal) {     MsgBox(Selection.Font.Kerning) } else {     MsgBox("You need to select some text.") } }` |

#### **Font.Ligatures**

返回或设置指定 **Font** 对象的连字设置。可读/写 WdLigatures。

**语法**

**express.Ligatures**

*express*   一个代表 **Font** 对象的变量。

**说明**

OpenType 字体支持使用连字。**Ligatures** 属性指定要应用于 OpenType 字体的 WPS 2015 连字设置。

下表列出了连字的四个基本值。

| 值     | 说明                                                         |
| ------ | ------------------------------------------------------------ |
| 标准   | 设计用来增强可读性和吸引力。拉丁语言中的标准连字包括“fi”、“fl”和“ff”。 |
| 上下文 | 设计用来通过根据周围文本提供最佳连字选择来增强可读性和吸引力。 |
| 历史   | 较老的、装饰用连字，对于现代读者可能看起来比较古老。不是专为可读性而设计的。 |
| 任意   | 旨在起装饰作用，不是为了可读性。                             |

这四个基本值的组合构成了 **Ligatures** 属性的可用值集合。该组值在 WdLigatures 枚举中提供。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码示例将任意连字应用于活动文档中的字体。*/ ActiveDocument.Range.Font.Ligatures = wdLigaturesDiscretional` |

#### **Font.Line**

返回 LineFormat 对象，该对象指定行的格式。可读/写。

**语法**

**express.Line**

*express*   一个代表 **Font** 对象的变量。

#### **Font.Name**

返回或设置指定对象的名称。**String** 类型，可读写。

**语法**

**express.Name**

*express*   一个代表 **Font** 对象的变量。

**说明**

本示例将选定内容设置为 Arial 字体的加粗格式。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将选定内容设置为 Arial 字体的加粗格式。*/ function test() {   Application.Selection.Font.Name = "Arial"   Application.Selection.Font.Bold = true }` |

#### **Font.NameAscii**

返回或设置拉丁语（字符代码从 0（零）到 127 的字符）所用的字体。**String** 类型，可读写。

**语法**

**express.NameAscii**

*express*   一个代表 **Font** 对象的变量。

**说明**

在 WPS 美国英语版中，该属性的默认值是 Times New Roman。使用 Name属性可以更改应用于所有文本的字体和**“格式”**工具栏**“字体”**框中显示的字体。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置拉丁语所用的字体。*/ Application.Selection.Font.NameAscii = "Century"` |

#### **Font.NameBi**

返回或设置使用从右向左语言的文档中的字体的名称。**String** 类型，可读写。

**语法**

**express.NameBi**

*express*   一个代表 **Font** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将选定内容的格式设为 Arial 字体。*/ Selection.Font.NameBi = "Arial"` |

#### **Font.NameFarEast**

返回或设置一种东亚字体名称。**String** 类型，可读写。

**语法**

**express.NameFarEast**

*express*   一个代表 **Font** 对象的变量。

**说明**

在英语（美国）版的 WPS 中，本属性的默认值是 Times New Roman。推荐使用本属性返回或设置文档中的亚洲文本所用的字体，该文档是用 WPS 的亚洲版创建的。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示应用于选定内容的东亚字体名称。*/ MsgBox(Selection.Font.NameFarEast)` |

#### **Font.NameOther**

返回或设置字符代码从 128 到 255 的字符所用的字体。**String** 类型，可读写。

**语法**

**express.NameOther**

*express*   一个代表 **Font** 对象的变量。

**说明**

在 WPS 美国英语版中，该属性的默认值是 Times New Roman。使用 **Name** 属性可以更改应用于所有文本的字体和**“格式”**工具栏**“字体”**框中显示的字体。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置字符代码从 128 到 255 的字符所用的字体。*/ Selection.Font.NameOther = "Century"` |

#### **Font.NumberForm**

返回或设置 OpenType 字体的数字形式设置。可读/写 WdNumberForm。

**语法**

**express.NumberForm**

*express*   一个代表 **Font** 对象的变量。

**说明**

可以通过一致的高度配合文本基线（称为“内衬”）来显示 OpenType 字体中的数字，也可以通过变动高度（称为“悬挂”或“旧样式”）来显示这些数字，其中数字显示在文本基线之上或之下。使用 **NumberForm** 属性可指定是使用内衬还是旧样式来显示数字。

设置此属性与选择**“数字形式:”**（ WPS 2015 中**“字体”**对话框的**“高级”**选项卡上的**“OpenType 功能”**组）旁边的下拉框中的项具有相同作用。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码示例将活动文档中字体的数字形式设置为“旧样式”。*/ ActiveDocument.Range.Font.NumberForm = wdNumberFormOldStyle` |

#### **Font.NumberSpacing**

返回或设置字体的数字间距设置。可读/写 WdNumberSpacing。

**语法**

**express.NumberSpacing**

*express*   一个代表 **Font** 对象的变量。

**说明**

OpenType 字体支持成比例和表格图表功能来控制数字间距。成比例数字间距按不同的宽度处理每个数字。例如，“1”的显示要比“5”窄。表格数字间距按宽度相等处理数字，从而使它们能够垂直对齐，以增加可读性，特别是对于财务信息。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码示例将活动文档中字体的数字间距设置为成比例。*/ ActiveDocument.Range.Font.NumberSpacing = wdNumberSpacingProportional` |

#### **Font.Outline**

如果字体格式为镂空，则该属性值为 **True**。**Long** 类型，可读写。

**语法**

**express.Outline**

*express*   一个代表 **Font** 对象的变量。

**说明**

返回 **True**、**False** 或 **wdUndefined**（当返回值既可为 **True**，也可为 **False** 时取该值）。可设置为 **True**、**False** 或 **wdToggle**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为活动文档前三个词设置镂空字体格式。*/ function test() { let myRange = ActiveDocument.Range(ActiveDocument.Words.Item(1).Start, ActiveDocument.Words.Item(3).End) myRange.Font.Outline = true }  ` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为选定文本切换镂空格式。*/ Selection.Font.Outline = wdToggle` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*当选定内容中的部分文本应用了镂空格式时，本示例清除选定内容中的镂空字体格式。*/ function test() { let myFont = Selection.Font if(myFont.Outline == wdUndefined){     myFont.Outline = false } }` |

#### **Font.Parent**

返回一个 **Object** 类型值，该值代表指定 **Font** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **Font** 对象的变量。

#### **Font.Position**

返回或设置相对于基准线的文本位置（以磅为单位）。可读写 **Long** 类型。

**语法**

**express.Position**

*express*   一个代表 **Font** 对象的变量。

**说明**

正值将文本提升，负值将文本降低。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将所选文本降低 2 磅。*/ Selection.Font.Position = -2` |

#### **Font.Reflection**

返回一个 ReflectionFormat 对象，该对象代表形状的反射格式。只读。

**语法**

**express.Reflection**

*express*   一个代表 **Font** 对象的变量。

#### **Font.Scaling**

返回或设置应用于字体的缩放百分比。**Long** 类型，可读写。

**语法**

**express.Scaling**

*express*   一个代表 **Font** 对象的变量。

**说明**

本属性以当前字体大小的百分比水平拉长或压缩文字（缩放范围从 1 到 600）。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档中的文字水平拉长到原大小的 110%。*/ ActiveDocument.Content.Font.Scaling = 110` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 Sales.doc 中第一段的文字压缩到原大小的 90%。*/ function test() { let myFont = Documents.Item("Sales.doc").Paragraphs.Item(1).Range.Font myFont.Scaling = 90 myFont.Bold = false }` |

#### **Font.Shading**

返回一个 **Shading** 对象，该对象代表指定字体的底纹格式。

**语法**

**express.Shading**

*express*   一个代表 **Font** 对象的变量。

#### **Font.Shadow**

如果将指定字体设置为阴影格式，则该属性值为 **True**。可读写 **Long** 类型。

**语法**

**express.Shadow**

*express*   一个代表 **Font** 对象的变量。

**说明**

此属性可以是 **True**、**False** 或 **wdUndefined**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例对所选内容应用阴影和加粗格式。*/ function test() { if(Selection.Type == wdSelectionNormal){     Selection.Font.Shadow = true     Selection.Font.Bold = true } else {     MsgBox("You need to select some text.") } }  ` |

#### **Font.Size**

返回或设置字号（以磅为单位）。**Single** 类型，可读写。

**语法**

**express.Size**

*express*   一个代表 **Font** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例插入文本并将该文本的第七个单词的字号设置为 20 磅。*/ function test() {   Application.Selection.Collapse(wdCollapseEnd)   let rng = Application.Selection.Range   rng.Font.Reset()   rng.InsertBefore( "This is a demonstration of font size.")   rng.Words.Item(7).Font.Size = 20 }  /*以下示例确定所选文本的字号。*/ function test() {   let mySel = Application.Selection.Font.Size   if(mySel == wdUndefined){       alert("there is a mix of font sizes in the selection.")   }   else {       alert(mySel + " points")   } }` |

#### **Font.SizeBi**

返回或设置字体大小（以磅为单位）。**Single** 类型，可读写。

**语法**

**express.SizeBi**

*express*   一个代表 **Font** 对象的变量。

**说明**

**SizeBi** 属性适用于使用从右向左语言的文本。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将第一个单词的字体大小设置为 20 磅。*/ ActiveDocument.Paragraphs.Item(1).Range.Words.Item(1).Font.SizeBi = 20` |

#### **Font.SmallCaps**

如果字体格式为小型大写字母，则该属性值为 **True**。**Long** 类型，可读写。

**语法**

**express.SmallCaps**

*express*   一个代表 **Font** 对象的变量。

**说明**

返回 **True**、**False** 或 **wdUndefined**（当返回值既可为 **True**，也可为 **False** 时取该值）。可设置为 **True**、**False** 或 **wdToggle**。

如果将 **SmallCaps** 属性设为 **True**，则 **AllCaps** 属性会设为 **False**，反之亦然。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例区分新文档中小型大写字母和全部大写字母之间的差别。*/ function test() { let myRange = Documents.Add().Content     myRange.InsertAfter("This is a demonstration of SmallCaps.")     myRange.Words.Item(6).Font.SmallCaps = true     myRange.InsertParagraphAfter()     myRange.InsertAfter("This is a demonstration of AllCaps.")     myRange.Words.Item(14).Font.AllCaps = true }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果部分选定内容中的字体格式已设置为小型大写字母格式，本示例将整个选定内容的字体格式设置为小型大写字母格式。*/ function test() { if(Selection.Type == wdSelectionNormal) {      let mySel = Selection.Font.SmallCaps     if(mySel == wdUndefined) {         Selection.Font.SmallCaps = true     } } else {     MsgBox("You need to select some text.") } }` |

#### **Font.Spacing**

返回或设置字符的间距（以磅为单位）。可读写 **Single** 类型。

**语法**

**express.Spacing**

*express*   一个代表 **Font** 对象的变量。

**说明**

------

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例演示活动文档开头的两种不同字符间距。*/ function test() { let myRange = ActiveDocument.Range(0, 0)     myRange.InsertAfter("Demonstration of no character spacing.")     myRange.InsertParagraphAfter()     myRange.InsertAfter("Demonstration of character spacing (1.5pt).")     myRange.InsertParagraphAfter()  ActiveDocument.Paragraphs.Item(2).Range.Font.Spacing = 1.5 }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将所选文本的字符间距设置为 2 磅。*/ function test() { if(Selection.Type == wdSelectionNormal){     Selection.Font.Spacing = 2 } else{     MsgBox("You need to select some text.") } }` |

#### **Font.StrikeThrough**

如果字体格式为带有删除线，则该属性值为 **True**。**Long** 类型，可读写。

**语法**

**express.StrikeThrough**

*express*   一个代表 **Font** 对象的变量。

**说明**

**StrikeThrough** 属性返回 **True**、**False** 或 **wdUndefined**（当返回值既可为 **True**，也可为 **False** 时取该值）。可设置为 **True**、**False** 或 **wdToggle**。

使用 **DoubleStrikeThrough** 属性可设置或返回双删除线格式。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为活动文档中的前三个单词设置删除线格式。*/ function test() { let myDoc = ActiveDocument let myRange = myDoc.Range(myDoc.Words.Item(1).Start, myDoc.Words.Item(3).End) myRange.Font.StrikeThrough = true }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为选定文本设置删除线格式。*/ function test() { if(Selection.Type == wdSelectionNormal){     Selection.Font.StrikeThrough = true } else{     MsgBox("You need to select some text.") } }` |

#### **Font.StylisticSet**

指明指定字体的样式集。可读/写 WdStylisticSet。

**语法**

**express.StylisticSet**

*express*   一个代表 **Font** 对象的变量。

**说明**

一些 OpenType 字体提供了样式集。样式集定义字体内要一起使用的一组字符，通常是为了获得视觉上的协调性，例如用于标题中。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码示例将活动文档的字体设置为 Gabriola，然后应用 Gabriola 字体提供的第六个样式集。*/ function test() { ActiveDocument.Range.Font.Name = "Gabriola" ActiveDocument.Range.Font.StylisticSet = wdStylisticSet06 }` |

#### **Font.Subscript**

如果字体格式为下标，则该属性值为 **True**。**Long** 类型，可读写。

**语法**

**express.Subscript**

*express*   一个代表 **Font** 对象的变量。

**说明**

返回 **True**、**False** 或 **wdUndefined**（当返回值既可为 **True**，也可为 **False** 时取该值）。可设置为 **True**、**False** 或 **wdToggle**。

如果将 **Subscript** 属性设为 **True**，则 **Superscript** 属性会设为 **False**，反之亦然。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在活动文档的开头插入文本，并将第十个字符设为下标。*/ function test() { let myRange = ActiveDocument.Range(0, 0) myRange.InsertAfter("Water = H20") myRange.Characters.Item(10).Font.Subscript = true }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例检查所选文本的下标格式。*/ function test() { if(Selection.Type == wdSelectionNormal){     let mySel = Selection.Font.Subscript     if(mySel == wdUndefined || mySel == true){         MsgBox("Subscript text exists in the selection.")     }     else{         MsgBox("No subscript text in the selection.")     } } else{     MsgBox("You need to select some text.") } }` |

#### **Font.Superscript**

如果字体格式为上标，则该属性值为 **True**。**Long** 类型，可读写。

**语法**

**express.Superscript**

*express*   一个代表 **Font** 对象的变量。

**说明**

返回 **True**、**False** 或 **wdUndefined**（当返回值既可为 **True**，也可为 **False** 时取该值）。可设置为 **True**、**False** 或 **wdToggle**。

如果将 **Superscript** 属性设为 **True**，则 **Subscript** 属性会设为 **False**，反之亦然。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在活动文档的开头插入文本，并将第四个单词的两个字符设为上标。*/ function test() { let myRange = ActiveDocument.Range(0, 0) myRange.InsertAfter("Superscript in the 4th word.") ActiveDocument.Range(20, 22).Font.Superscript = true }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将所选文本设为上标。*/ function test() { if(Selection.Type == wdSelectionNormal){     Selection.Font.Superscript = true } else{     MsgBox("You need to select some text.") } }` |

#### **Font.TextColor**

返回一个代表指定字体的颜色的 [ColorFormat](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/ColorFormat/ColorFormat%20.htm#jsObject_ColorFormat)[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/ColorFormat/ColorFormat%20.htm#jsObject_ColorFormat)对象。只读。

**语法**

**express.TextColor**

*express*   一个代表 **Font** 对象的变量。

#### **Font.TextShadow**

返回 ShadowFormat 对象，该对象指明指定字体的阴影格式。

**语法**

**express.TextShadow**

*express*   一个代表 **Font** 对象的变量。

#### **Font.ThreeD**

返回 ThreeDFormat 对象，该对象包含指定字体的三维效果格式属性。只读。

**语法**

**express.ThreeD**

*express*   一个代表 **Font** 对象的变量。

#### **Font.Underline**

返回或设置应用于字体的下划线类型。可读写 **WdUnderline**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdUnderline%20%E6%9E%9A%E4%B8%BE.html)类型。

**语法**

**express.Underline**

*express*   一个代表 **Font** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例给所选内容添加单下划线。*/ function test() {   if(Application.Selection.Type == wdSelectionNormal){       Application.Selection.Font.Underline = wdUnderlineSingle   }   else{       alert("You need to select some text.")   } }` |

#### **Font.UnderlineColor**

返回或设置指定 **Font** 对象的下划线的 24 位颜色。

**语法**

**express.UnderlineColor**

*express*   一个代表 **Font** 对象的变量。

**说明**

此属性可以是任何有效的 **WdColor** 常量，也可以是 Visual Basic 的 **RGB** 函数返回的值。将 **UnderlineColor** 属性设置为 **wdColorAutomatic** 会将下划线的颜色重新设置为其上文本的颜色。

适用环境：web

适用平台：windows/linux