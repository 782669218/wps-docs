**TextEffectFormat**



包含应用于艺术字对象的属性和方法。

**说明**

使用 **TextEffect** 属性可返回一个 **TextEffectFormat** 对象。

下例为 *myDocument* 上的形状一设置字体名称及格式。要运行本示例，形状一必须是艺术字对象。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let myDocument = Application.Worksheets.Item(1)     let TextEffect = myDocument.Shapes(1).TextEffect     TextEffect.FontName = "Courier New"     TextEffect.FontBold = true     TextEffect.FontItalic = true }  ` |

**方法**

|                                                              | 名称                   | 说明                                                     |
| ------------------------------------------------------------ | ---------------------- | -------------------------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ToggleVerticalText** | 将指定艺术字中的文字排列方向从水平切换为垂直，反之亦然。 |

**属性**

|                                                              | 名称                 | 说明                                                         |
| ------------------------------------------------------------ | -------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Alignment**        | 返回或设置一个 **MsoTextEffectAlignment** 值，它代表艺术字的对齐方式。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**      | 如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**          | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FontBold**         | 如果指定艺术字中的字体格式为加粗，则该属性值为 **True**。**MsoTriState** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FontItalic**       | 如果指定艺术字中的字体是倾斜的，则返回 **msoTrue**。**MsoTriState** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FontName**         | 返回或设置指定艺术字的字体名称。**String** 类型，可读写。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FontSize**         | 以磅为单位返回或设置指定艺术字的字体大小。**Single** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **KernedPairs**      | 如果指定艺术字中的字符对是自动缩紧的，则该属性值为 **True**。**MsoTriState** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NormalizedHeight** | 如果指定艺术字中所有字符（无论大小写）均为同一高度，则该属性值为 **True**。**MsoTriState** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**           | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PresetShape**      | 返回或设置指定艺术字的形状样式。**MsoPresetTextEffectShape** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PresetTextEffect** | 返回或设置指定艺术字的样式。**MsoPresetTextEffect** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RotatedChars**     | 如果指定艺术字对象中的字符相对于该对象旋转了 90 度，则该值为 **True**。如果指定艺术字对象中的字符相对于该对象保持原有方向，则该值为 **False**。**MsoTriState** 类型，可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Text**             | 返回或设置指定对象中的文本。**String** 型，可读写。          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Tracking**         | 返回或设置指定艺术字中每个字符的水平间隔相对于字符宽度的比例。该比例可为从 0 到 5 之间的值。（为该属性指定的值越大，字符间的空间越大；小于 1 的值会产生字符重叠）。可读写。**Single** 类型。 |

**成员方法**

#### **TextEffectFormat.ToggleVerticalText**

将指定艺术字中的文字排列方向从水平切换为垂直，反之亦然。

**语法**

**express.ToggleVerticalText()**

*express*   一个代表 **TextEffectFormat** 对象的变量。

**说明**

使用 **ToggleVerticalText** 方法可以交换代表艺术字的 **Shape** 对象的 **Width** 和 **Height** 属性的值，但不改变 **Left** 和 **Top** 属性。

**Shape** 对象的 **Flip** 方法和 **Rotation** 属性以及 **TextEffectFormat** 对象的 **RotatedChars** 属性和 **ToggleVerticalText** 方法都会影响代表艺术字的 **Shape** 对象中字符方向和文本排列方式。可能必须经过试验才能找到组合使用这些属性和方法的最佳途径以达到预期效果。

**示例**

本示例向 `myDocument` 中添加了包含文本“Test”的艺术字，并将其水平文字排列方向（对于指定的艺术字样式 **msoTextEffect1**，这是默认的文字排列方向）更改为垂直文字排列方向。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let myDocument = Application.Worksheets.Item(1)     let newWordArt = myDocument.Shapes.AddTextEffect(Application.Enum.msoTextEffect1, "Test", "Arial Black", 36, false, false, 100, 100)     newWordArt.TextEffect.ToggleVerticalText() } ` |

**成员属性**

#### **TextEffectFormat.Alignment**

返回或设置一个 **MsoTextEffectAlignment** 值，它代表艺术字的对齐方式。

**语法**

**express.Alignment**

*express*   一个代表 **TextEffectFormat** 对象的变量。

**示例**

此示例向第一张工作表中添加艺术字对象，然后将该对象右对齐

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let mySh = Application.Worksheets.Item(1).Shapes     let myTE = mySh.AddTextEffect(Application.Enum.msoTextEffect1, "Test Text", "Palatino", 54, true, false, 100, 50)     myTE.TextEffect.Alignment = Application.Enum.msoTextEffectAlignmentRight } ` |

#### **TextEffectFormat.Application**

如果不使用对象识别符，则该属性返回一个 **Application** 对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **TextEffectFormat** 对象的变量。

**示例**

myObject

 

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let myObject = Application.ActiveWorkbook     if (myObject.Application.Value == "ET") {         alert("This is an ET Application object.")     } else {         alert("This is not an ET Application object.")     } }` |

#### **TextEffectFormat.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **TextEffectFormat** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **TextEffectFormat.FontBold**

如果指定艺术字中的字体格式为加粗，则该属性值为 **True**。**MsoTriState** 类型，可读写。

**语法**

**express.FontBold**

*express*   一个代表 **TextEffectFormat** 对象的变量。

**说明**

| **MsoTriState** 可以是下列 **MsoTriState** 常量之一。 |
| ----------------------------------------------------- |
| **msoCTrue**                                          |
| **msoFalse**                                          |
| **msoTriStateMixed**                                  |
| **msoTriStateToggle**                                 |
| **msoTrue** 指定的艺术字为粗体。                      |

**示例**

本示例将myDocument中第三个形状（如果该形状为艺术字）的字体设置为粗体。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let myDocument = Application.Worksheets.Item(1)     let rng = myDocument.Shapes.Item(3)     if (rng.Type == Application.Enum.msoTextEffect) {         rng.TextEffect.FontBold = Application.Enum.msoTrue     } }` |

#### **TextEffectFormat.FontItalic**

如果指定艺术字中的字体是倾斜的，则返回 **msoTrue**。**MsoTriState** 类型，可读写。

**语法**

**express.FontItalic**

*express*   一个代表 **TextEffectFormat** 对象的变量。

**说明**

| **MsoTriState** 可以是下列 **MsoTriState** 常量之一。 |
| ----------------------------------------------------- |
| **msoCTrue** 不应用于该属性。                         |
| **msoFalse** 指定的艺术字对象中的文字不是倾斜的。     |
| **msoTriStateMixed** 不应用于该属性。                 |
| **msoTriStateToggle** 不应用于该属性。                |
| **msoTrue** 指定的艺术字对象中的文字是倾斜的。        |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let myDocument = Application.Worksheets.Item(1)     myDocument.Shapes.Item("WordArt 4").TextEffect.FontItalic = Application.Enum.msoTrue }` |

本示例将 `myDocument` 中形状“WordArt 4”的字体设置为倾斜

#### **TextEffectFormat.FontName**

返回或设置指定艺术字的字体名称。**String** 类型，可读写。

**语法**

**express.FontName**

*express*   一个代表 **TextEffectFormat** 对象的变量。

**示例**

 

myDocument

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let myDocument = Application.Worksheets.Item(1)     let rng = myDocument.Shapes.Item(3)     if (rng.Type == Application.Enum.msoTextEffect) {                 rng.TextEffect.FontName = "Courier New"     } }` |

#### **TextEffectFormat.FontSize**

以磅为单位返回或设置指定艺术字的字体大小。**Single** 类型，可读写。

**语法**

**express.FontSize**

*express*   一个代表 **TextEffectFormat** 对象的变量。

**示例**

本示例将 `myDocument` 中名为“WordArt 4”的形状中的字号设置为 16 磅。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let myDocument = Application.Worksheets.Item(1)     myDocument.Shapes.Item("WordArt 4").TextEffect.FontSize = 16 }` |

#### **TextEffectFormat.KernedPairs**

如果指定艺术字中的字符对是自动缩紧的，则该属性值为 **True**。**MsoTriState** 类型，可读写。

**语法**

**express.KernedPairs**

*express*   一个代表 **TextEffectFormat** 对象的变量。

**说明**

| **MsoTriState** 可以是下列 **MsoTriState** 常量之一。 |
| ----------------------------------------------------- |
| **msoCTrue**                                          |
| **msoFalse**                                          |
| **msoTriStateMixed**                                  |
| **msoTriStateToggle**                                 |
| **msoTrue** 指定艺术字中的字符对自动紧缩。            |

**示例**

本示例打开 `myDocument` 中第三个形状（如果该形状为艺术字）的字符对自动紧缩。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let myDocument = Application.Worksheets.Item(1)     let rng = myDocument.Shapes.Item(3)     if (rng.Type == Application.Enum.msoTextEffect) {         rng.TextEffect.KernedPairs = Application.Enum.msoTrue     } }` |

#### **TextEffectFormat.NormalizedHeight**

如果指定艺术字中所有字符（无论大小写）均为同一高度，则该属性值为 **True**。**MsoTriState** 类型，可读写。

**语法**

**express.NormalizedHeight**

*express*   一个代表 **TextEffectFormat** 对象的变量。

**说明**

| **MsoTriState** 可以是下列 **MsoTriState** 常量之一。        |
| ------------------------------------------------------------ |
| **msoCTrue**                                                 |
| **msoFalse**                                                 |
| **msoTriStateMixed**                                         |
| **msoTriStateToggle**                                        |
| **msoTrue** 指定艺术字中的所有字符（无论大小写）均为同一高度。 |

**示例**

 

myDocument

 

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let myDocument = Application.Worksheets.Item(1)     myDocument.Shapes.AddTextEffect(Application.Enum.msoTextEffect1, "Test Effect", "Courier New", 44, true, false, 10, 10).Name = "texteff1"     myDocument.Shapes.Item("texteff1").TextEffect.NormalizedHeight = Application.Enum.msoTrue } ` |

#### **TextEffectFormat.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **TextEffectFormat** 对象的变量。

#### **TextEffectFormat.PresetShape**

返回或设置指定艺术字的形状样式。**MsoPresetTextEffectShape** 类型，可读写。

**语法**

**express.PresetShape**

*express*   一个代表 **TextEffectFormat** 对象的变量。

**说明**

| **MsoPresetTextEffectShape** 可以是下列 **MsoPresetTextEffectShape** 常量之一。 |
| ------------------------------------------------------------ |
| **msoTextEffectShapeArchDownCurve**                          |
| **msoTextEffectShapeArchDownPour**                           |
| **msoTextEffectShapeArchUpCurve**                            |
| **msoTextEffectShapeArchUpPour**                             |
| **msoTextEffectShapeButtonCurve**                            |
| **msoTextEffectShapeButtonPour**                             |
| **msoTextEffectShapeCanDown**                                |
| **msoTextEffectShapeCanUp**                                  |
| **msoTextEffectShapeCascadeDown**                            |
| **msoTextEffectShapeCascadeUp**                              |
| **msoTextEffectShapeChevronDown**                            |
| **msoTextEffectShapeChevronUp**                              |
| **msoTextEffectShapeCircleCurve**                            |
| **msoTextEffectShapeCirclePour**                             |
| **msoTextEffectShapeCurveDown**                              |
| **msoTextEffectShapeCurveUp**                                |
| **msoTextEffectShapeDeflate**                                |
| **msoTextEffectShapeDeflateBottom**                          |
| **msoTextEffectShapeDeflateInflateDeflate**                  |
| **msoTextEffectShapeDoubleWave1**                            |
| **msoTextEffectShapeFadeDown**                               |
| **msoTextEffectShapeFadeRight**                              |
| **msoTextEffectShapeInflate**                                |
| **msoTextEffectShapeInflateTop**                             |
| **msoTextEffectShapePlainText**                              |
| **msoTextEffectShapeRingOutside**                            |
| **msoTextEffectShapeSlantUp**                                |
| **msoTextEffectShapeTriangleDown**                           |
| **msoTextEffectShapeWave1**                                  |
| **msoTextEffectShapeDeflateInflate**                         |
| **msoTextEffectShapeDeflateTop**                             |
| **msoTextEffectShapeDoubleWave2**                            |
| **msoTextEffectShapeFadeLeft**                               |
| **msoTextEffectShapeFadeUp**                                 |
| **msoTextEffectShapeInflateBottom**                          |
| **msoTextEffectShapeMixed**                                  |
| **msoTextEffectShapeRingInside**                             |
| **msoTextEffectShapeSlantDown**                              |
| **msoTextEffectShapeStop**                                   |
| **msoTextEffectShapeTriangleUp**                             |
| **msoTextEffectShapeWave2**                                  |

设置 **PresetTextEffect** 属性会自动设置 **PresetShape** 属性。

**示例**

本示例将 `myDocument` 中的所有艺术字形状设置为中心指向下方的 V 字形。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let myDocument = Application.Worksheets.Item(1)     for (let s = 1; s <= myDocument.Shapes.Count; s++) {         if (myDocument.Shapes.Item(s).Type == Application.Enum.msoTextEffect) {             myDocument.Shapes.Item(s).TextEffect.PresetShape = Application.Enum.msoTextEffectShapeChevronDown         }     } }` |

#### **TextEffectFormat.PresetTextEffect**

返回或设置指定艺术字的样式。**MsoPresetTextEffect** 类型，可读写。

**语法**

**express.PresetTextEffect**

*express*   一个代表 **TextEffectFormat** 对象的变量。

**说明**

该属性的值与**“艺术字库”**对话框内的格式（按从左至右和从上至下的顺序编号）相对应。

| **MsoPresetTextEffect** 可以是下列 **MsoPresetTextEffect** 常量之一。 |
| ------------------------------------------------------------ |
| **msoTextEffect1**                                           |
| **msoTextEffect10**                                          |
| **msoTextEffect11**                                          |
| **msoTextEffect12**                                          |
| **msoTextEffect13**                                          |
| **msoTextEffect14**                                          |
| **msoTextEffect15**                                          |
| **msoTextEffect16**                                          |
| **msoTextEffect17**                                          |
| **msoTextEffect18**                                          |
| **msoTextEffect19**                                          |
| **msoTextEffect2**                                           |
| **msoTextEffect20**                                          |
| **msoTextEffect21**                                          |
| **msoTextEffect22**                                          |
| **msoTextEffect23**                                          |
| **msoTextEffect24**                                          |
| **msoTextEffect25**                                          |
| **msoTextEffect26**                                          |
| **msoTextEffect27**                                          |
| **msoTextEffect28**                                          |
| **msoTextEffect29**                                          |
| **msoTextEffect3**                                           |
| **msoTextEffect30**                                          |
| **msoTextEffect4**                                           |
| **msoTextEffect5**                                           |
| **msoTextEffect6**                                           |
| **msoTextEffect7**                                           |
| **msoTextEffect8**                                           |
| **msoTextEffect9**                                           |
| **msoTextEffectMixed**                                       |

设置 **PresetTextEffect** 属性会自动设置指定形状的许多其他格式属性。

**示例**

本示例将 `myDocument` 中所有的艺术字样式设置为“艺术字库”对话框中列出的第一种样式。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let myDocument =  Application.Worksheets.Item(1)     for (let s = 1; s <= myDocument.Shapes.Count; s++) {         if (myDocument.Shapes.Item(s).Type == Application.Enum.msoTextEffect) {             myDocument.Shapes.Item(s).TextEffect.PresetTextEffect = Application.Enum.msoTextEffect1         }     } }` |

#### **TextEffectFormat.RotatedChars**

如果指定艺术字对象中的字符相对于该对象旋转了 90 度，则该值为 **True**。如果指定艺术字对象中的字符相对于该对象保持原有方向，则该值为 **False**。**MsoTriState** 类型，可读/写。

**语法**

**express.RotatedChars**

*express*   一个代表 **TextEffectFormat** 对象的变量。

**说明**

| **MsoTriState** 可以是下列 **MsoTriState** 常量之一。        |
| ------------------------------------------------------------ |
| **msoCTrue**                                                 |
| **msoFalse** 指定艺术字中的字符保留原有的相对于边框形状的方向。 |
| **msoTriStateMixed**                                         |
| **msoTriStateToggle**                                        |
| **msoTrue** 将指定艺术字中的字符相对于艺术字边框形状旋转 90 度。 |

如果艺术字中含有水平文字，将 **RotatedChars** 属性设置为 **msoTrue** 可将该字符逆时针旋转 90 度。如果艺术字中含有垂直文字，将 **RotatedChars** 属性设置为 **msoFalse** 可将该字符顺时针旋转 90 度。使用 **ToggleVerticalText** 方法在水平文字流和垂直文字流之间切换。

**Shape** 对象的 **Flip** 方法和 **Rotation** 属性以及 **TextEffectFormat** 对象的 **RotatedChars** 属性和 **ToggleVerticalText** 方法会对表示艺术字的 **Shape** 对象中的字符的方向和文字排列的方向产生影响。可能必须经过试验才能找到组合使用这些属性和方法的最佳途径以达到预期效果。

**示例**

本例向 `myDocument` 中添加带有“Test”文本的艺术字，并将该字符逆时针旋转 90 度。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let myDocument = Application.Worksheets.Item(1)     let newWordArt = myDocument.Shapes.AddTextEffect(Application.Enum.msoTextEffect1, "Test", "Arial Black", 36, false, false, 10, 10)     newWordArt.TextEffect.RotatedChars = Application.Enum.msoTrue }` |

#### **TextEffectFormat.Text**

返回或设置指定对象中的文本。**String** 型，可读写。

**语法**

**express.Text**

*express*   一个代表 **TextEffectFormat** 对象的变量。

#### **TextEffectFormat.Tracking**

返回或设置指定艺术字中每个字符的水平间隔相对于字符宽度的比例。该比例可为从 0 到 5 之间的值。（为该属性指定的值越大，字符间的空间越大；小于 1 的值会产生字符重叠）。可读写。**Single** 类型。

**语法**

**express.Tracking**

*express*   一个代表 **TextEffectFormat** 对象的变量。

**说明**

以下表格相对于用户界面上的可用设置给出 **Tracking** 属性的值。

| 用户界面设置 | 相应的 Tracking 属性值 |
| ------------ | ---------------------- |
| 很紧         | 0.8                    |
| 紧密         | 0.9                    |
| 正常         | 1.0                    |
| 稀疏         | 1.2                    |
| 很松         | 1.5                    |

**示例**

本例向 `myDocument` 中添加包含“Test”文本的艺术字，并且指定这些字符紧密排列。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let myDocument = Application.Worksheets.Item(1)     let newWordArt = myDocument.Shapes.AddTextEffect(Application.Enum.msoTextEffect1, "Test", "Arial Black", 36, false, false, 100, 100)     newWordArt.TextEffect.Tracking = 0.8 }` |

适用环境：web

适用平台：windows/linux