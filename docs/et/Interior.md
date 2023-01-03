**Interior**



代表一个对象的内部。

**说明**

使用 **Interior** 属性可返回 **Interior** 对象。下例将单元格 A1 的内部设置为红色。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.Worksheets.Item("Sheet1").Range("A1").Interior.ColorIndex = 3` |

**属性**

|                                                              | 名称                    | 说明                                                         |
| ------------------------------------------------------------ | ----------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**         | 如果不使用对象识别符，则该属性返回一个代表 ET 应用程序的 **Application** 对象。如果使用对象识别符，则该属性返回一个代表指定对象的创建程序的 **Application** 对象（可对一个 OLE 自动化对象使用该属性来返回该对象的应用程序）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Color**               | 返回或设置对象的主要颜色，如注释部分中的表格所示。使用 **RGB** 函数可创建颜色值。**Variant** 型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ColorIndex**          | 返回或设置一个**Variant** 值，它代表内部颜色。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**             | 返回一个 32 位整数，该整数指示创建对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Gradient**            | 返回或设置所选内容的 **Interior** 对象的 **Gradient** 属性。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **InvertIfNegative**    | 如果指定项与一个负数相对应时 ET 就将其反色，则为 **True**。**Variant** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**              | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Pattern**             | 返回或设置一个包含 **xlPattern** 常量的 **Variant** 值，它代表内部图案。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PatternColor**        | 以 RGB 值返回或设置内部图案的颜色。可读/写 **Variant** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PatternColorIndex**   | 返回或设置内部图案的颜色，表示为当前调色板中的颜色编号或下列 **XlColorIndex** 常量之一：**xlColorIndexAutomatic** 或 **xlColorIndexNone**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PatternThemeColor**   | 返回或设置 **Interior** 对象的主题颜色图案。可读/写 **Variant** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PatternTintAndShade** | 返回或设置 **Interior** 对象的淡色和底纹图案。可读/写 **Variant** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ThemeColor**          | 返回或设置已应用的配色方案中的主题颜色，该配色方案与指定对象相关联。可读/写 **Variant** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TintAndShade**        | 返回或设置一个 **Single**，使颜色变深或变浅。                |

**成员属性**

#### **Interior.Application**

如果不使用对象识别符，则该属性返回一个代表 ET 应用程序的 **Application** 对象。如果使用对象识别符，则该属性返回一个代表指定对象的创建程序的 **Application** 对象（可对一个 OLE 自动化对象使用该属性来返回该对象的应用程序）。只读。

**语法**

**express.Application**

*express*   一个代表 **Interior** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test{   /*本示例显示一条有关创建 myObject 的应用程序的消息。*/   let myObject = Application.ActiveWorkbook   if(myObject.Application.Value == "ET") {       alert("This is an ET Application object.")   }   else {       alert("This is not an ET Application object.")   } }` |

#### **Interior.Color**

返回或设置对象的主要颜色，如注释部分中的表格所示。使用 **RGB** 函数可创建颜色值。**Variant** 型，可读写。

**语法**

**express.Color**

*express*   一个代表 **Interior** 对象的变量。

**说明**

| **Border**   | 边框的颜色。                                                 |
| ------------ | ------------------------------------------------------------ |
| **Borders**  | 一个区域的所有四条边的颜色。如果四边不是同一种颜色，则 **Color** 返回的是 0（零）。 |
| **Font**     | 字体的颜色。                                                 |
| **Interior** | 单元格底纹的颜色或图形对象的填充颜色。                       |
| **Tab**      | 选项卡的颜色。                                               |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例对 Chart1 中数值坐标轴的刻度线标志颜色进行设置。*/ Application.Charts.Item("Chart1").Axes(xlValue).TickLabels.Font.Color = (0, 255, 0)` |

#### **Interior.ColorIndex**

返回或设置一个**Variant** 值，它代表内部颜色。

**语法**

**express.ColorIndex**

*express*   一个代表 **Interior** 对象的变量。

**说明**

颜色可指定为当前调色板中颜色的索引值，也可指定为下列 **XlColorIndex** 常量之一：

- **xlColorIndexAutomatic**
- **xlColorIndexNone**

#### **Interior.Creator**

返回一个 32 位整数，该整数指示创建对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Interior** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **Interior.Gradient**

返回或设置所选内容的 **Interior** 对象的 **Gradient** 属性。只读。

**语法**

**express.Gradient**

*express*   一个代表 **Interior** 对象的变量。

**说明**

使用 LinearGradient 或 RectangularGradient。

#### **Interior.InvertIfNegative**

如果指定项与一个负数相对应时 ET 就将其反色，则为 **True**。**Variant** 类型，可读写。

**语法**

**express.InvertIfNegative**

*express*   一个代表 **Interior** 对象的变量。

#### **Interior.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **Interior** 对象的变量。

#### **Interior.Pattern**

返回或设置一个包含 **xlPattern** 常量的 **Variant** 值，它代表内部图案。

**语法**

**express.Pattern**

*express*   一个代表 **Interior** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为 Sheet1 中 A1 单元格的内部添加十字图案。*/ Application.Worksheets.Item("Sheet1").Range("A1").Interior.Pattern = xlPatternCrissCross` |

#### **Interior.PatternColor**

以 RGB 值返回或设置内部图案的颜色。可读/写 **Variant** 类型。

**语法**

**express.PatternColor**

*express*   一个代表 **Interior** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test{   /*本示例对 Sheet1 中矩形一的内部图案的颜色进行设置。*/   let rectangles2 = Application.Worksheets.Item("Sheet1").Rectangles(1).Interior   rectangles2.Pattern = xlGrid   rectangles2.PatternColor = (255,0,0) }` |

#### **Interior.PatternColorIndex**

返回或设置内部图案的颜色，表示为当前调色板中的颜色编号或下列 **XlColorIndex** 常量之一：**xlColorIndexAutomatic** 或 **xlColorIndexNone**。**Long** 类型，可读写。

**语法**

**express.PatternColorIndex**

*express*   一个代表 **Interior** 对象的变量。

**说明**

将本属性设置为 **xlColorIndexAutomatic**，表示单元格使用自动图案或图形对象使用自动填充样式。将本属性设置为 **xlColorIndexNone**，则表示不希望使用图案（等价于将 **Interior** 对象的 **Pattern** 属性设置为 **xlPatternNone**）。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test{   /*本示例对 Sheet1 中矩形一的内部图案的颜色进行设置。*/   let rectangles2 = Worksheets.Item("Sheet1").Rectangles(1).Interior   rectangles2.Pattern = xlChecker   rectangles2.PatternColorIndex = 5 }` |

#### **Interior.PatternThemeColor**

返回或设置 **Interior** 对象的主题颜色图案。可读/写 **Variant** 类型。

**语法**

**express.PatternThemeColor**

*express*   一个代表 **Interior** 对象的变量。

#### **Interior.PatternTintAndShade**

返回或设置 **Interior** 对象的淡色和底纹图案。可读/写 **Variant** 类型。

**语法**

**express.PatternTintAndShade**

*express*   一个代表 **Interior** 对象的变量。

#### **Interior.ThemeColor**

返回或设置已应用的配色方案中的主题颜色，该配色方案与指定对象相关联。可读/写 **Variant** 类型。

**语法**

**express.ThemeColor**

*express*   一个代表 **Interior** 对象的变量。

#### **Interior.TintAndShade**

返回或设置一个 **Single**，使颜色变深或变浅。

**语法**

**express.TintAndShade**

*express*   一个代表 **Interior** 对象的变量。

**说明**

可以为 **TintAndShade** 属性输入 -1（最暗）到 1（最亮）之间的数字，零 (0) 为中间值。

如果将此属性设置为小于 -1 或大于 1 的值，则会引起运行时错误：“指定的值超出了范围”。此属性用于主题颜色和非主题颜色。

适用环境：web

适用平台：windows/linux