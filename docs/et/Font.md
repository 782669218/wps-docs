**Font**



包含对象的字体属性（字体名称、字号、颜色等等）。

**说明**

如果不想将单元格中的文本或图形设为相同的格式，则使用 **Characters** 属性返回文本的子集。

使用 **Font** 属性可返回 **Font** 对象。下例将单元格区域 A1:C5 的格式设置为加粗。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.Worksheets.Item("Sheet1").Range("A1:C5").Font.Bold = true` |

**属性**

|                                                              | 名称              | 说明                                                         |
| ------------------------------------------------------------ | ----------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**   | 如果不使用对象识别符，则该属性返回一个 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Background**    | 返回或设置图表中使用的文本的背景类型。**Variant** 类型，可读写。可将其设置为 **XlBackground**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/XlBackground%20%E6%9E%9A%E4%B8%BE.html)常量之一。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Bold**          | 如果字体格式为加粗，则该属性值为 **True**。**Variant** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Color**         | 返回或设置对象的主要颜色，如注释部分中的表格所示。使用 **RGB** 函数可创建颜色值。**Variant** 型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ColorIndex**    | 返回或设置一个 **Variant** 值，它代表字体的颜色。            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**       | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FontStyle**     | 返回或设置字体样式。**String** 类型，可读写。                |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Italic**        | 如果字体样式为倾斜，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**          | 返回或设置一个 **Variant** 值，它代表对象的名称。            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**        | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Size**          | 返回或设置字号。可读/写 **Variant** 类型。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Strikethrough** | 如果文字中间有一条水平删除线，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Subscript**     | 如果字体格式设置为下标，则该属性值为 **True**。默认值为 **False**。**Variant** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Superscript**   | 如果字体格式设置为上标字符，则该属性值为 **True**；默认值为 **False**。**Variant** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ThemeColor**    | 返回或设置已应用的配色方案中的主题颜色，该配色方案与指定对象相关联。可读/写 **Variant** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ThemeFont**     | 返回或设置与指定对象关联的应用字体方案中的主题字体。可读/写 **XlThemeFont**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/XlThemeFont%20%E6%9E%9A%E4%B8%BE.html)类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TintAndShade**  | 返回或设置一个 **Single**，使颜色变深或变浅。                |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Underline**     | 返回或设置应用于字体的下划线类型。可为以下 **XlUnderlineStyle**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/XlUnderlineStyle%20%E6%9E%9A%E4%B8%BE.html)常量之一。**Variant** 类型，可读写。 |

**成员属性**

#### **Font.Application**

如果不使用对象识别符，则该属性返回一个 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application** 对象。只读。

**语法**

**express.Application**

*express*   一个代表 **Font** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例显示一条有关创建 myObject 的应用程序的消息*/ let myObject = Application.ActiveWorkbook if(myObject.Application.Value == "ET") {     alert("This is an ET Application object.") } else {     alert("This is not an ET Application object.") }` |

#### **Font.Background**

返回或设置图表中使用的文本的背景类型。**Variant** 类型，可读写。可将其设置为 **XlBackground**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/XlBackground%20%E6%9E%9A%E4%B8%BE.html)常量之一。

**语法**

**express.Background**

*express*   一个代表 **Font** 对象的变量。

**说明**

| XlBackground 可为以下常量之一。 | 文本背景描述                                                 |
| ------------------------------- | ------------------------------------------------------------ |
| **xlBackgroundAutomatic**       | 字体背景将自动更改文本周围的背景区域颜色，以使在文本下方元素所应用的颜色之上显示的文本达到最佳效果 |
| **xlBackgroundOpaque**          | 如果文本颜色与文本下的填充颜色非常接近或者为同一颜色，则字体背景会设置为空，以便不显示文本 |
| **xlBackgroundTransparent**     | 字体背景设置为透明，以便当文本颜色接近文本下方的颜色时，文本背景不会更改 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例为第一张工作表上的第一张嵌入图表添加图表标题，然后设置标题的字体大小和背景类型。本示例假定图表在第一张工作表上*/ function test() {  let objects1 = Application.Worksheets.Item(1).ChartObjects(1).Chart objects1.HasTitle = true objects1.ChartTitle.Text = "Rainfall Totals by Month"     let objects2 = objects1.ChartTitle.Font     objects2.Size = 10     objects2.Background = xlBackgroundTransparent  }` |

#### **Font.Bold**

如果字体格式为加粗，则该属性值为 **True**。**Variant** 类型，可读写。

**语法**

**express.Bold**

*express*   一个代表 **Font** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例将工作表 Sheet1 上 A1:A5 区域中的字体设为加粗*/ Application.Worksheets.Item("Sheet1").Range("A1:A5").Font.Bold = true` |

#### **Font.Color**

返回或设置对象的主要颜色，如注释部分中的表格所示。使用 **RGB** 函数可创建颜色值。**Variant** 型，可读写。

**语法**

**express.Color**

*express*   一个代表 **Font** 对象的变量。

**说明**

| 对象         | 对应颜色                                                     |
| ------------ | ------------------------------------------------------------ |
| **Border**   | 边框的颜色。                                                 |
| **Borders**  | 一个区域的所有四条边的颜色。如果四边不是同一种颜色，则 **Color** 返回的是 0（零）。 |
| **Font**     | 字体的颜色。                                                 |
| **Interior** | 单元格底纹的颜色或图形对象的填充颜色。                       |
| **Tab**      | 选项卡的颜色。                                               |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 此示例对 Chart1 中数值坐标轴的刻度线标志颜色进行设置*/ Application.Charts.Item("Chart1").Axes(xlValue).TickLabels.Font.Color = (0, 255, 0)` |

#### **Font.ColorIndex**

返回或设置一个 **Variant** 值，它代表字体的颜色。

**语法**

**express.ColorIndex**

*express*   一个代表 **Font** 对象的变量。

**说明**

颜色可指定为当前调色板中颜色的索引值，也可指定为下列 **XlColorIndex** 常量之一：

- **xlColorIndexAutomatic**
- **xlColorIndexNone**

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例将 Sheet1 上 A1 单元格的字体颜色更改为红色*/ Application.Worksheets.Item("Sheet1").Range("A1").Font.ColorIndex = 3` |

#### **Font.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Font** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **Font.FontStyle**

返回或设置字体样式。**String** 类型，可读写。

**语法**

**express.FontStyle**

*express*   一个代表 **Font** 对象的变量。

**说明**

更改该属性可能会影响其他 **Font** 属性（例如 **Bold** 和 **Italic**）。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例将 Sheet1 中 A1 单元格的字体样式设置为加粗和倾斜*/ Application.Worksheets.Item("Sheet1").Range("A1").Font.FontStyle = "Bold Italic"` |

#### **Font.Italic**

如果字体样式为倾斜，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.Italic**

*express*   一个代表 **Font** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例将 Sheet1 中 A1:A5 区域的字形设置为倾斜*/ Application.Worksheets.Item("Sheet1").Range("A1:A5").Font.Italic = true` |

#### **Font.Name**

返回或设置一个 **Variant** 值，它代表对象的名称。

**语法**

**express.Name**

*express*   一个代表 **Font** 对象的变量。

**说明**

**Font** 对象的名称为 **String**。

#### **Font.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **Font** 对象的变量。

#### **Font.Size**

返回或设置字号。可读/写 **Variant** 类型。

**语法**

**express.Size**

*express*   一个代表 **Font** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例将 Sheet1 的 A1:D10 单元格的字体大小设为 12 磅*/ function test() {     let size = Application.Worksheets.Item("Sheet1").Range("A1:D10")     size.Value2 = "Test"     size.Font.Size = 12 }` |

#### **Font.Strikethrough**

如果文字中间有一条水平删除线，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.Strikethrough**

*express*   一个代表 **Font** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例将 Sheet1 中活动单元格的字体设置为加删除线*/ function test() {     Application.Worksheets.Item("Sheet1").Activate() Application.ActiveCell.Font.Strikethrough = true }` |

#### **Font.Subscript**

如果字体格式设置为下标，则该属性值为 **True**。默认值为 **False**。**Variant** 类型，可读写。

**语法**

**express.Subscript**

*express*   一个代表 **Font** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例将 A1 单元格的第二个字符设置为下标字符*/ Application.Worksheets.Item("Sheet1").Range("A1").Characters(2, 1).Font.Subscript = true` |

#### **Font.Superscript**

如果字体格式设置为上标字符，则该属性值为 **True**；默认值为 **False**。**Variant** 类型，可读写。

**语法**

**express.Superscript**

*express*   一个代表 **Font** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例将 A1 单元格的最后一个字符设置为上标字符*/ function test() {     let n = Application.Worksheets.Item("Sheet1").Range("A1").Characters().Count     Application.Worksheets.Item("Sheet1").Range("A1").Characters(n, 1).Font.Superscript = true }` |

#### **Font.ThemeColor**

返回或设置已应用的配色方案中的主题颜色，该配色方案与指定对象相关联。可读/写 **Variant** 类型。

**语法**

**express.ThemeColor**

*express*   一个代表 **Font** 对象的变量。

**说明**

如果对象当前未应用主题颜色，试图访问其主题颜色则会引起无效请求运行时错误。

#### **Font.ThemeFont**

返回或设置与指定对象关联的应用字体方案中的主题字体。可读/写 **XlThemeFont**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/XlThemeFont%20%E6%9E%9A%E4%B8%BE.html)类型。

**语法**

**express.ThemeFont**

*express*   一个代表 **Font** 对象的变量。

#### **Font.TintAndShade**

返回或设置一个 **Single**，使颜色变深或变浅。

**语法**

**express.TintAndShade**

*express*   一个代表 **Font** 对象的变量。

**说明**

可以为 **TintAndShade** 属性输入 -1（最暗）到 1（最亮）之间的数字，零 (0) 为中间值。

如果将此属性设置为小于 -1 或大于 1 的值，则会引起运行时错误：“指定的值超出了范围”。此属性用于主题颜色和非主题颜色。

#### **Font.Underline**

返回或设置应用于字体的下划线类型。可为以下 **XlUnderlineStyle**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/XlUnderlineStyle%20%E6%9E%9A%E4%B8%BE.html)常量之一。**Variant** 类型，可读写。

**语法**

**express.Underline**

*express*   一个代表 **Font** 对象的变量。

**说明**

| **XlUnderlineStyle** 可为以下 **XlUnderlineStyle** 常量之一。 |
| ------------------------------------------------------------ |
| **xlUnderlineStyleNone**                                     |
| **xlUnderlineStyleSingle**                                   |
| **xlUnderlineStyleDouble**                                   |
| **xlUnderlineStyleSingleAccounting**                         |
| **xlUnderlineStyleDoubleAccounting**                         |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例将 Sheet1 中活动单元格的字体设置为单下划线*/ function test() {     Application.Worksheets.Item("Sheet1").Activate()     Application.ActiveCell.Font.Underline = xlUnderlineStyleSingle }` |

适用环境：web

适用平台：windows/linux