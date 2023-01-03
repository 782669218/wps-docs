**Style**



代表区域的样式说明。

**说明**

**Style** 对象包含样式的所有属性（字体、数字格式、对齐方式，等等）。有几种内置样式，包括“常规”、“货币”和“百分比”。同时对多个单元格修改单元格格式属性时，使用 **Style** 对象是快捷高效的方法。

对于 [**Workbook** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Workbook/Workbook%20.htm#jsObject_Workbook)对象，**Style** 对象是 [**Styles** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Styles/Styles%20.htm#jsObject_Styles)集合的成员。**Styles** 集合包含该工作簿的所有已定义样式。

通过更改应用于单元格的样式的属性可更改单元格的外观。但要记住，更改样式的属性将影响所有以该样式格式化了的单元格。

样式按照名称的字母顺序排序。样式编号表明指定样式在样式名排序列表中的位置。`Styles(1)` 是排序列表中的第一个样式，而 `Styles(Styles.Count)` 是最后一个。

有关创建和修改样式的详细信息，请参阅 [**Styles** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Styles/Styles%20.htm#jsObject_Styles)对象。

使用 **Style** 属性可返回一个用于 **Range** 对象的 **Style** 对象。下例将“百分比”样式应用于 Sheet1 中的单元格区域 A1:A10。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.Worksheets.Item("Sheet1").Range("A1:A10").Style = "Percent"` |

使用 **Styles**(*index*)（其中 *index* 是样式索引号或名称）可从工作簿的 **Style** 集合中返回一个 **Styles** 对象。下例通过设置活动工作簿中“常规”样式的 **Bold** 属性来更改该样式。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActiveWorkbook.Styles.Item("Normal").Font.Bold = true` |

**方法**

|                                                              | 名称       | 说明                    |
| ------------------------------------------------------------ | ---------- | ----------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete** | 删除对象。返回Variant值 |

**属性**

|                                                              | 名称                    | 说明                                                         |
| ------------------------------------------------------------ | ----------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AddIndent**           | 返回或设置一个 **Boolean** 值，它指明当单元格中文本的对齐方式为水平或垂直等距分布时，文本是否为自动缩进。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**         | 如果不使用对象识别符，则该属性返回一个 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Borders**             | 返回一个 **Borders** 集合，它代表样式或单元格区域（包括定义为条件格式一部分的区域）的边框。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BuiltIn**             | 如果样式为内置样式，则为 **True**。只读 **Boolean** 类型。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**             | 返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Font**                | 返回一个[Font](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Font/Font%20.htm#jsObject_Font) 对象，它代表指定对象的字体。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FormulaHidden**       | 返回或设置一个 **Boolean** 值，它指明在工作表处于保护状态时是否隐藏公式。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HorizontalAlignment** | 返回或设置一个 [XlHAlign ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/XlHAlign%20%E6%9E%9A%E4%B8%BE.html)值，它代表指定对象的水平对齐方式。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IncludeAlignment**    | 如果样式包含 **AddIndent**、**HorizontalAlignment**、**VerticalAlignment**、**WrapText**、**IndentLevel** 和 **Orientation** 属性，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IncludeBorder**       | 如果指定样式中包含 **Color**、**ColorIndex**、**LineStyle** 和 **Weight** 边框属性，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IncludeFont**         | 如果样式包含 **Background**、**Bold**、**Color**、**ColorIndex**、**FontStyle**、**Italic**、**Name**、**Size**、**Strikethrough**、**Subscript**、**Superscript**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Font/Font%20.htm#Font.Superscript)和 **Underline**字体属性，则此属性为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IncludeNumber**       | 如果样式中包含 **NumberFormat** 属性，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IncludePatterns**     | 如果指定样式中包含 **Color**、**ColorIndex**、**InvertIfNegative**、**Pattern**、**PatternColor** 和 **PatternColorIndex** 对象的内部属性，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IncludeProtection**   | 如果指定样式中包含 **FormulaHidden** 和 **Locked** 保护属性，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IndentLevel**         | 返回或设置一个 **Long** 值，它代表样式的缩进量。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Interior**            | 返回一个 **Interior** 对象，它代表指定对象的内部。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Locked**              | 返回或设置一个 **Boolean** 值，它指明对象是否已被锁定。      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MergeCells**          | 如果样式包含合并的单元格，则为 **True**。**Variant** 型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**                | 返回一个 **String** 值，它代表对象的名称。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NameLocal**           | 以用户语言返回或设置对象的名称。**String** 型，只读。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NumberFormat**        | 返回或设置一个 **String** 值，它代表对象的格式代码。         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NumberFormatLocal**   | 以采用用户语言字符串的形式返回或设置一个 **String** 值，它代表对象的格式代码。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Orientation**         | 返回或设置一个 **XlOrientation**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/XlOrientation%20%E6%9E%9A%E4%B8%BE.html)值，它代表文本方向。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**              | 返回指定对象的父对象。只读。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ReadingOrder**        | 返回或设置指定对象的阅读次序。可为以下常量之一：**xlRTL**（从右到左）、**xlLTR**（从左到右）或 **xlContext**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShrinkToFit**         | 返回或设置一个 **Boolean** 值，它指明文本是否可以自动收缩为适当尺寸以适应可用列宽。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Value**               | 返回一个 **String** 值，它代表指定样式的名称。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **VerticalAlignment**   | 返回或设置一个 **XlVAlign**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/XlVAlign%20%E6%9E%9A%E4%B8%BE.html)值，它代表指定对象的垂直对齐方式。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **WrapText**            | 返回或设置一个 **Boolean** 值，它指明 ET 是否为对象中的文本自动换行。 |

**成员方法**

#### **Style.Delete**

删除对象。返回Variant值

**语法**

**express.Delete()**

*express*   一个代表 **Style** 对象的变量。

**成员属性**

#### **Style.AddIndent**

返回或设置一个 **Boolean** 值，它指明当单元格中文本的对齐方式为水平或垂直等距分布时，文本是否为自动缩进。

**语法**

**express.AddIndent**

*express*   一个代表 **Style** 对象的变量。

**说明**

如果将此属性的值设为 **True**，那么在单元格中文本的对齐方式设为水平或垂直等距分布时，将自动缩进文本。

要将文本对齐方式设为等距分布，可在 [**Orientation** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Style/Style%20.htm#Style.Orientation)属性的值为 [**xlVertical** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/XlOrientation%20%E6%9E%9A%E4%B8%BE.html)时，将 [**VerticalAlignment** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Style/Style%20.htm#Style.VerticalAlignment)属性设为 **xlVAlignDistributed**；在 [**Orientation** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Style/Style%20.htm#Style.Orientation)属性的值为 [**xlHorizontal** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/XlOrientation%20%E6%9E%9A%E4%B8%BE.html)时，将 [**HorizontalAlignment** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Style/Style%20.htm#Style.HorizontalAlignment)属性设为 **xlHAlignDistributed**。

#### **Style.Application**

如果不使用对象识别符，则该属性返回一个 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象，该对象表示 ET 应用程序。如果使用对象识别符，则该属性返回一个表示指定对象（可对一个 OLE 自动操作对象使用本属性来返回该对象的应用程序）创建者的 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象。只读。

**语法**

**express.Application**

*express*   一个代表 **Style** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){   /*本示例显示一条有关创建 myObject 的应用程序的消息。*/   let myObject = Application.ActiveWorkbook   if(myObject.Application.Value == "ET"){       alert("This is an ET Application object.")   }   else{       alert("This is not an ET Application object.")   } }` |

#### **Style.Borders**

返回一个 **Borders** 集合，它代表样式或单元格区域（包括定义为条件格式一部分的区域）的边框。

**语法**

**express.Borders**

*express*   一个代表 **Style** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例将 Sheet1 中单元格 B2 的底部边框颜色设置为红色细边框。*/ function test(){      let rng = Application.Worksheets.Item("Sheet1").Range("B2").Borders.Item(xlEdgeBottom)     rng.LineStyle = xlContinuous     rng.Weight = xlThin     rng.ColorIndex = 3 								 }` |

#### **Style.BuiltIn**

如果样式为内置样式，则为 **True**。只读 **Boolean** 类型。

**语法**

**express.BuiltIn**

*express*   一个代表 **Style** 对象的变量。

#### **Style.Creator**

返回一个 32 位整数，该整数指示在其中创建此对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Style** 对象的变量。

**说明**

如果该对象是在 ET 中创建的，则此属性返回字符串 XCEL，它等同于十六进制的数字 5843454C。**Creator** 属性是为 Macintosh 上的 ET 设计的，在 Macintosh 上，每个应用程序都具有一个四字符的创建者代码。例如，ET 的创建者代码为 XCEL。

#### **Style.Font**

返回一个[Font](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Font/Font%20.htm#jsObject_Font) 对象，它代表指定对象的字体。

**语法**

**express.Font**

*express*   一个代表 **Style** 对象的变量。

#### **Style.FormulaHidden**

返回或设置一个 **Boolean** 值，它指明在工作表处于保护状态时是否隐藏公式。

**语法**

**express.FormulaHidden**

*express*   一个代表 **Style** 对象的变量。

**说明**

请不要将此属性与 **Hidden** 属性混淆。如果工作簿受保护，而工作表不受保护，将不会隐藏公式。只有在工作表受保护时，才会隐藏公式。

#### **Style.HorizontalAlignment**

返回或设置一个 [XlHAlign ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/XlHAlign%20%E6%9E%9A%E4%B8%BE.html)值，它代表指定对象的水平对齐方式。

**语法**

**express.HorizontalAlignment**

*express*   一个代表 **Style** 对象的变量。

**说明**

某些常量可能不可用，这取决于所选择或安装的语言支持（例如，美国英语）。

#### **Style.IncludeAlignment**

如果样式包含 **AddIndent**、**HorizontalAlignment**、**VerticalAlignment**、**WrapText**、**IndentLevel** 和 **Orientation** 属性，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.IncludeAlignment**

*express*   一个代表 **Style** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在 Sheet1 的 A1 单元格样式中加入对齐格式。*/ Application.Worksheets.Item("Sheet1").Range("A1").Style.IncludeAlignment = true` |

#### **Style.IncludeBorder**

如果指定样式中包含 **Color**、**ColorIndex**、**LineStyle** 和 **Weight** 边框属性，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.IncludeBorder**

*express*   一个代表 **Style** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在 Sheet1 的 A1 单元格样式中加入边框格式。*/ Application.Worksheets.Item("Sheet1").Range("A1").Style.IncludeBorder = true` |

#### **Style.IncludeFont**

如果样式包含 **Background**、**Bold**、**Color**、**ColorIndex**、**FontStyle**、**Italic**、**Name**、**Size**、**Strikethrough**、**Subscript**、**Superscript**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Font/Font%20.htm#Font.Superscript)和 **Underline**字体属性，则此属性为 **True**。**Boolean** 类型，可读写。

**语法**

**express.IncludeFont**

*express*   一个代表 **Style** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在 Sheet1 的 A1 单元格样式中加入字体格式。*/ Application.Worksheets.Item("Sheet1").Range("A1").Style.IncludeFont = true` |

#### **Style.IncludeNumber**

如果样式中包含 **NumberFormat** 属性，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.IncludeNumber**

*express*   一个代表 **Style** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在 Sheet1 的 A1 单元格样式中加入数字格式。*/ Application.Worksheets.Item("Sheet1").Range("A1").Style.IncludeNumber = true ` |

#### **Style.IncludePatterns**

如果指定样式中包含 **Color**、**ColorIndex**、**InvertIfNegative**、**Pattern**、**PatternColor** 和 **PatternColorIndex** 对象的内部属性，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.IncludePatterns**

*express*   一个代表 **Style** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在 Sheet1 的 A1 单元格样式中加入图案格式。*/ Application.Worksheets.Item("Sheet1").Range("A1").Style.IncludePatterns = true` |

#### **Style.IncludeProtection**

如果指定样式中包含 **FormulaHidden** 和 **Locked** 保护属性，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.IncludeProtection**

*express*   一个代表 **Style** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在 Sheet1 的 A1 单元格样式中加入保护格式。*/ Application.Worksheets.Item("Sheet1").Range("A1").Style.IncludeProtection = true` |

#### **Style.IndentLevel**

返回或设置一个 **Long** 值，它代表样式的缩进量。

**语法**

**express.IndentLevel**

*express*   一个代表 **Style** 对象的变量。

**说明**

使用此属性将缩进量设为小于 0 或者大于 15 的数字，将导致发生错误。

#### **Style.Interior**

返回一个 **Interior** 对象，它代表指定对象的内部。

**语法**

**express.Interior**

*express*   一个代表 **Style** 对象的变量。

#### **Style.Locked**

返回或设置一个 **Boolean** 值，它指明对象是否已被锁定。

**语法**

**express.Locked**

*express*   一个代表 **Style** 对象的变量。

**说明**

如果对象已被锁定，此属性将返回 **True**；如果在工作表处于受保护状态时仍能修改对象，则返回 **False**。

#### **Style.MergeCells**

如果样式包含合并的单元格，则为 **True**。**Variant** 型，可读写。

**语法**

**express.MergeCells**

*express*   一个代表 **Style** 对象的变量。

#### **Style.Name**

返回一个 **String** 值，它代表对象的名称。

**语法**

**express.Name**

*express*   一个代表 **Style** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*首先用宏语言，然后用用户语言显示活动工作簿中的第一种样式的名称。*/ function test(){   let sty = Application.ActiveWorkbook.Styles.Item(1)   alert("The name of the style: " + sty.Name)   alert("The localized name of the style: " + sty.NameLocal) }    /*显示活动工作簿的 Sheet1 中的默认 ListObject 对象的名称。*/ function Test(){       let wrksht = Application.ActiveWorkbook.Worksheets.Item("Sheet1")     let oListObj = wrksht.ListObjects.Item(1)										     alert(oListObj.Name)  								  }` |

#### **Style.NameLocal**

以用户语言返回或设置对象的名称。**String** 型，只读。

**语法**

**express.NameLocal**

*express*   一个代表 **Style** 对象的变量。

**说明**

如果样式为内置样式，则此属性以当前系统所使用的地区语言返回样式的名称。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){   /*此示例显示活动工作簿上样式一的原名称和本地化以后的名称。*/   let sty = Application.ActiveWorkbook.Styles.Item(1)   alert("The name of the style is " + sty.Name)   alert("The localized name of the style is " + sty.NameLocal) }` |

#### **Style.NumberFormat**

返回或设置一个 **String** 值，它代表对象的格式代码。

**语法**

**express.NumberFormat**

*express*   一个代表 **Style** 对象的变量。

**说明**

格式代码与**“设置单元格格式”**对话框中的**“格式代码”**选项是同一个字符串。**Format** 函数使用的格式代码字符串与 **NumberFormat** 和 **NumberFormatLocal**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Style/Style%20.htm#Style.NumberFormatLocal)属性使用的格式代码字符串不同。

#### **Style.NumberFormatLocal**

以采用用户语言字符串的形式返回或设置一个 **String** 值，它代表对象的格式代码。

**语法**

**express.NumberFormatLocal**

*express*   一个代表 **Style** 对象的变量。

**说明**

**Format** 函数使用的格式代码字符串与 **NumberFormat**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/Style/Style%20.htm#Style.NumberFormat)和 **NumberFormatLocal** 属性使用的格式代码字符串不同。

#### **Style.Orientation**

返回或设置一个 **XlOrientation**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/XlOrientation%20%E6%9E%9A%E4%B8%BE.html)值，它代表文本方向。

**语法**

**express.Orientation**

*express*   一个代表 **Style** 对象的变量。

#### **Style.Parent**

返回指定对象的父对象。只读。

**语法**

**express.Parent**

*express*   一个代表 **Style** 对象的变量。

#### **Style.ReadingOrder**

返回或设置指定对象的阅读次序。可为以下常量之一：**xlRTL**（从右到左）、**xlLTR**（从左到右）或 **xlContext**。**Long** 类型，可读写。

**语法**

**express.ReadingOrder**

*express*   一个代表 **Style** 对象的变量。

#### **Style.ShrinkToFit**

返回或设置一个 **Boolean** 值，它指明文本是否可以自动收缩为适当尺寸以适应可用列宽。

**语法**

**express.ShrinkToFit**

*express*   一个代表 **Style** 对象的变量。

#### **Style.Value**

返回一个 **String** 值，它代表指定样式的名称。

**语法**

**express.Value**

*express*   一个代表 **Style** 对象的变量。

#### **Style.VerticalAlignment**

返回或设置一个 **XlVAlign**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E8%A1%A8%E6%A0%BC%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/XlVAlign%20%E6%9E%9A%E4%B8%BE.html)值，它代表指定对象的垂直对齐方式。

**语法**

**express.VerticalAlignment**

*express*   一个代表 **Style** 对象的变量。

**说明**

某些常量可能不可用，这取决于所选择或安装的语言支持（例如，美国英语）。

#### **Style.WrapText**

返回或设置一个 **Boolean** 值，它指明 ET 是否为对象中的文本自动换行。

**语法**

**express.WrapText**

*express*   一个代表 **Style** 对象的变量。

适用环境：web

适用平台：windows/linux