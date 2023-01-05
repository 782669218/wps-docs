#### **Footnotes**



以后的版本中将提供关于此成员的说明。

**方法**

|                                                              | 名称                           | 说明                                                         |
| ------------------------------------------------------------ | ------------------------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**                        | 返回一个 **Footnote** 对象，该对象代表添加到区域中的脚注。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Convert**                    | 将尾注转换为脚注，或将脚注转换为尾注。                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**                       | 返回集合中的单个 **Footnote** 对象。                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ResetContinuationNotice**    | 将脚注或尾注延续标记重新设置为默认标记。                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ResetContinuationSeparator** | 将脚注或尾注延续分隔符重新设置为默认分隔符。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ResetSeparator**             | 将脚注分隔符重新设置为默认分隔符。                           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SwapWithEndnotes**           | 将一篇文档中的所有尾注转换成脚注或将脚注转换为尾注。要将某一区域的尾注转换为脚注，请使用 **Convert** 方法。 |

**属性**

|                                                              | 名称         | 说明                       |
| ------------------------------------------------------------ | ------------ | -------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**    | 返回指定集合中的项目数。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Location** | 返回或设置所有脚注的位置。 |

**成员方法**

#### **Footnotes.Add**

返回一个 **Footnote** 对象，该对象代表添加到区域中的脚注。

**语法**

**express.Add(Range, Reference, Text)**

*express*   一个代表 **Footnotes** 对象的变量。

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
| `/*以下代码示例在选定内容的末尾添加一条自动编号的脚注。*/ Application.ActiveDocument.Footnotes.Add(Selection.Range,"The Willow Tree, (Lone Creek Press, 1996).")  /*以下代码示例为引用标记添加一条使用自定义符号的脚注。*/ Application.ActiveDocument.Footnotes.Add(Selection.Range ,"More information in the full report.","{Symbol -3998}")` |

#### **Footnotes.Convert**

将尾注转换为脚注，或将脚注转换为尾注。

**语法**

**express.Convert()**

*express*   一个代表 **Footnotes** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将选定内容的脚注转换为尾注。*/ function test() { if( Selection.Footnotes.Count > 0 ) {     Selection.Footnotes.Convert() } }` |

#### **Footnotes.Item**

返回集合中的单个 **Footnote** 对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **Footnotes** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Index*  | 必选          | **Long**     | 要返回的单个对象。可以是代表单个对象序号位置的 Long 类型值。 |

#### **Footnotes.ResetContinuationNotice**

将脚注或尾注延续标记重新设置为默认标记。

**语法**

**express.ResetContinuationNotice()**

*express*   一个代表 **Footnotes** 对象的变量。

**说明**

默认标记为空（无文本）。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在 Sales.doc 中重新设置脚注延续标记并将脚注引用标记的起始编号设置为 2。*/ function test() { let note= Documents.Item("Sales.doc").Sections.Item(1).Range.Footnotes note.ResetContinuationNotice() note.NumberingRule = wdRestartContinuous note.StartingNumber = 2 }` |

#### **Footnotes.ResetContinuationSeparator**

将脚注或尾注延续分隔符重新设置为默认分隔符。

**语法**

**express.ResetContinuationSeparator()**

*express*   一个代表 **Footnotes** 对象的变量。

**说明**

默认分隔符是一条长横线，该横线分隔文档文字与上接前一页的注释。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将脚注延续分隔符重新设置为默认分隔线。*/ ActiveDocument.Footnotes.ResetContinuationSeparator()` |

#### **Footnotes.ResetSeparator**

将脚注分隔符重新设置为默认分隔符。

**语法**

**express.ResetSeparator()**

*express*   一个代表 **Footnotes** 对象的变量。

**说明**

默认分隔符为一条短横线，该横线分隔文档文字和注释。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将脚注分隔符重新设置为默认分隔线。*/ ActiveDocument.Footnotes.ResetSeparator()` |

#### **Footnotes.SwapWithEndnotes**

将一篇文档中的所有尾注转换成脚注或将脚注转换为尾注。要将某一区域的尾注转换为脚注，请使用 **Convert** 方法。

**语法**

**express.SwapWithEndnotes()**

*express*   一个代表 **Footnotes** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档中的脚注转换为尾注，将尾注转换为脚注。*/ ActiveDocument.Footnotes.SwapWithEndnotes()` |

**成员属性**

#### **Footnotes.Count**

返回指定集合中的项目数。

**语法**

**express.Count**

*express*   一个代表 **Footnotes** 对象的变量。

#### **Footnotes.Location**

返回或设置所有脚注的位置。

**语法**

**express.Location**

*express*   一个代表 **Footnotes** 对象的变量。

适用环境：web

适用平台：windows/linux