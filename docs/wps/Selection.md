#### **Selection**



代表窗口或窗格中的当前所选内容。所选内容代表文档中选定（或突出显示）的区域，如果文档中没有选定任何内容，则代表插入点。每个文档窗格只能有一个 **Selection** 对象，并且在整个应用程序中只能有一个活动的 **Selection** 对象。

**说明**

可以使用 **Selection** 属性返回 **Selection** 对象。如果 **Selection** 属性未使用对象限定符，则 WPS 返回活动文档窗口的活动窗格中的选定内容。以下示例从活动文档中复制当前选定内容。

| 示例代码复制                   |
| ------------------------------ |
| `Application.Selection.Copy()` |

以下示例删除 **Documents** 集合中第三个文档的所选内容。访问该文档的当前所选内容时，该文档无需处于活动状态。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.Documents.Item(3).ActiveWindow.Selection.Cut()` |

以下示例复制活动文档第一个窗格中的所选内容，并将其粘贴到第二个窗格中。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     Application.ActiveDocument.ActiveWindow.Panes.Item(1).Selection.Copy()     Application.ActiveDocument.ActiveWindow.Panes.Item(2).Selection.Paste() }` |

**Text** 属性是 **Selection** 对象的默认属性。使用此属性可设置或返回当前所选内容中的文本。

| 示例代码复制                 |
| ---------------------------- |
| `Application.Selection.Text` |

**方法**

|                                                              | 名称                               | 说明                                                         |
| ------------------------------------------------------------ | ---------------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **BoldRun**                        | 在当前局部添加粗体字符格式或删除该格式                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Calculate**                      | 计算选定内容中的数学表达式。返回的结果为 **Single** 类型。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ClearCharacterAllFormatting**    | 从所选文本中删除所有字符格式（包括通过字符样式应用的格式和手动应用的格式）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ClearCharacterDirectFormatting** | 取消所选文本的字符格式，即通过功能区上的按钮或对话框手动应用的格式。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ClearCharacterStyle**            | 从所选文本中删除通过字符样式应用的字符格式。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ClearFormatting**                | 清除所选内容的文本格式和段落格式。                           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ClearParagraphAllFormatting**    | 从所选文本中删除所有段落格式（包括通过段落样式应用的格式和手动应用的格式）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ClearParagraphDirectFormatting** | 从所选文本中删除通过功能区上的按钮或者对话框手动应用的段落格式。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ClearParagraphStyle**            | 从所选文本中删除通过段落样式应用的段落格式。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Collapse**                       | 将选定内容折叠到起始位置或结束位置。折叠之后起始位置和结束位置相同。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ConvertToTable**                 | 将范围内的文本转换为表格。将该表格作为一个 **Table** 对象返回。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Copy**                           | 将指定的选定内容复制到剪贴板。                               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CopyAsPicture**                  | CopyAsPicture 方法与 Copy 方法的工作方式相同。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CopyFormat**                     | 复制选定文字第一个字符的字符格式。                           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CreateAutoTextEntry**            | 根据当前选定内容，将一个新的 AutoTextEntry 对象添加到 AutoTextEntries 集合。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CreateTextbox**                  | 在选定内容周围添加一个默认大小的文本框。                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Cut**                            | 从文档中删除指定对象，并将其移动到剪贴板上。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**                         | 删除指定数量的字符或单词。                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **DetectLanguage**                 | 分析指定文本，以确定书写文本的语言类型。                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Endkey**                         | 将选定内容移动或扩展到指定单位的末尾。                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **EndOf**                          | 将区域或选定内容的结束字符位置移动或扩展至最近的一个指定文本单元末尾。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **EscapeKey**                      | 取消某种模式，如取消扩展或列选定模式（与按 ESC 相同）。      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Expand**                         | 扩展指定区域或选定内容。返回添至该区域或选定内容的字符数。**Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ExportAsFixedFormat**            | 将当前选定内容另存为 PDF 或 XPS 格式。                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Extend**                         | 打开扩展模式。如果扩展模式已经打开，则将选定内容扩展到下一个更大的文本单位。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Goto**                           | 将插入点移至紧靠指定项之前的字符位置，并返回一个 **Range** 对象（除 **wdGoToGrammaticalError**、**wdGoToProofreadingError** 或 **wdGoToSpellingError** 常量之外）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **GoToEditableRange**              | 返回一个 **Range** 对象，该对象代表指定用户或用户组可以修改的文档区域。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **GoToNext**                       | 返回一个 **Range** 对象，该对象参考下一项目的起始位置或者由 *What* 参数确定的位置。如果将该方法应用于 **Selection** 对象，则该方法将选定内容移动到指定的项目处（但 **wdGoToGrammaticalError**、**wdGoToProofreadingError** 和 **wdGoToSpellingError** 常量除外）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **GoToPrevious**                   | 返回一个 **Range** 对象，该对象参考前一项目的起始位置或者由参数 *What* 指定的位置。如果对 **Selection** 对象采用此方法，**GoToPrevious** 将选定内容移动到所指定的项目处。**Range** 对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Homekey**                        | 将选定内容移动或扩展至指定单位的开始。该方法返回表明选定内容实际移动的字符数的整数；如果移动不成功，则返回 0（零）。该方法相当于 **Home** 键的功能。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **InRange**                        | 如果应用此方法的所选内容包含在 *Range* 参数所指定的范围内，则该方法为 **True**。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **InsertAfter**                    | 将指定文本插入范围或所选内容的末尾。                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **InsertBefore**                   | 在指定选定内容之前插入指定文本。                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **InsertBreak**                    | 插入分页符、分栏符或分节符。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **InsertCaption**                  | 紧靠在指定的所选内容之前或之后插入题注。                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **InsertCells**                    | 向原有表添加单元格。                                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **InsertColumns**                  | 在选定内容的左侧插入新列。                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **InsertColumnsRight**             | 在当前选定内容的右边插入列。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **InsertCrossReference**           | 插入对标题、书签、脚注、尾注或定义了题注标签的项（如公式、图表或表格）的交叉引用。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **InsertDateTime**                 | 以文本或 TIME 域的形式插入当前日期或时间，或将两者都插入。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **InsertFile**                     | 插入指定文件的全部或一部分。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **InsertFormula**                  | 插入包含选定内容的公式的 = (Formula) 域。                    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **InsertNewPage**                  | 在插入点的位置插入一个新页面。                               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **InsertParagraph**                | 用新段落替换指定的所选内容。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **InsertParagraphAfter**           | 在所选内容之后插入段落标记。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **InsertParagraphBefore**          | 在指定的所选内容或范围前插入一个新段落。                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **InsertRows**                     | 在所选内容的上方插入指定数量的新行。如果选定内容不在表格中，则会导致出错。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **InsertRowsAbove**                | 在当前选定内容上方插入行。                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **InsertRowsBelow**                | 在当前选定内容的下方插入行。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **InsertStyleSeparator**           | 插入特殊隐藏段落标记，使 WPS 可以使用不同的段落样式合并段落格式，以便在目录中插入内置标题。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **InsertSymbol**                   | 插入一个符号，用来替换指定的所选内容。                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **InsertXML**                      | 将指定的 XML 插入文档中的光标处，并且替换任何选定文本。      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **InStory**                        | 如果应用此方法的选择范围与 *Range* 参数指定的范围位于相同的文字部分中，则该方法为 **True**。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **IsEqual**                        | 如果应用此方法的选择范围等于 *Range* 参数所指定的范围，则该方法为 **True**。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ItalicRun**                      | 在当前局部中添加或删除斜体字符格式。                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **LtrPara**                        | 将指定段落的对齐方式和阅读顺序设置为从左向右。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **LtrRun**                         | 用于将指定局部的阅读顺序和对齐方式设置为从左向右。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Move**                           | 将指定的所选内容折叠到其起始位置或结束位置，然后将折叠的对象移动指定的单位数。此方法返回一个 **Long** 类型的值，该值代表所选内容移动的单位数；如果移动失败，则返回 0（零）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **MoveDown**                       | 将选定内容向下移动，并返回移动距离的单位数。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **MoveEnd**                        | 移动范围或所选内容的结束字符位置。                           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **MoveEndUntil**                   | 移动指定的所选内容的结束位置，直到在文档中找到任何指定的字符。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **MoveEndWhile**                   | 当在文档中找到任何指定的字符时，移动所选内容的结束字符位置。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **MoveLeft**                       | 将选定内容向左移动，并返回移动距离的单位数。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **MoveRight**                      | 将选定内容向右移动，并返回移动距离的单位数。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **MoveStart**                      | 移动指定的所选内容的起始位置。                               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **MoveStartUntil**                 | 移动指定的所选内容的起始位置，直到在文档中找到一个指定的字符。如果是在文档中向后移动，则扩展所选内容。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **MoveStartWhile**                 | 当在文档中找到任何指定的字符时，移动指定的所选内容的起始位置。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **MoveUntil**                      | 移动指定的所选内容，直到在文档中找到一个指定的字符。         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **MoveUp**                         | 将所选内容向上移动，并返回移动的单位数。                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **MoveWhile**                      | 当在文档中找到任何指定的字符时，移动指定的选择范围。         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Next**                           | 返回一个 **Range** 对象，该对象代表相对于指定的选择范围的下一个单位。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **NextField**                      | 选定下一个域。                                               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **NextRevision**                   | 找到下一处修订并作为一个 **Revision** 对象返回。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **NextSubdocument**                | 将所选内容移至下一个子文档。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Paste**                          | 将“剪贴板”的内容插入指定的选择范围处。                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PasteAndFormat**                 | 粘贴选定的表格单元格，并为其设置指定的格式。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PasteAppendTable**               | 通过在所选行之间插入粘贴的行，将粘贴的单元格合并到现有的表格中。不覆盖任何单元格。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PasteAsNestedTable**             | 将一个或一组单元格作为嵌套表格粘贴到所选内容中。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PasteExcelTable**                | 粘贴 ET 表格并设置其格式。                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PasteFormat**                    | 将以 **CopyFormat** 方法复制的格式应用于选定内容。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PasteSpecial**                   | 插入“剪贴板”中的内容。                                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Previous**                       | 将所选文本移动指定的单位数，并返回与折叠的所选内容相关的 **Range** 对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PreviousField**                  | 选择并返回前一域。                                           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PreviousRevision**               | 找到前一处修订并作为 **Revision** 对象返回。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PreviousSubdocument**            | 将所选内容移至上一个子文档。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ReadingModeGrowFont**            | 在以阅读模式显示文档时，将所显示文本的大小增大一磅。         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ReadingModeShrinkFont**          | 在以阅读模式显示文档时，将所显示文本的大小减小一磅。         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **RtlPara**                        | 将指定段落的阅读顺序和对齐方式设置为从右向左。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **RtlRun**                         | 将指定的局部的阅读顺序和对齐方式设置为从右向左。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Select**                         | 选择指定的文本。                                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SelectCell**                     | 选择包含当前选定内容的匹配单元格。                           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SelectColumn**                   | 选定包含插入点的列，或者选定包含选定内容的所有列。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SelectCurrentAlignment**         | 向前扩展选定部分，直到遇到另一种段落对齐方式为止。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SelectCurrentColor**             | 向前扩展选定内容，直至遇到另一种颜色的文字为止。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SelectCurrentFont**              | 向前扩展选定内容，直至遇到另一种字体或字号。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SelectCurrentIndent**            | 向前扩展选定内容，直至遇到具有另一种段落左右缩进量的文本为止。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SelectCurrentSpacing**           | 向前扩展选定内容，直至遇到具有另一种行间距的段落为止。       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SelectCurrentTabs**              | 向前扩展选定内容，直至遇到另一种制表位的段落为止。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SelectRow**                      | 选定插入点所在的行，或者选择选定内容所在行。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SetRange**                       | 设置所选内容的起始字符和结束字符的位置。                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Shrink**                         | 将所选内容缩减至下一级较小的文字单位。                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ShrinkDiscontiguousSelection**   | 当所选内容包括多个不连续的所选内容时，取消全部选择，只保留最近选择的文本。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Sort**                           | 对指定的所选内容中的段落进行排序。                           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SortAscending**                  | 按字母数字升序对段落或表格行进行排序。                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SortDescending**                 | 以字母数字降序排列所选内容中的段落或表格行。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SplitTable**                     | 在选定内容第一行上面插入一个空段落。                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **StartOf**                        | 将指定的区域或选定内容的开始位置移动或扩展至最近的指定文字单位的开头。该方法返回 **Long** 类型的值表明了区域或选定内容移动或扩展的字符数。如果是在文档中向后移动，则该方法返回负数 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ToggleCharacterCode**            | 在 Unicode 字符和其相应的十六进制值之间切换选定内容。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **TypeBackspace**                  | 删除折叠的选定内容（即一个插入点）前面的字符                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **TypeParagraph**                  | 插入一个新的空段落。                                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **TypeText**                       | 插入指定的文本。                                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **WholeStory**                     | 扩展某一所选内容，使其包括整个文章。                         |

**属性**

|                                                              | 名称                   | 说明                                                         |
| ------------------------------------------------------------ | ---------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Active**             | 如果指定窗口或窗格中的选定内容处于活动状态，则该属性值为 **True**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BookmarkID**         | 返回位于选定内容开始位置的书签编号。**Long** 类型，只读。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Bookmarks**          | 返回一个 **Bookmarks** 集合。该集合代表某一文档、区域或选定内容中的所有书签。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Borders**            | 返回一个 **Borders** 集合，该集合代表指定对象的所有边框。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Cells**              | 返回一个 **Cells** 集合，该集合代表选定内容中的表格单元格。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Characters**         | 返回一个 **Characters** 集合，该集合代表文档、区域或选定内容中的字符。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ChildShapeRange**    | 返回一个 **ShapeRange** 集合，该集合代表选定区域中包含的子图形。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Columns**            | 返回一个 **Columns** 集合，该集合代表所选内容中的所有表格列。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ColumnSelectMode**   | 如果列选定模式处于活动状态，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Comments**           | 返回一个 **Comments** 集合，该集合代表指定批注中的所有批注。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**            | 回一个 32 位整数，该整数指出用于创建指定对象的应用程序。**Long** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Document**           | 返回一个 **Document** 对象，该对象与指定的选定内容相关。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Editors**            | 回一个 **Editors** 对象，该对象代表已授权修改文档中选定内容的所有用户。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **End**                | 返回或设置选定内容的结束字符的位置。**Long** 类型，可读写。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EndnoteOptions**     | 返回一个 **EndnoteOptions** 对象，该对象代表选定内容中的尾注。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Endnotes**           | 返回一个 **Endnotes** 集合，该集合代表选定内容中包含的所有尾注。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EnhMetaFileBits**    | 返回一个 **Variant** 类型的值，该值代表选定文本或文本区域的显示方式的图片表示形式。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ExtendMode**         | 如果“扩展”模式处于活动状态，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Fields**             | 返回一个只读 **Fields** 集合，该集合代表选定内容中的所有域。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Find**               | 返回一个 **Find** 对象，该对象包含查找操作所需的条件。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FitTextWidth**       | 返回或设置 WPS 在当前选定内容中填入文字的宽度（使用当前度量单位）。**Single** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Flags**              | 返回或设置选定内容的属性。**WdSelectionFlags** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Font**               | 返回或设置一个 **Font** 对象，该对象代表指定对象的字符格式。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FootnoteOptions**    | 返回一个 **FootnoteOptions** 对象，该对象代表选定内容中的脚注。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Footnotes**          | 返回一个 **Footnotes** 集合，该集合代表区域、选定内容或文档中的所有脚注。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FormattedText**      | 返回或设置一个 **Range** 对象，该对象包含指定区域或选定内容中进行过格式编排的文字。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FormFields**         | 返回一个 **FormFields** 集合，该集合代表选定内容中所有窗体域。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Frames**             | 返回一个 **Frames** 集合，该集合代表选定内容中的所有框架。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HasChildShapeRange** | 如果选定内容包含子图形，则该属性值为 **True**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HeaderFooter**       | 为指定的选定内容返回 **HeaderFooter** 对象。只读。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HTMLDivisions**      | 返回一个 **HTMLDivisions** 对象，该对象代表 Web 文档中的一个 HTML 区域。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Hyperlinks**         | 返回一个 **Hyperlinks** 集合，该集合代表指定选定内容中的所有超链接。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Information**        | 返回有关指定的选定内容的信息。**Variant** 类型，只读。       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **InlineShapes**       | 返回一个 **InlineShapes** 集合，该集合代表选定内容中的所有 **InlineShape** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IPAtEndOfLine**      | 如果插入点位于行（该行折到了下一行）的末尾，则该属性值为 **True**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IsEndOfRowMark**     | 如果指定的选定内容或区域折叠且位于表格行的结束标志处，则该属性值为 **True**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LanguageDetected**   | 返回或设置一个指定 WPS 是否已经检测过选定文本的语言的 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LanguageID**         | 返回或设置指定对象的语言。可读写。                           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LanguageIDFarEast**  | 为指定的对象返回或设置东亚语言。**WdLanguageID** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LanguageIDOther**    | 返回或设置指定对象的语言。**WdLanguageID** 类型，可读写。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NoProofing**         | 如果拼写和语法检查程序忽略指定文本，则该属性值为 **True**。如果只将某些指定文本的 **NoProofing** 属性设为 **True**，则会返回 **wdUndefined**。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OMaths**             | 返回一个 **OMaths** 集合，该集合代表当前选定区域内的 **OMath** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Orientation**        | 在启用“文字方向”功能后返回或设置选定内容中文字的方向。**WdTextOrientation** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PageSetup**          | 返回一个 **PageSetup** 对象，该对象与指定选定内容相关联。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ParagraphFormat**    | 返回或设置一个 **ParagraphFormat** 对象，该对象代表指定选定内容的段落设置。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Paragraphs**         | 返回一个 **Paragraphs** 集合，该集合代表指定选定内容中的所有段落。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**             | 返回一个 **Object** 类型的值，该值代表指定 **Selection** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PreviousBookmarkID** | 返回位于指定的所选内容或区域中或之前的最后一个书签的编号。如果没有相应的书签，则返回 0（零）。**Long** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Range**              | 返回一个 **Range** 对象，该对象代表指定对象所含的部分文档。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Rows**               | 返回一个 **Rows** 集合，该集合代表某个区域、选定内容或表格中所有的表格行。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Sections**           | 返回一个 **Sections** 集合，该集合代表指定选定内容中的各节。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Sentences**          | 返回一个 **Sentences** 集合，该集合代表选定内容中的所有句子。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Shading**            | 返回一个 **Shading** 对象，该对象代表指定选定内容的底纹格式。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Start**              | 返回或设置选定内容的起始字符位置。**Long** 类型，可读写。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **StartIsActive**      | 如果选定内容的开始部分处于活动状态，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **StoryLength**        | 返回包含指定选定内容的文章所含的字符数。**Long** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **StoryType**          | 返回指定选定内容的文章类型。**WdStoryType** 类型，只读。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Style**              | 该属性返回或设置用于指定对象的样式。若要设置该属性，请指定样式的本地名称、一个整数值、一个 **WdBuiltinStyle** 常量或一个代表该样式的对象。若要获取有效常量的列表，请查询“Visual Basic 对象浏览器”。**Variant** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Tables**             | 返回一个 **Tables** 集合，该集合代表指定选定内容中的所有表格。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Text**               | 返回或设置指定的选定内容中的文本。**String** 类型，可读写。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TopLevelTables**     | 回一个 **Tables** 集合，该集合代表当前选定内容最外部嵌套层上的表格。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**               | 返回选择类型。**WdSelectionType** 类型，只读。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **WordOpenXML**        | 返回一个 **String** 类型的值，该值以 WPS Open XML 格式表示选定内容中包含的 XML。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Words**              | 返回一个 **Words** 集合，该集合代表选定内容中的所有字词。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **XML**                | 返回一个 **String** 类型的值，该值代表指定对象中的 XML 文本。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **XMLNodes**           | 返回一个 **XMLNodes** 集合，该集合代表选定内容中所有 XML 元素的集合，包括那些仅部分位于选定内容中的元素。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **XMLParentNode**      | 返回一个 **XMLNode** 对象，该对象代表选定内容的父节点。      |

**成员方法**

#### **Selection.BoldRun**

在当前局部添加粗体字符格式或删除该格式

**语法**

**express.BoldRun()**

*express*   一个代表 **Selection** 对象的变量。

**说明**

如果该局部含有加粗和非加粗两种文字，该方法将为整个局部添加粗体格式。

#### **Selection.Calculate**

计算选定内容中的数学表达式。返回的结果为 **Single** 类型。

**语法**

**express.Calculate()**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.ClearCharacterAllFormatting**

从所选文本中删除所有字符格式（包括通过字符样式应用的格式和手动应用的格式）。

**语法**

**express.ClearCharacterAllFormatting()**

*express*   一个代表 **Selection** 对象的变量。

**说明**

此方法会删除所有字符格式。如果您需要删除通过字符样式应用的格式，请使用 ClearCharacterStyle 方法。要删除用户使用 WPS 字符格式设置功能手动应用的字符格式，请使用 ClearCharacterDirectFormatting 方法。

#### **Selection.ClearCharacterDirectFormatting**

取消所选文本的字符格式，即通过功能区上的按钮或对话框手动应用的格式。

**语法**

**express.ClearCharacterDirectFormatting()**

*express*   一个代表 **Selection** 对象的变量。

**说明**

此方法不会删除通过字符样式应用的字符格式。要删除用户通过字符样式应用的字符格式，请使用 ClearCharacterStyle 方法。要删除所有字符格式，无论这些格式是用户通过字符样式应用的，还是用户通过 WPS 中的格式设置功能手动应用的，都可以使用 ClearCharacterAllFormatting 方法。

#### **Selection.ClearCharacterStyle**

从所选文本中删除通过字符样式应用的字符格式。

**语法**

**express.ClearCharacterStyle()**

*express*   一个代表 **Selection** 对象的变量。

**说明**

此方法不会删除用户手动应用的字符格式。 要删除手动应用的字符格式，请使用 ClearCharacterDirectFormatting 方法。要删除所有字符格式（包括样式和手动格式），请使用 ClearCharacterAllFormatting 方法。

#### **Selection.ClearFormatting**

清除所选内容的文本格式和段落格式。

**语法**

**express.ClearFormatting()**

*express*   一个代表 **Selection** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例清除活动文档中的所有文本格式和段落格式。 function test() {     Application.ActiveDocument.Select     Application.Selection.ClearFormatting }` |

#### **Selection.ClearParagraphAllFormatting**

从所选文本中删除所有段落格式（包括通过段落样式应用的格式和手动应用的格式）。

**语法**

**express.ClearParagraphAllFormatting()**

*express*   一个代表 **Selection** 对象的变量。

**说明**

此方法将删除所有段落格式。如果您需要删除通过段落样式应用的段落格式，请使用 ClearParagraphStyle 方法。要删除用户使用 WPS 段落格式设置功能手动应用的段落格式，请使用 ClearParagraphDirectFormatting 方法。

#### **Selection.ClearParagraphDirectFormatting**

从所选文本中删除通过功能区上的按钮或者对话框手动应用的段落格式。

**语法**

**express.ClearParagraphDirectFormatting()**

*express*   一个代表 **Selection** 对象的变量。

**说明**

此方法不会删除使用段落样式应用的段落格式。 要删除用户使用段落样式应用的段落格式，请使用 ClearParagraphStyle 方法。要删除所有段落格式（包括样式和手动格式），请使用 ClearParagraphAllFormatting 方法。

#### **Selection.ClearParagraphStyle**

从所选文本中删除通过段落样式应用的段落格式。

**语法**

**express.ClearParagraphStyle()**

*express*   一个代表 **Selection** 对象的变量。

**说明**

此方法不会删除用户手动应用的段落格式。 要删除手动应用的段落格式，请使用 ClearParagraphDirectFormatting 方法。要删除所有段落格式（包括样式和手动格式），请使用 ClearParagraphAllFormatting 方法。

#### **Selection.Collapse**

将选定内容折叠到起始位置或结束位置。折叠之后起始位置和结束位置相同。

**语法**

**express.Collapse(Direction)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Direction* | 可选          | **Variant**  | 折叠某区域或选定内容的方向。可以是下列 **WdCollapseDirection** 常量之一：**wdCollapseEnd** 或 **wdCollapseStart**。默认值为 **wdCollapseStart**。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将选定内容折叠为选定内容的开头。*/ Application.Selection.Collapse(wdCollapseStart)` |

#### **Selection.ConvertToTable**

将范围内的文本转换为表格。将该表格作为一个 **Table** 对象返回。

**语法**

**express.ConvertToTable(Separator, NumRows, NumColumns, InitialColumnWidth, Format, ApplyBorders, ApplyShading, ApplyFont, ApplyColor, ApplyHeadingRows, ApplyLastRow, ApplyFirstColumn, ApplyLastColumn, AutoFit, AutoFitBehavior, DefaultTableBehavior)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称**               | **必选/可选** | **数据类型** | **说明**                                                     |
| ---------------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Separator*            | 可选          | **Variant**  | 指定用于将文本分隔为多个单元格的字符。可以是一个字符，也可以是下列 WdTableFieldSeparator 常量之一。如果省略此参数，则使用 DefaultTableSeparator 属性的值。 |
| *NumRows*              | 可选          | **Variant**  | 表格中的行数。如果省略此参数，WPS 将根据该范围的内容设置行数。 |
| *NumColumns*           | 可选          | **Variant**  | 表格中的列数。如果省略此参数， WPS 将根据该范围的内容设置列数。 |
| *InitialColumnWidth*   | 可选          | **Variant**  | 每一列的初始宽度，以磅为单位。如果省略此参数， WPS 将计算并调整列宽，以使表格填满页面。 |
| *Format*               | 可选          | **Variant**  | 指定“表格自动套用格式”对话框中列出的一种预定义格式。可以是 WdTableFormat 常量之一。 |
| *ApplyBorders*         | 可选          | **Variant**  | 如果该参数值为 True，则应用指定格式的边框属性。              |
| *ApplyShading*         | 可选          | **Variant**  | 如果该参数值为 True，则应用指定格式的底纹属性。              |
| *ApplyFont*            | 可选          | **Variant**  | 如果该参数值为 True，则应用指定格式的字体属性。              |
| *ApplyColor*           | 可选          | **Variant**  | 如果该参数值为 True，则应用指定格式的颜色属性。              |
| *ApplyHeadingRows*     | 可选          | **Variant**  | 如果该参数值为 True，则应用指定格式的标题行属性。            |
| *ApplyLastRow*         | 可选          | **Variant**  | 如果该参数值为 True，则应用指定格式的最后一行的属性。        |
| *ApplyFirstColumn*     | 可选          | **Variant**  | 如果该参数值为 True，则应用指定格式的第一列的属性。          |
| *ApplyLastColumn*      | 可选          | **Variant**  | 如果该参数值为 True，则应用指定格式的最后一列的属性。        |
| *AutoFit*              | 可选          | **Variant**  | 如果该参数值为 True，则在不改变单元格内文字的换行方式的前提下尽可能缩小表格列宽。 |
| *AutoFitBehavior*      | 可选          | **Variant**  | 设置 WPS 调整表格大小的“自动调整”规则。可以是下列 WdAutoFitBehavior 常量之一。如果 DefaultTableBehavior 为 wdWord8TableBehavior，则忽略此参数。 |
| *DefaultTableBehavior* | 可选          | **Variant**  | 设置用于指定 WPS 是否会自动调整表格中单元格的尺寸以适应内容（自动调整）的值。可以是 WdDefaultTableBehavior 常量之一。 |

**返回值**

Table

#### **Selection.Copy**

将指定的选定内容复制到剪贴板。

**语法**

**express.Copy()**

*express*   一个代表 **Selection** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将选定内容复制到粘贴板上。*/ Application.Selection.Copy()` |

#### **Selection.CopyAsPicture**

CopyAsPicture 方法与 Copy 方法的工作方式相同。

**语法**

**express.CopyAsPicture()**

*express*   一个代表 **Selection** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//该示例将活动文档的内容复制为图片，并将其作为图片粘贴到文档的结尾。 function test(){     Application.ActiveDocument.Content.Select     Application.Selection.CopyAsPicture     Application.Collapse(Application.Enum.wdCollapseEnd)     Application.PasteSpecial(Application.Enum.wdPasteMetafilePicture) }` |

#### **Selection.CopyFormat**

复制选定文字第一个字符的字符格式。

**语法**

**express.CopyFormat()**

*express*   一个代表 **Selection** 对象的变量。

**说明**

如果选定了一个段落标记，则 WPS 除了复制字符格式，还复制段落格式。可以利用 [**PasteFormat** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Selection/Selection%20.htm#Selection.Pasteformat)方法，将复制的格式应用到另一个选择内容。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档第一段的格式复制到第二段。*/ function test() { 	Application.ActiveDocument.Paragraphs.Item(1).Range.Select() 	Application.Selection.CopyFormat() 	Application.ActiveDocument.Paragraphs.Item(2).Range.Select() 	Application.Selection.PasteFormat() }  /*本示例折叠选定内容，并将选定内容的字符格式复制到下一个单词。*/ function test() { 	let sel = Application.Selection; 	sel.Collapse(wdCollapseStart); 	sel.CopyFormat(); 	sel.Next(wdWord, 1).Select() 	sel.PasteFormat() } ` |

#### **Selection.CreateAutoTextEntry**

根据当前选定内容，将一个新的 AutoTextEntry 对象添加到 AutoTextEntries 集合。

**语法**

**express.CreateAutoTextEntry(Name, StyleName)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                   |
| ----------- | ------------- | ------------ | ---------------------------------------------------------- |
| *Name*      | 必选          | **String**   | 该文本必须由用户键入，以调用新的“自动图文集”词条。         |
| *StyleName* | 必选          | **String**   | 在该类别下，新的“自动图文集”词条将列在“自动图文集”菜单上。 |

**说明**

------

#### **Selection.CreateTextbox**

在选定内容周围添加一个默认大小的文本框。

**语法**

**express.CreateTextbox()**

*express*   一个代表 **Selection** 对象的变量。

**说明**

如果所选内容是一个插入点，本方法将鼠标指针设为十字型指针，以便绘制一个文本框。

该方法与单击**“绘图”**工具栏上的**“文本框”**按钮等效。文本框是一个具有关联文本框架的矩形。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例在选定内容周围添加一个文本框，然后改变该文本框样式。 function test(){     if(Application.Selection.Type == wdSelectionNormal){         Application.Selection.CreateTextbox()         Application.Selection.ShapeRange(1).Line.DashStyle = msoLineDashDot     }  }` |

#### **Selection.Cut**

从文档中删除指定对象，并将其移动到剪贴板上。

**语法**

**express.Cut()**

*express*   一个代表 **Selection** 对象的变量。

**说明**

将所选内容移动到剪贴板上，但折叠的所选内容仍保留在文档中。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例删除所选内容，并将其粘贴到新文档中。*/ function test() { 	let sel = Application.Selection; 	if(sel.Type == wdSelectionNormal) { 		sel.Cut() 		Application.Documents.Add().Content.Paste() 	} }` |

#### **Selection.Delete**

删除指定数量的字符或单词。

**语法**

**express.Delete(Unit, Count)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Unit*   | 可选          | **Variant**  | {删除折叠选定内容时基于的单位。可以是 WdUnits 常量之一。     |
| *Count*  | 可选          | **Variant**  | 要删除的单位数。要删除选定内容后的单位，折叠选定内容并使用一个正值。要删除选定内容前的单位，折叠选定内容并使用一个负值。 |

**说明**

该方法返回一个代表删除的项目数的 **Long** 类型值。或者，如果删除不成功，则返回 0（零）。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例选择并删除活动文档中的内容。*/ Application.Selection.Delete()` |

#### **Selection.DetectLanguage**

分析指定文本，以确定书写文本的语言类型。

**语法**

**express.DetectLanguage()**

*express*   一个代表 **Selection** 对象的变量。

**说明**

**DetectLanguage** 方法的结果逐个字符地保存在 **LanguageID** 属性中。若要读取 **LanguageID** 的属性，则必须先指定文本的选定内容或区域。

如果选定内容包含某个句子的一部分，则选定内容将扩展到该句的句末。

如果指定文本已应用了 **DetectLanguage** 方法，则 **LanguageDetected** 属性将设置为 **True**。要重新检测指定文本的语言，必须先将 LanguageDetected 属性设置为 **False**。

#### **Selection.Endkey**

将选定内容移动或扩展到指定单位的末尾。

**语法**

**express.Endkey(Unit, Extend)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Unit*   | 可选          | **Variant**  | 移动或扩展选定内容时基于的单位。可以是 **WdUnits** 常量之一。默认值为 **wdLine**。 |
| *Extend* | 可选          | **Variant**  | 指定移动所选内容的方式。可以是任意 **WdMovementType** 常量。如果该参数值为 wdMove，则所选内容折叠到一个插入点中并移至指定单位的末尾。如果该参数值为 **wdExtend**，则所选内容的末尾扩展到指定单位的末尾。默认值为 **wdMove**。 |

**说明**

该方法返回表明选定内容或活动结尾实际移动的字符数的整数；如果移动不成功，则返回 0（零）。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将选定内容移动到当前行尾，然后把移动的字符数赋给 pos 变量。*/ Application.Selection.EndKey(wdLine,wdMove)  /*本示例将所选内容移至当前表格列的开始，然后将所选内容扩展至列的末尾。*/ function test() { 	let sel = Application.Selection; 	if(sel.Information(wdWithInTable)) { 		sel.HomeKey(wdColumn,wdMove); 		sel.EndKey(wdColumn,wdExtend); 	} }` |

#### **Selection.EndOf**

将区域或选定内容的结束字符位置移动或扩展至最近的一个指定文本单元末尾。

**语法**

**express.EndOf(Unit, Extend)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Unit*   | 可选          | **Variant**  | 移动结束字符位置时所基于的单位。WdUnits。                    |
| *Extend* | 可选          | **Variant**  | 可以是任意 WdMovementType 常量。如果该参数值为 wdMove，则区域和选定内容对象的结尾都移至指定单位的末尾。如果使用 wdExtend，则区域或选定内容的末尾将扩展到指定单位的末尾。默认值为 wdMove。 |

**说明**

本方法返回该区域或所选内容所移动或扩展的字符位置数（移动方向为向前）。

如果区域或选定内容的开始和结束位置都已在指定单位的末尾，则该方法不移动或扩展此区域或选定内容。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本例将字符串 myRange 移动到选定内容的第一个单词的末尾。 function test(){     let myRange = Application.Selection.Characters(1)     myRange.EndOf.Unit = wdWord     myRange.EndOf.Extend = wdMove }` |

#### **Selection.EscapeKey**

取消某种模式，如取消扩展或列选定模式（与按 ESC 相同）。

**语法**

**express.EscapeKey()**

*express*   一个代表 **Selection** 对象的变量。

**示例**

| 示例代码复制                                             |
| -------------------------------------------------------- |
| `//本示例取消扩展模式。 Application.Selection.EscapeKey` |

#### **Selection.Expand**

扩展指定区域或选定内容。返回添至该区域或选定内容的字符数。**Long** 类型。

**语法**

**express.Expand(Unit)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Unit*   | 可选          | **Variant**  | **WdUnits 变量，表示指定区域扩展距离的度量单位。默认值为 wdWord。** |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例先将选定内容的首个字符设为大写，再将选定内容扩展至整句。*/ function test() { 	let sel = Application.Selection; 	sel.Characters.Item(1).Case = wdTitleSentence 	sel.Expand(wdSentence) }` |

#### **Selection.ExportAsFixedFormat**

将当前选定内容另存为 PDF 或 XPS 格式。

**语法**

**express.ExportAsFixedFormat(OutputFileName, ExportFormat, OpenAfterExport, OptimizeFor, ExportCurrentPage, Item, IncludeDocProps, KeepIRM, CreateBookmarks, DocStructureTags, BitmapMissingFonts, UseISO19005_1, FixedFormatExtClassPtr)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称**                 | **必选/可选** | **数据类型**                | **说明**                                                     |
| ------------------------ | ------------- | --------------------------- | ------------------------------------------------------------ |
| *OutputFileName*         | 必选          | **String**                  | 新的 PDF 或 XPS 文件的路径和文件名。                         |
| *ExportFormat*           | 必选          | **WdExportFormat**          | 指定采用 PDF 格式或 XPS 格式。                               |
| *OpenAfterExport*        | 可选          | **Boolean**                 | 导出内容后打开新文件。                                       |
| *OptimizeFor*            | 可选          | **WdExportOptimizeFor**     | 指定是针对屏幕显示还是打印进行优化。                         |
| *ExportCurrentPage*      | 可选          | **Boolean**                 | 指定是否导出当前页。如果为 True，则导出整个页面。如果为 False，则只导出当前所选内容。 |
| *Item*                   | 可选          | **WdExportItem**            | 指定导出过程是只包括文本还是包括文本和标记。                 |
| *IncludeDocProps*        | 可选          | **Boolean**                 | 指定在最新导出的文件中是否包括文档属性。                     |
| *KeepIRM*                | 可选          | **Boolean**                 | 指定在源文档受 IRM 保护的情况下，是否将 IRM 权限复制到 XPS 文档。默认值为 True。 |
| *CreateBookmarks*        | 可选          | **WdExportCreateBookmarks** | 指定是否导出书签和要导出的书签的类型。                       |
| *DocStructureTags*       | 可选          | **Boolean**                 | 指定是否包含额外数据（如有关内容的流和逻辑组织的信息）以便为屏幕读取器提供帮助。默认值为 True。 |
| *BitmapMissingFonts*     | 可选          | **Boolean**                 | 指定是否包含文本的位图。当字体许可不允许在 PDF 文件中嵌入某一字体时，将此参数设置为 True。如果为 False，则引用该字体，且查看者的计算机会替换适当的字体（如果此创作的字体不可用）。默认值为 True。 |
| *UseISO19005_1*          | 可选          | **Boolean**                 | 指定是否将 PDF 的使用范围限制在按照 ISO 19005-1 标准化的 PDF 子集。如果为 True，生成的文件会更加可靠地自我包含，但由于受到格式的限制，可能会导致该文件较大或显示更多的可视项目。默认值为 False。 |
| *FixedFormatExtClassPtr* | 可选          | **Variant**                 | 指定一个指针以指向一个允许对代码的备用实现进行调用的加载项。代码的备用实现将对应用程序生成的 EMF 和 EMF+ 页面描述进行解释，以生成其自身的 PDF 或 XPS。 |

#### **Selection.Extend**

打开扩展模式。如果扩展模式已经打开，则将选定内容扩展到下一个更大的文本单位。

**语法**

**express.Extend(Character)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Character* | 可选          | **Variant**  | 通过气扩展选定内容的字符。该参数区分大小写，其值必须为 String 类型，否则将发生错误。另外，如果该参数的值不止一个字符，则 WPS 完全忽略该命令。 |

**说明**

使用该方法会将 **ExtendMode** 属性设置为 **True**（如果尚不是此设置）。

逐级选定文本单位是：单词、句、段落、节和整篇文档。如果指定了 *Character*，则该方法正向扩展选定内容至指定字符下一次出现的位置。通过移动选定内容的活动结尾来扩展选定内容。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例激活选定内容的结尾并将选定内容扩展至下一个以“R”开头的实例。 function test(){     Application.Selection.StartIsActive = false     Application.Selection.Extend("R") }` |

#### **Selection.Goto**

将插入点移至紧靠指定项之前的字符位置，并返回一个 **Range** 对象（除 **wdGoToGrammaticalError**、**wdGoToProofreadingError** 或 **wdGoToSpellingError** 常量之外）。

**语法**

**express.Goto(What, Which, Count, Name)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *What*   | 可选          | **Variant**  | 范围或所选内容移动到的项目的种类。可以是 **WdGoToItem** 常量之一。 |
| *Which*  | 可选          | **Variant**  | 范围或所选内容要移动到的项目。可以是 **WdGoToDirection** 常量之一。 |
| *Count*  | 可选          | **Variant**  | 文档中的项数。默认值为 1。只有正值有效。要指定范围或所选内容之前的一个项目，可使用 **wdGoToPrevious** 作为 **Which** 参数，并指定一个 **Count** 值。 |
| *Name*   | 可选          | **Variant**  | 如果 What 参数为 **wdGoToBookmark、wdGoToComment、wdGoToField 或 wdGoToObject，则此参数指定一个名称。** |

**说明**

将 **GoTo** 方法与 **wdGoToGrammaticalError**、**wdGoToProofreadingError** 或 **wdGoToSpellingError** 常量一起使用时，返回的 [**Range** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)对象中包括所有含语法或拼写错误的文本。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下列示例功能等效，都将所选内容移至文档中的第一个标题。*/ Application.Selection.GoTo(wdGoToHeading,wdGoToFirst); Application.Selection.GoTo(wdGoToHeading,wdGoToAbsolute,1);  /*以下示例将插入点移至活动文档中第五个尾注引用标记的前面。*/ function test() { 	if(Application.ActiveDocument.Endnotes.Count > 5) { 		Application.Selection.GoTo(wdGoToEndnote,wdGoToAbsolute,5); 	} }  ` |

#### **Selection.GoToEditableRange**

返回一个 **Range** 对象，该对象代表指定用户或用户组可以修改的文档区域。

**语法**

**express.GoToEditableRange(EditorID)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称**   | **必选/可选** | **数据类型** | **说明**                                                     |
| ---------- | ------------- | ------------ | ------------------------------------------------------------ |
| *EditorID* | 可选          | **Variant**  | 可以是代表用户电子邮件别名（如果在相同的域中）和电子邮件地址的 String 类型值，或者是代表用户组的 WdEditorType 常量。如果省略该参数，则选择所有区域以便所有用户都拥有编辑权限。 |

**返回值**

Range

**说明**

还可使用 **Editor** 对象的 NextRange 属性来返回用户有权修改的下一区域。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例转到当前用户有权修改的下一区域。 Application.Selection.GoToEditableRange = wdEditorCurrent` |

#### **Selection.GoToNext**

返回一个 **Range** 对象，该对象参考下一项目的起始位置或者由 *What* 参数确定的位置。如果将该方法应用于 **Selection** 对象，则该方法将选定内容移动到指定的项目处（但 **wdGoToGrammaticalError**、**wdGoToProofreadingError** 和 **wdGoToSpellingError** 常量除外）。

**语法**

**express.GoToNext(What)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型**   | **说明**                             |
| -------- | ------------- | -------------- | ------------------------------------ |
| *What*   | 必选          | **WdGoToItem** | 指定的区域或选定内容将被移动到的项。 |

#### **Selection.GoToPrevious**

返回一个 **Range** 对象，该对象参考前一项目的起始位置或者由参数 *What* 指定的位置。如果对 **Selection** 对象采用此方法，**GoToPrevious** 将选定内容移动到所指定的项目处。**Range** 对象。

**语法**

**express.GoToPrevious(What)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型**   | **说明**                             |
| -------- | ------------- | -------------- | ------------------------------------ |
| *What*   | 必选          | **WdGoToItem** | 指定的区域或选定内容将被移动到的项。 |

#### **Selection.Homekey**

将选定内容移动或扩展至指定单位的开始。该方法返回表明选定内容实际移动的字符数的整数；如果移动不成功，则返回 0（零）。该方法相当于 **Home** 键的功能。

**语法**

**express.Homekey(Unit, Extend)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Unit*   | 可选          | **Variant**  | 移动或扩展选定内容时基于的单位。默认值为 **wdLine**。        |
| *Extend* | 可选          | **Variant**  | 指定移动所选内容的方式。可以是 **WdMovementType** 常量之一。如果该参数值为 wdMove，则所选内容折叠到一个插入点中并移至指定单位的开头。如果参数值为 **wdExtend，则所选内容的开头扩展到指定单位的开头。默认值为 wdMove。** |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将所选内容移至当前文字部分的开头。如果所选内容位于文档正文部分，则本示例将所选内容移至文档的开头。*/ Application.Selection.HomeKey(wdStory,wdMove)  /*本示例将选定内容移到当前行首，将移动的字符数赋给变量 pos。*/ let pos = Application.Selection.HomeKey(wdLine,wdMove)` |

#### **Selection.InRange**

如果应用此方法的所选内容包含在 *Range* 参数所指定的范围内，则该方法为 **True**。

**语法**

**express.InRange(Range)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                     |
| -------- | ------------- | ------------ | ---------------------------- |
| *Range*  | 必选          | **Variant**  | 要与所选内容进行比较的范围。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例确定所选内容是否包含在活动文档的第一段中。 function test() {     let mes = Application.Selection.InRange(Application.ActiveDocument.Paragraphs(1).Range)     alert(mes) }` |

#### **Selection.InsertAfter**

将指定文本插入范围或所选内容的末尾。

**语法**

**express.InsertAfter(Text)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**       |
| -------- | ------------- | ------------ | -------------- |
| *Text*   | 必选          | **String**   | 要插入的文本。 |

**说明**

使用此方法后，所选内容将扩展以包含新文本。

如果对引用整个段落的所选内容使用此方法，则在末段标记之后插入文本（插入文本将出现在下一段落的开头）。要在段尾插入文本，请先确定终点，再从该位置减去 1（因为段落标记是一个字符），如以下示例所示。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() { 	let activeDoc = Application.ActiveDocument; 	let paragraohs = activeDoc.Paragraphs.Item(1); 	activeDoc.Range(paragraohs.Range.Start,paragraohs.Range.End - 1); 	Application.Selection.InsertAfter(" This is now the last sentence in paragraph one."); }` |

然而，如果所选内容以一个段落标记结尾，而该段落标记正好又是文档的末尾，则 WPS 在末段标记前插入文本，而不是在文档末尾创建一个新段落。同样，如果所选内容是书签，则 WPS 会插入指定文本，但是不会扩展所选内容或书签使其包含新文本。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在所选内容的末尾插入文本，然后将所选内容折叠到插入点。*/ function test() { 	Application.Selection.InsertAfter("appended text") 	Application.Selection.Collapse(wdCollapseEnd) }` |

#### **Selection.InsertBefore**

在指定选定内容之前插入指定文本。

**语法**

**express.InsertBefore(Text)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**       |
| -------- | ------------- | ------------ | -------------- |
| *Text*   | 必选          | **String**   | 要插入的文本。 |

**说明**

使用此方法插入文字后，所选内容将扩展以包含新文字。如果所选内容是书签，则书签也会扩展以包含下一处文字。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在所选内容之前插入文字“Hamlet”（括在引号中），然后折叠所选内容。*/ function test() { 	Application.Selection.InsertBefore("\"Hamlet\"") 	Application.Selection.Collapse(wdCollapseEnd) }` |

#### **Selection.InsertBreak**

插入分页符、分栏符或分节符。

**语法**

**express.InsertBreak(Type)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型**    | **说明**                                                     |
| -------- | ------------- | --------------- | ------------------------------------------------------------ |
| *Type*   | 必选          | **WdBreakType** | 要插入的分隔符的类型。默认值为 wdPageBreak。根据您所选择或安装的语言支持（例如，美国英语），某些 WdBreakType 常量可能不可用。 |

**说明**

插入分页符或分栏符时，插入的分隔符将替换所选内容。如果您不希望替换所选内容，请在使用 **InsertBreak** 方法之前使用 Collapse 方法。

#### **Selection.InsertCaption**

紧靠在指定的所选内容之前或之后插入题注。

**语法**

**express.InsertCaption(Label, Title, TitleAutoText, Position, ExcludeLabel)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称**        | **必选/可选** | **数据类型** | **说明**                                                     |
| --------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Label*         | 必选          | **Variant**  | 要插入的题注标签。可以是一个 String，也可以是 WdCaptionLabelID 常量之一。如果未定义该标签，则会发生错误。可以使用 Add 方法和 CaptionLabels 对象定义新的题注标签。 |
| *Title*         | 可选          | **Variant**  | 要紧接在题注标签之后插入的字符串（如果已指定 TitleAutoText，则忽略该参数）。 |
| *TitleAutoText* | 可选          | **Variant**  | 要紧接在题注标签之后插入其内容的“自动图文集”词条（将覆盖由 Title 指定的任何文本）。 |
| *Position*      | 可选          | **Variant**  | 指定题注是插入到所选内容上方还是下方。可以是 WdCaptionPosition 常量之一。 |
| *ExcludeLabel*  | 可选          | **Variant**  | 如果该参数值为 True，则不包括 Label 参数中定义的文本标签。如果该参数值为 False，则包括指定的标签。 |

#### **Selection.InsertCells**

向原有表添加单元格。

**语法**

**express.InsertCells(ShiftCells)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称**     | **必选/可选** | **数据类型**      | **说明**                                 |
| ------------ | ------------- | ----------------- | ---------------------------------------- |
| *ShiftCells* | 可选          | **WdInsertCells** | 指定如何将单元格插入到表的现有列和行中。 |

**说明**

插入的单元格数量等于选定内容中单元格的数量。 您还可以使用 Cells 对象的 [Add](https://docs.microsoft.com/zh-cn/office/vba/api/word.cells.add) 方法 插入单元格。

#### **Selection.InsertColumns**

在选定内容的左侧插入新列。

**语法**

**express.InsertColumns()**

*express*   一个代表 **Selection** 对象的变量。

**说明**

插入的列数与选定列数相等。还可以使用 **Columns** 对象的 Add 方法插入列。

如果选定内容不在表格中，则会导致出错。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例在所选内容的左侧插入新列。插入的列数与选定列数相等。 function test(){     if(Application.Selection.Information(wdWithInTable)){         Application.Selection.InsertColumns()     } }` |

#### **Selection.InsertColumnsRight**

在当前选定内容的右边插入列。

**语法**

**express.InsertColumnsRight()**

*express*   一个代表 **Selection** 对象的变量。

**说明**

WPS 插入与当前选定内容中数量相同的列。

要使用此方法，当前所选内容必须处于表格中。

#### **Selection.InsertCrossReference**

插入对标题、书签、脚注、尾注或定义了题注标签的项（如公式、图表或表格）的交叉引用。

**语法**

**express.InsertCrossReference(ReferenceType, ReferenceKind, ReferenceItem, InsertAsHyperlink, IncludePosition, SeparateNumbers, SeparatorString)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称**            | **必选/可选** | **数据类型**        | **说明**                                                     |
| ------------------- | ------------- | ------------------- | ------------------------------------------------------------ |
| *ReferenceType*     | 必选          | **Variant**         | 要插入交叉引用的项目的类型。可以是任意 WdReferenceType 或 WdCaptionLabelID 常量，也可以是用户定义的题注标签。 |
| *ReferenceKind*     | 必选          | **WdReferenceKind** | 要包含在交叉引用中的信息。                                   |
| *ReferenceItem*     | 必选          | **Variant**         | 如果 ReferenceType 为 wdRefTypeBookmark，则此参数指定一个书签名称。对于所有其他 ReferenceType 值，此参数将指定“交叉引用”对话框的“引用类型”框中的项目编号或名称。使用 GetCrossReferenceItems 方法可返回能够用于此参数的项目名称的列表。 |
| *InsertAsHyperlink* | 可选          | **Variant**         | 如果该参数值为 True，则作为超链接插入交叉引用。              |
| *IncludePosition*   | 可选          | **Variant**         | 如果该参数值为 True，则根据引用项目相对于交叉引用的位置插入“above”或“below”。 |
| *SeparateNumbers*   | 可选          | **Variant**         | 如果该参数值为 True，则使用分隔符将数字从关联文本中分离出来。（仅在将 ReferenceType 参数设置为 wdRefTypeNumberedItem 并将 ReferenceKind 参数设置为 wdNumberFullContext 时才可使用）。 |
| *SeparatorString*   | 可选          | **Variant**         | 如果 SeparateNumbers 参数设置为 True，则指定字符串用作分隔符。 |

**说明**

如果指定 **wdPageNumber** 作为 *ReferenceKind* 的值，则可能需要对文档重新分页才能看到正确的交叉引用信息。

#### **Selection.InsertDateTime**

以文本或 TIME 域的形式插入当前日期或时间，或将两者都插入。

**语法**

**express.InsertDateTime(DateTimeFormat, InsertAsField, InsertAsFullWidth, DateLanguage, CalendarType)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称**            | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *DateTimeFormat*    | 可选          | **Variant**  | 显示日期或时间或者同时显示两者时所用的格式。如果省略该参数，WPS 将使用 Windows“控制面板”（“区域设置”图标）中的短日期样式。 |
| *InsertAsField*     | 可选          | **Variant**  | 如果该参数值为 True，则以 TIME 域的形式插入指定的信息。默认值为 True。 |
| *InsertAsFullWidth* | 可选          | **Variant**  | 如果该参数值为 True，则以双字节数字的形式插入指定信息。根据您所选择或安装的语言支持（例如，美国英语），此参数可能不可用。 |
| *DateLanguage*      | 可选          | **Variant**  | 设置显示日期或时间所使用的语言。可以是 WdDateLanguage 常量之一。根据您所选择或安装的语言支持（例如，美国英语），此参数可能不可用。 |
| *CalendarType*      | 可选          | **Variant**  | 设置显示日期或时间时使用的日历类型。可以是 WdCalendarTypeBi 常量之一。根据您所选择或安装的语言支持（例如，美国英语），此参数可能不可用。 |

**说明**

使用此方法将替换当前所选内容。为了避免出现此情况，请在使用此方法之前使用 Collapse 方法。

#### **Selection.InsertFile**

插入指定文件的全部或一部分。

**语法**

**express.InsertFile(FileName, Range, ConfirmConversions, Link, Attachment)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称**             | **必选/可选** | **数据类型** | **说明**                                                     |
| -------------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *FileName*           | 必选          | **String**   | 要插入的文件的路径及文件名。如果没有指定路径，则 WPS 假定文件位于当前文件夹中。 |
| *Range*              | 可选          | **Variant**  | 如果指定的文件是 WPS 文档，则此参数代表书签。如果该文件为其他类型（例如，ET 工作表），则此参数代表命名区域或单元格区域（例如，R1C1:R3C4）。 |
| *ConfirmConversions* | 可选          | **Variant**  | 如果该参数值为 True，则在插入非 WPS 文档格式的文件时， WPS 将提示您确认转换。 |
| *Link*               | 可选          | **Variant**  | 如果该参数值为 True，则使用 INCLUDETEXT 域插入该文件。       |
| *Attachment*         | 可选          | **Variant**  | 如果该参数值为 True，则将该文件以附件形式插入电子邮件中。    |

#### **Selection.InsertFormula**

插入包含选定内容的公式的 = (Formula) 域。

**语法**

**express.InsertFormula(Formula, NumberFormat)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称**       | **必选/可选** | **数据类型** | **说明**                                                     |
| -------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Formula*      | 可选          | **Variant**  | 希望 = (Formula) 域求值的数学公式。可以使用与电子表格类似的方式引用表格中的单元格。例如，“=SUM(A4:C4)”指定第四行中的前三个值。有关 = (Formula) 域的详细内容，请参阅域代码：= (Formula) 域。 |
| *NumberFormat* | 可选          | **Variant**  | = (Formula) 域的结果的格式。有关可应用的格式类型的信息，请参阅“数字图片 (\#)”域开关。 |

**说明**

如果所选内容未折叠，则该公式将替换所选内容。

如果正在使用电子表格应用程序（如 WPS OfficeExcel），那么在文档中嵌入全部或部分工作表比在表格中使用 = (Formula) 域更简单易行。

只有当所选内容在一个单元格中，并且插入点所在的单元格的上方或左边至少有一个包含值的单元格时，*Formula* 参数才是可选的。如果插入点上方的单元格包含值，则插入域 {=SUM(ABOVE)}。如果插入点左边的单元格包含值，则插入域 {=SUM(LEFT)}。如果插入点上方和左边的单元格都含有值，则 WPS 使用下面的规则确定插入何种 SUM 函数：

- 如果紧邻插入点上面的单元格中含有数值，则 WPS 插入 {=SUM(ABOVE)}。
- 如果紧邻插入点上方的单元格不包含值，但紧邻插入点左边的单元格包含值，则 WPS 插入 {=SUM(LEFT)}。
- 如果紧邻插入点上方的单元格和下方的单元格都不包含值，则 WPS 插入 {=SUM(ABOVE)}。
- 如果没有指定 **Formula**，并且插入点上方和左边的所有单元格都是空的，则使用 = (Formula) 会发生错误。

#### **Selection.InsertNewPage**

在插入点的位置插入一个新页面。

**语法**

**express.InsertNewPage()**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.InsertParagraph**

用新段落替换指定的所选内容。

**语法**

**express.InsertParagraph()**

*express*   一个代表 **Selection** 对象的变量。

**说明**

使用此方法后，所选内容将包含新段落。如果您不希望替换当前所选内容，请在使用此方法之前使用 **Collapse** 方法。您还可以使用 InsertParagraphBefore 或 InsertParagraphAfter 方法在所选内容之前或之后插入新段落。

#### **Selection.InsertParagraphAfter**

在所选内容之后插入段落标记。

**语法**

**express.InsertParagraphAfter()**

*express*   一个代表 **Selection** 对象的变量。

**说明**

使用此方法后，所选内容将扩展以包含新段落。

#### **Selection.InsertParagraphBefore**

在指定的所选内容或范围前插入一个新段落。

**语法**

**express.InsertParagraphBefore()**

*express*   一个代表 **Selection** 对象的变量。

**说明**

使用此方法后，所选内容将扩展以包含新段落。

#### **Selection.InsertRows**

在所选内容的上方插入指定数量的新行。如果选定内容不在表格中，则会导致出错。

**语法**

**express.InsertRows(NumRows)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型** | **说明**         |
| --------- | ------------- | ------------ | ---------------- |
| *NumRows* | 可选          | **Variant**  | 需要添加的行数。 |

**说明**

还可以通过使用 **Rows** 对象的 Add 方法插入行。

#### **Selection.InsertRowsAbove**

在当前选定内容上方插入行。

**语法**

**express.InsertRowsAbove()**

*express*   一个代表 **Selection** 对象的变量。

**说明**

WPS 可在当前选定内容中插入任意数量的行。

要使用此方法，当前所选内容必须处于表格中。

#### **Selection.InsertRowsBelow**

在当前选定内容的下方插入行。

**语法**

**express.InsertRowsBelow()**

*express*   一个代表 **Selection** 对象的变量。

**说明**

WPS 可在当前选定内容中插入任意数量的行。

要使用此方法，当前所选内容必须处于表格中。

#### **Selection.InsertStyleSeparator**

插入特殊隐藏段落标记，使 WPS 可以使用不同的段落样式合并段落格式，以便在目录中插入内置标题。

**语法**

**express.InsertStyleSeparator()**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.InsertSymbol**

插入一个符号，用来替换指定的所选内容。

**语法**

**express.InsertSymbol(CharacterNumber, Font, Unicode, Bias)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称**          | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *CharacterNumber* | 必选          | **Long**     | 指定符号的字符编号。该值应是 31 加上此符号在符号表中的位置（从左向右计算）相应的数字。例如，要指定一个位于“Symbol”字体符号表中第 37 号位置的 delta 字符，请将 CharacterNumber 设置为 68。 |
| *Font*            | 可选          | **Variant**  | 包含符号的字体的名称。                                       |
| *Unicode*         | 可选          | **Variant**  | 如果该参数值为 True，则插入由 CharacterNumber 指定的 Unicode 字符；如果该参数值为 False，则插入由 CharacterNumber 指定的 ANSI 字符。默认值为 False。 |
| *Bias*            | 可选          | **Variant**  | 为符号设置字体斜线。此参数可用于为东亚字符设置正确的字体斜线。可以是 WdFontBias 常量之一。根据您所选择或安装的语言支持（例如，美国英语），此参数可能不可用。 |

**说明**

如果不希望替换所选内容，请在使用此方法之前使用 **Collapse** 方法。

#### **Selection.InsertXML**

将指定的 XML 插入文档中的光标处，并且替换任何选定文本。

**语法**

**express.InsertXML(XML, Transform)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *XML*       | 必选          | **String**   | 指定要插入的 XML。这可以是任何有效的自定义 XML。             |
| *Transform* | 可选          | **Variant**  | 指定用于转换 XML 的 XML 转换 (XSLT)。如果省略，则将 XML 作为自定义 XML 插入，而不应用转换。 |

#### **Selection.InStory**

如果应用此方法的选择范围与 *Range* 参数指定的范围位于相同的文字部分中，则该方法为 **True**。

**语法**

**express.InStory(Range)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                           |
| -------- | ------------- | ------------ | -------------------------------------------------- |
| *Range*  | 必选          | **Range**    | 其文字部分与当前选择范围的文字部分进行比较的范围。 |

**返回值**

Boolean

**说明**

一个范围只能属于一个文字部分。

#### **Selection.IsEqual**

如果应用此方法的选择范围等于 *Range* 参数所指定的范围，则该方法为 **True**。

**语法**

**express.IsEqual(Range)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                       |
| -------- | ------------- | ------------ | ------------------------------ |
| *Range*  | 必选          | **Range**    | 与当前选择范围进行比较的范围。 |

**说明**

此方法比较起始字符和结束字符的位置以及文字部分的类型。如果所选内容和 *Range* 参数所指定的对象的所有三项都相同，则这两个对象相等。

#### **Selection.ItalicRun**

在当前局部中添加或删除斜体字符格式。

**语法**

**express.ItalicRun()**

*express*   一个代表 **Selection** 对象的变量。

**说明**

如果该局部既包含斜体文本，又有非斜体文本，则该方法将在整个局部中增添斜体字符格式。

#### **Selection.LtrPara**

将指定段落的对齐方式和阅读顺序设置为从左向右。

**语法**

**express.LtrPara()**

*express*   一个代表 **Selection** 对象的变量。

**说明**

对于所有已选定的段落，此方法将其阅读顺序设置为从左向右。如果一个段落的阅读顺序是从右向左并且为右对齐方式，则此方法会将其设置为相反的阅读顺序，并将其对齐方式设置为左对齐。

可用 **ReadingOrder** 属性来改变阅读顺序而不影响段落的对齐方式。

#### **Selection.LtrRun**

用于将指定局部的阅读顺序和对齐方式设置为从左向右。

**语法**

**express.LtrRun()**

*express*   一个代表 **Selection** 对象的变量。

**说明**

对指定的局部，此方法将其阅读顺序设置为从左向右。如果该局部中的某一段落的阅读顺序为从右向左且为右对齐的，则此方法会将其设置为相反的阅读顺序，并将段落对齐方式设置为左对齐。

可用 **ReadingOrder** 属性来改变阅读顺序而不影响段落的对齐方式。

#### **Selection.Move**

将指定的所选内容折叠到其起始位置或结束位置，然后将折叠的对象移动指定的单位数。此方法返回一个 **Long** 类型的值，该值代表所选内容移动的单位数；如果移动失败，则返回 0（零）。

**语法**

**express.Move(Unit, Count)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Unit*   | 可选          | **WdUnits**  | 移动结束字符位置时所基于的单位。                             |
| *Count*  | 可选          | **Variant**  | 指定的范围或所选内容要移动的单位数。如果 Count 为正数，则对象折叠到其结束位置，并在文档中向后移动指定的单位数；如果 Count 为负数，则对象折叠到其起始位置，并向前移动指定的单位数。默认值为 1。您也可以在使用 Move 方法之前通过使用 Collapse 方法来控制折叠的方向。如果区域或所选内容位于单位的中间位置或者未折叠，则可以认为将其移动到单位的开始或结束位置也就是将它移动了一个完整单位。 |

**说明**

折叠的范围或所选内容的开始和结束位置相同。

对一个范围应用 **Move** 方法不会重排文档中的文本，而是重新定义该范围，以引用文档中的一个新位置。

如果将 **Move** 方法应用于除 **Range** 对象变量以外的任意范围（例如，`Selection.Paragraphs(3).Range.Move`），则此方法不起作用。

移动 **Selection** 对象会折叠所选内容，并在文档中向前或向后移动插入点。

#### **Selection.MoveDown**

将选定内容向下移动，并返回移动距离的单位数。

**语法**

**express.MoveDown(Unit, Count, Extend)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Unit*   | 可选          | **WdUnits**  | 移动选定内容时所基于的单位。默认值为 wdLine。                |
| *Count*  | 可选          | **Variant**  | 选定内容移动的单位数。默认值为 1。                           |
| *Extend* | 可选          | **Variant**  | 可以是 wdMove 或 wdExtend。如果使用 wdMove，则所选内容折叠到结束位置并向下移动。如果使用 wdExtend，则所选内容向下扩展。默认值为 wdMove。 |

#### **Selection.MoveEnd**

移动范围或所选内容的结束字符位置。

**语法**

**express.MoveEnd(Unit, Count)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Unit*   | 可选          | **WdUnits**  | 结束字符位置要移动的单位。默认值为 wdCharacter。             |
| *Count*  | 可选          | **Variant**  | 要移动的单位数。如果该值为正，则结束字符在文档中向前移动；如果该值为负，则结束字符向后移动。如果结束位置超过了起始位置，则折叠该范围并将首末两个字符的位置移至同一位置。默认值为 1。 |

**说明**

此方法如果返回一个整数，表示范围或所选内容实际移动的单位数；如果返回 0（零），则移动失败。

#### **Selection.MoveEndUntil**

移动指定的所选内容的结束位置，直到在文档中找到任何指定的字符。

**语法**

**express.MoveEndUntil(Cset, Count)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Cset*   | 必选          | **Variant**  | 一个或多个字符。此参数区分大小写。                           |
| *Count*  | 可选          | **Variant**  | 指定的所选内容要移动的最大字符数。可以是一个数字，也可以是 wdForward 或 wdBackward 常量。如果 Count 为正数，则所选内容在文档中向前移动。如果为负数，则所选内容向后移动。默认值为 wdForward。 |

**返回值**

Long

**说明**

此方法返回一个 **Long** 类型的值，该值代表指定的所选内容的结束位置要移动的字符数。如果 *Count* 大于 0（零），则此方法返回移动的字符数加 1。如果 *Count* 小于 0（零），则此方法返回移动的字符数减 1。如果没有找到 *Cset* 字符，则不改变所选内容，并且此方法返回 0（零）。如果结束位置向后移至原来的起始位置之前，则将该起始位置设置为新的结束位置。

如果是在文档中向前移动，则扩展所选内容。

#### **Selection.MoveEndWhile**

当在文档中找到任何指定的字符时，移动所选内容的结束字符位置。

**语法**

**express.MoveEndWhile(Cset, Count)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Cset*   | 必选          | **Variant**  | 一个或多个字符。此参数区分大小写。                           |
| *Count*  | 可选          | **Variant**  | 所选内容要移动的最大字符数。可以是一个数字，也可以是 wdForward 或 wdBackward 常量。如果 Count 为正数，则所选内容在文档中向前移动。如果为负数，则所选内容向后移动。默认值为 wdForward。 |

**说明**

在找到 *Cset* 中的任何字符时，就移动指定所选内容的结束位置。此方法将该所选内容的结束位置移动的字符数作为 **Long** 类型的值返回。如果没有找到 *Cset* 字符，则不改变所选内容，并且此方法返回 0（零）。如果结束位置向后移至原来的起始位置之前，则将该起始位置设置为新的结束位置。

#### **Selection.MoveLeft**

将选定内容向左移动，并返回移动距离的单位数。

**语法**

**express.MoveLeft(Unit, Count, Extend)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Unit*   | 可选          | **WdUnits**  | 移动选定内容时所基于的单位。默认值为 wdCharacter。           |
| *Count*  | 可选          | **WdUnits**  | 选定内容移动的单位数。默认值为 1。                           |
| *Extend* | 可选          | **Variant**  | 可以是 wdMove 或 wdExtend。如果使用 wdMove，则所选内容折叠到结束位置并向左移动。如果使用 wdExtend，则所选内容向左扩展。默认值为 wdMove。 |

**说明**

如果 *Unit* 为 **wdCell**，则 *Extend* 参数只能是 **wdMove**。

#### **Selection.MoveRight**

将选定内容向右移动，并返回移动距离的单位数。

**语法**

**express.MoveRight(Unit, Count, Extend)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Unit*   | 可选          | **WdUnits**  | 移动选定内容时所基于的单位。默认值为 wdCharacter。           |
| *Count*  | 可选          | **Variant**  | 选定内容移动的单位数。默认值为 1。                           |
| *Extend* | 可选          | **Variant**  | 可以是 wdMove 或 wdExtend。如果使用 wdMove，则所选内容折叠到结束位置并向右移动。如果使用 wdExtend，则所选内容向右扩展。默认值为 wdMove。 |

**说明**

如果 *Unit* 为 **wdCell**，则 *Extend* 参数只能是 **wdMove**。

#### **Selection.MoveStart**

移动指定的所选内容的起始位置。

**语法**

**express.MoveStart(Unit, Count)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Unit*   | 可选          | **WdUnits**  | 指定的所选内容的起始位置要移动的单位。可以是 WdUnits 常量之一。默认值为 wdCharacter。 |
| *Count*  | 可选          | **Variant**  | 指定的所选内容要移动的最大单位数。如果 Count 为正数，则所选内容的起始位置在文档中向前移动。如果为负数，则起始位置向后移动。如果起始位置向前移至结束位置之前，则折叠所选内容，并将起始位置和结束位置移至同一位置。默认值为 1。 |

**说明**

此方法返回一个整数，该整数表明起始位置或该所选内容实际移动的单位数，如果移动不成功，则此方法返回 0（零）。

#### **Selection.MoveStartUntil**

移动指定的所选内容的起始位置，直到在文档中找到一个指定的字符。如果是在文档中向后移动，则扩展所选内容。

**语法**

**express.MoveStartUntil(Cset, Count)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Cset*   | 必选          | **Variant**  | 一个或多个字符。此参数区分大小写。                           |
| *Count*  | 可选          | **Variant**  | 指定的所选内容要移动的最大字符数。可以是一个数字，也可以是 wdForward 或 wdBackward 常量。如果 Count 为正数，则所选内容在文档中向前移动。如果为负数，则所选内容向后移动。默认值为 wdForward。 |

**说明**

此方法将指定的所选内容的起始位置移动的字符数作为 **Long** 类型的值返回。如果 *Count* 大于 0（零），则此方法返回移动的字符数加 1。如果 *Count* 小于 0（零），则此方法返回移动的字符数减 1。如果未找到 *Cset* 字符，则指定的所选内容不更改，并且此方法返回 0（零）。如果起始位置向前移至结束位置的前面，则折叠所选内容，并且开始和结束位置移至同一位置。

#### **Selection.MoveStartWhile**

当在文档中找到任何指定的字符时，移动指定的所选内容的起始位置。

**语法**

**express.MoveStartWhile(Cset, Count)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Cset*   | 必选          | **Variant**  | 一个或多个字符。此参数区分大小写。                           |
| *Count*  | 可选          | **Variant**  | 指定的所选内容要移动的最大字符数。可以是一个数字，也可以是 wdForward 或 wdBackward 常量。如果 Count 为正数，则所选内容在文档中向前移动。如果为负数，则所选内容向后移动。默认值为 wdForward。 |

**说明**

当找到 *Cset* 中的任何字符时，就移动所选内容的起始位置。此方法将该所选内容的起始位置移动的字符数作为 **Long** 类型的值返回。如果没有找到 *Cset* 字符，则不改变该所选内容，并且此方法返回 0（零）。如果起始位置向前移至原来的结束位置之前，则将该结束位置设置为新的起始位置。

#### **Selection.MoveUntil**

移动指定的所选内容，直到在文档中找到一个指定的字符。

**语法**

**express.MoveUntil(Cset, Count)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Cset*   | 必选          | **Variant**  | 一个或多个字符。如果在 Count 值终止之前找到了 Cset 中的任何字符，则指定的所选内容作为插入点紧靠在该字符之前放置。此参数区分大小写。 |
| *Count*  | 可选          | **Variant**  | 指定的所选内容要移动的最大字符数。可以是一个数字，也可以是 wdForward 或 wdBackward 常量。如果 Count 为正数，则所选内容在文档中从结束位置开始向前移动。如果为负数，则所选内容从起始位置开始向后移动。默认值为 wdForward。 |

**说明**

此方法将指定的所选内容移动的字符数作为 **Long** 类型的值返回。如果 *Count* 大于 0（零），则此方法返回字符数加 1；如果 *Count* 小于 0（零），则此方法返回字符数减 1。如果没有找到 *Cset* 字符，则不改变所选内容，并且此方法返回 0（零）。

#### **Selection.MoveUp**

将所选内容向上移动，并返回移动的单位数。

**语法**

**express.MoveUp(Unit, Count, Extend)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Unit*   | 可选          | **Variant**  | 移动选定内容时所基于的单位。可以是下列 WdUnits 常量之一：wdLine、wdParagraph、wdWindow 或 wdScreen。默认值为 wdLine。可以将 wdWindow 常量用于 Unit 参数，以便移至活动窗口的顶部或底部。不管 Count 的值是多少（大于 1 还是小于 -1），wdWindow 常量只移动一个单位。可使用 wdScreen 常量移动多屏。 |
| *Count*  | 可选          | **Variant**  | 选定内容移动的单位数。默认值为 1。                           |
| *Extend* | 可选          | **Variant**  | 指定是移动还是扩展所选内容。可以是 wdMove 或 wdExtend。如果使用 wdMove，则所选内容折叠到结束位置并向上移动。如果使用 wdExtend，则所选内容向上扩展。默认值为 wdMove。 |

**返回值**

Long

#### **Selection.MoveWhile**

当在文档中找到任何指定的字符时，移动指定的选择范围。

**语法**

**express.MoveWhile(Cset, Count)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Cset*   | 必选          | **Variant**  | 一个或多个字符。此参数区分大小写。                           |
| *Count*  | 可选          | **Variant**  | 指定的所选内容要移动的最大字符数。可以是一个数字，也可以是 wdForward 或 wdBackward 常量。如果 Count 为正数，则指定的所选内容在文档中从结束位置开始向前移动。如果为负数，则所选内容从起始位置开始向后移动。默认值为 wdForward。 |

**说明**

当找到 *Cset* 中的任何字符时，则移动指定的选择范围。找到任何 *Cset* 字符后，结果 **Selection** 对象作为插入点放置。此方法将指定的选择范围移动的字符数作为 **Long** 类型的值返回。如果没有找到 *Cset* 字符，则不改变所选内容，并且此方法返回 0（零）。

#### **Selection.Next**

返回一个 **Range** 对象，该对象代表相对于指定的选择范围的下一个单位。

**语法**

**express.Next(Unit, Count)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                    |
| -------- | ------------- | ------------ | ------------------------------------------- |
| *Unit*   | 可选          | **Variant**  | 要计算的单位类型。可以是任意 WdUnits 常量。 |
| *Count*  | 可选          | **Variant**  | 要向前移动的单位数。默认值为 1。            |

**说明**

如果选择范围恰好在指定的 *Unit* 之前，则选择范围移至下一个单位

#### **Selection.NextField**

选定下一个域。

**语法**

**express.NextField()**

*express*   一个代表 **Selection** 对象的变量。

**说明**

如果找到了一个域，则该方法返回 **Field** 对象；否则，该方法返回 **Nothing**。

#### **Selection.NextRevision**

找到下一处修订并作为一个 **Revision** 对象返回。

**语法**

**express.NextRevision(Wrap)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Wrap*   | 可选          | **Variant**  | 如果为 True，则在到达文档尾部时继续从文档开始查找修订。默认值为 False。 |

**说明**

修订的文本成为当前的选定区域。用生成的 **Revision** 对象的属性来查看修订的类型、由谁制作的修订，等等。用 **Revision** 对象的方法接受或拒绝修订。

如果没有发现有修订，当前的选择区域保持不变。

#### **Selection.NextSubdocument**

将所选内容移至下一个子文档。

**语法**

**express.NextSubdocument()**

*express*   一个代表 **Selection** 对象的变量。

**说明**

如果没有下一个子文档，将发生错误。

#### **Selection.Paste**

将“剪贴板”的内容插入指定的选择范围处。

**语法**

**express.Paste()**

*express*   一个代表 **Selection** 对象的变量。

**说明**

使用此方法替换当前选择范围的内容。如果您不希望替换所选范围的内容，请在使用此方法之前使用 [**Collapse** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Selection/Selection%20.htm#Selection.Collapse)方法。

如果您将此方法与 [**Range** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)对象一起使用，该范围将扩展以包含“剪贴板”中的内容。如果您使用 **Selection** 对象的此方法，所选内容不会扩展以包含“剪贴板”中的内容；相反，所选内容将放置到粘贴的“剪贴板”内容之后。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例复制文档中的第一段并将其粘贴到插入点。*/ function test() { 	let app = Application; 	app.ActiveDocument.Paragraphs.Item(1).Range.Copy(); 	app.Selection.Collapse(wdCollapseStart); 	app.Selection.Paste(); }` |

#### **Selection.PasteAndFormat**

粘贴选定的表格单元格，并为其设置指定的格式。

**语法**

**express.PasteAndFormat(Type)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型**       | **说明**                               |
| -------- | ------------- | ------------------ | -------------------------------------- |
| *Type*   | 必选          | **WdRecoveryType** | 粘贴所选的表格单元格时使用的格式类型。 |

#### **Selection.PasteAppendTable**

通过在所选行之间插入粘贴的行，将粘贴的单元格合并到现有的表格中。不覆盖任何单元格。

**语法**

**express.PasteAppendTable()**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.PasteAsNestedTable**

将一个或一组单元格作为嵌套表格粘贴到所选内容中。

**语法**

**express.PasteAsNestedTable()**

*express*   一个代表 **Selection** 对象的变量。

**说明**

仅当“剪贴板”中包含一个或一组单元格，并且所选范围是当前文档的一个或一组单元格时，才能使用 **PasteAsNestedTable** 方法。

#### **Selection.PasteExcelTable**

粘贴 ET 表格并设置其格式。

**语法**

**express.PasteExcelTable(LinkedToExcel, WordFormatting, RTF)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称**         | **必选/可选** | **数据类型** | **说明**                                                     |
| ---------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *LinkedToExcel*  | 必选          | **Boolean**  | 如果该参数值为 True，则将粘贴的表格链接到原始 ET 文件，以便对 ET 文件所做的更改可以反映到 WPS 中。 |
| *WordFormatting* | 必选          | **Boolean**  | 如果该参数值为 True，则按照 WPS 文档中的格式设置表格的格式。如果该参数值为 False，则按照原始 ET 文件中的格式设置表格的格式。 |
| *RTF*            | 必选          | **Boolean**  | 如果该参数值为 True，则使用 RTF 格式粘贴 ET 表格。如果该参数值为 False，则以 HTML 格式粘贴 ET 表格。 |

#### **Selection.PasteFormat**

将以 **CopyFormat** 方法复制的格式应用于选定内容。

**语法**

**express.PasteFormat()**

*express*   一个代表 **Selection** 对象的变量。

**说明**

如果用 **CopyFormat** 方法时也选定了段落标记，则 WPS 不仅应用字符格式，也应用段落格式。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将选定内容第一段的段落和字符格式复制到第二段。*/ function test() { 	let sel = Application.Selection; 	sel.Paragraphs.Item(1).Range.Select(); 	sel.CopyFormat(); 	sel.Paragraphs.Item(1).Next().Range.Select(); 	sel.PasteFormat(); }  /*本示例折叠选定内容，并将选定内容的字符格式复制到下一个单词。*/ function test() { 	let sel = Application.Selection; 	sel.Collapse(wdCollapseStart); 	sel.CopyFormat(); 	sel.Next(wdWord,1).Select(); 	sel.PasteFormat(); }` |

#### **Selection.PasteSpecial**

插入“剪贴板”中的内容。

**语法**

**express.PasteSpecial(IconIndex, Link, Placement, DisplayAsIcon, DataType, IconFileName, IconLabel)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称**        | **必选/可选** | **数据类型** | **说明**                                                     |
| --------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *IconIndex*     | 可选          | **Variant**  | 如果 DisplayAsIcon 为 True，则此参数是一个数字，该数字对应于要在程序文件中使用的由 IconFilename 指定的图标。 |
| *Link*          | 可选          | **Variant**  | 如果该参数值为 True，则创建指向“剪贴板”内容源文件的链接。默认值为 False。 |
| *Placement*     | 可选          | **Variant**  | 可以是任一 WdOLEPlacement 常量。                             |
| *DisplayAsIcon* | 可选          | **Variant**  | 如果该参数值为 True，则将链接显示为图标。默认值为 False。    |
| *DataType*      | 可选          | **Variant**  | “剪贴板”内容在插入到文档时采用的格式。WdPasteDataType。      |
| *IconFileName*  | 可选          | **Variant**  | 如果 DisplayAsIcon 为 True，则此参数是存储要显示的图标的文件的路径和文件名。 |
| *IconLabel*     | 可选          | **Variant**  | 如果 DisplayAsIcon 为 True，则此参数是显示在图标下方的文本。 |

**说明**

与 Paste 方法不同，使用 **PasteSpecial** 方法您可以控制粘贴的信息的格式，并且可以建立指向源文件（例如，ET 工作表）的链接（可选）。如果您不希望替换指定所选内容，请在使用此方法之前使用 Collapse 方法。使用此方法时，所选内容不会扩展以包含“剪贴板”中的内容。

#### **Selection.Previous**

将所选文本移动指定的单位数，并返回与折叠的所选内容相关的 **Range** 对象。

**语法**

**express.Previous(Unit, Count)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                |
| -------- | ------------- | ------------ | ------------------------------------------------------- |
| *Unit*   | 可选          | **Variant**  | 指定所选内容移动的单位的类型。可以是 WdUnits 常量之一。 |
| *Count*  | 可选          | **Variant**  | 要移动的单位数。默认值为 1。                            |

**返回值**

Range

#### **Selection.PreviousField**

选择并返回前一域。

**语法**

**express.PreviousField()**

*express*   一个代表 **Selection** 对象的变量。

**返回值**

Field

**说明**

如果找到了一个域，则该方法返回 **Field** 对象；否则，该方法返回 **Nothing**。

#### **Selection.PreviousRevision**

找到前一处修订并作为 **Revision** 对象返回。

**语法**

**express.PreviousRevision(Wrap)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Wrap*   | 可选          | **Variant**  | 如果为 True，则在到达文档开始时继续从文档尾部查找修订。默认值为 False。 |

**返回值**

Revision

#### **Selection.PreviousSubdocument**

将所选内容移至上一个子文档。

**语法**

**express.PreviousSubdocument()**

*express*   一个代表 **Selection** 对象的变量。

**说明**

如果没有上一个子文档，将发生错误。

#### **Selection.ReadingModeGrowFont**

在以阅读模式显示文档时，将所显示文本的大小增大一磅。

**语法**

**express.ReadingModeGrowFont()**

*express*   一个代表 **Selection** 对象的变量。

**说明**

使用 ReadingModeShrinkFont 方法可以减小文本的大小。这不会影响文档中的字号，而只会在以阅读模式查看文档时影响文本的大小。

#### **Selection.ReadingModeShrinkFont**

在以阅读模式显示文档时，将所显示文本的大小减小一磅。

**语法**

**express.ReadingModeShrinkFont()**

*express*   一个代表 **Selection** 对象的变量。

**说明**

使用 ReadingModeGrowFont 方法可以增大文本的大小。这不会影响文档中的字号，而只会在以阅读模式查看文档时影响文本的大小。

#### **Selection.RtlPara**

将指定段落的阅读顺序和对齐方式设置为从右向左。

**语法**

**express.RtlPara()**

*express*   一个代表 **Selection** 对象的变量。

**说明**

本方法将所有选定段落的阅读顺序设置为从右向左。如果某一段落的阅读顺序是从左向右且是左对齐的，本方法可使其阅读顺序相反，且将其对齐方式设置为右对齐。

可用 **ReadingOrder** 属性来改变阅读顺序而不影响段落的对齐方式。

#### **Selection.RtlRun**

将指定的局部的阅读顺序和对齐方式设置为从右向左。

**语法**

**express.RtlRun()**

*express*   一个代表 **Selection** 对象的变量。

**说明**

本示例将指定的局部的阅读顺序设置为从右向左。如果该局部中的某段落的阅读顺序为从左向右且为左对齐方式，则该方法将使其阅读顺序相反，并将其对齐方式设置为右对齐。

可用 **ReadingOrder** 属性来改变阅读顺序而不影响段落的对齐方式。

#### **Selection.Select**

选择指定的文本。

**语法**

**express.Select()**

*express*   一个代表 **Selection** 对象的变量。

**说明**

使用此方法后，可使用 **Selection** 对象处理所选项目。有关详细信息，请参阅 处理 Selection 对象。

#### **Selection.SelectCell**

选择包含当前选定内容的匹配单元格。

**语法**

**express.SelectCell()**

*express*   一个代表 **Selection** 对象的变量。

**说明**

使用此方法，要求当前选定内容必须包含在一个单独单元格中。

#### **Selection.SelectColumn**

选定包含插入点的列，或者选定包含选定内容的所有列。

**语法**

**express.SelectColumn()**

*express*   一个代表 **Selection** 对象的变量。

**说明**

如果选定内容不在表格中，则会导致出错。

#### **Selection.SelectCurrentAlignment**

向前扩展选定部分，直到遇到另一种段落对齐方式为止。

**语法**

**express.SelectCurrentAlignment()**

*express*   一个代表 **Selection** 对象的变量。

**说明**

段落的对齐方式有四种：左对齐、居中、右对齐和两端对齐。

#### **Selection.SelectCurrentColor**

向前扩展选定内容，直至遇到另一种颜色的文字为止。

**语法**

**express.SelectCurrentColor()**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.SelectCurrentFont**

向前扩展选定内容，直至遇到另一种字体或字号。

**语法**

**express.SelectCurrentFont()**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.SelectCurrentIndent**

向前扩展选定内容，直至遇到具有另一种段落左右缩进量的文本为止。

**语法**

**express.SelectCurrentIndent()**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.SelectCurrentSpacing**

向前扩展选定内容，直至遇到具有另一种行间距的段落为止。

**语法**

**express.SelectCurrentSpacing()**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.SelectCurrentTabs**

向前扩展选定内容，直至遇到另一种制表位的段落为止。

**语法**

**express.SelectCurrentTabs()**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.SelectRow**

选定插入点所在的行，或者选择选定内容所在行。

**语法**

**express.SelectRow()**

*express*   一个代表 **Selection** 对象的变量。

**说明**

如果选定内容不在表格中，则会导致出错。

#### **Selection.SetRange**

设置所选内容的起始字符和结束字符的位置。

**语法**

**express.SetRange(Start, End)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                 |
| -------- | ------------- | ------------ | ------------------------ |
| *Start*  | 必选          | **Long**     | 所选内容的起始字符位置。 |
| *End*    | 必选          | **Long**     | 所选内容的结束字符位置。 |

**说明**

字符位置的值从文档该部分开头计起，起始值为 0（零）。将计算所有的字符，包括非打印字符和未显示的隐藏字符。

**SetRange** 方法用于重新定义现有 **Selection** 对象的起始和结束位置。此方法不同于 **Range** 方法，后者用于在给定起始和结束位置的情况下创建 **Range** 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例选择文档中的前 10 个字符。*/ Application.Selection.SetRange(0,10);  /*以下示例将所选内容扩展至文档末尾。*/ function test() { 	let sel = Application.Selection; 	sel.SetRange(sel.Start,ActiveDocument.Content.End); } ` |

#### **Selection.Shrink**

将所选内容缩减至下一级较小的文字单位。

**语法**

**express.Shrink()**

*express*   一个代表 **Selection** 对象的变量。

**说明**

此方法的单位级数如下：整篇文档、节、段落、句子、单词及插入点。

#### **Selection.ShrinkDiscontiguousSelection**

当所选内容包括多个不连续的所选内容时，取消全部选择，只保留最近选择的文本。

**语法**

**express.ShrinkDiscontiguousSelection()**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.Sort**

对指定的所选内容中的段落进行排序。

**语法**

**express.Sort(ExcludeHeader, FieldNumber, SortFieldType, SortOrder, FieldNumber2, SortFieldType2, SortOrder2, SortColumn2, Separator, FieldNumber3, SortFieldType3, SortOrder3, CaseSensitive, BidiSort, IgnoreThe, IgnoreKashida, IgnoreDiacritics, IgnoreHe, LanguageID, SubFieldNumber, SubFieldNumber2, SubFieldNumber3)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称**           | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------------ | ------------- | ------------ | ------------------------------------------------------------ |
| *ExcludeHeader*    | 可选          | **Variant**  | 如果该参数值为 True，则不对首行或段落标题进行排序。默认值为 False。 |
| *FieldNumber*      | 可选          | **Variant**  | 用于排序的第一个域。                                         |
| *SortFieldType*    | 可选          | **Variant**  | FieldNumber 的排序类型。可以是 WdSortFieldType 常量之一。默认值为 wdSortFieldAlphanumeric。根据您选择或安装的语言支持（例如，美国英语），某些 WdSortFieldType 常量可能不可用。 |
| *SortOrder*        | 可选          | **Variant**  | 对 FieldNumber 进行排序时使用的排序顺序。可以是一个 WdSortOrder 常量。默认值为 wdSortOrderAscending。 |
| *FieldNumber2*     | 可选          | **Variant**  | 用于排序的第二个域。                                         |
| *SortFieldType2*   | 可选          | **Variant**  | FieldNumber2 的排序类型。可以是 WdSortFieldType 常量之一。默认值为 wdSortFieldAlphanumeric。根据您选择或安装的语言支持（例如，美国英语），某些 WdSortFieldType 常量可能不可用。 |
| *SortOrder2*       | 可选          | **Variant**  | 对 FieldNumber2 进行排序时使用的排序顺序。可以是一个 WdSortOrder 常量。默认值为 wdSortOrderAscending。 |
| *SortColumn2*      | 可选          | **Variant**  | 如果该参数值为 True，则只对由 Selection 对象指定的列进行排序。 |
| *Separator*        | 可选          | **Variant**  | 域分隔符的类型。                                             |
| *FieldNumber3*     | 可选          | **Variant**  | 用于排序的第三个域。                                         |
| *SortFieldType3*   | 可选          | **Variant**  | FieldNumber3 的排序类型。可以是 WdSortFieldType 常量之一。默认值为 wdSortFieldAlphanumeric。根据您选择或安装的语言支持（例如，美国英语），某些 WdSortFieldType 常量可能不可用。 |
| *SortOrder3*       | 可选          | **Variant**  | 对 FieldNumber3 进行排序时使用的排序顺序。可以是 WdSortOrder 常量之一。默认值为 wdSortOrderAscending。 |
| *CaseSensitive*    | 可选          | **Variant**  | 如果该参数值为 True，则排序时区分大小写。默认值为 False。    |
| *BidiSort*         | 可选          | **Variant**  | 如果该参数值为 True，则基于从右向左语言的规则进行排序。根据您所选择或安装的语言支持（例如，美国英语），此参数可能不可用。 |
| *IgnoreThe*        | 可选          | **Variant**  | 如果该参数值为 True，则对从右向左的语言文本进行排序时忽略阿拉伯字符 alef lam。根据您所选择或安装的语言支持（例如，美国英语），此参数可能不可用。 |
| *IgnoreKashida*    | 可选          | **Variant**  | 如果该参数值为 True，则对从右向左的语言文本进行排序时忽略 Kashidas。根据您所选择或安装的语言支持（例如，美国英语），此参数可能不可用。 |
| *IgnoreDiacritics* | 可选          | **Variant**  | 如果该参数值为 True，则对从右向左的语言文本进行排序时忽略双向控制字符。根据您所选择或安装的语言支持（例如，美国英语），此参数可能不可用。 |
| *IgnoreHe*         | 可选          | **Variant**  | 如果该参数值为 True，则对从右向左的语言文本进行排序时忽略希伯来字符 he。根据您所选择或安装的语言支持（例如，美国英语），此参数可能不可用。 |
| *LanguageID*       | 可选          | **Variant**  | 指定排序语言。可以是 WdLanguageID 常量之一。                 |
| *SubFieldNumber*   | 可选          | **Variant**  | 用于排序的二级域编号。                                       |
| *SubFieldNumber2*  | 可选          | **Variant**  | 用于排序的二级域编号。                                       |
| *SubFieldNumber3*  | 可选          | **Variant**  | 用于排序的二级域编号。                                       |

#### **Selection.SortAscending**

按字母数字升序对段落或表格行进行排序。

**语法**

**express.SortAscending()**

*express*   一个代表 **Selection** 对象的变量。

**说明**

此方法提供针对包含数据列的邮件合并数据源进行排序的一种简化形式。对于大多数排序任务，请使用 Sort 方法。

#### **Selection.SortDescending**

以字母数字降序排列所选内容中的段落或表格行。

**语法**

**express.SortDescending()**

*express*   一个代表 **Selection** 对象的变量。

**说明**

此方法提供针对包含数据列的邮件合并数据源进行排序的一种简化形式。对于大多数排序任务，请使用 Sort 方法。

#### **Selection.SplitTable**

在选定内容第一行上面插入一个空段落。

**语法**

**express.SplitTable()**

*express*   一个代表 **Selection** 对象的变量。

**说明**

如果选定内容不在表格中的第一行，则将该表格拆分为两个表格。如果选定内容不在表格中，则会导致出错。

#### **Selection.StartOf**

将指定的区域或选定内容的开始位置移动或扩展至最近的指定文字单位的开头。该方法返回 **Long** 类型的值表明了区域或选定内容移动或扩展的字符数。如果是在文档中向后移动，则该方法返回负数

**语法**

**express.StartOf(Unit, Extend)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Unit*   | 必选          | **Number**   | 移动指定区域或选定内容的起始位置时所基于的单位。             |
| *Extend* | 必选          | **Number**   | 如果使用了 wdMove，则该区域或选定内容的末尾都移至指定单位的开头；如果使用了 wdExtend，则该区域或选定内容的开始扩展到指定单位的开头。 |

**说明**

如果指定区域或选定内容的开始已经位于指定单位的开始，则该方法不移动或扩展区域或选定内容。例如，如果选定内容位于一行的开头，则下面的示例返回 0（零），并且不改变选定内容。

#### **Selection.ToggleCharacterCode**

在 Unicode 字符和其相应的十六进制值之间切换选定内容。

**语法**

**express.ToggleCharacterCode()**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.TypeBackspace**

删除折叠的选定内容（即一个插入点）前面的字符

**语法**

**express.TypeBackspace()**

*express*   一个代表 **Selection** 对象的变量。

**说明**

本方法与 Backspace 的功能相同。如果选定内容没有折叠为一个插入点，则删除选定内容。

#### **Selection.TypeParagraph**

插入一个新的空段落。

**语法**

**express.TypeParagraph()**

*express*   一个代表 **Selection** 对象的变量。

**说明**

本方法与 Enter 的功能相同。如果选定内容没有折叠为一个插入点，则新段落取代选定内容。

用 **InsertParagraphAfter** 或 **InsertParagraphBefore** 方法可插入一个新段而不删除选定内容。

#### **Selection.TypeText**

插入指定的文本。

**语法**

**express.TypeText(Text)**

*express*   一个代表 **Selection** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**       |
| -------- | ------------- | ------------ | -------------- |
| *Text*   | 必选          | **String**   | 要插入的文字。 |

**说明**

如果 **ReplaceSelection** 属性为 **True**，则用指定文本替换选定内容。如果 **ReplaceSelection** 为 **False**，则在选定内容之前插入指定的文本。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果将 WPS 设置为用键入文本替换选定文本，则本示例在插入“Hello”前先折叠选定内容。这样可以避免现有文档文本被替换。*/ function test() { 	let sel = Application.Selection; 	if(Application.Options.ReplaceSelection) { 		sel.Collapse(wdCollapseStart); 		sel.TypeText("Hello") 	} }  /*本示例插入“Title”和一个新段落。*/ function test() { 	let sel = Application.Selection; 	sel.TypeText("Title"); 	sel.TypeParagraph() }` |

#### **Selection.WholeStory**

扩展某一所选内容，使其包括整个文章。

**语法**

**express.WholeStory()**

*express*   一个代表 **Selection** 对象的变量。

**成员属性**

#### **Selection.Active**

如果指定窗口或窗格中的选定内容处于活动状态，则该属性值为 **True**。**Boolean** 类型，只读。

**语法**

**express.Active**

*express*   一个代表 **Selection** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//示例将活动窗口拆分为两个窗格，打印第一个窗格中的选定内容是否处于活动状态。  function test()  {      Application.ActiveDocument.ActiveWindow.Split = true;      let bAc = Application.ActiveDocument.ActiveWindow.Panes.Item(1).Selection.Active;       alert(bac);  }` |

#### **Selection.BookmarkID**

返回位于选定内容开始位置的书签编号。**Long** 类型，只读。

**语法**

**express.BookmarkID**

*express*   一个代表 **Selection** 对象的变量。

**说明**

果没有相应的书签，则返回 0（零）。编号对应于书签在文档中的位置，1 对应于第一个书签，2 对应于第二个书签，依此类推。

**示例**

| 示例代码复制                       |
| ---------------------------------- |
| `Application.Selection.BookmarkID` |

#### **Selection.Bookmarks**

返回一个 **Bookmarks** 集合。该集合代表某一文档、区域或选定内容中的所有书签。只读。

**语法**

**express.Bookmarks**

*express*   一个代表 **Selection** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//获取当前选取的书签个数 Application.Selection.Bookmarks.Count` |

#### **Selection.Borders**

返回一个 **Borders** 集合，该集合代表指定对象的所有边框。

**语法**

**express.Borders**

*express*   一个代表 **Selection** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

#### **Selection.Cells**

返回一个 **Cells** 集合，该集合代表选定内容中的表格单元格。只读。

**语法**

**express.Cells**

*express*   一个代表 **Selection** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//将前单元格的背景色设置为绿色。 Application.Selection.Cells.Item(1).Shading.BackgroundPatternColorIndex = 3` |

#### **Selection.Characters**

返回一个 **Characters** 集合，该集合代表文档、区域或选定内容中的字符。只读。

**语法**

**express.Characters**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.ChildShapeRange**

返回一个 **ShapeRange** 集合，该集合代表选定区域中包含的子图形。

**语法**

**express.ChildShapeRange**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.Columns**

返回一个 **Columns** 集合，该集合代表所选内容中的所有表格列。只读。

**语法**

**express.Columns**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.ColumnSelectMode**

如果列选定模式处于活动状态，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.ColumnSelectMode**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.Comments**

返回一个 **Comments** 集合，该集合代表指定批注中的所有批注。只读。

**语法**

**express.Comments**

*express*   一个代表 **Selection** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//获取当前的选取的批注数量 Application.Selection.Comments.Count` |

#### **Selection.Creator**

回一个 32 位整数，该整数指出用于创建指定对象的应用程序。**Long** 类型，只读。

**语法**

**express.Creator**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.Document**

返回一个 **Document** 对象，该对象与指定的选定内容相关。只读。

**语法**

**express.Document**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.Editors**

回一个 **Editors** 对象，该对象代表已授权修改文档中选定内容的所有用户。

**语法**

**express.Editors**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.End**

返回或设置选定内容的结束字符的位置。**Long** 类型，可读写。

**语法**

**express.End**

*express*   一个代表 **Selection** 对象的变量。

**说明**

如果将该属性设置为小于 **Start** 属性的值，则 **Start** 属性将设置为同一值（即 **Start** 与 **End** 属性值相等）。

[**Selection** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Selection/Selection%20.htm#jsObject_Selection)对象具有起始位置和结束位置。结束位置是距离文字部分最远的点。该属性返回相对于文字部分的开始的结束字符位置。文档主要文字部分 (**wdMainTextStory**) 的起始字符位置为 0。通过设置该属性可以更改选定内容的大小。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例检索选定内容的结束位置。该值用于创建一个区域，以便在选定内容之后插入一个字段。*/ function test() { 	let sel = Application.Selection; 	let pos = sel.End; 	let myRange = Application.ActiveDocument.Range(pos, pos) 	Application.ActiveDocument.Fields.Add(myRange,wdFieldAuthor); } ` |

#### **Selection.EndnoteOptions**

返回一个 **EndnoteOptions** 对象，该对象代表选定内容中的尾注。

**语法**

**express.EndnoteOptions**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.Endnotes**

返回一个 **Endnotes** 集合，该集合代表选定内容中包含的所有尾注。只读。

**语法**

**express.Endnotes**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.EnhMetaFileBits**

返回一个 **Variant** 类型的值，该值代表选定文本或文本区域的显示方式的图片表示形式。

**语法**

**express.EnhMetaFileBits**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.ExtendMode**

如果“扩展”模式处于活动状态，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.ExtendMode**

*express*   一个代表 **Selection** 对象的变量。

**说明**

如果“扩展”模式处于活动状态，则默认情况下，下列方法的 *Extend* 参数为 **True**：**EndKey**、**HomeKey**、**MoveDown**、**MoveLeft**、**MoveRight** 和 **MoveUp**。另外字母“EXT”会显示在状态栏。

该属性只能在运行时设置；不能在“即时”模式中设置该属性。该属性不影响 **EndOf** 方法和 **StartOf** 方法的 *Extend* 参数。

#### **Selection.Fields**

返回一个只读 **Fields** 集合，该集合代表选定内容中的所有域。

**语法**

**express.Fields**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.Find**

返回一个 **Find** 对象，该对象包含查找操作所需的条件。只读。

**语法**

**express.Find**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.FitTextWidth**

返回或设置 WPS 在当前选定内容中填入文字的宽度（使用当前度量单位）。**Single** 类型，可读写。

**语法**

**express.FitTextWidth**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.Flags**

返回或设置选定内容的属性。**WdSelectionFlags** 类型，可读写。

**语法**

**express.Flags**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.Font**

返回或设置一个 **Font** 对象，该对象代表指定对象的字符格式。可读写。

**语法**

**express.Font**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.FootnoteOptions**

返回一个 **FootnoteOptions** 对象，该对象代表选定内容中的脚注。

**语法**

**express.FootnoteOptions**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.Footnotes**

返回一个 **Footnotes** 集合，该集合代表区域、选定内容或文档中的所有脚注。只读。

**语法**

**express.Footnotes**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.FormattedText**

返回或设置一个 **Range** 对象，该对象包含指定区域或选定内容中进行过格式编排的文字。可读写。

**语法**

**express.FormattedText**

*express*   一个代表 **Selection** 对象的变量。

**说明**

此属性返回 **Range** 对象，以及指定的区域或所选内容中的字符格式和文本。如果在区域或所选内容中有一个段落标记，则 **Range** 对象中包含段落格式。在设置此属性时，区域中的文本会被格式文本替换。如果不需要替换现有的文本，则在使用本属性之前使用 **Collapse** 方法

#### **Selection.FormFields**

返回一个 **FormFields** 集合，该集合代表选定内容中所有窗体域。只读。

**语法**

**express.FormFields**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.Frames**

返回一个 **Frames** 集合，该集合代表选定内容中的所有框架。只读。

**语法**

**express.Frames**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.HasChildShapeRange**

如果选定内容包含子图形，则该属性值为 **True**。**Boolean** 类型，只读。

**语法**

**express.HasChildShapeRange**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.HeaderFooter**

为指定的选定内容返回 **HeaderFooter** 对象。只读。

**语法**

**express.HeaderFooter**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.HTMLDivisions**

返回一个 **HTMLDivisions** 对象，该对象代表 Web 文档中的一个 HTML 区域。

**语法**

**express.HTMLDivisions**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.Hyperlinks**

返回一个 **Hyperlinks** 集合，该集合代表指定选定内容中的所有超链接。只读。

**语法**

**express.Hyperlinks**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.Information**

返回有关指定的选定内容的信息。**Variant** 类型，只读。

**语法**

**express.Information**

*express*   一个代表 **Selection** 对象的变量。

**说明**

| **名称** | **必选/可选** | **数据类型**      | **说明**   |
| -------- | ------------- | ----------------- | ---------- |
| *Type*   | 必选          | **WdInformation** | 信息类型。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示当前页码和活动文档的总页数。*/ function test() { 	let sel = Application.Selection; 	alert("The selection is on page " + sel.Information(wdActiveEndPageNumber) +" of page " +  sel.Information(wdNumberOfPagesInDocument)) }  /*如果选定内容位于一个表格中，则本示例选定该表格。*/ function test() { 	let sel = Application.Selection; 	if(sel.Information(wdWithInTable)) { 		sel.Tables.Item(1).Select(); 	} } ` |

#### **Selection.InlineShapes**

返回一个 **InlineShapes** 集合，该集合代表选定内容中的所有 **InlineShape** 对象。只读。

**语法**

**express.InlineShapes**

*express*   一个代表 **Selection** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示活动文档中的图形和内嵌图形的数目。*/ function test() { 	let doc = Application.ActiveDocument; 	let msg = "InlineShape = " + doc.InlineShapes.Count + "\nShapes = " + doc.Shapes.Count; 	alert(msg) }` |

#### **Selection.IPAtEndOfLine**

如果插入点位于行（该行折到了下一行）的末尾，则该属性值为 **True**。**Boolean** 类型，只读。

**语法**

**express.IPAtEndOfLine**

*express*   一个代表 **Selection** 对象的变量。

**说明**

如果为 **False**，则选定内容不折叠；或者插入点不在行尾；或者插入点位于段落标记之前。

#### **Selection.IsEndOfRowMark**

如果指定的选定内容或区域折叠且位于表格行的结束标志处，则该属性值为 **True**。**Boolean** 类型，只读。

**语法**

**express.IsEndOfRowMark**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.LanguageDetected**

返回或设置一个指定 WPS 是否已经检测过选定文本的语言的 **Boolean** 类型。

**语法**

**express.LanguageDetected**

*express*   一个代表 **Selection** 对象的变量。

**说明**

检查以前所有语言检测结果的 **LanguageID** 属性。

调用 **DetectLanguage** 方法时，**LanguageDetected** 属性被设置为 **True**。要重新检测指定文本的语言，必须先将 **LanguageDetected** 属性设置为 **False**。

#### **Selection.LanguageID**

返回或设置指定对象的语言。可读写。

**语法**

**express.LanguageID**

*express*   一个代表 **Selection** 对象的变量。

**说明**

对于自定义词典，在指定 **LanguageID** 属性之前，必须将 **LanguageSpecific** 属性设置为 **True**。特定于语言的自定义词典只能检查用该语言格式编写的文本。

根据您选择或安装的语言支持（如，美国英语）的不同，以上列出的某些常量可能不可用。

#### **Selection.LanguageIDFarEast**

为指定的对象返回或设置东亚语言。**WdLanguageID** 类型，可读写。

**语法**

**express.LanguageIDFarEast**

*express*   一个代表 **Selection** 对象的变量。

**说明**

推荐使用本方法，来返回或设置用 WPS 中文版创建的文档中东亚文字的语言。

#### **Selection.LanguageIDOther**

返回或设置指定对象的语言。**WdLanguageID** 类型，可读写。

**语法**

**express.LanguageIDOther**

*express*   一个代表 **Selection** 对象的变量。

**说明**

推荐使用本方法，来设置或返回在 WPS 从右向左语言版本所创建的文档中西文文字所用的语言。

#### **Selection.NoProofing**

如果拼写和语法检查程序忽略指定文本，则该属性值为 **True**。如果只将某些指定文本的 **NoProofing** 属性设为 **True**，则会返回 **wdUndefined**。**Long** 类型，可读写。

**语法**

**express.NoProofing**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.OMaths**

返回一个 **OMaths** 集合，该集合代表当前选定区域内的 **OMath** 对象。只读。

**语法**

**express.OMaths**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.Orientation**

在启用“文字方向”功能后返回或设置选定内容中文字的方向。**WdTextOrientation** 类型，可读写。

**语法**

**express.Orientation**

*express*   一个代表 **Selection** 对象的变量。

**说明**

某些 **WdTextOrientation** 常量可能不可用，具体取决于所选择或安装的语言支持（例如，美国英语）。

您可以设置文本框架或者恰好位于文本框架内的选定内容的方向。有关文本框架和文本框之间的区别的信息，请参阅 **TextFrame** 对象。

#### **Selection.PageSetup**

返回一个 **PageSetup** 对象，该对象与指定选定内容相关联。

**语法**

**express.PageSetup**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.ParagraphFormat**

返回或设置一个 **ParagraphFormat** 对象，该对象代表指定选定内容的段落设置。可读写。

**语法**

**express.ParagraphFormat**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.Paragraphs**

返回一个 **Paragraphs** 集合，该集合代表指定选定内容中的所有段落。只读。

**语法**

**express.Paragraphs**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.Parent**

返回一个 **Object** 类型的值，该值代表指定 **Selection** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.PreviousBookmarkID**

返回位于指定的所选内容或区域中或之前的最后一个书签的编号。如果没有相应的书签，则返回 0（零）。**Long** 类型，只读。

**语法**

**express.PreviousBookmarkID**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.Range**

返回一个 **Range** 对象，该对象代表指定对象所含的部分文档。

**语法**

**express.Range**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.Rows**

返回一个 **Rows** 集合，该集合代表某个区域、选定内容或表格中所有的表格行。只读。

**语法**

**express.Rows**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.Sections**

返回一个 **Sections** 集合，该集合代表指定选定内容中的各节。只读。

**语法**

**express.Sections**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.Sentences**

返回一个 **Sentences** 集合，该集合代表选定内容中的所有句子。只读。

**语法**

**express.Sentences**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.Shading**

返回一个 **Shading** 对象，该对象代表指定选定内容的底纹格式。

**语法**

**express.Shading**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.Start**

返回或设置选定内容的起始字符位置。**Long** 类型，可读写。

**语法**

**express.Start**

*express*   一个代表 **Selection** 对象的变量。

**说明**

**Selection** 对象具有起始字符位置和结束字符位置。起始字符位置是指距文档开头部分最近的字符位置。如果将该属性的值设置为大于 **End** 属性的值，则 **End** 属性值会设置为与 **Start** 属性值相同。

该属性返回起始字符相对于文档开头部分的位置。文本主体部分 (**wdMainTextStory**) 的起始字符位置为 0（零）。通过设置该属性可以更改选定内容、区域或书签的大小。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例通过起始和结束字符位置确定选定内容的长度。*/ let SelLength = Application.Selection.End - Application.Selection.Start` |

#### **Selection.StartIsActive**

如果选定内容的开始部分处于活动状态，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.StartIsActive**

*express*   一个代表 **Selection** 对象的变量。

**说明**

如果选定内容没有折叠为插入点，则其开始部分和结束部分都处于活动状态。如果调用下列方法，选定内容的活动结尾将会移动：**EndKey**、**Extend**（带 *Characters* 参数）、**HomeKey**、**MoveDown**、**MoveLeft**、**MoveRight** 和 **MoveUp**。

该属性相当于使用带有 **wdSelStartActive** 常量的 **Flags** 属性。但是，使用 **Flags** 属性需要二元运算，这比使用 **StartIsActive** 属性更复杂。

#### **Selection.StoryLength**

返回包含指定选定内容的文章所含的字符数。**Long** 类型，只读。

**语法**

**express.StoryLength**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.StoryType**

返回指定选定内容的文章类型。**WdStoryType** 类型，只读。

**语法**

**express.StoryType**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.Style**

该属性返回或设置用于指定对象的样式。若要设置该属性，请指定样式的本地名称、一个整数值、一个 **WdBuiltinStyle** 常量或一个代表该样式的对象。若要获取有效常量的列表，请查询“Visual Basic 对象浏览器”。**Variant** 类型，可读写。

**语法**

**express.Style**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.Tables**

返回一个 **Tables** 集合，该集合代表指定选定内容中的所有表格。只读。

**语法**

**express.Tables**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.Text**

返回或设置指定的选定内容中的文本。**String** 类型，可读写。

**语法**

**express.Text**

*express*   一个代表 **Selection** 对象的变量。

**说明**

**Text** 属性返回选定内容的无格式纯文本。设置该属性，可替换该区域或选定内容的文本。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示选定区域中的文本。如果没有选中任何内容，则显示插入点之后的字符。*/ alert(Application.Selection.Test); ` |

#### **Selection.TopLevelTables**

回一个 **Tables** 集合，该集合代表当前选定内容最外部嵌套层上的表格。只读。

**语法**

**express.TopLevelTables**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.Type**

返回选择类型。**WdSelectionType** 类型，只读。

**语法**

**express.Type**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.WordOpenXML**

返回一个 **String** 类型的值，该值以 WPS Open XML 格式表示选定内容中包含的 XML。只读。

**语法**

**express.WordOpenXML**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.Words**

返回一个 **Words** 集合，该集合代表选定内容中的所有字词。只读。

**语法**

**express.Words**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.XML**

返回一个 **String** 类型的值，该值代表指定对象中的 XML 文本。

**语法**

**express.XML**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.XMLNodes**

返回一个 **XMLNodes** 集合，该集合代表选定内容中所有 XML 元素的集合，包括那些仅部分位于选定内容中的元素。

**语法**

**express.XMLNodes**

*express*   一个代表 **Selection** 对象的变量。

#### **Selection.XMLParentNode**

返回一个 **XMLNode** 对象，该对象代表选定内容的父节点。

**语法**

**express.XMLParentNode**

*express*   一个代表 **Selection** 对象的变量。

适用环境：web

适用平台：windows/linux