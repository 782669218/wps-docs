#### **Options**



代表 WPS 中的应用程序和文档选项。**Options** 对象的许多属性与**“选项”**对话框中的项对应。

**说明**

使用 **Options** 属性可返回 **Options** 对象。以下示例设置了 WPS 的三个应用程序选项。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     Application.Options.AllowDragAndDrop = true     Application.Options.ConfirmConversions = false     Application.Options.MeasurementUnit = wdPoints }` |

**方法**

|                                                              | 名称                | 说明                                                         |
| ------------------------------------------------------------ | ------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **DefaultFilePath** | 返回或设置各项（例如文档、模板和图形）的默认文件夹。**String** 类型，可读写。 |

**属性**

|                                                              | 名称                                            | 说明                                                         |
| ------------------------------------------------------------ | ----------------------------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AddBiDirectionalMarksWhenSavingTextFile**     | 如果 WPS 在将文档保存为文本文件时添加双向控制字符，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AddControlCharacters**                        | 如果 WPS 在剪切和复制文本时添加双向控制字符，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AddHebDoubleQuote**                           | 如果 WPS 会将数字格式括在双引号 (") 中，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AllowAccentedUppercase**                      | 如果当一个法语字符改成大写字母时重音仍保留，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AllowClickAndTypeMouse**                      | 如果启用“即点即输”功能，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AllowCombinedAuxiliaryForms**                 | 如果 WPS 在朝鲜语文档中执行拼写检查时将忽略辅助动词形式，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AllowCompoundNounProcessing**                 | 如果 WPS 在朝鲜语文档中执行拼写检查时将忽略复合名词，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AllowDragAndDrop**                            | 如果可以使用拖动来移动或复制所选内容，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AllowOpenInDraftView**                        | 返回或设置 **Boolean** 值，该值表示是否允许用户在草稿视图中打开文档。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AllowPixelUnits**                             | 如果 WPS 用像素作为默认度量单位（这是 HTML 功能所支持的度量单位），则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AllowReadingMode**                            | 如果为 **True**，则表明 WPS 将文档在“阅读版式”视图中打开。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AlwaysUseClearType**                          | 返回或设置 **Boolean** 值，该值表示是否使用 ClearType 来显示菜单、功能区和对话框文本中的字体。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AnimateScreenMovements**                      | 如果 WPS 激活动态鼠标功能、使用动态光标，并且激活后台保存与查找和替换操作之类的动态操作效果，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**                                 | 返回一个代表 WPS 应用程序的 **Application** 对象。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ApplyFarEastFontsToAscii**                    | 如果 WPS 将中文字体应用于西文文字，则该属性值为 **True**。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ArabicMode**                                  | 返回或设置阿拉伯语的拼写检查模式。**WdAraSpeller** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ArabicNumeral**                               | 返回或设置一篇阿拉伯语文档的数字样式。**WdArabicNumeral** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoCreateNewDrawings**                       | 如果为 **True**，则 WPS 在绘图画布上画出新创建的图形。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoFormatApplyBulletedLists**                | 如果当 WPS 自动设置文档或区域格式时，会用**“格式”**菜单的**“项目符号和编号”**对话框中的项目符号替换列表段落的开始字符（例如星号、连字符和大于号符号），则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoFormatApplyFirstIndents**                 | 如果为 **True**，则 WPS 在自动设置文档或区域的格式时，用首行缩进替换在段落开始处输入的空格。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoFormatApplyHeadings**                     | 如果当 WPS 自动设置文档或区域格式时，将自动为标题设置样式，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoFormatApplyLists**                        | 如果当 WPS 自动设置文档或区域格式时，自动为列表应用样式，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoFormatApplyOtherParas**                   | 如果当 WPS 自动设置文档或区域格式时，自动为非标题项或列表项的段落设置样式，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoFormatAsYouTypeApplyBorders**             | 如果按下 Enter 后，三个或更多的连字符 (-)、等号 (=) 或下划线 (_) 将自动替换成指定的边框线，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoFormatAsYouTypeApplyBulletedLists**       | 如果在键入时用**“格式”**菜单的**“项目符号和编号”**对话框中的项目符号替换当前的项目符号（例如星号、连字符和大于号），则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoFormatAsYouTypeApplyClosings**            | 如果为 **True**，则 WPS 会在键入时自动为信函的结束语应用“结束语”样式。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoFormatAsYouTypeApplyDates**               | 如果为 **True**，则 WPS 自动为键入的日期应用“日期”样式。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoFormatAsYouTypeApplyFirstIndents**        | 如果为 **True**，则 WPS 在键入时自动用首行缩进替代段落开始处输入的空格。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoFormatAsYouTypeApplyHeadings**            | 如果在键入标题时，自动为标题设置样式，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoFormatAsYouTypeApplyNumberedLists**       | 如果根据键入的内容，以**“格式”**菜单的“项目符号和编号”对话框中的编号方案将段落自动设置为编号列表，则该属性值为 **True**。例如，如果一个段落以“1.1”和一个制表符开始，则当按下 Enter 后 WPS 将自动插入“1.2”和一个制表符。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoFormatAsYouTypeApplyTables**              | 如果为 **True**，则当您键入一个加号、一串连字符、另一个加号等，然后按下 Enter 后， WPS 将自动创建一张表格。加号变为列边框，连字符变为列宽度。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoFormatAsYouTypeAutoLetterWizard**         | 如果为 **True**，则在用户输入信函称呼或结束语时，WPS 会自动启动“英文信函向导”。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoFormatAsYouTypeDefineStyles**             | 如果 WPS 自动根据手动设置的格式新建一种样式，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoFormatAsYouTypeDeleteAutoSpaces**         | 如果为 **True**，则 WPS 会在用户键入时自动删除日文和拉丁文文字之间插入的空格。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoFormatAsYouTypeFormatListItemBeginning**  | 如果 WPS 将列表项开始处的字符格式重复应用于下一个列表项，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoFormatAsYouTypeInsertClosings**           | 如果为 **True**，则当用户输入备忘录标题时，WPS 会自动插入相应的备忘录结束语。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoFormatAsYouTypeInsertOvers**              | 如果为 **True**，则当用户输入“ ![img]()”或“ ![img]()”时，WPS 将自动插入“ ![img]() ![img]()”。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoFormatAsYouTypeMatchParentheses**         | 如果为 **True**，WPS 会自动更正不正确成对的括号，可读写。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoFormatAsYouTypeReplaceFarEastDashes**     | 如果为 **True**，WPS 会自动更正长元音和划线，可读写。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoFormatAsYouTypeReplaceFractions**         | 如果在您键入时，将使用当前字符集的分数替换键入的分数，则该属性值为 **True**。例如，将“1/2”替换为“?”。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoFormatAsYouTypeReplaceHyperlinks**        | 如果本属性值为 **True**，则 WPS 会在键入时自动将电子邮件地址、服务器和共享名称（即 UNC 路径）和 Internet 地址（即 URL）改为超链接。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoFormatAsYouTypeReplaceOrdinals**          | 如果为 **True**，则键入时， WPS 会用相同字母的上标格式替换序数后缀“st”、“nd”、“rd”和“th”。例如，“1st”被替换为“1”后接上标“st”。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoFormatAsYouTypeReplacePlainTextEmphasis** | 如果键入时，手动强调的字符将自动替换为相应字符格式，则该属性值为 **True**。例如，将“*bold*”改为“**bold**”、将“_underline_”改为“underline”。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoFormatAsYouTypeReplaceQuotes**            | 如果键入时自动将所键入的直引号替换为弯引号，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoFormatAsYouTypeReplaceSymbols**           | 如果键入的两个连续的连字符 (--) 将替换为短划线 (–) 或长划线 (—)，则该属性值为 **True**。**Boolean** 类型，可读写。如果键入连字符时带有前导和尾部空格，则 WPS 将连字符替换为短划线。如果没有尾部空格，则连字符替换为长划线。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoFormatDeleteAutoSpaces**                  | 如果 WPS 在自动设置文档或区域的格式时，删除日文和西文文字之间插入的空格，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoFormatMatchParentheses**                  | 如果 WPS 在自动设置文档或区域的格式时，会更正不恰当匹配的括号，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoFormatPlainTextWordMail**                 | 如果在 WPS 中打开纯文本格式的电子邮件时， WPS 将自动进行格式设置，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoFormatPreserveStyles**                    | 如果当 WPS 自动设置文档或区域格式时，保留先前所用的样式，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoFormatReplaceFarEastDashes**              | 如果 WPS 在自动设置文档或区域的格式时，会更正长元音和划线的使用，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoFormatReplaceFractions**                  | 如果当 WPS 自动设置文档或区域格式时，会用当前字符集中的分数替换所键入的分数，则该属性值为 **True**。例如，将“1/2”替换为“?”。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoFormatReplaceHyperlinks**                 | 如果为 **True**，则 WPS 在对某一文档或范围自动套用格式时，将对其中的电子邮件地址、服务器和共享名称（即 UNC 路径）和 Internet 地址（即 URL）自动进行格式设置。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoFormatReplaceOrdinals**                   | 如果为 **True**，则当 WPS 自动设置文档或区域格式时，用相同字母的上标格式替换序数后缀“st”、“nd”、“rd”和“th”。例如，将“1st”替换为“1”后接上标“st”。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoFormatReplacePlainTextEmphasis**          | 如果当 WPS 自动设置文档或区域的格式时，将手动强调的字符替换为相应的字符格式（例如，“*bold*”更改为“bold”，“_underline_”更改为“underline”），则该属性值为 **True**。**Boolean** 类型，可读写。如果当 WPS 自动设置文档或区域的格式时，将手动强调的字符替换为相应的字符格式（例如，“*bold*”更改为“bold”，“_underline_”更改为“underline”），则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoFormatReplaceQuotes**                     | 如果当 WPS 自动设置文档或区域格式时，直引号会自动更改为弯引号，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoFormatReplaceSymbols**                    | 如果当 WPS 自动设置文档或区域的格式时，将两个连续的连字号 (--) 替换为短破折号 ( |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoKeyboardSwitching**                       | 如果 WPS 随时自动切换键盘语言并与键入的内容匹配，则该属性值为 **True**。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoWordSelection**                           | 如果在拖动时，一次选定一个单词，而不是选定一个字符，则该属性为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BackgroundOpen**                              | 如果为 **True**，WPS 在后台打开 Web 文档。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BackgroundSave**                              | 如果为 **True**，则 WPS 在后台保存文档。当 WPS 在后台保存文档时，用户可以继续键入内容或执行命令。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BibliographySort**                            | 返回或设置一个 **String** 类型的值，该值代表在**“源管理器”**对话框中显示源的次序。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BibliographyStyle**                           | 返回或设置一个 **String** 类型的值，该值代表用于设置书目格式的样式的名称。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BrazilReform**                                | 返回或设置巴西葡萄牙语拼写器的模式。可读/写。                |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ButtonFieldClicks**                           | 返回或设置一个数字，该数字是运行 GOTOBUTTON 或 MACROBUTTON 域所需的点击次数（单击或双击）。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CheckGrammarAsYouType**                       | 如果 WPS 在键入时自动对文档进行语法检查并自动标记错误，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CheckGrammarWithSpelling**                    | 如果 WPS 在检查拼写的同时也检查语法，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CheckHangulEndings**                          | 如果 WPS 自动检测朝鲜文字结尾，并在将朝鲜文字转换为朝鲜文汉字过程中忽略朝鲜文字结尾，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CheckSpellingAsYouType**                      | 如果 WPS 在您键入的同时自动进行拼写检查并标记错误，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CommentsColor**                               | 该属性返回或设置一个 **WdColorIndex** 常量，代表文档中批注的颜色。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ConfirmConversions**                          | 如果 WPS 在打开或者插入一个非 WPS 文档或模板的文件之前，会显示一个**“转换文件”**对话框，则该属性值为 **True**。在**“转换文件”**对话框中，用户可以选择需要转换的文件格式。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ContextualSpeller**                           | 返回或设置 **Boolean** 值，该值表示是否使用上下文拼写检查器，基于单词上下文和它前后的单词来检查拼写。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ConvertHighAnsiToFarEast**                    | 如果在打开文档时，WPS 将与东亚字体相关的文字转换为相应的字体，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CreateBackup**                                | 如果 WPS 在每次保存文档时都建立一个备份，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**                                     | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CtrlClickHyperlinkToOpen**                    | 如果 WPS 需要单击同时按住 Ctrl 键来打开超链接，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CursorMovement**                              | 返回或设置在双向文字中插入点移动的方式。**WdCursorMovement** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DefaultBorderColor**                          | 返回或设置用于新的 **Border** 对象的默认 24 位颜色。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DefaultBorderColorIndex**                     | 返回或设置边框的默认线条颜色。**WdColorIndex** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DefaultBorderLineStyle**                      | 返回或设置默认边框的样式。**WdLineStyle** 类型，可读写。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DefaultBorderLineWidth**                      | 返回或设置边框的默认线条宽度。**WdLineWidth** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DefaultEPostageApp**                          | 设置或返回一个 **String** 类型的值，该值代表默认的电子邮政应用程序的路径和文件名。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DefaultFilePath**                             | 返回或设置各项（例如文档、模板和图形）的默认文件夹。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DefaultHighlightColorIndex**                  | 返回或设置突出显示文本所用的颜色，用**“格式”**工具栏的**“突出显示”**按钮可将文本格式设为突出显示。**WdColorIndex** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DefaultOpenFormat**                           | 返回或设置打开文档的默认转换器。可为由 **OpenFormat** 属性返回的一个数字，也可为下列 **WdOpenFormat** 常量之一。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DefaultTextEncoding**                         | 返回或设置一个 **MsoEncoding** 常量，该常量代表 WPS 用于所有保存为编码文本文件的文档的代码页或字符集。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DefaultTray**                                 | 返回或设置默认纸盒，以便打印机使用该纸盒打印文档。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DefaultTrayID**                               | 返回或设置用于打印文档的默认纸盒。**WdPaperTray** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DeletedCellColor**                            | 返回或设置一个 **WdCellColor** 常量，该常量代表删除的单元格的颜色。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DeletedTextColor**                            | 当激活修订选项时，返回或设置要删除的文字的颜色。**WdColorIndex** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DeletedTextMark**                             | 当激活修订选项时，返回或设置要删除的文字的格式。**WdDeletedTextMark** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DiacriticColorVal**                           | 返回或设置用于从右向左语言文档中的音调符号的 24 位颜色。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisableFeaturesbyDefault**                    | 如果为 **True**，则 WPS 将在所有文档中禁用 **DisableFeaturesIntroducedAfterbyDefault** 中指定的 WPS 版本后引入的所有功能。默认值为 **False**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisableFeaturesIntroducedAfterbyDefault**     | 对所有文档禁用指定版本后引入的所有功能。**WdDisableFeaturesIntroducedAfter** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayGridLines**                            | 如果 WPS 显示文档网格，则该属性值为 **True**。该属性相当于**“视图”**菜单中的**“网格线”**命令。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayPasteOptions**                         | 如果为 **True**，则 WPS 直接在新粘贴的文本下显示**“粘贴选项”**按钮。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplaySmartTagButtons**                      | 如果为 **True**，则 WPS 在鼠标指针置于智能标记上方时显示一个按钮。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DocumentViewDirection**                       | 返回或设置整篇文档的对齐方式和阅读顺序。**WdDocumentViewDirection** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DoNotPromptForConvert**                       | 设置或返回一个 **Boolean**，它代表在为处于兼容模式下的文档调用“转换”命令时是否出现一个警告对话框。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EnableHangulHanjaRecentOrdering**             | 如果 WPS 在进行朝鲜文字和朝鲜文汉字间的转换时，在建议列表的顶端显示最近用过的单词，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EnableLegacyIMEMode**                         | 返回或设置一个 **Boolean** 类型的值，该值代表是否启用旧式 IME 模式。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EnableLivePreview**                           | 设置或返回 **Boolean** 值，该值表示显示还是隐藏在使用支持预览的库时所出现的库预览。如果为 **True**，将在应用命令之前在文档中显示预览。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EnableMisusedWordsDictionary**                | 如果 WPS 在检查文档的拼写和语法时，也检查误用的单词，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EnableSound**                                 | 如果 WPS 在导致出错时使计算机发出声音进行响应，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EnvelopeFeederInstalled**                     | 如果当前的打印机为信封指定了送纸器，则该属性值为 **True**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FormatScanning**                              | 如果为 **True**，则 WPS 对文档中的所有格式进行跟踪。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FrenchReform**                                | 返回或设置一个 **WdFrenchSpeller** 常量，该常量代表要对语言格式设置为法语的文本区域使用哪个拼写词典。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **GridDistanceHorizontal**                      | 当在新文档中绘制、移动和调整自选图形或东亚语言字符时，返回或设置 WPS 使用的不可见网格线的水平间距。**Single** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **GridDistanceVertical**                        | 在新文档中绘制、移动和调整自选图形或东亚语言字符时，返回或设置 WPS 所用的不可见网格线间的垂直间距。**Single** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **GridOriginHorizontal**                        | 返回或设置不可见网格相对于页面左边的起点位置，当在新文档中绘制、移动和调整自选图形或东亚语言字符时将使用该网格。**Single** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **GridOriginVertical**                          | 返回或设置不可见网格相对于页面顶边的起点位置，当在新文档中绘制、移动和调整自选图形或东亚语言字符时将使用该网格。**Single** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HangulHanjaFastConversion**                   | 如果 WPS 在朝鲜文字和朝鲜文汉字间进行转换时，自动转换具有单独建议的单词，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HebrewMode**                                  | 返回或设置希伯来语拼写检查工具的模式。**WdHebSpellStart** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HideChartDraftModeNotification**              | 如果在草图模式下插入的图表上隐藏**“草图模式”**通知标签，则该属性值为**“True”**。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IgnoreInternetAndFileAddresses**              | 如果为 **True**，则进行拼写检查时忽略文件扩展名、MS-DOS 路径、电子邮件地址、服务器名称、共享名（也称为 UNC 路径）和 Internet 地址（也称为 URL）。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IgnoreMixedDigits**                           | 如果进行拼写检查时忽略包含数字的单词，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IgnoreUppercase**                             | 如果进行拼写检查时忽略全部大写的单词，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IMEAutomaticControl**                         | 如果设置 WPS 自动打开和关闭日文输入法编辑器 (IME)，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **InlineConversion**                            | 如果 WPS 将中文输入法 (IME) 中的未确认字符串显示为已有（已确认的）字符串之间的插入项，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **InsertChartInDraftMode**                      | 如果使用草图模式在文档中插入图表，则该属性值为 **True**。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **InsertedCellColor**                           | 返回或设置一个 **WdCellColor** 常量，该常量代表插入的表格单元格的颜色。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **InsertedTextColor**                           | 当启用修订选项时，返回或设置插入文字的颜色。**WdColorIndex** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **InsertedTextMark**                            | 返回或设置启用修订时 WPS 设置插入文本格式的方式（**TrackRevisions** 属性为 **True**）。**WdInsertedTextMark** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **INSKeyForOvertype**                           | 如果可以用 Ins 键来打开和关闭“改写”，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **INSKeyForPaste**                              | 如果可用 INS 键来粘贴“剪贴板”的内容，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **InterpretHighAnsi**                           | 返回或设置高端字符译码方式。**WdHighAnsiText** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LabelSmartTags**                              | 如果为 **True**，WPS 使用智能标记信息标记文档中的文本。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LocalNetworkFile**                            | 如果在编辑保存在网络服务器上的文件时，WPS 在用户的计算机上创建文件的一个本地副本，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MapPaperSize**                                | 如果为 **True**，则自动调整采用其他国家/地区标准纸张大小（例如 A4）的文档格式，以便按照用户所在国家/地区的标准纸张大小（例如，Letter）正确打印该文档。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MatchFuzzyAY**                                | 如果在搜索过程中 WPS 将忽略 ![img]() 行和 ![img]() 行字符后“ ![img]()”和“ ![img]()”之间的差异，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MatchFuzzyBV**                                | 如果 WPS 将在搜索过程中忽略“ ![img]()”与“ ![img]() ![img]()”之间以及“ ![img]()”与“ ![img]() ![img]()”之间的差异，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MatchFuzzyByte**                              | 如果在搜索过程中 WPS 将忽略全角和半角字符（西文或日文）的差异，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MatchFuzzyCase**                              | 如果 WPS 在搜索过程中不区分字母大小写，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MatchFuzzyDash**                              | 如果 WPS 在搜索过程中忽略减号、划线以及长元音符号之间的差异，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MatchFuzzyDZ**                                | 如果 WPS 将在搜索过程中忽略“ ![img]()”与“ ![img]()”之间以及“ ![img]()”与“ ![img]()”之间的差异，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MatchFuzzyHF**                                | 如果 WPS 将在搜索过程中忽略“ ![img]() ![img]()”与“ ![img]() ![img]()”之间以及“ ![img]() ![img]()”与“ ![img]() ![img]()”之间的差异，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MatchFuzzyHiragana**                          | 如果 WPS 在搜索过程中将忽略平假名与片假名之间的差异，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MatchFuzzyIterationMark**                     | 如果 WPS 在搜索过程中将忽略重复的标记类型间的差异，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MatchFuzzyKanji**                             | 如果 WPS 在搜索过程中忽略标准和非标准日文汉字的差异，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MatchFuzzyKiKu**                              | 如果在搜索过程中 WPS 将忽略 ![img]() 行字符前“ ![img]()”和“ ![img]()”之间的差异，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MatchFuzzyOldKana**                           | 如果 WPS 在搜索过程中将忽略新旧假名字符之间的差异，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MatchFuzzyProlongedSoundMark**                | 如果在搜索过程中 WPS 将忽略短元音和长元音的差异，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MatchFuzzyPunctuation**                       | 如果在搜索过程中 WPS 忽略标点类型的差异，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MatchFuzzySmallKana**                         | 如果 WPS 在搜索过程中忽略双元音和双辅音的差异，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MatchFuzzySpace**                             | 如果 WPS 在搜索过程中忽略空格标记的差异，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MatchFuzzyTC**                                | 如果在搜索过程中 WPS 将忽略“ ![img]() ![img]()”、“ ![img]() ![img]()”与“ ![img]()”之间以及“ ![img]() ![img]()”与“ ![img]()”之间的差异，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MatchFuzzyZJ**                                | 如果 WPS 将在搜索过程中忽略“ ![img]()”与“ ![img]() ![img]()”之间以及“ ![img]()”与“ ![img]() ![img]()”之间的差异，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MeasurementUnit**                             | 设置或者返回 WPS 的标准度量单位。**WdMeasurementUnits** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MergedCellColor**                             | 返回或设置一个 **WdCellColor** 常量，该常量代表合并的表格单元格的颜色。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MonthNames**                                  | 返回或设置朝鲜文字和朝鲜文汉字之间的转换方向。**WdMonthNames** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MoveFromTextColor**                           | 返回或设置一个 **WdColorIndex** 常量，该常量代表所移动文本的颜色。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MoveFromTextMark**                            | 返回或设置一个 **WdMoveFromTextMark** 常量，该常量代表要用于所移动文本的修订标记的类型。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MoveToTextColor**                             | 返回或设置一个 **WdColorIndex** 常量，该常量代表所移动文本的颜色。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MoveToTextMark**                              | 返回或设置一个 **WdMoveToTextMark** 常量，该常量代表要用于所移动文本的修订标记的类型。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MultipleWordConversionsMode**                 | 返回或设置朝鲜文字和朝鲜文汉字之间的转换方向。**WdMultipleWordConversionsMode** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OMathAutoBuildUp**                            | 返回或设置一个 **Boolean** 类型的值，该值代表 WPS 是否自动将公式转换为专业格式。如果为 **True**，则表示 WPS 自动将公式转换为专业格式。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OMathCopyLF**                                 | 返回或设置一个 **Boolean** 类型的值，该值代表如何以纯文本的形式表示公式。如果为 **True**，则表示以线性格式表示公式。如果为 **False**，则表示以 MathML 格式表示公式。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OptimizeForWord97byDefault**                  | 如果 WPS 将禁用所有不兼容的格式，以便在 WPS 97 中查看新文档而对其进行优化，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Overtype**                                    | 如果激活改写模式，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Pagination**                                  | 如果 WPS 在后台重新为文档分页，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**                                      | 返回一个 **Object** 类型值，该值代表指定 **Options** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PasteAdjustParagraphSpacing**                 | 如果在剪切和粘贴选定内容时，WPS 自动调整段落间距，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PasteAdjustTableFormatting**                  | 如果在剪切和粘贴选定内容时，WPS 自动调整表格的格式，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PasteAdjustWordSpacing**                      | 如果在剪切和粘贴选定内容时，WPS 自动调整字符间距，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PasteFormatBetweenDocuments**                 | 返回或设置一个 [WdPasteOptions ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdPasteOptions%20%E6%9E%9A%E4%B8%BE.html)常量，该常量代表在从另一个 WPS Office WPS 文档复制文本时粘贴文本的方式。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PasteFormatBetweenStyledDocuments**           | 返回或设置一个 **WdPasteOptions** 常量，该常量代表在从使用样式的文档中复制文本时粘贴文本的方式。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PasteFormatFromExternalSource**               | 返回或设置一个 [WdPasteOptions ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdPasteOptions%20%E6%9E%9A%E4%B8%BE.html)常量，该常量代表在从外部源（如网页）复制文本时粘贴文本的方式。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PasteFormatWithinDocument**                   | 返回或设置一个 [WdPasteOptions ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdPasteOptions%20%E6%9E%9A%E4%B8%BE.html)常量，该常量代表在同一文档内复制或剪切文本然后进行粘贴时的文本粘贴方式。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PasteMergeFromPPT**                           | 如果为 **True**，则从 Microsoft PowerPoint 粘贴文本时合并文本格式。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PasteMergeFromXL**                            | 如果为 **True**，则从 ET 粘贴表格时合并表格的格式。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PasteMergeLists**                             | 如果为 **True**，则合并粘贴列表和周围列表的格式。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PasteOptionKeepBulletsAndNumbers**            | 返回或设置 **Boolean** 值，该值表示在从**“粘贴选项”**上下文菜单中选择**“仅保留文本”**时是否保留项目符号和编号。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PasteSmartCutPaste**                          | 如果 WPS 在文档中智能粘贴选定内容，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PasteSmartStyleBehavior**                     | 从其他文档粘贴选定文本时，如果 WPS 智能合并样式，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PictureEditor**                               | 返回或设置用于编辑图片的应用程序的名称。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PictureWrapType**                             | 设置或返回一个 **WdWrapTypeMerged** 类型的值，该值表示 WPS 在图片周围环绕文字的方式。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PortugalReform**                              | 返回或设置欧洲葡萄牙语拼写器的模式。可读/写 WdPortugueseReform。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PrecisePositioning**                          | 返回或设置一个 **Boolean** 类型的值，该值表示 WPS 是否针对打印版式而不是针对屏幕可读性优化字符位置。如果该属性的值为 **True**，则禁用压缩字符间距以便增强屏幕可读性的默认设置，并启用适用于打印介质的字符间距。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PrintBackground**                             | 如果 WPS 在后台打印，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PrintBackgrounds**                            | 返回一个 **Boolean** 类型的值，该值代表打印文档时，是否打印背景颜色和图像。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PrintComments**                               | 如果 WPS 在文档后另起一页打印批注，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PrintDraft**                                  | 如果 WPS 以最简单的格式打印文档，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PrintDrawingObjects**                         | 如果 WPS 打印图形对象，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PrintEvenPagesInAscendingOrder**              | 如果 WPS 在手动双面打印模式下按升序打印偶数页，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PrintFieldCodes**                             | 如果 WPS 打印域代码而不打印域的结果，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PrintHiddenText**                             | 如果打印隐藏文字，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PrintProperties**                             | 如果 WPS 在文档结尾处另起一页打印文档摘要信息，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PrintReverse**                                | 如果 WPS 逆页序打印页面，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PrintXMLTag**                                 | 返回一个 **Boolean** 类型的值，该值代表打印文档时是否打印 XML 标记。对应于**“选项”**对话框中**“打印”**选项卡上的**“XML 标记”**复选框。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PromptUpdateStyle**                           | 如果为 **True**，则在改变样式的格式时显示一条消息，请用户确认是否要重新设置样式的格式或重新应用原样式格式。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RepeatWPS**                                   | 返回或设置 **Boolean** 值，该值表示在检查拼写时是否标记重复的单词。如果为 **True**，将标记重复的单词。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ReplaceSelection**                            | 如果键入或粘贴的内容替换选定内容，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RevisedLinesColor**                           | 当激活修订选项时，返回或设置文档中修订线的颜色。**WdColorIndex** 类型，可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RevisedLinesMark**                            | 当激活修订选项时，返回或设置修订线在文档中的位置。**WdRevisedLinesMark** 类型，可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RevisedPropertiesColor**                      | 在启用修订功能时，返回或设置用于标记格式更改情况的颜色。**WdColorIndex** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RevisedPropertiesMark**                       | 在启用修订功能时，返回或设置用于显示格式修改情况的标记。**WdRevisedPropertiesMark** 类型，可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RevisionsBalloonPrintOrientation**            | 返回或设置一个 **WdRevisionsBalloonPrintOrientation** 常量，该常量表示打印时修订和批注气球的方向。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RTFInClipboard**                              | 您请求了仅用于 Macintosh 上的 Visual Basic 关键字的帮助。有关 **Options** 对象的 **RTFInClipboard** 属性的信息，请参阅 WPS OfficeMacintosh Edition 中附带的语言参考帮助。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SaveInterval**                                | 返回或设置保存“自动恢复”信息的时间间隔（以分钟为单位）。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SaveNormalPrompt**                            | 如果 WPS 在关闭之前提示用户确认是否保存对 Normal 模板所做的更改，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SavePropertiesPrompt**                        | 如果 WPS 在保存新的文档时，提示该文档的属性信息，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SendMailAttach**                              | 如果使用**“文件”**菜单上的**“发送”**命令将活动文档作为附件插入邮件中，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SequenceCheck**                               | 如果为 **True**，则检查南亚文本独立字符的顺序。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShortMenuNames**                              | 您请求了仅用于 Macintosh 上的 Visual Basic 关键字的帮助。有关 **Options** 对象的 **ShortMenuNames** 属性的信息，请参阅 WPS OfficeMacintosh Edition 中附带的语言参考帮助。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowControlCharacters**                       | 如果显示当前文档中的双向控制符，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowDevTools**                                | 返回或设置 **Boolean** 类型的值，该值代表是否在功能区中显示**“开发人员”**选项卡。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowDiacritics**                              | 如果在使用从右向左语言的文档中显示音调符号，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowFormatError**                             | 如果为 **True**，WPS 通过在文本（该文本的格式与文档中使用频繁的其他格式相似）下加弯曲下划线来标记不一致的格式。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowMarkupOpenSave**                          | 返回或设置一个 **Boolean** 类型的值，该值代表打开或保存文件时 WPS 是否显示隐藏标记。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowMenuFloaties**                            | 返回或设置 **Boolean** 值，该值表示当用户在文档窗口中右键单击时是否显示浮动工具栏。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowReadabilityStatistics**                   | 如果 WPS 在结束语法检查时，显示统计信息的摘要列表（包括可读性程度），则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowSelectionFloaties**                       | 返回或设置 **Boolean** 值，该值表示当用户选择文本时是否显示浮动工具栏。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SmartCursoring**                              | 返回或设置一个 **Boolean** 类型的值，该值表示是否启用智能指针。如果值为 **True**，则启用智能指针。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SmartCutPaste**                               | 如果在剪切和粘贴时，WPS 自动调整字词和标点符号之间的间距，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SmartParaSelection**                          | 如果为 **True**，则在选择段落中大多数或全部内容时，WPS 会在选定内容中包含段落标记。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SnapToGrid**                                  | 如果绘制、移动自选图形或中文字符或者调整其大小时，它们自动与不可见的网格对齐，则该属性值为 **True**。可读写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SnapToShapes**                                | 如果 WPS 自动将自选图形或中文字符与不可见的网格线对齐，这些网格线穿过其他自选图形或中文字符的垂直和水平边缘，则该属性值为 **True**。可读写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SpanishMode**                                 | 返回或设置西班牙语拼写器的模式。可读/写 WdSpanishSpeller。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SplitCellColor**                              | 返回或设置一个 **WdCellColor** 常量，该常量代表拆分的表格单元格的颜色。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **StoreRSIDOnSave**                             | 如果为 **True**，则 WPS 在每次保存文档时为文档中的变化指定一个随机数，以便比较与合并文档。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **StrictFinalYaa**                              | 如果拼写检查使用针对以字母“yaa”结尾的阿拉伯语单词的拼写规则，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **StrictInitialAlefHamza**                      | 如果拼写检查使用针对以“alef hamza”开头的阿拉伯语单词的拼写规则，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **StrictRussianE**                              | 如果拼写检查器使用有关使用严格 ? 字符的俄语单词的拼写规则，则该属性值为 **True**。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **StrictTaaMarboota**                           | 如果拼写检查器使用拼写规则来标记以 haa（而不是 taa marboota）结尾的阿拉伯语单词，则该属性值为 **True**。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SuggestFromMainDictionaryOnly**               | 如果 WPS 仅根据主词典提供拼写建议，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SuggestSpellingCorrections**                  | 如果 WPS 在检查拼写时，总是为每一个拼写错误的单词提供可选的拼写建议，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TabIndentKey**                                | 如果可分别使用 Tab 和 Backspace 键增加和减少段落的左缩进，并且可使用 Backspace 键将右对齐的段落更改为居中的段落而将居中的段落更改为左对齐的段落，则该属性值为 **True**。可读写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TypeNReplace**                                | 如果为 **True**，则 WPS 将替换非法的南亚字符。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **UpdateFieldsAtPrint**                         | 如果 WPS 在打印文档前自动更新域，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **UpdateFieldsWithTrackedChangesAtPrint**       | 如果 WPS 2015 允许在打印前更新包含修订的字段，则该属性值为 **True**。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **UpdateLinksAtOpen**                           | 如果 WPS 在打开文档时自动更新嵌入的所有 OLE 链接，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **UpdateLinksAtPrint**                          | 如果 WPS 在打印文档前更新指向其他文件的嵌入链接，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **UpdateStyleListBehavior**                     | 返回或设置 WdUpdateStyleListBehavior 常量，该常量指定在更新某种样式以与包含编号或项目符号的所选内容相匹配时， WPS 2015 应采取的行为。可读/写。rs">版本信息 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **UseCharacterUnit**                            | 如果 WPS 在当前文档中使用字符作为默认度量单位，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **UseDiffDiacColor**                            | 如果可以设置当前文档中音调符号的颜色，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **UseGermanSpellingReform**                     | 如果 WPS 在检查拼写时，使用德语后期修订拼写规则，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **UseNormalStyleForList**                       | 返回或设置 **Boolean** 值，该值表示 WPS 是否对项目符号和编号使用“正文”样式。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **VisualSelection**                             | 根据从右向左式语言文档中的可视光标移动，返回或设置选择行为。**WdVisualSelection** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **WarnBeforeSavingPrintingSendingMarkup**       | 如果为 **True**，则 WPS 在保存、打印或以电子邮件形式发送包含批注或修订的文档时，显示警告。**Boolean** 类型，可读写。 |

**成员方法**

#### **Options.DefaultFilePath**

返回或设置各项（例如文档、模板和图形）的默认文件夹。**String** 类型，可读写。

**语法**

**express.DefaultFilePath(Path)**

*express*   一个代表 **Options** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型**          | **说明**           |
| -------- | ------------- | --------------------- | ------------------ |
| *Path*   | 必选          | **WdDefaultFilePath** | 设置默认的文件夹。 |

**说明**

您可以使用空字符串 ("") 从 Windows 注册表中删除设置。新设置会立即生效。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为 WPS 文档设置默认文件夹。*/ Options.DefaultFilePath(wdDocumentsPath) = "C:\\Documents"` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回用户模板当前的默认路径（相当于“选项”对话框中“文件位置”选项卡上的默认路径设置）。*/ let strPath = Options.DefaultFilePath(wdUserTemplatesPath)` |

**成员属性**

#### **Options.AddBiDirectionalMarksWhenSavingTextFile**

如果 WPS 在将文档保存为文本文件时添加双向控制字符，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.AddBiDirectionalMarksWhenSavingTextFile**

*express*   一个代表 **Options** 对象的变量。

**说明**

如果保存文本文件时带有双向控制符，则可保留从右向左和从左向右的属性，以及中性字符的顺序。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将设定 WPS 在将文档保存为文本文件时添加双向控制字符。*/ Options.AddBiDirectionalMarksWhenSavingTextFile = true` |

#### **Options.AddControlCharacters**

如果 WPS 在剪切和复制文本时添加双向控制字符，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.AddControlCharacters**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 在剪切和复制文本时添加双向控制字符。*/ Options.AddControlCharacters = true` |

#### **Options.AddHebDoubleQuote**

如果 WPS 会将数字格式括在双引号 (") 中，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.AddHebDoubleQuote**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 将数字格式括在双引号 (") 中。*/ Options.AddHebDoubleQuote = true` |

#### **Options.AllowAccentedUppercase**

如果当一个法语字符改成大写字母时重音仍保留，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.AllowAccentedUppercase**

*express*   一个代表 **Options** 对象的变量。

**说明**

本属性只影响标准法语文本。对于其他语言的文本，尽管将 **AllowAccentedUppercase** 属性设为 **False**，仍然保留重音。

当一个重音符号被删除之后，如果把一个字符改回小写字母，则重音不会再出现。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*当法语字符被改为大写字母时，本示例将设置 WPS 删除重音符号。*/ Options.AllowAccentedUppercase = false` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回在“选项”对话框中的“编辑”选项卡上的“允许法语中大写字母带重音符号”选项的状态。*/ let blnUppercaseAccents = Options.AllowAccentedUppercase` |

#### **Options.AllowClickAndTypeMouse**

如果启用“即点即输”功能，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.AllowClickAndTypeMouse**

*express*   一个代表 **Options** 对象的变量。

**说明**

有关“即点即输”功能的详细信息，请参阅“关于即点即输”。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例检查并确认是否已启用“即点即输”功能。如果未启用该功能，则根据用户的选择来设置该功能。*/ function test() { if(Options.AllowClickAndTypeMouse == false) {      x = MsgBox("Do you want to use Click and Type?", jsYesNo)     if(x == jsResultYes) {         Options.AllowClickAndTypeMouse = true         MsgBox("Click and Type enabled!")     } } }` |

#### **Options.AllowCombinedAuxiliaryForms**

如果 WPS 在朝鲜语文档中执行拼写检查时将忽略辅助动词形式，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.AllowCombinedAuxiliaryForms**

*express*   一个代表 **Options** 对象的变量。

**说明**

有关在 WPS 中使用亚洲语言的详细信息，请参阅 WPS 中关于亚洲语言的功能。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例询问用户：WPS 在朝鲜语文档中执行拼写检查时，是否忽略辅助动词形式。*/ function test() { if(Options.AllowCombinedAuxiliaryForms == false){     let x = MsgBox("Do you want to ignore auxiliary verb forms when checking spelling?", jsYesNo)     if(x == jsResultYes){         Options.AllowCombinedAuxiliaryForms = true         MsgBox("Auxiliary verb forms will be ignored!")     } } }  ` |

#### **Options.AllowCompoundNounProcessing**

如果 WPS 在朝鲜语文档中执行拼写检查时将忽略复合名词，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.AllowCompoundNounProcessing**

*express*   一个代表 **Options** 对象的变量。

**说明**

有关在 WPS 中使用亚洲语言的详细信息，请参阅 WPS 中关于亚洲语言的功能。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例询问用户：WPS 在朝鲜语文档中执行拼写检查时，是否忽略复合名词。*/ function test() { if(Options.AllowCompoundNounProcessing == false){     let x = MsgBox("Do you want to ignore compound "+"nouns when checking spelling?",jsYesNo)     if(x == jsResultYes){         Options.AllowCompoundNounProcessing = true         MsgBox("Compound nouns will be ignored!")     } } }` |

#### **Options.AllowDragAndDrop**

如果可以使用拖动来移动或复制所选内容，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.AllowDragAndDrop**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例打开拖放编辑功能。*/ Options.AllowDragAndDrop = true` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“选项”对话框中“编辑”选项卡上的“拖放式文字编辑”选项的状态。*/ let blnDragAndDrop = Options.AllowDragAndDrop` |

#### **Options.AllowOpenInDraftView**

返回或设置 **Boolean** 值，该值表示是否允许用户在草稿视图中打开文档。可读/写。

**语法**

**express.AllowOpenInDraftView**

*express*   一个代表 **Options** 对象的变量。

**说明**

此属性对应于**“ WPS 选项”**对话框中的**“允许在草稿视图中打开文档”**复选框。

#### **Options.AllowPixelUnits**

如果 WPS 用像素作为默认度量单位（这是 HTML 功能所支持的度量单位），则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.AllowPixelUnits**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例实现的功能是：使 WPS 可以将像素作为 HTML 功能的默认度量单位。*/ Options.AllowPixelUnits = true` |

#### **Options.AllowReadingMode**

如果为 **True**，则表明 WPS 将文档在“阅读版式”视图中打开。**Boolean** 类型，可读写。

**语法**

**express.AllowReadingMode**

*express*   一个代表 **Options** 对象的变量。

**说明**

对应于**“选项”**对话框中**“常规”**选项卡上的**“允许从‘阅读版式’启动”**复选框。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下列示例切换“允许从‘阅读版式’启动”复选框。*/ function ToggleReadingMode(){     if(Options.AllowReadingMode == true){         Options.AllowReadingMode = false     }     else{         Options.AllowReadingMode = true     } }` |

#### **Options.AlwaysUseClearType**

返回或设置 **Boolean** 值，该值表示是否使用 ClearType 来显示菜单、功能区和对话框文本中的字体。可读/写。

**语法**

**express.AlwaysUseClearType**

*express*   一个代表 **Options** 对象的变量。

**说明**

即使关闭 Microsoft Windows 的 ClearType 设置，将此属性设置为 **True** 也会替代 Windows 设置，并在所有 WPS Office应用程序中使用 ClearType。此属性对应于**“ WPS 选项”**对话框的**“高级”**选项卡中的**“总是使用 ClearType”**复选框。

#### **Options.AnimateScreenMovements**

如果 WPS 激活动态鼠标功能、使用动态光标，并且激活后台保存与查找和替换操作之类的动态操作效果，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.AnimateScreenMovements**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 以激活屏幕上的动态移动功能。*/ Options.AnimateScreenMovements = true` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“工具”菜单“选项”对话框中“常规”选项卡的“提供动画反馈”选项的当前状态。*/ let blnAnimation = Options.AnimateScreenMovements` |

#### **Options.Application**

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

*express*   一个代表 **Options** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **Options.ApplyFarEastFontsToAscii**

如果 WPS 将中文字体应用于西文文字，则该属性值为 **True**。

**语法**

**express.ApplyFarEastFontsToAscii**

*express*   一个代表 **Options** 对象的变量。

**说明**

仅在选择一种东亚语言进行编辑时才应用该属性。如果为 **False**，并且对某个特定区域应用中文字体，则 WPS 将不对该区域中的任何西文文字应用该字体。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 对西文文字应用中文字体。*/ Options.ApplyFarEastFontsToAscii = true` |

#### **Options.ArabicMode**

返回或设置阿拉伯语的拼写检查模式。**WdAraSpeller** 类型，可读写。

**语法**

**express.ArabicMode**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置拼写检查忽略阿拉伯语中关于单词以 alef hamza 开头的拼写规则。*/ Options.ArabicMode = wdInitialAlef` |

#### **Options.ArabicNumeral**

返回或设置一篇阿拉伯语文档的数字样式。**WdArabicNumeral** 类型，可读写。

**语法**

**express.ArabicNumeral**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例把数字样式设为印地语样式。*/ Options.ArabicNumeral = wdNumeralHindi` |

#### **Options.AutoCreateNewDrawings**

如果为 **True**，则 WPS 在绘图画布上画出新创建的图形。**Boolean** 类型，可读写。

**语法**

**express.AutoCreateNewDrawings**

*express*   一个代表 **Options** 对象的变量。

**说明**

**AutoCreateNewDrawings** 属性只能影响从 WPS 内部添加的图形。如果图形是通过 示例代码 代码添加的，则此选项设置为 **True** 或 **False** 都没有关系，因为它们是按照代码的指定添加的。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 直接在文档中添加新创建的图形，而不通过绘图画布。*/ function NewDrawings(){     Application.Options.AutoCreateNewDrawings = false }` |

#### **Options.AutoFormatApplyBulletedLists**

如果当 WPS 自动设置文档或区域格式时，会用**“格式”**菜单的**“项目符号和编号”**对话框中的项目符号替换列表段落的开始字符（例如星号、连字符和大于号符号），则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.AutoFormatApplyBulletedLists**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例用项目符号替换当前选定内容中列表段落的段首字符。*/ function test() { Options.AutoFormatApplyBulletedLists = true Selection.Range.AutoFormat() }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“工具”菜单的“自动更正”对话框中“自动套用格式”选项卡上“自动项目符号列表”选项的状态。*/ let blnAutoFormat = Options.AutoFormatApplyBulletedLists` |

#### **Options.AutoFormatApplyFirstIndents**

如果为 **True**，则 WPS 在自动设置文档或区域的格式时，用首行缩进替换在段落开始处输入的空格。**Boolean** 类型，可读写。

**语法**

**express.AutoFormatApplyFirstIndents**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 用首行缩进替换在段落开始处输入的空格，并自动设置选定区域的格式。*/ function test() { Options.AutoFormatApplyFirstIndents = true Selection.Range.AutoFormat() }` |

#### **Options.AutoFormatApplyHeadings**

如果当 WPS 自动设置文档或区域格式时，将自动为标题设置样式，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.AutoFormatApplyHeadings**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例对选定内容中的标题应用标题 1 到标题 9 的样式。*/ function test() { Options.AutoFormatApplyHeadings = true Selection.Range.AutoFormat() }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“工具”菜单的“自动更正”对话框中“自动套用格式”选项卡上“标题”选项的状态。*/ let blnAutoFormat = Options.AutoFormatApplyHeadings` |

#### **Options.AutoFormatApplyLists**

如果当 WPS 自动设置文档或区域格式时，自动为列表应用样式，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.AutoFormatApplyLists**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为当前选定内容中的所有列表应用样式。*/ function test() { Options.AutoFormatApplyLists = true Selection.Range.AutoFormat() }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“工具”菜单的“自动更正”对话框中“自动套用格式”选项卡上“列表”选项的状态。*/ let blnAutoFormat = Options.AutoFormatApplyLists` |

#### **Options.AutoFormatApplyOtherParas**

如果当 WPS 自动设置文档或区域格式时，自动为非标题项或列表项的段落设置样式，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.AutoFormatApplyOtherParas**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例自动为当前选定段落应用样式。*/ function test() { Options.AutoFormatApplyOtherParas = true Selection.Range.AutoFormat() }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“工具”菜单的“自动更正”对话框中“自动套用格式”选项卡上“其他段落样式”选项的状态。*/ let blnAutoFormat = Options.AutoFormatApplyOtherParas` |

#### **Options.AutoFormatAsYouTypeApplyBorders**

如果按下 Enter 后，三个或更多的连字符 (-)、等号 (=) 或下划线 (_) 将自动替换成指定的边框线，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.AutoFormatAsYouTypeApplyBorders**

*express*   一个代表 **Options** 对象的变量。

**说明**

以一条 0.75 磅的直线替换连字符 (-)，以一条 0.75 磅的双线替换等号符号 (=)，以一条 1.5 磅的直线替换下划线 (_)。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将三个或更多的连字符 (-)、等号符号 (=) 或下划线字符 (_) 序列替换成边框。*/ Options.AutoFormatAsYouTypeApplyBorders = true` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“工具”菜单的“自动更正”对话框中“键入时自动套用格式”选项卡上“边框”选项的当前设置。*/ MsgBox(Options.AutoFormatAsYouTypeApplyBorders)` |

#### **Options.AutoFormatAsYouTypeApplyBulletedLists**

如果在键入时用**“格式”**菜单的**“项目符号和编号”**对话框中的项目符号替换当前的项目符号（例如星号、连字符和大于号），则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.AutoFormatAsYouTypeApplyBulletedLists**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例用项目符号替换在列表中键入的字符。*/ Options.AutoFormatAsYouTypeApplyBulletedLists = true` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“工具”菜单的“自动更正”对话框中“键入时自动套用格式”选项卡上“自动项目符号列表”选项的状态。*/ let blnAutoFormat = Options.AutoFormatAsYouTypeApplyBulletedLists ` |

#### **Options.AutoFormatAsYouTypeApplyClosings**

如果为 **True**，则 WPS 会在键入时自动为信函的结束语应用“结束语”样式。**Boolean** 类型，可读写。

**语法**

**express.AutoFormatAsYouTypeApplyClosings**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 在键入时自动为信函的结束语应用“结束语”样式。*/ function AutoClosings(){     Options.AutoFormatAsYouTypeApplyClosings = true }` |

#### **Options.AutoFormatAsYouTypeApplyDates**

如果为 **True**，则 WPS 自动为键入的日期应用“日期”样式。可读写。

**语法**

**express.AutoFormatAsYouTypeApplyDates**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 自动为键入时输入的日期应用“日期”样式。*/ function AutoApplyDates(){     Options.AutoFormatAsYouTypeApplyDates = true }` |

#### **Options.AutoFormatAsYouTypeApplyFirstIndents**

如果为 **True**，则 WPS 在键入时自动用首行缩进替代段落开始处输入的空格。可读写。

**语法**

**express.AutoFormatAsYouTypeApplyFirstIndents**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 自动将键入时在段落开始处输入的空格替换为首行缩进。*/ function ApplyFirstIndents(){     Options.AutoFormatAsYouTypeApplyFirstIndents = true }` |

#### **Options.AutoFormatAsYouTypeApplyHeadings**

如果在键入标题时，自动为标题设置样式，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.AutoFormatAsYouTypeApplyHeadings**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*当键入文字时，本示例对选定内容中的标题应用标题 1 到标题 9 的样式。*/ Options.AutoFormatAsYouTypeApplyHeadings = true` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“工具”菜单的“自动更正”对话框中“键入时自动套用格式”选项卡上“标题”选项的状态。*/ let blnAutoFormat = Options.AutoFormatAsYouTypeApplyHeadings` |

#### **Options.AutoFormatAsYouTypeApplyNumberedLists**

如果根据键入的内容，以**“格式”**菜单的“项目符号和编号”对话框中的编号方案将段落自动设置为编号列表，则该属性值为 **True**。例如，如果一个段落以“1.1”和一个制表符开始，则当按下 Enter 后 WPS 将自动插入“1.2”和一个制表符。**Boolean** 类型，可读写。

**语法**

**express.AutoFormatAsYouTypeApplyNumberedLists**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*键入文字时，本示例自动对列表进行编号。*/ Options.AutoFormatAsYouTypeApplyNumberedLists = true` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“工具”菜单的“自动更正”对话框中“键入时自动套用格式”选项卡上“自动编号列表”选项的状态。*/ let blnAutoFormat = Options.AutoFormatAsYouTypeApplyNumberedLists` |

#### **Options.AutoFormatAsYouTypeApplyTables**

如果为 **True**，则当您键入一个加号、一串连字符、另一个加号等，然后按下 Enter 后， WPS 将自动创建一张表格。加号变为列边框，连字符变为列宽度。**Boolean** 类型，可读写。

**语法**

**express.AutoFormatAsYouTypeApplyTables**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 在您键入时自动创建表格。*/ Options.AutoFormatAsYouTypeApplyTables = true` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“工具”菜单的“自动更正”对话框中“键入时自动套用格式”选项卡上“表格”选项的状态。*/ let blnAutoFormat = Options.AutoFormatAsYouTypeApplyTables` |

#### **Options.AutoFormatAsYouTypeAutoLetterWizard**

如果为 **True**，则在用户输入信函称呼或结束语时，WPS 会自动启动“英文信函向导”。可读写。

**语法**

**express.AutoFormatAsYouTypeAutoLetterWizard**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 WPS 设置为在用户输入信函称呼或结束语时，自动启动“英文信函向导”。*/ function AutoLeterWizard(){     Options.AutoFormatAsYouTypeAutoLetterWizard = true }` |

#### **Options.AutoFormatAsYouTypeDefineStyles**

如果 WPS 自动根据手动设置的格式新建一种样式，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.AutoFormatAsYouTypeDefineStyles**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*键入文本时，本示例使 WPS 自动创建样式。*/ Options.AutoFormatAsYouTypeDefineStyles = true` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“工具”菜单的“自动更正”对话框中“键入时自动套用格式”选项卡上“基于所用格式定义样式”选项的状态。*/ let blnAutoFormat = Options.AutoFormatAsYouTypeDefineStyles` |

#### **Options.AutoFormatAsYouTypeDeleteAutoSpaces**

如果为 **True**，则 WPS 会在用户键入时自动删除日文和拉丁文文字之间插入的空格。可读写。

**语法**

**express.AutoFormatAsYouTypeDeleteAutoSpaces**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 在用户键入时自动删除日文和拉丁文文字之间插入的空格。*/ function AutoDeleteSpaces() {     Options.AutoFormatAsYouTypeDeleteAutoSpaces = true }` |

#### **Options.AutoFormatAsYouTypeFormatListItemBeginning**

如果 WPS 将列表项开始处的字符格式重复应用于下一个列表项，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.AutoFormatAsYouTypeFormatListItemBeginning**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 自动重复列表项开始处的字符格式。*/ Options.AutoFormatAsYouTypeFormatListItemBeginning = true` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“工具”菜单的“自动更正”对话框中“键入时自动套用格式”选项卡上“将列表项开始的格式设为与其前一项相似”选项的状态。*/ let blnAutoFormat = Options.AutoFormatAsYouTypeFormatListItemBeginning` |

#### **Options.AutoFormatAsYouTypeInsertClosings**

如果为 **True**，则当用户输入备忘录标题时，WPS 会自动插入相应的备忘录结束语。可读写。

**语法**

**express.AutoFormatAsYouTypeInsertClosings**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 WPS 设置为在用户输入备忘录标题时自动插入相应的备忘录结束语。*/ function AutoInsertClosings() {     Options.AutoFormatAsYouTypeInsertClosings = true }` |

#### **Options.AutoFormatAsYouTypeInsertOvers**

如果为 **True**，则当用户输入“ ![img]()”或“ ![img]()”时，WPS 将自动插入“ ![img]() ![img]()”。**Boolean** 类型，可读写。

**语法**

**express.AutoFormatAsYouTypeInsertOvers**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 WPS 设置为在用户输入“ ki”或“ an”时自动插入“ ijou ijou”。*/ Options.AutoFormatAsYouTypeInsertOvers = true` |

#### **Options.AutoFormatAsYouTypeMatchParentheses**

如果为 **True**，WPS 会自动更正不正确成对的括号，可读写。

**语法**

**express.AutoFormatAsYouTypeMatchParentheses**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 在用户键入时自动更正不正确成对的括号。*/ function AutoMatchParentheses() {     Options.AutoFormatAsYouTypeMatchParentheses = true }` |

#### **Options.AutoFormatAsYouTypeReplaceFarEastDashes**

如果为 **True**，WPS 会自动更正长元音和划线，可读写。

**语法**

**express.AutoFormatAsYouTypeReplaceFarEastDashes**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 在用户键入时自动更正长元音和划线。*/ function AutoFarEastDashes() {     Options.AutoFormatAsYouTypeReplaceFarEastDashes = true }` |

#### **Options.AutoFormatAsYouTypeReplaceFractions**

如果在您键入时，将使用当前字符集的分数替换键入的分数，则该属性值为 **True**。例如，将“1/2”替换为“?”。**Boolean** 类型，可读写。

**语法**

**express.AutoFormatAsYouTypeReplaceFractions**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例关闭键入分数的自动替换功能。*/ Options.AutoFormatAsYouTypeReplaceFractions = false` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“工具”菜单的“自动更正”对话框中“键入时自动套用格式”选项卡上“分数 (1/2) 替换为分数字符 (?)”选项的状态。*/ let blnAutoFormat = Options.AutoFormatAsYouTypeReplaceFractions` |

#### **Options.AutoFormatAsYouTypeReplaceHyperlinks**

如果本属性值为 **True**，则 WPS 会在键入时自动将电子邮件地址、服务器和共享名称（即 UNC 路径）和 Internet 地址（即 URL）改为超链接。**Boolean** 类型，可读写。

**语法**

**express.AutoFormatAsYouTypeReplaceHyperlinks**

*express*   一个代表 **Options** 对象的变量。

**说明**

WPS 将任何看起来像电子邮件地址、UNC 或 URL 的文字更改为超链接。 WPS 并不检查超链接是否有效。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例使 WPS 自动将键入的任何 Internet 或者网络路径替换为超链接。*/ Options.AutoFormatAsYouTypeReplaceHyperlinks = true` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“工具”菜单的“自动更正”对话框中“键入时自动套用格式”选项卡上“Internet 及网络路径替换为超链接”选项的状态。*/ let blnAutoFormat = Options.AutoFormatAsYouTypeReplaceHyperlinks` |

#### **Options.AutoFormatAsYouTypeReplaceOrdinals**

如果为 **True**，则键入时， WPS 会用相同字母的上标格式替换序数后缀“st”、“nd”、“rd”和“th”。例如，“1st”被替换为“1”后接上标“st”。**Boolean** 类型，可读写。

**语法**

**express.AutoFormatAsYouTypeReplaceOrdinals**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例启用将序数字母自动替换为上标字母的功能。*/ Options.AutoFormatAsYouTypeReplaceOrdinals = true` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“工具”菜单的“自动更正”对话框中“键入时自动套用格式”选项卡上“序号 (1st) 替换为上标”选项的状态。*/ let blnAutoFormat = Options.AutoFormatAsYouTypeReplaceOrdinals` |

#### **Options.AutoFormatAsYouTypeReplacePlainTextEmphasis**

如果键入时，手动强调的字符将自动替换为相应字符格式，则该属性值为 **True**。例如，将“*bold*”改为“**bold**”、将“_underline_”改为“underline”。**Boolean** 类型，可读写。

**语法**

**express.AutoFormatAsYouTypeReplacePlainTextEmphasis**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例启用将手动强调的字符替换为相应字符格式的功能。*/ Options.AutoFormatAsYouTypeReplacePlainTextEmphasis = true ` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“工具”菜单的“自动更正”对话框中“键入时自动套用格式”选项卡上“*加粗* 和 _下划线_ 替换为真正格式”选项的状态。*/ let blnAutoFormat = Options.AutoFormatAsYouTypeReplacePlainTextEmphasis  ` |

#### **Options.AutoFormatAsYouTypeReplaceQuotes**

如果键入时自动将所键入的直引号替换为弯引号，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.AutoFormatAsYouTypeReplaceQuotes**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例启用将所键入的直引号自动替换为弯引号的功能。*/ Options.AutoFormatAsYouTypeReplaceQuotes = true` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“工具”菜单的“自动更正”对话框中“键入时自动套用格式”选项卡上“直引号替换为弯引号”选项的状态。*/ let blnAutoFormat = Options.AutoFormatReplaceQuotes` |

#### **Options.AutoFormatAsYouTypeReplaceSymbols**

如果键入的两个连续的连字符 (--) 将替换为短划线 (–) 或长划线 (—)，则该属性值为 **True**。**Boolean** 类型，可读写。如果键入连字符时带有前导和尾部空格，则 WPS 将连字符替换为短划线。如果没有尾部空格，则连字符替换为长划线。

**语法**

**express.AutoFormatAsYouTypeReplaceSymbols**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例启用将键入的将连字符替换为符号的功能。*/ Options.AutoFormatAsYouTypeReplaceSymbols = true` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“工具”菜单的“自动更正”对话框中“键入时自动套用格式”选项卡上“连字符 (--) 替换为长划线 (—)”选项的状态。*/ let blnAutoFormat = Options.AutoFormatAsYouTypeReplaceSymbols` |

#### **Options.AutoFormatDeleteAutoSpaces**

如果 WPS 在自动设置文档或区域的格式时，删除日文和西文文字之间插入的空格，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.AutoFormatDeleteAutoSpaces**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 自动删除日文和西文文字之间的空格，然后设置当前选定内容的格式。*/ function test() { Options.AutoFormatDeleteAutoSpaces = true Selection.Range.AutoFormat() }` |

#### **Options.AutoFormatMatchParentheses**

如果 WPS 在自动设置文档或区域的格式时，会更正不恰当匹配的括号，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.AutoFormatMatchParentheses**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 自动更正不恰当匹配的括号，然后设置当前选定内容的格式。*/ function test() { Options.AutoFormatMatchParentheses = true Selection.Range.AutoFormat() }` |

#### **Options.AutoFormatPlainTextWordMail**

如果在 WPS 中打开纯文本格式的电子邮件时， WPS 将自动进行格式设置，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.AutoFormatPlainTextWordMail**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 WPS 设置为可对打开的纯文本格式的电子邮件自动进行格式设置。*/ Options.AutoFormatPlainTextWordMail = true` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“工具”菜单的“自动更正”对话框中“自动套用格式”选项卡上“纯文本型 WordMail 文档”选项的状态。*/ let blnAutoFormat = Options.AutoFormatPlainTextWordMail` |

#### **Options.AutoFormatPreserveStyles**

如果当 WPS 自动设置文档或区域格式时，保留先前所用的样式，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.AutoFormatPreserveStyles**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例中，先设置 Word，使其可保留原有样式，并自动设置标题、列表和其他段落的样式，然后 WPS 自动设置当前所选内容的格式。*/ function test() { Options.AutoFormatPreserveStyles = true Options.AutoFormatApplyHeadings = true Options.AutoFormatApplyLists = true Options.AutoFormatApplyOtherParas = true  Selection.Range.AutoFormat() }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“工具”菜单的“自动更正”对话框中“自动套用格式”选项卡上“样式”选项的状态。*/ let blnAutoFormat = Options.AutoFormatPreserveStyles` |

#### **Options.AutoFormatReplaceFarEastDashes**

如果 WPS 在自动设置文档或区域的格式时，会更正长元音和划线的使用，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.AutoFormatReplaceFarEastDashes**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 自动更正长元音和划线的使用，然后设置当前选定内容的格式。*/ function test() { Options.AutoFormatReplaceFarEastDashes = true Selection.Range.AutoFormat() } ` |

#### **Options.AutoFormatReplaceFractions**

如果当 WPS 自动设置文档或区域格式时，会用当前字符集中的分数替换所键入的分数，则该属性值为 **True**。例如，将“1/2”替换为“?”。**Boolean** 类型，可读写。

**语法**

**express.AutoFormatReplaceFractions**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例激活键入分数的替换功能，然后自动设置当前选定内容的格式。*/ function test() { Options.AutoFormatReplaceFractions = true Selection.Range.AutoFormat() }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“工具”菜单的“自动更正”对话框中“自动套用格式”选项卡上“分数 (1/2) 替换为分数字符 (?)”选项的状态。*/ let blnAutoFormat = Options.AutoFormatReplaceFractions` |

#### **Options.AutoFormatReplaceHyperlinks**

如果为 **True**，则 WPS 在对某一文档或范围自动套用格式时，将对其中的电子邮件地址、服务器和共享名称（即 UNC 路径）和 Internet 地址（即 URL）自动进行格式设置。**Boolean** 类型，可读写。

**语法**

**express.AutoFormatReplaceHyperlinks**

*express*   一个代表 **Options** 对象的变量。

**说明**

WPS 将任何看起来像电子邮件地址、UNC 或 URL 的文字更改为超链接。 WPS 并不检查超链接是否有效。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例使 WPS 能自动将任何 Internet 或网络路径替换为超链接，然后对选定内容自动套用格式。*/ function test() { Options.AutoFormatReplaceHyperlinks = true Selection.Range.AutoFormat() }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“工具”菜单的“自动更正”对话框中“自动套用格式”选项卡上“Internet 及网络路径替换为超链接”选项的状态。*/ let blnAutoFormat = Options.AutoFormatReplaceHyperlinks` |

#### **Options.AutoFormatReplaceOrdinals**

如果为 **True**，则当 WPS 自动设置文档或区域格式时，用相同字母的上标格式替换序数后缀“st”、“nd”、“rd”和“th”。例如，将“1st”替换为“1”后接上标“st”。**Boolean** 类型，可读写。

**语法**

**express.AutoFormatReplaceOrdinals**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将序数字母自动替换为上标，然后自动设置当前选定内容的格式。*/ function test() { Options.AutoFormatReplaceOrdinals = true Selection.Range.AutoFormat() }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“工具”菜单的“自动更正”对话框中“自动套用格式”选项卡上“序号(1st)替换为上标”选项的状态。*/ let blnAutoFormat = Options.AutoFormatReplaceOrdinals` |

#### **Options.AutoFormatReplacePlainTextEmphasis**

如果当 WPS 自动设置文档或区域的格式时，将手动强调的字符替换为相应的字符格式（例如，“*bold*”更改为“bold”，“_underline_”更改为“underline”），则该属性值为 **True**。**Boolean** 类型，可读写。

如果当 WPS 自动设置文档或区域的格式时，将手动强调的字符替换为相应的字符格式（例如，“*bold*”更改为“bold”，“_underline_”更改为“underline”），则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.AutoFormatReplacePlainTextEmphasis**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例启用将手动强调字符替换为相应字符格式的功能。*/ function test() { Options.AutoFormatReplacePlainTextEmphasis = true Selection.Range.AutoFormat() }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“工具”菜单的“自动更正”对话框中“自动套用格式”选项卡上“*加粗*和 _下划线_ 替换为真正格式”选项的状态。*/ let blnAutoFormat = Options.AutoFormatReplacePlainTextEmphasis` |

#### **Options.AutoFormatReplaceQuotes**

如果当 WPS 自动设置文档或区域格式时，直引号会自动更改为弯引号，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.AutoFormatReplaceQuotes**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例启用将直引号自动替换为弯引号的功能，然后自动设置当前所选内容的格式。*/ function test() { Options.AutoFormatReplaceQuotes = true Selection.Range.AutoFormat() }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“工具”菜单的“自动更正”对话框中“自动套用格式”选项卡上“直引号替换为弯引号”选项的状态。*/ let blnAutoFormat = Options.AutoFormatReplaceQuotes` |

#### **Options.AutoFormatReplaceSymbols**

如果当 WPS 自动设置文档或区域的格式时，将两个连续的连字号 (--) 替换为短破折号 (

**语法**

**express.AutoFormatReplaceSymbols**

*express*   一个代表 **Options** 对象的变量。

#### **Options.AutoKeyboardSwitching**

如果 WPS 随时自动切换键盘语言并与键入的内容匹配，则该属性值为 **True**。可读/写 **Boolean** 类型。

**语法**

**express.AutoKeyboardSwitching**

*express*   一个代表 **Options** 对象的变量。

**说明**

若要使用该属性，必须将 **CheckLanguage** 属性设为 **True**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例请用户选择是否启用多语言文档的键盘自动切换功能。*/ function test() { let x = MsgBox("Enable automatic keyboard switching?", jsYesNo) if (x == jsResultYes) {     Application.CheckLanguage = true     Options.AutoKeyboardSwitching = true     MsgBox("Automatic keyboard switching enabled!") } }` |

#### **Options.AutoWordSelection**

如果在拖动时，一次选定一个单词，而不是选定一个字符，则该属性为 **True**。**Boolean** 类型，可读写。

**语法**

**express.AutoWordSelection**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 在拖动时不选定整个单词，而是选定单个字符。*/ Options.AutoWordSelection = false` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“选项”对话框中“编辑”选项卡上“选定时自动选定整个单词”选项的状态。*/ let blnAutoSelect = Options.AutoWordSelection` |

#### **Options.BackgroundOpen**

如果为 **True**，WPS 在后台打开 Web 文档。**Boolean** 类型，可读写。

**语法**

**express.BackgroundOpen**

*express*   一个代表 **Options** 对象的变量。

**说明**

当 WPS 在后台打开较大的 Web 文档时，用户可以在另一个文档中继续键入和选择命令。然而，Web 文档完全打开后，则不能在打开的文档中使用 WPS 的 示例代码 函数。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在不在后台打开较大的 Web 文档和在后台打开较大文档之间切换。*/ function BackOpen() {     if ( Options.BackgroundOpen == false ) {         Options.BackgroundOpen = true     }     else {         Options.BackgroundOpen  = false     } }` |

#### **Options.BackgroundSave**

如果为 **True**，则 WPS 在后台保存文档。当 WPS 在后台保存文档时，用户可以继续键入内容或执行命令。**Boolean** 类型，可读写。

**语法**

**express.BackgroundSave**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例让用户在 WPS 保存文档时可以继续处理文档。*/ Options.BackgroundSave = true` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“选项”对话框中“保存”选项卡上“允许后台保存”选项的当前状态。*/ let blnAutoSave = Options.BackgroundSave` |

#### **Options.BibliographySort**

返回或设置一个 **String** 类型的值，该值代表在**“源管理器”**对话框中显示源的次序。可读写。

**语法**

**express.BibliographySort**

*express*   一个代表 **Options** 对象的变量。

**说明**

可能的 **BibliographySort** 属性值有 `Author`、`Title`、`Tag` 或 `Year`。

#### **Options.BibliographyStyle**

返回或设置一个 **String** 类型的值，该值代表用于设置书目格式的样式的名称。可读写。

**语法**

**express.BibliographyStyle**

*express*   一个代表 **Options** 对象的变量。

#### **Options.BrazilReform**

返回或设置巴西葡萄牙语拼写器的模式。可读/写。

**语法**

**express.BrazilReform**

*express*   一个代表 **Options** 对象的变量。

**说明**

设置该属性与在**“ WPS 选项”**对话框中**“巴西葡萄牙语模式:”**旁边的下拉框中选择项目具有相同作用（**“校对”**、**“在 WPS Office程序中更正拼写时”**）。

| 注释                                                         |
| ------------------------------------------------------------ |
| 该属性不设置欧洲葡萄牙语拼写器的模式。若要设置欧洲葡萄牙语拼写器模式，请使用 [Options.PortugalReform](https://qn.cache.wpscdn.cn/encs/doc/office_v19/PortugalReform%20%E5%B1%9E%E6%80%A7.htm) 属性。 |

#### **Options.ButtonFieldClicks**

返回或设置一个数字，该数字是运行 GOTOBUTTON 或 MACROBUTTON 域所需的点击次数（单击或双击）。**Long** 类型，可读写。

**语法**

**express.ButtonFieldClicks**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将运行 GOTOBUTTON 或 MACROBUTTON 域所需的点击次数设置为 1。*/ Options.ButtonFieldClicks = 1` |

#### **Options.CheckGrammarAsYouType**

如果 WPS 在键入时自动对文档进行语法检查并自动标记错误，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.CheckGrammarAsYouType**

*express*   一个代表 **Options** 对象的变量。

**说明**

本属性标记语法错误，但必须在将 **ShowGrammaticalErrors** 的属性设置为 **True** 后才能在屏幕上显示错误标记。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例对 WPS 进行设置，使其在您键入时检查活动文档中的语法错误并显示发现的所有错误。*/ function test() { Options.CheckGrammarAsYouType = true ActiveDocument.ShowGrammaticalErrors = true }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“工具”菜单的“选项”对话框中“拼写和语法”选项卡上“键入时检查语法”选项的状态。*/ let blnCheck = Options.CheckGrammarAsYouType` |

#### **Options.CheckGrammarWithSpelling**

如果 WPS 在检查拼写的同时也检查语法，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.CheckGrammarWithSpelling**

*express*   一个代表 **Options** 对象的变量。

**说明**

当使用**“工具”**菜单中的**“拼写”**命令检查拼写时，可用本属性控制 WPS 是否同时检查语法。

若要通过 Visual Basic 过程检查拼写或语法，可用 **CheckSpelling** 方法只检查拼写，而用 **CheckGrammar** 方法同时检查拼写和语法。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“选项”对话框“拼写和语法”选项卡上“随拼写检查语法”选项的状态。如果选中该选项，则检查过程会检查活动文档的拼写和语法；否则只检查拼写。*/ function test() { if (Options.CheckGrammarWithSpelling == true) {     ActiveDocument.CheckGrammar() } else {     ActiveDocument.CheckSpelling() } }` |

#### **Options.CheckHangulEndings**

如果 WPS 自动检测朝鲜文字结尾，并在将朝鲜文字转换为朝鲜文汉字过程中忽略朝鲜文字结尾，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.CheckHangulEndings**

*express*   一个代表 **Options** 对象的变量。

**说明**

如果将朝鲜文汉字转换为朝鲜文字，则忽略该属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例询问用户是否设置 WPS 自动检测朝鲜文字结尾，并在将朝鲜文字转换为朝鲜文汉字过程中忽略朝鲜文字结尾。*/ function test() { let x = MsgBox("Check Hangul endings during conversion from Hangul to Hanja?", jsYesNo) if ( x == jsResultYes ) {     Options.CheckHangulEndings = true } else {     Options.CheckHangulEndings = false } }` |

#### **Options.CheckSpellingAsYouType**

如果 WPS 在您键入的同时自动进行拼写检查并标记错误，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.CheckSpellingAsYouType**

*express*   一个代表 **Options** 对象的变量。

**说明**

本属性将标记拼写错误。如要在屏幕上查看标记，则必须将 **ShowSpellingErrors** 属性设置为 **True**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例关闭 WPS 中的自动拼写检查功能。*/ Options.CheckSpellingAsYouType = false` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 在键入的同时检查拼写错误，并在活动文档中显示所找到的错误。*/ function test() { Options.CheckSpellingAsYouType = true ActiveDocument.ShowSpellingErrors = true }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“工具”菜单的“选项”对话框中“拼写和语法”选项卡上“键入时检查拼写”选项的状态。*/ let blnCheck = Options.CheckSpellingAsYouType` |

#### **Options.CommentsColor**

该属性返回或设置一个 **WdColorIndex** 常量，代表文档中批注的颜色。可读写。

**语法**

**express.CommentsColor**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 的全局选项，根据批注的作者设置批注的颜色。*/ function ColorCodeComments() {     Options.CommentsColor = wdByAuthor }` |

#### **Options.ConfirmConversions**

如果 WPS 在打开或者插入一个非 WPS 文档或模板的文件之前，会显示一个**“转换文件”**对话框，则该属性值为 **True**。在**“转换文件”**对话框中，用户可以选择需要转换的文件格式。**Boolean** 类型，可读写。

**语法**

**express.ConfirmConversions**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 WPS 设置为在打开一个非 WPS 文档或模板的文件之前，显示“转换文件”对话框。*/ Options.ConfirmConversions = true` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“选项”对话框中“常规”选项卡上“打开时确认转换”选项的当前状态。*/ let blnConfirm = Options.ConfirmConversions` |

#### **Options.ContextualSpeller**

返回或设置 **Boolean** 值，该值表示是否使用上下文拼写检查器，基于单词上下文和它前后的单词来检查拼写。可读/写。

**语法**

**express.ContextualSpeller**

*express*   一个代表 **Options** 对象的变量。

**说明**

上下文拼写检查器可以识别出单词的结构和它们在句子中的位置，以确定拼写是否正确。它可以找到标准拼写检查器找不到的错误。例如，用户可能键入“superb owl”，而不是“super bowl”。由于“superb”和“owl”都是拼写正确的单词，因此标准拼写检查器找不到错误。基于句子上下文中的单词以及它们前后的单词，上下文拼写检查器就能确定正确单词是“super”和“bowl”，并自动进行更改。

此属性对应于**“ WPS 选项”**对话框的**“校对”**选项卡中的**“使用上下文拼写检查”**复选框。

#### **Options.ConvertHighAnsiToFarEast**

如果在打开文档时，WPS 将与东亚字体相关的文字转换为相应的字体，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.ConvertHighAnsiToFarEast**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 在打开文档时，将与东亚字体相关的文字转换为相应的字体。*/ Options.ConvertHighAnsiToFarEast = true` |

#### **Options.CreateBackup**

如果 WPS 在每次保存文档时都建立一个备份，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.CreateBackup**

*express*   一个代表 **Options** 对象的变量。

**说明**

**CreateBackup** 和 **AllowFastSave** 属性不能同时设置为 **True**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 自动创建一个备份，然后开始保存活动文档。*/ function test() { Options.CreateBackup = true ActiveDocument.Save() }` |

#### **Options.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Options** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **Options.CtrlClickHyperlinkToOpen**

如果 WPS 需要单击同时按住 Ctrl 键来打开超链接，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.CtrlClickHyperlinkToOpen**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例禁用需要在单击的同时按住 Ctrl 键来打开超链接的选项。*/ function ToggleHyperlinkOption() {     Options.CtrlClickHyperlinkToOpen = false }` |

#### **Options.CursorMovement**

返回或设置在双向文字中插入点移动的方式。**WdCursorMovement** 类型，可读写。

**语法**

**express.CursorMovement**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*当插入点在双向文字中移动时，本示例使其前进至下一个可见的邻近字符。*/ Options.CursorMovement = wdCursorMovementVisual` |

#### **Options.DefaultBorderColor**

返回或设置用于新的 **Border** 对象的默认 24 位颜色。可读写。

**语法**

**express.DefaultBorderColor**

*express*   一个代表 **Options** 对象的变量。

**说明**

可以是任何有效的 **WdColor** 常量或 Visual Basic 的 **RGB** 函数返回的值。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将新边框的默认颜色设为蓝绿色。*/ Options.DefaultBorderColor = wdColorTeal` |

#### **Options.DefaultBorderColorIndex**

返回或设置边框的默认线条颜色。**WdColorIndex** 类型，可读写。

**语法**

**express.DefaultBorderColorIndex**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例更改边框的默认颜色和样式，并为活动文档的第一段设置一个边框。*/ function test() { ActiveDocument.Paragraphs.Item(1).Borders.Enable = true Options.DefaultBorderColorIndex = wdRed Options.DefaultBorderLineStyle = wdLineStyleDouble  }` |

#### **Options.DefaultBorderLineStyle**

返回或设置默认边框的样式。**WdLineStyle** 类型，可读写。

**语法**

**express.DefaultBorderLineStyle**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将默认样式设置为双线。*/ Options.DefaultBorderLineStyle = wdLineStyleDouble` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回当前默认的样式。*/ let lngTemp = Options.DefaultBorderLineStyle` |

#### **Options.DefaultBorderLineWidth**

返回或设置边框的默认线条宽度。**WdLineWidth** 类型，可读写。

**语法**

**express.DefaultBorderLineWidth**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例更改边框的默认线条宽度，然后为选定内容的每一段周围添加边框。*/ function test() { Options.DefaultBorderLineWidth = wdLineWidth050pt Selection.Borders.Enable = true }` |

#### **Options.DefaultEPostageApp**

设置或返回一个 **String** 类型的值，该值代表默认的电子邮政应用程序的路径和文件名。可读写。

**语法**

**express.DefaultEPostageApp**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例指定默认的电子邮政应用程序的路径和文件名。*/ function DefaultEPostage() {     Application.Options.DefaultEPostageApp = "C:\\MyApp\\EPostage.exe" } ` |

#### **Options.DefaultFilePath**

返回或设置各项（例如文档、模板和图形）的默认文件夹。**String** 类型，可读写。

**语法**

**express.DefaultFilePath**

*express*   一个代表 **Options** 对象的变量。

**说明**

| **名称** | **必选/可选** | **数据类型**          | **说明**           |
| -------- | ------------- | --------------------- | ------------------ |
| *Path*   | 必选          | **WdDefaultFilePath** | 设置默认的文件夹。 |

您可以使用空字符串 ("") 从 Windows 注册表中删除设置。新设置会立即生效。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为 WPS 文档设置默认文件夹。*/ Application.Options.DefaultFilePath(wdDocumentsPath,"C:\\Documents")   /*本示例返回用户模板当前的默认路径（相当于“选项”对话框中“文件位置”选项卡上的默认路径设置）。*/ let strPath = Application.Options.DefaultFilePath(wdUserTemplatesPath)` |

#### **Options.DefaultHighlightColorIndex**

返回或设置突出显示文本所用的颜色，用**“格式”**工具栏的**“突出显示”**按钮可将文本格式设为突出显示。**WdColorIndex** 类型，可读写。

**语法**

**express.DefaultHighlightColorIndex**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将默认的突出显示颜色设置为浅绿色。新的颜色不应用于以前突出显示的任何文字。*/ Options.DefaultHighlightColorIndex = wdBrightGreen` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回当前默认的突出显示颜色索引。*/ let lngTemp = Options.DefaultHighlightColorIndex` |

#### **Options.DefaultOpenFormat**

返回或设置打开文档的默认转换器。可为由 **OpenFormat** 属性返回的一个数字，也可为下列 **WdOpenFormat** 常量之一。

**语法**

**express.DefaultOpenFormat**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例首先将 WPS 文档格式设置为打开文档的默认转换器，然后打开文档 Forecast.doc。*/ function test() { Options.DefaultOpenFormat = wdOpenFormatDocument Documents.Open("C:\\Sales\\Forecast.doc") }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将打开文档所需的默认转换器设为在打开文档时自动选用相应的文件转换器。*/ Options.DefaultOpenFormat = wdOpenFormatAuto` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 WordPerfect 6.x 格式设为打开文档的默认转换器。*/ function test() { Options.DefaultOpenFormat =     FileConverters.Item("WordPerfect6x").OpenFormat }` |

#### **Options.DefaultTextEncoding**

返回或设置一个 **MsoEncoding** 常量，该常量代表 WPS 用于所有保存为编码文本文件的文档的代码页或字符集。可读写。

**语法**

**express.DefaultTextEncoding**

*express*   一个代表 **Options** 对象的变量。

**说明**

使用 **TextEncoding** 属性设置个别文档的编码。若要设置 HTML 文档的编码，请使用 **Encoding** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将全局文本编码设为 Western 代码页，这表示 WPS 将使用 Western 代码页来保存所有编码文本文件。*/ function test() { function DefaultEncode() {     Application.Options.DefaultTextEncoding = msoEncodingWestern } }` |

#### **Options.DefaultTray**

返回或设置默认纸盒，以便打印机使用该纸盒打印文档。**String** 类型，可读写。

**语法**

**express.DefaultTray**

*express*   一个代表 **Options** 对象的变量。

**说明**

要设置该属性，必须在**“选项”**对话框**“打印”**选项卡上**“默认纸盒”**框中指定一个字符串。使用 **DefaultTrayID** 属性并指定一个 **WdPaperTray** 常量也可设置该选项。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 使用下层纸盒。*/ Options.DefaultTray = "Lower tray"` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回在“选项”对话框中“打印”选项卡上“默认纸盒”框中找到的字符串。*/ Msgbox(Options.DefaultTray)` |

#### **Options.DefaultTrayID**

返回或设置用于打印文档的默认纸盒。**WdPaperTray** 类型，可读写。

**语法**

**express.DefaultTrayID**

*express*   一个代表 **Options** 对象的变量。

**说明**

将 **DefaultTray** 属性与**“选项”**对话框**“打印”**选项卡上**“默认纸盒”**框中的字符串结合使用，也可设置该选项。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/**/ function test() { Options.DefaultTrayID = wdPrinterUpperBin ActiveDocument.PrintOut() }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回 “选项”对话框中“打印”选项卡上“默认纸盒”选项的当前设置。*/ let lngTray = Options.DefaultTrayID` |

#### **Options.DeletedCellColor**

返回或设置一个 **WdCellColor** 常量，该常量代表删除的单元格的颜色。可读写。

**语法**

**express.DeletedCellColor**

*express*   一个代表 **Options** 对象的变量。

#### **Options.DeletedTextColor**

当激活修订选项时，返回或设置要删除的文字的颜色。**WdColorIndex** 类型，可读写。

**语法**

**express.DeletedTextColor**

*express*   一个代表 **Options** 对象的变量。

**说明**

如果将 **DeletedTextColor** 属性设置为 **wdByAuthor**， WPS 会自动为文档的前八名审阅者分配不同的颜色。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将要删除文本的颜色设置为鲜绿色。*/ Options.DeletedTextColor = wdBrightGreen` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“选项”对话框中“修订”选项卡上“删除的文字”下“颜色”选项的当前状态。*/ let lngTemp = Options.DeletedTextColor` |

#### **Options.DeletedTextMark**

当激活修订选项时，返回或设置要删除的文字的格式。**WdDeletedTextMark** 类型，可读写。

**语法**

**express.DeletedTextMark**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将要删除的文本的格式设置为删除线。*/ Options.DeletedTextMark = wdDeletedTextMarkStrikeThrough` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“选项”对话框中“修订”选项卡上“删除的文字”下“标记”选项的当前状态。*/ let lngTemp = Options.DeletedTextMark` |

#### **Options.DiacriticColorVal**

返回或设置用于从右向左语言文档中的音调符号的 24 位颜色。

**语法**

**express.DiacriticColorVal**

*express*   一个代表 **Options** 对象的变量。

**说明**

此属性可以是任何有效的 **WdColor** 常量或 Visual Basic 的 **RGB** 函数返回的值。必须将 **UseDiffDiacColor** 属性的值设为 **True** 才能使用此属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将音调字符的颜色设为鲜绿色。*/ function test() { if ( Options.UseDiffDiacColor == true ) {     Options.DiacriticColorVal = wdColorBrightGreen } }` |

#### **Options.DisableFeaturesbyDefault**

如果为 **True**，则 WPS 将在所有文档中禁用 **DisableFeaturesIntroducedAfterbyDefault** 中指定的 WPS 版本后引入的所有功能。默认值为 **False**。**Boolean** 类型，可读写。

**语法**

**express.DisableFeaturesbyDefault**

*express*   一个代表 **Options** 对象的变量。

**说明**

**DisableFeaturesByDefault** 属性设置应用程序的全局选项。如果仅对该文档禁用 WPS 97 for Windows 版本后引入的所有功能，可使用 **DisableFeatures** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例对所有文档禁用 WPS for Windows 的 7.0 和 7.0a 版之后引入的所有功能。*/ function FeaturesDisableByDefault() {     let  options = Application.Options      //Checks whether features are disabled     if ( options.DisableFeaturesbyDefault == true ) {          //If they are, disables all features after  WPS for Windows         options.DisableFeaturesIntroducedAfterbyDefault = wd70     }     else {          //If not, turns on the disable features option and disables         //all features introduced after  WPS for Windows         options.DisableFeaturesbyDefault = true         options.DisableFeaturesIntroducedAfterbyDefault = wd70     } }` |

#### **Options.DisableFeaturesIntroducedAfterbyDefault**

对所有文档禁用指定版本后引入的所有功能。**WdDisableFeaturesIntroducedAfter** 类型，可读写。

**语法**

**express.DisableFeaturesIntroducedAfterbyDefault**

*express*   一个代表 **Options** 对象的变量。

**说明**

设置 **DisableFeaturesIntroducedAfterByDefault** 属性之前，必须先将 **DisableFeaturesByDefault** 属性设为 **True**。否则该设置无效，并保持默认设置 WPS 97 for Windows。

**DisableFeaturesIntroducedAfterByDefault** 属性设置应用程序的全局选项并影响所有文档。如果仅对某文档禁用指定版本后引入的所有功能，可使用 **DisableFeaturesIntroducedAfter** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例对所有文档禁用 WPS for Windows 的 7.0 和 7.0a 版之后引入的所有功能。*/ function FeaturesDisableByDefault() {     let options = Application.Options      //Checks whether features are disabled     if ( options.DisableFeaturesbyDefault == true ) {          //If they are, disables all features after  WPS for Windows         options.DisableFeaturesIntroducedAfterbyDefault = wd70     }     else {          //If not, turns on the disable features option and disables         //all features introduced after  WPS for Windows         options.DisableFeaturesbyDefault = true         options.DisableFeaturesIntroducedAfterbyDefault = wd70     } }` |

#### **Options.DisplayGridLines**

如果 WPS 显示文档网格，则该属性值为 **True**。该属性相当于**“视图”**菜单中的**“网格线”**命令。**Boolean** 类型，可读写。

**语法**

**express.DisplayGridLines**

*express*   一个代表 **Options** 对象的变量。

**说明**

该属性只影响文档网格。对于表格虚框，请使用 **TableGridlines** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例实现的功能是：在显示或隐藏活动窗口的文档网格之间切换。*/ Options.DisplayGridLines = ( !Options.DisplayGridLines )` |

#### **Options.DisplayPasteOptions**

如果为 **True**，则 WPS 直接在新粘贴的文本下显示**“粘贴选项”**按钮。**Boolean** 类型，可读写。

**语法**

**express.DisplayPasteOptions**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果“粘贴选项”按钮无效，则本示例将启用该选项。*/ function ShowPasteOptionsButton() {     if ( Options.DisplayPasteOptions == false ) {         Options.DisplayPasteOptions = true     } }` |

#### **Options.DisplaySmartTagButtons**

如果为 **True**，则 WPS 在鼠标指针置于智能标记上方时显示一个按钮。**Boolean** 类型，可读写。

**语法**

**express.DisplaySmartTagButtons**

*express*   一个代表 **Options** 对象的变量。

**说明**

智能标记按钮提供一个下拉菜单，以便用户访问智能标记选项和操作。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例隐藏在鼠标指针置于智能标记上方时按钮的显示。*/ function DontShowSmartTagButton() {     Options.DisplaySmartTagButtons = false }` |

#### **Options.DocumentViewDirection**

返回或设置整篇文档的对齐方式和阅读顺序。**WdDocumentViewDirection** 类型，可读写。

**语法**

**express.DocumentViewDirection**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将整篇文档设为右对齐方式和从右向左的阅读顺序。*/ Options.DocumentViewDirection = wdDocumentViewRtl` |

#### **Options.DoNotPromptForConvert**

设置或返回一个 **Boolean**，它代表在为处于兼容模式下的文档调用“转换”命令时是否出现一个警告对话框。可读写。

**语法**

**express.DoNotPromptForConvert**

*express*   一个代表 **Options** 对象的变量。

#### **Options.EnableHangulHanjaRecentOrdering**

如果 WPS 在进行朝鲜文字和朝鲜文汉字间的转换时，在建议列表的顶端显示最近用过的单词，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.EnableHangulHanjaRecentOrdering**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例询问用户是否设置 WPS 在进行朝鲜文字和朝鲜文汉字间的转换时，在建议列表的顶端显示最近用过的单词。*/ function test() { let x = MsgBox("Display most recently used words at the top of the suggestions list?", jsYesNo) if ( x == jsResultYes ) {     Options.EnableHangulHanjaRecentOrdering = true } else {     Options.EnableHangulHanjaRecentOrdering = false } }` |

#### **Options.EnableLegacyIMEMode**

返回或设置一个 **Boolean** 类型的值，该值代表是否启用旧式 IME 模式。可读写。

**语法**

**express.EnableLegacyIMEMode**

*express*   一个代表 **Options** 对象的变量。

#### **Options.EnableLivePreview**

设置或返回 **Boolean** 值，该值表示显示还是隐藏在使用支持预览的库时所出现的库预览。如果为 **True**，将在应用命令之前在文档中显示预览。可读/写。

**语法**

**express.EnableLivePreview**

*express*   一个代表 **Options** 对象的变量。

**说明**

此属性对应于**“ WPS 选项”**对话框中的**“启用实时预览”**复选框。

#### **Options.EnableMisusedWordsDictionary**

如果 WPS 在检查文档的拼写和语法时，也检查误用的单词，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.EnableMisusedWordsDictionary**

*express*   一个代表 **Options** 对象的变量。

**说明**

在检查误用单词时， WPS 将检查以下内容：形容词和副词、比较级和最高级、“like”作为连词、“nor”与“or”、“what”与“which”、“who”与“whom”、度量单位、连词、介词和代名词的不正确用法。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设定 WPS 在拼写和语法检查中忽略误用的单词。*/ Options.EnableMisusedWordsDictionary = false` |

#### **Options.EnableSound**

如果 WPS 在导致出错时使计算机发出声音进行响应，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.EnableSound**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例根据用户的输入，设置“选项”对话框“常规”选项卡上的“提供声音反馈”选项。*/ function test() { if ( MsgBox("Do you want  WPS to beep on errors?", 36) == jsResultYes ) {     Options.EnableSound = true } else {     Options.EnableSound = false } }` |

#### **Options.EnvelopeFeederInstalled**

如果当前的打印机为信封指定了送纸器，则该属性值为 **True**。**Boolean** 类型，只读。

**语法**

**express.EnvelopeFeederInstalled**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果已安装了信封送纸器，本示例以信封方式打印活动文档。*/ function test() { if ( Options.EnvelopeFeederInstalled == true ) {     ActiveDocument.Envelope.PrintOut(null,null,null,null,null,null,null,null,null,null,null,null,InchesToPoints(3),InchesToPoints(1.5)) } else {     Msgbox("No envelope feeder available.") } }` |

#### **Options.FormatScanning**

如果为 **True**，则 WPS 对文档中的所有格式进行跟踪。**Boolean** 类型，可读写。

**语法**

**express.FormatScanning**

*express*   一个代表 **Options** 对象的变量。

**说明**

启用 **FormatScanning** 属性允许您识别文档中所有单独的格式，这样您可以方便地将相同的格式应用于新文本，并快速替换或修改文档中出现的所有给定的格式。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例启用 WPS 对文档中格式的跟踪，但禁止在文字下方显示波浪下划线，这些文字的格式类似于文档中使用频率更高的其他格式。*/ function ShowFormatErrors() {     Options.FormatScanning = true     Options.ShowFormatError = false  //Disables displaying squiggly underline }` |

#### **Options.FrenchReform**

返回或设置一个 **WdFrenchSpeller** 常量，该常量代表要对语言格式设置为法语的文本区域使用哪个拼写词典。可读写。

**语法**

**express.FrenchReform**

*express*   一个代表 **Options** 对象的变量。

#### **Options.GridDistanceHorizontal**

当在新文档中绘制、移动和调整自选图形或东亚语言字符时，返回或设置 WPS 使用的不可见网格线的水平间距。**Single** 类型，可读写。

**语法**

**express.GridDistanceHorizontal**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例实现的功能是：设置网格线的水平和垂直间距，并为新文档启动“对象与网格对齐”功能。*/ function test() { Options.GridDistanceHorizontal = InchesToPoints(0.2) Options.GridDistanceVertical = InchesToPoints(0.2) Options.SnapToGrid = true Documents.Add() }` |

#### **Options.GridDistanceVertical**

在新文档中绘制、移动和调整自选图形或东亚语言字符时，返回或设置 WPS 所用的不可见网格线间的垂直间距。**Single** 类型，可读写。

**语法**

**express.GridDistanceVertical**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例实现的功能是：设置网格线的水平和垂直间距，并为新文档启动“对象与网格对齐”功能。*/ function test() { Options.GridDistanceHorizontal = InchesToPoints(0.2) Options.GridDistanceVertical = InchesToPoints(0.2) Options.SnapToGrid = true Documents.Add() }` |

#### **Options.GridOriginHorizontal**

返回或设置不可见网格相对于页面左边的起点位置，当在新文档中绘制、移动和调整自选图形或东亚语言字符时将使用该网格。**Single** 类型，可读写。

**语法**

**express.GridOriginHorizontal**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例实现的功能是：设置网格的水平和垂直起点，以及网格线之间的水平和垂直距离，并为新文档启用“对象与网格对齐”功能。*/ function test() { Options.GridOriginHorizontal = InchesToPoints(1) Options.GridOriginVertical = InchesToPoints(2) Options.GridDistanceHorizontal = InchesToPoints(0.1) Options.GridDistanceVertical = InchesToPoints(0.1) Options.SnapToGrid = true Documents.Add() }` |

#### **Options.GridOriginVertical**

返回或设置不可见网格相对于页面顶边的起点位置，当在新文档中绘制、移动和调整自选图形或东亚语言字符时将使用该网格。**Single** 类型，可读写。

**语法**

**express.GridOriginVertical**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例实现的功能是：设置网格的水平和垂直起点，以及网格线之间的水平和垂直距离，并为新文档启用“对象与网格对齐”功能。*/ function test() { Options.GridOriginHorizontal = InchesToPoints(1) Options.GridOriginVertical = InchesToPoints(2) Options.GridDistanceHorizontal = InchesToPoints(0.2) Options.GridDistanceVertical = InchesToPoints(0.2) Options.SnapToGrid = true Documents.Add() }` |

#### **Options.HangulHanjaFastConversion**

如果 WPS 在朝鲜文字和朝鲜文汉字间进行转换时，自动转换具有单独建议的单词，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.HangulHanjaFastConversion**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例询问用户，在朝鲜文字和朝鲜文汉字间进行转换时，是否设置 WPS 使用快速转换功能。*/ function test() { let x = MsgBox("Use fast conversion?", jsYesNo) if ( x == jsResultYes ) {     Options.HangulHanjaFastConversion = true } else {     Options.HangulHanjaFastConversion = false } }` |

#### **Options.HebrewMode**

返回或设置希伯来语拼写检查工具的模式。**WdHebSpellStart** 类型，可读写。

**语法**

**express.HebrewMode**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置拼写检查工具在进行拼写检查时，遵循希伯来语言学院规定的包含音调符号的文字的常规书写规则。*/ Options.HebrewMode = wdFullScript` |

#### **Options.HideChartDraftModeNotification**

如果在草图模式下插入的图表上隐藏**“草图模式”**通知标签，则该属性值为**“True”**。可读/写。

**语法**

**express.HideChartDraftModeNotification**

*express*   一个代表 **Options** 对象的变量。

**说明**

设置该属性不会影响先前在草图模式下插入的图表。

#### **Options.IgnoreInternetAndFileAddresses**

如果为 **True**，则进行拼写检查时忽略文件扩展名、MS-DOS 路径、电子邮件地址、服务器名称、共享名（也称为 UNC 路径）和 Internet 地址（也称为 URL）。**Boolean** 类型，可读写。

**语法**

**express.IgnoreInternetAndFileAddresses**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 忽略文件名和 Internet 地址，然后对活动文档进行拼写检查。*/ function test() { Options.IgnoreInternetAndFileAddresses = true ActiveDocument.CheckSpelling() }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“选项”对话框中“拼写和语法”选项卡上“忽略 Internet 和文件地址”选项的当前状态。*/ let blnTemp = Options.IgnoreInternetAndFileAddresses` |

#### **Options.IgnoreMixedDigits**

如果进行拼写检查时忽略包含数字的单词，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.IgnoreMixedDigits**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 忽略包含数字的单词，然后对活动文档进行拼写检查。*/ function test() { Options.IgnoreMixedDigits = true ActiveDocument.CheckSpelling() }  ` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“选项”对话框中“拼写和语法”选项卡上“忽略带数字的单词”选项当前的状态。*/ let blnTemp = Options.IgnoreMixedDigits` |

#### **Options.IgnoreUppercase**

如果进行拼写检查时忽略全部大写的单词，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.IgnoreUppercase**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 忽略全部大写的单词，然后对活动文档进行拼写检查。*/ function test() { Options.IgnoreUppercase = true ActiveDocument.CheckSpelling() }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“选项”对话框中“拼写和语法”选项卡上“忽略所有字母都大写的单词”选项当前的状态。*/ let blnTemp = Options.IgnoreUppercase` |

#### **Options.IMEAutomaticControl**

如果设置 WPS 自动打开和关闭日文输入法编辑器 (IME)，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.IMEAutomaticControl**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 自动打开和关闭日文输入法编辑器 (IME)。*/ Options.IMEAutomaticControl = true` |

#### **Options.InlineConversion**

如果 WPS 将中文输入法 (IME) 中的未确认字符串显示为已有（已确认的）字符串之间的插入项，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.InlineConversion**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设定 WPS 将中文输入法 (IME) 中的未确认字符串显示为已有（已确认的）字符串之间的插入项。*/ Options.InlineConversion = true` |

#### **Options.InsertChartInDraftMode**

如果使用草图模式在文档中插入图表，则该属性值为 **True**。可读/写。

**语法**

**express.InsertChartInDraftMode**

*express*   一个代表 **Options** 对象的变量。

**说明**

设置 **InsertChartInDraftMode** 属性与单击 WPS 2015 中**“ WPS 选项”**对话框中的**“使用草图模式插入图表”**具有相同作用（单击“高级”，滚动到“图表”）。

#### **Options.InsertedCellColor**

返回或设置一个 **WdCellColor** 常量，该常量代表插入的表格单元格的颜色。可读写。

**语法**

**express.InsertedCellColor**

*express*   一个代表 **Options** 对象的变量。

#### **Options.InsertedTextColor**

当启用修订选项时，返回或设置插入文字的颜色。**WdColorIndex** 类型，可读写。

**语法**

**express.InsertedTextColor**

*express*   一个代表 **Options** 对象的变量。

**说明**

如果将 **InsertedTextColor** 属性设置为 **wdByAuthor**，WPS 会自动为文档的前八个审阅者各赋予唯一的颜色。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将插入文本的颜色设置为深红色。*/ Options.InsertedTextColor = wdDarkRed` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“选项”对话框中“修订”选项卡上“修订选项”下“颜色”选项的当前状态。*/ let lngColor = Options.InsertedTextColor` |

#### **Options.InsertedTextMark**

返回或设置启用修订时 WPS 设置插入文本格式的方式（**TrackRevisions** 属性为 **True**）。**WdInsertedTextMark** 类型，可读写。

**语法**

**express.InsertedTextMark**

*express*   一个代表 **Options** 对象的变量。

**说明**

如果未启用修订，则会忽略该属性。将该属性与 **InsertedTextColor** 属性结合使用可控制文档中插入文本的外观。

要在编辑过程中查看插入文本的格式，**ShowRevisions** 属性必须为 **True**。如果要 WPS 在打印文档时使用插入文本的格式，**PrintRevisions** 属性必须为 **True**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 为插入的文本应用斜体格式。*/ Options.InsertedTextMark = wdInsertedTextMarkItalic` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 为插入的文本应用加粗格式（如果尚未将其设置为加粗格式）。*/ function test() { if ( Options.InsertedTextMark != wdInsertedTextMarkBold ) {     Options.InsertedTextMark = wdInsertedTextMarkBold } else {     MsgBox("Inserted text is already bold!") } }` |

#### **Options.INSKeyForOvertype**

如果可以用 Ins 键来打开和关闭“改写”，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.INSKeyForOvertype**

*express*   一个代表 **Options** 对象的变量。

#### **Options.INSKeyForPaste**

如果可用 INS 键来粘贴“剪贴板”的内容，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.INSKeyForPaste**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例启用 INS 来粘贴“剪贴板”的内容。*/ Options.INSKeyForPaste = true` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“选项”对话框中“编辑”选项卡上“用 INS 粘贴”选项的状态。*/ let blnTemp = Options.INSKeyForPaste` |

#### **Options.InterpretHighAnsi**

返回或设置高端字符译码方式。**WdHighAnsiText** 类型，可读写。

**语法**

**express.InterpretHighAnsi**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 不将所有高端字符译码为东亚字符。*/ Options.InterpretHighAnsi = wdHighAnsiIsFarEast` |

#### **Options.LabelSmartTags**

如果为 **True**，WPS 使用智能标记信息标记文档中的文本。**Boolean** 类型，可读写。

**语法**

**express.LabelSmartTags**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在文档中关闭智能标记功能。*/ function MarkSmartTags() {     Application.Options.LabelSmartTags = false }` |

#### **Options.LocalNetworkFile**

如果在编辑保存在网络服务器上的文件时，WPS 在用户的计算机上创建文件的一个本地副本，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.LocalNetworkFile**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例指示 WPS 不对保存在服务器上的文件创建本地副本文件。*/ function LocalFile() {     Application.Options.LocalNetworkFile = false }` |

#### **Options.MapPaperSize**

如果为 **True**，则自动调整采用其他国家/地区标准纸张大小（例如 A4）的文档格式，以便按照用户所在国家/地区的标准纸张大小（例如，Letter）正确打印该文档。**Boolean** 类型，可读写。

**语法**

**express.MapPaperSize**

*express*   一个代表 **Options** 对象的变量。

**说明**

本属性只影响您的文档的打印输出，文档的格式不变。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例允许 WPS 根据国家/地区设置调整纸张大小。*/ Options.MapPaperSize = true` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“工具”菜单的“选项”对话框中“打印”选项卡上“允许重调 A4/Letter 纸型”选项的状态。*/ let temp = Options.MapPaperSize` |

#### **Options.MatchFuzzyAY**

如果在搜索过程中 WPS 将忽略 ![img]() 行和 ![img]() 行字符后“ ![img]()”和“ ![img]()”之间的差异，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.MatchFuzzyAY**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设定 WPS 在搜索过程中忽略位于 i 行和 e 行字符后面的字符“ a”与“ ya”的差异。*/ Options.MatchFuzzyAY = true` |

#### **Options.MatchFuzzyBV**

如果 WPS 将在搜索过程中忽略“ ![img]()”与“ ![img]() ![img]()”之间以及“ ![img]()”与“ ![img]() ![img]()”之间的差异，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.MatchFuzzyBV**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设定 WPS 在搜索过程中忽略“ ba”与“ vua vua”之间以及“ ha”与“ fua fua”之间的差异。*/ Options.MatchFuzzyBV = true` |

#### **Options.MatchFuzzyByte**

如果在搜索过程中 WPS 将忽略全角和半角字符（西文或日文）的差异，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.MatchFuzzyByte**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 WPS 设置为在搜索过程中忽略全角和半角字符（西文或日文）的差异。*/ Options.MatchFuzzyByte = true` |

#### **Options.MatchFuzzyCase**

如果 WPS 在搜索过程中不区分字母大小写，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.MatchFuzzyCase**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设定 WPS 在搜索过程中不区分字母的大小写。*/ Options.MatchFuzzyCase = true` |

#### **Options.MatchFuzzyDash**

如果 WPS 在搜索过程中忽略减号、划线以及长元音符号之间的差异，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.MatchFuzzyDash**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 WPS 设置为在搜索过程中忽略减号、划线以及长元音符号之间的差异。*/ Options.MatchFuzzyDash = true` |

#### **Options.MatchFuzzyDZ**

如果 WPS 将在搜索过程中忽略“ ![img]()”与“ ![img]()”之间以及“ ![img]()”与“ ![img]()”之间的差异，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.MatchFuzzyDZ**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设定 WPS 在搜索过程中忽略“ di”与“ zi”之间以及“ du”与“ zu”之间的差异。*/ Options.MatchFuzzyDZ = true` |

#### **Options.MatchFuzzyHF**

如果 WPS 将在搜索过程中忽略“ ![img]() ![img]()”与“ ![img]() ![img]()”之间以及“ ![img]() ![img]()”与“ ![img]() ![img]()”之间的差异，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.MatchFuzzyHF**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设定 WPS 在搜索过程中忽略“ heyu heyu”与“ fuyu fuyu”之间以及“ beyu beyu”与“ vuyu vuyu”之间的差异。*/ Options.MatchFuzzyHF = true` |

#### **Options.MatchFuzzyHiragana**

如果 WPS 在搜索过程中将忽略平假名与片假名之间的差异，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.MatchFuzzyHiragana**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 WPS 设置为在搜索过程中忽略平假名与片假名之间的差异。*/ Options.MatchFuzzyHiragana = true` |

#### **Options.MatchFuzzyIterationMark**

如果 WPS 在搜索过程中将忽略重复的标记类型间的差异，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.MatchFuzzyIterationMark**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 WPS 设置为在搜索过程中忽略重复的标记类型间的差异。*/ Options.MatchFuzzyIterationMark = true` |

#### **Options.MatchFuzzyKanji**

如果 WPS 在搜索过程中忽略标准和非标准日文汉字的差异，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.MatchFuzzyKanji**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 WPS 设置为在搜索过程中忽略标准和非标准日文汉字的差异。*/ Options.MatchFuzzyKanji = true` |

#### **Options.MatchFuzzyKiKu**

如果在搜索过程中 WPS 将忽略 ![img]() 行字符前“ ![img]()”和“ ![img]()”之间的差异，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.MatchFuzzyKiKu**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设定 WPS 在搜索过程中忽略位于 sa 行字符之前的字符“ ki”和“ ku”之间的差异。*/ Options.MatchFuzzyKiKu = true` |

#### **Options.MatchFuzzyOldKana**

如果 WPS 在搜索过程中将忽略新旧假名字符之间的差异，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.MatchFuzzyOldKana**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 WPS 设置为在搜索过程中忽略新旧假名字符的差异。*/ Options.MatchFuzzyOldKana = true` |

#### **Options.MatchFuzzyProlongedSoundMark**

如果在搜索过程中 WPS 将忽略短元音和长元音的差异，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.MatchFuzzyProlongedSoundMark**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 WPS 设置为在搜索过程中忽略短元音和长元音的差异。*/ Options.MatchFuzzyProlongedSoundMark = true` |

#### **Options.MatchFuzzyPunctuation**

如果在搜索过程中 WPS 忽略标点类型的差异，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.MatchFuzzyPunctuation**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 WPS 设置为在搜索过程中忽略标点的类型差异。*/ Options.MatchFuzzyPunctuation = true` |

#### **Options.MatchFuzzySmallKana**

如果 WPS 在搜索过程中忽略双元音和双辅音的差异，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.MatchFuzzySmallKana**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 WPS 设置为在搜索过程中忽略双元音和双辅音的差异。*/ Options.MatchFuzzySmallKana = true` |

#### **Options.MatchFuzzySpace**

如果 WPS 在搜索过程中忽略空格标记的差异，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.MatchFuzzySpace**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 WPS 设置为在搜索过程中忽略使用的空格标记之间的差异。*/ Options.MatchFuzzySpace = true` |

#### **Options.MatchFuzzyTC**

如果在搜索过程中 WPS 将忽略“ ![img]() ![img]()”、“ ![img]() ![img]()”与“ ![img]()”之间以及“ ![img]() ![img]()”与“ ![img]()”之间的差异，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.MatchFuzzyTC**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设定 WPS 在搜索过程中忽略“ tsui tsui”、“ tei tei”与“ chi”之间以及“ dei dei”与“ ji”之间的差异。*/ Options.MatchFuzzyTC = true` |

#### **Options.MatchFuzzyZJ**

如果 WPS 将在搜索过程中忽略“ ![img]()”与“ ![img]() ![img]()”之间以及“ ![img]()”与“ ![img]() ![img]()”之间的差异，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.MatchFuzzyZJ**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设定 WPS 在搜索过程中忽略“ se”与“ shie shie”之间以及“ ze”与“ jie jie”之间的差异。*/ Options.MatchFuzzyZJ = true` |

#### **Options.MeasurementUnit**

设置或者返回 WPS 的标准度量单位。**WdMeasurementUnits** 类型，可读写。

**语法**

**express.MeasurementUnit**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 WPS 的标准度量单位设置为磅。*/ Options.MeasurementUnit = wdPoints` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“选项”对话框（“工具”菜单）中“常用”选项卡上当前所选定的度量单位。*/ let CurrUnit = Options.MeasurementUnit` |

#### **Options.MergedCellColor**

返回或设置一个 **WdCellColor** 常量，该常量代表合并的表格单元格的颜色。可读写。

**语法**

**express.MergedCellColor**

*express*   一个代表 **Options** 对象的变量。

**说明**

此属性仅适用于对其运行了 **CompareDocuments**、**Compare** 或 **Merge** 方法的文档。

#### **Options.MonthNames**

返回或设置朝鲜文字和朝鲜文汉字之间的转换方向。**WdMonthNames** 类型，可读写。

**语法**

**express.MonthNames**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设定 WPS 在默认情况下将朝鲜文字转换成朝鲜文汉字。*/ Options.MultipleWordConversionsMode = wdHangulToHanja` |

#### **Options.MoveFromTextColor**

返回或设置一个 **WdColorIndex** 常量，该常量代表所移动文本的颜色。可读写。

**语法**

**express.MoveFromTextColor**

*express*   一个代表 **Options** 对象的变量。

#### **Options.MoveFromTextMark**

返回或设置一个 **WdMoveFromTextMark** 常量，该常量代表要用于所移动文本的修订标记的类型。可读写。

**语法**

**express.MoveFromTextMark**

*express*   一个代表 **Options** 对象的变量。

#### **Options.MoveToTextColor**

返回或设置一个 **WdColorIndex** 常量，该常量代表所移动文本的颜色。可读写。

**语法**

**express.MoveToTextColor**

*express*   一个代表 **Options** 对象的变量。

#### **Options.MoveToTextMark**

返回或设置一个 **WdMoveToTextMark** 常量，该常量代表要用于所移动文本的修订标记的类型。可读写。

**语法**

**express.MoveToTextMark**

*express*   一个代表 **Options** 对象的变量。

#### **Options.MultipleWordConversionsMode**

返回或设置朝鲜文字和朝鲜文汉字之间的转换方向。**WdMultipleWordConversionsMode** 类型，可读写。

**语法**

**express.MultipleWordConversionsMode**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设定 WPS 在默认情况下将朝鲜文字转换成朝鲜文汉字。*/ Options.MultipleWordConversionsMode = wdHangulToHanja` |

#### **Options.OMathAutoBuildUp**

返回或设置一个 **Boolean** 类型的值，该值代表 WPS 是否自动将公式转换为专业格式。如果为 **True**，则表示 WPS 自动将公式转换为专业格式。可读写。

**语法**

**express.OMathAutoBuildUp**

*express*   一个代表 **Options** 对象的变量。

#### **Options.OMathCopyLF**

返回或设置一个 **Boolean** 类型的值，该值代表如何以纯文本的形式表示公式。如果为 **True**，则表示以线性格式表示公式。如果为 **False**，则表示以 MathML 格式表示公式。可读写。

**语法**

**express.OMathCopyLF**

*express*   一个代表 **Options** 对象的变量。

#### **Options.OptimizeForWord97byDefault**

如果 WPS 将禁用所有不兼容的格式，以便在 WPS 97 中查看新文档而对其进行优化，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.OptimizeForWord97byDefault**

*express*   一个代表 **Options** 对象的变量。

**说明**

若要为 WPS 97 优化单个文档，可用 **OptimizeForWord97** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例实现的功能是：先设定 WPS 禁用新文档中所有与 WPS 97 不兼容的格式，然后新建一个文档，其 OptimizeForWord97 属性自动设置为 True。*/ function test() { Options.OptimizeForWord97byDefault = true MsgBox(Documents.Add(undefined, undefined, wdNewBlankDocument).OptimizeForWord97) }  ` |

#### **Options.Overtype**

如果激活改写模式，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.Overtype**

*express*   一个代表 **Options** 对象的变量。

**说明**

在改写模式中，键入的字符逐一替代已有的字符。如果改写模式没有被激活，则键入的字符将已有文本右移。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果“改写”模式已激活，则本示例显示一个信息框，询问是否将“改写”模式改为非激活状态。如果用户单击“是”按钮，则“改写”模式变为非激活状态。*/ function test() { if ( Options.Overtype == true ) {     let aButton = MsgBox("Overtype is on. Turn off?", 4)     if ( aButton == jsResultYes ) {         Options.Overtype = false     } } }` |

#### **Options.Pagination**

如果 WPS 在后台重新为文档分页，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.Pagination**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 在后台重新分页。*/ Options.Pagination = true` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“工具”菜单的“选项”对话框中“常规”选项卡上“后台重新分页”选项的当前状态。*/ let temp = Options.Pagination` |

#### **Options.Parent**

返回一个 **Object** 类型值，该值代表指定 **Options** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **Options** 对象的变量。

#### **Options.PasteAdjustParagraphSpacing**

如果在剪切和粘贴选定内容时，WPS 自动调整段落间距，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.PasteAdjustParagraphSpacing**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果禁用该选项，本示例设置 WPS 在剪切和粘贴选定内容时自动调整段落间距。*/ function AdjustParaSpace() {     if ( Options.PasteAdjustParagraphSpacing == false ) {         Options.PasteAdjustParagraphSpacing = true     } }` |

#### **Options.PasteAdjustTableFormatting**

如果在剪切和粘贴选定内容时，WPS 自动调整表格的格式，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.PasteAdjustTableFormatting**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果禁用该选项，本示例设置 WPS 在剪切和粘贴时自动设置表格的格式。*/ function AdjustTableFormatting() {     if ( Options.PasteAdjustTableFormatting == false ) {         Options.PasteAdjustTableFormatting = true     } }` |

#### **Options.PasteAdjustWordSpacing**

如果在剪切和粘贴选定内容时，WPS 自动调整字符间距，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.PasteAdjustWordSpacing**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果禁用该选项，本示例设置 WPS 在剪切和粘贴选定文本时自动调整字符间距。*/ function AdjustWordSpace() {     if ( Options.PasteAdjustWordSpacing == false ) {         Options.PasteAdjustWordSpacing = true     } }` |

#### **Options.PasteFormatBetweenDocuments**

返回或设置一个 [WdPasteOptions ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdPasteOptions%20%E6%9E%9A%E4%B8%BE.html)常量，该常量代表在从另一个 WPS Office WPS 文档复制文本时粘贴文本的方式。可读写。

**语法**

**express.PasteFormatBetweenDocuments**

*express*   一个代表 **Options** 对象的变量。

**说明**

相当于**“ WPS 选项”**对话框的**“高级”**选项卡中的**“在两个文档之间粘贴(不带样式)”**选项。

#### **Options.PasteFormatBetweenStyledDocuments**

返回或设置一个 **WdPasteOptions** 常量，该常量代表在从使用样式的文档中复制文本时粘贴文本的方式。可读写。

**语法**

**express.PasteFormatBetweenStyledDocuments**

*express*   一个代表 **Options** 对象的变量。

**说明**

相当于**“ WPS 选项”**对话框的**“高级”**选项卡中的**“在两个文档之间粘贴(带样式)”**选项。

#### **Options.PasteFormatFromExternalSource**

返回或设置一个 [WdPasteOptions ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdPasteOptions%20%E6%9E%9A%E4%B8%BE.html)常量，该常量代表在从外部源（如网页）复制文本时粘贴文本的方式。可读写。

**语法**

**express.PasteFormatFromExternalSource**

*express*   一个代表 **Options** 对象的变量。

**说明**

相当于**“ WPS 选项”**对话框的**“高级”**选项卡中的**“从其他程序粘贴”**选项。

#### **Options.PasteFormatWithinDocument**

返回或设置一个 [WdPasteOptions ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdPasteOptions%20%E6%9E%9A%E4%B8%BE.html)常量，该常量代表在同一文档内复制或剪切文本然后进行粘贴时的文本粘贴方式。可读写。

**语法**

**express.PasteFormatWithinDocument**

*express*   一个代表 **Options** 对象的变量。

**说明**

相当于**“ WPS 选项”**对话框的**“高级”**选项卡中的**“在同一文档内粘贴”**选项。

#### **Options.PasteMergeFromPPT**

如果为 **True**，则从 Microsoft PowerPoint 粘贴文本时合并文本格式。**Boolean** 类型，可读写。

**语法**

**express.PasteMergeFromPPT**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果禁用该选项，本示例设置 WPS 从 PowerPoint 粘贴内容时自动合并文本格式。*/ function AdjustPPTFormatting() {     if ( Options.PasteMergeFromPPT == false ) {         Options.PasteMergeFromPPT = true     } }` |

#### **Options.PasteMergeFromXL**

如果为 **True**，则从 ET 粘贴表格时合并表格的格式。**Boolean** 类型，可读写。

**语法**

**express.PasteMergeFromXL**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果禁用该选项，本示例设置 WPS 在粘贴 ET 表格时自动合并表格的格式。*/ function AdjustExcelFormatting() {     if ( Options.PasteMergeFromXL == false ) {         Options.PasteMergeFromXL = true     } }` |

#### **Options.PasteMergeLists**

如果为 **True**，则合并粘贴列表和周围列表的格式。**Boolean** 类型，可读写。

**语法**

**express.PasteMergeLists**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果禁用该选项，则本示例设置 WPS 自动合并粘贴列表和周围列表的格式。*/ function UseSmartStyle() {     if ( Options.PasteMergeLists == false ) {         Options.PasteMergeLists = true     } }` |

#### **Options.PasteOptionKeepBulletsAndNumbers**

返回或设置 **Boolean** 值，该值表示在从**“粘贴选项”**上下文菜单中选择**“仅保留文本”**时是否保留项目符号和编号。可读/写。

**语法**

**express.PasteOptionKeepBulletsAndNumbers**

*express*   一个代表 **Options** 对象的变量。

**说明**

此属性对应于**“ WPS 选项”**对话框的**“高级”**选项卡中的**“使用‘仅保留文本’选项粘贴文本时保留项目符号和编号”**复选框。

#### **Options.PasteSmartCutPaste**

如果 WPS 在文档中智能粘贴选定内容，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.PasteSmartCutPaste**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果禁用该选项，本示例设置 WPS 智能粘贴选定文本。*/ function EnableSmartCutPaste() {     if ( Options.PasteSmartCutPaste == false ) {         Options.PasteSmartCutPaste = true     } }` |

#### **Options.PasteSmartStyleBehavior**

从其他文档粘贴选定文本时，如果 WPS 智能合并样式，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.PasteSmartStyleBehavior**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果禁用该选项，本示例会将 WPS 设置为从其他文档的选定文本中智能粘贴样式。*/ function UseSmartStyle() {     if ( Options.PasteSmartStyleBehavior == false ) {         Options.PasteSmartStyleBehavior = true     } }` |

#### **Options.PictureEditor**

返回或设置用于编辑图片的应用程序的名称。**String** 类型，可读写。

**语法**

**express.PictureEditor**

*express*   一个代表 **Options** 对象的变量。

**说明**

必须使用**“工具”**菜单**“选项”**对话框中**“编辑”**选项卡上“图片编辑器”框中显示的正确单词。否则，将使用默认设置“WPS ”。

如果所需的图形应用程序的名称没有显示在列表中，请和图形应用程序的制造商联系咨询。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置用于编辑图片的应用程序。*/ Options.PictureEditor = "WPS "` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回用于编辑图片的应用程序的名称。*/ MsgBox(Options.PictureEditor)` |

#### **Options.PictureWrapType**

设置或返回一个 **WdWrapTypeMerged** 类型的值，该值表示 WPS 在图片周围环绕文字的方式。可读写。

**语法**

**express.PictureWrapType**

*express*   一个代表 **Options** 对象的变量。

**说明**

这是默认的选项设置并对所有插入的图片有效，除非为图片单独定义图片环绕。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果尚未指定嵌入格式，本示例将 WPS 设置为以嵌入方式在文本中插入并粘贴所有图片。*/ function PicWrap() {     let  options = Application.Options     if ( options.PictureWrapType != wdWrapMergeInline ) {         options.PictureWrapType = wdWrapMergeInline     } }` |

#### **Options.PortugalReform**

返回或设置欧洲葡萄牙语拼写器的模式。可读/写 WdPortugueseReform。

**语法**

**express.PortugalReform**

*express*   一个代表 **Options** 对象的变量。

**说明**

设置该属性与在**“ WPS 选项”**对话框中**“欧洲葡萄牙语模式:”**旁边的下拉框中选择项目具有相同作用（**“校对”**、**“在 WPS Office程序中更正拼写时”**）。

| 注释                                                         |
| ------------------------------------------------------------ |
| 该属性不设置巴西葡萄牙语拼写器的模式。若要设置巴西葡萄牙语拼写器模式，请使用 [Options.BrazilReform](https://qn.cache.wpscdn.cn/encs/doc/office_v19/BrazilReform%20%E5%B1%9E%E6%80%A7.htm) 属性。 |

#### **Options.PrecisePositioning**

返回或设置一个 **Boolean** 类型的值，该值表示 WPS 是否针对打印版式而不是针对屏幕可读性优化字符位置。如果该属性的值为 **True**，则禁用压缩字符间距以便增强屏幕可读性的默认设置，并启用适用于打印介质的字符间距。可读/写。

**语法**

**express.PrecisePositioning**

*express*   一个代表 **Options** 对象的变量。

#### **Options.PrintBackground**

如果 WPS 在后台打印，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.PrintBackground**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 WPS 设置为在后台打印文档，然后打印活动文档。*/ function test() { Options.PrintBackground = true ActiveDocument.PrintOut() }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“工具”菜单的“选项”对话框中“打印”选项卡上“后台打印”选项的当前状态。*/ let temp = Options.PrintBackground` |

#### **Options.PrintBackgrounds**

返回一个 **Boolean** 类型的值，该值代表打印文档时，是否打印背景颜色和图像。

**语法**

**express.PrintBackgrounds**

*express*   一个代表 **Options** 对象的变量。

**说明**

如果为 **True**，代表打印背景颜色和图像。如果为 **False**，则代表不打印背景颜色和图像。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下列示例指定打印文档的同时要打印背景颜色和图像。*/ Options.PrintBackgrounds = true` |

#### **Options.PrintComments**

如果 WPS 在文档后另起一页打印批注，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.PrintComments**

*express*   一个代表 **Options** 对象的变量。

**说明**

如果将 **PrintComments** 属性设置为 **True**，则 **PrintHiddenText** 属性将自动设置为 **True**。但是，如果将 **PrintComments** 属性设置为 **False**，不会影响到 **PrintHiddenText** 属性的设置。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 WPS 设置为打印批注，然后打印活动文档。*/ function test() { Options.PrintComments = true ActiveDocument.PrintOut() }  ` |

#### **Options.PrintDraft**

如果 WPS 以最简单的格式打印文档，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.PrintDraft**

*express*   一个代表 **Options** 对象的变量。

**说明**

不是所有的打印机都支持草稿输出。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 WPS 设置为使用草稿输出，然后打印活动文档。*/ function test() { Options.PrintDraft = true ActiveDocument.PrintOut() }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“工具”菜单“选项”对话框中“打印”选项卡上“草稿输出”选项的当前状态。*/ let temp = Options.PrintDraft` |

#### **Options.PrintDrawingObjects**

如果 WPS 打印图形对象，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.PrintDrawingObjects**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 WPS 设置为打印图形对象，然后打印活动文档。*/ function test() { Options.PrintDrawingObjects = true ActiveDocument.PrintOut() }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“工具”菜单“选项”对话框中“打印”选项卡上“图形对象”选项的当前状态。*/ let temp = Options.PrintDrawingObjects` |

#### **Options.PrintEvenPagesInAscendingOrder**

如果 WPS 在手动双面打印模式下按升序打印偶数页，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.PrintEvenPagesInAscendingOrder**

*express*   一个代表 **Options** 对象的变量。

**说明**

如果将 **PrintOut** 方法的 *ManualDuplexPrint* 参数设为 **False**，该属性将被忽略。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例先设定 WPS 在手动双面打印模式下按升序打印奇数页，按降序打印偶数页，然后以此方式打印活动文档。*/ function test() { Options.PrintOddPagesInAscendingOrder = true Options.PrintEvenPagesInAscendingOrder = false ActiveDocument.PrintOut(null,null,null,null,null,null,null,null,null,null,null,null,null,null,true) }` |

#### **Options.PrintFieldCodes**

如果 WPS 打印域代码而不打印域的结果，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.PrintFieldCodes**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 WPS 设置为打印域代码，然后打印活动文档。*/ function test() { Options.PrintFieldCodes = true ActiveDocument.PrintOut() }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“工具”菜单“选项”对话框中“打印”选项卡上“域代码”选项的当前状态。*/ let temp = Options.PrintFieldCodes` |

#### **Options.PrintHiddenText**

如果打印隐藏文字，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.PrintHiddenText**

*express*   一个代表 **Options** 对象的变量。

**说明**

如果将 **PrintHiddenText** 属性设置为 **False**，则**PrintComments**属性将自动设置为 **False**。但是，如果将 **PrintHiddenText** 属性设置为 **True**，不会影响到**PrintComments**属性的设置。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 WPS 设置为打印隐藏文字，然后打印活动文档。*/ Application.Options.PrintHiddenText = true Application.ActiveDocument.PrintOut()  /*本示例返回“选项”对话框中“打印”选项卡上“隐藏文字”选项的当前状态。*/ let temp = Application.Options.PrintHiddenText` |

#### **Options.PrintProperties**

如果 WPS 在文档结尾处另起一页打印文档摘要信息，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.PrintProperties**

*express*   一个代表 **Options** 对象的变量。

**说明**

如果不打印文档摘要信息，则为 **False**。可在**“文件”**菜单的**“属性”**对话框中找到摘要信息。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 在文档结尾另起一页打印文档的摘要信息，然后打印活动文档。*/ function test() { Options.PrintProperties = true ActiveDocument.PrintOut() }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“工具”菜单“选项”对话框中“打印”选项卡上“文档属性”选项的当前状态。*/ let temp = Options.PrintProperties` |

#### **Options.PrintReverse**

如果 WPS 逆页序打印页面，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.PrintReverse**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 WPS 设置为逆页序打印页面，然后打印活动文档。*/ function test() { Options.PrintReverse = true ActiveDocument.PrintOut() }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“工具”菜单“选项”对话框中“打印”选项卡上“逆页序打印”选项的当前状态。*/ let temp = Options.PrintReverse` |

#### **Options.PrintXMLTag**

返回一个 **Boolean** 类型的值，该值代表打印文档时是否打印 XML 标记。对应于**“选项”**对话框中**“打印”**选项卡上的**“XML 标记”**复选框。

**语法**

**express.PrintXMLTag**

*express*   一个代表 **Options** 对象的变量。

**说明**

如果为 **True**，则代表打印标记。如果为 **False**，则代表不打印标记。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下列示例指定打印文档的同时要打印标记。*/ Options.PrintXMLTag = true` |

#### **Options.PromptUpdateStyle**

如果为 **True**，则在改变样式的格式时显示一条消息，请用户确认是否要重新设置样式的格式或重新应用原样式格式。**Boolean** 类型，可读写。

**语法**

**express.PromptUpdateStyle**

*express*   一个代表 **Options** 对象的变量。

**说明**

如果为 **False**，则对选定内容重新应用样式的格式，而无需确认用户是否要改变样式。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例检查在更新样式时用户是否会收到消息，如果没有，则启用该功能。*/ function UpdateStylePrompt() {     let opt = Application.Options     if(opt.PromptUpdateStyle == false) {        opt.PromptUpdateStyle = true     } }` |

#### **Options.RepeatWPS**

返回或设置 **Boolean** 值，该值表示在检查拼写时是否标记重复的单词。如果为 **True**，将标记重复的单词。可读/写。

**语法**

**express.RepeatWPS**

*express*   一个代表 **Options** 对象的变量。

**说明**

此属性对应于**“ WPS 选项”**对话框的**“校对”**选项卡上的**“标记重复单词”**复选框。

#### **Options.ReplaceSelection**

如果键入或粘贴的内容替换选定内容，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.ReplaceSelection**

*express*   一个代表 **Options** 对象的变量。

**说明**

如果在选定内容之前添加键入或粘贴的内容，并保留选定内容，则为 **False**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 WPS 设置为在选定内容之前添加键入或粘贴的内容，并保留选定内容。*/ Application.Options.ReplaceSelection = false  /*本示例返回“工具”菜单“选项”对话框中“编辑”选项卡上“键入内容替换选定内容”选项的状态。*/ let temp = Application.Options.ReplaceSelection` |

#### **Options.RevisedLinesColor**

当激活修订选项时，返回或设置文档中修订线的颜色。**WdColorIndex** 类型，可读/写。

**语法**

**express.RevisedLinesColor**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将修订线的颜色设置为粉红色。*/ Options.RevisedLinesColor = wdPink` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“选项”对话框（“工具”菜单）中“修订”选项卡上的“修订行”下的“颜色”选项的当前状态。*/ let temp = Options.RevisedLinesColor` |

#### **Options.RevisedLinesMark**

当激活修订选项时，返回或设置修订线在文档中的位置。**WdRevisedLinesMark** 类型，可读/写。

**语法**

**express.RevisedLinesMark**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置修订线显示在每页的左边距上。*/ Options.RevisedLinesMark = wdRevisedLinesMarkLeftBorder` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“选项”对话框（“工具”菜单）中“修订”选项卡上的“修订行”下的“标记”选项的当前状态。*/ let temp = Options.RevisedLinesMark` |

#### **Options.RevisedPropertiesColor**

在启用修订功能时，返回或设置用于标记格式更改情况的颜色。**WdColorIndex** 类型，可读写。

**语法**

**express.RevisedPropertiesColor**

*express*   一个代表 **Options** 对象的变量。

**说明**

如果删除或插入的文本在格式上有所更改，则使用 **DeletedTextColor** 或 **InsertedTextColor** 属性覆盖 **RevisedPropertiesColor** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例跟踪活动文档的修订，将已改变格式的文本的颜色设置为兰绿色，并将加粗格式应用于所选内容。*/ function test() { ActiveDocument.TrackRevisions = true Options.RevisedPropertiesColor = wdTeal Selection.Font.Bold = true }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“选项”对话框（“工具”菜单）中“修订”选项卡上“修订选项”下“颜色”框中所选中的选项。*/ let temp = Options.RevisedPropertiesColor` |

#### **Options.RevisedPropertiesMark**

在启用修订功能时，返回或设置用于显示格式修改情况的标记。**WdRevisedPropertiesMark** 类型，可读/写。

**语法**

**express.RevisedPropertiesMark**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在启用修订功能后，用双下划线标记修改过格式的文本。*/ Options.RevisedPropertiesMark = wdRevisedPropertiesMarkDoubleUnderline    ` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“选项”对话框（“工具”菜单）中“修订”选项卡上“修订”下“格式”框中所选中的选项。*/ let temp = Options.RevisedPropertiesMark` |

#### **Options.RevisionsBalloonPrintOrientation**

返回或设置一个 **WdRevisionsBalloonPrintOrientation** 常量，该常量表示打印时修订和批注气球的方向。可读写。

**语法**

**express.RevisionsBalloonPrintOrientation**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例以横向格式打印带有批注的文档，修订和批注气球在页面的一侧，文本在另一侧。*/ function PrintLandscapeCommentBalloons() {     Options.RevisionsBalloonPrintOrientation = wdBalloonPrintOrientationForceLandscape }` |

#### **Options.RTFInClipboard**

您请求了仅用于 Macintosh 上的 Visual Basic 关键字的帮助。有关 **Options** 对象的 **RTFInClipboard** 属性的信息，请参阅 WPS OfficeMacintosh Edition 中附带的语言参考帮助。

**语法**

**express.RTFInClipboard**

*express*   一个代表 **Options** 对象的变量。

#### **Options.SaveInterval**

返回或设置保存“自动恢复”信息的时间间隔（以分钟为单位）。**Long** 类型，可读写。

**语法**

**express.SaveInterval**

*express*   一个代表 **Options** 对象的变量。

**说明**

将 **SaveInterval** 属性设置为 0（零）时，将关闭保存“自动恢复”信息的功能。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 WPS 设置为每隔五分钟保存所有打开文档的“自动恢复”信息。*/ Options.SaveInterval = 5` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 WPS 设置为不保存“自动恢复”信息。*/ Options.SaveInterval = 0` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“工具”菜单“选项”对话框中“保存”选项卡上“自动保存时间间隔:”选项的当前状态。*/ let temp = Options.SaveInterval` |

#### **Options.SaveNormalPrompt**

如果 WPS 在关闭之前提示用户确认是否保存对 Normal 模板所做的更改，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.SaveNormalPrompt**

*express*   一个代表 **Options** 对象的变量。

**说明**

如果 WPS 在关闭之前自动保存对 Normal 模板所做的更改，则为 **False**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 WPS 设置为在关闭之前自动保存对 Normal 模板所做的更改，然后退出。*/ function test() { Options.SaveNormalPrompt = false Application.Quit() }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“工具”菜单“选项”对话框中“保存”选项卡上“提示保存 Normal 模板”选项的当前状态。*/ let temp = Options.SaveNormalPrompt` |

#### **Options.SavePropertiesPrompt**

如果 WPS 在保存新的文档时，提示该文档的属性信息，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.SavePropertiesPrompt**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 WPS 设置为在保存新文档时，提示该文档的属性信息。*/ Options.SavePropertiesPrompt = true` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“工具”菜单“选项”对话框中“保存”选项卡上“提示保存文档属性”选项的当前状态。*/ let temp = Options.SavePropertiesPrompt` |

#### **Options.SendMailAttach**

如果使用**“文件”**菜单上的**“发送”**命令将活动文档作为附件插入邮件中，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.SendMailAttach**

*express*   一个代表 **Options** 对象的变量。

**说明**

如果使用**“发送”**命令将活动文档的内容作为文本插入邮件中，则为 **False**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例打开一个新电子邮件，并将活动文档作为邮件附件。*/ function test() { Options.SendMailAttach = true ActiveDocument.SendMail() }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“选项”对话框中“常规”选项卡上“作为附件发送”选项的状态。*/ MsgBox(Options.SendMailAttach)` |

#### **Options.SequenceCheck**

如果为 **True**，则检查南亚文本独立字符的顺序。**Boolean** 类型，可读写。

**语法**

**express.SequenceCheck**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例启用顺序检查，使用户可以键入有效的独立字符顺序以组成南亚文本的有效字符单元格。*/ function CheckSequence() {     Options.SequenceCheck = true }` |

#### **Options.ShortMenuNames**

您请求了仅用于 Macintosh 上的 Visual Basic 关键字的帮助。有关 **Options** 对象的 **ShortMenuNames** 属性的信息，请参阅 WPS OfficeMacintosh Edition 中附带的语言参考帮助。

**语法**

**express.ShortMenuNames**

*express*   一个代表 **Options** 对象的变量。

#### **Options.ShowControlCharacters**

如果显示当前文档中的双向控制符，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.ShowControlCharacters**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在当前文档中隐藏双向控制符。*/ Options.ShowControlCharacters = false` |

#### **Options.ShowDevTools**

返回或设置 **Boolean** 类型的值，该值代表是否在功能区中显示**“开发人员”**选项卡。可读/写。

**语法**

**express.ShowDevTools**

*express*   一个代表 **Options** 对象的变量。

**说明**

此属性对应于**“ WPS 选项”** 对话框中的**“在功能区中显示‘开发人员’选项卡”**复选框。

#### **Options.ShowDiacritics**

如果在使用从右向左语言的文档中显示音调符号，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.ShowDiacritics**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例隐藏当前文档中的音调符号。*/ Options.ShowDiacritics = false` |

#### **Options.ShowFormatError**

如果为 **True**，WPS 通过在文本（该文本的格式与文档中使用频繁的其他格式相似）下加弯曲下划线来标记不一致的格式。**Boolean** 类型，可读写。

**语法**

**express.ShowFormatError**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例使 WPS 可以保留文档中格式的修订，但不在文本下显示弯曲的下划线。*/ function ShowFormatErrors() {     let opt =  Options     opt.FormatScanning = true  //Enables keeping track of formatting     opt.ShowFormatError = false }` |

#### **Options.ShowMarkupOpenSave**

返回或设置一个 **Boolean** 类型的值，该值代表打开或保存文件时 WPS 是否显示隐藏标记。

**语法**

**express.ShowMarkupOpenSave**

*express*   一个代表 **Options** 对象的变量。

**说明**

**ShowMarkupOpenSave** 属性对应于**“选项”**对话框的**“安全性”**选项卡上的**“打开或保存时标记可见”**选项。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例启用“打开或保存时标记可见”选项。*/ Options.ShowMarkupOpenSave = true` |

#### **Options.ShowMenuFloaties**

返回或设置 **Boolean** 值，该值表示当用户在文档窗口中右键单击时是否显示浮动工具栏。可读/写。

**语法**

**express.ShowMenuFloaties**

*express*   一个代表 **Options** 对象的变量。

#### **Options.ShowReadabilityStatistics**

如果 WPS 在结束语法检查时，显示统计信息的摘要列表（包括可读性程度），则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.ShowReadabilityStatistics**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 WPS 设置为显示可读性统计信息，然后检查活动文档的拼写和语法。*/ function test() { Options.ShowReadabilityStatistics = true ActiveDocument.CheckGrammar() }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“工具”菜单“选项”对话框中“拼写和语法”选项卡上“显示可读性统计信息”选项的当前状态。*/ let temp = Options.ShowReadabilityStatistics` |

#### **Options.ShowSelectionFloaties**

返回或设置 **Boolean** 值，该值表示当用户选择文本时是否显示浮动工具栏。可读/写。

**语法**

**express.ShowSelectionFloaties**

*express*   一个代表 **Options** 对象的变量。

**说明**

对应于**“ WPS 选项”**对话框中的**“选择时显示浮动工具栏”**复选框。

#### **Options.SmartCursoring**

返回或设置一个 **Boolean** 类型的值，该值表示是否启用智能指针。如果值为 **True**，则启用智能指针。

**语法**

**express.SmartCursoring**

*express*   一个代表 **Options** 对象的变量。

**说明**

**SmartCursoring** 属性对应于**“选项”**对话框的**“编辑”**选项卡中的**“使用智能指针”**选项，该选项在默认情况下处于选中状态。

**SmartCursoring** 属性为 **True** 时，使用 Page Down 键在文档中滚动会使指针移到当前页面。如果 **SmartCursoring** 属性为 **False**，则指针停留在上次编辑的位置。

**示例**

| 示例代码复制                                                |
| ----------------------------------------------------------- |
| `/*以下示例禁用智能指针。*/ Options.SmartCursoring = false` |

#### **Options.SmartCutPaste**

如果在剪切和粘贴时，WPS 自动调整字词和标点符号之间的间距，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.SmartCutPaste**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 WPS 设置为在剪切和粘贴时自动调整字词和标点符号之间的间距，然后删除和粘贴新建文档中的部分文本。如果将 SmartCutPaste 属性设置为 False，那么第二个和第三个词将一起处理。*/ function test() { Options.SmartCutPaste = true let myDoc = Documents.Add() myDoc.Content.InsertAfter("The brown quick fox") myDoc.Words.Item(2).Cut() myDoc.Characters.Item(10).Paste() }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“工具”菜单“选项”对话框中“编辑”选项卡上“智能剪贴”选项的状态。*/ let temp = Options.SmartCutPaste` |

#### **Options.SmartParaSelection**

如果为 **True**，则在选择段落中大多数或全部内容时，WPS 会在选定内容中包含段落标记。**Boolean** 类型，可读写。

**语法**

**express.SmartParaSelection**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例禁用智能段落选择功能。*/ function SetSmartParagraphSelection() {     Options.SmartParaSelection = false }` |

#### **Options.SnapToGrid**

如果绘制、移动自选图形或中文字符或者调整其大小时，它们自动与不可见的网格对齐，则该属性值为 **True**。可读写 **Boolean** 类型。

**语法**

**express.SnapToGrid**

*express*   一个代表 **Options** 对象的变量。

**说明**

绘制或移动自选图形或者调整其大小时，按下 Alt 可以暂时使该设置无效。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例设置 Word，使新文档的自选图形自动与不可见的网格对齐。*/ function test() { Options.SnapToGrid = true Documents.Add() }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例返回“对齐网格”对话框中的“对齐网格”选项的状态。*/ let Temp = Options.SnapToGrid` |

#### **Options.SnapToShapes**

如果 WPS 自动将自选图形或中文字符与不可见的网格线对齐，这些网格线穿过其他自选图形或中文字符的垂直和水平边缘，则该属性值为 **True**。可读写 **Boolean** 类型。

**语法**

**express.SnapToShapes**

*express*   一个代表 **Options** 对象的变量。

**说明**

此属性为每个自选图形创建附加的不可见网格线。**SnapToShapes** 属性与 **SnapToGrid** 属性无关。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例设置 Word，使新文档中的自选图形自动与不可见的网格线对齐，这些网格线穿过其他自选图形的垂直和水平边缘。*/ function test() { Options.SnapToShapes = true Documents.Add() }` |

#### **Options.SpanishMode**

返回或设置西班牙语拼写器的模式。可读/写 WdSpanishSpeller。

**语法**

**express.SpanishMode**

*express*   一个代表 **Options** 对象的变量。

**说明**

设置该属性与选择**“ WPS 选项”**对话框中**“西班牙语模式:”**下的选项具有相同作用（**“校对”**、**“在 WPS Office程序中更正拼写时”**）。

#### **Options.SplitCellColor**

返回或设置一个 **WdCellColor** 常量，该常量代表拆分的表格单元格的颜色。可读写。

**语法**

**express.SplitCellColor**

*express*   一个代表 **Options** 对象的变量。

**说明**

此属性仅适用于对其运行了 **CompareDocuments**、**Compare** 或 **Merge** 方法的文档。

#### **Options.StoreRSIDOnSave**

如果为 **True**，则 WPS 在每次保存文档时为文档中的变化指定一个随机数，以便比较与合并文档。**Boolean** 类型，可读写。

**语法**

**express.StoreRSIDOnSave**

*express*   一个代表 **Options** 对象的变量。

**说明**

WPS 将这些随机数保存在表格中并在每次保存后更新该表格。**StoreRSIDOnSave** 属性的默认值为 **True**。但是不保存有关 HTML 文档的 RSID 信息。

如果要删除文档中有关作者和审阅者的信息，可使用 **RemovePersonalInformation** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在保存文档时关闭保存随机数功能。*/ function SaveRandomNumber() {     Application.Options.StoreRSIDOnSave = false }` |

#### **Options.StrictFinalYaa**

如果拼写检查使用针对以字母“yaa”结尾的阿拉伯语单词的拼写规则，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.StrictFinalYaa**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置拼写检查使用针对以字符 yaa 结尾的阿拉伯语单词的拼写规则。*/ Options.StrictFinalYaa = true` |

#### **Options.StrictInitialAlefHamza**

如果拼写检查使用针对以“alef hamza”开头的阿拉伯语单词的拼写规则，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.StrictInitialAlefHamza**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置拼写检查使用针对以“alef hamza”开头的阿拉伯语单词的拼写规则。*/ Options.StrictInitialAlefHamza = true` |

#### **Options.StrictRussianE**

如果拼写检查器使用有关使用严格 ? 字符的俄语单词的拼写规则，则该属性值为 **True**。可读/写。

**语法**

**express.StrictRussianE**

*express*   一个代表 **Options** 对象的变量。

**说明**

设置该属性与选中或取消选中**“ WPS 选项”**对话框中的**“俄语: 强制严格 ?”**复选框具有相同作用（**“校对”**、**“在 WPS Office程序中更正拼写时”**）。

#### **Options.StrictTaaMarboota**

如果拼写检查器使用拼写规则来标记以 haa（而不是 taa marboota）结尾的阿拉伯语单词，则该属性值为 **True**。可读/写。

**语法**

**express.StrictTaaMarboota**

*express*   一个代表 **Options** 对象的变量。

**说明**

设置该属性与选中或清除**“ WPS 选项”**对话框中的**“阿拉伯语: 严格 taa marboota”**复选框具有相同作用（**“校对”**、**“在 WPS Office程序中更正拼写时”**）。

#### **Options.SuggestFromMainDictionaryOnly**

如果 WPS 仅根据主词典提供拼写建议，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.SuggestFromMainDictionaryOnly**

*express*   一个代表 **Options** 对象的变量。

**说明**

如果根据主词典和添加的任何自定义词典提供拼写建议，则该属性值返回 **False**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 WPS 设置为仅根据主词典选择建议单词，然后在活动文档中进行拼写检查。*/ function test() { Options.SuggestFromMainDictionaryOnly = true ActiveDocument.CheckSpelling() }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“工具”菜单“选项”对话框中“拼写和语法”选项卡上“仅根据主词典提供建议”选项的当前状态。*/ let temp = Options.SuggestFromMainDictionaryOnly` |

#### **Options.SuggestSpellingCorrections**

如果 WPS 在检查拼写时，总是为每一个拼写错误的单词提供可选的拼写建议，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.SuggestSpellingCorrections**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 WPS 设置为总是为拼写错误的单词提出可选的拼写建议，然后在活动文档中进行拼写检查。*/ function test() { Options.SuggestSpellingCorrections = true ActiveDocument.CheckSpelling() }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“工具”菜单“选项”对话框中“拼写和语法”选项卡上“总提出更正建议”选项的当前状态。 */ let temp = Options.SuggestSpellingCorrections` |

#### **Options.TabIndentKey**

如果可分别使用 Tab 和 Backspace 键增加和减少段落的左缩进，并且可使用 Backspace 键将右对齐的段落更改为居中的段落而将居中的段落更改为左对齐的段落，则该属性值为 **True**。可读写 **Boolean** 类型。

**语法**

**express.TabIndentKey**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例设置 WPS 使用 Tab 和 Backspace 键来设置段落的左缩进而不是插入和删除制表位。*/ Options.TabIndentKey = true` |

#### **Options.TypeNReplace**

如果为 **True**，则 WPS 将替换非法的南亚字符。**Boolean** 类型，可读写。

**语法**

**express.TypeNReplace**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 WPS 设置为替换非法的南亚字符。*/ function TypeReplace() {     Application.Options.TypeNReplace = true }` |

#### **Options.UpdateFieldsAtPrint**

如果 WPS 在打印文档前自动更新域，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.UpdateFieldsAtPrint**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 WPS 设置为在打印文档前自动更新域，然后打印活动文档。*/ function test() { Options.UpdateFieldsAtPrint = true ActiveDocument.PrintOut() }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“工具”菜单“选项”对话框中“打印”选项卡上“更新域”选项的当前状态。*/ let temp = Options.UpdateFieldsAtPrint` |

#### **Options.UpdateFieldsWithTrackedChangesAtPrint**

如果 WPS 2015 允许在打印前更新包含修订的字段，则该属性值为 **True**。可读/写。

**语法**

**express.UpdateFieldsWithTrackedChangesAtPrint**

*express*   一个代表 **Options** 对象的变量。

**说明**

设置该属性与在 WPS 2015 中选中或取消选中**“ WPS 选项”**对话框中**“允许在打印之前更新包含修订的字段”**（**“高级”**选项卡，**“打印”**）旁边的复选框具有相同作用。

#### **Options.UpdateLinksAtOpen**

如果 WPS 在打开文档时自动更新嵌入的所有 OLE 链接，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.UpdateLinksAtOpen**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 WPS 设置为在打开文档时自动更新嵌入的 OLE 链接。*/ Options.UpdateLinksAtOpen = true` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“选项”对话框中“常规”选项卡上“打开时更新自动方式的链接”选项的当前状态。*/ let temp = Options.UpdateLinksAtOpen` |

#### **Options.UpdateLinksAtPrint**

如果 WPS 在打印文档前更新指向其他文件的嵌入链接，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.UpdateLinksAtPrint**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 WPS 设置为在打印文档前自动更新嵌入链接，然后打印活动文档。*/ function test() { Options.UpdateLinksAtPrint = true ActiveDocument.PrintOut() }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“工具”菜单“选项”对话框中“打印”选项卡上“更新链接”选项的当前状态。*/ let temp = Options.UpdateLinksAtPrint` |

#### **Options.UpdateStyleListBehavior**

返回或设置 WdUpdateStyleListBehavior 常量，该常量指定在更新某种样式以与包含编号或项目符号的所选内容相匹配时， WPS 2015 应采取的行为。可读/写。rs">版本信息

**语法**

**express.UpdateStyleListBehavior**

*express*   一个代表 **Options** 对象的变量。

**说明**

设置该属性与选择 WPS 2015**“选项”**对话框中下拉列表中的项目具有相同作用（**“高级”**选项卡，**“编辑选项”**，**“更新样式以匹配所选内容:”**）。

#### **Options.UseCharacterUnit**

如果 WPS 在当前文档中使用字符作为默认度量单位，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.UseCharacterUnit**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 WPS 设置为使用字符作为默认度量单位。*/ Options.UseCharacterUnit = true` |

#### **Options.UseDiffDiacColor**

如果可以设置当前文档中音调符号的颜色，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.UseDiffDiacColor**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在设置当前选定内容中的音调符号的颜色前，检查 UseDiffDiacColor 属性的值。*/ function test() { if(Options.UseDiffDiacColor == true) {     Selection.Font.DiacriticColor = wdColorBlue } }` |

#### **Options.UseGermanSpellingReform**

如果 WPS 在检查拼写时，使用德语后期修订拼写规则，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.UseGermanSpellingReform**

*express*   一个代表 **Options** 对象的变量。

**说明**

由于您选择或安装的语言支持不同（例如美国英语），该属性可能无法使用。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 WPS 设置为使用德语后期修订规则检查拼写。*/ Options.UseGermanSpellingReform = true` |

#### **Options.UseNormalStyleForList**

返回或设置 **Boolean** 值，该值表示 WPS 是否对项目符号和编号使用“正文”样式。可读写。

**语法**

**express.UseNormalStyleForList**

*express*   一个代表 **Options** 对象的变量。

**说明**

对应于**“ WPS 选项”**对话框的**“高级”**选项卡中的**“对项目符号或编号列表使用正文样式”**复选框。

#### **Options.VisualSelection**

根据从右向左式语言文档中的可视光标移动，返回或设置选择行为。**WdVisualSelection** 类型，可读写。

**语法**

**express.VisualSelection**

*express*   一个代表 **Options** 对象的变量。

**说明**

必须将 **CursorMovement** 属性设置为 **wdCursorMovementVisual** 才能使用此属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置选定行为，以使选定内容换行。*/ function test() { if(Options.CursorMovement == wdCursorMovementVisual) {    Options.VisualSelection = wdVisualSelectionContinuous } }` |

#### **Options.WarnBeforeSavingPrintingSendingMarkup**

如果为 **True**，则 WPS 在保存、打印或以电子邮件形式发送包含批注或修订的文档时，显示警告。**Boolean** 类型，可读写。

**语法**

**express.WarnBeforeSavingPrintingSendingMarkup**

*express*   一个代表 **Options** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例打印活动文档，但如果文档包含修订或批注，则允许用户停止打印。*/ function SaferPrint() {     //Save old state in variable     let blnOldState = Application.Options.WarnBeforeSavingPrintingSendingMarkup 								     //Turn on warning     Application.Options.WarnBeforeSavingPrintingSendingMarkup = true 	     //Print document							     ActiveDocument.PrintOut() 								     //Restore original warning state     Application.Options.WarnBeforeSavingPrintingSendingMarkup = blnOldState }` |

适用环境：web

适用平台：windows/linux