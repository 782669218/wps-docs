**Application** 



代表 WPS 应用程序。**Application** 对象包含可返回顶级对象的属性和方法。例如，[**ActiveDocument**](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#Application.ActiveDocument) 属性返回 [**Document**](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Document/Document%20.htm#jsObject_Document) 对象。

**说明**

整个wps应用程序的api对象结构是一个树状结构，而Application对象是树状结构的根对象，同时，Application对象也提供了对于应用程序相关的各种访问接口，例如：

- 应用程序的设置选项、环境、版本号等相关信息
- 一些常见的属性，如ActiveWindow、ActiveDocument、Name等

以下示例显示 WPS 用户名。

| 示例代码复制                   |
| ------------------------------ |
| `alert(Application.UserName);` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 添加一个文档，并且打印该文档名 */ function test() { 	Application.Documents.Add(); 	alert(Application.ActiveDocument.Name); }` |

WPS宏编辑器环境下，许多返回最常见用户界面对象（如活动文档，**ActiveDocument** 属性）的属性和方法可在没有 **Application** 对象识别符的情况下使用。例如，您可以编写 ActiveDocument.PrintOut，而不是编写 Application.ActiveDocument.PrintOut。可以在没有 **Application** 对象识别符的情况下使用的属性和方法被视为“全局”属性和方法。

**方法**

|                                                              | 名称                        | 说明                                                         |
| ------------------------------------------------------------ | --------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Activate**                | 激活指定的对象。                                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddAddress**              | 本方法在通讯簿中新增一项。每项都有一个或多个标签 ID 值。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AutomaticChange**         | 当“WPS 助手”建议进行更改时，将执行**“自动套用格式”**操作。如果未激活任何“自动套用格式”操作，该方法将导致出错。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **BuildKeyCode**            | 为指定的组合键返回唯一的代码。                               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CentimetersToPoints**     | 将度量单位由厘米转换为磅数（1 厘米=28.35 磅）。所返回的转换结果为 **Single** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ChangeFileOpenDirectory** | 设置 WPS 在其中搜索文档的文件夹。                            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CheckGrammar**            | 检查字符串的语法错误。返回一个 **Boolean** 类型的值，以表示字符串中是否包含语法错误。如果字符串未包含语法错误，则该值为 **True**。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CheckSpelling**           | 检查字符串的拼写错误。返回一个 **Boolean** 类型的值，以显示字符串是否包含拼写错误。如果字符串未包含拼写错误，则该值为 **True**。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CleanString**             | 从指定字符串中删除非打印字符（字符代码为 1-29）及 WPS 的特殊字符，或将它们替换为空格（字符代码为 32）。返回的结果为 **String** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CompareDocuments**        | 比较两个文档，并返回一个 **Document** 对象，该对象表示包含两个文档之间的差别的文档，该文档用修订进行标记。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **DDEExecute**              | 通过指定的 DDE（即动态数据交换）通道，向应用程序发送一条或一组命令。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **DDEInitiate**             | 打开通向其他应用程序的 DDE（动态数据交换）通道，并返回通道序号。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **DDEPoke**                 | 通过打开的 DDE（动态数据交换）通道向应用程序发送数据。       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **DDERequest**              | 通过打开的 DDE（动态数据交换）通道向接收应用程序查询信息，并以 **String** 类型返回该信息。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **DDETerminate**            | 关闭指定的通向另一应用程序的 DDE（动态数据交换）通道。       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **DDETerminateAll**         | 关闭所有由 WPS 打开的动态数据交换 (DDE) 通道。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **DefaultWebOptions**       | 返回 **DefaultWebOptions** 对象，该对象包含 WPS 每当将文档保存为网页或打开网页时使用的应用程序级全局属性。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **FileDialog**              | 返回一个 **FileDialog** 对象，该对象代表文件对话框的单个实例。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **FindKey**                 | 返回 **KeyBinding** 对象，该对象代表指定的组合键。只读。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **GetAddress**              | 返回默认通讯簿中的地址。                                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **GetDefaultTheme**         | 返回代表默认主题名称的 **String**，以及 WPS 用于新文档、电子邮件或网页的主题格式选项。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **GetSpellingSuggestions**  | 返回一个 **SpellingSuggestions** 集合，该集合代表指定单词的建议拼写替换单词。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **GoBack**                  | 将插入点在活动文档中最近进行过编辑操作的三个位置间移动，其作用等同于按 Shift+F5。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **GoForward**               | 将插入点在活动文档中进行编辑的最后三个位置之间向前移动。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Help**                    | 显示安装的帮助信息。                                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **HelpTool**                | 指针从箭头改为问号，表明这是关于将要单击的命令或屏幕元素的快捷帮助。如果单击文本， WPS 将显示描述当前段落及字符格式的对话框。按 Esc 可将指针恢复为原始状态。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **InchesToPoints**          | 将度量单位从英寸转换为磅（1 英寸 = 72 磅）。以 **Single** 类型返回转换结果。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **International**           | 返回当前的国家/地区和国际设置信息。**Variant** 类型，只读。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **KeysBoundTo**             | 返回一个 **KeysBoundTo** 对象，该对象代表指定项的所有组合键。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Keyboard**                | 返回或设置键盘的语言和布局设置。                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **KeyboardBidi**            | 将键盘语言设置为从右向左的语言，并将文字的输入方向设置为从右向左。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **KeyboardLatin**           | 将键盘语言设置为从左向右的语言，并将文字的输入方向设置为从左向右。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **KeyString**               | 为指定的键返回组合键字符串（例如，Ctrl+Shift+A）。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **LinesToPoints**           | 将度量单位从行转换为磅（1 行 = 12 磅）。以 **Single** 类型返回转换结果。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ListCommands**            | 新建文档，然后插入带有相关快捷键和菜单设定的 WPS 命令表。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **LoadMasterList**          | 加载书目源文件。                                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **LookupNameProperties**    | 在全局通讯簿列表中查找一个姓名，并显示**“属性”**对话框，其中包含了有关指定姓名的信息。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **MergeDocuments**          | 比较两个文档，并返回一个 **Document** 对象，该对象表示包含两个文档之间的差别的文档，该文档用修订进行标记。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **MillimetersToPoints**     | 把度量单位从毫米转换为磅值（1 毫米 = 2.85 磅）。以 **Single** 类型返回转换结果。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Move**                    | 设置任务窗口或活动文档窗口的位置。                           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **NewWindow**               | 打开一个新的窗口作为指定窗口，并显示原文档。返回一个 **Window** 对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **NextLetter**              | 您已经就仅在 Macintosh 上使用的 Visual Basic 关键字请求帮助。有关 **Application** 对象的 **NextLetter** 方法的信息，请参阅 WPS OfficeMacintosh Edition 所附带的语言参考帮助。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **OnTime**                  | 启动在指定的时间运行宏的后台计时器。                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **OrganizerCopy**           | 将指定的“自动图文集”词条、工具栏、样式或宏方案项从源文档或模板中复制到目标文档或模板上。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **OrganizerDelete**         | 从文档或模板中删除指定的样式、“自动图文集”词条、工具栏或宏方案项。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **OrganizerRename**         | 重新命名文档或模板中指定的样式、“自动图文集”词条、工具栏或宏方案项。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PicasToPoints**           | 将度量单位从十二点活字转换为磅值（1 十二点活字 = 12 磅）。以 **Single** 类型返回转换结果。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PixelsToPoints**          | 将长度值的单位由像素转换为磅。以 **Single** 类型返回转换结果。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PointsToCentimeters**     | 将度量单位由磅转换为厘米（1 厘米 = 28.35 磅）。以 **Single** 类型返回转换结果。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PointsToInches**          | 将磅转换为英寸（1 英寸 = 72 磅）。以 **Single** 类型返回转换结果。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PointsToLines**           | 将磅转换为行（1 行 = 12 磅）。以 **Single** 类型返回转换结果。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PointsToMillimeters**     | 将度量单位由磅转换为毫米（1 毫米 = 2.835 磅）。以 **Single** 类型返回转换结果。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PointsToPicas**           | 将度量单位由磅转换为十二点活字（1 十二点活字 = 12 磅）。以 **Single** 类型返回转换结果。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PointsToPixels**          | 将长度值的单位由磅转换为像素。以 **Single** 类型返回转换结果。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PrintOut**                | 打印指定文档的全部或部分内容。                               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ProductCode**             | 返回一个 **String** 类型的 WPS 全球唯一标识符 (GUID)。       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PutFocusInMailHeader**    | 如果活动窗口中的文档是电子邮件文档，则将插入点置于邮件标题的**“收件人”**行。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Quit**                    | 退出 WPS ，并可选择保存或传送打开的文档。                    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Repeat**                  | 一次或多次重复最近的编辑操作。如果成功地重复了命令，则返回 **True**。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ResetIgnoreAll**          | 在拼写检查时，取消以前的忽略词汇表。                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Resize**                  | 调整 WPS 应用程序或某一任务的窗口大小。                      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Run**                     | 运行JS函数                                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ScreenRefresh**           | 使用视频内存缓冲区的当前信息更新显示器的显示。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SetDefaultTheme**         | 为 WPS 设置用于新文档、电子邮件或网页的默认主题。            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ShowClipboard**           | 显示**“剪贴板”**任务窗格。                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ShowMe**                  | 当存在更多可用信息时，显示“WPS 助手”或“帮助”窗口。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SubstituteFont**          | 设置字体映射选项。                                           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SynonymInfo**             | 返回一个 **SynonymInfo** 对象，该对象包含来自同义词库的、与指定字词或短语的同义词、反义词或相关词语及表达相关的信息。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ToggleKeyboard**          | 在从左向右的语言和从右向左语言之间切换键盘语言设置。         |

**属性**

|                                                              | 名称                                  | 说明                                                         |
| ------------------------------------------------------------ | ------------------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ActiveDocument**                    | 返回一个 [**Document** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Document/Document%20.htm#jsObject_Document)对象，该对象代表活动文档。如果没有打开的文档，就会导致出错。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ActiveEncryptionSession**           | 返回一个 **Long** 类型的值，该值代表与活动文档相关的加密会话。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ActivePrinter**                     | 返回或设置活动打印机名称。可读/写 **String** 类型。          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ActiveProtectedViewWindow**         | 返回代表活动的受保护的视图窗口的 ProtectedViewWindow 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ActiveWindow**                      | 返回 **Window** 对象，该对象代表活动窗口（焦点所在的窗口）。如果没有打开的窗口，则会导致出错。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AddIns**                            | 返回一个 **AddIns** 集合，该集合代表所有有效加载项，而不考虑当前是否已加载它们。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AnswerWizard**                      | 返回一个 **AnswerWizard** 对象，该对象包含联机“帮助”搜索引擎所用的文件。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**                       | 返回一个 **Application** 对象，该对象代表 WPS 应用程序。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ArbitraryXMLSupportAvailable**      | 返回一个 **Boolean** 类型的值，代表 WPS 是否接受自定义的 XML 架构。值为 **True** 表明 WPS 接受自定义的 XML 架构。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Assistance**                        | 返回一个代表 WPS Office帮助查看器的 **Assistance** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Assistant**                         | 返回一个 **Assistant** 对象，该对象代表“WPS Office 助手”。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoCaptions**                      | 返回一个 **AutoCaptions** 集合。当诸如表格或图片之类的项目插入到文档中时，此集合代表自动添加的题注。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoCorrect**                       | 返回一个 **AutoCorrect** 对象，该对象包含当前“自动更正”的选项、词条和例外项。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoCorrectEmail**                  | 返回一个 **AutoCorrect** 对象，该对象代表对电子邮件消息进行的自动更正。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutomationSecurity**                | 返回或设置一个 **MsoAutomationSecurity** 常量，该常量代表当用编程方法打开文件时 WPS 所使用的安全设置。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BackgroundPrintingStatus**          | 返回后台打印队列中的打印任务数目。**Long** 类型，只读。      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BackgroundSavingStatus**            | 返回后台保存队列中的文件数。**Long** 类型，只读。            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Bibliography**                      | 返回一个 **Bibliography** 对象，该对象代表 WPS 中存储的书目参考源。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BrowseExtraFileTypes**              | 如果将该属性设置为“text/html”，则用 WPS （而不是默认的 Internet 浏览器）可以打开超链接的 HTML 文件。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Browser**                           | 返回一个 **Browser** 对象，该对象代表垂直滚动条上的**“选择浏览对象”**工具。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Build**                             | 返回 WPS 应用程序的版本号及编译序号。**String** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CapsLock**                          | 如果已打开 Caps Lock 键，则该属性值为 **True**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Caption**                           | 返回或设置出现在应用程序窗口的标题栏中的文本。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CaptionLabels**                     | 返回一个 **CaptionLabels** 集合，该集合包括全部有效题注标签。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CheckLanguage**                     | 如果 WPS 在键入时自动检测使用的语言，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **COMAddIns**                         | 返回对 **COMAddIns** 集合的引用，该集合代表 WPS 中当前加载的所有组件对象模型 (COM) 加载项?（COM 加载项：通过添加自定义命令和指定的功能来扩展 WPS Office程序的功能的补充程序。COM 加载项可在一个或多个 Office 程序中运行。COM 加载项使用文件扩展名 .dll 或 .exe。）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CommandBars**                       | 返回一个 **CommandBars** 集合，该集合代表 WPS 中的菜单栏以及所有工具栏。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**                           | 返回一个 32 位整数，该整数指出用于创建指定对象的应用程序。**Long** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CustomDictionaries**                | 返回一个 **Dictionaries** 对象，该对象代表活动自定义词典的集合。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CustomizationContext**              | 返回或设置一个 **Template** 或 **Document** 对象，该对象代表存储菜单栏、工具栏和键绑定的更改的模板或文档。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DefaultLegalBlackline**             | 如果为 **True**，则 WPS 使用**“比较并合并文档”**对话框中的**“精确比较”**选项比较和合并文档。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DefaultSaveFormat**                 | 返回或设置将在**“另存为”**对话框上的**“保存类型”**框中显示的默认格式。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DefaultTableSeparator**             | 返回或设置一个字符；在将文本转换为表格时，该字符用来将文本分隔为单元格。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Dialogs**                           | 返回一个 **Dialogs** 集合，该集合代表 WPS 中所有内置对话框。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayAlerts**                     | 返回或设置运行宏时产生的一些警告和消息的处理方式。**WdAlertLevel** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayAutoCompleteTips**           | 如果 WPS 显示有关所键入整个单词、日期或词组的建议的提示，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayDocumentInformationPanel**   | 返回或设置 **Boolean** 值，该值表示是否显示文档属性面板。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayRecentFiles**                | 如果在**“文件”**菜单中显示最近使用的文件名，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayScreenTips**                 | 如果为 **True**，则批注、脚注、尾注和超链接以提示形式显示。突出显示那些标记有批注的文本。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayScrollBars**                 | 如果为 **True**，则 WPS 在至少一个文档窗口中显示滚动条；如果为 **False**，则在任何窗口中均不显示滚动条。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Documents**                         | 返回一个 **Documents** 集合，该集合代表所有打开的文档。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DontResetInsertionPointProperties** | 返回或设置一个 **Boolean** 类型的值，该值代表 WPS 在运行其他代码后是否保持位于插入点位置的文本的格式属性。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EmailOptions**                      | 返回一个 **EmailOptions** 对象，该对象代表创作电子邮件的全局首选项。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EmailTemplate**                     | 返回或设置一个 **String** 类型的值，该值代表发送电子邮件时使用的文档模板。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EnableCancelKey**                   | 返回或设置 WPS 处理 Ctrl+Break 用户中断的方式。**WdEnableCancelKey** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FeatureInstall**                    | 返回或设置 WPS 将如何处理对所需功能尚未安装的方法和属性的调用。**MsoReatureInstall** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FileConverters**                    | 返回一个 **FileConverters** 集合，该集合代表可用于 WPS 的所有文件转换器。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FileValidation**                    | 返回或设置 WPS 在打开文件之前验证这些文件的方式。可读/写 MsoFileValidationMode。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FindKey**                           | 返回 **KeyBinding** 对象，该对象代表指定的组合键。只读。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FocusInMailHeader**                 | 如果插入点位于电子邮件标题域中（如“收件人”域），则该属性值为 **True**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FontNames**                         | 返回一个 **FontNames** 对象，该对象包含所有有效字体的名称。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HangulHanjaDictionaries**           | 返回一个 **HangulHanjaConversionDictionaries** 集合，该集合代表所有活动的自定义转换词典。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Height**                            | 返回或设置活动文档窗口的高度。**Long** 类型，可读写。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IsObjectValid**                     | 如果引用某个对象的指定变量有效，则该属性值为 **True**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IsSandboxed**                       | 如果应用程序窗口为受保护的视图窗口，则该属性值为 **True**。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **KeyBindings**                       | 返回一个 **KeyBindings** 集合，该集合代表自定义的按键指定方案，包含了键代码、键类别和命令。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LandscapeFontNames**                | 返回一个 **FontNames** 对象，该对象包括了所有有效的横向字体的名称。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Language**                          | 返回一个 **MsoLanguageID** 常量，该常量代表为 WPS 用户界面选定的语言。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Languages**                         | 返回一个 **Languages** 集合，该集合代表列在**“语言”**对话框（在**“工具”**菜单上，单击**“语言”**，再单击**“设置语言”**）中的校对语言。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LanguageSettings**                  | 返回一个 **LanguageSettings** 对象，该对象包含 WPS 中语言设置的信息。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Left**                              | 返回或设置一个 **Long** 类型的值，该值代表活动文档的水平位置，以磅为单位。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ListGalleries**                     | 返回一个 **ListGalleries** 集合，该集合代表三个列表模板库。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MacroContainer**                    | 返回 **Template** 或 **Document** 对象，该对象代表其中存储模块（该模块中包含着运行的过程）的模板或文档。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MailingLabel**                      | 返回一个 **MailingLabel** 对象，该对象代表一个邮件标签。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MailMessage**                       | 返回一个 **MailMessage** 对象，该对象代表活动的电子邮件。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MailSystem**                        | 返回主机上安装的邮件系统（或系统）。**WdMailSystem** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MAPIAvailable**                     | 如果已经安装了 MAPI，则该属性值为 **True**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MathCoprocessorAvailable**          | 如果安装了数学协处理器并且该处理器适用于 WPS ，则该属性值为 **True**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MouseAvailable**                    | 如果系统有可用的鼠标，则该属性值为 **True**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**                              | 返回指定对象的名称。**String** 类型，只读。                  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NewDocument**                       | 返回一个 NewFile 对象，该对象代表在“新建文档”任务窗格上所列的一个文档。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NormalTemplate**                    | 返回一个 **Template** 对象，该对象代表 Normal 模板。         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NumLock**                           | 此属性返回 Num Lock 键的状态。如果为 **True**，则数字键盘可用于输入数字；如果为 **False**，则该键盘用于移动插入点。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OMathAutoCorrect**                  | 返回一个 **OMathAutoCorrect** 对象，该对象代表公式的自动更正条目。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OpenAttachmentsInFullScreen**       | 返回或设置 **Boolean** 值，该值表示 WPS 是否在读取模式下打开电子邮件附件。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Options**                           | 返回一个 **Options** 对象，该对象代表 WPS 中的应用程序设置。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**                            | 返回一个 **Object** 类型的值，该值代表指定 **Application** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Path**                              | 返回指定对象的磁盘或 Web 路径。只读 **String** 类型。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PathSeparator**                     | 返回用于分隔文件夹名称的字符。在 Windows 中，该属性返回一个反斜线 (\)。只读 **String** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PickerDialog**                      | 返回 PickerDialog 对象，以提供在对话框中选择用户或数据的功能。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PortraitFontNames**                 | 返回一个 **FontNames** 对象，该对象包括所有有效纵向字体的名称。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PrintPreview**                      | 如果当前视图是打印预览，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ProtectedViewWindows**              | 返回代表所有受保护的视图窗口的 ProtectedViewWindows 集合。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RecentFiles**                       | 返回一个 **RecentFiles** 集合，该集合代表最近存取过的文件。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RestrictLinkedStyles**              | 返回或设置 **Boolean** 值，该值表示 WPS 是否允许链接样式。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ScreenUpdating**                    | 如果启用屏幕更新，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Selection**                         | 返回 **Selection** 对象，该对象代表一选定的范围或插入点。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowStartupDialog**                 | 如果为 **True**，则在启动 WPS 时显示**“任务窗格”**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowStylePreviews**                 | 返回或设置 **Boolean** 值，该值表示 WPS 是否在**“样式”**对话框中显示样式的格式预览。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowVisualBasicEditor**             | 如果显示“Visual Basic 编辑器”窗口，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowWindowsInTaskbar**              | 如果为 **True**，则在任务栏中显示打开的文档，默认值为 Single Document Interface (SDI)。如果为 **False**，则只在**“窗口”**菜单中列出打开的文档，提供 Multiple Document Interface (MDI) 的外观。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SmartArtColors**                    | 返回 SmartArtColors 对象，该对象代表应用程序中当前加载的颜色样式集。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SmartArtLayouts**                   | 返回 SmartArtLayouts 对象，该对象代表应用程序中当前加载的 SmartArt 布局集。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SmartArtQuickStyles**               | 返回 SmartArtQuickStyles 对象，该对象代表应用程序中当前加载的 SmartArt 样式集。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SmartTagRecognizers**               | 返回应用程序的一个 **SmartTagRecognizers** 集合。            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SmartTagTypes**                     | 返回一个 **SmartTagTypes** 集合，代表 WPS 中所安装的智能标记组件的智能标记类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SpecialMode**                       | 如果 WPS 处于某个特殊模式（如 CopyText 模式或 MoveText 模式）下，则该属性值为 **True**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **StartupPath**                       | 返回或设置启动文件夹的完整路径，不包括最后的分隔符。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **System**                            | 返回一个 **System** 对象，该对象可用于返回与系统相关的信息，并执行与系统相关的任务。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TaskPanes**                         | 返回一个 **TaskPanes** 集合，该集合代表 WPS 中最常执行的任务。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Tasks**                             | 返回一个 **Tasks** 集合，该集合代表所有正在运行的应用程序。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Templates**                         | 返回一个 **Templates** 集合，该集合代表所有可用的模板，包括共用模板和附加到打开文档的模板。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Top**                               | 返回或设置活动文档的垂直位置。**Long** 类型，可读写。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **UndoRecord**                        | 返回 UndoRecord 对象，该对象提供撤消堆栈的自定义入口点。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **UsableHeight**                      | 返回 WPS 文档窗口高度可设置的最大值（以磅为单位）。**Long** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **UsableWidth**                       | 返回 WPS 文档窗口可设置的最大宽度（以磅为单位）。**Long** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **UserAddress**                       | 该属性返回或设置用户的通讯地址。**String** 类型，可读写。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **UserControl**                       | 如果文档或应用程序是由用户创建或打开的，则该属性值为 **True**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **UserName**                          | 该属性返回或设置用户姓名， WPS 将其用于信封和文档的“作者”属性。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **VBE**                               | 返回一个 VBE 对象，该对象代表“Visual Basic 编辑器”。         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Version**                           | 返回 WPS 版本号。**String** 类型，只读。                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Visible**                           | 如果指定对象可见，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Width**                             | 返回或设置应用程序窗口的宽度（以磅为单位）。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Windows**                           | 返回一个 **Windows** 集合，该集合代表所有文档窗口。只读。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **WindowState**                       | 返回或设置指定文档窗口或任务窗口的状态。**WdWindowState** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **WordBasic**                         | 返回一个自动化对象 (Word.Basic)，其中包含适用于 WPS 6.0 和 WPS for Windows 中所有 WordBasic 语句和函数的方法。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **XMLNamespaces**                     | 返回一个 **XMLNamespaces** 集合，代表“架构库”中的 XML 架构。 |

**成员方法**

#### **Application.Activate**

激活指定的对象。

**语法**

**express.Activate()**

*express*   一个代表 **Application** 对象的变量。

#### **Application.AddAddress**

本方法在通讯簿中新增一项。每项都有一个或多个标签 ID 值。

**语法**

**express.AddAddress(TagID, Value)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型**     | **说明**                                                     |
| -------- | ------------- | ---------------- | ------------------------------------------------------------ |
| *TagID*  | 必选          | **String array** | 新地址项的标签 ID 值。数组中的每个元素可包含下表所列的字符串之一。只有显示名称是必需的，其余项都是可选的。标签 ID说明PR_DISPLAY_NAME**“通讯簿”**对话框中显示的姓名PR_DISPLAY_NAME_PREFIX称谓（例如，“Ms.”或“Dr.”）PR_GIVEN_NAME名PR_SURNAME姓PR_STREET_ADDRESS街道地址PR_LOCALITY市/县或地方PR_STATE_OR_PROVINCE省/市/自治区PR_POSTAL_CODE邮政编码PR_COUNTRY国家/地区PR_TITLE职务PR_COMPANY_NAME公司名称PR_DEPARTMENT_NAME公司内部门名称PR_OFFICE_LOCATION办公室位置PR_PRIMARY_TELEPHONE_NUMBER主要电话号码PR_PRIMARY_FAX_NUMBER主要传真号码PR_OFFICE_TELEPHONE_NUMBER办公室电话号码PR_OFFICE2_TELEPHONE_NUMBER第二个办公室电话号码PR_HOME_TELEPHONE_NUMBER住宅电话号码PR_CELLULAR_TELEPHONE_NUMBER移动电话号码PR_BEEPER_TELEPHONE_NUMBERBP 机号码PR_COMMENT地址项的**“注释”**选项卡包含的文字PR_EMAIL_ADDRESS电子邮件地址PR_ADDRTYPE电子邮件地址类型PR_OTHER_TELEPHONE_NUMBER其他电话号码（区别于住宅或办公室电话号码）PR_BUSINESS_FAX_NUMBER商务传真号码PR_HOME_FAX_NUMBER住宅传真号码PR_RADIO_TELEPHONE_NUMBER无线电话号码PR_INITIALS缩写PR_LOCATION地点，格式为楼号/房间号（例如，7/3007 代表 7 号楼 3007 房间）PR_CAR_TELEPHONE_NUMBER汽车电话号码 |
| *Value*  | 必选          | **String array** | 新地址项的值。每个元素对应于 TagID 数组中的一个元素。有关详细信息，请参阅以下示例。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在通讯簿中添加一项。*/ function test() {     let tagIDArray = [0, 1, 2, 3]     let valueArray = [0, 1, 2, 3]      tagIDArray[0] = "PR_DISPLAY_NAME"     tagIDArray[1] = "PR_GIVEN_NAME"     tagIDArray[2] = "PR_SURNAME"     tagIDArray[3] = "PR_COMMENT"     valueArray[0] = "Kim Buhler"     valueArray[1] = "Kim"     valueArray[2] = "Buhler"     valueArray[3] = "This is a comment"      Application.AddAddress(tagIDArray, valueArray) }` |

#### **Application.AutomaticChange**

当“WPS 助手”建议进行更改时，将执行**“自动套用格式”**操作。如果未激活任何“自动套用格式”操作，该方法将导致出错。

**语法**

**express.AutomaticChange()**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将完成“WPS 助手”建议进行的“自动套用格式”操作（如果存在一个待执行的操作）。*/ Application.AutomaticChange()` |

#### **Application.BuildKeyCode**

为指定的组合键返回唯一的代码。

**语法**

**express.BuildKeyCode(Arg1, Arg2 - Arg4)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称**      | **必选/可选** | **数据类型** | **说明**                          |
| ------------- | ------------- | ------------ | --------------------------------- |
| *Arg1*        | 必选          | **WdKey**    | 使用 **WdKey** 常量之一指定的键。 |
| *Arg2 - Arg4* | 可选          | **WdKey**    | 使用 **WdKey** 常量之一指定的键。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为 Organizer 命令指定组合键 Alt+F1。*/ Application.CustomizationContext = NormalTemplate Application.KeyBindings.Add(wdKeyCategoryCommand,"Organizer",BuildKeyCode(wdKeyAlt,wdKeyF1))  /*本示例从 Normal 模板中删除 Alt+F1 按键指定方案。*/ Application.CustomizationContext = NormalTemplate Application.FindKey(BuildKeyCode(wdKeyAlt,wdKeyF1)).Clear()` |

#### **Application.CentimetersToPoints**

将度量单位由厘米转换为磅数（1 厘米=28.35 磅）。所返回的转换结果为 **Single** 类型。

**语法**

**express.CentimetersToPoints(Centimeters)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称**      | **必选/可选** | **数据类型** | **说明**             |
| ------------- | ------------- | ------------ | -------------------- |
| *Centimeters* | 必选          | **Single**   | 要转换为磅的厘米值。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例用于为选定内容的所有段落添加居中制表位。此制表位位于距页面左边距 1.5 厘米处。*/ Application.Selection.Paragraphs.TabStops.Add(Application.CentimetersToPoints(1.5),Application.Enum.wdAlignTabCenter)` |

#### **Application.ChangeFileOpenDirectory**

设置 WPS 在其中搜索文档的文件夹。

**语法**

**express.ChangeFileOpenDirectory(Path)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                  |
| -------- | ------------- | ------------ | ----------------------------------------- |
| *Path*   | 必选          | **String**   | 文件夹路径， WPS 将在该文件夹中搜索文档。 |

**说明**

下次显示**“文件”**菜单的**“打开”**对话框时，将列出指定文件夹的内容。用户在**“打开”**对话框中更改该文件夹或当前 WPS 会话结束前， WPS 将一直在指定的文件夹中搜索文档。使用 **DefaultFilePath**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Options/Options%20.htm#Options.DefaultFilePath)属性更改每个 WPS 会话中的文档默认文件夹。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例更改文件夹，以使 WPS 在该文件夹中搜索文档，然后打开“Test.doc”文档。*/ function test() { 	Application.ChangeFileOpenDirectory("C:\\Documents") 	Application.Documents.Open("Test.doc") }` |

#### **Application.CheckGrammar**

检查字符串的语法错误。返回一个 **Boolean** 类型的值，以表示字符串中是否包含语法错误。如果字符串未包含语法错误，则该值为 **True**。

**语法**

**express.CheckGrammar(String)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                     |
| -------- | ------------- | ------------ | ---------------------------- |
| *String* | 必选          | **String**   | 要对其进行语法检查的字符串。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示选定部分的语法检查结果。*/ function test() {     let strPass = Application.CheckGrammar(Application.Selection.Text)     alert("Selection is grammatically correct: " + strPass) }` |

#### **Application.CheckSpelling**

检查字符串的拼写错误。返回一个 **Boolean** 类型的值，以显示字符串是否包含拼写错误。如果字符串未包含拼写错误，则该值为 **True**。

**语法**

**express.CheckSpelling(Word, CustomDictionary, IgnoreUppercase, MainDictionary, CustomDictionary2, CustomDictionary3, CustomDictionary4, CustomDictionary5, CustomDictionary6, CustomDictionary7)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称**            | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Word*              | 必选          | **String**   | 要对其进行拼写检查的文本。                                   |
| *CustomDictionary*  | 可选          | **Variant**  | 可以是返回 Dictionary 对象的表达式，也可以是自定义词典的文件名。 |
| *IgnoreUppercase*   | 可选          | **Variant**  | 如果忽略大小写，则该属性值为 True。如果省略该参数，则使用 **IgnoreUppercase** 属性的当前值。 |
| *MainDictionary*    | 可选          | **Variant**  | 可以是返回 **Dictionary** 对象的表达式，也可以是主词典的文件名。 |
| *CustomDictionary2* | 可选          | **Variant**  | 可以是返回 **Dictionary** 对象的表达式，也可以是主词典的文件名。 |
| *CustomDictionary3* | 可选          | **Variant**  | 可以是返回 **Dictionary** 对象的表达式，也可以是主词典的文件名。 |
| *CustomDictionary4* | 可选          | **Variant**  | 可以是返回 **Dictionary** 对象的表达式，也可以是主词典的文件名。 |
| *CustomDictionary5* | 可选          | **Variant**  | 可以是返回 **Dictionary** 对象的表达式，也可以是主词典的文件名。 |
| *CustomDictionary6* | 可选          | **Variant**  | 可以是返回 **Dictionary** 对象的表达式，也可以是主词典的文件名。 |
| *CustomDictionary7* | 可选          | **Variant**  | 可以是返回 **Dictionary** 对象的表达式，也可以是主词典的文件名。 |

#### **Application.CleanString**

从指定字符串中删除非打印字符（字符代码为 1-29）及 WPS 的特殊字符，或将它们替换为空格（字符代码为 32）。返回的结果为 **String** 类型。

**语法**

**express.CleanString(String)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**   |
| -------- | ------------- | ------------ | ---------- |
| *String* | 必选          | **String**   | 源字符串。 |

**说明**

以下字符将按此表所示进行转换。

| 字符代码          | 说明                                                         |
| ----------------- | ------------------------------------------------------------ |
| 7（蜂鸣）         | 如果前导字符代码不是 13（段落），则将其删除并转换为字符 9（制表符）。 |
| 10（换行）        | 如果前导字符代码不是 13，则转换为字符 13（段落），然后将其删除。 |
| 13（段落）        | 不改变。                                                     |
| 31（可选连字符）  | 删除。                                                       |
| 160（不间断空格） | 转换为字符 32（空格）。                                      |
| 172（可选连字符） | 删除。                                                       |
| 176（不间断空格） | 转换为字符 32（空格）。                                      |
| 182（段落标记）   | 删除。                                                       |
| 183（项目符号）   | 转换为字符 32（空格）。                                      |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例删除选定文本的非打印字符，并将结果插入新文档中。*/ function test() { 	let strClean = Application.CleanString(Application.Selection.Text) 	let docNew = Application.Documents.Add() 	docNew.Content.InsertAfter(strClean) }` |

#### **Application.CompareDocuments**

比较两个文档，并返回一个 **Document** 对象，该对象表示包含两个文档之间的差别的文档，该文档用修订进行标记。

**语法**

**express.CompareDocuments(OriginalDocument, RevisedDocument, Destination, Granularity, CompareFormatting, CompareCaseChanges, CompareWhitespace, CompareTables, CompareHeaders, CompareFootnotes, CompareTextboxes, CompareFields, CompareComments, RevisedAuthor, IgnoreAllComparisonWarnings)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称**                      | **必选/可选** | **数据类型**             | **说明**                                                     |
| ----------------------------- | ------------- | ------------------------ | ------------------------------------------------------------ |
| *OriginalDocument*            | 必选          | **Document**             | 指定原始文档的路径和文件名。                                 |
| *RevisedDocument*             | 必选          | **Document**             | 指定与原始文档进行比较的修订后文档的路径和文件名。           |
| *Destination*                 | 可选          | **WdCompareDestination** | 指定是创建新文件还是在原始文档或修订后文档中标记两个文档之间的差别。默认值为 **wdCompareDestinationNew**。 |
| *Granularity*                 | 可选          | **WdGranularity**        | 指定是按字符还是按字词进行修订。默认值为 **wdGranularityWordLevel**。 |
| *CompareFormatting*           | 可选          | **Boolean**              | 指定是否标记两个文档之间的格式差别。默认值为 **True**。      |
| *CompareCaseChanges*          | 可选          |                          | 指定是否标记两个文档之间的大小写差别。默认值为 **True**。    |
| *CompareWhitespace*           | 可选          | **Boolean**              | 指定是否标记两个文档之间的空白区域（如段落或间距）差别。默认值为 **True**。 |
| *CompareTables*               | 可选          | **Boolean**              | 指定是否比较两个文档之间的表中所包含数据的差别。默认值为 **True**。 |
| *CompareHeaders*              | 可选          | **Boolean**              | 指定是否比较两个文档之间的页眉和页脚的差别。默认值为 **True**。 |
| *CompareFootnotes*            | 可选          | **Boolean**              | 指定是否比较两个文档之间的脚注和尾注的差别。默认值为 **True**。 |
| *CompareTextboxes*            | 可选          | **Boolean**              | 指定是否比较两个文档之间的文本框中所包含数据的差别。默认值为 **True**。 |
| *CompareFields*               | 可选          | **Boolean**              | 指定是否比较两个文档之间的域的差别。默认值为 **True**。      |
| *CompareComments*             | 可选          | **Boolean**              | 指定是否比较两个文档之间的批注的差别。默认值为 **True**。    |
| *RevisedAuthor*               | 可选          | **String**               | 指定在比较两个文档时做出更改的人员的姓名。                   |
| *IgnoreAllComparisonWarnings* | 可选          | **Boolean**              | 指定在比较两个文档时是否忽视警告消息。                       |

#### **Application.DDEExecute**

通过指定的 DDE（即动态数据交换）通道，向应用程序发送一条或一组命令。

**语法**

**express.DDEExecute(Channel, Command)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型** | **说明**                                                     |
| --------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Channel* | 必选          | **Long**     | **DDEInitiate** 方法返回的通道号。                           |
| *Command* | 必选          | **String**   | 由接收应用程序（DDE 服务器）识别的一条或一组命令。如果接收应用程序不能执行指定的命令，则会出现错误。 |

**说明**

**安全性**??动态数据交换 (DDE) 是一种不安全的陈旧技术。如果可能，请使用比 DDE 更加安全的技术，如对象链接和嵌入 (OLE)。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在 ET 中创建新的工作表。创建新工作表的 XLM 宏指令是 New(1)。*/ function test() { 	let lngChannel = Application.DDEInitiate("ET","System") 	Application.DDEExecute(lngChannel,"[New(1)]") 	Application.DDETerminate(lngChannel) }` |

#### **Application.DDEInitiate**

打开通向其他应用程序的 DDE（动态数据交换）通道，并返回通道序号。

**语法**

**express.DDEInitiate(App, Topic)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *App*    | 必选          | **String**   | 应用程序名。                                                 |
| *Topic*  | 必选          | **String**   | DDE 主题名称（如某打开文档的名称），打开通道所指向的应用程序将识别该名称。 |

**说明**

**安全性**??动态数据交换 (DDE) 是一种不安全的陈旧技术。如果可能，请使用比 DDE 更加安全的技术，如对象链接和嵌入 (OLE)。

如果成功，**DDEInitiate** 方法将返回打开通道的序号。所有后续 DDE 函数使用该序号来指定通道。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例用 System 主题启动 DDE 会话，并打开ET 工作簿 Sales.xls。然后，本示例终止 DDE 通道，启动通向 Sales.xls 的通道，并在单元格 R1C1 中插入文本。*/ function test() { 	let lngChannel = Application.DDEInitiate("ET", "System") 	Application.DDEExecute(lngChannel, "[OPEN(" + "C:\Sales.xls)]") 	Application.DDETerminate(lngChannel) 	lngChannel = Application.DDEInitiate("ET", "Sales.xls") 	Application.DDEPoke(lngChannel, "R1C1", "1996 Sales") 	Application.DDETerminate(lngChannel) }` |

#### **Application.DDEPoke**

通过打开的 DDE（动态数据交换）通道向应用程序发送数据。

**语法**

**express.DDEPoke(Channel, Item, Data)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型** | **说明**                                        |
| --------- | ------------- | ------------ | ----------------------------------------------- |
| *Channel* | 必选          | **Long**     | DDEInitiate 方法返回的通道号。                  |
| *Item*    | 必选          | **String**   | DDE 主题中的项，指定的数据将发送给该项。        |
| *Data*    | 必选          | **String**   | 需要发送至接收应用程序（即 DDE 服务器）的数据。 |

**说明**

**安全性**??动态数据交换 (DDE) 是一种不安全的陈旧技术。如果可能，请使用比 DDE 更加安全的技术，如对象链接和嵌入 (OLE)。

如果使用 **DDEPoke** 方法不成功，将导致出错。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例打开 ET 工作表 Sales.xls，然后在 R1C1 单元格内插入“1996 Sales”。*/ function test() { 	let lngChannel = Application.DDEInitiate("ET","System") 	Application.DDEExecute(lngChannel,"[OPEN(C:\Sales.xls)]") 	Application.DDETerminate(lngChannel) 	lngChannel = Application.DDEInitiate("ET", "Sales.xls") 	Application.DDEPoke(lngChannel,"R1C1", "1996 Sales") 	Application.DDETerminate(lngChannel) }` |

#### **Application.DDERequest**

通过打开的 DDE（动态数据交换）通道向接收应用程序查询信息，并以 **String** 类型返回该信息。

**语法**

**express.DDERequest(Channel, Item)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型** | **说明**                       |
| --------- | ------------- | ------------ | ------------------------------ |
| *Channel* | 必选          | **Long**     | DDEInitiate 方法返回的通道号。 |
| *Item*    | 必选          | **String**   | 需要进行查询的项。             |

**说明**

**安全性**??动态数据交换 (DDE) 是一种不安全的陈旧技术。如果可能，请使用比 DDE 更加安全的技术，如对象链接和嵌入 (OLE)。

请求服务器应用程序的主题信息时，必须指定包含所请求内容的主题中的项目。例如，在ET 中，单元格是合法项目，可以通过“R1C1”格式或已命名的引用来引用单元格。

ET 和其他支持 DDE 的应用程序可识别名为“System”的主题。下表对 System 主题中的三个标准项目加以说明。请注意：使用 SysItems 项目可得到 System 主题中的其他项目列表。

| System 主题中的项目 | 效果                                  |
| ------------------- | ------------------------------------- |
| SysItems            | 返回 System 主题中的所有项目的列表。  |
| Topics              | 返回所有有效主题的列表。              |
| Formats             | 返回 WPS 支持的所有剪贴板格式的列表。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例打开 ET 工作簿 Book1.xls，然后检索单元格 R1C1 的内容。*/ function test() { 	let lngChannel = DDEInitiate("ET", "System") 	Application.DDEExecute(lngChannel,"[OPEN(C:\Documents\Book1.xls)]") 	Application.DDETerminate(lngChannel) 	lngChannel = Application.DDEInitiate("ET","Book1.xls") 	alert(Application.DDERequest(lngChannel,"R1C1")) 	Application.DDETerminateAll() }` |

#### **Application.DDETerminate**

关闭指定的通向另一应用程序的 DDE（动态数据交换）通道。

**语法**

**express.DDETerminate(Channel)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型** | **说明**                       |
| --------- | ------------- | ------------ | ------------------------------ |
| *Channel* | 必选          | **Long**     | DDEInitiate 方法返回的通道号。 |

**说明**

**安全性**??动态数据交换 (DDE) 是一种不安全的陈旧技术。如果可能，请使用比 DDE 更加安全的技术，如对象链接和嵌入 (OLE)。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在 ET 中创建一张新工作簿，然后终止 DDE 会话。*/ function test() { 	let lngChannel = DDEInitiate("ET", "System") 	Application.DDEExecute(lngChannel,"[New(1)]") 	Application.DDETerminate(lngChannel) }` |

#### **Application.DDETerminateAll**

关闭所有由 WPS 打开的动态数据交换 (DDE) 通道。

**语法**

**express.DDETerminateAll()**

*express*   一个代表 **Application** 对象的变量。

**说明**

此方法不关闭由客户端应用程序打开的到 WPS 的通道。使用此方法与对每一个打开的通道使用 **DDETerminate** 方法等效。

**安全性**??动态数据交换 (DDE) 是一种不安全的陈旧技术。如果可能，请使用比 DDE 更加安全的技术，如对象链接和嵌入 (OLE)。

如果中断打开 DDE 通道的宏，可能会无意中使一个通道处于打开状态。宏结束时打开的通道不会自行关闭，并且每一个打开的通道都会占用系统资源。因此，在调试打开一个或多个 DDE 通道的宏时，最好使用此方法关闭 DDE 通道。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例打开ET 工作簿 Book1.xls，在单元格 R2C3 中插入文本，然后保存此工作簿，再关闭所有的 DDE 通道。*/ function test() { 	let lngChannel = DDEInitiate("ET", "System") 	Application.DDEExecute(lngChannel,"[OPEN(C:\Documents\Book1.xls)]") 	Application.DDETerminate(lngChannel) 	lngChannel = Application.DDEInitiate("ET","Book1.xls") 	Application.DDEPoke(lngChannel, "R2C3", "Hello World") 	Application.DDEExecute(lngChannel,"[Save]") 	Application.DDETerminateAll() }` |

#### **Application.DefaultWebOptions**

返回 **DefaultWebOptions** 对象，该对象包含 WPS 每当将文档保存为网页或打开网页时使用的应用程序级全局属性。

**语法**

**express.DefaultWebOptions()**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例检查文档编码的默认设置是否为“Western”，然后根据检查结果设置 strDocEncoding 字符串。*/ function test() { 	let strDocEncoding  	if(Application.DefaultWebOptions.Encoding == msoEncodingWestern) { 	    strDocEncoding = "Western" 	} else { 	    strDocEncoding = "Other" 	} }` |

#### **Application.FileDialog**

返回一个 **FileDialog** 对象，该对象代表文件对话框的单个实例。

**语法**

**express.FileDialog(FileDialogType)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称**         | **必选/可选** | **数据类型**          | **说明**     |
| ---------------- | ------------- | --------------------- | ------------ |
| *FileDialogType* | 必选          | **MsoFileDialogType** | 对话框类型。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*该示例显示“另存为”对话框。*/ function ShowSaveAsDialog() {     let dlgSaveAs = Application.FileDialog(msoFileDialogSaveAs)     dlgSaveAs.Show() }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示“打开”对话框并允许用户选择打开多个文件。*/ function ShowFileDialog() {     let dlgOpen = Application.FileDialog(msoFileDialogOpen)     dlgOpen.AllowMultiSelect = true     dlgOpen.Show() }` |

#### **Application.FindKey**

返回 **KeyBinding** 对象，该对象代表指定的组合键。只读。

**语法**

**express.FindKey(KeyCode, KeyCode2)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称**   | **必选/可选** | **数据类型** | **说明**                          |
| ---------- | ------------- | ------------ | --------------------------------- |
| *KeyCode*  | 必选          | **Long**     | 用 WdKey 常量之一指定的一个键。   |
| *KeyCode2* | 可选          | **Variant**  | 用 WdKey 常量之一指定的第二个键。 |

**说明**

可用 **BuildKeyCode** 方法创建 *KeyCode* 或 *KeyCode2* 参数。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例取消活动文档模板中的 Alt+Shift+F12。要返回包含不止两个键的 KeyBinding 对象，请用 BuildKeyCode 方法，如本例所示。*/ function test() {     let CustomizationContext = Application.ActiveDocument.AttachedTemplate     Application.FindKey(Application.BuildKeyCode(Application.Enum.wdKeyAlt,Application.Enum.wdKeyShift,Application.Enum.wdKeyF12).Disable() }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示 F1 指向的命令。*/ function test() {     CustomizationContext = NormalTemplate     alert((Application.FindKey(Application.Enum.wdKeyF1).Command)) }` |

#### **Application.GetAddress**

返回默认通讯簿中的地址。

**语法**

**express.GetAddress(Name, AddressProperties, UseAutoText, DisplaySelectDialog, SelectDialog, CheckNamesDialog, RecentAddressesChoice, UpdateRecentAddresses)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称**                | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Name*                  | 可选          | **Variant**  | 收件人姓名，与通讯簿中“查找姓名”对话框中所指定的相同。       |
| *AddressProperties*     | 可选          | **Variant**  | 如果 UseAutoText 为 True，则该参数代表定义通讯簿属性系列的“自动图文集”词条的名称。 如果 UseAutoText 为 False 或被省略，则该参数定义自定义布局。有效通讯簿属性名称或属性名称集合用尖括号（“<”和“>”）括起来并用空格或段落标记隔开（例如，“ ”& vbCr &“”）。 如果省略 AddressProperties 参数，则使用名为“AddressLayout”的默认“自动图文集”词条。如果还没有定义“AddressLayout”，则使用下面的地址布局定义：" " & vbCr & "" & vbCr & "" & ", " & "" & " " & "" & vbCr & ""。 要获取有效的通讯簿属性名列表，请参阅 AddAddress 方法。 |
| *UseAutoText*           | 可选          | **Variant**  | 如果 AddressProperties 指定了定义通讯簿属性系列的“自动图文集”词条的名称，则为 **True**；如果它指定了一个自定义布局，则为 False。 |
| *DisplaySelectDialog*   | 可选          | **Variant**  | 指定是否显示“选择姓名”对话框，如下表所示。值说明0（零）不显示**“选择姓名”**对话框。1 或省略显示**“选择姓名”**对话框。2不显示**“选择姓名”**对话框，并且不搜索指定的姓名。该方法返回的地址将是以前选择的地址。 |
| *SelectDialog*          | 可选          | **Variant**  | 指定“选择姓名”对话框的显示方式（即以何种模式显示），如下表所示。值显示模式0（零）或省略浏览模式1紧凑模式，只显示**“收件人:”**框2紧凑模式，**“收件人:”**和**“抄送:”**框都显示 |
| *CheckNamesDialog*      | 可选          | **Variant**  | 如果该属性值为 **True**，则在 Name 参数的值不够具体时显示“检查姓名”对话框。 |
| *RecentAddressesChoice* | 可选          | **Variant**  | 如果该属性值为 **True**，则使用最近使用的回信地址列表。      |
| *UpdateRecentAddresses* | 可选          | **Variant**  | 如果该属性值为 **True**，则向最近使用的地址列表中添加一个地址；如果该属性值为 False，则不添加地址。如果 SelectDialog 设置为 1 或 2，则忽略该参数。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 John Smith 的地址赋给变量 strAddress，将插入点移到文档的开头，并插入地址。插入的地址包括默认的通讯簿属性。*/ function test() {     let strAddress = Application.GetAddress("John Smith", null, null, null, null, true)     Application.ActiveDocument.Range(0, 0).InsertAfter(strAddress) }` |

#### **Application.GetDefaultTheme**

返回代表默认主题名称的 **String**，以及 WPS 用于新文档、电子邮件或网页的主题格式选项。

**语法**

**express.GetDefaultTheme(DocumentType)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称**       | **必选/可选** | **数据类型**         | **说明**                             |
| -------------- | ------------- | -------------------- | ------------------------------------ |
| *DocumentType* | 必选          | **WdDocumentMedium** | 要为其检索默认主题名的新文档的类型。 |

**说明**

也可使用 **ThemeName** 属性为新电子邮件返回并设置默认主题。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示 WPS 用于新网页的主题名。*/ alert(Application.GetDefaultTheme(wdWebPage))` |

#### **Application.GetSpellingSuggestions**

返回一个 **SpellingSuggestions** 集合，该集合代表指定单词的建议拼写替换单词。

**语法**

**express.GetSpellingSuggestions(Word, IgnoreUppercase, SuggestionMode)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称**          | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Word*            | 必选          | **String**   | 将接受拼写检查的单词。                                       |
| *IgnoreUppercase* | 可选          | **Variant**  | 如果为 **TrueIgnoreUppercase**                               |
| *SuggestionMode*  | 可选          | **Variant**  | 指定 WPS 提出拼写建议的方式。可以是下列 **WdSpellingWordType**wdAnagramwdSpellwordwdWildcardWdSpellword。 |

**说明**

如果此单词拼写正确，则 **SpellingSuggestions** 对象的 **Count** 属性返回 0（零）。

#### **Application.GoBack**

将插入点在活动文档中最近进行过编辑操作的三个位置间移动，其作用等同于按 Shift+F5。

**语法**

**express.GoBack()**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例打开最近使用过的文件，并将插入点移至文档中最后进行编辑操作的位置。*/ function test() {     Application.RecentFiles.Item(1).Open()     Application.GoBack() } ` |

#### **Application.GoForward**

将插入点在活动文档中进行编辑的最后三个位置之间向前移动。

**语法**

**express.GoForward()**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将插入点移至曾经进行过编辑的下一个位置。*/ Application.GoForward()` |

#### **Application.Help**

显示安装的帮助信息。

**语法**

**express.Help(HelpType)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称**   | **必选/可选** | **数据类型** | **说明**                                                     |
| ---------- | ------------- | ------------ | ------------------------------------------------------------ |
| *HelpType* | 可选          | **Variant**  | 联机帮助主题或窗口。可以是下列 **WdHelpType** 常量之一：**wdHelp**、**wdHelpAbout**、**wdHelpActiveWindow**、**wdHelpContents**、**wdHelpHWP**、**wdHelpIchitaro**、**wdHelpIndex**、**wdHelpPE2**、**wdHelpPSSHelp**、**wdHelpSearch** 和 **wdHelpUsingHelp**（由于选择或安装的语言不同，此处列出的某些常量可能无法使用）。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示“帮助主题”对话框。*/ Application.Help(Application.Enum.wdHelp)` |

#### **Application.HelpTool**

指针从箭头改为问号，表明这是关于将要单击的命令或屏幕元素的快捷帮助。如果单击文本， WPS 将显示描述当前段落及字符格式的对话框。按 Esc 可将指针恢复为原始状态。

**语法**

**express.HelpTool()**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将鼠标指针从箭头变为问号。*/ Application.HelpTool()` |

#### **Application.InchesToPoints**

将度量单位从英寸转换为磅（1 英寸 = 72 磅）。以 **Single** 类型返回转换结果。

**语法**

**express.InchesToPoints(Inches)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**               |
| -------- | ------------- | ------------ | ---------------------- |
| *Inches* | 必选          | **Single**   | 要转换为磅值的英寸值。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将选定段落的段前间距设置为 0.25 英寸。*/ Application.Selection.ParagraphFormat.SpaceBefore =Application. InchesToPoints(0.25)` |

#### **Application.International**

返回当前的国家/地区和国际设置信息。**Variant** 类型，只读。

**语法**

**express.International(Index)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型**             | **说明**                       |
| -------- | ------------- | ------------------------ | ------------------------------ |
| *Index*  | 必选          | **WdInternationalIndex** | 当前的国家/地区和/或国际设置。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在状态栏中显示货币格式。*/ StatusBar = "Currency Format: " + Application.International(Application.Enum.wdCurrencyCode)` |

#### **Application.KeysBoundTo**

返回一个 **KeysBoundTo** 对象，该对象代表指定项的所有组合键。

**语法**

**express.KeysBoundTo(KeyCategory, Command, CommandParameter)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称**           | **必选/可选** | **数据类型**      | **说明**                                                     |
| ------------------ | ------------- | ----------------- | ------------------------------------------------------------ |
| *KeyCategory*      | 必选          | **WdKeyCategory** | 组合键的类别。                                               |
| *Command*          | 必选          | **String**        | 命令的名称。                                                 |
| *CommandParameter* | 可选          | **Variant**       | 由 Command 指定的命令所需的附加文字（如果有）。有关详细信息，请参阅 KeyBindings 对象的 Add 方法的“说明”部分。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示活动文档附加的模板中为 FileOpen 命令指定的所有组合键。*/ function test() {     let kbLoop = Application.KeysBoundTo(Application.Enum.wdKeyCategoryCommand, "FileOpen")     CustomizationContext = Application.ActiveDocument.AttachedTemplate      for (let i = 1; i <= kbLoop.Count; i++) {         strOutput = strOutput + kbLoop.Item(i).KeyString + "\r"     }     alert(strOutput) }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例删除“Normal”模板中“Macro1”的所有按键指定方案。*/ function test() {     let kb = KeysBoundTo(Application.Enum.wdKeyCategoryMacro, "Macro1")     CustomizationContext = NormalTemplate     for (let i = 1; i <= kb.Count; i++) {         kb.Item(i).Disable()     } }` |

#### **Application.Keyboard**

返回或设置键盘的语言和布局设置。

**语法**

**express.Keyboard(LangId)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *LangId* | 可选          | **Long**     | {description}WPS 要针对其设置键盘的语言和布局组合。如果省略该参数，则该方法返回当前的语言和布局设置。 |

**说明**

Microsoft Windows 用一种名为“输入语言句柄”的变量类型来跟踪键盘的语言和布局设置，常把这种变量称为“HKL”。句柄的低端字是语言 ID，句柄的高端字是键盘布局。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将当前的键盘语言和布局设置赋给一个变量。*/ let lngKeyboard = Application.Keyboard()` |

#### **Application.KeyboardBidi**

将键盘语言设置为从右向左的语言，并将文字的输入方向设置为从右向左。

**语法**

**express.KeyboardBidi()**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*将键盘语言设置为从右向左的语言，并将文字的输入方向设置为从右向左。*/ Application.KeyboardBidi()` |

#### **Application.KeyboardLatin**

将键盘语言设置为从左向右的语言，并将文字的输入方向设置为从左向右。

**语法**

**express.KeyboardLatin()**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为从左向右的语言输入配置键盘。*/ Application.KeyboardLatin()` |

#### **Application.KeyString**

为指定的键返回组合键字符串（例如，Ctrl+Shift+A）。

**语法**

**express.KeyString(KeyCode, KeyCode2)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称**   | **必选/可选** | **数据类型** | **说明**                              |
| ---------- | ------------- | ------------ | ------------------------------------- |
| *KeyCode*  | 必选          | **Long**     | 用某个 **WdKey** 常量指定的键。       |
| *KeyCode2* | 可选          | **Variant**  | 用某个 **WdKey** 常量指定的另一个键。 |

**说明**

可以使用 **BuildKeyCode** 方法创建 *KeyCode* 或 *KeyCode2* 参数。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为以下三个 <b>wdKeyA</b> 常量显示组合键字符串 (Ctrl+Shift+A)：<b>WdKey</b>、<b>wdKeyControl</b> 和 <b>wdKeyShift</b>。*/ function test() {     Application.CustomizationContext = Application.ActiveDocument.AttachedTemplate     alert(KeyString(BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyA))) }` |

#### **Application.LinesToPoints**

将度量单位从行转换为磅（1 行 = 12 磅）。以 **Single** 类型返回转换结果。

**语法**

**express.LinesToPoints(Lines)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**             |
| -------- | ------------- | ------------ | -------------------- |
| *Lines*  | 必选          | **Single**   | 要转换为磅值的行值。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将选定段落的行距设置为 3 倍行距。*/ function test() {     Application.Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceMultiple     Application.Selection.ParagraphFormat.LineSpacing = LinesToPoints(3) }` |

#### **Application.ListCommands**

新建文档，然后插入带有相关快捷键和菜单设定的 WPS 命令表。

**语法**

**express.ListCommands(ListAllCommands)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称**          | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *ListAllCommands* | 必选          | **Boolean**  | 如果该属性值为 **True**，则包括所有的 WPS 命令及其设定（不论是自定义的还是内置的）。如果该属性值为 **False**，则只包括自定义设定的命令。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将新建文档，该文档将列出所有的 WPS 命令及其快捷键和菜单设定，然后打印并关闭新文档，不保存所进行的修改。*/ function test() {     Application.ListCommands(true)     Application.ActiveDocument.PrintOut()     Application.ActiveDocument.Close(wdDoNotSaveChanges) }` |

#### **Application.LoadMasterList**

加载书目源文件。

**语法**

**express.LoadMasterList(FileName)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称**   | **必选/可选** | **数据类型** | **说明**                   |
| ---------- | ------------- | ------------ | -------------------------- |
| *FileName* | 必选          | **String**   | 书目源文件的路径和文件名。 |

#### **Application.LookupNameProperties**

在全局通讯簿列表中查找一个姓名，并显示**“属性”**对话框，其中包含了有关指定姓名的信息。

**语法**

**express.LookupNameProperties(Name)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**               |
| -------- | ------------- | ------------ | ---------------------- |
| *Name*   | 必选          | **String**   | 全局通讯簿中的一个姓名 |

**说明**

如果该方法找到了多个匹配姓名，则显示**“检查姓名”**对话框。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在通讯簿中查找“Don Funk”这一姓名，并显示“Don Funk”的“属性”对话框。*/ Application.LookupNameProperties("Don Funk")` |

#### **Application.MergeDocuments**

比较两个文档，并返回一个 **Document** 对象，该对象表示包含两个文档之间的差别的文档，该文档用修订进行标记。

**语法**

**express.MergeDocuments(OriginalDocument, RevisedDocument, Destination, Granularity, CompareFormatting, CompareCaseChanges, CompareWhitespace, CompareTables, CompareHeaders, CompareFootnotes, CompareFields, CompareComments, OriginalAuthor, RevisedAuthor, FormatFrom)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称**             | **必选/可选** | **数据类型**             | **说明**                                                     |
| -------------------- | ------------- | ------------------------ | ------------------------------------------------------------ |
| *OriginalDocument*   | 必选          | **Document**             | 指定原始文档的路径和文件名。                                 |
| *RevisedDocument*    | 必选          | **Document**             | 指定与原始文档进行比较的修订后文档的路径和文件名。           |
| *Destination*        | 可选          | **WdCompareDestination** | 指定是创建新文件还是在原始文档或修订后文档中标记两个文档之间的差别。默认值为 **wdCompareDestinationNew**。 |
| *Granularity*        | 可选          | **WdGranularity**        | 指定是按字符还是按字词进行修订。默认值为 **wdGranularityWordLevel**。 |
| *CompareFormatting*  | 可选          | **Boolean**              | 指定是否标记两个文档之间的格式差别。默认值为 **True**。      |
| *CompareCaseChanges* | 可选          | **Boolean**              | 指定是否标记两个文档之间的大小写差别。默认值为 **True**。    |
| *CompareWhitespace*  | 可选          | **Boolean**              | 指定是否标记两个文档之间的空白区域（如段落或间距）差别。默认值为 **True**。 |
| *CompareTables*      | 可选          | **Boolean**              | 指定是否比较两个文档之间的表中所包含数据的差别。默认值为 **True**。 |
| *CompareHeaders*     | 可选          | **Boolean**              | 指定是否比较两个文档之间的页眉和页脚的差别。默认值为 **True**。 |
| *CompareFootnotes*   | 可选          | **Boolean**              | 指定是否比较两个文档之间的脚注和尾注的差别。默认值为 **True**。 |
| *CompareFields*      | 可选          | **Boolean**              | 指定是否比较两个文档之间的域的差别。默认值为 **True**。      |
| *CompareComments*    | 可选          | **Boolean**              | 指定是否比较两个文档之间的批注的差别。默认值为 **True**。    |
| *OriginalAuthor*     | 可选          | **String**               | 指定原始文档作者的姓名。                                     |
| *RevisedAuthor*      | 可选          | **String**               | 指定合并两个文档之后进行非属性化更改的人员的姓名。           |
| *FormatFrom*         | 可选          | **WdMergeFormatFrom**    | 指定要保留其格式的文档。                                     |

#### **Application.MillimetersToPoints**

把度量单位从毫米转换为磅值（1 毫米 = 2.85 磅）。以 **Single** 类型返回转换结果。

**语法**

**express.MillimetersToPoints(Millimeters)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称**      | **必选/可选** | **数据类型** | **说明**               |
| ------------- | ------------- | ------------ | ---------------------- |
| *Millimeters* | 必选          | **Single**   | 要转换为磅值的毫米值。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档的断字区设置为 8.8 毫米。*/ Application.ActiveDocument.HyphenationZone = Application.MillimetersToPoints(8.8)` |

#### **Application.Move**

设置任务窗口或活动文档窗口的位置。

**语法**

**express.Move(Left, Top)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                 |
| -------- | ------------- | ------------ | ------------------------ |
| *Left*   | 必选          | **Long**     | 指定窗口的水平屏幕位置。 |
| *Top*    | 必选          | **Long**     | 指定窗口的垂直屏幕位置。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例启动“计算器”应用程序 (Calc.exe)，并使用 Move 方法重新调整应用程序窗口的位置。*/ function test() { 	Application.Tasks("Calculator"); 	Application.WindowState = Application.Enum.wdWindowStateNormal 	Application.Move(50,50) }` |

#### **Application.NewWindow**

打开一个新的窗口作为指定窗口，并显示原文档。返回一个 **Window** 对象。

**语法**

**express.NewWindow()**

*express*   一个代表 **Application** 对象的变量。

**说明**

对同一文档打开多个窗口时，窗口标题将出现一个冒号 (:) 和一个数字。如果将 **NewWindow** 方法与 **Application** 对象配合使用，则为活动窗口打开一个新的窗口。下列两条指令在功能上是等效的。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `let myWindow = Application.ActiveDocument.ActiveWindow.NewWindow() myWindow = Application.NewWindow()` |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 新增一个窗口，并且打印新增前后窗口数量 */ function test() { 	alert(Application.Windows.Count); 	Application.NewWindow(); 	alert(Application.Windows.Count); }` |

#### **Application.NextLetter**

您已经就仅在 Macintosh 上使用的 Visual Basic 关键字请求帮助。有关 **Application** 对象的 **NextLetter** 方法的信息，请参阅 WPS OfficeMacintosh Edition 所附带的语言参考帮助。

**语法**

**express.NextLetter()**

*express*   一个代表 **Application** 对象的变量。

#### **Application.OnTime**

启动在指定的时间运行宏的后台计时器。

**语法**

**express.OnTime(When, Name, Tolerance)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *When*      | 必选          | **Variant**  | 运行宏的时间。                                               |
| *Name*      | 必选          | **String**   | 要运行的宏的名称。                                           |
| *Tolerance* | 可选          | **Variant**  | 表示未在 When 指定的时间运行的宏最多经过多长时间（单位为秒）被取消。宏不一定总在指定时间运行。例如，如果正在进行排序操作，或者正在显示某一对话框，则宏将延迟到 WPS 完成当前任务后才运行。如果该参数值为 0（零）或被省略，则不管距 When 指定的时间逾期多久，宏都将运行。 |

**说明**

*When* 参数可以是指定时间的字符串（如 `"4:30 pm"` 或 `"16:30"`），也可以是诸如 **TimeValue** 或 **TimeSerial** 等函数返回的一系列数字（如 `TimeValue("2:30 pm")` 或 `TimeSerial(14, 30, 00)`）。也可以包括日期（如 `"6/30 4:15 pm"` 或 `TimeValue("6/30 4:15 pm")`）。

对于 *Name* 参数，应使用完整的宏路径来确保运行正确的宏（如 `"Project.Module1.Macro1"`）。对于要运行的宏，文档或模板必须同时满足以下条件：既在 **OnTime** 指令运行时可用，又在到达 *When* 指定的时间时可用。因此，最好将宏存储在 Normal.dot 或其他能够自动加载的共用模板上。

使用 **Now** 函数以及 **TimeValue** 或 **TimeSerial** 函数的返回值的和可以设置计时器，使宏在语句运行后的指定时间运行。例如，使用 `Now+TimeValue("00:05:30")` 可以使宏在语句运行后再过 5 分 30 秒运行。

WPS 只能保留一个由 **OnTime** 设置的后台计时器。如果在现有计时器运行前启动另一个计时器，则将取消现有计时器。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在 3:55 P.M 运行当前模块中名为“Macro1”的宏。*/ Application.OnTime("15:55:00","Macro1")` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在运行 15 秒后运行名为“Macro1”的宏。*/ function test() { let now = new Date() let now_tmp = now.getTime() //获取15秒后的时间 now.setTime(now_tmp + 1000 * 15) Application.OnTime(now.toLocaleString(),"Macro1") }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在 1:30 P.M 运行名为“Start”的宏。宏的名称包括了项目名称和模块名称。*/ Application.OnTime("1:30 pm", "Start")` |

#### **Application.OrganizerCopy**

将指定的“自动图文集”词条、工具栏、样式或宏方案项从源文档或模板中复制到目标文档或模板上。

**语法**

**express.OrganizerCopy(Source, Destination, Name, Object)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称**      | **必选/可选** | **数据类型**          | **说明**                                           |
| ------------- | ------------- | --------------------- | -------------------------------------------------- |
| *Source*      | 必选          | **String**            | 含有要复制的条目的文档或模板的名称。               |
| *Destination* | 必选          | **String**            | 要向其复制条目的文档或模板的名称。                 |
| *Name*        | 必选          | **String**            | 要复制的“自动图文集”词条、工具栏、样式或宏的名称。 |
| *Object*      | 必选          | **WdOrganizerObject** | 要复制的项目的类型。                               |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将与活动文档相关的模板中的所有“自动图文集”词条复制到 Normal 模板中。*/ function test() {     let atEntry = Application.ActiveDocument.AttachedTemplate.AutoTextEntries     for (let i = 1; i <= atEntry.Count; i++) {         Application.OrganizerCopy(Application.ActiveDocument.AttachedTemplate.FullName, NormalTemplate.FullName, atEntry.Item(i).Name, wdOrganizerObjectAutoText)     } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果活动文档中含有名为“SubText”的样式，本示例将该样式复制到 C:\Templates\Template1.dot 中。*/ function test() {     let styleLoop = Application.ActiveDocument.Styles     for (let i = 1; i <= styleLoop.Count; i++) {         if (styleLoop.Item(i) == "SubText") {             Application.OrganizerCopy(Application.ActiveDocument.Name, "C:\\Templates\\Template1.dot", "SubText", wdOrganizerObjectStyles)         }     } }` |

#### **Application.OrganizerDelete**

从文档或模板中删除指定的样式、“自动图文集”词条、工具栏或宏方案项。

**语法**

**express.OrganizerDelete(Source, Name, Object)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型**          | **说明**                                           |
| -------- | ------------- | --------------------- | -------------------------------------------------- |
| *Source* | 必选          | **String**            | 含有需要删除的条目的文档或模板的名称。             |
| *Name*   | 必选          | **String**            | 要删除的样式、“自动图文集”词条、工具栏或宏的名称。 |
| *Object* | 必选          | **WdOrganizerObject** | 要复制的项目的类型。                               |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将从 Normal 模板中删除名为“Custom 1”的工具栏。*/ function test() {     for (let i = 1; i <= Application.CommandBars.Count; i++) {         if (Application.CommandBars.Item(i).Name == "Custom 1") {             Application.OrganizerDelete(NormalTemplate.Name, "Custom 1", wdOrganizerObjectCommandBars)         }     } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例提示用户删除活动文档的相关模板中的每一个“自动图文集”词条。如果用户单击“确定”按钮，则将删除“自动图文集”词条。*/ function test() {     let intResponse     let temp = Application.ActiveDocument.AttachedTemplate     let atEntry = Application.ActiveDocument.AttachedTemplate.AutoTextEntries      for (let i = 1; i <= atEntry.Count; i++) {         intResponse = alert("Do you want to delete the " + atEntry.Item(i).Name + " AutoText entry?", jsYesNoCancel)         if (intResponse == jsResultYes) {             Application.OrganizerDelete(temp.Path + "\\" + temp.Name, atEntry.Item(i).Name, wdOrganizerObjectAutoText)         }         if (intResponse == jsResultCancel) {             break         }     } }` |

#### **Application.OrganizerRename**

重新命名文档或模板中指定的样式、“自动图文集”词条、工具栏或宏方案项。

**语法**

**express.OrganizerRename(Source, Name, NewName, Object)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型**          | **说明**                                             |
| --------- | ------------- | --------------------- | ---------------------------------------------------- |
| *Source*  | 必选          | **String**            | 含有需要重新命名的项目的文档或模板的名称。           |
| *Name*    | 必选          | **String**            | 要重命名的样式、“自动图文集”词条、工具栏或宏的名称。 |
| *NewName* | 必选          | **String**            | 项目的新名称。                                       |
| *Object*  | 必选          | **WdOrganizerObject** | 要复制的项目的类型。                                 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档中名为“SubText”的样式改为“SubText2”。*/ function test() {     let styleLoop = Application.ActiveDocument.Styles     for (let i = 1; i <= styleLoop.Count; i++) {         if (styleLoop.Item(i).NameLocal == "SubText") {             Application.OrganizerRename(Application.ActiveDocument.Name, "SubText", "SubText2", wdOrganizerObjectStyles)         }     } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将附加模板中名为“Module1”的宏改为“Macros1”。*/ function test() {     let dotTemp = Application.ActiveDocument.AttachedTemplate.Name     Application.OrganizerRename(dotTemp,"Module1","Macros1",wdOrganizerObjectProjectItems) }` |

#### **Application.PicasToPoints**

将度量单位从十二点活字转换为磅值（1 十二点活字 = 12 磅）。以 **Single** 类型返回转换结果。

**语法**

**express.PicasToPoints(Picas)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                     |
| -------- | ------------- | ------------ | ---------------------------- |
| *Picas*  | 必选          | **Single**   | 要转换为磅值的十二点活字数。 |

**返回值**

Single

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `本示例将为活动文档添加行号，并将行号和文档文本之间的距离设置为 4 个十二点活字。 function test() { ActiveDocument.PageSetup.LineNumbering.Active = true ActiveDocument.PageSetup.LineNumbering.DistanceFromText = PicasToPoints(4) }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将选定段落的首行缩进设置为 3 个十二点活字。*/ Application.Selection.ParagraphFormat.FirstLineIndent = PicasToPoints(3)` |

#### **Application.PixelsToPoints**

将长度值的单位由像素转换为磅。以 **Single** 类型返回转换结果。

**语法**

**express.PixelsToPoints(Pixels, fVertical)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Pixels*    | 必选          | **Single**   | 要转换为磅值的像素值。                                       |
| *fVertical* | 可选          | **Variant**  | 如果该属性值为 True，则转换垂直像素；如果该属性值为 False，则转换水平像素。 |

**返回值**

Single

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将显示一个以像素度量的对象在以磅为单位时的高度和宽度。*/ alert("320x240 pixels is equivalent to " + PixelsToPoints(320, false) + "x" + PixelsToPoints(240, true) + " points on this display.")` |

#### **Application.PointsToCentimeters**

将度量单位由磅转换为厘米（1 厘米 = 28.35 磅）。以 **Single** 类型返回转换结果。

**语法**

**express.PointsToCentimeters(Points)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**             |
| -------- | ------------- | ------------ | -------------------- |
| *Points* | 必选          | **Single**   | 以磅为单位的度量值。 |

**返回值**

Single

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 30 磅的度量数值转换为相应的厘米数。*/ alert(PointsToCentimeters(30) + " centimeters")` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例根据 sngData 变量的值（从 1 到 5 的值，表明结果的度量单位），将 intUnit 变量的值（以磅为单位）转换为厘米、英寸、行、毫米或十二点活字。*/ function test() { let intUnit let sngData 									 function ConvertPoints(intUnit,sngData){     switch(intUnit) { 	case 1: 	    ConvertPoints = PointsToCentimeters(sngData) 	case 2: 	    ConvertPoints = PointsToInches(sngData) 	case 3: 	    ConvertPoints = PointsToLines(sngData) 	case 4: 	    ConvertPoints = PointsToMillimeters(sngData) 	case 5: 	    ConvertPoints = PointsToPicas(sngData) 	default: 	    alert("无效的过程调用")     } } } ` |

#### **Application.PointsToInches**

将磅转换为英寸（1 英寸 = 72 磅）。以 **Single** 类型返回转换结果。

**语法**

**express.PointsToInches(Points)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**             |
| -------- | ------------- | ------------ | -------------------- |
| *Points* | 必选          | **Single**   | 以磅为单位的度量值。 |

**返回值**

Single

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档上边距转换为英寸，并在消息框中显示结果。*/ alert(PointsToInches(Applicatiion.ActiveDocument.Sections.Item(1).PageSetup.TopMargin))` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例根据 sngData 变量的值（从 1 到 5 的值，表明结果的度量单位），将 intUnit 变量的值（以磅为单位）转换为厘米、英寸、行、毫米或十二点活字。*/ function test() { let intUnit let sngData 									 function ConvertPoints(intUnit,sngData){     switch(intUnit) { 	case 1: 	    ConvertPoints = PointsToCentimeters(sngData) 	case 2: 	    ConvertPoints = PointsToInches(sngData) 	case 3: 	    ConvertPoints = PointsToLines(sngData) 	case 4: 	    ConvertPoints = PointsToMillimeters(sngData) 	case 5: 	    ConvertPoints = PointsToPicas(sngData) 	default: 	    alert("无效的过程调用")     } } }` |

#### **Application.PointsToLines**

将磅转换为行（1 行 = 12 磅）。以 **Single** 类型返回转换结果。

**语法**

**express.PointsToLines(Points)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**             |
| -------- | ------------- | ------------ | -------------------- |
| *Points* | 必选          | **Single**   | 以磅为单位的度量值。 |

**返回值**

Single

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将选定内容中第一段的行间距的值由磅转换为行。*/ alert(PointsToLines(Application.Selection.Paragraphs.Item(1).LineSpacing) + " lines")` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例根据 sngData 变量的值（从 1 到 5 的值，表明结果的度量单位），将 intUnit 变量的值（以磅为单位）转换为厘米、英寸、行、毫米或十二点活字。*/ function test() { let intUnit let sngData 									 function ConvertPoints(intUnit,sngData){     switch(intUnit) { 	case 1: 	    ConvertPoints = PointsToCentimeters(sngData) 	case 2: 	    ConvertPoints = PointsToInches(sngData) 	case 3: 	    ConvertPoints = PointsToLines(sngData) 	case 4: 	    ConvertPoints = PointsToMillimeters(sngData) 	case 5: 	    ConvertPoints = PointsToPicas(sngData) 	default: 	    alert("无效的过程调用")     } } } ` |

#### **Application.PointsToMillimeters**

将度量单位由磅转换为毫米（1 毫米 = 2.835 磅）。以 **Single** 类型返回转换结果。

**语法**

**express.PointsToMillimeters(Points)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**             |
| -------- | ------------- | ------------ | -------------------- |
| *Points* | 必选          | **Single**   | 以磅为单位的度量值。 |

**返回值**

Single

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 72 磅转换为相应的毫米数。*/ alert(PointsToMillimeters(72) + " millimeters")` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例根据 sngData 变量的值（从 1 到 5 的值，表明结果的度量单位），将 intUnit 变量的值（以磅为单位）转换为厘米、英寸、行、毫米或十二点活字。*/ function test() { let intUnit let sngData 									 function ConvertPoints(intUnit,sngData){     switch(intUnit) { 	case 1: 	    ConvertPoints = PointsToCentimeters(sngData) 	case 2: 	    ConvertPoints = PointsToInches(sngData) 	case 3: 	    ConvertPoints = PointsToLines(sngData) 	case 4: 	    ConvertPoints = PointsToMillimeters(sngData) 	case 5: 	    ConvertPoints = PointsToPicas(sngData) 	default: 	    alert("无效的过程调用")     } } }` |

#### **Application.PointsToPicas**

将度量单位由磅转换为十二点活字（1 十二点活字 = 12 磅）。以 **Single** 类型返回转换结果。

**语法**

**express.PointsToPicas(Points)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**             |
| -------- | ------------- | ------------ | -------------------- |
| *Points* | 必选          | **Single**   | 以磅为单位的度量值。 |

**返回值**

Single

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 36 磅转换为相应的十二点活字数。*/ alert(Application.PointsToPicas(36) + " picas")` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例根据 sngData 变量的值（从 1 到 5 的值，表明结果的度量单位），将 intUnit 变量的值（以磅为单位）转换为厘米、英寸、行、毫米或十二点活字。*/ function test() { let intUnit let sngData 									 function ConvertPoints(intUnit,sngData){     switch(intUnit) { 	case 1: 	    ConvertPoints = PointsToCentimeters(sngData) 	case 2: 	    ConvertPoints = PointsToInches(sngData) 	case 3: 	    ConvertPoints = PointsToLines(sngData) 	case 4: 	    ConvertPoints = PointsToMillimeters(sngData) 	case 5: 	    ConvertPoints = PointsToPicas(sngData) 	default: 	    alert("无效的过程调用")     } } }` |

#### **Application.PointsToPixels**

将长度值的单位由磅转换为像素。以 **Single** 类型返回转换结果。

**语法**

**express.PointsToPixels(Points, fVertical)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Points*    | 必选          | **Single**   | 要转换为像素值的磅值。                                       |
| *fVertical* | 可选          | **Variant**  | 如果该属性值为 True，则返回的结果是垂直像素值；如果该属性值为 False，则返回的结果是水平像素值。 |

**返回值**

Single

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例以像素为单位显示以磅为单位的对象的高度和宽度。*/ alert("180x120 points is equivalent to " + Application.PointsToPixels(180, false) + "x" + Application.PointsToPixels(120, true) + " pixels on this display.")` |

#### **Application.PrintOut**

打印指定文档的全部或部分内容。

**语法**

**express.PrintOut(Background, Append, Range, OutputFileNameOutputFileName, From, To, Item, Copies, Pages, PageType, PrintToFile, Collate, FileName, ActivePrinterMacGX, ManualDuplexPrint, PrintZoomColumn, PrintZoomRow, PrintZoomPaperWidth, PrintZoomPaperHeight)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称**                       | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------------------------ | ------------- | ------------ | ------------------------------------------------------------ |
| *Background*                   | 可选          | **Variant**  | 如果将该属性设置为 True，则 WPS 在打印文档时继续运行宏。     |
| *Append*                       | 可选          | **Variant**  | 如果将该属性设置为 True，则将指定文档附加到 OutputFileName 参数所指定的文件名。如果将该属性设置为 False，则覆盖 OutputFileName 参数所指定文件的内容。 |
| *Range*                        | 可选          | **Variant**  | 页码范围。可以是任意 WdPrintOutRange 常量。                  |
| *OutputFileNameOutputFileName* | 可选          | **Variant**  | 如果 PrintToFile 为 True，则该参数指定输出文件的路径和文件名。 |
| *From*                         | 可选          | **Variant**  | 如果将 Range 设置为 wdPrintFromTo，则该参数指定起始页码。    |
| *To*                           | 可选          | **Variant**  | 如果将 Range 设置为 wdPrintFromTo，则该参数指定结束页码。    |
| *Item*                         | 可选          | **Variant**  | 要打印的项目。可以是任意 WdPrintOutItem 常量。               |
| *Copies*                       | 可选          | **Variant**  | 要打印的份数。                                               |
| *Pages*                        | 可选          | **Variant**  | 要打印的页码和页码范围，中间用逗号分开。例如，“2, 6-10”表示打印第 2 页以及第 6 至第 10 页。 |
| *PageType*                     | 可选          | **Variant**  | 要打印的页面类型。可以是任意 WdPrintOutPages 常量。          |
| *PrintToFile*                  | 可选          | **Variant**  | 如果该属性值为 True，则将打印指令发送到文件。请确保使用 OutputFileName 指定文件名。 |
| *Collate*                      | 可选          | **Variant**  | 在打印文档的多份副本时，如果该属性值为 True，则完成打印所有页面后再打印下一份副本。 |
| *FileName*                     | 可选          | **Variant**  | 要打印的文档的路径和文件名。如果省略该参数， WPS 将打印活动文档（仅应用于 Application 对象）。 |
| *ActivePrinterMacGX*           | 可选          | **Variant**  | 该参数仅适用于 WPS OfficeMacintosh Edition。有关该参数的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。 |
| *ManualDuplexPrint*            | 可选          | **Variant**  | 如果该属性值为 True，则在无双面打印组件的打印机上打印双面文档。如果该属性值为 True，则忽略 PrintBackground 和 PrintReverse 属性。使用 PrintOddPagesInAscendingOrder 和 PrintEvenPagesInAscendingOrder 属性可在手动双面打印时控制输出。由于选择或安装的语言支持（如美国英语）不同，该参数可能不可用。 |
| *PrintZoomColumn*              | 可选          | **Variant**  | 表示 WPS 在一页纸上水平放置的页数。可以是 1、2、3 或 4 页。与 PrintZoomRow 参数一同使用可在单张纸上打印多页文档。 |
| *PrintZoomRow*                 | 可选          | **Variant**  | 表示 WPS 在一页纸上垂直放置的页数。可以是 1、2 或 4 页。与 PrintZoomColumn 参数一同使用可在单张纸上打印多页文档。 |
| *PrintZoomPaperWidth*          | 可选          | **Variant**  | WPS 将打印页面缩放到的宽度，以缇为单位（20 缇 = 1 磅；72 磅 = 1 英寸）。 |
| *PrintZoomPaperHeight*         | 可选          | **Variant**  | WPS 将打印页面缩放到的高度，以缇为单位（20 缇 = 1 磅；72 磅 = 1 英寸）。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例打印活动文档的当前页面。*/ Application.ActiveDocument.PrintOut(null,null,wdPrintCurrentPage)` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例打印活动窗口中文档的前三页。*/ Application.ActiveDocument.ActiveWindow.PrintOut(null,null,wdPrintFromTo,null,"1","3")` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例打印活动文档中的备注。*/ function test() { if(Application.ActiveDocument.Comments.Count >= 1) {     Application.ActiveDocument.PrintOut(null,null,null,null,null,null,wdPrintComments) } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将打印活动文档，每张纸上打印六页文档。*/ Application.ActiveDocument.PrintOut(null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,3,2)` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例按实际尺寸的 75% 打印活动文档。*/ Application.ActiveDocument.PrintOut(null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,null,0.75 * (8.5 * 1440),0.75 * (11 * 1440))` |

#### **Application.ProductCode**

返回一个 **String** 类型的 WPS 全球唯一标识符 (GUID)。

**语法**

**express.ProductCode()**

*express*   一个代表 **Application** 对象的变量。

**返回值**

String

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将显示 WPS 的全球唯一标识符 (GUID)。*/ alert(Application.ProductCode())` |

#### **Application.PutFocusInMailHeader**

如果活动窗口中的文档是电子邮件文档，则将插入点置于邮件标题的**“收件人”**行。

**语法**

**express.PutFocusInMailHeader()**

*express*   一个代表 **Application** 对象的变量。

**说明**

为了获得最佳结果，可使用具有 **PutFocusInMailHeader** 属性的 **EnvelopeVisible** 方法。**EnvelopeVisible** 属性设置为 **True** 时，**PutFocusInMailHeader** 方法会将插入点置于邮件标题中。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示活动文档的邮件标题，然后将插入点置于邮件标题的“收件人”行。*/ function test() {     Application.ActiveDocument.ActiveWindow.EnvelopeVisible = true     Application.PutFocusInMailHeader() }` |

#### **Application.Quit**

退出 WPS ，并可选择保存或传送打开的文档。

**语法**

**express.Quit(SaveChanges, OriginalFormat, RouteDocument)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称**         | **必选/可选** | **数据类型** | **说明**                                                     |
| ---------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *SaveChanges*    | 可选          | **Variant**  | 指定 WPS 在关闭前是否保存更改过的文档。可以是 **WdSaveOptions** 常量之一。 |
| *OriginalFormat* | 可选          | **Variant**  | 指定 WPS 保存其原格式不是 WPS 文档格式的文档的方式。可以是 **WdOriginalFormat** 常量之一。 |
| *RouteDocument*  | 可选          | **Variant**  | 如果为 **True**，则将文档传送给下一个收件人。如果文档没有附加传送名单，则忽略该参数。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例关闭 WPS，并提示用户保存自上次保存后已更改过的每篇文档。 */ Application.Quit(wdPromptToSaveChanges)` |

#### **Application.Repeat**

一次或多次重复最近的编辑操作。如果成功地重复了命令，则返回 **True**。

**语法**

**express.Repeat(Times)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                     |
| -------- | ------------- | ------------ | ---------------------------- |
| *Times*  | 可选          | **Variant**  | 为需要重复最近的命令的次数。 |

**返回值**

Boolean

**说明**

使用这种方法等同于使用**“编辑”**菜单上的“重复”命令。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例插入文本“Hello”及两个段落（重复一次第二个键入操作）。*/ function test() {     Application.Selection.TypeText("Hello")     Application.Selection.TypeParagraph()     Application.Repeat() }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例重复三次最后的命令（如果该命令可以重复的话）。*/ function test() {     try {         if (Application.Repeat(3) == true) {             StatusBar = "Action repeated"         }     }     catch (exception) {         Application.console.log(Application.exception.message)     } }` |

#### **Application.ResetIgnoreAll**

在拼写检查时，取消以前的忽略词汇表。

**语法**

**express.ResetIgnoreAll()**

*express*   一个代表 **Application** 对象的变量。

**说明**

运行该方法后，以前被忽略的词现在同其他所有词一同受到检查。为使 **ResetIgnoreAll** 方法工作，必须首先将 **SpellingChecked** 属性设置为 **False**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例取消在上一次拼写检查中忽略的词汇列表，然后对活动文档重新进行拼写检查。*/ function test() {     Application.ResetIgnoreAll()     Application.ActiveDocument.SpellingChecked = false     Application.ActiveDocument.CheckSpelling() }` |

#### **Application.Resize**

调整 WPS 应用程序或某一任务的窗口大小。

**语法**

**express.Resize(Width, Height)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                        |
| -------- | ------------- | ------------ | ------------------------------- |
| *Width*  | 必选          | **Long**     | 窗口的宽度，以磅值为单位。      |
| *Height* | 必选          | **Long**     | Long 窗口的高度，以磅值为单位。 |

**说明**

如果该窗口被最大化或最小化，将出现错误。使用 **Width** 或 **Height** 属性为窗口单独设置宽度和高度。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 ET 应用程序窗口调整为 6 英寸宽、4 英寸高。*/ function test() {     if (Application.Tasks.Exists("ET") == true) {         Application.Tasks("ET").WindowState = wdWindowStateNormal         Application.Tasks("ET").Resize(Application.InchesToPoints(6), Application.InchesToPoints(4))     } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 WPS 应用程序窗口调整为 7 英寸宽、6 英寸高。*/ function test() {     Application.WindowState = wdWindowStateNormal     Application.Resize(InchesToPoints(7), InchesToPoints(6)) }` |

#### **Application.Run**

运行JS函数

**语法**

**express.Run(MacroName)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                 |
| ----------- | ------------- | ------------ | ------------------------ |
| *MacroName* | 必选          | **String**   | 函数的名称，可以带参数。 |

**说明**

*MacroName* 参数可以是任意模板、模块和宏名的组合。

**示例**

| 示例代码复制                                       |
| -------------------------------------------------- |
| `/*本示例运行test函数。*/ Application.Run(test())` |

#### **Application.ScreenRefresh**

使用视频内存缓冲区的当前信息更新显示器的显示。

**语法**

**express.ScreenRefresh()**

*express*   一个代表 **Application** 对象的变量。

**说明**

可以在使用 **ScreenUpdating** 属性禁用屏幕更新后使用该方法。**ScreenRefresh** 只对某一指令打开屏幕更新，然后立即关闭。直到再次使用 **ScreenUpdating** 属性打开屏幕更新后，后续指令才更新屏幕。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例关闭屏幕更新，打开文档 Test.doc，插入文本，刷新屏幕，然后关闭该文档（保存所做的更改）。*/ function test() {     let rngTemp     Application.ScreenUpdating = false     Application.Documents.Open("C:\\DOCS\\TEST.DOC")      rngTemp = Application.ActiveDocument.Range(0, 0)      rngTemp.InsertBefore("new")     Application.ScreenRefresh()     Application.ActiveDocument.Close(wdSaveChanges)     Application.ScreenUpdating = true }` |

#### **Application.SetDefaultTheme**

为 WPS 设置用于新文档、电子邮件或网页的默认主题。

**语法**

**express.SetDefaultTheme(Name, DocumentType)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称**       | **必选/可选** | **数据类型**         | **说明**                                                     |
| -------------- | ------------- | -------------------- | ------------------------------------------------------------ |
| *Name*         | 必选          | **String**           | 由要指定为默认主题的主题名称加上任何要应用的主题格式选项组成。该字符串格式为“themennn”，其中 theme 和 nnn 的定义如下：主题 :包含所需主题的数据的文件夹名称（主题数据文件夹的默认位置是 C:\Program Files\Kingsoft\WPS Office Professional\office6\document theme）。主题必须使用该文件夹名，不能使用“主题”对话框（“格式”菜单中的“主题”命令）中显示的名称。nnn :三个数字的字符串，表示要激活的主题格式选项（1 为激活，0 为停用）。每个数字分别对应“主题”对话框（“格式”菜单中的“主题”命令）中的“鲜艳颜色”、“动态图形”和“背景图像”复选框。如果省略该字符串，则 nnn 的默认值为“011”（“动态图形”和“背景图像”均被激活）。 |
| *DocumentType* | 必选          | **WdDocumentMedium** | 为其指定默认主题的新文档的类型。                             |

**说明**

如果设置了一个默认主题，则此主题将不用于启动 WPS 时自动创建的空白文档。此后所创建的任何文档都将具有该默认主题。

也可使用 **ThemeName** 属性为新电子邮件返回并设置默认主题。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例指定 WPS 在所有的新电子邮件中使用“Blueprint”主题。*/ Application.SetDefaultTheme("blueprnt", wdEmailMessage)` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例指定 WPS 在所有新网页中的“动态图形”中使用“Expedition”主题。*/ Application.SetDefaultTheme("expeditn 010", wdWebPage)` |

#### **Application.ShowClipboard**

显示**“剪贴板”**任务窗格。

**语法**

**express.ShowClipboard()**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示“剪贴板”任务窗格。*/ Application.ShowClipboard()` |

#### **Application.ShowMe**

当存在更多可用信息时，显示“WPS 助手”或“帮助”窗口。

**语法**

**express.ShowMe()**

*express*   一个代表 **Application** 对象的变量。

**说明**

如果没有进一步的有效信息，该方法将显示一则信息以提示不存在相关“帮助”主题。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果可能的话，本示例将完成一个“操作向导详细说明”操作。*/ Application.ShowMe()` |

#### **Application.SubstituteFont**

设置字体映射选项。

**语法**

**express.SubstituteFont(UnavailableFont, SubstituteFont)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称**          | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *UnavailableFont* | 必选          | **String**   | 在本计算机上无效的字体名称（因而您希望映射为其他字体以进行显示和打印输出）。 |
| *SubstituteFont*  | 必选          | **String**   | 用于替换无效字体并在本计算机上有效的字体的名称。             |

**说明**

可以在**“字体替换”**对话框中查找字体映射选项。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 Courier 字体替换为 CustomFont1 字体。*/ Application.SubstituteFont("CustomFont1","Courier")` |

#### **Application.SynonymInfo**

返回一个 **SynonymInfo** 对象，该对象包含来自同义词库的、与指定字词或短语的同义词、反义词或相关词语及表达相关的信息。

**语法**

**express.SynonymInfo(Word, LanguageID)**

*express*   一个代表 **Application** 对象的变量。

**参数**

| **名称**     | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------ | ------------- | ------------ | ------------------------------------------------------------ |
| *Word*       | 必选          | **String**   | 指定的字词或短语。                                           |
| *LanguageID* | 可选          | **Variant**  | 用于同义词库的语言。可以是 WdLanguageID 常量之一。（但是根据您选择或安装的语言支持不同（例如美国英语），部分列出的常量可能无法使用） |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回美国英语中单词“big”的反义词列表。*/ function test() {     let Alist = Application.SynonymInfo("big", wdEnglishUS).AntonymList     for (let i = 1; i <= Alist.length - 1; i++) {         alert(Alist[i])     } }` |

#### **Application.ToggleKeyboard**

在从左向右的语言和从右向左语言之间切换键盘语言设置。

**语法**

**express.ToggleKeyboard()**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将键盘语言设置由从右向左的语言切换为从左向右的语言。*/ function test() {     let x = alert("Switch the keyboard language setting")      Application.ToggleKeyboard() }` |

**成员属性**

#### **Application.ActiveDocument**

返回一个 [**Document** ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Document/Document%20.htm#jsObject_Document)对象，该对象代表活动文档。如果没有打开的文档，就会导致出错。只读。

**语法**

**express.ActiveDocument**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例打印当前激活的文档名*/ // 当前有打开文档 function test() { 	if(Application.Documents.Count >= 1) {  		alert(ActiveDocument.Name) 	}else { 		alert("No documents are open") 	} }  /*本示例将选定内容折叠为插入点，然后创建一个区域，包括选定内容的后面五个字符*/ function test() { 	Application.Selection.Collapse(Application.Enum.wdCollapseStart) 	let rng = Application.ActiveDocument.Range(Application.Selection.Start, Application.Selection.Start + 5) } ` |

#### **Application.ActiveEncryptionSession**

返回一个 **Long** 类型的值，该值代表与活动文档相关的加密会话。只读。

**语法**

**express.ActiveEncryptionSession**

*express*   一个代表 **Application** 对象的变量。

**说明**

加密提供程序机制在单独的进程中管理每个文件，因此每个文件均与单独的加密会话相关。

| 注释                              |
| --------------------------------- |
| 此属性仅用于文档执行自定义加密时. |

#### **Application.ActivePrinter**

返回或设置活动打印机名称。可读/写 **String** 类型。

**语法**

**express.ActivePrinter**

*express*   一个代表 **Application** 对象的变量。

**说明**

使用 **ActivePrinter** 属性设置打印机会更改默认打印机。有关详细信息，请参阅 设置 ActivePrinter 会更改系统默认打印机。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示当前打印机的名称。*/ alert( "The name of the active printer is " + ActivePrinter)` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将网络 HP LaserJet IIISi 打印机设置为活动打印机。*/ Application.ActivePrinter = "HP LaserJet IIISi on \\printers\laser"` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 LPT1 端口的本地 HP LaserJet 4 打印机设置为活动打印机。*/ Application.ActivePrinter = "HP LaserJet 4 local on LPT1:"` |

#### **Application.ActiveProtectedViewWindow**

返回代表活动的受保护的视图窗口的 ProtectedViewWindow 对象。只读。

**语法**

**express.ActiveProtectedViewWindow**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码示例显示活动的受保护的视图窗口的标题文本。*/ alert(Application.ActiveProtectedViewWindow.Caption)` |

#### **Application.ActiveWindow**

返回 **Window** 对象，该对象代表活动窗口（焦点所在的窗口）。如果没有打开的窗口，则会导致出错。只读。

**语法**

**express.ActiveWindow**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例显示活动窗口的标题文本。  Application.ActiveWindow.Caption` |

#### **Application.AddIns**

返回一个 **AddIns** 集合，该集合代表所有有效加载项，而不考虑当前是否已加载它们。只读。

**语法**

**express.AddIns**

*express*   一个代表 **Application** 对象的变量。

**说明**

此 **AddIns** 集合包含全局模板和**“工具”**菜单的**“模板和加载项”**对话框中所列的 WPS 加载项库（即 WLL）。有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例返回加载项的总数。*/ let intAddIns = Application.AddIns.Count` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例显示 Addins 集合中每一加载项的名称。*/ function test() {     for (i = 1; i <= Application.AddIns.Count; i++) {         alert(Application.AddIns.Item(i).Name)     } }` |

#### **Application.AnswerWizard(已弃用)**

返回一个 **AnswerWizard** 对象，该对象包含联机“帮助”搜索引擎所用的文件。

**语法**

**express.AnswerWizard**

*express*   一个代表 **Application** 对象的变量。

**说明**

“WPS Office 助手”和“应答向导”在 WPS Office 2015 版中已弃用。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例重新设置“操作向导”文件列表。*/ function test() {     function AnswerWizardReset()         Application.AnswerWizard.ResetFileList()     } }` |

#### **Application.Application**

返回一个 **Application** 对象，该对象代表 WPS 应用程序。

**语法**

**express.Application**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例使 WPS 显示滚动条、屏幕提示和状态栏。*/ function test() {     let app = Application     app.DisplayScrollBars = true     app.DisplayScreenTips = true     app.DisplayStatusBar = true }` |

#### **Application.ArbitraryXMLSupportAvailable**

返回一个 **Boolean** 类型的值，代表 WPS 是否接受自定义的 XML 架构。值为 **True** 表明 WPS 接受自定义的 XML 架构。

**语法**

**express.ArbitraryXMLSupportAvailable**

*express*   一个代表 **Application** 对象的变量。

**说明**

WPS Office 2003 使用 WPS XML 架构提供 XML 支持，但不提供对自定义 XML 架构的支持。仅在 WPS Office 2003 或更高版本提供对自定义 XML 架构的支持。使用 **ArbitraryXMLSupportAvailable** 属性可确定安装了哪种版本。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果安装的 WPS 版本不支持自定义的 XML 架构，则以下代码将显示一条消息。*/ function test() {     if(Application.ArbitraryXMLSupportAvailable == false) {         alert("Custom XML schemas are not supported in this version of WPS.")     } }` |

#### **Application.Assistance**

返回一个代表 WPS Office帮助查看器的 **Assistance** 对象。只读。

**语法**

**express.Assistance**

*express*   一个代表 **Application** 对象的变量。

**说明**

**Assistance** 对象使开发人员可以在 Office 帮助查看器中显示自定义帮助以及随 Office 一起安装的帮助。

#### **Application.Assistant**

返回一个 **Assistant** 对象，该对象代表“WPS Office 助手”。

**语法**

**express.Assistant**

*express*   一个代表 **Application** 对象的变量。

**说明**

WPS Office 助手和 AnswerWizard 在 2013 版的 WPS Office 中的作用已经弱化。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示 WPS 助手。*/ Application.Assistant.Visible = true` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示 WPS 助手，并将它移至屏幕的左上角。*/ function test() { assistant.Visible = true assistant.Move(100,100) }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示 WPS 助手和在气球中的自定义消息。*/ function test() { Assistant.Visible = true let bln = Assistant.NewBalloon bln.Mode = msoModeAutoDown bln.Text = "Hello" bln.Button = msoButtonSetNone bln.Show() }` |

#### **Application.AutoCaptions**

返回一个 **AutoCaptions** 集合。当诸如表格或图片之类的项目插入到文档中时，此集合代表自动添加的题注。只读。

**语法**

**express.AutoCaptions**

*express*   一个代表 **Application** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示在插入到文档时自动添加题注的各个项目的名称。*/ function test() {     for (let i = 1; i <= Application.AutoCaptions.Count; i++) {         if (Application.AutoCaptions.Item(i).AutoInsert == true) {             alert(Application.AutoCaptions.Item(i).name)         }     } }` |

#### **Application.AutoCorrect**

返回一个 **AutoCorrect** 对象，该对象包含当前“自动更正”的选项、词条和例外项。只读。

**语法**

**express.AutoCorrect**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例添加一个“自动更正”替换词条。运行此代码后，在文档中键入的所有“sr”都会自动替换为“Stella Richards”。*/ Application.AutoCorrect.Entries.Add("sr","Stella Richards")` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将删除指定的原有“自动更正”词条。*/ function test() {     let strInput     let aceLoop     let blnMatch     let intConfirm      blnMatch = false      strInput = alert("Enter the AutoCorrect entry to delete.")      for (let i = 1; i <= Application.AutoCorrect.Entries.Count; i++) {         aceLoop = Application.AutoCorrect.Entries.Item(i)         if (aceLoop.Name == strInput) {             blnMatch == true             intConfirm = alert("Are you sure you want to delete " + aceLoop.Name, jsYesNo)             if (intConfirm == jsResultYes) {                 aceLoop.Delete()             }         }     }      if (blnMatch != true) {         alert("There was no AutoCorrect entry: " + strInput)     } }` |

#### **Application.AutoCorrectEmail**

返回一个 **AutoCorrect** 对象，该对象代表对电子邮件消息进行的自动更正。

**语法**

**express.AutoCorrectEmail**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例添加电子邮件消息的“自动更正”词条。运行该代码后，在电子邮件消息中键入的所有“allways”、“hte”和“hwen”将被分别替换为“always”、“the”和“when”。*/ function AutoCorrectEMailAddress(){     var Email = Application.AutoCorrectEmail     Email.Entries.Add("allways","always")     Email.Entries.Add("hte","the")     Email.Entries.Add("hwen","when") }` |

#### **Application.AutomationSecurity**

返回或设置一个 **MsoAutomationSecurity** 常量，该常量代表当用编程方法打开文件时 WPS 所使用的安全设置。

**语法**

**express.AutomationSecurity**

*express*   一个代表 **Application** 对象的变量。

**说明**

**AutomationSecurity** 属性的默认设置为 **msoAutomationSecurityLow**。因此，若要避免更改用户的安全设置或中断依赖于默认设置的解决方案，在程序打开文件后，应谨慎地将此属性设置回其初始设置。

将 **ScreenUpdating** 设置为 **False** 不会影响警告提醒和安全警告。**DisplayAlerts**设置不会应用于安全警告。例如，如果用户将 **DisplayAlerts** 设置为等于 **False**，将 **AutomationSecurity** 设置为 **msoAutomationSecurityByUI**，同时用户处于“中等”安全级别，则在运行宏时会显示安全警告。这使宏可以捕获文件打开错误，而即使文件成功打开，仍然会显示安全警告。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例更改设置以禁用宏，显示“打开”对话框，然后将 AutomationSecurity 属性设置回其初始设置。*/ function Security(){     let lngAutomation = Application.AutomationSecurity     Application.AutomationSecurity = msoAutomationSecurityForceDisable     let dialog = Application.FileDialog(msoFileDialogOpen)     dialog.Show()     dialog.Execute()     Application.AutomationSecurity = lngAutomation }` |

#### **Application.BackgroundPrintingStatus**

返回后台打印队列中的打印任务数目。**Long** 类型，只读。

**语法**

**express.BackgroundPrintingStatus**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回当前排队等候后台打印的 WPS 打印任务数目。*/ function test() {     let lngStatus     if (Application.Options.PrintBackground == true) {         lngStatus = Application.BackgroundPrintingStatus     } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果打印任务数大于 0（零），本示例将在状态栏中显示消息。*/ function test() {     if (Application.BackgroundPrintingStatus > 0) {         Application.StatusBar = Application.BackgroundPrintingStatus + " print jobs are queued up"     } } ` |

#### **Application.BackgroundSavingStatus**

返回后台保存队列中的文件数。**Long** 类型，只读。

**语法**

**express.BackgroundSavingStatus**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在状态栏中显示当前正在保存的文档数。*/ function test() {     Application.Options.BackgroundSave = true     Application.Documents.Add()     Application.ActiveDocument.SaveAs()     while (Application.BackgroundSavingStatus != 0) {         Application.StatusBar = "Documents remaining to save: " + Application.BackgroundSavingStatus     } }` |

#### **Application.Bibliography**

返回一个 **Bibliography** 对象，该对象代表 WPS 中存储的书目参考源。只读。

**语法**

**express.Bibliography**

*express*   一个代表 **Application** 对象的变量。

#### **Application.BrowseExtraFileTypes**

如果将该属性设置为“text/html”，则用 WPS （而不是默认的 Internet 浏览器）可以打开超链接的 HTML 文件。**String** 类型，可读写。

**语法**

**express.BrowseExtraFileTypes**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例使 WPS（而不是默认的 Internet 浏览器）打开超链接的 HTML 文件。*/ function test() { Application.BrowseExtraFileTypes = "text/html" }  ` |

#### **Application.Browser**

返回一个 **Browser** 对象，该对象代表垂直滚动条上的**“选择浏览对象”**工具。只读。

**语法**

**express.Browser**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例跳转到活动文档的下一个脚注引用标记。*/ function test() {     let Borwser = Application.Browser     Borwser.Target = wdBrowseFootnote     Borwser.Next() }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例跳转到活动文档的下一个域。将从所选内容的开始处到下一个域之间的文本设置为加粗格式。*/ function test() {     Application.Selection.ExtendMode = true     let ExtendMode = Application.Browser     ExtendMode.Target = wdBrowseField     ExtendMode.Next()     Application.Selection.Font.Bold = true     Application.Selection.ExtendMode = false     Application.Selection.Collapse(wdCollapseEnd) }` |

#### **Application.Build**

返回 WPS 应用程序的版本号及编译序号。**String** 类型，只读。

**语法**

**express.Build**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例用于显示 WPS 的版本号和编译序号。*/ alert(Application.Build+"WPS  Version")` |

#### **Application.CapsLock**

如果已打开 Caps Lock 键，则该属性值为 **True**。**Boolean** 类型，只读。

**语法**

**express.CapsLock**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例检索 Caps Lock 键的当前状态。*/ function test() {     let blnCapsLock     blnCapsLock = Application.CapsLock }` |

#### **Application.Caption**

返回或设置出现在应用程序窗口的标题栏中的文本。**String** 类型，可读写。

**语法**

**express.Caption**

*express*   一个代表 **Application** 对象的变量。

**说明**

要将应用程序窗口的题注改为默认文本，请将该属性设置为空字符串（""）。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例重新设置应用程序窗口的题注。*/ Application.Caption = ""` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例改变 WPS 应用程序窗口的题注，使其包括用户名。*/ Application.Caption = Application.UserName + "'s copy of Word"` |

#### **Application.CaptionLabels**

返回一个 **CaptionLabels** 集合，该集合包括全部有效题注标签。只读。

**语法**

**express.CaptionLabels**

*express*   一个代表 **Application** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为表格题注设置编号样式。*/ Application.CaptionLabels(wdCaptionTable).NumberStyle = Application.Enum.wdCaptionNumberStyleLowercaseRoman` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例先添加一个名为“Photo”的新题注标签，然后插入一个照片题注。*/ function test() {     Application.CaptionLabels.Add("Photo")     Application.Selection.InsertParagraphAfter()     Application.Selection.InsertCaption("Photo") }` |

#### **Application.CheckLanguage**

如果 WPS 在键入时自动检测使用的语言，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.CheckLanguage**

*express*   一个代表 **Application** 对象的变量。

**说明**

如果未将 WPS 设为可用于多语言编辑，则 **CheckLanguage** 属性总返回 **False**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将检查是否已激活自动语言检测功能。*/ function test() {     If(Application.CheckLanguage == true) {         alert("Automatic language detection is activated!")     } } ` |

#### **Application.COMAddIns**

返回对 **COMAddIns** 集合的引用，该集合代表 WPS 中当前加载的所有组件对象模型 (COM) 加载项?（COM 加载项：通过添加自定义命令和指定的功能来扩展 WPS Office程序的功能的补充程序。COM 加载项可在一个或多个 Office 程序中运行。COM 加载项使用文件扩展名 .dll 或 .exe。）。

**语法**

**express.COMAddIns**

*express*   一个代表 **Application** 对象的变量。

**说明**

COM 加载项在**“COM 加载项”**对话框中列出。可以使用**“自定义”**对话框（**“工具”**菜单），向**“工具”**菜单添加**“COM 加载项”**命令。

#### **Application.CommandBars**

返回一个 **CommandBars** 集合，该集合代表 WPS 中的菜单栏以及所有工具栏。

**语法**

**express.CommandBars**

*express*   一个代表 **Application** 对象的变量。

**说明**

在访问 **CommandBars** 集合之前，可用 **CustomizationContext** 属性设置模板或文档上下文。

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例增大所有命令栏按钮并激活工具提示。*/ function test() {     let Com = Application.CommandBars     Com.LargeButtons = true     Com.DisplayTooltips = true }  ` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在应用程序窗口底部显示“绘图”工具栏。*/ function test() {     let Com = Application.CommandBars.Item("Drawing")     Com.Visible = true     Com.Position = msoBarBottom }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将“版本”命令按钮添至“常用”工具栏。*/ function test() {     CustomizationContext = NormalTemplate     Application.CommandBars.Item("Standard").Controls.Add(msoControlButton, 2522, null, 4) }` |

#### **Application.Creator**

返回一个 32 位整数，该整数指出用于创建指定对象的应用程序。**Long** 类型，只读。

**语法**

**express.Creator**

*express*   一个代表 **Application** 对象的变量。

**说明**

如果对象是在 WPS 中创建的，则 **Creator** 属性返回十六进制数 4D535744，代表字符串“WPS”。该属性主要是为 Macintosh 机的应用设计的，在该机上每个应用程序都有一个四字符的创建者代码。例如， WPS 的创建者代码是 WPS。有关该属性的其他信息，请参考 WPS OfficeMacintosh Edition 附带的语言参考帮助。

| 注释                                  |
| ------------------------------------- |
| 值也可用常量 **wdCreatorCode** 表示。 |

#### **Application.CustomDictionaries**

返回一个 **Dictionaries** 对象，该对象代表活动自定义词典的集合。只读。

**语法**

**express.CustomDictionaries**

*express*   一个代表 **Application** 对象的变量。

**说明**

在**“自定义词典”**对话框中选中复选框的词典即为活动自定义词典。有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在该集合中添加一个新的空白自定义词典。然后在消息框中显示该自定义词典的路径和文件名。*/ function test() {     let dicHome = Application.CustomDictionaries.Add("Home.dic")     alert(Application.CustomDictionaries.Item(1).Path + Application.PathSeparator + Application.CustomDictionaries.Item(1).Name) }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例清除所有的自定义词典以便没有激活状态的自定义词典。但并非删除自定义词典文件。*/ Application.CustomDictionaries.ClearAll()` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示该集合中所有自定义词典的名称。*/ function test() {     for (let i = 1; i <= Application.CustomDictionaries.Count; i++) {         alert(Application.CustomDictionaries.Item(i).Name)     } }` |

#### **Application.CustomizationContext**

返回或设置一个 **Template** 或 **Document** 对象，该对象代表存储菜单栏、工具栏和键绑定的更改的模板或文档。可读/写。

**语法**

**express.CustomizationContext**

*express*   一个代表 **Application** 对象的变量。

**说明**

对应于**“工具”**菜单**“自定义”**对话框**“命令”**选项卡上**“保存于”**框中的值。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 Alt+Ctrl+W 添至 FileClose 命令，然后将该键盘自定义方案保存于“Normal”模板中。*/ function test() {     CustomizationContext = NormalTemplate     KeyBindings.Add(wdKeyCategoryCommand, "FileClose", BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyW)) }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例向“常用”工具栏中添加“文件版本”按钮。该命令栏自定义方案保存于活动文档所用的模板中。*/ function test() {     CustomizationContext = Application.ActiveDocument.AttachedTemplate     Application.CommandBars("Standard").Controls.Add(msoControlButton, 2522, null, 8) }` |

#### **Application.DefaultLegalBlackline**

如果为 **True**，则 WPS 使用**“比较并合并文档”**对话框中的**“精确比较”**选项比较和合并文档。**Boolean** 类型，可读写。

**语法**

**express.DefaultLegalBlackline**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例启用 WPS 的“精确比较”选项比较和合并法律文档。*/ function CreateLegalBlackline() {     Application.DefaultLegalBlackline = true }` |

#### **Application.DefaultSaveFormat**

返回或设置将在**“另存为”**对话框上的**“保存类型”**框中显示的默认格式。**String** 类型，可读写。

**语法**

**express.DefaultSaveFormat**

*express*   一个代表 **Application** 对象的变量。

**说明**

本属性所用字符串是文件转换程序类名。 WPS 内部格式的类名列于下表。

| wps格式          | 文件转化程序类名 |
| ---------------- | ---------------- |
| wps文档          | ""               |
| 文档模板         | "Dot"            |
| 纯文本           | "Text"           |
| 带换行符的纯文本 | "CRText"         |
| MS-DOS 文本      | "9Text"          |
| RTF格式          | "Rtf"            |
| Unicode文本      | "Unicode"        |

使用 **FileConverter** 对象的 **ClassName** 属性可以确定外部文件转换器的类名。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将默认的保存格式设置为 WPS 文档格式。*/ Application.DefaultSaveFormat = ""` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例返回 WPS 用来保存文档的当前设置。*/ alert(Application.DefaultSaveFormat)` |

#### **Application.DefaultTableSeparator**

返回或设置一个字符；在将文本转换为表格时，该字符用来将文本分隔为单元格。**String** 类型，可读写。

**语法**

**express.DefaultTableSeparator**

*express*   一个代表 **Application** 对象的变量。

**说明**

如果省略 **ConvertToTable** 方法、**Range** 或 **Selection** 对象中的 *Separator* 参数，则使用 **DefaultTableSeparator** 属性的值。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例更改默认的表格分隔字符。*/ Application.DefaultTableSeparator = "%"` |

#### **Application.Dialogs**

返回一个 **Dialogs** 集合，该集合代表 WPS 中所有内置对话框。只读。

**语法**

**express.Dialogs**

*express*   一个代表 **Application** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 显示“打开文件”对话框 */ Application.Dialogs.Item(Application.Enumerate.wdDialogFileOpen).Show()` |

#### **Application.DisplayAlerts**

返回或设置运行宏时产生的一些警告和消息的处理方式。**WdAlertLevel** 类型，可读写。

**语法**

**express.DisplayAlerts**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 WPS 设置为在运行宏时显示所有警告和消息框。*/ Application.DisplayAlerts = wdAlertsAll` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回 DisplayAlerts 属性的当前设置。*/ let lngTemp = Application.DisplayAlerts` |

#### **Application.DisplayAutoCompleteTips**

如果 WPS 显示有关所键入整个单词、日期或词组的建议的提示，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.DisplayAutoCompleteTips**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 在您键入文字时，显示有关所键入的整个单词、日期或词组的建议的提示。*/ Application.DisplayAutoCompleteTips = true` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“工具”菜单的“自动更正”对话框中“自动图文集”选项卡的“显示有关‘自动图文集’和日期的‘记忆式键入’提示”选项的状态。*/ let blnTemp = Application.DisplayAutoCompleteTips` |

#### **Application.DisplayDocumentInformationPanel**

返回或设置 **Boolean** 值，该值表示是否显示文档属性面板。可读/写。

**语法**

**express.DisplayDocumentInformationPanel**

*express*   一个代表 **Application** 对象的变量。

#### **Application.DisplayRecentFiles**

如果在**“文件”**菜单中显示最近使用的文件名，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.DisplayRecentFiles**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例使 WPS 在“文件”菜单上显示至多 6 个文件名。*/ function test() {     Application.DisplayRecentFiles = true     Application.RecentFiles.Maximum = 6 }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例用于从“文件”菜单中删除最近使用的文件列表。*/ Application.DisplayRecentFiles = false` |

#### **Application.DisplayScreenTips**

如果为 **True**，则批注、脚注、尾注和超链接以提示形式显示。突出显示那些标记有批注的文本。**Boolean** 类型，可读写。

**语法**

**express.DisplayScreenTips**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例使 WPS 将批注、脚注和尾注显示为提示。而且突出显示那些标记有批注的文本。*/ Application.DisplayScreenTips = true` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“选项”对话框中“视图”选项卡上的“显示”区域中“屏幕提示”复选框当前的状态。*/ let temp = Application.DisplayScreenTips` |

#### **Application.DisplayScrollBars**

如果为 **True**，则 WPS 在至少一个文档窗口中显示滚动条；如果为 **False**，则在任何窗口中均不显示滚动条。**Boolean** 类型，可读写。

**语法**

**express.DisplayScrollBars**

*express*   一个代表 **Application** 对象的变量。

**说明**

将 **DisplayScrollBars** 属性设置为 **True**，可在所有窗口中显示垂直滚动条和水平滚动条。将该属性设置为 **False** 将关闭所有窗口内的滚动条。

使用 **DisplayHorizontalScrollBar** 和 **DisplayVerticalScrollBar** 属性在指定的窗口中显示单独的滚动条。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在所有窗口中显示垂直滚动条和水平滚动条。*/ Application.DisplayScrollBars = true` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果当前在任何窗口中显示有滚动条，则本示例返回 True。*/ let blnTemp = Application.DisplayScrollBars` |

#### **Application.Documents**

返回一个 **Documents** 集合，该集合代表所有打开的文档。只读。

**语法**

**express.Documents**

*express*   一个代表 **Application** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 新增一篇文档，并显示“另存为”对话框 */ Application.Documents.Add().Save()` |

#### **Application.DontResetInsertionPointProperties**

返回或设置一个 **Boolean** 类型的值，该值代表 WPS 在运行其他代码后是否保持位于插入点位置的文本的格式属性。可读写。

**语法**

**express.DontResetInsertionPointProperties**

*express*   一个代表 **Application** 对象的变量。

**说明**

有时，在运行其他 示例代码 (VBA) 代码后， WPS 会丢失插入点的格式。如果发生这样的情况，则会给依赖屏幕读取器应用程序的用户造成困难。当用户的辅助应用程序执行看起来不相关的任务时，他们会丢失格式。当运行包含 WPS 对象模型中的属性或方法的其他代码时，此属性可防止 WPS 丢失或更改位于插入点位置的文本的格式。

**要点**??除非明确需要使用此属性来纠正解决方案功能，否则不要使用此属性。

#### **Application.EmailOptions**

返回一个 **EmailOptions** 对象，该对象代表创作电子邮件的全局首选项。只读。

**语法**

**express.EmailOptions**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 标记电子邮件中的批注。*/ Application.EmailOptions.MarkComments = true` |

#### **Application.EmailTemplate**

返回或设置一个 **String** 类型的值，该值代表发送电子邮件时使用的文档模板。可读写。

**语法**

**express.EmailTemplate**

*express*   一个代表 **Application** 对象的变量。

**说明**

如果在 Microsoft Outlook 中指定 WPS 为电子邮件编辑器，可使用 **EmailTemplate** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 对所有的新电子邮件使用名为“Email”的模板。本示例假定有一个名为“Email”的模板保存于默认的模板位置。*/ function MessageTemplate() {     Application.EmailTemplate = "Email" }` |

#### **Application.EnableCancelKey**

返回或设置 WPS 处理 Ctrl+Break 用户中断的方式。**WdEnableCancelKey** 类型，可读写。

**语法**

**express.EnableCancelKey**

*express*   一个代表 **Application** 对象的变量。

**说明**

使用此属性时应非常小心。如果使用 **wdCancelDisabled**，则无法中断失控的循环或其他非自我中断的代码。另外，当代码停止运行时，除非明确地重新设置其值，否则 **EnableCancelKey** 属性不会重新设置为 **wdCancelInterrupt**，该属性在 WPS 会话期间内仍为 **wdCancelDisabled**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例使 Ctrl+Break 不能中断计数器循环。*/ function test() {     Application.EnableCancelKey = wdCancelDisabled     for (let intWait = 1; intWait <= 10000; intWait++) {         StatusBar = intWait     }     Application.EnableCancelKey = wdCancelInterrupt }` |

#### **Application.FeatureInstall**

返回或设置 WPS 将如何处理对所需功能尚未安装的方法和属性的调用。**MsoReatureInstall** 类型，可读写。

**语法**

**express.FeatureInstall**

*express*   一个代表 **Application** 对象的变量。

**说明**

当某个功能正在安装时，可以使用 **msoFeatureInstallOnDemandWithUI** 常量来防止用户认为应用程序没有响应。如果希望只有开发者才能安装新功能，则使用 **msoFeatureInstallNone** 常量。

如果将 **DisplayAlerts** 属性设为 **False**，则即使将 **FeatureInstall** 属性设为 **msoFeatureInstallOnDemand**，也不会提示用户安装新功能。如果将 **DisplayAlerts** 属性设置为 **True**，同时将 **FeatureInstall** 属性设置为 **msoFeatureInstallOnDemand**，则会显示安装进度指示器。

#### **Application.FileConverters**

返回一个 **FileConverters** 集合，该集合代表可用于 WPS 的所有文件转换器。只读。

**语法**

**express.FileConverters**

*express*   一个代表 **Application** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条消息，以表明 FileConverters 集合的第三个转换器能否用来保存文件。*/ function test() {     if (Application.FileConverters.Item(3).CanSave == true) {         alert(Application.FileConverters.Item(3).FormatName + " can save files")     }     else {         alert(Application.FileConverters.Item(3).FormatName + " cannot save files")     } }  ` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示最后一个文件转换器的名称。*/ function test() {     let fcTemp = Application.FileConverters.Item(Application.FileConverters.Count)     alert("The file name extensions for " + fcTemp.FormatName + " files are: " + fcTemp.Extensions) }` |

#### **Application.FileValidation**

返回或设置 WPS 在打开文件之前验证这些文件的方式。可读/写 MsoFileValidationMode。

**语法**

**express.FileValidation**

*express*   一个代表 **Application** 对象的变量。

**说明**

未通过验证的文件将在受保护的视图窗口中打开。**FileValidation** 属性仅针对每个会话。如果设置了 **FileValidation** 属性，该设置将在应用程序处于打开状态的整个会话中保持有效。

#### **Application.FindKey**

返回 **KeyBinding** 对象，该对象代表指定的组合键。只读。

**语法**

**express.FindKey**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 显示F1指向的命令 */ alert(Application.FindKey(Application.Enum.wdKeyF1).Command)` |

#### **Application.FocusInMailHeader**

如果插入点位于电子邮件标题域中（如“收件人”域），则该属性值为 **True**。**Boolean** 类型，只读。

**语法**

**express.FocusInMailHeader**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果插入点位于电子邮件标题域中，则本示例在状态栏显示一条消息。*/ function test() {     If(Application.FocusInMailHeader == true){         StatusBar = "Selection is in message header"     } }` |

#### **Application.FontNames**

返回一个 **FontNames** 对象，该对象包含所有有效字体的名称。只读。

**语法**

**express.FontNames**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例用于显示 FontNames 集合中的字体名称。*/ function test() {     let intResponse      for (let i = 1; i <= Application.FontNames.Count; i++) {         intResponse = alert(Application.FontNames.Item(i))         if (intResponse == jsResultCancel) {             break         }     } }` |

#### **Application.HangulHanjaDictionaries**

返回一个 **HangulHanjaConversionDictionaries** 集合，该集合代表所有活动的自定义转换词典。

**语法**

**express.HangulHanjaDictionaries**

*express*   一个代表 **Application** 对象的变量。

**说明**

在**“自定义词典”**对话框中，活动的自定义转换词典使用复选框进行标记（在**“工具”**菜单上，单击**“选项”**，单击**“拼写和语法”**选项卡，然后单击**“自定义词典”**按钮可显示该对话框）。

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在该集合中添加一个新的空白自定义词典。然后在消息框中显示该自定义词典的路径和文件名。*/ function test() {     let myHome = Application.HangulHanjaDictionaries.Add("Home.hhd")     alert(myHome.Path + Application.PathSeparator + myHome.Name) }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例取消所有自定义词典的激活状态，但不删除自定义词典文件。*/ Application.HangulHanjaDictionaries.ClearAll()` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示该集合中所有自定义词典的名称。*/ function test() {     for (let i = 1; i <= Application.HangulHanjaDictionaries.Count; i++) {         alert(Application.HangulHanjaDictionaries.Item(i).Name)     } } ` |

#### **Application.Height**

返回或设置活动文档窗口的高度。**Long** 类型，可读写。

**语法**

**express.Height**

*express*   一个代表 **Application** 对象的变量。

#### **Application.IsObjectValid**

如果引用某个对象的指定变量有效，则该属性值为 **True**。**Boolean** 类型，只读。

**语法**

**express.IsObjectValid**

*express*   一个代表 **Application** 对象的变量。

**说明**

如果该变量引用的对象已被删除，则 **IsObjectValid** 属性返回 **False**。

**示例**

本示例向活动文档添加一个表格，并将该表格赋给 `aTable` 变量。然后，本示例删除文档中的第一个表格。如果 aTable 引用的表格不是文档中的第一个表格（也就是说，`aTable` 仍是一个有效对象），则本示例将同时删除该表格中的所有边框。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let aTable = Application.ActiveDocument.Tables.Add(Selection.Range, 2, 3)     Application.ActiveDocument.Tables.Item(1).Delete()     if (IsObjectValid(aTable) == true) {         aTable.Borders.Enable = false     } } ` |

#### **Application.IsSandboxed**

如果应用程序窗口为受保护的视图窗口，则该属性值为 **True**。只读。

**语法**

**express.IsSandboxed**

*express*   一个代表 **Application** 对象的变量。

**说明**

使用 **IsSandboxed** 属性可确定文档是否在受保护的视图窗口内打开。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码示例显示指定的文档是否在受保护的视图窗口中打开。*/ function CheckIfSandboxed(doc) {     alert(doc.Application.IsSandboxed)	 }   ` |

#### **Application.KeyBindings**

返回一个 **KeyBindings** 集合，该集合代表自定义的按键指定方案，包含了键代码、键类别和命令。

**语法**

**express.KeyBindings**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 Ctrl+Alt+W 指定给 FileClose 命令。这个键盘自定义方案保存在“Normal”模板中。*/ function test() {     CustomizationContext = NormalTemplate     Application.KeyBindings.Add(wdKeyCategoryCommand, "FileClose", BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyW)) }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为 KeyBindings 集合中的每项插入命令名称和组合键字符串。*/ function test() {     CustomizationContext = NormalTemplate      for (let i = 1; i <= Application.KeyBindings.Count; i++) {         Application.Selection.InsertAfter(Application.KeyBindings.Item(i).Command + "\t" + KeyBindings.Item(i).KeyString + "\r")         Application.Selection.Collapse(wdCollapseEnd)     } }` |

#### **Application.LandscapeFontNames**

返回一个 **FontNames** 对象，该对象包括了所有有效的横向字体的名称。

**语法**

**express.LandscapeFontNames**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例创建在 FontNames 对象中命名的横向字体名称新文档的排序列表。*/ function ListLandscapeFonts() {     let docNew = Application.Documents.Add()     docNew.Content.InsertAfter("Landscape Fonts" + "\n")      for (let i = 1; i <= LandscapeFontNames.Count; i++) {         docNew.Content.InsertAfter(LandscapeFontNames.Item(i) + "\n")     }      docNew.Range(docNew.Paragraphs.Item(2).Range.Start, docNew.Paragraphs.Item(docNew.Paragraphs.Count).Range.End).Select()      Selection.Sort() }` |

#### **Application.Language**

返回一个 **MsoLanguageID** 常量，该常量代表为 WPS 用户界面选定的语言。

**语法**

**express.Language**

*express*   一个代表 **Application** 对象的变量。

**说明**

该属性的值与下列表达式的返回值相同：

| 示例代码复制                                               |
| ---------------------------------------------------------- |
| `Application.LanguageSettings.LanguageID(msoLanguageIDUI)` |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条消息，表明 WPS 用户界面所选语言是否是英语（美国）。*/ function LangSetting(){     if(Application.Language == msoLanguageIDEnglishUS){         alert("The user interface language is U.S. English.")     }     else{         alert("The user interface language is not U.S. English.")     } }` |

#### **Application.Languages**

返回一个 **Languages** 集合，该集合代表列在**“语言”**对话框（在**“工具”**菜单上，单击**“语言”**，再单击**“设置语言”**）中的校对语言。

**语法**

**express.Languages**

*express*   一个代表 **Application** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回活动拼写词典的完整路径和文件名。*/ function test() {     let dicSpell = Application.Languages.Item(Selection.LanguageID).ActiveSpellingDictionary     alert(dicSpell.Path + Application.PathSeparator + dicSpell.Name) }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例用 aLang() 数组来存储校对语言的名称。*/ function test() {     let aLang = []     intCount = 0      for (let i = 1; i <= Languages.Count; i++) {         aLang[intCount] = Languages.Item(i).NameLocal         intCount++     } }` |

#### **Application.LanguageSettings**

返回一个 **LanguageSettings** 对象，该对象包含 WPS 中语言设置的信息。

**语法**

**express.LanguageSettings**

*express*   一个代表 **Application** 对象的变量。

#### **Application.Left**

返回或设置一个 **Long** 类型的值，该值代表活动文档的水平位置，以磅为单位。可读写。

**语法**

**express.Left**

*express*   一个代表 **Application** 对象的变量。

#### **Application.ListGalleries**

返回一个 **ListGalleries** 集合，该集合代表三个列表模板库。

**语法**

**express.ListGalleries**

*express*   一个代表 **Application** 对象的变量。

**说明**

每个模板库（项目符号、编号和多级符号）都对应于**“项目符号和编号”**对话框（**“格式”**菜单）上的一个选项卡。有关返回集合中单个成员的信息，请参阅 从集合返回一个对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将变量 mylsttmp 设置为“项目符号和编号”对话框的“多级符号”选项卡中的第二种列表模板。然后，该示例将此模板应用到活动文档中的第一个列表。*/ function test() {     let mylsttmp = Application.ListGalleries.Item(wdOutlineNumberGallery).ListTemplates.Item(2)     Application.ActiveDocument.Lists.Item(1).ApplyListTemplate(mylsttmp) }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在 ListGalleries 集合中循环搜索，并将每个列表模板库中的模板重新设置为内置模板。*/ function test() {     for (let i = 1; i <= Application.ListGalleries.Count; i++) {         for (let j = 1; j <= 7; j++) {             Application.ListGalleries.Item(i).Reset(j)         }     } } ` |

#### **Application.MacroContainer**

返回 **Template** 或 **Document** 对象，该对象代表其中存储模块（该模块中包含着运行的过程）的模板或文档。

**语法**

**express.MacroContainer**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将显示文档或模板名称，其中包含正在运行的过程。*/ function test() {     let cntnr = Application.MacroContainer     alert(cntnr.Name) }` |

#### **Application.MailingLabel**

返回一个 **MailingLabel** 对象，该对象代表一个邮件标签。

**语法**

**express.MailingLabel**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为指定地址创建一个新的 Avery 2160 小型标签文档。*/ function test() {     let addr = "Dave Edson" + "\r" + "123 Skye St." + "\r" + "Our Town, WA  98004"     Application.MailingLabel.CreateNewDocument("2160 mini", addr, null, false) }` |

#### **Application.MailMessage**

返回一个 **MailMessage** 对象，该对象代表活动的电子邮件。

**语法**

**express.MailMessage**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为活动电子邮件显示“选择姓名”对话框。*/ Application.MailMessage.DisplaySelectNamesDialog()` |

#### **Application.MailSystem**

返回主机上安装的邮件系统（或系统）。**WdMailSystem** 类型，只读。

**语法**

**express.MailSystem**

*express*   一个代表 **Application** 对象的变量。

**说明**

某些 **WdMailSystem** 常量仅可用于 WPS OfficeMacintosh Edition。有关这些常量的其他信息，请参阅 WPS OfficeMacintosh Edition 中的语言参考的帮助文件。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条消息，指示邮件系统是否安装在此计算机上。*/ function test() {     let ms = Application.MailSystem     if (ms != wdNoMailSystem) {         alert("This computer has a mail system installed.")     }     else {         alert("This computer has no mail system installed.")     } }` |

#### **Application.MAPIAvailable**

如果已经安装了 MAPI，则该属性值为 **True**。**Boolean** 类型，只读。

**语法**

**express.MAPIAvailable**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果已安装了 MAPI，则本示例显示一条消息。*/ function test() {     if (Application.MAPIAvailable == true) {         alert("MAPI is available")     } }` |

#### **Application.MathCoprocessorAvailable**

如果安装了数学协处理器并且该处理器适用于 WPS ，则该属性值为 **True**。**Boolean** 类型，只读。

**语法**

**express.MathCoprocessorAvailable**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条消息以表明是否安装了数学协处理器并且是否适用于 WPS。*/ function test() { if(Application.MathCoprocessorAvailable == true){     alert("A math coprocessor is available.") } else{     alert("A math coprocessor is not installed.") } }` |

#### **Application.MouseAvailable**

如果系统有可用的鼠标，则该属性值为 **True**。**Boolean** 类型，只读。

**语法**

**express.MouseAvailable**

*express*   一个代表 **Application** 对象的变量。

**示例**

------

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一条消息，表明没有可用的鼠标。*/ function test() { if(Application.MouseAvailable == false){     alert("Make sure your mouse is plugged in.") } else{     alert("Mouse is available") } }` |

#### **Application.Name**

返回指定对象的名称。**String** 类型，只读。

**语法**

**express.Name**

*express*   一个代表 **Application** 对象的变量。

#### **Application.NewDocument**

返回一个 NewFile 对象，该对象代表在“新建文档”任务窗格上所列的一个文档。

**语法**

**express.NewDocument**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在“新建文档”任务窗格的“根据现有文档新建”部分中创建一个文档列表项目。*/ function CreateNewDocument() {     Application.NewDocument.Add("C:\\NewFile.doc", msoNewfromExistingFile, "New File", msoCreateNewFile) }` |

#### **Application.NormalTemplate**

返回一个 **Template** 对象，该对象代表 Normal 模板。

**语法**

**express.NormalTemplate**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果 AutoTextEntries 集合中包含一个名为“Test”的“自动图文集”词条，则以下示例从 Normal 模板插入该词条。*/ function test() {     for (let i = 1; i <= Application.NormalTemplate.AutoTextEntries.Count; i++) {         let entry = Application.NormalTemplate.AutoTextEntries.Item(i)         if (entry.Name == "Test") {             entry.Insert(Selection.Range)         }     } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果修改了 Normal 模板，本示例将保存该模板。*/ function test() {     if (Application.NormalTemplate.Saved == false) {         Application.NormalTemplate.Save()     } }` |

#### **Application.NumLock**

此属性返回 Num Lock 键的状态。如果为 **True**，则数字键盘可用于输入数字；如果为 **False**，则该键盘用于移动插入点。**Boolean** 类型，只读。

**语法**

**express.NumLock**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回 Num Lock 键的当前状态。*/ let theState = Application.NumLock` |

#### **Application.OMathAutoCorrect**

返回一个 **OMathAutoCorrect** 对象，该对象代表公式的自动更正条目。只读。

**语法**

**express.OMathAutoCorrect**

*express*   一个代表 **Application** 对象的变量。

#### **Application.OpenAttachmentsInFullScreen**

返回或设置 **Boolean** 值，该值表示 WPS 是否在读取模式下打开电子邮件附件。可读写。

**语法**

**express.OpenAttachmentsInFullScreen**

*express*   一个代表 **Application** 对象的变量。

**说明**

此属性对应于**“ WPS 选项”**对话框中的**“在读取模式下打开电子邮件附件”**复选框。

#### **Application.Options**

返回一个 **Options** 对象，该对象代表 WPS 中的应用程序设置。

**语法**

**express.Options**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例打印 Sales.doc 文档及其批注和域结果。 */ function test() { 	Application.Options.PrintFieldCodes = false 	Application.Options.PrintComments = true 	Application.Documents("Sales.doc").PrintOut() }` |

#### **Application.Parent**

返回一个 **Object** 类型的值，该值代表指定 **Application** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **Application** 对象的变量。

#### **Application.Path**

返回指定对象的磁盘或 Web 路径。只读 **String** 类型。

**语法**

**express.Path**

*express*   一个代表 **Application** 对象的变量。

**说明**

此路径不包括尾部字符，例如，“C:\WPSOffice”或“http://MyServer”。使用 **PathSeparator** 属性可添加分隔文件夹与驱动器号的字符。使用 **Document** 对象的 **Name** 属性可返回不带路径的文件名，而用 **FullName** 属性可同时返回文件名和路径。

| 注释                                                         |
| ------------------------------------------------------------ |
| 可以用 **PathSeparator** 属性建立网站地址，即使此地址中会包含正斜线 (/)，而 **PathSeparator** 属性默认使用反斜线 (\)。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示活动文档的路径和文件名。*/ MsgBox(ActiveDocument.Path + Application.PathSeparator + ActiveDocument.Name)` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将当前文件夹路径改为活动文档附加的模板的路径。*/ ChDir ActiveDocument.AttachedTemplate.Path` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示 AddIns 集合中第一个加载项的路径。*/ function test() { if(AddIns.Count >= 1 ) {     MsgBox(AddIns.Item(1).Path) } }` |

#### **Application.PathSeparator**

返回用于分隔文件夹名称的字符。在 Windows 中，该属性返回一个反斜线 (\)。只读 **String** 类型。

**语法**

**express.PathSeparator**

*express*   一个代表 **Application** 对象的变量。

**说明**

可以用 **PathSeparator** 属性建立网站地址，尽管在地址中可能包含正斜线 (/)。

| 注释                                                         |
| ------------------------------------------------------------ |
| **FullName** 属性以单个字符串形式返回路径和文件名，其中包含路径分隔符。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示活动文档的路径和文件名。*/ alert(Application.ActiveDocument.Path + Application.PathSeparator + Application.ActiveDocument.Name)    ` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果第一个加载项是模板，本示例卸载该模板并打开它。*/ function test() {     if (Application.Addins.Item(1).Compiled == false) {         Application.Addins.Item(1).Installed = false         Application.Documents.Open(Application.AddIns.Item(1).Path + Application.PathSeparator + Application.AddIns.Item(1).Name)     } }` |

#### **Application.PickerDialog**

返回 PickerDialog 对象，以提供在对话框中选择用户或数据的功能。只读。

**语法**

**express.PickerDialog**

*express*   一个代表 **Application** 对象的变量。

**说明**

**PickerDialog** 对象可用于 WPS Office类型库。有关详细信息，请参阅下列对象及其成员：

- PickerDialog
- PickerField
- PickerFields
- PickerProperties
- PickerProperty
- PickerResult
- PickerResults

#### **Application.PortraitFontNames**

返回一个 **FontNames** 对象，该对象包括所有有效纵向字体的名称。

**语法**

**express.PortraitFontNames**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在插入点插入一组纵向字体。*/ function test() {     let sel = Application.Selection     for (let i = 1; i <= Application.PortraitFontNames.Count; i++) {         sel.Collapse(wdCollapseEnd)         sel.InsertAfter(PortraitFontNames.Item(i))         sel.InsertParagraphAfter()         sel.Collapse(wdCollapseEnd)     } }` |

#### **Application.PrintPreview**

如果当前视图是打印预览，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.PrintPreview**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将视图切换为打印预览。*/ Application.PrintPreview = true` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动窗口从打印预览切换至普通视图。*/ function test() {     Application.PrintPreview = false     Application.ActiveDocument.ActiveWindow.View.Type = wdNormalView }` |

#### **Application.ProtectedViewWindows**

返回代表所有受保护的视图窗口的 ProtectedViewWindows 集合。只读。

**语法**

**express.ProtectedViewWindows**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码示例显示打开的受保护的视图窗口数量。*/ alert("There are " + Application.ProtectedViewWindows.Count + " protected view windows open.")` |

#### **Application.RecentFiles**

返回一个 **RecentFiles** 集合，该集合代表最近存取过的文件。

**语法**

**express.RecentFiles**

*express*   一个代表 **Application** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例打开 RecentFiles 集合中的第一项（即“文件”菜单上列出的第一个文档名）。*/ function test() {     if (Application.RecentFiles.Count >= 1) {         Application.RecentFiles.Item(1).Open()     } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示 RecentFiles 集合中的每个文件的名称。*/ function test() {     for (let i = 1; i <= Application.RecentFiles.Count; i++) {         alert(Applicaiton.RecentFiles.Item(i).Name)     } }` |

#### **Application.RestrictLinkedStyles**

返回或设置 **Boolean** 值，该值表示 WPS 是否允许链接样式。可读写。

**语法**

**express.RestrictLinkedStyles**

*express*   一个代表 **Application** 对象的变量。

**说明**

链接样式是可以作为字符样式或段落样式应用的样式。此属性对应于**“样式”**对话框中的**“禁用链接样式”**复选框。

#### **Application.ScreenUpdating**

如果启用屏幕更新，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.ScreenUpdating**

*express*   一个代表 **Application** 对象的变量。

**说明**

当运行某一过程时，**ScreenUpdating** 属性控制显示器的主要变化情况。如果关闭屏幕更新，屏幕依然显示工具栏，并且 WPS 仍允许此过程用状态栏提示、输入框、对话框或信息框等工具来显示和检索信息。关闭屏幕更新可提高某些过程的执行速度。如果此过程已结束，或由于某些错误而停止，请务必将 **ScreenUpdating** 属性设为 **True**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将关闭屏幕更新，并添加一篇新文档。将 500 行文本添加至文档中。宏每隔 50 行选定一行并刷新屏幕。*/ function test() {     Application.ScreenUpdating = false     Application.Documents.Add()      for (let x = 1; x <= 500; x++) {         Application.ActiveDocument.Content.InsertAfter("This is line " + x + ".")         Application.ActiveDocument.Content.InsertParagraphAfter()          if (x % 50 == 0) {             Application.ActiveDocument.Paragraphs.Item(x).Range.Select()             Application.ScreenRefresh()         }     }      Application.ScreenUpdating = true }` |

#### **Application.Selection**

返回 **Selection** 对象，该对象代表一选定的范围或插入点。只读。

**语法**

**express.Selection**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 打印选区文本 */ function test() { 	if(Application.Selection.Type == wdSelectionNormal) {     	alert(Application.Selection.Text); 	} }` |

#### **Application.ShowStartupDialog**

如果为 **True**，则在启动 WPS 时显示**“任务窗格”**。**Boolean** 类型，可读写。

**语法**

**express.ShowStartupDialog**

*express*   一个代表 **Application** 对象的变量。

**说明**

**ShowStartupDialog** 属性是一个全局选项，新的设置仅在重新启动 WPS 后生效。使用 **CommandBars** 集合的 **Visible** 属性可以在不重新启动 WPS 的情况下显示或隐藏任务窗格。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例关闭“任务窗格”，使其不在启动 WPS 时显示。仅在用户下一次启动 WPS 时才生效。*/ function HideStartUpDlg(){     Application.ShowStartupDialog = false }` |

#### **Application.ShowStylePreviews**

返回或设置 **Boolean** 值，该值表示 WPS 是否在**“样式”**对话框中显示样式的格式预览。可读/写。

**语法**

**express.ShowStylePreviews**

*express*   一个代表 **Application** 对象的变量。

**说明**

此属性对应于**“样式”**对话框中的**“显示预览”**复选框。

#### **Application.ShowVisualBasicEditor**

如果显示“Visual Basic 编辑器”窗口，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.ShowVisualBasicEditor**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示“Visual Basic 编辑器”窗口。*/ Application.ShowVisualBasicEditor = true` |

#### **Application.ShowWindowsInTaskbar**

如果为 **True**，则在任务栏中显示打开的文档，默认值为 Single Document Interface (SDI)。如果为 **False**，则只在**“窗口”**菜单中列出打开的文档，提供 Multiple Document Interface (MDI) 的外观。**Boolean** 类型，可读写。

**语法**

**express.ShowWindowsInTaskbar**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例切换界面，以只在“窗口”菜单上列出打开的文档。*/ function SDIToMDI(){     Application.ShowWindowsInTaskbar = false }` |

#### **Application.SmartArtColors**

返回 SmartArtColors 对象，该对象代表应用程序中当前加载的颜色样式集。只读。

**语法**

**express.SmartArtColors**

*express*   一个代表 **Application** 对象的变量。

**说明**

**SmartArtColors** 属性代表的颜色集与**“更改颜色”**按钮上的可用颜色样式相对应，该按钮位于 WPS 2015 中**“SmartArt 工具”**上下文选项卡上的**“设计”**选项卡中。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码示例向活动文档中添加 SmartArt 图形，然后将该 SmartArt 图形颜色设置为“深色 2 轮廓”。*/ function test() {     let myShape     let mySmartArt     myShape = Application.ActiveDocument.Shapes.AddSmartArt(Application.SmartArtLayouts.Item(1), 50, 50, 200, 200)     mySmartArt = myShape.SmartArt     mySmartArt.Color = Application.SmartArtColors.Item(2) }` |

#### **Application.SmartArtLayouts**

返回 SmartArtLayouts 对象，该对象代表应用程序中当前加载的 SmartArt 布局集。只读。

**语法**

**express.SmartArtLayouts**

*express*   一个代表 **Application** 对象的变量。

**说明**

**SmartArtLayouts** 属性代表的布局集与**“布局”**组中的可用布局相对应，该组位于 WPS 2015 中**“SmartArt 工具”**上下文选项卡上的**“设计”**选项卡中。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码示例向活动文档中添加 SmartArt 图形，然后将该 SmartArt 图形布局设置为“分组列表”。*/ function test() {     let myShape     let mySmartArt     myShape = Application.ActiveDocument.Shapes.AddSmartArt(Application.SmartArtLayouts.Item(1), 50, 50, 200, 200, null)     mySmartArt = myShape.SmartArt     mySmartArt.Layout = Application.SmartArtLayouts.Item(15) }` |

#### **Application.SmartArtQuickStyles**

返回 SmartArtQuickStyles 对象，该对象代表应用程序中当前加载的 SmartArt 样式集。只读。

**语法**

**express.SmartArtQuickStyles**

*express*   一个代表 **Application** 对象的变量。

**说明**

**SmartArtQuickStyles** 属性代表的样式集与**“样式”**组中的可用样式相对应，该组位于 WPS 2015 中**“SmartArt 工具”**上下文选项卡上的**“设计”**选项卡中。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码示例向活动文档中添加 SmartArt 图形，然后将该 SmartArt 图形样式设置为“抛光”。*/ function test() {     let myShape     let mySmartArt     myShape = Application.ActiveDocument.Shapes.AddSmartArt(Application.SmartArtLayouts.Item(1), 50, 50, 200, 200, null)     mySmartArt = myShape.SmartArt     mySmartArt.QuickStyle = Application.SmartArtQuickStyles.Item(6) }` |

#### **Application.SmartTagRecognizers**

返回应用程序的一个 **SmartTagRecognizers** 集合。

**语法**

**express.SmartTagRecognizers**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下列示例访问 WPS 所安装的智能标记识别器列表中的第一个智能标记识别器。*/ Application.SmartTagRecognizers.Item(1)` |

#### **Application.SmartTagTypes**

返回一个 **SmartTagTypes** 集合，代表 WPS 中所安装的智能标记组件的智能标记类型。

**语法**

**express.SmartTagTypes**

*express*   一个代表 **Application** 对象的变量。

**说明**

智能标记类型是智能标记组件中的单个项目。智能标记组件可以包含多个智能标记类型。例如，安装在英文系统上的 Address (English) 智能标记列表默认情况下包含 name 智能标记类型、street 智能标记类型和 city 智能标记类型，等等。**SmartTagTypes** 集合包含用户计算机上安装的所有组件的所有智能标记类型。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下列示例循环遍历 SmartTagTypes 集合。如果 SmartTagType 是 Address 智能标记，则重新加载该智能标记的识别器和处理程序。*/ function GetSmartTagsTypes(){     let strSmartTagType     let objSmartTagType = Application.SmartTagTypes     strSmartTagType = "urn:schemas-microsoft-com" + ":office:smarttags#address"     for(let i=1; i<=objSmartTagType.Count; i++){         if(objSmartTagType == strSmartTagType){             objSmartTagType.Item(i).smartTagActions.ReloadActions()             objSmartTagType.Item(i).SmartTagRecognizers.ReloadRecognizers()         }     } }` |

#### **Application.SpecialMode**

如果 WPS 处于某个特殊模式（如 CopyText 模式或 MoveText 模式）下，则该属性值为 **True**。**Boolean** 类型，只读。

**语法**

**express.SpecialMode**

*express*   一个代表 **Application** 对象的变量。

**说明**

在选定文本后，按 F2 或 Shift+F2， WPS 将进入特殊的复制或移动模式。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例检查 WPS 是否处于某个特殊模式下。如果是，则在删除和粘贴当前所选内容之前，激活 Esc。*/ function test() {     if (Application.SpecialMode == true) {         SendKeys("ESC")         Selection.Cut()         Selection.EndKey(wdStory)         Selection.Paste()     } }` |

#### **Application.StartupPath**

返回或设置启动文件夹的完整路径，不包括最后的分隔符。**String** 类型，可读写。

**语法**

**express.StartupPath**

*express*   一个代表 **Application** 对象的变量。

**说明**

启动 WPS 时，将自动装载启动文件夹中的模板和加载项。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示启动文件夹的完整路径。*/ alert(Application.StartupPath)` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例使用户可更改启动文件夹的路径。*/ function test() {     let x = alert("Do you want to change the startup path?")     let newStartup = alert("Type a startup path")     Application.StartupPath = newStartup }` |

#### **Application.System**

返回一个 **System** 对象，该对象可用于返回与系统相关的信息，并执行与系统相关的任务。

**语法**

**express.System**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回有关系统的信息。*/ function test() { let processor = System.ProcessorType let enviro = System.OperatingSystem }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例建立到网络驱动器的连接。*/ Application.System.Connect("\\Project\\Info")` |

#### **Application.TaskPanes**

返回一个 **TaskPanes** 集合，该集合代表 WPS 中最常执行的任务。

**语法**

**express.TaskPanes**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下列示例显示包含文档中格式信息的任务栏。*/ function showFormatting() {     Application.TaskPanes.Item(wdTaskPaneFormatting).Visible = true }` |

#### **Application.Tasks**

返回一个 **Tasks** 集合，该集合代表所有正在运行的应用程序。

**语法**

**express.Tasks**

*express*   一个代表 **Application** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例查看 ET 是否正在运行。如果该任务正在运行，则本示例激活 ET；否则将显示一个消息框。*/ function test() {     if (Application.Tasks.Exists("ET") == true) {         Application.Tasks("ET").Activate()         Application.Tasks("ET").WindowState = wdWindowStateMaximize     } else {         alert("ET is not currently running.")     } }` |

#### **Application.Templates**

返回一个 **Templates** 集合，该集合代表所有可用的模板，包括共用模板和附加到打开文档的模板。

**语法**

**express.Templates**

*express*   一个代表 **Application** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示 Templates 集合中每个模板的名称。*/ function test() {     let Count = 1     for (let i = 1; i <= Application.Templates.Count; i++) {         alert(Application.Templates.Item(i).Name + " is template number " + Count)         Count = Count + 1     } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例中，如果模板 1 是一个共用模板，则其路径存储在 thePath 中。ChDir 语句用于将具有存储在 thePath 中的路径的文件夹设为当前文件夹。进行更改后，将显示“打开”对话框。*/ function test() {     if (Application.Templates.Item(1).Type == wdGlobalTemplate) {         let thePath = Application.Templates.Item(1).Path         if (thePath != "") {             ChDir(thePath)         }         Dialogs(wdDialogFileOpen).Show()     } }` |

#### **Application.Top**

返回或设置活动文档的垂直位置。**Long** 类型，可读写。

**语法**

**express.Top**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 WPS 应用程序窗口置于距屏幕顶部 100 磅的位置。*/ function test() {     Application.WindowState = wdWindowStateNormal     Application.Top = 100 }` |

#### **Application.UndoRecord**

返回 UndoRecord 对象，该对象提供撤消堆栈的自定义入口点。只读。

**语法**

**express.UndoRecord**

*express*   一个代表 **Application** 对象的变量。

**说明**

使用 **UndoRecord** 对象可在 WPS 2015 撤消堆栈中创建和修改自定义撤消记录。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码示例将实例化 UndoRecord 对象。*/ let objUndo = Application.UndoRecord` |

#### **Application.UsableHeight**

返回 WPS 文档窗口高度可设置的最大值（以磅为单位）。**Long** 类型，只读。

**语法**

**express.UsableHeight**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档窗口的尺寸设置为屏幕最大区域的四分之一。*/ function test() {     let Win = Application.ActiveDocument.ActiveWindow     Win.WindowState = wdWindowStateNormal     Win.Top = 5     Win.Left = 5     Win.Height = (Application.UsableHeight * 0.5)     Win.Width = (Application.UsableWidth * 0.5) }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示活动文档窗口中工作区的大小。*/ function test() {     let Win = Application.ActiveDocument.ActiveWindow     alert("Working area height = " + Win.UsableHeight + "\n" + "Working area width = " + Win.UsableWidth) }` |

#### **Application.UsableWidth**

返回 WPS 文档窗口可设置的最大宽度（以磅为单位）。**Long** 类型，只读。

**语法**

**express.UsableWidth**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档窗口的尺寸设置为屏幕最大区域的四分之一。*/ function test() { let Win = ActiveDocument.ActiveWindow Win.WindowState = wdWindowStateNormal Win.Top = 5 Win.Left = 5 Win.Height = (Application.UsableHeight*0.5) Win.Width = (Application.UsableWidth*0.5) }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示活动文档窗口中工作区的大小。*/ function test() {     let Win = Application.ActiveDocument.ActiveWindow     alert("Working area height = " + Win.UsableHeight + "\n" + "Working area width = " + Win.UsableWidth) }  ` |

#### **Application.UserAddress**

该属性返回或设置用户的通讯地址。**String** 类型，可读写。

**语法**

**express.UserAddress**

*express*   一个代表 **Application** 对象的变量。

**说明**

使用该通讯地址作为信封上的寄信人地址。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置用户的寄信人地址。Chr 函数用于返回一个换行符。*/ Application.UserAddress = "4200 Third Street NE" + "\n" + "Anytown, WA  98999"` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“工具”菜单的“选项”对话框中“用户信息”选项卡的“通讯地址”框中的地址。*/ alert(Application.UserAddress)` |

#### **Application.UserControl**

如果文档或应用程序是由用户创建或打开的，则该属性值为 **True**。**Boolean** 类型，只读。

**语法**

**express.UserControl**

*express*   一个代表 **Application** 对象的变量。

**说明**

如果应用程序是由另一 WPS Office应用程序使用 **Open** 方法、**CreateObject** 或 **GetObject** 方法以编程方式创建或打开的，则 **UserControl** 属性返回 **False**。

| 注释                                                         |
| ------------------------------------------------------------ |
| 如果 WPS 对用户可见，或者从代码模块中调用 **Application** 对象的 **UserControl** 属性，则该属性将始终返回 **True**。 |

**示例**

#### **Application.UserName**

该属性返回或设置用户姓名， WPS 将其用于信封和文档的“作者”属性。**String** 类型，可读写。

**语法**

**express.UserName**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置用户姓名。*/ Application.UserName = "Andrew Fuller"` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“工具”菜单上“选项”对话框中“用户信息”选项卡上“姓名”框中的姓名。*/ MsgBox(Application.UserName)` |

#### **Application.VBE**

返回一个 VBE 对象，该对象代表“Visual Basic 编辑器”。

**语法**

**express.VBE**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的示例将显示活动项目中可用的引用个数。*/ alert ("References = "+ Application..ActiveVBProject.References.Count)` |

#### **Application.Version**

返回 WPS 版本号。**String** 类型，只读。

**语法**

**express.Version**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例在信息框中显示 WPS 的版本号。 */ alert("The version of  WPS is " + Application.Version)` |

#### **Application.Visible**

如果指定对象可见，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.Visible**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                        |
| --------------------------------------------------- |
| `/*本示例隐藏 WPS 。*/ Application.Visible = false` |

#### **Application.Width**

返回或设置应用程序窗口的宽度（以磅为单位）。**Long** 类型，可读写。

**语法**

**express.Width**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 应用程序窗口的宽度和高度。*/ function test() {     Application.WindowState = wdWindowStateNormal     Application.Width = 500     Application.Height = 400 }` |

#### **Application.Windows**

返回一个 **Windows** 集合，该集合代表所有文档窗口。只读。

**语法**

**express.Windows**

*express*   一个代表 **Application** 对象的变量。

**说明**

此属性返回的集合中既包括可见窗口，也包括隐藏窗口。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 获取第一个窗口的题注 */ Application.Windows.Item(1).Caption` |

#### **Application.WindowState**

返回或设置指定文档窗口或任务窗口的状态。**WdWindowState** 类型，可读写。

**语法**

**express.WindowState**

*express*   一个代表 **Application** 对象的变量。

**说明**

**wdWindowStateNormal** 常量表示没有最大化或最小化的窗口。无法设置不活动窗口的状态。在设置窗口状态之前可用 **Activate** 方法激活窗口。

#### **Application.WordBasic**

返回一个自动化对象 (Word.Basic)，其中包含适用于 WPS 6.0 和 WPS for Windows 中所有 WordBasic 语句和函数的方法。只读。

**语法**

**express.WordBasic**

*express*   一个代表 **Application** 对象的变量。

**说明**

在 WPS 2000 及其后续版本中，当打开一个包含 WordBasic 宏的 WPS 6.0 或 WPS for Windows 模板时，宏将自动转换为 VB 模块。宏内部的每个 WordBasic 语句和函数都转换为相应的 Word.Basic 方法。

有关 WordBasic 语句和函数的信息，请参阅 WPS 6.0 或 WPS for Windows 中的“WordBasic 帮助”。有关将 WordBasic 转换为 Visual Basic 的信息，请参阅 将 WordBasic 宏转换为 Visual Basic。有关常规信息，请参阅 WordBasic 和 Visual Basic 在概念上的区别。

**示例**

#### **Application.XMLNamespaces**

返回一个 **XMLNamespaces** 集合，代表“架构库”中的 XML 架构。

**语法**

**express.XMLNamespaces**

*express*   一个代表 **Application** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下列示例返回“架构库”中的第一个架构。*/ let objSchema = Application.XMLNamespaces.Item(1)` |

适用环境：web

适用平台：windows/linux