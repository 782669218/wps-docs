**Document** 



代表一个文档。**Document** 对象是 **Documents** 集合的成员。**Documents** 集合中包含当前在 WPS 中打开的所有 **Document** 对象。

**说明**

使用 **Documents**(*Index*) 返回单个 **Document** 对象（其中 *Index* 是文档名称或索引编号）。以下示例关闭名为“Report.doc”的文档但不保存更改。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.Documents.Item("Report.doc").Close(wdDoNotSaveChanges)` |

索引编号代表文档在 **Documents** 集合中的位置。以下示例激活 **Documents** 集合中的第一篇文档。

| 示例代码复制                               |
| ------------------------------------------ |
| `Application.Documents.Item(1).Activate()` |

可以使用**ActiveDocument**属性引用具有焦点的文档。以下示例使用 **Activate**方法激活名为“Document 1”的文档，然后将页面方向设置为横向，并打印该文档。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test(){   Application.Documents.Item("Document1").Activate()   Application.ActiveDocument.PageSetup.Orientation = wdOrientLandscape   Application.ActiveDocument.PrintOut() }` |

**方法**

|                                                              | 名称                             | 说明                                                         |
| ------------------------------------------------------------ | -------------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AcceptAllRevisions**           | 接受对指定文档的所有修订。                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AcceptAllRevisionsShown**      | 接受显示在屏幕上的指定文档中的所有修订。                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Activate**                     | 激活指定的文档，使其成为活动文档。                           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddToFavorites**               | 创建文档或超链接的快捷方式，并将其添加到“收藏夹”。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ApplyQuickStyleSet2**          | 将指定的快速样式集应用于文档。                               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ApplyTheme**                   | 将主题应用于打开的文档。                                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AutoFormat**                   | 自动给文档套用格式。                                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CanCheckin**                   | 如果该方法返回值为 **True**，则 WPS 可将指定文档签入服务器。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CheckConsistency**             | 搜索日语文档中的所有文字，并显示对相同单词的字符用法不一致的实例。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CheckGrammar**                 | 开始对指定的文档或区域执行拼写和语法检查。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CheckIn**                      | 该方法将文档从本地计算机返回到服务器，并将本地文档设为只读，使之无法在本地进行编辑。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CheckInWithVersion**           | 将文档从本地计算机保存到服务器，并将本地文档设为只读，使之无法在本地进行编辑。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CheckSpelling**                | 开始对指定的文档或区域进行拼写检查。                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Close**                        | 关闭指定的文档。                                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ClosePrintPreview**            | 将指定文档从打印预览视图切换到以前的视图。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Compare**                      | 显示修订标记，以表明指定的文档与另一个文档的区别。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Compatibility**                | 如果为 **True**，则启用由 *Type* 参数指定的兼容性选项。兼容性选项将影响 WPS 中文档的显示方式。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ComputeStatistics**            | 返回基于指定文档内容的统计值。**Long** 类型。                |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Convert**                      | 将文件转换成最新的文件格式，并启用所有新功能。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ConvertAutoHyphens**           | 将由自动断字功能创建的连字符转换为手动连字符。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ConvertNumbersToText**         | 将指定文档中的列表编号和 LISTNUM 域更改为文本。              |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ConvertVietDoc**               | 使用非默认的代码页将一篇越南语文档重新转换为 Unicode。       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CopyStylesFromTemplate**       | 将指定模板中的样式复制到某篇文档。                           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CountNumberedItems**           | 返回指定 **Document** 对象中项目符号项或编号项以及 LISTNUM 域的数目。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CreateLetterContent**          | 生成并返回一个 **LetterContent** 对象，该对象基于指定信函内容。**LetterContent** 对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **DataForm**                     | 显示**“数据表单”**对话框，在此对话框中可添加、删除或修改记录。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **DeleteAllComments**            | 删除文档中 **Comments**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Comments/Comments%20.htm#jsObject_Comments)集合内的所有备注。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **DeleteAllCommentsShown**       | 删除显示在屏幕上的指定文档的所有修订。                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **DeleteAllEditableRanges**      | 删除（指定用户或用户组有权修改其权限的）所有区域中的权限。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **DeleteAllInkAnnotations**      | 删除文档中所有手写墨迹注释。                                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **DetectLanguage**               | 分析指定文本，以确定书写文本的语言类型。                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **DowngradeDocument**            | 将文档降级为 WPS 97-2003 文档格式，以便可以在以前版本的 WPS 中进行编辑。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **EndReview**                    | 终止审阅使用 **SendForReview** 方法发送以进行审阅的文件，或通过电子邮件将文档发送给另一用户而自动进行审阅的文件。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ExportAsFixedFormat**          | 将文档保存为 PDF 或 XPS 格式。                               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **FitToPages**                   | 适当缩小文本的字体大小，使文档的页数减少。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **FollowHyperlink**              | 显示一个缓存文档（如果该文档已下载）。否则，此方法将解析该超链接，下载目标文档，并在相应的应用程序中显示此文档。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ForwardMailer**                | 您请求就仅在 Macintosh 上使用的 Visual Basic 关键字获得帮助。有关 **Document** 对象的 **ForwardMailer** 方法的信息，请参阅 WPS OfficeMacintosh Edition 所附带的语言参考帮助。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **FreezeLayout**                 | 在 Web 视图中，按照文档当前所显示的样子固定文档的版式，以使换行保持不变，且在调整窗口大小时墨迹不会移动。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **GetCrossReferenceItems**       | 返回一个项目数组，根据指定的交叉引用类型可对该数组中的项目进行交叉引用。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **GetLetterContent**             | 检索指定文档中的信函内容并返回一个 **LetterContent** 对象。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **GetWorkflowTasks**             | 返回一个 **WorkflowTasks** 集合，该集合表示分配给文档的工作流任务。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **GetWorkflowTemplates**         | 返回一个 **WorkflowTemplates** 集合，该集合表示附加到文档的工作流模板。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **GoTo**                         | 返回一个 **Range**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)对象，该对象代表指定项（如页、书签或域）的起始位置。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **LockServerFile**               | 在服务器上锁定文件，以避免任何其他人进行编辑。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **MakeCompatibilityDefault**     | 兼容性选项位于**“选项”**对话框的**“兼容性”**选项卡上，它们是新文档的默认设置。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ManualHyphenation**            | 启动文档的手动断字设置，每次一行。                           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Merge**                        | 将文档中用修订标记标识的修改合并到另一篇文档。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Post**                         | 将指定的文档传送到 Microsoft Exchange 中的公用文件夹。       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PresentIt**                    | 打开 WPP，并加载指定的 WPS 文档。                            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PrintOut**                     | 打印指定文档的全部或部分内容。                               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PrintPreview**                 | 将视图切换到打印预览。                                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Protect**                      | 帮助保护指定的文档不被更改。保护文档后，用户只能进行有限的更改，例如添加注释，进行修订或填写表格。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Range**                        | 通过使用指定的开始和结束字符位置返回一个 **Range**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Redo**                         | 重复最后一次撤消的操作（**Undo** 方法的逆操作）。如果重复操作成功，则返回 **True**。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **RejectAllRevisions**           | 拒绝指定文档的所有修订。                                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **RejectAllRevisionsShown**      | 拒绝显示在屏幕上的文档的所有修订。                           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Reload**                       | 通过解析指向缓冲文档的超链接再下载该文档来重新加载该文档。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ReloadAs**                     | 使用指定的编码重新加载一篇基于 HTML 文档的文档。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **RemoveDocumentInformation**    | 从文档中删除敏感性信息、属性、批注和其他元数据。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **RemoveLockedStyles**           | 当文档中应用了格式设置限制时，清除锁定样式的文档。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **RemoveNumbers**                | 删除指定文档中的编号或项目符号。                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **RemoveTheme**                  | 删除当前文档中的活动主题。                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Repaginate**                   | 对整篇文档重新分页。                                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Reply**                        | 打开一封新的电子邮件（**“收件人”**行已含发件人地址）以答复活动邮件。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ReplyAll**                     | 打开一封新的电子邮件（**“收件人”**和**“抄送”**行已含发件人和所有其他收件人的地址）以答复活动邮件。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ReplyWithChanges**             | 为文档（该文档以被发送出去进行审阅）的作者发送电子邮件，通知他们审阅者已经审阅完文档。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ResetFormFields**              | 清除文档中的所有窗体域，准备再次填充窗体。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **RunAutoMacro**                 | 运行存储于指定文档的自动宏。如果没有自动宏，则不做任何操作。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **RunLetterWizard**              | 对指定的文档运行英文信函向导。                               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Save**                         | 保存当前文档。                                               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SaveAs2**                      | 使用新的名称或格式保存指定的文档。此方法的一些参数与**“另存为”**对话框（**“文件”**选项卡）中的选项相对应。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SaveAsQuickStyleSet**          | 保存当前所用的快速样式组。                                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Select**                       | 选择指定文档的内容。                                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SelectAllEditableRanges**      | 选择指定用户或用户组有权修改的所有区域。                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SelectContentControlsByTag**   | 返回一个 **ContentControls** 集合，该集合代表文档中具有 *Tag* 参数中指定标记值的所有内容控件。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SelectContentControlsByTitle** | 返回一个 **ContentControls** 集合，该集合代表文档中具有 *Title* 参数中指定标题的所有内容控件。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SelectLinkedControls**         | 返回一个 **ContentControls** 集合，该集合代表文档中根据 Node 参数的指定，链接到文档 XML 数据存储区中特定自定义 XML 节点的所有内容控件。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SelectNodes**                  | 返回一个 **XMLNodes** 集合，代表所有与 *XPath* 参数相匹配的节点并按其在文档或区域中显示的顺序排列。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SelectSingleNode**             | 返回一个 **XMLNode** 对象，该对象代表画布与指定文档中 *XPath* 参数相匹配的第一个节点。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SelectUnlinkedControls**       | 返回一个 **ContentControls** 集合，该集合代表文档中没有链接到文档 XML 数据存储区中 XML 节点的所有内容控件。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SendFax**                      | 将指定的文档作为传真发送，且无需用户交互。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SendFaxOverInternet**          | 将文档发送给传真服务提供商，该提供商将该文档传真给一个或多个指定收件人。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SendForReview**                | 将文档通过电子邮件发送给指定收件人进行审阅。                 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SendMail**                     | 打开一个邮件窗口以通过 Microsoft Exchange 发送指定文档。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SendMailer**                   | 您请求就仅在 Macintosh 上使用的 Visual Basic 关键字获得帮助。有关 **Document** 对象的 **SendMailer** 方法的信息，请参阅 WPS OfficeMacintosh Edition 所附带的语言参考帮助。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SetCompatibilityMode**         | 设置文档的兼容模式。                                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SetDefaultTableStyle**         | 指定用于文档中新建表的表样式。                               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SetLetterContent**             | 将指定 **LetterContent** 对象的内容插入文档。                |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SetPasswordEncryptionOptions** | 设置 WPS 用于通过密码加密文档的选项                          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ToggleFormsDesign**            | 在打开和关闭窗体设计模式之间切换。                           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **TransformDocument**            | 将 Extensible Stylesheet Language Transformation (XSLT) 文件应用于指定文档，并用结果替换该文档。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Undo**                         | 消最后一次操作或最后一系列操作，这些操作显示在**“撤消”**列表中。返回 **Boolean** 。如果撤消操作成功，则返回 **True**。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **UndoClear**                    | 清除可对指定文档撤消的操作列表。                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Unprotect**                    | 清除对指定文档的保护。                                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **UpdateStyles**                 | 将所有样式从附加模板复制到文档中，同时覆盖文档中所有现有同名的样式。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **UpdateSummaryProperties**      | 更新**“文件”**菜单**“属性”**对话框中的关键词和备注文字，以反映为指定文档编写的自动摘要内容。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ViewCode**                     | 显示指定文档中所选 Microsoft ActiveX 控件的代码窗口。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ViewPropertyBrowser**          | 显示指定文档中所选 Microsoft ActiveX 控件的属性窗口。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **WebPagePreview**               | 显示将当前文档保存为网页时的预览。                           |

**属性**

|                                                              | 名称                                 | 说明                                                         |
| ------------------------------------------------------------ | ------------------------------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ActiveTheme**                      | 返回指定文档的活动主题名称及该主题的格式选项。**String** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ActiveThemeDisplayName**           | 返回指定文档的活动主题的显示名称。**String** 类型，只读。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ActiveWindow**                     | 返回一个 **Window** 对象，该对象代表活动窗口（焦点所在的窗口）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ActiveWritingStyle**               | 返回或设置指定文档中指定语言的写作风格。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**                      | 返回一个代表 WPS 应用程序的 **Application** 对象。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AttachedTemplate**                 | 返回一个 **Template** 对象，该对象代表与指定文档相关联的模板。可读/写 **Variant** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoFormatOverride**               | 返回或设置一个 **Boolean** 类型的值，该数值代表是否在已应用格式设置限制的文档中使用自动设置格式替代格式设置限制。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoHyphenation**                  | 如果代表指定文档已启动自动断字功能，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Background**                       | 返回一个 **Shape** 对象，该对象代表指定文档的背景图像。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Bibliography**                     | 返回一个 **Bibliography** 对象，代表文档中包含的书目参考。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Bookmarks**                        | 返回一个 **Bookmarks** 集合，该集合代表文档中的所有书签。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BuiltInDocumentProperties**        | 返回一个 **DocumentProperties** 集合，该集合代表指定文档的所有内置的文档属性。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Characters**                       | 返回一个 **Characters** 集合，该集合代表文档中的字符。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ChildNodeSuggestions**             | 返回一个 **XMLChildNodeSuggestions** 集合，该集合代表允许在文档中使用的元素列表。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ClickAndTypeParagraphStyle**       | 返回或设置“即点即输”功能在指定文档中应用于文字的默认段落样式。**Variant** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CoAuthoring**                      | 返回 CoAuthoring 对象，该对象提供文档中与共同创作相关的对象模型的入口点。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CodeName**                         | 返回指定文档的代码名称。**String** 类型，只读。              |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CommandBars**                      | 返回一个 **CommandBars** 集合，该集合代表 WPS 中的菜单栏以及所有工具栏。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Comments**                         | 返回一个 **Comments** 集合，该集合代表指定文档的所有批注。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CompatibilityMode**                | 返回 **Number** 类型值，该值指定打开文档时 WPS 2015 使用的兼容模式。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ConsecutiveHyphensLimit**          | 返回或设置以连字符结束的连续行的最大数目。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Container**                        | 返回一个对象，该对象代表指定文档的容器应用程序。**Object** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Content**                          | 返回一个 **Range** 对象，该对象代表主文档文章。只读。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ContentControls**                  | 返回一个 **ContentControls** 集合，该集合代表文档中的所有内容控件。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ContentTypeProperties**            | 返回一个 **MetaProperties** 集合，该集合代表文档中所存储的元数据，如作者姓名、主题和公司。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**                          | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CurrentRsid**                      | 返回一个 **Long** 类型的值，该值代表 WPS 为文档中的更改所分配的随机数。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CustomDocumentProperties**         | 返回一个 **DocumentProperties** 集合，该集合代表指定文档的所有自定义文档属性。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CustomXMLParts**                   | 返回一个 **CustomXMLParts** 集合，该集合代表 XML 数据存储区中的自定义 XML。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DefaultTableStyle**                | 返回一个 **Variant** 类型的值，该值代表可以应用于文档中新创建的所有表格的表格样式。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DefaultTabStop**                   | 以磅为单位返回或设置指定文档中的默认制表位间隔。**Single** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DefaultTargetFrame**               | 返回或设置一个 **String** 类型的值，该值代表通过超链接可到达显示网页的浏览器框架。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisableFeatures**                  | 如果为 **True**，则禁用 **DisableFeaturesIntroducedAfter** 属性中指定的版本之后的所有功能。默认值为 **False**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisableFeaturesIntroducedAfter**   | 仅对该文档禁用指定的 WPS 版本后引入的所有功能。**WdDisableFeaturesIntroducedAfter** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DocumentInspectors**               | 返回一个 **DocumentInspectors** 集合，该集合使您能够查找隐藏的个人信息，如作者姓名、公司名称和修订日期。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DocumentLibraryVersions**          | 返回一个 **DocumentLibraryVersions** 集合，该集合表示共享文档的版本集合，该共享文档启用了版本控制并存储在服务器的文档库中。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DocumentTheme**                    | 返回一个 **OfficeTheme** 对象，代表应用于文档的 WPS Office主题。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DoNotEmbedSystemFonts**            | 如果为 **True**，则 WPS 不嵌入常规系统字体。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Email**                            | 返回一个 **Email** 对象，该对象包含当前文档的所有与电子邮件相关的属性。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EmbedLinguisticData**              | 如果为 **True**，则 WPS 会嵌入语音和手写功能，以便数据可以转换回语音或手写。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EmbedTrueTypeFonts**               | 如果在保存文档时，WPS 在文档中嵌入 TrueType 字体，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EncryptionProvider**               | 返回一个 **String** 类型的值，该值指定 WPS 在对文档进行加密时使用的算法加密提供程序的名称。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Endnotes**                         | 返回一个 **Endnotes** 集合，该集合代表文档中的所有尾注。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EnforceStyle**                     | 返回或设置一个 **Boolean** 类型的值，该值代表是否在受保护的文档中实施格式设置限制。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Envelope**                         | 返回一个 **Envelope** 对象，该对象代表文档中的信封和信封功能。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FarEastLineBreakLanguage**         | 返回或设置一个 **WdFarEastLineBreakLanguageID** 类型的值，该值代表在指定文档或模板中对文本进行换行时使用的东亚语言。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FarEastLineBreakLevel**            | 返回或设置一个 **WdFarEastLineBreakLevel** 类型的值，该值代表指定文档的换行控制级别。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Fields**                           | 返回一个 **Fields**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Fields/Fields%20.htm#jsObject_Fields)集合，该集合代表文档中的所有域。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Final**                            | 返回或设置 **Boolean** 值，该值指示文档是否是最终的。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Footnotes**                        | 返回一个 **Footnotes**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Footnotes/Footnotes%20.htm#jsObject_Footnotes)集合，该集合代表文档中的所有脚注。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FormattingShowClear**              | 如果为 **True**，则 WPS 在**“样式和格式”**任务窗格中显示“清除格式”。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FormattingShowFilter**             | 设置或返回一个 **WdShowFilter** 常量，代表**“样式和格式”**任务窗格中显示的样式和格式。**Boolean** 类型，可读写 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FormattingShowFont**               | 如果为 **True**，则 WPS 显示**“样式和格式”**任务窗格中的字体格式。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FormattingShowNextLevel**          | 返回或设置 **Boolean** 值，该值表示 WPS 在使用上一个标题级别时是否显示下一个标题级别。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FormattingShowNumbering**          | 如果为 **True**，则 WPS 在**“样式和格式”**任务窗格中显示编号格式。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FormattingShowParagraph**          | 如果为 **True**，则 WPS 在**“样式和格式”**任务窗格中显示段落格式。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FormattingShowUserStyleName**      | 返回或设置 **Boolean** 值，该值表示是否显示用户定义的样式。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FormFields**                       | 返回一个 **FormFields** 集合，该集合代表文档中的所有窗体域。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FormsDesign**                      | 如果指定的文档处于窗体设计模式，则该属性值为 **True**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FormsDesignFormsDesign**           | 如果指定的文档处于窗体设计模式，则该属性值为 **True**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Frames**                           | 返回一个 **Frames** 集合，该集合代表文档中的所有图文框。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Frameset**                         | 返回一个 **Frameset** 对象，该对象代表一个整体框架页或框架页的单个框架。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FullName**                         | 返回一个 **String** 类型的值，该值代表包括路径的文档的名称。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **GrammarChecked**                   | 如果已经检查了指定范围或文档的语法，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **GrammaticalErrors**                | 返回一个 **ProofreadingErrors** 集合，该集合代表指定的文档中有语法检查错误的句子。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **GridDistanceHorizontal**           | 在指定文档中绘制、移动和调整自选图形或东亚语言字符时，返回或设置代表 WPS 所用的不可见的网格线之间的水平距离的 **Single** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **GridDistanceVertical**             | 在指定的文档中绘制、移动和调整自选图形或东亚语言字符时，返回或设置代表 WPS 所用的不可见网格线之间的垂直距离的 **Single** 类型。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **GridOriginFromMargin**             | 如果 WPS 将在页面左上角开始使用字符网格，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **GridOriginHorizontal**             | 返回或设置一个 **Single** 类型的值，该值代表不可见网格相对于页面左边的起点位置，当在指定的文档中绘制、移动和调整自选图形或东亚语言字符时将使用该网格。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **GridOriginVertical**               | 返回或设置一个 **Single** 类型的值，该值代表不可见网格相对于页面顶边的起点位置，当在指定的文档中绘制、移动和调整自选图形或东亚语言字符时将使用该网格。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **GridSpaceBetweenHorizontalLines**  | 返回或设置 WPS 在页面视图中显示的水平字符网格的间隔。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **GridSpaceBetweenVerticalLines**    | 返回或设置 WPS 在页面视图中显示的垂直字符网格线间隔。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HasMailer**                        | 您请求就仅在 Macintosh 上使用的 Visual Basic 关键字获得帮助。有关 **Document** 对象的 **HasMailer** 属性的信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HasPassword**                      | 如果需要密码才能打开指定文档，则该属性值为 **True**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HasVBProject**                     | 返回一个 **Boolean** 类型的值，该值代表文档是否具有附加的 示例代码 项目。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HTMLDivisions**                    | 返回一个 **HTMLDivisions** 集合，该集合代表 Web 文档中的 HTML DIV 元素。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Hyperlinks**                       | 返回一个 **Hyperlinks** 集合，该集合代表指定文档中的所有超链接。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HyphenateCaps**                    | 如果全是大写字母时可对其进行断字，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HyphenationZone**                  | 返回或设置断字区的宽度（以磅为单位）。**Long** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Indexes**                          | 返回一个 **Indexes** 集合，该集合代表指定文档的所有索引。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **InlineShapes**                     | 返回一个 **InlineShapes** 集合，该集合代表文档中的所有 **InlineShape** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IsMasterDocument**                 | 如果指定的文档为一篇主文档，则该属性值为 **True**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IsSubdocument**                    | 如果指定文档是主文档的子文档，则该属性值为 **True**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **JustificationMode**                | 返回或设置指定文档字符间距的调整量。可读/写 **WdJustificationMode** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **KerningByAlgorithm**               | 如果 WPS 在指定文档中调整半角拉丁字符和标点符号的间距，则为 **True**。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Kind**                             | 返回或设置 WPS 在自动设置指定文档的格式时使用的格式类型。**WdDocumentKind** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LanguageDetected**                 | 返回或设置一个指定 WPS 是否对指定文本的语言进行检测的值。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ListParagraphs**                   | 返回一个 **ListParagraphs** 对象，该对象代表文档中的所有编号段落。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Lists**                            | 返回一个 **Lists** 集合，该集合含有指定文档中的所有项目符号和编号列表。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ListTemplates**                    | 返回一个 **ListTemplates** 集合，该集合代表指定文档中的所有列表格式。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LockQuickStyleSet**                | 返回或设置一个 **Boolean**，它代表用户是否可以更改使用的快速样式集。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LockTheme**                        | 返回或设置一个 **Boolean** 类型的值，该值代表用户是否可以更改文档主题。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MailEnvelope**                     | 返回一个 **MsoEnvelope** 对象，该对象代表文档的电子邮件标题。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Mailer**                           | 您请求就仅在 Macintosh 上使用的 Visual Basic 关键字获得帮助。有关 **Document** 对象的 **Mailer** 属性的信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **MailMerge**                        | 返回一个 **MailMerge** 对象，该对象代表指定文档的邮件合并功能。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**                             | 返回指定文档的名称。**String** 类型，只读。                  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NoLineBreakAfter**                 | 返回或设置 WPS 不在其后换行的首尾字符。可读/写 **String** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NoLineBreakBefore**                | 返回或设置 WPS 不在其前换行的首尾字符。可读/写 **String** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OMathBreakBin**                    | 返回或设置一个 **WdOMathBreakBin** 常量，该常量代表当公式跨越两行或多行时 WPS 用来放置二元运算符的位置。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OMathBreakSub**                    | 返回或设置一个 **WdOMathBreakSub** 常量，该常量代表 WPS 如何处理位于换行符之前的减法运算符。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OMathFontName**                    | 返回或设置一个 **String** 类型的值，该值代表文档中用于显示公式的字体的名称。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OMathIntSubSupLim**                | 返回或设置一个 **Boolean** 类型的值，该值代表积分的默认极限位置。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OMathJc**                          | 返回或设置一个 **WdOMathJc** 常量，该常量代表一组公式的默认对齐方式：左对齐、右对齐、居中或作为一个组居中。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OMathLeftMargin**                  | 返回或设置一个代表公式的左边距的 **Single**。可读写。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OMathNarySupSubLim**               | 返回或设置一个 **Boolean** 类型的值，该值代表 *n* 元对象（而不是积分）的默认极限位置。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OMathRightMargin**                 | 返回或设置一个 **Single** 类型的值，该值代表公式的右边距。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OMaths**                           | 返回一个 **OMaths** 集合，该集合代表指定区域内的 **OMath** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OMathSmallFrac**                   | 返回或设置一个 **Boolean** 类型的值，该值代表是否在文档所包含的公式中使用小型分数。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OMathWrap**                        | 返回或设置一个 **Single** 类型的值，该值代表换到新的一行的公式的第二行位置。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OpenEncoding**                     | 返回用于打开指定文档的编码。**MsoEncoding** 类型，只读。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OptimizeForWord97**                | 如果 WPS 通过禁用所有不兼容的格式来优化当前文档，以便在 WPS 97 中查看，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OriginalDocumentTitle**            | 返回一个 **String** 类型的值，该值代表在运行“精确比较”文档比较功能之后原文档的文档标题。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PageSetup**                        | 返回一个与指定文档相关联的 **PageSetup** 对象。              |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Paragraphs**                       | 返回一个 **Paragraphs** 集合，该集合代表指定文档中的所有段落。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**                           |                                                              |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Password**                         | 设置在打开文档必须使用的密码。**String** 类型，只写。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PasswordEncryptionAlgorithm**      | 返回一个 **String** 类型的值，该值代表 WPS 用密码加密文档时所用的算法。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PasswordEncryptionFileProperties** | 如果 WPS 为密码保护的文档的文件属性加密，则该属性值为 **True**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PasswordEncryptionKeyLength**      | 返回一个 **Long** 类型的值，该值代表 WPS 用密码加密文档时所用的密码键长。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PasswordEncryptionProvider**       | 返回一个 **String** 类型的值，该值指定 WPS 用密码加密文档时所用的加密算法提供程序的名称。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Path**                             | 返回文档的磁盘或 Web 路径。只读 **String** 类型。            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Permission**                       | 返回一个 **Permission** 对象，该对象代表指定文档中的权限设置。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PrintFormsData**                   | 如果 WPS 仅将相应联机窗体中的键入的数据打印在预打印窗体上，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PrintFractionalWidths**            | 如果设置指定文档的格式，使用分式间距来显示和打印字符，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PrintPostScriptOverText**          | 如果为 **True**，则使用 PostScript 打印机时，文档中的 PRINT 域指令（例如，PostScript 命令）将打印在文字和图形的上方。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PrintRevisions**                   | 如果为 **True**，则在打印文档的同时打印修订标记。如果返回 **False**，则不打印修订标记（即打印接受修订后的状态）。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Protect**                          | 返回或设置与指定传送名单相关联的文档的保护类型。**WdProtectionType** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ProtectionType**                   | 该属性返回指定文档的保护类型。可以是下列 **WdProtectionType** 常量之一：**wdAllowOnlyComments**、**wdAllowOnlyFormFields**、**wdAllowOnlyReading**、**wdAllowOnlyRevisions** 或 **wdNoProtection**。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ReadabilityStatistics**            | 返回一个 **ReadabilityStatistics** 集合，该集合代表指定文档或区域的可读性统计信息。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ReadingLayoutSizeX**               | 设置或返回一个 **Long**类型的值，该值表示当在阅读版式视图中显示文档并对其冻结了输入手写标记时文档中页面的宽度。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ReadingLayoutSizeY**               | 设置或返回一个 **Long** 类型的值，该值表示当在阅读版式视图中显示文档并对其冻结了输入手写标记时文档中页面的高度。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ReadingModeLayoutFrozen**          | 设置或返回一个 **Boolean** 类型的值，该值表示是否将阅读版式视图中显示的页面冻结为指定大小以向文档插入手写标记。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ReadOnly**                         | 如果对文档所作修改不能保存到原始文档中，则该属性为 **True**。只读 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ReadOnlyRecommended**              | 如果无论用户何时打开文档，WPS 都显示一个消息框，提示以只读方式打开，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RemoveDateAndTime**                | 设置或返回一个 **Boolean** 类型的值，该值指示文档是否存储修订的日期和时间元数据。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RemovePersonalInformation**        | 如果 WPS 在保存文档时从批注、修订和“属性”对话框中删除所有用户信息，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Research**                         | 返回一个 **Research** 对象，该对象代表文档的检索服务。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RevisedDocumentTitle**             | 返回一个 **String** 类型的值，该值代表在运行“精确比较”文档比较功能之后修订的文档的文档标题。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Revisions**                        | 返回一个 **Revisions**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Revisions/Revisions%20.htm#jsObject_Revisions)集合，该集合代表文档或区域中的修订。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Saved**                            | 如果指定的文档或模板从上次保存后一直没有更改，则该属性值为 **True**。如果关闭文档时，WPS 提示保存对文档所做的更改，则该属性值为 **False**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SaveEncoding**                     | 返回或设置保存文档时使用的编码。**MsoEncoding** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SaveFormat**                       | 返回指定文档或文件转换器的文件格式。只读 **Long** 类型。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SaveFormsData**                    | 如果 WPS 将输入到一个窗体中的数据作为以制表位分隔的数据库记录保存，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SaveSubsetFonts**                  | 如果 WPS 将嵌入的 TrueType 字体子集和文档一起保存，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Scripts**                          | 返回一个 **Scripts** 集合，该集合代表指定对象中的 HTML 脚本的集合。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Sections**                         | 返回一个 **Section** 集合，该集合代表指定文档中的节。只读。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Sentences**                        | 返回一个 **Sentences** 集合，该集合代表文档中的所有句子。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ServerPolicy**                     | 返回一个 **ServerPolicy** 对象，该对象代表为运行 SharePoint Server 2007 的服务器上存储的文档指定的策略。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Shapes**                           | 返回一个 **Shapes** 集合，该集合代表指定文档中的所有 **Shape** 对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowGrammaticalErrors**            | 如果指定文档中的语法错误用绿色波浪线标记，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowRevisions**                    | 如果在屏幕上显示对指定文档的修订，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowSpellingErrors**               | 如果 WPS 为文档中的拼写错误添加下划线，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Signatures**                       | 返回一个 **SignatureSet** 集合，该集合代表文档的数字签名。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SmartDocument**                    | 返回一个 **SmartDocument** 对象，该对象代表智能文档解决方案的设置。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SmartTagsAsXMLProps**              | 如果该属性的值为 **True**，当将包含智能标记的文档保存为 HTML 格式时，WPS 会创建一个包含智能标记信息的 XML 页眉。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SnapToGrid**                       | 如果在指定文档中绘制、移动自选图形/东亚字符或者调整它们的大小时，自选图形/东亚字符自动与不可见的网格线对齐，则该属性为 **True**。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SnapToShapes**                     | 如果 WPS 在指定文档中使自选图形或东亚字符自动与不可见网格线对齐，这些网格线穿过其他自选图形或东亚字符的水平和垂直边缘，则该属性为 **True**。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SpellingChecked**                  | 如果已对指定的区域或文档完成拼写检查，则该属性值为 **True**。如果所有或部分区域或文档尚未进行拼写检查，则该属性值为 **False**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SpellingErrors**                   | 返回一个 **ProofreadingErrors** 集合，该集合代表在指定文档或区域中标识为拼写错误的单词。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **StoryRanges**                      | 返回一个 **StoryRanges** 集合，该集合代表指定文档中的所有的文字部分。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Styles**                           | 本示例返回一个指定文档的 **Styles** 集合。只读。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **StyleSortMethod**                  | 返回或设置一个 **WdStyleSort** 常量，该常量代表要在对**“样式”**任务窗格中的样式进行排序时使用的排序方法。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Subdocuments**                     | 返回一个 **Subdocuments** 集合，该集合代表指定文档中的所有子文档。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Sync**                             | 返回一个 **Sync** 对象，该对象提供对构成“文档工作区”的文档的方法和属性的访问。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Tables**                           | 返回一个 **Table**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Table/Table%20.htm#jsObject_Table)集合，该集合代表指定文档中的所有表格。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TablesOfAuthorities**              | 返回一个 **TableOfAuthorities** 集合，该集合代表指定文档中的引文目录。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TablesOfAuthoritiesCategories**    | 返回一个 **TablesOfAuthoritiesCategories** 集合，该集合代表指定文档中可用的引文目录类别。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TablesOfContents**                 | 返回一个 **TablesOfContents** 集合，该集合代表指定文档中的目录。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TablesOfFigures**                  | 返回一个 **TablesOfFigures** 集合，该集合代表指定文档中的图表目录。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextEncoding**                     | 返回或设置代码页或字符集，WPS 应用该代码页或字符集将文档保存为编码文本文件。**MsoEncoding** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TextLineEnding**                   | 返回或设置一个 **WdLineEndingType** 常量，该常量代表 WPS 在保存为文本文件的文档中标记直线和段落分隔符的方式。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TrackFormatting**                  | 返回或设置 **Boolean** 值，该值表示在打开更改跟踪功能时是否跟踪格式更改。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TrackMoves**                       | 返回或设置 **Boolean** 值，该值表示在打开跟踪更改功能时是否对移动文本进行标记。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TrackRevisions**                   | 如果标记对指定文档的修改，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**                             | 返回文档的类型（模板或文档）。只读 **WdDocumentType**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdDocumentType%20%E6%9E%9A%E4%B8%BE.html)类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **UpdateStylesOnOpen**               | 如果在每次打开指定文档时会依照其附属的模板更新文档样式，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **UseMathDefaults**                  | 返回或设置一个代表在创建新的公式时是否使用默认的数学设置的 **Boolean**。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **UserControl**                      | 如果文档是由用户创建或打开的，则该属性为 **True**。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Variables**                        | 返回一个 **Variables** 集合，该集合代表指定文档所存的变量。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **VBASigned**                        | 如果指定文档的 示例代码 (VBA) 项目具有数字签名，则该属性值为 **True**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **VBProject**                        | 返回指定模板或文档的 **VBProject** 对象。                    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **WebOptions**                       | 将文档保存为网页或打开网页时，该属性将返回 **WebOptions** 对象，该对象包含 WPS 所用的文档属性。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Windows**                          | 返回一个 **Windows** 集合，该集合代表指定文档的所有窗口。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **WordOpenXML**                      | 返回一个 **String** 类型的值，该值代表文档的 WPS Open XML 内容的 Flat XML 格式。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Words**                            | 返回一个 **Words** 集合，该集合代表文档中的所有字词。只读。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **WritePassword**                    | 该属性设置一个保存对指定文档所做的修改时所需的密码。**String** 类型，只写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **WriteReserved**                    | 如果指定文档设有写权限密码，则该属性值为 **True**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **XMLHideNamespaces**                | 返回一个 **Boolean** 类型的值，该值表示是否隐藏 XML 命名空间，该命名空间位于“XML 结构”任务窗格的元素列表中。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **XMLNodes**                         | 返回一个 **XMLNodes** 集合，该集合代表文档中的所有 XML 元素。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **XMLSaveDataOnly**                  | 设置或返回一个 **Boolean** 类型的值，保存文档时包含 XML 标记还是仅保存文本。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **XMLSaveThroughXSLT**               | 设置或返回一个 **String** 类型的值，指定用户保存文档时，Extensible Stylesheet Language Transformation (XSLT) 要应用的路径和文件名。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **XMLSchemaReferences**              | 返回一个 XMLSchemaReferences 集合，代表附加到文档的架构。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **XMLSchemaViolations**              | 返回一个 XMLNodes 集合，代表文档中所有存在验证错误的节点。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **XMLShowAdvancedErrors**            | 返回或设置一个 **Boolean** 类型的值，该值表示是由内置的 WPS 错误消息还是由包括在 Office 中的 Microsoft XML Core Services (MSXML) 5.0 组件来生成错误消息文本。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **XMLUseXSLTWhenSaving**             | 返回一个 **Boolean** 类型的值，代表是否通过 Extensible Stylesheet Language Transformation (XSLT) 保存文档。如果通过 XSLT 保存文档，则该值为 **True**。 |

**成员方法**

#### **Document.AcceptAllRevisions**

接受对指定文档的所有修订。

**语法**

**express.AcceptAllRevisions()**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例检查文档是否有修订，有的话接受所有修订 */ function test() { 	if(Application.ActiveDocument.Revisions.Count >= 1){   		Application.ActiveDocument.AcceptAllRevisions() 	}	 }` |

#### **Document.AcceptAllRevisionsShown**

接受显示在屏幕上的指定文档中的所有修订。

**语法**

**express.AcceptAllRevisionsShown()**

*express*   一个代表 **Document** 对象的变量。

**说明**

可用 **RejectAllRevisionsShown** 方法来拒绝显示在屏幕上的指定文档中的所有修订。

#### **Document.Activate**

激活指定的文档，使其成为活动文档。

**语法**

**express.Activate()**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 激活第一个文档 */ Application.Documents.Item(1).Activate() ` |

#### **Document.AddToFavorites**

创建文档或超链接的快捷方式，并将其添加到“收藏夹”。

**语法**

**express.AddToFavorites()**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为活动文档的所有超链接创建快捷方式，并将其添加到“收藏夹”文件夹。*/ function test() {     for (let i = 1; i <= Application.ActiveDocument.Hyperlinks.Count; i++) {         Application.ActiveDocument.Hyperlinks.Item(i).AddToFavorites()     } } ` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为 Sales.doc 创建快捷方式，并将其添加至“收藏夹”文件夹。如果 Sales.doc 还未打开，本示例将从 C:\Documents 文件夹打开该文档。*/ function test() {     let doc     for (let i = 1; i <= Application.Documents.Count; i++) {         doc = Application.Documents.Item(i)         if (doc.Name.toLowerCase() == "sales.doc") {             isOpen = true         }     }     if (isOpen != true) {         Application.Documents.Open("C:\\Documents\\Sales.doc")     }     Application.Documents.Item("Sales.doc").AddToFavorites() }` |

#### **Document.ApplyQuickStyleSet2**

将指定的快速样式集应用于文档。

**语法**

**express.ApplyQuickStyleSet2(Style)**

*express*   一个代表 **Document** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Style*  | 必选          | **Variant**  | 可以是指定要使用的集合名称（与“样式集”列表中列出的名称对应）的 String，也可以是 WdApplyQuickStyleSets 枚举中的常量。 |

#### **Document.ApplyTheme**

将主题应用于打开的文档。

**语法**

**express.ApplyTheme(Name)**

*express*   一个代表 **Document** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Name*   | 必选          | **String**   | 主题名称及要应用的任何主题格式选项。该字符串格式为“themennn”，其中 theme 和 nnn 的定义如下：主题: 包含请求主题的数据的文件夹名称（主题数据文件夹的默认位置是 C:\Program Files\Common Files\WPS Shared\Themes）。主题必须使用该文件夹名，不能使用“主题”对话框中显示的名称。nnn: 三个数字的字符串，表示要激活的主题格式选项（1 为激活，0 为停用）。每个数字分别对应“主题”对话框中的“鲜艳颜色”、“动态图形”和“背景图像”复选框。如果省略该字符串，则 nnn 的默认值为“011”（“动态图形”和“背景图像”均被激活）。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例实现的功能是：向活动文档添加“Artsy”主题，并激活“鲜艳颜色”选项。*/ Application.ActiveDocument.ApplyTheme("artsy 100")` |

#### **Document.AutoFormat**

自动给文档套用格式。

**语法**

**express.AutoFormat()**

*express*   一个代表 **Document** 对象的变量。

**说明**

使用 **Kind** 属性指定文档类型。

#### **Document.CanCheckin**

如果该方法返回值为 **True**，则 WPS 可将指定文档签入服务器。可读/写 **Boolean** 类型。

**语法**

**express.CanCheckin()**

*express*   一个代表 **Document** 对象的变量。

**返回值**

Boolean

**说明**

要利用 WPS 中内置的协作功能，文档必须存储在 Microsoft SharePoint Portal Server 上。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例检查服务器并确定指定的文档是否可以签入；如果可以，则关闭文档并将其签回服务器。*/ function CheckInOut(docCheckIn) {     if (Application.Documents.Item(docCheckIn).CanCheckin() == true) {         Application.Documents.Item(docCheckIn).CheckIn()         alert(docCheckIn + " has been checked in.")     }     else {         alert("This file cannot be checked in " + "at this time.  Please try again later.")     } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*若要调用上述 CheckInOut 子程序，请使用下列子程序并将文件名“http://servername/workspace/report.doc”替换为以上“说明”一节中提到的服务器上实际的文件。*/ function CheckDocInOut(){     CheckInOut ("http://servername/workspace/report.doc") }` |

#### **Document.CheckConsistency**

搜索日语文档中的所有文字，并显示对相同单词的字符用法不一致的实例。

**语法**

**express.CheckConsistency()**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例检查活动文档中日语字符的一致性。*/ ActiveDocument.CheckConsistency()` |

#### **Document.CheckGrammar**

开始对指定的文档或区域执行拼写和语法检查。

**语法**

**express.CheckGrammar()**

*express*   一个代表 **Document** 对象的变量。

**说明**

如果文档或区域中包含错误，则该方法显示**“拼写和语法”**对话框，其中**“检查语法”**复选框处于选定状态。该方法应用于文档时，将检查文档中所有可用文字的拼写（如页眉、页脚和文本框）。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例开始对活动文档中的所有文字执行拼写和语法检查。*/ ActiveDocument.CheckGrammar()` |

#### **Document.CheckIn**

该方法将文档从本地计算机返回到服务器，并将本地文档设为只读，使之无法在本地进行编辑。

**语法**

**express.CheckIn(SaveChanges, MakePublic, Comments, VersionType)**

*express*   一个代表 **Document** 对象的变量。

**参数**

| **名称**      | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *SaveChanges* | 可选          | **Boolean**  | 如果该属性值为 True，则将文档保存到服务器位置。默认值为 True。 |
| *MakePublic*  | 可选          | **Boolean**  | 如果该属性值为 True，则允许用户在文档签入后进行文档发布。这将提交文档以进行批准，并最终生成发布给用户的一个文档版本，用户对该文档只具有可读权限（仅当 SaveChanges 为 True 时才应用)。 |
| *Comments*    | 可选          | **Variant**  | 正在签入的文档的修订的批注（仅当 SaveChanges 为 True 时才应用）。 |
| *VersionType* | 可选          | **Variant**  | 指定文档的版本信息。                                         |

**说明**

要利用 WPS 中内置的协作功能，文档必须存储在 Microsoft SharePoint Portal Server 上。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例检查服务器并确定指定的文档是否可以签入。如果可以，则保存并关闭文档，然后将文档签回服务器。*/ function CheckInOut(docCheckIn){   if(Documents.Item(docCheckIn).CanCheckin() == true){     Documents.Item(docCheckIn).CheckIn()     MsgBox(docCheckIn + " has been checked in.")   }   else{     MsgBox("This file cannot be checked in at this time.  Please try again later.")   } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*若要调用 CheckInOut 子程序，请使用下列子程序并将“http://servername/workspace/report.doc”替换为以上“注解”一节中提到的服务器上实际文件的文件名。*/ function CheckDocInOut(){     CheckInOut("http://servername/workspace/report.doc") }` |

#### **Document.CheckInWithVersion**

将文档从本地计算机保存到服务器，并将本地文档设为只读，使之无法在本地进行编辑。

**语法**

**express.CheckInWithVersion(SaveChanges, Comments, MakePublic, VersionType)**

*express*   一个代表 **Document** 对象的变量。

**参数**

| **名称**      | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *SaveChanges* | 可选          | **Boolean**  | 如果值为 True，则将文档保存到服务器位置。默认值为 True。     |
| *Comments*    | 可选          | **Variant**  | 正在签入的文档的修订批注（仅适用于 SaveChanges 设置为 True 的情况）。 |
| *MakePublic*  | 可选          | **Boolean**  | 如果值为 True，则允许用户在文档签入后发布文档。              |
| *VersionType* | 可选          | **Variant**  | 指定文档的版本信息。                                         |

**说明**

如果将 *MakePublic* 参数设置为 **True**，则会提交文档以进行审批，审批过程最终产生的文档版本将发布给对该文档具有只读权限的用户（仅适用于 *SaveChanges* 设置为 **True** 的情况）。

要利用 WPS 中内置的协作功能，文档必须存储在 Microsoft SharePoint Server 上。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例使用 CanCheckin 方法确定文档是否已存储在 Microsoft SharePoint Server 上。如果文档已存储在服务器上，则此示例调用 CheckInWithVersion 方法签入文档以及指定的批注和版本号、将更改保存到服务器位置并提交文档以进行审批。   此示例适用于文档级自定义。  */ function DocumentCheckIn() {     if (Application.ActiveDocument.CanCheckin()) {         Application.ActiveDocument.CheckInWithVersion(true, "My updates.", true, wdCheckInMinorVersion)     }     else {         alert("This document cannot be checked in")     } } ` |

#### **Document.CheckSpelling**

开始对指定的文档或区域进行拼写检查。

**语法**

**express.CheckSpelling(IgnoreUppercase, AlwaysSuggest)**

*express*   一个代表 **Document** 对象的变量。

**参数**

| **名称**          | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *IgnoreUppercase* | 可选          | **Variant**  | 如果忽略大小写，则该属性值为 True。如果省略该参数，则使用 IgnoreUppercase 属性的当前值。 |
| *AlwaysSuggest*   | 可选          | **Variant**  | 如果该属性值为 True，则 WPS 总是给出建议的拼写。如果省略该参数，则使用 SuggestSpellingCorrections 属性的当前值。 |

**说明**

如果该文档或区域中包含错误，则该方法显示**“拼写和语法”**对话框（**“工具”**菜单），其中**“检查语法”**复选框处于清除状态。该方法应用于文档时，将检查文档中所有可用文字的拼写（如页眉、页脚和文本框）。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例检查活动文档中的拼写。*/ Application.ActiveDocument.CheckSpelling()` |

#### **Document.Close**

关闭指定的文档。

**语法**

**express.Close(SaveChanges, OriginalFormat, RouteDocument)**

*express*   一个代表 **Document** 对象的变量。

**参数**

| **名称**         | **必选/可选** | **数据类型**         | **说明**                                                     |
| ---------------- | ------------- | -------------------- | ------------------------------------------------------------ |
| *SaveChanges*    | 可选          | **WdSaveOptions**    | 指定保存文档的操作。可以是下列 WdSaveOptions 常量之一：wdDoNotSaveChanges、wdPromptToSaveChanges 或 wdSaveChanges。 |
| *OriginalFormat* | 可选          | **WdOriginalFormat** | 指定保存文档的格式。可以是下列 WdOriginalFormat 常量之一：wdOriginalDocumentFormat、wdPromptUser 或 wdWordDocument。 |
| *RouteDocument*  | 可选          | **Boolean**          | 如果为 True，则将文档传送给下一个收件人。如果文档没有附加的传送名单，则忽略该参数。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例在关闭活动文档前提示用户保存该文档 */ Application.ActiveDocument.Close(wdPromptToSaveChanges,wdPromptUser)` |

#### **Document.ClosePrintPreview**

将指定文档从打印预览视图切换到以前的视图。

**语法**

**express.ClosePrintPreview()**

*express*   一个代表 **Document** 对象的变量。

**说明**

如果指定文档未处于打印预览视图，将出现错误。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动窗口从打印预览视图切换至普通视图。*/ function test() {     if (Application.ActiveDocument.PrintPreview == true) {         Application.ActiveDocument.ClosePrintPreview()     }     Application.ActiveDocument.ActiveWindow.View.Type = wdNormalView }` |

#### **Document.Compare**

显示修订标记，以表明指定的文档与另一个文档的区别。

**语法**

**express.Compare(Name, AuthorName, CompareTarget, DetectFormatChanges, IgnoreAllComparisonWarnings, AddToRecentFiles, RemovePersonalInformation, RemoveDateAndTime)**

*express*   一个代表 **Document** 对象的变量。

**参数**

| **名称**                      | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------------------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Name*                        | 可选          | **String**   | 与指定文档作比较的文档的名称。                               |
| *AuthorName*                  | 可选          | **Variant**  | 与比较生成的区别相关联的审阅者姓名。如果没有指定，则默认值为修订的文档的作者姓名或字符串“Comparison”（如果没有提供作者信息）。 |
| *CompareTarget*               | 可选          | **Variant**  | 用于比较的目标文档。可以是任意 WdCompareTarget 常量。        |
| *DetectFormatChanges*         | 可选          | **Boolean**  | 如果该属性值为 True（默认值），则比较结果包括格式变化检测。  |
| *IgnoreAllComparisonWarnings* | 可选          | **Variant**  | 如果该属性值为 True，则比较文档，而不通知用户问题。默认值为 False。 |
| *AddToRecentFiles*            | 可选          | **Variant**  | 如果该属性值为 True，则将文档添加到“文件”菜单上最近使用的文件列表中。 |
| *RemovePersonalInformation*   | 可选          | **Boolean**  | 如果该属性值为 True，则从返回的 Document 对象的备注、修订和属性对话框中删除所有用户信息。默认值为 False。 |
| *RemoveDateAndTime*           | 可选          | **Boolean**  | 如果该属性值为 True，则从返回的 Document 对象的跟踪更改中删除日期和时间戳记。默认值为 False。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档与位于 Draft 文件夹中名为“FirstRev.doc”的文档进行比较，并将比较结果区别置于一个新文档中。*/ function CompareDocument(){   Application.ActiveDocument.Compare("C:\\Draft\\FirstRev.doc", null, wdCompareTargetNew) }` |

#### **Document.Compatibility**

如果为 **True**，则启用由 *Type* 参数指定的兼容性选项。兼容性选项将影响 WPS 中文档的显示方式。**Boolean** 类型，可读写。

**语法**

**express.Compatibility(Type)**

*express*   一个代表 **Document** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型**        | **说明**     |
| -------- | ------------- | ------------------- | ------------ |
| *Type*   | 必选          | **WdCompatibility** | 兼容性选项。 |

**说明**

由于您选择或安装的语言支持不同（例如美国英语），上述部分常量可能无法使用。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为活动文档启用“工具”菜单中“选项”对话框“兼容性”选项卡上的“取消硬分页符或分栏符前后的空格”选项。*/ Applicatiion.ActiveDocument.Compatibility(wdSuppressSpBfAfterPgBrk) = true` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在启用或关闭“不为悬挂式缩进添加自动制表位”选项之间进行切换。*/ Application.ActiveDocument.Compatibility(wdNoTabHangIndent) !Application.ActiveDocument.Compatibility(wdNoTabHangIndent)` |

#### **Document.ComputeStatistics**

返回基于指定文档内容的统计值。**Long** 类型。

**语法**

**express.ComputeStatistics(Statistic, IncludeFootnotesAndEndnotes)**

*express*   一个代表 **Document** 对象的变量。

**参数**

| **名称**                      | **必选/可选** | **数据类型**    | **说明**                                                     |
| ----------------------------- | ------------- | --------------- | ------------------------------------------------------------ |
| *Statistic*                   | 必选          | **WdStatistic** | 要计算的统计值。                                             |
| *IncludeFootnotesAndEndnotes* | 可选          | **Variant**     | 如果为 True，则统计时将包括脚注和尾注；如果省略该参数，则默认值为 False。 |

**说明**

由于您选择或安装的语言支持不同（例如美国英语），上述部分常量可能无法使用。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示活动文档的字数（包括脚注）。*/ alert(Application.ActiveDocument.ComputeStatistics(wdStatisticWords, true) + " words")` |

#### **Document.Convert**

将文件转换成最新的文件格式，并启用所有新功能。

**语法**

**express.Convert()**

*express*   一个代表 **Document** 对象的变量。

#### **Document.ConvertAutoHyphens**

将由自动断字功能创建的连字符转换为手动连字符。

**语法**

**express.ConvertAutoHyphens()**

*express*   一个代表 **Document** 对象的变量。

#### **Document.ConvertNumbersToText**

将指定文档中的列表编号和 LISTNUM 域更改为文本。

**语法**

**express.ConvertNumbersToText()**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档中的列表编号和 LISTNUM 域转换为文本。*/ Application.ActiveDocument.ConvertNumbersToText()` |

#### **Document.ConvertVietDoc**

使用非默认的代码页将一篇越南语文档重新转换为 Unicode。

**语法**

**express.ConvertVietDoc(CodePageOrigin)**

*express*   一个代表 **Document** 对象的变量。

**参数**

| **名称**         | **必选/可选** | **数据类型** | **说明**                         |
| ---------------- | ------------- | ------------ | -------------------------------- |
| *CodePageOrigin* | 必选          | **Long**     | 用于对文档进行编码的原始代码页。 |

**说明**

如果要使一篇文档在另一台计算机或平台上具有可浏览性的话，可以使用 **ConvertVietDoc** 方法。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例将活动文档从越南语 ABC 代码页转换为 Unicode。本示例假定活动文档是使用越南语 ABC 代码页进行编码。function ConvertToVietCodePage(){   Application.ActiveDocument.ConvertVietDoc(5) }` |

#### **Document.CopyStylesFromTemplate**

将指定模板中的样式复制到某篇文档。

**语法**

**express.CopyStylesFromTemplate(Template)**

*express*   一个代表 **Document** 对象的变量。

**参数**

| **名称**   | **必选/可选** | **数据类型** | **说明**     |
| ---------- | ------------- | ------------ | ------------ |
| *Template* | 必选          | **String**   | 模板文件名。 |

**说明**

在将某个模板的样式复制到文档时，文档中同名的样式将重新定义为与模板中描述的样式一致。模板的特有样式将复制到文档，文档特有的样式则保持不变。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档的模板的样式复制到文档。*/ Application.ActiveDocument.CopyStylesFromTemplate(ActiveDocument.AttachedTemplate.FullName)` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将模板 Sales96.dot 的样式复制到 Sales.doc。*/ Application.Documents.Item("Sales.doc").CopyStylesFromTemplate ("C:\\WPSOffice\\Templates\\Sales96.dot")` |

#### **Document.CountNumberedItems**

返回指定 **Document** 对象中项目符号项或编号项以及 LISTNUM 域的数目。

**语法**

**express.CountNumberedItems(NumberType, Level)**

*express*   一个代表 **Document** 对象的变量。

**参数**

| **名称**     | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------ | ------------- | ------------ | ------------------------------------------------------------ |
| *NumberType* | 可选          | **Variant**  | 要计数的编号类型。可以是 WdNumberType 常量之一。默认值为 wdNumberAllNumbers。 |
| *Level*      | 可选          | **Variant**  | 与要计数的编号级别相对应的数。如果省略该参数，将计算所有级别的数目。 |

**说明**

将 *NumberType* 指定为 **wdNumberParagraph** 或 **wdNumberAllNumbers**（默认值）时，将计算项目符号项的数目。

有两种编号类型：一种是预设编号 (**wdNumberParagraph**)，这种编号可通过在**“项目符号和编号”**对话框中选择模板来添加到段落中；另一种是 LISTNUM 域 (**wdNumberListNum**)，这种编号允许为每段添加多个编号。

#### **Document.CreateLetterContent**

生成并返回一个 **LetterContent** 对象，该对象基于指定信函内容。**LetterContent** 对象。

**语法**

**express.CreateLetterContent(DateFormat, IncludeHeaderFooter, PageDesign, LetterStyle, Letterhead, LetterheadLocation, LetterheadSize, RecipientName, RecipientAddress, Salutation, SalutationType, RecipientReference, MailingInstructions, AttentionLine, Subject, CCList, ReturnAddress, SenderName, Closing, SenderCompany, SenderJobTitle, SenderInitials, EnclosureNumber, InfoBlock, RecipientCode, RecipientGender, ReturnAddressShortForm, SenderCity, SenderCode, SenderGender, SenderReference)**

*express*   一个代表 **Document** 对象的变量。

**参数**

| **名称**                 | **必选/可选** | **数据类型**             | **说明**                                                     |
| ------------------------ | ------------- | ------------------------ | ------------------------------------------------------------ |
| *DateFormat*             | 必选          | **String**               | 信函的日期。                                                 |
| *IncludeHeaderFooter*    | 必选          | **Boolean**              | 如果该属性值为 True，则页面设计模板中将包括页眉和页脚。      |
| *PageDesign*             | 必选          | **String**               | 附加到文档的模板的名称。                                     |
| *LetterStyle*            | 必选          | **WdLetterStyle**        | 文档版式。                                                   |
| *Letterhead*             | 必选          | **String**               | 如果该属性值为 True，则为预先打印的信头保留空间。            |
| *LetterheadLocation*     | 必选          | **WdLetterheadLocation** | 预先打印信头的位置。                                         |
| *LetterheadSize*         | 必选          | **Single**               | 为预先打印的信头保留的空间（以磅为单位）。                   |
| *RecipientName*          | 必选          | **String**               | 收信人的姓名。                                               |
| *RecipientAddress*       | 必选          | **String**               | 收信人的地址。                                               |
| *Salutation*             | 必选          | **String**               | 信函的称呼文本。                                             |
| *SalutationType*         | 必选          | **WdSalutationType**     | 信函称呼的类型。                                             |
| *RecipientReference*     | 必选          | **String**               | 信函的参考词组（例如，“In reply to”；）。                    |
| *MailingInstructions*    | 必选          | **String**               | 信函邮寄格式（例如，“Certified Mail”）。                     |
| *AttentionLine*          | 必选          | **String**               | 信函注意用语（例如，“Attention”；）。                        |
| *Subject*                | 必选          | **String**               | 指定信函的主题文本。                                         |
| *CCList*                 | 必选          | **String**               | 信函副本收件人的姓名（转寄人）。                             |
| *ReturnAddress*          | 必选          | **String**               | 回信的通讯地址。                                             |
| *SenderName*             | 必选          | **String**               | 发信人姓名。                                                 |
| *Closing*                | 必选          | **String**               | 信函的结束语。                                               |
| *SenderCompany*          | 必选          | **String**               | 写信人所在公司的名称。                                       |
| *SenderJobTitle*         | 必选          | **String**               | 写信人的职务。                                               |
| *SenderInitials*         | 必选          | **String**               | 写信人的姓名缩写。                                           |
| *EnclosureNumber*        | 必选          | **Long**                 | 信函的附件数。                                               |
| *InfoBlock*              | 可选          | **Variant**              | 由于选择或安装的语言支持（如美国英语）不同，该参数可能不可用。 |
| *RecipientCode*          | 可选          | **Variant**              | 由于选择或安装的语言支持（如美国英语）不同，该参数可能不可用。 |
| *RecipientGender*        | 可选          | **Variant**              | 由于选择或安装的语言支持（如美国英语）不同，该参数可能不可用。 |
| *ReturnAddressShortForm* | 可选          | **Variant**              | 由于选择或安装的语言支持（如美国英语）不同，该参数可能不可用。 |
| *SenderCity*             | 可选          | **Variant**              | 由于选择或安装的语言支持（如美国英语）不同，该参数可能不可用。 |
| *SenderCode*             | 可选          | **Variant**              | 由于选择或安装的语言支持（如美国英语）不同，该参数可能不可用。 |
| *SenderGender*           | 可选          | **Variant**              | 由于选择或安装的语言支持（如美国英语）不同，该参数可能不可用。 |
| *SenderReference*        | 可选          | **Variant**              | 由于选择或安装的语言支持（如美国英语）不同，该参数可能不可用。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例实现的功能是：利用 CreateLetterContent 方法在活动文档中新建 LetterContent 对象，然后用 RunLetterWizard 方法使用该对象。*/ function test() {     let myLetter = Application.ActiveDocument.CreateLetterContent("July 31,1996", false, "", wdFullBlock, true, wdLetterTop, InchesToPoints(1.5), "Dave Edson", "436 SE Main St." + "\r" + "Bellevue, WA 98004", "Dear Dave,", +         wdSalutationInformal, "", "", "", "End of year report", "", "", "", "Sincerely yours,", "", "", "", 0)     Application.ActiveDocument.RunLetterWizard(myLetter) }` |

#### **Document.DataForm**

显示**“数据表单”**对话框，在此对话框中可添加、删除或修改记录。

**语法**

**express.DataForm()**

*express*   一个代表 **Document** 对象的变量。

**说明**

可在邮件合并主文档、邮件合并数据源或其他含有以表格单元格或分隔符分开的数据的文档中应用此方法。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*当活动文档为邮件合并文档时，本示例将显示“数据表单”对话框。*/ function test() {     if (Application.ActiveDocument.MailMerge.State != wdNormalDocument) {         Application.ActiveDocument.DataForm()     } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在一个新文档中创建一个表格并显示“数据表单”对话框。*/ function test() {     let aDoc = Application.Documents.Add()     aDoc.Tables.Add(aDoc.Content, 2, 2)     aDoc.Tables.Item(1).Cell(1, 1).Range.Text = "Name"     aDoc.Tables.Item(1).Cell(1, 2).Range.Text = "Age"     aDoc.DataForm() }` |

#### **Document.DeleteAllComments**

删除文档中 **Comments**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Comments/Comments%20.htm#jsObject_Comments)集合内的所有备注。

**语法**

**express.DeleteAllComments()**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 删除当前文档所有批注 */ Application.ActiveDocument.DeleteAllComments()` |

#### **Document.DeleteAllCommentsShown**

删除显示在屏幕上的指定文档的所有修订。

**语法**

**express.DeleteAllCommentsShown()**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 删除当前文档所有显示的批注 */ Application.ActiveDocument.DeleteAllCommentsShown() ` |

#### **Document.DeleteAllEditableRanges**

删除（指定用户或用户组有权修改其权限的）所有区域中的权限。

**语法**

**express.DeleteAllEditableRanges(EditorID)**

*express*   一个代表 **Document** 对象的变量。

**参数**

| **名称**   | **必选/可选** | **数据类型** | **说明**                                                     |
| ---------- | ------------- | ------------ | ------------------------------------------------------------ |
| *EditorID* | 可选          | **Variant**  | 可以是代表用户电子邮件别名（如果在相同的域中）和电子邮件地址的 String 类型值，或者是代表用户组的 WdEditorType 常量。如果省略该参数，则不删除文档中的任何权限。 |

**说明**

还可以使用 **DeleteAll** 方法来删除（指定用户或用户组有权修改其权限的）所有区域中的权限。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例删除当前用户在所有区域中的所有权限。*/ Application.ActiveDocument.DeleteAllEditableRanges(wdEditorCurrent)` |

#### **Document.DeleteAllInkAnnotations**

删除文档中所有手写墨迹注释。

**语法**

**express.DeleteAllInkAnnotations()**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*要处理墨迹注释，必须在平板电脑上运行 WPS 。*/ Application.ActiveDocument.DeleteAllInkAnnotations()` |

#### **Document.DetectLanguage**

分析指定文本，以确定书写文本的语言类型。

**语法**

**express.DetectLanguage()**

*express*   一个代表 **Document** 对象的变量。

**说明**

应用于 **Document** 对象时，**DetectLanguage** 方法将检查文档中所有可用文本（页眉、页脚、文本框等）。如果指定文本包含了某个句子的一部分，则选定内容或区域将扩展到该句的句末。

如果指定文本已应用了 **DetectLanguage** 方法，则 **LanguageDetected** 属性将设置为 **True**。要重新检测指定文本的语言，必须先将 **LanguageDetected** 属性设置为 **False**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例检查活动文档，以确定其所用的语言类型并显示检查结果。*/ function test() {     if (Application.ActiveDocument.LanguageDetected == true) {         alert("This document has already ")         Application.ActiveDocument.LanguageDetected = false         Application.ActiveDocument.DetectLanguage()      }     else {         Application.ActiveDocument.DetectLanguage()     }     if (Application.ActiveDocument.Range.LanguageID == wdEnglishUS) {         alert("This is a U.S. English document.")     }     else {         alert("This is not a U.S. English document.")     } }` |

#### **Document.DowngradeDocument**

将文档降级为 WPS 97-2003 文档格式，以便可以在以前版本的 WPS 中进行编辑。

**语法**

**express.DowngradeDocument()**

*express*   一个代表 **Document** 对象的变量。

#### **Document.EndReview**

终止审阅使用 **SendForReview** 方法发送以进行审阅的文件，或通过电子邮件将文档发送给另一用户而自动进行审阅的文件。

**语法**

**express.EndReview()**

*express*   一个代表 **Document** 对象的变量。

**说明**

当执行 **EndReview** 方法时将显示一笑消息，询问用户是否要终止审阅。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例终止对活动文档的审阅，本示例假定活动文档处于审阅循环中。*/ function EndDocRev(){   Application.ActiveDocument.EndReview() }` |

#### **Document.ExportAsFixedFormat**

将文档保存为 PDF 或 XPS 格式。

**语法**

**express.ExportAsFixedFormat(OutputFileName, ExportFormat, OpenAfterExport, OptimizeFor, Range, From, To, Item, IncludeDocProps, KeepIRM, CreateBookmarks, DocStructureTags, BitmapMissingFonts, UseISO19005_1, FixedFormatExtClassPtr)**

*express*   一个代表 **Document** 对象的变量。

**参数**

| **名称**                 | **必选/可选** | **数据类型**                | **说明**                                                     |
| ------------------------ | ------------- | --------------------------- | ------------------------------------------------------------ |
| *OutputFileName*         | 必选          | **String**                  | 新的 PDF 或 XPS 文件的路径和文件名。                         |
| *ExportFormat*           | 可选          | **WdExportFormat**          | 指定采用 PDF 格式或 XPS 格式。                               |
| *OpenAfterExport*        | 可选          | **Boolean**                 | 导出内容后打开新文件。                                       |
| *OptimizeFor*            | 可选          | **WdExportOptimizeFor**     | 指定是针对屏幕显示还是打印进行优化。                         |
| *Range*                  | 可选          | **WdExportRange**           | 指定导出区域是整个文档、当前页面、文本区域还是当前所选内容。默认情况下导出整个文档。 |
| *From*                   | 可选          | **Long**                    | 如果 Range 参数设置为 wdExportFromTo，则指定起始页码。       |
| *To*                     | 可选          | **Long**                    | 如果 Range 参数设置为 wdExportFromTo，则指定结束页码。       |
| *Item*                   | 可选          | **WdExportItem**            | 指定导出过程是只包括文本还是包括文本和标记。                 |
| *IncludeDocProps*        | 可选          | **Boolean**                 | 指定在最新导出的文件中是否包括文档属性。                     |
| *KeepIRM*                | 可选          | **Boolean**                 | 指定在源文档受 IRM 保护的情况下，是否将 IRM 权限复制到 XPS 文档。默认值为 True。 |
| *CreateBookmarks*        | 可选          | **WdExportCreateBookmarks** | 指定是否导出书签和要导出的书签的类型。                       |
| *DocStructureTags*       | 可选          | **Boolean**                 | 指定是否包含额外数据（如有关内容的流和逻辑组织的信息）以便为屏幕读取器提供帮助。默认值为 True。 |
| *BitmapMissingFonts*     | 可选          | **Boolean**                 | 指定是否包含文本的位图。当字体许可不允许在 PDF 文件中嵌入某一字体时，将此参数设置为 True。如果为 False，则引用该字体，且查看者的计算机会替换适当的字体（如果此创作的字体不可用）。默认值为 True。 |
| *UseISO19005_1*          | 可选          | **Boolean**                 | 指定是否将 PDF 的使用范围限制在按照 ISO 19005-1 标准化的 PDF 子集。如果为 True，生成的文件会更加可靠地自我包含，但由于受到格式的限制，可能会导致该文件较大或显示更多的可视项目。默认值为 False。 |
| *FixedFormatExtClassPtr* | 可选          | **Variant**                 | 指定一个指针以指向一个允许对代码的备用实现进行调用的加载项。代码的备用实现将对应用程序生成的 EMF 和 EMF+ 页面描述进行解释，以生成其自身的 PDF 或 XPS。有关详细信息，请参阅 MSDN 上的“扩展 WPS (2015) 固定格式导出功能”。 |

#### **Document.FitToPages**

适当缩小文本的字体大小，使文档的页数减少。

**语法**

**express.FitToPages()**

*express*   一个代表 **Document** 对象的变量。

**说明**

如果 WPS 无法将文档的页数减少一页，将出现错误。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例试图将活动文档的页数减少一页。*/ function test() {     try {         Application.ActiveDocument.FitToPages()     }     catch (exception) {         alert("Fit to pages failed")     } } ` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例试图将每一打开文档的页数减少一页。*/ function test() {     for (let i = 1; i <= Application.Documents.Count; i++) {         Application.Documents.Item(i).FitToPages()     } }` |

#### **Document.FollowHyperlink**

显示一个缓存文档（如果该文档已下载）。否则，此方法将解析该超链接，下载目标文档，并在相应的应用程序中显示此文档。

**语法**

**express.FollowHyperlink(Address, SubAddress, NewWindow, AddHistory, ExtraInfo, Method, HeaderInfo)**

*express*   一个代表 **Document** 对象的变量。

**参数**

| **名称**     | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------ | ------------- | ------------ | ------------------------------------------------------------ |
| *Address*    | 必选          | **String**   | String 目标文档的地址。                                      |
| *SubAddress* | 可选          | **Variant**  | Variant 目标文档中的位置。默认值为一个空字符串。             |
| *NewWindow*  | 可选          | **Variant**  | 如果该属性值为 True，则在新窗口中显示目标位置。默认值为 False。 |
| *AddHistory* | 可选          | **Variant**  | 如果该属性值为 True，则向当前的历史记录文件夹中加入链接。    |
| *ExtraInfo*  | 可选          | **Variant**  | 指定用于解析超链接的 HTTP 附加信息的字符串或字节数组。例如，您可以使用 ExtraInfo 来指定图像映射的坐标、窗体的内容或 FAT 文件的名称。根据 Method 的值来决定是发布还是附加该字符串。可使用 ExtraInfoRequired 属性确定是否需要额外信息。 |
| *Method*     | 可选          | **Variant**  | 指定 HTTP 附加信息的处理方式。可以是任何 MsoExtraInfoMethod 常量。 |
| *HeaderInfo* | 可选          | **Variant**  | 指定 HTTP 请求的标题信息的字符串。默认值为一个空字符串。可以使用以下语法将几个标题行合并成一个字符串："string1" & vbCr & "string2"。指定的字符串自动转换为 ANSI 字符。请注意，HeaderInfo 参数可能会覆盖默认的 HTTP 标题字段。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例根据指定的 URL 地址在新窗口显示 WPS 主页。*/ Application.ActiveDocument.FollowHyperlink("http://www.wps.cn", null, true, true)` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示名为“Default.htm”的 HTML 文档。/ Application.ActiveDocument.FollowHyperlink("file:C:\\Pages\\Default.htm")` |

#### **Document.ForwardMailer**

您请求就仅在 Macintosh 上使用的 Visual Basic 关键字获得帮助。有关 **Document** 对象的 **ForwardMailer** 方法的信息，请参阅 WPS OfficeMacintosh Edition 所附带的语言参考帮助。

**语法**

**express.ForwardMailer()**

*express*   一个代表 **Document** 对象的变量。

#### **Document.FreezeLayout**

在 Web 视图中，按照文档当前所显示的样子固定文档的版式，以使换行保持不变，且在调整窗口大小时墨迹不会移动。

**语法**

**express.FreezeLayout()**

*express*   一个代表 **Document** 对象的变量。

#### **Document.GetCrossReferenceItems**

返回一个项目数组，根据指定的交叉引用类型可对该数组中的项目进行交叉引用。

**语法**

**express.GetCrossReferenceItems(ReferenceType)**

*express*   一个代表 **Document** 对象的变量。

**参数**

| **名称**        | **必选/可选** | **数据类型** | **说明**                                                     |
| --------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *ReferenceType* | 必选          | **Variant**  | 要在其中插入交叉引用的项目类型。可以是任何 WdReferenceType 常量。 |

**说明**

该方法返回的数组与**“交叉引用”**对话框的**“引用哪一个”**框中所列的项目相对应。该方法返回的值可以用作 **Range** 或 **Selection** 对象的 **InsertCrossReference** 方法的 *ReferenceWhich* 参数的值。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示活动文档中第一个可以进行交叉引用的书签名称。*/ function test() {     if (Application.ActiveDocument.Bookmarks.Count >= 1) {         myBookmarks = Application.ActiveDocument.GetCrossReferenceItems(wdRefTypeBookmark)         alert(myBookmarks(1))     } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例首先用 GetCrossReferenceItems 方法检索可以进行交叉引用的标题的列表，然后在包含标题“Introduction”页中插入一条交叉引用。*/ function test() {     let myHeadings = Application.ActiveDocument.GetCrossReferenceItems(wdRefTypeHeading)     for (let i = 0; i <= myHeadings.length - 1; i++) {         if (myHeadings[i].toLowerCase().indexOf("introduction") != -1) {             Application.Selection.InsertCrossReference(wdRefTypeHeading, wdPageNumber, i + 1)             Application.Selection.InsertParagraphAfter()         }     } }` |

#### **Document.GetLetterContent**

检索指定文档中的信函内容并返回一个 **LetterContent** 对象。

**语法**

**express.GetLetterContent()**

*express*   一个代表 **Document** 对象的变量。

**返回值**

LetterContent

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示活动文档信件中的问候语以及收信人的姓名。*/ alert(Application.ActiveDocument.GetLetterContent().Salutation + Application.ActiveDocument.GetLetterContent().RecipientName)` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例从活动文档获取信函元素，并通过设置“CClist”属性来改变抄送人列表，然后用 SetLetterContent 方法来更新活动文档，以便反映这些更改。*/ function test() {     let myLetterContent = Application.ActiveDocument.GetLetterContent()     myLetterContent.CCList = "J. Burns, L. Scarpaczyk, K. Wong"     myLetterContent.RecipientName = "Amy Anderson"     myLetterContent.RecipientAddress = "123 Main" + "\r" + "Bellevue, WA  98004"     myLetterContent.LetterStyle = wdFullBlock     Application.ActiveDocument.SetLetterContent(myLetterContent) }` |

#### **Document.GetWorkflowTasks**

返回一个 **WorkflowTasks** 集合，该集合表示分配给文档的工作流任务。

**语法**

**express.GetWorkflowTasks()**

*express*   一个代表 **Document** 对象的变量。

**返回值**

WorkflowTasks

#### **Document.GetWorkflowTemplates**

返回一个 **WorkflowTemplates** 集合，该集合表示附加到文档的工作流模板。

**语法**

**express.GetWorkflowTemplates()**

*express*   一个代表 **Document** 对象的变量。

**返回值**

WorkflowTemplates

#### **Document.GoTo**

返回一个 **Range**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)对象，该对象代表指定项（如页、书签或域）的起始位置。

**语法**

**express.GoTo(What, Which, Count, Name)**

*express*   一个代表 **Document** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型**        | **说明**                                                     |
| -------- | ------------- | ------------------- | ------------------------------------------------------------ |
| *What*   | 可选          | **WdGoToItem**      | 指定区域或选定内容要移动到的项的类别。可以是下列 WdGoToItem 常量之一。 |
| *Which*  | 可选          | **WdGoToDirection** | 指定区域或选定内容要移动到的项。可以是下列 WdGoToDirection 常量之一。 |
| *Count*  | 可选          | **Number**          | 文档中的项数。默认值为 1。只有正值是有效的。若要指定一个在该区域或选定内容之前的项，可将 Which 参数指定为 wdGoToPrevious，并指定一个 Count 值。 |
| *Name*   | 可选          | **Variant**         | 如果 What 参数是 wdGoToBookmark、wdGoToComment、wdGoToField 或 wdGoToObject，则此参数指定一个名称。若要指定一个在该区域或选定内容之前的项，可将 Which 参数指定为 wdGoToPrevious，并指定一个 Count 参数值。 |

**说明**

将 **GoTo** 方法与 **wdGoToGrammaticalError**、**wdGoToProofreadingError** 或 **wdGoToSpellingError** 常量一起使用时，返回的 **Range** 对象中包括所有含语法或拼写错误的文本。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 切换到第一个作者名为"authorname"的备注开始位置 */ Application.ActiveDocument.GoTo(wdGoToComment,wdGoToFirst,1,"authorname").Select()  /* 选中当前文档的第一个脚注引用处 */ function test() {     if (Application.ActiveDocument.Footnotes.Count >= 1) {         let r1 = Application.ActiveDocument.GoTo(wdGoToFootnote, wdGoToFirst);         r1.Select();     } }` |

#### **Document.LockServerFile**

在服务器上锁定文件，以避免任何其他人进行编辑。

**语法**

**express.LockServerFile()**

*express*   一个代表 **Document** 对象的变量。

**说明**

此方法仅适用于在运行 SharePoint Server 2007 的服务器上存储的文档。

#### **Document.MakeCompatibilityDefault**

兼容性选项位于**“选项”**对话框的**“兼容性”**选项卡上，它们是新文档的默认设置。

**语法**

**express.MakeCompatibilityDefault()**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例首先为活动文档设置兼容性选项，然后将当前选项设置为默认设置。*/ function test() {     Application.ActiveDocument.Compatibility(wdSuppressSpBfAfterPgBrk)     Application.ActiveDocument.Compatibility(wdExpandShiftReturn)      Application.ActiveDocument.Compatibility(wdUsePrinterMetrics)     Application.ActiveDocument.Compatibility(wdNoLeading)     Application.ActiveDocument.MakeCompatibilityDefault() }  ` |

#### **Document.ManualHyphenation**

启动文档的手动断字设置，每次一行。

**语法**

**express.ManualHyphenation()**

*express*   一个代表 **Document** 对象的变量。

**说明**

使用 **ManualHyphenation** 方法时， WPS 提示用户接受还是拒绝建议的断字。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例启动活动文档的手动断字操作。*/ ActiveDocument.ManualHyphenation()` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例先设置断字选项，然后对 MyDoc.doc 进行手动断字。*/ function test() {     let rng = Application.Documents.Item("MyDoc.doc")     rng.HyphenationZone = InchesToPoints(0.25)     rng.HyphenateCaps = false     rng.ManualHyphenation() }` |

#### **Document.Merge**

将文档中用修订标记标识的修改合并到另一篇文档。

**语法**

**express.Merge(Name, MergeTarget, DetectFormatChanges, UseFormattingFrom, AddToRecentFiles)**

*express*   一个代表 **Document** 对象的变量。

**参数**

| **名称**              | **必选/可选** | **数据类型**            | **说明**                                             |
| --------------------- | ------------- | ----------------------- | ---------------------------------------------------- |
| *Name*                | 必选          | **String**              | 要合并的文档的路径和文件名。                         |
| *MergeTarget*         | 可选          | **WdMergeTarget**       | 指定放置最终合并的内容的位置。                       |
| *DetectFormatChanges* | 可选          | **Boolean**             | 指定是否标记格式差别。                               |
| *UseFormattingFrom*   | 可选          | **WdUseFormattingFrom** | 指定将哪篇文档用于设置合并文档的格式。               |
| *AddToRecentFiles*    | 可选          | **Boolean**             | 指定是否将 Name 参数中的文档添加到最近的文件列表中。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将在 Sales2.doc（活动文档）中所做的修改合并到 Sales1.doc。*/ function test() {     if (Application.ActiveDocument.Name.indexOf("sales2.doc")) {         Application.ActiveDocument.Merge("C:\\Docs\\Sales1.doc")     } }` |

#### **Document.Post**

将指定的文档传送到 Microsoft Exchange 中的公用文件夹。

**语法**

**express.Post()**

*express*   一个代表 **Document** 对象的变量。

**说明**

该方法显示**“发送到 Exchange 文件夹”**对话框，以便选择一个文件夹。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示“发送到 Exchange 文件夹”对话框，以便将活动文档发送到公用文件夹。*/ Application.ActiveDocument.Post()` |

#### **Document.PresentIt**

打开 WPP，并加载指定的 WPS 文档。

**语法**

**express.PresentIt()**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例向 WPP 中发送名为“MyPresentation.doc”的文档的副本。*/ Application.Documents.Item("MyPresentation.doc").PresentIt()` |

#### **Document.PrintOut**

打印指定文档的全部或部分内容。

**语法**

**express.PrintOut(Background, Append, Range, OutputFileName, From, To, Item, Copies, Pages, PageType, PrintToFile, Collate, FileName, ActivePrinterMacGX, ManualDuplexPrint, PrintZoomColumn, PrintZoomRow, PrintZoomPaperWidth, PrintZoomPaperHeight)**

*express*   一个代表 **Document** 对象的变量。

**参数**

| **名称**               | **必选/可选** | **数据类型** | **说明**                                                     |
| ---------------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Background*           | 可选          | **Variant**  | 如果将该属性设置为 True，则 WPS 在打印文档时继续运行宏。     |
| *Append*               | 可选          | **Variant**  | 如果将该属性设置为 True，则将指定文档附加到 OutputFileName 参数所指定的文件名。如果将该属性设置为 False，则覆盖 OutputFileName 参数所指定文件的内容。 |
| *Range*                | 可选          | **Variant**  | 页码范围。可以是任意 WdPrintOutRange 常量。                  |
| *OutputFileName*       | 可选          | **Variant**  | 如果 PrintToFile 为 True，则该参数指定输出文件的路径和文件名。 |
| *From*                 | 可选          | **Variant**  | 如果将 Range 设置为 wdPrintFromTo，则该参数指定起始页码。    |
| *To*                   | 可选          | **Variant**  | 如果将 Range 设置为 wdPrintFromTo，则该参数指定结束页码。    |
| *Item*                 | 可选          | **Variant**  | 要打印的项目。可以是任意 WdPrintOutItem 常量。               |
| *Copies*               | 可选          | **Variant**  | 要打印的份数。                                               |
| *Pages*                | 可选          | **Variant**  | 要打印的页码和页码范围，中间用逗号分开。例如，“2, 6-10”表示打印第 2 页以及第 6 至第 10 页。 |
| *PageType*             | 可选          | **Variant**  | 要打印的页面类型。可以是任意 WdPrintOutPages 常量。          |
| *PrintToFile*          | 可选          | **Variant**  | 如果该属性值为 True，则将打印指令发送到文件。请确保使用 OutputFileName 指定文件名。 |
| *Collate*              | 可选          | **Variant**  | 在打印文档的多份副本时，如果该属性值为 True，则完成打印所有页面后再打印下一份副本。 |
| *FileName*             | 可选          | **Variant**  | 要打印的文档的路径和文件名。如果省略该参数， WPS 将打印活动文档（仅应用于 Application 对象）。 |
| *ActivePrinterMacGX*   | 可选          | **Variant**  | 该参数仅适用于 WPS OfficeMacintosh Edition。有关该参数的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。 |
| *ManualDuplexPrint*    | 可选          | **Variant**  | 如果该属性值为 True，则在无双面打印组件的打印机上打印双面文档。如果该属性值为 True，则忽略 PrintBackground 和 PrintReverse 属性。使用 PrintOddPagesInAscendingOrder 和 PrintEvenPagesInAscendingOrder 属性可在手动双面打印时控制输出。由于选择或安装的语言支持（如美国英语）不同，该参数可能不可用。 |
| *PrintZoomColumn*      | 可选          | **Variant**  | 表示 WPS 在一页纸上水平放置的页数。可以是 1、2、3 或 4 页。与 PrintZoomRow 参数一同使用可在单张纸上打印多页文档。 |
| *PrintZoomRow*         | 可选          | **Variant**  | 表示 WPS 在一页纸上垂直放置的页数。可以是 1、2 或 4 页。与 PrintZoomColumn 参数一同使用可在单张纸上打印多页文档。 |
| *PrintZoomPaperWidth*  | 可选          | **Variant**  | WPS 将打印页面缩放到的宽度，以缇为单位（20 缇 = 1 磅；72 磅 = 1 英寸）。 |
| *PrintZoomPaperHeight* | 可选          | **Variant**  | WPS 将打印页面缩放到的高度，以缇为单位（20 缇 = 1 磅；72 磅 = 1 英寸）。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例打印活动文档的当前页面 */ Application.ActiveDocument.PrintOut(null, null, wdPrintCurrentPage)  /* 打印当前活动文档的第1-3页 */ Application.ActiveDocument.PrintOut(null,null,wdPrintFromTo,"","1","3")` |

#### **Document.PrintPreview**

将视图切换到打印预览。

**语法**

**express.PrintPreview()**

*express*   一个代表 **Document** 对象的变量。

**说明**

除了使用 **PrintPreview** 方法外，还可以将**PrintPreview**属性设置为 **True** 或 **False**，前者切换到打印预览，后者从打印预览切回。还可以通过将 **View** 对象的 **Type** 属性设置为 **wdPrintPreview**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdViewType%20%E6%9E%9A%E4%B8%BE.html)来更改视图。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 切换到打印预览模式 */ Application.ActiveDocument.PrintPreview() ` |

#### **Document.Protect**

帮助保护指定的文档不被更改。保护文档后，用户只能进行有限的更改，例如添加注释，进行修订或填写表格。

**语法**

**express.Protect(Type, NoReset, Password, UseIRM, EnforceStyleLock)**

*express*   一个代表 **Document** 对象的变量。

**参数**

| **名称**           | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------------ | ------------- | ------------ | ------------------------------------------------------------ |
| *Type*             | 必选          | **Variant**  | 指定文档的保护类型。 **WdProtectionType**。                  |
| *NoReset*          | 可选          | **Variant**  | False将表单字段重置为其默认值。如果指定的文档受保护，则为true以保留当前表单字段值。如果Type不是wdAllowOnlyFormFields，则将忽略NoReset参数。 |
| *Password*         | 可选          | **Variant**  | 从指定文档中删除保护所需的密码。 （请参阅下面的备注。）      |
| *UseIRM*           | 可选          | **Variant**  | 指定在保护文档免受更改时是否使用信息权限管理（IRM）。        |
| *EnforceStyleLock* | 可选          | **Variant**  | 指定是否在受保护的文档中实施格式限制。                       |

#### **Document.Range**

通过使用指定的开始和结束字符位置返回一个 **Range**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)对象。

**语法**

**express.Range()**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例为活动文档中的前 10 个字符设置加粗格式 */ Application.ActiveDocument.Range(0, 10).Bold = true` |

#### **Document.Redo**

重复最后一次撤消的操作（**Undo** 方法的逆操作）。如果重复操作成功，则返回 **True**。

**语法**

**express.Redo()**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 将当前文档前两个字符加粗然后倾斜，然后撤销倾斜，撤销加粗，重做加粗，那么文字最终加粗，并没倾斜 */ function test() { 	let doc = Application.ActiveDocument; 	let rng = doc.Range(0,2); 	rng.Bold = true; 	rng.Italic = true; 	doc.Undo(2); 	doc.Redo(1); }` |

#### **Document.RejectAllRevisions**

拒绝指定文档的所有修订。

**语法**

**express.RejectAllRevisions()**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 拒绝当前文档的所有修订 */ Application.ActiveDocument.RejectAllRevisions()` |

#### **Document.RejectAllRevisionsShown**

拒绝显示在屏幕上的文档的所有修订。

**语法**

**express.RejectAllRevisionsShown()**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 隐藏所有"authorname"做的修订，拒绝所有其他显示的修订 */ function test() { 	let rev = Application.ActiveWindow.View.Reviewers; 	rev.Item("authorname").Visible = false; 	Application.ActiveDocument.RejectAllRevisionsShown(); }` |

#### **Document.Reload**

通过解析指向缓冲文档的超链接再下载该文档来重新加载该文档。

**语法**

**express.Reload()**

*express*   一个代表 **Document** 对象的变量。

**说明**

该方法将同步重新加载文档，也就是说，在真正开始重新加载文档之前，可以执行此过程中 **Reload** 方法后紧跟的语句。因此，在宏中使用此方法将得到意想不到的结果。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例打开并重新加载至本地内部网上地址为“main”的超链接。*/ function test() {     Application.ActiveDocument.FollowHyperlink("http://main")     Application.ActiveDocument.Reload() }` |

#### **Document.ReloadAs**

使用指定的编码重新加载一篇基于 HTML 文档的文档。

**语法**

**express.ReloadAs(Encoding)**

*express*   一个代表 **Document** 对象的变量。

**参数**

| **名称**   | **必选/可选** | **数据类型**    | **说明**                       |
| ---------- | ------------- | --------------- | ------------------------------ |
| *Encoding* | 必选          | **MsoEncoding** | 指定重新加载文档时使用的编码。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例用西里尔文编码重新加载当前文档。*/ Application.ActiveDocument.ReloadAs(msoEncodingCyrillic)` |

#### **Document.RemoveDocumentInformation**

从文档中删除敏感性信息、属性、批注和其他元数据。

**语法**

**express.RemoveDocumentInformation(RemoveDocInfoType)**

*express*   一个代表 **Document** 对象的变量。

**参数**

| **名称**            | **必选/可选** | **数据类型**            | **说明**           |
| ------------------- | ------------- | ----------------------- | ------------------ |
| *RemoveDocInfoType* | 必选          | **WdRemoveDocInfoType** | 指定要删除的内容。 |

#### **Document.RemoveLockedStyles**

当文档中应用了格式设置限制时，清除锁定样式的文档。

**语法**

**express.RemoveLockedStyles()**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例清除活动文档中的锁定样式。*/ Application.ActiveDocument.RemoveLockedStyles()` |

#### **Document.RemoveNumbers**

删除指定文档中的编号或项目符号。

**语法**

**express.RemoveNumbers(NumberType)**

*express*   一个代表 **Document** 对象的变量。

**参数**

| **名称**     | **必选/可选** | **数据类型**     | **说明**           |
| ------------ | ------------- | ---------------- | ------------------ |
| *NumberType* | 可选          | **WdNumberType** | 要删除的编号类型。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例从头删除活动文档中已编号段落的编号。*/ Application.ActiveDocument.RemoveNumbers(wdNumberParagraph)` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例删除 MyDocument.doc 中第三个列表的项目符号或编号。*/ function test() {     if (Application.Documents.Item("MyDocument.doc").Lists.Count >= 3) {         Application.Documents.Item("MyDocument.doc").Lists.Item(3).RemoveNumbers()     } } ` |

#### **Document.RemoveTheme**

删除当前文档中的活动主题。

**语法**

**express.RemoveTheme()**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将删除当前文档中的活动主题。*/ Application.ActiveDocument.RemoveTheme()` |

#### **Document.Repaginate**

对整篇文档重新分页。

**语法**

**express.Repaginate()**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果活动文档在上次保存后进行了更改，本示例将对其重新分页。*/ function test() {     if (Application.ActiveDocument.Saved == false) {         Application.ActiveDocument.Repaginate()     } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将对所有打开文档重新分页。*/ function test() {     for (let i = 1; i <= Application.Documents.Count; i++) {         Application.Documents.Item(i).Repaginate()     } }` |

#### **Document.Reply**

打开一封新的电子邮件（**“收件人”**行已含发件人地址）以答复活动邮件。

**语法**

**express.Reply()**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例打开一封新的电子邮件以答复活动邮件。*/ Application.MailMessage.Reply()` |

#### **Document.ReplyAll**

打开一封新的电子邮件（**“收件人”**和**“抄送”**行已含发件人和所有其他收件人的地址）以答复活动邮件。

**语法**

**express.ReplyAll()**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例打开一封新的电子邮件以答复活动邮件。*/ Application.MailMessage.ReplyAll()` |

#### **Document.ReplyWithChanges**

为文档（该文档以被发送出去进行审阅）的作者发送电子邮件，通知他们审阅者已经审阅完文档。

**语法**

**express.ReplyWithChanges(ShowMessage)**

*express*   一个代表 **Document** 对象的变量。

**参数**

| **名称**      | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *ShowMessage* | 可选          | **Variant**  | 如果该属性值为 True，则发送前先显示邮件。如果属性值为 False，则自动发送邮件而不先显示它。默认值为 True。 |

**说明**

可用 **SendForReview** 方法开始对文档进行合作审阅。如果对不属于合作审阅阶段的文档执行 **ReplyWithChanges** 方法，则 WPS 会显示错误消息。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例发送一条信息，通知作者审阅者已经完成审阅（而不先向审阅者显示该电子邮件）。本示例假定当前文档是合作审阅循环中的一部分。*/ function ReplyMsg() {     Application.ActiveDocument.ReplyWithChanges(false) }` |

#### **Document.ResetFormFields**

清除文档中的所有窗体域，准备再次填充窗体。

**语法**

**express.ResetFormFields()**

*express*   一个代表 **Document** 对象的变量。

**说明**

当文档中的域没有锁定时，可用 **ResetFormFields** 方法清除域。要在文档中的域处于锁定状态时清除域，请使用 **Protect** 方法。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例清除活动文档中的域。本示例假定活动文档包含窗体域。*/ function ClearFormFields() {     Application.ActiveDocument.ResetFormFields() }` |

#### **Document.RunAutoMacro**

运行存储于指定文档的自动宏。如果没有自动宏，则不做任何操作。

**语法**

**express.RunAutoMacro(Which)**

*express*   一个代表 **Document** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型**     | **说明**         |
| -------- | ------------- | ---------------- | ---------------- |
| *Which*  | 必选          | **WdAutoMacros** | 要运行的自动宏。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例运行活动文档中的 AutoOpen 宏。*/ Application.ActiveDocument.RunAutoMacro(wdAutoOpen)` |

#### **Document.RunLetterWizard**

对指定的文档运行英文信函向导。

**语法**

**express.RunLetterWizard(LetterContent, WizardMode)**

*express*   一个代表 **Document** 对象的变量。

**参数**

| **名称**        | **必选/可选** | **数据类型** | **说明**                                                     |
| --------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *LetterContent* | 可选          | **Variant**  | 一个 LetterContent 对象。LetterContent 对象中的任何填充属性在“英文信函向导”对话框中显示为预填充元素。如果省略该参数，则 GetLetterContent 方法将自动用于从指定文档获取 LetterContent 对象。 |
| *WizardMode*    | 可选          | **Variant**  | Variant 如果该属性值为 True，则将“英文信函向导”对话框显示为一系列的步骤，并且提供“下一步”、“上一步”、“完成”按钮。如果该属性值为 False，则显示“英文信函向导”对话框就像从“工具”菜单上打开它一样（包含一个“确定”按钮和“取消”按钮的属性对话框）。默认值为 True。 |

**说明**

使用 **CreateLetterContent** 方法返回一个 **LetterContent** 对象，该对象给出各种信函内容属性。使用 **GetLetterContent** 方法返回一个基于指定文档内容的 **LetterContent** 对象。可以使用返回的 **LetterContent** 对象和 **RunLetterWizard** 方法来预先设置**“英文信函向导”**对话框中的元素。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例创建一个新的 LetterContent 对象并设置该对象的若干属性，再使用 RunLetterWizard 方法运行信函向导。*/ function test() {     let myContent = Application.ActiveDocument.GetLetterContent()     myContent.Salutation = "Hello"     myContent.SalutationType = wdSalutationOther     myConent.SenderName = Application.UserName     myContent.SenderInitials = Application.UserInitials     Application.Documents.Add().RunLetterWizard(myContent, true) }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的示例使用 CreateLetterContent 方法在活动文档中创建一个新的 LetterContent 对象，再使用该对象的 RunLetterWizard 方法。*/ function test() {     let myLetter = Application.ActiveDocument.CreateLetterContent("July 31, 1999", false,         +"C:\\WPSOffice\\Templates" + "\\Letters + Faxes\\Contemporary Letter.dot", wdFullBlock, true, wdLetterTop, InchesToPoints(1.5),         +"Dave Edson", "436 SE Main St." + "\r" + "Bellevue, WA 98004", "Dear Dave,", wdSalutationInformal,         + "", "", "", "End of year report", "", "", "", "Sincerely yours,", "", "", "", 0)         Application.ActiveDocument.RunLetterWizard(myLetter) } ` |

#### **Document.Save**

保存当前文档。

**语法**

**express.Save(NoPrompt, OriginalFormat)**

*express*   一个代表 **Document** 对象的变量。

**参数**

| **名称**         | **必选/可选** | **数据类型** | **说明**                                                     |
| ---------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *NoPrompt*       | 可选          | **Variant**  | 如果该属性值为 True，则 WPS 将自动保存所有文档。如果该属性值为 False， WPS 将提示用户保存自上次保存后已修改过的每篇文档。 |
| *OriginalFormat* | 可选          | **Variant**  | 指定文档的保存方式。可以是 WdOriginalFormat 常量之一。       |

**说明**

如果文档以前未保存过，则**“另存为”**对话框将提示用户键入文件名。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 如果活动文档在上次保存后进行了更改，本示例将保存活动文档 */ function test() { 	if(Application.ActiveDocument.Saved == false) {     	  Application.ActiveDocument.Save() 	} }` |

#### **Document.SaveAs2**

使用新的名称或格式保存指定的文档。此方法的一些参数与**“另存为”**对话框（**“文件”**选项卡）中的选项相对应。

**语法**

**express.SaveAs2(FileName, FileFormat, LockComments, Password, AddToRecentFiles, WritePassword, ReadOnlyRecommended, EmbedTrueTypeFonts, SaveNativePictureFormat, SaveFormsData, SaveAsAOCELetter, Encoding, InsertLineBreaks, AllowSubstitutions, LineEnding, AddBiDiMarks, CompatibilityMode)**

*express*   一个代表 **Document** 对象的变量。

**参数**

| **名称**                  | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *FileName*                | 可选          | **Variant**  | 文档的名称。默认值为当前文件夹和文件名。如果从未保存过文档，将使用默认名称（例如，Doc1.doc）。如果已经存在具有指定文件名的文档，将覆盖该文档，而不提示用户。 |
| *FileFormat*              | 可选          | **Variant**  | 文档的保存格式。可以是任何 WdSaveFormat 常量。若要以另一种格式保存文档，请为 FileConverter 对象的 SaveFormat 属性指定适当的值。 |
| *LockComments*            | 可选          | **Variant**  | 如果该属性值为 True，则锁定文档备注。默认值为 False。        |
| *Password*                | 可选          | **Variant**  | 用于打开文档的密码字符串（请参阅下面的“说明”。）             |
| *AddToRecentFiles*        | 可选          | **Variant**  | 如果该属性值为 True，则将文档添加到“文件”菜单上最近使用的文件列表中。默认值为 True。 |
| *WritePassword*           | 可选          | **Variant**  | 用于将更改保存到文档的密码字符串（请参阅下面的“说明”。）     |
| *ReadOnlyRecommended*     | 可选          | **Variant**  | 如果该属性值为 True，则每次打开文档时，WPS 都将建议采用只读方式。默认值为 False。 |
| *EmbedTrueTypeFonts*      | 可选          | **Variant**  | 如果该属性值为 True，则 TrueType 字体随文档一起保存。如果省略，则 EmbedTrueTypeFonts 参数将假定为 EmbedTrueTypeFonts 属性的值。 |
| *SaveNativePictureFormat* | 可选          | **Variant**  | 如果该属性值为 True，则对于从其他系统平台（如 Macintosh）导入的图形，只保存其 Microsoft Windows 版本。 |
| *SaveFormsData*           | 可选          | **Variant**  | 如果该属性值为 True，则将用户在窗体中输入的数据保存为记录。  |
| *SaveAsAOCELetter*        | 可选          | **Variant**  | 如果文档包含附加的邮件，当该属性值为 True 时，将文档存为 AOCE 信函（同时保存邮件）。 |
| *Encoding*                | 可选          | **Variant**  | 用于另存为编码文本文件的文档的代码页或字符集。默认为系统代码页。不能将所有 MsoEncoding 常量用于该参数。 |
| *InsertLineBreaks*        | 可选          | **Variant**  | 在将文档保存为文本文件时，如果该属性值为 True，则在每一行文本后插入换行符。 |
| *AllowSubstitutions*      | 可选          | **Variant**  | 将文档保存为文本文件时，如果该属性值为 True，则 WPS 可以将一些符号替换为相似的文本。例如，将版权符号显示为 (c)。默认值为 False。 |
| *LineEnding*              | 可选          | **Variant**  | WPS 在另存为文本文件的文档中标记换行符和分段符的方式。可为下列 WdLineEndingType 常量之一：wdCRLF（默认）或 wdCROnly。 |
| *AddBiDiMarks*            | 可选          | **Variant**  | 如果该属性值为 True，则为输出文件添加控制符，以保留原始文档中文本的双向版式。 |
| *CompatibilityMode*       | 可选          | **Variant**  | 打开文档时， WPS 2015 使用的兼容模式。WdCompatibilityMode 常量。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 下面的代码示例将活动文档另存为 Test.rtf，格式为 RTF*/ Application.ActiveDocument.SaveAs2("Text.rtf", wdFormatRTF)  /* 将文档保存为带密码 */ function test(strPWD) {     Application.ActiveDocument.SaveAs2("", "", "", "", "", strPWD) }` |

#### **Document.SaveAsQuickStyleSet**

保存当前所用的快速样式组。

**语法**

**express.SaveAsQuickStyleSet(FileName)**

*express*   一个代表 **Document** 对象的变量。

**参数**

| **名称**   | **必选/可选** | **数据类型** | **说明**                       |
| ---------- | ------------- | ------------ | ------------------------------ |
| *FileName* | 必选          | **String**   | 快速样式集文件的路径和文件名。 |

#### **Document.Select**

选择指定文档的内容。

**语法**

**express.Select()**

*express*   一个代表 **Document** 对象的变量。

**说明**

使用该方法后，可用 [Selection](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Selection/Selection%20.htm#jsObject_Selection) 属性来处理选定项目。有关详细信息，请参阅 处理 [Selection ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Selection/Selection%20.htm#jsObject_Selection)对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 选中当前文档所有文本 */ Application.ActiveDocument.Select()` |

#### **Document.SelectAllEditableRanges**

选择指定用户或用户组有权修改的所有区域。

**语法**

**express.SelectAllEditableRanges(EditorID)**

*express*   一个代表 **Document** 对象的变量。

**参数**

| **名称**   | **必选/可选** | **数据类型** | **说明**                                                     |
| ---------- | ------------- | ------------ | ------------------------------------------------------------ |
| *EditorID* | 必选          | **Variant**  | 可以是代表用户电子邮件别名（如果在相同的域中）和电子邮件地址的 String 类型值，或者是代表用户组的 WdEditorType 常量。如果省略，则只选择所有用户有权修改的区域。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例选定当前用户有权修改的所有区域。*/ Application.ActiveDocument.SelectAllEditableRanges(wdEditorCurrent)` |

#### **Document.SelectContentControlsByTag**

返回一个 **ContentControls** 集合，该集合代表文档中具有 *Tag* 参数中指定标记值的所有内容控件。只读。

**语法**

**express.SelectContentControlsByTag(Tag)**

*express*   一个代表 **Document** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                   |
| -------- | ------------- | ------------ | -------------------------- |
| *Tag*    | 必选          | **String**   | 要返回的内容控件的标记值。 |

**返回值**

ContentControls

#### **Document.SelectContentControlsByTitle**

返回一个 **ContentControls** 集合，该集合代表文档中具有 *Title* 参数中指定标题的所有内容控件。只读。

**语法**

**express.SelectContentControlsByTitle()**

*express*   一个代表 **Document** 对象的变量。

**返回值**

ContentControls

#### **Document.SelectLinkedControls**

返回一个 **ContentControls** 集合，该集合代表文档中根据 Node 参数的指定，链接到文档 XML 数据存储区中特定自定义 XML 节点的所有内容控件。只读。

**语法**

**express.SelectLinkedControls(Node)**

*express*   一个代表 **Document** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型**      | **说明**                                      |
| -------- | ------------- | ----------------- | --------------------------------------------- |
| *Node*   | 必选          | **CustomXMLNode** | 与内容控件链接的文档数据存储区中的 XML 节点。 |

**返回值**

ContentControls

#### **Document.SelectNodes**

返回一个 **XMLNodes** 集合，代表所有与 *XPath* 参数相匹配的节点并按其在文档或区域中显示的顺序排列。

**语法**

**express.SelectNodes(XPath, PrefixMapping, FastSearchSkippingTextNodes)**

*express*   一个代表 **Document** 对象的变量。

**参数**

| **名称**                      | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------------------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *XPath*                       | 必选          | **String**   | 有效的 XPath 字符串。有关 XPath 的详细信息，请参阅 Microsoft Developer Network (MSDN) 网站上的 XPath 参考文档。 |
| *PrefixMapping*               | 可选          | **Variant**  | 提供用于搜索的架构中的前缀。如果 XPath 参数使用名称来搜索元素，则可使用 PrefixMapping 参数。 |
| *FastSearchSkippingTextNodes* | 可选          | **Boolean**  | 如果该参数值为 True，则搜索指定节点时跳过所有文本节点。如果该参数值为 False，则搜索时包含文本节点。默认值为 False。 |

**返回值**

XMLNodes

**说明**

将 *FastSearchSkippingTextNodes* 参数设置为 **True** 会降低性能，因为 WPS 将根据节点中包含的文本对文档中所有节点进行搜索。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例返回活动文档中所有 book 元素的集合。*/ function test() {     let strElement     let strPrefix      strElement = "/x:catalog/x:book"     strPrefix = "xmlns:x=\"\"" + ActiveDocument.XMLSchemaReferences.Item(1).NamespaceURI + "\"\""      let objElements = Application.ActiveDocument.SelectNodes(strElement, strPrefix) }` |

#### **Document.SelectSingleNode**

返回一个 **XMLNode** 对象，该对象代表画布与指定文档中 *XPath* 参数相匹配的第一个节点。

**语法**

**express.SelectSingleNode(XPath, PrefixMapping, FastSearchSkippingTextNodes)**

*express*   一个代表 **Document** 对象的变量。

**参数**

| **名称**                      | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------------------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *XPath*                       | 必选          | **String**   | 有效的 XPath 字符串。有关 XPath 的详细信息，请参阅 Microsoft Developer Network (MSDN) 网站上的 XPath 参考文档。 |
| *PrefixMapping*               | 可选          | **Variant**  | 提供用于搜索的架构中的前缀。如果 XPath 参数使用名称来搜索元素，则可使用 PrefixMapping 参数。 |
| *FastSearchSkippingTextNodes* | 可选          | **Boolean**  | 如果该参数值为 True，则搜索指定节点时跳过所有文本节点。如果该参数值为 False，则搜索时包含文本节点。默认值为 True。 |

**返回值**

XMLNode

**说明**

将 *FastSearchSkippingTextNodes* 参数设置为 **False** 会降低性能，因为 WPS 将根据节点中包含的文本对文档中所有节点进行搜索。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例返回在活动文档中找到的第一个 title 元素，该元素是 book 元素的子元素。*/ function test() {     let strElement     let strPrefix      strElement = "/x:catalog/x:book"     strPrefix = "xmlns:x=\"\"" + ActiveDocument.XMLSchemaReferences.Item(1).NamespaceURI + "\"\""      let objElements = ActiveDocument.SelectSingleNode(strElement, strPrefix) }` |

#### **Document.SelectUnlinkedControls**

返回一个 **ContentControls** 集合，该集合代表文档中没有链接到文档 XML 数据存储区中 XML 节点的所有内容控件。只读。

**语法**

**express.SelectUnlinkedControls(Stream)**

*express*   一个代表 **Document** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型**      | **说明**                                                     |
| -------- | ------------- | ----------------- | ------------------------------------------------------------ |
| *Stream* | 可选          | **CustomXMLPart** | 自定义 XML 部件引用。设置此参数会对返回的内容控件进行筛选，以只包括那些在其 XMLMapping 定义中引用了此 CustomXMLPart 的内容控件。 |

**返回值**

ContentControls

#### **Document.SendFax**

将指定的文档作为传真发送，且无需用户交互。

**语法**

**express.SendFax(Address, Subject)**

*express*   一个代表 **Document** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型** | **说明**                           |
| --------- | ------------- | ------------ | ---------------------------------- |
| *Address* | 必选          | **String**   | 收件人的传真号码。                 |
| *Subject* | 可选          | **Variant**  | 主题行的文字。字符数不可超过 255。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档作为传真发送。*/ Application.ActiveDocument.SendFax("12065551234",  "Important Fax")` |

#### **Document.SendFaxOverInternet**

将文档发送给传真服务提供商，该提供商将该文档传真给一个或多个指定收件人。

**语法**

**express.SendFaxOverInternet(Recipients, Subject, ShowMessage)**

*express*   一个代表 **Document** 对象的变量。

**参数**

| **名称**      | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Recipients*  | 可选          | **Variant**  | 一个代表传真收件人的传真号码和电子邮件地址的 String 类型值。如果有多个收件人，则用分号隔开。 |
| *Subject*     | 可选          | **Variant**  | 一个代表传真文档的主题行的String 类型值。                    |
| *ShowMessage* | 可选          | **Variant**  | 如果该属性值为 True，则发送传真之前显示传真邮件。如果该属性值为 False，则发送传真而不显示传真邮件。 |

**说明**

使用 **SendFaxOverInternet** 方法需要在用户计算机上启用传真服务。如果未启用传真服务，则 **SendFaxOverInternet** 方法将会导致运行时错误。

可使用两种格式在 *Recipients* 参数中指定传真号码：一种是 recipientsfaxnumber@usersfaxprovider，另一种是 recipientsname@recipientsfaxnumber。可通过使用以下注册表路径访问用户的传真提供商信息：

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `HKEY_CURRENT_USER\Software\Microsoft\Office\11.0\Common\Services\Fax` |

使用此注册表位置的 FaxAddress 键值可确定用于某个用户的格式。如果此注册表项不存在，则没有可用的传真服务。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将传真发送给传真服务提供商，该提供商将邮件传真给收件人。*/ Application.ActiveDocument.SendFaxOverInternet("14255550101@consolidatedmessenger.com", "For your review", true)    ` |

#### **Document.SendForReview**

将文档通过电子邮件发送给指定收件人进行审阅。

**语法**

**express.SendForReview(Recipients, Subject, ShowMessage, IncludeAttachment)**

*express*   一个代表 **Document** 对象的变量。

**参数**

| **名称**            | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Recipients*        | 可选          | **Variant**  | 列出邮件收件人的字符串。这些字符串可以是在电子邮件电话簿或完整电子邮件地址中未解析的名称和别名。用分号 (;) 分隔开多个收件人。如果留有空格并且 ShowMessage 为 False，则将收到错误消息，邮件也不会发送出去。 |
| *Subject*           | 可选          | **Variant**  | 邮件主题字符串。如果为空，则主题将为：请审阅“文件名”。       |
| *ShowMessage*       | 可选          | **Variant**  | 一个代表执行该方法时是否显示邮件的 Boolean 值。默认值为 True。如果设置为 False，则邮件自动发送给接收者，而不先向发送者显示邮件。 |
| *IncludeAttachment* | 可选          | **Variant**  | 一个代表邮件是否应包含附件或到服务器位置的链接的 Boolean 值。默认值为 True。如果设置为 False，文档必须保存在共享位置中。 |

**说明**

使用 **SendForReview** 方法可启动合作审阅流程。使用 **EndReview** 方法可结束审阅流程。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例自动将当前文档作为电子邮件附件发送给指定收件人。*/ function WebReview() {     Application.ActiveDocument.SendForReview("someone@example.com; amy jones", "Please review this document.", false, true) }` |

#### **Document.SendMail**

打开一个邮件窗口以通过 Microsoft Exchange 发送指定文档。

**语法**

**express.SendMail()**

*express*   一个代表 **Document** 对象的变量。

**说明**

使用 **SendMailAttach** 属性可控制文档是作为邮件窗口中的文本还是作为附件发送。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档作为邮件的附件发送。*/ function test() {     Application.Options.SendMailAttach = true     Application.ActiveDocument.SendMail() }` |

#### **Document.SendMailer**

您请求就仅在 Macintosh 上使用的 Visual Basic 关键字获得帮助。有关 **Document** 对象的 **SendMailer** 方法的信息，请参阅 WPS OfficeMacintosh Edition 所附带的语言参考帮助。

**语法**

**express.SendMailer()**

*express*   一个代表 **Document** 对象的变量。

#### **Document.SetCompatibilityMode**

设置文档的兼容模式。

**语法**

**express.SetCompatibilityMode(Mode)**

*express*   一个代表 **Document** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Mode*   | 必选          | **Long**     | 指定要靠近的 WPS 版本。使用 WdCompatibilityMode 枚举的常量作为该参数的一个参数。 |

**说明**

在 WPS 2015 中打开在以前版本的 WPS 中创建的文档时，将启用“兼容模式”。“兼容模式”确保在处理文档时无法使用 WPS 2015 中的任何新的或增强的功能，从而使用以前版本的 WPS 编辑该文档的用户将具有完全编辑能力。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的示例代码将 WPS 2015 置于 WPS 2003 兼容模式。*/ Application.ActiveDocument.SetCompatibilityMode(wdWord2003)` |

#### **Document.SetDefaultTableStyle**

指定用于文档中新建表的表样式。

**语法**

**express.SetDefaultTableStyle(Style, SetInTemplate)**

*express*   一个代表 **Document** 对象的变量。

**参数**

| **名称**        | **必选/可选** | **数据类型** | **说明**                                              |
| --------------- | ------------- | ------------ | ----------------------------------------------------- |
| *Style*         | 必选          | **Variant**  | 指定样式名称的字符串。                                |
| *SetInTemplate* | 必选          | **Boolean**  | 如果该属性值为 True，则在文档附加的模板中保存表样式。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例检查活动文档中使用的默认表格样式是否是 Table Normal，如果是，将默认表格样式改为 TableStyle1。本示例假定存在名为 TableStyle1 的表格样式。*/ function TableDefaultStyle() {     if (Application.ActiveDocument.DefaultTableStyle == "Table Normal") {         Application.ActiveDocument.SetDefaultTableStyle("TableStyle1", true)     } }` |

#### **Document.SetLetterContent**

将指定 **LetterContent** 对象的内容插入文档。

**语法**

**express.SetLetterContent(LetterContent)**

*express*   一个代表 **Document** 对象的变量。

**参数**

| **名称**        | **必选/可选** | **数据类型**      | **说明**                 |
| --------------- | ------------- | ----------------- | ------------------------ |
| *LetterContent* | 必选          | **LetterContent** | 包括不同信函内容的对象。 |

**说明**

该方法与 **RunLetterWizard** 方法类似，但前者不显示“英文信函向导”对话框。该方法基于 **LetterContent** 对象的内容，在指定文档中添加、删除信函内容或重新设置信函格式。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例检索活动文档的“英文信函向导”元素，并更改信封用语文本；然后用 SetLetterContent 方法来更新活动文档，以便反映这些更改。*/ function test() {     let myLetterContent = Application.ActiveDocument.GetLetterContent()     myLetterContent.AttentionLine = "Greetings"     Application.ActiveDocument.SetLetterContent(myLetterContent) }` |

#### **Document.SetPasswordEncryptionOptions**

设置 WPS 用于通过密码加密文档的选项

**语法**

**express.SetPasswordEncryptionOptions(PasswordEncryptionProvider, PasswordEncryptionAlgorithm, PasswordEncryptionKeyLength, PasswordEncryptionFileProperties)**

*express*   一个代表 **Document** 对象的变量。

**参数**

| **名称**                           | **必选/可选** | **数据类型** | **说明**                                                    |
| ---------------------------------- | ------------- | ------------ | ----------------------------------------------------------- |
| *PasswordEncryptionProvider*       | 必选          | **String**   | 加密提供程序的名称。                                        |
| *PasswordEncryptionAlgorithm*      | 必选          | **String**   | 加密算法的名称。 WPS 支持流式加密算法。                     |
| *PasswordEncryptionKeyLength*      | 必选          | **Long**     | 加密密钥长度。必须是从 40 开始的 8 的倍数。                 |
| *PasswordEncryptionFileProperties* | 可选          | **Variant**  | 如果该属性值为 True，则 WPS 对文件属性加密。默认值为 True。 |

**说明**

为了增强安全性，不要使用不可靠的加密 (XOR)（又称作“OfficeXor”）或“Office97/2000 Compatible”（又称作“OfficeStandard”）算法。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果使用的密码加密算法是“OfficeXor”或“OfficeStandard”，本示例将密码加密设置为更强的加密。*/ function PasswordSettings() {     if (Application.ActiveDocument.PasswordEncryptionAlgorithm == "OfficeXor" || Application.ActiveDocument.PasswordEncryptionAlgorithm == "OfficeStandard") {         Application.ActiveDocument.SetPasswordEncryptionOptions("Microsoft RSA SChannel Cryptographic Provider", "RC4", 56, true)     } }` |

#### **Document.ToggleFormsDesign**

在打开和关闭窗体设计模式之间切换。

**语法**

**express.ToggleFormsDesign()**

*express*   一个代表 **Document** 对象的变量。

**说明**

当 WPS 处于窗体设计模式时，将显示**“控件工具箱”**工具栏。可以使用该工具箱插入各种 Microsoft ActiveX 控件，如命令按钮、滚动条和选项按钮。在窗体设计模式下，不会运行事件过程，并且单击一个嵌入的控件后，会出现控件的尺寸控点。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果当前没有显示“控件工具箱”工具栏，本示例将切换到窗体设计模式。*/ function test() {     if (Application.CommandBars.Item("Control Toolbox").Visible == false) {         Application.ActiveDocument.ToggleFormsDesign()     } }` |

#### **Document.TransformDocument**

将 Extensible Stylesheet Language Transformation (XSLT) 文件应用于指定文档，并用结果替换该文档。

**语法**

**express.TransformDocument(Path, DataOnly)**

*express*   一个代表 **Document** 对象的变量。

**参数**

| **名称**   | **必选/可选** | **数据类型** | **说明**                                                     |
| ---------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Path*     | 必选          | **String**   | XSLT 使用的路径。                                            |
| *DataOnly* | 可选          | **Boolean**  | 如果该属性值为 True，则仅将转换应用于文档（不包括 WPS XML）中的数据。如果该属性值为 False，则将转换应用于整篇文档（包括 WPS XML）。默认值为 True。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例使用指定的 XSLT 文件转换活动文档。*/ Application.ActiveDocument.TransformDocument("c:\\schemas\\simplesample.xslt")` |

#### **Document.Undo**

消最后一次操作或最后一系列操作，这些操作显示在**“撤消”**列表中。返回 **Boolean** 。如果撤消操作成功，则返回 **True**。

**语法**

**express.Undo(Times)**

*express*   一个代表 **Document** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**             |
| -------- | ------------- | ------------ | -------------------- |
| *Times*  | 可选          | **Variant**  | 要撤消的操作的数量。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 加粗，倾斜后，再撤销一部，仅保留加粗 */ function test() { 	let rng = Application.ActiveDocument.Range(0,3); 	rng.Bold = true; 	rng.Italic = true; 	Application.ActiveDocument.Undo(); }` |

#### **Document.UndoClear**

清除可对指定文档撤消的操作列表。

**语法**

**express.UndoClear()**

*express*   一个代表 **Document** 对象的变量。

**说明**

该方法对应于单击**“常用”**工具栏上的**“撤消”**按钮旁边的箭头时出现的项目列表。在宏的末尾包括该方法可防止如“VBA-Selection.InsertAfter”等 Visual Basic 操作显示在**“撤消”**框中。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例清除对活动文档撤消的操作列表。*/ Application.ActiveDocument.UndoClear()` |

#### **Document.Unprotect**

清除对指定文档的保护。

**语法**

**express.Unprotect(Password)**

*express*   一个代表 **Document** 对象的变量。

**参数**

| **名称**   | **必选/可选** | **数据类型** | **说明**                                                     |
| ---------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Password* | 可选          | **Variant**  | 用于保护文档的密码字符串。密码是区分大小写的。如果用户在使用一篇设置有密码的文档时没有提供正确的密码，就会显示一个对话框，提示用户输入密码。 |

**说明**

如果对文档没有加以保护，则该方法会导致出错。

## 安全性

请避免在应用程序中使用硬编码的密码。如果过程中要求输入密码，请从用户处请求密码，并将密码作为变量存储起来，然后在代码中使用该变量。有关如何做到这一点的建议最佳操作，请参阅 WPS Office解决方案开发人员安全注意事项。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例使用 strPassword 变量的值作为密码，清除对活动文档的保护。*/ function test() {     let strPassword     if (Application.ActiveDocument.ProtectionType != wdNoProtection) {         Application.ActiveDocument.Unprotect(strPassword)     } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例清除对活动文档的保护。然后插入文本并对文档进行修订保护。*/ function test() {     let strPassword     let aDoc = Application.ActiveDocument     if (aDoc.ProtectionType != wdNoProtection) {         aDoc.Unprotect()         Selection.InsertBefore("department six")         aDoc.Protect(wdAllowOnlyRevisions, null, strPassword)     } }` |

#### **Document.UpdateStyles**

将所有样式从附加模板复制到文档中，同时覆盖文档中所有现有同名的样式。

**语法**

**express.UpdateStyles()**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将相关模板中的样式复制到每一打开文档中，然后关闭每一文档。*/ function test() {     for (let i = 1; i <= Application.Documents.Count; i++) {         Application.Documents.Item(i).UpdateStyles()         Application.Documents.Item(i).Close(wdSaveChanges)     } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例更改活动文档所选用的模板中“标题 1”的样式设置。用 UpdateStyles 方法更新活动文档中的样式（包括“标题 1”样式）。*/ function test() {     let aDoc = Application.ActiveDocument.AttachedTemplate.OpenAsDocument     let font = aDoc.Styles.Item(wdStyleHeading1).Font     font.Name = "Arial"     font.Bold = false     aDoc.Close(wdSaveChanges)     Application.ActiveDocument.UpdateStyles() }` |

#### **Document.UpdateSummaryProperties**

更新**“文件”**菜单**“属性”**对话框中的关键词和备注文字，以反映为指定文档编写的自动摘要内容。

**语法**

**express.UpdateSummaryProperties()**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例突出显示活动文档的要点，并且更新“文件”菜单“属性”对话框中的摘要信息。*/ function test() {     Application.ActiveDocument.AutoSummarize(wd25Percent, wdSummaryModeHighlight)     Application.ActiveDocument.UpdateSummaryProperties() }` |

#### **Document.ViewCode**

显示指定文档中所选 Microsoft ActiveX 控件的代码窗口。

**语法**

**express.ViewCode()**

*express*   一个代表 **Document** 对象的变量。

**说明**

该方法只有在 WPS 之外才可用。

#### **Document.ViewPropertyBrowser**

显示指定文档中所选 Microsoft ActiveX 控件的属性窗口。

**语法**

**express.ViewPropertyBrowser()**

*express*   一个代表 **Document** 对象的变量。

**说明**

该方法只有在 WPS 之外才可用。

#### **Document.WebPagePreview**

显示将当前文档保存为网页时的预览。

**语法**

**express.WebPagePreview()**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示将文档保存为网页时的预览。*/ Application.ActiveDocument.WebPagePreview()` |

**成员属性**

#### **Document.ActiveTheme**

返回指定文档的活动主题名称及该主题的格式选项。**String** 类型，只读。

**语法**

**express.ActiveTheme**

*express*   一个代表 **Document** 对象的变量。

**说明**

如果该文档无活动主题，则 **ActiveTheme** 属性返回“none”。有关该属性返回值的详细解释，请参阅 **ApplyTheme** 方法的 *Name* 参数。该属性的返回值可能与主题的显示名称不对应。若要返回主题的显示名称，可使用 **ActiveThemeDisplayName** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例实现的功能是：为当前文档应用主题并显示活动主题的名称及主题格式选项。*/ function CheckTheme() {     Application.ActiveDocument.ApplyTheme("artsy 100")     alert(Application.ActiveDocument.ActiveTheme) }` |

#### **Document.ActiveThemeDisplayName**

返回指定文档的活动主题的显示名称。**String** 类型，只读。

**语法**

**express.ActiveThemeDisplayName**

*express*   一个代表 **Document** 对象的变量。

**说明**

如果该文档无活动主题，则 **ActiveThemeDisplayName** 属性返回“none”。主题的显示名称指显示在**“主题”**对话框中的名称。此名称可能与用来设置默认主题或为文档应用主题的字符串不对应。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回当前文档活动主题的显示名称。*/ function CheckTheme() {     Application.ActiveDocument.ApplyTheme("artsy 100")     alert(Application.ActiveDocument.ActiveThemeDisplayName) }` |

#### **Document.ActiveWindow**

返回一个 **Window** 对象，该对象代表活动窗口（焦点所在的窗口）。只读。

**语法**

**express.ActiveWindow**

*express*   一个代表 **Document** 对象的变量。

**说明**

如果没有打开的窗口，则使用 **ActiveWindow**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#Application.ActiveWindow)属性将出现错误。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例显示活动窗口的标题文本 */ function CheckTheme() {     Application.ActiveDocument.ActiveWindow.Caption     Application.ActiveDocument.ActiveWindow.Split = true }` |

#### **Document.ActiveWritingStyle**

返回或设置指定文档中指定语言的写作风格。**String** 类型，可读写。

**语法**

**express.ActiveWritingStyle**

*express*   一个代表 **Document** 对象的变量。

**说明**

**WritingStyleList** 属性返回一个可用写作风格名称的数组。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置活动文档中法语、德语和美国英语的写作风格。若要运行本示例，必须安装法语、德语和美国英语的语法文件。*/ function test() {     Application.ActiveDocument.ActiveWritingStyle(wdFrench).value = "Commercial"     Application.ActiveDocument.ActiveWritingStyle(wdGerman).value = "Technisch/Wiss"     Application.ActiveDocument.ActiveWritingStyle(wdEnglishUS).value = "Technical" }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回选定内容的语言的写作风格。*/ function WhichLanguage() {     let varLang = Application.Selection.LanguageID     alert(Application.ActiveDocument.ActiveWritingStyle(varLang)) }` |

#### **Document.Application**

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

*express*   一个代表 **Document** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **Document.AttachedTemplate**

返回一个 **Template** 对象，该对象代表与指定文档相关联的模板。可读/写 **Variant** 类型。

**语法**

**express.AttachedTemplate**

*express*   一个代表 **Document** 对象的变量。

**说明**

若要设置该属性，请指定模板名称或返回 **Template** 对象的表达式。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示与活动文档相关联的模板名称和路径。*/ function test() {     let myTemplate = Application.ActiveDocument.AttachedTemplate     alert(myTemplate.Path + Application.PathSeparator + myTemplate.Name) }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在文档 1 的开始处插入图文场的内容（一个内置“自动图文集”词条）。*/ function test() {     let myRange = Application.Documents.Item(1).Range(0, 0)     Applicatiion.Documents.Item(1).AttachedTemplate.AutoTextEntries.Item("Spike").Insert(myRange) }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为活动文档附加名为“Letter.dot”模板。*/ function test() {     Application.ActiveDocument.AttachedTemplate = "C:\\Templates\\Letter.dot" }` |

#### **Document.AutoFormatOverride**

返回或设置一个 **Boolean** 类型的值，该数值代表是否在已应用格式设置限制的文档中使用自动设置格式替代格式设置限制。

**语法**

**express.AutoFormatOverride**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例指定在受保护的文档中使用自动设置格式替代格式设置限制。*/ Application.ActiveDocument.AutoFormatOverride = true` |

#### **Document.AutoHyphenation**

如果代表指定文档已启动自动断字功能，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.AutoHyphenation**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例启动自动断字功能，断字区为 0.25 英寸。全为大写的单词不断字。*/ function test() { ActiveDocument.HyphenationZone = InchesToPoints(0.25) ActiveDocument.HyphenateCaps = false ActiveDocument.AutoHyphenation = true }` |

#### **Document.Background**

返回一个 **Shape** 对象，该对象代表指定文档的背景图像。只读。

**语法**

**express.Background**

*express*   一个代表 **Document** 对象的变量。

**说明**

只有在 Web 版式视图中才能看到背景。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动窗口的 Web 版式视图的背景色设置为浅灰色。*/ function test() { ActiveDocument.ActiveWindow.View.Type = wdWebView let fill = ActiveDocument.Background.Fill fill.Visible = true fill.ForeColor.RGB = 192, 192, 192 }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将 Web 版式视图的背景位图图像设置为 Bubbles.bmp。*/ function test() { ActiveDocument.ActiveWindow.View.Type = wdWebView ActiveDocument.Background.Fill.UserPicture("C:\\Windows\\Bubbles.bmp") }` |

#### **Document.Bibliography**

返回一个 **Bibliography** 对象，代表文档中包含的书目参考。只读。

**语法**

**express.Bibliography**

*express*   一个代表 **Document** 对象的变量。

#### **Document.Bookmarks**

返回一个 **Bookmarks** 集合，该集合代表文档中的所有书签。只读。

**语法**

**express.Bookmarks**

*express*   一个代表 **Document** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 返回第一个书签的开始结束位置 */ function test() { 	let myMark = Application.ActiveDocument.Bookmarks.Item(1) 	BookStart = myMark.Start 	BookEnd = myMark.End 	alert("start:" + BookStart + " end" + BookEnd) }` |

#### **Document.BuiltInDocumentProperties**

返回一个 **DocumentProperties** 集合，该集合代表指定文档的所有内置的文档属性。

**语法**

**express.BuiltInDocumentProperties**

*express*   一个代表 **Document** 对象的变量。

**说明**

要返回单个代表特定内置文档属性的 **DocumentProperty** 对象，可使用 **BuiltinDocumentProperties** 属性。如果 WPS 没有为一个内置的文档属性定义一个值，则读取这个文档属性的 **Value** 属性时会产生一个错误。

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

用 **CustomDocumentProperties** 属性返回自定义文档属性的集合。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在活动文档的末尾插入一个内置属性列表。*/ function ListProperties() {     let rngDoc = ActiveDocument.Content     rngDoc.Collapse(wdCollapseEnd)     let proDoc = ActiveDocument.BuiltInDocumentProperties     for(let i=1;i <= proDoc.Count;i++) {         try {             rngDoc.InsertParagraphAfter()             rngDoc.InsertAfter(proDoc.Item(i).Name + "= ")         }         catch(exception) {             rngDoc.InsertAfter(proDoc.Item(i).Value)         }     } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示活动文档中的单词数。*/ function DisplayTotalWords() {     let intWords = ActiveDocument.BuiltInDocumentProperties.Item(wdPropertyWords).Value     MsgBox("This document contains " + intWords + " words.") }` |

#### **Document.Characters**

返回一个 **Characters** 集合，该集合代表文档中的字符。只读。

**语法**

**express.Characters**

*express*   一个代表 **Document** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 返回文档中字符数，包括空格 */ Application.ActiveDocument.Characters.Count` |

#### **Document.ChildNodeSuggestions**

返回一个 **XMLChildNodeSuggestions** 集合，该集合代表允许在文档中使用的元素列表。

**语法**

**express.ChildNodeSuggestions**

*express*   一个代表 **Document** 对象的变量。

**说明**

该属性返回所有附加架构的根元素。

**XMLChildNodeSuggestions** 集合中的每个 **XMLChildNodeSuggestion** 对象都是允许的、可能做为子元素的 XML 元素列表中的一项，该列表位于**“XML 结构”**任务窗格底部。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下列示例循环遍历活动文档中选定的第一个元素的建议子元素，并在插入点插入所有允许的元素。*/ function GetChildNodeSuggestions() {     let objNode = Selection.XMLParentNode     let objSuggestion = objNode.ChildNodeSuggestions     for(let i=1; i <= objSuggestion.Count; i++) {         objSuggestion.Insert()         Selection.MoveRight()     } }` |

#### **Document.ClickAndTypeParagraphStyle**

返回或设置“即点即输”功能在指定文档中应用于文字的默认段落样式。**Variant** 类型，可读写。

**语法**

**express.ClickAndTypeParagraphStyle**

*express*   一个代表 **Document** 对象的变量。

**说明**

要设置 **ClickAndTypeParagraphStyle** 属性，请指定样式的本地名称、一个整数、一个 **WdBuiltinStyle** 常量或一个代表样式的对象。有关 **WdBuiltinStyle** 常量的列表，请参阅“Visual Basic 对象浏览器”。

如果将指定样式的 **InUse** 属性设为 **False**，则会出现错误。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置“即点即输”应用于纯文本文件的默认段落样式。*/ function test() { if(ActiveDocument.Styles.Item("Plain Text").InUse) {     ActiveDocument.ClickAndTypeParagraphStyle = "Plain Text" } else {     MsgBox("Sorry, this style is not in use yet.") } }` |

#### **Document.CoAuthoring**

返回 CoAuthoring 对象，该对象提供文档中与共同创作相关的对象模型的入口点。只读。

**语法**

**express.CoAuthoring**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的代码示例返回活动文档中的所有合著者。*/ let allAuthors = ActiveDocument.CoAuthoring.Authors` |

#### **Document.CodeName**

返回指定文档的代码名称。**String** 类型，只读。

**语法**

**express.CodeName**

*express*   一个代表 **Document** 对象的变量。

**说明**

代码名称是保存文档的事件宏的模块名。该模块的默认名为“ThisDocument”，可在“工程”窗口中查看该模块。有关使用 Document 对象的事件的内容，请参阅 使用 Document 对象的事件。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回活动文档的代码窗口名称。*/ MsgBox(ActiveDocument.CodeName)` |

#### **Document.CommandBars**

返回一个 **CommandBars** 集合，该集合代表 WPS 中的菜单栏以及所有工具栏。

**语法**

**express.CommandBars**

*express*   一个代表 **Document** 对象的变量。

**说明**

在访问 **CommandBars** 集合之前，可用 **CustomizationContext** 属性设置模板或文档上下文。

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例增大所有命令栏按钮并激活工具提示。*/ function test() { CommandBars.LargeButtons = true CommandBars.DisplayTooltips = true }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在应用程序窗口底部显示“绘图”工具栏。*/ function test() { CommandBars.Item("Drawing").Visible = true CommandBars.Item("Drawing").Position = msoBarBottom }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将“版本”命令按钮添至“常用”工具栏。*/ function test() { CustomizationContext = NormalTemplate CommandBars.Item("Standard").Controls.Add(msoControlButton, 2522, null, 4) }` |

#### **Document.Comments**

返回一个 **Comments** 集合，该集合代表指定文档的所有批注。只读。

**语法**

**express.Comments**

*express*   一个代表 **Document** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例将活动文档中的每一批注的作者姓名与“工具”菜单中“选项”对话框的“用户信息”选项卡的用户姓名作比较。如果两者不同,则设置备注内容为"test" */ function test() { 	let comm = Application.ActiveDocument.Comments 	for(let i=1;i <= comm.Count;i++) { 	    if(comm.Item(i).Author != Application.UserName) { 	        comm.Item(i).Range.Text = 'test' 	    } 	} } ` |

#### **Document.CompatibilityMode**

返回 **Number** 类型值，该值指定打开文档时 WPS 2015 使用的兼容模式。只读。 

**语法**

**express.CompatibilityMode**

*express*   一个代表 **Document** 对象的变量。

**说明**

在 WPS 2015 中打开在以前版本的 WPS 中创建的文档时，将启用“兼容模式”。“兼容模式”确保在处理文档时无法使用 WPS 2015 中的任何新的或增强的功能，从而使用以前版本的 WPS 编辑该文档的用户将具有完全编辑能力。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 获取文档的兼容文档 */ Application.ActiveDocument.CompatibilityMode` |

#### **Document.ConsecutiveHyphensLimit**

返回或设置以连字符结束的连续行的最大数目。**Long** 类型，可读写。

**语法**

**express.ConsecutiveHyphensLimit**

*express*   一个代表 **Document** 对象的变量。

**说明**

如果将 **ConsecutiveHyphensLimit** 属性设置为 0（零），则以连字符结束的连续行可以是任何数目。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为 MyReport.doc 设置自动断字功能并将以连字符结束的连续行的数目限制为 2 个。*/ function test() { Documents.Item("MyReport.doc").AutoHyphenation = true Documents.Item("MyReport.doc").ConsecutiveHyphensLimit = 2 } ` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例不限制以连字符结束的连续行的数目。*/ ActiveDocument.ConsecutiveHyphensLimit = 0` |

#### **Document.Container**

返回一个对象，该对象代表指定文档的容器应用程序。**Object** 类型，只读。

**语法**

**express.Container**

*express*   一个代表 **Document** 对象的变量。

**说明**

如果指定的文档作为 OLE 对象嵌入到另一个应用程序中，则通过 **Container** 属性可以访问该文档的容器应用程序。如果一个 WPS 文档是作为 ActiveX 文档打开的（如在 WPS Office活页夹或 Internet Explorer 中打开一个 WPS 文档），则通过该属性也可访问容器应用程序中的对象模型。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示活动文档第一个图形的容器应用程序名。若要让本示例正常运行，则该图形必须是一个 OLE 对象。*/ Msgbox(ActiveDocument.Shapes.Item(1).OLEFormat.Object.Container.Name)` |

#### **Document.Content**

返回一个 **Range** 对象，该对象代表主文档文章。只读。

**语法**

**express.Content**

*express*   一个代表 **Document** 对象的变量。

**说明**

以下两个语句是等价的：

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() { let mainStory mainStory = ActiveDocument.Content mainStory = ActiveDocument.StoryRanges(wdMainTextStory)  }` |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档中的字体和字号分别改为 Arial 和 10 磅。*/ function test() { let myRange = ActiveDocument.Content myRange.Font.Name = "Arial" myRange.Font.Size = 10 }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在名为“Changes.doc”的文档末尾插入文本。For Each...Next 语句用来判断此文档是否已打开。*/ function test() { for(let i=1;i <= Documents.Count;i++) {     if(Documents.Item(i).Name.toLowerCase().indexOf("changes.doc") != -1) {         let myRange = Documents.Item("Changes.doc").Content         myRange.InsertAfter("the end.")     } } }` |

#### **Document.ContentControls**

返回一个 **ContentControls** 集合，该集合代表文档中的所有内容控件。只读。

**语法**

**express.ContentControls**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 下拉列表内容控件插入到活动文档中 */ let objCC = Application.ActiveDocument.ContentControls.Add(wdContentControlDropdownList)` |

#### **Document.ContentTypeProperties**

返回一个 **MetaProperties** 集合，该集合代表文档中所存储的元数据，如作者姓名、主题和公司。只读。

**语法**

**express.ContentTypeProperties**

*express*   一个代表 **Document** 对象的变量。

#### **Document.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Document** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **Document.CurrentRsid**

返回一个 **Long** 类型的值，该值代表 WPS 为文档中的更改所分配的随机数。只读。

**语法**

**express.CurrentRsid**

*express*   一个代表 **Document** 对象的变量。

**说明**

如果 **StoreRSIDOnSave** 属性为 **True**，则每次保存文档时， WPS 生成一个随机数，应用程序使用该随机数来比较和合并文档。 WPS 将这些随机数保存在表格中并在每次保存后更新该表格。**CurrentRsid** 属性将返回 WPS 分配给文档的上一个数。

#### **Document.CustomDocumentProperties**

返回一个 **DocumentProperties** 集合，该集合代表指定文档的所有自定义文档属性。

**语法**

**express.CustomDocumentProperties**

*express*   一个代表 **Document** 对象的变量。

**说明**

使用 **BuiltInDocumentProperties** 属性可返回内置文档属性的集合。

**msoPropertyTypeString** 类型的属性的长度不能超过 255 个字符。

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在活动文档的末尾插入一个自定义内置属性列表。*/ function test() { let myRange = ActiveDocument.Content myRange.Collapse(wdCollapseEnd)									 let prop = ActiveDocument.CustomDocumentProperties for(let i=1;i <= prop.Count;i++) {     myRange.InsertParagraphAfter()     myRange.InsertAfter(prop.Item(i).Name + "= ")     myRange.InsertAfter(prop.Item(i).Value) } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为 Sales.doc 添加一个自定义内置属性。*/ function test() { let thename = prompt("Please type your name", "Name") Documents.Item("Sales.doc").CustomDocumentProperties.Add("YourName", false, msoPropertyTypeString, thename) }` |

#### **Document.CustomXMLParts**

返回一个 **CustomXMLParts** 集合，该集合代表 XML 数据存储区中的自定义 XML。只读。

**语法**

**express.CustomXMLParts**

*express*   一个代表 **Document** 对象的变量。

#### **Document.DefaultTableStyle**

返回一个 **Variant** 类型的值，该值代表可以应用于文档中新创建的所有表格的表格样式。只读。

**语法**

**express.DefaultTableStyle**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例检查活动文档中使用的默认表格样式是否是“Table Normal”，如果是，将默认表格样式改为“TableStyle1”。本示例假定存在名为“TableStyle1”的表格样式。*/ function TableDefaultStyle() {     if(ActiveDocument.DefaultTableStyle == "Table Normal") {         ActiveDocument.SetDefaultTableStyle("TableStyle1", true)     } }` |

#### **Document.DefaultTabStop**

以磅为单位返回或设置指定文档中的默认制表位间隔。**Single** 类型，可读写。

**语法**

**express.DefaultTabStop**

*express*   一个代表 **Document** 对象的变量。

**示例**

------

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档中的默认制表位设置为 1 英寸。其中 InchesToPoints 方法用来将英寸转化为磅值。*/ ActiveDocument.DefaultTabStop = InchesToPoints(1)` |

#### **Document.DefaultTargetFrame**

返回或设置一个 **String** 类型的值，该值代表通过超链接可到达显示网页的浏览器框架。可读写。

**语法**

**express.DefaultTargetFrame**

*express*   一个代表 **Document** 对象的变量。

**说明**

当 **DefaultTargetFrame** 属性可以使用任意用户定义的字符串时，它有下列预定义的字符串：“_top”、“_blank”、“_parent”和“_self”。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例的功能如下：当用户单击活动文档中的超链接时，WPS 打开一个新的空白浏览器窗口。*/ function DefaultFrame() {     ActiveDocument.DefaultTargetFrame = "_blank" }` |

#### **Document.DisableFeatures**

如果为 **True**，则禁用 **DisableFeaturesIntroducedAfter** 属性中指定的版本之后的所有功能。默认值为 **False**。**Boolean** 类型，可读写。

**语法**

**express.DisableFeatures**

*express*   一个代表 **Document** 对象的变量。

**说明**

**DisableFeatures** 属性只影响设置属性的文档。如果计划与使用 WPS 较早版本的用户共享文档，请使用该属性，这样就无需禁用其 WPS 版本不支持的文档功能。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在当前文档中禁用在 WPS for Windows 、7.0 和 7.0a 版本之后添加的功能，全局默认设置保持不变。*/ function FeaturesDisable() {      //Checks whether features are disabled     if(ActiveDocument.DisableFeatures == true) {          //If they are, disables all features after  WPS for Windows         ActiveDocument.DisableFeaturesIntroducedAfter = wd70 }     else {          //If not, turns on the disable features option and disables         //all features introduced after  WPS for Windows         ActiveDocument.DisableFeatures = true         ActiveDocument.DisableFeaturesIntroducedAfter = wd70 } }` |

#### **Document.DisableFeaturesIntroducedAfter**

仅对该文档禁用指定的 WPS 版本后引入的所有功能。**WdDisableFeaturesIntroducedAfter** 类型，可读写。

**语法**

**express.DisableFeaturesIntroducedAfter**

*express*   一个代表 **Document** 对象的变量。

**说明**

设置 **DisableFeaturesIntroducedAfter** 属性之前，必须先将 **DisableFeatures** 属性设为 **True**。否则，该设置无效，并保持其默认值 WPS 97 for Windows。

**DisableFeaturesIntroducedAfter** 属性仅影响该属性所设置的文档。如果要设置应用程序的全局选项以禁用所有文档的功能。可使用 **DisableFeaturesIntroducedAfterByDefault** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例仅对当前文档禁用 WPS for Windows 的 7.0 和 7.0a 版本之后引入的所有功能。全局默认设置保持不变。*/ function FeaturesDisable() {      //Checks whether features are disabled     if(ActiveDocument.DisableFeatures == true) {          //If they are, disables all features after  WPS for Windows         ActiveDocument.DisableFeaturesIntroducedAfter = wd70 }     else {          /*If not, turns on the disable features option and disables         all features introduced after  WPS for Windows*/         ActiveDocument.DisableFeatures = true         ActiveDocument.DisableFeaturesIntroducedAfter = wd70 } }` |

#### **Document.DocumentInspectors**

返回一个 **DocumentInspectors** 集合，该集合使您能够查找隐藏的个人信息，如作者姓名、公司名称和修订日期。只读。

**语法**

**express.DocumentInspectors**

*express*   一个代表 **Document** 对象的变量。

#### **Document.DocumentLibraryVersions**

返回一个 **DocumentLibraryVersions** 集合，该集合表示共享文档的版本集合，该共享文档启用了版本控制并存储在服务器的文档库中。

**语法**

**express.DocumentLibraryVersions**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例返回活动文档的版本集合。本示例假设活动文档启用了版本控制并存储在服务器的共享文档库中。*/ let objVersions = ActiveDocument.DocumentLibraryVersions` |

#### **Document.DocumentTheme**

返回一个 **OfficeTheme** 对象，代表应用于文档的 WPS Office主题。只读。

**语法**

**express.DocumentTheme**

*express*   一个代表 **Document** 对象的变量。

**说明**

使用 **ApplyDocumentTheme** 方法可应用 Office 主题。

#### **Document.DoNotEmbedSystemFonts**

如果为 **True**，则 WPS 不嵌入常规系统字体。**Boolean** 类型，可读写。

**语法**

**express.DoNotEmbedSystemFonts**

*express*   一个代表 **Document** 对象的变量。

**说明**

如果用户使用中文系统创建文档，并希望该文档对于那些系统中没有该语言字体的用户也可读，可将 **Document** 属性设为 **False**。例如，日语系统的用户可以选择在文档中嵌入字体，使日文文档在所有系统中可读。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在当前文档中嵌入所有字体。*/ function EmbedFonts() {     if(ActiveDocument.EmbedTrueTypeFonts == false) {         ActiveDocument.EmbedTrueTypeFonts = true         ActiveDocument.DoNotEmbedSystemFonts = false }     else {         ActiveDocument.DoNotEmbedSystemFonts = false } }` |

#### **Document.Email**

返回一个 **Email** 对象，该对象包含当前文档的所有与电子邮件相关的属性。只读。

**语法**

**express.Email**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将返回与当前电子邮件作者相关联的样式名。*/ MsgBox(ActiveDocument.Email.CurrentEmailAuthor.Style.NameLocal)` |

#### **Document.EmbedLinguisticData**

如果为 **True**，则 WPS 会嵌入语音和手写功能，以便数据可以转换回语音或手写。**Boolean** 类型，可读写。

**语法**

**express.EmbedLinguisticData**

*express*   一个代表 **Document** 对象的变量。

**说明**

使用 **EmbedLinguisticData** 属性还能保存东亚 IME 键击，以便使用 Windows 文本服务框架应用程序编程接口提高对文本服务数据（源自与 WPS Office连接的设备）的更正和控制能力。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将文档中可能出现的任何语音和手写输入嵌入活动文档中。*/ function EmbedSpeechHandwriting() {     ActiveDocument.EmbedLinguisticData = true }` |

#### **Document.EmbedTrueTypeFonts**

如果在保存文档时，WPS 在文档中嵌入 TrueType 字体，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.EmbedTrueTypeFonts**

*express*   一个代表 **Document** 对象的变量。

**说明**

通过嵌入 TrueType 字体，其他人可以使用创建文档时所用的字体来查看该文档。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 在保存文档时，在文档中自动嵌入 TrueType 字体，然后保存活动文档。*/ function test() { ActiveDocument.EmbedTrueTypeFonts = true ActiveDocument.Save() }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“选项”对话框中“保存”选项卡上“保存选项”区域中“嵌入 TrueType 字体”复选框的当前状态。*/ let temp = ActiveDocument.EmbedTrueTypeFonts` |

#### **Document.EncryptionProvider**

返回一个 **String** 类型的值，该值指定 WPS 在对文档进行加密时使用的算法加密提供程序的名称。可读写。

**语法**

**express.EncryptionProvider**

*express*   一个代表 **Document** 对象的变量。

#### **Document.Endnotes**

返回一个 **Endnotes** 集合，该集合代表文档中的所有尾注。只读。

**语法**

**express.Endnotes**

*express*   一个代表 **Document** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档中的尾注放在文档的结尾并将尾注引用编号设置为小写的罗马数字。*/ function test() { let endnotes = ActiveDocument.Endnotes endnotes.Location = wdEndOfDocument endnotes.NumberStyle = 2 }` |

#### **Document.EnforceStyle**

返回或设置一个 **Boolean** 类型的值，该值代表是否在受保护的文档中实施格式设置限制。

**语法**

**express.EnforceStyle**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例启用活动文档中的格式设置限制。*/ ActiveDocument.EnforceStyle = true` |

#### **Document.Envelope**

返回一个 **Envelope** 对象，该对象代表文档中的信封和信封功能。只读。

**语法**

**express.Envelope**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将默认信封尺寸设为 C4（229 x 324 毫米）。*/ ActiveDocument.Envelope.DefaultSize = "C4"` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果文档中添加了一个信封，则本示例显示收信人地址，否则显示一个消息框。*/ function test() { try {     let addr = ActiveDocument.Envelope.Address.Text     MsgBox(addr, null, "Delivery Address") } catch(exception) {     MsgBox("Add an envelope to the document") } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例新建一篇文档，并添加有预定义收信人地址和寄信人地址的一个信封。*/ function test() { let addr = "Don Funk" + "\r" + "123 Skye St." + "\r" + "Our Town, WA  98040" let retaddr = "Karin Gallagher" + "\r" + "123 Main" + "\r" + "Other Town, WA  98004" Documents.Add().Envelope.Insert(null, addr, null, null, retaddr) ActiveDocument.ActiveWindow.View.Type = wdPrintView }` |

#### **Document.FarEastLineBreakLanguage**

返回或设置一个 **WdFarEastLineBreakLanguageID** 类型的值，该值代表在指定文档或模板中对文本进行换行时使用的东亚语言。可读写。

**语法**

**express.FarEastLineBreakLanguage**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 在当前文档中依照朝鲜语的规则换行。*/ ActiveDocument.FarEastLineBreakLanguage = wdLineBreakKorean` |

#### **Document.FarEastLineBreakLevel**

返回或设置一个 **WdFarEastLineBreakLevel** 类型的值，该值代表指定文档的换行控制级别。可读写。

**语法**

**express.FarEastLineBreakLevel**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 在活动文档中第一级首尾字符处换行。*/ ActiveDocument.FarEastLineBreakLevel = wdJustificationModeCompressKana    ` |

#### **Document.Fields**

返回一个 **Fields**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Fields/Fields%20.htm#jsObject_Fields)集合，该集合代表文档中的所有域。只读。

**语法**

**express.Fields**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                 |
| -------------------------------------------- |
| `Application.ActiveDocument.Fields.Update()` |

#### **Document.Final**

返回或设置 **Boolean** 值，该值指示文档是否是最终的。可读/写。

**语法**

**express.Final**

*express*   一个代表 **Document** 对象的变量。

**说明**

**True** 将文档标记为最终版本，以通知收件人（如果有）文档是最终的，并将文档设置为只读。

#### **Document.Footnotes**

返回一个 **Footnotes**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Footnotes/Footnotes%20.htm#jsObject_Footnotes)集合，该集合代表文档中的所有脚注。只读。

**语法**

**express.Footnotes**

*express*   一个代表 **Document** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 设置当前文档脚注的起始编号 */ Application.ActiveDocument.Footnotes.StartingNumber = 3` |

#### **Document.FormattingShowClear**

如果为 **True**，则 WPS 在**“样式和格式”**任务窗格中显示“清除格式”。**Boolean** 类型，可读写。

**语法**

**express.FormattingShowClear**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例禁用样式列表中“清除格式”按钮的显示。*/ function ShowClearFormatting() {     ActiveDocument.FormattingShowClear = false     ActiveDocument.FormattingShowFilter = wdShowFilterFormattingInUse     ActiveDocument.FormattingShowFont = true     ActiveDocument.FormattingShowNumbering = true     ActiveDocument.FormattingShowParagraph = true }` |

#### **Document.FormattingShowFilter**

设置或返回一个 **WdShowFilter** 常量，代表**“样式和格式”**任务窗格中显示的样式和格式。**Boolean** 类型，可读写

**语法**

**express.FormattingShowFilter**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例对格式进行筛选，在“样式和格式”任务窗格中仅显示活动文档中使用的格式。*/ function ShowClearFormatting() {     ActiveDocument.FormattingShowClear = false     ActiveDocument.FormattingShowFilter = wdShowFilterFormattingInUse     ActiveDocument.FormattingShowFont = true     ActiveDocument.FormattingShowNumbering = true     ActiveDocument.FormattingShowParagraph = true }` |

#### **Document.FormattingShowFont**

如果为 **True**，则 WPS 显示**“样式和格式”**任务窗格中的字体格式。**Boolean** 类型，可读写。

**语法**

**express.FormattingShowFont**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例启用“样式和格式”任务窗格中字体样式的显示功能。*/ function ShowClearFormatting() {     ActiveDocument.FormattingShowClear = false     ActiveDocument.FormattingShowFilter = wdShowFilterFormattingInUse     ActiveDocument.FormattingShowFont = true     ActiveDocument.FormattingShowNumbering = true     ActiveDocument.FormattingShowParagraph = true }` |

#### **Document.FormattingShowNextLevel**

返回或设置 **Boolean** 值，该值表示 WPS 在使用上一个标题级别时是否显示下一个标题级别。可读/写。

**语法**

**express.FormattingShowNextLevel**

*express*   一个代表 **Document** 对象的变量。

**说明**

此属性对应于**“样式库选项”**对话框中的**“使用上一个级别时显示下一级标题”**复选框。

#### **Document.FormattingShowNumbering**

如果为 **True**，则 WPS 在**“样式和格式”**任务窗格中显示编号格式。**Boolean** 类型，可读写。

**语法**

**express.FormattingShowNumbering**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例启用“样式和格式”任务窗格中编号格式的显示功能。*/ function ShowClearFormatting() {     ActiveDocument.FormattingShowClear = false     ActiveDocument.FormattingShowFilter = wdShowFilterFormattingInUse     ActiveDocument.FormattingShowFont = true     ActiveDocument.FormattingShowNumbering = true     ActiveDocument.FormattingShowParagraph = true }` |

#### **Document.FormattingShowParagraph**

如果为 **True**，则 WPS 在**“样式和格式”**任务窗格中显示段落格式。**Boolean** 类型，可读写。

**语法**

**express.FormattingShowParagraph**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例启用“样式和格式”任务窗格中段落格式的显示功能。*/ function ShowClearFormatting() {     ActiveDocument.FormattingShowClear = false     ActiveDocument.FormattingShowFilter = wdShowFilterFormattingInUse     ActiveDocument.FormattingShowFont = true     ActiveDocument.FormattingShowNumbering = true     ActiveDocument.FormattingShowParagraph = true }` |

#### **Document.FormattingShowUserStyleName**

返回或设置 **Boolean** 值，该值表示是否显示用户定义的样式。可读/写。

**语法**

**express.FormattingShowUserStyleName**

*express*   一个代表 **Document** 对象的变量。

**说明**

此属性对应于**“样式库选项”**对话框中的**“存在可选名称时隐藏内置名称”**复选框。当存在可选的用户定义样式时，将 **FormattingShowUserStyleName** 属性设置为 **True** 会隐藏内置样式。例如，如果用户创建标题样式（如“标题 1”或“标题 2”）以利用 WPS 的其他内置功能（例如，目录），则用户定义的样式优先，并且在样式名称列表中不会显示以相似方式命名的内置样式。

#### **Document.FormFields**

返回一个 **FormFields** 集合，该集合代表文档中的所有窗体域。只读。

**语法**

**express.FormFields**

*express*   一个代表 **Document** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将名为“Text1”的窗体域的内容设置为“Name”。*/ ActiveDocument.FormFields.Item("Text1").Result = "Name"` |

#### **Document.FormsDesign**

如果指定的文档处于窗体设计模式，则该属性值为 **True**。**Boolean** 类型，只读。

**语法**

**express.FormsDesign**

*express*   一个代表 **Document** 对象的变量。

**说明**

如果在从 WPS 中运行的代码中使用 **FormsDesign** 属性，则本属性总是返回 **False**，但如果通过自动功能运行，则可返回正确的值。例如，如果您在 ET 中运行示例，则处于设计模式时将返回 **True**。

当 WPS 处于窗体设计模式时，将显示**“控件工具箱”**工具栏。使用该工具箱可以插入各种 ActiveX 控件，如命令按钮、滚动条和选项按钮。在窗体设计模式下，不会运行事件过程，并且单击一个嵌入的控件后，就会出现控件的尺寸控点。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一个消息框，以报告活动文档是否处于窗体设计模式。本示例总是返回 False。*/ MsgBox(ActiveDocument.FormsDesign)` |

#### **Document.FormsDesignFormsDesign**

如果指定的文档处于窗体设计模式，则该属性值为 **True**。**Boolean** 类型，只读。

**语法**

**express.FormsDesignFormsDesign**

*express*   一个代表 **Document** 对象的变量。

**说明**

如果在从 WPS 中运行的代码中使用 **FormsDesign** 属性，则本属性总是返回 **False**，但如果通过自动功能运行，则可返回正确的值。例如，如果您在 ET 中运行示例，则处于设计模式时将返回 **True**。

当 WPS 处于窗体设计模式时，将显示**“控件工具箱”**工具栏。使用该工具箱可以插入各种 ActiveX 控件，如命令按钮、滚动条和选项按钮。在窗体设计模式下，不会运行事件过程，并且单击一个嵌入的控件后，就会出现控件的尺寸控点。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示一个消息框，以报告活动文档是否处于窗体设计模式。本示例总是返回 False。*/ MsgBox(ActiveDocument.FormsDesign)` |

#### **Document.Frames**

返回一个 **Frames** 集合，该集合代表文档中的所有图文框。只读。

**语法**

**express.Frames**

*express*   一个代表 **Document** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在选定内容的四周添加一个图文框，并将图文框对象返回给 myFrame 变量。*/ let myFrame = ActiveDocument.Frames.Add(Selection.Range)` |

#### **Document.Frameset**

返回一个 **Frameset** 对象，该对象代表一个整体框架页或框架页的单个框架。只读。

**语法**

**express.Frameset**

*express*   一个代表 **Document** 对象的变量。

**说明**

有关创建框架页的详细信息，请参阅 创建框架页。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将指定框架页中框架边框的颜色设为茶色。*/ function test() { let frameset = ActiveWindow.Document.Frameset frameset.FramesetBorderColor = wdColorTan frameset.FramesetBorderWidth = 6 }` |

#### **Document.FullName**

返回一个 **String** 类型的值，该值代表包括路径的文档的名称。只读。

**语法**

**express.FullName**

*express*   一个代表 **Document** 对象的变量。

**说明**

使用该属性与按顺序使用 **Path**、**PathSeparator** 和 **Name** 属性的作用相同。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示活动文档的路径和文件名。*/ function DocName() {     MsgBox(ActiveDocument.FullName) }` |

#### **Document.GrammarChecked**

如果已经检查了指定范围或文档的语法，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.GrammarChecked**

*express*   一个代表 **Document** 对象的变量。

**说明**

如果指定文档的部分或全部没有经过语法检查，则返回 **False**。要重复检查文档的语法，可将 **GrammarChecked** 属性设为 **False**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例判断是否对活动文档进行了语法检查。如果已检查，则显示文档的字数。如果尚未检查语法，则开始进行拼写和语法检查。*/ function test() { let myStat = ActiveDocument.ReadabilityStatistics let passGram = ActiveDocument.GrammarChecked if(passGram == true) {     MsgBox(myStat.Item(1).Name + " - " + myStat.Item(1).Value) } else {     ActiveDocument.CheckGrammar() } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置活动文档的 GrammarChecked 属性为 False，然后再次进行语法检查。*/ function test() { ActiveDocument.GrammarChecked = false ActiveDocument.CheckGrammar() }` |

#### **Document.GrammaticalErrors**

返回一个 **ProofreadingErrors** 集合，该集合代表指定的文档中有语法检查错误的句子。只读。

**语法**

**express.GrammaticalErrors**

*express*   一个代表 **Document** 对象的变量。

**说明**

在每个句子中可有多个错误。如果没有语法错误，由 **GrammaticalErrors** 属性返回的 **ProofreadingErrors** 集合的 **Count** 属性的返回值为 0。

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例检查活动文档的语法错误。如果发现任何错误，则重新开始拼写和语法检查。*/ function test() { if(ActiveDocument.GrammaticalErrors.Count == 0) {     MsgBox("There are no grammatical errors.") } else {     ActiveDocument.CheckGrammar() } }` |

#### **Document.GridDistanceHorizontal**

在指定文档中绘制、移动和调整自选图形或东亚语言字符时，返回或设置代表 WPS 所用的不可见的网格线之间的水平距离的 **Single** 类型，可读写。

**语法**

**express.GridDistanceHorizontal**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例实现的功能是：设置网格线之间的水平和垂直间距，并为当前文档启用“对象与网格对齐”功能。*/ function test() { ActiveDocument.GridDistanceHorizontal = 9 ActiveDocument.GridDistanceVertical = 9 ActiveDocument.SnapToGrid = true }` |

#### **Document.GridDistanceVertical**

在指定的文档中绘制、移动和调整自选图形或东亚语言字符时，返回或设置代表 WPS 所用的不可见网格线之间的垂直距离的 **Single** 类型。可读写。

**语法**

**express.GridDistanceVertical**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例实现的功能是：设置网格线之间的水平和垂直间距，并为当前文档启用“对象与网格对齐”功能。*/ function test() { ActiveDocument.GridDistanceHorizontal = 9 ActiveDocument.GridDistanceVertical = 9 ActiveDocument.SnapToGrid = true }` |

#### **Document.GridOriginFromMargin**

如果 WPS 将在页面左上角开始使用字符网格，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.GridOriginFromMargin**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 从活动文档页面的左上角开始使用字符网格。*/ ActiveDocument.GridOriginFromMargin = true` |

#### **Document.GridOriginHorizontal**

返回或设置一个 **Single** 类型的值，该值代表不可见网格相对于页面左边的起点位置，当在指定的文档中绘制、移动和调整自选图形或东亚语言字符时将使用该网格。可读写。

**语法**

**express.GridOriginHorizontal**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例实现的功能是：设置网格的水平和垂直起点，以及网格线之间的水平和垂直距离，并为当前文档启用“对齐网格”功能。*/ function test() { let ment = ActiveDocument ment.GridOriginHorizontal = 80 ment.GridOriginVertical = 90 ment.GridDistanceHorizontal = 9 ment.GridDistanceVertical = 9 ment.SnapToGrid = true }` |

#### **Document.GridOriginVertical**

返回或设置一个 **Single** 类型的值，该值代表不可见网格相对于页面顶边的起点位置，当在指定的文档中绘制、移动和调整自选图形或东亚语言字符时将使用该网格。可读写。

**语法**

**express.GridOriginVertical**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例实现的功能是：设置网格的水平和垂直起点，以及网格线之间的水平和垂直距离，并为当前文档启用“对象与网格对齐”功能。*/ function test() { let ment = ActiveDocument ment.GridOriginHorizontal = 80 ment.GridOriginVertical = 90 ment.GridDistanceHorizontal = 9 ment.GridDistanceVertical = 9 ment.SnapToGrid = true }` |

#### **Document.GridSpaceBetweenHorizontalLines**

返回或设置 WPS 在页面视图中显示的水平字符网格的间隔。**Long** 类型，可读写。

**语法**

**express.GridSpaceBetweenHorizontalLines**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 每隔四格显示一个水平字符网格线。*/ ActiveDocument.GridSpaceBetweenHorizontalLines = 5` |

#### **Document.GridSpaceBetweenVerticalLines**

返回或设置 WPS 在页面视图中显示的垂直字符网格线间隔。**Long** 类型，可读写。

**语法**

**express.GridSpaceBetweenVerticalLines**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 每隔一格显示一个垂直字符网格线。*/ ActiveDocument.GridSpaceBetweenVerticalLines = 2` |

#### **Document.HasMailer**

您请求就仅在 Macintosh 上使用的 Visual Basic 关键字获得帮助。有关 **Document** 对象的 **HasMailer** 属性的信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

**语法**

**express.HasMailer**

*express*   一个代表 **Document** 对象的变量。

#### **Document.HasPassword**

如果需要密码才能打开指定文档，则该属性值为 **True**。**Boolean** 类型，只读。

**语法**

**express.HasPassword**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例首先将“kittycat”设置为活动文档的密码，然后显示确认信息。*/ function test() { ActiveDocument.Password = "kittycat" if(ActiveDocument.HasPassword == true) {     MsgBox("The password is set.") } }` |

#### **Document.HasVBProject**

返回一个 **Boolean** 类型的值，该值代表文档是否具有附加的 示例代码 项目。只读。

**语法**

**express.HasVBProject**

*express*   一个代表 **Document** 对象的变量。

**说明**

此属性非常有用，可用来以编程方式确定是否需要将文档保存为支持宏的文件格式。如果用其他格式保存，则文档中包含的宏和代码项目可能丢失。

#### **Document.HTMLDivisions**

返回一个 **HTMLDivisions** 集合，该集合代表 Web 文档中的 HTML DIV 元素。

**语法**

**express.HTMLDivisions**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置活动文档中三个嵌套的区域的格式。本示例假定该活动文档为具有至少三个区域的 HTML 文档。*/ function FormatHTMLDivisions() {     let htmldivisions = ActiveDocument.HTMLDivisions.Item(1)     let htmldivisions1 = htmldivisions.HTMLDivisions.Item(1)     let htmldivisions2 = htmldivisions1.HTMLDivisions.Item(1)     let borderleft = htmldivisions.Borders.Item(wdBorderLeft)     let borderright = htmldivisions.Borders.Item(wdBorderRight)     let bordertop = htmldivisions1.Borders.Item(wdBorderTop)     let borderbottom = htmldivisions1.Borders.Item(wdBorderBottom)     borderleft.Color = wdColorRed     borderleft.LineStyle = wdLineStyleSingle     borderright.Color = wdColorRed     borderright.LineStyle = wdLineStyleSingle     htmldivisions1.LeftIndent = InchesToPoints(1)     htmldivisions1.RightIndent = InchesToPoints(1)     bordertop.Color = wdColorBlue     bordertop.LineStyle = wdLineStyleDouble     borderbottom.Color = wdColorBlue     borderbottom.LineStyle = wdLineStyleDouble     htmldivisions2.LeftIndent = InchesToPoints(1)     htmldivisions2.RightIndent = InchesToPoints(1)     htmldivisions2.Borders.Item(wdBorderLeft).LineStyle = wdLineStyleDot     htmldivisions2.Borders.Item(wdBorderRight).LineStyle = wdLineStyleDot     htmldivisions2.Borders.Item(wdBorderTop).LineStyle = wdLineStyleDot     htmldivisions2.Borders.Item(wdBorderBottom).LineStyle = wdLineStyleDot }` |

#### **Document.Hyperlinks**

返回一个 **Hyperlinks** 集合，该集合代表指定文档中的所有超链接。只读。

**语法**

**express.Hyperlinks**

*express*   一个代表 **Document** 对象的变量。

**说明**

有关返回集合的单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示 Home.doc 中第二个超链接的目标地址。*/ function test() { if(Documents.Item("Home.doc").Hyperlinks.Count >= 2) {     MsgBox(Documents.Item("Home.doc").Hyperlinks.Item(2).Name) } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示活动文档中地址包含“Microsoft”一词的每个超链接的名字。*/ function test() { let aHyperlink = ActiveDocument.Hyperlinks for(let i = 1; i <= aHyperlink.Count; i++) {     if((aHyperlink.Item(i).Address.toLowerCase().indexOf("microsoft") != -1) != 0) {         MsgBox(aHyperlink.Item(i).Name) } } }` |

#### **Document.HyphenateCaps**

如果全是大写字母时可对其进行断字，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.HyphenateCaps**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例自动为指定的文档断字，并允许对全是大写字母的单词进行断字。*/ function test() { ActiveDocument.AutoHyphenation = true ActiveDocument.HyphenateCaps = true }` |

#### **Document.HyphenationZone**

返回或设置断字区的宽度（以磅为单位）。**Long** 类型，可读写。

**语法**

**express.HyphenationZone**

*express*   一个代表 **Document** 对象的变量。

**说明**

断字区是 WPS 在一行最后一个单词与右边距间的保留的最大距离。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为 MyReport.doc 启用自动断字功能。并将断字区的宽度设置为 36 磅（0.5 英寸）。*/ function test() { Documents.Item("MyReport.doc").AutoHyphenation = true Documents.Item("MyReport.doc").HyphenationZone = 36 }  ` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将断字区的宽度设置为 0.25 英寸（18 磅），并启用活动文档的手动断字功能。*/ function test() { ActiveDocument.HyphenationZone = InchesToPoints(0.25) ActiveDocument.ManualHyphenation() }` |

#### **Document.Indexes**

返回一个 **Indexes** 集合，该集合代表指定文档的所有索引。只读。

**语法**

**express.Indexes**

*express*   一个代表 **Document** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在活动文档的末尾添加索引。*/ function test() { let MyRange = ActiveDocument.Range(ActiveDocument.Content.End - 1, ActiveDocument.Content.End - 1) ActiveDocument.Indexes.Add(MyRange, false, null, null, 1) }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例插入关于选定文本的索引项。*/ function test() { if(Selection.Type == wdSelectionNormal) {     ActiveDocument.Indexes.MarkEntry(Selection.Range, Selection.Range.Text) } }` |

#### **Document.InlineShapes**

返回一个 **InlineShapes** 集合，该集合代表文档中的所有 **InlineShape** 对象。只读。

**语法**

**express.InlineShapes**

*express*   一个代表 **Document** 对象的变量。

**说明**

有关返回集合的单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例显示活动文档中形状和嵌入式形状的数目 */ Application.ActiveDocument.InlineShapes.Count` |

#### **Document.IsMasterDocument**

如果指定的文档为一篇主文档，则该属性值为 **True**。**Boolean** 类型，只读。

**语法**

**express.IsMasterDocument**

*express*   一个代表 **Document** 对象的变量。

**说明**

主文档包括一篇或多篇子文档。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果活动文档是一篇主文档，则本示例切换至主控文档视图，然后打开第一篇子文档。*/ function test() { if(ActiveDocument.IsMasterDocument == true) {     ActiveDocument.ActiveWindow.View.Type = wdMasterView     ActiveDocument.Subdocuments.Item(1).Open() } else {     MsgBox("This document is not a master document.") } }` |

#### **Document.IsSubdocument**

如果指定文档是主文档的子文档，则该属性值为 **True**。**Boolean** 类型，只读。

**语法**

**express.IsSubdocument**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例判断 Sales.doc 是否为子文档，然后显示一条消息，表明 Sales.doc 的状态。*/ function test() { if(Documents.Item("Sales.doc").IsSubdocument == true) {     MsgBox("Sales.doc is a subdocument.") } else {     MsgBox("Sales.doc is not a subdocument.") } }` |

#### **Document.JustificationMode**

返回或设置指定文档字符间距的调整量。可读/写 **WdJustificationMode** 类型。

**语法**

**express.JustificationMode**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设定 WPS 在调整字符间距时只压缩标点符号。*/ ActiveDocument.JustificationMode = wdJustificationModeCompressKana` |

#### **Document.KerningByAlgorithm**

如果 WPS 在指定文档中调整半角拉丁字符和标点符号的间距，则为 **True**。可读/写 **Boolean** 类型。

**语法**

**express.KerningByAlgorithm**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设定 WPS 调整活动文档中的半角拉丁字符与标点符号的间距。*/ ActiveDocument.KerningByAlgorithm = true` |

#### **Document.Kind**

返回或设置 WPS 在自动设置指定文档的格式时使用的格式类型。**WdDocumentKind** 类型，可读写。

**语法**

**express.Kind**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例询问用户活动文档是否为电子邮件。如果是，则将此文档的格式设置为电子邮件格式。*/ function test() { let response = MsgBox("Is this document an e-mail message?", vbYesNo) if(response == vbYes) {     ActiveDocument.Kind = wdDocumentEmail     ActiveDocument.Content.AutoFormat() } }` |

#### **Document.LanguageDetected**

返回或设置一个指定 WPS 是否对指定文本的语言进行检测的值。**Boolean** 类型，可读写。

**语法**

**express.LanguageDetected**

*express*   一个代表 **Document** 对象的变量。

**说明**

检查以前任何语言检测结果的 **LanguageID** 属性。

调用 **DetectLanguage** 方法时，**LanguageDetected** 属性会设置为 **True**。要重新评估指定文本的语言，请将 **LanguageDetected** 属性设置为 **False**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例检查活动文档，以确定其所用的语言类型并显示检查结果。*/ function test() { if(ActiveDocument.LanguageDetected == true) {     x = MsgBox("This document has already " + "been checked. Do you want to check " + "it again?", vbYesNo)     if(x == vbYes) {         ActiveDocument.LanguageDetected = false         ActiveDocument.DetectLanguage() } } else {     ActiveDocument.DetectLanguage() } if(ActiveDocument.Range.LanguageID == wdEnglishUS) {     MsgBox("This is a U.S. English document.") } else {     MsgBox("This is not a U.S. English document.") } }` |

#### **Document.ListParagraphs**

返回一个 **ListParagraphs** 对象，该对象代表文档中的所有编号段落。只读。

**语法**

**express.ListParagraphs**

*express*   一个代表 **Document** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的单个对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例为第一个文档中的每个编号段落或有项目符号的段落添加黄色背景。*/ function test() { let numpar = Documents.Item(1).ListParagraphs for(let i = 1; i <= numpar.Count; i++) {     numpar.Item(i).Shading.BackgroundPatternColorIndex = wdYellow } }` |

#### **Document.Lists**

返回一个 **Lists** 集合，该集合含有指定文档中的所有项目符号和编号列表。只读。

**语法**

**express.Lists**

*express*   一个代表 **Document** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置所选内容的格式为编号列表，然后在一个消息框中显示活动文档中列表数。*/ function test() { Selection.Range.ListFormat.ApplyListTemplate(ListGalleries.Item(wdNumberGallery).ListTemplates.Item(2)) MsgBox("This document has " + ActiveDocument.Lists.Count + " lists.") }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置活动文档中第三个列表的格式为默认的项目符号列表格式。如果该列表已设置为项目符号列表格式，该示例清除此格式。*/ function test() { if(ActiveDocument.Lists.Count >= 3) {     ActiveDocument.Lists.Item(3).Range.ListFormat.ApplyBulletDefault() } }  ` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在一个消息框中显示 MyLetter.doc 的每一列表的项目数。*/ function test() { let myDoc = Documents.Item("MyLetter.doc") let i = myDoc.Lists.Count for(let j = 1; j <= i; j++) {     MsgBox("List " + i + " has " + myDoc.Lists.Item(j).CountNumberedItems() + " items.")     i = i - 1 } }` |

#### **Document.ListTemplates**

返回一个 **ListTemplates** 集合，该集合代表指定文档中的所有列表格式。只读。

**语法**

**express.ListTemplates**

*express*   一个代表 **Document** 对象的变量。

**说明**

有关返回集合的单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示活动文档中使用的列表模板数量。*/ MsgBox(ActiveDocument.ListTemplates.Count)` |

#### **Document.LockQuickStyleSet**

返回或设置一个 **Boolean**，它代表用户是否可以更改使用的快速样式集。可读写。

**语法**

**express.LockQuickStyleSet**

*express*   一个代表 **Document** 对象的变量。

#### **Document.LockTheme**

返回或设置一个 **Boolean** 类型的值，该值代表用户是否可以更改文档主题。可读/写。

**语法**

**express.LockTheme**

*express*   一个代表 **Document** 对象的变量。

**说明**

使用 **ApplyDocumentTheme** 方法可应用 WPS Office主题。使用 **DocumentTheme** 属性可返回应用于文档的 Office 主题。

#### **Document.MailEnvelope**

返回一个 **MsoEnvelope** 对象，该对象代表文档的电子邮件标题。

**语法**

**express.MailEnvelope**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为活动文档的电子邮件标题设置批注。*/ function HeaderComments(){     ActiveDocument.MailEnvelope.Introduction = "Please review this document and let me know " + "what you think.  I need your input by Friday." + "  Thanks." }` |

#### **Document.Mailer**

您请求就仅在 Macintosh 上使用的 Visual Basic 关键字获得帮助。有关 **Document** 对象的 **Mailer** 属性的信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

**语法**

**express.Mailer**

*express*   一个代表 **Document** 对象的变量。

#### **Document.MailMerge**

返回一个 **MailMerge** 对象，该对象代表指定文档的邮件合并功能。只读。

**语法**

**express.MailMerge**

*express*   一个代表 **Document** 对象的变量。

**说明**

不管指定文档是否是一个邮件合并主文档，**MailMerge** 对象都有效。用 **State** 属性可确定邮件合并操作当前的状态。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果活动文档是一个带有附加数据源的主文档，则本示例执行邮件合并。*/ function test() { let myMerge = ActiveDocument.MailMerge if(myMerge.State == wdMainAndDataSource) {     myMerge.Execute() } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将记录 1 到 4 与主文档合并，并将合并后的文档发送到打印机。*/ function test() { let mailmerge = ActiveDocument.MailMerge     mailmerge.DataSource.FirstRecord = 1     mailmerge.DataSource.LastRecord = 4     mailmerge.Destination = wdSendToPrinter     mailmerge.SuppressBlankLines = true     mailmerge.Execute()  }  ` |

#### **Document.Name**

返回指定文档的名称。**String** 类型，只读。

**语法**

**express.Name**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 返回当前文档的文件名 */ Application.ActiveDocument.FullName` |

#### **Document.NoLineBreakAfter**

返回或设置 WPS 不在其后换行的首尾字符。可读/写 **String** 类型。

**语法**

**express.NoLineBreakAfter**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例实现的功能是：在活动文档中，设定以“$”、“(”、“[”、“\”和“{”作为首尾字符时，WPS 不在其后面进行换行。*/ ActiveDocument.NoLineBreakAfter = "$([\{"` |

#### **Document.NoLineBreakBefore**

返回或设置 WPS 不在其前换行的首尾字符。可读/写 **String** 类型。

**语法**

**express.NoLineBreakBefore**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在活动文档中设定，以“!”、“)”和“]”作为首尾字符时，WPS 不在这些首尾字符的前面进行换行。*/ ActiveDocument.NoLineBreakBefore = "!)]"` |

#### **Document.OMathBreakBin**

返回或设置一个 **WdOMathBreakBin** 常量，该常量代表当公式跨越两行或多行时 WPS 用来放置二元运算符的位置。可读写。

**语法**

**express.OMathBreakBin**

*express*   一个代表 **Document** 对象的变量。

**说明**

如果公式在二元运算符（如加号、减号或乘号）处换行，则该运算符有三种不同的放置方式：换行符之前、换行符之后以及在换行符前后重复。

如果此属性设置为 **wdOMathBreakBinRepeat**，则可以使用 **OMathBreakSub** 属性指定 WPS 如何处理位于换行符之前的减法运算符。

#### **Document.OMathBreakSub**

返回或设置一个 **WdOMathBreakSub** 常量，该常量代表 WPS 如何处理位于换行符之前的减法运算符。可读写。

**语法**

**express.OMathBreakSub**

*express*   一个代表 **Document** 对象的变量。

**说明**

此属性仅在 **OMathBreakBin** 属性设置为 **wdOMathBreakBinRepeat** 时使用。当换行符刚好位于减法运算符之后，且文档设置为在下一行重复减法运算符时，由于负负得正的缘故，有时需要特殊对待减法运算符。有些作者选择将其中一个减号转变为加号，而有些作者则选择保留两个减号。

#### **Document.OMathFontName**

返回或设置一个 **String** 类型的值，该值代表文档中用于显示公式的字体的名称。可读写。

**语法**

**express.OMathFontName**

*express*   一个代表 **Document** 对象的变量。

#### **Document.OMathIntSubSupLim**

返回或设置一个 **Boolean** 类型的值，该值代表积分的默认极限位置。可读写。

**语法**

**express.OMathIntSubSupLim**

*express*   一个代表 **Document** 对象的变量。

#### **Document.OMathJc**

返回或设置一个 **WdOMathJc** 常量，该常量代表一组公式的默认对齐方式：左对齐、右对齐、居中或作为一个组居中。可读写。

**语法**

**express.OMathJc**

*express*   一个代表 **Document** 对象的变量。

#### **Document.OMathLeftMargin**

返回或设置一个代表公式的左边距的 **Single**。可读写。

**语法**

**express.OMathLeftMargin**

*express*   一个代表 **Document** 对象的变量。

#### **Document.OMathNarySupSubLim**

返回或设置一个 **Boolean** 类型的值，该值代表 *n* 元对象（而不是积分）的默认极限位置。可读写。

**语法**

**express.OMathNarySupSubLim**

*express*   一个代表 **Document** 对象的变量。

#### **Document.OMathRightMargin**

返回或设置一个 **Single** 类型的值，该值代表公式的右边距。可读写。

**语法**

**express.OMathRightMargin**

*express*   一个代表 **Document** 对象的变量。

#### **Document.OMaths**

返回一个 **OMaths** 集合，该集合代表指定区域内的 **OMath** 对象。只读。

**语法**

**express.OMaths**

*express*   一个代表 **Document** 对象的变量。

#### **Document.OMathSmallFrac**

返回或设置一个 **Boolean** 类型的值，该值代表是否在文档所包含的公式中使用小型分数。可读写。

**语法**

**express.OMathSmallFrac**

*express*   一个代表 **Document** 对象的变量。

#### **Document.OMathWrap**

返回或设置一个 **Single** 类型的值，该值代表换到新的一行的公式的第二行位置。可读写。

**语法**

**express.OMathWrap**

*express*   一个代表 **Document** 对象的变量。

**说明**

如果值为 -1，则指定公式的第二行为右对齐。如果为正值，则表示从左边距缩进。

#### **Document.OpenEncoding**

返回用于打开指定文档的编码。**MsoEncoding** 类型，只读。

**语法**

**express.OpenEncoding**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例实现的功能是：测试当前文档是否使用编码 UTF7 打开。*/ function test() { if(ActiveDocument.OpenEncoding == msoEncodingUTF7){     MsgBox("This is a UTF7-encoded text file!") } else{     MsgBox("This is not a UTF7-encoded text file!") } }` |

#### **Document.OptimizeForWord97**

如果 WPS 通过禁用所有不兼容的格式来优化当前文档，以便在 WPS 97 中查看，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.OptimizeForWord97**

*express*   一个代表 **Document** 对象的变量。

**说明**

若要为 WPS 97 按默认设置优化所有的新文档，可用 **OptimizeForWord97byDefault** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例的功能如下：检查当前文档是否针对 WPS 97 进行过优化。如果没有，则询问用户是否进行优化。*/ function test() { if(ActiveDocument.OptimizeForWord97 == false) {     let x = MsgBox("Is this document targeted at WPS 97 users?", jsYesNo)     if (x == jsResultYes) {         ActiveDocument.OptimizeForWord97 = true     } } }` |

#### **Document.OriginalDocumentTitle**

返回一个 **String** 类型的值，该值代表在运行“精确比较”文档比较功能之后原文档的文档标题。只读。

**语法**

**express.OriginalDocumentTitle**

*express*   一个代表 **Document** 对象的变量。

**说明**

要执行“精确比较”文档比较，请使用 **CompareDocuments** 方法。

#### **Document.PageSetup**

返回一个与指定文档相关联的 **PageSetup** 对象。

**语法**

**express.PageSetup**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例将活动文档的右边距设置为 72 磅（1 英寸） */ Application.ActiveDocument.PageSetup.RightMargin = Application.InchesToPoints(2)` |

#### **Document.Paragraphs**

返回一个 **Paragraphs** 集合，该集合代表指定文档中的所有段落。只读。

**语法**

**express.Paragraphs**

*express*   一个代表 **Document** 对象的变量。

**说明**

有关返回集合的单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 将当前文档首段设置为双倍行距 */ Application.ActiveDocument.Paragraphs.Item(1).LineSpacingRule = wdLineSpaceDouble` |

#### **Document.Parent**

**语法**

**express.Parent**

*express*   一个代表 **Document** 对象的变量。

**说明**

返回一个 **Object** 类型值，该值代表指定 **Document** 对象的父对象。

#### **Document.Password**

设置在打开文档必须使用的密码。**String** 类型，只写。

**语法**

**express.Password**

*express*   一个代表 **Document** 对象的变量。

**说明**

**要点**??请避免在应用程序中使用硬编码的密码。如果过程中要求输入密码，请从用户处请求密码，并将密码作为变量存储起来，然后在代码中使用该变量。有关如何做到这一点的建议最佳操作，请参阅 WPS Office解决方案开发人员安全注意事项。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例首先打开文档 Earnings.doc 并为其设置密码，然后关闭该文档。*/ function test() { let myDoc = Documents.Open("C:\\My Documents\\Earnings.doc") myDoc.Password = "strPassword" myDoc.Close() }` |

#### **Document.PasswordEncryptionAlgorithm**

返回一个 **String** 类型的值，该值代表 WPS 用密码加密文档时所用的算法。只读。

**语法**

**express.PasswordEncryptionAlgorithm**

*express*   一个代表 **Document** 对象的变量。

**说明**

使用 **SetPasswordEncryptionOptions** 方法指定 WPS 用密码加密文档时所用的算法。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果当前使用的密码加密算法是 WPS 97 for Windows 以前版本所使用的加密算法“OfficeXor”，则本示例将重新设置密码加密选项。*/ function PasswordSettings(){     let ad = ActiveDocument     if(ad.PasswordEncryptionAlgorithm == "OfficeXor"){         ad.SetPasswordEncryptionOptions("Microsoft RSA SChannel Cryptographic Provider", "RC4", 56, true)     } }` |

#### **Document.PasswordEncryptionFileProperties**

如果 WPS 为密码保护的文档的文件属性加密，则该属性值为 **True**。**Boolean** 类型，只读。

**语法**

**express.PasswordEncryptionFileProperties**

*express*   一个代表 **Document** 对象的变量。

**说明**

使用 **SetPasswordEncryptionOptions** 方法指定 WPS 是否将密码保护的文档的文件属性加密。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果未对密码保护的文档的文件属性加密，本示例将设置密码加密选项。*/ function PasswordSettings(){     let ad = ActiveDocument     if(ad.PasswordEncryptionFileProperties == false){         ad.SetPasswordEncryptionOptions("Microsoft RSA SChannel Cryptographic Provider", "RC4", 56, true)     } }` |

#### **Document.PasswordEncryptionKeyLength**

返回一个 **Long** 类型的值，该值代表 WPS 用密码加密文档时所用的密码键长。只读。

**语法**

**express.PasswordEncryptionKeyLength**

*express*   一个代表 **Document** 对象的变量。

**说明**

使用 **SetPasswordEncryptionOptions** 方法指定 WPS 用密码加密文档时所用的密码键长。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果密码加密键长小于 40，则本示例将设置密码加密选项。*/ function PasswordSettings(){     let ad = ActiveDocument     if(ad.PasswordEncryptionKeyLength < 40){         ad.SetPasswordEncryptionOptions("Microsoft RSA SChannel Cryptographic Provider", "RC4", 56, true)     } }` |

#### **Document.PasswordEncryptionProvider**

返回一个 **String** 类型的值，该值指定 WPS 用密码加密文档时所用的加密算法提供程序的名称。只读。

**语法**

**express.PasswordEncryptionProvider**

*express*   一个代表 **Document** 对象的变量。

**说明**

使用 **SetPasswordEncryptionOptions** 方法指定 WPS 用密码加密文档时所用的加密算法提供程序的名称。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果当前使用的密码加密算法提供程序不是“Microsoft RSA SChannel Cryptographic Provider”，则本示例将设置密码加密选项。*/ function PasswordSettings(){     let ad = ActiveDocument     if(ad.PasswordEncryptionProvider != "Microsoft RSA SChannel Cryptographic Provider"){         ad.SetPasswordEncryptionOptions("Microsoft RSA SChannel Cryptographic Provider", "RC4", 56, true)     } }` |

#### **Document.Path**

返回文档的磁盘或 Web 路径。只读 **String** 类型。

**语法**

**express.Path**

*express*   一个代表 **Document** 对象的变量。

**说明**

此路径不包括尾随字符，例如，“C:\WPSOffice”或“http://MyServer”。使用 **PathSeparator** 属性可添加分隔文件夹与驱动器号的字符。使用 **Name** 属性可返回不带路径的文件名，而使用 **FullName** 属性可同时返回文件名和路径。

| 注释                                                         |
| ------------------------------------------------------------ |
| 可以使用 **PathSeparator** 属性建立 Web 地址，即使地址中包含正斜杠 (/)（**PathSeparator** 属性默认使用反斜杠 (\)）。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示活动文档的路径和文件名。*/ MsgBox(ActiveDocument.Path + Application.PathSeparator + ActiveDocument.Name)` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将当前文件夹改为附加到活动文档的模板的路径。*/ ChDir ActiveDocument.AttachedTemplate.Path` |

#### **Document.Permission**

返回一个 **Permission** 对象，该对象代表指定文档中的权限设置。

**语法**

**express.Permission**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下例返回活动文档的权限设置。*/ let objPermission = ActiveDocument.Permission` |

#### **Document.PrintFormsData**

如果 WPS 仅将相应联机窗体中的键入的数据打印在预打印窗体上，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.PrintFormsData**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 仅打印联机窗体域的数据内容，然后开始打印活动文档。*/ function test() { ActiveDocument.PrintFormsData = true ActiveDocument.PrintOut() }  ` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“选项”对话框“打印”选项卡上“只用于当前文档的选项”区域中“仅打印窗体域的内容”复选框当前的状态。*/ let temp = ActiveDocument.PrintFormsData` |

#### **Document.PrintFractionalWidths**

如果设置指定文档的格式，使用分式间距来显示和打印字符，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.PrintFractionalWidths**

*express*   一个代表 **Document** 对象的变量。

**说明**

**PrintFractionalWidths** 属性只适用于 Apple Macintosh 计算机。在 Windows 中，**PrintFractionalWidths** 属性始终返回 **False**。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 的语言参考帮助。

#### **Document.PrintPostScriptOverText**

如果为 **True**，则使用 PostScript 打印机时，文档中的 PRINT 域指令（例如，PostScript 命令）将打印在文字和图形的上方。**Boolean** 类型，可读写。

**语法**

**express.PrintPostScriptOverText**

*express*   一个代表 **Document** 对象的变量。

**说明**

**PrintPostScriptOverText** 属性控制是否将 Macintosh 文档中的 postscript 代码以转换的 WPS 形式打印。如果文档不包含 PRINT 域，则该属性无效。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 可在文字和图形的上方打印 PRINT 域指令，然后打印该活动文档。*/ function test() { ActiveDocument.PrintPostScriptOverText = true ActiveDocument.PrintOut() }  ` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“选项”对话框“打印”选项卡上“打印选项”区域中的“在文本上方打印 postscript”复选框当前的状态。*/ let currSet = ActiveDocument.PrintPostScriptOverText` |

#### **Document.PrintRevisions**

如果为 **True**，则在打印文档的同时打印修订标记。如果返回 **False**，则不打印修订标记（即打印接受修订后的状态）。**Boolean** 类型，可读写。

**语法**

**express.PrintRevisions**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在打印活动文档时不带有修订标记。*/ function test() { ActiveDocument.PrintRevisions = false ActiveDocument.PrintOut() }` |

#### **Document.Protect**

返回或设置与指定传送名单相关联的文档的保护类型。**WdProtectionType** 类型，可读写。

**语法**

**express.Protect**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例先指定用于活动文档（只允许批注）的保护类型，然后传送该文档。*/ function test() { ActiveDocument.HasRoutingSlip = true let RoutingSlip = ActiveDocument.RoutingSlip     RoutingSlip.Subject = "Status Doc"     RoutingSlip.Protect = wdAllowOnlyComments     RoutingSlip.AddRecipient("Kim Johnson") ActiveDocument.Route() }` |

#### **Document.ProtectionType**

该属性返回指定文档的保护类型。可以是下列 **WdProtectionType** 常量之一：**wdAllowOnlyComments**、**wdAllowOnlyFormFields**、**wdAllowOnlyReading**、**wdAllowOnlyRevisions** 或 **wdNoProtection**。

**语法**

**express.ProtectionType**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果没有保护活动文档，该示例保护文档，使审阅者只能插入批注而不改变文档内容。*/ function test() { if(ActiveDocument.ProtectionType == wdNoProtection) {     ActiveDocument.Protect(wdAllowOnlyComments) } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果活动文档受到保护，则本示例解除文档的保护。*/ function test() { let doc = ActiveDocument if(doc.ProtectionType != wdNoProtection) {     doc.Unprotect() } }` |

#### **Document.ReadabilityStatistics**

返回一个 **ReadabilityStatistics** 集合，该集合代表指定文档或区域的可读性统计信息。只读。

**语法**

**express.ReadabilityStatistics**

*express*   一个代表 **Document** 对象的变量。

**说明**

有关返回集合的单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示文档 1 的每一种可读性统计信息及其值。*/ function test() { let rs = Documents.Item(1).ReadabilityStatistics for(let i = 1; i <= rs.Count; i++) {     MsgBox(rs.Item(i).Name + " - " + rs.Item(i).Value) } }` |

#### **Document.ReadingLayoutSizeX**

设置或返回一个 **Long**类型的值，该值表示当在阅读版式视图中显示文档并对其冻结了输入手写标记时文档中页面的宽度。

**语法**

**express.ReadingLayoutSizeX**

*express*   一个代表 **Document** 对象的变量。

**说明**

设置完 **ReadingLayoutSizeX** 和 **ReadingLayoutSizeY** 属性后，可使用 **ReadingModeLayoutFrozen** 属性以指定的高度和宽度显示页面。使用 **ReadingLayout** 属性可在阅读版式视图中显示文档。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在阅读版式视图中显示活动文档，然后设置所显示页面的大小。*/ function test() { ActiveWindow.View.ReadingLayout = true ActiveDocument.ReadingLayoutSizeX = 300 ActiveDocument.ReadingLayoutSizeY = 300 ActiveDocument.ReadingModeLayoutFrozen = true }` |

#### **Document.ReadingLayoutSizeY**

设置或返回一个 **Long** 类型的值，该值表示当在阅读版式视图中显示文档并对其冻结了输入手写标记时文档中页面的高度。

**语法**

**express.ReadingLayoutSizeY**

*express*   一个代表 **Document** 对象的变量。

**说明**

设置完 **ReadingLayoutSizeX** 和 **ReadingLayoutSizeY** 属性后，可使用 **ReadingModeLayoutFrozen** 属性以指定的高度和宽度显示页面。使用 **ReadingLayout** 属性可在阅读版式视图中显示文档。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在阅读版式视图中显示活动文档，然后设置所显示页面的大小。*/ function test() { ActiveWindow.View.ReadingLayout = true ActiveDocument.ReadingLayoutSizeX = 300 ActiveDocument.ReadingLayoutSizeY = 300 ActiveDocument.ReadingModeLayoutFrozen = true }` |

#### **Document.ReadingModeLayoutFrozen**

设置或返回一个 **Boolean** 类型的值，该值表示是否将阅读版式视图中显示的页面冻结为指定大小以向文档插入手写标记。

**语法**

**express.ReadingModeLayoutFrozen**

*express*   一个代表 **Document** 对象的变量。

**说明**

使用 **ReadingLayoutSizeX** 和 **ReadingLayoutSizeY** 属性可指定将阅读版式大小冻结以向文档插入手写标记时所显示的页面的大小。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在阅读版式视图中显示活动文档，然后设置所显示的页面的大小。*/ function test() { ActiveWindow.View.ReadingLayout = true ActiveDocument.ReadingLayoutSizeX = 300 ActiveDocument.ReadingLayoutSizeY = 300 ActiveDocument.ReadingModeLayoutFrozen = true }` |

#### **Document.ReadOnly**

如果对文档所作修改不能保存到原始文档中，则该属性为 **True**。只读 **Boolean** 类型。

**语法**

**express.ReadOnly**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在活动文档不是只读文档时保存该文档。*/ function test() { if(ActiveDocument.ReadOnly == false) {     ActiveDocument.Save() } }` |

#### **Document.ReadOnlyRecommended**

如果无论用户何时打开文档，WPS 都显示一个消息框，提示以只读方式打开，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.ReadOnlyRecommended**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例的功能如下：在打开文档时， WPS 建议以只读方式打开。*/ ActiveDocument.ReadOnlyRecommended = true` |

#### **Document.RemoveDateAndTime**

设置或返回一个 **Boolean** 类型的值，该值指示文档是否存储修订的日期和时间元数据。

**语法**

**express.RemoveDateAndTime**

*express*   一个代表 **Document** 对象的变量。

**说明**

如果为 **True**，则删除修订的日期和时间戳信息。如果为 **False**，则不删除修订的日期和时间戳信息。将 **RemoveDateAndTime** 属性与 **RemovePersonalInformation** 属性结合使用，可帮助删除文档属性中的个人信息。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例从活动文档中删除个人信息，并从文档的所有修订中删除日期和时间信息。*/ function test() { ActiveDocument.RemovePersonalInformation = true ActiveDocument.RemoveDateAndTime = true }` |

#### **Document.RemovePersonalInformation**

如果 WPS 在保存文档时从批注、修订和“属性”对话框中删除所有用户信息，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.RemovePersonalInformation**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置当前文档，在用户下一次保存文档时删除所有的个人信息。*/ function RemovePersonalInfo() {     ActiveDocument.RemovePersonalInformation = true }` |

#### **Document.Research**

返回一个 **Research** 对象，该对象代表文档的检索服务。只读。

**语法**

**express.Research**

*express*   一个代表 **Document** 对象的变量。

#### **Document.RevisedDocumentTitle**

返回一个 **String** 类型的值，该值代表在运行“精确比较”文档比较功能之后修订的文档的文档标题。只读。

**语法**

**express.RevisedDocumentTitle**

*express*   一个代表 **Document** 对象的变量。

**说明**

要执行“精确比较”文档比较，请使用 **CompareDocuments** 方法。

#### **Document.Revisions**

返回一个 **Revisions**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Revisions/Revisions%20.htm#jsObject_Revisions)集合，该集合代表文档或区域中的修订。只读。

**语法**

**express.Revisions**

*express*   一个代表 **Document** 对象的变量。

**说明**

有关返回集合的单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例显示活动文档第一节中的修订数 */ Application.ActiveDocument.Sections.Item(1).Range.Revisions.Count` |

#### **Document.Saved**

如果指定的文档或模板从上次保存后一直没有更改，则该属性值为 **True**。如果关闭文档时，WPS 提示保存对文档所做的更改，则该属性值为 **False**。**Boolean** 类型，可读写。

**语法**

**express.Saved**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在活动文档含有以前未保存的更改的情况下保存该文档。*/ function test() { if(ActiveDocument.Saved == false) {     ActiveDocument.Save() } }` |

#### **Document.SaveEncoding**

返回或设置保存文档时使用的编码。**MsoEncoding** 类型，可读写。

**语法**

**express.SaveEncoding**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例实现的功能是：指定用公历编码保存当前文档。*/ ActiveDocument.SaveEncoding = msoEncodingWestern` |

#### **Document.SaveFormat**

返回指定文档或文件转换器的文件格式。只读 **Long** 类型。

**语法**

**express.SaveFormat**

*express*   一个代表 **Document** 对象的变量。

**说明**

**SaveFormat** 属性将是一个指定外部文件转换器的唯一数字或一个 **WdSaveFormat** 常量。

如果对 **SaveAs** 方法的 *FileFormat* 参数使用 **SaveFormat** 属性的值，可以将文档保存为没有对应 **WdSaveFormat** 常量的某种文件格式。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果活动文档为 RTF 格式的文档，本示例将其另存为 WPS 文档。*/ function test() { if(ActiveDocument.SaveFormat == wdFormatRTF) {     ActiveDocument.SaveAs(null, wdFormatDocument) } }` |

#### **Document.SaveFormsData**

如果 WPS 将输入到一个窗体中的数据作为以制表位分隔的数据库记录保存，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.SaveFormsData**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 只保存输入到一个窗体中的数据。*/ ActiveDocument.SaveFormsData = true` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“选项”对话框“保存”选项卡上“保存选项”区域中“仅保存窗体域内容”复选框当前的状态。*/ let temp = ActiveDocument.SaveFormsData` |

#### **Document.SaveSubsetFonts**

如果 WPS 将嵌入的 TrueType 字体子集和文档一起保存，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.SaveSubsetFonts**

*express*   一个代表 **Document** 对象的变量。

**说明**

如果在一个文档中使用的 TrueType 字体少于 32 个字符，则 WPS 将在文档中嵌入 TrueType 字体子集（只包含所用的字体）；如果使用的 TrueType 字体多于 32 个字符，则 WPS 在文档中嵌入整个 TrueType 字体。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置一个名为“MyDoc”的文档只保存嵌入的 TrueType 字体子集（如果使用很少的字符），然后再保存“MyDoc”。*/ function test() { let md = Documents.Item("MyDoc")     md.EmbedTrueTypeFonts = true     md.SaveSubsetFonts = true     md.Save()  }` |

#### **Document.Scripts**

返回一个 **Scripts** 集合，该集合代表指定对象中的 HTML 脚本的集合。

**语法**

**express.Scripts**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示活动文档的第一个 Script 对象中的文本。*/ Debug.Print(ActiveDocument.Scripts.Item(1).ScriptText)` |

#### **Document.Sections**

返回一个 **Section** 集合，该集合代表指定文档中的节。只读。

**语法**

**express.Sections**

*express*   一个代表 **Document** 对象的变量。

**说明**

有关返回集合的单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例设置活动文档中所有节的页面方向 */ function test() { 	let sec = Application.ActiveDocument.Sections 	for(let i = 1; i<= sec.Count; i++) { 	    sec.Item(i).PageSetup.Orientation = wdOrientLandscape 	} }` |

#### **Document.Sentences**

返回一个 **Sentences** 集合，该集合代表文档中的所有句子。只读。

**语法**

**express.Sentences**

*express*   一个代表 **Document** 对象的变量。

**说明**

有关返回集合的单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例复制活动文档中的第一个句子。*/ ActiveDocument.Sentences.Item(1).Copy()` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例删除活动文档中的最后一个句子。*/ ActiveDocument.Sentences.Last.Delete()` |

#### **Document.ServerPolicy**

返回一个 **ServerPolicy** 对象，该对象代表为运行 SharePoint Server 2007 的服务器上存储的文档指定的策略。只读。

**语法**

**express.ServerPolicy**

*express*   一个代表 **Document** 对象的变量。

#### **Document.Shapes**

返回一个 **Shapes** 集合，该集合代表指定文档中的所有 **Shape** 对象。只读。

**语法**

**express.Shapes**

*express*   一个代表 **Document** 对象的变量。

**说明**

该集合可以包含绘图、形状、图片、OLE 对象、ActiveX 控件、文本对象和标注。有关返回集合的单个成员的信息，请参阅 返回集合中的对象。

**Shapes** 属性应用于文档时将返回文档主体部分的所有 **Shape** 对象，不包括页眉和页脚。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例新建一篇文档，为其添加一个矩形 */ Application.Documents.Add().Shapes.AddShape(msoShapeRectangle, 5, 25, 100, 50)` |

#### **Document.ShowGrammaticalErrors**

如果指定文档中的语法错误用绿色波浪线标记，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.ShowGrammaticalErrors**

*express*   一个代表 **Document** 对象的变量。

**说明**

只有将 **CheckGrammarAsYouType** 属性设为 **True**，才能查看文档中的语法错误。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例对 WPS 进行设置，使其在您键入时检查活动文档中的语法错误并显示发现的所有错误。*/ function test() { Options.CheckGrammarAsYouType = true ActiveDocument.ShowGrammaticalErrors = true }` |

#### **Document.ShowRevisions**

如果在屏幕上显示对指定文档的修订，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.ShowRevisions**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例对活动文档进行设置，以便能标记对该文档的修改，并在屏幕上显示修订标记。*/ function test() { ActiveDocument.TrackRevisions = true ActiveDocument.ShowRevisions = true  }` |

#### **Document.ShowSpellingErrors**

如果 WPS 为文档中的拼写错误添加下划线，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.ShowSpellingErrors**

*express*   一个代表 **Document** 对象的变量。

**说明**

必须将 **CheckSpellingAsYouType** 属性设置为 **True**，才能查看文档中的拼写错误。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 隐藏活动文档中的拼写错误标记，该标记为红色波浪线。*/ ActiveDocument.ShowSpellingErrors = false` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示活动文档中的拼写错误标记。*/ function test() { Options.CheckSpellingAsYouType = true ActiveDocument.ShowSpellingErrors = true }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回“选项”对话框中“拼写和语法”选项卡的“拼写”区域中的“隐藏文档中的拼写错误”复选框的当前状态。*/ let temp = ActiveDocument.ShowSpellingErrors` |

#### **Document.Signatures**

返回一个 **SignatureSet** 集合，该集合代表文档的数字签名。

**语法**

**express.Signatures**

*express*   一个代表 **Document** 对象的变量。

**说明**

若要对 WPS 文档进行数字签名，并验证文档中的其他签名，您需要 Microsoft CryptoAPI 和唯一的数字签名证书。CryptoAPI 随着 Microsoft Internet Explorer 4.01 或更高版本一起安装。可从证书颁发机构获得数字签名证书。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示“签名”对话框，可在其中为文档添加数字签名。*/ function AddSignature() {     ActiveDocument.Signatures.Add() }` |

#### **Document.SmartDocument**

返回一个 **SmartDocument** 对象，该对象代表智能文档解决方案的设置。

**语法**

**express.SmartDocument**

*express*   一个代表 **Document** 对象的变量。

**说明**

有关智能文档的详细信息，请参阅 Microsoft Developer Network (MSDN) 网站上的 Smart Document Software Development Kit。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下列示例显示一个对话框，其中包含用于活动文档的有效 XML 扩展包列表。*/ ActiveDocument.SmartDocument.PickSolution()` |

#### **Document.SmartTagsAsXMLProps**

如果该属性的值为 **True**，当将包含智能标记的文档保存为 HTML 格式时，WPS 会创建一个包含智能标记信息的 XML 页眉。**Boolean** 类型，可读写。

**语法**

**express.SmartTagsAsXMLProps**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果活动文档被保存为 HTML 格式，本示例可以在 XML 页眉中保存智能标记信息。*/ function SaveXMLForSmartTags() {     ActiveDocument.SmartTagsAsXMLProps = true }` |

#### **Document.SnapToGrid**

如果在指定文档中绘制、移动自选图形/东亚字符或者调整它们的大小时，自选图形/东亚字符自动与不可见的网格线对齐，则该属性为 **True**。可读/写 **Boolean** 类型。

**语法**

**express.SnapToGrid**

*express*   一个代表 **Document** 对象的变量。

**说明**

绘制、移动自选图形或调整其大小时，按下 Alt 可以暂时使该设置无效。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS ，使当前文档中的东亚字符自动与不可见的网格线对齐。*/ ActiveDocument.SnapToGrid = true` |

#### **Document.SnapToShapes**

如果 WPS 在指定文档中使自选图形或东亚字符自动与不可见网格线对齐，这些网格线穿过其他自选图形或东亚字符的水平和垂直边缘，则该属性为 **True**。可读/写 **Boolean** 类型。

**语法**

**express.SnapToShapes**

*express*   一个代表 **Document** 对象的变量。

**说明**

该属性为每个自选图形创建附加的不可见网格线。**SnapToShapes** 属性与 **SnapToGrid** 属性独立工作。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS ，使当前文档中的中文字符自动与不可见的网格线对齐，这些网格线穿过其他中文字符的水平和垂直边缘。*/ ActiveDocument.SnapToShapes = true` |

#### **Document.SpellingChecked**

如果已对指定的区域或文档完成拼写检查，则该属性值为 **True**。如果所有或部分区域或文档尚未进行拼写检查，则该属性值为 **False**。**Boolean** 类型，可读写。

**语法**

**express.SpellingChecked**

*express*   一个代表 **Document** 对象的变量。

**说明**

若要重新检查某一区域或文档的拼写，请将 **SpellingChecked** 属性设置为 **False**。

若要查看该区域或文档中是否含有拼写错误，请使用 **SpellingErrors** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将“MyDocument.doc”的 SpellingChecked 属性设置为 False，然后对该文档进行另一轮拼写检查。*/ function test() { Documents.Item("MyDocument.doc").SpellingChecked = false Documents.Item("MyDocument.doc").CheckSpelling(null,false) }` |

#### **Document.SpellingErrors**

返回一个 **ProofreadingErrors** 集合，该集合代表在指定文档或区域中标识为拼写错误的单词。只读。

**语法**

**express.SpellingErrors**

*express*   一个代表 **Document** 对象的变量。

**说明**

有关返回集合的单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例检查活动文档的拼写错误，并显示找到的错误数目。*/ function test() { let myErr = ActiveDocument.SpellingErrors.Count if(myErr == 0) {     MsgBox ("No spelling errors found.") } else{     MsgBox(myErr + " spelling errors found.") } }` |

#### **Document.StoryRanges**

返回一个 **StoryRanges** 集合，该集合代表指定文档中的所有的文字部分。只读。

**语法**

**express.StoryRanges**

*express*   一个代表 **Document** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例检查 StoryRanges 集合中的每一个元素，判定 wdPrimaryFooterStory 是否为 StoryRanges 集合的一部分。*/ function test() { let aStory = ActiveDocument.StoryRanges for(let i = 1; i <= aStory.Count; i++) {     if(aStory.Item(i).StoryType == wdEvenPagesFooterStory) {         MsgBox("Document includes an even page footer")     } } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将文字添至文档页眉部分，然后显示该文字。*/ function test() { ActiveDocument.Sections.Item(1).Headers.Item(wdHeaderFooterPrimary).Range.Text = "Header text" MsgBox(ActiveDocument.StoryRanges.Item(wdPrimaryHeaderStory).Text) }` |

#### **Document.Styles**

本示例返回一个指定文档的 **Styles** 集合。只读。

**语法**

**express.Styles**

*express*   一个代表 **Document** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档中所有以“Chapter”开始的段落样式设置为“标题 1”。*/ function test() { let para = ActiveDocument.Paragraphs for(let i = 1; i <= para.Count; i++) {     if(para.Item(i).Range.Words.Item(1).Text == "Chapter ") {        para.Item(i).Style = ActiveDocument.Styles.Item(wdStyleHeading1)     } } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例打开活动文档所用的模板，并修改“标题 1”样式，然后保存模板，更新活动文档中所有的样式。*/ function test() { let tempDoc = ActiveDocument.AttachedTemplate.OpenAsDocument() let font = tempDoc.Styles.Item(wdStyleHeading1).Font     font.Name = "Arial"     font.Size = 16 tempDoc.Close(wdSaveChanges) ActiveDocument.UpdateStyles() }` |

#### **Document.StyleSortMethod**

返回或设置一个 **WdStyleSort** 常量，该常量代表要在对**“样式”**任务窗格中的样式进行排序时使用的排序方法。可读写。

**语法**

**express.StyleSortMethod**

*express*   一个代表 **Document** 对象的变量。

#### **Document.Subdocuments**

返回一个 **Subdocuments** 集合，该集合代表指定文档中的所有子文档。只读。

**语法**

**express.Subdocuments**

*express*   一个代表 **Document** 对象的变量。

**说明**

有关返回集合的单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示嵌入在活动文档中的子文档的数量。*/ MsgBox(ActiveDocument.Subdocuments.Count)` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示活动文档中每篇子文档的路径和文件名。*/ function test() { let subdoc = ActiveDocument.Subdocuments for(let i = 1; i <= subdoc.Count; i++) {     if(subdoc.Item(i).HasFile == true) {         MsgBox(subdoc.Item(i).Path + Application.PathSeparator + subdoc.Item(i).Name)     }     else {         MsgBox("This subdocument has not been saved.")     }    } }` |

#### **Document.Sync**

返回一个 **Sync** 对象，该对象提供对构成“文档工作区”的文档的方法和属性的访问。

**语法**

**express.Sync**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果活动文档是“文档工作区”中的共享文档，则以下示例显示活动文档最后一个修改者的姓名。*/ function test() { let strLastUser let eStatus = ActiveDocument.Sync.Status  if(eStatus == msoSyncStatusLatest) {     strLastUser = ActiveDocument.Sync.WorkspaceLastChangedBy     MsgBox("You have the most up-to-date copy." + "This file was last modified by " + strLastUser) } }` |

#### **Document.Tables**

返回一个 **Table**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Table/Table%20.htm#jsObject_Table)集合，该集合代表指定文档中的所有表格。只读。

**语法**

**express.Tables**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例在活动文档第一个表格的第一列中插入数字和文本 */ function test() { 	let num = 90 	let acell = Application.ActiveDocument.Tables.Item(1).Columns.Item(1).Cells 	for(let i = 1; i <= acell.Count; i++) { 	    acell.Item(i).Range.Text = num + " Sales" 	    num = num + 1 	} }` |

#### **Document.TablesOfAuthorities**

返回一个 **TableOfAuthorities** 集合，该集合代表指定文档中的引文目录。只读。

**语法**

**express.TablesOfAuthorities**

*express*   一个代表 **Document** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在文件 Sales.doc 的起始处添加一个引文目录。在该引文目录中汇编了所有类别的引文。*/ function test() { let myRange = Documents.Item("Sales.doc").Range(0, 0) Documents.Item("Sales.doc").TablesOfAuthorities.Add(myRange, 0, null, true, null, null, null, null, null, true) }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将更新活动文档中的所有引文目录。*/ function test() { let myTOA = ActiveDocument.TablesOfAuthorities for(let i = 1; i <= myTOA.Count; i++) {     myTOA.Item(i).Update() } }` |

#### **Document.TablesOfAuthoritiesCategories**

返回一个 **TablesOfAuthoritiesCategories** 集合，该集合代表指定文档中可用的引文目录类别。只读。

**语法**

**express.TablesOfAuthoritiesCategories**

*express*   一个代表 **Document** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例更改活动文档的引文目录类别列表中的八个元素的名称。*/ ActiveDocument.TablesOfAuthoritiesCategories.Item(8).Name = "Other case"` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在最后一个引文目录类别的名称发生了变化的情况下，显示该类别的名称。*/ function test() { let last = ActiveDocument.TablesOfAuthoritiesCategories.Count if(ActiveDocument.TablesOfAuthoritiesCategories.Item(last).Name != "16") {     MsgBox(ActiveDocument.TablesOfAuthoritiesCategories.Item(last).Name) } }` |

#### **Document.TablesOfContents**

返回一个 **TablesOfContents** 集合，该集合代表指定文档中的目录。只读。

**语法**

**express.TablesOfContents**

*express*   一个代表 **Document** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在 Sales.doc 文档的开始处添加目录。目录从 TC 域收集项目文本。*/ function test() { let myRange = Documents.Item("Sales.doc").Range(0, 0) Documents.Item("Sales.doc").TablesOfContents.Add(myRange, false, null, null, true)     }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例更新活动文档的目录中的各项目页码。*/ function test() { let myTOC = ActiveDocument.TablesOfContents for(let i = 1; i <= myTOC.Count; i++) {     myTOC.Item(i).UpdatePageNumbers() } }` |

#### **Document.TablesOfFigures**

返回一个 **TablesOfFigures** 集合，该集合代表指定文档中的图表目录。只读。

**语法**

**express.TablesOfFigures**

*express*   一个代表 **Document** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在活动文档的插入点位置添加图表目录。*/ function test() { Selection.Collapse(wdCollapseStart) ActiveDocument.TablesOfFigures.Add(Selection.Range, wdCaptionFigure) }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例更新 Report.doc 文档中第一个图表目录的内容。*/ Documents.Item("Report.doc").TablesOfFigures.Item(1).Update()` |

#### **Document.TextEncoding**

返回或设置代码页或字符集，WPS 应用该代码页或字符集将文档保存为编码文本文件。**MsoEncoding** 类型，可读写。

**语法**

**express.TextEncoding**

*express*   一个代表 **Document** 对象的变量。

**说明**

**TextEncoding** 属性用不同于 HTML 编码的方式设置文本编码，用户可以用 **Encoding** 属性设置 HTML 编码。如果要对保存为文本文件的所有文档设置文本编码，可使用 **DefaultTextEncoding** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果把活动文档保存为文本文件，本示例将设置活动文档的文本编码为日语。*/ function EncodeText() {     ActiveDocument.TextEncoding = msoEncodingJapaneseShiftJIS }` |

#### **Document.TextLineEnding**

返回或设置一个 **WdLineEndingType** 常量，该常量代表 WPS 在保存为文本文件的文档中标记直线和段落分隔符的方式。可读写。

**语法**

**express.TextLineEnding**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*将活动文档保存为文本文件时，本示例将为该活动文档中的直线或段落分隔符输入回车。*/ function LineEndings() {     ActiveDocument.TextLineEnding = wdCROnly }` |

#### **Document.TrackFormatting**

返回或设置 **Boolean** 值，该值表示在打开更改跟踪功能时是否跟踪格式更改。可读/写。

**语法**

**express.TrackFormatting**

*express*   一个代表 **Document** 对象的变量。

#### **Document.TrackMoves**

返回或设置 **Boolean** 值，该值表示在打开跟踪更改功能时是否对移动文本进行标记。可读/写。

**语法**

**express.TrackMoves**

*express*   一个代表 **Document** 对象的变量。

**说明**

默认情况下，如果打开跟踪更改功能，则移动文本将标记为删除和插入。如果 **TrackMoves** 为 **True**，则移动文本将标记为移动，并以删除线格式来标记原始文本。此属性对应于**“跟踪更改选项”**对话框中的**“跟踪移动”**复选框。

#### **Document.TrackRevisions**

如果标记对指定文档的修改，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.TrackRevisions**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例对活动文档进行设置，以便能标记对该文档的修改，并在屏幕上显示修订标记。*/ function test() { ActiveDocument.TrackRevisions = true ActiveDocument.ShowRevisions = true }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将在没有启用修订标记的情况下插入文本。*/ function test() { if(ActiveDocument.TrackRevisions == false) {     Selection.InsertBefore("new text") } }` |

#### **Document.Type**

返回文档的类型（模板或文档）。只读 **WdDocumentType**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdDocumentType%20%E6%9E%9A%E4%B8%BE.html)类型。

**语法**

**express.Type**

*express*   一个代表 **Document** 对象的变量。

#### **Document.UpdateStylesOnOpen**

如果在每次打开指定文档时会依照其附属的模板更新文档样式，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.UpdateStylesOnOpen**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将启用所有打开文档的更新文档样式选项，然后关闭这些文档。当重新打开这些文档时，对该文档的附属模板中样式的修改将自动显示在文档中。*/ function test() { let doc = Documents for(let i = 1; i <= doc.Count; i++) {     doc.Item(i).UpdateStylesOnOpen = true     doc.Item(i).Close(wdSaveChanges) } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将禁用更新文档样式的选项，这样，对文档 Report.doc 的附属模板的样式进行的修改就不会反映在该文档中。*/ function test() { Documents.Item("Report.doc").UpdateStylesOnOpen = false }` |

#### **Document.UseMathDefaults**

返回或设置一个代表在创建新的公式时是否使用默认的数学设置的 **Boolean**。可读写。

**语法**

**express.UseMathDefaults**

*express*   一个代表 **Document** 对象的变量。

#### **Document.UserControl**

如果文档是由用户创建或打开的，则该属性为 **True**。可读/写 **Boolean** 类型。

**语法**

**express.UserControl**

*express*   一个代表 **Document** 对象的变量。

**说明**

如果文档是由其他 WPS Office应用程序用 **Open** 方法或 Visual Basic **CreateObject** 或 **GetObject** 命令以编程方式创建或打开的，则该属性返回 **False**。

| 注释                                                         |
| ------------------------------------------------------------ |
| 如果 WPS 对用户可见，或者您在 WPS 代码模块内部调用 **UserControl** 属性，则该属性将始终返回 **True**。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*该示例显示活动文档的 UserControl 属性的状态。本示例仅在从装载 WPS 对象库的其他 Office 应用程序运行的情况下有效。*/ function test() { Set wd = New Kwps.application Set wdDoc = _     wd.Documents.Open("C:\My Documents\doc1.doc") If wdDoc.UserControl = True Then     MsgBox "This document was created or opened by the user." Else     MsgBox "This document was created programmatically." End If }  ` |

#### **Document.Variables**

返回一个 **Variables** 集合，该集合代表指定文档所存的变量。只读。

**语法**

**express.Variables**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例为活动文档添加一个名为“Value1”的文档变量 */ Application.ActiveDocument.Variables.Add('num', 3)` |

#### **Document.VBASigned**

如果指定文档的 示例代码 (VBA) 项目具有数字签名，则该属性值为 **True**。**Boolean** 类型，只读。

**语法**

**express.VBASigned**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例加载一个名为“Temp.doc”的文档，并测试该文档是否具有数字签名。如果没有数字签名，本示例将显示警告信息。*/ function test() { Documents.Open("C:\\My Documents\\Temp.doc") if(ActiveDocument.VBASigned == false) {     MsgBox("Warning! This document has not been digitally signed.", vbCritical, "Digital Signature Warning") } }` |

#### **Document.VBProject**

返回指定模板或文档的 **VBProject** 对象。

**语法**

**express.VBProject**

*express*   一个代表 **Document** 对象的变量。

**说明**

使用该属性可访问代码模块和用户窗体。

要在对象浏览器中查看 **VBProject** 对象，必须在“Visual Basic 编辑器”的**“引用”**对话框中选定**“示例代码 Extensibility”**复选框（在**“工具”**菜单上）。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示活动文档的 Visual Basic 项目的名称。*/ function test() { Set currProj = ActiveDocument.VBProject MsgBox currProj.Name }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将一个标准代码模块添至活动文档，并将其重命名为“MyModule”。*/ function test() { Set newModule = ActiveDocument.VBProject.VBComponents _     .Add(vbext_ct_StdModule) NewModule.Name = "MyModule" }` |

#### **Document.WebOptions**

将文档保存为网页或打开网页时，该属性将返回 **WebOptions** 对象，该对象包含 WPS 所用的文档属性。只读。

**语法**

**express.WebOptions**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例指定将活动文档的项目保存为网页时，使用级联样式表和西文文档编码。*/ function test() { let objWO = ActiveDocument.WebOptions objWO.RelyOnCSS = true objWO.Encoding = msoEncodingWestern }` |

#### **Document.Windows**

返回一个 **Windows** 集合，该集合代表指定文档的所有窗口。只读。

**语法**

**express.Windows**

*express*   一个代表 **Document** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示 NewWindow 方法运行前后活动文档的窗口数目。*/ function test() { MsgBox(ActiveDocument.Windows.Count + " window(s)", ActiveDocument.Name) ActiveDocument.ActiveWindow.NewWindow() MsgBox(ActiveDocument.Windows.Count + " windows", ActiveDocument.Name) }` |

#### **Document.WordOpenXML**

返回一个 **String** 类型的值，该值代表文档的 WPS Open XML 内容的 Flat XML 格式。只读。

**语法**

**express.WordOpenXML**

*express*   一个代表 **Document** 对象的变量。

#### **Document.Words**

返回一个 **Words** 集合，该集合代表文档中的所有字词。只读。

**语法**

**express.Words**

*express*   一个代表 **Document** 对象的变量。

**说明**

文档中的标点符号和段落标记也包括在 **Words** 集合中。

有关返回集合的单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示所选内容的字数。段落标记、部分单词和标点符号都统计在内。*/ MsgBox("There are " + Selection.Words.Count + " words.")` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例遍历 myRange 中的单词（从活动文档的开始之所选内容的末尾），并删除该区域中出现的任何“Franklin”（包含尾部空格）。*/ function test() { let myRange = ActiveDocument.Range(0, Selection.End)		 let aWord = myRange.Words					 for(let i = 1; i <= aWord.Count; i++) {     if(aWord.Item(i).Text == "Franklin ") {         aWord.Item(i).Delete() } } }` |

#### **Document.WritePassword**

该属性设置一个保存对指定文档所做的修改时所需的密码。**String** 类型，只写。

**语法**

**express.WritePassword**

*express*   一个代表 **Document** 对象的变量。

**说明**

**要点**??请避免在应用程序中使用硬编码的密码。如果过程中要求输入密码，请从用户处请求密码，并将密码作为变量存储起来，然后在代码中使用该变量。有关如何做到这一点的建议最佳操作，请参阅 WPS Office解决方案开发人员安全注意事项。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果活动文档没有设置密码以防修改文档，则本示例将“secret”设置为该文档的写权限密码。*/ function test() { let myDoc = ActiveDocument if(myDoc.WriteReserved == false) {     myDoc.WritePassword = "secret" } }` |

#### **Document.WriteReserved**

如果指定文档设有写权限密码，则该属性值为 **True**。**Boolean** 类型，只读。

**语法**

**express.WriteReserved**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果活动文档设有写权限密码，则本示例显示一条消息。*/ function test() { if(ActiveDocument.WriteReserved == true) {     MsgBox("Changes cannot be made to this document.") } }` |

#### **Document.XMLHideNamespaces**

返回一个 **Boolean** 类型的值，该值表示是否隐藏 XML 命名空间，该命名空间位于“XML 结构”任务窗格的元素列表中。

**语法**

**express.XMLHideNamespaces**

*express*   一个代表 **Document** 对象的变量。

**说明**

如果为 **True**，则显示元素，并在元素名称右侧显示该元素 XML 架构命名空间。如果为 **False**，则不显示 XML 架构命名空间。

当包含相似或相同元素名称的多个架构附加到同一文档时，将 **XMLHideNamespaces** 属性值设置为 **False** 可能会很有帮助。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下列示例中， WPS 不显示活动文档中的 XML 架构命名空间。*/ ActiveDocument.XMLHideNamespaces = false` |

#### **Document.XMLNodes**

返回一个 **XMLNodes** 集合，该集合代表文档中的所有 XML 元素。

**语法**

**express.XMLNodes**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例返回活动文档中的第一个 XML 节点。*/ let objNode = ActiveDocument.XMLNodes.Item(1)` |

#### **Document.XMLSaveDataOnly**

设置或返回一个 **Boolean** 类型的值，保存文档时包含 XML 标记还是仅保存文本。

**语法**

**express.XMLSaveDataOnly**

*express*   一个代表 **Document** 对象的变量。

**说明**

如果为 **True**，则表明 WPS 在保存文档时仅包含自定义 XML 标记及其相关内容。如果为 **False**，则表明 WPS 在保存文档时将包含全部 XML 标记。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下列示例指定活动文档仅与自定义 XML 标记及相关内容一起保存。*/ ActiveDocument.XMLSaveDataOnly = true` |

#### **Document.XMLSaveThroughXSLT**

设置或返回一个 **String** 类型的值，指定用户保存文档时，Extensible Stylesheet Language Transformation (XSLT) 要应用的路径和文件名。

**语法**

**express.XMLSaveThroughXSLT**

*express*   一个代表 **Document** 对象的变量。

**说明**

仅当将 **XMLUseXSLTWhenSaving** 属性值设置为 **True** 时，**XMLSaveThroughXSLT** 属性才可用。如果将 **XMLUseXSLTWhenSaving** 属性值设置为 **False**，则 WPS 会忽略 **XMLSaveThroughXSLT** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下列示例指定 WPS 保存活动文档时应使用 XSLT，并指定了要使用的 XSLT。*/ function test() { ActiveDocument.XMLUseXSLTWhenSaving = true ActiveDocument.XMLSaveThroughXSLT = "C:\\schemas\\book.xsl" }` |

#### **Document.XMLSchemaReferences**

返回一个 XMLSchemaReferences 集合，代表附加到文档的架构。

**语法**

**express.XMLSchemaReferences**

*express*   一个代表 **Document** 对象的变量。

**说明**

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下列示例重新加载附加到活动文档的第一个架构。*/ function test() { let objSchema = ActiveDocument.XMLSchemaReferences.Item(1) objSchema.Reload() }` |

#### **Document.XMLSchemaViolations**

返回一个 XMLNodes 集合，代表文档中所有存在验证错误的节点。

**语法**

**express.XMLSchemaViolations**

*express*   一个代表 **Document** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下列示例创建一个对活动文档中存在验证错误的 XML 元素的引用。*/ let objNodes = ActiveDocument.XMLSchemaViolations` |

#### **Document.XMLShowAdvancedErrors**

返回或设置一个 **Boolean** 类型的值，该值表示是由内置的 WPS 错误消息还是由包括在 Office 中的 Microsoft XML Core Services (MSXML) 5.0 组件来生成错误消息文本。

**语法**

**express.XMLShowAdvancedErrors**

*express*   一个代表 **Document** 对象的变量。

**说明**

使用 MSXML 5.0 组件的高级错误信息可提供更具体的错误信息。当 **XMLShowAdvancedErrors** 属性为 **True** 时，通过 XML Core Services 可提供大约 500 条错误信息。

WPS 大约有 50 条内置的常规架构错误。当 **XMLShowAdvancedErrors** 属性为 **False** 时， WPS 对于 XML 文档中生成的错误使用内置的错误信息。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在活动文档中启用高级错误消息。*/ ActiveDocument.XMLShowAdvancedErrors = true` |

#### **Document.XMLUseXSLTWhenSaving**

返回一个 **Boolean** 类型的值，代表是否通过 Extensible Stylesheet Language Transformation (XSLT) 保存文档。如果通过 XSLT 保存文档，则该值为 **True**。

**语法**

**express.XMLUseXSLTWhenSaving**

*express*   一个代表 **Document** 对象的变量。

**说明**

在将 XMLUseXSLTWhenSaving 属性值设置为 **True** 时，使用 **XMLSaveThroughXSLT** 属性可指定要使用的 XSLT 的路径和文件名。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下列示例指定 WPS 保存活动文档时应使用 XSLT，并指定了要使用的 XSLT。*/ function test() { ActiveDocument.XMLUseXSLTWhenSaving = true ActiveDocument.XMLSaveThroughXSLT = "C:\\schemas\\book.xslt" }` |

适用环境：web

适用平台：windows/linux