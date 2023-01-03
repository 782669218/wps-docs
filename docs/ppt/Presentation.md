**Presentation**



代表一个WPP 演示文稿。

**说明**

**Presentation** 对象是 [Presentations ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Presentations/Presentations%20.htm#jsObject_Presentations)集合的成员。**Presentations** 集合包含代表WPP 中打开的演示文稿的所有 **Presentation** 对象。

以下示例说明如何执行下列操作：

- 返回通过名称或索引号指定的演示文稿
- 返回活动窗口中的演示文稿
- 返回指定的任一文档窗口或幻灯片放映窗口中的演示文稿

## 示例：

使用 **Presentations**(*index*) 返回单个 **Presentation** 对象，其中 *index* 是演示文稿的名称或索引号。演示文稿的名称就是其文件名，是否带文件扩展名均可，但不包含路径。以下示例在“Sample Presentation”的开头添加一张幻灯片。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.Presentations.Item("Sample Presentation").Slides.Add(1, 1)` |

请注意，如果同时打开了多个具有相同名称的演示文稿，则会返回集合中具有指定名称的第一个演示文稿。

使用 **ActivePresentation** 属性可返回活动窗口中的演示文稿。以下示例将保存活动演示文稿。

| 示例代码复制                            |
| --------------------------------------- |
| `Application.ActivePresentation.Save()` |

使用**Presentation** 属性可返回指定文档窗口或幻灯片放映窗口中的演示文稿。以下示例将显示运行于第一个幻灯片放映窗口中的幻灯片放映的名称。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `alert(Application.SlideShowWindows.Item(1).Presentation.Name)` |

**方法**

|                                                              | 名称                             | 说明                                                         |
| ------------------------------------------------------------ | -------------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddTitleMaster**               | 为指定的演示文稿添加一个标题母版。返回一个代表该标题母版的 **Master** 对象。如果演示文稿已有标题母版，则会出错。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddToFavorites**               | 在 Program Files 文件夹的 Favorites 文件夹中，添加一个快捷方式，以代表指定演示文稿的当前选择内容（对于 **Presentation** 对象）或指定超链接的目标文档（对于 **Hyperlink** 对象）。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ApplyTemplate**                | 将设计模板应用于指定的演示文稿。                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ApplyTheme**                   | 将主题或设计模板应用于指定的演示文稿。                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CheckIn**                      | 从本地计算机将演示文稿返回到服务器，并且设置本地文件为只读，以便使得其不能在本地编辑。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CheckIn12**                    | 从本地计算机将演示文稿返回到服务器，并且设置本地文件为只读，以便使得其不能在本地编辑。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Close**                        | 关闭指定的文档窗口、演示文稿或打开的任意多边形               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Convert**                      | 将图示转换为另一种类型。                                     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Export**                       | 使用指定的图形筛选器导出演示文稿中的每张幻灯片，并将导出的文件保存在指定的文件夹中。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ExportAsFixedFormat**          | 将 演示文稿的副本发布为固定格式的文件（PDF 或 XPS）。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **FollowHyperlink**              | 如果缓存文档已被下载，本方法将其显示出来。否则，该方法将识别超链接，下载目标文档并将其显示在适当的应用程序中。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **GetWorkflowTasks**             | 返回 WorkflowTasks 集合。 可读写。                           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **GetWorkflowTemplates**         | 返回 WorkflowTemplates 集合。 可读写。                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **LockServerFile**               | 锁定服务器上的演示文稿，以阻止对其进行的修改。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **NewWindow**                    | 打开一个包含指定演示文稿的新窗口。返回一个代表新窗口的 **DocumentWindow** 对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PrintOut**                     | 打印指定演示文稿。                                           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **PublishSlides**                | 通过加载的任意演示文稿创建包含幻灯片的 Web 演示文稿（HTML 格式）。可以在 Web 浏览器中查看发布的演示文稿。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **ReloadAs**                     | 基于指定的 HTML 文档编码重新加载演示文稿。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **RemoveDocumentInformation**    | 删除 演示文稿中的文档信息，如个人信息、批注和文档属性。      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Save**                         | 保存指定的演示文稿。                                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SaveAs**                       | 保存未保存过的演示文稿，或以其他文件名保存已保存过的演示文稿。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SaveCopyAs**                   | 在不修改原文件的情况下将指定演示文稿的副本保存至另一文件。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SendFaxOverInternet**          | 向指定的收件人以传真形式发送演示文稿。                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SetPasswordEncryptionOptions** | 设置 WPP 通过密码加密演示文稿时使用的选项。                  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **UpdateLinks**                  | 更新指定演示文稿中链接的 OLE 对象。                          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **WebPagePreview**               | 在当前 Web 浏览器中预览演示文稿。                            |

**属性**

|                                                              | 名称                                 | 说明                                                         |
| ------------------------------------------------------------ | ------------------------------------ | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**                      | 返回一个 **Application** 对象，该对象表示指定对象的创建者。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BuiltInDocumentProperties**        | 返回一个 **DocumentProperties** 集合，该集合代表指定演示文稿的所有内置文档属性。只读。有关返回集合中单个成员的信息，请参阅 返回集合中的对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ColorSchemes**                     | 返回一个 **ColorSchemes** 集合，该集合代表指定演示文稿中的配色方案。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CommandBars**                      | 返回一个 **CommandBars** 集合，该集合代表主机容器应用程序和 WPP 中合并的命令栏集。仅当容器是 DocObject 服务器（如 活页夹）并且 WPP 作为 OLE 服务器时，该属性返回有效对象。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Container**                        | 返回包含指定嵌入演示文稿的对象。只读 **Object** 类型。如果容器不支持 OLE 自动化，或指定的演示文稿不是嵌入在“活页夹”文件中，则该属性无效。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ContentTypeProperties**            | 返回 MetaProperties 集合，该集合描述演示文稿中存储的元数据。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CustomDocumentProperties**         | 返回一个 **DocumentProperties** 集合，该集合代表指定演示文稿的所有自定义文档属性。只读。有关返回集合中单个成员的信息，请参阅 返回集合中的对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CustomerData**                     | 返回一个 CustomerData 对象。只读。                           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CustomXMLParts**                   | 返回一个 **CustomXMLParts** 对象。                           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DefaultLanguageID**                | 返回或设置演示文稿的默认语言。在您设置一个演示文稿的 **DefaultLanguageID** 属性时，同时也为所有后续新演示文稿设置了该属性。可读/写 **MsoLanguageID** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DefaultShape**                     | 返回一个 **Shape** 对象，该对象代表演示文稿的默认形状。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Designs**                          | 返回一个 **Designs** 对象，该对象代表一个设计集合。          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DisplayComments**                  | 决定是否在指定的演示文稿中显示批注。可读/写 **MsoTriState** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DocumentInspectors**               | 返回 DocumentInspectors 集合。只读。                         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DocumentLibraryVersions**          | 返回一个 **DocumentLibraryVersions** 集合，该集合代表共享演示文稿的版本集合，该共享演示文稿启用了版本控制功能，并存储在服务器上的文档库中。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EnvelopeVisible**                  | 决定电子邮件的邮件头在文档窗口中是否可见。可读/写 **MsoTriState** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ExtraColors**                      | 返回一个 **ExtraColors** 对象，该对象代表指定演示文稿中可用的其他颜色。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FarEastLineBreakLanguage**         | 返回或设置启用了换行控制选项时用于确定所用换行级别的语言。可读/写 **MsoFarEastLineBreakLanguageID** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FarEastLineBreakLevel**            | 返回或设置基于亚洲字符级别的换行符。可读/写 **Long** 类型。可读/写 **PpFarEastLineBreakLevel** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Final**                            | 确定是否将演示文稿标记为最终版本（只读）。可读/写。          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Fonts**                            | 返回一个 **Fonts** 集合，该集合代表用于指定演示文稿的所有字体。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FullName**                         | 返回指定的加载宏或保存的演示文稿的名称，包括路径、当前文件系统分隔符和文件扩展名。只读。**String** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **GridDistance**                     | 设置或返回一个 **Single** 类型的值，该值代表网格线之间的距离。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HandoutMaster**                    | 返回一个代表讲义母版的 **Master** 对象。只读。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HasTitleMaster**                   | 如果指定的演示文稿具有标题母版，则为 **MsoTrue**。只读。**MsoTriState** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HasVBProject**                     | 返回一个值，该值指示当前演示文稿是否包含 示例代码 (VBA) 项目。只读 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LayoutDirection**                  | 为用户界面返回或设置版式方向。可为以下值之一，默认值取决于已选择或安装的语言支持。可读/写 **PpDirection** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**                             | 演示文稿的名称包括文件扩展名（对于已注册的文件类型），但不包括其路径。不能使用该属性来设置名称。如果需要更改名称，请使用 **SaveAs** 方法将演示文稿另存为其他名称。**String** 类型，只读 。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NoLineBreakAfter**                 | 返回或设置前置字符。可读/写 **String** 类型。                |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NoLineBreakBefore**                | 返回或设置后置标点字符。可读/写 **String** 类型。            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NotesMaster**                      | 返回一个代表备注母版的 **Master** 对象。只读。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PageSetup**                        | 返回一个 **PageSetup** 对象，该对象的属性控制指定演示文稿的幻灯片的页面设置属性。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**                           | 返回指定对象的父对象。                                       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Password**                         | 返回或设置一个 **String** 类型的值，该值代表要打开指定演示文稿必须提供的密码。可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PasswordEncryptionAlgorithm**      | 返回一个 **String** 类型的值，该值代表 WPP 通过密码加密文档所使用的算法。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PasswordEncryptionFileProperties** | 如果 WPP 对使用密码保护的文档的文件属性进行加密，则返回 **MsoTrue** 属性值。只读 **MsoTriState** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PasswordEncryptionKeyLength**      | 返回一个 **Long** 类型的值，该值表示 WPP 通过密码加密文档时所使用算法的密钥长度。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PasswordEncryptionProvider**       | 返回一个 **String** 类型的值，该值指定 WPP 通过密码加密文档时所使用的加密算法提供商的名称。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Path**                             | 返回 **String**，该值代表到指定的 **Presentation** 对象的路径，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Permission**                       | 返回一个 **Permission** 对象，该对象可用于限制对活动演示文稿的权限，还可用于返回或设置特定权限设置。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PrintOptions**                     | 返回一个 **PrintOptions** 对象，该对象代表与指定演示文稿共同保存的打印选项。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **PublishObjects**                   | 返回一个 **PublishObjects** 集合，该集合代表已完全或部分加载的演示文稿的集合，这些演示文稿可作为 HTML 进行发布。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ReadOnly**                         | 返回指定的演示文稿是否为只读。只读 **MsoTriState** 类型。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **RemovePersonalInformation**        | 如果该属性值为 **MsoTrue**，则 WPP 在保存演示文稿时从批注和修订中删除所有用户信息。**MsoTriState** 类型，可读/写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Research**                         | 返回一个 Research 对象，利用该对象可访问信息检索服务功能。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Saved**                            | 决定演示文稿自上次保存以来是否进行过更改。可读/写 **MsoTriState** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ServerPolicy**                     | 返回一个ServerPolicy 对象。只读。                            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SharedWorkspace**                  | 返回一个 **SharedWorkspace** 对象，该对象代表指定演示文稿所在的文档工作区。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Signatures**                       | 返回一个 **SignatureSet** 对象，该对象代表数字签名的集合。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SlideMaster**                      | 返回一个 [**Master**](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Master/Master%20.htm#jsObject_Master) 对象，该对象代表幻灯片母版。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Slides**                           | 返回一个 **Slides** 集合，该集合代表指定演示文稿中的所有幻灯片。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SlideShowSettings**                | 返回一个 **SlideShowSettings** 对象，该对象代表指定演示文稿的幻灯片放映设置。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SlideShowWindow**                  | 返回一个 **SlideShowWindow** 对象，该对象代表正在运行的指定演示文稿所在的幻灯片放映窗口。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **SnapToGrid**                       | 该属性值为 **MsoTrue** 时，在指定的演示文稿中将形状与网格线对齐。可读/写 **MsoTriState** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Sync**                             | 返回一个 **Sync** 对象，使用该对象可对存储在 Windows SharePoint Services 共享工作区中的共享演示文稿的本地副本和服务器副本的同步进行管理。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Tags**                             | 返回一个代表指定对象的标签的 **Tags** 对象。只读。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TemplateName**                     | 返回与指定演示文稿相关的设计模板的名称。只读 **String** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **TitleMaster**                      | 返回 [Master](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Master/Master%20.htm#jsObject_Master)对象，该对象代表指定演示文稿的标题母版。如果演示文稿没有标题母版，则会产生一个错误。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **WebOptions**                       | 返回一个 **WebOptions** 对象，该对象中包含演示文稿级别的属性，当以网页的方式发布或保存整个或部分演示文稿，或打开某个网页时，WPP 将使用这些属性。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Windows**                          | 返回一个 **DocumentWindows** 集合，该集合代表与指定演示文稿相关联的所有文档窗口。该属性不返回与该演示文稿相关联的任何幻灯片放映窗口。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **WritePassword**                    | 设置或返回一个 **String** 类型的值，该值代表保存对指定文档所做的更改时使用的密码。可读/写。 |

**成员方法**

#### **Presentation.AddTitleMaster**

为指定的演示文稿添加一个标题母版。返回一个代表该标题母版的 **Master** 对象。如果演示文稿已有标题母版，则会出错。

**语法**

**express.AddTitleMaster()**

*express*   一个代表 **Presentation** 对象的变量。

**返回值**

Master

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例中，如果当前演示文稿没有标题母版，则为其添加一个标题母版。 function test() {     if (Application.ActivePresentation.HasTitleMaster == false) {         Application.ActivePresentation.AddTitleMaster()     } }` |

#### **Presentation.AddToFavorites**

在 Program Files 文件夹的 Favorites 文件夹中，添加一个快捷方式，以代表指定演示文稿的当前选择内容（对于 **Presentation** 对象）或指定超链接的目标文档（对于 **Hyperlink** 对象）。

**语法**

**express.AddToFavorites()**

*express*   一个代表 **Presentation** 对象的变量。

**说明**

如果文档具有显示名称，则快捷方式的名称是该显示名称；否则，快捷方式名称由 HLINK.DLL 决定。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例在 Windows 程序文件夹的 Favorites 中添加一个指向当前演示文稿的超链接。 Application.ActivePresentation.AddToFavorites()` |

#### **Presentation.ApplyTemplate**

将设计模板应用于指定的演示文稿。

**语法**

**express.ApplyTemplate(FileName)**

*express*   一个代表 **Presentation** 对象的变量。

**参数**

| **名称**   | **必选/可选** | **数据类型** | **说明**             |
| ---------- | ------------- | ------------ | -------------------- |
| *FileName* | 必选          | **String**   | 指定设计模板的名称。 |

**说明**

如果在字符串中引用未安装的演示文稿设计模板，则将产生运行时错误。无论 **Application.****FeatureInstall** 属性设置是什么，该模板都不自动安装。要对当前未安装的模板使用 **ApplyTemplate** 方法，必须先安装附加设计模板。为此，请通过运行 WPS Office安装程序（单击 Windows 控制面板中的**“添加/删除程序”**图标）安装WPP 的附加设计模板。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将“Professional”设计模板应用于当前演示文稿。*/ Application.ActivePresentation.ApplyTemplate("c:\\program files\\office\\templates" + "\\presentation designs\\professional.pot")` |

#### **Presentation.ApplyTheme**

将主题或设计模板应用于指定的演示文稿。

**语法**

**express.ApplyTheme(themeName)**

*express*   一个代表 **Presentation** 对象的变量。

**参数**

| **名称**    | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------- | ------------- | ------------ | ------------------------------------------------------------ |
| *themeName* | 必选          | **String**   | 应用于 Presentation 对象的主题文件 (.thmx) 或设计模板文件 (.pot) 的路径和名称。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将保存的主题应用于当前演示文稿：*/ Application.ActivePresentation.ApplyTheme("C:\\Program Files\\Office\\Templates\\MyTheme.thmx")` |

#### **Presentation.CheckIn**

从本地计算机将演示文稿返回到服务器，并且设置本地文件为只读，以便使得其不能在本地编辑。

**语法**

**express.CheckIn(SaveChanges, Comments, MakePublic)**

*express*   一个代表 **Presentation** 对象的变量。

**参数**

| **名称**      | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *SaveChanges* | 可选          | **Boolean**  | 如果为 True，则将演示文稿保存到服务器位置。默认值为 False。  |
| *Comments*    | 可选          | **Boolean**  | 正被签入的演示文稿的修订批注（仅在 SaveChanges 等于 True 时应用）。 |
| *MakePublic*  | 可选          | **Variant**  | 如果为 True，则允许用户在签入之后发布演示文稿。这样会提交文档使之进入审批流程，最终可将一种演示文稿版本发布给对演示文稿具有只读权限的用户（仅在 SaveChanges 等于 True 时应用）。 |

**说明**

若要利用 WPP 中的协作功能，演示文稿必须保存在 SharePoint Portal Server 上。

**示例**

本示例检查服务器是否可以签入指定的演示文稿；如果可以，它将关闭此演示文稿并且将其签回到服务器中。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `function CheckInPresentation(strPresentation){     if(Application.Presentations.Item(strPresentation).CanCheckIn == true ){         Application.Presentations.Item(strPresentation).CheckIn()         alert(strPresentation + " has been checked in.")     }     else{         alert(strPresentation + " cannot be checked in " + "at this time.  Please try again later.")     } }` |

若要调用上述子程序，请使用下列子程序并用在上面备注部分中提到的服务器上的实际文件替换“http://servername/workspace/report.ppt”文件名。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `function CheckInPresentation(){     CheckInPresentation("http://servername/workspace/report.ppt") }` |

#### **Presentation.CheckIn12**

从本地计算机将演示文稿返回到服务器，并且设置本地文件为只读，以便使得其不能在本地编辑。

**语法**

**express.CheckIn12(SaveChanges, Comments, MakePublic, VersionType)**

*express*   一个代表 **Presentation** 对象的变量。

**参数**

| **名称**      | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *SaveChanges* | 可选          | **Boolean**  | 如果为 True，则将演示文稿保存到服务器位置。默认值为 False。  |
| *Comments*    | 可选          | **Variant**  | 正被签入的演示文稿的修订批注（仅在 SaveChanges 等于 True 时应用）。 |
| *MakePublic*  | 可选          | **Variant**  | 如果为 True，则允许用户在签入之后发布演示文稿。这样会提交文档使之进入审批流程，最终可将一种演示文稿版本发布给对演示文稿具有只读权限的用户（仅在 SaveChanges 等于 True 时应用）。 |
| *VersionType* | 可选          | **Variant**  | 演示文稿的版本号。                                           |

**说明**

要利用 WPP 中内置的协作功能，演示文稿必须存储在 SharePoint Portal Server 上。

**示例**

本示例检查服务器是否可以签入指定的演示文稿；如果可以，它将关闭此演示文稿并且将其签回到服务器中。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `function CheckInPresentation(strPresentation){     if(Application.Presentations.Item(strPresentation).CanCheckIn == true ){         Application.Presentations.Item(strPresentation).CheckIn()         alert(strPresentation + " has been checked in.")     }     else{         alert(strPresentation + " cannot be checked in " + "at this time.  Please try again later.")     } }` |

若要调用上述子程序，请使用下列子程序并用在上面备注部分中提到的服务器上的实际文件替换“http://servername/workspace/report.ppt”文件名。

| 示例代码                                                     |
| ------------------------------------------------------------ |
| `function CheckInPresentation(){     CheckInPresentation("http://servername/workspace/report.ppt") }` |

#### **Presentation.Close**

关闭指定的文档窗口、演示文稿或打开的任意多边形

**语法**

**express.Close()**

*express*   一个代表 **Presentation** 对象的变量。

**说明**

使用此方法，WPP关闭打开的演示文档，并且不提示用户保存所做的工作。为防止工作丢失，您应在使用 **Close** 方法之前使用 **Save** 或 **SaveAs** 方法。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在不保存更改的情况下关闭 Pres1.ppt。*/ function test() {     let myprtion = Application.Presentations.Item("pres1.ppt")     myprtion.Saved = true     myprtion.Close() }  /*本示例关闭所有打开的演示文稿。*/ function test() {     for(let i = Application.Presentations.Count; i >= 1; i--)     {         Application.Presentations.Item(i).Close()     } }` |

#### **Presentation.Convert**

将图示转换为另一种类型。

**语法**

**express.Convert(Type)**

*express*   一个代表 **Presentation** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型**       | **说明**             |
| -------- | ------------- | ------------------ | -------------------- |
| *Type*   | 必选          | **MsoDiagramType** | 要转换为的图示类型。 |

**说明**

如果目标图示的 **Type** 属性是组织结构图 (**msoDiagramTypeOrgChart**)，此方法将产生错误。

| **MsoDiagramType** 可以是下列 **MsoDiagramType** 常数之一。 |
| ----------------------------------------------------------- |
| **msoDiagramCycle**                                         |
| **msoDiagramMixed**                                         |
| **msoDiagramOrgChart**                                      |
| **msoDiagramPyramid**                                       |
| **msoDiagramRadial**                                        |
| **msoDiagramTarget**                                        |
| **msoDiagramVenn**                                          |

#### **Presentation.Export**

使用指定的图形筛选器导出演示文稿中的每张幻灯片，并将导出的文件保存在指定的文件夹中。

**语法**

**express.Export(Path, FilterName, ScaleWidth, ScaleHeight)**

*express*   一个代表 **Presentation** 对象的变量。

**参数**

| **名称**      | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Path*        | 必选          | **String**   | 要用来保存导出的幻灯片的文件夹的路径。可以包括完整路径；如果不包括完整路径，WPP 就会在当前文件夹中为导出的幻灯片创建一个子文件夹。 |
| *FilterName*  | 必选          | **String**   | 要导出幻灯片的图形格式。指定的图形格式必须在 Windows 注册表中注册导出筛选器。可以指定注册的扩展名或注册的筛选器名称。WPP将首先在注册表中搜索匹配的扩展名。如果没有找到与指定字符串匹配的扩展名，WPP将查找匹配的筛选器名称。 |
| *ScaleWidth*  | 可选          | **Number**   | 导出幻灯片的宽度（以像素为单位）。                           |
| *ScaleHeight* | 可选          | **Number**   | 导出幻灯片的高度（以像素为单位）。                           |

**说明**

导出一个演示文稿不会将演示文稿的 **Saved** 属性设置为 **True**。

WPP使用指定的图形筛选器保存演示文稿中每张单独的幻灯片。导出并保存到磁盘的幻灯片的名称由WPP 决定。通常情况下，这些文件保存为 Slide1.wmf、Slide2.wmf 等。保存文件的路径在 **Path** 参数中指定。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动演示文稿保存为 WPP 演示文稿，然后将其中的每个幻灯片导出为移植网络图形 (PNG) 文件并保存在 Current Work 文件夹中*/ /*本示例还将每个导出的幻灯片的高度和宽度均设为 100 像素。*/ function test() {     let mypretion = Application.ActivePresentation     mypretion.SaveAs("c:\\Current Work\\Annual Sales", Application.Enum.ppSaveAsPresentation)     mypretion.Export("c:\\Current Work", "png", 100, 100) } ` |

#### **Presentation.ExportAsFixedFormat**

将 演示文稿的副本发布为固定格式的文件（PDF 或 XPS）。

**语法**

**express.ExportAsFixedFormat(Path, FixedFormatType, Intent, FrameSlides, HandoutOrder, OutputType, PrintHiddenSlides, PrintRange, RangeType, SlideShowName, IncludeDocProperties, KeepIRMSettings)**

*express*   一个代表 **Presentation** 对象的变量。

**参数**

| **名称**               | **必选/可选** | **数据类型**            | **说明**                                   |
| ---------------------- | ------------- | ----------------------- | ------------------------------------------ |
| *Path*                 | 必选          | **String**              | 导出的路径。                               |
| *FixedFormatType*      | 必选          | **PpFixedFormatType**   | 幻灯片应导出为的格式。                     |
| *Intent*               | 可选          | **PpFixedFormatIntent** | 导出的目的。                               |
| *FrameSlides*          | 可选          | **MsoTriState**         | 要导出的幻灯片。                           |
| *HandoutOrder*         | 可选          | **PpPrintHandoutOrder** | 打印讲义的顺序。                           |
| *OutputType*           | 可选          | **PpPrintOutputType**   | 讲义的类型。                               |
| *PrintHiddenSlides*    | 可选          | **MsoTriState**         | 打印隐藏的幻灯片。                         |
| *PrintRange*           | 可选          | **PrintRange**          | 幻灯片范围。                               |
| *RangeType*            | 可选          | **PpPrintRangeType**    | 幻灯片范围的类型。                         |
| *SlideShowName*        | 可选          | **String**              | 幻灯片的名称。                             |
| *IncludeDocProperties* | 可选          | **Boolean**             | 如果还应导出文档属性，则该参数值为 True。  |
| *KeepIRMSettings*      | 可选          | **Boolean**             | 如果还应导出 IRM 设置，则该参数值为 True。 |

#### **Presentation.FollowHyperlink**

如果缓存文档已被下载，本方法将其显示出来。否则，该方法将识别超链接，下载目标文档并将其显示在适当的应用程序中。

**语法**

**express.FollowHyperlink(Address, SubAddress, NewWindow, AddHistory, ExtraInfo, Method, HeaderInfo)**

*express*   一个代表 **Presentation** 对象的变量。

**参数**

| **名称**     | **必选/可选** | **数据类型**           | **说明**                                                     |
| ------------ | ------------- | ---------------------- | ------------------------------------------------------------ |
| *Address*    | 必选          | **String**             | 目标文档的地址。                                             |
| *SubAddress* | 可选          | **String**             | 在目标文档中的位置。此参数默认为一个空字符串。               |
| *NewWindow*  | 可选          | **Boolean**            | 值为 True 时，在新窗口中打开目标应用程序。默认值为 False。   |
| *AddHistory* | 可选          | **Boolean**            | 值为 True 时，将链接添加到当天的历史记录文件夹中。           |
| *ExtraInfo*  | 可选          | **String**             | 指定 HTTP 信息的字符串或字节数组。例如，该参数可用于指定图像映射的坐标或窗体的内容等，还可用于指定 FAT 文件名。Method 参数将决定如何处理此附加信息。 |
| *Method*     | 可选          | **MsoExtraInfoMethod** | 指定发布或附加 ExtraInfo 的方式。                            |
| *HeaderInfo* | 可选          | **String**             | 指定 HTTP 请求的标头信息的字符串。默认值为一个空字符串。可以通过使用下列语法将几个标头行组合为单个字符串："string1" & vbCr & "string2"。所指定的字符串自动转换为 ANSI 字符。请注意，HeaderInfo 参数可能会覆盖默认的 HTTP 标头字段。 |

**说明**

*MsoExtraInfoMethod* 可为以下 **MsoExtraInfoMethod** 常量之一。

| 常量              | 说明                                                  |
| ----------------- | ----------------------------------------------------- |
| **msoMethodGet**  | 默认值。*ExtraInfo* 是追加到该地址的一个 **String**。 |
| **msoMethodPost** | *ExtraInfo* 以 **String** 或字节数组形式发布。        |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例将 example.microsoft.com 中的文档加载到一个新窗口，并将其添加到“历史记录”文件夹。 Application.ActivePresentation.FollowHyperlink("http://example.microsoft.com",undefined,true,true)` |

#### **Presentation.GetWorkflowTasks**

返回 WorkflowTasks 集合。 可读写。

**语法**

**express.GetWorkflowTasks()**

*express*   一个代表 **Presentation** 对象的变量。

**返回值**

WorkFlowTasks

#### **Presentation.GetWorkflowTemplates**

返回 WorkflowTemplates 集合。 可读写。

**语法**

**express.GetWorkflowTemplates()**

*express*   一个代表 **Presentation** 对象的变量。

**返回值**

WordFlowTemplates

#### **Presentation.LockServerFile**

锁定服务器上的演示文稿，以阻止对其进行的修改。

**语法**

**express.LockServerFile()**

*express*   一个代表 **Presentation** 对象的变量。

#### **Presentation.NewWindow**

打开一个包含指定演示文稿的新窗口。返回一个代表新窗口的 **DocumentWindow** 对象。

**语法**

**express.NewWindow()**

*express*   一个代表 **Presentation** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例创建一个与当前窗口有相同内容的新窗口（同时激活新窗口）并返回到原窗口。*/ function test() {     let oldW = Application.ActiveWindow     let newW = oldW.NewWindow()     oldW.Activate() }` |

#### **Presentation.PrintOut**

打印指定演示文稿。

**语法**

**express.PrintOut(From, To, PrintToFile, Copies, Collate)**

*express*   一个代表 **Presentation** 对象的变量。

**参数**

| **名称**      | **必选/可选** | **数据类型**    | **说明**                                                     |
| ------------- | ------------- | --------------- | ------------------------------------------------------------ |
| *From*        | 可选          | **Number**      | 要打印的第一页的页码。如果忽略此参数，将从演示文稿的开头开始打印。指定 To 和 From 参数会设置 PrintRanges 对象的内容，并设置演示文稿的 RangeType 属性的值。 |
| *To*          | 可选          | **Number**      | 要打印的最后一页的页码。如果忽略此参数，将一直打印到演示文稿的末尾。指定 To 和 From 参数会设置 PrintRanges 对象的内容，并设置演示文稿的 RangeType 属性的值。 |
| *PrintToFile* | 可选          | **String**      | 要输出到的文件名。如果指定此参数，则文件将输出到某个文件，而不是发送到打印机。如果忽略此参数，文件将发送至打印机。 |
| *Copies*      | 可选          | **Number**      | 要打印的份数。如果忽略此参数，将只打印一份。指定此参数会设置 NumberOfCopies 属性的值。 |
| *Collate*     | 可选          | **MsoTriState** | 如果忽略此参数，将逐份打印多份。指定此参数会设置 Collate 属性的值。 |

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。                |
| ------------------------------------------------------------ |
| **msoCTrue**                                                 |
| **msoFalse**                                                 |
| **msoTriStateMixed**                                         |
| **msoTriStateToggle**                                        |
| **msoTrue**：先打印该演示文稿的完整副本，然后再打印下一副本的第一页。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例打印每个幻灯片（从当前演示文稿的第二个幻灯片到第五个幻灯片）的两个不分页复制的副本（不管是可见的还是隐藏的）。*/ function test() {     let mypertion = Application.ActivePresentation     mypertion.PrintOptions.PrintHiddenSlides = true     mypertion.PrintOut(2,5,"test",2,msoFalse) }  /*本示例将当前演示文稿中的所有幻灯片的单个副本输出到文件 Testprnt.prn。*/ Application.ActivePresentation.PrintOut(undefined,undefined,"TestPrnt")` |

#### **Presentation.PublishSlides**

通过加载的任意演示文稿创建包含幻灯片的 Web 演示文稿（HTML 格式）。可以在 Web 浏览器中查看发布的演示文稿。

**语法**

**express.PublishSlides(SlideLibraryUrl, Overwrite)**

*express*   一个代表 **Presentation** 对象的变量。

**参数**

| **名称**          | **必选/可选** | **数据类型** | **说明**                                  |
| ----------------- | ------------- | ------------ | ----------------------------------------- |
| *SlideLibraryUrl* | 必选          | **String**   | 幻灯片库的 URL。                          |
| *Overwrite*       | 可选          | **Boolean**  | 如果覆盖原始演示文稿，则该参数值为 True。 |

#### **Presentation.ReloadAs**

基于指定的 HTML 文档编码重新加载演示文稿。

**语法**

**express.ReloadAs(cp)**

*express*   一个代表 **Presentation** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型**    | **说明**                       |
| -------- | ------------- | --------------- | ------------------------------ |
| *cp*     | 必选          | **MsoEncoding** | 重新加载网页时使用的文档编码。 |

**说明**

| MsoEncoding 可以是下列 MsoEncoding 常量之一。 |
| --------------------------------------------- |
| **msoEncodingArabicAutoDetect**               |
| **msoEncodingAutoDetect**                     |
| **msoEncodingCyrillicAutoDetect**             |
| **msoEncodingGreekAutoDetect**                |
| **msoEncodingJapaneseAutoDetect**             |
| **msoEncodingKoreanAutoDetect**               |
| **msoEncodingSimplifiedChineseAutoDetect**    |
| **msoEncodingTraditionalChineseAutoDetect**   |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例使用西方编码重新加载当前演示文稿。 Application.ActivePresentation.ReloadAs(Application.Enum.msoEncodingWestern)` |

#### **Presentation.RemoveDocumentInformation**

删除 演示文稿中的文档信息，如个人信息、批注和文档属性。

**语法**

**express.RemoveDocumentInformation(Type)**

*express*   一个代表 **Presentation** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型**            | **说明**             |
| -------- | ------------- | ----------------------- | -------------------- |
| *Type*   | 必选          | **PpRemoveDocInfoType** | 要删除的信息的类型。 |

#### **Presentation.Save**

保存指定的演示文稿。

**语法**

**express.Save()**

*express*   一个代表 **Presentation** 对象的变量。

**说明**

使用 **SaveAs** 方法可保存尚未保存的演示文稿。若要判断演示文稿是否已保存，请测试 **FullName** 或 **Path** 属性是否非空。如果磁盘上已存在与指定演示文稿同名的文件，将覆盖该文件且不显示警告信息。

若要将演示文稿标记为已保存，但不将其写入磁盘，请将 **Saved** 属性设为 **True**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果当前演示文稿在上次保存后进行了更改，则本示例保存当前演示文稿。*/ function test() {     let tion = Application.ActivePresentation     if(tion.Saved == false && tion.Path != "")      {         tion.Save()     } } ` |

#### **Presentation.SaveAs**

保存未保存过的演示文稿，或以其他文件名保存已保存过的演示文稿。

**语法**

**express.SaveAs(Filename, FileFormat, EmbedFonts)**

*express*   一个代表 **Presentation** 对象的变量。

**参数**

| **名称**     | **必选/可选** | **数据类型**         | **说明**                                                     |
| ------------ | ------------- | -------------------- | ------------------------------------------------------------ |
| *Filename*   | 必选          | **String**           | 指定文件的保存名称。如果不指定完整路径，WPP会将该文件保存于当前文件夹中。 |
| *FileFormat* | 可选          | **PpSaveAsFileType** | 指定文件的保存格式。如果省略该参数，则文件将以WPP 当前版本中的演示文稿格式保存 (ppSaveAsPresentation)。 |
| *EmbedFonts* | 可选          | **MsoTriState**      | 指定WPP 是否将 TrueType 字体嵌入保存的演示文稿中。           |

**说明**

| PpSaveAsFileType 可以是下列 PpSaveAsFileType 常量之一。 |
| ------------------------------------------------------- |
| ppSaveAsHTMLv3                                          |
| ppSaveAsAddIn                                           |
| ppSaveAsBMP                                             |
| ppSaveAsDefault                                         |
| ppSaveAsGIF                                             |
| ppSaveAsHTML                                            |
| ppSaveAsHTMLDual                                        |
| ppSaveAsJPG                                             |
| ppSaveAsMetaFile                                        |
| ppSaveAsPNG                                             |
| ppSaveAsPowerPoint3                                     |
| ppSaveAsPowerPoint4                                     |
| ppSaveAsPowerPoint4FarEast                              |
| ppSaveAsPowerPoint7                                     |
| ppSaveAsPresentation 默认值。                           |
| ppSaveAsRTF                                             |
| ppSaveAsShow                                            |
| ppSaveAsTemplate                                        |
| ppSaveAsTIF                                             |
| ppSaveAsWebArchive                                      |

| MsoTriState 可以是下列 MsoTriState 常量之一。           |
| ------------------------------------------------------- |
| **msoCTrue**                                            |
| **msoFalse**                                            |
| **msoTriStateMixed** 默认值。                           |
| **msoTriStateToggle**                                   |
| **msoTrue**WPP 将 TrueType 字体嵌入到保存的演示文稿中。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例以文件名“New Format Copy.ppt”保存当前演示文稿的一个副本。*/ /*默认情况下，该副本以当前版本的WPP 演示文稿格式保存，然后以文件名“Old Format Copy”另存为WPP 4.0 格式文件。*/ function test() {     let tion= Application.ActivePresentation     tion.SaveCopyAs("New Format Copy")     tion.SaveAs("Old Format Copy", ppSaveAsPowerPoint4) }` |

#### **Presentation.SaveCopyAs**

在不修改原文件的情况下将指定演示文稿的副本保存至另一文件。

**语法**

**express.SaveCopyAs(FileName, FileFormat, EmbedTrueTypeFonts)**

*express*   一个代表 **Presentation** 对象的变量。

**参数**

| **名称**             | **必选/可选** | **数据类型**         | **说明**                                                     |
| -------------------- | ------------- | -------------------- | ------------------------------------------------------------ |
| *FileName*           | 必选          | **String**           | 指定文件的保存名称。如果不指定完整路径，WPP会将该文件保存于当前文件夹中。 FileFormat |
| *FileFormat*         | 可选          | **PpSaveAsFileType** | 文件格式。                                                   |
| *EmbedTrueTypeFonts* | 可选          | **MsoTriState**      | 指定是否嵌入 TrueType 字体。                                 |

**说明**

| FileFormat 可以是下列常量之一。 |
| ------------------------------- |
| ppSaveAsHTMLv3                  |
| ppSaveAsAddIn                   |
| ppSaveAsBMP                     |
| ppSaveAsDefault 默认值          |
| ppSaveAsGIF                     |
| ppSaveAsHTML                    |
| ppSaveAsHTMLDual                |
| ppSaveAsJPG                     |
| ppSaveAsMetaFile                |
| ppSaveAsPNG                     |
| ppSaveAsPowerPoint3             |
| ppSaveAsPowerPoint4             |
| ppSaveAsPowerPoint4FarEast      |
| ppSaveAsPowerPoint7             |
| ppSaveAsPresentation            |
| ppSaveAsRTF                     |
| ppSaveAsShow                    |
| ppSaveAsTemplate                |
| ppSaveAsTIF                     |
| ppSaveAsWebArchive              |

MsoTriState 可以是下列 MsoTriState 常量之一。**msoCTrue****msoFalse****msoTriStateMixed** 默认值**msoTriStateToggle****msoTrue**

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例以文件名“New Format Copy.ppt”保存当前演示文稿的一个副本。*/ /*默认情况下，该副本以当前版本的WPP 演示文稿格式保存，然后以文件名“Old Format Copy”另存为WPP 4.0 格式文件。*/ function test() {     let tion= Application.ActivePresentation     tion.SaveCopyAs("New Format Copy")     tion.SaveAs("Old Format Copy", ppSaveAsPowerPoint4) }` |

#### **Presentation.SendFaxOverInternet**

向指定的收件人以传真形式发送演示文稿。

**语法**

**express.SendFaxOverInternet(Recipients, Subject, ShowMessage)**

*express*   一个代表 **Presentation** 对象的变量。

**参数**

| **名称**      | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Recipients*  | 可选          | **Variant**  | 一个代表传真收件人的传真号码和电子邮件地址的 String 类型值。如果有多个收件人，则用分号隔开。 |
| *Subject*     | 可选          | **Variant**  | 一个代表传真演示文稿的主题行的 String 类型值。               |
| *ShowMessage* | 可选          | **Variant**  | 如果值为 True，则发送传真之前显示传真邮件。如果值为 False，则发送传真而不显示传真邮件。 |

**说明**

要使用 **SendFaxOverInternet** 方法，需要在用户的计算机上启用传真服务。

用于指定 **Recipients** 参数中传真号码的格式可为 *recipientsfaxnumber*@*usersfaxprovider* 或 *recipientsname*@*recipientsfaxnumber*。可使用以下注册表路径访问用户的传真提供商信息：

`HKEY_CURRENT_USER\Software\Microsoft\Office\11.0\Common\Services\Fax`

使用上面的注册表路径下的 FaxAddress 键值确定用户可使用的格式。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例将传真发送给传真服务提供商，该提供商将邮件传真给收件人。 Application.ActivePresentation.SendFaxOverInternet("14255550101@consolidatedmessenger.com","For your review", true)` |

#### **Presentation.SetPasswordEncryptionOptions**

设置 WPP 通过密码加密演示文稿时使用的选项。

**语法**

**express.SetPasswordEncryptionOptions(PasswordEncryptionProvider, PasswordEncryptionAlgorithm, PasswordEncryptionKeyLength, PasswordEncryptionFileProperties)**

*express*   一个代表 **Presentation** 对象的变量。

**参数**

| **名称**                           | **必选/可选** | **数据类型**    | **说明**                                       |
| ---------------------------------- | ------------- | --------------- | ---------------------------------------------- |
| *PasswordEncryptionProvider*       | 必选          | **String**      | 加密提供程序的名称。                           |
| *PasswordEncryptionAlgorithm*      | 必选          | **String**      | 加密算法的名称。WPP支持流加密算法。            |
| *PasswordEncryptionKeyLength*      | 必选          | **Long**        | 加密密钥的长度。必须为 8 的倍数，且不小于 40。 |
| *PasswordEncryptionFileProperties* | 必选          | **MsoTriState** | 如果值为 MsoTrue，则WPP 会对文件属性进行加密。 |

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。 |
| --------------------------------------------- |
| **msoCTrue** 不适用于此方法。                 |
| **msoFalse**                                  |
| **msoTriStateMixed** 不适用于此方法。         |
| **msoTriStateToggle** 不适用于此方法。        |
| **msoTrue**                                   |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例中，如果使用密码保护的文档的文件属性没有加密，则对密码加密选项进行设置。 function PasswordSettings() {     let tion = Application.ActivePresentation     if (tion.PasswordEncryptionFileProperties == msoFalse) {         tion.SetPasswordEncryptionOptions(" RSA SChannel Cryptographic Provider", "RC4", 56, true)     } }` |

#### **Presentation.UpdateLinks**

更新指定演示文稿中链接的 OLE 对象。

**语法**

**express.UpdateLinks()**

*express*   一个代表 **Presentation** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例更新当前演示文稿中的所有 OLE 链接。*/ Application.ActivePresentation.UpdateLinks()` |

#### **Presentation.WebPagePreview**

在当前 Web 浏览器中预览演示文稿。

**语法**

**express.WebPagePreview()**

*express*   一个代表 **Presentation** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例将第二篇演示文稿作为网页进行预览。 Application.Presentations.Item(2).WebPagePreview()` |

**成员属性**

#### **Presentation.Application**

返回一个 **Application** 对象，该对象表示指定对象的创建者。

**语法**

**express.Application**

*express*   一个代表 **Presentation** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例中在活动的演示文稿中添加一张幻灯片，然后将该演示文稿保存在运行 WPP 的文件夹中。*/ function test()  {     let pptPres = Application.ActivePresentation     pptPres.Slides.Add(1, 1)     pptPres.SaveAs(pptPres.Application.Path + "\\Added Slide") }  /*以下示例显示在当前演示文稿的第一张幻灯片上创建每个链接的 OLE 对象的应用程序的名称。*/ function test() {     for(let i=1; i<= Application.ActivePresentation.Slides.Item(1).Shapes.Count; i++)      {         let shpOle = Application.ActivePresentation.Slides.Item(1).Shapes.Item(i)         if(shpOle.Type == msoLinkedOLEObject)          {             alert(shpOle.OLEFormat.Application.Name)         }     } }` |

#### **Presentation.BuiltInDocumentProperties**

返回一个 **DocumentProperties** 集合，该集合代表指定演示文稿的所有内置文档属性。只读。有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**语法**

**express.BuiltInDocumentProperties**

*express*   一个代表 **Presentation** 对象的变量。

**说明**

使用 **CustomDocumentProperties** 属性返回自定义文档属性集合。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例显示活动演示文稿的所有内置文档属性的名称。 function test() {     let bidpList     for (let i = 1; i <= Application.ActivePresentation.BuiltInDocumentProperties.Count; i++) {         let p = Application.ActivePresentation.BuiltInDocumentProperties.Item(i)         bidpList = bidpList + p.Name + "\r"     }     alert(bidpList) }  //如果当前演示文稿的作者是“Jake Jarmel”，则本示例将该演示文稿的内置属性设置为“Category”。 function test() {     let ties = Application.ActivePresentation.BuiltInDocumentProperties     if (ties.Item("author").Value == "Jake Jarmel") {         ties.Item("category").Value = "Creative Writing"     } }` |

#### **Presentation.ColorSchemes**

返回一个 **ColorSchemes** 集合，该集合代表指定演示文稿中的配色方案。只读。

**语法**

**express.ColorSchemes**

*express*   一个代表 **Presentation** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `////本示例为返回活动演示文稿配色方案。 Application.ActivePresentation.ColorSchemes` |

#### **Presentation.CommandBars**

返回一个 **CommandBars** 集合，该集合代表主机容器应用程序和 WPP 中合并的命令栏集。仅当容器是 DocObject 服务器（如 活页夹）并且 WPP 作为 OLE 服务器时，该属性返回有效对象。只读。

**语法**

**express.CommandBars**

*express*   一个代表 **Presentation** 对象的变量。

#### **Presentation.Container**

返回包含指定嵌入演示文稿的对象。只读 **Object** 类型。如果容器不支持 OLE 自动化，或指定的演示文稿不是嵌入在“活页夹”文件中，则该属性无效。

**语法**

**express.Container**

*express*   一个代表 **Presentation** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例将隐藏包含嵌入活动演示文稿的“活页夹”文件的第二部分。演示文稿的 Container 属性将返回一个 Section 对象，该 Parent 对象的 Section 属性将返回一个 Binder 对象。 Application.ActivePresentation.Container.Parent.Sections.Item(2).Visible = false` |

#### **Presentation.ContentTypeProperties**

返回 MetaProperties 集合，该集合描述演示文稿中存储的元数据。只读。

**语法**

**express.ContentTypeProperties**

*express*   一个代表 **Presentation** 对象的变量。

#### **Presentation.CustomDocumentProperties**

返回一个 **DocumentProperties** 集合，该集合代表指定演示文稿的所有自定义文档属性。只读。有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**语法**

**express.CustomDocumentProperties**

*express*   一个代表 **Presentation** 对象的变量。

**说明**

使用 **BuiltInDocumentProperties** 属性可以返回内置文档属性集合。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例为活动演示文稿添加一个名为“Complete”的静态自定义属性。 Application.ActivePresentation.CustomDocumentProperties.Add("Complete",false,msoPropertyTypeBoolean,false)//如果自定义属性“Complete”的值为 True，本示例将显示当前演示文稿。 function test() {     let tion = Application.ActivePresentation     if (tion.CustomDocumentProperties.Item("complete")) {         tion.PrintOut()     } } ` |

#### **Presentation.CustomerData**

返回一个 CustomerData 对象。只读。

**语法**

**express.CustomerData**

*express*   一个代表 **Presentation** 对象的变量。

#### **Presentation.CustomXMLParts**

返回一个 **CustomXMLParts** 对象。

**语法**

**express.CustomXMLParts**

*express*   一个代表 **Presentation** 对象的变量。

#### **Presentation.DefaultLanguageID**

返回或设置演示文稿的默认语言。在您设置一个演示文稿的 **DefaultLanguageID** 属性时，同时也为所有后续新演示文稿设置了该属性。可读/写 **MsoLanguageID** 类型。

**语法**

**express.DefaultLanguageID**

*express*   一个代表 **Presentation** 对象的变量。

**说明**

| MsoLanguageID 可以是下列 MsoLanguageID 常量之一。 |
| ------------------------------------------------- |
| **msoLanguageIDAfrikaans**                        |
| **msoLanguageIDAlbanian**                         |
| **msoLanguageIDAmharic**                          |
| **msoLanguageIDArabic**                           |
| **msoLanguageIDArabicAlgeria**                    |
| **msoLanguageIDArabicBahrain**                    |
| **msoLanguageIDArabicEgypt**                      |
| **msoLanguageIDArabicIraq**                       |
| **msoLanguageIDArabicJordan**                     |
| **msoLanguageIDArabicKuwait**                     |
| **msoLanguageIDArabicLebanon**                    |
| **msoLanguageIDArabicLibya**                      |
| **msoLanguageIDArabicMorocco**                    |
| **msoLanguageIDArabicOman**                       |
| **msoLanguageIDArabicQatar**                      |
| **msoLanguageIDArabicSyria**                      |
| **msoLanguageIDArabicTunisia**                    |
| **msoLanguageIDArabicUAE**                        |
| **msoLanguageIDArabicYemen**                      |
| **msoLanguageIDArmenian**                         |
| **msoLanguageIDAssamese**                         |
| **msoLanguageIDAzeriCyrillic**                    |
| **msoLanguageIDAzeriLatin**                       |
| **msoLanguageIDBasque**                           |
| **msoLanguageIDBelgianDutch**                     |
| **msoLanguageIDBelgianFrench**                    |
| **msoLanguageIDBengali**                          |
| **msoLanguageIDBrazilianPortuguese**              |
| **msoLanguageIDBulgarian**                        |
| **msoLanguageIDBurmese**                          |
| **msoLanguageIDByelorussian**                     |
| **msoLanguageIDCatalan**                          |
| **msoLanguageIDCherokee**                         |
| **msoLanguageIDChineseHongKong**                  |
| **msoLanguageIDChineseMacao**                     |
| **msoLanguageIDChineseSingapore**                 |
| **msoLanguageIDCroatian**                         |
| **msoLanguageIDCzech**                            |
| **msoLanguageIDDanish**                           |
| **msoLanguageIDDutch**                            |
| **msoLanguageIDEnglishAUS**                       |
| **msoLanguageIDEnglishBelize**                    |
| **msoLanguageIDEnglishCanadian**                  |
| **msoLanguageIDEnglishCaribbean**                 |
| **msoLanguageIDEnglishIreland**                   |
| **msoLanguageIDEnglishJamaica**                   |
| **msoLanguageIDEnglishNewZealand**                |
| **msoLanguageIDEnglishPhilippines**               |
| **msoLanguageIDEnglishSouthAfrica**               |
| **msoLanguageIDEnglishTrinidad**                  |
| **msoLanguageIDEnglishUK**                        |
| **msoLanguageIDEnglishUS**                        |
| **msoLanguageIDEnglishZimbabwe**                  |
| **msoLanguageIDEstonian**                         |
| **msoLanguageIDFaeroese**                         |
| **msoLanguageIDFarsi**                            |
| **msoLanguageIDFinnish**                          |
| **msoLanguageIDFrench**                           |
| **msoLanguageIDFrenchCameroon**                   |
| **msoLanguageIDFrenchCanadian**                   |
| **msoLanguageIDFrenchCotedIvoire**                |
| **msoLanguageIDFrenchLuxembourg**                 |
| **msoLanguageIDFrenchMali**                       |
| **msoLanguageIDFrenchMonaco**                     |
| **msoLanguageIDFrenchReunion**                    |
| **msoLanguageIDFrenchSenegal**                    |
| **msoLanguageIDFrenchWestIndies**                 |
| **msoLanguageIDFrenchZaire**                      |
| **msoLanguageIDFrisianNetherlands**               |
| **msoLanguageIDGaelicIreland**                    |
| **msoLanguageIDGaelicScotland**                   |
| **msoLanguageIDGalician**                         |
| **msoLanguageIDGeorgian**                         |
| **msoLanguageIDGerman**                           |
| **msoLanguageIDGermanAustria**                    |
| **msoLanguageIDGermanLiechtenstein**              |
| **msoLanguageIDGermanLuxembourg**                 |
| **msoLanguageIDGreek**                            |
| **msoLanguageIDGujarati**                         |
| **msoLanguageIDHebrew**                           |
| **msoLanguageIDHindi**                            |
| **msoLanguageIDHungarian**                        |
| **msoLanguageIDIcelandic**                        |
| **msoLanguageIDIndonesian**                       |
| **msoLanguageIDInuktitut**                        |
| **msoLanguageIDItalian**                          |
| **msoLanguageIDJapanese**                         |
| **msoLanguageIDKannada**                          |
| **msoLanguageIDKashmiri**                         |
| **msoLanguageIDKazakh**                           |
| **msoLanguageIDKhmer**                            |
| **msoLanguageIDKirghiz**                          |
| **msoLanguageIDKonkani**                          |
| **msoLanguageIDKorean**                           |
| **msoLanguageIDLao**                              |
| **msoLanguageIDLatvian**                          |
| **msoLanguageIDLithuanian**                       |
| **msoLanguageIDMacedonian**                       |
| **msoLanguageIDMalayalam**                        |
| **msoLanguageIDMalayBruneiDarussalam**            |
| **msoLanguageIDMalaysian**                        |
| **msoLanguageIDMaltese**                          |
| **msoLanguageIDManipuri**                         |
| **msoLanguageIDMarathi**                          |
| **msoLanguageIDMexicanSpanish**                   |
| **msoLanguageIDMixed**                            |
| **msoLanguageIDMongolian**                        |
| **msoLanguageIDNepali**                           |
| **msoLanguageIDNone**                             |
| **msoLanguageIDNoProofing**                       |
| **msoLanguageIDNorwegianBokmol**                  |
| **msoLanguageIDNorwegianNynorsk**                 |
| **msoLanguageIDOriya**                            |
| **msoLanguageIDPolish**                           |
| **msoLanguageIDPunjabi**                          |
| **msoLanguageIDRhaetoRomanic**                    |
| **msoLanguageIDRomanian**                         |
| **msoLanguageIDRomanianMoldova**                  |
| **msoLanguageIDRussian**                          |
| **msoLanguageIDRussianMoldova**                   |
| **msoLanguageIDSamiLappish**                      |
| **msoLanguageIDSanskrit**                         |
| **msoLanguageIDSerbianCyrillic**                  |
| **msoLanguageIDSerbianLatin**                     |
| **msoLanguageIDSesotho**                          |
| **msoLanguageIDSimplifiedChinese**                |
| **msoLanguageIDSindhi**                           |
| **msoLanguageIDSlovak**                           |
| **msoLanguageIDSlovenian**                        |
| **msoLanguageIDSorbian**                          |
| **msoLanguageIDSpanish**                          |
| **msoLanguageIDSpanishArgentina**                 |
| **msoLanguageIDSpanishBolivia**                   |
| **msoLanguageIDSpanishChile**                     |
| **msoLanguageIDSpanishColombia**                  |
| **msoLanguageIDSpanishCostaRica**                 |
| **msoLanguageIDSpanishDominicanRepublic**         |
| **msoLanguageIDSpanishEcuador**                   |
| **msoLanguageIDSpanishElSalvador**                |
| **msoLanguageIDSpanishGuatemala**                 |
| **msoLanguageIDSpanishHonduras**                  |
| **msoLanguageIDSpanishModernSort**                |
| **msoLanguageIDSpanishNicaragua**                 |
| **msoLanguageIDSpanishPanama**                    |
| **msoLanguageIDSpanishParaguay**                  |
| **msoLanguageIDSpanishPeru**                      |
| **msoLanguageIDSpanishPuertoRico**                |
| **msoLanguageIDSpanishUruguay**                   |
| **msoLanguageIDSpanishVenezuela**                 |
| **msoLanguageIDSutu**                             |
| **msoLanguageIDSwahili**                          |
| **msoLanguageIDSwedish**                          |
| **msoLanguageIDSwedishFinland**                   |
| **msoLanguageIDSwissFrench**                      |
| **msoLanguageIDSwissGerman**                      |
| **msoLanguageIDSwissItalian**                     |
| **msoLanguageIDTajik**                            |
| **msoLanguageIDTamil**                            |
| **msoLanguageIDTatar**                            |
| **msoLanguageIDTelugu**                           |
| **msoLanguageIDThai**                             |
| **msoLanguageIDTibetan**                          |
| **msoLanguageIDTraditionalChinese**               |
| **msoLanguageIDTsonga**                           |
| **msoLanguageIDTswana**                           |
| **msoLanguageIDTurkish**                          |
| **msoLanguageIDTurkmen**                          |
| **msoLanguageIDUkrainian**                        |
| **msoLanguageIDUrdu**                             |
| **msoLanguageIDUzbekCyrillic**                    |
| **msoLanguageIDUzbekLatin**                       |
| **msoLanguageIDVenda**                            |
| **msoLanguageIDVietnamese**                       |
| **msoLanguageIDWelsh**                            |
| **msoLanguageIDXhosa**                            |
| **msoLanguageIDZulu**                             |
| **msoLanguageIDPortuguese**                       |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例将活动演示文稿以及随后的所有新演示文稿的默认语言设置为 German。 Application.ActivePresentation.DefaultLanguageID = msoLanguageIDGerman` |

#### **Presentation.DefaultShape**

返回一个 **Shape** 对象，该对象代表演示文稿的默认形状。只读。

**语法**

**express.DefaultShape**

*express*   一个代表 **Presentation** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例为活动演示文稿的第一张幻灯片添加一个形状，将演示文稿中形状的默认填充颜色设为红色。然后添加另一个形状，默认填充颜色将自动应用于此形状。 function test() {     let tion = Application.ActivePresentation     let sld1Shapes = tion.Slides.Item(1).Shapes     sld1Shapes.AddShape(msoShape16pointStar, 20, 20, 100, 100)     tion.DefaultShape.Fill.ForeColor.RGB = (255, 0, 0)     sld1Shapes.AddShape(msoShape16pointStar, 150, 20, 100, 100) }` |

#### **Presentation.Designs**

返回一个 **Designs** 对象，该对象代表一个设计集合。

**语法**

**express.Designs**

*express*   一个代表 **Presentation** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例对于当前演示文稿中的每个设计都会显示一条消息。*/ function test() {     let pre = Application.ActivePresentation     for(let i=1; i<=pre.Designs.Count; i++)      {         alert("The design name is " + pre.Designs.Item(pre.Designs.Item(i).Index).Name)     } }` |

#### **Presentation.DisplayComments**

决定是否在指定的演示文稿中显示批注。可读/写 **MsoTriState** 类型。

**语法**

**express.DisplayComments**

*express*   一个代表 **Presentation** 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。 |
| --------------------------------------------- |
| **msoCTrue**                                  |
| **msoFalse**                                  |
| **msoTriStateMixed**                          |
| **msoTriStateToggle**                         |
| **msoTrue** 在指定的演示文稿中显示批注。      |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例隐藏活动演示文稿中的批注。 Application.ActivePresentation.DisplayComments = msoFalse` |

#### **Presentation.DocumentInspectors**

返回 DocumentInspectors 集合。只读。

**语法**

**express.DocumentInspectors**

*express*   一个代表 **Presentation** 对象的变量。

#### **Presentation.DocumentLibraryVersions**

返回一个 **DocumentLibraryVersions** 集合，该集合代表共享演示文稿的版本集合，该共享演示文稿启用了版本控制功能，并存储在服务器上的文档库中。

**语法**

**express.DocumentLibraryVersions**

*express*   一个代表 **Presentation** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//下面的示例返回活动演示文稿的版本集合。本示例假定活动演示文稿已启用了版本控制功能，并存储在服务器上的共享文档库中。 let objVersions = Application.ActivePresentation.DocumentLibraryVersions` |

#### **Presentation.EnvelopeVisible**

决定电子邮件的邮件头在文档窗口中是否可见。可读/写 **MsoTriState** 类型。

**语法**

**express.EnvelopeVisible**

*express*   一个代表 **Presentation** 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。  |
| ---------------------------------------------- |
| **msoCTrue**                                   |
| **msoFalse** 默认值。                          |
| **msoTriStateMixed**                           |
| **msoTriStateToggle**                          |
| **msoTrue** 电子邮件的邮件头在文档窗口中可见。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例显示电子邮件的邮件头。 Application.ActivePresentation.EnvelopeVisible = msoTrue` |

#### **Presentation.ExtraColors**

返回一个 **ExtraColors** 对象，该对象代表指定演示文稿中可用的其他颜色。只读。

**语法**

**express.ExtraColors**

*express*   一个代表 **Presentation** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例在当前演示文稿的第一张幻灯片中添加一个矩形，并将其前景填充颜色设为第一种其他颜色。演示文稿中必须已定义至少一种其他颜色，否则本示例将无效。 function test() {     let tion = Application.ActivePresentation     let rect = tion.Slides.Item(1).Shapes.AddShape(msoShapeRectangle, 50, 50, 100, 200)     rect.Fill.ForeColor.RGB = tion.ExtraColors.Item(1) }` |

#### **Presentation.FarEastLineBreakLanguage**

返回或设置启用了换行控制选项时用于确定所用换行级别的语言。可读/写 **MsoFarEastLineBreakLanguageID** 类型。

**语法**

**express.FarEastLineBreakLanguage**

*express*   一个代表 **Presentation** 对象的变量。

**说明**

| MsoFarEastLineBreakLanguageID 可以是下列 MsoFarEastLineBreakLanguageID 常量之一。 |
| ------------------------------------------------------------ |
| **MsoFarEastLineBreakLanguageJapanese**                      |
| **MsoFarEastLineBreakLanguageKorean**                        |
| **MsoFarEastLineBreakLanguageSimplifiedChinese**             |
| **MsoFarEastLineBreakLanguageTraditionalChinese**            |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例将换行语言设置为日语。 Application.ActivePresentation.FarEastLineBreakLanguage = MsoFarEastLineBreakLanguageJapanese` |

#### **Presentation.FarEastLineBreakLevel**

返回或设置基于亚洲字符级别的换行符。可读/写 **Long** 类型。可读/写 **PpFarEastLineBreakLevel** 类型。

**语法**

**express.FarEastLineBreakLevel**

*express*   一个代表 **Presentation** 对象的变量。

**说明**

| PpFarEastLineBreakLevel 可以是下列 PpFarEastLineBreakLevel 常量之一。 |
| ------------------------------------------------------------ |
| **ppFarEastLineBreakLevelCustom**                            |
| **ppFarEastLineBreakLevelNormal**                            |
| **ppFarEastLineBreakLevelStrict**                            |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例将换行控制设置为使用第一级首尾字符。 Application.ActivePresentation.FarEastLineBreakLevel = ppFarEastLineBreakLevelNormal` |

#### **Presentation.Final**

确定是否将演示文稿标记为最终版本（只读）。可读/写。

**语法**

**express.Final**

*express*   一个代表 **Presentation** 对象的变量。

**说明**

返回值 Boolean

#### **Presentation.Fonts**

返回一个 **Fonts** 集合，该集合代表用于指定演示文稿的所有字体。只读。

**语法**

**express.Fonts**

*express*   一个代表 **Presentation** 对象的变量。

#### **Presentation.FullName**

返回指定的加载宏或保存的演示文稿的名称，包括路径、当前文件系统分隔符和文件扩展名。只读。**String** 类型。

**语法**

**express.FullName**

*express*   一个代表 **Presentation** 对象的变量。

**说明**

此属性与 **Path** 属性等效，附加上当前文件系统分隔符和 **Name** 属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示当前演示文稿的路径和文件名（假设演示文稿已保存）。*/ alert(Application.ActivePresentation.FullName)` |

#### **Presentation.GridDistance**

设置或返回一个 **Single** 类型的值，该值代表网格线之间的距离。可读/写。

**语法**

**express.GridDistance**

*express*   一个代表 **Presentation** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例显示网格线，然后指定网格线之间的距离并启用对齐网格的设置。 function SetGridLines() {     Application.DisplayGridLines = msoTrue     let tion = Application.ActivePresentation     tion.GridDistance = 18     tion.SnapToGrid = msoTrue }` |

#### **Presentation.HandoutMaster**

返回一个代表讲义母版的 **Master** 对象。只读。

**语法**

**express.HandoutMaster**

*express*   一个代表 **Presentation** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例设置活动演示文稿的讲义母版的背景图案。 Application.ActivePresentation.HandoutMaster.Background.Fill.Patterned(Application.Enum.msoPatternDarkHorizontal)` |

#### **Presentation.HasTitleMaster**

如果指定的演示文稿具有标题母版，则为 **MsoTrue**。只读。**MsoTriState** 类型。

**语法**

**express.HasTitleMaster**

*express*   一个代表 **Presentation** 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。 |
| --------------------------------------------- |
| **msoCTrue**                                  |
| **msoFalse**                                  |
| **msoTriStateMixed**                          |
| **msoTriStateToggle**                         |
| **msoTrue**：指定的演示文稿具有标题母版。     |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例中，如果当前演示文稿没有标题母版，则为其添加一个标题母版。 function test() {     let tion = Application.ActivePresentation     if (!tion.HasTitleMaster) {         tion.AddTitleMaster()     } }` |

#### **Presentation.HasVBProject**

返回一个值，该值指示当前演示文稿是否包含 示例代码 (VBA) 项目。只读

**语法**

**express.HasVBProject**

*express*   一个代表 **Presentation** 对象的变量。

#### **Presentation.LayoutDirection**

为用户界面返回或设置版式方向。可为以下值之一，默认值取决于已选择或安装的语言支持。可读/写 **PpDirection** 类型。

**语法**

**express.LayoutDirection**

*express*   一个代表 **Presentation** 对象的变量。

**说明**

| PpDirection 可以是下列 PpDirection 常量之一。 |
| --------------------------------------------- |
| **ppDirectionLeftToRight**                    |
| **ppDirectionMixed**                          |
| **ppDirectionRightToLeft**                    |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将版式方向设置为从右向左。*/ Application.ActivePresentation.LayoutDirection = ppDirectionRightToLeft` |

#### **Presentation.Name**

演示文稿的名称包括文件扩展名（对于已注册的文件类型），但不包括其路径。不能使用该属性来设置名称。如果需要更改名称，请使用 **SaveAs** 方法将演示文稿另存为其他名称。**String** 类型，只读 。

**语法**

**express.Name**

*express*   一个代表 **Presentation** 对象的变量。

#### **Presentation.NoLineBreakAfter**

返回或设置前置字符。可读/写 **String** 类型。

**语法**

**express.NoLineBreakAfter**

*express*   一个代表 **Presentation** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例将“$”、“(”、“[”、“\”和“{”设置为前置字符。 function test() {     let tion = Application.ActivePresentation     tion.FarEastLineBreakLevel = ppFarEastLineBreakLevelCustom     tion.NoLineBreakAfter = "$([\{" }` |

#### **Presentation.NoLineBreakBefore**

返回或设置后置标点字符。可读/写 **String** 类型。

**语法**

**express.NoLineBreakBefore**

*express*   一个代表 **Presentation** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例将“!”、“)”和“]”设置为后置字符。 function test() {     let tion = Application.ActivePresentation     tion.FarEastLineBreakLevel = ppFarEastLineBreakLevelCustom     tion.NoLineBreakBefore = "!)]" }` |

#### **Presentation.NotesMaster**

返回一个代表备注母版的 **Master** 对象。只读。

**语法**

**express.NotesMaster**

*express*   一个代表 **Presentation** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置活动演示文稿的备注母版的页眉和页脚文。*/ function test() {     let ster = Application.ActivePresentation.NotesMaster.HeadersFooters     ster.Header.Text = "Employee Guidelines"     ster.Footer.Text = "Volcano Coffee" }` |

#### **Presentation.PageSetup**

返回一个 **PageSetup** 对象，该对象的属性控制指定演示文稿的幻灯片的页面设置属性。只读。

**语法**

**express.PageSetup**

*express*   一个代表 **Presentation** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例为名为“Pres1”的演示文稿设置幻灯片的大小和方向。*/ function test() {     let setup = Application.Presentations.Item("pres1").PageSetup     setup.SlideSize = ppSlideSize35MM     setup.SlideOrientation = msoOrientationHorizontal }` |

#### **Presentation.Parent**

返回指定对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **Presentation** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例在活动演示文稿的第一张幻灯片中添加一个包含文本的椭圆并将该椭圆及文本旋转 45 度。文本框的父对象就是包含文本的 Shape 对象。 function test() {     let myShapes = Application.ActivePresentation.Slides.Item(1).Shapes     let addShp = myShapes.AddShape(msoShapeOval, 50, 50, 300, 150).TextFrame     addShp.TextRange.Text = "Test text"     addShp.Parent.Rotation = 45 }` |

#### **Presentation.Password**

返回或设置一个 **String** 类型的值，该值代表要打开指定演示文稿必须提供的密码。可读/写。

**语法**

**express.Password**

*express*   一个代表 **Presentation** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例打开 Earnings.ppt，为其设置密码，然后关闭该演示文稿。*/ function test() {     let pen = Application.Presentations.Open("C:\\My Documents\\Earnings.ppt")     pen.Password = "complexstrPWD"     pen.Save()     pen.Close() }` |

#### **Presentation.PasswordEncryptionAlgorithm**

返回一个 **String** 类型的值，该值代表 WPP 通过密码加密文档所使用的算法。只读。

**语法**

**express.PasswordEncryptionAlgorithm**

*express*   一个代表 **Presentation** 对象的变量。

**说明**

使用 **SetPasswordEncryptionOptions** 方法可指定WPP 通过密码加密文档所使用的算法。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例中，如果使用的密码加密算法不是 RC4，则对密码加密选项进行设置。 function PasswordSettings() {     let tion = Application.ActivePresentation     if(tion.PasswordEncryptionAlgorithm!="RC4") {         tion.SetPasswordEncryptionOptions(" RSA SChannel Cryptographic Provider","RC4",56,true)     } }` |

#### **Presentation.PasswordEncryptionFileProperties**

如果 WPP 对使用密码保护的文档的文件属性进行加密，则返回 **MsoTrue** 属性值。只读 **MsoTriState** 类型。

**语法**

**express.PasswordEncryptionFileProperties**

*express*   一个代表 **Presentation** 对象的变量。

**说明**

使用 **SetPasswordEncryptionOptions** 方法可指定WPP 通过密码加密文档所使用的算法。

| MsoTriState 可以是下列 MsoTriState 常量之一。 |
| --------------------------------------------- |
| **msoCTrue**                                  |
| **msoFalse**                                  |
| **msoTriStateMixed**                          |
| **msoTriStateToggle**                         |
| **msoTrue**                                   |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例中，如果使用密码保护的文档的文件属性没有加密，则对密码加密选项进行设置。 function PasswordSettings() {     let tion = Application.ActivePresentation     if(tion.PasswordEncryptionFileProperties == msoFalse) {         tion.SetPasswordEncryptionOptions(" RSA SChannel Cryptographic Provider","RC4",56,true)     } }` |

#### **Presentation.PasswordEncryptionKeyLength**

返回一个 **Long** 类型的值，该值表示 WPP 通过密码加密文档时所使用算法的密钥长度。只读。

**语法**

**express.PasswordEncryptionKeyLength**

*express*   一个代表 **Presentation** 对象的变量。

**说明**

使用 **SetPasswordEncryptionOptions** 方法可指定WPP 通过密码加密文档所使用的算法。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例中，如果密码加密密钥的长度小于 40，则对密码加密选项进行设置。 function PasswordSettings() {     let tion = Application.ActivePresentation     if(tion.PasswordEncryptionKeyLength < 40) {         tion.SetPasswordEncryptionOptions(" RSA SChannel Cryptographic Provider","RC4",56,true)     } }` |

#### **Presentation.PasswordEncryptionProvider**

返回一个 **String** 类型的值，该值指定 WPP 通过密码加密文档时所使用的加密算法提供商的名称。只读。

**语法**

**express.PasswordEncryptionProvider**

*express*   一个代表 **Presentation** 对象的变量。

**说明**

使用 **SetPasswordEncryptionOptions** 方法可指定WPP 通过密码加密文档所使用的算法。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例中，如果使用的密码加密算法的提供商不是 RSA SChannel Cryptographic Provider，则对密码加密选项进行设置。 function PasswordSettings() {     let tion = Application.ActivePresentation     if(tion.PasswordEncryptionProvider != " RSA SChannel Cryptographic Provider") {         tion.SetPasswordEncryptionOptions(" RSA SChannel Cryptographic Provider","RC4",56,true)     } }` |

#### **Presentation.Path**

返回 **String**，该值代表到指定的 **Presentation** 对象的路径，只读。

**语法**

**express.Path**

*express*   一个代表 **Presentation** 对象的变量。

**说明**

如果使用此属性返回一个尚未保存的演示文稿的路径，则会得到一个空字符串。

路径中不包括最后的反斜线 (\) 或指定对象的名称。使用 **Presentation** 对象的 **Name** 属性可返回不带路径的文件名，使用 **FullName** 属性则可返回带路径的文件名。

#### **Presentation.Permission**

返回一个 **Permission** 对象，该对象可用于限制对活动演示文稿的权限，还可用于返回或设置特定权限设置。只读。

**语法**

**express.Permission**

*express*   一个代表 **Presentation** 对象的变量。

**说明**

使用 **Permission** 对象可限制对活动文档的权限，还可返回或设置特定权限设置。

使用 **Enabled** 属性可决定是否限制对活动文档的权限。使用 **Count** 属性可返回拥有权限的用户数量，使用 **RemoveAll** 方法可重置现有的所有权限。

**DocumentAuthor**、**EnableTrustedBrowser**、**RequestPermissionURL** 和 **StoreLicenses** 属性提供有关权限设置的其他信息。

**Permission** 对象分配对 **UserPermission** 对象集合的访问权限。使用 **UserPermission** 对象可将特定的权限集合与各个用户关联。当某些通过用户界面（如 **msoPermissionPrint**）分配的权限应用到所有用户时，可使用 **UserPermission** 对象为每个用户分配权限，并为每个用户指定权限的到期日期。

“信息版权管理”支持管理权限策略的使用，这些策略列出用户和组及其文档权限。使用 **ApplyPolicy** 方法可应用权限策略，使用 **PermissionFromPolicy**、**PolicyName** 和 **PolicyDescription** 属性可返回策略信息。

无论是否限制了对活动文档的权限，都可使用 **Permission**。活动文档没有受限权限时，**Presentation** 对象的 **Permission** 属性不返回 **Nothing**。使用 **Enabled** 属性可决定文档是否有受限权限。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//下面的示例创建一个新的演示文稿，并向电子邮件地址为“someone@example.com”的用户分配对新建演示文稿的读取权限。该示例将显示所有者和新用户的权限。 function AddUserPermissions() {     let myPres = Application.Presentations.Add(msoTrue)     let myPer = myPres.Permission     myPer.Enabled = true     let NewOwnerPer = myPer.Add("someone@example.com", msoPermissionRead )     alert(myPer.Item(1).UserId + " " + Str(myPer.Item(1).Permission))     alert(myPer.Item(2).UserId + " " + Str(myPer.Item(2).Permission)) }` |

#### **Presentation.PrintOptions**

返回一个 **PrintOptions** 对象，该对象代表与指定演示文稿共同保存的打印选项。只读。

**语法**

**express.PrintOptions**

*express*   一个代表 **Presentation** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例打印当前演示文稿中的隐藏幻灯片，并调整幻灯片的尺寸以适应纸张大小。 function test() {     let tion = Application.ActivePresentation     let ons = tion.PrintOptions     ons.PrintHiddenSlides = true     ons.FitToPage = true     tion.PrintOut() }` |

#### **Presentation.PublishObjects**

返回一个 **PublishObjects** 集合，该集合代表已完全或部分加载的演示文稿的集合，这些演示文稿可作为 HTML 进行发布。只读。

**语法**

**express.PublishObjects**

*express*   一个代表 **Presentation** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例将当前演示文稿的第三张幻灯片到第五张幻灯片发布为 HTML 格式，并将发布的幻灯片命名为“Mallard.htm”。 function test() {     let ject = Application.ActivePresentation.PublishObjects.Item(1)     ject.FileName = "C:\\Test\\Mallard.htm"     ject.SourceType = ppPublishSlideRange     ject.RangeStart = 3     ject.RangeEnd = 5     ject.Publish() }` |

#### **Presentation.ReadOnly**

返回指定的演示文稿是否为只读。只读 **MsoTriState** 类型。

**语法**

**express.ReadOnly**

*express*   一个代表 **Presentation** 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。 |
| --------------------------------------------- |
| **msoCTrue**                                  |
| **msoFalse**                                  |
| **msoTriStateMixed**                          |
| **msoTriStateToggle**                         |
| **msoTrue** 指定的演示文稿为只读。            |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果活动演示文稿为只读，本示例将其保存为文件“Newfile.ppt”。*/ function test() {     let pre = Application.ActivePresentation     if(pre.ReadOnly)      {         pre.SaveAs("newfile")     } }` |

#### **Presentation.RemovePersonalInformation**

如果该属性值为 **MsoTrue**，则 WPP 在保存演示文稿时从批注和修订中删除所有用户信息。**MsoTriState** 类型，可读/写。

**语法**

**express.RemovePersonalInformation**

*express*   一个代表 **Presentation** 对象的变量。

**说明**

| **MsoTriState** 可以是下列 **MsoTriState** 常量之一。 |
| ----------------------------------------------------- |
| **msoCTrue** 不应用于此属性。                         |
| **msoFalse** 批注、修订和个人信息保留在演示文稿中。   |
| **msoTriStateMixed** 不应用于该属性。                 |
| **msoTriStateToggle** 不应用于此属性。                |
| **msoTrue** 保存演示文稿时删除批注、修订和个人信息。  |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例设置在用户下次保存当前演示文稿时删除所有个人信息。 function RemovePersonalInfo() {     Application.ActivePresentation.RemovePersonalInformation = msoTrue }` |

#### **Presentation.Research**

返回一个 Research 对象，利用该对象可访问信息检索服务功能。只读。

**语法**

**express.Research**

*express*   一个代表 **Presentation** 对象的变量。

#### **Presentation.Saved**

决定演示文稿自上次保存以来是否进行过更改。可读/写 **MsoTriState** 类型。

**语法**

**express.Saved**

*express*   一个代表 **Presentation** 对象的变量。

**说明**

如果将已修改演示文稿的 **Saved** 属性设置为 **msoTrue**，则关闭演示文稿时不会提示用户保存所做的更改，并将丢失自上次保存后进行的所有更改。

| MsoTriState 可以是下列 MsoTriState 常量之一。      |
| -------------------------------------------------- |
| **msoCTrue**                                       |
| **msoFalse**                                       |
| **msoTriStateMixed**                               |
| **msoTriStateToggle**                              |
| **msoTrue** 演示文稿自上次保存以来没有进行过更改。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果当前演示文稿在上次保存后进行了更改，则以下示例保存当前演示文稿。*/ function test() {     let pre = Application.ActivePresentation     if(pre.Saved == false && pre.Path != "")      {         pre.Save()     } }` |

#### **Presentation.ServerPolicy**

返回一个ServerPolicy 对象。只读。

**语法**

**express.ServerPolicy**

*express*   一个代表 **Presentation** 对象的变量。

#### **Presentation.SharedWorkspace**

返回一个 **SharedWorkspace** 对象，该对象代表指定演示文稿所在的文档工作区。只读。

**语法**

**express.SharedWorkspace**

*express*   一个代表 **Presentation** 对象的变量。

**说明**

使用 **SharedWorkspace** 对象可以将当前的 WPP 演示文稿添加到服务器上的 Windows SharePoint Services 文档工作区网站中，以便利用工作区的协作功能，或者断开文档与工作区的连接或从工作区删除文档。使用 **SharedWorkspace** 对象的集合可以管理与共享文档关联的文件、文件夹、链接、成员和任务。

无论文档是否存储在工作区中，均可以使用 **SharedWorkspace** 对象模型。当文档不共享时，**Presentation** 对象的 **SharedWorkspace** 属性不返回 **Nothing**。使用 **SharedWorkspace** 对象的 **Connected** 属性可以确定当前演示文稿是否实际保存和连接到共享工作区。

用户必须有相应的权限才能使用 **SharedWorkspace** 对象层次结构中的对象、属性和方法。

使用 **SharedWorkspaceFiles** 集合（通过 **SharedWorkspace** 对象的 **Files** 属性访问）可以管理保存在共享工作区中的演示文稿和文件。

使用 **SharedWorkspaceFolders** 集合（通过 **SharedWorkspace** 对象的 **Folders** 属性访问）可以管理共享工作区主文档库文件夹中的子文件夹。

使用 **SharedWorkspaceLinks** 集合（通过 **SharedWorkspace** 对象的 **Links** 属性访问）可以管理附加文档和信息的链接，这些文档和信息是共享工作区中文档的协作成员所关心的。

使用 **SharedWorkspaceMembers** 集合（通过 **SharedWorkspace** 对象的 **Members** 属性访问）可以管理具有共享工作区参与权限的用户和具有工作区中保存的共享文档的协作权限的用户。

使用 **SharedWorkspaceTasks** 集合（通过 **SharedWorkspace** 对象的 **Tasks** 属性访问）可以管理分配给共享工作区中文档的协作成员的任务。

使用 **CreateNew** 方法可以新建文档工作区，并将活动文档添加到工作区中。使用 **Name** 和 **URL** 属性可以返回有关工作区的信息。

**SharedWorkspace** 对象通过服务器使用对象和属性的本地缓存。开发人员可能需要在执行某些操作之前更新此缓存，也可能需要将经过缓存的属性更改重新保存到服务器。使用 **SharedWorkspace** 对象的 **Refresh** 方法可以通过服务器刷新本地缓存，使用 **LastRefreshed** 属性可以确定上次刷新操作发生的时间。在本地修改了 **SharedWorkspaceLink** 和 **SharedWorkspaceTask** 对象的属性之后，使用这两类对象的 **Save** 方法可以将所做更改上载到服务器。

使用 **Disconnect** 方法可以断开活动文档的本地副本与共享工作区的连接，同时，使工作区中的共享副本保持原样。使用 **RemoveDocument** 方法可以将共享文档从共享工作区中彻底删除。

用户必须有相应的权限才能使用 **SharedWorkspace** 对象层次结构中的对象、属性和方法。向 **SharedWorkspaceMembers** 集合添加成员时，可以使用 *Role* 参数指定特定于每个工作区成员的权限集。

使用 **SharedWorkspace** 对象模型时可能会创建以下条件，即 **SharedWorkspace** 对象缓存与活动文档的**“共享工作区”**窗格中显示的用户界面不同步。例如，如果在**“共享工作区”**窗格打开的同时，**CreateNew** 方法以编程方式将活动文档添加到新的工作区中，则**“共享工作区”**窗格将继续显示**“新建”**按钮。在这种情况下，如果用户在**“共享工作区”**窗格进行的选择已不再有效，就会产生错误，此时，执行刷新操作以使显示与当前文档状态和共享工作区数据同步。

此外，**Presentation** 对象还具有返回 **Sync** 对象的 **Sync** 属性。使用 **Sync** 对象及其属性和方法可以管理共享文档的本地副本和服务器副本之间的同步。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//下面的示例返回一个对存储活动文档的文档工作区的引用。本示例假定活动文档属于文档工作区。 let objWorkspace = Application.ActivePresentation.SharedWorkspace` |

#### **Presentation.Signatures**

返回一个 **SignatureSet** 对象，该对象代表数字签名的集合。

**语法**

**express.Signatures**

*express*   一个代表 **Presentation** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//下面的代码行显示数字签名的个数。 function DisplayNumberOfSignatures() {     alert("Number of digital signatures: " + Application.ActivePresentation.Signatures.Count) }` |

#### **Presentation.SlideMaster**

返回一个 [**Master**](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Master/Master%20.htm#jsObject_Master) 对象，该对象代表幻灯片母版。

**语法**

**express.SlideMaster**

*express*   一个代表 **Presentation** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例设置当前演示文稿的幻灯片母版的背景图案。*/ Application.ActivePresentation.SlideMaster.Background.Fill.PresetTextured(Application.Enum.msoTextureGreenMarble)` |

#### **Presentation.Slides**

返回一个 **Slides** 集合，该集合代表指定演示文稿中的所有幻灯片。只读。

**语法**

**express.Slides**

*express*   一个代表 **Presentation** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例向活动演示文稿添加一张幻灯片。*/ Application.ActivePresentation.Slides.Add(1, Application.Enum.ppLayoutTitle)` |

#### **Presentation.SlideShowSettings**

返回一个 **SlideShowSettings** 对象，该对象代表指定演示文稿的幻灯片放映设置。只读。

**语法**

**express.SlideShowSettings**

*express*   一个代表 **Presentation** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例启动一个预设使用扬声器的幻灯片放映。运行该幻灯片放映时关闭动画和旁白。*/ function test() {     let ting = Application.ActivePresentation.SlideShowSettings     ting.ShowType = ppShowTypeSpeaker     ting.ShowWithNarration = false     ting.ShowWithAnimation = false     ting.Run() }` |

#### **Presentation.SlideShowWindow**

返回一个 **SlideShowWindow** 对象，该对象代表正在运行的指定演示文稿所在的幻灯片放映窗口。只读。

**语法**

**express.SlideShowWindow**

*express*   一个代表 **Presentation** 对象的变量。

#### **Presentation.SnapToGrid**

该属性值为 **MsoTrue** 时，在指定的演示文稿中将形状与网格线对齐。可读/写 **MsoTriState** 类型。

**语法**

**express.SnapToGrid**

*express*   一个代表 **Presentation** 对象的变量。

**说明**

| MsoTriState 可以是下列 MsoTriState 常量之一。 |
| --------------------------------------------- |
| **msoCTrue**                                  |
| **msoFalse**                                  |
| **msoTriStateMixed**                          |
| **msoTriStateToggle**                         |
| **msoTrue**                                   |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例切换在活动演示文稿中形状与网格线的对齐设置。 function ToggleSnapToGrid() {     let tion = Application.ActivePresentation     if(tion.SnapToGrid == msoTrue) {         tion.SnapToGrid = msoFalse     }     else {         tion.SnapToGrid = msoTrue     } }` |

#### **Presentation.Sync**

返回一个 **Sync** 对象，使用该对象可对存储在 Windows SharePoint Services 共享工作区中的共享演示文稿的本地副本和服务器副本的同步进行管理。只读。

**语法**

**express.Sync**

*express*   一个代表 **Presentation** 对象的变量。

**说明**

**Sync** 对象的 **Status** 属性返回有关同步的当前状态的重要信息。使用 **GetUpdate** 方法可刷新同步状态。使用 **LastSyncTime**、**ErrorType** 和 **WorkspaceLastChangedBy** 属性可返回其他信息。

有关共享演示文稿的本地副本和服务器副本之间可存在的差异和冲突的详细信息，请参阅 **Status** 属性。

使用 **PutUpdate** 方法可将本地更改保存到服务器。未进行任何本地更改时，关闭并重新打开文档可从服务器检索最新版本。使用 **ResolveConflict** 方法可解决本地副本和服务器副本之间的差异，或者使用 **OpenVersion** 方法在当前打开文档的本地版本的情况下再打开其他版本。

**Sync** 对象的 **GetUpdate**、**PutUpdate** 和 **ResolveConflict** 方法不返回状态代码，因为它们以异步方式完成任务。**Sync** 对象通过单个事件，即 **Application** 对象的 **PresentationSync** 事件，提供重要的状态信息。

上述 **Sync** 事件返回一个 **msoSyncEventType** 类型的值。

| MsoSyncEventType 可以是下列 **msoSyncEventType** 常量之一。 |
| ----------------------------------------------------------- |
| **msoSyncEventDownloadInitiated** (0)                       |
| **msoSyncEventDownloadSucceeded** (1)                       |
| **msoSyncEventDownloadFailed** (2)                          |
| **msoSyncEventUploadInitiated** (3)                         |
| **msoSyncEventUploadSucceeded** (4)                         |
| **msoSyncEventUploadFailed** (5)                            |
| **msoSyncEventDownloadNoChange** (6)                        |
| **msoSyncEventOffline** (7)                                 |

无论活动文档上启用或禁用共享和同步，都可使用 **Sync** 对象模型。未共享活动文档或未启用同步时，**Presentation** 对象的 **Sync** 属性不返回 **Nothing**。使用 **Status** 属性可决定是否共享文档以及是否启用同步。

并不是所有文档同步问题都会导致可捕获的运行时错误。在使用了 **Sync** 对象的方法后，最好对 **Status** 属性进行检查。如果 **Status** 属性为 **msoSyncStatusError**，请检查 **ErrorType** 属性以获取有关所发生的错误类型的其他信息。

在许多情况下，解决错误条件的最好方法是调用 **GetUpdate** 方法。例如，如果对 **PutUpdate** 的调用导致错误条件的产生，则对 **GetUpdate** 的调用会将状态重置为 **msoSyncStatusLocalChanges**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//下面的示例显示最后一个修改活动演示文稿的人的名称（如果活动演示文稿是文档工作区中的共享文档）。 let eStatus = Application.ActivePresentation.Sync.Status if(eStatus == msoSyncStatusLatest) {     strLastUser = Application.ActivePresentation.Sync.WorkspaceLastChangedBy     alert("You have the most up-to-date copy.This file was last modified by " + strLastUser) }` |

#### **Presentation.Tags**

返回一个代表指定对象的标签的 **Tags** 对象。只读。

**语法**

**express.Tags**

*express*   一个代表 **Presentation** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//以下示例向当前演示文稿的第一张幻灯片中添加名为“REGION”和“PRIORITY”的标签。 function test() {     let tag = Application.ActivePresentation.Slides.Item(1).Tags     tag.Add("Region", "East")   //Adds "Region" tag with value "East"     tag.Add("Priority", "Low")    //Adds "Priority" tag with value "Low" }` |

#### **Presentation.TemplateName**

返回与指定演示文稿相关的设计模板的名称。只读 **String** 类型。

**语法**

**express.TemplateName**

*express*   一个代表 **Presentation** 对象的变量。

**说明**

返回的字符串包含 MS-DOS 文件扩展名（对于已注册的文件类型），但不包括完整路径。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例中，如果设计模板“Professional.pot”尚未应用于演示文稿“Pres1.ppt”，则将该模板应用于该演示文稿。*/ function test() {     let tion = Application.Presentations.Item("Pres1.ppt")     if(tion.TemplateName != "Professional.pot")      {         tion.ApplyTemplate("C:\\program files\\office\\templates\\presentation designs\\Professional.pot")     } }` |

#### **Presentation.TitleMaster**

返回 [Master](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%BC%94%E7%A4%BA%20API%20%E5%8F%82%E8%80%83/Master/Master%20.htm#jsObject_Master)对象，该对象代表指定演示文稿的标题母版。如果演示文稿没有标题母版，则会产生一个错误。

**语法**

**express.TitleMaster**

*express*   一个代表 **Presentation** 对象的变量。

**说明**

可以使用 **AddTitleMaster**方法为演示文稿添加一个标题母版。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例中，如果当前演示文稿有一个标题母版，则为该标题母版设置页脚文本。*/ function test() {     let tion = Application.ActivePresentation     if(tion.HasTitleMaster)      {         tion.TitleMaster.HeadersFooters.Footer.Text = "Introduction"     } }` |

#### **Presentation.WebOptions**

返回一个 **WebOptions** 对象，该对象中包含演示文稿级别的属性，当以网页的方式发布或保存整个或部分演示文稿，或打开某个网页时，WPP 将使用这些属性。只读。

**语法**

**express.WebOptions**

*express*   一个代表 **Presentation** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `//本示例的功能：当以网页保存或发布当前演示文稿时，允许使用可移植网络图形 (PNG)，并将大纲窗格的文字颜色设置为白色，将大纲窗格和幻灯片窗格的背景色设置为黑色。 function test() {     let tion = Application.ActivePresentation.WebOptions     tion.FrameColors = ppFrameColorsWhiteTextOnBlack     tion.AllowPNG = true }` |

#### **Presentation.Windows**

返回一个 **DocumentWindows** 集合，该集合代表与指定演示文稿相关联的所有文档窗口。该属性不返回与该演示文稿相关联的任何幻灯片放映窗口。只读。

**语法**

**express.Windows**

*express*   一个代表 **Presentation** 对象的变量。

#### **Presentation.WritePassword**

设置或返回一个 **String** 类型的值，该值代表保存对指定文档所做的更改时使用的密码。可读/写。

**语法**

**express.WritePassword**

*express*   一个代表 **Presentation** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置保存对活动演示文稿所做的更改时使用的密码。*/ Application.ActivePresentation.WritePassword = "complexstrPWD"` |

适用环境：web

适用平台：windows/linux