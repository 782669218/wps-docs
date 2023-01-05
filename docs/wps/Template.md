#### **Template**



代表文档模板。**Template** 对象是 **Templates** 集合的一个成员。**Templates** 集合包含所有可用的 **Template** 对象。

**说明**

可以使用 **Templates**(*Index*) 返回单个 **Template** 对象，其中 *Index* 是模板名称或索引编号。如果 Memo2.dot 模板在 **Templates** 集合中，以下示例将保存该模板。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `for(let i = 1; i <= Application.Templates.Count; i++){     if(Application.Templates.Item(i).Name.toLowerCase() == "memo2.dot"){         Application.Templates.Item(i).Save()     } }` |

索引号代表模板在 

Templates

 集合中的位置。以下示例打开 

Templates

 集合中的第一个模板。

| 示例代码复制                                     |
| ------------------------------------------------ |
| `Application.Templates.Item(1).OpenAsDocument()` |

Add

 方法不可用于 

Templates

 集合。可以改用下列方式之一向 

Templates

 集合添加模板：

- 使用 **Open** 方法和 **Documents** 集合打开基于模板的文档或模板
- 使用 **Add** 方法和 **Documents** 集合打开基于模板的新文档
- 使用 **Add** 方法和 **Addins** 集合加载共用模板
- 使用 **AttachedTemplate** 属性和 **Document** 对象将模板附加到文档

使用 **NormalTemplate** 属性可返回代表 Normal 模板的模板对象。使用 **AttachedTemplate** 属性可返回附加到指定文档的模板。

可以使用 **DefaultFilePath** 属性返回或设置用户或工作组模板的位置（即，用于存储这些模板的文件夹）。以下示例在**“选项”**对话框（通过**“工具”**菜单打开）中的**“文件位置”**选项卡上显示用户模板文件夹。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `alert(Application.Options.DefaultFilePath(wdUserTemplatesPath))` |

**方法**

|                                                              | 名称               | 说明                                                 |
| ------------------------------------------------------------ | ------------------ | ---------------------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **OpenAsDocument** | 将指定模板作为文档打开并返回一个 **Document** 对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Save**           | 保存指定的模板。                                     |

**属性**

|                                                              | 名称                          | 说明                                                         |
| ------------------------------------------------------------ | ----------------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**               | 返回一个代表 WPS 应用程序的 **Application** 对象。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BuildingBlockEntries**      | 返回一个 **BuildingBlockEntries** 集合，该集合代表模板中构建基块条目的集合。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BuildingBlockTypes**        | 返回一个 **BuildingBlockTypes** 集合，该集合代表模板中包含的构建基块类型的集合。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **BuiltInDocumentProperties** | 返回一个 **DocumentProperties** 集合，该集合代表指定文档的所有内置的文档属性。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**                   | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CustomDocumentProperties**  | 返回一个 **DocumentProperties** 集合，该集合代表指定文档的所有自定义文档属性。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FarEastLineBreakLanguage**  | 返回或设置在指定的文档或模板中对文本进行换行时使用的东亚语言。**WdFarEastLineBreakLanguageID** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FarEastLineBreakLevel**     | 返回或设置指定文档的换行控制级别。**WdFarEastLineBreakLevel** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **FullName**                  | 指定模板的名称，其中包含驱动器或 Web 路径。**String** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **JustificationMode**         | 返回或设置指定模板的字符间距调整。可读写 **WdJustificationMode** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **KerningByAlgorithm**        | 如果 WPS 调整指定文档中的半角西文字符与标点符号的字距，则该属性值为 **True**。可读写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LanguageID**                | 返回或设置一个 **WdLanguageID** 常量，该常量代表指定范围的语言。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LanguageIDFarEast**         | 返回或设置指定对象的东亚语言。可读写 **WdLanguageID** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ListTemplates**             | 返回一个 **ListTemplates** 集合，该集合代表指定文档、模板或列表库的所有列表格式。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**                      | 返回指定对象的名称。**String** 类型，只读。                  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NoLineBreakAfter**          | 返回或设置 WPS 不在其后面进行换行的首尾字符。可读写 **String** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NoLineBreakBefore**         | 返回或设置 WPS 不在其前面进行换行的首尾字符。可读写 **String** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NoProofing**                | 如果拼写和语法检查程序忽略基于此模板的文档，则该属性值为 **True**。可读写 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**                    | 返回一个 **Object** 类型值，该值代表指定 **Template** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Path**                      | 返回指定文档模板的路径。只读 **String** 类型。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Saved**                     | 如果自上次保存之后指定的模板未更改，则该属性值为 **True**。如果关闭文档时 WPS 提示保存对文档所做的更改，则该属性值为 **False**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**                      | 返回模板的类型。只读 [WdTemplateType ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdTemplateType%20%E6%9E%9A%E4%B8%BE.html)类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **VBProject**                 | 返回指定模板的 **VBProject** 对象。                          |

**成员方法**

#### **Template.OpenAsDocument**

将指定模板作为文档打开并返回一个 **Document** 对象。

**语法**

**express.OpenAsDocument ()**

*express*   一个代表 **Template** 对象的变量。

**返回值**

Document

**说明**

将模板作为文档打开则允许编辑该模板的内容。当 **Template** 对象中的某属性或方法（例如 **Styles** 属性）无效时，就可能需要使用这种方法。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例打开活动文档所选用的模板，如果该模板包含了不止一个段落标记的内容，则显示一个消息框，然后关闭该模板。*/ function test() { let docNew = ActiveDocument.AttachedTemplate.OpenAsDocument()  if(docNew.Content.Text != "\r"){     MsgBox("Template is not empty") } else{     MsgBox("Template is empty") }  docNew.Close(wdDoNotSaveChanges) }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为 Normal 模板保存一个副本“Backup.dot”。*/ function test() { let docNew = NormalTemplate.OpenAsDocument() docNew.SaveAs("Backup.dot") docNew.Close(wdDoNotSaveChanges) }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例更改活动文档所选用的模板中“标题 1”的样式设置。UpdateStyles 方法用来更新活动文档中的样式。*/ function test() { let docNew = ActiveDocument.AttachedTemplate.OpenAsDocument() let fontStyle = docNew.Styles.Item(wdStyleHeading1).Font  fontStyle.Name = "Arial" fontStyle.Size = 16 fontStyle.Bold = false  docNew.Close(wdSaveChanges) ActiveDocument.UpdateStyles() }` |

#### **Template.Save**

保存指定的模板。

**语法**

**express.Save()**

*express*   一个代表 **Template** 对象的变量。

**说明**

如果之前未保存过此模板，则**“另存为”**对话框将提示用户输入文件名。

**成员属性**

#### **Template.Application**

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

*express*   一个代表 **Template** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **Template.BuildingBlockEntries**

返回一个 **BuildingBlockEntries** 集合，该集合代表模板中构建基块条目的集合。只读。

**语法**

**express.BuildingBlockEntries**

*express*   一个代表 **Template** 对象的变量。

#### **Template.BuildingBlockTypes**

返回一个 **BuildingBlockTypes** 集合，该集合代表模板中包含的构建基块类型的集合。只读。

**语法**

**express.BuildingBlockTypes**

*express*   一个代表 **Template** 对象的变量。

#### **Template.BuiltInDocumentProperties**

返回一个 **DocumentProperties** 集合，该集合代表指定文档的所有内置的文档属性。

**语法**

**express.BuiltInDocumentProperties**

*express*   一个代表 **Template** 对象的变量。

**说明**

要返回单个代表特定内置文档属性的 **DocumentProperty** 对象，可使用 **BuiltinDocumentProperties** 属性。如果 WPS 没有为一个内置的文档属性定义一个值，则读取这个文档属性的 **Value** 属性时会产生一个错误。

用 **CustomDocumentProperties** 属性返回自定义文档属性的集合。

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

#### **Template.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Template** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **Template.CustomDocumentProperties**

返回一个 **DocumentProperties** 集合，该集合代表指定文档的所有自定义文档属性。

**语法**

**express.CustomDocumentProperties**

*express*   一个代表 **Template** 对象的变量。

**说明**

使用 **BuiltInDocumentProperties** 属性可返回内置文档属性的集合。**msoPropertyTypeString** 类型的属性的长度不能超过 255 个字符。

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

#### **Template.FarEastLineBreakLanguage**

返回或设置在指定的文档或模板中对文本进行换行时使用的东亚语言。**WdFarEastLineBreakLanguageID** 类型，可读写。

**语法**

**express.FarEastLineBreakLanguage**

*express*   一个代表 **Template** 对象的变量。

#### **Template.FarEastLineBreakLevel**

返回或设置指定文档的换行控制级别。**WdFarEastLineBreakLevel** 类型，可读写。

**语法**

**express.FarEastLineBreakLevel**

*express*   一个代表 **Template** 对象的变量。

**说明**

如果 **FarEastLineBreakControl** 属性设为 **False**，则忽略该属性。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例设置 WPS 在活动文档中第一级首尾字符处换行。*/ ActiveDocument.FarEastLineBreakLevel = wdJustificationModeCompressKana` |

#### **Template.FullName**

指定模板的名称，其中包含驱动器或 Web 路径。**String** 类型，只读。

**语法**

**express.FullName**

*express*   一个代表 **Template** 对象的变量。

**说明**

使用该属性与按顺序使用 **Path、PathSeparator** 和 **Name** 属性的作用相同。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示活动文档附加模板的路径和文件名*/ function TemplateName(){     alert(Application.ActiveDocument.AttachedTemplate.FullName) }` |

#### **Template.JustificationMode**

返回或设置指定模板的字符间距调整。可读写 **WdJustificationMode** 类型。

**语法**

**express.JustificationMode**

*express*   一个代表 **Template** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例设置 WPS 在调整字符间距时只压缩标点符号。*/ NormalTemplate.JustificationMode = wdJustificationModeCompressKana` |

#### **Template.KerningByAlgorithm**

如果 WPS 调整指定文档中的半角西文字符与标点符号的字距，则该属性值为 **True**。可读写 **Boolean** 类型。

**语法**

**express.KerningByAlgorithm**

*express*   一个代表 **Template** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例设置 WPS 调整 Normal 模板中的半角西文字符与标点符号的字距。*/ NormalTemplate.KerningByAlgorithm = true` |

#### **Template.LanguageID**

返回或设置一个 **WdLanguageID** 常量，该常量代表指定范围的语言。可读写。

**语法**

**express.LanguageID**

*express*   一个代表 **Template** 对象的变量。

**说明**

根据您所选择或安装的语言支持（例如，美国英语），某些 **WdLanguageID** 常量可能不可用。

#### **Template.LanguageIDFarEast**

返回或设置指定对象的东亚语言。可读写 **WdLanguageID** 类型。

**语法**

**express.LanguageIDFarEast**

*express*   一个代表 **Template** 对象的变量。

**说明**

推荐使用本方法，以返回或设置在 WPS 东亚版中创建的文档中的东亚文字的语言。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将所选内容的语言设置为朝鲜语。*/ NormalTemplate.LanguageIDFarEast = wdKorean` |

#### **Template.ListTemplates**

返回一个 **ListTemplates** 集合，该集合代表指定文档、模板或列表库的所有列表格式。只读。

**语法**

**express.ListTemplates**

*express*   一个代表 **Template** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

#### **Template.Name**

返回指定对象的名称。**String** 类型，只读。

**语法**

**express.Name**

*express*   一个代表 **Template** 对象的变量。

#### **Template.NoLineBreakAfter**

返回或设置 WPS 不在其后面进行换行的首尾字符。可读写 **String** 类型。

**语法**

**express.NoLineBreakAfter**

*express*   一个代表 **Template** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例实现的功能是：在活动文档中，设置“$”、“(”、“[”、“\”和“{”作为首尾字符时，WPS 不在这些首尾字符的后面进行换行。*/ ActiveDocument.NoLineBreakAfter = "$([\{"` |

#### **Template.NoLineBreakBefore**

返回或设置 WPS 不在其前面进行换行的首尾字符。可读写 **String** 类型。

**语法**

**express.NoLineBreakBefore**

*express*   一个代表 **Template** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例实现的功能是：在活动文档中，设置“!”、“)”和“]”作为首尾字符时， WPS 不在这些首尾字符的前面进行换行。*/ NormalTemplate.NoLineBreakBefore = "!)]"` |

#### **Template.NoProofing**

如果拼写和语法检查程序忽略基于此模板的文档，则该属性值为 **True**。可读写 **Long** 类型。

**语法**

**express.NoProofing**

*express*   一个代表 **Template** 对象的变量。

#### **Template.Parent**

返回一个 **Object** 类型值，该值代表指定 **Template** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **Template** 对象的变量。

#### **Template.Path**

返回指定文档模板的路径。只读 **String** 类型。

**语法**

**express.Path**

*express*   一个代表 **Template** 对象的变量。

**说明**

该路径不包括尾部字符，例如，“C:\WPSOffice”或“http://MyServer”。使用 **PathSeparator** 属性可添加分隔文件夹与驱动器号的字符。使用 **Name** 属性可返回不带路径的文件名，而使用 **FullName** 属性可同时返回文件名和路径。

#### **Template.Saved**

如果自上次保存之后指定的模板未更改，则该属性值为 **True**。如果关闭文档时 WPS 提示保存对文档所做的更改，则该属性值为 **False**。**Boolean** 类型，可读写。

**语法**

**express.Saved**

*express*   一个代表 **Template** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将 Normal 模板的状态设置为未更改。如果更改了 Normal 模板，则退出 WPS 时不保存所做的更改。*/ function test() { NormalTemplate.Saved = true Application.Quit() }` |

#### **Template.Type**

返回模板的类型。只读 [WdTemplateType ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdTemplateType%20%E6%9E%9A%E4%B8%BE.html)类型。

**语法**

**express.Type**

*express*   一个代表 **Template** 对象的变量。

#### **Template.VBProject**

返回指定模板的 **VBProject** 对象。

**语法**

**express.VBProject**

*express*   一个代表 **Template** 对象的变量。

**说明**

使用该属性可访问代码模块和用户窗体。

要在“对象浏览器”中查看 **VBProject** 对象，必须选中“Visual Basic 编辑器”的**“引用”**对话框（**“工具”**菜单）中的**“示例代码 Extensibility”**复选框。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示 Normal 模板的 Visual Basic 项目的名称。*/ function test() { Set normProj = NormalTemplate.VBProject MsgBox normProj.Name }  ` |

适用环境：web

适用平台：windows/linux