#### **Documents**



当前在 WPS 中打开的所有 **Document** 对象的集合。

**方法**

|                                                              | 名称                   | 说明                                                         |
| ------------------------------------------------------------ | ---------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**                | 返回一个 **Document** 对象，该对象代表添加到打开文档集合中的新建空文档。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **AddBlogDocument**    | 返回一个 **Document** 对象，该对象表示 WPS 向前三个参数所描述的帐户发布的新博客文档。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **CanCheckOut**        | 如果该属性值为 **True**，则 WPS 可将指定的文档从服务器签出。可读/写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Close**              | 关闭指定的文档。                                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item**               | 返回集合中的单个 **Document**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Document/Document%20.htm#jsObject_Document)对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Open**               | 打开指定的文档并将其添加到 **Documents** 集合。返回一个 **Document** 对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **OpenNoRepairDialog** | 打开指定的文档并将其添加到 Documents 集合。                  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Parent**             | 返回一个 **Object** 类型值，该值代表指定 **Documents** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Save**               | 保存 **Documents** 集合中的所有文档。                        |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Add**         | 返回一个 **Document** 对象，该对象代表添加到打开文档集合中的新建空文档。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个代表 WPS 应用程序的 **Application** 对象。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回一个 **Number** 类型的值，该值代表集合中的文档数。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |

**成员方法**

#### **Documents.Add**

返回一个 **Document** 对象，该对象代表添加到打开文档集合中的新建空文档。

**语法**

**express.Add(Template, NewTemplate, DocumentType, Visible)**

*express*   一个代表 **Documents** 对象的变量。

**参数**

| **名称**       | **必选/可选** | **数据类型** | **说明**                                                     |
| -------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Template*     | 可选          | **Variant**  | 要用于新文档的模板名。如果省略该参数，则使用 Normal 模板。   |
| *NewTemplate*  | 可选          | **Variant**  | 如果该属性值为 True，则将文档作为模板打开。默认值为 False。  |
| *DocumentType* | 可选          | **Variant**  | 可以是下列 WdNewDocumentType 常量之一：wdNewBlankDocument、wdNewEmailMessage、wdNewFrameset 或 wdNewWebPage。默认常量是 wdNewBlankDocument。 |
| *Visible*      | 可选          | **Variant**  | 如果该属性值为 True，将在可见窗口中打开文档。如果该属性值为 False，则 WPS 打开此文档，但将文档窗口的 Visible 属性设置为 False。默认值为 True。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例基于 Normal 模板新建一篇文档*/ Application.Documents.Add()` |

#### **Documents.AddBlogDocument**

返回一个 **Document** 对象，该对象表示 WPS 向前三个参数所描述的帐户发布的新博客文档。

**语法**

**express.AddBlogDocument(ProviderID, PostURL, BlogName, PostID)**

*express*   一个代表 **Documents** 对象的变量。

**参数**

| **名称**     | **必选/可选** | **数据类型** | **说明**                                                     |
| ------------ | ------------- | ------------ | ------------------------------------------------------------ |
| *ProviderID* | 必选          | **String**   | GUID，即提供商在向 WPS 进行注册时所用的唯一值。              |
| *PostURL*    | 必选          | **String**   | 用于向博客添加帖子的 URL。                                   |
| *BlogName*   | 必选          | **String**   | 将用于 WPS 中的博客的显示名称。                              |
| *PostID*     | 可选          | **String**   | 用来填充使用 AddBlogDocument 方法创建的文档的现有张贴内容的 ID。 |

**说明**

此方法不仅创建一个新文档，还会向 WPS 注册指定的博客帐户（如果未注册）。此外，如果指定了 *PostID* 参数，则从提供商网站中用 *PostID* 参数的值所指定的张贴内容来填充新文档。

#### **Documents.CanCheckOut**

如果该属性值为 **True**，则 WPS 可将指定的文档从服务器签出。可读/写 **Boolean** 类型。

**语法**

**express.CanCheckOut(FileName)**

*express*   一个代表 **Documents** 对象的变量。

**参数**

| **名称**   | **必选/可选** | **数据类型** | **说明**                   |
| ---------- | ------------- | ------------ | -------------------------- |
| *FileName* | 必选          | **String**   | 文档的服务器路径和文件名。 |

**返回值**

Boolean

**说明**

要利用 WPS 中内置的协作功能，文档必须存储在 Microsoft SharePoint Portal Server 上。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例验证没有其他用户编辑文档，并可将其签出。如果可以签出文档，则将文档复制到本地计算机进行编辑。*/ function CheckDocInOut(docCheckOut) { if(Documents.CanCheckOut(docCheckOut) == true) {         Documents.CheckOut(docCheckOut)     }     else {         MsgBox("You are unable to check out this document at this time.")     } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*若要调用 CheckInOut 子程序，请使用下列子程序并将文件名“http://servername/workspace/report.doc”替换为以上“说明”一节中提到的服务器上实际的文件。 */ function CheckDocInOut() {     Call CheckInOut("http://servername/workspace/report.doc") }` |

#### **Documents.Close**

关闭指定的文档。

**语法**

**express.Close(SaveChanges, OriginalFormat, RouteDocument)**

*express*   一个代表 **Documents** 对象的变量。

**参数**

| **名称**         | **必选/可选** | **数据类型** | **说明**                                                     |
| ---------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *SaveChanges*    | 可选          | **Variant**  | 指定保存文档的操作。可以是下列 WdSaveOptions 常量之一：wdDoNotSaveChanges、wdPromptToSaveChanges 或 wdSaveChanges。 |
| *OriginalFormat* | 可选          | **Variant**  | 指定保存文档的格式。可以是下列 WdOriginalFormat 常量之一：wdOriginalDocumentFormat、wdPromptUser 或 wdWordDocument。 |
| *RouteDocument*  | 可选          | **Variant**  | 如果为 True，则将文档传送给下一个收件人。如果文档没有附加的传送名单，则忽略该参数。 |

#### **Documents.Item**

返回集合中的单个 **Document**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Document/Document%20.htm#jsObject_Document)对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **Documents** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Index*  | 可选          | **Number**   | 要返回的单个对象。可以是代表单个对象序号位置的 Number 类型值。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例显示 Documents 集合的第一篇文档的名称 */ Application.Documents.Item(1).Name` |

#### **Documents.Open**

打开指定的文档并将其添加到 **Documents** 集合。返回一个 **Document** 对象。

**语法**

**express.Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate, Revert, WritePasswordDocument, WritePasswordTemplate, Format, Encoding, Visible, OpenConflictDocument, OpenAndRepair, DocumentDirection, NoEncodingDialog)**

*express*   一个代表 **Documents** 对象的变量。

**参数**

| **名称**                | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *FileName*              | 必选          | **Variant**  | 文档名（可包含路径）。                                       |
| *ConfirmConversions*    | 可选          | **Variant**  | 如果该属性为 True，则当文件不是 WPS 格式时，将显示“文件转换”对话框。 |
| *ReadOnly*              | 可选          | **Variant**  | 如果该属性值为 True，则以只读方式打开文档。该参数不会覆盖保存的文档的只读建议设置。例如，如果文档在只读建议启用的情况下保存，则将 ReadOnly 参数设置为 False 不会导致文件以可读写方式打开。 |
| *AddToRecentFiles*      | 可选          | **Variant**  | 如果该属性值为 True，则将文件名添加到“文件”菜单底部最近使用过的文件列表中。 |
| *PasswordDocument*      | 可选          | **Variant**  | 打开文档时所需的密码。                                       |
| *PasswordTemplate*      | 可选          | **Variant**  | 打开模板时所需的密码。                                       |
| *Revert*                | 可选          | **Variant**  | 控制当 FileName 是一篇打开的文档的名称时应采取的操作。如果该属性值为 True，则放弃对打开文档进行的所有尚未保存的更改，并将重新打开该文件。如果该属性值为 False，则激活打开的文档。 |
| *WritePasswordDocument* | 可选          | **Variant**  | 用于保存文档更改的密码。                                     |
| *WritePasswordTemplate* | 可选          | **Variant**  | 用于保存模板更改的密码。                                     |
| *Format*                | 可选          | **Variant**  | 用于打开文档的文件转换器。可以是 WdOpenFormat 常量之一。默认值为 wdOpenFormatAuto。 |
| *Encoding*              | 可选          | **Variant**  | 当您查看保存的文档时 WPS 所使用的文档编码（代码页或字符集）。可以是任何有效的 MsoEncoding 常量。要查看有效 MsoEncoding 常量的列表，请参阅“Visual Basic 编辑器”中的“对象浏览器”。默认值为系统代码页。 |
| *Visible*               | 可选          | **Variant**  | 如果该属性值为 True，则可在可见窗口中打开文档。默认值为 True。 |
| *OpenConflictDocument*  | 可选          | **Variant**  | 指定是否打开具有脱机冲突的文档的冲突文件。                   |
| *OpenAndRepair*         | 可选          | **Variant**  | 如果该属性为 True，则修复文档，以防止文档毁坏。              |
| *DocumentDirection*     | 可选          | **Variant**  | 表示文档中的横排文字。默认值为 wdLeftToRight。               |
| *NoEncodingDialog*      | 可选          | **Variant**  | 如果该属性值为 True，则跳过显示当文字编码无法识别时 WPS 所显示的“编码”对话框。默认值为 False。 |

**说明**

请避免在应用程序中使用硬编码的密码。如果过程中要求输入密码，请从用户处请求密码，并将密码作为变量存储起来，然后在代码中使用该变量。有关如何做到这一点的建议最佳操作，请参阅 WPS Office解决方案开发人员安全注意事项。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例以只读方式打开文档 MyDoc.doc*/ Application.Documents.Open("C:\\MyDoc.doc", null, true)` |

#### **Documents.OpenNoRepairDialog**

打开指定的文档并将其添加到 Documents 集合。

**语法**

**express.OpenNoRepairDialog(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate, Revert, WritePasswordDocument, WritePasswordTemplate, Format, Encoding, Visible, OpenAndRepair, DocumentDirection, NoEncodingDialog, XMLTransform)**

*express*   一个代表 **Documents** 对象的变量。

**参数**

| **名称**                | **必选/可选** | **数据类型** | **说明**                                                     |
| ----------------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *FileName*              | 必选          | **Variant**  | 文档名（可包含路径）。                                       |
| *ConfirmConversions*    | 可选          | **Variant**  | 如果该属性值为 True，则当文件不是 WPS 格式时，将显示“转换文件”对话框。 |
| *ReadOnly*              | 可选          | **Variant**  | 如果值为 True，则以只读方式打开文档。该参数不会覆盖已保存文档的只读建议设置。例如，如果在只读建议设置打开的情况下保存了文档，则将 ReadOnly 参数设置为 False 不会导致文件以可读/写方式打开。 |
| *AddToRecentFiles*      | 可选          | **Variant**  | 如果该属性值为 True，则将文件名添加到“文件”菜单底部最近使用过的文件列表中。 |
| *PasswordDocument*      | 可选          | **Variant**  | 打开文档时所需的密码。                                       |
| *PasswordTemplate*      | 可选          | **Variant**  | 打开模板时所需的密码。                                       |
| *Revert*                | 可选          | **Variant**  | 控制当 FileName 是一篇打开的文档的名称时应采取的操作。如果该属性值为 True，则放弃对打开文档进行的所有尚未保存的更改，并将重新打开该文件。如果该属性值为 False，则激活打开的文档。 |
| *WritePasswordDocument* | 可选          | **Variant**  | 用于保存文档更改的密码。                                     |
| *WritePasswordTemplate* | 可选          | **Variant**  | 用于保存模板更改的密码。                                     |
| *Format*                | 可选          | **Variant**  | 用于打开文档的文件转换器。可以是 WdOpenFormat 常量之一。默认值为 wdOpenFormatAuto。 |
| *Encoding*              | 可选          | **Variant**  | 当您查看已保存文档时 WPS 所使用的文档编码（代码页或字符集）。可以是任何有效的 MsoEncoding 常量。要查看有效 MsoEncoding 常量的列表，请参阅“Visual Basic 编辑器中的对象浏览器”。默认为系统代码页。 |
| *Visible*               | 可选          | **Variant**  | 如果值为 True，则在可见窗口中打开文档。默认值为 True。       |
| *OpenAndRepair*         | 可选          | **Variant**  | 如果值为 True，则修复文档，以防止文档毁坏。                  |
| *DocumentDirection*     | 可选          | **Variant**  | 表示文档中文字的水平排列方式。可以是任何有效的 WdDocumentDirection 常量。默认值为 wdLeftToRight。 |
| *NoEncodingDialog*      | 可选          | **Variant**  | 如果值为 True，则跳过显示当无法识别文字编码时 WPS 所显示的“编码”对话框。默认值为 False。 |
| *XMLTransform*          | 可选          | **Variant**  | 指定要使用的转换。                                           |

**返回值**

一个代表指定文档的 Document 对象

**说明**

## 安全性

请避免在应用程序中使用硬编码的密码。如果过程中要求输入密码，请从用户处请求密码，并将密码作为变量存储起来，然后在代码中使用该变量。有关如何做到这一点的建议最佳操作，请参阅 WPS Office解决方案开发人员安全注意事项。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例以只读方式打开文档 MyDoc.doc。*/ function OpenDoc() {     Documents.OpenNoRepairDialog("C:\\MyFiles\\MyDoc.doc", null, true) }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例利用 WordPerfect 6.x 文件转换器打开 Test.wp。*/ function OpenDoc2() {     let fmt = Application.FileConverters.Item("WordPerfect6x").OpenFormat     Documents.OpenNoRepairDialog("C:\\MyFiles\\Test.wp", null, null, null, null, null, null, null, null, fmt) }` |

#### **Documents.Parent**

返回一个 **Object** 类型值，该值代表指定 **Documents** 对象的父对象。

**语法**

**express.Parent()**

*express*   一个代表 **Documents** 对象的变量。

#### **Documents.Save**

保存 **Documents** 集合中的所有文档。

**语法**

**express.Save(NoPrompt, OriginalFormat)**

*express*   一个代表 **Documents** 对象的变量。

**参数**

| **名称**         | **必选/可选** | **数据类型** | **说明**                                                     |
| ---------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *NoPrompt*       | 可选          | **Variant**  | 如果该属性值为 True，则 WPS 将自动保存所有文档。如果该属性值为 False， WPS 将提示用户保存自上次保存后已修改过的每篇文档。 |
| *OriginalFormat* | 可选          | **Variant**  | 指定文档的保存方式。可以是 WdOriginalFormat 常量之一。       |

**说明**

如果此文档以前未保存过，则**“另存为”**对话框将提示用户键入文件名。

**成员属性**

#### **Documents.Add**

返回一个 **Document** 对象，该对象代表添加到打开文档集合中的新建空文档。

**语法**

**express.Add**

*express*   一个代表 **Documents** 对象的变量。

**示例**

| 示例代码复制                                               |
| ---------------------------------------------------------- |
| `/*本示例基于 Normal 模板新建一篇文档。*/ Documents.Add()` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例基于“专业型备忘录”模板新建一篇文档。*/ Documents.Add("C:\\Program Files\\Kingsoft\\WPS Office Professional\\office6\\mui\\default\\templates\\Normal.wpt")` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例以活动文档的模板为样例，创建并打开一个新模板。*/ function test() { let tmpName = ActiveDocument.AttachedTemplate.FullName Documents.Add(tmpName, true) }  ` |

#### **Documents.Application**

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

*express*   一个代表 **Documents** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **Documents.Count**

返回一个 **Number** 类型的值，该值代表集合中的文档数。只读。

**语法**

**express.Count**

*express*   一个代表 **Documents** 对象的变量。

#### **Documents.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Documents** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

适用环境：web

适用平台：windows/linux