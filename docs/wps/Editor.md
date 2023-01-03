**Editor** 



表示已被分配特定权限可编辑文档部分的单个用户。

**说明**

可授予权限的用户包括单独的参与者以及为“文档工作区”站点定义的用户组。

分配给范围和所选内容的权限仅在文档受到保护之后生效。使用 **Editors** 集合和 **Editor** 对象可将特定权限分配到文档的节。然后使用 **Protect** 方法来保护该文档。

使用 **Add** 集合的 **Editors** 方法可向指定的用户或组授予权限以修改文档中的范围或所选内容。以下示例为当前用户分配编辑权限以修改活动的所选内容。

| 示例代码复制                                             |
| -------------------------------------------------------- |
| `let objEditor = Selection.Editors.Add(wdEditorCurrent)` |

**方法**

|                                                              | 名称          | 说明                                       |
| ------------------------------------------------------------ | ------------- | ------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**    | 删除指定的 **Editor** 对象。               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **DeleteAll** | 删除特定用户在文档中的所有编辑权限。       |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **SelectAll** | 选择文档中由单个用户插入或编辑的所有形状。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ID**          | 将父文档另存为网页时，返回或设置指定对象的标识标签。只读 **String** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**        | 返回或设置指定对象的名称。只读 **String** 类型。             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **NextRange**   | 返回一个 **Range** 对象，该对象代表用户有权修改的下一个区域。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回一个 **Object** 类型值，该值代表指定 **Editor** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Range**       | 返回一个 **Range** 对象，该对象代表指定对象所包含的文档部分。 |

**成员方法**

#### **Editor.Delete**

删除指定的 **Editor** 对象。

**语法**

**express.Delete()**

*express*   一个代表 **Editor** 对象的变量。

#### **Editor.DeleteAll**

删除特定用户在文档中的所有编辑权限。

**语法**

**express.DeleteAll()**

*express*   一个代表 **Editor** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例删除第一个用户在活动文档 Editors 集合中的所有编辑权限。*/ function test() { let objEditor = Selection.Editors.Item(1) objEditor.DeleteAll() }` |

#### **Editor.SelectAll**

选择文档中由单个用户插入或编辑的所有形状。

**语法**

**express.SelectAll()**

*express*   一个代表 **Editor** 对象的变量。

**说明**

此方法不能选择 **InlineShape** 对象。

**成员属性**

#### **Editor.Application**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

**语法**

**express.Application**

*express*   一个代表 **Editor** 对象的变量。

#### **Editor.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Editor** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **Editor.ID**

将父文档另存为网页时，返回或设置指定对象的标识标签。只读 **String** 类型。

**语法**

**express.ID**

*express*   一个代表 **Editor** 对象的变量。

#### **Editor.Name**

返回或设置指定对象的名称。只读 **String** 类型。

**语法**

**express.Name**

*express*   一个代表 **Editor** 对象的变量。

#### **Editor.NextRange**

返回一个 **Range** 对象，该对象代表用户有权修改的下一个区域。

**语法**

**express.NextRange**

*express*   一个代表 **Editor** 对象的变量。

**说明**

还可使用 **Range** 对象的 **GoToEditableRange** 方法返回用户有权修改的下一区域。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例返回活动选定内容中第一个编辑者的下一区域。*/ function test() { let objEditor = Selection.Editors.Item(1) let objRange = objEditor.NextRange }` |

#### **Editor.Parent**

返回一个 **Object** 类型值，该值代表指定 **Editor** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **Editor** 对象的变量。

#### **Editor.Range**

返回一个 **Range** 对象，该对象代表指定对象所包含的文档部分。

**语法**

**express.Range**

*express*   一个代表 **Editor** 对象的变量。

**说明**

有关从文档返回范围或从形状集合返回形状范围的信息，请参阅 **Range** 方法。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例为活动文档中的第一段应用“标题 1”样式。*/ ActiveDocument.Paragraphs.Item(1).Range.Style = wdStyleHeading1` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例复制表格 1 中的第一行。*/ function test() { if(ActiveDocument.Tables.Count >= 1) {     ActiveDocument.Tables.Item(1).Rows.Item(1).Range.Copy() } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例更改文档中第一个批注的文本。*/ function test() { let comments = ActiveDocument.Comments.Item(1).Range comments.Delete() comments.InsertBefore("new comment text") }  ` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在第一节的结尾插入文本。*/ function test() { let myRange = ActiveDocument.Sections.Item(1).Range myRange.MoveEnd(wdCharacter, -1) myRange.Collapse(wdCollapseEnd) myRange.InsertParagraphAfter() myRange.InsertAfter("End of section") }` |

适用环境：web

适用平台：windows/linux