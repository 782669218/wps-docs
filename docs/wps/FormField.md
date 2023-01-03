#### **FormField**



代表单个窗体域。**FormField** 对象是 **FormFields** 集合的一个成员。

**说明**

使用 **FormFields**(index) 可返回一个 **FormField** 对象，其中 index 是书签名称或索引号。以下示例将 Text1 窗体域的结果设置为“Don Funk”。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActiveDocument.FormFields.Item("Text1").Result = "Don Funk"` |

索引号代表窗体域在所选内容、范围或文档中的位置。以下示例显示所选内容中第一个窗体域的名称。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     if(Application.Selection.FormFields.Count >= 1) {         alert(Application.Selection.FormFields.Item(1).Name)     } }` |

使用 

Add

 方法和 

FormFields

 对象可添加一个窗体域。以下示例在活动文档的开始处添加一个复选框，然后选中该复选框。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let ffield = Application.ActiveDocument.FormFields.Add(ActiveDocument.Range(0,0),wdFieldFormCheckBox)     ffield.CheckBox.Value = true }` |

使用 

FormField

 对象和 

CheckBox

、

DropDown

 和 

TextInput

 属性可返回 

CheckDown

、

DropDown

 和 

TextInput

 对象。以下示例选中名为“Check1”的复选框。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActiveDocument.FormFields.Item("Check1").CheckBox.Value = true` |

**方法**

|                                                              | 名称       | 说明                               |
| ------------------------------------------------------------ | ---------- | ---------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Copy**   | 将指定窗体域复制到剪贴板。         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Cut**    | 将指定窗体域从文档中移到剪贴板上。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete** | 删除指定窗体域。                   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Select** | 选择指定的对象。                   |

**属性**

|                                                              | 名称                | 说明                                                         |
| ------------------------------------------------------------ | ------------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application**     | 返回一个代表 WPS 应用程序的 **Application** 对象。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CalculateOnExit** | 如果退出指定的窗体域时自动更新对该域的引用，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **CheckBox**        | 返回一个 **CheckBox** 对象，该对象代表复选框型窗体域。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**         | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **DropDown**        | 返回一个 **DropDown** 对象，该对象代表下拉型窗体域。只读。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Enabled**         | 如果启用窗体域，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **EntryMacro**      | 返回或设置指定窗体域（CheckBox、DropDown 或 TextInput）的输入宏的名称。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ExitMacro**       | 返回或设置指定窗体域（CheckBox、DropDown 或 TextInput）的退出宏的名称。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **HelpText**        | 返回或设置当窗体域获得焦点并且用户按 F1 时在消息框中显示的文字。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Name**            | 返回或设置指定对象的名称。**String** 类型，可读写。          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Next**            | 返回集合中的下一个窗体域。只读。                             |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OwnHelp**         | 指定当窗体域具有焦点并且用户按 F1 时在消息框中显示的文本的源。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Range**           | 返回一个 **Range** 对象，该对象代表窗体域中包含的文档部分。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Result**          | 返回一个代表指定的窗体域的结果的 **String**。可读写。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**            | 返回域的类型。只读 [WdFieldType ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdFieldType%20%E6%9E%9A%E4%B8%BE.html)类型。 |

**成员方法**

#### **FormField.Copy**

将指定窗体域复制到剪贴板。

**语法**

**express.Copy()**

*express*   一个代表 **FormField** 对象的变量。

#### **FormField.Cut**

将指定窗体域从文档中移到剪贴板上。

**语法**

**express.Cut()**

*express*   一个代表 **FormField** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例删除活动文档的第一个域，并将该域粘贴到插入点。*/ function test() {   if(Application.ActiveDocument.Fields.Count >= 1) {       Application.ActiveDocument.Fields.Item(1).Cut()       Application.Selection.Collapse(wdCollapseEnd)       Application.Selection.Paste()   } }  /*本示例删除第一段的第一个单词，并将该单词粘贴到该段的末尾。*/ function test() {   let ran= Application.ActiveDocument.Paragraphs.Item(1).Range   ran.Words.Item(1).Cut()   ran.Collapse(wdCollapseEnd)   ran.Move(wdCharacter,-1)   ran.Paste() }  /*本示例删除所选内容，并将其粘贴到新文档中。*/ function test() {   if(Application.Selection.Type == wdSelectionNormal) {       Application.Selection.Cut()       Application.Documents.Add().Content.Paste()   } }` |

#### **FormField.Delete**

删除指定窗体域。

**语法**

**express.Delete()**

*express*   一个代表 **FormField** 对象的变量。

#### **FormField.Select**

选择指定的对象。

**语法**

**express.Select()**

*express*   一个代表 **FormField** 对象的变量。

**说明**

使用此方法后，可使用 [Selection ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Selection/Selection%20.htm#jsObject_Selection)对象处理所选项目。有关详细信息，请参阅 处理 Selection 对象。

**成员属性**

#### **FormField.Application**

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

*express*   一个代表 **FormField** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **FormField.CalculateOnExit**

如果退出指定的窗体域时自动更新对该域的引用，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.CalculateOnExit**

*express*   一个代表 **FormField** 对象的变量。

**说明**

一个 REF 域可用来引用窗体域中的内容。例如，{REF SubTotal} 引用由 SubTotal 书签标识的窗体域。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在退出 Form.doc 中的窗体域时可以自动更新对该域的引用。*/ function test() { for(let i=1;i<=Documents.Item("Form.doc").FormFields.Count;i++) {     Documents.Item("Form.doc").FormFields.Item(i).CalculateOnExit = false } } ` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在新文档中添加一个文字型窗体域和一个 REF 域。每次键入文字时并退出 Text1 域时，都将自动更新 REF 域。*/ function test() { let abc= Documents.Add() abc.FormFields.Add(Selection.Range,wdFieldFormTextInput) abc.Fields.Add(Selection.Range,wdFieldRef,"Text1") abc.FormFields.Item("Text1").CalculateOnExit = true abc.Protect(wdAllowOnlyFormFields) }` |

#### **FormField.CheckBox**

返回一个 **CheckBox** 对象，该对象代表复选框型窗体域。只读。

**语法**

**express.CheckBox**

*express*   一个代表 **FormField** 对象的变量。

**说明**

如果将 **CheckBox** 属性应用于非复选框型窗体域的 **FormField** 对象，则该属性有效，但是返回的对象的 **Valid** 属性将为 **False**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例清除名为“Blue”的复选框。*/ ActiveDocument.FormFields.Item("Blue").CheckBox.Value = false` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将当前值与名为“Check1”的复选框的默认值进行比较。如果两值相等，则 blnSame 变量会设置为 True。*/ function test() { let ffTemp = ActiveDocument.FormFields.Item("Check1").CheckBox let blnSame if(ffTemp.Default == ffTemp.Value) {     blnSame = true } else {     blnSame = false } }` |

#### **FormField.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **FormField** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **FormField.DropDown**

返回一个 **DropDown** 对象，该对象代表下拉型窗体域。只读。

**语法**

**express.DropDown**

*express*   一个代表 **FormField** 对象的变量。

**说明**

如果将 **DropDown** 属性应用于非下拉型窗体域的 **FormField** 对象，则该属性有效，但是返回的对象的 **Valid** 属性将为 **False**。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示名为“Colors”的下拉型窗体域中所选项的文本。*/ function test() { let ffDrop = ActiveDocument.FormFields.Item("Colors").DropDown MsgBox(ffDrop.ListEntries.Item(ffDrop.Value).Name) }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例向 Form.doc 中名为“Places”的下拉型窗体域中添加“Seattle”。*/ function test() { let trie=Documents.Item("Form.doc").FormFields.Item("Places").DropDown.ListEntries trie.Add("Seattle") }` |

#### **FormField.Enabled**

如果启用窗体域，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.Enabled**

*express*   一个代表 **FormField** 对象的变量。

**说明**

如果启用了一个窗体域，则在填充该窗体时，可以更改其内容。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果活动文档中的第一个窗体域是一个已启用的复选框，则以下示例选中该复选框。*/ function test() { let ffFirst = ActiveDocument.FormFields.Item(1) if(ffFirst.Enabled == true &&ffFirst.Type == wdFieldFormCheckBox) {     ffFirst.CheckBox.Value = true } }` |

#### **FormField.EntryMacro**

返回或设置指定窗体域（CheckBox、DropDown 或 TextInput）的输入宏的名称。**String** 类型，可读写。

**语法**

**express.EntryMacro**

*express*   一个代表 **FormField** 对象的变量。

**说明**

当窗体域获得焦点时，输入宏开始运行。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为“Form.doc”中的第一个窗体域指定名为“Blue”的宏。*/ Documents.Item("Form.doc").FormFields.Item(1).EntryMacro = "Blue"` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为活动文档中名为“Text1”的窗体域指定一个名为“Breadth”的宏。*/ ActiveDocument.FormFields.Item("Text1").EntryMacro = "Breadth"` |

#### **FormField.ExitMacro**

返回或设置指定窗体域（CheckBox、DropDown 或 TextInput）的退出宏的名称。**String** 类型，可读写。

**语法**

**express.ExitMacro**

*express*   一个代表 **FormField** 对象的变量。

**说明**

当窗体域失去焦点时，退出宏开始运行。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为选定内容的第一个窗体域指定名为“Reformat”的宏。*/ function test() { if(Selection.FormFields.Count > 0 ) {     Selection.FormFields.Item(1).ExitMacro = "Reformat" } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为“Form.doc”中最后一个窗体域指定名为“Blue”的宏。*/ function test() { let intMax = Documents.Item("Form.doc").FormFields.Count Documents.Item("Form.doc").FormFields.Item(intMax).ExitMacro = "Blue" }` |

#### **FormField.HelpText**

返回或设置当窗体域获得焦点并且用户按 F1 时在消息框中显示的文字。**String** 类型，可读写。

**语法**

**express.HelpText**

*express*   一个代表 **FormField** 对象的变量。

**说明**

如果将 **OwnHelp** 属性设置为 **True**，则 **HelpText** 会指定该文字字符串的值。如果将 **OwnHelp** 设置为 **False**，则 **HelpText** 会指定包含窗体域的帮助文字的“自动图文集”词条的名称。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为名为“Name”的窗体域设置帮助文字。*/ function test() { ActiveDocument.FormFields.Item("Name").OwnHelp = true ActiveDocument.FormFields.Item("Name").HelpText = "Type your full legal name." }` |

#### **FormField.Name**

返回或设置指定对象的名称。**String** 类型，可读写。

**语法**

**express.Name**

*express*   一个代表 **FormField** 对象的变量。

#### **FormField.Next**

返回集合中的下一个窗体域。只读。

**语法**

**express.Next**

*express*   一个代表 **FormField** 对象的变量。

#### **FormField.OwnHelp**

指定当窗体域具有焦点并且用户按 F1 时在消息框中显示的文本的源。**Boolean** 类型，可读写。

**语法**

**express.OwnHelp**

*express*   一个代表 **FormField** 对象的变量。

**说明**

如果为 **True**，则显示 **HelpText** 属性指定的文本。如果为 **False**，则显示 **HelpText** 属性指定的“自动图文集”词条中的文本。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将当前节的第一个窗体域的帮助文字设置为名为“Sample”的“自动图文集”词条的内容。*/ function test() { let help=Selection.Sections.Item(1).Range.FormFields.Item(1) help.OwnHelp = false help.HelpText = "Sample" }` |

#### **FormField.Range**

返回一个 **Range** 对象，该对象代表窗体域中包含的文档部分。

**语法**

**express.Range**

*express*   一个代表 **FormField** 对象的变量。

#### **FormField.Result**

返回一个代表指定的窗体域的结果的 **String**。可读写。

**语法**

**express.Result**

*express*   一个代表 **FormField** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示活动文档中每个窗体域的结果。*/ function test() {   for(let i=1;i<=Application.ActiveDocument.FormFields.Count;i++) {       alert(Application.ActiveDocument.FormFields.Item(i).Result)   } }` |

#### **FormField.Type**

返回域的类型。只读 [WdFieldType ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdFieldType%20%E6%9E%9A%E4%B8%BE.html)类型。

**语法**

**express.Type**

*express*   一个代表 **FormField** 对象的变量。

**说明**

**安全性:**动态数据交换 (DDE) 是一种不安全的陈旧技术。如果可能，请使用比 DDE 更加安全的技术，如对象链接和嵌入 (OLE)。

适用环境：web

适用平台：windows/linux