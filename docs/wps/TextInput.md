#### **TextInput**



代表一个文本型窗体域。

**说明**

使用 **FormFields**(*Index*) 可返回一个 **FormField** 对象，其中 *Index* 是与文本型窗体域或索引号相关联的书签名称。使用 **TextInput** 属性和 **FormField** 对象可返回一个 **TextInput** 对象。以下示例删除活动文档中名为“Text1”的文本型窗体域的内容。

| 示例代码复制                                                |
| ----------------------------------------------------------- |
| `ActiveDocument.FormFields.Item("Text1").TextInput.Clear()` |

索引号代表窗体域在 **FormFields** 集合中的位置。以下示例检查活动文档中第一个窗体域的类型。如果该窗体域是文本型窗体域，则该示例将“Mission Critical”设置为该域的值。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() { if(ActiveDocument.FormFields.Item(1).Type == wdFieldFormTextInput){     ActiveDocument.FormFields.Item(1).Result = "Mission Critical" } }` |

以下示例确定 *ffield* 变量在设置为默认的文本以前，是否在活动文档中代表一个有效的文本型窗体域。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() { let ffield = ActiveDocument.FormFields.Item(1).TextInput  if(ffield.Valid == true){     ffield.Default = "Type your name here" } else{     MsgBox("First field is not a text box") } }` |

使用 **Add** 方法和 **FormFields** 对象可添加一个文本型窗体域。以下示例在活动文档的开始处添加文本型窗体域，然后将该窗体域的名称设置为“FirstName”。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() { let ffield = ActiveDocument.FormFields.Add(ActiveDocument.Range(0, 0), wdFieldFormTextInput) ffield.Name = "FirstName" }` |

**方法**

|                                                              | 名称         | 说明                             |
| ------------------------------------------------------------ | ------------ | -------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Clear**    | 删除指定的文字型窗体域中的文字。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **EditType** | 为指定的文字型窗体域设置选项。   |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个代表 WPS 应用程序的 **Application** 对象。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Default**     | 返回或设置代表默认文本框内容的文本。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Format**      | 返回指定文本框的文字格式。**String** 类型，只读。            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回一个 **Object** 类型值，该值代表指定 **TextInput** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**        | 返回文字型窗体域的类型。只读 **WdTextFormFieldType** 类型。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Valid**       | 如果指定的窗体域对象为有效的复选框窗体域，则该属性值为 **True**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Width**       | 返回或设置指定文本输入域的宽度（以磅为单位）。可读写 **Long** 类型。 |

**成员方法**

#### **TextInput.Clear**

删除指定的文字型窗体域中的文字。

**语法**

**express.Clear()**

*express*   一个代表 **TextInput** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*此示例保护文档中的窗体，并当第一个窗体域是文字型窗体域时，删除其中的文本。*/ function test() { ActiveDocument.Protect(wdAllowOnlyFormFields, true)  if(ActiveDocument.FormFields.Item(1).Type == wdFieldFormTextInput){     ActiveDocument.FormFields.Item(1).TextInput.Clear() }  }` |

#### **TextInput.EditType**

为指定的文字型窗体域设置选项。

**语法**

**express.EditType(Type, Default, Format, Enabled)**

*express*   一个代表 **TextInput** 对象的变量。

**参数**

| **名称**  | **必选/可选** | **数据类型**            | **说明**                                                     |
| --------- | ------------- | ----------------------- | ------------------------------------------------------------ |
| *Type*    | 必选          | **WdTextFormFieldType** | 文本框的类型。                                               |
| *Default* | 可选          | **Variant**             | 出现在文本框中的默认文字。                                   |
| *Format*  | 可选          | **Variant**             | 设置文本、数字或日期格式的格式字符串（例如，“0.00”、“Title Case”或“M/d/yy”）。有关格式的更多示例，请查看“文字型窗体域选项”对话框中指定文字型窗体域类型的格式列表。 |
| *Enabled* | 可选          | **Variant**             | 如果为 True，则可以向窗体域中输入文字。                      |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在活动文档的开头添加一个名为“Today”的文字型窗体域。EditType 方法将窗体域的类型设置为 wdCurrentDateText，并将日期格式设置为“M/d/yy”。*/ function test() { let DateText = ActiveDocument.FormFields.Add(ActiveDocument.Range(0, 0), wdFieldFormTextInput)  DateText.Name = "Today" DateText.TextInput.EditType(wdCurrentDateText, null, "M/d/yy", false) }` |

**成员属性**

#### **TextInput.Application**

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

*express*   一个代表 **TextInput** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **TextInput.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **TextInput** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **TextInput.Default**

返回或设置代表默认文本框内容的文本。**String** 类型，可读写。

**语法**

**express.Default**

*express*   一个代表 **TextInput** 对象的变量。

#### **TextInput.Format**

返回指定文本框的文字格式。**String** 类型，只读。

**语法**

**express.Format**

*express*   一个代表 **TextInput** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示活动文档中第一个域中的文本格式。*/ function test() { if(ActiveDocument.FormFields.Item(1).Type == wdFieldFormTextInput){     MsgBox(ActiveDocument.FormFields.Item(1).TextInput.Format) } else{     MsgBox("First field is not a text form field") } }` |

#### **TextInput.Parent**

返回一个 **Object** 类型值，该值代表指定 **TextInput** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **TextInput** 对象的变量。

#### **TextInput.Type**

返回文字型窗体域的类型。只读 **WdTextFormFieldType** 类型。

**语法**

**express.Type**

*express*   一个代表 **TextInput** 对象的变量。

#### **TextInput.Valid**

如果指定的窗体域对象为有效的复选框窗体域，则该属性值为 **True**。**Boolean** 类型，只读。

**语法**

**express.Valid**

*express*   一个代表 **TextInput** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例确定活动文档中第一个窗体域是否为文字型窗体域。如果 Valid 属性值为 True，则该文字型窗体域的内容更改为“Hello”。 */ function test() { if(ActiveDocument.FormFields.Item(1).TextInput.Valid == true){     ActiveDocument.FormFields.Item(1).Result = "Hello" } }` |

#### **TextInput.Width**

返回或设置指定文本输入域的宽度（以磅为单位）。可读写 **Long** 类型。

**语法**

**express.Width**

*express*   一个代表 **TextInput** 对象的变量。

适用环境：web

适用平台：windows/linux