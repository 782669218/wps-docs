**CheckBox** 



代表一个复选框窗体域。

**说明**

使用 **FormFields**(*Index*) 返回单个 **FormField** 对象（其中 *Index* 是与复选框关联的索引编号或书签名称）。可将 **CheckBox** 属性与 **FormField** 对象配合使用，返回 **CheckBox** 对象。以下示例从活动文档中选择名为“Check1”的复选框窗体域。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.ActiveDocument.FormFields.Item("Check1").CheckBox.Value = true` |

索引编号代表窗体域在 **FormFields** 集合中的位置。以下示例检查第一个窗体域的类型，如果是复选框，则选中该复选框。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     if (Application.ActiveDocument.FormFields.Item(1).Type == wdFieldFormCheckBox) {         Application.ActiveDocument.FormFields.Item(1).CheckBox.Value = true     } }` |

以下示例先判断 *ffield* 对象是否有效，然后再将复选框大小修改为 14 磅。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let ffield = Application.ActiveDocument.FormField.Item(1).CheckBox     if (ffield.Valid == true) {         ffield.AutoSize = false         ffield.Size = 14     }     else {         alert("First field is not a check box")     } }` |

可将 **Add** 方法与 **FormFields** 对象配合使用，添加一个复选框窗体域。以下示例在活动文档的开头添加一个复选框，并将其命名为“Color”，然后选定该复选框。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     let rng = Application.ActiveDocument.FormFields.Add(ActiveDocument.Range(0, 0), wdFieldFormCheckBox)     rng.Name = "Color"     rng.CheckBox.Value = true }` |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个 **Application** 对象，该对象代表 WPS 应用程序。     |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **AutoSize**    | 如果为 **True**，则根据环绕文字的字体大小调整复选框或文本框的大小。如果为 **False**，则根据 **Size** 属性调整复选框或文本框的大小。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数指出用于创建指定对象的应用程序。**Long** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Default**     | 返回或设置默认的复选框值。如果选中默认值，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回一个 **Object** 类型的值，该值代表指定 **CheckBox** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Size**        | 返回或设置复选框的大小。**Single** 类型，可读写。            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Valid**       | 如果指定的窗体域对象为有效的复选框窗体域，则该属性值为 **True**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Value**       | 如果选中此复选框，则该属性值为 **True**。**Boolean** 类型，可读写。 |

**成员属性**

#### **CheckBox.Application**

返回一个 **Application** 对象，该对象代表 WPS 应用程序。

**语法**

**express.Application**

*express*   一个代表 **CheckBox** 对象的变量。

#### **CheckBox.AutoSize**

如果为 **True**，则根据环绕文字的字体大小调整复选框或文本框的大小。如果为 **False**，则根据 **Size** 属性调整复选框或文本框的大小。**Boolean** 类型，可读写。

**语法**

**express.AutoSize**

*express*   一个代表 **CheckBox** 对象的变量。

#### **CheckBox.Creator**

返回一个 32 位整数，该整数指出用于创建指定对象的应用程序。**Long** 类型，只读。

**语法**

**express.Creator**

*express*   一个代表 **CheckBox** 对象的变量。

**说明**

如果对象是在 WPS 中创建的，则 **Creator** 属性返回十六进制数 4D535744，代表字符串“WPS”。该属性主要是为 Macintosh 机的应用设计的，在该机上每个应用程序都有一个四字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参考 WPS OfficeMacintosh Edition 附带的语言参考帮助。

| 注释                                    |
| --------------------------------------- |
| 该值也可用常量 **wdCreatorCode** 表示。 |

#### **CheckBox.Default**

返回或设置默认的复选框值。如果选中默认值，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.Default**

*express*   一个代表 **CheckBox** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果活动文档的第一个窗体域是复选框类型，本示例将检索其默认值。*/ function test() {     let blnDefault     if (Application.ActiveDocument.FormFields.Item(1).Type == wdFieldFormCheckBox) {         blnDefault = Application.ActiveDocument.FormFields.Item(1).CheckBox.Default     } }` |

#### **CheckBox.Parent**

返回一个 **Object** 类型的值，该值代表指定 **CheckBox** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **CheckBox** 对象的变量。

#### **CheckBox.Size**

返回或设置复选框的大小。**Single** 类型，可读写。

**语法**

**express.Size**

*express*   一个代表 **CheckBox** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档中名为“Check1”的复选框的大小设置为 14 磅，然后选中该复选框。*/ function test() {     let rng = Application.ActiveDocument.FormFields.Item("Check1").CheckBox     rng.AutoSize = false     rng.Size = 14     rng.Value = true }` |

#### **CheckBox.Valid**

如果指定的窗体域对象为有效的复选框窗体域，则该属性值为 **True**。**Boolean** 类型，只读。

**语法**

**express.Valid**

*express*   一个代表 **CheckBox** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在插入点添加一个文字型窗体域。由于 myFormField 是一个文字输入域而不是一个复选框，所以消息框将显示“False”。*/ function test() {     Application.Selection.Collapse(wdCollapseStart)     let myFormField = Application.ActiveDocument.FormFields.Add(Application.Selection.Range, wdFieldFormTextInput)     alert(myFormField.CheckBox.Valid) }` |

#### **CheckBox.Value**

如果选中此复选框，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.Value**

*express*   一个代表 **CheckBox** 对象的变量。

适用环境：web

适用平台：windows/linux