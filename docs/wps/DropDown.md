**DropDown** 



代表包含一个窗体选项列表的下拉型窗体域。

**说明**

使用 **FormFields**(index) 返回单个 **FormField** 对象（其中 index 是与下拉型窗体域相关的索引编号或书签名称）。可将 **DropDown** 属性与 **FormField** 对象配合使用，返回 **DropDown** 的对象。以下示例选择活动文档中名为“DropDown”的下拉型窗体域中的第一项。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `ActiveDocument.FormFields.Item("DropDown1").DropDown.Value = 1` |

索引编号代表窗体域在 **FormFields** 集合中的位置。以下示例检查活动文档中第一个窗体域的类型。如果是下拉型窗体域，则选中第二项。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() { if(ActiveDocument.FormFields.Item(1).Type == wdFieldFormDropDown) {     ActiveDocument.FormFields.Item(1).DropDown.Value = 2 } }  ` |

以下示例在向 *ffield* 窗体域中添加选项以前，判断该下拉型窗体域是否有效。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() { let ffield = ActiveDocument.FormFields.Item(1).DropDown if(ffield.Valid == true) {     ffield.ListEntries.Add("Hello") } else {     MsgBox("First field is not a drop down") } }  ` |

可将 **Add** 方法用于 **FormFields** 集合来添加下拉型窗体域。以下示例在活动文档的开头添加一个下拉型窗体域，然后在该窗体域中添加项目。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() { let ffield = ActiveDocument.FormFields.Add( ActiveDocument.Range(0, 0), wdFieldFormDropDown) let ff = ActiveDocument.FormFields.Item(1) ff.Name = "Colors" ff.DropDown.ListEntries ff.DropDown.ListEntries.Add("Blue") ff.DropDown.ListEntries.Add("Green") ff.DropDown.ListEntries.Add("Red") }` |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个代表 WPS 应用程序的 **Application** 对象。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Default**     | 返回或设置一个 **Long** 类型值，该值代表默认的下拉型项目。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ListEntries** | 返回一个 **ListEntries** 集合，该集合代表 **DropDown** 对象中所有项。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回一个 **Object** 类型值，该值代表指定 **DropDown** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Valid**       | 如果指定的窗体域对象为有效的下拉窗体域，则该属性为 **True**。**Boolean** 类型，只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Value**       | 返回或设置下拉型窗体域中所选项目的编号。可读写 **Long** 类型。 |

**成员属性**

#### **DropDown.Application**

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

*express*   一个代表 **DropDown** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **DropDown.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **DropDown** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **DropDown.Default**

返回或设置一个 **Long** 类型值，该值代表默认的下拉型项目。可读写。

**语法**

**express.Default**

*express*   一个代表 **DropDown** 对象的变量。

**说明**

下拉型窗体域的第一项为 1，第二项为 2，以此类推。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例为 Sales.doc 中名为“Colors”的下拉型窗体域设置默认项。*/ Documents.Item("Sales.doc").FormFields.Item("Colors").DropDown.Default = 2` |

#### **DropDown.ListEntries**

返回一个 **ListEntries** 集合，该集合代表 **DropDown** 对象中所有项。

**语法**

**express.ListEntries**

*express*   一个代表 **DropDown** 对象的变量。

**说明**

有关返回集合中单个成员的信息，请参阅 返回集合中的对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在名为“DropDown1”的下拉型窗体域中检索活动项文字。*/ function test() { let myField = ActiveDocument.FormFields.Item("DropDown1").DropDown let num = myField.Value let myName = myField.ListEntries.Item(Number(num)).Name }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在活动下拉型窗体域中检索项的总数（应该先保护文档的窗体）。如果有两个或更多的项，本示例将第二项设置为活动项。*/ function test() { let myField = Selection.FormFields.Item(1) if(myField.Type == wdFieldFormDropDown) {     num = myField.DropDown.ListEntries.Count     if(num > = 2) {         myField.DropDown.Value = 2 } }  ` |

#### **DropDown.Parent**

返回一个 **Object** 类型值，该值代表指定 **DropDown** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **DropDown** 对象的变量。

#### **DropDown.Valid**

如果指定的窗体域对象为有效的下拉窗体域，则该属性为 **True**。**Boolean** 类型，只读。

**语法**

**express.Valid**

*express*   一个代表 **DropDown** 对象的变量。

**说明**

在应用 **DropDown** 属性之前，请使用 **FormField** 对象的 **Type** 属性来确定窗体域的类型 (**wdFieldFormDropDown**)。该措施可确保 **FormField** 对象为要求的类型。

#### **DropDown.Value**

返回或设置下拉型窗体域中所选项目的编号。可读写 **Long** 类型。

**语法**

**express.Value**

*express*   一个代表 **DropDown** 对象的变量。

适用环境：web

适用平台：windows/linux