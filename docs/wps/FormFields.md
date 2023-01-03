#### **FormFields**



**FormField** 对象的集合，这些对象代表所选内容、范围或文档中的所有窗体域。

**方法**

|                                                              | 名称     | 说明                                                         |
| ------------------------------------------------------------ | -------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**  | 返回一个 **FormField** 对象，该对象代表添加到区域中的新的窗体域。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item** | 返回集合中的单个 **FormField** 对象。                        |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个代表 WPS 应用程序的 **Application** 对象。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回一个 **Long** 类型的值，该值代表集合中域的数量。只读。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回一个 **Object** 类型值，该值代表指定 **FormFields** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Shaded**      | 如果将底纹应用于窗体域，则该属性值为 **True**。**Boolean** 类型，可读写。 |

**成员方法**

#### **FormFields.Add**

返回一个 **FormField** 对象，该对象代表添加到区域中的新的窗体域。

**语法**

**express.Add(Range, Type)**

*express*   一个代表 **FormFields** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型**    | **说明**                                                     |
| -------- | ------------- | --------------- | ------------------------------------------------------------ |
| *Range*  | 必选          | **Range 对象**  | 要添加窗体域的区域。如果该区域未折叠，那么窗体域将替换此区域。 |
| *Type*   | 必选          | **WdFieldType** | 要添加的窗体域的类型。                                       |

**说明**

**安全性:**动态数据交换 (DDE) 是一种不安全的陈旧技术。如果可能，请使用比 DDE 更加安全的技术，如对象链接和嵌入 (OLE)。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例在选定内容的末尾处添加一个复选框，然后命名并选定该复选框。*/ function test() {   Application.Selection.Collapse(wdCollapseEnd)      let ffield = Application.ActiveDocument.FormFields.Add(Selection.Range, wdFieldFormCheckBox)   ffield.Name = "Check_Box_1"   ffield.CheckBox.Value = true }` |

#### **FormFields.Item**

返回集合中的单个 **FormField** 对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **FormFields** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Index*  | 必选          | **Variant**  | 要返回的单个对象。可以是代表序号位置的 Long 类型值，或代表单个对象名称的 String 类型值。 |

**成员属性**

#### **FormFields.Application**

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

*express*   一个代表 **FormFields** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **FormFields.Count**

返回一个 **Long** 类型的值，该值代表集合中域的数量。只读。

**语法**

**express.Count**

*express*   一个代表 **FormFields** 对象的变量。

#### **FormFields.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **FormFields** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **FormFields.Parent**

返回一个 **Object** 类型值，该值代表指定 **FormFields** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **FormFields** 对象的变量。

#### **FormFields.Shaded**

如果将底纹应用于窗体域，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.Shaded**

*express*   一个代表 **FormFields** 对象的变量。

**说明**

使用底纹便于在文档中查找窗体域，且不影响打印输出。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例从 Employment Form.doc 的窗体域中清除底纹。*/ Application.Documents.Item("Employment Form.doc").FormFields.Shaded = false  /*本示例向活动文档的窗体域添加底纹，并且保护文档中的窗体。*/ function test() {   Application.ActiveDocument.FormFields.Shaded = true   Application.ActiveDocument.Protect(wdAllowOnlyFormFields, true) }` |

适用环境：web

适用平台：windows/linux