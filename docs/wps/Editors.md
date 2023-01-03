#### **Editors**



**Editor** 对象的集合代表已被授予特定权限可编辑文档部分的用户或用户组集合。

**方法**

|                                                              | 名称     | 说明                                                         |
| ------------------------------------------------------------ | -------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**  | 返回一个 **Editor** 对象，该对象代表指定用户用于修改文档中区域或选定内容的新权限。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item** | 返回一个 **Editor** 对象，该对象代表已分配给权限来编辑部分文档的特定用户或用户组。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个代表 WPS 应用程序的 **Application** 对象。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Code**        | 返回一个代表域代码的 **Range** 对象。可读写。                |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回一个 **Long** 类型值，该值代表集合中 Editor 对象的数目。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Data**        | 返回或设置 ADDIN 域中的数据。可读写 **String** 类型。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Index**       | 返回一个 **Long** 类型的值，该值代表项目在集合中的位置。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **InlineShape** | 返回一个 **InlineShape** 对象，该对象可以代表图片、OLE 对象以及作为 INCLUDEPICTURE 或 EMBED 域结果的 ActiveX 控件。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Kind**        | 返回 **Field** 对象的链接类型。只读 **WdFieldKind** 类型。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LinkFormat**  | 返回一个 **LinkFormat** 对象，该对象代表指定域的链接选项。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Locked**      | 如果指定域已被锁定，则该属性值为 **True**。可读写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Next**        | 返回集合中的下一个对象。只读。                               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OLEFormat**   | 返回一个 **OLEFormat** 对象，该对象代表指定域的 OLE 特征（链接除外）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回一个 **Object** 类型值，该值代表指定 **Editors** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Previous**    | 返回集合中的上一个对象。只读。                               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Result**      | 返回一个 **Range** 对象，该对象代表域的结果。可读写。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowCodes**   | 如果显示指定域的域代码而不是域结果，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**        | 返回域的类型。只读 **WdFieldType** 类型。                    |

**成员方法**

#### **Editors.Add**

返回一个 **Editor** 对象，该对象代表指定用户用于修改文档中区域或选定内容的新权限。

**语法**

**express.Add(EditorID)**

*express*   一个代表 **Editors** 对象的变量。

**参数**

| **名称**   | **必选/可选** | **数据类型** | **说明**                                                     |
| ---------- | ------------- | ------------ | ------------------------------------------------------------ |
| *EditorID* | 必选          | **Variant**  | 可为代表用户电子邮件别名（如果在相同的域中）和电子邮件地址的 String 字符串或代表用户组的 WdEditorType 常量。 |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将选定文本的编辑权限分配给当前用户。*/ let objEditor = Selection.Editors.Add(wdEditorCurrent)` |

#### **Editors.Item**

返回一个 **Editor** 对象，该对象代表已分配给权限来编辑部分文档的特定用户或用户组。

**语法**

**express.Item(Index)**

*express*   一个代表 **Editors** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Index*  | 必选          | **Variant**  | 可以是代表用户电子邮件别名（如果在相同的域中）和电子邮件地址的 String 类型值，或者是代表用户组的 WdEditorType 常量。 |

**返回值**

Editor

**成员属性**

#### **Editors.Application**

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

*express*   一个代表 **Editors** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **Editors.Code**

返回一个代表域代码的 **Range** 对象。可读写。

**语法**

**express.Code**

*express*   一个代表 **Editors** 对象的变量。

**说明**

域代码是包含在域字符 (**{ }**) 中的所有内容，包括前导和尾部空格。无需更改域结果的显示方式即可访问域代码。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示活动文档中每一个域的域代码。 */ function test() { for (let i=1;i<= ActiveDocument.Fields.Count;i++) {     MsgBox("\"" + ActiveDocument.Fields.Item(i).Code.Text + "\"") } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将活动文档中第一个域的域代码改为 CREATEDATE。*/ function test() { let rngTemp = ActiveDocument.Fields.Item(1).Code     rngTemp.Text = " CREATEDATE "     ActiveDocument.Fields.Item(1).Update() }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例判断活动文档中是否包含一个名为“Title”的邮件合并域。*/ function test() { for(let i=1;i<= ActiveDocument.MailMerge.Fields.Count;i++) {     let str=ActiveDocument.MailMerge.Fields.Item(i).Code.Text     if(str.indexOf("Title",1)) {         MsgBox("A Title merge field is in this document")     } } }` |

#### **Editors.Count**

返回一个 **Long** 类型值，该值代表集合中 Editor 对象的数目。只读。

**语法**

**express.Count**

*express*   一个代表 **Editors** 对象的变量。

#### **Editors.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Editors** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **Editors.Data**

返回或设置 ADDIN 域中的数据。可读写 **String** 类型。

**语法**

**express.Data**

*express*   一个代表 **Editors** 对象的变量。

**说明**

该数据在域代码或结果中是不可见的，只有通过返回 **Data** 属性的值才能进行访问。如果该域不是 ADDIN 域，该属性将发生错误。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在活动文档的插入点插入一个 ADDIN 域，然后设置该域的数据。*/ function test() { Selection.Collapse(wdCollapseStart) let fldTemp =ActiveDocument.Fields.Add(Selection.Range,wdFieldAddin) fldTemp.Data = "Hidden information" }  ` |

#### **Editors.Index**

返回一个 **Long** 类型的值，该值代表项目在集合中的位置。只读。

**语法**

**express.Index**

*express*   一个代表 **Editors** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回选定域在 Fields 集合中的位置。*/ let num = Selection.Fields.Item(1).Index` |

#### **Editors.InlineShape**

返回一个 **InlineShape** 对象，该对象可以代表图片、OLE 对象以及作为 INCLUDEPICTURE 或 EMBED 域结果的 ActiveX 控件。

**语法**

**express.InlineShape**

*express*   一个代表 **Editors** 对象的变量。

**说明**

可将 **InlineShape** 对象作为字符处理，并作为字符放置在一行文本中。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回活动文档中与第一个域有关的嵌入式图形的宽度。若要使此示例运行，此域必须是 INCLUDEPICTURE 域。*/ function test() { if( ActiveDocument.Fields.Item(1).Type == wdFieldIncludePicture) {     MsgBox( ActiveDocument.Fields.Item(1).InlineShape.Width) } }` |

#### **Editors.Kind**

返回 **Field** 对象的链接类型。只读 **WdFieldKind** 类型。

**语法**

**express.Kind**

*express*   一个代表 **Editors** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例更新活动文档中的所有热链接域。*/ function test() { let aField = ActiveDocument.Fields for(let i=1; i<=aField.Count; i++) {     if(aField.Item(i).Kind == wdFieldKindWarm) {         aField.Item(i).Update()     } } }` |

#### **Editors.LinkFormat**

返回一个 **LinkFormat** 对象，该对象代表指定域的链接选项。只读。

**语法**

**express.LinkFormat**

*express*   一个代表 **Editors** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例更新活动文档中尚未自动更新的域。*/ function test() { for(let i=1;i<= ActiveDocument.Fields.Count;i++) {     if(ActiveDocument.Fields.Item(i).LinkFormat.AutoUpdate == false) {           ActiveDocument.Fields.Item(i).LinkFormat.Update()     } } }` |

#### **Editors.Locked**

如果指定域已被锁定，则该属性值为 **True**。可读写 **Boolean** 类型。

**语法**

**express.Locked**

*express*   一个代表 **Editors** 对象的变量。

**说明**

当域处于锁定状态时，无法更新该域的结果。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在所选内容的开头插入一个 DATE 域，然后锁定该域。*/ function test() { Selection.Collapse(wdCollapseStart) let myField = ActiveDocument.Fields.Add(Selection.Range,wdFieldDate) myField.Locked = true }` |

#### **Editors.Next**

返回集合中的下一个对象。只读。

**语法**

**express.Next**

*express*   一个代表 **Editors** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果 Next 方法返回一个 Field 对象，并且该域不是 FILLIN 域，则以下示例更新活动文档的第一节中的域。*/ function test() { let myField if (ActiveDocument.Sections.Item(1).Range.Fields.Count >= 1) {     myField = ActiveDocument.Fields.Item(1)     while(myField!=null) {         if(myField.Type != wdFieldFillIn ) {             myField.Update()             myField = myField.Next         }     } } }` |

#### **Editors.OLEFormat**

返回一个 **OLEFormat** 对象，该对象代表指定域的 OLE 特征（链接除外）。只读。

**语法**

**express.OLEFormat**

*express*   一个代表 **Editors** 对象的变量。

#### **Editors.Parent**

返回一个 **Object** 类型值，该值代表指定 **Editors** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **Editors** 对象的变量。

#### **Editors.Previous**

返回集合中的上一个对象。只读。

**语法**

**express.Previous**

*express*   一个代表 **Editors** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示活动文档中倒数第二个域的域代码。*/ function test() { let aField = ActiveDocument.Fields.Item(ActiveDocument.Fields.Count).Previous MsgBox( "Field code = " + aField.Code) }` |

#### **Editors.Result**

返回一个 **Range** 对象，该对象代表域的结果。可读写。

**语法**

**express.Result**

*express*   一个代表 **Editors** 对象的变量。

**说明**

无需改变域代码的显示方式即可获得一个域结果。使用 **Text** 属性可返回 **Range** 对象中的文本。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例对所选内容中的第一个域应用加粗格式。*/ function test() { if(Selection.Fields.Count >= 1) {      let myRange = Selection.Fields.Item(1).Result     myRange.Bold = true } }` |

#### **Editors.ShowCodes**

如果显示指定域的域代码而不是域结果，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.ShowCodes**

*express*   一个代表 **Editors** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例选定下一个域，并显示其域代码。*/ function test() { Selection.GoTo(wdGoToField) Selection.Expand(wdWord) if( Selection.Fields.Count == 1 ) {     Selection.Fields.Item(1).ShowCodes = true } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例更新并显示活动文档中第一个域的结果。*/ function test() { if( ActiveDocument.Fields.Count >= 1) {     ActiveDocument.Fields.Item(1).Update()     ActiveDocument.Fields.Item(1).ShowCodes = false } }` |

#### **Editors.Type**

返回域的类型。只读 **WdFieldType** 类型。

**语法**

**express.Type**

*express*   一个代表 **Editors** 对象的变量。

**说明**

**安全性**??动态数据交换 (DDE) 是一种不安全的陈旧技术。如果可能，请使用比 DDE 更加安全的技术，如对象链接和嵌入 (OLE)。

适用环境：web

适用平台：windows/linux