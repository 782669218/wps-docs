#### **Field**



代表一个域。**Field** 对象是 **Fields** 集合的一个成员。 **Fields** 集合代表所选内容、范围或文档中的域。

**说明**

使用 **Fields**(*Index*) 可返回一个 **Field** 对象，其中 *Index* 是索引号。索引号代表域在所选内容、范围或文档中的位置。以下示例显示活动文档中第一个域的域代码和结果。

**方法**

|                                                              | 名称             | 说明                                              |
| ------------------------------------------------------------ | ---------------- | ------------------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Copy**         | 将指定域复制到剪贴板。                            |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Cut**          | 将指定域从文档中移到剪贴板上。                    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete**       | 删除指定的域。                                    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **DoClick**      | 单击指定域。                                      |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Select**       | 选择指定的域。                                    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Unlink**       | 将指定域替换为其最新结果。                        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Update**       | 更新域的结果。如果该域更新成功，则返回 **True**。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **UpdateSource** | 将对 INCLUDETEXT 域所做的更改保存回源文档。       |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个代表 WPS 应用程序的 **Application** 对象。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Code**        | 返回一个代表域代码的 **Range**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)对象。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Data**        | 返回或设置 ADDIN 域中的数据。可读写 **String** 类型。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Index**       | 返回一个 **Number** 类型的值，该值代表项目在集合中的位置。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **InlineShape** | 返回一个 **InlineShape** 对象，该对象可以代表图片、OLE 对象以及作为 INCLUDEPICTURE 或 EMBED 域结果的 ActiveX 控件。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Kind**        | 返回 **Field** 对象的链接类型。只读 **WdFieldKind** 类型。   |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **LinkFormat**  | 返回一个 **LinkFormat** 对象，该对象代表指定域的链接选项。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Locked**      | 如果指定域已被锁定，则该属性值为 **True**。可读写 **Boolean** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Next**        | 返回集合中的下一个对象。只读。                               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **OLEFormat**   | 返回一个 **OLEFormat** 对象，该对象代表指定域的 OLE 特征（链接除外）。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回一个 **Object** 类型值，该值代表指定 **Field** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Previous**    | 返回集合中的上一个对象。只读。                               |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Result**      | 返回一个 **Range** 对象，该对象代表域的结果。可读写。        |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowCodes**   | 如果显示指定域的域代码而不是域结果，则该属性值为 **True**。**Boolean** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Type**        | 返回域的类型。只读 [WdFieldType ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdFieldType%20%E6%9E%9A%E4%B8%BE.html)类型。 |

**成员方法**

#### **Field.Copy**

将指定域复制到剪贴板。

**语法**

**express.Copy()**

*express*   一个代表 **Field** 对象的变量。

#### **Field.Cut**

将指定域从文档中移到剪贴板上。

**语法**

**express.Cut()**

*express*   一个代表 **Field** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例删除活动文档的第一个域，并将该域粘贴到插入点 */ function test() { 	if(Application.ActiveDocument.Fields.Count>=1){ 		Application.ActiveDocument.Fields.Item(1).Cut() 		Application.Selection.Collapse(wdCollapseEnd) 		Application.Selection.Paste() 	} }` |

#### **Field.Delete**

删除指定的域。

**语法**

**express.Delete()**

*express*   一个代表 **Field** 对象的变量。

#### **Field.DoClick**

单击指定域。

**语法**

**express.DoClick()**

*express*   一个代表 **Field** 对象的变量。

**说明**

如果是 GOTOBUTTON 域，该方法将插入点移至指定位置或选定指定的书签。如果是 MACROBUTTON 域，该方法将运行指定的宏。如果是 HYPERLINK 域，该方法将跳转至目标位置。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果所选内容中的第一个域为 GOTOBUTTON 域，本示例可实现：单击此域（将插入点移至指定位置，或选择指定的书签）。*/ function test() { let fldTemp = Selection.Fields.Item(1) if(fldTemp.Type == wdFieldGoToButton){     fldTemp.DoClick() } }` |

#### **Field.Select**

选择指定的域。

**语法**

**express.Select()**

*express*   一个代表 **Field** 对象的变量。

**说明**

使用此方法后，可使用 **Selection** 对象处理所选项目。有关详细信息，请参阅 处理 [Selection](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Selection/Selection%20.htm#jsObject_Selection) 对象。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 以下示例更新并选定活动文档中的第一个域 */ function test() { 	if(Application.ActiveDocument.Fields.Count >= 1 ){ 		Application.ActiveDocument.Fields.Item(1).Update() 		Application.ActiveDocument.Fields.Item(1).Select() 	} }` |

#### **Field.Unlink**

将指定域替换为其最新结果。

**语法**

**express.Unlink()**

*express*   一个代表 **Field** 对象的变量。

**说明**

如果断开一个域的链接，则当前结果会转换为文字或图形，并且无法再自动更新。注意，某些域（例如，XE（索引项）域和 SEQ（序列）域）不能断开链接。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例断开“Sales.doc”中的第一个域的链接。*/ Documents.Item("Sales.doc").Fields.Item(1).Unlink()` |

#### **Field.Update**

更新域的结果。如果该域更新成功，则返回 **True**。

**语法**

**express.Update()**

*express*   一个代表 **Field** 对象的变量。

**返回值**

Boolean

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例更新活动文档中的第一个域。返回 0（零）值表示更新域时没有出错。*/ function test() { if (ActiveDocument.Fields.Item(1).Update() == 0){     MsgBox( "Update Successful") } else{     MsgBox("Field "+ActiveDocument.Fields.Item(1).Update()+" has an error") } }` |

#### **Field.UpdateSource**

将对 INCLUDETEXT 域所做的更改保存回源文档。

**语法**

**express.UpdateSource()**

*express*   一个代表 **Field** 对象的变量。

**说明**

源文档必须是 WPS 文档格式。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例更新活动文档中的 INCLUDETEXT 域。*/ function test() { for (let i=1; i<= ActiveDocument.Fields.Count; i++) {     if (ActiveDocument.Fields.Item(i).Type == wdFieldIncludeText) {          ActiveDocument.Fields.Item(i).UpdateSource()         } } }` |

**成员属性**

#### **Field.Application**

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

*express*   一个代表 **Field** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **Field.Code**

返回一个代表域代码的 **Range**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Range/Range%20.htm#jsObject_Range)对象。可读写。

**语法**

**express.Code**

*express*   一个代表 **Field** 对象的变量。

**说明**

域代码是包含在域字符 (**{ }**) 中的所有内容，包括前导和尾部空格。无需更改域结果的显示方式即可访问域代码。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 获取当前文档的域代码 */ Application.ActiveDocument.Fields.Item(1).Code` |

#### **Field.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Field** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **Field.Data**

返回或设置 ADDIN 域中的数据。可读写 **String** 类型。

**语法**

**express.Data**

*express*   一个代表 **Field** 对象的变量。

**说明**

该数据在域代码或结果中是不可见的，只有通过返回 **Data** 属性的值才能进行访问。如果该域不是 ADDIN 域，该属性将发生错误。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在活动文档的插入点插入一个 ADDIN 域，然后设置该域的数据。*/ function test() { Selection.Collapse(wdCollapseStart) let fldTemp =ActiveDocument.Fields.Add(Selection.Range,wdFieldAddin) fldTemp.Data = "Hidden information" }` |

#### **Field.Index**

返回一个 **Number** 类型的值，该值代表项目在集合中的位置。只读。

**语法**

**express.Index**

*express*   一个代表 **Field** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 返回选定域再Fields集合中的位置 */ Application.Selection.Fields.Item(1).Index` |

#### **Field.InlineShape**

返回一个 **InlineShape** 对象，该对象可以代表图片、OLE 对象以及作为 INCLUDEPICTURE 或 EMBED 域结果的 ActiveX 控件。

**语法**

**express.InlineShape**

*express*   一个代表 **Field** 对象的变量。

**说明**

可将 **InlineShape** 对象作为字符处理，并作为字符放置在一行文本中。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例返回活动文档中与第一个域有关的嵌入式图形的宽度。若要使此示例运行，此域必须是 INCLUDEPICTURE 域。*/ function test() { if( ActiveDocument.Fields.Item(1).Type == wdFieldIncludePicture) {     MsgBox( ActiveDocument.Fields.Item(1).InlineShape.Width) } }` |

#### **Field.Kind**

返回 **Field** 对象的链接类型。只读 **WdFieldKind** 类型。

**语法**

**express.Kind**

*express*   一个代表 **Field** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例更新活动文档中的所有热链接域。*/ function test() { let aField = ActiveDocument.Fields for(let i=1; i<=aField.Count; i++) {     if(aField.Item(i).Kind == wdFieldKindWarm) {         aField.Item(i).Update()     } } }` |

#### **Field.LinkFormat**

返回一个 **LinkFormat** 对象，该对象代表指定域的链接选项。只读。

**语法**

**express.LinkFormat**

*express*   一个代表 **Field** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例更新活动文档中尚未自动更新的域。*/ function test() { for(let i=1;i<= ActiveDocument.Fields.Count;i++) {     if(ActiveDocument.Fields.Item(i).LinkFormat.AutoUpdate == false) {           ActiveDocument.Fields.Item(i).LinkFormat.Update()     } } }` |

#### **Field.Locked**

如果指定域已被锁定，则该属性值为 **True**。可读写 **Boolean** 类型。

**语法**

**express.Locked**

*express*   一个代表 **Field** 对象的变量。

**说明**

当域处于锁定状态时，无法更新该域的结果。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例在所选内容的开头插入一个 DATE 域，然后锁定该域。*/ function test() { Selection.Collapse(wdCollapseStart) let myField = ActiveDocument.Fields.Add(Selection.Range,wdFieldDate) myField.Locked = true }` |

#### **Field.Next**

返回集合中的下一个对象。只读。

**语法**

**express.Next**

*express*   一个代表 **Field** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*如果 Next 方法返回一个 Field 对象，并且该域不是 FILLIN 域，则以下示例更新活动文档的第一节中的域。*/ function test() { let myField if (ActiveDocument.Sections.Item(1).Range.Fields.Count >= 1) {     myField = ActiveDocument.Fields.Item(1)     while(myField!=null) {         if(myField.Type != wdFieldFillIn ) {             myField.Update()             myField = myField.Next         }     } } }` |

#### **Field.OLEFormat**

返回一个 **OLEFormat** 对象，该对象代表指定域的 OLE 特征（链接除外）。只读。

**语法**

**express.OLEFormat**

*express*   一个代表 **Field** 对象的变量。

#### **Field.Parent**

返回一个 **Object** 类型值，该值代表指定 **Field** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **Field** 对象的变量。

#### **Field.Previous**

返回集合中的上一个对象。只读。

**语法**

**express.Previous**

*express*   一个代表 **Field** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例显示活动文档中倒数第二个域的域代码。*/ function test() { let aField = ActiveDocument.Fields.Item(ActiveDocument.Fields.Count).Previous MsgBox( "Field code = " + aField.Code) }` |

#### **Field.Result**

返回一个 **Range** 对象，该对象代表域的结果。可读写。

**语法**

**express.Result**

*express*   一个代表 **Field** 对象的变量。

**说明**

无需改变域代码的显示方式即可获得一个域结果。使用 **Text** 属性可返回 **Range** 对象中的文本。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 将选区第一个域结果加粗 */ function test() { 	if(Application.Selection.Fields.Count >= 1) {  		let myRange = Application.Selection.Fields.Item(1).Result 		myRange.Bold = true 	} }` |

#### **Field.ShowCodes**

如果显示指定域的域代码而不是域结果，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.ShowCodes**

*express*   一个代表 **Field** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例更新当前文档第一个域，并显示其域代码 */ function test() { 	if( Application.ActiveDocument.Fields.Count >= 1) { 		Application.ActiveDocument.Fields.Item(1).Update() 		Application.ActiveDocument.Fields.Item(1).ShowCodes = true 	} }` |

#### **Field.Type**

返回域的类型。只读 [WdFieldType ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/%E6%9E%9A%E4%B8%BE/WdFieldType%20%E6%9E%9A%E4%B8%BE.html)类型。

**语法**

**express.Type**

*express*   一个代表 **Field** 对象的变量。

适用环境：web

适用平台：windows/linux