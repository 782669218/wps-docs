**Comment** 



代表单个备注。**Comment** 对象是 **Comments** 集合的成员。**Comments** 集合中包括选定内容、区域或文档中的备注。

**说明**

使用 **Add** 方法在指定区域添加一个备注。以下示例在紧挨着选定内容后添加一个备注。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `Application.Selection.Collapse(wdCollapseEnd) Application.ActiveDocument.Comments.Add(Application.Selection.Range, "review this")` |

使用 **Reference** 属性返回与指定备注关联的引用标记。使用 **Range** 属性返回与指定备注关联的文本。以下示例显示与活动文档的第一个备注关联的文本。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `alert(Application.ActiveDocument.Comments.Item(1).Range.Text)` |

**方法**

|                                                              | 名称       | 说明                                                      |
| ------------------------------------------------------------ | ---------- | --------------------------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete** | 删除指定的批注。                                          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Edit**   | 在创建指定 OLE 对象的应用程序中打开该对象，以便进行编辑。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个代表 WPS 应用程序的 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Author**      | 返回或设置一个 **String** 类型的值，该值代表某一批注的作者姓名。可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Date**        | 返回代表批注的插入日期和时间的 **Date** 类型。只读。         |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Index**       | 返回一个 **Long** 类型的值，该值代表项目在集合中的位置。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Initial**     | 返回或设置与指定批注相关联的用户的姓名缩写。**String** 类型，可读写。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **IsInk**       | 返回一个 **Boolean** 类型的值，该值代表批注是否为手写批注。  |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回一个 **Object** 类型值，该值代表指定 **Comment** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Range**       | 返回代表备注内容的 **Range** 对象。                          |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Reference**   | 返回代表备注的引用标记的 **Range** 对象。                    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Scope**       | 返回一个 **Range** 对象，该对象代表指定批注所标记的文本范围。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowTip**     | 如果在“屏幕提示”中显示与批注相关联的文本，则该属性值为 **True**。**Boolean** 类型，可读写。 |

**成员方法**

#### **Comment.Delete**

删除指定的批注。

**语法**

**express.Delete()**

*express*   一个代表 **Comment** 对象的变量。

#### **Comment.Edit**

在创建指定 OLE 对象的应用程序中打开该对象，以便进行编辑。

**语法**

**express.Edit()**

*express*   一个代表 **Comment** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例打开（为了进行编辑）活动文档中第一个嵌入式 OLE 对象（定义为图形）。*/ function test() {     let shapesAll = Application.ActiveDocument.Shapes     if (shapesAll.Count >= 1) {         if (shapesAll.Item(1).Type == msoEmbeddedOLEObject) {             shapesAll.Item(1).OLEFormat.Edit()         }     } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例打开（为了进行编辑）活动文档中第一个链接的 OLE 对象（定义为嵌入式图形）。*/ function test() {     let colIS = Application.ActiveDocument.InlineShapes     if (colIS.Count >= 1) {         if (colIS.Item(1).Type == wdInlineShapeLinkedOLEObject) {             colIS.Item(1).OLEFormat.Edit()         }     } }` |

**成员属性**

#### **Comment.Application**

返回一个代表 WPS 应用程序的 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象。

**语法**

**express.Application**

*express*   一个代表 **Comment** 对象的变量。

#### **Comment.Author**

返回或设置一个 **String** 类型的值，该值代表某一批注的作者姓名。可读写。

**语法**

**express.Author**

*express*   一个代表 **Comment** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 本示例设置作者姓名，并且为活动文档的第一个批注设置作者姓名缩写。 */ function test() {     if (Application.ActiveDocument.Comments.Count >= 1) {         let rng = Application.ActiveDocument.Comments.Item(1)         rng.Author = ("Joe Smith")         rng.Initial = ("JAS")     } }` |

#### **Comment.Date**

返回代表批注的插入日期和时间的 **Date** 类型。只读。

**语法**

**express.Date**

*express*   一个代表 **Comment** 对象的变量。

#### **Comment.Index**

返回一个 **Long** 类型的值，该值代表项目在集合中的位置。只读。

**语法**

**express.Index**

*express*   一个代表 **Comment** 对象的变量。

#### **Comment.Initial**

返回或设置与指定批注相关联的用户的姓名缩写。**String** 类型，可读写。

**语法**

**express.Initial**

*express*   一个代表 **Comment** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示选定内容的第一条批注的用户姓名缩写。*/ function test() {     if (Application.Selection.Comments.Count >= 1) {         alert("Comment made by " + Selection.Comments.Item(1).Initial)     } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例检查文档第一节中每条备注的作者姓名缩写。如果作者姓名缩写为“WPSOffice”，本示例则将其改为“KAE”。*/ function test() {     let rngTemp = Application.ActiveDocument.Sections.Item(1).Range     for (let i = 1; i <= rngTemp.Comments.Count; i++) {         if (rngTemp.Comments.Item(i).Initial == "MSOffice") {             rngTemp.Comments.Item(i).Initial = "KAE"         }     } }` |

#### **Comment.IsInk**

返回一个 **Boolean** 类型的值，该值代表批注是否为手写批注。

**语法**

**express.IsInk**

*express*   一个代表 **Comment** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 删除文档中的手写签批 */ function test() {     for(let i = 1; i <=  Application.ActiveDocument.Comments.count; i++){     if(Application.ActiveDocument.Comments.Item(i).IsInk == true){         objComment.Delete()     } }` |

#### **Comment.Parent**

返回一个 **Object** 类型值，该值代表指定 **Comment** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **Comment** 对象的变量。

#### **Comment.Range**

返回代表备注内容的 **Range** 对象。

**语法**

**express.Range**

*express*   一个代表 **Comment** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 修改第一个备注的文本，在前面附加上test文本 */ Application.ActiveDocument.Comments.Item(1).Range.InsertBefore("test")` |

#### **Comment.Reference**

返回代表备注的引用标记的 **Range** 对象。

**语法**

**express.Reference**

*express*   一个代表 **Comment** 对象的变量。

#### **Comment.Scope**

返回一个 **Range** 对象，该对象代表指定批注所标记的文本范围。

**语法**

**express.Scope**

*express*   一个代表 **Comment** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示与选定内容中第一条批注相关联的文本。*/ function test() {     if (Application.Selection.Comments.Count >= 1) {         let myRange = Application.Selection.Comments.Item(1).Scope         alert(myRange.Text)     } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例复制与活动文档中最后一条批注相关联的文本。*/ function test() {     total = Application.ActiveDocument.Comments.Count     if (total >= 1) {         Application.ActiveDocument.Comments.Item(total).Scope.Copy()     } } ` |

#### **Comment.ShowTip**

如果在“屏幕提示”中显示与批注相关联的文本，则该属性值为 **True**。**Boolean** 类型，可读写。

**语法**

**express.ShowTip**

*express*   一个代表 **Comment** 对象的变量。

**说明**

当单击鼠标或按下任意键时，“屏幕提示”才会消失。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示活动文档中的第一条批注的“屏幕提示”。*/ function test() {     if (Application.ActiveDocument.Comments.Count >= 1) {         Application.ActiveDocument.Comments.Item(1).ShowTip = true     } }` |

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例显示活动文档中下一条批注的“屏幕提示”。*/ function test() {     if (Application.ActiveDocument.Comments.Count >= 1) {         Application.Selection.GoTo(wdGoToComment, wdGoToNext)         Application.Selection.MoveEnd(wdWord, 1)         Application.Selection.Comments.Item(1).ShowTip == true     } }` |

适用环境：web

适用平台：windows/linux