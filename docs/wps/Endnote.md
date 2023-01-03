#### **Endnote**



代表一个尾注。**Endnote** 对象是 **Endnotes** 集合的一个成员，该集合代表所选内容、范围或文档的尾注。

**说明**

使用 **Endnotes**(*Index*) 可返回一个 **Endnote** 对象，其中 *Index* 是索引号。索引号代表尾注在所选内容、范围或文档中的位置。以下示例将所选内容中的第一个尾注设置成红色。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     if(Application.Selection.Endnotes.Count >= 1) {         Application.Selection.Endnotes.Item(1).Reference.Font.ColorIndex = wdRed     } }` |

使用 

Add

 方法可向 

Endnotes

 集合添加尾注。以下示例紧接着所选内容添加一个尾注。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     Application.Selection.Collapse(wdCollapseEnd)     Application.ActiveDocument.Endnotes.Add(Selection.Range , null, "The Willow Tree, (Lone Creek Press, 1996).") }` |

**方法**

|                                                              | 名称       | 说明             |
| ------------------------------------------------------------ | ---------- | ---------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete** | 删除指定的尾注。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个代表 WPS 应用程序的 [Application](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application) 对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Index**       | 返回一个 **Long** 类型的值，该值代表项目在集合中的位置。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回一个 **Object** 类型值，该值代表指定 **Endnote** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Range**       | 返回 **Range** 对象，该对象表示文档包含在指定对象中的一部分。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Reference**   | 返回 **Range** 对象，该对象表示尾注引用标记。                |

**成员方法**

#### **Endnote.Delete**

删除指定的尾注。

**语法**

**express.Delete()**

*express*   一个代表 **Endnote** 对象的变量。

**成员属性**

#### **Endnote.Application**

返回一个代表 WPS 应用程序的 [Application](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application) 对象。

**语法**

**express.Application**

*express*   一个代表 **Endnote** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **Endnote.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 

**语法**

**express.Creator**

*express*   一个代表 **Endnote** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **Endnote.Index**

返回一个 **Long** 类型的值，该值代表项目在集合中的位置。只读。

**语法**

**express.Index**

*express*   一个代表 **Endnote** 对象的变量。

#### **Endnote.Parent**

返回一个 **Object** 类型值，该值代表指定 **Endnote** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **Endnote** 对象的变量。

#### **Endnote.Range**

返回 **Range** 对象，该对象表示文档包含在指定对象中的一部分。

**语法**

**express.Range**

*express*   一个代表 **Endnote** 对象的变量。

**说明**

有关从文档返回一个区域或从形状集合返回形状区域的信息，请参阅 **Range** 方法。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例更改活动文档中第一个尾注的文本。*/ function test() {   let range = Application.ActiveDocument.Endnotes.Item(1).Range   range.Delete()   range.Text = "new endnote text" }` |

#### **Endnote.Reference**

返回 **Range** 对象，该对象表示尾注引用标记。

**语法**

**express.Reference**

*express*   一个代表 **Endnote** 对象的变量。

**说明**

此示例将 *myRange* 设置为活动文档中的第一个尾注引用标记，然后复制该引用标记。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {   let myRange = Application.ActiveDocument.Endnotes.Item(1).Reference   myRange.Copy() }` |

适用环境：web

适用平台：windows/linux