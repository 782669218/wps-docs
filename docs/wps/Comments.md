#### **Comments**



 

[Comment ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Comment/Comment%20.htm#jsObject_Comment)对象的集合，这些对象代表选定内容、区域或文档中的批注。

**方法**

|                                                              | 名称     | 说明                                                      |
| ------------------------------------------------------------ | -------- | --------------------------------------------------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**  | 返回一个 **Comment** 对象，该对象代表添加到区域中的备注。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item** | 返回集合中的单个 **Comment** 对象。                       |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个代表 WPS 应用程序的 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回一个 **Number** 类型的值，该值代表 **Comments** 集合中的项目数。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回一个 **Object** 类型值，该值代表指定 **Comments** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **ShowBy**      | 返回或设置审阅者的名字，该审阅者的批注显示在批注窗格中。**String** 类型，可读写。 |

**成员方法**

#### **Comments.Add**

返回一个 **Comment** 对象，该对象代表添加到区域中的备注。

**语法**

**express.Add(Range, Text)**

*express*   一个代表 **Comments** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**           |
| -------- | ------------- | ------------ | ------------------ |
| *Range*  | 必选          | **Range**    | 要添加备注的区域。 |
| *Text*   | 可选          | **Variant**  | 备注文本。         |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/* 在选区开始位置，插入一条备注 */ function test(){   Application.Selection.Collapse(wdCollapseEnd);   Application.ActiveDocument.Comments.Add(Application.Selection.Range,"review this"); }  /* 在活动文档的第三段添加一条备注 */ function test(){   let myRange = Application.ActiveDocument.Paragraphs.Item(3).Range   Application.ActiveDocument.Comments.Add(myRange, "original third paragraph") }` |

#### **Comments.Item**

返回集合中的单个 **Comment** 对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **Comments** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Index*  | 必选          | **Number**   | 要返回的单个对象。可以是代表单个对象序号位置的 Number 类型值。 |

**成员属性**

#### **Comments.Application**

返回一个代表 WPS 应用程序的 **Application**[ ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象。

**语法**

**express.Application**

*express*   一个代表 **Comments** 对象的变量。

#### **Comments.Count**

返回一个 **Number** 类型的值，该值代表 **Comments** 集合中的项目数。只读。

**语法**

**express.Count**

*express*   一个代表 **Comments** 对象的变量。

#### **Comments.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Comments** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **Comments.Parent**

返回一个 **Object** 类型值，该值代表指定 **Comments** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **Comments** 对象的变量。

#### **Comments.ShowBy**

返回或设置审阅者的名字，该审阅者的批注显示在批注窗格中。**String** 类型，可读写。

**语法**

**express.ShowBy**

*express*   一个代表 **Comments** 对象的变量。

**说明**

可选择显示一位审阅者或者所有审阅者所做的批注。若要查看所有审阅者的批注，可将本属性设置为“所有审阅者”。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*下面的示例显示“Don Funk”所做的批注。*/ function test() {     if (Application.ActiveDocument.Comments.Count >= 1) {         Application.ActiveDocument.ActiveWindow.View.SplitSpecial = wdPaneComments         Application.ActiveDocument.Comments.ShowBy = "Don Funk"     } } ` |

适用环境：web

适用平台：windows/linux