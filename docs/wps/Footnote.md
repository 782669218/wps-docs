#### **Footnote**



代表位于页面底端或文字下方的脚注。**Footnote** 对象是 **Footnotes** 集合的一个成员。 **Footnotes** 集合代表所选内容、范围或文档中的脚注。

**说明**

使用 **Footnotes**(*Index*) 可返回一个 **Footnote** 对象，其中 *Index* 是索引号。索引号代表脚注在所选内容、范围或文档中的位置。以下示例对所选内容中的第一个脚注应用红色格式。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `if(Application.Selection.Footnotes.Count >= 1) {     Application.Selection.Footnotes.Item(1).Reference.Font.ColorIndex = wdRed }` |

使用 

Add

 方法可将脚注添加到 

Footnotes

 集合。以下示例紧接所选内容之后插入一条自动编号的脚注。

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `function test() {     Application.Selection.Collapse(wdCollapseEnd)     Application.ActiveDocument.Footnotes.Add(Selection.Range ,"The Willow Tree, (Lone Creek Press, 1996).") }` |

**方法**

|                                                              | 名称       | 说明             |
| ------------------------------------------------------------ | ---------- | ---------------- |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Delete** | 删除指定的脚注。 |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个代表 WPS 应用程序的 **Application** 对象。           |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Index**       | 返回一个 **Long** 类型的值，该值代表项目在集合中的位置。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回一个 **Object** 类型值，该值代表指定 **Footnote** 对象的父对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Range**       | 返回一个 **Range** 对象，该对象代表脚注中包含的文档部分。    |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Reference**   | 返回一个 **Range** 对象，该对象代表脚注引用标记。            |

**成员方法**

#### **Footnote.Delete**

删除指定的脚注。

**语法**

**express.Delete()**

*express*   一个代表 **Footnote** 对象的变量。

**成员属性**

#### **Footnote.Application**

返回一个代表 WPS 应用程序的 **Application** 对象。

**语法**

**express.Application**

*express*   一个代表 **Footnote** 对象的变量。

**说明**

Visual Basic 的 **CreateObject** 和 **GetObject** 函数使您可以从 示例代码 项目中访问 OLE 自动化对象。

#### **Footnote.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Footnote** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **Footnote.Index**

返回一个 **Long** 类型的值，该值代表项目在集合中的位置。只读。

**语法**

**express.Index**

*express*   一个代表 **Footnote** 对象的变量。

#### **Footnote.Parent**

返回一个 **Object** 类型值，该值代表指定 **Footnote** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **Footnote** 对象的变量。

#### **Footnote.Range**

返回一个 **Range** 对象，该对象代表脚注中包含的文档部分。

**语法**

**express.Range**

*express*   一个代表 **Footnote** 对象的变量。

#### **Footnote.Reference**

返回一个 **Range** 对象，该对象代表脚注引用标记。

**语法**

**express.Reference**

*express*   一个代表 **Footnote** 对象的变量。

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*以下示例将 myRange 设置为活动文档中的第一个脚注引用标记，然后复制该引用标记。*/ function test() { let myRange = ActiveDocument.Footnotes.Item(1).Reference myRange.Copy() }` |

适用环境：web

适用平台：windows/linux