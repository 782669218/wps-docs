#### **Hyperlinks**



代表文档、范围或所选内容中 **Hyperlink** 对象的集合。

**方法**

|                                                              | 名称     | 说明                                                         |
| ------------------------------------------------------------ | -------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Add**  | 返回一个 **Hyperlink** 对象，该对象代表添加到区域、选定内容或文档的新的超链接。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/methods.gif) | **Item** | 返回集合中的单个 **Hyperlink** 对象。                        |

**属性**

|                                                              | 名称            | 说明                                                         |
| ------------------------------------------------------------ | --------------- | ------------------------------------------------------------ |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Application** | 返回一个代表 WPS 应用程序的 [Application ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Count**       | 返回一个 **Long** 类型的值，该值代表集合中超链接的数目。只读。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Creator**     | 返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。 |
| ![img](https://qn.cache.wpscdn.cn/encs/doc/office_v19/gif/properties.gif) | **Parent**      | 返回一个 **Object** 类型值，该值代表指定 **Hyperlinks** 对象的父对象。 |

**成员方法**

#### **Hyperlinks.Add**

返回一个 **Hyperlink** 对象，该对象代表添加到区域、选定内容或文档的新的超链接。

**语法**

**express.Add(Anchor, Address, SubAddress, ScreenTip, TextToDisplay, Target)**

*express*   一个代表 **Hyperlinks** 对象的变量。

**参数**

| **名称**        | **必选/可选** | **数据类型** | **说明**                                                     |
| --------------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Anchor*        | 必选          | **Object**   | 要转换为超链接的文本或图形。                                 |
| *Address*       | 可选          | **Variant**  | 指定链接的地址。此地址可以是电子邮件地址、Internet 地址或文件名。请注意，WPS 不检查该地址的正确性。 |
| *SubAddress*    | 可选          | **Variant**  | 目标文件内的位置名，如书签、已命名的区域或幻灯片编号。       |
| *ScreenTip*     | 可选          | **Variant**  | 当鼠标指针放在指定超链接上时显示为“屏幕提示”的文本。默认值为“地址”。 |
| *TextToDisplay* | 可选          | **Variant**  | 指定超链接的显示文本。该参数的值将取代由 Anchor 指定的文本或图形。 |
| *Target*        | 可选          | **Variant**  | 要在其中加载指定超链接的框架或窗口的名字。                   |

**示例**

| 示例代码复制                                                 |
| ------------------------------------------------------------ |
| `/*本示例将选定内容转换为指向万维网上 Microsoft 公司地址的超链接。*/ Application.ActiveDocument.Hyperlinks.Add(Selection.Range, "https://www.wps.cn")  /*本示例将选定内容转换为链接到 MyFile.doc 中名为 MyBookMark 的书签的超链接。*/ Application.ActiveDocument.Hyperlinks.Add(Selection.Range, "C:\\My Documents\\MyFile.doc", "MyBookMark")  /*本示例将选定内容中的第一个图形定义为超链接。*/ Application.ActiveDocument.Hyperlinks.Add(Selection.ShapeRange.Item(1), "https://www.wps.cn")` |

#### **Hyperlinks.Item**

返回集合中的单个 **Hyperlink** 对象。

**语法**

**express.Item(Index)**

*express*   一个代表 **Hyperlinks** 对象的变量。

**参数**

| **名称** | **必选/可选** | **数据类型** | **说明**                                                     |
| -------- | ------------- | ------------ | ------------------------------------------------------------ |
| *Index*  | 必选          | **Variant**  | 要返回的单个对象。可以是代表序号位置的 Long 类型值，或代表单个对象名称的 String 类型值。 |

**成员属性**

#### **Hyperlinks.Application**

返回一个代表 WPS 应用程序的 [Application ](https://qn.cache.wpscdn.cn/encs/doc/office_v19/apiObjectTemplate.htm?page=topics/WPS%20%E5%9F%BA%E7%A1%80%E6%8E%A5%E5%8F%A3/%E6%96%87%E5%AD%97%20API%20%E5%8F%82%E8%80%83/Application/Application%20.htm#jsObject_Application)对象。

**语法**

**express.Application**

*express*   一个代表 **Hyperlinks** 对象的变量。

#### **Hyperlinks.Count**

返回一个 **Long** 类型的值，该值代表集合中超链接的数目。只读。

**语法**

**express.Count**

*express*   一个代表 **Hyperlinks** 对象的变量。

#### **Hyperlinks.Creator**

返回一个 32 位整数，该整数代表在其中创建特定对象的应用程序。只读 **Long** 类型。

**语法**

**express.Creator**

*express*   一个代表 **Hyperlinks** 对象的变量。

**说明**

如果该对象在 WPS 中创建，则 **Creator** 属性返回十六进制数字 4D535744，代表字符串“WPS”。该属性主要设计用于 Macintosh，在 Macintosh 中，每个应用程序都具有四个字符的创建者代码。例如，WPS 的创建者代码是 WPS。有关该属性的其他信息，请参阅 WPS OfficeMacintosh Edition 附带的语言参考帮助。

#### **Hyperlinks.Parent**

返回一个 **Object** 类型值，该值代表指定 **Hyperlinks** 对象的父对象。

**语法**

**express.Parent**

*express*   一个代表 **Hyperlinks** 对象的变量。

适用环境：web

适用平台：windows/linux